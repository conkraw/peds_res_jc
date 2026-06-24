from __future__ import annotations

import json
import re
from copy import deepcopy
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from pptx_builder import build_powerpoint
from docx_builder import build_word_summary
from printable_form_builder import build_printable_planning_form
from github_storage import (
    GitHubDraftLoadError,
    GitHubDraftSaveError,
    github_backup_is_configured,
    github_config_status_message,
    list_drafts_from_github,
    load_draft_from_github,
    save_draft_to_github,
    save_article_to_github,
    load_file_bytes_from_github,
)
from slide_schema import SLIDES, make_default_deck


APP_TITLE = "Journal Club PowerPoint Builder"
PROJECT_VERSION = "0.2.6"


# -----------------------------
# Utility functions
# -----------------------------


def count_words(text: Any) -> int:
    return len(re.findall(r"\b\w+\b", str(text or "")))


def nonempty_lines(text: Any) -> List[str]:
    return [line.strip() for line in str(text or "").splitlines() if line.strip()]


def clear_widget_state() -> None:
    for key in list(st.session_state.keys()):
        if key.startswith(("widget__", "table__", "cell__")):
            del st.session_state[key]


def normalize_deck(deck_or_payload: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """Merge uploaded/old draft data into the current schema.

    Supports both older draft JSON files that contain only the deck and newer
    GitHub backup files that wrap the deck with presenter/session metadata.
    """
    deck = deck_or_payload.get("deck", deck_or_payload) if isinstance(deck_or_payload, dict) else {}
    default = make_default_deck()
    for slide in SLIDES:
        sid = slide["id"]
        if isinstance(deck, dict) and sid in deck and isinstance(deck[sid], dict):
            for field in slide["fields"]:
                fkey = field["key"]
                if fkey in deck[sid]:
                    default[sid][fkey] = deck[sid][fkey]
    return default


def initialize_state() -> None:
    if "deck" not in st.session_state:
        st.session_state.deck = make_default_deck()
    if "selected_slide_id" not in st.session_state:
        st.session_state.selected_slide_id = SLIDES[0]["id"]
    if "include_facilitator_notes" not in st.session_state:
        st.session_state.include_facilitator_notes = True


def nav_label(slide: Dict[str, Any]) -> str:
    """Short, user-friendly labels for the sidebar only."""
    labels = {
        "title_goal": "Title",
        "opening_case": "Opening Case",
        "patient_problem": "Slide 1: Patient Problem",
        "pico": "Slide 2: PICO",
        "study_design": "Slide 3: Study Design",
        "main_result": "Slide 4: Main Result",
        "clinical_bottom_line": "Slide 5: Clinical Bottom Line",
        "paper_framework": "PAPER Framework",
        "month_skill": "Monthly Skill",
        "apply_back": "Apply Back to Patient",
        "final_bottom_line": "Final Bottom Line",
    }
    return labels.get(slide["id"], slide["label"])


def slide_display_title(slide: Dict[str, Any]) -> str:
    """Actual slide title shown in the main workspace and preview.

    This intentionally does not use nav_label(), because nav_label() is only
    for the sidebar. Keeping these separate prevents duplicate headings like
    'Opening Case: Opening patient case'.
    """
    return str(slide.get("export_title") or slide.get("label") or "Untitled slide").strip()


def field_is_visible(slide_data: Dict[str, Any], field: Dict[str, Any]) -> bool:
    condition = field.get("show_if")
    if not condition:
        return True
    for controlling_key, expected_value in condition.items():
        if slide_data.get(controlling_key) != expected_value:
            return False
    return True


def parse_table_columns(raw_columns: Any) -> List[str]:
    """Convert a comma-separated column field into a clean list of column names."""
    if isinstance(raw_columns, list):
        candidates = raw_columns
    else:
        candidates = str(raw_columns or "").replace("\n", ",").split(",")

    columns: List[str] = []
    for col in candidates:
        cleaned = str(col).strip()
        if cleaned and cleaned not in columns:
            columns.append(cleaned)
    return columns


def sync_selected_slide() -> None:
    """Update the selected slide before the rest of the page renders.

    Streamlit reruns the script after each widget interaction. Using an
    on_change callback prevents the sidebar radio button from feeling like it
    needs a second click, especially after editing text/table widgets.
    """
    label_to_id = {nav_label(slide): slide["id"] for slide in SLIDES}
    selected_label = st.session_state.get("selected_slide_label")
    if selected_label in label_to_id:
        st.session_state.selected_slide_id = label_to_id[selected_label]


def validate_field(value: Any, field: Dict[str, Any]) -> List[str]:
    problems: List[str] = []
    label = field.get("label", field.get("key", "Field"))

    if field.get("required"):
        if field.get("type") == "table":
            if not value:
                problems.append(f"{label} is required.")
        elif str(value or "").strip() == "":
            problems.append(f"{label} is required.")

    if field.get("type") == "number":
        try:
            numeric_value = float(value)
            min_value = field.get("min_value")
            max_value = field.get("max_value")
            if min_value is not None and numeric_value < float(min_value):
                problems.append(f"{label} must be at least {min_value}.")
            if max_value is not None and numeric_value > float(max_value):
                problems.append(f"{label} must be no more than {max_value}.")
        except (TypeError, ValueError):
            problems.append(f"{label} must be a number.")
        return problems

    if field.get("type") == "table":
        rows = value if isinstance(value, list) else []
        max_rows = field.get("max_rows")
        if max_rows is not None and len(rows) > max_rows:
            problems.append(f"{label} has {len(rows)}/{max_rows} rows.")
        for row_idx, row in enumerate(rows, start=1):
            for col, cell_value in row.items():
                if count_words(cell_value) > 10:
                    problems.append(f"{label} row {row_idx}, {col} is long ({count_words(cell_value)} words).")
        return problems

    text = str(value or "")

    if "max_words" in field:
        words = count_words(text)
        if words > int(field["max_words"]):
            problems.append(f"{label} has {words}/{field['max_words']} words.")

    if "max_chars" in field:
        chars = len(text)
        if chars > int(field["max_chars"]):
            problems.append(f"{label} has {chars}/{field['max_chars']} characters.")

    if "max_lines" in field:
        lines = len(nonempty_lines(text))
        if lines > int(field["max_lines"]):
            problems.append(f"{label} has {lines}/{field['max_lines']} lines.")

    if "max_words_per_line" in field:
        for line_number, line in enumerate(nonempty_lines(text), start=1):
            words = count_words(line)
            if words > int(field["max_words_per_line"]):
                problems.append(
                    f"{label} line {line_number} has {words}/{field['max_words_per_line']} words."
                )

    return problems


def validate_deck(deck: Dict[str, Dict[str, Any]]) -> List[str]:
    problems: List[str] = []
    for slide in SLIDES:
        sid = slide["id"]
        slide_data = deck.get(sid, {})
        for field in slide["fields"]:
            if not field_is_visible(slide_data, field):
                continue
            value = slide_data.get(field["key"], "")
            for problem in validate_field(value, field):
                problems.append(f"{slide['label']} — {problem}")
    return problems


def progress_summary(deck: Dict[str, Dict[str, Any]]) -> tuple[int, int]:
    visible_fields = 0
    filled_fields = 0
    for slide in SLIDES:
        slide_data = deck.get(slide["id"], {})
        for field in slide["fields"]:
            if not field_is_visible(slide_data, field):
                continue
            visible_fields += 1
            value = slide_data.get(field["key"], "")
            if field.get("type") == "table":
                if value:
                    filled_fields += 1
            elif str(value or "").strip():
                filled_fields += 1
    return filled_fields, visible_fields


# -----------------------------
# Rendering functions
# -----------------------------


def render_text_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    key = field["key"]
    widget_key = f"widget__{slide_id}__{key}"
    if widget_key not in st.session_state:
        st.session_state[widget_key] = slide_data.get(key, field.get("default", ""))

    label = field["label"]
    help_text = field.get("guide")

    if field["type"] == "textarea":
        #value = st.text_area(label, key=widget_key, help=help_text, height=130)
        value = st.text_area(label,key=widget_key,help=help_text,height=field.get("height", 160))
    else:
        value = st.text_input(label, key=widget_key, help=help_text)

    slide_data[key] = value
    return value


def render_number_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    key = field["key"]
    widget_key = f"widget__{slide_id}__{key}"
    default_value = float(slide_data.get(key, field.get("default", 0.0)) or 0.0)

    value = st.number_input(
        field["label"],
        min_value=float(field.get("min_value", 0.0)),
        value=default_value,
        key=widget_key,
        help=field.get("guide"),
    )
    slide_data[key] = value
    return value


def render_select_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    key = field["key"]
    options = field.get("options", [])
    current = slide_data.get(key, field.get("default", options[0] if options else ""))
    if current not in options and options:
        current = options[0]
    widget_key = f"widget__{slide_id}__{key}"
    index = options.index(current) if current in options else 0

    value = st.selectbox(
        field["label"],
        options=options,
        index=index,
        key=widget_key,
        help=field.get("guide"),
    )
    slide_data[key] = value
    return value


def slug_for_widget(value: Any) -> str:
    """Create a stable widget-key fragment from a table column name."""
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", str(value or "").strip()).strip("_")
    return slug or "column"


def clear_table_cell_state(slide_id: str, table_key: str) -> None:
    """Clear only the widget keys that belong to one manually edited table."""
    prefix = f"cell__{slide_id}__{table_key}__"
    for session_key in list(st.session_state.keys()):
        if session_key.startswith(prefix):
            del st.session_state[session_key]


def normalize_table_rows(rows: Any, columns: List[str]) -> List[Dict[str, str]]:
    """Return table rows with exactly the selected columns and string values.

    This keeps old draft JSON files compatible even when the user renames or
    removes columns in the app. If a user renames a column, values are carried
    over by column position when an exact column-name match is not available.
    """
    if not isinstance(rows, list):
        rows = []

    normalized: List[Dict[str, str]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue

        old_columns = list(row.keys())
        normalized_row: Dict[str, str] = {}
        for col_index, col in enumerate(columns):
            if col in row:
                value = row.get(col, "")
            elif col_index < len(old_columns):
                value = row.get(old_columns[col_index], "")
            else:
                value = ""
            normalized_row[col] = str(value or "")

        normalized.append(normalized_row)
    return normalized


def render_table_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    """Render a stable, beginner-friendly table editor.

    Earlier versions used st.data_editor. That looks elegant, but with Streamlit
    reruns it can feel like a cell needs to be deleted/typed twice before the
    change sticks. For this app, individual text inputs are more reliable and
    easier for non-technical users.
    """
    key = field["key"]

    current_rows = deepcopy(slide_data.get(key, field.get("default", [])))
    current_df = pd.DataFrame(current_rows)

    columns_key = field.get("columns_key")
    configured_columns = parse_table_columns(slide_data.get(columns_key, "")) if columns_key else []
    if not configured_columns:
        configured_columns = list(current_df.columns) or list(field.get("columns", []))

    if not configured_columns:
        st.info("Add at least one table column above.")
        slide_data[key] = []
        return []

    rows = normalize_table_rows(current_rows, configured_columns)
    if not rows:
        rows = [{col: "" for col in configured_columns}]

    max_rows = int(field.get("max_rows", 8) or 8)

    st.caption(field.get("guide", ""))

    control_cols = st.columns([1, 1, 3])
    with control_cols[0]:
        if st.button(
            "Add row",
            key=f"table_add__{slide_id}__{key}",
            disabled=len(rows) >= max_rows,
            use_container_width=True,
        ):
            rows.append({col: "" for col in configured_columns})
            slide_data[key] = rows
            clear_table_cell_state(slide_id, key)
            st.rerun()

    with control_cols[1]:
        if st.button(
            "Remove row",
            key=f"table_remove__{slide_id}__{key}",
            disabled=len(rows) <= 1,
            use_container_width=True,
        ):
            rows = rows[:-1]
            slide_data[key] = rows
            clear_table_cell_state(slide_id, key)
            st.rerun()

    with control_cols[2]:
        st.caption(f"{len(rows)} / {max_rows} rows")

    header_cols = st.columns([1 for _ in configured_columns])
    for header_col, column_name in zip(header_cols, configured_columns):
        header_col.markdown(f"**{column_name}**")

    updated_rows: List[Dict[str, str]] = []
    column_signature = "__".join(slug_for_widget(col) for col in configured_columns)

    for row_index, row in enumerate(rows):
        row_cols = st.columns([1 for _ in configured_columns])
        updated_row: Dict[str, str] = {}

        for col_index, column_name in enumerate(configured_columns):
            cell_key = (
                f"cell__{slide_id}__{key}__{row_index}__"
                f"{column_signature}__{slug_for_widget(column_name)}"
            )
            if cell_key not in st.session_state:
                st.session_state[cell_key] = row.get(column_name, "")

            updated_row[column_name] = row_cols[col_index].text_input(
                f"{column_name} row {row_index + 1}",
                key=cell_key,
                label_visibility="collapsed",
            )

        updated_rows.append(updated_row)

    slide_data[key] = updated_rows
    return updated_rows


def render_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> None:
    if not field_is_visible(slide_data, field):
        return

    ftype = field["type"]
    if ftype in {"text", "textarea"}:
        value = render_text_field(slide_id, slide_data, field)
    elif ftype == "number":
        value = render_number_field(slide_id, slide_data, field)
    elif ftype == "select":
        value = render_select_field(slide_id, slide_data, field)
    elif ftype == "table":
        value = render_table_field(slide_id, slide_data, field)
    else:
        st.error(f"Unsupported field type: {ftype}")
        return

    problems = validate_field(value, field)

    if ftype in {"text", "textarea"}:
        metric_parts = []
        if "max_words" in field:
            metric_parts.append(f"{count_words(value)} / {field['max_words']} words")
        if "max_lines" in field:
            metric_parts.append(f"{len(nonempty_lines(value))} / {field['max_lines']} lines")
        if metric_parts:
            st.caption(" · ".join(metric_parts))

    for problem in problems:
        st.warning(problem)


def render_slide_preview(slide: Dict[str, Any], slide_data: Dict[str, Any]) -> None:
    preview_title = slide_display_title(slide)
    st.caption(f"PowerPoint slide title: {preview_title}")

    for field in slide["fields"]:
        if not field_is_visible(slide_data, field):
            continue
        value = slide_data.get(field["key"], "")
        if field["type"] == "table":
            st.markdown(f"**{field['label']}**")
            st.dataframe(pd.DataFrame(value), hide_index=True, use_container_width=True)
        elif field["type"] == "select":
            st.caption(f"{field['label']}: {value}")
        else:
            st.markdown(f"**{field['label']}**")
            if field["type"] == "textarea":
                st.write(value if str(value).strip() else "—")
            else:
                st.write(value if str(value).strip() else "—")


def render_validation_panel(deck: Dict[str, Dict[str, Any]]) -> None:
    problems = validate_deck(deck)
    filled, total = progress_summary(deck)
    st.progress(filled / total if total else 0)
    st.caption(f"{filled}/{total} visible fields completed")

    if problems:
        with st.expander(f"Validation warnings ({len(problems)})", expanded=True):
            for problem in problems:
                st.write(f"- {problem}")
    else:
        st.success("All visible fields are within limits.")


def default_session_title(deck: Dict[str, Dict[str, Any]]) -> str:
    title_slide = deck.get("title_goal", {})
    return str(title_slide.get("session_title") or title_slide.get("article_title") or "Journal Club").strip()


def friendly_draft_label(filename: str) -> str:
    """Turn 2026-06-22_jane-smith_oxykids-trial.json into a readable option."""
    stem = str(filename or "").removesuffix(".json")
    parts = stem.split("_", 2)
    if len(parts) == 3:
        saved_date, presenter_slug, title_slug = parts
        presenter = presenter_slug.replace("-", " ").title()
        title = title_slug.replace("-", " ").title()
        return f"{saved_date} — {title} ({presenter})"
    return filename


def apply_loaded_payload_to_session(loaded: Dict[str, Any]) -> None:
    """Load a draft payload into the active app session and refresh widgets.

    Do not directly assign Streamlit widget keys here. This function can be
    called after widgets are already rendered in the same run, and Streamlit
    disallows changing widget-backed session_state keys after instantiation.
    """
    st.session_state.deck = normalize_deck(loaded)
    clear_widget_state()


def render_github_backup(deck: Dict[str, Dict[str, Any]]) -> None:
    with st.expander("Backup draft to GitHub", expanded=False):
        st.caption(
            "Optional safety net. This saves the editable JSON draft to your configured private GitHub drafts repo."
        )

        presenter_name = st.text_input(
            "Presenter name",
            key="github_presenter_name",
            placeholder="Jane Smith",
            help="Used only for the saved JSON filename and metadata.",
        )

        current_deck_title = default_session_title(deck)
        previous_deck_title = st.session_state.get("_last_github_deck_title", "")
        current_saved_title = st.session_state.get("github_session_title", "")
        if (
            "github_session_title" not in st.session_state
            or not str(current_saved_title).strip()
            or current_saved_title == previous_deck_title
        ):
            st.session_state.github_session_title = current_deck_title
        st.session_state._last_github_deck_title = current_deck_title

        session_title = st.text_input(
            "Session title for saved filename",
            key="github_session_title",
            help="Used for the saved JSON filename. You can keep the default.",
        )

        if github_backup_is_configured():
            st.success(github_config_status_message())
        else:
            st.info(github_config_status_message())
            st.caption("Add Streamlit secrets first. See README.md for setup instructions.")

        if st.button("Save draft to GitHub", use_container_width=True):
            if not presenter_name.strip():
                st.error("Please enter the presenter name before saving.")
                return
            if not session_title.strip():
                st.error("Please enter the session title before saving.")
                return

            try:
                result = save_draft_to_github(
                    deck=deck,
                    presenter_name=presenter_name,
                    session_title=session_title,
                    app_version=PROJECT_VERSION,
                )
                
                messages = [f"Draft saved to GitHub: {result.path}"]
                
                article_file = st.session_state.get("uploaded_article_pdf")
                if article_file is not None:
                    article_result = save_article_to_github(
                        article_bytes=article_file.getvalue(),
                        original_filename=article_file.name,
                        presenter_name=presenter_name,
                        session_title=session_title,
                    )
                    messages.append(f"Article saved to GitHub: {article_result.path}")
                
                st.success("\n".join(messages))
                if result.html_url:
                    st.caption("You can retrieve it from the drafts repo later and upload it with Load a saved draft JSON.")
            except GitHubDraftSaveError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Unexpected GitHub save error: {exc}")

            article_file = st.session_state.get("uploaded_article_pdf")

            if article_file is None:
                st.warning("No journal article PDF uploaded yet.")


def render_github_recovery() -> None:
    with st.expander("Reload draft from GitHub", expanded=False):
        st.caption(
            "Find a saved JSON draft in the private GitHub drafts repo and load it back into the app."
        )

        if github_backup_is_configured():
            st.success(github_config_status_message())
        else:
            st.info(github_config_status_message())
            st.caption("Add Streamlit secrets first. See README.md for setup instructions.")

        recover_name = st.text_input(
            "Presenter name to search",
            key="github_recover_presenter_name",
            placeholder="Jane Smith",
            help="Searches filenames saved with this presenter name.",
        )

        if st.button("Find saved drafts", use_container_width=True):
            if not recover_name.strip():
                st.error("Please enter a presenter name to search.")
            else:
                try:
                    drafts = list_drafts_from_github(recover_name)
                    st.session_state.github_draft_results = drafts
                    if drafts:
                        st.success(f"Found {len(drafts)} saved draft(s).")
                    else:
                        st.info("No saved drafts found for that presenter name.")
                except GitHubDraftLoadError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Unexpected GitHub load error: {exc}")

        drafts = st.session_state.get("github_draft_results", [])
        if drafts:
            label_to_path = {friendly_draft_label(d["name"]): d["path"] for d in drafts}
            option_labels = list(label_to_path.keys())
            if st.session_state.get("selected_github_draft_label") not in label_to_path:
                st.session_state.selected_github_draft_label = option_labels[0]

            selected_label = st.selectbox(
                "Choose a saved draft",
                option_labels,
                key="selected_github_draft_label",
            )

            selected_path = label_to_path[selected_label]
            st.caption(f"GitHub path: {selected_path}")

            if st.button("Load selected draft", use_container_width=True):
                try:
                    loaded = load_draft_from_github(selected_path)
                    apply_loaded_payload_to_session(loaded)
                    st.success(f"Loaded draft: {selected_label}")
                    st.rerun()
                except GitHubDraftLoadError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Unexpected GitHub load error: {exc}")


def render_downloads(deck: Dict[str, Dict[str, Any]]) -> None:
    problems = validate_deck(deck)
    timestamp = datetime.now().strftime("%Y%m%d")

    include_notes = st.checkbox(
        "Include facilitator notes appendix slide",
        value=st.session_state.include_facilitator_notes,
        key="include_facilitator_notes",
    )

    pptx_bytes = build_powerpoint(deck, include_facilitator_notes=include_notes)
    st.download_button(
        "Download PowerPoint",
        data=pptx_bytes,
        file_name=f"journal_club_deck_{timestamp}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        disabled=bool(problems),
        use_container_width=True,
    )

    docx_bytes = build_word_summary(deck)
    st.download_button(
        "Download 1-page summary",
        data=docx_bytes,
        file_name=f"journal_club_summary_{timestamp}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        disabled=bool(problems),
        use_container_width=True,
    )

    printable_form = build_printable_planning_form(deck)
    st.download_button(
        "Download printable planning form",
        data=printable_form,
        file_name="journal_club_printable_planning_form.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    draft_json = json.dumps(deck, indent=2, ensure_ascii=False).encode("utf-8")
    st.download_button(
        "Download editable draft JSON",
        data=draft_json,
        file_name=f"journal_club_draft_{timestamp}.json",
        mime="application/json",
        use_container_width=True,
    )

    st.divider()
    render_github_backup(deck)
    render_github_recovery()


# -----------------------------
# Main app
# -----------------------------


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    initialize_state()

    st.title(APP_TITLE)
    st.caption(
        "Choose a slide on the left, complete the fields in the main workspace, then export a standardized PowerPoint and one-page summary. Feedback links are added automatically."
    )

    with st.sidebar:
        st.header("Slides")
        st.caption("Pick a slide here. Edit the fields on the main page.")

        label_to_id = {nav_label(slide): slide["id"] for slide in SLIDES}
        id_to_label = {slide["id"]: nav_label(slide) for slide in SLIDES}

        # Keep the displayed radio value and the internal slide id synchronized.
        # The explicit key + callback prevents occasional "double-click" behavior
        # when moving between slides after editing fields.
        if "selected_slide_label" not in st.session_state:
            st.session_state.selected_slide_label = id_to_label[st.session_state.selected_slide_id]
        elif st.session_state.selected_slide_label not in label_to_id:
            st.session_state.selected_slide_label = id_to_label[SLIDES[0]["id"]]
            st.session_state.selected_slide_id = SLIDES[0]["id"]

        st.radio(
            "Choose slide",
            list(label_to_id.keys()),
            key="selected_slide_label",
            on_change=sync_selected_slide,
        )

        st.divider()

        with st.expander("Advanced: drafts/reset", expanded=False):
            uploaded = st.file_uploader("Load a saved draft JSON", type=["json"])
            if uploaded is not None:
                if st.button("Load uploaded draft", use_container_width=True):
                    try:
                        loaded = json.loads(uploaded.getvalue().decode("utf-8"))
                        st.session_state.deck = normalize_deck(loaded)
                        clear_widget_state()
                        st.success("Draft loaded.")
                    except Exception as exc:
                        st.error(f"Could not load draft: {exc}")

            if st.button("Reset to OxyKids example", use_container_width=True):
                st.session_state.deck = make_default_deck()
                clear_widget_state()
                st.success("Reset complete.")

            if st.button("Clear all fields", use_container_width=True):
                st.session_state.deck = make_default_deck()
            
                for slide in SLIDES:
                    for field in slide["fields"]:
                        field_key = field["key"]
                        if field.get("type") == "table":
                            st.session_state.deck[slide["id"]][field_key] = []
                        else:
                            st.session_state.deck[slide["id"]][field_key] = ""
            
                clear_widget_state()
                st.success("All fields cleared.")
                #st.rerun()

    selected_slide = next(slide for slide in SLIDES if slide["id"] == st.session_state.selected_slide_id)
    selected_slide_data = st.session_state.deck[selected_slide["id"]]

    editor_col, export_col = st.columns([2.2, 0.9])

    with editor_col:
        st.markdown(f"## {slide_display_title(selected_slide)}")
        st.caption("Fill out the fields below. The sidebar is only for moving between slides.")

        with st.container(border=True):
            st.markdown("### Edit this slide")
            for field in selected_slide["fields"]:
                render_field(selected_slide["id"], selected_slide_data, field)
            if selected_slide["id"] == "title_goal":
                st.markdown("### Journal article upload")
            
                saved_article = st.session_state.get("saved_article", {}) or {}
            
                if saved_article.get("path"):
                    st.success(
                        f"Article already saved in GitHub: "
                        f"{saved_article.get('filename', saved_article.get('path'))}"
                    )
                    st.caption(saved_article.get("path"))
            
                    try:
                        article_bytes = load_file_bytes_from_github(saved_article["path"])
            
                        st.download_button(
                            "Download saved article PDF",
                            data=article_bytes,
                            file_name=saved_article.get("filename", "journal_article.pdf"),
                            mime="application/pdf",
                            use_container_width=True,
                        )
                    except Exception as exc:
                        st.warning(f"Article is listed, but could not be downloaded: {exc}")
            
                uploaded_article = st.file_uploader(
                    "Upload or replace journal article PDF",
                    type=["pdf"],
                    key="uploaded_article_pdf",
                    help=(
                        "Strongly recommended. If this draft already has an article saved, "
                        "you do not need to re-upload unless replacing it."
                    ),
                )
            
                if uploaded_article is None and not saved_article.get("path"):
                    st.warning(
                        "Please upload the journal article PDF before saving."
                    )
                elif uploaded_article is not None:
                    st.success(f"New article ready to save: {uploaded_article.name}")

        with st.expander("Preview this slide", expanded=False):
            render_slide_preview(selected_slide, selected_slide_data)

    with export_col:
        st.markdown("### Export")
        render_validation_panel(st.session_state.deck)
        st.divider()
        render_downloads(st.session_state.deck)

    st.caption(f"Version {PROJECT_VERSION}")


if __name__ == "__main__":
    main()
