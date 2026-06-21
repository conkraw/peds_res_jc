from __future__ import annotations

import json
import re
from copy import deepcopy
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from pptx_builder import build_powerpoint
from slide_schema import SLIDES, make_default_deck


APP_TITLE = "Journal Club PowerPoint Builder"
PROJECT_VERSION = "0.1.3"


# -----------------------------
# Utility functions
# -----------------------------


def count_words(text: Any) -> int:
    return len(re.findall(r"\b\w+\b", str(text or "")))


def nonempty_lines(text: Any) -> List[str]:
    return [line.strip() for line in str(text or "").splitlines() if line.strip()]


def clear_widget_state() -> None:
    for key in list(st.session_state.keys()):
        if key.startswith("widget__") or key.startswith("table__"):
            del st.session_state[key]


def normalize_deck(deck: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """Merge uploaded/old draft data into the current schema."""
    default = make_default_deck()
    for slide in SLIDES:
        sid = slide["id"]
        if sid in deck and isinstance(deck[sid], dict):
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
    """Short labels for the sidebar so the main page can carry the instructions."""
    labels = {
        "title_goal": "Title: JC Overview",
        "opening_case": "Opening: The Patient Case",
        "patient_problem": "Slide 1: The Patient Problem",
        "pico": "Slide 2: PICO",
        "study_design": "Slide 3: Study Design",
        "main_result": "Slide 4: Main Results",
        "clinical_bottom_line": "Slide 5: Clinical Bottom Line",
        "paper_framework": "PAPER",
        "month_skill": "Monthly Focus Skill",
        "apply_back": "Apply Back To Patient",
        "final_bottom_line": "Final Bottom Line",
    }
    return labels.get(slide["id"], slide["label"])


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
        value = st.text_area(label, key=widget_key, help=help_text, height=130)
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


def render_table_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    key = field["key"]
    current_rows = slide_data.get(key, deepcopy(field.get("default", [])))
    df = pd.DataFrame(current_rows)

    # Optional: allow a companion text field to control the table columns.
    # For Slide 4, this is results_table_columns.
    columns_key = field.get("columns_key")
    configured_columns = parse_table_columns(slide_data.get(columns_key, "")) if columns_key else []
    if not configured_columns:
        configured_columns = list(df.columns) or list(field.get("columns", []))

    if configured_columns:
        for col in configured_columns:
            if col not in df.columns:
                df[col] = ""
        df = df[configured_columns]

    # Include the column names in the widget key so Streamlit redraws the table
    # immediately when the user edits the column list.
    column_signature = "__".join(configured_columns) if configured_columns else "default"
    widget_key = f"table__{slide_id}__{key}__{column_signature}"

    st.caption(field.get("guide", ""))
    edited = st.data_editor(
        df,
        key=widget_key,
        hide_index=True,
        num_rows="dynamic",
        use_container_width=True,
    )
    rows = edited.fillna("").to_dict(orient="records")
    slide_data[key] = rows
    return rows


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
    # Show one clean preview heading. The slide label is for navigation/sidebar;
    # export_title is what appears as the actual slide title.
    preview_title = slide.get("export_title") or slide["label"]
    st.subheader(preview_title)

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

    draft_json = json.dumps(deck, indent=2, ensure_ascii=False).encode("utf-8")
    st.download_button(
        "Download editable draft JSON",
        data=draft_json,
        file_name=f"journal_club_draft_{timestamp}.json",
        mime="application/json",
        use_container_width=True,
    )


# -----------------------------
# Main app
# -----------------------------


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    initialize_state()

    st.title(APP_TITLE)
    st.caption(
        "Choose a slide on the left, complete the fields in the main workspace, then export a standardized PowerPoint."
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

    selected_slide = next(slide for slide in SLIDES if slide["id"] == st.session_state.selected_slide_id)
    selected_slide_data = st.session_state.deck[selected_slide["id"]]

    editor_col, export_col = st.columns([2.2, 0.9])

    with editor_col:
        st.markdown(f"## {nav_label(selected_slide)}: {selected_slide.get('export_title') or selected_slide['label']}")
        st.caption("Fill out the fields below. The sidebar is only for moving between slides.")

        with st.container(border=True):
            st.markdown("### Edit this slide")
            for field in selected_slide["fields"]:
                render_field(selected_slide["id"], selected_slide_data, field)

        with st.expander("Preview this slide", expanded=True):
            render_slide_preview(selected_slide, selected_slide_data)

    with export_col:
        st.markdown("### Export")
        render_validation_panel(st.session_state.deck)
        st.divider()
        render_downloads(st.session_state.deck)

    st.caption(f"Version {PROJECT_VERSION}")


if __name__ == "__main__":
    main()
