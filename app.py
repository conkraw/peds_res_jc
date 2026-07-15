from __future__ import annotations

import json
import re
import uuid
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from github_storage import (
    GitHubDraftLoadError,
    GitHubDraftSaveError,
    github_backup_is_configured,
    github_config_status_message,
    list_drafts_from_github,
    load_draft_from_github,
    load_file_bytes_from_github,
    save_draft_to_github,
    save_article_to_github,
    generate_archive_id,
    delete_draft_and_article_from_github,

)
from slide_schema import SLIDES, make_default_deck

APP_TITLE = "Journal Club PowerPoint Builder"
PROJECT_VERSION = "0.2.9"


CUSTOM_SLIDES_KEY = "_custom_slides"
MAX_CUSTOM_SLIDES = 2
DEFAULT_CUSTOM_SLIDE_INSERT_AFTER = "final_bottom_line"

CUSTOM_SLIDE_FIELDS: List[Dict[str, Any]] = [
    {
        "key": "insert_after",
        "label": "Place optional slide after",
        "type": "placement",
        "required": True,
        "guide": "Choose where this optional slide should appear in the exported PowerPoint.",
    },
    {
        "key": "title",
        "label": "Optional slide title",
        "type": "text",
        "required": True,
        "max_words": 12,
        "guide": "Short title for the extra slide. Keep this focused.",
    },
    {
        "key": "body",
        "label": "Optional slide main text",
        "type": "textarea",
        "required": True,
        "max_words": 90,
        "height": 190,
        "guide": "Use only when the core slide set is not enough. Write concise teaching text or one short bullet per line.",
    },
    {
        "key": "reason",
        "label": "Why is this extra slide needed?",
        "type": "textarea",
        "max_words": 35,
        "height": 90,
        "guide": "Mentor-facing note. This is included in the review DOCX but not displayed on the PowerPoint slide.",
    },
]


STATS_DEFINITIONS_KEY = "_stats_definitions"
STATS_DEFINITIONS_SLIDE_ID = "_statistics_definitions"
STATS_DEFINITION_COLUMNS = ["Term", "Definition", "How to use clinically"]

DEFAULT_STATS_DEFINITIONS: List[Dict[str, str]] = [
    {
      "Term": "p value",
      "Definition": "Answers the question: 'If there really were no difference between the groups, how likely is it that we would see results like this just by chance?'",
      "How to use clinically": "Smaller p values mean the results are less likely to be explained by chance alone. They do not tell you how big or clinically important the effect is."
    },
    {
      "Term": "Confidence interval",
      "Definition": "A range showing how big the true effect could realistically be.",
      "How to use clinically": "Look at the entire range, not just the average result. Narrow intervals mean more precise estimates; wide intervals mean more uncertainty. Ask whether the possible benefit or harm across the range would change patient care."
    },
    {
        "Term": "RR / risk ratio",
        "Definition": "Compares the probability of an outcome in one group with the probability in another group.",
        "How to use clinically": "RR = 1 means no difference. Above 1 means higher risk; below 1 means lower risk. Always pair it with baseline risk.",
    },
    {
      "Term": "Odds ratio (OR)",
      "Definition": "Shows how much more or less likely an outcome is in one group compared with another using odds.",
      "How to use clinically": "1 = no difference. Greater than 1 = higher odds of the outcome. Less than 1 = lower odds of the outcome. Remember that ORs can exaggerate the size of an effect when the outcome is common."
    },
    {
        "Term": "Absolute risk difference",
        "Definition": "The actual percentage-point difference in outcome rates between groups.",
        "How to use clinically": "This is often the easiest number to use at the bedside because it shows how many more or fewer patients are affected.",
    },
]

STATISTICS_DEFINITIONS_SCHEMA: Dict[str, Any] = {
    "id": STATS_DEFINITIONS_SLIDE_ID,
    "label": "Statistics Definitions",
    "export_title": "Statistics Definitions",
    "is_stats_definitions": True,
    "fields": [
        {
            "key": STATS_DEFINITIONS_KEY,
            "label": "Statistics definitions table",
            "type": "table",
            "required": False,
            "max_rows": 12,
            "columns": STATS_DEFINITION_COLUMNS,
            "multiline_columns": ["Definition", "How to use clinically"],
            "cell_height": 82,
            "guide": (
                "These resident-friendly definitions appear only in the 1-page summary DOCX, "
                "on a separate statistics definitions page. They do not appear in the PowerPoint."
            ),
        }
    ],
}



# -----------------------------
# Performance / export helpers
# -----------------------------

EXPORT_CACHE_KEYS = [
    "prepared_pptx_bytes",
    "prepared_pptx_signature",
    "prepared_pptx_timestamp",
    "prepared_summary_docx_bytes",
    "prepared_summary_docx_signature",
    "prepared_summary_docx_timestamp",
    "prepared_review_docx_bytes",
    "prepared_review_docx_signature",
    "prepared_review_docx_timestamp",
]

ARTICLE_DOWNLOAD_CACHE_PREFIX = "prepared_article_pdf__"


def deck_export_signature(deck: Dict[str, Any], exclude_keys: List[str] | None = None) -> str:
    """Return a stable signature for the current editable deck.

    Export files are expensive to build. This signature lets the app reuse a
    previously prepared PowerPoint/DOCX until the relevant deck content changes.
    Optional exclude_keys are useful for content that appears only in certain
    exports, such as statistics definitions that appear only in the summary DOCX.
    """
    snapshot = deepcopy(deck) if isinstance(deck, dict) else {}
    for key in exclude_keys or []:
        if isinstance(snapshot, dict):
            snapshot.pop(key, None)
    return json.dumps(snapshot, sort_keys=True, ensure_ascii=False, default=str)


def clear_export_cache() -> None:
    """Remove prepared export bytes when loading/resetting drafts."""
    for key in EXPORT_CACHE_KEYS:
        st.session_state.pop(key, None)
    for key in list(st.session_state.keys()):
        if str(key).startswith(ARTICLE_DOWNLOAD_CACHE_PREFIX):
            st.session_state.pop(key, None)


@st.cache_data(show_spinner=False)
def load_static_printable_form_bytes(path: str = "journal_club_printable_planning_form.docx") -> bytes | None:
    """Load the static planning form once instead of reading it on every rerun."""
    form_path = Path(path)
    if form_path.exists():
        return form_path.read_bytes()
    return None


def get_prepared_export_bytes(kind: str, signature: str) -> bytes | None:
    """Return cached export bytes only if they match the current deck."""
    data_key = f"prepared_{kind}_bytes"
    sig_key = f"prepared_{kind}_signature"
    if st.session_state.get(sig_key) == signature:
        data = st.session_state.get(data_key)
        if isinstance(data, bytes) and data:
            return data
    return None


def mark_export_prepared(kind: str, signature: str, data: bytes) -> None:
    """Store prepared export bytes in session state."""
    st.session_state[f"prepared_{kind}_bytes"] = data
    st.session_state[f"prepared_{kind}_signature"] = signature
    st.session_state[f"prepared_{kind}_timestamp"] = datetime.now().strftime("%H:%M:%S")


def export_needs_refresh(kind: str, signature: str) -> bool:
    """Return True when an old prepared export exists but no longer matches."""
    sig_key = f"prepared_{kind}_signature"
    old_signature = st.session_state.get(sig_key)
    return bool(old_signature and old_signature != signature)


def build_powerpoint_export(deck: Dict[str, Any]) -> bytes:
    """Lazy import PowerPoint builder only when the user prepares an export."""
    from pptx_builder import build_powerpoint

    return build_powerpoint(deck, include_facilitator_notes=True)


def build_summary_export(deck: Dict[str, Any]) -> bytes:
    """Lazy import Word summary builder only when the user prepares an export."""
    from docx_builder import build_word_summary

    return build_word_summary(deck)


def build_review_export(deck: Dict[str, Any]) -> bytes:
    """Lazy import mentor review builder only when the user prepares it."""
    from docx_builder import build_review_text_docx

    return build_review_text_docx(deck)


def build_printable_form_fallback(deck: Dict[str, Any]) -> bytes | None:
    """Fallback for older deployments that still use printable_form_builder.py."""
    try:
        from printable_form_builder import build_printable_planning_form
    except Exception:
        return None
    return build_printable_planning_form(deck)


# -----------------------------
# Autosave helpers
# -----------------------------


def autosave_is_ready() -> bool:
    """Autosave can update an existing Archive draft after the first manual save/reload."""
    return bool(
        st.session_state.get("autosave_enabled", True)
        and github_backup_is_configured()
        and str(st.session_state.get("archive_path", "") or "").strip()
        and str(st.session_state.get("archive_id", "") or "").strip()
    )


def autosave_current_draft(deck: Dict[str, Any], reason: str = "autosave") -> bool:
    """Update the current GitHub Archive draft if autosave is active.

    Autosave intentionally updates only the JSON draft. If the resident uploads
    or replaces the article PDF, they should still use Save Draft To Archive so
    the article file itself is uploaded and linked correctly.
    """
    if not autosave_is_ready():
        return False

    signature = deck_export_signature(deck)
    if signature == str(st.session_state.get("autosave_last_signature", "") or ""):
        return False

    archive_path = str(st.session_state.get("archive_path", "") or "").strip().lstrip("/")
    archive_id = str(st.session_state.get("archive_id", "") or "").strip()
    saved_article = st.session_state.get("saved_article", {}) or {}

    presenter_name = str(st.session_state.get("archive_presenter_name", "") or "").strip()
    if not presenter_name:
        presenter_name = str(st.session_state.get("github_presenter_name", "") or "").strip()
    if not presenter_name:
        presenter_name = "Autosaved Draft"

    session_title = str(st.session_state.get("archive_session_title", "") or "").strip()
    if not session_title:
        session_title = str(st.session_state.get("github_session_title", "") or "").strip()
    if not session_title:
        session_title = default_session_title(deck)

    try:
        result = save_draft_to_github(
            deck=deck,
            presenter_name=presenter_name,
            session_title=session_title,
            app_version=PROJECT_VERSION,
            article=saved_article,
            archive_id=archive_id,
            existing_path=archive_path,
        )
        st.session_state.archive_path = result.path
        st.session_state.archive_presenter_name = presenter_name
        st.session_state.archive_session_title = session_title
        st.session_state.autosave_last_signature = signature
        st.session_state.autosave_last_at = datetime.now().strftime("%I:%M:%S %p")
        st.session_state.autosave_last_reason = reason
        st.session_state.autosave_last_error = ""
        return True
    except Exception as exc:
        st.session_state.autosave_last_error = str(exc)
        return False


def render_autosave_status(deck: Dict[str, Any]) -> None:
    """Show autosave controls/status in the sidebar."""
    with st.container(border=True):
        st.markdown("**Autosave**")
        st.checkbox(
            "Autosave when changing slides",
            key="autosave_enabled",
            help=(
                "After the first manual Archive save, the app updates the same "
                "draft JSON when you move to another slide. Article PDF uploads "
                "still require Save Draft To Archive."
            ),
        )

        if autosave_is_ready():
            st.success("Active after first Archive save")
            last_at = str(st.session_state.get("autosave_last_at", "") or "").strip()
            last_reason = str(st.session_state.get("autosave_last_reason", "") or "").strip()
            if last_at:
                st.caption(f"Last autosave: {last_at}" + (f" ({last_reason})" if last_reason else ""))
            else:
                st.caption("No autosave yet this session.")

            if st.button("Autosave now", key="autosave_now_button", use_container_width=True):
                saved = autosave_current_draft(deck, reason="manual")
                if saved:
                    st.success("Autosaved.")
                else:
                    st.info("No changes to autosave.")
        else:
            if github_backup_is_configured():
                st.info("Autosave starts after the first Save Draft To Archive or after you reload an archived draft.")
            else:
                st.info("Autosave starts after GitHub Archive is configured.")

        error = str(st.session_state.get("autosave_last_error", "") or "").strip()
        if error:
            st.warning(f"Last autosave issue: {error}")


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
    default[CUSTOM_SLIDES_KEY] = []
    default[STATS_DEFINITIONS_KEY] = make_default_statistics_definitions()

    for slide in SLIDES:
        sid = slide["id"]
        if isinstance(deck, dict) and sid in deck and isinstance(deck[sid], dict):
            for field in slide["fields"]:
                fkey = field["key"]
                if fkey in deck[sid]:
                    default[sid][fkey] = deck[sid][fkey]

    if isinstance(deck, dict):
        raw_custom_slides = deck.get(CUSTOM_SLIDES_KEY, [])
        if isinstance(raw_custom_slides, list):
            default[CUSTOM_SLIDES_KEY] = [
                normalize_custom_slide(item)
                for item in raw_custom_slides
                if isinstance(item, dict)
            ][:MAX_CUSTOM_SLIDES]

        if STATS_DEFINITIONS_KEY in deck:
            default[STATS_DEFINITIONS_KEY] = normalize_statistics_definitions(
                deck.get(STATS_DEFINITIONS_KEY, [])
            )

    return default

def delete_passcode_is_valid(passcode: str) -> bool:
    """Return True when the entered delete passcode matches Streamlit secrets.

    Supports either:
    [admin]
    delete_passcode = "..."

    or the older location:
    [github]
    delete_passcode = "..."
    """
    provided = str(passcode or "").strip()
    if not provided:
        return False

    expected_values: List[str] = []
    try:
        expected_values.append(str(st.secrets.get("admin", {}).get("delete_passcode", "")).strip())
        expected_values.append(str(st.secrets.get("github", {}).get("delete_passcode", "")).strip())
    except Exception:
        expected_values = []

    return any(expected and provided == expected for expected in expected_values)

def initialize_state() -> None:
    if "deck" not in st.session_state:
        st.session_state.deck = make_blank_deck()
    if "selected_slide_id" not in st.session_state:
        st.session_state.selected_slide_id = SLIDES[0]["id"]
    if "include_facilitator_notes" not in st.session_state:
        st.session_state.include_facilitator_notes = True
    # Metadata for an article PDF already saved with a GitHub draft.
    # Streamlit cannot pre-fill st.file_uploader(), so this is how the app
    # remembers and surfaces the saved article after a draft is reloaded.
    if "saved_article" not in st.session_state:
        st.session_state.saved_article = {}
    if "archive_panel" not in st.session_state:
        st.session_state.archive_panel = ""
    if "archive_id" not in st.session_state:
        st.session_state.archive_id = ""
    if "archive_path" not in st.session_state:
        st.session_state.archive_path = ""
    if "advanced_panel" not in st.session_state:
        st.session_state.advanced_panel = ""
    if "archive_index_rows" not in st.session_state:
        st.session_state.archive_index_rows = []
    if "archive_index_warnings" not in st.session_state:
        st.session_state.archive_index_warnings = []
    if "archive_index_loaded" not in st.session_state:
        st.session_state.archive_index_loaded = False
    if "export_cache_initialized" not in st.session_state:
        st.session_state.export_cache_initialized = True
    if "_slide_radio_version" not in st.session_state:
        st.session_state._slide_radio_version = 0
    if "autosave_enabled" not in st.session_state:
        st.session_state.autosave_enabled = True
    if "autosave_last_signature" not in st.session_state:
        st.session_state.autosave_last_signature = ""
    if "autosave_last_at" not in st.session_state:
        st.session_state.autosave_last_at = ""
    if "autosave_last_reason" not in st.session_state:
        st.session_state.autosave_last_reason = ""
    if "autosave_last_error" not in st.session_state:
        st.session_state.autosave_last_error = ""
    if "archive_presenter_name" not in st.session_state:
        st.session_state.archive_presenter_name = ""
    if "archive_session_title" not in st.session_state:
        st.session_state.archive_session_title = ""



def make_default_statistics_definitions() -> List[Dict[str, str]]:
    """Return editable default rows for the statistics definitions table."""
    return deepcopy(DEFAULT_STATS_DEFINITIONS)


def normalize_statistics_definitions(rows: Any) -> List[Dict[str, str]]:
    """Normalize saved statistics-definition rows into the current three columns."""
    if not isinstance(rows, list):
        return make_default_statistics_definitions()

    normalized: List[Dict[str, str]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        normalized.append(
            {
                "Term": str(row.get("Term") or row.get("term") or "").strip(),
                "Definition": str(row.get("Definition") or row.get("definition") or "").strip(),
                "How to use clinically": str(
                    row.get("How to use clinically")
                    or row.get("Clinical interpretation")
                    or row.get("clinical_interpretation")
                    or row.get("how_to_use")
                    or ""
                ).strip(),
            }
        )

    return normalized


def get_statistics_definitions(deck: Dict[str, Any]) -> List[Dict[str, str]]:
    """Return the editable statistics definitions list, creating defaults if missing."""
    rows = deck.get(STATS_DEFINITIONS_KEY)
    if not isinstance(rows, list):
        deck[STATS_DEFINITIONS_KEY] = make_default_statistics_definitions()
    else:
        deck[STATS_DEFINITIONS_KEY] = normalize_statistics_definitions(rows)
    return deck[STATS_DEFINITIONS_KEY]


def get_custom_slides(deck: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Return the optional slide list, creating a safe empty list if needed."""
    custom_slides = deck.get(CUSTOM_SLIDES_KEY)
    if not isinstance(custom_slides, list):
        deck[CUSTOM_SLIDES_KEY] = []
        custom_slides = deck[CUSTOM_SLIDES_KEY]
    return custom_slides



def core_slide_ids() -> List[str]:
    """Return valid core slide IDs for optional-slide placement."""
    return [slide["id"] for slide in SLIDES]


def normalize_custom_slide_position(value: Any) -> str:
    """Return a safe core slide ID for optional-slide placement."""
    candidate = str(value or "").strip()
    if candidate in core_slide_ids():
        return candidate
    return DEFAULT_CUSTOM_SLIDE_INSERT_AFTER


def custom_slide_insert_after(custom_slide: Dict[str, Any]) -> str:
    """Return the core slide ID after which this optional slide should appear."""
    return normalize_custom_slide_position(custom_slide.get("insert_after"))


def default_insert_after_for_new_custom_slide(deck: Dict[str, Any]) -> str:
    """Default a new optional slide to appear after the currently selected slide when possible."""
    selected_id = str(st.session_state.get("selected_slide_id", "") or "")
    if selected_id in core_slide_ids():
        return selected_id
    for custom_slide in get_custom_slides(deck):
        if str(custom_slide.get("id", "")) == selected_id:
            return custom_slide_insert_after(custom_slide)
    return DEFAULT_CUSTOM_SLIDE_INSERT_AFTER


def ordered_custom_slides(deck: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Return optional slides in the same order they will export in PowerPoint."""
    custom_slides = get_custom_slides(deck)
    ordered: List[Dict[str, Any]] = []
    for core_slide in SLIDES:
        core_id = core_slide["id"]
        for custom_slide in custom_slides:
            if custom_slide_insert_after(custom_slide) == core_id:
                ordered.append(custom_slide)
    return ordered

def make_new_custom_slide(insert_after: str | None = None) -> Dict[str, Any]:
    """Create one blank optional slide record."""
    return {
        "id": f"custom_{uuid.uuid4().hex[:8]}",
        "insert_after": normalize_custom_slide_position(insert_after),
        "title": "",
        "body": "",
        "reason": "",
    }


def normalize_custom_slide(raw: Dict[str, Any]) -> Dict[str, Any]:
    """Normalize older or partial optional slide records."""
    return {
        "id": str(raw.get("id") or f"custom_{uuid.uuid4().hex[:8]}").strip(),
        "insert_after": normalize_custom_slide_position(raw.get("insert_after") or raw.get("after_slide_id")),
        "title": str(raw.get("title") or "").strip(),
        "body": str(raw.get("body") or "").strip(),
        "reason": str(raw.get("reason") or "").strip(),
    }


def make_custom_slide_schema(custom_slide: Dict[str, Any], index: int) -> Dict[str, Any]:
    """Create a schema-like object so optional slides use the same field renderer."""
    title = str(custom_slide.get("title") or "").strip()
    return {
        "id": custom_slide.get("id", f"custom_{index}"),
        "label": f"Optional Slide {index}",
        "export_title": title or f"Optional Slide {index}",
        "fields": CUSTOM_SLIDE_FIELDS,
        "is_custom": True,
        "custom_index": index,
    }


def get_navigation_slides(deck: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Return core slides plus optional slides in their selected export position."""
    navigation_slides: List[Dict[str, Any]] = []
    custom_slides = get_custom_slides(deck)

    for core_slide in SLIDES:
        navigation_slides.append(core_slide)
        for stored_index, custom_slide in enumerate(custom_slides, start=1):
            if custom_slide_insert_after(custom_slide) == core_slide["id"]:
                # Keep the optional-slide number stable by creation order, even
                # when the slide is moved earlier/later in the deck.
                navigation_slides.append(make_custom_slide_schema(custom_slide, stored_index))

    navigation_slides.append(STATISTICS_DEFINITIONS_SCHEMA)
    return navigation_slides


def get_selected_slide_and_data(deck: Dict[str, Any]) -> tuple[Dict[str, Any], Dict[str, Any]]:
    """Return the selected schema and data dict for either a core or optional slide."""
    selected_id = st.session_state.get("selected_slide_id", SLIDES[0]["id"])

    for slide in SLIDES:
        if slide["id"] == selected_id:
            return slide, deck[slide["id"]]

    for idx, custom_slide in enumerate(get_custom_slides(deck), start=1):
        if custom_slide.get("id") == selected_id:
            return make_custom_slide_schema(custom_slide, idx), custom_slide

    if selected_id == STATS_DEFINITIONS_SLIDE_ID:
        get_statistics_definitions(deck)
        return STATISTICS_DEFINITIONS_SCHEMA, deck

    st.session_state.selected_slide_id = SLIDES[0]["id"]
    return SLIDES[0], deck[SLIDES[0]["id"]]


def remove_custom_slide(deck: Dict[str, Any], slide_id: str) -> None:
    """Remove one optional slide from the active deck."""
    deck[CUSTOM_SLIDES_KEY] = [
        slide for slide in get_custom_slides(deck)
        if slide.get("id") != slide_id
    ]


def request_slide_selection(slide_id: str, slide_label: str = "") -> None:
    """Select a slide on the next rerun without touching a widget-backed key.

    Streamlit does not allow assigning to a radio widget's session_state key
    after that widget has been created. To avoid that error, we store only the
    internal selected_slide_id and bump a radio-version counter. On the next
    rerun, the radio gets a fresh key and uses selected_slide_id to choose its
    default index.
    """
    st.session_state.selected_slide_id = slide_id
    st.session_state._slide_radio_version = int(st.session_state.get("_slide_radio_version", 0) or 0) + 1

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
        STATS_DEFINITIONS_SLIDE_ID: "Statistics Definitions",
    }
    return labels.get(slide["id"], slide["label"])


def slide_display_title(slide: Dict[str, Any]) -> str:
    """Actual slide title shown in the main workspace and preview.

    This intentionally does not use nav_label(), because nav_label() is only
    for the sidebar. Keeping these separate prevents duplicate headings like
    'Opening Case: Opening patient case'.
    """
    return str(slide.get("export_title") or slide.get("label") or "Untitled slide").strip()

def core_slide_position_label(slide_id: str) -> str:
    """Human-readable label for optional-slide placement."""
    safe_id = normalize_custom_slide_position(slide_id)
    for slide in SLIDES:
        if slide["id"] == safe_id:
            return f"After {nav_label(slide)}"
    return "After Final Bottom Line"


def core_slide_position_options() -> List[str]:
    """Core slide IDs used by the optional-slide placement selectbox."""
    return [slide["id"] for slide in SLIDES]


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
    """Update the selected slide from the sidebar radio selection."""
    label_to_id = {
        nav_label(slide): slide["id"]
        for slide in get_navigation_slides(st.session_state.deck)
    }
    radio_key = st.session_state.get("_slide_radio_key", "")
    selected_label = st.session_state.get(radio_key)
    if selected_label in label_to_id:
        new_slide_id = label_to_id[selected_label]
        old_slide_id = str(st.session_state.get("selected_slide_id", "") or "")
        if new_slide_id != old_slide_id:
            autosave_current_draft(st.session_state.deck, reason="slide change")
        st.session_state.selected_slide_id = new_slide_id


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

    for idx, custom_slide in enumerate(get_custom_slides(deck), start=1):
        slide = make_custom_slide_schema(custom_slide, idx)
        for field in slide["fields"]:
            value = custom_slide.get(field["key"], "")
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

    for idx, custom_slide in enumerate(get_custom_slides(deck), start=1):
        slide = make_custom_slide_schema(custom_slide, idx)
        for field in slide["fields"]:
            visible_fields += 1
            value = custom_slide.get(field["key"], "")
            if str(value or "").strip():
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


def render_placement_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    """Render the optional-slide placement selector with friendly labels."""
    key = field["key"]
    options = core_slide_position_options()
    current = normalize_custom_slide_position(slide_data.get(key, DEFAULT_CUSTOM_SLIDE_INSERT_AFTER))
    widget_key = f"widget__{slide_id}__{key}"

    value = st.selectbox(
        field["label"],
        options=options,
        index=options.index(current) if current in options else options.index(DEFAULT_CUSTOM_SLIDE_INSERT_AFTER),
        key=widget_key,
        help=field.get("guide"),
        format_func=core_slide_position_label,
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
    multiline_columns = set(field.get("multiline_columns", []))
    cell_height = int(field.get("cell_height", 80) or 80)

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

            if column_name in multiline_columns:
                updated_row[column_name] = row_cols[col_index].text_area(
                    f"{column_name} row {row_index + 1}",
                    key=cell_key,
                    height=cell_height,
                    label_visibility="collapsed",
                )
            else:
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
    elif ftype == "placement":
        value = render_placement_field(slide_id, slide_data, field)
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
        elif field["type"] == "placement":
            st.caption(f"{field['label']}: {core_slide_position_label(value)}")
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
        with st.expander(f"Validation Warnings ({len(problems)})", expanded=False):
            for problem in problems:
                st.write(f"- {problem}")
    else:
        st.success("All visible fields are within limits.")


def default_session_title(deck: Dict[str, Dict[str, Any]]) -> str:
    title_slide = deck.get("title_goal", {})
    return str(title_slide.get("session_title") or title_slide.get("article_title") or "Journal Club").strip()


def friendly_draft_label(filename: str) -> str:
    """Turn an archive filename into a readable dropdown option.

    Newer filenames end with a unique archive ID, such as:
    2026-06-25_jane-smith_ouch-trial_a1b2c3d4e5f6.json
    """
    stem = str(filename or "").removesuffix(".json")
    parts = stem.split("_", 2)
    if len(parts) == 3:
        saved_date, presenter_slug, title_slug = parts
        archive_id = ""
        title_part = title_slug

        maybe_title, sep, maybe_id = title_slug.rpartition("_")
        if sep and re.fullmatch(r"[a-f0-9]{8,24}", maybe_id):
            title_part = maybe_title
            archive_id = maybe_id

        presenter = presenter_slug.replace("-", " ").title()
        title = title_part.replace("-", " ").title()
        if archive_id:
            return f"{saved_date} — {title} ({presenter}) · ID {archive_id}"
        return f"{saved_date} — {title} ({presenter})"
    return filename


def apply_loaded_payload_to_session(loaded: Dict[str, Any], source_path: str = "") -> None:
    """Load a draft payload into the active app session and refresh widgets.

    Do not directly assign Streamlit widget keys here. This function can be
    called after widgets are already rendered in the same run, and Streamlit
    disallows changing widget-backed session_state keys after instantiation.

    Important: st.file_uploader() cannot be pre-filled from GitHub. Instead,
    the saved article metadata is kept separately in session_state so the app
    can show a Download saved article button after reloading a draft.
    """
    st.session_state.deck = normalize_deck(loaded)

    article_metadata: Dict[str, Any] = {}
    archive_id = ""
    archive_path = str(source_path or "").strip().lstrip("/")
    if isinstance(loaded, dict):
        archive_id = str(loaded.get("archive_id", "") or "").strip()
        archive_path = archive_path or str(loaded.get("archive_path", "") or "").strip().lstrip("/")
        if isinstance(loaded.get("article"), dict):
            article_metadata = loaded.get("article", {}) or {}
            archive_id = archive_id or str(article_metadata.get("archive_id", "") or "").strip()

    st.session_state.saved_article = article_metadata
    st.session_state.archive_id = archive_id
    st.session_state.archive_path = archive_path

    if isinstance(loaded, dict):
        st.session_state.archive_presenter_name = str(loaded.get("presenter_name", "") or "").strip()
        st.session_state.archive_session_title = str(loaded.get("session_title", "") or "").strip()

    st.session_state.autosave_last_signature = deck_export_signature(st.session_state.deck)
    st.session_state.autosave_last_at = ""
    st.session_state.autosave_last_reason = ""
    st.session_state.autosave_last_error = ""

    clear_export_cache()
    clear_widget_state()


def render_github_backup(deck: Dict[str, Dict[str, Any]]) -> None:
    """Render the save-to-archive form after the user clicks the archive button."""
    with st.container(border=True):
        st.markdown("#### Save Draft To Archive")
        st.caption("Save the editable draft JSON and the article PDF, if one is attached.")

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

        existing_archive_id = str(st.session_state.get("archive_id", "") or "").strip()
        existing_archive_path = str(st.session_state.get("archive_path", "") or "").strip().lstrip("/")
        if existing_archive_path:
            st.caption(f"This will update the existing archive file: {existing_archive_path}")
        elif existing_archive_id:
            st.caption(f"Archive ID for this draft: {existing_archive_id}")
        else:
            st.caption("A unique Archive ID will be created when this draft is saved.")

        if github_backup_is_configured():
            st.success(github_config_status_message())
        else:
            st.info(github_config_status_message())
            st.caption("Add Streamlit secrets first. See README.md for setup instructions.")

        article_file = st.session_state.get("uploaded_article_pdf")
        saved_article = st.session_state.get("saved_article", {}) or {}

        if article_file is None and not saved_article.get("path"):
            st.warning("No journal article PDF uploaded yet.")
        elif article_file is not None:
            st.success(f"Article PDF ready to save: {article_file.name}")
        elif saved_article.get("path"):
            st.success(
                f"Existing article will be kept: "
                f"{saved_article.get('filename', saved_article.get('path'))}"
            )

        button_cols = st.columns([1, 1])
        with button_cols[0]:
            save_clicked = st.button(
                "Save draft to Archive",
                key="save_draft_to_archive_button",
                use_container_width=True,
            )
        with button_cols[1]:
            if st.button("Close Save Archive Panel", key="cancel_save_archive_button", use_container_width=True):
                st.session_state.archive_panel = ""
                st.rerun()

        if save_clicked:
            if not presenter_name.strip():
                st.error("Please enter the presenter name before saving.")
                return
            if not session_title.strip():
                st.error("Please enter the session title before saving.")
                return

            try:
                archive_id = (
                    str(st.session_state.get("archive_id", "") or "").strip()
                    or str(saved_article.get("archive_id", "") or "").strip()
                    or generate_archive_id()
                )
                archive_path = str(st.session_state.get("archive_path", "") or "").strip().lstrip("/")
                st.session_state.archive_id = archive_id

                article_metadata = dict(saved_article or {})
                if article_metadata:
                    article_metadata["archive_id"] = archive_id

                # 1. Save article first if a new one was uploaded.
                if article_file is not None:
                    article_result = save_article_to_github(
                        article_bytes=article_file.getvalue(),
                        original_filename=article_file.name,
                        presenter_name=presenter_name,
                        session_title=session_title,
                        archive_id=archive_id,
                        existing_path=str(saved_article.get("path", "") or "").strip().lstrip("/"),
                    )

                    article_metadata = {
                        "archive_id": archive_id,
                        "filename": article_file.name,
                        "path": article_result.path,
                        "html_url": article_result.html_url,
                        "commit_sha": article_result.commit_sha,
                    }

                    st.session_state.saved_article = article_metadata

                # 2. Save JSON draft after article metadata exists.
                result = save_draft_to_github(
                    deck=deck,
                    presenter_name=presenter_name,
                    session_title=session_title,
                    app_version=PROJECT_VERSION,
                    article=article_metadata,
                    archive_id=archive_id,
                    existing_path=archive_path,
                )
                st.session_state.archive_path = result.path
                st.session_state.archive_presenter_name = presenter_name.strip()
                st.session_state.archive_session_title = session_title.strip()
                st.session_state.autosave_last_signature = deck_export_signature(deck)
                st.session_state.autosave_last_at = datetime.now().strftime("%I:%M:%S %p")
                st.session_state.autosave_last_reason = "manual save"
                st.session_state.autosave_last_error = ""

                messages = ["Draft saved to Archive."]

                if article_file is not None:
                    messages.append("Article saved to Archive.")
                elif article_metadata.get("path"):
                    messages.append("Existing article kept in Archive.")
                else:
                    messages.append("No article PDF uploaded.")

                st.success("\n".join(messages))

                if result.html_url:
                    st.caption("You can reload it later from the Archive.")

            except GitHubDraftSaveError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Unexpected GitHub save error: {exc}")


def render_github_recovery() -> None:
    """Render the reload/delete archive form after the user clicks the archive button."""
    with st.container(border=True):
        st.markdown("#### Reload Saved Draft From Archive")

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

        search_cols = st.columns([1, 1])
        with search_cols[0]:
            find_clicked = st.button(
                "Find saved drafts",
                key="find_saved_drafts_button",
                use_container_width=True,
            )
        with search_cols[1]:
            if st.button("Close Reload Archive Panel", key="cancel_reload_archive_button", use_container_width=True):
                st.session_state.archive_panel = ""
                st.rerun()

        if find_clicked:
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
        if not drafts:
            return

        label_to_path = {friendly_draft_label(d["name"]): d["path"] for d in drafts}
        option_labels = list(label_to_path.keys())

        if (
            st.session_state.get("selected_github_draft_label") not in label_to_path
            and option_labels
        ):
            st.session_state.selected_github_draft_label = option_labels[0]

        selected_label = st.selectbox(
            "Choose a saved draft",
            option_labels,
            key="selected_github_draft_label",
        )

        selected_path = label_to_path[selected_label]

        if st.button("Load selected draft", key="load_selected_archive_draft_button", use_container_width=True):
            try:
                loaded = load_draft_from_github(selected_path)
                apply_loaded_payload_to_session(loaded, source_path=selected_path)
                st.success(f"Loaded draft: {selected_label}")
                st.rerun()
            except GitHubDraftLoadError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Unexpected GitHub load error: {exc}")

        # The delete controls remain behind a separate safety expander.
        with st.expander("Danger zone: delete selected draft", expanded=False):
            st.warning(
                "This will delete the selected draft from Archive. "
                "If the draft has a saved article PDF, you can delete that too."
            )

            delete_article_too = st.checkbox(
                "Also delete the associated article PDF, if one exists",
                value=True,
                key="delete_article_too",
            )

            delete_passcode = st.text_input(
                "Delete passcode",
                type="password",
                key="delete_passcode",
                help="Only mentors/admins with the delete passcode can delete Archive files.",
            )

            confirm_delete = st.checkbox(
                "I understand this will delete files from the Archive",
                key="confirm_delete_github_draft",
            )

            st.caption(f"Selected for deletion: {selected_label}")

            if st.button("Delete selected draft from Archive", key="delete_selected_archive_draft_button", use_container_width=True):
                if not confirm_delete:
                    st.error("Please check the confirmation box before deleting.")
                elif not delete_passcode_is_valid(delete_passcode):
                    st.error("Incorrect or missing delete passcode.")
                else:
                    try:
                        delete_result = delete_draft_and_article_from_github(
                            draft_path=selected_path,
                            delete_article=delete_article_too,
                        )

                        deleted = delete_result.get("deleted", [])
                        warnings = delete_result.get("warnings", [])

                        if deleted:
                            st.success("Deleted:\n" + "\n".join(deleted))

                        for warning in warnings:
                            st.warning(warning)

                        st.session_state.github_draft_results = [
                            draft for draft in st.session_state.get("github_draft_results", [])
                            if draft.get("path") != selected_path
                        ]

                        for key in [
                            "delete_article_too",
                            "delete_passcode",
                            "confirm_delete_github_draft",
                            "selected_github_draft_label",
                        ]:
                            if key in st.session_state:
                                del st.session_state[key]

                        st.session_state.archive_index_loaded = False
                        st.session_state.archive_index_rows = []
                        st.session_state.archive_index_warnings = []

                        st.rerun()

                    except GitHubDraftSaveError as exc:
                        st.error(str(exc))
                    except GitHubDraftLoadError as exc:
                        st.error(str(exc))
                    except Exception as exc:
                        st.error(f"Unexpected GitHub delete error: {exc}")



def extract_archive_index_row(draft_info: Dict[str, Any], payload: Dict[str, Any]) -> Dict[str, Any]:
    """Pull mentor-facing index fields out of one saved draft payload."""
    payload = payload if isinstance(payload, dict) else {}
    deck = payload.get("deck", payload) if isinstance(payload, dict) else {}
    deck = deck if isinstance(deck, dict) else {}

    title_slide = deck.get("title_goal", {}) if isinstance(deck.get("title_goal", {}), dict) else {}
    month_skill_slide = deck.get("month_skill", {}) if isinstance(deck.get("month_skill", {}), dict) else {}
    article = payload.get("article", {}) if isinstance(payload.get("article", {}), dict) else {}

    article_status = "No"
    if article.get("filename"):
        article_status = str(article.get("filename"))
    elif article.get("path"):
        article_status = "Yes"

    return {
        "Saved date": payload.get("saved_date", ""),
        "Presenter": payload.get("presenter_name", ""),
        #"Session Title": payload.get("session_title") or title_slide.get("session_title", ""),
        "Article Title": title_slide.get("article_title", ""),
        "Monthly skill": month_skill_slide.get("skill_title", ""),
        #"Article PDF": article_status,
        #"Archive ID": payload.get("archive_id") or article.get("archive_id", ""),
        #"Draft path": draft_info.get("path", ""),
    }


def load_archive_index_rows() -> tuple[List[Dict[str, Any]], List[str]]:
    """Load every saved draft JSON and build an archive index table."""
    drafts = list_drafts_from_github("")
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []

    for draft in drafts:
        draft_path = str(draft.get("path", "") or "")
        if not draft_path:
            continue
        try:
            payload = load_draft_from_github(draft_path)
            rows.append(extract_archive_index_row(draft, payload))
        except Exception as exc:
            warnings.append(f"Could not read {draft_path}: {exc}")

    return rows, warnings


def render_archive_index() -> None:
    """Render a mentor-facing table of saved journal club archive contents."""
    with st.container(border=True):
        st.markdown("#### Archive Index")
        st.caption("Lists saved drafts with presenter, article title, monthly skill, and article PDF status.")

        if github_backup_is_configured():
            st.success(github_config_status_message())
        else:
            st.info(github_config_status_message())
            st.caption("Add Streamlit secrets first. See README.md for setup instructions.")

        action_cols = st.columns([1, 1])
        with action_cols[0]:
            refresh_clicked = st.button(
                "Refresh Archive Index",
                key="refresh_archive_index_button",
                use_container_width=True,
            )
        with action_cols[1]:
            if st.button("Close Archive Index Panel", key="cancel_archive_index_button", use_container_width=True):
                st.session_state.archive_panel = ""
                st.rerun()

        if refresh_clicked or not st.session_state.get("archive_index_loaded", False):
            try:
                with st.spinner("Building archive index..."):
                    rows, warnings = load_archive_index_rows()
                st.session_state.archive_index_rows = rows
                st.session_state.archive_index_warnings = warnings
                st.session_state.archive_index_loaded = True
                if rows:
                    st.success(f"Found {len(rows)} archived draft(s).")
                else:
                    st.info("No archived drafts found.")
            except GitHubDraftLoadError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Unexpected archive index error: {exc}")

        rows = st.session_state.get("archive_index_rows", []) or []
        warnings = st.session_state.get("archive_index_warnings", []) or []

        if rows:
            df = pd.DataFrame(rows)
            preferred_columns = [
                "Saved date",
                "Presenter",
                "Article Title",
                "Monthly skill",
                #"Session title",
                #"Article PDF",
                #"Archive ID",
                #"Draft path",
            ]
            df = df[[col for col in preferred_columns if col in df.columns]]
            st.dataframe(df, hide_index=True, use_container_width=True)

            csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Download Archive Index CSV",
                data=csv_bytes,
                file_name="journal_club_archive_index.csv",
                mime="text/csv",
                use_container_width=True,
            )

        for warning in warnings:
            st.warning(warning)

def render_archive_controls(deck: Dict[str, Dict[str, Any]]) -> None:
    """Use persistent buttons instead of expanders for archive actions."""
    if "archive_panel" not in st.session_state:
        st.session_state.archive_panel = ""
    if "archive_id" not in st.session_state:
        st.session_state.archive_id = ""
    if "archive_path" not in st.session_state:
        st.session_state.archive_path = ""

    if st.button(
        "Save Draft To Archive",
        key="open_save_archive_panel_button",
        use_container_width=True,
    ):
        st.session_state.archive_panel = "save"

    if st.button(
        "Reload Saved Draft From Archive",
        key="open_reload_archive_panel_button",
        use_container_width=True,
    ):
        st.session_state.archive_panel = "reload"

    if st.button(
        "View Archive Index",
        key="open_archive_index_panel_button",
        use_container_width=True,
    ):
        st.session_state.archive_panel = "index"

    panel = st.session_state.get("archive_panel", "")
    if panel == "save":
        render_github_backup(deck)
    elif panel == "reload":
        render_github_recovery()
    elif panel == "index":
        render_archive_index()
        
def make_blank_deck() -> Dict[str, Dict[str, Any]]:
    deck = make_default_deck()
    for slide in SLIDES:
        for field in slide["fields"]:
            field_key = field["key"]
            if field.get("type") == "table":
                deck[slide["id"]][field_key] = []
            elif field.get("type") == "number":
                deck[slide["id"]][field_key] = field.get("default", 0)
            elif field.get("type") == "select":
                deck[slide["id"]][field_key] = field.get("default", field.get("options", [""])[0])
            else:
                deck[slide["id"]][field_key] = ""
    deck[CUSTOM_SLIDES_KEY] = []
    deck[STATS_DEFINITIONS_KEY] = make_default_statistics_definitions()
    return deck
    
def render_downloads(deck: Dict[str, Dict[str, Any]]) -> None:
    problems = validate_deck(deck)
    timestamp = datetime.now().strftime("%Y%m%d")
    pptx_signature = deck_export_signature(deck, exclude_keys=[STATS_DEFINITIONS_KEY])
    summary_signature = deck_export_signature(deck)

    render_archive_controls(deck)
    st.divider()

    st.caption(
        "Exports are prepared only when you click a Prepare button. "
        "This keeps the app fast while you edit."
    )

    export_cols = st.columns(2)
    with export_cols[0]:
        prepare_pptx = st.button(
            "Prepare PowerPoint",
            key="prepare_powerpoint_button",
            disabled=bool(problems),
            use_container_width=True,
        )
    with export_cols[1]:
        prepare_summary = st.button(
            "Prepare 1-Page Summary",
            key="prepare_summary_docx_button",
            disabled=bool(problems),
            use_container_width=True,
        )

    if prepare_pptx:
        try:
            with st.spinner("Preparing PowerPoint..."):
                mark_export_prepared("pptx", pptx_signature, build_powerpoint_export(deck))
            st.success("PowerPoint prepared.")
        except Exception as exc:
            st.error(f"Could not prepare PowerPoint: {exc}")

    if prepare_summary:
        try:
            with st.spinner("Preparing 1-page summary..."):
                mark_export_prepared("summary_docx", summary_signature, build_summary_export(deck))
            st.success("1-page summary prepared.")
        except Exception as exc:
            st.error(f"Could not prepare 1-page summary: {exc}")

    pptx_bytes = get_prepared_export_bytes("pptx", pptx_signature)
    if export_needs_refresh("pptx", pptx_signature):
        st.warning("PowerPoint needs to be prepared again because the deck changed.")

    st.download_button(
        "Download PowerPoint",
        data=pptx_bytes or b"",
        file_name=f"journal_club_deck_{timestamp}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        disabled=bool(problems) or pptx_bytes is None,
        use_container_width=True,
    )

    summary_docx_bytes = get_prepared_export_bytes("summary_docx", summary_signature)
    if export_needs_refresh("summary_docx", summary_signature):
        st.warning("1-page summary needs to be prepared again because the deck changed.")

    st.download_button(
        "Download 1-Page Summary",
        data=summary_docx_bytes or b"",
        file_name=f"journal_club_summary_{timestamp}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        disabled=bool(problems) or summary_docx_bytes is None,
        use_container_width=True,
    )

    printable_form = load_static_printable_form_bytes()
    if printable_form is None:
        # Slow fallback only used when the static form file is missing.
        printable_form = build_printable_form_fallback(deck)

    if printable_form is not None:
        st.download_button(
            "Download Printable Planning Form",
            data=printable_form,
            file_name="journal_club_printable_planning_form.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    else:
        st.warning("Printable planning form file is missing from the app folder.")

    draft_json = json.dumps(deck, indent=2, ensure_ascii=False).encode("utf-8")
    # st.download_button("Download editable draft JSON", data=draft_json, file_name=f"journal_club_draft_{timestamp}.json", mime="application/json", use_container_width=True)




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

        navigation_slides = get_navigation_slides(st.session_state.deck)
        label_to_id = {nav_label(slide): slide["id"] for slide in navigation_slides}
        id_to_label = {slide["id"]: nav_label(slide) for slide in navigation_slides}

        # Keep the displayed radio value and the internal slide id synchronized.
        # Do not assign to the radio widget's session_state key directly.
        # Instead, choose the default index from selected_slide_id.
        option_labels = list(label_to_id.keys())
        current_label = id_to_label.get(st.session_state.get("selected_slide_id"), option_labels[0])
        if current_label not in label_to_id:
            st.session_state.selected_slide_id = SLIDES[0]["id"]
            current_label = id_to_label[SLIDES[0]["id"]]

        radio_version = int(st.session_state.get("_slide_radio_version", 0) or 0)
        radio_key = f"selected_slide_radio_{radio_version}"
        st.session_state._slide_radio_key = radio_key

        st.radio(
            "Choose slide",
            option_labels,
            index=option_labels.index(current_label),
            key=radio_key,
            on_change=sync_selected_slide,
        )

        st.divider()
        st.markdown("**Optional slides**")
        custom_slides = get_custom_slides(st.session_state.deck)
        st.caption(f"{len(custom_slides)} / {MAX_CUSTOM_SLIDES} optional slides added")
        if st.button(
            "Add Optional Slide",
            key="add_optional_slide_button",
            disabled=len(custom_slides) >= MAX_CUSTOM_SLIDES,
            use_container_width=True,
        ):
            new_slide = make_new_custom_slide(default_insert_after_for_new_custom_slide(st.session_state.deck))
            custom_slides.append(new_slide)
            autosave_current_draft(st.session_state.deck, reason="optional slide added")
            request_slide_selection(new_slide["id"], f"Optional Slide {len(custom_slides)}")
            clear_export_cache()
            clear_widget_state()
            st.rerun()

        st.divider()
        render_autosave_status(st.session_state.deck)

        st.divider()

        if st.button(
            "Advanced Drafts/Reset",
            key="open_advanced_panel_button",
            use_container_width=True,
        ):
            st.session_state.advanced_panel = (
                "" if st.session_state.get("advanced_panel") == "advanced" else "advanced"
            )

        if st.session_state.get("advanced_panel") == "advanced":
            with st.container(border=True):
                st.markdown("**Advanced drafts/reset**")

                #uploaded = st.file_uploader("Load a saved draft JSON",type=["json"],key="advanced_uploaded_draft_json",)
                #if uploaded is not None:
                #    if st.button("Load uploaded draft", key="load_uploaded_json_button", use_container_width=True):
                #        try:
                #            loaded = json.loads(uploaded.getvalue().decode("utf-8"))
                #            apply_loaded_payload_to_session(loaded)
                #            st.success("Draft loaded.")
                #        except Exception as exc:
                #            st.error(f"Could not load draft: {exc}")

                st.caption(
                    "Mentor review export creates an editable DOCX with the slide text, reviewer guidelines, and Track Changes enabled."
                )
                review_signature = deck_export_signature(st.session_state.deck)
                if st.button(
                    "Prepare PowerPoint Text Review DOCX",
                    key="prepare_review_docx_button",
                    use_container_width=True,
                ):
                    try:
                        with st.spinner("Preparing mentor review DOCX..."):
                            mark_export_prepared(
                                "review_docx",
                                review_signature,
                                build_review_export(st.session_state.deck),
                            )
                        st.success("Mentor review DOCX prepared.")
                    except Exception as exc:
                        st.error(f"Could not prepare mentor review DOCX: {exc}")

                review_docx_bytes = get_prepared_export_bytes("review_docx", review_signature)
                if export_needs_refresh("review_docx", review_signature):
                    st.warning("Mentor review DOCX needs to be prepared again because the deck changed.")

                st.download_button(
                    "Download PowerPoint Text Review DOCX",
                    data=review_docx_bytes or b"",
                    file_name=f"journal_club_text_review_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    disabled=review_docx_bytes is None,
                    use_container_width=True,
                )

                st.divider()

                if st.button("Reset to OxyKids Example", key="reset_oxykids_button", use_container_width=True):
                    st.session_state.deck = make_default_deck()
                    st.session_state.deck[CUSTOM_SLIDES_KEY] = []
                    st.session_state.deck[STATS_DEFINITIONS_KEY] = make_default_statistics_definitions()
                    st.session_state.saved_article = {}
                    st.session_state.archive_id = ""
                    st.session_state.archive_path = ""
                    st.session_state.archive_presenter_name = ""
                    st.session_state.archive_session_title = ""
                    st.session_state.autosave_last_signature = ""
                    st.session_state.autosave_last_at = ""
                    st.session_state.autosave_last_reason = ""
                    st.session_state.autosave_last_error = ""
                    clear_export_cache()
                    clear_widget_state()
                    st.success("Reset complete.")
                    st.rerun()

                if st.button("Clear All Fields", key="clear_all_fields_button", use_container_width=True):
                    st.session_state.deck = make_blank_deck()
                    st.session_state.saved_article = {}
                    st.session_state.archive_id = ""
                    st.session_state.archive_path = ""
                    st.session_state.archive_presenter_name = ""
                    st.session_state.archive_session_title = ""
                    st.session_state.autosave_last_signature = ""
                    st.session_state.autosave_last_at = ""
                    st.session_state.autosave_last_reason = ""
                    st.session_state.autosave_last_error = ""
                    clear_export_cache()
                
                    for slide in SLIDES:
                        for field in slide["fields"]:
                            field_key = field["key"]
                            if field.get("type") == "table":
                                st.session_state.deck[slide["id"]][field_key] = []
                            else:
                                st.session_state.deck[slide["id"]][field_key] = ""
                
                    clear_widget_state()
                    st.success("All fields cleared.")
                    st.rerun()

                if st.button("Close Advanced Panel", key="close_advanced_panel_button", use_container_width=True):
                    st.session_state.advanced_panel = ""
                    st.rerun()

    selected_slide, selected_slide_data = get_selected_slide_and_data(st.session_state.deck)

    editor_col, export_col = st.columns([2.2, 0.9])

    with editor_col:
        st.markdown(f"## {slide_display_title(selected_slide)}")
        st.caption("Fill out the fields below. The sidebar is only for moving between slides.")

        with st.container(border=True):
            st.markdown("### Edit this slide")

            if selected_slide.get("is_custom"):
                st.info(
                    f"This optional slide will export {core_slide_position_label(selected_slide_data.get('insert_after')).lower()}. "
                    "Use it only when the standard slide set does not cover an important teaching point."
                )
                if st.button(
                    "Remove this optional slide",
                    key=f"remove_optional_slide_{selected_slide['id']}",
                    use_container_width=True,
                ):
                    remove_custom_slide(st.session_state.deck, selected_slide["id"])
                    autosave_current_draft(st.session_state.deck, reason="optional slide removed")
                    request_slide_selection(SLIDES[0]["id"], nav_label(SLIDES[0]))
                    clear_export_cache()
                    clear_widget_state()
                    st.rerun()

            for field in selected_slide["fields"]:
                render_field(selected_slide["id"], selected_slide_data, field)
            if selected_slide["id"] == "title_goal":
                st.markdown("### Journal article upload")
            
                saved_article = st.session_state.get("saved_article", {}) or {}
            
                if saved_article.get("path"):
                    st.success(
                        f"Article already saved in Archive: "
                        f"{saved_article.get('filename', saved_article.get('path'))}"
                    )
                    saved_archive_id = saved_article.get("archive_id") or st.session_state.get("archive_id", "")
                    if saved_archive_id:
                        st.caption(f"Archive ID: {saved_archive_id}")
            
                    article_path = str(saved_article.get("path", "") or "")
                    article_cache_key = ARTICLE_DOWNLOAD_CACHE_PREFIX + article_path
                    if st.button(
                        "Prepare saved article PDF download",
                        key="prepare_saved_article_pdf_button",
                        use_container_width=True,
                    ):
                        try:
                            with st.spinner("Loading saved article PDF..."):
                                st.session_state[article_cache_key] = load_file_bytes_from_github(article_path)
                            st.success("Saved article PDF is ready to download.")
                        except Exception as exc:
                            st.warning(f"Article is listed, but could not be downloaded: {exc}")

                    article_bytes = st.session_state.get(article_cache_key)
                    if isinstance(article_bytes, bytes) and article_bytes:
                        st.download_button(
                            "Download saved article PDF",
                            data=article_bytes,
                            file_name=saved_article.get("filename", "journal_article.pdf"),
                            mime="application/pdf",
                            use_container_width=True,
                        )
            
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
