"""Word summary export logic for Journal Club Builder."""

from __future__ import annotations

from io import BytesIO
from typing import Any, Dict, List

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import qrcode

from feedback_config import REDCAP_DISPLAY_URL, REDCAP_QR_URL


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _first_line(value: Any) -> str:
    lines = [line.strip() for line in _safe_text(value).splitlines() if line.strip()]
    return lines[0] if lines else ""


def _lines(value: Any, limit: int | None = None) -> List[str]:
    lines = [line.strip() for line in _safe_text(value).splitlines() if line.strip()]
    return lines[:limit] if limit else lines


def _add_heading(doc: Document, text: str, size: int = 11) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(30, 80, 130)


def _add_small_text(doc: Document, text: str, bold_label: str | None = None) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.0
    if bold_label:
        r = p.add_run(f"{bold_label}: ")
        r.bold = True
        r.font.size = Pt(9)
    r = p.add_run(_safe_text(text))
    r.font.size = Pt(9)


def _add_bullets(doc: Document, items: List[str], limit: int = 4) -> None:
    for item in items[:limit]:
        p = doc.add_paragraph(style=None)
        p.paragraph_format.left_indent = Inches(0.18)
        p.paragraph_format.first_line_indent = Inches(-0.12)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(f"• {item}")
        r.font.size = Pt(8.5)


def _add_compact_table(doc: Document, rows: List[tuple[str, str]]) -> None:
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.autofit = True
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
        for idx, cell in enumerate(cells):
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    run.font.size = Pt(8.5)
                    if idx == 0:
                        run.bold = True


def _make_qr_image(url: str) -> BytesIO:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    bio = BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return bio


def _add_feedback_block(doc: Document) -> None:
    """Add fixed feedback link and QR code to the one-page summary."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run("Feedback: ")
    run.bold = True
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(30, 80, 130)
    link = p.add_run(REDCAP_DISPLAY_URL)
    link.font.size = Pt(8.5)
    link.font.color.rgb = RGBColor(30, 80, 130)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    run = p2.add_run()
    run.add_picture(_make_qr_image(REDCAP_QR_URL), width=Inches(0.72))


def build_word_summary(deck: Dict[str, Dict[str, Any]]) -> BytesIO:
    """Build a compact one-page DOCX summary for session handout/records."""
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.45)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9)

    title_data = deck.get("title_goal", {})
    pico = deck.get("pico", {})
    design = deck.get("study_design", {})
    result = deck.get("main_result", {})
    bottom = deck.get("clinical_bottom_line", {})
    final = deck.get("final_bottom_line", {})

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(_safe_text(title_data.get("session_title", "Journal Club")))
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(30, 80, 130)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(_safe_text(title_data.get("article_title", "")))
    r.bold = True
    r.font.size = Pt(10)

    _add_heading(doc, "Session purpose")
    _add_small_text(doc, _safe_text(title_data.get("teaching_goal", "")))

    _add_heading(doc, "Article in one view")
    _add_compact_table(
        doc,
        [
            ("Patient/problem", _safe_text(pico.get("patient", ""))),
            ("Study question", _safe_text(pico.get("plain_question", ""))),
            ("Study design", _safe_text(design.get("design", ""))),
            ("Primary outcome", _safe_text(pico.get("outcome", ""))),
        ],
    )

    _add_heading(doc, "Main result")
    _add_small_text(doc, _safe_text(result.get("main_result", "")), "Headline")
    _add_small_text(doc, _safe_text(result.get("plain_result", "")), "Plain language")

    _add_heading(doc, "Clinical bottom line")
    _add_small_text(doc, _safe_text(bottom.get("bottom_line", "")))
    _add_small_text(doc, _safe_text(bottom.get("practice_statement", "")), "Practice implication")

    _add_heading(doc, "Why trust it / why be cautious")
    _add_small_text(doc, "Trust factors", None)
    _add_bullets(doc, _lines(bottom.get("trust_bullets", ""), limit=3), limit=3)
    _add_small_text(doc, "Cautions", None)
    _add_bullets(doc, _lines(bottom.get("caution_bullets", ""), limit=3), limit=3)

    _add_heading(doc, "Discussion questions")
    questions = [
        _safe_text(deck.get("patient_problem", {}).get("discussion_question", "")),
        _safe_text(pico.get("discussion_question", "")),
        _safe_text(result.get("discussion_question", "")),
        _safe_text(deck.get("apply_back", {}).get("return_question", "")),
    ]
    _add_bullets(doc, [q for q in questions if q], limit=4)

    _add_heading(doc, "Resident take-home")
    _add_small_text(doc, _safe_text(final.get("resident_take_home", "")))

    _add_feedback_block(doc)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output
