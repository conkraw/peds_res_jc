"""Word summary export logic for Journal Club Builder."""

from __future__ import annotations

from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile
from typing import Any, Dict, List

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import qrcode

from feedback_config import REDCAP_DISPLAY_URL, REDCAP_QR_URL
from slide_schema import SLIDES


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

    _add_heading(doc, "Session Purpose")
    _add_small_text(doc, _safe_text(title_data.get("teaching_goal", "")))

    _add_heading(doc, "Article In One View")
    _add_compact_table(
        doc,
        [
            ("Patient/Problem", _safe_text(pico.get("patient", ""))),
            ("Study Question", _safe_text(pico.get("plain_question", ""))),
            ("Study Design", _safe_text(design.get("design", ""))),
            ("Primary Outcome", _safe_text(pico.get("outcome", ""))),
        ],
    )

    _add_heading(doc, "Main Result")
    _add_small_text(doc, _safe_text(result.get("main_result", "")), "Headline")
    _add_small_text(doc, _safe_text(result.get("plain_result", "")), "Plain language")

    _add_heading(doc, "Clinical Bottom Line")
    _add_small_text(doc, _safe_text(bottom.get("bottom_line", "")))
    _add_small_text(doc, _safe_text(bottom.get("practice_statement", "")), "Practice implication")

    _add_heading(doc, "Why Trust It / Why Be Cautious")
    _add_small_text(doc, "Trust Factors", None)
    _add_bullets(doc, _lines(bottom.get("trust_bullets", ""), limit=3), limit=3)
    _add_small_text(doc, "Cautions", None)
    _add_bullets(doc, _lines(bottom.get("caution_bullets", ""), limit=3), limit=3)

    _add_heading(doc, "Discussion Questions")
    questions = [
        _safe_text(deck.get("patient_problem", {}).get("discussion_question", "")),
        _safe_text(pico.get("discussion_question", "")),
        _safe_text(result.get("discussion_question", "")),
        _safe_text(deck.get("apply_back", {}).get("return_question", "")),
    ]
    _add_bullets(doc, [q for q in questions if q], limit=4)

    _add_heading(doc, "Resident Take-Home")
    _add_small_text(doc, _safe_text(final.get("resident_take_home", "")))

    #_add_feedback_block(doc)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# -----------------------------
# Mentor review text export
# -----------------------------


def _add_review_heading(doc: Document, text: str, level: int = 1) -> None:
    """Add a readable heading for the mentor review document."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8 if level == 1 else 5)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(15 if level == 1 else 12)
    run.font.color.rgb = RGBColor(30, 80, 130)


def _add_review_paragraph(doc: Document, text: str, font_size: float = 10, italic: bool = False) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(_safe_text(text))
    run.font.name = "Arial"
    run.font.size = Pt(font_size)
    run.italic = italic


def _add_review_label(doc: Document, label: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(label)
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(30, 80, 130)


def _add_review_text_block(doc: Document, label: str, text: Any) -> None:
    """Add one editable slide-text field with a clear label."""
    _add_review_label(doc, label)
    content = _safe_text(text)
    if not content:
        content = "[blank]"
    for idx, paragraph_text in enumerate(content.splitlines() or [content]):
        if not paragraph_text.strip():
            continue
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.15)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(paragraph_text.strip())
        run.font.name = "Arial"
        run.font.size = Pt(10)


def _add_review_table_block(doc: Document, label: str, rows: Any) -> None:
    """Add a readable editable representation of a slide table."""
    _add_review_label(doc, label)
    if not isinstance(rows, list) or not rows:
        _add_review_paragraph(doc, "[blank table]", font_size=10, italic=True)
        return

    columns: List[str] = []
    for row in rows:
        if isinstance(row, dict):
            for key in row.keys():
                if str(key) not in columns:
                    columns.append(str(key))
    if not columns:
        _add_review_paragraph(doc, "[blank table]", font_size=10, italic=True)
        return

    table = doc.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"
    header_cells = table.rows[0].cells
    for idx, column in enumerate(columns):
        header_cells[idx].text = column
        for paragraph in header_cells[idx].paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(8.5)

    for row in rows:
        if not isinstance(row, dict):
            continue
        cells = table.add_row().cells
        for idx, column in enumerate(columns):
            cells[idx].text = _safe_text(row.get(column, ""))
            for paragraph in cells[idx].paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_reviewer_guidelines(doc: Document) -> None:
    """Add mentor/reviewer instructions at the top of the review document."""
    _add_review_heading(doc, "Reviewer Guidelines", level=2)
    guidelines = [
        "Use Track Changes and/or comments so the resident can see your suggestions.",
        "Focus on clarity, accuracy, educational value, clinical reasoning, and whether the message is easy to present aloud.",
        "Avoid stylistic preference edits unless they improve readability, reduce confusion, or make the teaching point clearer.",
        "Preserve the resident's voice when possible; suggest targeted edits rather than rewriting the entire slide.",
        "Flag overstatements, missing limitations, unclear applicability, jargon, or places where the clinical takeaway is too broad.",
        "Keep slide text concise. If content is correct but too dense, suggest what to cut or move to facilitator notes.",
    ]
    for item in guidelines:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.22)
        p.paragraph_format.first_line_indent = Inches(-0.14)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(f"• {item}")
        run.font.name = "Arial"
        run.font.size = Pt(9.5)

    _add_review_paragraph(
        doc,
        "Reviewer workflow: edit the text below directly, or add comments beside sections that need discussion. "
        "The goal is not to perfect the slide design; the goal is to help the resident make the content clearer, more accurate, and more useful for learners.",
        font_size=9.5,
        italic=True,
    )


def _enable_track_changes(docx_stream: BytesIO) -> BytesIO:
    """Turn on Word track-revisions setting for the generated review document.

    This does not create tracked changes by itself. It nudges Word to open the
    document with Track Changes enabled so mentor edits are easier for the
    resident to review.
    """
    try:
        source = BytesIO(docx_stream.getvalue())
        target = BytesIO()
        with ZipFile(source, "r") as zin, ZipFile(target, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/settings.xml":
                    xml = data.decode("utf-8")
                    if "w:trackRevisions" not in xml:
                        xml = xml.replace("</w:settings>", "<w:trackRevisions/></w:settings>")
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
        target.seek(0)
        return target
    except Exception:
        docx_stream.seek(0)
        return docx_stream


def build_review_text_docx(deck: Dict[str, Dict[str, Any]]) -> BytesIO:
    """Build an editable DOCX containing the PowerPoint text for mentor review.

    This is intentionally different from the one-page summary. It includes the
    full text from the slide-builder fields, organized by slide title, so a
    mentor can provide tracked edits or comments before the resident finalizes
    the PowerPoint.
    """
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)

    title_data = deck.get("title_goal", {})
    session_title = _safe_text(title_data.get("session_title", "Journal Club")) or "Journal Club"
    article_title = _safe_text(title_data.get("article_title", ""))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run("Journal Club PowerPoint Text Review")
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(17)
    run.font.color.rgb = RGBColor(30, 80, 130)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(session_title)
    r.bold = True
    r.font.name = "Arial"
    r.font.size = Pt(12)

    if article_title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run(article_title)
        r.font.name = "Arial"
        r.font.size = Pt(10)
        r.italic = True

    _add_reviewer_guidelines(doc)

    for slide_number, slide in enumerate(SLIDES, start=1):
        slide_id = slide["id"]
        slide_data = deck.get(slide_id, {}) or {}
        slide_title = str(slide.get("export_title") or slide.get("label") or slide_id).strip()
        _add_review_heading(doc, f"Slide {slide_number}: {slide_title}", level=1)

        for field in slide.get("fields", []):
            # Respect simple show_if conditions so the review file matches the visible app fields.
            show_if = field.get("show_if") or {}
            if show_if:
                visible = all(slide_data.get(key) == expected for key, expected in show_if.items())
                if not visible:
                    continue

            field_label = str(field.get("label") or field.get("key") or "Field")
            field_key = field.get("key")
            value = slide_data.get(field_key, "") if field_key else ""

            if field.get("type") == "table":
                _add_review_table_block(doc, field_label, value)
            else:
                _add_review_text_block(doc, field_label, value)

        _add_review_label(doc, "Mentor notes / comments")
        _add_review_paragraph(doc, "[Add comments here or use Word comments in the margin.]", font_size=9.5, italic=True)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return _enable_track_changes(output)
