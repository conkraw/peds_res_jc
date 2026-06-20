"""PowerPoint export logic for Journal Club Builder."""

from __future__ import annotations

from io import BytesIO
from typing import Any, Dict, Iterable, List

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

from slide_schema import SLIDES

# Simple, conservative styling. Change these values if you want a Penn State-like theme.
COLOR_DARK = RGBColor(35, 35, 35)
COLOR_MID = RGBColor(95, 95, 95)
COLOR_LIGHT_GRAY = RGBColor(242, 242, 242)
COLOR_HEADER = RGBColor(70, 70, 70)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_ACCENT = RGBColor(30, 80, 130)
COLOR_ACCENT_LIGHT = RGBColor(226, 236, 246)
COLOR_WARNING_LIGHT = RGBColor(250, 238, 218)

SLIDE_W = 13.333
SLIDE_H = 7.5


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _lines(value: Any) -> List[str]:
    return [line.strip() for line in _safe_text(value).splitlines() if line.strip()]


def add_textbox(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text: Any,
    font_size: int = 20,
    bold: bool = False,
    color: RGBColor = COLOR_DARK,
    align=PP_ALIGN.LEFT,
    fill: RGBColor | None = None,
    margin: float = 0.08,
):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = MSO_ANCHOR.TOP

    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = _safe_text(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color

    if fill is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
        shape.line.color.rgb = fill

    return shape


def add_title(slide, title: str, subtitle: str | None = None):
    add_textbox(slide, 0.55, 0.22, 12.2, 0.5, title, font_size=30, bold=True, color=COLOR_DARK)
    line = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.6), Inches(0.82), Inches(12.1), Inches(0.03)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_ACCENT
    line.line.color.rgb = COLOR_ACCENT
    if subtitle:
        add_textbox(slide, 0.65, 0.9, 12.0, 0.4, subtitle, font_size=16, color=COLOR_MID)


def add_bullets(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    items: Iterable[str],
    font_size: int = 18,
    bullet: bool = True,
):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)

    usable_items = [str(item).strip() for item in items if str(item).strip()]
    if not usable_items:
        usable_items = [""]

    for i, item in enumerate(usable_items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.size = Pt(font_size)
        p.font.color.rgb = COLOR_DARK
        p.level = 0
        # python-pptx doesn't expose bullet toggling consistently across versions;
        # prefixing is predictable and remains editable.
        if bullet and not item.startswith(("•", "-", "A.", "B.", "C.", "D.", "1.", "2.", "3.", "4.", "5.")):
            p.text = f"• {item}"

    return shape


def add_section_label(slide, x: float, y: float, w: float, label: str, fill: RGBColor = COLOR_ACCENT_LIGHT):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.38))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = fill
    tf = shape.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = COLOR_DARK
    return shape


def add_footer(slide, text: str = "Journal Club Builder"):
    add_textbox(slide, 0.6, 7.08, 12.0, 0.24, text, font_size=8, color=COLOR_MID, align=PP_ALIGN.RIGHT)


def add_results_table(slide, rows_data: List[Dict[str, Any]], x=0.55, y=1.35, w=12.25, h=3.55):
    columns = ["Outcome", "88% threshold", "92% threshold", "Difference", "Interpretation"]
    cleaned_rows = []
    for row in rows_data or []:
        if any(_safe_text(row.get(col, "")).strip() for col in columns):
            cleaned_rows.append({col: _safe_text(row.get(col, "")) for col in columns})
    cleaned_rows = cleaned_rows[:5]
    if not cleaned_rows:
        cleaned_rows = [{col: "" for col in columns}]

    shape = slide.shapes.add_table(
        len(cleaned_rows) + 1,
        len(columns),
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )
    table = shape.table
    widths = [2.65, 1.55, 1.55, 1.8, 4.7]
    for idx, width in enumerate(widths):
        table.columns[idx].width = Inches(width)

    for c, col in enumerate(columns):
        cell = table.cell(0, c)
        cell.text = col
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_HEADER
        cell.margin_left = Inches(0.04)
        cell.margin_right = Inches(0.04)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.color.rgb = COLOR_WHITE

    for r, row in enumerate(cleaned_rows, start=1):
        for c, col in enumerate(columns):
            cell = table.cell(r, c)
            cell.text = _safe_text(row.get(col, ""))
            cell.margin_left = Inches(0.04)
            cell.margin_right = Inches(0.04)
            if r % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = COLOR_LIGHT_GRAY
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT if c in (0, 4) else PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(10.5)
                    run.font.color.rgb = COLOR_DARK

    return shape


def add_big_number_card(slide, big_number: str, caption: str):
    card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(2.1), Inches(1.55), Inches(9.2), Inches(2.55))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_ACCENT_LIGHT
    card.line.color.rgb = COLOR_ACCENT

    add_textbox(slide, 2.35, 1.82, 8.7, 0.8, big_number, font_size=44, bold=True, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, 2.55, 2.82, 8.3, 0.8, caption, font_size=20, bold=False, align=PP_ALIGN.CENTER)


def add_simple_bar_chart(slide, title: str, label1: str, value1: float, label2: str, value2: float, units: str):
    add_textbox(slide, 0.8, 1.25, 11.8, 0.35, title, font_size=18, bold=True, align=PP_ALIGN.CENTER)
    max_value = max(float(value1 or 0), float(value2 or 0), 1.0)
    chart_x = 2.3
    chart_y = 1.9
    chart_w = 8.8
    bar_h = 0.52
    gap = 0.75

    def draw_bar(row_idx: int, label: str, value: float):
        y = chart_y + row_idx * gap
        add_textbox(slide, 0.85, y - 0.02, 1.35, 0.35, label, font_size=12, bold=True, align=PP_ALIGN.RIGHT)
        bg = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(chart_x), Inches(y), Inches(chart_w), Inches(bar_h))
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_LIGHT_GRAY
        bg.line.color.rgb = COLOR_LIGHT_GRAY
        bar_w = chart_w * (float(value or 0) / max_value)
        bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(chart_x), Inches(y), Inches(bar_w), Inches(bar_h))
        bar.fill.solid()
        bar.fill.fore_color.rgb = COLOR_ACCENT
        bar.line.color.rgb = COLOR_ACCENT
        add_textbox(slide, chart_x + chart_w + 0.15, y + 0.02, 1.2, 0.35, f"{value:g} {units}", font_size=13, bold=True)

    draw_bar(0, label1, value1)
    draw_bar(1, label2, value2)


def build_title_goal_slide(prs, deck):
    data = deck["title_goal"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_textbox(slide, 0.75, 0.6, 11.8, 0.55, data.get("session_title"), font_size=34, bold=True, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, 1.0, 1.28, 11.3, 0.85, data.get("article_title"), font_size=24, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, 1.25, 2.45, 10.9, 0.95, data.get("teaching_goal"), font_size=20, align=PP_ALIGN.CENTER, fill=COLOR_ACCENT_LIGHT)
    add_section_label(slide, 2.65, 4.0, 8.0, "Five questions residents should answer")
    add_bullets(slide, 2.6, 4.55, 8.5, 1.55, _lines(data.get("five_questions")), font_size=17, bullet=False)
    add_footer(slide)
    return slide


def build_opening_case_slide(prs, deck):
    data = deck["opening_case"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "Opening patient case")
    add_textbox(slide, 0.75, 1.18, 11.9, 1.5, data.get("case_stem"), font_size=20, fill=COLOR_LIGHT_GRAY)
    add_textbox(slide, 0.75, 2.95, 11.9, 0.45, data.get("question"), font_size=22, bold=True, color=COLOR_ACCENT)
    add_bullets(slide, 1.0, 3.55, 6.4, 1.6, _lines(data.get("answer_choices")), font_size=19, bullet=False)
    add_textbox(slide, 0.85, 5.65, 11.6, 0.75, data.get("facilitator_prompt"), font_size=18, bold=True, fill=COLOR_WARNING_LIGHT, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_patient_problem_slide(prs, deck):
    data = deck["patient_problem"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "The patient problem")
    add_textbox(slide, 0.75, 1.15, 11.9, 0.75, data.get("headline"), font_size=25, bold=True, color=COLOR_ACCENT)
    add_section_label(slide, 0.8, 2.18, 3.2, "Clinical problem")
    add_bullets(slide, 0.95, 2.72, 11.5, 2.05, _lines(data.get("problem_bullets")), font_size=20)
    add_textbox(slide, 0.85, 5.45, 11.6, 0.75, data.get("discussion_question"), font_size=21, bold=True, fill=COLOR_ACCENT_LIGHT, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_pico_slide(prs, deck):
    data = deck["pico"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "The study question", "PICO")
    labels = [("Patient/problem", "patient"), ("Intervention", "intervention"), ("Comparison", "comparison"), ("Outcome", "outcome")]
    y = 1.35
    for label, key in labels:
        add_section_label(slide, 0.75, y, 2.2, label)
        add_textbox(slide, 3.1, y - 0.02, 9.55, 0.55, data.get(key), font_size=15)
        y += 0.85
    add_textbox(slide, 0.85, 4.95, 11.6, 0.75, data.get("plain_question"), font_size=20, bold=True, fill=COLOR_LIGHT_GRAY, align=PP_ALIGN.CENTER)
    add_textbox(slide, 0.85, 6.05, 11.6, 0.5, data.get("discussion_question"), font_size=19, bold=True, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_study_design_slide(prs, deck):
    data = deck["study_design"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "What they did")
    add_textbox(slide, 0.75, 1.12, 11.9, 0.55, data.get("design"), font_size=24, bold=True, color=COLOR_ACCENT)
    add_section_label(slide, 0.75, 1.95, 3.0, "What that means")
    add_bullets(slide, 0.9, 2.42, 5.7, 2.2, _lines(data.get("design_bullets")), font_size=14)
    add_section_label(slide, 6.95, 1.95, 2.7, "Who was included")
    add_bullets(slide, 7.05, 2.42, 5.2, 1.25, _lines(data.get("included")), font_size=14)
    add_section_label(slide, 6.95, 3.9, 2.7, "Important exclusions")
    add_bullets(slide, 7.05, 4.35, 5.2, 1.55, _lines(data.get("excluded")), font_size=13)
    add_textbox(slide, 0.85, 6.25, 11.6, 0.5, data.get("discussion_question"), font_size=17, bold=True, fill=COLOR_ACCENT_LIGHT, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_main_result_slide(prs, deck):
    data = deck["main_result"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "What they found")
    add_textbox(slide, 0.75, 1.05, 11.9, 0.45, data.get("main_result"), font_size=24, bold=True, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)
    visual_type = data.get("visual_type", "Results table")

    if visual_type == "Results table":
        add_results_table(slide, data.get("results_table", []), y=1.62)
    elif visual_type == "Big-number card":
        add_big_number_card(slide, data.get("big_number", ""), data.get("big_number_caption", ""))
        add_bullets(slide, 1.2, 4.55, 11.0, 0.8, _lines(data.get("key_results"))[:3], font_size=16)
    elif visual_type == "Simple bar chart":
        add_simple_bar_chart(
            slide,
            data.get("chart_title", ""),
            data.get("chart_group_1_label", "Group 1"),
            float(data.get("chart_group_1_value") or 0),
            data.get("chart_group_2_label", "Group 2"),
            float(data.get("chart_group_2_value") or 0),
            data.get("chart_units", ""),
        )
        add_bullets(slide, 1.2, 3.9, 11.0, 1.1, _lines(data.get("key_results"))[:3], font_size=16)
    else:
        add_bullets(slide, 1.0, 1.75, 11.4, 2.6, _lines(data.get("key_results")), font_size=20)

    add_textbox(slide, 0.85, 5.55, 11.6, 0.55, data.get("plain_result"), font_size=17, bold=True, fill=COLOR_LIGHT_GRAY, align=PP_ALIGN.CENTER)
    add_textbox(slide, 0.85, 6.35, 11.6, 0.42, data.get("discussion_question"), font_size=15, bold=True, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_clinical_bottom_line_slide(prs, deck):
    data = deck["clinical_bottom_line"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "What should we do?")
    add_textbox(slide, 0.75, 1.1, 11.9, 0.85, data.get("bottom_line"), font_size=18, bold=True, fill=COLOR_ACCENT_LIGHT, align=PP_ALIGN.CENTER)
    add_section_label(slide, 0.9, 2.25, 2.6, "I trust it because")
    add_bullets(slide, 0.95, 2.75, 5.55, 1.55, _lines(data.get("trust_bullets")), font_size=15)
    add_section_label(slide, 6.85, 2.25, 2.7, "I am cautious because")
    add_bullets(slide, 6.9, 2.75, 5.55, 1.9, _lines(data.get("caution_bullets")), font_size=14)
    add_textbox(slide, 0.85, 5.25, 11.6, 0.65, data.get("practice_statement"), font_size=16, bold=True, fill=COLOR_LIGHT_GRAY, align=PP_ALIGN.CENTER)
    add_textbox(slide, 0.85, 6.15, 11.6, 0.55, data.get("family_explanation"), font_size=13, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_paper_framework_slide(prs, deck):
    data = deck["paper_framework"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "PAPER framework discussion")
    boxes = [
        ("P", "Patient problem", data.get("patient_problem_answer")),
        ("A", "Article type", data.get("article_type_answer")),
        ("P", "Primary question/outcome", data.get("primary_question_answer")),
        ("E", "Evidence quality", data.get("evidence_quality_answer")),
        ("R", "Real-world use", data.get("real_world_answer")),
    ]
    coords = [(0.75, 1.2), (5.0, 1.2), (9.25, 1.2), (2.85, 4.05), (7.1, 4.05)]
    for (letter, title, body), (x, y) in zip(boxes, coords):
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(3.35), Inches(2.25))
        shape.fill.solid()
        shape.fill.fore_color.rgb = COLOR_LIGHT_GRAY
        shape.line.color.rgb = COLOR_ACCENT
        add_textbox(slide, x + 0.1, y + 0.08, 0.45, 0.4, letter, font_size=22, bold=True, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)
        add_textbox(slide, x + 0.55, y + 0.12, 2.55, 0.35, title, font_size=12, bold=True)
        add_textbox(slide, x + 0.18, y + 0.58, 3.0, 1.45, body, font_size=11)
    add_footer(slide)
    return slide


def build_month_skill_slide(prs, deck):
    data = deck["month_skill"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, data.get("skill_title", "Monthly Focus Skill"))
    add_section_label(slide, 0.9, 1.25, 3.6, "Find only five things")
    add_bullets(slide, 0.95, 1.75, 5.3, 2.3, _lines(data.get("reading_questions")), font_size=18, bullet=False)
    add_section_label(slide, 6.8, 1.25, 3.5, "Use this paper as the example")
    add_bullets(slide, 6.85, 1.75, 5.6, 2.3, _lines(data.get("this_paper_summary")), font_size=14, bullet=False)
    add_textbox(slide, 0.85, 5.4, 11.6, 0.9, data.get("teaching_pearl"), font_size=18, bold=True, fill=COLOR_ACCENT_LIGHT, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_apply_back_slide(prs, deck):
    data = deck["apply_back"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "Apply back to the patient")
    add_textbox(slide, 0.85, 1.35, 11.6, 0.8, data.get("return_question"), font_size=24, bold=True, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)
    add_section_label(slide, 3.0, 2.7, 7.3, "Closing vote")
    add_bullets(slide, 3.15, 3.22, 7.3, 1.4, _lines(data.get("vote_options")), font_size=20, bullet=False)
    add_textbox(slide, 0.85, 5.55, 11.6, 0.8, data.get("facilitator_synthesis"), font_size=17, bold=True, fill=COLOR_WARNING_LIGHT, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_final_bottom_line_slide(prs, deck):
    data = deck["final_bottom_line"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "Final bottom line")
    add_textbox(slide, 1.0, 1.45, 11.3, 1.55, data.get("final_summary"), font_size=22, bold=True, fill=COLOR_ACCENT_LIGHT, align=PP_ALIGN.CENTER)
    add_section_label(slide, 2.2, 4.25, 8.9, "Resident take-home sentence")
    add_textbox(slide, 1.4, 4.9, 10.6, 0.95, data.get("resident_take_home"), font_size=24, bold=True, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def build_facilitator_notes_slide(prs, deck):
    """Create an appendix-style slide for facilitator notes.

    python-pptx does not reliably expose full speaker-notes editing across versions,
    so notes are exported as an editable final slide instead of hidden PowerPoint notes.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "Facilitator notes appendix")
    notes = []
    patient_note = deck.get("patient_problem", {}).get("speaker_note", "")
    if patient_note:
        notes.append(f"Slide 1 patient problem: {patient_note}")
    opening_prompt = deck.get("opening_case", {}).get("facilitator_prompt", "")
    if opening_prompt:
        notes.append(f"Opening case: {opening_prompt}")
    family = deck.get("clinical_bottom_line", {}).get("family_explanation", "")
    if family:
        notes.append(f"Family-facing explanation: {family}")
    add_bullets(slide, 0.85, 1.25, 11.6, 5.25, notes, font_size=15)
    add_footer(slide)
    return slide


def build_powerpoint(deck: Dict[str, Dict[str, Any]], include_facilitator_notes: bool = True) -> BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    build_title_goal_slide(prs, deck)
    build_opening_case_slide(prs, deck)
    build_patient_problem_slide(prs, deck)
    build_pico_slide(prs, deck)
    build_study_design_slide(prs, deck)
    build_main_result_slide(prs, deck)
    build_clinical_bottom_line_slide(prs, deck)
    build_paper_framework_slide(prs, deck)
    build_month_skill_slide(prs, deck)
    build_apply_back_slide(prs, deck)
    build_final_bottom_line_slide(prs, deck)

    if include_facilitator_notes:
        build_facilitator_notes_slide(prs, deck)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


def slide_labels() -> List[str]:
    return [slide["label"] for slide in SLIDES]
