"""Slide schema for the Journal Club PowerPoint Builder.

The schema is intentionally plain Python so non-developers can edit slide labels,
field limits, helper text, and defaults without digging through the app logic.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List


DEFAULT_RESULTS_TABLE: List[Dict[str, str]] = [
    {
        "Outcome": "Time to discharge criteria",
        "88% threshold": "27.6 h",
        "92% threshold": "46.6 h",
        "Difference": "~16.8 h shorter",
        "Interpretation": "Faster discharge readiness",
    },
    {
        "Outcome": "Length of stay",
        "88% threshold": "39.8 h",
        "92% threshold": "60.8 h",
        "Difference": "~17.6 h shorter",
        "Interpretation": "Shorter hospitalization",
    },
    {
        "Outcome": "Oxygen use",
        "88% threshold": "Less",
        "92% threshold": "More",
        "Difference": "—",
        "Interpretation": "Less treatment burden",
    },
    {
        "Outcome": "Safety outcomes",
        "88% threshold": "Similar",
        "92% threshold": "Similar",
        "Difference": "—",
        "Interpretation": "No clear short-term safety signal",
    },
]


DEFAULT_RESULTS_TABLE_COLUMNS = "Outcome, 88% threshold, 92% threshold, Difference, Interpretation"


SLIDES: List[Dict[str, Any]] = [
    {
        "id": "title_goal",
        "label": "Title + teaching goal",
        "export_title": "Journal Club",
        "fields": [
            {
                "key": "session_title",
                "label": "Session Title",
                "type": "text",
                "required": True,
                "max_words": 12,
                "default": "First Journal Club: OxyKids Trial",
                "guide": "Short name for the session.",
            },
            {
                "key": "article_title",
                "label": "Article Title / Topic",
                "type": "text",
                "required": True,
                "max_words": 30,
                "default": "Oxygen Saturation Thresholds in Children With Acute Respiratory Distress",
                "guide": "Use the article title or a clean teaching title.",
            },
            {
                "key": "teaching_goal",
                "label": "Teaching Goal",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "height":100,
                "default": "Today is not about mastering every statistical method. Today is about learning how to read a clinical article well enough to answer five core questions.",
                "guide": "Set expectations for the first session.",
            },
            {
                "key": "five_questions",
                "label": "Five Reading Questions",
                "type": "textarea",
                "required": True,
                "max_lines": 5,
                "max_words_per_line": 20,
                "default": "What question is the article asking?\nWhat kind of study is this?\nWhat was the primary outcome?\nWhat was the main result?\nWhat should we do with this information clinically?",
                "guide": "One question per line.",
            },
        ],
    },
    {
        "id": "opening_case",
        "label": "Opening case",
        "export_title": "Opening Patient Case",
        "fields": [
            {
                "key": "case_stem",
                "label": "Patient Case",
                "type": "textarea",
                "required": True,
                "max_words": 95,
                "default": "A 9-month-old, previously healthy child is admitted to the general pediatrics floor with bronchiolitis. The child has mild to moderate work of breathing, is drinking enough to avoid IV fluids, and is otherwise improving. Overnight, the pulse oximeter reads 88–91% while asleep. The nurse asks whether to restart oxygen to keep the saturation above 92%.",
                "guide": "Brief clinical setup. Keep it realistic and readable aloud.",
                "height":100
            },
            {
                "key": "question",
                "label": "Opening Question",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "Before reading the article, what would you do?",
                "guide": "This should be answerable before knowing the article.",
            },
            {
                "key": "answer_choices",
                "label": "Answer Choices",
                "type": "textarea",
                "required": True,
                "max_lines": 5,
                "max_words_per_line": 20,
                "default": "A. Restart oxygen to keep SpO₂ ≥92%\nB. Accept SpO₂ ≥90%\nC. Accept SpO₂ ≥88% if the child otherwise looks well\nD. I am unsure",
                "guide": "One option per line. Usually A–D.",
                "height":130
            },
            {
                "key": "facilitator_prompt",
                "label": "Facilitator Prompt",
                "type": "textarea",
                "required": False,
                "max_words": 50,
                "height":100,
                "default": "Most of us have seen this exact situation. The question is whether the number on the monitor is helping the child or prolonging the admission.",
                "guide": "What should the facilitator say to frame the case?",
            },
        ],
    },
    {
        "id": "patient_problem",
        "label": "Slide 1: Patient Problem",
        "export_title": "The Patient Problem",
        "fields": [
            {
                "key": "headline",
                "label": "Slide Headline",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "Oxygen saturation thresholds may drive care more than the child’s appearance",
                "guide": "One clear sentence that frames the clinical problem.",
            },
            {
                "key": "problem_bullets",
                "label": "Clinical Problem Bullets",
                "type": "textarea",
                "required": True,
                "max_lines": 5,
                "max_words_per_line": 20,
                "default": "Oxygen thresholds are often based on tradition or guidelines with limited evidence.\nA higher threshold may prolong oxygen use and hospital stay.\nA lower threshold could reduce treatment burden.\nThe safety question matters: are we missing children who need oxygen?",
                "guide": "One bullet per line. Avoid full paragraphs.",
            },
            {
                "key": "discussion_question",
                "label": "Discussion Question",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "What SpO₂ number makes you uncomfortable, and why?",
                "guide": "Should invite judgment, not recall.",
            },
            {
                "key": "speaker_note",
                "label": "Speaker Note / Facilitator Note",
                "type": "textarea",
                "required": False,
                "max_words": 100,
                "default": "This article is useful because it tests a common real-world decision: should we tolerate a lower oxygen saturation threshold in otherwise appropriate children on the general pediatric ward?",
                "guide": "Stored for the facilitator guide slide at the end.",
            },
        ],
    },
    {
        "id": "pico",
        "label": "Slide 2: Study question / PICO",
        "export_title": "The Study Question",
        "fields": [
            {
                "key": "patient",
                "label": "Patient/Problem",
                "type": "textarea",
                "required": True,
                "max_words": 50,
                "default": "Children 6 weeks to 12 years old hospitalized on general pediatric wards with acute respiratory distress from bronchiolitis, lower respiratory tract infection, or acute viral-induced wheeze.",
                "height":100
            },
            {
                "key": "intervention",
                "label": "Intervention",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "Use an SpO₂ threshold of 88% to start, continue, or wean oxygen.",
                "height":100
            },
            {
                "key": "comparison",
                "label": "Comparison",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "Use an SpO₂ threshold of 92%.",
                "height":100
            },
            {
                "key": "outcome",
                "label": "Outcome",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "Primary outcome: time from admission to meeting predefined discharge criteria.",
                "height":100
            },
            {
                "key": "plain_question",
                "label": "Plain-Language Study Question",
                "type": "textarea",
                "required": True,
                "max_words": 50,
                "default": "In children admitted with common acute respiratory illnesses, can we safely use an oxygen saturation threshold of 88% instead of 92% and get them ready for discharge sooner?",
                "height":100
            },
            {
                "key": "discussion_question",
                "label": "Discussion Question",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "Is this the right question for our patients?",
                "height":100
            },
        ],
    },
    {
        "id": "study_design",
        "label": "Slide 3: What they did",
        "export_title": "What They Did",
        "fields": [
            {
                "key": "design",
                "label": "Study Design",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "Pragmatic, multicenter, open-label randomized clinical trial.",
            },
            {
                "key": "design_bullets",
                "label": "What That Means",
                "type": "textarea",
                "required": True,
                "max_lines": 6,
                "max_words_per_line": 20,
                "default": "Pragmatic: designed to reflect real clinical care.\nMulticenter: done at 10 hospitals.\nRandomized: assigned to 88% or 92% threshold.\nOpen-label: clinicians knew the assigned threshold.\nGeneral pediatric wards, not PICUs.\nOther care followed local practice.",
            },
            {
                "key": "included",
                "label": "Who Was Included",
                "type": "textarea",
                "required": True,
                "max_lines": 10,
                "max_words_per_line": 20,
                "default": "Children 6 weeks to 12 years.\nBronchiolitis, lower respiratory tract infection, or viral-induced wheeze.\nRequired oxygen by the standard 92% threshold.",
            },
            {
                "key": "excluded",
                "label": "Important Exclusions",
                "type": "textarea",
                "required": True,
                "max_lines": 10,
                "max_words_per_line": 20,
                "default": "Significant pre-existing cardiopulmonary, neurologic, immunologic, or hematologic conditions.\nBorn before 32 weeks’ gestation.\nOlder children with acute asthma.\nFamilies without Dutch or English language ability or stable internet access.",
            },
            {
                "key": "discussion_question",
                "label": "Discussion Question",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "Which patients on our wards look like the study population, and which patients do not?",
            },
        ],
    },
    {
        "id": "main_result",
        "label": "Slide 4: What they found",
        "export_title": "What They Found",
        "fields": [
            {
                "key": "main_result",
                "label": "Main Result Headline",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "The 88% group met discharge criteria sooner.",
            },
            {
                "key": "visual_type",
                "label": "Visual Format",
                "type": "select",
                "required": True,
                "options": ["Results table", "Big-number card", "Simple bar chart", "No visual"],
                "default": "Results table",
                "guide": "The table is the recommended default.",
            },
            {
                "key": "results_table_columns",
                "label": "Results Table Columns",
                "type": "text",
                "required": False,
                "max_words": 50,
                "default": DEFAULT_RESULTS_TABLE_COLUMNS,
                "show_if": {"visual_type": "Results table"},
                "guide": "Edit the column names here. Separate columns with commas.",
            },
            {
                "key": "results_table",
                "label": "Results Table",
                "type": "table",
                "required": False,
                "default": DEFAULT_RESULTS_TABLE,
                "columns_key": "results_table_columns",
                "max_rows": 5,
                "show_if": {"visual_type": "Results table"},
                "guide": "Keep to 4–5 rows. The PowerPoint table stays editable.",
            },
            {
                "key": "big_number",
                "label": "Big-Number Text",
                "type": "text",
                "required": False,
                "max_words": 8,
                "default": "≈17 hours sooner",
                "show_if": {"visual_type": "Big-number card"},
            },
            {
                "key": "big_number_caption",
                "label": "Big-Number Caption",
                "type": "textarea",
                "required": False,
                "max_words": 22,
                "default": "Children in the 88% group met discharge criteria sooner than children in the 92% group.",
                "show_if": {"visual_type": "Big-number card"},
            },
            {
                "key": "chart_title",
                "label": "Chart Title",
                "type": "text",
                "required": False,
                "max_words": 10,
                "default": "Time to discharge criteria",
                "show_if": {"visual_type": "Simple bar chart"},
            },
            {
                "key": "chart_group_1_label",
                "label": "Group 1 Label",
                "type": "text",
                "required": False,
                "max_words": 5,
                "default": "88% threshold",
                "show_if": {"visual_type": "Simple bar chart"},
            },
            {
                "key": "chart_group_1_value",
                "label": "Group 1 Value",
                "type": "number",
                "required": False,
                "default": 27.6,
                "min_value": 0.0,
                "show_if": {"visual_type": "Simple bar chart"},
            },
            {
                "key": "chart_group_2_label",
                "label": "Group 2 Label",
                "type": "text",
                "required": False,
                "max_words": 5,
                "default": "92% threshold",
                "show_if": {"visual_type": "Simple bar chart"},
            },
            {
                "key": "chart_group_2_value",
                "label": "Group 2 Value",
                "type": "number",
                "required": False,
                "default": 46.6,
                "min_value": 0.0,
                "show_if": {"visual_type": "Simple bar chart"},
            },
            {
                "key": "chart_units",
                "label": "Chart Units",
                "type": "text",
                "required": False,
                "max_words": 3,
                "default": "hours",
                "show_if": {"visual_type": "Simple bar chart"},
            },
            {
                "key": "key_results",
                "label": "Key Results Bullets",
                "type": "textarea",
                "required": True,
                "max_lines": 7,
                "max_words_per_line": 20,
                "default": "Time to meeting discharge criteria: 27.6 h with 88% vs 46.6 h with 92%.\nLength of stay: 39.8 h with 88% vs 60.8 h with 92%.\nMore children in the 88% group never received oxygen.\nOxygen duration was shorter in the 88% group.\nSafety outcomes did not significantly differ.",
                "guide": "Used when no visual is selected and in the facilitator guide.",
            },
            {
                "key": "plain_result",
                "label": "Plain-Language Result",
                "type": "textarea",
                "required": True,
                "max_words": 50,
                "default": "Using 88% instead of 92% got children ready for discharge about 17 hours sooner, with less oxygen use and no clear short-term safety signal in this trial.",
                "height":100
            },
            {
                "key": "discussion_question",
                "label": "Discussion Question",
                "type": "text",
                "required": True,
                "max_words": 50,
                "default": "What result matters most to you: discharge readiness, actual length of stay, oxygen duration, or safety outcomes?",
            },
        ],
    },
    {
        "id": "clinical_bottom_line",
        "label": "Slide 5: What Should We Do?",
        "export_title": "What should we do?",
        "fields": [
            {
                "key": "bottom_line",
                "label": "Clinical Bottom Line",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "This article suggests that an SpO₂ threshold of 88% may be reasonable for selected, otherwise healthy children admitted to general pediatric wards with bronchiolitis, lower respiratory tract infection, or viral-induced wheeze.",
            },
            {
                "key": "trust_bullets",
                "label": "Why I Trust It",
                "type": "textarea",
                "required": True,
                "max_lines": 10,
                "max_words_per_line": 20,
                "default": "Randomized and pragmatic.\nClinically relevant outcome.\nMeaningful effect size.\nShort-term safety outcomes were reassuring.",
            },
            {
                "key": "caution_bullets",
                "label": "Why I Am Cautious",
                "type": "textarea",
                "required": True,
                "max_lines": 10,
                "max_words_per_line": 20,
                "default": "Open-label design.\nPrimary outcome included SpO₂ threshold.\nConducted in the Netherlands.\nFew children with darker skin tones.\nExcluded important groups we care for.\nNot designed to prove rare harms or long-term safety.",
            },
            {
                "key": "practice_statement",
                "label": "Practice Statement",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "I would not change practice alone based on this one article. I would use it to start a local discussion about whether selected, otherwise healthy ward patients could safely tolerate a lower oxygen threshold as part of a pathway.",
            },
            {
                "key": "family_explanation",
                "label": "Family-Facing Explanation",
                "type": "textarea",
                "required": False,
                "max_words": 100,
                "default": "Some children who are otherwise improving do not need oxygen just because the monitor briefly reads below 92%. We care about the whole child: breathing effort, feeding, alertness, hydration, and trend over time — not just one number.",
            },
        ],
    },
    {
        "id": "paper_framework",
        "label": "PAPER Framework Discussion",
        "export_title": "PAPER framework discussion",
        "fields": [
            {
                "key": "patient_problem_answer",
                "label": "P — Patient Problem",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "We may be treating a number instead of treating the child. Higher oxygen thresholds may keep children hospitalized longer.",
                "height":100
            },
            {
                "key": "article_type_answer",
                "label": "A — Article Type",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "This is a randomized clinical trial, which is appropriate because the authors are testing an intervention: one oxygen threshold versus another.",
                "height":100
            },
            {
                "key": "primary_question_answer",
                "label": "P — Primary Question / Outcome",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "Can an 88% SpO₂ threshold safely reduce time to discharge readiness compared with 92%? The primary outcome was time from admission to meeting predefined discharge criteria.",
                "height":100
            },
            {
                "key": "evidence_quality_answer",
                "label": "E — Evidence Quality",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "Mostly trustworthy, but not perfect. The trial was randomized, multicenter, pragmatic, and clinically meaningful, but it was open-label and the primary outcome partly depended on the assigned oxygen threshold.",
                "height":100
            },
            {
                "key": "real_world_answer",
                "label": "R — Real-World Use",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "Maybe, but only through a local pathway or guideline discussion, not by individual clinicians freelancing. Nursing, ED, hospital medicine, pulmonology, and family communication would matter.",
                "height":100
            },
        ],
    },
    {
        "id": "month_skill",
        "label": "Monthly Focus Skill",
        "export_title": "Monthly Focus Skill",
        "fields": [
            {
                "key": "skill_title",
                "label": "Skill title",
                "type": "text",
                "required": True,
                "max_words": 20,
                "default": "How to read an article without drowning in details",
            },
            {
                "key": "reading_questions",
                "label": "Five Things To Find",
                "type": "textarea",
                "required": True,
                "max_lines": 5,
                "max_words_per_line": 20,
                "default": "What question is being asked?\nWhat type of study is this?\nWhat was the primary outcome?\nWhat was the main result?\nWhat is the clinical bottom line?",
            },
            {
                "key": "this_paper_summary",
                "label": "Use This Paper As An Example",
                "type": "textarea",
                "required": True,
                #"max_lines": 5,
                #"max_words_per_line": 20,
                "default": "Question: Can we use 88% instead of 92% as the oxygen threshold?\nStudy type: Randomized clinical trial.\nPrimary outcome: Time to meeting discharge criteria.\nMain result: About 17 hours sooner with the 88% threshold.\nClinical bottom line: Selected children may need less oxygen and shorter hospitalization.",
            },
            {
                "key": "teaching_pearl",
                "label": "Teaching Pearl",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "You do not need to understand every statistical detail to learn from an article. Start with the patient problem, the question, the outcome, the main result, the biggest limitation, and the clinical takeaway.",
                "height":100
            },
        ],
    },
    {
        "id": "apply_back",
        "label": "Apply Back To The Patient",
        "export_title": "Apply back to the patient",
        "fields": [
            {
                "key": "return_question",
                "label": "Return-To-Case Question",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "After reviewing the article, would you restart oxygen for asleep SpO₂ 88–91% if the child otherwise looks well?",
                "height":100
            },
            {
                "key": "vote_options",
                "label": "Closing Vote Options",
                "type": "textarea",
                "required": True,
                "max_lines": 5,
                "max_words_per_line": 20,
                "default": "A. Change practice\nB. Do not change practice\nC. Maybe, but only in selected patients\nD. More evidence or local guidance is needed",
            },
            {
                "key": "facilitator_synthesis",
                "label": "Facilitator Synthesis",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "I would be most comfortable with ‘maybe, but only in selected patients.’ This article gives strong support for a local pathway discussion, but we should be careful about applying it to patients who were excluded or underrepresented.",
                "height":100
            },
        ],
    },
    {
        "id": "final_bottom_line",
        "label": "Final Bottom Line",
        "export_title": "Final bottom line",
        "fields": [
            {
                "key": "final_summary",
                "label": "Final Summary",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "This article suggests that an 88% oxygen saturation threshold can reduce oxygen use and shorten hospital stay in selected children admitted with acute respiratory illness. I partially trust it because it was randomized, pragmatic, and multicenter, but I am cautious because it was open-label and important patient groups were excluded or underrepresented.",
            },
            {
                "key": "resident_take_home",
                "label": "Resident Take-Home Sentence",
                "type": "textarea",
                "required": True,
                "max_words": 100,
                "default": "My take-home from this article is that oxygen saturation thresholds can drive hospitalization, and we should think carefully before treating a number instead of the child.",
            },
        ],
    },
]


def make_default_deck() -> Dict[str, Dict[str, Any]]:
    """Return a fresh copy of the default deck data."""
    deck: Dict[str, Dict[str, Any]] = {}
    for slide in SLIDES:
        deck[slide["id"]] = {}
        for field in slide["fields"]:
            deck[slide["id"]][field["key"]] = deepcopy(field.get("default", ""))
    return deck


def get_slide_by_id(slide_id: str) -> Dict[str, Any]:
    for slide in SLIDES:
        if slide["id"] == slide_id:
            return slide
    raise KeyError(f"Unknown slide_id: {slide_id}")
