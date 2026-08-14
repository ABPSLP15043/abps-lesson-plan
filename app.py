import io
import json
import os
from xml.sax.saxutils import escape as xml_escape

import streamlit as st
import docx
import graphviz
from pypdf import PdfReader
from google import genai

from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# MODELS  (current as of the Gemini Interactions API)
# ---------------------------------------------------------
CANDIDATE_MODELS = [
    "gemini-3.7-flash",
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
]

# ---------------------------------------------------------
# JSON SCHEMA FOR THE LESSON PLAN
# ---------------------------------------------------------
_STR_LIST = {"type": "array", "items": {"type": "string"}}

LESSON_PLAN_SCHEMA = {
    "type": "object",
    "properties": {
        "curriculum_goal": {"type": "string", "description": "NCF-SE 2023 curriculum goal for this chapter."},
        "relevant_competencies": _STR_LIST,
        "learning_objectives": _STR_LIST,
        "expected_learning_outcomes": _STR_LIST,
        "formulas_and_equations": _STR_LIST,
        "teaching_methodology": _STR_LIST,
        "teaching_aids": _STR_LIST,
        "art_integration": _STR_LIST,
        "previous_knowledge": _STR_LIST,
        "innovative_techniques": _STR_LIST,
        "content_points": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"section": {"type": "string"}, "topics": _STR_LIST},
                "required": ["section", "topics"],
            },
        },
        "projects_experiential": _STR_LIST,
        "skills_acquired": _STR_LIST,
        "values_inculcated": _STR_LIST,
        "multiple_assessment": {
            "type": "object",
            "properties": {
                "oral_questions": _STR_LIST,
                "worksheet": _STR_LIST,
                "practical": _STR_LIST,
                "exit_ticket": _STR_LIST,
            },
            "required": ["oral_questions", "worksheet", "practical", "exit_ticket"],
        },
        "class_work": _STR_LIST,
        "home_work": _STR_LIST,
        "remedial_measures": {
            "type": "object",
            "properties": {"slow_learners": _STR_LIST, "advanced_learners": _STR_LIST},
            "required": ["slow_learners", "advanced_learners"],
        },
        "resources": {
            "type": "object",
            "properties": {"books": _STR_LIST, "websites": _STR_LIST, "videos": _STR_LIST},
            "required": ["books", "websites", "videos"],
        },
    },
    "required": [
        "curriculum_goal", "relevant_competencies", "learning_objectives",
        "expected_learning_outcomes", "teaching_methodology", "teaching_aids",
        "art_integration", "previous_knowledge", "innovative_techniques",
        "content_points", "projects_experiential", "skills_acquired",
        "values_inculcated", "multiple_assessment", "class_work", "home_work",
        "remedial_measures", "resources",
    ],
}

# ---------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------
st.set_page_config(
    page_title="ABPS Baikunth - AI Lesson Plan Generator",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🏫 The Aditya Birla Public School, Baikunth")
st.caption("AI-Powered NCF-SE 2023 Lesson Plan & Mind Map Generator")

system_api_key = os.getenv("GEMINI_API_KEY", "")

# ---------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------
st.sidebar.header("📋 Lesson Plan Details")

teacher_name = st.sidebar.text_input("Teacher Name", "Educator Name")

user_api_key = st.sidebar.text_input(
    "Manual API Key (Optional)", type="password",
    help="Leave blank if GEMINI_API_KEY is set in your environment.",
)
active_api_key = user_api_key.strip() if user_api_key.strip() else system_api_key

subject = st.sidebar.selectbox("Subject", [
    "SCIENCE", "MATHEMATICS", "SOCIAL SCIENCE", "ENGLISH", "HINDI", "SANSKRIT",
    "MUSIC", "ART & CRAFT", "DANCE",
    "PHYSICS", "CHEMISTRY", "BIOLOGY", "COMPUTER SCIENCE", "IP",
    "ACCOUNTANCY", "BUSINESS STUDIES", "ECONOMICS", "HISTORY",
    "POLITICAL SCIENCE", "GEOGRAPHY",
    "ARTIFICIAL INTELLIGENCE (AI)", "INFORMATION TECHNOLOGY (IT)",
])

col_g, col_s = st.sidebar.columns(2)
with col_g:
    grade = st.selectbox("Class", ["VI", "VII", "VIII", "IX", "X", "XI", "XII"])
with col_s:
    section = st.text_input("Section", "A")

month = st.sidebar.selectbox("Month", [
    "APRIL", "MAY", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER",
    "NOVEMBER", "DECEMBER", "JANUARY", "FEBRUARY",
])
chapter = st.sidebar.text_input("Chapter / Topic", "Force and Laws of Motion")
periods = st.sidebar.number_input("No. of Periods", min_value=1, max_value=25, value=8)

st.sidebar.divider()

# Diagnostic: ask the API what models this key can actually use.
if st.sidebar.button("🔍 Check available models"):
    if not active_api_key:
        st.sidebar.error("Enter an API key first.")
    else:
        try:
            diag_client = genai.Client(api_key=active_api_key)
            names = []
            for m in diag_client.models.list():
                raw = getattr(m, "name", str(m))
                names.append(raw.replace("models/", ""))
            if names:
                st.sidebar.success(f"{len(names)} models available:")
                st.sidebar.code("\n".join(sorted(names)))
            else:
                st.sidebar.warning("No models returned for this key.")
        except Exception as exc:
            st.sidebar.error(f"Could not list models: {exc}")

# ---------------------------------------------------------
# CHAPTER INPUT
# ---------------------------------------------------------
st.subheader("📂 Step 1: Provide Chapter Content")
tab1, tab2 = st.tabs(["📄 Upload Chapter PDF (Max 25MB)", "📝 Paste Chapter Text/Notes"])

uploaded_pdf = None
pasted_text = ""

with tab1:
    uploaded_pdf = st.file_uploader(
        "Upload NCERT / Textbook Chapter PDF (up to 25 MB)", type=["pdf"]
    )

with tab2:
    pasted_text = st.text_area(
        "Paste text, outline, or key topics from the chapter here:", height=200
    )


# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------
def extract_text_from_pdf(uploaded_file, max_pages=40):
    reader = PdfReader(uploaded_file)
    parts = []
    for page in reader.pages[:max_pages]:
        text = page.extract_text()
        if text:
            parts.append(text)
    return "\n".join(parts)


def strip_code_fences(text):
    text = text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[-1] if "\n" in text else text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()


def build_prompt(subject, grade, section, chapter, month, periods, chapter_content):
    return (
        "You are an expert curriculum designer for Indian schools following the "
        "NCF-SE 2023 and NEP 2020 guidelines for NCERT textbooks.\n\n"
        "Using the chapter content supplied below, produce a HIGHLY SPECIFIC, UNIQUE and "
        "ACCURATE lesson plan. Every item must refer to the actual content of this chapter, "
        "never generic placeholders. Include real formulas, laws, definitions, dates or "
        "terms wherever the subject calls for them.\n\n"
        f"Subject: {subject}\n"
        f"Class / Section: {grade} - {section}\n"
        f"Chapter: {chapter}\n"
        f"Month: {month}\n"
        f"Number of periods: {periods}\n\n"
        "--- CHAPTER CONTENT ---\n"
        f"{chapter_content[:12000]}\n"
        "--- END CHAPTER CONTENT ---\n\n"
        "Give 6 to 10 items in each list where the chapter supports it. "
        "Give at least 4 sections in content_points, each with 3 or more topics. "
        "If the chapter has no formulas or equations, return an empty list for that field."
    )


def call_model(client, model_name, prompt):
    """Works with google-genai 2.x (Interactions API) and falls back to 1.x."""
    if hasattr(client, "interactions"):
        interaction = client.interactions.create(
            model=model_name,
            input=prompt,
            response_format={
                "type": "text",
                "mime_type": "application/json",
                "schema": LESSON_PLAN_SCHEMA,
            },
        )
        return interaction.output_text

    # Legacy SDK path (google-genai 1.x)
    from google.genai import types
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=LESSON_PLAN_SCHEMA,
        ),
    )
    return response.text


def generate_ai_lesson_plan(api_key, subject, grade, section, chapter,
                            month, periods, chapter_content):
    client = genai.Client(api_key=api_key)
    prompt = build_prompt(subject, grade, section, chapter, month, periods, chapter_content)

    failures = []
    for model_name in CANDIDATE_MODELS:
        try:
            raw = call_model(client, model_name, prompt)
            return json.loads(strip_code_fences(raw)), model_name
        except Exception as exc:
            failures.append(f"• {model_name} → {exc}")

    raise RuntimeError(
        "Every candidate model failed.\n\n"
        + "\n".join(failures)
        + "\n\nClick 'Check available models' in the sidebar to see what your "
          "API key can actually use, then update CANDIDATE_MODELS at the top of app.py."
    )


# ---------------------------------------------------------
# WORD EXPORT
# ---------------------------------------------------------
def set_cell_background(cell, fill_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_color)
    tc_pr.append(shd)


def bullets(items):
    if isinstance(items, list):
        return "\n".join(f"• {i}" for i in items)
    return str(items)


def generate_docx(meta, plan):
    doc = docx.Document()
    for sec in doc.sections:
        sec.top_margin = Inches(0.5)
        sec.bottom_margin = Inches(0.5)
        sec.left_margin = Inches(0.6)
        sec.right_margin = Inches(0.6)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = title_p.add_run("THE ADITYA BIRLA PUBLIC SCHOOL\n")
    run1.bold = True
    run1.font.size = Pt(14)
    run1.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    run2 = title_p.add_run(
        "BAIKUNTH, TAH- TILDA, DIST. RAIPUR (C.G.) – 493116\nLesson Plan\n"
    )
    run2.font.size = Pt(10)
    run2.bold = True

    info = doc.add_table(rows=3, cols=2)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info.style = "Table Grid"
    rows_data = [
        [f"Name of the Chapter: {meta['chapter']}", f"Subject: {meta['subject']}"],
        [f"Class: {meta['grade']} {meta['section']}", f"Month: {meta['month']}"],
        [f"Teacher: {meta['teacher']}", f"No. of Periods: {meta['periods']}"],
    ]
    for r_idx, row in enumerate(rows_data):
        for c_idx, val in enumerate(row):
            cell = info.cell(r_idx, c_idx)
            cell.text = val
            set_cell_background(cell, "F7FAFC")

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def add_section(title, content):
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"

        left = table.cell(0, 0)
        left.width = Inches(2.2)
        left.paragraphs[0].add_run(title).bold = True
        set_cell_background(left, "EDF2F7")

        right = table.cell(0, 1)
        right.width = Inches(4.8)
        if isinstance(content, list):
            right.paragraphs[0].text = ""
            for item in content:
                right.add_paragraph(f"• {item}")
        else:
            right.paragraphs[0].text = str(content)

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_section("Curriculum Goal (NCF-SE 2023)", plan.get("curriculum_goal", ""))
    add_section("Relevant Competencies", plan.get("relevant_competencies", []))
    add_section("Learning Objectives", plan.get("learning_objectives", []))
    add_section("Expected Learning Outcomes", plan.get("expected_learning_outcomes", []))

    if plan.get("formulas_and_equations"):
        add_section("Key Formulas, Laws & Equations", plan["formulas_and_equations"])

    add_section("Teaching Methodology", plan.get("teaching_methodology", []))

    aids = ("Teaching Aids:\n" + bullets(plan.get("teaching_aids", []))
            + "\n\nArt Integration:\n" + bullets(plan.get("art_integration", [])))
    add_section("Teaching Aids & Integration of Arts", aids)

    add_section("Connecting Previous Knowledge", plan.get("previous_knowledge", []))
    add_section("Innovative Techniques", plan.get("innovative_techniques", []))

    cp_lines = []
    for item in plan.get("content_points", []):
        cp_lines.append(f"{item.get('section', '')}:")
        for topic in item.get("topics", []):
            cp_lines.append(f"   • {topic}")
    add_section("Content / Teaching Points", "\n".join(cp_lines))

    add_section("Project / Experiential Learning", plan.get("projects_experiential", []))
    add_section("Skills Acquired", plan.get("skills_acquired", []))
    add_section("Values Inculcated", plan.get("values_inculcated", []))

    ass = plan.get("multiple_assessment", {})
    ass_text = (
        "Oral Questions:\n" + bullets(ass.get("oral_questions", []))
        + "\n\nWorksheet:\n" + bullets(ass.get("worksheet", []))
        + "\n\nPractical Assessment:\n" + bullets(ass.get("practical", []))
        + "\n\nExit Ticket:\n" + bullets(ass.get("exit_ticket", []))
    )
    add_section("Multiple / Periodic Assessment", ass_text)

    add_section("Class Work", plan.get("class_work", []))
    add_section("Home Work", plan.get("home_work", []))

    rem = plan.get("remedial_measures", {})
    rem_text = ("Slow Learners:\n" + bullets(rem.get("slow_learners", []))
                + "\n\nAdvanced Learners:\n" + bullets(rem.get("advanced_learners", [])))
    add_section("Remedial Measures", rem_text)

    res = plan.get("resources", {})
    res_text = ("Books:\n" + bullets(res.get("books", []))
                + "\n\nWebsites:\n" + bullets(res.get("websites", []))
                + "\n\nVideos:\n" + bullets(res.get("videos", [])))
    add_section("Resources & References", res_text)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    sig = doc.add_table(rows=1, cols=3)
    sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig.cell(0, 0).text = "___________________\nSubject Teacher"
    sig.cell(0, 1).text = "___________________\nHOD"
    sig.cell(0, 2).text = "___________________\nPrincipal"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# PDF EXPORT
# ---------------------------------------------------------
def esc(text):
    """Escape &, < and > so ReportLab never crashes on chapter text."""
    return xml_escape(str(text))


def html_bullets(items):
    if isinstance(items, list):
        return "<br/>".join(f"• {esc(i)}" for i in items)
    return esc(items)


def generate_pdf(meta, plan):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28,
    )
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleStyle", parent=styles["Heading1"], fontName="Helvetica-Bold",
        fontSize=13, alignment=1, textColor=colors.HexColor("#1A365D"),
    )
    sub_style = ParagraphStyle(
        "SubStyle", parent=styles["Normal"], fontName="Helvetica",
        fontSize=8.5, alignment=1,
    )
    body = ParagraphStyle(
        "BodyStyle", parent=styles["Normal"], fontName="Helvetica",
        fontSize=8, leading=11,
    )

    story.append(Paragraph("THE ADITYA BIRLA PUBLIC SCHOOL", title_style))
    story.append(Paragraph(
        "BAIKUNTH, TAH- TILDA, DIST. RAIPUR (C.G.) – 493116<br/><b>Lesson Plan</b>",
        sub_style,
    ))
    story.append(Spacer(1, 8))

    info_data = [
        [Paragraph(f"<b>Chapter:</b> {esc(meta['chapter'])}", body),
         Paragraph(f"<b>Subject:</b> {esc(meta['subject'])}", body)],
        [Paragraph(f"<b>Class:</b> {esc(meta['grade'])} {esc(meta['section'])}", body),
         Paragraph(f"<b>Month:</b> {esc(meta['month'])}", body)],
        [Paragraph(f"<b>Teacher:</b> {esc(meta['teacher'])}", body),
         Paragraph(f"<b>No. of Periods:</b> {esc(meta['periods'])}", body)],
    ]
    info_table = Table(info_data, colWidths=[260, 260])
    info_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7FAFC")),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 6))

    cp_html = ""
    for item in plan.get("content_points", []):
        cp_html += f"<b>{esc(item.get('section', ''))}</b><br/>"
        for topic in item.get("topics", []):
            cp_html += f"• {esc(topic)}<br/>"

    ass = plan.get("multiple_assessment", {})
    ass_html = (
        f"<b>Oral Questions:</b><br/>{html_bullets(ass.get('oral_questions', []))}<br/><br/>"
        f"<b>Worksheet:</b><br/>{html_bullets(ass.get('worksheet', []))}<br/><br/>"
        f"<b>Practical:</b><br/>{html_bullets(ass.get('practical', []))}<br/><br/>"
        f"<b>Exit Ticket:</b><br/>{html_bullets(ass.get('exit_ticket', []))}"
    )

    rem = plan.get("remedial_measures", {})
    rem_html = (
        f"<b>Slow Learners:</b><br/>{html_bullets(rem.get('slow_learners', []))}<br/><br/>"
        f"<b>Advanced Learners:</b><br/>{html_bullets(rem.get('advanced_learners', []))}"
    )

    res = plan.get("resources", {})
    res_html = (
        f"<b>Books:</b><br/>{html_bullets(res.get('books', []))}<br/><br/>"
        f"<b>Websites:</b><br/>{html_bullets(res.get('websites', []))}<br/><br/>"
        f"<b>Videos:</b><br/>{html_bullets(res.get('videos', []))}"
    )

    aids_html = (
        f"<b>Teaching Aids:</b><br/>{html_bullets(plan.get('teaching_aids', []))}<br/><br/>"
        f"<b>Art Integration:</b><br/>{html_bullets(plan.get('art_integration', []))}"
    )

    rows = [
        ["Curriculum Goal (NCF-SE 2023)", esc(plan.get("curriculum_goal", ""))],
        ["Relevant Competencies", html_bullets(plan.get("relevant_competencies", []))],
        ["Learning Objectives", html_bullets(plan.get("learning_objectives", []))],
        ["Expected Learning Outcomes", html_bullets(plan.get("expected_learning_outcomes", []))],
    ]
    if plan.get("formulas_and_equations"):
        rows.append(["Key Formulas & Equations", html_bullets(plan["formulas_and_equations"])])
    rows.extend([
        ["Teaching Methodology", html_bullets(plan.get("teaching_methodology", []))],
        ["Teaching Aids & Art Integration", aids_html],
        ["Connecting Previous Knowledge", html_bullets(plan.get("previous_knowledge", []))],
        ["Innovative Techniques", html_bullets(plan.get("innovative_techniques", []))],
        ["Content / Teaching Points", cp_html],
        ["Project / Experiential Learning", html_bullets(plan.get("projects_experiential", []))],
        ["Skills Acquired", html_bullets(plan.get("skills_acquired", []))],
        ["Values Inculcated", html_bullets(plan.get("values_inculcated", []))],
        ["Multiple / Periodic Assessment", ass_html],
        ["Class Work", html_bullets(plan.get("class_work", []))],
        ["Home Work", html_bullets(plan.get("home_work", []))],
        ["Remedial Measures", rem_html],
        ["Resources & References", res_html],
    ])

    details_data = [
        [Paragraph(f"<b>{label}</b>", body), Paragraph(value, body)]
        for label, value in rows
    ]

    details_table = Table(details_data, colWidths=[135, 385], repeatRows=0)
    details_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EDF2F7")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(details_table)
    story.append(Spacer(1, 14))

    sig_table = Table(
        [["___________________\nSubject Teacher",
          "___________________\nHOD",
          "___________________\nPrincipal"]],
        colWidths=[173, 173, 173],
    )
    sig_table.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    story.append(sig_table)

    doc.build(story)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# GENERATE
# ---------------------------------------------------------
st.divider()

if st.button("✨ Generate AI Lesson Plan & Mind Map", type="primary", use_container_width=True):
    if not active_api_key:
        st.error("⚠️ Gemini API key is missing. Add it in the sidebar or set GEMINI_API_KEY.")
    elif uploaded_pdf is not None and uploaded_pdf.size > 25 * 1024 * 1024:
        st.error("⚠️ File exceeds 25 MB. Please upload a smaller PDF.")
    else:
        chapter_content = ""
        if uploaded_pdf is not None:
            with st.spinner("Extracting text from the uploaded PDF..."):
                chapter_content = extract_text_from_pdf(uploaded_pdf)
            if not chapter_content.strip():
                st.warning(
                    "No selectable text found in that PDF (it may be a scanned image). "
                    "Paste the chapter text in the second tab instead."
                )
        elif pasted_text.strip():
            chapter_content = pasted_text.strip()
        else:
            chapter_content = (
                f"Standard NCERT syllabus content for {subject}, Class {grade}, "
                f"Chapter: {chapter}."
            )

        with st.spinner("🧠 Analysing the chapter with Gemini..."):
            try:
                plan_data, used_model = generate_ai_lesson_plan(
                    active_api_key, subject, grade, section,
                    chapter, month, periods, chapter_content,
                )
                st.session_state["plan_data"] = plan_data
                st.session_state["meta"] = {
                    "teacher": teacher_name, "subject": subject, "grade": grade,
                    "section": section, "chapter": chapter,
                    "periods": periods, "month": month,
                }
                st.success(f"🎉 Lesson plan generated using **{used_model}**")
            except Exception as exc:
                st.error("Could not generate the lesson plan.")
                st.code(str(exc))

# ---------------------------------------------------------
# OUTPUT
# ---------------------------------------------------------
if "plan_data" in st.session_state:
    data = st.session_state["plan_data"]
    meta = st.session_state["meta"]

    st.subheader(f"📋 Preview: {meta['chapter']} ({meta['subject']})")
    st.markdown(f"**Curriculum Goal:** {data.get('curriculum_goal', '')}")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Learning Objectives")
        for obj in data.get("learning_objectives", []):
            st.markdown(f"- {obj}")
    with col2:
        st.markdown("### Expected Outcomes")
        for outcome in data.get("expected_learning_outcomes", []):
            st.markdown(f"- {outcome}")

    if data.get("formulas_and_equations"):
        st.divider()
        st.subheader("📐 Key Formulas & Laws Extracted")
        for formula in data["formulas_and_equations"]:
            st.info(f"⚡ {formula}")

    st.divider()
    st.subheader(f"🧠 Visual Mind Map: {meta['chapter']}")
    try:
        dot = graphviz.Digraph(comment=meta["chapter"])
        dot.attr(rankdir="LR", splines="ortho")
        dot.attr("node", style="filled", fillcolor="#EBF8FF",
                 color="#2B6CB0", fontname="Helvetica", fontsize="10", shape="box")
        dot.node("CENTER", meta["chapter"], fillcolor="#2B6CB0", fontcolor="white")

        for s_idx, item in enumerate(data.get("content_points", [])):
            section_id = f"SEC{s_idx}"
            dot.node(section_id, item.get("section", "Core Concepts"),
                     shape="ellipse", fillcolor="#E2E8F0")
            dot.edge("CENTER", section_id)
            for t_idx, topic in enumerate(item.get("topics", [])):
                topic_id = f"SEC{s_idx}_T{t_idx}"
                dot.node(topic_id, topic, fillcolor="#FFFFFF")
                dot.edge(section_id, topic_id)

        st.graphviz_chart(dot, use_container_width=True)
    except Exception:
        st.warning("Mind map rendering skipped. The downloads below still contain everything.")

    st.divider()
    safe_chapter = meta["chapter"].replace(" ", "_").replace("/", "-")

    col_pdf, col_docx = st.columns(2)
    with col_pdf:
        st.download_button(
            label="📄 Download Official PDF",
            data=generate_pdf(meta, data),
            file_name=f"ABPS_Lesson_Plan_{meta['grade']}_{safe_chapter}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with col_docx:
        st.download_button(
            label="📝 Download Editable Word (.docx)",
            data=generate_docx(meta, data),
            file_name=f"ABPS_Lesson_Plan_{meta['grade']}_{safe_chapter}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
