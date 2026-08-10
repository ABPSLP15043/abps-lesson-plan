import streamlit as st
import json
import io
import os
import time
from google import genai
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(page_title="ABPS Baikunth - NEP 2023 Lesson Plan Generator", layout="wide")

st.title("🏫 The Aditya Birla Public School, Baikunth")
st.caption("Integrated NCF-SE 2023 | NEP 2020 Lesson Plan & Mind Map Generator")

# Retrieve API Key from Environment or Streamlit Secrets
API_KEY = os.getenv("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY", None)

# ---------------------------------------------------------
# PDF GENERATION FUNCTION (ABPS OFFICIAL FORMAT)
# ---------------------------------------------------------
def generate_pdf(meta, plan_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold',
        fontSize=12, alignment=1, textColor=colors.HexColor('#1A365D')
    )
    body_style = ParagraphStyle(
        'BodyStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=11
    )

    # Header Banner
    story.append(Paragraph("<b>THE ADITYA BIRLA PUBLIC SCHOOL</b>", title_style))
    story.append(Paragraph("BAIKUNTH, TAH- TILDA, DIST. RAIPUR (C.G.) – 493116", ParagraphStyle('Sub', parent=title_style, fontSize=8, fontName='Helvetica')))
    story.append(Spacer(1, 8))

    # Basic Info Block
    info_data = [
        [Paragraph(f"<b>Name of the Chapter:</b> {meta['chapter']}", body_style), Paragraph(f"<b>Subject:</b> {meta['subject']}", body_style)],
        [Paragraph(f"<b>Class:</b> {meta['grade']}", body_style), Paragraph(f"<b>Month:</b> {meta['month']}", body_style)],
        [Paragraph(f"<b>Teacher:</b> {meta['teacher']}", body_style), Paragraph(f"<b>No. of Periods:</b> {meta['periods']}", body_style)]
    ]
    info_table = Table(info_data, colWidths=[270, 270])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F7FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E0')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 8))

    # NCF 2023 Curriculum Goals & Competencies Section
    story.append(Paragraph(f"<b>Curriculum Goal (NCF-SE 2023):</b> {plan_data.get('curriculum_goal', '')}", body_style))
    story.append(Spacer(1, 4))
    
    comps_text = "<br/>".join([f"• {c}" for c in plan_data.get("relevant_competencies", [])])
    story.append(Paragraph(f"<b>Relevant Competencies:</b><br/>{comps_text}", body_style))
    story.append(Spacer(1, 8))

    # Detailed Content Table
    objs_text = "<br/>".join([f"• {o}" for o in plan_data.get("learning_objectives", [])])
    outcomes_text = "<br/>".join([f"• {e}" for e in plan_data.get("expected_learning_outcomes", [])])

    details_data = [
        [Paragraph("<b>Learning Objectives</b>", body_style), Paragraph(objs_text, body_style)],
        [Paragraph("<b>Expected Learning Outcomes</b>", body_style), Paragraph(outcomes_text, body_style)],
        [Paragraph("<b>Teaching Methodology</b>", body_style), Paragraph(", ".join(plan_data.get("teaching_methodology", [])), body_style)],
        [Paragraph("<b>Teaching Aids & Art Integration</b>", body_style), Paragraph(plan_data.get("teaching_aids_art", ""), body_style)],
        [Paragraph("<b>Connecting Previous Knowledge</b>", body_style), Paragraph(plan_data.get("previous_knowledge", ""), body_style)],
        [Paragraph("<b>Content / Teaching Points</b>", body_style), Paragraph(plan_data.get("content_points", ""), body_style)],
        [Paragraph("<b>Mind Map / Concept Flow</b>", body_style), Paragraph(plan_data.get("mind_map_structure", ""), body_style)],
        [Paragraph("<b>Experiential Learning / Projects</b>", body_style), Paragraph(plan_data.get("projects_experiential", ""), body_style)],
        [Paragraph("<b>Class Work & Home Work</b>", body_style), Paragraph(f"<b>Class Work:</b> {plan_data.get('class_work', '')}<br/><br/><b>Home Work:</b> {plan_data.get('home_work', '')}", body_style)],
        [Paragraph("<b>Remedial Measures</b>", body_style), Paragraph(plan_data.get("remedial_measures", ""), body_style)],
    ]
    
    details_table = Table(details_data, colWidths=[140, 400])
    details_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E0')),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#EDF2F7')),
        ('PADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(details_table)
    story.append(Spacer(1, 15))

    # Signatures
    sig_data = [["___________________\nSubject Teacher", "___________________\nHOD", "___________________\nPrincipal"]]
    sig_table = Table(sig_data, colWidths=[180, 180, 180])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(sig_table)

    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# SIDEBAR CONTROLS (Class VI to XII + Skill Subjects)
# ---------------------------------------------------------
st.sidebar.header("Lesson Details")
teacher_name = st.sidebar.text_input("Teacher Name", "Educator Name")

subject = st.sidebar.selectbox("Subject", [
    # Academic Subjects
    "Science", "Mathematics", "Social Science", "English", "Hindi", "Sanskrit",
    "Physics", "Chemistry", "Biology", "Computer Science", "IP",
    "Accountancy", "Business Studies", "Economics", "History", "Political Science", "Geography",
    # CBSE Skill Subjects
    "Artificial Intelligence (AI)", "Information Technology (IT)", "Financial Literacy",
    "Coding / Data Science", "Web Application", "Tourism", "Marketing & Sales", "Design Thinking"
])

grade = st.sidebar.selectbox("Class", [
    "Class VI", "Class VII", "Class VIII", "Class IX", "Class X", "Class XI", "Class XII"
])

month = st.sidebar.selectbox("Month", ["APRIL", "MAY", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER", "JANUARY", "FEBRUARY"])
chapter = st.sidebar.text_input("Chapter / Topic", "Heat Transfer in Nature")
periods = st.sidebar.number_input("No. of Periods", min_value=1, max_value=25, value=8)

# Fallback UI if API key isn't found in secrets/environment
if not API_KEY:
    API_KEY = st.sidebar.text_input("Gemini API Key (Optional Override)", type="password")

# ---------------------------------------------------------
# GENERATION ENGINE
# ---------------------------------------------------------
if st.sidebar.button("✨ Generate Lesson Plan", type="primary"):
    if not API_KEY:
        st.error("API Key is missing. Please configure GEMINI_API_KEY in Streamlit Secrets or environment variables.")
    else:
        with st.spinner("Generating NCF-SE 2023 aligned lesson plan with Competencies and Mind Map..."):
            try:
                client = genai.Client(api_key=API_KEY)
                prompt = f"""
                You are an expert CBSE curriculum planner for The Aditya Birla Public School, Baikunth.
                Generate a structured lesson plan for Class: {grade}, Subject: {subject}, Chapter: {chapter}, Month: {month}, Periods: {periods}.

                Follow NCF-SE 2023 / CBSE guidelines strictly. Output valid JSON with these exact keys:
                - curriculum_goal: string (e.g. CG-2: Develops procedural and conceptual fluency in subject-specific knowledge)
                - relevant_competencies: list of 3 strings (e.g. C-2.1: Demonstrates understanding of primary principles, C-2.2: Applies problem-solving techniques)
                - learning_objectives: list of 6-8 strings
                - expected_learning_outcomes: list of 6-8 strings
                - teaching_methodology: list of strings
                - teaching_aids_art: string
                - previous_knowledge: string
                - content_points: string
                - mind_map_structure: string (A structured text flowchart / mind map breakdown of main concepts)
                - projects_experiential: string
                - class_work: string
                - home_work: string
                - remedial_measures: string
                """

                response = None
                for attempt in range(3):
                    try:
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=prompt,
                            config={'response_mime_type': 'application/json'}
                        )
                        break
                    except Exception as e:
                        if "429" in str(e) and attempt < 2:
                            time.sleep(5)
                        else:
                            raise e

                st.session_state['plan_data'] = json.loads(response.text)
                st.session_state['meta'] = {
                    'teacher': teacher_name, 'subject': subject, 'grade': grade,
                    'chapter': chapter, 'periods': periods, 'month': month
                }
                st.success("Lesson Plan Generated Successfully!")

            except Exception as e:
                st.error(f"Error generating lesson plan: {str(e)}")

# ---------------------------------------------------------
# DISPLAY & DOWNLOAD
# ---------------------------------------------------------
if 'plan_data' in st.session_state:
    data = st.session_state['plan_data']
    meta = st.session_state['meta']
    
    st.subheader("📋 Generated Plan Preview")
    
    st.markdown(f"**Curriculum Goal (NCF 2023):** {data.get('curriculum_goal', '')}")
    st.markdown("**Relevant Competencies:**")
    for c in data.get("relevant_competencies", []):
        st.markdown(f"- {c}")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Learning Objectives")
        for obj in data.get("learning_objectives", []):
            st.markdown(f"- {obj}")
            
    with col2:
        st.markdown("### Expected Outcomes")
        for outcome in data.get("expected_learning_outcomes", []):
            st.markdown(f"- {outcome}")

    st.markdown("### 🧠 Concept Flow / Mind Map")
    st.code(data.get("mind_map_structure", ""), language="text")

    pdf_file = generate_pdf(meta, data)
    st.download_button(
        label="📄 Download ABPS Official Lesson Plan PDF",
        data=pdf_file,
        file_name=f"ABPS_Lesson_Plan_{meta['grade']}_{meta['chapter'].replace(' ', '_')}.pdf",
        mime="application/pdf",
        type="primary"
    )