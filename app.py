import streamlit as st
import io
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(page_title="ABPS Baikunth - Lesson Plan Generator", layout="wide")

st.title("🏫 The Aditya Birla Public School, Baikunth")
st.caption("Integrated NCF-SE 2023 | NEP 2020 Lesson Plan Generator (PDF & Word)")

# ---------------------------------------------------------
# COMPREHENSIVE CONTENT BUILDER ENGINE
# ---------------------------------------------------------
def build_comprehensive_plan(subject, grade, section, chapter, month, periods):
    return {
        "curriculum_goal": f"CG-3: Explores {subject} by understanding foundational concepts, processes, and interactions, while promoting responsible and analytical practices.",
        "relevant_competencies": [
            f"C-3.1: Explains core principles of {chapter} through observation and scientific understanding.",
            "C-3.2: Relates theoretical concepts to real-world phenomena and practical applications.",
            "C-3.3: Demonstrates responsible behavior, informed decision-making, and analytical practices based on knowledge."
        ],
        "learning_objectives": [
            f"Explain the primary concepts and mechanisms of {chapter}.",
            "Identify key terms, definitions, and underlying scientific/academic principles.",
            "Understand practical applications in daily life and industrial contexts.",
            "Relate theoretical models with environmental and real-world interactions.",
            "Appreciate sustainable practices and problem-solving methodologies."
        ],
        "expected_learning_outcomes": [
            f"Explain key definitions and core mechanisms of {chapter} with examples.",
            "Demonstrate understanding using practical experiments or case studies.",
            "Differentiate between primary concepts, classifications, and components.",
            "Describe natural cycles, applications, or functional structures related to the topic.",
            "Apply acquired concepts in daily life and academic assessments."
        ],
        "teaching_methodology": [
            "Inquiry Based Learning", "Activity Based Learning", "Demonstration Method",
            "Cooperative Learning", "Think-Pair-Share", "Experiential Learning",
            "Competency Based Learning", "Discussion Method", "Observation Method"
        ],
        "teaching_aids": [
            "Smart Board / Interactive PPT", "NCERT Animations & Videos",
            "Worksheets & Handouts", "Demonstration Equipment & Charts"
        ],
        "art_integration": [
            f"{chapter} Mind Map / Flow Chart", "Concept Posters & Infographics",
            "Illustrative Diagrams & Comic Strips"
        ],
        "previous_knowledge": [
            f"Why does this phenomenon occur in daily life related to {subject}?",
            f"Where do we observe the practical applications of {chapter} around us?",
            "How do environmental conditions affect these processes?"
        ],
        "innovative_techniques": [
            "Blended Learning", "QR Code Videos", "Interactive PPT", "Mind Mapping",
            "Virtual Laboratory", "Quizizz / Kahoot", "Exit Tickets"
        ],
        "content_points": [
            {"section": "Introduction & Fundamentals", "topics": [f"Basic definitions and importance of {chapter}", "Everyday life examples and historical context"]},
            {"section": "Core Concept Analysis", "topics": ["Detailed theoretical framework", "Mechanisms, formulas, and structural breakdown"]},
            {"section": "Practical Applications & Phenomena", "topics": ["Real-world applications and observational studies", "Environmental and industrial relevance"]},
            {"section": "Review & Synthesis", "topics": ["Concept mapping and summary worksheets", "Interactive revision and quiz sessions"]}
        ],
        "projects_experiential": [
            f"Demonstrate core principles of {chapter} using simple materials.",
            f"Construct a working model or visual poster detailing {chapter}.",
            "Prepare a case study report on practical applications in your locality."
        ],
        "skills_acquired": [
            "Observation", "Scientific Inquiry", "Critical Thinking",
            "Classification", "Experimentation", "Data Analysis", "Teamwork"
        ],
        "values_inculcated": [
            "Environmental Awareness", "Scientific Temper", "Curiosity",
            "Teamwork", "Responsibility", "Sustainable Living"
        ],
        "multiple_assessment": {
            "oral_questions": [f"What is the core principle of {chapter}?", "Give two real-life examples."],
            "worksheet": ["MCQs", "Assertion–Reason", "Match the Following", "Case Study Questions"],
            "practical": ["Experiment Performance", "Observation Skills", "Scientific Reasoning"],
            "exit_ticket": ["Write one new concept learned today and one real-life application."]
        },
        "class_work": ["Concept Notes & Diagrams", "Observation Tables", "Mind Maps & Flowcharts", "Solving NCERT Worksheet Questions"],
        "home_work": [f"Prepare a detailed concept map of {chapter}.", "Answer review questions from the textbook.", "List 5 daily life examples of the topic."],
        "remedial_measures": {
            "slow_learners": ["Picture cards and visual aids", "Peer tutoring & guided study", "Simplified worksheets and oral questioning"],
            "advanced_learners": ["Research assignments on advanced applications", "Preparation of presentations / solar models", "In-depth case study analysis"]
        },
        "resources": {
            "books": [f"NCERT {grade} {subject} Textbook", "CBSE Lab Manual & Exemplar"],
            "websites": ["DIKSHA Portal", "NCERT e-Resources", "CBSE Academic", "National Digital Library of India"],
            "videos": ["NCERT Official Videos", "DIKSHA Learning Modules", "Khan Academy"]
        }
    }

# ---------------------------------------------------------
# WORD DOCUMENT GENERATOR (.DOCX)
# ---------------------------------------------------------
def set_cell_background(cell, fill_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

def generate_docx(meta, plan_data):
    doc = docx.Document()
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.6)
        s.right_margin = Inches(0.6)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t1 = p_title.add_run("THE ADITYA BIRLA PUBLIC SCHOOL\n")
    run_t1.bold = True
    run_t1.font.size = Pt(14)
    run_t1.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)

    run_t2 = p_title.add_run("BAIKUNTH, TAH- TILDA, DIST. RAIPUR (C.G.) – 493116\nLesson Plan\n")
    run_t2.font.size = Pt(10)
    run_t2.bold = True

    info_table = doc.add_table(rows=3, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.style = 'Table Grid'
    
    rows_data = [
        [f"Name of the Chapter: {meta['chapter']}", f"Subject: {meta['subject']}"],
        [f"Class: {meta['grade']} {meta['section']}", f"Month: {meta['month']}"],
        [f"Teacher: {meta['teacher']}", f"No. of Periods: {meta['periods']}"]
    ]
    
    for r_idx, r in enumerate(rows_data):
        for c_idx, val in enumerate(r):
            cell = info_table.cell(r_idx, c_idx)
            cell.text = val
            set_cell_background(cell, "F7FAFC")

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def add_section_table(title, content):
        t = doc.add_table(rows=1, cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.style = 'Table Grid'
        
        c0 = t.cell(0, 0)
        c0.width = Inches(2.2)
        c0.paragraphs[0].add_run(title).bold = True
        set_cell_background(c0, "EDF2F7")
        
        c1 = t.cell(0, 1)
        c1.width = Inches(4.8)
        
        if isinstance(content, list):
            for item in content:
                c1.add_paragraph(f"• {item}")
        else:
            c1.paragraphs[0].text = str(content)
            
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_section_table("Curriculum Goal (NCF-SE 2023)", plan_data.get("curriculum_goal", ""))
    add_section_table("Relevant Competencies", plan_data.get("relevant_competencies", []))
    add_section_table("Learning Objectives", plan_data.get("learning_objectives", []))
    add_section_table("Expected Learning Outcomes", plan_data.get("expected_learning_outcomes", []))
    add_section_table("Teaching Methodology", plan_data.get("teaching_methodology", []))
    
    aids_text = "Teaching Aids:\n" + "\n".join([f"• {a}" for a in plan_data.get("teaching_aids", [])])
    aids_text += "\n\nArt Integration:\n" + "\n".join([f"• {a}" for a in plan_data.get("art_integration", [])])
    add_section_table("Teaching Aids &\nIntegration of Arts", aids_text)

    add_section_table("Connecting Previous Knowledge", plan_data.get("previous_knowledge", []))
    add_section_table("Innovative Techniques\n(Blended/Mind Map)", plan_data.get("innovative_techniques", []))

    cp_text = ""
    for item in plan_data.get("content_points", []):
        cp_text += f"{item.get('section','')}:\n"
        for t in item.get("topics", []):
            cp_text += f"  • {t}\n"
    add_section_table("Content / Teaching Points", cp_text.strip())

    add_section_table("Project / Experiential Learning", plan_data.get("projects_experiential", []))
    add_section_table("Skills Acquired", plan_data.get("skills_acquired", []))
    add_section_table("Values Inculcated", plan_data.get("values_inculcated", []))

    ass = plan_data.get("multiple_assessment", {})
    ass_text = "Oral Questions:\n" + "\n".join([f"• {q}" for q in ass.get("oral_questions", [])])
    ass_text += "\n\nWorksheet:\n" + "\n".join([f"• {w}" for w in ass.get("worksheet", [])])
    ass_text += "\n\nPractical Assessment:\n" + "\n".join([f"• {p}" for p in ass.get("practical", [])])
    ass_text += "\n\nExit Ticket:\n" + "\n".join([f"• {e}" for e in ass.get("exit_ticket", [])])
    add_section_table("Multiple / Periodic Assessment", ass_text)

    add_section_table("Class Work", plan_data.get("class_work", []))
    add_section_table("Home Work", plan_data.get("home_work", []))

    rem = plan_data.get("remedial_measures", {})
    rem_text = "Slow Learners:\n" + "\n".join([f"• {s}" for s in rem.get("slow_learners", [])])
    rem_text += "\n\nAdvanced Learners:\n" + "\n".join([f"• {a}" for a in rem.get("advanced_learners", [])])
    add_section_table("Remedial Measures", rem_text)

    res = plan_data.get("resources", {})
    res_text = "Books:\n" + "\n".join([f"• {b}" for b in res.get("books", [])])
    res_text += "\n\nWebsites:\n" + "\n".join([f"• {w}" for w in res.get("websites", [])])
    res_text += "\n\nVideos:\n" + "\n".join([f"• {v}" for v in res.get("videos", [])])
    add_section_table("Resources & References", res_text)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    sig_table = doc.add_table(rows=1, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.cell(0, 0).text = "___________________\nSubject Teacher"
    sig_table.cell(0, 1).text = "___________________\nHOD"
    sig_table.cell(0, 2).text = "___________________\nPrincipal"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# PDF GENERATOR
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

    story.append(Paragraph("<b>THE ADITYA BIRLA PUBLIC SCHOOL</b>", title_style))
    story.append(Paragraph("BAIKUNTH, TAH- TILDA, DIST. RAIPUR (C.G.) – 493116<br/><b>Lesson Plan</b>", ParagraphStyle('Sub', parent=title_style, fontSize=8, fontName='Helvetica')))
    story.append(Spacer(1, 8))

    info_data = [
        [Paragraph(f"<b>Name of the Chapter:</b> {meta['chapter']}", body_style), Paragraph(f"<b>Subject:</b> {meta['subject']}", body_style)],
        [Paragraph(f"<b>Class:</b> {meta['grade']} {meta['section']}", body_style), Paragraph(f"<b>Month:</b> {meta['month']}", body_style)],
        [Paragraph(f"<b>Teacher:</b> {meta['teacher']}", body_style), Paragraph(f"<b>No. of Periods:</b> {meta['periods']}", body_style)]
    ]
    info_table = Table(info_data, colWidths=[270, 270])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F7FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E0')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 6))

    def make_bullet_list(items):
        return "<br/>".join([f"• {i}" for i in items]) if isinstance(items, list) else str(items)

    cp_formatted = ""
    for item in plan_data.get("content_points", []):
        cp_formatted += f"<b>{item.get('section','')}</b><br/>"
        for t in item.get("topics", []):
            cp_formatted += f"• {t}<br/>"

    ass = plan_data.get("multiple_assessment", {})
    ass_formatted = f"<b>Oral Questions:</b><br/>{make_bullet_list(ass.get('oral_questions',[]))}<br/><br/>"
    ass_formatted += f"<b>Worksheet:</b><br/>{make_bullet_list(ass.get('worksheet',[]))}<br/><br/>"
    ass_formatted += f"<b>Practical Assessment:</b><br/>{make_bullet_list(ass.get('practical',[]))}<br/><br/>"
    ass_formatted += f"<b>Exit Ticket:</b><br/>{make_bullet_list(ass.get('exit_ticket',[]))}"

    rem = plan_data.get("remedial_measures", {})
    rem_formatted = f"<b>Slow Learners:</b><br/>{make_bullet_list(rem.get('slow_learners',[]))}<br/><br/><b>Advanced Learners:</b><br/>{make_bullet_list(rem.get('advanced_learners',[]))}"

    res = plan_data.get("resources", {})
    res_formatted = f"<b>Books:</b><br/>{make_bullet_list(res.get('books',[]))}<br/><br/><b>Websites:</b><br/>{make_bullet_list(res.get('websites',[]))}<br/><br/><b>Videos:</b><br/>{make_bullet_list(res.get('videos',[]))}"

    aids_formatted = f"<b>Teaching Aids:</b><br/>{make_bullet_list(plan_data.get('teaching_aids',[]))}<br/><br/><b>Art Integration:</b><br/>{make_bullet_list(plan_data.get('art_integration',[]))}"

    details_data = [
        [Paragraph("<b>Curriculum Goal (NCF 2023)</b>", body_style), Paragraph(plan_data.get("curriculum_goal",""), body_style)],
        [Paragraph("<b>Relevant Competencies</b>", body_style), Paragraph(make_bullet_list(plan_data.get("relevant_competencies",[])), body_style)],
        [Paragraph("<b>Learning Objectives</b>", body_style), Paragraph(make_bullet_list(plan_data.get("learning_objectives",[])), body_style)],
        [Paragraph("<b>Expected Learning Outcomes</b>", body_style), Paragraph(make_bullet_list(plan_data.get("expected_learning_outcomes",[])), body_style)],
        [Paragraph("<b>Teaching Methodology</b>", body_style), Paragraph(make_bullet_list(plan_data.get("teaching_methodology",[])), body_style)],
        [Paragraph("<b>Teaching Aids & Integration of Arts</b>", body_style), Paragraph(aids_formatted, body_style)],
        [Paragraph("<b>Connecting Previous Knowledge</b>", body_style), Paragraph(make_bullet_list(plan_data.get("previous_knowledge",[])), body_style)],
        [Paragraph("<b>Innovative Techniques</b>", body_style), Paragraph(make_bullet_list(plan_data.get("innovative_techniques",[])), body_style)],
        [Paragraph("<b>Content / Teaching Points</b>", body_style), Paragraph(cp_formatted, body_style)],
        [Paragraph("<b>Project / Experiential Learning</b>", body_style), Paragraph(make_bullet_list(plan_data.get("projects_experiential",[])), body_style)],
        [Paragraph("<b>Skills Acquired</b>", body_style), Paragraph(make_bullet_list(plan_data.get("skills_acquired",[])), body_style)],
        [Paragraph("<b>Values Inculcated</b>", body_style), Paragraph(make_bullet_list(plan_data.get("values_inculcated",[])), body_style)],
        [Paragraph("<b>Multiple / Periodic Assessment</b>", body_style), Paragraph(ass_formatted, body_style)],
        [Paragraph("<b>Class Work</b>", body_style), Paragraph(make_bullet_list(plan_data.get("class_work",[])), body_style)],
        [Paragraph("<b>Home Work</b>", body_style), Paragraph(make_bullet_list(plan_data.get("home_work",[])), body_style)],
        [Paragraph("<b>Remedial Measures</b>", body_style), Paragraph(rem_formatted, body_style)],
        [Paragraph("<b>Resources & References</b>", body_style), Paragraph(res_formatted, body_style)],
    ]
    
    details_table = Table(details_data, colWidths=[140, 400])
    details_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E0')),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#EDF2F7')),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(details_table)
    story.append(Spacer(1, 12))

    sig_data = [["___________________\nSubject Teacher", "___________________\nHOD", "___________________\nPrincipal"]]
    sig_table = Table(sig_data, colWidths=[180, 180, 180])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(sig_table)

    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# UI CONTROLS
# ---------------------------------------------------------
st.sidebar.header("Lesson Form Details")
teacher_name = st.sidebar.text_input("Teacher Name", "Educator Name")

subject = st.sidebar.selectbox("Subject", [
    "SCIENCE", "MATHEMATICS", "SOCIAL SCIENCE", "ENGLISH", "HINDI", "SANSKRIT",
    "PHYSICS", "CHEMISTRY", "BIOLOGY", "COMPUTER SCIENCE", "IP",
    "ACCOUNTANCY", "BUSINESS STUDIES", "ECONOMICS", "HISTORY", "POLITICAL SCIENCE", "GEOGRAPHY",
    "ARTIFICIAL INTELLIGENCE (AI)", "INFORMATION TECHNOLOGY (IT)", "FINANCIAL LITERACY"
])

col_g, col_s = st.sidebar.columns(2)
with col_g:
    grade = st.selectbox("Class", ["VI", "VII", "VIII", "IX", "X", "XI", "XII"])
with col_s:
    section = st.text_input("Section", "B")

month = st.sidebar.selectbox("Month", ["APRIL", "MAY", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER", "JANUARY", "FEBRUARY"])
chapter = st.sidebar.text_input("Chapter / Topic", "Heat Transfer in Nature")
periods = st.sidebar.number_input("No. of Periods", min_value=1, max_value=25, value=8)

# ---------------------------------------------------------
# GENERATION ENGINE
# ---------------------------------------------------------
if st.sidebar.button("✨ Generate Official Lesson Plan", type="primary"):
    with st.spinner("Building official ABPS lesson plan format..."):
        plan_data = build_comprehensive_plan(subject, grade, section, chapter, month, periods)
        st.session_state['plan_data'] = plan_data
        st.session_state['meta'] = {
            'teacher': teacher_name, 'subject': subject, 'grade': grade,
            'section': section, 'chapter': chapter, 'periods': periods, 'month': month
        }
        st.success("Lesson Plan Generated Successfully!")

# ---------------------------------------------------------
# DISPLAY & DOWNLOADS
# ---------------------------------------------------------
if 'plan_data' in st.session_state:
    data = st.session_state['plan_data']
    meta = st.session_state['meta']
    
    st.subheader(f"📋 Preview: {meta['chapter']} ({meta['subject']})")
    st.markdown(f"**Curriculum Goal:** {data['curriculum_goal']}")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Learning Objectives")
        for obj in data['learning_objectives']:
            st.markdown(f"- {obj}")
            
    with col2:
        st.markdown("### Expected Outcomes")
        for outcome in data['expected_learning_outcomes']:
            st.markdown(f"- {outcome}")

    st.divider()
    
    col_pdf, col_docx = st.columns(2)
    
    with col_pdf:
        pdf_file = generate_pdf(meta, data)
        st.download_button(
            label="📄 Download Official Lesson Plan (PDF)",
            data=pdf_file,
            file_name=f"ABPS_Lesson_Plan_{meta['grade']}_{meta['chapter'].replace(' ', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        
    with col_docx:
        docx_file = generate_docx(meta, data)
        st.download_button(
            label="📝 Download Editable Lesson Plan (Word .docx)",
            data=docx_file,
            file_name=f"ABPS_Lesson_Plan_{meta['grade']}_{meta['chapter'].replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
