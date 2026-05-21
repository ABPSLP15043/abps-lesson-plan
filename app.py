import streamlit as st
import google.generativeai as genai
import docx
from io import BytesIO

# Interface styling matching ABPS Branding
st.set_page_config(page_title="ABPS NCF Lesson Plan Generator", page_icon="📝", layout="wide")
st.markdown("""
    <style>
    .main-title {color: #003366; font-size: 32px; font-weight: bold; text-align: center;}
    .subtitle {color: #F7931E; font-size: 20px; text-align: center; margin-bottom: 25px; font-weight: 500;}
    </style>
""", unsafe_allowed_html=True)

st.markdown('<div class="main-title">The Aditya Birla Public School, Baikunth</div>', unsafe_allowed_html=True)
st.markdown('<div class="subtitle">AI-Powered NCF Curricular Goal & Lesson Plan Builder</div>', unsafe_allow_html=True)

# Sidebar for Key Configuration
st.sidebar.header("Configuration")
api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-pro')

    # School Template Inputs
    with st.container():
        col1, col2, col3 = st.columns(3)
        with col1:
            subject = st.text_input("Subject")
            grade = st.text_input("Class")
        with col2:
            chapter = st.text_input("Chapter/Topic Name")
            month = st.text_input("Month")
        with col3:
            periods = st.text_input("No. of Periods Required")

    core_concept = st.text_area("Briefly describe your teaching focus or paste rough lecture notes:")

    if st.button("Generate NCF Compliant Lesson Plan"):
        with st.spinner("Analyzing NCF 2023 Guidelines & Building Document Structure..."):
            
            prompt = f"""
            You are an expert curriculum designer for CBSE schools following the National Curriculum Framework (NCF 2023).
            Generate an exhaustive lesson plan based on this input:
            Subject: {subject}
            Class: {grade}
            Chapter: {chapter}
            Core Topic: {core_concept}

            Provide comprehensive text content for each of the following exact school template headers:
            
            1. [NCF Curricular Goals & Competencies]: Map the explicit NCF 2023 Curricular Goals (CG) and Competencies codes relevant to this grade level and subject domain.
            2. [Learning Objectives]: State clearly what students will learn about.
            3. [Expected Learning Outcomes]: Write "At the end of this lesson, students will be able to:" followed by distinct, measurable behavioral milestones.
            4. [Teaching Methodology]: Detail the precise step-by-step pedagogical delivery plan.
            5. [Skill and Competencies developed]: Highlight subject-specific core skills targeted.
            6. [Teaching Aids & Integration of Arts]: List concrete physical/digital aids and cross-curricular art elements.
            7. [Connecting Previous Knowledge]: Formulate opening diagnostic check questions.
            8. [Innovative Techniques]: Outline a structural map layout for a Blended Learning path, Mind Map, or Flow Chart.
            9. [Content/ Teaching points]: Bulleted summary of high-yield instructional facts.
            10. [Project/ Enrichment / Experiential Learning Activity]: Design a hands-on, realistic assignment.
            11. [Skills Acquired]: Core lifelong learning skills gained.
            12. [Values Inculcated]: Value education and ethics connected to this lesson.
            13. [Multiple Assessment / Periodic Assessment]: Diagnostic items or quick check for understanding parameters.
            14. [Class Work]: Suggested class room exercises.
            15. [Home Work]: Homework problems.
            16. [Remedial Measures]: Dedicated intervention tasks for diverse learners.
            17. [Resources / Suggested References]: Books, verified educational websites, and interactive animation video directions.
            """
            
            response = model.generate_content(prompt)
            st.success("Plan Drafted and Aligned!")
            st.markdown(response.text)
            
            # Assembly of the Microsoft Word file matching the exact layout
            doc = docx.Document()
            doc.add_heading('The Aditya Birla Public School, Baikunth', 0)
            doc.add_heading('Official Lesson Plan Document', level=1)
            
            p = doc.add_paragraph()
            p.add_run(f"Chapter: {chapter} | Subject: {subject}\n").bold = True
            p.add_run(f"Class: {grade} | Month: {month} | No. of Periods: {periods}\n")
            doc.add_paragraph(response.text)
            
            bio = BytesIO()
            doc.save(bio)
            st.download_button(
                label="📥 Download Official .docx Lesson Plan File",
                data=bio.getvalue(),
                file_name=f"ABPS_Baikunth_Lesson_Plan_{chapter}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.info("Please enter your Gemini API Key in the left sidebar to run the generation system.")
