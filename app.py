import streamlit as st
import google.generativeai as genai
from docx import Document
from io import BytesIO

# PAGE SETTINGS
st.set_page_config(
    page_title="ABPS Lesson Plan Generator",
    page_icon="📘",
    layout="wide"
)

# HEADER
st.markdown("""
<h1 style='text-align:center; color:#003366;'>
The Aditya Birla Public School, Baikunth
</h1>
<h3 style='text-align:center; color:#F7931E;'>
AI Powered NCF 2023 Lesson Plan Generator
</h3>
""", unsafe_allow_html=True)

# SIDEBAR
st.sidebar.header("Gemini API Configuration")

api_key = st.sidebar.text_input(
    "Enter Gemini API Key",
    type="password"
)

# MAIN APP
if api_key:

    genai.configure(api_key=api_key)

    model = genai.GenerativeModel("models/gemini-1.5-flash")

    st.subheader("Teacher Input Section")

    col1, col2 = st.columns(2)

    with col1:
        subject = st.text_input("Subject")
        grade = st.text_input("Class")
        chapter = st.text_input("Chapter / Topic")

    with col2:
        periods = st.text_input("Number of Periods")
        month = st.text_input("Month")
        teacher_name = st.text_input("Teacher Name")

    notes = st.text_area(
        "Paste Rough Notes / Teaching Ideas",
        height=200
    )

    if st.button("Generate ABPS Lesson Plan"):

        with st.spinner("Generating Lesson Plan..."):

            prompt = f"""
            You are an expert CBSE curriculum planner.

            Generate a complete lesson plan for:

            Subject: {subject}
            Class: {grade}
            Chapter: {chapter}
            Month: {month}
            Number of Periods: {periods}

            Teaching Notes:
            {notes}

            Generate detailed content under these headings:

            1. NCF 2023 Curricular Goals & Competencies
            2. Learning Objectives
            3. Expected Learning Outcomes
            4. Teaching Methodology
            5. Skills & Competencies Developed
            6. Teaching Aids
            7. Innovative Techniques
            8. Experiential Learning Activities
            9. Classroom Discussion Questions
            10. Multiple Assessment Techniques
            11. Class Work
            12. Homework
            13. Remedial Measures
            14. Values Inculcated
            15. Suggested Resources

            Use professional school language suitable for CBSE schools.
            """

            response = model.generate_content(prompt)

            lesson_text = response.text

            st.success("Lesson Plan Generated Successfully")

            st.markdown(lesson_text)

            # CREATE WORD DOCUMENT
            doc = Document()

            doc.add_heading(
                "The Aditya Birla Public School, Baikunth",
                level=1
            )

            doc.add_heading(
                "NCF 2023 Integrated Lesson Plan",
                level=2
            )

            doc.add_paragraph(f"Teacher Name: {teacher_name}")
            doc.add_paragraph(f"Subject: {subject}")
            doc.add_paragraph(f"Class: {grade}")
            doc.add_paragraph(f"Chapter: {chapter}")
            doc.add_paragraph(f"Month: {month}")
            doc.add_paragraph(f"No. of Periods: {periods}")

            doc.add_paragraph(lesson_text)

            buffer = BytesIO()

            doc.save(buffer)

            st.download_button(
                label="📥 Download Lesson Plan DOCX",
                data=buffer.getvalue(),
                file_name=f"{chapter}_Lesson_Plan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

else:

    st.info("Please Enter Gemini API Key in Sidebar")
