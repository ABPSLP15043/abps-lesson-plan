import streamlit as st

st.set_page_config(page_title="ABPS Lesson Plan Generator")

st.title("ABPS Lesson Plan Generator")

subject = st.text_input("Enter Subject")
grade = st.text_input("Enter Class")
chapter = st.text_input("Enter Chapter")

notes = st.text_area("Enter Teaching Notes")

if st.button("Generate Lesson Plan"):
    st.success("Lesson Plan Generated Successfully")

    st.write("### Lesson Plan")
    st.write(f"Subject: {subject}")
    st.write(f"Class: {grade}")
    st.write(f"Chapter: {chapter}")
    st.write(f"Teaching Notes: {notes}")

    st.write("### NCF 2023 Curriculum Goals")
    st.write("- Competency Based Learning")
    st.write("- Experiential Learning")
    st.write("- Student Centric Pedagogy")
