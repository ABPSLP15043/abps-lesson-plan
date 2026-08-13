import streamlit as st
import os
import io
import json
import docx
import graphviz
from pypdf import PdfReader
from google import genai
from google.genai import types

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
st.set_page_config(
    page_title="ABPS Baikunth - AI Lesson Plan Generator",
    page_icon="🏫",
    layout="wide"
)

st.title("🏫 The Aditya Birla Public School, Baikunth")
st.caption("Free AI-Powered NCF-SE 2023 Lesson Plan & Mind Map Generator")

# Retrieve GEMINI_API_KEY from Render environment variables automatically
system_api_key = os.getenv("GEMINI_API_KEY", "")

# ---------------------------------------------------------
# HELPER: EXTRACT TEXT FROM PDF (Memory Safe for <= 25 MB)
# ---------------------------------------------------------
def extract_text_from_pdf(uploaded_file, max_pages=40):
    reader = PdfReader(uploaded_file)
    extracted_text = ""
    total_pages = min(len(reader.pages), max_pages)
    for i in range(total_pages):
        text = reader.pages[i].extract_text()
        if text:
            extracted_text += text + "\n"
    return extracted_text

# ---------------------------------------------------------
# HELPER: CALL GEMINI API (GEMINI 1.5 FLASH)
# ---------------------------------------------------------
def generate_ai_lesson_plan(api_key, subject, grade, section, chapter, month, periods, chapter_content):
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    You are an expert curriculum designer for Indian schools following NCF-SE 2023 and NEP 2020 guidelines for NCERT books.
    Analyze the following details and provided chapter content/text to create a HIGHLY SPECIFIC, UNIQUE, and ACCURATE lesson plan.
    
    Subject: {subject}
    Grade/Class: {grade} - {section}
    Chapter Name: {chapter}
    Month: {month}
    Number of Periods: {periods}
    
    --- CHAPTER CONTENT / EXTRACTED TEXT ---
    {chapter_content[:12000]}
    ---------------------------------------
    
    Return a strictly valid JSON object with the following keys:
    {{
      "curriculum_goal": "NCF Curriculum Goal specific to this chapter",
      "relevant_competencies": ["Competency 1", "Competency 2", "Competency 3"],
      "learning_objectives": ["Objective 1 specific to content", "Objective 2", "Objective 3"],
      "expected_learning_outcomes": ["Outcome 1", "Outcome 2", "Outcome 3"],
      "formulas_and_equations": ["Specific formula/law/definition 1 from chapter", "Formula/law 2"],
      "teaching_methodology": ["Method 1", "Method 2"],
      "teaching_aids": ["Aid 1", "Aid 2"],
      "art_integration": ["Art integration task specific to chapter"],
      "previous_knowledge": ["Prerequisite 1", "Prerequisite 2"],
      "innovative_techniques": ["Technique 1", "Technique 2"],
      "content_points": [
        {{"section": "Section 1 Title", "topics": ["Subtopic A", "Subtopic B"]}},
        {{"section": "Section 2 Title", "topics": ["Subtopic C", "Subtopic D"]}}
      ],
      "projects_experiential": ["Experiential activity for this chapter"],
      "skills_acquired": ["Skill 1", "Skill 2"],
      "values_inculcated": ["Value 1", "Value 2"],
      "multiple_assessment": {{
        "oral_questions": ["Chapter Question 1", "Chapter Question 2"],
        "worksheet": ["Task 1", "Task 2"],
        "practical": ["Practical task / activity"],
        "exit_ticket": ["Exit ticket prompt"]
      }},
      "class_work": ["Classwork task 1", "Classwork task 2"],
      "home_work": ["Homework task 1", "Homework task 2"],
      "remedial_measures": {{
        "slow_learners": ["Support strategy 1"],
        "advanced_learners": ["Enrichment task 1"]
      }},
      "resources": {{
        "books": ["NCERT Class {grade} {subject}"],
        "websites": ["DIKSHA Portal"],
        "videos": ["Relevant topic educational video"]
      }}
    }}
    Do not add markdown backticks like ```json. Return raw JSON string only.
    """

    response = client.models.generate_content(
        model='gemini-1.5-flash',
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json"
        )
    )
    
    return json.loads(response.text)

# ---------------------------------------------------------
# WORD GENERATOR (.DOCX)
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
    for s in doc.sections:
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
    
    if "formulas_and_equations" in plan_data and plan_data["formulas_and_equations"]:
        add_section_table("Key Formulas, Laws &\nEquations", plan_data.get("formulas_and_equations", []))

    add_section_table("Teaching Methodology", plan_data.get("teaching_methodology", []))
    
    aids_text = "Teaching Aids:\n" + "\n".join([f"• {a}" for a in plan_data.get("teaching_aids", [])])
    aids_text += "\n\nArt Integration:\n"
