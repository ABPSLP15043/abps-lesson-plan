import streamlit as st
import io
import docx
import graphviz
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
st.set_page_config(page_title="ABPS Baikunth - Advanced Lesson Plan Generator", layout="wide")

st.title("🏫 The Aditya Birla Public School, Baikunth")
st.caption("Integrated NCF-SE 2023 | NEP 2020 Multi-Subject Lesson Plan & Mind Map Generator")

# ---------------------------------------------------------
# DYNAMIC MULTI-SUBJECT GENERATION ENGINE (WITH FORMULAS)
# ---------------------------------------------------------
def build_comprehensive_plan(subject, grade, section, chapter, month, periods):
    
    # -----------------------------------------------------
    # 1. MATHEMATICS (Formulas, Theorems & Proofs)
    # -----------------------------------------------------
    if subject == "MATHEMATICS":
        return {
            "curriculum_goal": f"CG-MATH-1: Develops logical reasoning, abstract thinking, mathematical modeling, and problem-solving skills through quantitative analysis.",
            "relevant_competencies": [
                f"C-M1: Applies core mathematical theorems, formulas, and algebraic identities related to '{chapter}'.",
                "C-M2: Demonstrates accuracy in step-by-step calculations, graphical representations, and geometric proofs.",
                "C-M3: Solves real-world contextual problems using mathematical modeling and quantitative logic."
            ],
            "learning_objectives": [
                f"Understand the theoretical derivation and practical application of formulas in '{chapter}'.",
                "Master step-by-step algorithms and computational techniques.",
                "Apply mathematical theorems to solve standard and exemplar numerical problems.",
                "Develop error-analysis and verification strategies for calculated results."
            ],
            "expected_learning_outcomes": [
                f"Students will accurately state and apply formulas associated with '{chapter}'.",
                "Solve multi-step numerical exercises and geometric/algebraic proofs independently.",
                "Construct accurate graphs, geometric constructions, or statistical tables."
            ],
            "formulas_and_equations": [
                f"Core Formula 1 ({chapter}): Area / Perimeter / Algebraic Identity: A = ½ × b × h | (a+b)² = a² + 2ab + b²",
                f"Core Formula 2 ({chapter}): Standard Equation / Theorem: Pyhtagoras Theorem: a² + b² = c²",
                "Conversion & Constant Ratios: π ≈ 22/7 or 3.14159, Percentage / Ratio Formulae",
                "Verification Identity: L.H.S = R.H.S verification protocols"
            ],
            "teaching_methodology": [
                "Inductive-Deductive Problem Solving", "Step-by-Step Board Work",
                "Guided Exemplar Practice", "Math Lab Hands-On Activities"
            ],
            "teaching_aids": [
                "GeoGebra / Interactive Geometry Software", "Math Lab Geometry Kits & Graph Boards", "Formula Reference Flashcards"
            ],
            "art_integration": [
                f"Symmetry & Tessellation Art Project based on mathematical shapes in '{chapter}'", "Geometrical Pattern Posters"
            ],
            "previous_knowledge": [
                "Prerequisite operations (addition, multiplication, basic algebraic simplification).",
                "Understanding of basic geometric shapes and foundational formula applications."
            ],
            "innovative_techniques": [
                "GeoGebra Visualizer", "Step-by-step Formula Wheel", "Peer Error-Analysis Drills"
            ],
            "content_points": [
                {"section": "Theoretical Concepts & Formulas", "topics": [f"Derivation of core formulas and key definitions in '{chapter}'"]},
                {"section": "Guided Exemplar Problems", "topics": ["Step-by-step walkthrough of NCERT textbook sample problems"]},
                {"section": "Independent Practice & Graphs", "topics": ["Solving exercise numericals, graphical plots, and word problems"]}
            ],
            "projects_experiential": [
                f"Construct a real-world mathematical model or scale drawing demonstrating '{chapter}'."
            ],
            "skills_acquired": [
                "Computational Speed", "Logical Deduction", "Spatial Visualization", "Error Checking"
            ],
            "values_inculcated": [
                "Precision & Accuracy", "Persistence in Problem Solving", "Systematic Thinking"
            ],
            "multiple_assessment": {
                "oral_questions": [f"State the main formula used in '{chapter}'.", "What is the condition for formula applicability?"],
                "worksheet": ["Formula substitution numericals", "NCERT Exemplar short answer questions"],
                "practical": ["Math Lab Activity & Graphical Verification Worksheet"],
                "exit_ticket": ["Write down the core formula learned today with its SI units or variable definitions."]
            },
            "class_work": ["Solving Exercise Questions in notebook", "Board numerical practice"],
            "home_work": [f"Complete textbook exercise questions for '{chapter}' and create a Formula Revision Chart."],
            "remedial_measures": {
                "slow_learners": ["Formula memory flashcards", "Step-by-step guided calculation templates"],
                "advanced_learners": ["CBSE High-Order Thinking Skills (HOTS) and Olympiad problems"]
            },
            "resources": {
                "books": [f"NCERT Class {grade} Mathematics", "CBSE Exemplar Problems"],
                "websites": ["DIKSHA Math Modules", "GeoGebra Online"],
                "videos": ["Khan Academy Mathematics", "NCERT Virtual Lab Videos"]
            }
        }

    # -----------------------------------------------------
    # 2. PHYSICS (Laws, Numerical Formulas & SI Units)
    # -----------------------------------------------------
    elif subject in ["PHYSICS", "SCIENCE"] and ("FORCE" in chapter.upper() or "MOTION" in chapter.upper() or "LIGHT" in chapter.upper() or "ELECTRICITY" in chapter.upper() or "ENERGY" in chapter.upper() or subject == "PHYSICS"):
        return {
            "curriculum_goal": f"CG-PHY-1: Explores natural physical phenomena through systematic inquiry, quantitative laws, mathematical modeling, and experimental physics.",
            "relevant_competencies": [
                f"C-P1: Understands and applies foundational laws of physics, equations of motion, and field principles for '{chapter}'.",
                "C-P2: Derives numerical formulas, performs vector/scalar computations, and verifies physical constants.",
                "C-P3: Conducts laboratory experiments, records observation tables, and analyzes physical graph trends."
            ],
            "learning_objectives": [
                f"State the foundational physical laws and principles governing '{chapter}'.",
                "Understand the mathematical derivation of key equations and SI unit dimensions.",
                "Solve conceptual problems and numerical exercises using standard units.",
                "Analyze experimental graphs (e.g., v-t graphs, V-I graphs, ray diagrams)."
            ],
            "expected_learning_outcomes": [
                f"Students will accurately state laws and apply equations associated with '{chapter}'.",
                "Convert physical quantities into standard SI units and calculate numerical outputs.",
                "Construct accurate ray diagrams, circuit diagrams, or motion graphs."
            ],
            "formulas_and_equations": [
                f"Kinematics / Dynamics: v = u + at | s = ut + ½at² | v² - u² = 2as | F = m·a",
                f"Work & Energy / Power: W = F·d·cos(θ) | P = W/t | E = mgh | KE = ½mv²",
                f"Electricity / Waves: V = I·R | P = V·I | f = 1/T | v = f·λ",
                "Standard Constants & SI Units: g = 9.8 m/s², SI Units: Newton (N), Joule (J), Watt (W), Volt (V), Ohm (Ω)"
            ],
            "teaching_methodology": [
                "Demonstration & Laboratory Experimentation", "Inquiry-Based Learning",
                "Mathematical Problem-Solving Sessions", "Interactive Physics Simulations"
            ],
            "teaching_aids": [
                "PhET Interactive Physics Simulations", "Physics Lab Apparatus & Multimeters", "Circuit Boards & Optical Benches"
            ],
            "art_integration": [
                f"Infographic Poster on 'Physics Laws in Daily Life' for '{chapter}'", "Ray/Circuit Diagram Art"
            ],
            "previous_knowledge": [
                "Understanding of basic measurement units (meters, seconds, kilograms).",
                "Basic algebraic rearrangement of equations."
            ],
            "innovative_techniques": [
                "PhET Digital Simulations", "Slow-Motion Video Analysis of Motion", "Exit Ticket Quiz"
            ],
            "content_points": [
                {"section": "Laws & Principles", "topics": [f"Statement of fundamental physical laws and definitions in '{chapter}'"]},
                {"section": "Mathematical Derivations & Formulas", "topics": ["Deriving equations, dimensional analysis, and SI unit units"]},
                {"section": "Experimental Setup & Numericals", "topics": ["Lab demonstration, graph plotting, and numerical walkthroughs"]}
            ],
            "projects_experiential": [
                f"Build a working physics toy model or working circuit demonstration for '{chapter}'."
            ],
            "skills_acquired": [
                "Scientific Inquiry", "Graph Interpretation", "Numerical Precision", "Lab Safety"
            ],
            "values_inculcated": [
                "Objectivity", "Curiosity about Natural Laws", "Safety Awareness"
            ],
            "multiple_assessment": {
                "oral_questions": [f"State the physical law studied in '{chapter}'.", "What is the SI unit of key variables?"],
                "worksheet": ["Formula substitution and unit conversion numericals", "Ray/Circuit diagram completion"],
                "practical": ["Physics Lab Experiment & Observation Table Recording"],
                "exit_ticket": ["Write down the core equation learned today with SI units."]
            },
            "class_work": ["Solving physics numericals in notebook", "Lab experiment recording"],
            "home_work": [f"Solve numerical exercises for '{chapter}' and construct a Law & Formula Summary Table."],
            "remedial_measures": {
                "slow_learners": ["Formula triangle visual guides (e.g., V=I·R triangle)", "Units matching worksheets"],
                "advanced_learners": ["CBSE HOTS questions and multi-concept combined numericals"]
            },
            "resources": {
                "books": [f"NCERT Class {grade} Physics/Science", "CBSE Lab Manual"],
                "websites": ["PhET Interactive Simulations", "DIKSHA Portal"],
                "videos": ["NCERT Physics Lab Demonstrations", "Khan Academy Physics"]
            }
        }

    # -----------------------------------------------------
    # 3. CHEMISTRY (Chemical Reactions, Formulas & Equations)
    # -----------------------------------------------------
    elif subject in ["CHEMISTRY", "SCIENCE"] and ("CHEMICAL" in chapter.upper() or "ACID" in chapter.upper() or "ATOM" in chapter.upper() or "ORGANIC" in chapter.upper() or "PERIODIC" in chapter.upper() or subject == "CHEMISTRY"):
        return {
            "curriculum_goal": f"CG-CHEM-1: Investigates atomic structure, chemical transformations, reaction mechanisms, and stoichiometry through scientific experimentation.",
            "relevant_competencies": [
                f"C-C1: Balances chemical equations, predicts reaction products, and understands stoichiometric ratios in '{chapter}'.",
                "C-C2: Applies periodic trends, chemical bonding concepts, and molecular structure representations.",
                "C-C3: Conducts lab experiments safely, observing color changes, gas evolution, and precipitate formation."
            ],
            "learning_objectives": [
                f"Understand atomic/molecular structures and chemical formulas relevant to '{chapter}'.",
                "Master balancing chemical equations and applying the Law of Conservation of Mass.",
                "Identify types of reactions (Combination, Decomposition, Redox, Substitution, Acid-Base).",
                "Perform stoichiometric mole concept calculations."
            ],
            "expected_learning_outcomes": [
                f"Students will write balanced chemical equations and formulas for '{chapter}'.",
                "Identify chemical reaction indicators (pH, precipitate, gas evolution, heat transfer).",
                "Calculate molar masses, mole ratios, or concentration terms accurately."
            ],
            "formulas_and_equations": [
                f"General Reaction Representation: Reactants (A + B) ──► Products (C + D)",
                f"Mole Concept Formulas: Moles (n) = Given Mass (m) / Molar Mass (M) | N = n × 6.022 × 10²³",
                f"pH & Concentration Equations: pH = -log[H⁺] | Molarity (M) = Moles of Solute / Volume of Solution (L)",
                "Key Balanced Reactions: CaCO₃ ──► CaO + CO₂ ↑ | 2H₂ + O₂ ──► 2H₂O | Acid + Base ──► Salt + Water"
            ],
            "teaching_methodology": [
                "Lab Experimentation & Demonstration", "Equation Balancing Drills",
                "Molecular Model Building", "Inquiry-Based Chemical Observation"
            ],
            "teaching_aids": [
                "Chemistry Lab Reagents & Test Tubes", "3D Molecular Model Kits", "Periodic Table Visual Charts"
            ],
            "art_integration": [
                f"Periodic Table Art / Reaction Flowchart Wall Chart for '{chapter}'", "Color-coded pH Scale Art"
            ],
            "previous_knowledge": [
                "Symbols of elements and basic valency rules.",
                "Distinction between physical and chemical changes."
            ],
            "innovative_techniques": [
                "3D Molecular AR/VR Apps", "Equation Balancing Card Game", "Interactive Periodic Table"
            ],
            "content_points": [
                {"section": "Chemical Concepts & Valency", "topics": [f"Chemical symbols, valency, and molecular formulas for '{chapter}'"]},
                {"section": "Balanced Equations & Mechanisms", "topics": ["Step-by-step balancing of equations and reaction types"]},
                {"section": "Lab Observation & Stoichiometry", "topics": ["Experimental observations, heat changes, and mole calculations"]}
            ],
            "projects_experiential": [
                f"Perform a natural indicator or safe home-chemistry reaction demonstration for '{chapter}'."
            ],
            "skills_acquired": [
                "Chemical Equation Balancing", "Lab Reagent Handling", "Observation & Inferences", "Precise Measurement"
            ],
            "values_inculcated": [
                "Laboratory Safety", "Environmental Green Chemistry", "Scientific Accuracy"
            ],
            "multiple_assessment": {
                "oral_questions": [f"Name the key chemical compound in '{chapter}'.", "How do you test for gas evolution?"],
                "worksheet": ["Equation balancing exercises", "Mole concept numericals"],
                "practical": ["Chemical Reaction Test Tube Experiment & Observation Recording"],
                "exit_ticket": ["Write down one balanced chemical equation learned today."]
            },
            "class_work": ["Balancing chemical equations on board", "Notebook reaction charts"],
            "home_work": [f"Complete textbook exercise equations for '{chapter}' and write a Chemical Formula Table."],
            "remedial_measures": {
                "slow_learners": ["Valency cross-multiplication method flashcards", "Pre-balanced equation templates"],
                "advanced_learners": ["Redox reaction balancing (ion-electron method) and stoichiometry challenges"]
            },
            "resources": {
                "books": [f"NCERT Class {grade} Chemistry/Science", "CBSE Science Manual"],
                "websites": ["Royal Society of Chemistry", "DIKSHA Portal"],
                "videos": ["NCERT Chemistry Practical Videos", "Khan Academy Chemistry"]
            }
        }

    # -----------------------------------------------------
    # 4. MUSIC, ART & DANCE
    # -----------------------------------------------------
    elif subject == "MUSIC":
        return {
            "curriculum_goal": f"CG-ART-1: Develops aesthetic sensibility, rhythmic awareness, and cultural appreciation through vocal and instrumental music.",
            "relevant_competencies": [
                f"C-1.1: Demonstrates mastery of basic Swaras, Taal, and rhythmic patterns in '{chapter}'.",
                "C-1.2: Expresses emotional and artistic nuances through vocal/instrumental practice.",
                "C-1.3: Understands historical and cultural context of Indian classical/folk music styles."
            ],
            "learning_objectives": [
                f"Identify and sing/play the core Swaras/Taal associated with '{chapter}'.",
                "Understand the tempo (Laya) and rhythmic structure (Matra/Tali/Khali).",
                "Develop ear training, pitch accuracy, and voice modulation."
            ],
            "expected_learning_outcomes": [
                f"Students will perform '{chapter}' with accurate pitch (Swar) and rhythm (Taal).",
                "Identify and count Matras using hand beats (Tali/Khali)."
            ],
            "formulas_and_equations": [
                "Rhythmic Structure (Taal Notation): Total Matras = Tali Beats + Khali Beats",
                "Scale Ratio (Saptak): Mandra (Lower) : Madhya (Middle) : Tara (Higher) Saptak",
                "Tempo Ratios (Laya): Vilambit (1x) ──► Madhya (2x) ──► Drut (4x)"
            ],
            "teaching_methodology": ["Demonstration & Imitation Method (Riyaz)", "Guided Vocal Practice", "Group Choral Singing"],
            "teaching_aids": ["Harmonium / Tanpura / Tabla", "Audio Recordings & Metronome", "Notation Charts"],
            "art_integration": [f"Musical Rendition & Rhythmic Ensemble on '{chapter}'"],
            "previous_knowledge": ["Familiarity with basic Saptak (Swara) and simple rhythmic patterns."],
            "innovative_techniques": ["Digital Metronome Practice", "Audio Self-Recording Analysis"],
            "content_points": [
                {"section": "Introduction & Aaroh-Avaroh", "topics": [f"Basic scale/Thaat introduction and Swara alignment for '{chapter}'"]},
                {"section": "Composition & Bandish Practice", "topics": ["Line-by-line vocal demonstration, lyrics meaning, and Riyaz"]}
            ],
            "projects_experiential": [f"Prepare a small group choir or solo musical presentation on '{chapter}'."],
            "skills_acquired": ["Aural Perception", "Rhythmic Precision", "Vocal Control"],
            "values_inculcated": ["Cultural Heritage Respect", "Patience & Discipline"],
            "multiple_assessment": {
                "oral_questions": [f"Name the primary Taal or Swaras used in '{chapter}'?"],
                "worksheet": ["Identify missing Swaras in notation"],
                "practical": ["Solo/Group Vocal Performance Test"],
                "exit_ticket": ["Demonstrate 1 Avartan of the assigned Taal with hand beats."]
            },
            "class_work": ["Notation writing in notebook", "Group Riyaz"],
            "home_work": [f"Practice 15 minutes daily Riyaz of '{chapter}'."],
            "remedial_measures": {
                "slow_learners": ["Individual Swara tuning assistance"],
                "advanced_learners": ["Aalap & Taam practice"]
            },
            "resources": {
                "books": ["CBSE Music Curriculum Guide"],
                "websites": ["NCERT e-Music Resources"],
                "videos": ["Classical Vocal Demonstrations"]
            }
        }

    # -----------------------------------------------------
    # 5. HINDI & SANSKRIT
    # -----------------------------------------------------
    elif subject == "HINDI":
        return {
            "curriculum_goal": f"लक्ष्य-3: {subject} शिक्षण द्वारा भाषा-कौशल, मौलिक चिंतन, रचनात्मकता एवं साहित्यिक समझ का विकास करना।",
            "relevant_competencies": [
                f"दक्षता-3.1: '{chapter}' पाठ का स्पष्ट उच्चारण, पठन एवं भावग्रहण करना।",
                "दक्षता-3.2: व्याकरणिक नियमों एवं भाषिक संरचनाओं का सही प्रयोग करना।"
            ],
            "learning_objectives": [
                f"पाठ '{chapter}' के मुख्य भाव, विचार एवं केंद्रीय विषय को समझना।",
                "कठिन शब्दों के अर्थ एवं व्याकरणिक तत्वों का ज्ञान प्राप्त करना।"
            ],
            "expected_learning_outcomes": [
                f"विद्यार्थी '{chapter}' के प्रश्नों के उत्तर स्पष्ट रूप से देने में सक्षम होंगे।",
                "नए शब्दावली का वाक्यों में प्रयोग कर सकेंगे।"
            ],
            "formulas_and_equations": [
                "व्याकरणिक नियम: संधि = वर्ण + वर्ण | समास = पद + पद",
                "वाक्य संरचना: कर्ता + कर्म + क्रिया"
            ],
            "teaching_methodology": ["वाचन विधि", "व्याख्यान विधि", "प्रश्न उत्तर विधि"],
            "teaching_aids": ["स्मार्ट बोर्ड / वीडियो", "चित्र एवं फ्लैश कार्ड"],
            "art_integration": [f"'{chapter}' पर आधारित माइंड मैप या पोस्टर निर्माण"],
            "previous_knowledge": ["पाठ से संबंधित बुनियादी शब्दावली की समझ की जाँच।"],
            "innovative_techniques": ["माइंड मैपिंग (Mind Mapping)", "डिजिटल वाचन"],
            "content_points": [
                {"section": "भूमिका एवं परिचय", "topics": [f"लेखक/कवि का परिचय एवं '{chapter}' का मुख्य सार"]},
                {"section": "व्याख्या एवं भावार्थ", "topics": ["पाठ का वाचन, कठिन शब्दार्थ एवं व्याख्या"]}
            ],
            "projects_experiential": [f"'{chapter}' के विषय पर एक लघु अनुच्छेद या कहानी लिखें।"],
            "skills_acquired": ["श्रवण", "वाचन", "पठन", "लेखन"],
            "values_inculcated": ["नैतिक मूल्य", "संवेदनशीलता"],
            "multiple_assessment": {
                "oral_questions": [f"'{chapter}' का मुख्य संदेश क्या है?"],
                "worksheet": ["बहुविकल्पीय प्रश्न (MCQs)"],
                "practical": ["सस्वर वाचन अभ्यास"],
                "exit_ticket": ["आज सीखे गए नए शब्द से एक वाक्य बनाएं।"]
            },
            "class_work": ["पाठ्यपुस्तक के प्रश्नोत्तर लिखना"],
            "home_work": [f"'{chapter}' का सारांश अपने शब्दों में लिखें।"],
            "remedial_measures": {
                "slow_learners": ["सस्वर वाचन अभ्यास"],
                "advanced_learners": ["मौलिक रचनात्मक लेखन"]
            },
            "resources": {
                "books": [f"एनसीईआरटी कक्षा {grade} हिंदी पाठ्यपुस्तक"],
                "websites": ["दीक्षा पोर्टल (DIKSHA Portal)"],
                "videos": ["शैक्षणिक वीडियो"]
            }
        }

    # -----------------------------------------------------
    # 6. GENERAL ACADEMIC SUBJECTS (Default)
    # -----------------------------------------------------
    else:
        return {
            "curriculum_goal": f"CG-3: Explores core domain knowledge in {subject} through structured scientific inquiry, analytical problem solving, and real-world application.",
            "relevant_competencies": [
                f"C-3.1: Analyzes fundamental theories and functional mechanisms of '{chapter}'.",
                "C-3.2: Conducts structured problem-solving, data interpretation, and practical experimentation.",
                "C-3.3: Evaluates real-world impact, environmental sustainability, and ethical implications."
            ],
            "learning_objectives": [
                f"Define and explain core concepts, scientific laws, or theoretical models of '{chapter}'.",
                "Apply analytical formulas, textual analysis, or logical frameworks to solve academic problems.",
                "Analyze cause-and-effect relationships in practical contexts."
            ],
            "expected_learning_outcomes": [
                f"Students will accurately articulate definitions, mechanisms, and key terms for '{chapter}'.",
                "Solve conceptual problems and numerical/textual exercises independently."
            ],
            "formulas_and_equations": [
                f"Key Formula / Analytical Relation ({chapter}): Core Input ──► Functional Process ──► Output / Result",
                "Standard Equations & Variable Definitions for the chapter domain"
            ],
            "teaching_methodology": ["Inquiry-Based Learning", "Problem-Solving Method", "Collaborative Peer Discussion"],
            "teaching_aids": ["Smart Board / Interactive PPTs", "Worksheets & Mind Maps"],
            "art_integration": [f"Concept Mapping & Graphical Flowchart of '{chapter}'"],
            "previous_knowledge": ["What foundational concepts from earlier grades relate to this chapter?"],
            "innovative_techniques": ["Blended Learning Modules", "Exit Tickets & Quizzes"],
            "content_points": [
                {"section": "Theoretical Framework", "topics": [f"Introduction, definitions, and scope of '{chapter}'"]},
                {"section": "Mechanism & Analysis", "topics": ["Step-by-step structural analysis and working principles"]}
            ],
            "projects_experiential": [f"Conduct a mini-research survey investigating '{chapter}' in your local context."],
            "skills_acquired": ["Critical Thinking", "Analytical Reasoning", "Data Interpretation"],
            "values_inculcated": ["Scientific Temper", "Curiosity", "Intellectual Honesty"],
            "multiple_assessment": {
                "oral_questions": [f"Explain the main principle behind '{chapter}'."],
                "worksheet": ["Short Answer & Conceptual Questions"],
                "practical": ["Case Study / Observation Table Recording"],
                "exit_ticket": ["Summarize today's core concept in 2 sentences."]
            },
            "class_work": ["Concept notes drafting", "Solving exercise questions"],
            "home_work": [f"Complete review questions for '{chapter}' and construct a summary mind map."],
            "remedial_measures": {
                "slow_learners": ["Peer tutoring and visual flashcards"],
                "advanced_learners": ["Advanced research assignments"]
            },
            "resources": {
                "books": [f"NCERT Class {grade} {subject} Textbook"],
                "websites": ["DIKSHA Portal"],
                "videos": ["NCERT Educational Videos"]
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
    
    if "formulas_and_equations" in plan_data:
        add_section_table("Key Formulas, Laws &\nChemical Equations", plan_data.get("formulas_and_equations", []))

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
    ]

    if "formulas_and_equations" in plan_data:
        details_data.append([Paragraph("<b>Key Formulas & Equations</b>", body_style), Paragraph(make_bullet_list(plan_data.get("formulas_and_equations",[])), body_style)])

    details_data.extend([
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
    ])
    
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
    "MUSIC", "ART & CRAFT", "DANCE",
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
chapter = st.sidebar.text_input("Chapter / Topic", "Force and Laws of Motion")
periods = st.sidebar.number_input("No. of Periods", min_value=1, max_value=25, value=8)

# ---------------------------------------------------------
# GENERATION ENGINE
# ---------------------------------------------------------
if st.sidebar.button("✨ Generate Official Lesson Plan", type="primary"):
    with st.spinner("Building subject-specific ABPS lesson plan..."):
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

    # =========================================================
    # 📐 FORMULAS, LAWS & CHEMICAL EQUATIONS SECTION
    # =========================================================
    if "formulas_and_equations" in data:
        st.divider()
        st.subheader("📐 Key Formulas, Laws & Chemical Equations")
        for f in data["formulas_and_equations"]:
            st.info(f"⚡ {f}")

    # =========================================================
    # 🧠 AUTOMATIC VISUAL MIND MAP GENERATOR
    # =========================================================
    st.divider()
    st.subheader(f"🧠 Visual Mind Map: {meta['chapter']}")

    dot = graphviz.Digraph(comment=meta['chapter'])
    dot.attr(rankdir='LR', size='8,5', node_style='filled', fillcolor='#EBF8FF', color='#2B6CB0', fontname='Helvetica')

    # Central Chapter Node
    dot.node('CENTER', meta['chapter'], shape='box', fillcolor='#2B6CB0', fontcolor='white')

    # Dynamically extract and connect subtopics from content points
    subtopic_idx = 0
    for item in data.get('content_points', []):
        section_name = item.get('section', 'Core Concepts')
        section_id = f"SEC_{subtopic_idx}"
        
        # Add Section Node
        dot.node(section_id, section_name, shape='ellipse', fillcolor='#E2E8F0')
        dot.edge('CENTER', section_id)
        
        # Add Subtopic Nodes under each section
        for topic in item.get('topics', []):
            topic_id = f"TOP_{subtopic_idx}"
            dot.node(topic_id, topic, shape='plaintext')
            dot.edge(section_id, topic_id)
            subtopic_idx += 1

    st.graphviz_chart(dot, use_container_width=True)
    # =========================================================

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
