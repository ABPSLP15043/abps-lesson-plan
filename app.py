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
st.set_page_config(page_title="ABPS Baikunth - Advanced Lesson Plan Generator", layout="wide")

st.title("🏫 The Aditya Birla Public School, Baikunth")
st.caption("Integrated NCF-SE 2023 | NEP 2020 Multi-Subject Lesson Plan Generator")

# ---------------------------------------------------------
# DYNAMIC MULTI-SUBJECT GENERATION ENGINE
# ---------------------------------------------------------
def build_comprehensive_plan(subject, grade, section, chapter, month, periods):
    # --- MUSIC ---
    if subject == "MUSIC":
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
                "Develop ear training, pitch accuracy, and voice modulation.",
                "Appreciate the cultural significance and composer history of the composition."
            ],
            "expected_learning_outcomes": [
                f"Students will perform '{chapter}' with accurate pitch (Swar) and rhythm (Taal).",
                "Identify and count Matras using hand beats (Tali/Khali).",
                "Demonstrate improved vocal flexibility or instrumental coordination."
            ],
            "teaching_methodology": [
                "Demonstration & Imitation Method (Riyaz)", "Guided Vocal Practice",
                "Group Choral Singing", "Rhythmic Clapping & Beat Counting"
            ],
            "teaching_aids": [
                "Harmonium / Tanpura / Tabla", "Audio Recordings & Metronome", "Notation Charts"
            ],
            "art_integration": [
                f"Musical Rendition & Rhythmic Ensemble on '{chapter}'", "Lyric Illustration Poster"
            ],
            "previous_knowledge": [
                "Are students familiar with basic Saptak (Swara) and simple rhythmic patterns?",
                "Prior experience in community or assembly singing."
            ],
            "innovative_techniques": [
                "Digital Metronome Practice", "Audio Self-Recording Analysis", "Interactive Music Quiz"
            ],
            "content_points": [
                {"section": "Introduction & Aaroh-Avaroh", "topics": [f"Basic scale/Thaat introduction and Swara alignment for '{chapter}'"]},
                {"section": "Composition & Bandish Practice", "topics": ["Line-by-line vocal demonstration, lyrics meaning, and Riyaz"]},
                {"section": "Taal & Rhythm Sync", "topics": ["Hand beats, Matra alignment, and tempo variation (Vilambit/Drut)"]}
            ],
            "projects_experiential": [
                f"Prepare a small group choir or solo musical presentation on '{chapter}'."
            ],
            "skills_acquired": [
                "Aural Perception", "Rhythmic Precision", "Vocal Control", "Stage Performance"
            ],
            "values_inculcated": [
                "Cultural Heritage Respect", "Patience & Discipline", "Team Harmony"
            ],
            "multiple_assessment": {
                "oral_questions": [f"Name the primary Taal or Swaras used in '{chapter}'?", "Define Aaroh and Avaroh."],
                "worksheet": ["Identify missing Swaras in notation", "Match Taal with its total Matras"],
                "practical": ["Solo/Group Vocal & Instrumental Performance Test"],
                "exit_ticket": ["Demonstrate 1 Avartan of the assigned Taal with hand beats."]
            },
            "class_work": ["Notation writing in notebook", "Group Riyaz and vocal warm-ups"],
            "home_work": [f"Practice 15 minutes daily Riyaz of '{chapter}'.", "Memorize lyrics and notation."],
            "remedial_measures": {
                "slow_learners": ["Individual Swara tuning assistance", "Slower tempo practice"],
                "advanced_learners": ["Aalap & Taam practice", "Harmonium accompaniment training"]
            },
            "resources": {
                "books": ["CBSE Music Curriculum Guide", "Sangeet Visharad Basics"],
                "websites": ["NCERT e-Music Resources", "SWAYAM Portal"],
                "videos": ["Classical Vocal Demonstrations", "Interactive Tabla Beat Loops"]
            }
        }

    # --- ART & CRAFT ---
    elif subject == "ART & CRAFT":
        return {
            "curriculum_goal": f"CG-ART-2: Enhances visual literacy, creative expression, fine motor skills, and spatial awareness through hands-on art exploration.",
            "relevant_competencies": [
                f"C-2.1: Applies fundamental elements of art (line, shape, color, texture) in '{chapter}'.",
                "C-2.2: Demonstrates control over tools, mediums, and crafting techniques.",
                "C-2.3: Evaluates artwork through aesthetic reflection and creative problem-solving."
            ],
            "learning_objectives": [
                f"Master color mixing, shading, or structural crafting techniques for '{chapter}'.",
                "Understand perspective, balance, and proportions in visual composition.",
                "Explore sustainable, upcycled, or local craft materials for artistic creation.",
                "Develop hand-eye coordination and spatial arrangement skills."
            ],
            "expected_learning_outcomes": [
                f"Students will create an original artwork or craft model based on '{chapter}'.",
                "Demonstrate correct usage of shading, blending, or sculpting mediums.",
                "Explain the artistic choices and color harmonies used in their art piece."
            ],
            "teaching_methodology": [
                "Step-by-Step Visual Demonstration", "Studio Practice & Hands-on Creation",
                "Peer Gallery Walk & Critique", "Experimental Medium Exploration"
            ],
            "teaching_aids": [
                "Drawing Sheets, Paints, Brushes & Craft Tools", "Sample Artwork Models", "Visual Reference Charts"
            ],
            "art_integration": [
                f"Cross-curricular poster art linking '{chapter}' with Environmental/Social Studies"
            ],
            "previous_knowledge": [
                "Basic understanding of primary colors and geometric shapes.",
                "Handling safety with scissors, brushes, and glue."
            ],
            "innovative_techniques": [
                "Upcycled Material Crafting", "Digital Art Reference Boards", "Visual Storytelling"
            ],
            "content_points": [
                {"section": "Concept & Color Theory", "topics": [f"Understanding color harmonies and design layouts for '{chapter}'"]},
                {"section": "Hands-on Crafting / Sketching", "topics": ["Guided line drawing, proportion drafting, and texture building"]},
                {"section": "Finishing & Presentation", "topics": ["Outlining, highlights, border presentation, and display setup"]}
            ],
            "projects_experiential": [
                f"Create a 3D craft model or detailed canvas painting based on '{chapter}'."
            ],
            "skills_acquired": [
                "Visual-Spatial Reasoning", "Fine Motor Control", "Creative Innovation", "Aesthetic Judgement"
            ],
            "values_inculcated": [
                "Resourcefulness", "Environmental Upcycling", "Patience", "Appreciation for Local Crafts"
            ],
            "multiple_assessment": {
                "oral_questions": ["What are complementary colors?", "How does texture add depth to artwork?"],
                "worksheet": ["Label color wheel zones", "Match art styles with technique descriptions"],
                "practical": ["Final Artwork Evaluation based on Neatness, Creativity, and Technique"],
                "exit_ticket": ["State one new color blending technique learned today."]
            },
            "class_work": ["Sketchbook drafting", "Color mixing and final artwork composition"],
            "home_work": [f"Collect natural/recyclable materials for the next '{chapter}' craft project."],
            "remedial_measures": {
                "slow_learners": ["Pre-drawn guidelines for proportion help", "Simplified color schemes"],
                "advanced_learners": ["3D perspective rendering", "Mixed-media experimentation"]
            },
            "resources": {
                "books": ["NCERT Art Education Handbook", "CBSE Fine Arts Manual"],
                "websites": ["National Gallery of Modern Art Portal", "DIKSHA Art Modules"],
                "videos": ["Step-by-Step Craft Tutorials", "Famous Artist Documentaries"]
            }
        }

    # --- DANCE ---
    elif subject == "DANCE":
        return {
            "curriculum_goal": f"CG-ART-3: Promotes kinesthetic intelligence, body posture, rhythm synchronization, and expressive narrative skills through dance.",
            "relevant_competencies": [
                f"C-3.1: Executes core footwork (Tatkkar), hand gestures (Mudras), and postures in '{chapter}'.",
                "C-3.2: Synchronizes body movements precisely with rhythm (Taal) and musical cues.",
                "C-3.3: Expresses emotions (Bhava/Rasa) effectively through facial expressions and body language."
            ],
            "learning_objectives": [
                f"Learn and execute the fundamental dance steps and Mudras for '{chapter}'.",
                "Develop physical stamina, flexibility, balance, and spatial alignment.",
                "Understand the narrative or cultural story behind the dance form.",
                "Build group coordination, timing, and stage presence."
            ],
            "expected_learning_outcomes": [
                f"Students will perform the choreographed sequence of '{chapter}' in sync with music.",
                "Identify key Asamyutta/Samyutta Hastamudras used in the routine.",
                "Demonstrate proper posture (Aramandi/Posture) and rhythmic accuracy."
            ],
            "teaching_methodology": [
                "Demonstration & Mirroring Method", "Rhythmic Movement Practice",
                "Group Formation Choreography", "Expression (Abhinaya) Drill"
            ],
            "teaching_aids": [
                "Dance Studio Mirror / Open Space", "Audio Sound System & Rhythm Track", "Mudra Reference Wall Charts"
            ],
            "art_integration": [
                f"Dance Drama Adaptation on '{chapter}' integrated with Literature/History"
            ],
            "previous_knowledge": [
                "Basic body warm-ups and rhythm listening skills.",
                "Familiarity with simple counts (1-2-3-4 / Ta-Dhei-Dhei-Tat)."
            ],
            "innovative_techniques": [
                "Video Movement Analysis", "Choreography Formation Mapping", "Rhythm Clapping Warm-ups"
            ],
            "content_points": [
                {"section": "Body Warm-up & Posture", "topics": [f"Stretching, stamina building, and posture alignment for '{chapter}'"]},
                {"section": "Footwork & Mudra Learning", "topics": ["Step-by-step footwork counts, hand gestures, and body coordination"]},
                {"section": "Choreography & Expressions", "topics": ["Linking steps into full sequence with musical track and Bhava (expressions)"]}
            ],
            "projects_experiential": [
                f"Choreograph a 2-minute thematic group performance on '{chapter}'."
            ],
            "skills_acquired": [
                "Kinesthetic Balance", "Rhythmic Synchronization", "Expressive Communication", "Spatial Awareness"
            ],
            "values_inculcated": [
                "Physical Fitness", "Cultural Respect", "Team Synergy", "Self-Confidence"
            ],
            "multiple_assessment": {
                "oral_questions": ["Name two hand Mudras used today.", "What is Tatkkar in dance?"],
                "worksheet": ["Match Mudras with their meanings", "Sequence the dance steps correctly"],
                "practical": ["Group Dance Performance Assessment on Rhythm, Posture, and Expressions"],
                "exit_ticket": ["Demonstrate one hand gesture learned today with correct name."]
            },
            "class_work": ["Floor practice of dance routine", "Mudra memorization and group formation practice"],
            "home_work": [f"Practice the 8-count footwork sequence of '{chapter}' at home."],
            "remedial_measures": {
                "slow_learners": ["Slow-motion breakdown of steps", "Peer-assisted buddy practice"],
                "advanced_learners": ["Solo improvisation", "Lead choreography assignment"]
            },
            "resources": {
                "books": ["CBSE Dance & Performing Arts Curriculum", "Natya Shastra Overview"],
                "websites": ["Sangeet Natak Akademi Portal", "DIKSHA Performing Arts"],
                "videos": ["Classical & Folk Dance Tutorials", "Cultural Performance Archives"]
            }
        }

    # --- HINDI ---
    elif subject == "HINDI":
        return {
            "curriculum_goal": f"लक्ष्य-3: {subject} शिक्षण द्वारा भाषा-कौशल, मौलिक चिंतन, रचनात्मकता एवं साहित्यिक समझ का विकास करना।",
            "relevant_competencies": [
                f"दक्षता-3.1: '{chapter}' पाठ का स्पष्ट उच्चारण, पठन एवं भावग्रहण करना।",
                "दक्षता-3.2: व्याकरणिक नियमों एवं भाषिक संरचनाओं का सही प्रयोग करना।",
                "दक्षता-3.3: व्यावहारिक जीवन एवं लेखन में नए शब्दों और विचारों का अनुप्रयोग करना।"
            ],
            "learning_objectives": [
                f"पाठ '{chapter}' के मुख्य भाव, विचार एवं केंद्रीय विषय को समझना।",
                "कठिन शब्दों के अर्थ एवं व्याकरणिक तत्वों का ज्ञान प्राप्त करना।",
                "शुद्ध उच्चारण, वाचन एवं अभिव्यक्ति क्षमता को बढ़ाना।"
            ],
            "expected_learning_outcomes": [
                f"विद्यार्थी '{chapter}' के प्रश्नों के उत्तर स्पष्ट रूप से देने में सक्षम होंगे।",
                "नए शब्दावली का वाक्यों में प्रयोग कर सकेंगे।"
            ],
            "teaching_methodology": ["वाचन विधि", "व्याख्यान विधि", "प्रश्न उत्तर विधि", "सामूहिक चर्चा"],
            "teaching_aids": ["स्मार्ट बोर्ड / वीडियो", "चित्र एवं फ्लैश कार्ड", "पाठ्यपुस्तक"],
            "art_integration": [f"'{chapter}' पर आधारित माइंड मैप या पोस्टर निर्माण"],
            "previous_knowledge": ["पाठ से संबंधित बुनियादी शब्दावली की समझ की जाँच।"],
            "innovative_techniques": ["माइंड मैपिंग (Mind Mapping)", "डिजिटल वाचन", "क्विज़"],
            "content_points": [
                {"section": "भूमिका एवं परिचय", "topics": [f"लेखक/कवि का परिचय एवं '{chapter}' का मुख्य सार"]},
                {"section": "व्याख्या एवं भावार्थ", "topics": ["पाठ का वाचन, कठिन शब्दार्थ एवं व्याख्या"]}
            ],
            "projects_experiential": [f"'{chapter}' के विषय पर एक लघु अनुच्छेद या कहानी लिखें।"],
            "skills_acquired": ["श्रवण (Listening)", "वाचन (Speaking)", "पठन (Reading)", "लेखन (Writing)"],
            "values_inculcated": ["नैतिक मूल्य", "सांस्कृतिक चेतना", "संवेदनशीलता"],
            "multiple_assessment": {
                "oral_questions": [f"'{chapter}' का मुख्य संदेश क्या है?"],
                "worksheet": ["बहुविकल्पीय प्रश्न (MCQs)", "शब्दार्थ एवं वाक्य प्रयोग"],
                "practical": ["सस्वर वाचन एवं उच्चारण सुधार"],
                "exit_ticket": ["आज सीखे गए नए शब्द से एक वाक्य बनाएं।"]
            },
            "class_work": ["पाठ्यपुस्तक के प्रश्नोत्तर लिखना", "शब्दार्थ अभ्यास"],
            "home_work": [f"'{chapter}' का सारांश अपने शब्दों में लिखें।"],
            "remedial_measures": {
                "slow_learners": ["सस्वर वाचन अभ्यास", "चित्रों के माध्यम से समझाना"],
                "advanced_learners": ["मौलिक रचनात्मक लेखन"]
            },
            "resources": {
                "books": [f"एनसीईआरटी कक्षा {grade} हिंदी पाठ्यपुस्तक"],
                "websites": ["दीक्षा पोर्टल (DIKSHA Portal)"],
                "videos": ["शैक्षणिक वीडियो एवं ऑडियो क्लिप्स"]
            }
        }

    # --- SANSKRIT ---
    elif subject == "SANSKRIT":
        return {
            "curriculum_goal": f"लक्ष्यम्-3: संस्कृतभाषायाः बोधगहनम्, नैतिकमूल्यानां विकासः तथा भाषाकौशलानां सम्पादनम्।",
            "relevant_competencies": [
                f"दक्षता-3.1: '{chapter}' पाठस्य शुद्धोच्चारणम्, वाचनम् एवं अर्थग्रहणम्।",
                "दक्षता-3.2: व्याकरणनियमानां तथा शब्दरूपाणां सम्यक् प्रयोगः।"
            ],
            "learning_objectives": [
                f"पाठस्य '{chapter}' मूलभावस्य अवगमनम्।",
                "नूतनशब्दानां श्लोकानां च अन्वयसहितं अर्थबोधः।"
            ],
            "expected_learning_outcomes": [
                f"छात्राः '{chapter}' पाठस्य श्लोकानाम् सरधार्थं वक्तुं समर्थाः भविष्यन्ति।",
                "प्रश्नोत्तराणि संस्कृतभाषायाम् एव लेखिष्यन्ति।"
            ],
            "teaching_methodology": ["अन्वय विधिः", "पाठ्यपुस्तक विधिः", "अभ्यास विधिः"],
            "teaching_aids": ["श्यामपट्टः / स्मार्ट बोर्ड", "शब्दरूप/धातु-चित्रपटम्"],
            "art_integration": [f"'{chapter}' श्लोकगायनम् / चित्रांकनम्"],
            "previous_knowledge": ["पूर्वाधीतशब्दानां तथा व्याकरणस्य पुनरावृत्तिः।"],
            "innovative_techniques": ["संस्कृत सम्भाषणम्", "डिजिटल-श्लोकगायनम्"],
            "content_points": [
                {"section": "पाठपरिचयः", "topics": [f"'{chapter}' पाठस्य पृष्ठभूमिः सारः च"]},
                {"section": "वाचनम् एवं व्याख्या", "topics": ["शुद्ध वाचनम्, अन्वयः, कठिनशब्दानाम् अर्थः"]}
            ],
            "projects_experiential": [f"'{chapter}' पाठस्य श्लोकान् कण्ठस्थीकृत्य कुरुत।"],
            "skills_acquired": ["उच्चारणम्", "अवबोधनम्", "संस्कृत-भाषणम्", "लेखनम्"],
            "values_inculcated": ["भारतीयसंस्कृतिः", "सदाचारः", "अनुशासनम्"],
            "multiple_assessment": {
                "oral_questions": [f"'{chapter}' पाठे मुख्यः संदेशः कः अस्ति?"],
                "worksheet": ["प्रश्नोत्तराणि", "रिक्तस्थानपूर्तिः"],
                "practical": ["श्लोकगायनम् तथा शुद्धोच्चारणम्"],
                "exit_ticket": ["एकं नूतनं संस्कृतशब्दं तस्य अर्थं च लिखत।"]
            },
            "class_work": ["अभ्यासकार्याणां समाधानम्", "शब्दार्थानां लेखनम्"],
            "home_work": [f"'{chapter}' पाठस्य श्लोकानां सरधार्थं लिखत।"],
            "remedial_measures": {
                "slow_learners": ["वर्णमाला पुनरावृत्तिः", "व्यक्तिगतध्यानम्"],
                "advanced_learners": ["अतिरिक्त-संस्कृत-कथावाचनम्"]
            },
            "resources": {
                "books": [f"NCERT कक्षा {grade} संस्कृत पाठ्यपुस्तकम्"],
                "websites": ["दीक्षा पोर्टल (DIKSHA Portal)"],
                "videos": ["संस्कृत-शैक्षणिक-वीडियो"]
            }
        }

    # --- GENERAL ACADEMIC SUBJECTS (Science, Math, SST, English, etc.) ---
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
                "Analyze cause-and-effect relationships in practical and industrial contexts.",
                "Synthesize learnings through diagrams, data tables, and structured writing."
            ],
            "expected_learning_outcomes": [
                f"Students will accurately articulate definitions, mechanisms, and key terms for '{chapter}'.",
                "Solve conceptual problems and numerical/textual exercises independently.",
                "Demonstrate practical understanding through experiments, case studies, or flowcharts."
            ],
            "teaching_methodology": [
                "Inquiry-Based Learning", "Problem-Solving & Case Study Method",
                "Demonstration & Lab Experimentation", "Collaborative Peer Discussion"
            ],
            "teaching_aids": [
                "Smart Board / Interactive PPTs", "NCERT Lab Kits & Experimental Apparatus", "Worksheets & Mind Maps"
            ],
            "art_integration": [
                f"Concept Mapping & Graphical Flowchart of '{chapter}'", "Infographic Poster Creation"
            ],
            "previous_knowledge": [
                f"What foundational concepts from earlier grades relate to '{chapter}'?",
                "Where do we observe this concept in daily real-world activities?"
            ],
            "innovative_techniques": [
                "Blended Learning Modules", "QR Code Interactive Simulations", "Exit Tickets & Kahoot Quizzes"
            ],
            "content_points": [
                {"section": "Theoretical Framework", "topics": [f"Introduction, definitions, and scope of '{chapter}'"]},
                {"section": "Mechanism & Analysis", "topics": ["Step-by-step structural analysis, formulas, and working principles"]},
                {"section": "Practical Application & Review", "topics": ["Case studies, numerical practice, and summary concept mapping"]}
            ],
            "projects_experiential": [
                f"Conduct a mini-experiment or research survey investigating '{chapter}' in your local context."
            ],
            "skills_acquired": [
                "Critical Thinking", "Analytical Reasoning", "Experimental Inquiry", "Data Interpretation"
            ],
            "values_inculcated": [
                "Scientific Temper", "Environmental Responsibility", "Curiosity", "Intellectual Honesty"
            ],
            "multiple_assessment": {
                "oral_questions": [f"Explain the main principle behind '{chapter}'.", "Give two practical examples."],
                "worksheet": ["MCQs & Assertion-Reasoning", "Short Answer & Conceptual Questions"],
                "practical": ["Experiment Performance & Observation Table Recording"],
                "exit_ticket": ["Summarize today's core concept in 2 sentences."]
            },
            "class_work": ["Concept notes drafting", "Solving textbook numericals and exercise questions"],
            "home_work": [f"Complete review questions for '{chapter}' and construct a summary mind map."],
            "remedial_measures": {
                "slow_learners": ["Peer tutoring and visual flashcards", "Step-by-step formula worksheets"],
                "advanced_learners": ["Exemplar problem solving", "Advanced research assignments"]
            },
            "resources": {
                "books": [f"NCERT Class {grade} {subject} Textbook", "CBSE Exemplar Manual"],
                "websites": ["DIKSHA Portal", "NCERT e-Pathshala"],
                "videos": ["NCERT Demonstration Videos", "Khan Academy Lessons"]
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
chapter = st.sidebar.text_input("Chapter / Topic", "Fundamental Rhythms & Scale")
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
