import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from doc_styles import (
    add_body_p, add_chapter_heading, add_section_heading,
    set_cell_margins, set_cell_shading, format_table, set_table_borders
)

def create_cover_page(doc, logo_path="nmims_logo.png"):
    """Create the official university Cover Page strictly adhering to the template."""
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes")
    r_title.font.name = 'Times New Roman'
    r_title.font.size = Pt(19)
    r_title.font.bold = True
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(3)
    p_sub.paragraph_format.space_after = Pt(3)
    p_sub.paragraph_format.line_spacing = 1.15
    p_sub.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Project Report submitted in the partial fulfilment\nof\nBachelor of Technology\nin\nInformation Technology")
    r_sub.font.name = 'Times New Roman'
    r_sub.font.size = Pt(12)

    p_to = doc.add_paragraph()
    p_to.paragraph_format.space_before = Pt(3)
    p_to.paragraph_format.space_after = Pt(3)
    p_to.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_to = p_to.add_run("Submitted To")
    r_to.font.name = 'Times New Roman'
    r_to.font.size = Pt(11)
    r_to.font.bold = True

    # Logo
    p_logo = doc.add_paragraph()
    p_logo.paragraph_format.space_before = Pt(2)
    p_logo.paragraph_format.space_after = Pt(2)
    p_logo.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_logo = p_logo.add_run()
    r_logo.add_picture(logo_path, width=Inches(1.3))

    p_inst = doc.add_paragraph()
    p_inst.paragraph_format.space_before = Pt(2)
    p_inst.paragraph_format.space_after = Pt(4)
    p_inst.paragraph_format.line_spacing = 1.15
    p_inst.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_inst.add_run("SVKM’s NMIMS,\nMukesh Patel School of Technology Management & Engineering,\nShirpur Campus (M.H.)")
    r_inst.font.name = 'Times New Roman'
    r_inst.font.size = Pt(11)
    r_inst.font.bold = True

    p_by = doc.add_paragraph()
    p_by.paragraph_format.space_before = Pt(3)
    p_by.paragraph_format.space_after = Pt(2)
    p_by.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_by = p_by.add_run("Submitted By:")
    r_by.font.name = 'Times New Roman'
    r_by.font.size = Pt(11)
    r_by.font.bold = True

    p_students = doc.add_paragraph()
    p_students.paragraph_format.space_before = Pt(2)
    p_students.paragraph_format.space_after = Pt(4)
    p_students.paragraph_format.line_spacing = 1.15
    p_students.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_s1 = p_students.add_run("Aman Choudhary Enrollment No – 70012300166\n")
    r_s2 = p_students.add_run("Pranav Patil Enrollment No – 70012400176\n")
    r_s3 = p_students.add_run("Mohit Patil Enrollment No – 70012400178")
    for r in [r_s1, r_s2, r_s3]:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11.5)
        r.font.bold = True

    p_sup = doc.add_paragraph()
    p_sup.paragraph_format.space_before = Pt(3)
    p_sup.paragraph_format.space_after = Pt(2)
    p_sup.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sup = p_sup.add_run("Under The Supervision Of:")
    r_sup.font.name = 'Times New Roman'
    r_sup.font.size = Pt(11)
    r_sup.font.bold = True

    p_mentor = doc.add_paragraph()
    p_mentor.paragraph_format.space_before = Pt(2)
    p_mentor.paragraph_format.space_after = Pt(6)
    p_mentor.paragraph_format.line_spacing = 1.15
    p_mentor.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_m1 = p_mentor.add_run("Dr. Pravin Landge\n")
    r_m1.font.bold = True
    r_m2 = p_mentor.add_run("(Faculty Mentor, Information Technology Department, MPSTME)")
    for r in [r_m1, r_m2]:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11.5)

    p_dept = doc.add_paragraph()
    p_dept.paragraph_format.space_before = Pt(4)
    p_dept.paragraph_format.space_after = Pt(0)
    p_dept.paragraph_format.line_spacing = 1.15
    p_dept.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_d1 = p_dept.add_run("INFORMATION TECHNOLOGY DEPARTMENT\n")
    r_d1.font.size = Pt(12)
    r_d1.font.bold = True
    r_d2 = p_dept.add_run("Mukesh Patel School of Technology Management & Engineering\n")
    r_d2.font.size = Pt(10.5)
    r_d3 = p_dept.add_run("ACADEMIC SESSION: 2026-27")
    r_d3.font.size = Pt(10.5)
    r_d3.font.bold = True
    for r in [r_d1, r_d2, r_d3]:
        r.font.name = 'Times New Roman'

    doc.add_page_break()

def create_certificate_page(doc, logo_path="nmims_logo.png"):
    """Create the official Certificate page with required academic text and signature fields."""
    p_cert = doc.add_paragraph()
    p_cert.paragraph_format.space_before = Pt(0)
    p_cert.paragraph_format.space_after = Pt(4)
    p_cert.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cert = p_cert.add_run("CERTIFICATE")
    r_cert.font.name = 'Times New Roman'
    r_cert.font.size = Pt(16)
    r_cert.font.bold = True

    p_logo = doc.add_paragraph()
    p_logo.paragraph_format.space_before = Pt(2)
    p_logo.paragraph_format.space_after = Pt(6)
    p_logo.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_logo = p_logo.add_run()
    r_logo.add_picture(logo_path, width=Inches(1.25))

    p_body = doc.add_paragraph()
    p_body.paragraph_format.space_before = Pt(4)
    p_body.paragraph_format.space_after = Pt(6)
    p_body.paragraph_format.line_spacing = 1.2
    p_body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    r1 = p_body.add_run("This is to certify that the project entitled ")
    r2 = p_body.add_run("“AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes” ")
    r2.font.bold = True
    r3 = p_body.add_run("has been done by\n")
    for r in [r1, r2, r3]:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11.5)

    p_names = doc.add_paragraph()
    p_names.paragraph_format.space_before = Pt(2)
    p_names.paragraph_format.space_after = Pt(6)
    p_names.paragraph_format.line_spacing = 1.15
    p_names.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_n1 = p_names.add_run("“Aman Choudhary” (Roll No.: 70012300166)\n")
    r_n2 = p_names.add_run("“Pranav Patil” (Roll No.: 70012400176)\n")
    r_n3 = p_names.add_run("“Mohit Patil” (Roll No.: 70012400178)\n")
    for r in [r_n1, r_n2, r_n3]:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11.5)
        r.font.bold = True

    p_foot = doc.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(2)
    p_foot.paragraph_format.space_after = Pt(14)
    p_foot.paragraph_format.line_spacing = 1.2
    p_foot.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_f1 = p_foot.add_run("under my guidance and supervision & has been submitted in partial fulfilment of the degree of ")
    r_f2 = p_foot.add_run("“Bachelor of Technology in Information Technology” ")
    r_f2.font.bold = True
    r_f3 = p_foot.add_run("of MPSTME, SVKM’S NMIMS (Deemed-to-be University), Shirpur (M.H.), India.")
    for r in [r_f1, r_f2, r_f3]:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11.5)

    # Signature Block Table
    tbl = doc.add_table(rows=4, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    col_w = Inches(3.2)
    for row in tbl.rows:
        for cell in row.cells:
            cell.width = col_w
            set_cell_margins(cell, top=30, bottom=30, left=30, right=30)

    # Row 0: Signature lines
    tbl.cell(0, 0).paragraphs[0].text = "________________________"
    tbl.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    tbl.cell(0, 1).paragraphs[0].text = "________________________"
    tbl.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Row 1: Mentor & Examiner
    p_m = tbl.cell(1, 0).paragraphs[0]
    p_m.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_m.add_run("Project Mentor (Internal Guide)\nDr. Pravin Landge")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(10.5)
    r.font.bold = True

    p_e = tbl.cell(1, 1).paragraphs[0]
    p_e.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_e.add_run("Examiner\n ")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(10.5)
    r.font.bold = True

    # Row 2: Spacing & Place/Date
    p_d = tbl.cell(2, 0).paragraphs[0]
    p_d.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_d.paragraph_format.space_before = Pt(8)
    r = p_d.add_run("Date: ____________________\nPlace: Shirpur")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(10.5)

    p_h = tbl.cell(2, 1).paragraphs[0]
    p_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_h.paragraph_format.space_before = Pt(8)
    r = p_h.add_run("________________________\nH.O.D.\nInformation Technology Dept.")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(10.5)
    r.font.bold = True

    # Department footer
    p_dfoot = doc.add_paragraph()
    p_dfoot.paragraph_format.space_before = Pt(10)
    p_dfoot.paragraph_format.space_after = Pt(0)
    p_dfoot.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_df1 = p_dfoot.add_run("INFORMATION TECHNOLOGY DEPARTMENT\n")
    r_df1.font.size = Pt(11.5)
    r_df1.font.bold = True
    r_df2 = p_dfoot.add_run("Mukesh Patel School of Technology Management & Engineering")
    r_df2.font.size = Pt(10.5)
    for r in [r_df1, r_df2]:
        r.font.name = 'Times New Roman'

def create_acknowledgement_page(doc):
    """Create genuine, respectful acknowledgement page fitting exactly on 1 page."""
    p_ack = doc.add_paragraph()
    p_ack.paragraph_format.space_before = Pt(12)
    p_ack.paragraph_format.space_after = Pt(10)
    p_ack.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ack = p_ack.add_run("ACKNOWLEDGEMENT")
    r_ack.font.name = 'Times New Roman'
    r_ack.font.size = Pt(15)
    r_ack.font.bold = True

    ack_paragraphs = [
        "After the completion of this capstone project work, words are not enough to express our sincere feelings about all those who helped us reach our goal; feeling above this is our indebtedness to the Almighty for providing us this opportunity and moment in life.",
        "It gives us immense pleasure and profound satisfaction to express our deepest gratitude to our faculty guide, Dr. Pravin Landge, Information Technology Department, Mukesh Patel School of Technology Management & Engineering (MPSTME), Shirpur, whose constant encouragement, constructive technical critiques, and scholarly guidance enabled us to work with rigorous dedication. His perpetual motivation, patience, and profound technical insight during discussions on software architecture, audio loopback routing, real-time speech processing, and language model integration benefited us immensely. Working under his supervision has been an exceptionally enriching, educational, and unforgettable academic journey.",
        "We convey our heartfelt thanks and sincere gratitude to the Head of Department, Information Technology Department, MPSTME, Shirpur, for providing the necessary infrastructural facilities, high-performance computing laboratories, and technical resources essential for successfully implementing and validating this capstone project.",
        "We also extend our earnest gratitude to the Associate Dean, SVKM’s NMIMS, Shirpur Campus, for fostering an academic environment dedicated to technical innovation, research excellence, and hands-on software engineering.",
        "We sincerely thank all the faculty members and technical staff of the Information Technology Department who imparted valuable knowledge and foundational computing principles during the various semesters of our Bachelor of Technology program. Finally, we express our warmest appreciation to our family members and peers who supported us through their constant encouragement, insightful discussions, and moral support throughout this project."
    ]

    for text in ack_paragraphs:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.2
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11)

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(12)
    p_sig.paragraph_format.space_after = Pt(0)
    p_sig.paragraph_format.line_spacing = 1.15
    p_sig.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sig = p_sig.add_run("Aman Choudhary (Enrollment No. 70012300166)\nPranav Patil (Enrollment No. 70012400176)\nMohit Patil (Enrollment No. 70012400178)")
    r_sig.font.name = 'Times New Roman'
    r_sig.font.size = Pt(11)
    r_sig.font.bold = True

    doc.add_page_break()

def create_abstract_page(doc):
    """Create a comprehensive, academically rigorous abstract fitting exactly on 1 page."""
    p_abs = doc.add_paragraph()
    p_abs.paragraph_format.space_before = Pt(12)
    p_abs.paragraph_format.space_after = Pt(10)
    p_abs.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_abs = p_abs.add_run("ABSTRACT")
    r_abs.font.name = 'Times New Roman'
    r_abs.font.size = Pt(15)
    r_abs.font.bold = True

    abs_paragraphs = [
        "The rapid proliferation of digital learning platforms and synchronous virtual conferencing systems has established online education as a mainstream pedagogical paradigm across modern universities. However, attending interactive live lectures imposes severe cognitive strain on students, who must simultaneously comprehend complex instructional derivations, maintain comprehensive handwritten or digital notes, monitor communication channels, and remain vigilant for verbal roll-calls or targeted technical cold-calls from instructors. This multitasking split-attention effect frequently degrades conceptual comprehension, induces cognitive fatigue, and causes critical lecture insights to be lost. Although contemporary generative artificial intelligence (AI) tools offer powerful reasoning capabilities, existing educational solutions remain largely passive and reactive—requiring continuous human manual prompting and lacking ambient auditory awareness, real-time classroom agency, and persistent knowledge synthesis.",
        "To resolve these systemic challenges, this capstone project presents AI Student Agent, an autonomous, multi-threaded AI learning assistant engineered specifically for real-time online classes conducted via Google Meet. The system implements a robust, hybrid local-cloud architecture designed to bridge the fundamental autonomy, memory, knowledge, and evaluation gaps inherent in modern educational software. Operating on standard consumer personal computers without requiring dedicated graphics processing hardware, the agent programmatically accesses scheduled Google Meet sessions through Selenium WebDriver, autonomously bypassing media permissions and navigating interface controls. To achieve continuous auditory awareness without recurring cloud streaming costs, the agent routes meeting audio through a unidirectional VB-Audio Virtual Cable loopback stream into an optimized local offline speech-to-text pipeline powered by the VOSK acoustic engine with Kaldi decoding libraries, operating entirely on the host central processing unit (CPU).",
        "The core intelligence engine utilizes a sliding context window buffer maintaining a rolling history of lecture dialogue to preserve conversational discourse. Spoken audio is analyzed in real time using fuzzy phonetic matching and acoustic confidence gating (confidence >= 0.70) to detect personalized student name callouts, attendance roll-calls, and direct questions. Upon trigger verification, the agent dynamically synthesizes grounded, context-aware responses using OpenAI’s GPT-4o-mini large language model and automatically dispatches answers directly into the active Google Meet chat stream through resilient browser Document Object Model (DOM) automation. Furthermore, the system captures a lossless pulse-code modulation (PCM) audio recording throughout the lecture, which undergoes post-session high-fidelity transcription via OpenAI Whisper and hierarchical summarization into structured lecture notes. These notes are programmatically synchronized with the student's personal Notion workspace via the official Notion REST API.",
    ]

    for text in abs_paragraphs:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.2
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10.5)

    p_kw = doc.add_paragraph()
    p_kw.paragraph_format.space_before = Pt(6)
    p_kw.paragraph_format.space_after = Pt(0)
    p_kw.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_kwh = p_kw.add_run("Keywords: ")
    r_kwh.font.name = 'Times New Roman'
    r_kwh.font.size = Pt(10)
    r_kwh.font.bold = True
    r_kwt = p_kw.add_run("Autonomous Agents, Speech Recognition, VOSK, Large Language Models, Google Meet, Robotic Process Automation, Notion API, Educational Technology.")
    r_kwt.font.name = 'Times New Roman'
    r_kwt.font.size = Pt(10)
    r_kwt.font.italic = True

    doc.add_page_break()

def create_table_of_contents_page(doc):
    """Create Table of Contents with visible borders strictly matching the template."""
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.space_before = Pt(14)
    p_toc.paragraph_format.space_after = Pt(10)
    p_toc.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_toc = p_toc.add_run("TABLE OF CONTENTS")
    r_toc.font.name = 'Times New Roman'
    r_toc.font.size = Pt(15)
    r_toc.font.bold = True

    toc_items = [
        ("Topics", "Page", True),
        ("List of Figures", "v", False),
        ("List of Tables", "vi", False),
        ("Abbreviations", "vii", False),
        ("Chapter 1  INTRODUCTION", "1", True),
        ("  1.1  Background of the Project Topic", "1", False),
        ("  1.2  Motivation and Scope of the Report", "2", False),
        ("  1.3  Problem Statement", "3", False),
        ("  1.4  Salient Contributions", "4", False),
        ("  1.5  Organization of the Report", "5", False),
        ("Chapter 2  LITERATURE SURVEY", "6", True),
        ("  2.1  Thematic Categorization of Literature", "6", False),
        ("  2.2  Detailed Review of Existing Research", "7", False),
        ("  2.3  Research Gap Synthesis and Technical Motivation", "8", False),
        ("Chapter 3  METHODOLOGY AND IMPLEMENTATION", "10", True),
        ("  3.1  Block Diagram and Architectural Framework", "10", False),
        ("  3.2  Hardware Description", "15", False),
        ("  3.3  Software Description, Audio Math and Algorithms", "17", False),
        ("  3.4  Technical Implementation Challenges and Solutions", "20", False),
        ("Chapter 4  RESULTS AND ANALYSIS", "23", True),
        ("  4.1  System Verification and Functional Test Matrix", "23", False),
        ("  4.2  Qualitative Demonstration and Live Session Outcomes", "24", False),
        ("  4.3  Latency, Resource Consumption, and Performance Analysis", "26", False),
        ("  4.4  Comparative Evaluation and Economic Cost Analysis", "27", False),
        ("  4.5  Architectural Trade-offs and Quality Assurance", "27", False),
        ("  4.6  Project Development Timeline and Milestone Tracking", "28", False),
        ("Chapter 5  ADVANTAGES, LIMITATIONS AND APPLICATIONS", "30", True),
        ("  5.1  Advantages", "30", False),
        ("  5.2  Limitations", "31", False),
        ("  5.3  Applications", "32", False),
        ("Chapter 6  CONCLUSION AND FUTURE SCOPE", "33", True),
        ("  6.1  Conclusion", "33", False),
        ("  6.2  Future Scope", "34", False),
        ("References", "36", True),
        ("Appendix A: Sample Code Excerpts", "37", True),
        ("Appendix B: Technical Specifications and Component Data Sheets", "38", True),
        ("Appendix C: Comprehensive List of Components", "39", True),
        ("Appendix D: List of Papers Presented and Published", "40", True),
    ]

    tbl = doc.add_table(rows=len(toc_items), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(4.8)
    tbl.columns[1].width = Inches(0.97)

    set_table_borders(tbl, color="7F7F7F", sz="4")

    for i, (topic, page, is_head) in enumerate(toc_items):
        cell_t = tbl.cell(i, 0)
        cell_p = tbl.cell(i, 1)
        set_cell_margins(cell_t, top=25, bottom=25, left=50, right=50)
        set_cell_margins(cell_p, top=25, bottom=25, left=50, right=50)
        
        p_t = cell_t.paragraphs[0]
        p_t.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_t.paragraph_format.line_spacing = 1.15
        p_t.paragraph_format.space_before = Pt(1)
        p_t.paragraph_format.space_after = Pt(1)
        r_t = p_t.add_run(topic)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(9.5 if is_head else 9)
        r_t.font.bold = is_head
        
        p_p = cell_p.paragraphs[0]
        p_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_p.paragraph_format.line_spacing = 1.15
        p_p.paragraph_format.space_before = Pt(1)
        p_p.paragraph_format.space_after = Pt(1)
        r_p = p_p.add_run(page)
        r_p.font.name = 'Times New Roman'
        r_p.font.size = Pt(9.5 if is_head else 9)
        r_p.font.bold = is_head

        if i == 0:
            set_cell_shading(cell_t, "E8EEF5")
            set_cell_shading(cell_p, "E8EEF5")
            p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()

def create_list_of_figures_page(doc):
    """Create List of Figures with visible borders fitting on 1 page."""
    p_lof = doc.add_paragraph()
    p_lof.paragraph_format.space_before = Pt(14)
    p_lof.paragraph_format.space_after = Pt(10)
    p_lof.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_lof = p_lof.add_run("LIST OF FIGURES")
    r_lof.font.name = 'Times New Roman'
    r_lof.font.size = Pt(15)
    r_lof.font.bold = True

    figs = [
        ("Sr. No.", "Figure No.", "Name of Figures", "Page"),
        ("1", "Fig. 3.1", "High-Level System Architecture and Layered Processing Pipeline", "11"),
        ("2", "Fig. 3.2", "UML Use Case Diagram of the AI Student Agent System", "13"),
        ("3", "Fig. 3.3", "UML Activity Diagram Depicting Meeting Lifecycle Workflows", "13"),
        ("4", "Fig. 3.4", "Object-Oriented Class Diagram and Structural Hierarchy", "14"),
        ("5", "Fig. 3.5", "Entity-Relationship (ER) Diagram of Session Data and Knowledge Models", "14"),
        ("6", "Fig. 3.6", "System Sequence Diagram for Real-Time Trigger and Chat Response", "15"),
        ("7", "Fig. 3.7", "Software Ecosystem and Technical Stack Integration Architecture", "17"),
        ("8", "Fig. 4.1", "Live Student Agent Operational Dashboard Interface", "25"),
        ("9", "Fig. 4.2", "Automated Google Meet In-Call Session and Chat Stream Interface", "25"),
        ("10", "Fig. 4.3", "Project Development Timeline and Milestone Execution (Gantt Chart)", "28"),
    ]

    tbl = doc.add_table(rows=len(figs), cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(0.7)
    tbl.columns[1].width = Inches(0.9)
    tbl.columns[2].width = Inches(3.4)
    tbl.columns[3].width = Inches(0.77)

    set_table_borders(tbl, color="7F7F7F", sz="4")

    for i, row in enumerate(figs):
        for c, text in enumerate(row):
            cell = tbl.cell(i, c)
            set_cell_margins(cell, top=30, bottom=30, left=45, right=45)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(text)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5 if i == 0 else 9)
            if i == 0:
                r.font.bold = True
                set_cell_shading(cell, "E8EEF5")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                if c in [0, 1, 3]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_page_break()

def create_list_of_tables_page(doc):
    """Create List of Tables with visible borders fitting on 1 page."""
    p_lot = doc.add_paragraph()
    p_lot.paragraph_format.space_before = Pt(14)
    p_lot.paragraph_format.space_after = Pt(10)
    p_lot.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_lot = p_lot.add_run("LIST OF TABLES")
    r_lot.font.name = 'Times New Roman'
    r_lot.font.size = Pt(15)
    r_lot.font.bold = True

    tables = [
        ("Sr. No.", "Table No.", "Name of Table", "Page"),
        ("1", "Table 2.1", "Comprehensive Comparative Summary of Reviewed Literature", "8"),
        ("2", "Table 3.1", "Hardware and Computational Environment Specifications", "15"),
        ("3", "Table 3.2", "System Architecture Component Execution and Status Matrix", "16"),
        ("4", "Table 3.3", "Audio Ingestion and Buffer Dimension Parameters", "18"),
        ("5", "Table 4.1", "Comprehensive Functional Verification and Test Matrix", "23"),
        ("6", "Table 4.2", "Computational Resource Footprint and Latency Metrics", "26"),
        ("7", "Table 4.3", "Economic Compute Cost Comparison: Continuous Cloud vs Hybrid Local", "27"),
        ("8", "Table 4.4", "Architectural Trade-off Analysis Across Alternative Implementations", "28"),
        ("9", "Table B.1", "VOSK Kaldi Acoustic Model Specifications", "38"),
        ("10", "Table B.2", "OpenAI GPT-4o-mini Operational Parameter Specifications", "38"),
        ("11", "Table C.1", "Comprehensive Hardware Component Manifest", "39"),
        ("12", "Table C.2", "Python Package Ecosystem and Installed Dependency Manifest", "39"),
    ]

    tbl = doc.add_table(rows=len(tables), cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(0.7)
    tbl.columns[1].width = Inches(0.9)
    tbl.columns[2].width = Inches(3.4)
    tbl.columns[3].width = Inches(0.77)

    set_table_borders(tbl, color="7F7F7F", sz="4")

    for i, row in enumerate(tables):
        for c, text in enumerate(row):
            cell = tbl.cell(i, c)
            set_cell_margins(cell, top=30, bottom=30, left=45, right=45)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(text)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5 if i == 0 else 9)
            if i == 0:
                r.font.bold = True
                set_cell_shading(cell, "E8EEF5")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                if c in [0, 1, 3]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_page_break()

def create_abbreviations_page(doc):
    """Create comprehensive List of Abbreviations formatted in a clean 2-column table fitting on 1 page."""
    p_abb = doc.add_paragraph()
    p_abb.paragraph_format.space_before = Pt(14)
    p_abb.paragraph_format.space_after = Pt(10)
    p_abb.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_abb = p_abb.add_run("ABBREVIATIONS")
    r_abb.font.name = 'Times New Roman'
    r_abb.font.size = Pt(15)
    r_abb.font.bold = True

    # 2-column layout like the friend's report: Abbr | Full Form | Abbr | Full Form
    abbrevs_pairs = [
        ("Abbr.", "Full Expansion", "Abbr.", "Full Expansion"),
        ("AI", "Artificial Intelligence", "ML", "Machine Learning"),
        ("ASR", "Automatic Speech Recognition", "LLM", "Large Language Model"),
        ("NLP", "Natural Language Processing", "STT", "Speech-to-Text"),
        ("TTS", "Text-to-Speech", "RPA", "Robotic Process Automation"),
        ("RAG", "Retrieval-Augmented Generation", "DOM", "Document Object Model"),
        ("PCM", "Pulse-Code Modulation", "WAV", "Waveform Audio File Format"),
        ("API", "Application Programming Interface", "REST", "Representational State Transfer"),
        ("JSON", "JavaScript Object Notation", "HTTP", "Hypertext Transfer Protocol"),
        ("CORS", "Cross-Origin Resource Sharing", "UI/UX", "User Interface / Experience"),
        ("CPU", "Central Processing Unit", "GPU", "Graphics Processing Unit"),
        ("RAM", "Random Access Memory", "OS", "Operating System"),
        ("WER", "Word Error Rate", "WFST", "Weighted Finite-State Transducer"),
        ("BVA", "Boundary Value Analysis", "EPP", "Equivalence Partitioning"),
    ]

    tbl = doc.add_table(rows=len(abbrevs_pairs), cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(0.8)
    tbl.columns[1].width = Inches(2.05)
    tbl.columns[2].width = Inches(0.8)
    tbl.columns[3].width = Inches(2.12)

    set_table_borders(tbl, color="7F7F7F", sz="4")

    for i, row in enumerate(abbrevs_pairs):
        for c, text in enumerate(row):
            cell = tbl.cell(i, c)
            set_cell_margins(cell, top=25, bottom=25, left=40, right=40)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(text)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5 if i == 0 else 9)
            if i == 0:
                r.font.bold = True
                set_cell_shading(cell, "E8EEF5")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                if c in [0, 2]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r.font.bold = True
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
