import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Exact brand palette matching MedIntel AI & NMIMS
COLOR_RED = RGBColor(188, 49, 44)        # #BC312C (NMIMS Brand Red)
COLOR_BLUE = RGBColor(46, 95, 138)       # #2E5F8A (Steel Navy Blue)
COLOR_GREEN = RGBColor(63, 143, 92)      # #3F8F5C (Forest Green)
COLOR_AMBER = RGBColor(201, 138, 44)     # #C98A2C (Amber Gold)
COLOR_DARK = RGBColor(51, 51, 51)        # #333333 (Primary Dark Text)
COLOR_MUTED = RGBColor(102, 102, 102)    # #666666 (Subtext Gray)
COLOR_CARD_BG = RGBColor(244, 246, 248)  # #F4F6F8 (Neutral Card Fill)
COLOR_WHITE = RGBColor(255, 255, 255)    # #FFFFFF (Pure White)
COLOR_BORDER = RGBColor(208, 208, 208)   # #D0D0D0 (Subtle Gray Border)
COLOR_LIGHT_RED = RGBColor(253, 242, 242)

FONT_FAMILY = "Times New Roman"

def add_header(slide, title_text, slide_num):
    """Adds title and slide number matching MedIntel template."""
    tb = slide.shapes.add_textbox(Inches(0.50), Inches(1.15), Inches(8.50), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK
    
    num_box = slide.shapes.add_textbox(Inches(7.50), Inches(6.92), Inches(2.00), Inches(0.35))
    ntf = num_box.text_frame
    ntf.margin_left = ntf.margin_top = ntf.margin_right = ntf.margin_bottom = 0
    np = ntf.paragraphs[0]
    np.alignment = PP_ALIGN.RIGHT
    np.text = str(slide_num)
    np.font.name = FONT_FAMILY
    np.font.size = Pt(13)
    np.font.color.rgb = COLOR_MUTED

def add_card(slide, left, top, width, height, fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER, border_width=Pt(1)):
    """Creates a rounded rectangle card."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = fill_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = border_width
    else:
        card.line.fill.background()
    return card

def add_badge(slide, left, top, size, number_str, color=COLOR_RED):
    """Creates a circular numbered badge."""
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    oval.fill.solid()
    oval.fill.fore_color.rgb = color
    oval.line.fill.background()
    tf = oval.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = str(number_str)
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_FAMILY
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    return oval

def style_table(table, col_widths=None):
    """Styles tables with NMIMS Red header and alternating row fills."""
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    for cell in table.rows[0].cells:
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_RED
        for p in cell.text_frame.paragraphs:
            p.font.name = FONT_FAMILY
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = COLOR_WHITE
            p.alignment = PP_ALIGN.CENTER
    for r_idx in range(1, len(table.rows)):
        row = table.rows[r_idx]
        bg_color = COLOR_WHITE if r_idx % 2 == 1 else COLOR_CARD_BG
        for cell in row.cells:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_color
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT_FAMILY
                p.font.size = Pt(10)
                p.font.color.rgb = COLOR_DARK

def add_bullet_point(tf, bold_prefix, text_content, font_size=Pt(12), space_after=Pt(8)):
    p = tf.add_paragraph() if len(tf.paragraphs[0].text) > 0 else tf.paragraphs[0]
    p.space_after = space_after
    if bold_prefix:
        r1 = p.add_run()
        r1.text = bold_prefix + " "
        r1.font.name = FONT_FAMILY
        r1.font.size = font_size
        r1.font.bold = True
        r1.font.color.rgb = COLOR_DARK
    r2 = p.add_run()
    r2.text = text_content
    r2.font.name = FONT_FAMILY
    r2.font.size = font_size
    r2.font.bold = False
    r2.font.color.rgb = COLOR_DARK

def generate_full_presentation():
    template_path = r'C:\Users\ganesh jadhav\Downloads\MedIntel_AI_Final_Presentation.pptx'
    prs = Presentation(template_path)
    
    # Remove existing slides to preserve master and image1.jpeg background
    for i in range(len(prs.slides) - 1, -1, -1):
        rId = prs.slides._sldIdLst[i].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[i]
        
    blank_layout = prs.slide_layouts[6]
    
    # =========================================================================
    # SLIDE 1: Title Slide
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    logo_path = 'extracted_media/image1.png'
    if os.path.exists(logo_path):
        s1.shapes.add_picture(logo_path, Inches(4.50), Inches(0.95), Inches(1.00), Inches(1.04))
    
    dept_box = s1.shapes.add_textbox(Inches(1.00), Inches(2.05), Inches(8.00), Inches(0.80))
    dtf = dept_box.text_frame
    dtf.word_wrap = True
    dp1 = dtf.paragraphs[0]
    dp1.alignment = PP_ALIGN.CENTER
    dp1.text = "Department of Information Technology"
    dp1.font.name = FONT_FAMILY
    dp1.font.size = Pt(15)
    dp1.font.bold = True
    dp1.font.color.rgb = COLOR_DARK
    dp2 = dtf.add_paragraph()
    dp2.alignment = PP_ALIGN.CENTER
    dp2.text = "Mukesh Patel School of Technology Management & Engineering"
    dp2.font.name = FONT_FAMILY
    dp2.font.size = Pt(13)
    dp2.font.color.rgb = COLOR_MUTED
    dp3 = dtf.add_paragraph()
    dp3.alignment = PP_ALIGN.CENTER
    dp3.text = "SVKM's NMIMS Shirpur Campus | Academic Year 2026-27"
    dp3.font.name = FONT_FAMILY
    dp3.font.size = Pt(12)
    dp3.font.color.rgb = COLOR_MUTED

    title_box = s1.shapes.add_textbox(Inches(0.60), Inches(3.05), Inches(8.80), Inches(1.60))
    ttf = title_box.text_frame
    ttf.word_wrap = True
    tp1 = ttf.paragraphs[0]
    tp1.alignment = PP_ALIGN.CENTER
    tp1.text = "Final Project Presentation"
    tp1.font.name = FONT_FAMILY
    tp1.font.size = Pt(16)
    tp1.font.color.rgb = COLOR_RED
    tp1.font.bold = True
    tp2 = ttf.add_paragraph()
    tp2.alignment = PP_ALIGN.CENTER
    tp2.text = "on"
    tp2.font.name = FONT_FAMILY
    tp2.font.size = Pt(13)
    tp2.font.color.rgb = COLOR_MUTED
    tp3 = ttf.add_paragraph()
    tp3.alignment = PP_ALIGN.CENTER
    tp3.text = "“AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes”"
    tp3.font.name = FONT_FAMILY
    tp3.font.size = Pt(21)
    tp3.font.bold = True
    tp3.font.color.rgb = COLOR_DARK

    info_box = s1.shapes.add_textbox(Inches(0.80), Inches(4.85), Inches(8.40), Inches(1.80))
    itf = info_box.text_frame
    itf.word_wrap = True
    ip1 = itf.paragraphs[0]
    ip1.alignment = PP_ALIGN.CENTER
    ip1.text = "Presented By:"
    ip1.font.name = FONT_FAMILY
    ip1.font.size = Pt(13)
    ip1.font.bold = True
    ip1.font.color.rgb = COLOR_DARK
    ip2 = itf.add_paragraph()
    ip2.alignment = PP_ALIGN.CENTER
    ip2.text = "1. Aman Choudhary (70012300166)     2. Pranav Patil (70012400176)     3. Mohit Patil (70012400178)"
    ip2.font.name = FONT_FAMILY
    ip2.font.size = Pt(12)
    ip2.font.color.rgb = COLOR_DARK
    ip3 = itf.add_paragraph()
    ip3.alignment = PP_ALIGN.CENTER
    ip3.text = "\nUnder the Guidance of:"
    ip3.font.name = FONT_FAMILY
    ip3.font.size = Pt(13)
    ip3.font.bold = True
    ip3.font.color.rgb = COLOR_DARK
    ip4 = itf.add_paragraph()
    ip4.alignment = PP_ALIGN.CENTER
    ip4.text = "Dr. Pravin Landge (Faculty Mentor, Department of Information Technology)"
    ip4.font.name = FONT_FAMILY
    ip4.font.size = Pt(12)
    ip4.font.color.rgb = COLOR_RED

    # =========================================================================
    # SLIDE 2: Problem Statement
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Problem Statement", 2)
    
    # Left column: 5 Systemic gaps
    left_tb = s2.shapes.add_textbox(Inches(0.50), Inches(1.90), Inches(5.10), Inches(4.80))
    ltf = left_tb.text_frame
    ltf.word_wrap = True
    ltf.margin_left = ltf.margin_top = ltf.margin_right = ltf.margin_bottom = 0
    add_bullet_point(ltf, "Split-Attention & Cognitive Load:", "Students suffer acute mental fatigue trying to simultaneously follow complex lecture derivations, take running notes, and watch chat streams (Sweller's CLT).", font_size=Pt(12.5), space_after=Pt(10))
    add_bullet_point(ltf, "Autonomy Gap in EdTech:", "Existing AI tools (ChatGPT, Copilot) are passive web forms; they cannot listen to live meeting audio or take autonomous action in the call.", font_size=Pt(12.5), space_after=Pt(10))
    add_bullet_point(ltf, "Ephemeral Class Discussions:", "Crucial verbal explanations, Q&A interactions, and attendance roll-calls vanish immediately once the Google Meet session terminates.", font_size=Pt(12.5), space_after=Pt(10))
    add_bullet_point(ltf, "Unbounded LLM Hallucinations:", "General-purpose models hallucinate unless strictly grounded in the rolling context of the instructor's immediate lecture discourse.", font_size=Pt(12.5), space_after=Pt(10))
    add_bullet_point(ltf, "Participation & Cold-Call Risk:", "Missing a verbal name callout due to distraction or audio lag leads to penalized participation grades and attendance loss.", font_size=Pt(12.5), space_after=Pt(6))

    # Right column: "What happens in a live class" card (matches MedIntel "What the patient sees")
    add_card(s2, Inches(5.80), Inches(1.90), Inches(3.70), Inches(4.75), fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER)
    r_hdr = s2.shapes.add_textbox(Inches(6.00), Inches(2.05), Inches(3.30), Inches(0.40))
    rh_tf = r_hdr.text_frame
    rp = rh_tf.paragraphs[0]
    rp.text = "What Happens in a Live Online Class"
    rp.font.name = FONT_FAMILY
    rp.font.size = Pt(13)
    rp.font.bold = True
    rp.font.color.rgb = COLOR_RED

    # Scenario cards
    def add_mini_scenario(top_in, title, detail, tag="?"):
        add_card(s2, Inches(6.00), Inches(top_in), Inches(3.30), Inches(0.80), fill_color=COLOR_WHITE, border_color=COLOR_BORDER)
        tb = s2.shapes.add_textbox(Inches(6.10), Inches(top_in + 0.05), Inches(2.70), Inches(0.70))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_FAMILY
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_DARK
        p2 = tf.add_paragraph()
        p2.text = detail
        p2.font.name = FONT_FAMILY
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = COLOR_MUTED
        
        tag_tb = s2.shapes.add_textbox(Inches(8.85), Inches(top_in + 0.15), Inches(0.35), Inches(0.40))
        ttp = tag_tb.text_frame.paragraphs[0]
        ttp.text = tag
        ttp.alignment = PP_ALIGN.CENTER
        ttp.font.name = FONT_FAMILY
        ttp.font.size = Pt(12)
        ttp.font.bold = True
        ttp.font.color.rgb = COLOR_RED

    add_mini_scenario(2.55, "Instructor Spoken Cold-Call", "“Justin, can you explain how paging prevents external fragmentation?”", tag="!")
    add_mini_scenario(3.45, "Cognitive Split-Attention Conflict", "Student transcribing diagram -> must stop notes -> search memory -> formulate answer", tag="X")
    add_mini_scenario(4.35, "10–15 Second Response Delay", "Typing in chat takes too long -> Teacher assumes student is absent or inattentive", tag="?")
    
    # Bottom callout in right container
    add_card(s2, Inches(6.00), Inches(5.25), Inches(3.30), Inches(1.20), fill_color=COLOR_LIGHT_RED, border_color=COLOR_RED)
    b_tb = s2.shapes.add_textbox(Inches(6.10), Inches(5.30), Inches(3.10), Inches(1.10))
    btf = b_tb.text_frame
    btf.word_wrap = True
    btf.margin_left = btf.margin_top = btf.margin_right = btf.margin_bottom = 0
    bp1 = btf.paragraphs[0]
    bp1.text = "The Assistive Necessity:"
    bp1.font.name = FONT_FAMILY
    bp1.font.size = Pt(11)
    bp1.font.bold = True
    bp1.font.color.rgb = COLOR_RED
    bp2 = btf.add_paragraph()
    bp2.text = "An ambient agent must listen continuously, identify student name callouts, formulate grounded answers in <4s, and sync notes to Notion."
    bp2.font.name = FONT_FAMILY
    bp2.font.size = Pt(10)
    bp2.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 3: Motivation and Need of the Project
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Motivation and Need of the Project", 3)
    
    # 4 Cards Layout
    cards_data = [
        ("Cognitive Overload & Multitasking", 
         "Students in virtual classes face extreme cognitive strain balancing real-time auditory intake, visual slide comprehension, manual note-taking, and active chat participation.\n\nConstant task switching leads to information omission, mental exhaustion, and degraded concept retention.",
         COLOR_RED, Inches(0.50), Inches(1.90)),
        ("Ephemeral Academic Knowledge", 
         "Synchronous virtual lectures generate thousands of words of spoken instruction that disappear the instant the call ends.\n\nWithout an automated capture and structuring pipeline, critical insights, derivations, and verbal Q&A exchanges are permanently lost to the learner.",
         COLOR_BLUE, Inches(5.10), Inches(1.90)),
        ("Inadequacy of Traditional AI Assistants", 
         "Current flagship AI platforms (ChatGPT, Claude, Copilot) are completely blind and deaf to the live online classroom environment.\n\nThey require manual copy-pasting of questions into external browser tabs, a process far too slow and distracting for spontaneous in-class interactions.",
         COLOR_AMBER, Inches(0.50), Inches(4.30)),
        ("Engineering Need & Solution Vision", 
         "A real-time autonomous student agent that runs locally on consumer hardware, captures lecture audio with zero feedback, decodes speech offline at $0 cloud cost, auto-replies to cold-calls in <4s, and persists notes to Notion.\n\nEliminates cognitive burden while preserving full academic agency.",
         COLOR_GREEN, Inches(5.10), Inches(4.30))
    ]
    
    for title, desc, border_c, left, top in cards_data:
        add_card(s3, left, top, Inches(4.40), Inches(2.20), fill_color=COLOR_WHITE, border_color=border_c, border_width=Pt(1.5))
        tb = s3.shapes.add_textbox(left + Inches(0.20), top + Inches(0.15), Inches(4.00), Inches(1.90))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(13)
        p0.font.bold = True
        p0.font.color.rgb = border_c
        p0.space_after = Pt(6)
        
        p1 = tf.add_paragraph()
        p1.text = desc
        p1.font.name = FONT_FAMILY
        p1.font.size = Pt(10.5)
        p1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 4: Objectives
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Objectives", 4)
    
    obj_cards = [
        ("1", "Zero-Cloud-Cost Ambient Listening & ASR", 
         [("Virtual Audio Conduit: ", "Intercept Google Meet speaker audio via VB-Audio Virtual Cable as a clean 16 kHz mono 16-bit PCM stream without microphone feedback loops."),
          ("Offline Speech Recognition: ", "Deploy VOSK toolkit with Kaldi WFST acoustic modeling locally on CPU, ensuring continuous monitoring with sub-500ms latency and $0 cloud fees."),
          ("Noise Filtering: ", "Apply SciPy Butterworth 4th-order bandpass filter (300 Hz–3,400 Hz) to eliminate 50 Hz hum and enhance speech formants.")],
         COLOR_RED, Inches(0.50)),
        ("2", "Contextual Trigger & Autonomous In-Call Chat", 
         [("Sliding Context Deque: ", "Maintain an in-memory 50-word rolling history to preserve discourse flow, dynamically extracting 15 words upon trigger detection."),
          ("Fuzzy Phonetic Matching: ", "Utilize normalized Levenshtein distance (Sim >= 0.80) and acoustic confidence gating (>= 0.70) with lookahead to detect student name callouts."),
          ("Grounded Q&A Dispatch: ", "Query OpenAI GPT-4o-mini for grounded responses and inject into Google Meet chat via Selenium DOM automation in < 4.0 seconds.")],
         COLOR_BLUE, Inches(3.60)),
        ("3", "Post-Session Synthesis & Notion Persistence", 
         [("Lossless Serialization: ", "Compile all raw lecture PCM audio frames into a master debug_audio.wav file upon meeting termination."),
          ("Whisper-1 Master Transcript: ", "Execute asynchronous batch transcription via OpenAI Whisper-1 API to yield verbatim, timestamped lecture transcripts."),
          ("Structured Notion Sync: ", "Synthesize hierarchical study guides (Topic, Key Concepts, Definitions) and export formatted block trees to Notion via REST API.")],
         COLOR_GREEN, Inches(6.70))
    ]
    
    for num_str, title, bullets, color, left in obj_cards:
        add_card(s4, left, Inches(2.00), Inches(2.80), Inches(4.65), fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        add_badge(s4, left + Inches(1.15), Inches(2.15), Inches(0.50), num_str, color=color)
        
        tb = s4.shapes.add_textbox(left + Inches(0.15), Inches(2.75), Inches(2.50), Inches(0.70))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(12.5)
        p0.font.bold = True
        p0.font.color.rgb = color
        
        b_tb = s4.shapes.add_textbox(left + Inches(0.15), Inches(3.50), Inches(2.50), Inches(3.05))
        btf = b_tb.text_frame
        btf.word_wrap = True
        btf.margin_left = btf.margin_top = btf.margin_right = btf.margin_bottom = 0
        for b_prefix, b_text in bullets:
            add_bullet_point(btf, b_prefix, b_text, font_size=Pt(10), space_after=Pt(6))

    # =========================================================================
    # SLIDE 5: Scope of the Project
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Scope of the Project", 5)
    
    scope_data = [
        ("Client Automation & Meet Interface", 
         [("Target Platform: ", "Google Meet web client accessed via Selenium Chrome WebDriver on Windows 10/11 consumer PCs."),
          ("Modal Bypass: ", "Automated arguments (--use-fake-ui-for-media-stream) suppress camera/microphone permission prompts."),
          ("Local State Broker: ", "Lightweight Flask HTTP server (localhost:5000) orchestrates session state, thread lifecycle, and live logs.")],
         COLOR_RED, Inches(0.50), Inches(1.90)),
        ("Hardware Audio Routing & DSP Conditioning", 
         [("Virtual Conduit: ", "Unidirectional loopback routing using VB-Audio Virtual Cable (CABLE Output -> PyAudio stream)."),
          ("Format Standards: ", "16 kHz sampling rate, 16-bit mono PCM chunks (4,000 bytes / 125ms non-blocking reads)."),
          ("SciPy DSP Pipeline: ", "4th-order digital Butterworth bandpass filter (300 Hz–3,400 Hz) eliminates mains hum and loopback echo.")],
         COLOR_BLUE, Inches(5.10), Inches(1.90)),
        ("Edge Speech & Cloud Intelligence", 
         [("Local Offline ASR: ", "VOSK 0.3.45 with Kaldi TDNN acoustic model running on CPU without GPU acceleration ($0 cost)."),
          ("Sliding Context Deque: ", "Thread-safe collections.deque(maxlen=50) isolates recent 15-word context for prompt grounding."),
          ("Frontier Models: ", "OpenAI GPT-4o-mini for sub-2s chat reasoning; OpenAI Whisper-1 for master post-lecture transcription.")],
         COLOR_AMBER, Inches(0.50), Inches(4.30)),
        ("Knowledge Base & Scope Boundaries", 
         [("Notion REST Integration: ", "Automated schema-compliant block trees exported to user's private cloud workspace via Notion API."),
          ("Strict Safety Boundaries: ", "Educational co-pilot and proxy; no adversarial exam evasion; credentials isolated in local .env."),
          ("Hardware Feasibility: ", "Engineered to run seamlessly on standard student laptops (8GB RAM, integrated graphics).")],
         COLOR_GREEN, Inches(5.10), Inches(4.30))
    ]
    
    for title, bullets, color, left, top in scope_data:
        add_card(s5, left, top, Inches(4.40), Inches(2.20), fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        tb = s5.shapes.add_textbox(left + Inches(0.20), top + Inches(0.12), Inches(4.00), Inches(2.00))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(13)
        p0.font.bold = True
        p0.font.color.rgb = color
        p0.space_after = Pt(6)
        for b_prefix, b_text in bullets:
            add_bullet_point(tf, b_prefix, b_text, font_size=Pt(10), space_after=Pt(4))

    # =========================================================================
    # SLIDE 6: Literature Review
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Literature Review", 6)
    
    table_shape = s6.shapes.add_table(7, 4, Inches(0.40), Inches(1.85), Inches(9.20), Inches(4.60))
    table = table_shape.table
    col_w = [Inches(1.80), Inches(2.70), Inches(2.20), Inches(2.50)]
    
    lit_data = [
        ["Paper / Source", "Key Contribution", "Limitation", "How AI Student Agent Addresses It"],
        ["Radford et al. (2023)\n[OpenAI Whisper]", "Weakly supervised Transformer ASR; human-level accuracy across accents.", "High streaming latency and 25 MB file payload boundary.", "Employed for post-session master notes; lightweight VOSK used for live speech."],
        ["Povey et al. (2011) /\nAlpha Cephei (2026)", "Kaldi WFST toolkit & VOSK: sub-500ms local ASR on consumer CPUs.", "Phonetic ambiguity on proper names and homophones.", "Two-stage lookahead + normalized Levenshtein distance string matching."],
        ["Dewan et al. (2023)", "End-to-end verbatim meeting transcription architecture.", "Virtual loopback echo and microphone feedback corruption.", "Isolated unidirectional VB-Audio Cable conduit eliminating virtual loopback echo."],
        ["Laskar et al. (2023)", "Role-directed LLM prompting for meeting summarization.", "Lacks real-time in-call grounding; context limits on long calls.", "Rolling 50-word context deque feeds immediate lecture discourse to GPT-4o-mini."],
        ["Wollny et al. (2021) /\nTaneja et al. (2024)", "Taxonomy of educational chatbots (90% operate at passive Level 1).", "Passive web forms; deaf/blind to live audio interactions.", "Implements Level 3 autonomous agent acting directly inside Google Meet chat."],
        ["Bhardwaj & Kumar (2025)", "Robotic process automation in higher education administration.", "Rigid DOM selectors break during continuous Google web updates.", "Resilient multi-selector fallback chains anchored to immutable ARIA attributes."]
    ]
    
    for r_idx, row in enumerate(lit_data):
        for c_idx, val in enumerate(row):
            table.cell(r_idx, c_idx).text = val
    style_table(table, col_w)
    
    fn_tb = s6.shapes.add_textbox(Inches(0.40), Inches(6.55), Inches(9.20), Inches(0.35))
    fnp = fn_tb.text_frame.paragraphs[0]
    fnp.text = "Survey of 18 peer-reviewed publications (2011–2026); full references detailed on final slides."
    fnp.font.name = FONT_FAMILY
    fnp.font.size = Pt(10)
    fnp.font.italic = True
    fnp.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 7: Research Gaps Addressed
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Research Gaps Addressed", 7)
    
    gaps_data = [
        ("Passive EdTech Tools", "Event-driven autonomous agent acts proactively inside Google Meet chat"),
        ("Costly Cloud Speech Streaming", "100% free offline CPU decoding via VOSK ($0/hr vs $23/mo cloud streaming)"),
        ("Virtual Audio Feedback & Echo", "Isolated unidirectional VB-Audio Virtual Cable conduit prevents feedback loops"),
        ("Phonetic Ambiguity on Names", "Two-stage lookahead ('just in' -> 'Justin') + normalized Levenshtein distance"),
        ("Ephemeral Meeting Discourse", "Lossless PCM audio serialization (debug_audio.wav) + Whisper-1 transcription"),
        ("Unstructured Raw Transcripts", "Hierarchical tripartite lecture note synthesis (Topic, Concepts, Definitions)"),
        ("Volatile Browser Web DOM", "Multi-selector XPath fallback chains anchored to immutable accessibility tags")
    ]
    
    for idx, (gap, sol) in enumerate(gaps_data):
        top_y = Inches(1.85 + idx * 0.68)
        # Gap card (left)
        add_card(s7, Inches(0.50), top_y, Inches(3.40), Inches(0.60), fill_color=COLOR_LIGHT_RED, border_color=COLOR_RED)
        gtb = s7.shapes.add_textbox(Inches(0.60), top_y + Inches(0.10), Inches(3.20), Inches(0.40))
        gtf = gtb.text_frame
        gtf.word_wrap = True
        gtf.margin_left = gtf.margin_top = gtf.margin_right = gtf.margin_bottom = 0
        gp = gtf.paragraphs[0]
        gp.text = gap
        gp.font.name = FONT_FAMILY
        gp.font.size = Pt(11)
        gp.font.bold = True
        gp.font.color.rgb = COLOR_RED
        
        # Arrow shape / badge
        arrow = s7.shapes.add_textbox(Inches(3.95), top_y + Inches(0.08), Inches(0.40), Inches(0.40))
        ap = arrow.text_frame.paragraphs[0]
        ap.text = "->"
        ap.font.name = FONT_FAMILY
        ap.font.size = Pt(14)
        ap.font.bold = True
        ap.font.color.rgb = COLOR_MUTED
        ap.alignment = PP_ALIGN.CENTER
        
        # Solution card (right)
        add_card(s7, Inches(4.40), top_y, Inches(5.10), Inches(0.60), fill_color=COLOR_WHITE, border_color=COLOR_GREEN)
        stb = s7.shapes.add_textbox(Inches(4.55), top_y + Inches(0.08), Inches(4.85), Inches(0.45))
        stf = stb.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
        sp = stf.paragraphs[0]
        sp.text = sol
        sp.font.name = FONT_FAMILY
        sp.font.size = Pt(10.5)
        sp.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 8: System Architecture
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "System Architecture", 8)
    
    arch_img = 'extracted_media/image2.jpg'
    if os.path.exists(arch_img):
        add_card(s8, Inches(0.45), Inches(1.85), Inches(9.10), Inches(3.60), fill_color=COLOR_WHITE, border_color=COLOR_BORDER)
        s8.shapes.add_picture(arch_img, Inches(0.50), Inches(1.90), Inches(9.00), Inches(3.50))
    
    # 3 Bullet Summary Box below diagram
    desc_tb = s8.shapes.add_textbox(Inches(0.50), Inches(5.55), Inches(9.00), Inches(1.30))
    dtf = desc_tb.text_frame
    dtf.word_wrap = True
    dtf.margin_left = dtf.margin_top = dtf.margin_right = dtf.margin_bottom = 0
    add_bullet_point(dtf, "Layers 1 & 2 (Client & Audio):", "Selenium Chrome driver bypasses media prompts; VB-Audio Cable isolates 16 kHz mono stream; SciPy Butterworth bandpass filter suppresses noise.", font_size=Pt(10.5), space_after=Pt(4))
    add_bullet_point(dtf, "Layers 3 & 4 (Perception & Execution):", "Offline VOSK Kaldi decodes speech locally on CPU; 50-word context deque buffers discourse; Levenshtein detector triggers GPT-4o-mini; Selenium types chat reply in < 4s.", font_size=Pt(10.5), space_after=Pt(4))
    add_bullet_point(dtf, "Layer 5 (Knowledge Base Persistence):", "Accumulates lossless PCM frames to debug_audio.wav; Whisper-1 transcribes full lecture; GPT-4o-mini formats study notes; Notion REST API publishes blocks to database.", font_size=Pt(10.5), space_after=Pt(2))

    # =========================================================================
    # SLIDE 9: Proposed Methodology: End-to-End Workflow
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Proposed Methodology: End-to-End Workflow", 9)
    
    steps = [
        ("1. Join Meet", "Selenium Chrome\nbypasses modals"),
        ("2. Ingest Audio", "VB-Audio Cable\n16 kHz mono PCM"),
        ("3. DSP Filter", "SciPy Butterworth\n300-3400 Hz band"),
        ("4. Offline ASR", "VOSK Kaldi WFST\non host CPU (<0.5s)"),
        ("5. Trigger & Q&A", "Levenshtein sim + \nGPT-4o-mini reason"),
        ("6. Post Notes", "Whisper transcript\n-> Notion API sync")
    ]
    
    step_w = Inches(1.38)
    gap_w = Inches(0.14)
    start_x = Inches(0.50)
    
    for i, (title, sub) in enumerate(steps):
        x = start_x + i * (step_w + gap_w)
        add_card(s9, x, Inches(1.90), step_w, Inches(1.25), fill_color=COLOR_WHITE, border_color=COLOR_RED if i in [0, 3, 5] else COLOR_BLUE, border_width=Pt(1.5))
        stb = s9.shapes.add_textbox(x + Inches(0.05), Inches(1.95), step_w - Inches(0.10), Inches(1.15))
        stf = stb.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
        p0 = stf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = COLOR_RED if i in [0, 3, 5] else COLOR_BLUE
        p1 = stf.add_paragraph()
        p1.alignment = PP_ALIGN.CENTER
        p1.text = sub
        p1.font.name = FONT_FAMILY
        p1.font.size = Pt(9)
        p1.font.color.rgb = COLOR_MUTED

    # Two Major Parallel Workflows
    # Left Box: Real-Time In-Call Loop
    add_card(s9, Inches(0.50), Inches(3.35), Inches(4.35), Inches(3.20), fill_color=COLOR_WHITE, border_color=COLOR_RED, border_width=Pt(1.5))
    ltb = s9.shapes.add_textbox(Inches(0.65), Inches(3.45), Inches(4.05), Inches(3.00))
    ltf = ltb.text_frame
    ltf.word_wrap = True
    ltf.margin_left = ltf.margin_top = ltf.margin_right = ltf.margin_bottom = 0
    lp0 = ltf.paragraphs[0]
    lp0.text = "Real-Time In-Call Turn-Taking Loop (< 3.5s)"
    lp0.font.name = FONT_FAMILY
    lp0.font.size = Pt(13)
    lp0.font.bold = True
    lp0.font.color.rgb = COLOR_RED
    lp0.space_after = Pt(6)
    add_bullet_point(ltf, "Audio Ingestion: ", "PyAudio streams non-blocking 4,000-byte chunks (125ms) into conditioned deque.", font_size=Pt(10), space_after=Pt(4))
    add_bullet_point(ltf, "Local Decoding: ", "VOSK processes 96,000-byte buffers (3.0s) in sub-500ms without cloud latency.", font_size=Pt(10), space_after=Pt(4))
    add_bullet_point(ltf, "Discourse Ring Buffer: ", "Rolling 50-word deque dynamically extracts recent 15-word active context.", font_size=Pt(10), space_after=Pt(4))
    add_bullet_point(ltf, "Trigger & Reasoning: ", "Levenshtein check (Sim >= 0.80) fires GPT-4o-mini; Selenium types reply in Meet chat.", font_size=Pt(10), space_after=Pt(2))

    # Right Box: Post-Session Knowledge Loop
    add_card(s9, Inches(5.15), Inches(3.35), Inches(4.35), Inches(3.20), fill_color=COLOR_WHITE, border_color=COLOR_BLUE, border_width=Pt(1.5))
    rtb = s9.shapes.add_textbox(Inches(5.30), Inches(3.45), Inches(4.05), Inches(3.00))
    rtf = rtb.text_frame
    rtf.word_wrap = True
    rtf.margin_left = rtf.margin_top = rtf.margin_right = rtf.margin_bottom = 0
    rp0 = rtf.paragraphs[0]
    rp0.text = "Post-Session Synthesis & Notion Sync Loop"
    rp0.font.name = FONT_FAMILY
    rp0.font.size = Pt(13)
    rp0.font.bold = True
    rp0.font.color.rgb = COLOR_BLUE
    rp0.space_after = Pt(6)
    add_bullet_point(rtf, "Lossless WAV Dump: ", "Master bytearray serializes cleanly to 16 kHz 16-bit mono debug_audio.wav.", font_size=Pt(10), space_after=Pt(4))
    add_bullet_point(rtf, "Whisper Transcription: ", "OpenAI Whisper-1 API outputs timestamped master transcript with jargon accuracy.", font_size=Pt(10), space_after=Pt(4))
    add_bullet_point(rtf, "Hierarchical Notes: ", "GPT-4o-mini transforms text into structured outline (# Topic, ## Concepts, Definitions).", font_size=Pt(10), space_after=Pt(4))
    add_bullet_point(rtf, "Notion Schema Adapter: ", "Notion SDK maps markdown into blocks and creates page in student database.", font_size=Pt(10), space_after=Pt(2))

    # =========================================================================
    # SLIDE 10: Module 1: Hardware Audio Routing & DSP Conditioning
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Module 1: Audio Capture, Loopback & DSP Conditioning", 10)
    
    # Left Table: Audio Ingestion Parameters
    tbl_shape10 = s10.shapes.add_table(6, 3, Inches(0.50), Inches(1.90), Inches(5.20), Inches(4.50))
    tbl10 = tbl_shape10.table
    w10 = [Inches(1.80), Inches(1.20), Inches(2.20)]
    t10_data = [
        ["Parameter Name", "Value", "Engineering Purpose"],
        ["Sampling Rate (R)", "16,000 Hz", "Standard Nyquist rate capturing human speech up to 8 kHz"],
        ["Channel Count (C)", "1 (Mono)", "Single-channel mono stream required by Kaldi acoustic model"],
        ["Bit Depth (Sb)", "16 Bits", "Provides 96 dB dynamic range representing amplitude accurately"],
        ["Chunk Read Size", "4,000 Bytes", "125 ms non-blocking I/O read ensuring real-time streaming"],
        ["Buffer Threshold", "96,000 Bytes", "Accumulates exactly 3.0s audio before feeding VOSK engine"]
    ]
    for r_idx, row in enumerate(t10_data):
        for c_idx, val in enumerate(row):
            tbl10.cell(r_idx, c_idx).text = val
    style_table(tbl10, w10)
    
    # Right Cards: Audio Loopback & Butterworth Filter
    add_card(s10, Inches(5.90), Inches(1.90), Inches(3.60), Inches(2.15), fill_color=COLOR_WHITE, border_color=COLOR_RED, border_width=Pt(1.5))
    tb10_r1 = s10.shapes.add_textbox(Inches(6.05), Inches(2.00), Inches(3.30), Inches(1.95))
    tf10_r1 = tb10_r1.text_frame
    tf10_r1.word_wrap = True
    tf10_r1.margin_left = tf10_r1.margin_top = tf10_r1.margin_right = tf10_r1.margin_bottom = 0
    p = tf10_r1.paragraphs[0]
    p.text = "Unidirectional VB-Cable Loopback"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(4)
    add_bullet_point(tf10_r1, "Echo Elimination: ", "Standard Windows 'Stereo Mix' leaked speaker audio back into microphone, causing infinite reverb and double tokens ('Justin Justin').", font_size=Pt(9.5), space_after=Pt(3))
    add_bullet_point(tf10_r1, "Virtual Bridge: ", "VB-Audio Cable creates an isolated, one-way conduit directly from Google Meet to PyAudio.", font_size=Pt(9.5), space_after=Pt(2))

    add_card(s10, Inches(5.90), Inches(4.25), Inches(3.60), Inches(2.15), fill_color=COLOR_WHITE, border_color=COLOR_BLUE, border_width=Pt(1.5))
    tb10_r2 = s10.shapes.add_textbox(Inches(6.05), Inches(4.35), Inches(3.30), Inches(1.95))
    tf10_r2 = tb10_r2.text_frame
    tf10_r2.word_wrap = True
    tf10_r2.margin_left = tf10_r2.margin_top = tf10_r2.margin_right = tf10_r2.margin_bottom = 0
    p = tf10_r2.paragraphs[0]
    p.text = "SciPy 4th-Order Butterworth Filter"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE
    p.space_after = Pt(4)
    add_bullet_point(tf10_r2, "Speech Formant Band: ", "Digital bandpass coefficients (b, a) configured with cutoffs at 300 Hz and 3,400 Hz.", font_size=Pt(9.5), space_after=Pt(3))
    add_bullet_point(tf10_r2, "Noise Suppression: ", "Filters out 50/60 Hz electrical mains hum and high-frequency hiss, boosting VOSK decoding confidence.", font_size=Pt(9.5), space_after=Pt(2))

    # =========================================================================
    # SLIDE 11: Module 2: Offline ASR & Sliding Context Deque
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "Module 2: Offline ASR & Sliding Context Deque", 11)
    
    # Left: VOSK Kaldi Engine Specs
    add_card(s11, Inches(0.50), Inches(1.90), Inches(4.35), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_RED, border_width=Pt(1.5))
    tb11_l = s11.shapes.add_textbox(Inches(0.70), Inches(2.05), Inches(3.95), Inches(4.40))
    tf11_l = tb11_l.text_frame
    tf11_l.word_wrap = True
    tf11_l.margin_left = tf11_l.margin_top = tf11_l.margin_right = tf11_l.margin_bottom = 0
    p = tf11_l.paragraphs[0]
    p.text = "VOSK Kaldi Offline Acoustic Engine"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(6)
    add_bullet_point(tf11_l, "Acoustic Topology: ", "Factorized Time-Delay Neural Network (F-TDNN) with 40-dimensional MFCCs and pitch feature extraction.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf11_l, "WFST Dynamic Decoding: ", "Kaldi Online2 Weighted Finite-State Transducer search graph covering ~128,000 English vocabulary words.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf11_l, "100% Local CPU Execution: ", "Average decoding latency is 350–480 ms on standard quad-core Intel i5 (RTF < 0.25). Zero cloud API reliance.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf11_l, "Confidence Scored Tokens: ", "Outputs finalized JSON containing token-level acoustic probabilities for trigger verification.", font_size=Pt(10), space_after=Pt(4))

    # Right: Sliding Context Deque
    add_card(s11, Inches(5.15), Inches(1.90), Inches(4.35), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_BLUE, border_width=Pt(1.5))
    tb11_r = s11.shapes.add_textbox(Inches(5.35), Inches(2.05), Inches(3.95), Inches(4.40))
    tf11_r = tb11_r.text_frame
    tf11_r.word_wrap = True
    tf11_r.margin_left = tf11_r.margin_top = tf11_r.margin_right = tf11_r.margin_bottom = 0
    p = tf11_r.paragraphs[0]
    p.text = "Sliding Context Window Management"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE
    p.space_after = Pt(6)
    add_bullet_point(tf11_r, "In-Memory Ring Buffer: ", "Python collections.deque(maxlen=50) maintains a continuous rolling history of the last 50 transcribed lecture words.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf11_r, "Discourse Preservation: ", "Automatically purges stale speech as new words arrive, retaining complete teacher sentence context without RAM bloat.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf11_r, "Dynamic Slice Extraction: ", "Upon verified trigger detection, isolates the most recent 15–20 words to ground prompt construction.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf11_r, "Hallucination Defense: ", "Grounding LLM reasoning strictly in the extracted discourse prevents generic, unanchored hallucinations.", font_size=Pt(10), space_after=Pt(4))

    # =========================================================================
    # SLIDE 12: Module 3: Fuzzy Phonetic Trigger & Intent Detection
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_header(s12, "Module 3: Fuzzy Phonetic Trigger & Intent Detection", 12)
    
    # Left Card: Trigger Logic & Formula
    add_card(s12, Inches(0.50), Inches(1.90), Inches(4.35), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_RED, border_width=Pt(1.5))
    tb12_l = s12.shapes.add_textbox(Inches(0.70), Inches(2.05), Inches(3.95), Inches(4.40))
    tf12_l = tb12_l.text_frame
    tf12_l.word_wrap = True
    tf12_l.margin_left = tf12_l.margin_top = tf12_l.margin_right = tf12_l.margin_bottom = 0
    p = tf12_l.paragraphs[0]
    p.text = "Multi-Tier Trigger Detection Mechanism"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(6)
    add_bullet_point(tf12_l, "Acoustic Confidence Gating: ", "Rejects all tokens with confidence score < 0.70 to eliminate background chatter and acoustic hallucinations.", font_size=Pt(10), space_after=Pt(5))
    add_bullet_point(tf12_l, "Normalized Levenshtein Metric: ", "Computes Sim(w_i, S_name) = 1 - (Lev(w_i, S_name) / max(len(w_i), len(S_name))). Threshold >= 0.80.", font_size=Pt(10), space_after=Pt(5))
    add_bullet_point(tf12_l, "Compound Token Lookahead: ", "VOSK frequently splits 'Justin' into two tokens: 'just' + 'in'. The lookahead buffer inspects w_i == 'just' and w_{i+1} == 'in'.", font_size=Pt(10), space_after=Pt(5))
    add_bullet_point(tf12_l, "Empirical Validation: ", "Compound lookahead raised cold-call detection from 60% to 100% in controlled classroom trials.", font_size=Pt(10), space_after=Pt(2))

    # Right Card: Intent Classification & Answering
    add_card(s12, Inches(5.15), Inches(1.90), Inches(4.35), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_GREEN, border_width=Pt(1.5))
    tb12_r = s12.shapes.add_textbox(Inches(5.35), Inches(2.05), Inches(3.95), Inches(4.40))
    tf12_r = tb12_r.text_frame
    tf12_r.word_wrap = True
    tf12_r.margin_left = tf12_r.margin_top = tf12_r.margin_right = tf12_r.margin_bottom = 0
    p = tf12_r.paragraphs[0]
    p.text = "Intent Classification & Response Synthesis"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN
    p.space_after = Pt(6)
    add_bullet_point(tf12_r, "Intent Branch 1: Oral Attendance: ", "If context contains roll-call indicators ('present', 'attendance', roll number), immediately outputs: 'Present, Professor. I am attending online.'", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf12_r, "Intent Branch 2: Technical Question: ", "If question words detected ('how', 'why', 'what', 'explain'), builds bounded prompt for GPT-4o-mini with the extracted 15-word lecture slice.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf12_r, "Concise Prompt Constraints: ", "System prompt enforces: 'Provide a direct, technically accurate 2-sentence response grounded strictly in the lecture context.'", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf12_r, "Sub-2s Cloud Inference: ", "GPT-4o-mini returns answer in 1.85s–2.65s, within natural conversation turn-taking limits.", font_size=Pt(10), space_after=Pt(2))

    # =========================================================================
    # SLIDE 13: Module 4: Browser Automation & In-Call Chat Injection
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_header(s13, "Module 4: Browser Automation & In-Call Chat Injection", 13)
    
    # Left Card: DOM Resilience
    add_card(s13, Inches(0.50), Inches(1.90), Inches(4.35), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_BLUE, border_width=Pt(1.5))
    tb13_l = s13.shapes.add_textbox(Inches(0.70), Inches(2.05), Inches(3.95), Inches(4.40))
    tf13_l = tb13_l.text_frame
    tf13_l.word_wrap = True
    tf13_l.margin_left = tf13_l.margin_top = tf13_l.margin_right = tf13_l.margin_bottom = 0
    p = tf13_l.paragraphs[0]
    p.text = "Google Meet DOM Mutation Resilience"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE
    p.space_after = Pt(6)
    add_bullet_point(tf13_l, "Obfuscated Class Names: ", "Google's CI/CD dynamically obfuscates CSS class names, causing standard static Selenium XPath scripts to fail.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf13_l, "Accessibility Attribute Anchoring: ", "Automation anchors exclusively to immutable accessibility tags: @aria-label, @aria-pressed, and role='button'.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf13_l, "Multi-Selector Fallback Chains: ", "Hierarchical fallback selectors guarantee button discovery even if interface layouts shift between updates.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf13_l, "Dynamic WebDriverWait Polling: ", "Explicit polling waits (up to 5s) ensure elements are interactable before dispatching actions.", font_size=Pt(10), space_after=Pt(2))

    # Right Card: Chat Injection
    add_card(s13, Inches(5.15), Inches(1.90), Inches(4.35), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_AMBER, border_width=Pt(1.5))
    tb13_r = s13.shapes.add_textbox(Inches(5.35), Inches(2.05), Inches(3.95), Inches(4.40))
    tf13_r = tb13_r.text_frame
    tf13_r.word_wrap = True
    tf13_r.margin_left = tf13_r.margin_top = tf13_r.margin_right = tf13_r.margin_bottom = 0
    p = tf13_r.paragraphs[0]
    p.text = "In-Call Chat Injection Architecture"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_AMBER
    p.space_after = Pt(6)
    add_bullet_point(tf13_r, "Chat Sidebar Detection: ", "Inspects DOM for //div[@aria-label='Chat with everyone']. If closed, locates toggle button and clicks.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf13_r, "Target Input Location: ", "Finds chat textarea via //textarea[@aria-label='Send a message to everyone'].", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf13_r, "Simulated Keystroke Dispatch: ", "Clears field, injects synthesized text via send_keys, and dispatches Keys.RETURN.", font_size=Pt(10), space_after=Pt(6))
    add_bullet_point(tf13_r, "Rapid Injection Timing: ", "DOM discovery and text submission executes in 650–920 ms, keeping total end-to-end response under 4.0s.", font_size=Pt(10), space_after=Pt(2))

    # =========================================================================
    # SLIDE 14: Module 5: Post-Session Synthesis & Notion Knowledge Base
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_header(s14, "Module 5: Post-Session Synthesis & Notion Knowledge Base", 14)
    
    # 4 Steps Cards Grid
    mod5_steps = [
        ("Lossless WAV Serialization", 
         "Throughout the lecture, PyAudio appends all raw 16 kHz 16-bit mono PCM chunks into an in-memory bytearray buffer.\n\nUpon session exit signal, the writer configures standard WAV headers and serializes debug_audio.wav to disk without data loss.",
         COLOR_RED, Inches(0.50), Inches(1.90)),
        ("Whisper-1 Master Transcription", 
         "The serialized audio is submitted to OpenAI's Whisper-1 API.\n\nGenerates a verbatim, punctuation-accurate master transcript with high fidelity on complex technical terminology and multi-speaker exchanges.",
         COLOR_BLUE, Inches(5.10), Inches(1.90)),
        ("Hierarchical Study Note Synthesis", 
         "GPT-4o-mini is prompted with a structured academic schema to transform raw speech disfluencies into concise study material:\n# Lecture Topic\n## Key Concepts & Principles\n* Bulleted definitions & formulas\n### Executive Summary",
         COLOR_AMBER, Inches(0.50), Inches(4.30)),
        ("Notion REST API Schema Adapter", 
         "Custom parser maps markdown syntax line-by-line into official Notion block objects (heading_1, heading_2, bulleted_list_item, callout).\n\nNotion-Client SDK issues client.pages.create() to publish a clean, searchable page into the student's Notion database.",
         COLOR_GREEN, Inches(5.10), Inches(4.30))
    ]
    for title, desc, color, left, top in mod5_steps:
        add_card(s14, left, top, Inches(4.40), Inches(2.20), fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        tb = s14.shapes.add_textbox(left + Inches(0.20), top + Inches(0.15), Inches(4.00), Inches(1.90))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(13)
        p0.font.bold = True
        p0.font.color.rgb = color
        p0.space_after = Pt(6)
        p1 = tf.add_paragraph()
        p1.text = desc
        p1.font.name = FONT_FAMILY
        p1.font.size = Pt(10)
        p1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 15: Tools, Technologies, Software and Hardware
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    add_header(s15, "Tools, Technologies, Software and Hardware", 15)
    
    table_shape15 = s15.shapes.add_table(10, 2, Inches(0.50), Inches(1.85), Inches(9.00), Inches(4.90))
    t15 = table_shape15.table
    w15 = [Inches(2.50), Inches(6.50)]
    
    t15_data = [
        ["Category", "Technologies & Specifications"],
        ["Client Automation", "Selenium WebDriver 4.49.0, Google Chrome (Headful / Headless)"],
        ["Audio Routing & Capture", "VB-Audio Cable v1.0.3.8, PyAudio 0.2.14, SoundDevice 0.5.5"],
        ["Digital Signal Processing", "SciPy 1.17.0 (Butterworth bandpass filter), NumPy 2.3.5"],
        ["Local Offline STT", "VOSK 0.3.45 + Kaldi WFST acoustic decoder (vosk-model-small-en-us-0.15)"],
        ["Frontier Cloud AI", "OpenAI GPT-4o-mini (chat reasoning), OpenAI Whisper-1 (batch transcription)"],
        ["Knowledge Persistence", "Notion REST API, Notion-Client SDK 3.1.0"],
        ["Internal State Broker", "Python 3.12+, Flask 3.1.3, Flask-CORS 6.0.5 (localhost:5000)"],
        ["Testing & Verification", "pytest, Boundary Value Analysis (BVA), Equivalence Partitioning (EPP)"],
        ["Hardware Environment", "Quad-Core Intel i5 / Ryzen 5, 8 GB DDR4 RAM, Realtek HD Audio, Zero GPU Required"]
    ]
    for r_idx, row in enumerate(t15_data):
        for c_idx, val in enumerate(row):
            t15.cell(r_idx, c_idx).text = val
    style_table(t15, w15)

    # =========================================================================
    # SLIDE 16: Database Design and Object-Oriented Modeling
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    add_header(s16, "Database Design and REST API", 16)
    
    # Left: Class Diagram picture
    class_img = 'extracted_media/image5.jpg'
    if os.path.exists(class_img):
        add_card(s16, Inches(0.45), Inches(1.85), Inches(4.45), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_BORDER)
        s16.shapes.add_picture(class_img, Inches(0.50), Inches(1.90), Inches(4.35), Inches(4.60))
    
    # Right: REST Endpoints Table & Domain Info
    tbl_shape16 = s16.shapes.add_table(6, 2, Inches(5.10), Inches(1.85), Inches(4.40), Inches(2.90))
    t16 = tbl_shape16.table
    w16 = [Inches(2.00), Inches(2.40)]
    t16_data = [
        ["Endpoint / Method", "Operational Purpose"],
        ["POST /api/session/start", "Spawns Chrome & initializes audio loopback"],
        ["GET /api/session/status", "Health monitoring of PyAudio & VOSK threads"],
        ["POST /api/session/stop", "Terminates session & triggers post-processing"],
        ["POST /api/chat/inject", "Direct chat injection fallback endpoint"],
        ["POST /api/notion/sync", "Exports parsed markdown blocks to Notion"]
    ]
    for r_idx, row in enumerate(t16_data):
        for c_idx, val in enumerate(row):
            t16.cell(r_idx, c_idx).text = val
    style_table(t16, w16)
    
    # Domain entities description below table
    add_card(s16, Inches(5.10), Inches(4.90), Inches(4.40), Inches(1.65), fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER)
    tb16_b = s16.shapes.add_textbox(Inches(5.25), Inches(5.00), Inches(4.10), Inches(1.45))
    tf16_b = tb16_b.text_frame
    tf16_b.word_wrap = True
    tf16_b.margin_left = tf16_b.margin_top = tf16_b.margin_right = tf16_b.margin_bottom = 0
    p = tf16_b.paragraphs[0]
    p.text = "Relational Domain Schema Architecture:"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK
    p.space_after = Pt(4)
    add_bullet_point(tf16_b, "Domain Entities: ", "User (1:N) ClassSession (1:1) AudioRecording (1:N) TranscriptChunk -> FullTranscript (1:1) LectureNote -> NotionExport.", font_size=Pt(9.5), space_after=Pt(3))
    add_bullet_point(tf16_b, "Decoupled Singleton Services: ", "Clean isolation between state entities and service orchestrators (SeleniumMeetService, VoskSpeechService, OpenAIService, NotionIntegrationService).", font_size=Pt(9.5), space_after=Pt(2))

    # =========================================================================
    # SLIDE 17: Implementation: Application Screens (1/2)
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    add_header(s17, "Implementation: Application Screens (1/2)", 17)
    
    dash_img = 'extracted_media/image9.jpg'
    if os.path.exists(dash_img):
        add_card(s17, Inches(0.45), Inches(1.85), Inches(9.10), Inches(3.60), fill_color=COLOR_WHITE, border_color=COLOR_BORDER)
        s17.shapes.add_picture(dash_img, Inches(0.50), Inches(1.90), Inches(9.00), Inches(3.50))
    
    # 5 Dashboard Features Callouts
    tb17_desc = s17.shapes.add_textbox(Inches(0.50), Inches(5.55), Inches(9.00), Inches(1.30))
    tf17 = tb17_desc.text_frame
    tf17.word_wrap = True
    tf17.margin_left = tf17.margin_top = tf17.margin_right = tf17.margin_bottom = 0
    add_bullet_point(tf17, "1. Active Session Monitor: ", "Continuous session timer (00:48:12) with real-time audio loopback health diagnostic.", font_size=Pt(10.5), space_after=Pt(3))
    add_bullet_point(tf17, "2. Real-Time Speech-to-Text: ", "Live acoustic waveform visualization and continuous stream of local VOSK transcribed tokens.", font_size=Pt(10.5), space_after=Pt(3))
    add_bullet_point(tf17, "3. Trigger Alert & Q&A Panel: ", "Displays captured instructor query ('Justin, can you explain how paging prevents external fragmentation?') and GPT-4o-mini answer auto-submitted to Meet chat in 3.6s.", font_size=Pt(10.5), space_after=Pt(2))

    # =========================================================================
    # SLIDE 18: Implementation: Application Screens (2/2)
    # =========================================================================
    s18 = prs.slides.add_slide(blank_layout)
    add_header(s18, "Implementation: Application Screens (2/2)", 18)
    
    meet_img = 'extracted_media/image10.png'
    if os.path.exists(meet_img):
        add_card(s18, Inches(0.45), Inches(1.85), Inches(9.10), Inches(3.60), fill_color=COLOR_WHITE, border_color=COLOR_BORDER)
        s18.shapes.add_picture(meet_img, Inches(0.50), Inches(1.90), Inches(9.00), Inches(3.50))
    
    tb18_desc = s18.shapes.add_textbox(Inches(0.50), Inches(5.55), Inches(9.00), Inches(1.30))
    tf18 = tb18_desc.text_frame
    tf18.word_wrap = True
    tf18.margin_left = tf18.margin_top = tf18.margin_right = tf18.margin_bottom = 0
    add_bullet_point(tf18, "Automated Google Meet In-Call Entry: ", "Selenium WebDriver launched Chrome, dismissed permission modals, entered user name, and successfully clicked 'Ask to join'.", font_size=Pt(10.5), space_after=Pt(3))
    add_bullet_point(tf18, "Autonomous In-Call Chat Injection: ", "Upon trigger recognition, Selenium located the chat input DOM element and typed the synthesized explanation in 0.8s.", font_size=Pt(10.5), space_after=Pt(3))
    add_bullet_point(tf18, "Verifiable Classroom Outcome: ", "The instructor accepted the technically grounded explanation of memory paging and fragmentation without realizing assistive intervention occurred.", font_size=Pt(10.5), space_after=Pt(2))

    # =========================================================================
    # SLIDE 19: Results: Latency & Performance Analysis
    # =========================================================================
    s19 = prs.slides.add_slide(blank_layout)
    add_header(s19, "Results: Latency and Resource Usage", 19)
    
    tbl_shape19 = s19.shapes.add_table(8, 4, Inches(0.50), Inches(1.85), Inches(9.00), Inches(3.70))
    t19 = tbl_shape19.table
    w19 = [Inches(2.50), Inches(2.20), Inches(2.10), Inches(2.20)]
    t19_data = [
        ["Processing Pipeline Stage", "Execution Subsystem", "Measured Latency", "Engineering Target"],
        ["Audio Ingestion & Chunking", "PyAudio & SoundDevice", "125 ms – 250 ms", "< 300 ms (Met)"],
        ["DSP Bandpass Filtering", "SciPy Butterworth Filter", "15 ms – 25 ms", "< 50 ms (Met)"],
        ["VOSK Offline Speech Decoding", "Kaldi Acoustic WFST Graph", "350 ms – 480 ms", "< 500 ms (Met)"],
        ["Trigger & Intent Detection", "Levenshtein Distance Engine", "8 ms – 15 ms", "< 20 ms (Met)"],
        ["Prompt Synthesis & OpenAI API", "OpenAI GPT-4o-mini", "1,850 ms – 2,650 ms", "< 3,000 ms (Met)"],
        ["Selenium DOM Chat Injection", "Selenium WebDriver Chrome", "650 ms – 920 ms", "< 1,200 ms (Met)"],
        ["Total End-to-End Response Time", "Full Pipeline (Audio -> Chat)", "3.20 s – 4.54 s", "3.0 s – 5.0 s (Met)"]
    ]
    for r_idx, row in enumerate(t19_data):
        for c_idx, val in enumerate(row):
            t19.cell(r_idx, c_idx).text = val
    style_table(t19, w19)
    
    # Resource metrics card below
    add_card(s19, Inches(0.50), Inches(5.70), Inches(9.00), Inches(0.95), fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER)
    tb19_b = s19.shapes.add_textbox(Inches(0.65), Inches(5.75), Inches(8.70), Inches(0.85))
    tf19_b = tb19_b.text_frame
    tf19_b.word_wrap = True
    tf19_b.margin_left = tf19_b.margin_top = tf19_b.margin_right = tf19_b.margin_bottom = 0
    p = tf19_b.paragraphs[0]
    p.text = "Host Hardware Resource Footprint Benchmarks:"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK
    p.space_after = Pt(3)
    add_bullet_point(tf19_b, "Memory (RAM): ", "1.18 GB – 1.34 GB total usage (well below 1.50 GB budget) | CPU Load: 18% – 26% on Intel i5 (leaves > 70% free) | Dedicated GPU: 0.0% (strictly software-only).", font_size=Pt(10), space_after=Pt(2))

    # =========================================================================
    # SLIDE 20: Results: Testing and Verification
    # =========================================================================
    s20 = prs.slides.add_slide(blank_layout)
    add_header(s20, "Results: Testing and Verification", 20)
    
    # 4 Stat Cards at top (exact layout matching MedIntel Slide 20)
    stats = [
        ("10 / 10", "Subsystems Fully Verified", COLOR_RED),
        ("100%", "Functional Tests Passed", COLOR_BLUE),
        ("3.6 s", "Average Live Response", COLOR_GREEN),
        ("0", "GPU Accelerators Needed", COLOR_AMBER)
    ]
    card_w = Inches(2.10)
    card_gap = Inches(0.20)
    for i, (val, label, color) in enumerate(stats):
        cx = Inches(0.50) + i * (card_w + card_gap)
        add_card(s20, cx, Inches(1.85), card_w, Inches(1.15), fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        stb = s20.shapes.add_textbox(cx, Inches(1.90), card_w, Inches(0.60))
        sp = stb.text_frame.paragraphs[0]
        sp.alignment = PP_ALIGN.CENTER
        sp.text = val
        sp.font.name = FONT_FAMILY
        sp.font.size = Pt(22)
        sp.font.bold = True
        sp.font.color.rgb = color
        
        ltb = s20.shapes.add_textbox(cx, Inches(2.50), card_w, Inches(0.45))
        lp = ltb.text_frame.paragraphs[0]
        lp.alignment = PP_ALIGN.CENTER
        lp.text = label
        lp.font.name = FONT_FAMILY
        lp.font.size = Pt(10)
        lp.font.color.rgb = COLOR_MUTED

    # Functional Verification Matrix Table (Table 4.1 from Black Book)
    tbl_shape20 = s20.shapes.add_table(6, 4, Inches(0.50), Inches(3.20), Inches(9.00), Inches(3.40))
    t20 = tbl_shape20.table
    w20 = [Inches(1.00), Inches(2.40), Inches(4.40), Inches(1.20)]
    t20_data = [
        ["Test ID", "Function Under Test", "Observed Empirical Result", "Status"],
        ["TC-01", "Automated Meet Access", "Chrome launches, dismisses permission modals, enters lobby cleanly", "PASS"],
        ["TC-02", "Virtual Audio Capture", "Clean 16 kHz mono 16-bit PCM captured via VB-Cable with zero echo", "PASS"],
        ["TC-03", "Offline VOSK STT", "Transcribes speech locally on CPU (<500 ms latency; ~22% CPU load)", "PASS"],
        ["TC-05", "Fuzzy Trigger Detection", "Compound lookahead resolves 'just in' to 'Justin' (Sim >= 0.80, Conf >= 0.70)", "PASS"],
        ["TC-07", "In-Call Chat Injection", "Selenium locates chat textarea and dispatches technical answer in 0.8s", "PASS"]
    ]
    for r_idx, row in enumerate(t20_data):
        for c_idx, val in enumerate(row):
            t20.cell(r_idx, c_idx).text = val
            if c_idx == 3 and r_idx > 0:
                p = t20.cell(r_idx, c_idx).text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                p.font.bold = True
                p.font.color.rgb = COLOR_GREEN
    style_table(t20, w20)

    # =========================================================================
    # SLIDE 21: Results: Economic Cost Analysis
    # =========================================================================
    s21 = prs.slides.add_slide(blank_layout)
    add_header(s21, "Results: Economic Cost Analysis", 21)
    
    # Cost Comparison Table (Table 4.3 from Black Book)
    tbl_shape21 = s21.shapes.add_table(4, 5, Inches(0.50), Inches(1.85), Inches(9.00), Inches(2.90))
    t21 = tbl_shape21.table
    w21 = [Inches(2.50), Inches(1.60), Inches(1.50), Inches(1.70), Inches(1.70)]
    t21_data = [
        ["Architecture Paradigm", "Continuous Listening", "In-Meeting Q&A", "Post-Meeting Notes", "Monthly Cost (60 Classes)"],
        ["Pure Cloud Streaming\n(Google STT + GPT-4)", "$0.360 / hr\n($0.006/min)", "$0.008\n(avg 2 queries)", "$0.015\n(audio + notes)", "$22.98 / month\n(per student)"],
        ["AI Student Agent Hybrid\n(Local VOSK + GPT-4o-mini)", "$0.000 / hr\n(100% Free on CPU)", "$0.001\n(GPT-4o-mini)", "$0.006\n(Whisper + notes)", "$0.42 / month\n(per student)"],
        ["Financial Savings &\nEfficiency Gain", "100.0% Reduction\nin Live Listening", "87.5% Reduction\nin Q&A Cost", "60.0% Reduction\nin Post-Notes", "98.17% Total Cost\nReduction (>98% Savings)"]
    ]
    for r_idx, row in enumerate(t21_data):
        for c_idx, val in enumerate(row):
            t21.cell(r_idx, c_idx).text = val
            if r_idx == 3:
                p = t21.cell(r_idx, c_idx).text_frame.paragraphs[0]
                p.font.bold = True
                p.font.color.rgb = COLOR_GREEN
    style_table(t21, w21)
    
    # Callout card below table
    add_card(s21, Inches(0.50), Inches(5.00), Inches(9.00), Inches(1.65), fill_color=COLOR_WHITE, border_color=COLOR_GREEN, border_width=Pt(1.5))
    tb21_b = s21.shapes.add_textbox(Inches(0.70), Inches(5.10), Inches(8.60), Inches(1.45))
    tf21_b = tb21_b.text_frame
    tf21_b.word_wrap = True
    tf21_b.margin_left = tf21_b.margin_top = tf21_b.margin_right = tf21_b.margin_bottom = 0
    p = tf21_b.paragraphs[0]
    p.text = "Key Economic & Architectural Takeaways:"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN
    p.space_after = Pt(4)
    add_bullet_point(tf21_b, "98.17% Operational Cost Reduction: ", "Offloading continuous acoustic listening to local CPU eliminates ongoing cloud ASR meter charges ($0.36/hr). Paid cloud APIs are invoked strictly on verified trigger events.", font_size=Pt(10), space_after=Pt(3))
    add_bullet_point(tf21_b, "Affordability for Students: ", "A full semester of 60 lecture hours costs less than $0.45 total on AI Student Agent, compared to ~$23.00 on pure cloud pipelines, proving large-scale accessibility.", font_size=Pt(10), space_after=Pt(2))

    # =========================================================================
    # SLIDE 22: Project Development Timeline (Gantt Chart)
    # =========================================================================
    s22 = prs.slides.add_slide(blank_layout)
    add_header(s22, "Project Development Timeline (Gantt Chart)", 22)
    
    gantt_img = 'extracted_media/image11.png'
    if os.path.exists(gantt_img):
        add_card(s22, Inches(0.45), Inches(1.85), Inches(9.10), Inches(3.60), fill_color=COLOR_WHITE, border_color=COLOR_BORDER)
        s22.shapes.add_picture(gantt_img, Inches(0.50), Inches(1.90), Inches(9.00), Inches(3.50))
    
    tb22_desc = s22.shapes.add_textbox(Inches(0.50), Inches(5.55), Inches(9.00), Inches(1.30))
    tf22 = tb22_desc.text_frame
    tf22.word_wrap = True
    tf22.margin_left = tf22.margin_top = tf22.margin_right = tf22.margin_bottom = 0
    add_bullet_point(tf22, "Phase 1 (Weeks 1–3): ", "Problem Formulation, Requirements Elicitation, Literature Survey (18 papers) and 5-Layer Architecture design.", font_size=Pt(10.5), space_after=Pt(3))
    add_bullet_point(tf22, "Phase 2 & 3 (Weeks 4–10): ", "Audio Loopback (VB-Cable), PyAudio streaming, SciPy filter, local VOSK setup, Selenium automation, and Levenshtein trigger engine.", font_size=Pt(10.5), space_after=Pt(3))
    add_bullet_point(tf22, "Phase 4 & 5 (Weeks 11–16): ", "Post-meeting notes (WAV serialization, Whisper-1, Notion REST API), stress testing, Black Book writing, and conference paper submission.", font_size=Pt(10.5), space_after=Pt(2))

    # =========================================================================
    # SLIDE 23: Advantages and Limitations
    # =========================================================================
    s23 = prs.slides.add_slide(blank_layout)
    add_header(s23, "Advantages and Limitations", 23)
    
    # Left: Advantages Card
    add_card(s23, Inches(0.50), Inches(1.90), Inches(4.35), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_GREEN, border_width=Pt(1.5))
    tb23_l = s23.shapes.add_textbox(Inches(0.70), Inches(2.05), Inches(3.95), Inches(4.40))
    tf23_l = tb23_l.text_frame
    tf23_l.word_wrap = True
    tf23_l.margin_left = tf23_l.margin_top = tf23_l.margin_right = tf23_l.margin_bottom = 0
    p = tf23_l.paragraphs[0]
    p.text = "System Advantages"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN
    p.space_after = Pt(6)
    add_bullet_point(tf23_l, "Cognitive Load Reduction: ", "Eliminates split-attention between note taking and listening, freeing mental bandwidth for conceptual learning.", font_size=Pt(9.5), space_after=Pt(4))
    add_bullet_point(tf23_l, "Zero-Cloud-Cost Listening: ", "Local CPU speech decoding via VOSK reduces monthly compute expenses by 98.17% ($0.42 vs $22.98).", font_size=Pt(9.5), space_after=Pt(4))
    add_bullet_point(tf23_l, "Phonetic Ambiguity Robustness: ", "Levenshtein distance + compound lookahead resolves homophones ('just in' -> 'Justin') with zero false activations.", font_size=Pt(9.5), space_after=Pt(4))
    add_bullet_point(tf23_l, "Contextually Grounded Answering: ", "Rolling 50-word context window prevents LLM hallucinations during in-class professor cold-calls.", font_size=Pt(9.5), space_after=Pt(4))
    add_bullet_point(tf23_l, "Automated Knowledge Persistence: ", "Synthesizes tripartite notes and syncs formatted blocks to Notion via REST API without student effort.", font_size=Pt(9.5), space_after=Pt(4))
    add_bullet_point(tf23_l, "Consumer Hardware Viability: ", "Operates smoothly on standard 8GB RAM laptops with zero dedicated GPU requirements.", font_size=Pt(9.5), space_after=Pt(2))

    # Right: Limitations Card
    add_card(s23, Inches(5.15), Inches(1.90), Inches(4.35), Inches(4.70), fill_color=COLOR_WHITE, border_color=COLOR_RED, border_width=Pt(1.5))
    tb23_r = s23.shapes.add_textbox(Inches(5.35), Inches(2.05), Inches(3.95), Inches(4.40))
    tf23_r = tb23_r.text_frame
    tf23_r.word_wrap = True
    tf23_r.margin_left = tf23_r.margin_top = tf23_r.margin_right = tf23_r.margin_bottom = 0
    p = tf23_r.paragraphs[0]
    p.text = "Operational Limitations"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(6)
    add_bullet_point(tf23_r, "Web DOM Selector Volatility: ", "Major UI redesigns by Google to Meet's DOM can temporarily affect XPath button clicks and chat textarea injection.", font_size=Pt(9.5), space_after=Pt(5))
    add_bullet_point(tf23_r, "Acoustic Noise Sensitivity: ", "VOSK local acoustic decoding accuracy degrades under severe ambient room noise or heavy VoIP packet loss.", font_size=Pt(9.5), space_after=Pt(5))
    add_bullet_point(tf23_r, "Cloud API Network Latency Jitter: ", "OpenAI API round-trips average 2.1s, but severe network congestion can push latency to 4.5s near conversational limits.", font_size=Pt(9.5), space_after=Pt(5))
    add_bullet_point(tf23_r, "Whisper File Upload Limit: ", "OpenAI Whisper-1 enforces a strict 25.0 MB payload limit (~12 minutes of 16 kHz WAV), requiring audio chunking for longer sessions.", font_size=Pt(9.5), space_after=Pt(5))
    add_bullet_point(tf23_r, "OS Virtual Audio Dependency: ", "Relies on Windows Core Audio & VB-Audio Cable; deployment on macOS or Linux requires native virtual daemons.", font_size=Pt(9.5), space_after=Pt(2))

    # =========================================================================
    # SLIDE 24: Conclusion and Future Scope
    # =========================================================================
    s24 = prs.slides.add_slide(blank_layout)
    add_header(s24, "Conclusion and Future Scope", 24)
    
    # Top Conclusion Card (matching MedIntel Slide 24)
    add_card(s24, Inches(0.50), Inches(1.85), Inches(9.00), Inches(1.50), fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER)
    tb24_top = s24.shapes.add_textbox(Inches(0.70), Inches(1.95), Inches(8.60), Inches(1.30))
    tf24_top = tb24_top.text_frame
    tf24_top.word_wrap = True
    tf24_top.margin_left = tf24_top.margin_top = tf24_top.margin_right = tf24_top.margin_bottom = 0
    p = tf24_top.paragraphs[0]
    p.text = "Project Conclusion:"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK
    p.space_after = Pt(4)
    cp = tf24_top.add_paragraph()
    cp.text = "All primary objectives were achieved: AI Student Agent was successfully developed and verified as an autonomous, multi-threaded assistant for Google Meet. By harmonizing edge acoustic decoding (VOSK/Kaldi) with frontier reasoning (GPT-4o-mini & Whisper), the system achieves sub-4s in-call response times and delivers 98.17% cloud cost reductions ($0.42/month) on standard consumer PC hardware without dedicated GPU acceleration."
    cp.font.name = FONT_FAMILY
    cp.font.size = Pt(10.5)
    cp.font.color.rgb = COLOR_DARK

    # Bottom 6 Future Scope Cards
    fs_cards = [
        ("Dynamic Audio Chunking", "Automate PyDub temporal segmentation to process multi-hour WAVs within Whisper's 25 MB limit.", COLOR_RED),
        ("Course-Grounded RAG", "Index course syllabus and textbook PDFs via ChromaDB / FAISS for deep factual grounding.", COLOR_BLUE),
        ("Synthetic Spoken Voice (TTS)", "Integrate Edge-TTS / ElevenLabs through a virtual microphone for spoken verbal participation.", COLOR_GREEN),
        ("Longitudinal Knowledge Graph", "Track misconception patterns over time to dynamically personalize review summaries.", COLOR_AMBER),
        ("Cross-Platform Conferencing", "Extend automation adapters to Zoom, MS Teams, Cisco Webex, macOS and Linux.", COLOR_BLUE),
        ("On-Device Quantized SLMs", "Deploy local quantized SLMs (Phi-3, Llama 3.2 3B) for completely offline edge reasoning.", COLOR_RED)
    ]
    
    fw = Inches(2.85)
    fh = Inches(1.40)
    for idx, (title, desc, color) in enumerate(fs_cards):
        col_i = idx % 3
        row_i = idx // 3
        cx = Inches(0.50) + col_i * (fw + Inches(0.22))
        cy = Inches(3.55) + row_i * (fh + Inches(0.18))
        
        add_card(s24, cx, cy, fw, fh, fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        tb = s24.shapes.add_textbox(cx + Inches(0.15), cy + Inches(0.12), fw - Inches(0.30), fh - Inches(0.20))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = color
        p0.space_after = Pt(3)
        p1 = tf.add_paragraph()
        p1.text = desc
        p1.font.name = FONT_FAMILY
        p1.font.size = Pt(9.5)
        p1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 25: Research Paper Submission
    # =========================================================================
    s25 = prs.slides.add_slide(blank_layout)
    add_header(s25, "Research Paper Submission", 25)
    
    # Manuscript metadata banner card
    add_card(s25, Inches(0.50), Inches(1.85), Inches(9.00), Inches(1.55), fill_color=COLOR_LIGHT_RED, border_color=COLOR_RED, border_width=Pt(1.5))
    tb25_top = s25.shapes.add_textbox(Inches(0.70), Inches(1.95), Inches(8.60), Inches(1.35))
    tf25_top = tb25_top.text_frame
    tf25_top.word_wrap = True
    tf25_top.margin_left = tf25_top.margin_top = tf25_top.margin_right = tf25_top.margin_bottom = 0
    p = tf25_top.paragraphs[0]
    p.text = "Scholarly Research Paper Prepared from Capstone Project Work"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(4)
    cp = tf25_top.add_paragraph()
    cp.text = "In accordance with academic guidelines of the Information Technology Department, MPSTME, SVKM’s NMIMS, a research manuscript documenting the software architecture, digital loopback routing, fuzzy phonetic acoustic confidence gating, and empirical latency benchmarks has been prepared and submitted for peer-reviewed publication."
    cp.font.name = FONT_FAMILY
    cp.font.size = Pt(10.5)
    cp.font.color.rgb = COLOR_DARK

    # Table of Submission Details (matching MedIntel Slide 25 Table)
    tbl_shape25 = s25.shapes.add_table(6, 2, Inches(0.50), Inches(3.60), Inches(9.00), Inches(3.00))
    t25 = tbl_shape25.table
    w25 = [Inches(2.50), Inches(6.50)]
    t25_data = [
        ["Item", "Details"],
        ["Paper Title", "AI Student Agent: An Autonomous Multi-Threaded Architecture for Real-Time Online Classroom Participation and Lecture Note Synthesis"],
        ["Target Conference", "International Conference on Artificial Intelligence and Emerging Technology 3.0 (AISUMMIT-2026)"],
        ["Conference Track", "Track 1.8 — AI and ML in Autonomous Systems"],
        ["Authors & Affiliation", "Aman Choudhary, Pranav Patil, Mohit Patil, Dr. Pravin Landge\nDepartment of Information Technology, MPSTME, SVKM’s NMIMS Shirpur Campus"],
        ["Manuscript Status", "Prepared and submitted for peer review"]
    ]
    for r_idx, row in enumerate(t25_data):
        for c_idx, val in enumerate(row):
            t25.cell(r_idx, c_idx).text = val
    style_table(t25, w25)

    # =========================================================================
    # SLIDE 26: References
    # =========================================================================
    s26 = prs.slides.add_slide(blank_layout)
    add_header(s26, "References", 26)
    
    # 2-Column layout for 12 IEEE references
    tb26_l = s26.shapes.add_textbox(Inches(0.50), Inches(1.85), Inches(4.35), Inches(4.80))
    tf26_l = tb26_l.text_frame
    tf26_l.word_wrap = True
    tf26_l.margin_left = tf26_l.margin_top = tf26_l.margin_right = tf26_l.margin_bottom = 0
    
    refs_left = [
        "[1] A. Radford, J. W. Kim, T. Xu, G. Brockman, C. McLeavey, and I. Sutskever, “Robust speech recognition via large-scale weak supervision,” in Proc. ICML, PMLR vol. 202, 2023.",
        "[2] D. Povey, A. Ghoshal, G. Boulianne, et al., “The Kaldi speech recognition toolkit,” in Proc. IEEE ASRU Workshop, 2011, pp. 1–4.",
        "[3] Alpha Cephei, “Vosk: Offline open-source speech recognition toolkit,” Alpha Cephei Inc., 2026. [Online]. Available: https://alphacephei.com/vosk/.",
        "[4] Y. Wang, C. S. Leow, A. Kobayashi, et al., “ExKaldi-RT: A real-time automatic speech recognition extension toolkit of Kaldi,” arXiv:2104.01384, 2021.",
        "[5] A. Dewan, M. Ziemski, H. Meylan, et al., “Developing automatic verbatim transcripts for international multilingual meetings,” arXiv:2309.15609, 2023.",
        "[6] M. T. R. Laskar, X.-Y. Fu, C. Chen, and S. Bhushan TN, “Building real-world meeting summarization systems using LLMs,” arXiv:2310.19233, 2023."
    ]
    for r in refs_left:
        p = tf26_l.add_paragraph() if len(tf26_l.paragraphs[0].text) > 0 else tf26_l.paragraphs[0]
        p.text = r
        p.font.name = FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK
        p.space_after = Pt(7)

    tb26_r = s26.shapes.add_textbox(Inches(5.15), Inches(1.85), Inches(4.35), Inches(4.80))
    tf26_r = tb26_r.text_frame
    tf26_r.word_wrap = True
    tf26_r.margin_left = tf26_r.margin_top = tf26_r.margin_right = tf26_r.margin_bottom = 0
    
    refs_right = [
        "[7] V. Rennard, G. Shang, J. Hunter, and M. Vazirgiannis, “Abstractive meeting summarization: A survey,” arXiv:2208.04163, 2022.",
        "[8] M. Zhong, D. Yin, T. Yu, et al., “QMSum: A new benchmark for query-based multi-domain meeting summarization,” in Proc. NAACL, 2021.",
        "[9] A. Madhavi, A. Chilakamarri, et al., “Automatic running notes generation from audio lecture: A systematic review,” LNNS, Springer, 2025.",
        "[10] K. Taneja, P. Maiti, S. Kakar, et al., “Jill Watson: A virtual teaching assistant powered by ChatGPT,” arXiv:2405.11070, 2024.",
        "[11] L. Labadze, M. Grigolia, and L. Machaidze, “Role of AI chatbots in education: Systematic literature review,” IJETHE, vol. 20, 2023.",
        "[12] OpenAI, “GPT-4o mini: Advancing cost-efficient intelligence,” OpenAI Technical Report, 2024. [Online]. Available: https://openai.com/."
    ]
    for r in refs_right:
        p = tf26_r.add_paragraph() if len(tf26_r.paragraphs[0].text) > 0 else tf26_r.paragraphs[0]
        p.text = r
        p.font.name = FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK
        p.space_after = Pt(7)

    # =========================================================================
    # SLIDE 27: Thank You & Questions & Answers
    # =========================================================================
    s27 = prs.slides.add_slide(blank_layout)
    
    ty_tb = s27.shapes.add_textbox(Inches(0.50), Inches(2.20), Inches(9.00), Inches(1.20))
    ty_tf = ty_tb.text_frame
    ty_p = ty_tf.paragraphs[0]
    ty_p.alignment = PP_ALIGN.CENTER
    ty_p.text = "Thank You"
    ty_p.font.name = FONT_FAMILY
    ty_p.font.size = Pt(40)
    ty_p.font.bold = True
    ty_p.font.color.rgb = COLOR_RED
    
    qa_tb = s27.shapes.add_textbox(Inches(0.50), Inches(3.40), Inches(9.00), Inches(0.80))
    qa_tf = qa_tb.text_frame
    qa_p = qa_tf.paragraphs[0]
    qa_p.alignment = PP_ALIGN.CENTER
    qa_p.text = "Questions & Answers"
    qa_p.font.name = FONT_FAMILY
    qa_p.font.size = Pt(22)
    qa_p.font.bold = True
    qa_p.font.color.rgb = COLOR_DARK
    
    recap_tb = s27.shapes.add_textbox(Inches(0.50), Inches(4.70), Inches(9.00), Inches(1.00))
    rtf = recap_tb.text_frame
    rtf.word_wrap = True
    rp1 = rtf.paragraphs[0]
    rp1.alignment = PP_ALIGN.CENTER
    rp1.text = "AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes"
    rp1.font.name = FONT_FAMILY
    rp1.font.size = Pt(13)
    rp1.font.bold = True
    rp1.font.color.rgb = COLOR_DARK
    
    rp2 = rtf.add_paragraph()
    rp2.alignment = PP_ALIGN.CENTER
    rp2.text = "Aman Choudhary (70012300166)  |  Pranav Patil (70012400176)  |  Mohit Patil (70012400178)"
    rp2.font.name = FONT_FAMILY
    rp2.font.size = Pt(11.5)
    rp2.font.color.rgb = COLOR_MUTED
    
    rp3 = rtf.add_paragraph()
    rp3.alignment = PP_ALIGN.CENTER
    rp3.text = "Under the Guidance of Dr. Pravin Landge | Department of Information Technology, MPSTME, SVKM's NMIMS"
    rp3.font.name = FONT_FAMILY
    rp3.font.size = Pt(11)
    rp3.font.color.rgb = COLOR_RED
    
    # Slide number
    num_box27 = s27.shapes.add_textbox(Inches(7.50), Inches(6.92), Inches(2.00), Inches(0.35))
    np27 = num_box27.text_frame.paragraphs[0]
    np27.alignment = PP_ALIGN.RIGHT
    np27.text = "27"
    np27.font.name = FONT_FAMILY
    np27.font.size = Pt(13)
    np27.font.color.rgb = COLOR_MUTED

    # Save to both workspace and Downloads
    out_workspace = "AI_Student_Agent_Final_Presentation.pptx"
    out_downloads = r"C:\Users\ganesh jadhav\Downloads\AI_Student_Agent_Final_Presentation_NMIMS.pptx"
    
    prs.save(out_workspace)
    print(f"Saved {out_workspace} successfully!")
    prs.save(out_downloads)
    print(f"Saved {out_downloads} successfully!")

if __name__ == '__main__':
    generate_full_presentation()
