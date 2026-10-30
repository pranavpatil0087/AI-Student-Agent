import os
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Exact Brand Colors
COLOR_RED = RGBColor(188, 49, 44)        # #BC312C (NMIMS Brand Red)
COLOR_BLUE = RGBColor(46, 95, 138)       # #2E5F8A (Steel Navy Blue)
COLOR_GREEN = RGBColor(63, 143, 92)      # #3F8F5C (Forest Green)
COLOR_AMBER = RGBColor(201, 138, 44)     # #C98A2C (Amber Gold)
COLOR_DARK = RGBColor(51, 51, 51)        # #333333 (Charcoal Dark Text)
COLOR_MUTED = RGBColor(102, 102, 102)    # #666666 (Subtext Gray)
COLOR_CARD_BG = RGBColor(244, 246, 248)  # #F4F6F8 (Neutral Card Fill)
COLOR_WHITE = RGBColor(255, 255, 255)    # #FFFFFF (Pure White)
COLOR_BORDER = RGBColor(208, 208, 208)   # #D0D0D0 (Subtle Gray Border)
COLOR_LIGHT_RED = RGBColor(253, 242, 242)

FONT_FAMILY = "Times New Roman"

def add_card(slide, left, top, width, height, fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER, border_width=Pt(1)):
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

def add_bullet_point(tf, bold_prefix, text_content, font_size=Pt(11), space_after=Pt(6)):
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

def build_external_review_presentation():
    template_src = r'C:\Users\ganesh jadhav\Downloads\Sample PPT format External Review.pptx'
    prs = Presentation(template_src)

    # -------------------------------------------------------------
    # SLIDE 1: Title Slide
    # -------------------------------------------------------------
    s1 = prs.slides[0]
    # Find Title 1 placeholder and format text cleanly
    for shp in s1.shapes:
        if shp.name == "Title 1" and shp.has_text_frame:
            tf = shp.text_frame
            tf.clear()
            
            p0 = tf.paragraphs[0]
            p0.alignment = PP_ALIGN.CENTER
            p0.text = "A Presentation"
            p0.font.name = FONT_FAMILY
            p0.font.size = Pt(15)
            p0.font.bold = True
            p0.font.color.rgb = COLOR_DARK
            
            p1 = tf.add_paragraph()
            p1.alignment = PP_ALIGN.CENTER
            p1.text = "on"
            p1.font.name = FONT_FAMILY
            p1.font.size = Pt(13)
            p1.font.color.rgb = COLOR_MUTED
            
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.text = "“AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes”"
            p2.font.name = FONT_FAMILY
            p2.font.size = Pt(20)
            p2.font.bold = True
            p2.font.color.rgb = COLOR_RED
            p2.space_after = Pt(10)
            
            p3 = tf.add_paragraph()
            p3.alignment = PP_ALIGN.CENTER
            p3.text = "Presented By:"
            p3.font.name = FONT_FAMILY
            p3.font.size = Pt(13)
            p3.font.bold = True
            p3.font.color.rgb = COLOR_DARK
            
            p4 = tf.add_paragraph()
            p4.alignment = PP_ALIGN.CENTER
            p4.text = "1. Aman Choudhary (70012300166)     2. Pranav Patil (70012400176)     3. Mohit Patil (70012400178)"
            p4.font.name = FONT_FAMILY
            p4.font.size = Pt(11.5)
            p4.font.color.rgb = COLOR_DARK
            p4.space_after = Pt(8)
            
            p5 = tf.add_paragraph()
            p5.alignment = PP_ALIGN.CENTER
            p5.text = "Under the Guidance of:"
            p5.font.name = FONT_FAMILY
            p5.font.size = Pt(13)
            p5.font.bold = True
            p5.font.color.rgb = COLOR_DARK
            
            p6 = tf.add_paragraph()
            p6.alignment = PP_ALIGN.CENTER
            p6.text = "Dr. Pravin Landge (Faculty Mentor, Department of Information Technology)"
            p6.font.name = FONT_FAMILY
            p6.font.size = Pt(12)
            p6.font.bold = True
            p6.font.color.rgb = COLOR_BLUE

    # -------------------------------------------------------------
    # SLIDE 2: Outline
    # -------------------------------------------------------------
    s2 = prs.slides[1]
    # Remove default plain bullet placeholder text if present
    for shp in list(s2.shapes):
        if shp.name == "Content Placeholder 2":
            shp.text_frame.clear()
            
    outline_items = [
        ("01", "Objective of Project", "Goals, ambient listening, and persistence"),
        ("02", "Literature Survey", "Comparative analysis of 18 publications"),
        ("03", "Problem Statement", "Split-attention, autonomy & memory gaps"),
        ("04", "Proposed Solution", "Hybrid edge-cloud multi-threaded architecture"),
        ("05", "System Architecture", "5-Layer pipeline and audio routing conduit"),
        ("06", "Result Analysis / Implementation", "10/10 test matrix, latency & cost savings"),
        ("07", "Conclusions and Future Scope", "Engineering achievements & roadmap"),
        ("08", "References", "Standard IEEE academic citations")
    ]
    card_w = Inches(4.30)
    card_h = Inches(0.95)
    for idx, (num, title, desc) in enumerate(outline_items):
        col = idx // 4
        row = idx % 4
        cx = Inches(0.50) + col * (card_w + Inches(0.40))
        cy = Inches(2.20) + row * (card_h + Inches(0.20))
        add_card(s2, cx, cy, card_w, card_h, fill_color=COLOR_WHITE, border_color=COLOR_BORDER, border_width=Pt(1.2))
        add_badge(s2, cx + Inches(0.15), cy + Inches(0.20), Inches(0.55), num, color=COLOR_RED if col == 0 else COLOR_BLUE)
        
        tb = s2.shapes.add_textbox(cx + Inches(0.85), cy + Inches(0.12), card_w - Inches(0.95), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(12)
        p0.font.bold = True
        p0.font.color.rgb = COLOR_DARK
        p1 = tf.add_paragraph()
        p1.text = desc
        p1.font.name = FONT_FAMILY
        p1.font.size = Pt(9.5)
        p1.font.color.rgb = COLOR_MUTED

    # -------------------------------------------------------------
    # SLIDE 3: Objective of Project
    # -------------------------------------------------------------
    s3 = prs.slides[2]
    for shp in list(s3.shapes):
        if shp.name == "Text Placeholder 2":
            shp.text_frame.clear()
            
    obj_cards = [
        ("1", "Zero-Cloud-Cost Ambient Listening & ASR", 
         [("Virtual Audio Conduit: ", "Intercept Google Meet audio via VB-Audio Virtual Cable as a 16 kHz mono 16-bit PCM stream without microphone feedback loops."),
          ("Offline Speech Recognition: ", "Deploy VOSK toolkit with Kaldi WFST acoustic modeling locally on CPU, achieving sub-500ms latency at $0 cloud cost."),
          ("Noise Filtering: ", "Apply SciPy Butterworth 4th-order bandpass filter (300 Hz–3,400 Hz) to eliminate 50 Hz hum and enhance speech formants.")],
         COLOR_RED, Inches(0.50)),
        ("2", "Contextual Trigger & Autonomous In-Call Chat", 
         [("Sliding Context Deque: ", "Maintain in-memory 50-word rolling history, extracting 15 words upon trigger detection to prevent hallucinations."),
          ("Fuzzy Phonetic Matching: ", "Utilize normalized Levenshtein distance (Sim >= 0.80) and confidence score gating (>= 0.70) with compound lookahead ('just in' -> 'Justin')."),
          ("Grounded Q&A Dispatch: ", "Query OpenAI GPT-4o-mini for grounded responses and inject into Google Meet chat via Selenium in < 4.0 seconds.")],
         COLOR_BLUE, Inches(3.60)),
        ("3", "Post-Session Synthesis & Notion Persistence", 
         [("Lossless Serialization: ", "Compile all raw lecture PCM audio frames into a master debug_audio.wav file upon session termination."),
          ("Whisper-1 Master Transcript: ", "Execute asynchronous batch transcription via OpenAI Whisper-1 API to yield verbatim timestamped transcripts."),
          ("Structured Notion Sync: ", "Synthesize hierarchical study guides (Topic, Concepts, Definitions) and export formatted block trees to Notion via REST API.")],
         COLOR_GREEN, Inches(6.70))
    ]
    for num_str, title, bullets, color, left in obj_cards:
        add_card(s3, left, Inches(2.20), Inches(2.80), Inches(4.55), fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        add_badge(s3, left + Inches(1.15), Inches(2.35), Inches(0.50), num_str, color=color)
        tb = s3.shapes.add_textbox(left + Inches(0.15), Inches(2.95), Inches(2.50), Inches(0.70))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(12)
        p0.font.bold = True
        p0.font.color.rgb = color
        
        b_tb = s3.shapes.add_textbox(left + Inches(0.15), Inches(3.65), Inches(2.50), Inches(2.95))
        btf = b_tb.text_frame
        btf.word_wrap = True
        btf.margin_left = btf.margin_top = btf.margin_right = btf.margin_bottom = 0
        for b_prefix, b_text in bullets:
            add_bullet_point(btf, b_prefix, b_text, font_size=Pt(9.5), space_after=Pt(5))

    # -------------------------------------------------------------
    # SLIDE 4: Literature Survey
    # -------------------------------------------------------------
    s4 = prs.slides[3]
    for shp in list(s4.shapes):
        if shp.name == "Text Placeholder 2":
            shp.text_frame.clear()
            
    tbl_shape4 = s4.shapes.add_table(7, 4, Inches(0.40), Inches(2.15), Inches(9.20), Inches(4.50))
    tbl4 = tbl_shape4.table
    col_w4 = [Inches(1.80), Inches(2.70), Inches(2.20), Inches(2.50)]
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
            tbl4.cell(r_idx, c_idx).text = val
    style_table(tbl4, col_w4)

    # -------------------------------------------------------------
    # SLIDE 5: Problem Statement
    # -------------------------------------------------------------
    s5 = prs.slides[4]
    for shp in list(s5.shapes):
        if shp.name == "Text Placeholder 2":
            shp.text_frame.clear()
            
    # Left column: 4 Core Gaps
    left_tb5 = s5.shapes.add_textbox(Inches(0.50), Inches(2.20), Inches(5.10), Inches(4.60))
    ltf5 = left_tb5.text_frame
    ltf5.word_wrap = True
    ltf5.margin_left = ltf5.margin_top = ltf5.margin_right = ltf5.margin_bottom = 0
    add_bullet_point(ltf5, "Cognitive Split-Attention Effect:", "Students struggle to simultaneously absorb complex lecture derivations, take running notes, and watch chat panels (Sweller's Cognitive Load Theory).", font_size=Pt(12), space_after=Pt(8))
    add_bullet_point(ltf5, "Autonomy Gap in EdTech:", "Existing AI tools (ChatGPT, Copilot) are passive web forms; they cannot listen to live meeting audio or take autonomous action in the call.", font_size=Pt(12), space_after=Pt(8))
    add_bullet_point(ltf5, "Ephemeral Meeting Discourse:", "Crucial verbal explanations, Q&A interactions, and attendance roll-calls vanish immediately once the Google Meet session terminates.", font_size=Pt(12), space_after=Pt(8))
    add_bullet_point(ltf5, "Unbounded LLM Hallucinations:", "General-purpose models hallucinate unless strictly grounded in the rolling context of the instructor's immediate lecture discourse.", font_size=Pt(12), space_after=Pt(6))

    # Right column: In-Call Scenario Box
    add_card(s5, Inches(5.80), Inches(2.20), Inches(3.70), Inches(4.55), fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER)
    r_hdr5 = s5.shapes.add_textbox(Inches(6.00), Inches(2.35), Inches(3.30), Inches(0.40))
    rp5 = r_hdr5.text_frame.paragraphs[0]
    rp5.text = "Live Classroom Scenario"
    rp5.font.name = FONT_FAMILY
    rp5.font.size = Pt(13)
    rp5.font.bold = True
    rp5.font.color.rgb = COLOR_RED

    def add_mini_scenario5(top_in, title, detail, tag="?"):
        add_card(s5, Inches(6.00), Inches(top_in), Inches(3.30), Inches(0.80), fill_color=COLOR_WHITE, border_color=COLOR_BORDER)
        tb = s5.shapes.add_textbox(Inches(6.10), Inches(top_in + 0.05), Inches(2.70), Inches(0.70))
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
        
        tag_tb = s5.shapes.add_textbox(Inches(8.85), Inches(top_in + 0.15), Inches(0.35), Inches(0.40))
        ttp = tag_tb.text_frame.paragraphs[0]
        ttp.text = tag
        ttp.alignment = PP_ALIGN.CENTER
        ttp.font.name = FONT_FAMILY
        ttp.font.size = Pt(12)
        ttp.font.bold = True
        ttp.font.color.rgb = COLOR_RED

    add_mini_scenario5(2.85, "Instructor Spoken Cold-Call", "“Justin, can you explain how paging prevents external fragmentation?”", tag="!")
    add_mini_scenario5(3.75, "Split-Attention Multitasking", "Student writing notes -> delayed response -> missing participation grade", tag="X")
    
    add_card(s5, Inches(6.00), Inches(4.70), Inches(3.30), Inches(1.85), fill_color=COLOR_LIGHT_RED, border_color=COLOR_RED)
    b_tb5 = s5.shapes.add_textbox(Inches(6.10), Inches(4.80), Inches(3.10), Inches(1.65))
    btf5 = b_tb5.text_frame
    btf5.word_wrap = True
    btf5.margin_left = btf5.margin_top = btf5.margin_right = btf5.margin_bottom = 0
    bp1 = btf5.paragraphs[0]
    bp1.text = "Engineering Solution:"
    bp1.font.name = FONT_FAMILY
    bp1.font.size = Pt(11)
    bp1.font.bold = True
    bp1.font.color.rgb = COLOR_RED
    bp2 = btf5.add_paragraph()
    bp2.text = "An autonomous, ambient edge-cloud assistant that joins Google Meet, listens locally with zero API fees, answers cold-calls in <4s, and syncs study guides to Notion."
    bp2.font.name = FONT_FAMILY
    bp2.font.size = Pt(10)
    bp2.font.color.rgb = COLOR_DARK

    # -------------------------------------------------------------
    # SLIDE 6: Proposed Solution
    # -------------------------------------------------------------
    s6 = prs.slides[5]
    for shp in list(s6.shapes):
        if shp.name == "Text Placeholder 2":
            shp.text_frame.clear()
            
    sol_cards = [
        ("Hybrid Edge-Cloud Architecture", 
         [("Edge Continuous Listening: ", "Local CPU speech decoding via VOSK runs 100% free ($0/hr), eliminating costly continuous cloud audio streaming."),
          ("Frontier Cloud Reasoning: ", "OpenAI GPT-4o-mini is invoked strictly on verified trigger events ($0.001/query), saving over 98% in operational costs.")],
         COLOR_RED, Inches(0.50), Inches(2.20)),
        ("Unidirectional Loopback & DSP Filtering", 
         [("VB-Audio Virtual Cable: ", "Routes meeting speaker output directly into PyAudio recording stream without physical microphone feedback or echo."),
          ("SciPy Butterworth Filter: ", "4th-order bandpass filter (300 Hz–3,400 Hz) isolates human speech formants and suppresses ambient electrical hum.")],
         COLOR_BLUE, Inches(5.10), Inches(2.20)),
        ("Sliding Context Deque & Fuzzy Trigger", 
         [("Rolling Discourse Buffer: ", "In-memory collections.deque(maxlen=50) retains the last 50 words, extracting 15 words to ground prompt synthesis."),
          ("Fuzzy Phonetic Matching: ", "Normalized Levenshtein metric (Sim >= 0.80) + acoustic confidence (>= 0.70) with lookahead ('just in' -> 'Justin').")],
         COLOR_AMBER, Inches(0.50), Inches(4.55)),
        ("Resilient Automation & Notion Persistence", 
         [("DOM Mutation Resilience: ", "Selenium WebDriver anchors to immutable accessibility attributes (@aria-label) and dynamic WebDriverWait polling."),
          ("Notion REST API Adapter: ", "Converts post-session Whisper-1 transcript into hierarchical academic study notes published to student Notion DB.")],
         COLOR_GREEN, Inches(5.10), Inches(4.55))
    ]
    for title, bullets, color, left, top in sol_cards:
        add_card(s6, left, top, Inches(4.40), Inches(2.15), fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        tb = s6.shapes.add_textbox(left + Inches(0.20), top + Inches(0.12), Inches(4.00), Inches(1.95))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(12)
        p0.font.bold = True
        p0.font.color.rgb = color
        p0.space_after = Pt(4)
        for b_prefix, b_text in bullets:
            add_bullet_point(tf, b_prefix, b_text, font_size=Pt(9.5), space_after=Pt(3))

    # -------------------------------------------------------------
    # SLIDE 7: System Architecture
    # -------------------------------------------------------------
    s7 = prs.slides[6]
    for shp in list(s7.shapes):
        if shp.name == "Text Placeholder 2":
            shp.text_frame.clear()
            
    arch_img = 'system_architecture_minimalist.png'
    if os.path.exists(arch_img):
        add_card(s7, Inches(0.45), Inches(2.10), Inches(9.10), Inches(3.60), fill_color=COLOR_WHITE, border_color=COLOR_BORDER)
        s7.shapes.add_picture(arch_img, Inches(0.50), Inches(2.15), Inches(9.00), Inches(3.50))
        
    tb7_desc = s7.shapes.add_textbox(Inches(0.50), Inches(5.80), Inches(9.00), Inches(1.10))
    tf7 = tb7_desc.text_frame
    tf7.word_wrap = True
    tf7.margin_left = tf7.margin_top = tf7.margin_right = tf7.margin_bottom = 0
    add_bullet_point(tf7, "Layer 1 & 2 (Client & Audio Loopback): ", "Selenium Chrome automation bypasses modals; VB-Audio Cable isolates 16 kHz mono stream; SciPy Butterworth bandpass filter suppresses noise.", font_size=Pt(10), space_after=Pt(2))
    add_bullet_point(tf7, "Layer 3 (Perception Engine): ", "Offline VOSK Kaldi decodes speech locally on CPU; 50-word context deque buffers discourse; Levenshtein detector triggers GPT-4o-mini; Selenium types chat reply in < 4s.", font_size=Pt(10), space_after=Pt(2))
    add_bullet_point(tf7, "Dual Execution Paths: ", "Real-time in-call Q&A loop (<3.5s turn-taking) and post-session lossless serialization (debug_audio.wav -> Whisper-1 -> Notion database sync).", font_size=Pt(10), space_after=Pt(1))

    # -------------------------------------------------------------
    # SLIDE 8: Result Analysis/Implementation
    # -------------------------------------------------------------
    s8 = prs.slides[7]
    for shp in list(s8.shapes):
        if shp.name == "Text Placeholder 2":
            shp.text_frame.clear()
            
    # 4 Stat Metric Cards at top
    stats8 = [
        ("10 / 10", "Subsystems Verified", COLOR_RED),
        ("3.6 s", "Avg Live Turn-Taking", COLOR_BLUE),
        ("98.17%", "Cloud Cost Savings", COLOR_GREEN),
        ("0", "GPU Accelerators Needed", COLOR_AMBER)
    ]
    card_w8 = Inches(2.10)
    card_gap8 = Inches(0.20)
    for i, (val, label, color) in enumerate(stats8):
        cx = Inches(0.50) + i * (card_w8 + card_gap8)
        add_card(s8, cx, Inches(2.15), card_w8, Inches(1.05), fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        stb = s8.shapes.add_textbox(cx, Inches(2.20), card_w8, Inches(0.55))
        sp = stb.text_frame.paragraphs[0]
        sp.alignment = PP_ALIGN.CENTER
        sp.text = val
        sp.font.name = FONT_FAMILY
        sp.font.size = Pt(20)
        sp.font.bold = True
        sp.font.color.rgb = color
        
        ltb = s8.shapes.add_textbox(cx, Inches(2.75), card_w8, Inches(0.40))
        lp = ltb.text_frame.paragraphs[0]
        lp.alignment = PP_ALIGN.CENTER
        lp.text = label
        lp.font.name = FONT_FAMILY
        lp.font.size = Pt(9.5)
        lp.font.color.rgb = COLOR_MUTED

    # Middle Latency Table
    tbl_shape8 = s8.shapes.add_table(5, 4, Inches(0.50), Inches(3.35), Inches(9.00), Inches(2.05))
    t8 = tbl_shape8.table
    w8 = [Inches(2.50), Inches(2.20), Inches(2.10), Inches(2.20)]
    t8_data = [
        ["Pipeline Stage", "Subsystem", "Measured Latency", "Engineering Target"],
        ["Audio Ingestion & DSP", "PyAudio & SciPy Butterworth", "140 ms – 275 ms", "< 350 ms (Met)"],
        ["Offline Speech Decoding", "VOSK Kaldi WFST Graph", "350 ms – 480 ms", "< 500 ms (Met)"],
        ["Context & LLM Q&A", "OpenAI GPT-4o-mini API", "1,850 ms – 2,650 ms", "< 3,000 ms (Met)"],
        ["Total End-to-End Latency", "Full Perception to Chat Loop", "3.20 s – 4.54 s (Avg: 3.6s)", "3.0 s – 5.0 s (Met)"]
    ]
    for r_idx, row in enumerate(t8_data):
        for c_idx, val in enumerate(row):
            t8.cell(r_idx, c_idx).text = val
    style_table(t8, w8)

    # Bottom Cost & Host Metrics Card
    add_card(s8, Inches(0.50), Inches(5.55), Inches(9.00), Inches(1.35), fill_color=COLOR_WHITE, border_color=COLOR_GREEN, border_width=Pt(1.5))
    tb8_b = s8.shapes.add_textbox(Inches(0.65), Inches(5.65), Inches(8.70), Inches(1.15))
    tf8_b = tb8_b.text_frame
    tf8_b.word_wrap = True
    tf8_b.margin_left = tf8_b.margin_top = tf8_b.margin_right = tf8_b.margin_bottom = 0
    p = tf8_b.paragraphs[0]
    p.text = "Economic Cost Model & Host Hardware Footprint:"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN
    p.space_after = Pt(3)
    add_bullet_point(tf8_b, "98.17% Compute Cost Reduction: ", "Pure cloud streaming costs $22.98/month (60 classes), while AI Student Agent hybrid costs $0.42/month ($0.45/semester).", font_size=Pt(9.5), space_after=Pt(2))
    add_bullet_point(tf8_b, "Resource Footprint: ", "Host RAM: 1.18–1.34 GB (<1.5 GB budget) | Host CPU: 18%–26% on Intel i5 | Dedicated GPU: 0% (standard laptop viability).", font_size=Pt(9.5), space_after=Pt(1))

    # -------------------------------------------------------------
    # SLIDE 9: Conclusions and Future Scope
    # -------------------------------------------------------------
    s9 = prs.slides[8]
    for shp in list(s9.shapes):
        if shp.name == "Text Placeholder 2":
            shp.text_frame.clear()
            
    # Top Conclusion Card
    add_card(s9, Inches(0.50), Inches(2.15), Inches(9.00), Inches(1.60), fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER)
    tb9_top = s9.shapes.add_textbox(Inches(0.70), Inches(2.25), Inches(8.60), Inches(1.40))
    tf9_top = tb9_top.text_frame
    tf9_top.word_wrap = True
    tf9_top.margin_left = tf9_top.margin_top = tf9_top.margin_right = tf9_top.margin_bottom = 0
    p = tf9_top.paragraphs[0]
    p.text = "Project Conclusions:"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK
    p.space_after = Pt(4)
    cp9 = tf9_top.add_paragraph()
    cp9.text = "AI Student Agent has successfully demonstrated that combining local edge acoustic processing (VOSK/Kaldi) with frontier cloud generative models (GPT-4o-mini & Whisper-1) creates an effective autonomous classroom assistant. The system resolves the split-attention dilemma in online learning, maintains sub-4s response latency in live Google Meet chat, and delivers 98.17% cloud cost reductions ($0.42/month) on standard consumer PC hardware without dedicated GPU acceleration."
    cp9.font.name = FONT_FAMILY
    cp9.font.size = Pt(10)
    cp9.font.color.rgb = COLOR_DARK

    # Bottom 4 Future Scope Cards
    fs_cards9 = [
        ("Dynamic Audio Chunking", "Automate PyDub temporal segmentation to process multi-hour WAVs within Whisper's 25 MB limit.", COLOR_RED),
        ("Course-Grounded RAG", "Index syllabus and textbook PDFs via ChromaDB / FAISS for deep factual grounding during cold-calls.", COLOR_BLUE),
        ("Synthetic Spoken Voice (TTS)", "Integrate Edge-TTS / ElevenLabs through a virtual microphone for spoken verbal participation.", COLOR_GREEN),
        ("Cross-Platform Support", "Extend automation adapters to Zoom, MS Teams, Cisco Webex, macOS (BlackHole), and Linux.", COLOR_AMBER)
    ]
    card_w9 = Inches(2.15)
    card_gap9 = Inches(0.18)
    for idx, (title, desc, color) in enumerate(fs_cards9):
        cx = Inches(0.50) + idx * (card_w9 + card_gap9)
        cy = Inches(3.95)
        add_card(s9, cx, cy, card_w9, Inches(2.70), fill_color=COLOR_WHITE, border_color=color, border_width=Pt(1.5))
        tb = s9.shapes.add_textbox(cx + Inches(0.12), cy + Inches(0.15), card_w9 - Inches(0.24), Inches(2.40))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.name = FONT_FAMILY
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = color
        p0.space_after = Pt(6)
        p1 = tf.add_paragraph()
        p1.text = desc
        p1.font.name = FONT_FAMILY
        p1.font.size = Pt(9.5)
        p1.font.color.rgb = COLOR_DARK

    # -------------------------------------------------------------
    # SLIDE 10: References (IEEE Standard)
    # -------------------------------------------------------------
    s10 = prs.slides[9]
    # Remove instructions box Rectangle 2
    for shp in list(s10.shapes):
        if shp.name == "Rectangle 2":
            # Remove shape via parent element
            sp_elem = shp._element
            sp_elem.getparent().remove(sp_elem)
            
    # Add 2 columns of IEEE references
    tb10_l = s10.shapes.add_textbox(Inches(0.50), Inches(2.20), Inches(4.35), Inches(4.50))
    tf10_l = tb10_l.text_frame
    tf10_l.word_wrap = True
    tf10_l.margin_left = tf10_l.margin_top = tf10_l.margin_right = tf10_l.margin_bottom = 0
    
    refs10_left = [
        "[1] A. Radford, J. W. Kim, T. Xu, G. Brockman, C. McLeavey, and I. Sutskever, “Robust speech recognition via large-scale weak supervision,” in Proc. 40th Int. Conf. Machine Learning (ICML), PMLR vol. 202, pp. 28492–28518, 2023.",
        "[2] D. Povey, A. Ghoshal, G. Boulianne, et al., “The Kaldi speech recognition toolkit,” in Proc. IEEE Workshop Automatic Speech Recognition and Understanding (ASRU), pp. 1–4, 2011.",
        "[3] Alpha Cephei, “Vosk: Offline open-source speech recognition toolkit,” Alpha Cephei Inc., Sep. 2026. [Online]. Available: https://alphacephei.com/vosk/.",
        "[4] A. Dewan, M. Ziemski, H. Meylan, et al., “Developing automatic verbatim transcripts for international multilingual meetings: An end-to-end solution,” arXiv:2309.15609, 2023."
    ]
    for r in refs10_left:
        p = tf10_l.add_paragraph() if len(tf10_l.paragraphs[0].text) > 0 else tf10_l.paragraphs[0]
        p.text = r
        p.font.name = FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK
        p.space_after = Pt(10)

    tb10_r = s10.shapes.add_textbox(Inches(5.15), Inches(2.20), Inches(4.35), Inches(4.50))
    tf10_r = tb10_r.text_frame
    tf10_r.word_wrap = True
    tf10_r.margin_left = tf10_r.margin_top = tf10_r.margin_right = tf10_r.margin_bottom = 0
    
    refs10_right = [
        "[5] M. T. R. Laskar, X.-Y. Fu, C. Chen, and S. Bhushan TN, “Building real-world meeting summarization systems using large language models: A practical perspective,” arXiv:2310.19233, 2023.",
        "[6] S. Wollny, J. Schneider, D. Di Mitri, et al., “Are we there yet? — A systematic literature review on chatbots in education,” Frontiers in Artificial Intelligence, vol. 4, art. 654924, 2021.",
        "[7] V. Bhardwaj and M. Kumar, “Transforming higher education with robotic process automation: Enhancing efficiency and student learning,” Discover Sustainability, vol. 6, no. 1, 2025.",
        "[8] OpenAI, “GPT-4o mini: Advancing cost-efficient intelligence,” OpenAI Technical Report, 2024. [Online]. Available: https://openai.com/index/gpt-4o-mini-advancing-cost-efficient-intelligence/."
    ]
    for r in refs10_right:
        p = tf10_r.add_paragraph() if len(tf10_r.paragraphs[0].text) > 0 else tf10_r.paragraphs[0]
        p.text = r
        p.font.name = FONT_FAMILY
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK
        p.space_after = Pt(10)

    # -------------------------------------------------------------
    # SLIDE 11: THANK YOU
    # -------------------------------------------------------------
    s11 = prs.slides[10]
    # Clean default text on title placeholder if present
    for shp in s11.shapes:
        if shp.has_text_frame and "THANK YOU" in shp.text_frame.text:
            shp.text_frame.clear()
            p = shp.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = "THANK YOU"
            p.font.name = FONT_FAMILY
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

    qa_tb11 = s11.shapes.add_textbox(Inches(0.50), Inches(3.30), Inches(9.00), Inches(0.80))
    qa_tf11 = qa_tb11.text_frame
    qa_p11 = qa_tf11.paragraphs[0]
    qa_p11.alignment = PP_ALIGN.CENTER
    qa_p11.text = "Questions & Answers"
    qa_p11.font.name = FONT_FAMILY
    qa_p11.font.size = Pt(22)
    qa_p11.font.bold = True
    qa_p11.font.color.rgb = COLOR_DARK

    recap_tb11 = s11.shapes.add_textbox(Inches(0.50), Inches(4.50), Inches(9.00), Inches(1.50))
    rtf11 = recap_tb11.text_frame
    rtf11.word_wrap = True
    rp1 = rtf11.paragraphs[0]
    rp1.alignment = PP_ALIGN.CENTER
    rp1.text = "AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes"
    rp1.font.name = FONT_FAMILY
    rp1.font.size = Pt(13)
    rp1.font.bold = True
    rp1.font.color.rgb = COLOR_DARK
    
    rp2 = rtf11.add_paragraph()
    rp2.alignment = PP_ALIGN.CENTER
    rp2.text = "Aman Choudhary (70012300166)  |  Pranav Patil (70012400176)  |  Mohit Patil (70012400178)"
    rp2.font.name = FONT_FAMILY
    rp2.font.size = Pt(11.5)
    rp2.font.color.rgb = COLOR_MUTED
    
    rp3 = rtf11.add_paragraph()
    rp3.alignment = PP_ALIGN.CENTER
    rp3.text = "Under the Guidance of Dr. Pravin Landge | Department of Information Technology, MPSTME, SVKM's NMIMS"
    rp3.font.name = FONT_FAMILY
    rp3.font.size = Pt(11)
    rp3.font.color.rgb = COLOR_RED

    # Save to workspace and Downloads
    out_ws = "AI_Student_Agent_External_Review_11Slides.pptx"
    out_dl = r"C:\Users\ganesh jadhav\Downloads\AI_Student_Agent_External_Review_11Slides.pptx"
    
    prs.save(out_ws)
    print(f"Saved workspace file: {out_ws}")
    prs.save(out_dl)
    print(f"Saved downloads file: {out_dl}")

if __name__ == '__main__':
    build_external_review_presentation()
