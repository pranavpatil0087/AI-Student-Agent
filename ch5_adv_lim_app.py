import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from doc_styles import (
    add_body_p, add_bullet_p, add_chapter_heading,
    add_section_heading
)

def build_chapter_5(doc):
    """Generate Chapter 5: Advantages, Limitations and Applications (~3 pages)."""
    add_chapter_heading(doc, "Chapter 5\nAdvantages, Limitations and Applications")
    
    # Mandatory Introductory Paragraph
    add_body_p(doc, "A comprehensive engineering evaluation of any newly developed software system necessitates a balanced, objective appraisal of its functional merits, operational constraints, and potential deployment horizons. While AI Student Agent successfully realizes a multi-threaded autonomous assistant capable of real-time auditory perception, in-call chat execution, and cloud knowledge persistence, its practical adoption must be contextualized within real-world educational ecosystems. This chapter presents a structured analysis of the system across three core dimensions: the verified technical and pedagogical advantages that distinguish it from conventional educational software, the genuine engineering and ethical limitations identified during implementation and live testing, and its prospective applications across academic institutions, corporate training, and assistive technology.")

    # 5.1 Advantages
    add_section_heading(doc, "5.1 Advantages")
    
    add_body_p(doc, "The architectural and algorithmic choices implemented in AI Student Agent provide substantial technical, cognitive, and financial benefits over existing educational tools:")

    add_bullet_p(doc, " Delegating continuous ambient listening, attendance confirmation, and note transcription to the autonomous agent eliminates the split-attention effect [14], freeing working memory for conceptual comprehension.", bold_prefix="1. Cognitive Load Reduction:")
    add_bullet_p(doc, " Executing speech decoding locally on CPU via VOSK eliminates recurring cloud streaming expenses, reducing operating costs by over 98% (from ~$23.00 down to $0.42 per month for 60 classes).", bold_prefix="2. Zero-Cloud-Cost Ambient Listening:")
    add_bullet_p(doc, " Normalized Levenshtein distance metrics combined with acoustic confidence gating (>= 0.70) reliably resolve homophones and phonetic ambiguity ('just in' vs 'Justin') without false activations.", bold_prefix="3. Resilient Multi-Tier Trigger Detection:")
    add_bullet_p(doc, " Dynamically extracting the rolling 15-word context window feeds immediate lecture discourse to GPT-4o-mini, grounding generated answers in the professor's live discussion and mitigating hallucinations.", bold_prefix="4. Contextually Grounded Real-Time Answering:")
    add_bullet_p(doc, " Post-session synthesis transforms disfluent spoken dialogue into pristine, structured academic study guides featuring conceptual headings, bulleted principles, and technical definitions.", bold_prefix="5. Hierarchical Lecture Note Synthesis:")
    add_bullet_p(doc, " Automatically structures synthesized notes into schema-compliant Notion block trees and publishes them directly to the student's cloud database via the official Notion REST API.", bold_prefix="6. Direct Notion Knowledge Base Integration:")
    add_bullet_p(doc, " Five-tier architecture cleanly separates audio ingestion, speech recognition, browser automation, and reasoning into thread-safe microservices, ensuring modularity and fault isolation.", bold_prefix="7. Decoupled Modular Architecture:")

    # 5.2 Limitations
    add_section_heading(doc, "5.2 Limitations")
    
    add_body_p(doc, "In accordance with rigorous academic integrity, several technical constraints and operational dependencies in the current implementation must be explicitly documented:")

    add_bullet_p(doc, " The browser automation layer relies on Selenium interacting with Google Meet's DOM; major UI redesigns by Google could temporarily impair automated button clicks or chat injection.", bold_prefix="1. Web DOM Selector Volatility:")
    add_bullet_p(doc, " VOSK's local acoustic model decoding accuracy degrades under severe ambient background noise, heavy packet loss, or non-standard regional accents, occasionally delaying trigger recognition.", bold_prefix="2. Acoustic Noise and Accent Sensitivity:")
    add_bullet_p(doc, " Cloud inference round-trips to OpenAI average 2.1s, but network congestion or provider rate-limiting can extend latency to 4.5s, approaching the boundary of natural conversational timing.", bold_prefix="3. Cloud API Network Latency Variations:")
    add_bullet_p(doc, " Capturing classroom audio streams touches upon institutional wiretapping and consent regulations; deploying the agent requires explicit consent from instructors and participants.", bold_prefix="4. Privacy and Institutional Consent:")
    add_bullet_p(doc, " OpenAI Whisper enforces a 25.0 MB file upload limit (~12 minutes of 16 kHz WAV audio), requiring multimodal streaming (Gemini) or temporal audio segmentation for extended lectures.", bold_prefix="5. Cloud Upload File Size Boundaries:")
    add_bullet_p(doc, " The current implementation utilizes VB-Audio Virtual Cable configured for Windows Core Audio; cross-platform deployment on macOS/Linux requires platform-native virtual daemons.", bold_prefix="6. Operating System Driver Dependencies:")

    # 5.3 Applications
    add_section_heading(doc, "5.3 Applications")
    
    add_body_p(doc, "The autonomous architecture pioneered by AI Student Agent possesses broad applicability across diverse educational and professional environments:")

    add_bullet_p(doc, " Directly applicable to remote undergraduate and postgraduate degree programs conducted via Google Meet, providing students with active copilot assistance during complex lectures.", bold_prefix="1. Synchronous Distance Learning:")
    add_bullet_p(doc, " Provides students who miss classes due to illness, sports, or lab conflicts with structured Notion study notes derived from authentic spoken lectures, ensuring academic continuity.", bold_prefix="2. Automated Lecture Archiving:")
    add_bullet_p(doc, " Represents transformative assistive technology for students with ADHD, auditory processing difficulties, or hearing impairments by preventing missed cold-calls and providing real-time text.", bold_prefix="3. Assistive Technology for Neurodivergent Learners:")
    add_bullet_p(doc, " Extensible to corporate training webinars and multi-stakeholder video conferences where participants must track action items, capture discussion points, and verify attendance.", bold_prefix="4. Corporate Training and Webinars:")
    add_bullet_p(doc, " Accreditation bodies (NBA, NAAC) can deploy the note synthesis pipeline to generate standardized, searchable documentation of curriculum delivery and lecture content coverage.", bold_prefix="5. Institutional Compliance and Course Auditing:")

    # Mandatory Chapter Summary Paragraph
    add_body_p(doc, "In summary, this chapter has critically evaluated the multidimensional impact of AI Student Agent. It highlighted the distinct advantages of cognitive load reduction, zero-cost continuous listening, contextually grounded answering, and automated Notion documentation. Concurrently, it honestly delineated real-world limitations, including DOM selector volatility, cloud payload boundaries, and ethical consent requirements. Finally, the chapter demonstrated that the underlying architecture extends far beyond student assistance into assistive accessibility technology and enterprise meeting automation. Building upon these insights, Chapter 6 articulates the definitive conclusion of the report and outlines an ambitious, technically actionable future scope.")

    doc.add_page_break()
