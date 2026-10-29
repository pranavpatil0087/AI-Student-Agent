import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from doc_styles import (
    add_body_p, add_chapter_heading, add_section_heading,
    add_subsection_heading, add_table_title, format_table
)

def build_chapter_2(doc):
    """Generate Chapter 2: Literature Survey (Compact ~5 pages)."""
    add_chapter_heading(doc, "Chapter 2\nLiterature Survey")
    
    # Mandatory Introductory Paragraph
    add_body_p(doc, "The conceptualization, architecture, and engineering realization of AI Student Agent rely on the intersection of several established and emerging sub-disciplines within computer science, artificial intelligence, and educational technology. Specifically, creating an autonomous agent capable of listening, reasoning, interacting, and documenting in a live online classroom requires foundational insights from robust automatic speech recognition (ASR), real-time acoustic streaming, abstractive meeting summarization, large language model (LLM) reasoning, educational chatbot interfaces, robotic process automation (RPA), and multi-agent systems. This chapter conducts a comprehensive literature survey across these interconnected domains. It examines eighteen foundational publications—comprising major journal articles, peer-reviewed international conference proceedings, and technical reports—evaluating their methodological contributions, empirical findings, and architectural constraints. The chapter synthesizes these studies through a structured comparative analysis matrix, explicitly identifies unresolved research gaps, and demonstrates how these limitations directly motivate the technical design of AI Student Agent.")

    # 2.1 Thematic Categorization
    add_section_heading(doc, "2.1 Thematic Categorization of Literature")
    
    add_body_p(doc, "The literature surveyed in this chapter is organized into five distinct thematic clusters: (1) Automatic Speech Recognition and Acoustic Modeling, investigating weakly supervised end-to-end transformers for robust multilingual transcription [1], [5]; (2) Real-Time Streaming Speech Recognition on the Edge, exploring lightweight WFST acoustic toolkits optimized for low-latency CPU execution [2], [3], [4]; (3) Meeting Discourse Summarization and Lecture Note Synthesis, examining hierarchical prompt schemas and abstractive condensation of conversational dialogue [6], [7], [8], [9]; (4) Conversational Educational Agents and Tutoring Systems, analyzing pedagogical chatbots, privacy guardrails, and autonomous interaction taxonomies in higher education [10], [11], [12], [13]; and (5) Process Automation, Attendance Verification, and Frontier Reasoning, evaluating web automation resilience, non-intrusive attendance verification, multi-agent frameworks, and cost-efficient LLM reasoning architectures [14], [15], [16], [17], [18].")

    # 2.2 Detailed Review of Existing Research
    add_section_heading(doc, "2.2 Detailed Review of Existing Research")
    
    add_body_p(doc, "In Speech Recognition and Acoustic Modeling, Radford et al. [1] demonstrated human-level transcription invariance across accents using Whisper, though acknowledging high streaming latency and strict 25 MB payload boundaries. Povey et al. [2] developed Kaldi, pioneering modular WFST speech decoding. Alpha Cephei [3] created VOSK, packaging Kaldi into a CPU-efficient offline engine with sub-500 ms latency, though vulnerable to phonetic ambiguity on proper nouns. Wang et al. [4] investigated ExKaldi-RT, showing that acoustic streaming on consumer processors requires buffering between 1.0 and 3.0 seconds to balance context and latency. Dewan et al. [5] resolved loopback echo corruption in teleconferencing through isolated unidirectional audio routing.")

    add_body_p(doc, "In Meeting Summarization and Note Synthesis, Laskar et al. [6] established that role-directed LLM prompting prevents the loss of key concepts and action items. Rennard et al. [7] identified that spoken disfluencies corrupt raw abstractive summarization unless conditioned by bandpass filtering. Zhong et al. [8] developed the QMSum benchmark, proving that query-directed context extraction outperforms whole-transcript summarization. Madhavi et al. [9] established the tripartite notes standard (headings, bullets, definitions) and identified unintegrated cloud persistence as a primary technical gap.")

    add_body_p(doc, "In Educational Agents and Process Automation, Taneja et al. [10] achieved a 94% satisfaction rate with Jill Watson in asynchronous forums but noted its inability to operate in synchronous classes. Wang et al. [11] found that educational chatbots lacking live lecture context produce ungrounded hallucinations. Labadze et al. [12] mandated secure local credential isolation to protect student privacy. Wollny et al. [13] established that 90% of tools operate at passive Level 1 autonomy, calling for Level 3 autonomous agents. Bhardwaj and Kumar [14] resolved web RPA volatility via resilient DOM fallback chains. Toma [15] reduced attendance logging errors by 98% via automation. Kuang and Baul [16] validated non-intrusive chat attendance over invasive biometric surveillance. Tran et al. [17] proved modular multi-agent architectures exhibit superior fault tolerance. Finally, OpenAI [18] released GPT-4o-mini, achieving 82.0% MMLU at sub-second latency and 60% lower cost ($0.15/M tokens), making real-time classroom reasoning economically viable.")

    # Table 2.1
    add_table_title(doc, "Table 2.1: Comprehensive Comparative Summary of Reviewed Literature")
    
    lit_table = [
        ("Paper & Reference", "Domain / Focus", "Methodology / Approach", "Key Empirical Findings", "Identified Gap & Relevance"),
        ("[1] Radford et al. (2023)", "Speech Recognition", "Weakly supervised Transformer", "Human-level WER; noise invariance", "Batch-only; 25MB limit. Used for post-notes."),
        ("[2] Povey et al. (2011)", "Acoustic Modeling", "WFST modular toolkit (Kaldi)", "State-of-the-art C++ decoding", "Complex setup. Core engine behind VOSK."),
        ("[3] Alpha Cephei (2026)", "Offline Streaming STT", "Lightweight Kaldi port (VOSK)", "Sub-500ms CPU latency; zero cost", "Homophone errors. Core live speech engine."),
        ("[4] Wang et al. (2021)", "Streaming Kaldi ASR", "Python wrapper (ExKaldi-RT)", "Sub-second frame decoding on CPU", "Small buffers elevate WER. Guided 3.0s buffer."),
        ("[5] Dewan et al. (2023)", "Meeting Transcription", "End-to-end verbatim pipeline", "Virtual loopback is main echo cause", "No client bot. Motivated VB-Cable loopback."),
        ("[6] Laskar et al. (2023)", "Meeting Summaries", "Role-directed LLM prompting", "Prevents loss of key action items", "Context limits. Provided notes prompt schema."),
        ("[7] Rennard et al. (2022)", "Abstractive Summaries", "Survey of meeting models", "Disfluencies corrupt raw summaries", "No real-time mode. Informed DSP filtering."),
        ("[8] Zhong et al. (2021)", "Query Summaries", "QMSum query benchmark", "Targeted query context outperforms", "Offline only. Guided 15-word context buffer."),
        ("[9] Madhavi et al. (2025)", "Lecture Running Notes", "Systematic review of note tools", "Tripartite notes hierarchy required", "Unintegrated. Resolved via Notion API export."),
        ("[10] Taneja et al. (2024)", "Teaching Assistants", "Jill Watson via ChatGPT", "94% satisfaction; cuts TA load", "Forum-bound only. Extended to live meetings."),
        ("[11] Wang et al. (2024)", "Higher Ed Chatbots", "ChatEd conversational tool", "Boosted student comprehension", "Hallucinates. Fixed via rolling context deque."),
        ("[12] Labadze et al. (2023)", "AI Education Review", "Review of 42 chatbot studies", "24/7 availability; privacy risks", "Cloud logging risks. Mandated local .env setup."),
        ("[13] Wollny et al. (2021)", "Chatbot Readiness", "Taxonomy of ed chatbots", "90% operate at Level 1 autonomy", "No Level 3 bots. Direct motivation for agent."),
        ("[14] Bhardwaj & Kumar (2025)", "Academic RPA", "Robotic process automation", "Substantial admin efficiency gains", "Rigid DOM breaks. Tackled via fallback chains."),
        ("[15] Toma (2023)", "Administrative RPA", "RPA bot for attendance sheets", "98% reduction in logging errors", "Centralized. Decentralized to student agent."),
        ("[16] Kuang & Baul (2021)", "Biometric Attendance", "Deep learning face recognition", "Accurate; heavy GPU and privacy push", "Privacy fears. Adopted non-intrusive chat bot."),
        ("[17] Tran et al. (2025)", "Multi-Agent LLMs", "Collaborative LLM survey", "Decoupled agents yield robustness", "No audio tools. Blueprint for 5-layer design."),
        ("[18] OpenAI (2024)", "Cost-Efficient LLMs", "GPT-4o-mini architecture", "82.0% MMLU; sub-2s latency; -60% cost", "Cloud API network need. Primary cloud brain."),
    ]

    tbl = doc.add_table(rows=len(lit_table), cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    col_widths = [Inches(1.10), Inches(0.95), Inches(1.25), Inches(1.25), Inches(1.22)]
    for row in tbl.rows:
        for idx, w in enumerate(col_widths):
            row.cells[idx].width = w

    for i, row in enumerate(lit_table):
        for c, text in enumerate(row):
            tbl.cell(i, c).text = text

    format_table(tbl, header_bg="E8EEF5", font_size=6.8, header_font_size=7.5)

    # Research Gap Synthesis
    add_section_heading(doc, "2.3 Research Gap Synthesis and Technical Motivation")
    
    add_body_p(doc, "The comprehensive examination of existing literature across speech recognition, educational technology, and process automation synthesizes four distinct research gaps that directly motivate this capstone project:")

    add_body_p(doc, " While offline models like Whisper [1] achieve high transcription fidelity, they are computationally heavy, require cloud streaming, and enforce strict payload limits. Conversely, lightweight engines like VOSK [3] run efficiently on CPUs but suffer from phonetic ambiguity on proper nouns. Existing literature reveals an absence of hybrid architectures that combine continuous local listening for real-time triggers with cloud-based batch processing for final session synthesis.", bold_prefix="1. The Gap in Low-Cost Ambient Audio Perception:")

    add_body_p(doc, " As emphasized by Wollny et al. [13] and Taneja et al. [10], conversational educational AI tools remain overwhelmingly passive. They function as static web forms waiting for human input, disconnected from the active meeting interface. There is an evident engineering void for an autonomous agent capable of perceiving speech, evaluating intent, and taking executive action inside the meeting platform.", bold_prefix="2. The Gap in Real-Time Autonomous Classroom Agency:")

    add_body_p(doc, " Studies evaluating educational chatbots [11] consistently warn of ungrounded hallucinations when answering technical questions. Commercial meeting summarizers [6] operate globally across full transcripts but lack the mechanism to isolate immediate preceding conversational turns necessary to ground a real-time answer during an active cold-call.", bold_prefix="3. The Gap in Contextually Grounded Real-Time Answering:")

    add_body_p(doc, " While academic note generation has been explored theoretically [9], existing tools fail to deliver an automated end-to-end bridge between live lecture speech and structured, searchable cloud repositories like Notion. Students are still required to manually copy, format, and organize post-session transcripts.", bold_prefix="4. The Gap in Automated Knowledge Persistence:")

    # Mandatory Chapter Summary Paragraph
    add_body_p(doc, "In summary, this chapter critically surveyed eighteen foundational research publications across speech recognition, meeting summarization, educational chatbots, and RPA. The comparative matrix in Table 2.1 highlights trade-offs between local and cloud computing, passive bots and autonomous agents, and raw transcripts versus structured notes. The identified research gaps—the lack of hybrid edge-cloud perception, the absence of real-time in-meeting agency, and the disconnect from external knowledge repositories—firmly substantiate the engineering necessity of AI Student Agent. Armed with these scholarly insights, Chapter 3 details the comprehensive methodology, architectural block diagrams, unified UML models, hardware specifications, and algorithmic pipelines implemented to realize the system.")

    doc.add_page_break()
