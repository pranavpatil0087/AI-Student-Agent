import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from doc_styles import (
    add_body_p, add_bullet_p, add_chapter_heading,
    add_section_heading
)

def build_chapter_6(doc):
    """Generate Chapter 6: Conclusion and Future Scope (strictly 2 pages)."""
    add_chapter_heading(doc, "Chapter 6\nConclusion and Future Scope")
    
    # Mandatory Introductory Paragraph
    add_body_p(doc, "As synchronous online education cements its position as a central pillar of global engineering pedagogy, resolving the cognitive friction between real-time comprehension and administrative multitasking has become a pressing necessity. This final chapter consolidates the engineering accomplishments, empirical findings, and architectural principles established throughout the design, implementation, and verification of AI Student Agent. It presents the single, unified academic conclusion of this capstone project report, providing a synthesis of the verified functional outcomes. Subsequently, the chapter outlines a lucid, actionable future research roadmap, detailing five strategic architectural enhancements designed to expand the system from a real-time lecture copilot into an adaptive educational intelligence platform.")

    # 6.1 Conclusion
    add_section_heading(doc, "6.1 Conclusion")
    
    add_body_p(doc, "This capstone project has successfully designed, implemented, and empirically validated AI Student Agent, an autonomous, multi-threaded artificial intelligence learning assistant engineered for real-time online classes conducted via Google Meet. By addressing the foundational Autonomy, Memory, Knowledge, and Evaluation gaps that restrict conventional educational software, the project demonstrates that modern language models and edge acoustic processors can be unified into an ambient, proactive digital proxy that actively supports undergraduate students during synchronous instruction.")

    add_body_p(doc, "The primary technical achievement lies in the hybrid edge-cloud architecture. By deploying the open-source VOSK toolkit with Kaldi acoustic models locally on the host CPU and intercepting speaker audio through an isolated, unidirectional VB-Audio Virtual Cable loopback stream, the system maintains continuous auditory monitoring with sub-500 ms latency while incurring zero recurring cloud streaming costs. An in-memory double-ended queue maintains a rolling 50-word discourse history, enabling the agent to resolve phonetic name ambiguity via normalized Levenshtein distance metrics and isolate the immediate 15-word conversational context surrounding instructor inquiries.")

    add_body_p(doc, "Upon trigger verification, the agent dynamically leverages OpenAI GPT-4o-mini to synthesize contextually grounded technical explanations or attendance confirmations, dispatching them into the Google Meet chat stream via Selenium WebDriver in under 4.0 seconds. Furthermore, the post-session serialization pipeline captures lossless 16 kHz 16-bit mono PCM audio (debug_audio.wav), executes high-fidelity batch transcription via Whisper, transforms raw spoken dialogue into structured, hierarchical lecture summaries, and automatically publishes them directly to the student's personal Notion database via the official Notion REST API.")

    add_body_p(doc, "Empirical verification across ten rigorous functional test cases confirmed that the agent operates reliably on standard consumer PCs without dedicated GPU acceleration, maintaining an average memory footprint of ~1.25 GB RAM and reducing monthly operational cloud compute expenses by 98.17% compared to pure cloud streaming alternatives. In conclusion, AI Student Agent proves that ambient, autonomous software agents can dramatically alleviate cognitive fatigue, eliminate the split-attention effect, and establish durable academic knowledge assets, making a significant, verifiable contribution to computer-assisted educational technology.")

    # 6.2 Future Scope
    add_section_heading(doc, "6.2 Future Scope")
    
    add_body_p(doc, "While the completed implementation delivers a fully functional, verified software product for real-time Google Meet classes, five ambitious engineering enhancements have been identified for future development iterations:")

    add_bullet_p(doc, " Integrate an automated temporal audio segmentation pipeline using PyDub to split multi-hour WAV recordings into sequential 10-minute chunks (<20 MB) at silent pauses, transcribing them in parallel via Whisper to bypass the 25 MB API limit.", bold_prefix="1. Dynamic Multi-Chunk Audio Segmentation Pipeline:")
    add_bullet_p(doc, " Implement an in-memory vector store (ChromaDB / FAISS) indexing course syllabi, textbooks, and past lecture slides to perform dense semantic retrieval during cold-calls, grounding LLM prompts in verified course definitions.", bold_prefix="2. Course-Grounded Retrieval-Augmented Generation (RAG):")
    add_bullet_p(doc, " Augment the agent with an episodic and semantic memory graph tracking student performance and historical misconceptions, tailoring post-session summaries to emphasize individual learning gaps.", bold_prefix="3. Persistent Longitudinal Student Knowledge Profiling:")
    add_bullet_p(doc, " Extend the execution layer with an acoustic text-to-speech engine (Edge-TTS / ElevenLabs) routed through a virtual microphone channel for spoken verbal participation, gated by student toggle controls.", bold_prefix="4. Synthetic Spoken Voice Response Engine (TTS):")
    add_bullet_p(doc, " Abstract the browser automation and audio layers to support Zoom, Microsoft Teams, and Cisco Webex, incorporating BlackHole (macOS) and PulseAudio (Linux) for cross-platform portability.", bold_prefix="5. Cross-Platform Conferencing and OS Portability:")

    # Mandatory Chapter Summary Paragraph
    add_body_p(doc, "In summary, this concluding chapter has synthesized the overarching achievements, architectural validations, and empirical outcomes of the AI Student Agent capstone project. By uniting local acoustic perception, sliding context buffering, generative reasoning, and automated cloud knowledge persistence, this work demonstrates the immense potential of ambient artificial intelligence in higher education. The outlined future scope provides a concrete, technically viable roadmap for expanding the system into multi-chunk audio processing, course-grounded RAG, and cross-platform multi-modal tutoring.")

    doc.add_page_break()
