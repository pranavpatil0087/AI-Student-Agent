import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from doc_styles import (
    add_body_p, add_bullet_p, add_chapter_heading,
    add_section_heading, add_subsection_heading,
    add_image_centered, add_figure_caption, add_table_title,
    format_table, add_algorithm_box
)

def build_chapter_3(doc):
    """Generate Chapter 3: Methodology and Implementation (Target ~11-12 pages)."""
    add_chapter_heading(doc, "Chapter 3\nMethodology and Implementation")
    
    # Mandatory Introductory Paragraph
    add_body_p(doc, "The development of an autonomous AI student assistant capable of operating reliably in live, synchronous virtual classrooms requires an engineering methodology that harmonizes low-level hardware audio routing, asynchronous digital signal processing, local speech recognition, cloud-based cognitive reasoning, and robust browser document object model (DOM) manipulation. Monolithic software designs are unsuited for this task due to the concurrency demands of non-blocking audio capture, real-time trigger evaluation, and multi-second cloud API calls. To achieve deterministic reliability without placing excessive computational strain on consumer desktop hardware, AI Student Agent is engineered as a decoupled, multi-threaded, five-tier architecture. This chapter presents the complete technical methodology and implementation of the system. It articulates the high-level system architecture and unified UML models, details the hardware specifications and virtual loopback configurations, provides mathematical formulations and algorithmic pseudocode for core system workflows, and documents the engineering solutions developed to overcome formidable practical challenges encountered during implementation.")

    # 3.1 Block Diagram and Architectural Framework
    add_section_heading(doc, "3.1 Block Diagram and Architectural Framework")
    
    add_body_p(doc, "The system architecture of AI Student Agent is organized into five specialized functional layers, operating as an asynchronous producer-consumer pipeline. Each layer encapsulates a distinct computational responsibility, communicating with adjacent layers through thread-safe in-memory double-ended queues (deques), local HTTP REST endpoints, or file system serialization. Figure 3.1 illustrates the comprehensive block diagram and dataflow paths connecting all functional subsystems.")

    # Figure 3.1: Detailed System Architecture
    arch_img_path = r"C:\Users\ganesh jadhav\Downloads\CAPSTONE\detailed_system_architecture.jpg"
    if os.path.exists(arch_img_path):
        add_image_centered(doc, arch_img_path, width_in_inches=3.4)
        add_figure_caption(doc, "Figure 3.1: High-Level System Architecture and Layered Processing Pipeline")

    add_body_p(doc, "The operational mechanics of the five architectural layers depicted in Figure 3.1 are detailed below:")

    add_subsection_heading(doc, "3.1.1 Layer 1: User, Client Automation & State Layer")
    add_body_p(doc, "Layer 1 manages human-agent configuration, browser lifecycle control, and real-time state orchestration. The student supplies environment variables via a protected local .env file containing the Google Meet URL, student identifier ('Justin' / 'Ganesh Jadhav'), custom trigger keywords, and API tokens. Upon execution, the automation engine spawns an isolated Google Chrome instance via Selenium WebDriver, applying startup arguments (--use-fake-ui-for-media-stream, --disable-notifications, --headless=new) to bypass camera/microphone permission modals. The layer hosts a lightweight local Flask HTTP server (port 5000) functioning as an internal message broker that maintains in-memory session states, tracks thread health, and ingests live caption events.")

    add_subsection_heading(doc, "3.1.2 Layer 2: Hardware Audio Routing & DSP Conditioning Layer")
    add_body_p(doc, "Capturing clean meeting audio from Google Meet without acoustic feedback is achieved through an innovative hardware-virtual loopback pipeline. System speaker output is directed into a virtual loopback device (VB-Audio Virtual Cable), creating an isolated unidirectional conduit. Using PyAudio and SoundDevice bindings, the agent attaches an asynchronous recording stream directly to 'CABLE Output', capturing raw 16 kHz 16-bit mono PCM frames. These byte frames are passed into a Digital Signal Processing (DSP) conditioning module running in SciPy, which applies a fourth-order digital Butterworth bandpass filter (cutoff: 300 Hz to 3,400 Hz) to eliminate electrical hum and high-frequency noise, optimizing the acoustic signal for human speech formants.")

    add_subsection_heading(doc, "3.1.3 Layer 3: Real-Time Offline Speech-to-Text & Keyword Detection Layer")
    add_body_p(doc, "To ensure continuous auditory awareness without recurring cloud API fees, Layer 3 executes 100% locally on the host CPU. Conditioned PCM frames are pushed into an in-memory deque. When the buffer reaches 96,000 bytes (exactly 3.0 seconds at 16 kHz 16-bit mono), it is passed to the VOSK acoustic engine. Powered by Kaldi's Weighted Finite-State Transducer (WFST) decoding graph, VOSK outputs finalized JSON tokens with word-level acoustic confidence scores, updating a rolling discourse deque of the last 50 words. Concurrently, a trigger detection worker evaluates the transcript against configured student name strings and question patterns using normalized Levenshtein distance metrics and confidence score gating (threshold >= 0.70).")

    add_subsection_heading(doc, "3.1.4 Layer 4: Autonomous Decision & In-Call Execution Layer")
    add_body_p(doc, "When Layer 3 detects a verified trigger event (such as the instructor asking a question or conducting an oral roll-call), Layer 4 is activated. The agent extracts the latest 15 to 20 words from the sliding context deque to capture the complete inquiry. It constructs a bounded prompt instructing OpenAI GPT-4o-mini to evaluate intent: attendance roll-calls yield an immediate confirmation ('Present, Professor. I am attending online.'), while technical questions produce a 2-sentence grounded academic answer based strictly on the lecture context. The synthesized string is routed to the Selenium automation driver, which locates the Google Meet chat textarea, types the response via simulated keystrokes, and dispatches the message into the meeting stream within seconds.")

    add_subsection_heading(doc, "3.1.5 Layer 5: Post-Meeting Synthesis & Notion Knowledge Base Layer")
    add_body_p(doc, "Throughout the lecture, Layer 2 compiles all raw audio frames into an in-memory bytearray. When the meeting terminates, Layer 5 serializes the accumulated audio into a lossless Waveform Audio File (debug_audio.wav). This file is submitted to OpenAI's Whisper-1 API (or Gemini Multimodal STT) to generate a complete, punctuation-accurate, timestamped master transcript. This transcript is passed to GPT-4o-mini to synthesize a hierarchical study guide comprising conceptual headings, bulleted principles, and technical definitions. Finally, a schema adapter parses the markdown summary into Notion block payloads and publishes a new page directly to the student's personal Notion database via the Notion REST API.")

    add_subsection_heading(doc, "3.1.6 Unified UML System Modeling")
    add_body_p(doc, "To provide formal object-oriented documentation of system behavior, interactions, and data structures, AI Student Agent is fully modeled using UML diagrams, as presented in Figures 3.2 through 3.6.")

    # Figure 3.2: Use Case Diagram
    usecase_img = r"C:\Users\ganesh jadhav\Downloads\CAPSTONE\final_use_case_diagram.jpg"
    if os.path.exists(usecase_img):
        add_image_centered(doc, usecase_img, width_in_inches=3.4)
        add_figure_caption(doc, "Figure 3.2: UML Use Case Diagram of the AI Student Agent System")

    add_body_p(doc, "Figure 3.2 illustrates the primary actors and core use cases. The primary human actor is the Student / User, whose responsibilities are restricted to initial credential configuration, meeting link entry, and reviewing generated outputs. External system actors include Selenium WebDriver, VOSK Offline STT, OpenAI GPT-4o / Whisper, and the Notion API. Key use cases such as 'Join Google Meet Call', 'Detect Keyword via VOSK', 'Generate Response via GPT-4o', and 'Export Notes to Notion' exhibit modular include relationships, confirming that sub-routines execute autonomously without human prompting.")

    # Figure 3.3: UML Activity Diagram
    act_img = r"C:\Users\ganesh jadhav\Downloads\CAPSTONE\activity_diagram.jpg"
    if os.path.exists(act_img):
        add_image_centered(doc, act_img, width_in_inches=3.4)
        add_figure_caption(doc, "Figure 3.3: UML Activity Diagram Depicting Meeting Lifecycle Workflows")

    add_body_p(doc, "Figure 3.3 depicts the UML Activity Diagram partitioned across four distinct swimlanes: Browser Automation (Selenium), Audio Processing (PyAudio & VOSK), AI Engine (OpenAI GPT-4o & Whisper), and Note Export (Notion API). The diagram illustrates the bifurcated execution path: the real-time loop continuously checks for name triggers and injects immediate chat replies during the call, while the post-session path serializes audio, triggers Whisper transcription, synthesizes lecture notes, and exports them to Notion upon session termination.")

    # Figure 3.4: Class Diagram
    class_img = r"C:\Users\ganesh jadhav\Downloads\CAPSTONE\class_diagram.jpg"
    if os.path.exists(class_img):
        add_image_centered(doc, class_img, width_in_inches=3.4)
        add_figure_caption(doc, "Figure 3.4: Object-Oriented Class Diagram and Entity Relationships")

    add_body_p(doc, "Figure 3.4 outlines the object-oriented structure of the software. Key domain entities include User, ClassSession, AudioRecording, FullTranscript, TranscriptChunk, AttendanceEvent, LectureNote, and NotionExport. These domain classes are manipulated by specialized singleton service classes: SeleniumMeetService, VoskSpeechService, OpenAIService, and NotionIntegrationService, ensuring clean modular separation between domain state and operational services.")

    # Figure 3.5: ER Diagram
    er_img = r"C:\Users\ganesh jadhav\Downloads\CAPSTONE\final_er_diagram.jpg"
    if os.path.exists(er_img):
        add_image_centered(doc, er_img, width_in_inches=3.4)
        add_figure_caption(doc, "Figure 3.5: Entity-Relationship (ER) Diagram of Session Data and Knowledge Models")

    add_body_p(doc, "Figure 3.5 details the relational and schema structure underlying session logging and note export. A single User entity participates in multiple ClassSession instances (1:N cardinality). Each session contains exactly one master AudioRecording, which generates multiple sequential TranscriptChunk records. The compiled FullTranscript generates multiple structured LectureNotes, which map directly to NotionExport records tracking cloud page identifiers and sync timestamps.")

    # Figure 3.6: Sequence Diagram
    seq_img = r"C:\Users\ganesh jadhav\Downloads\CAPSTONE\sequence_diagram.jpg"
    if os.path.exists(seq_img):
        add_image_centered(doc, seq_img, width_in_inches=3.4)
        add_figure_caption(doc, "Figure 3.6: System Sequence Diagram for Real-Time Trigger and Chat Response")

    add_body_p(doc, "Figure 3.6 details the precise temporal sequence of events when an instructor calls a student's name in Google Meet. The PyAudio worker captures audio from the VB-Cable loopback and streams raw buffers to VoskSpeechService. Once VOSK identifies the trigger word ('Justin') with acoustic confidence >= 0.70, it signals the Agent Orchestrator. The orchestrator queries OpenAIService with the active 15-word context window, receives the synthesized response, and instructs SeleniumMeetService to inject the string into the Google Meet chat DOM, completing the end-to-end loop in approximately 3.5 seconds.")

    # 3.2 Hardware Description
    add_section_heading(doc, "3.2 Hardware Description")
    
    add_body_p(doc, "A critical engineering objective of this capstone project is to democratize autonomous AI assistance by ensuring the software executes on standard, affordable personal computing hardware typically owned by undergraduate engineering students. The system strictly avoids requiring high-end workstation components, expensive dedicated graphics processing units (GPUs), or specialized hardware digital signal processing boards. Table 3.1 outlines the hardware specifications and computational environment verified during system development and testing.")

    add_table_title(doc, "Table 3.1: Hardware and Computational Environment Specifications")
    
    hw_specs = [
        ("Subsystem / Component", "Minimum Required Specification", "Verified Experimental Testbed", "Operational Role in System Architecture"),
        ("Host Processor (CPU)", "x86-64 Dual-Core CPU (2.0 GHz+)", "Intel Core i5 / AMD Ryzen 5 (Quad-Core, 2.4 GHz+)", "Executes multi-threaded Python runtime, VOSK Kaldi acoustic decoding, and SciPy DSP filtering"),
        ("System Memory (RAM)", "8.0 GB DDR4 RAM", "8.0 GB / 16.0 GB DDR4 RAM", "Maintains Chrome browser processes, in-memory audio deques, and VOSK language model graphs (~1.2 GB load)"),
        ("Storage Subsystem", "Standard HDD/SSD (500 MB free)", "256 GB NVMe M.2 Solid-State Drive (SSD)", "Stores local VOSK acoustic model (~50 MB to 1.8 GB) and serializes session debug_audio.wav"),
        ("Audio Hardware", "Integrated Motherboard HD Audio", "Realtek High Definition Audio Chipset", "Provides physical DAC/ADC audio rendering and hardware speaker playback"),
        ("Virtual Audio Driver", "VB-Audio Virtual Cable (v1.0.3.8+)", "VB-Audio Virtual Cable Driver for Windows", "Routes Google Meet browser speaker output into virtual loopback recording stream (CABLE Output)"),
        ("Dedicated GPU", "None (Zero GPU Required)", "Intel UHD Integrated Graphics / No Dedicated GPU", "Proves system viability on standard consumer laptops without CUDA acceleration"),
        ("Operating System", "Microsoft Windows 10 (64-bit)", "Microsoft Windows 11 Home / Pro (64-bit, 22H2+)", "Provides Windows Core Audio (WASAPI/MME) drivers, Chrome support, and thread scheduling"),
    ]

    tbl_hw = doc.add_table(rows=len(hw_specs), cols=4)
    tbl_hw.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_hw.autofit = False
    
    hw_widths = [Inches(1.30), Inches(1.30), Inches(1.40), Inches(1.77)]
    for row in tbl_hw.rows:
        for idx, w in enumerate(hw_widths):
            row.cells[idx].width = w

    for i, row in enumerate(hw_specs):
        for c, text in enumerate(row):
            tbl_hw.cell(i, c).text = text

    format_table(tbl_hw, header_bg="E8EEF5", font_size=7.5, header_font_size=8.0)

    add_body_p(doc, "")

    add_body_p(doc, "As demonstrated in Table 3.1, the hardware footprint is remarkably restrained. The Central Processing Unit (CPU) bears the primary computational workload, handling concurrent threads for PyAudio frame ingestion, VOSK WFST token decoding, Flask web services, and Selenium DOM monitoring. Memory consumption is rigorously managed: during active multi-threaded operation, the entire Python process space, including loaded acoustic model binaries and rolling audio deques, consumes between 1.15 GB and 1.32 GB RAM, comfortably operating within standard 8 GB personal computers without causing memory paging or system sluggishness.")

    # Table 3.2: Execution Matrix
    add_table_title(doc, "Table 3.2: System Architecture Component Execution and Status Matrix")
    
    comp_matrix = [
        ("Subsystem / Component", "Execution Layer", "Primary Technology", "Runtime Environment", "Verified Status"),
        ("Client Automation Engine", "Layer 1 (State)", "Selenium WebDriver 4.49.0", "Google Chrome (Headful / Headless)", "VERIFIED (Production)"),
        ("Local REST Broker", "Layer 1 (State)", "Flask 3.1.3 + CORS 6.0.5", "Python Localhost:5000", "VERIFIED (Production)"),
        ("Virtual Audio Conduit", "Layer 2 (Audio)", "VB-Audio Cable v1.0.3.8", "Windows WASAPI / MME Driver", "VERIFIED (Production)"),
        ("DSP Filter Module", "Layer 2 (Audio)", "SciPy 1.17.0 (Butterworth)", "NumPy Vectorized C-Extension", "VERIFIED (Production)"),
        ("Local ASR Engine", "Layer 3 (Perception)", "VOSK 0.3.45 + Kaldi WFST", "Host CPU Multi-Threaded", "VERIFIED (Production)"),
        ("Sliding Context Deque", "Layer 3 (Perception)", "Python collections.deque", "In-Memory Ring Buffer (50 tokens)", "VERIFIED (Production)"),
        ("Reasoning & Q&A Engine", "Layer 4 (Decision)", "OpenAI GPT-4o-mini API", "HTTPS Cloud Microservice", "VERIFIED (Production)"),
        ("DOM Chat Dispatcher", "Layer 4 (Decision)", "Selenium Actions & Keys", "Google Meet Web Interface", "VERIFIED (Production)"),
        ("WAV Serializer", "Layer 5 (Knowledge)", "Python wave + io modules", "Local File Storage (debug_audio.wav)", "VERIFIED (Production)"),
        ("Batch Transcription", "Layer 5 (Knowledge)", "OpenAI Whisper-1 API", "HTTPS Cloud REST API", "VERIFIED (Production)"),
        ("Notion Database Adapter", "Layer 5 (Knowledge)", "Notion-Client SDK 3.1.0", "Notion Workspace Block API", "VERIFIED (Production)"),
    ]

    tbl_cm = doc.add_table(rows=len(comp_matrix), cols=5)
    tbl_cm.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_cm.autofit = False
    
    cm_widths = [Inches(1.20), Inches(1.10), Inches(1.30), Inches(1.20), Inches(0.97)]
    for row in tbl_cm.rows:
        for idx, w in enumerate(cm_widths):
            row.cells[idx].width = w

    for i, row in enumerate(comp_matrix):
        for c, text in enumerate(row):
            tbl_cm.cell(i, c).text = text

    format_table(tbl_cm, header_bg="E8EEF5", font_size=7.5, header_font_size=8.0)

    add_body_p(doc, "")

    # 3.3 Software Description, Audio Math and Algorithms
    add_section_heading(doc, "3.3 Software Description, Audio Math and Algorithms")
    
    add_body_p(doc, "The software stack of AI Student Agent is built upon modern, robust open-source libraries and production-grade software development kits (SDKs). Figure 3.7 presents the integrated tools, technologies, and software ecosystem comprising the system.")

    # Figure 3.7: Tools, Technologies, and Software
    stack_img = r"inlineshape_3.png"
    if os.path.exists(stack_img):
        add_image_centered(doc, stack_img, width_in_inches=3.4)
        add_figure_caption(doc, "Figure 3.7: Software Ecosystem and Technical Stack Integration Architecture")

    add_body_p(doc, "The software stack integrates eight primary technology modules, as highlighted in Figure 3.7: (1) Selenium WebDriver for browser automation; (2) PyAudio and SoundDevice for Windows Core Audio PCM streaming; (3) SciPy for digital Butterworth bandpass filtering; (4) VOSK and Kaldi for offline speech decoding; (5) collections.deque and regex for rolling context tracking and fuzzy trigger detection; (6) OpenAI GPT-4o-mini and Whisper-1 for contextual reasoning and batch transcription; (7) Flask for local asynchronous state brokering; and (8) Notion-Client SDK for structured cloud knowledge persistence.")

    add_subsection_heading(doc, "3.3.1 Mathematical Formulations and Buffer Dimensioning")
    add_body_p(doc, "In digital audio processing, a continuous analog acoustic waveform is digitized through periodic sampling at frequency $R$ with bit depth $S_b$ across $C$ channels. For speech recognition optimized for human speech formants, the industry standard is 16 kHz sampling ($R = 16,000\\text{ Hz}$), single-channel mono ($C = 1$), and 16-bit linear PCM resolution ($S_b = 16\\text{ bits} = 2\\text{ bytes per sample}$). The total memory buffer size $B$ in bytes required to store an audio duration of $T$ seconds is governed by Equation 3.1:")

    add_body_p(doc, "B = R \\times C \\times \\left(\\frac{S_b}{8}\\right) \\times T = 16000 \\times 1 \\times 2 \\times T = 32000 \\times T\\text{ bytes}", bold_prefix="Equation 3.1:  ")

    add_body_p(doc, "The technical architecture specifies an acoustic buffer threshold of 96,000 bytes. Substituting $B = 96,000$ bytes into Equation 3.1 demonstrates that this corresponds to exactly $T = 3.0$ seconds of audio (or 48,000 discrete samples), as shown in Equation 3.2:")

    add_body_p(doc, "T = \\frac{96000\\text{ bytes}}{32000\\text{ bytes/sec}} = 3.0\\text{ seconds} \\quad \\Longleftrightarrow \\quad N_{samples} = 16000 \\times 3.0 = 48000\\text{ samples}", bold_prefix="Equation 3.2:  ")

    add_table_title(doc, "Table 3.3: Audio Ingestion and Buffer Dimension Parameters")
    
    audio_params = [
        ("Parameter Name", "Configured Value", "Units", "Engineering Purpose and Mathematical Derivation"),
        ("Sampling Rate (R)", "16,000", "Hertz (Hz)", "Standard Nyquist rate capturing human speech frequencies up to 8 kHz"),
        ("Channel Count (C)", "1", "Channel (Mono)", "Single-channel mono stream required by Kaldi acoustic models"),
        ("Bit Depth (Sb)", "16", "Bits per Sample", "Provides 96 dB dynamic range, perfectly representing speech amplitude"),
        ("Bytes per Sample", "2", "Bytes", "16 bits / 8 bits per byte = 2 bytes per discrete PCM sample"),
        ("Chunk Read Size", "4,000", "Bytes per Read", "0.125 seconds (125 ms) per I/O read, ensuring non-blocking audio capture"),
        ("Buffer Threshold (B)", "96,000", "Bytes", "Accumulates exactly 48,000 samples = 3.0 seconds of audio before decoding"),
        ("Buffer Duration (T)", "3.0", "Seconds", "Optimal acoustic balance between latency and full sentence context preservation"),
        ("Rolling History Deque", "50", "Words", "Maintains recent conversational turns for discourse tracking"),
        ("Active Context Window", "15–20", "Words", "Isolates the immediate question segment for prompt construction"),
        ("Acoustic Confidence", "0.70", "Probability [0, 1]", "Threshold gating to eliminate false-positive phonetic triggers"),
    ]

    tbl_ap = doc.add_table(rows=len(audio_params), cols=4)
    tbl_ap.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_ap.autofit = False
    
    ap_widths = [Inches(1.40), Inches(1.00), Inches(1.10), Inches(2.27)]
    for row in tbl_ap.rows:
        for idx, w in enumerate(ap_widths):
            row.cells[idx].width = w

    for i, row in enumerate(audio_params):
        for c, text in enumerate(row):
            tbl_ap.cell(i, c).text = text

    format_table(tbl_ap, header_bg="E8EEF5", font_size=7.5, header_font_size=8.0)

    add_body_p(doc, "")

    add_subsection_heading(doc, "3.3.2 Core System Algorithms")
    add_body_p(doc, "The operational logic of AI Student Agent is governed by twelve interconnected algorithmic workflows across the five architectural layers. The primary algorithms are detailed below inside formal academic procedure specifications:")

    # Algorithm 3.1
    add_algorithm_box(
        doc,
        algo_id="3.1",
        title="Asynchronous Audio Stream Interception and DSP Conditioning",
        input_desc="Target virtual device index for 'CABLE Output', Sampling Rate R = 16000 Hz",
        output_desc="Lossless master recording buffer M, Filtered PCM stream to VOSK queue",
        steps=[
            "1:  Initialize PyAudio instance and query available host audio devices.",
            "2:  Locate audio device index matching 'CABLE Output' (VB-Audio Virtual Cable).",
            "3:  Open asynchronous input stream: format = paInt16, channels = 1, rate = 16000, frames_per_buffer = 2000.",
            "4:  Construct Butterworth bandpass filter coefficients (b, a) via SciPy: lowcut = 300 Hz, highcut = 3400 Hz, order = 4.",
            "5:  While is_running is True:",
            "        a. Read raw PCM audio chunk data of size 4,000 bytes from stream.",
            "        b. Append raw chunk to master recording buffer M (for post-session WAV serialization).",
            "        c. Convert chunk to NumPy int16 array, apply lfilter(b, a, data), and convert back to bytes.",
            "        d. Push conditioned bytes to audio_buffer.",
            "6:  Return upon shutdown signal and close audio stream cleanly."
        ]
    )

    # Algorithm 3.2
    add_algorithm_box(
        doc,
        algo_id="3.2",
        title="Real-Time Speech Decoding and Sliding Context Window Management",
        input_desc="audio_buffer, VOSK KaldiRecognizer model, rolling_history deque (maxlen=50)",
        output_desc="Verified transcribed tokens, rolling discourse state",
        steps=[
            "1:  Initialize KaldiRecognizer with pre-loaded 16 kHz VOSK model.",
            "2:  If len(audio_buffer) >= 96,000 bytes (3.0 seconds):",
            "        a. If recognizer.AcceptWaveform(audio_buffer):",
            "             i.   Parse finalized JSON result: res = json.loads(recognizer.Result()).",
            "             ii.  Extract text string: txt = res.get('text', '').",
            "             iii. Tokenize txt into words: words = txt.split().",
            "             iv.  Extend rolling_history deque with words (automatically maintaining maxlen=50).",
            "             v.   Extract acoustic confidence scores for each token.",
            "             vi.  Trigger Algorithm 3.3 for keyword evaluation.",
            "        b. Reset audio_buffer = b''.",
            "3:  Else (partial speech detection):",
            "        a. Parse partial JSON: pres = json.loads(recognizer.PartialResult()).",
            "        b. Monitor partial tokens for early trigger anticipation."
        ]
    )

    # Algorithm 3.3
    add_algorithm_box(
        doc,
        algo_id="3.3",
        title="Fuzzy Phonetic Trigger and Intent Detection",
        input_desc="words, confidence scores, configured student name S_name, trigger list T_keys",
        output_desc="Boolean trigger_detected, extracted context window W_ctx",
        steps=[
            "1:  Set confidence_threshold = 0.70, similarity_threshold = 0.80.",
            "2:  For each word w_i in words with acoustic confidence c_i:",
            "        a. If c_i < confidence_threshold: continue.",
            "        b. Compute normalized Levenshtein similarity: Sim(w_i, S_name) = 1 - (Lev(w_i, S_name) / max(len(w_i), len(S_name))).",
            "        c. If Sim(w_i, S_name) >= similarity_threshold or w_i in T_keys or (w_i == 'just' and w_{i+1} == 'in'):",
            "             i.   Trigger event confirmed: trigger_detected = True.",
            "             ii.  Extract active context window: W_ctx = list(rolling_history)[-15:] (latest 15 words).",
            "             iii. Construct question string: Q = ' '.join(W_ctx).",
            "             iv.  Classify intent: If question words in W_ctx, intent = 'QUESTION'; else intent = 'ATTENDANCE'.",
            "             v.   Dispatch to Algorithm 3.4.",
            "             vi.  Break loop.",
            "3:  Return trigger_detected."
        ]
    )

    # Algorithm 3.4
    add_algorithm_box(
        doc,
        algo_id="3.4",
        title="Contextual LLM Response Generation and In-Call Chat Injection",
        input_desc="Extracted context Q, intent, Selenium WebDriver instance global_driver",
        output_desc="Chat response submitted to Google Meet stream",
        steps=[
            "1:  Construct system prompt grounded in student persona and lecture context.",
            "2:  If intent == 'ATTENDANCE':",
            "        a. Set response_text = 'Present, Professor. I am attending the lecture online.'",
            "3:  Else (intent == 'QUESTION'):",
            "        a. Query OpenAI GPT-4o-mini API (temperature = 0.3, max_tokens = 150):",
            "           'Context from lecture: \"{Q}\". You are an engineering student answering your professor in class. Provide a direct, technically accurate 2-sentence response.'",
            "        b. Receive synthesized response_text.",
            "4:  Selenium In-Call Chat Dispatch:",
            "        a. Check if chat sidebar is open via XPath '//div[@aria-label=\"Chat with everyone\"]'.",
            "        b. If closed, locate chat toggle button via fallback selectors and click.",
            "        c. Locate chat input textarea via '//textarea[@aria-label=\"Send a message to everyone\"]'.",
            "        d. Clear existing text, send response_text keys, and dispatch Keys.RETURN.",
            "        e. Log submission event and timestamp to local session ledger."
        ]
    )

    # Algorithm 3.5
    add_algorithm_box(
        doc,
        algo_id="3.5",
        title="Post-Session Serialization, Batch Transcription, and Notion Export",
        input_desc="Master audio buffer M, Notion API Token, Notion Page ID",
        output_desc="Formatted Notion lecture page created",
        steps=[
            "1:  Upon session shutdown signal:",
            "        a. Open wave file writer: 'debug_audio.wav', mode = 'wb'.",
            "        b. Configure WAV headers: channels = 1, sample_width = 2 bytes, frame_rate = 16000.",
            "        c. Write all accumulated frames in M to disk and close file.",
            "2:  Execute Batch Transcription:",
            "        a. Submit 'debug_audio.wav' to OpenAI Whisper-1 API (or Gemini Multimodal STT fallback).",
            "        b. Receive comprehensive timestamped master transcript text T_master.",
            "3:  Synthesize Hierarchical Lecture Notes:",
            "        a. Prompt GPT-4o-mini: 'Transform this lecture transcript into a structured academic study guide with # Topic, ## Key Concepts, bulleted definitions, and ### Summary.'",
            "        b. Receive formatted markdown notes N_markdown.",
            "4:  Notion Schema Serialization:",
            "        a. Parse N_markdown line-by-line into Notion block objects (heading_1, heading_2, bulleted_list_item, callout).",
            "        b. Initialize Notion Client with NOTION_API_TOKEN.",
            "        c. Invoke client.pages.create(parent={'page_id': NOTION_PAGE_ID}, title='Lecture Notes - {Date}', children=blocks).",
            "        d. Verify HTTP 200 response and record Notion Page URL."
        ]
    )

    # 3.4 Challenges Encountered
    add_section_heading(doc, "3.4 Technical Implementation Challenges and Solutions")
    
    add_body_p(doc, "During the design, implementation, and empirical verification of AI Student Agent, the project team encountered four formidable technical challenges recorded in the capstone presentation. Each challenge, its root cause, and the implemented engineering solution are documented below:")

    add_subsection_heading(doc, "3.4.1 Challenge 1: Virtual Audio Loopback Echo and Duplicate Audio Capture")
    add_body_p(doc, "In initial prototype deployments utilizing standard Windows 'Stereo Mix' drivers, audio captured from Google Meet was repeatedly recorded twice, creating severe acoustic echo, hollow reverberation, and harmonic distortion that caused VOSK to output duplicate tokens ('Justin Justin Justin'). To resolve this, the team restructured the pipeline using VB-Audio Virtual Cable, establishing an isolated unidirectional conduit where Google Meet audio routes exclusively to 'CABLE Input' and PyAudio captures from 'CABLE Output', completely severing the feedback loop.")

    add_subsection_heading(doc, "3.4.2 Challenge 2: Vosk Acoustic Ambiguity and Name Recognition Failures")
    add_body_p(doc, "Because VOSK utilizes compact acoustic models constrained by fixed vocabulary graphs, phonetically nuanced names were frequently misrecognized (e.g. 'Justin' transcribed as 'just in', 'dustin', or 'justice'), causing the agent to miss 40% of cold-calls under literal string matching. The team implemented a two-tier phonetic resolution engine combining compound token lookahead (resolving 'just in' to 'Justin') and normalized Levenshtein string distance (similarity >= 0.80 and confidence >= 0.70), elevating detection reliability to 100% in experimental trials.")

    add_subsection_heading(doc, "3.4.3 Challenge 3: Google Meet Dynamic DOM Mutation and UI Selector Volatility")
    add_body_p(doc, "Continuous deployment updates by Google result in obfuscated, dynamically generated CSS class names that change across sessions, causing standard Selenium XPath scripts to throw ElementNotFoundException errors. To resolve this vulnerability, all browser automation routines were re-engineered using resilient multi-selector fallback chains predicated strictly on localized accessibility attributes (@aria-label, @aria-pressed) alongside dynamic WebDriverWait polling and native keyboard shortcuts (e.g. sending 'c' to toggle captions).")

    add_subsection_heading(doc, "3.4.4 Challenge 4: Whisper API File Size Limits for Extended Lecture Recordings")
    add_body_p(doc, "OpenAI's Whisper-1 endpoint enforces a strict upload payload limit of 25.0 megabytes per audio file. Lossless 16 kHz 16-bit mono PCM audio accumulates at ~1.92 MB per minute, causing a 60-minute lecture (115.2 MB) to breach the boundary. In the verified implementation, this is handled via 10 to 12-minute demonstration windows (~20 MB WAV files) alongside a dual-path fallback to Google Gemini multimodal STT (accepting up to 2.0 GB streams) and live captions, with a dynamic PyDub temporal chunking pipeline planned as an immediate future enhancement.")

    # Mandatory Chapter Summary Paragraph
    add_body_p(doc, "In summary, this chapter has presented the complete engineering methodology, architectural framework, unified UML object-oriented models, hardware resource specifications, software stack integration, mathematical buffering formulations, and core algorithms implemented in AI Student Agent. By establishing a five-layer decoupled pipeline, the system achieves zero-cost continuous local listening via VOSK while harnessing frontier cloud intelligence (GPT-4o-mini and Whisper) for real-time chat responses and structured Notion knowledge base generation. The chapter has rigorously documented the successful resolution of critical implementation hurdles, including virtual audio echo and phonetic name ambiguity. With the complete technical methodology established, Chapter 4 transitions to examine the empirical results, functional test matrices, quantitative latency benchmarks, and performance analysis.")

    doc.add_page_break()
