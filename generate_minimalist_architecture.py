import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np

def create_minimalist_architecture(output_path="system_architecture_minimalist.png"):
    # High resolution canvas (16:9 ratio)
    fig, ax = plt.subplots(figsize=(16, 9), dpi=300)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')
    
    # Remove all axis borders and ticks
    ax.axis('off')
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)

    # Color Palette (Minimalist, modern academic)
    BG_CARD = '#F8FAFC'        # Slate 50
    BORDER_CARD = '#CBD5E1'    # Slate 300
    TEXT_MAIN = '#0F172A'      # Slate 900
    TEXT_MUTED = '#475569'     # Slate 600
    ACCENT_RED = '#BC312C'     # NMIMS Red
    ACCENT_BLUE = '#2563EB'    # Royal Blue
    ACCENT_GREEN = '#16A34A'   # Emerald Green
    ACCENT_AMBER = '#D97706'   # Amber
    ARROW_COLOR = '#64748B'    # Slate 500

    FONT_FAMILY = 'DejaVu Sans'

    # Helper: Draw rounded card box
    def draw_card(x, y, w, h, title, subtitle=None, details=None, 
                  border_color=BORDER_CARD, fill_color=BG_CARD, 
                  badge_text=None, badge_color=ACCENT_RED, title_color=TEXT_MAIN):
        # Card body
        card = FancyBboxPatch((x, y), w, h,
                              boxstyle="round,pad=0.3,rounding_size=1.2",
                              facecolor=fill_color, edgecolor=border_color,
                              linewidth=1.5, zorder=2)
        ax.add_patch(card)
        
        # Badge (optional)
        if badge_text:
            badge = FancyBboxPatch((x + 1.2, y + h - 2.8), len(badge_text)*1.0 + 1.5, 2.0,
                                   boxstyle="round,pad=0.2,rounding_size=0.6",
                                   facecolor=badge_color, edgecolor='none', zorder=3)
            ax.add_patch(badge)
            ax.text(x + 1.2 + (len(badge_text)*1.0 + 1.5)/2, y + h - 1.8, badge_text,
                    color='#FFFFFF', fontsize=7.5, fontweight='bold',
                    ha='center', va='center', zorder=4, fontfamily=FONT_FAMILY)

        # Title
        title_y = y + h - (3.8 if badge_text else 2.2)
        ax.text(x + w/2, title_y, title, color=title_color,
                fontsize=11, fontweight='bold', ha='center', va='center',
                zorder=3, fontfamily=FONT_FAMILY)
        
        # Subtitle
        curr_y = title_y - 2.0
        if subtitle:
            ax.text(x + w/2, curr_y, subtitle, color=TEXT_MUTED,
                    fontsize=8.5, fontstyle='italic', ha='center', va='center',
                    zorder=3, fontfamily=FONT_FAMILY)
            curr_y -= 1.8
            
        # Divider line if details exist
        if details:
            ax.plot([x + 1.5, x + w - 1.5], [curr_y + 0.6, curr_y + 0.6], 
                    color=border_color, linewidth=0.8, linestyle='--', zorder=3)
            curr_y -= 1.4
            for det in details:
                ax.text(x + 2.0, curr_y, "• " + det, color=TEXT_MUTED,
                        fontsize=8, ha='left', va='center',
                        zorder=3, fontfamily=FONT_FAMILY)
                curr_y -= 1.8

    # Helper: Draw clean arrow
    def draw_arrow(x1, y1, x2, y2, label=None, label_side='right', color=ARROW_COLOR, style='->', lw=1.5):
        arrow = FancyArrowPatch((x1, y1), (x2, y2),
                                arrowstyle=style, mutation_scale=14,
                                color=color, linewidth=lw, zorder=5)
        ax.add_patch(arrow)
        if label:
            mx, my = (x1 + x2)/2, (y1 + y2)/2
            offset_x = 1.0 if label_side == 'right' else -1.0
            ha = 'left' if label_side == 'right' else 'right'
            ax.text(mx + offset_x, my, label, color=color,
                    fontsize=7.5, fontweight='bold', ha=ha, va='center',
                    zorder=6, fontfamily=FONT_FAMILY,
                    bbox=dict(boxstyle="round,pad=0.2", facecolor="#FFFFFF", edgecolor="none", alpha=0.9))

    # ==========================================
    # HEADER SECTION
    # ==========================================
    ax.text(50, 96.5, "AI STUDENT AGENT — SYSTEM ARCHITECTURE", 
            color=TEXT_MAIN, fontsize=18, fontweight='bold', ha='center', va='center', fontfamily=FONT_FAMILY)
    ax.text(50, 93.8, "An Autonomous Multi-Threaded Edge-Cloud Architecture for Real-Time Google Meet Classrooms", 
            color=TEXT_MUTED, fontsize=10.5, ha='center', va='center', fontfamily=FONT_FAMILY)
    ax.plot([15, 85], [92.0, 92.0], color=BORDER_CARD, linewidth=1.0)

    # ==========================================
    # 1. TOP ACTOR & MEETING LAYER (y: 77 to 89)
    # ==========================================
    # Background Grouping Box
    layer1_box = FancyBboxPatch((4, 76), 92, 14.5,
                                boxstyle="round,pad=0.5,rounding_size=1.5",
                                facecolor='#F1F5F9', edgecolor='#E2E8F0', linewidth=1.0, zorder=1)
    ax.add_patch(layer1_box)
    ax.text(6, 89.2, "LAYER 1: CLIENT AUTOMATION & MEETING INTERFACE", 
            color='#64748B', fontsize=8.5, fontweight='bold', ha='left', va='center', fontfamily=FONT_FAMILY)

    # Student Node
    draw_card(7, 78, 18, 10.5, "Student / Attendee", "Local User & Config", 
              details=["Loads credentials (.env)", "Monitors agent dashboard"],
              border_color='#94A3B8', fill_color='#FFFFFF', badge_text="HUMAN ACTOR", badge_color='#64748B')

    # Selenium WebDriver
    draw_card(32, 78, 28, 10.5, "Selenium WebDriver Chrome", "Browser Automation Engine",
              details=["Launches Chrome with bypass flags", "Automates lobby 'Ask to join'", "Resilient DOM fallback chains"],
              border_color=ACCENT_BLUE, fill_color='#FFFFFF', badge_text="AUTOMATION", badge_color=ACCENT_BLUE)

    # Google Meet
    draw_card(67, 78, 26, 10.5, "Google Meet Web Client", "Live Synchronous Classroom",
              details=["Live instructor speech audio", "In-call public chat stream", "Participant state & events"],
              border_color=ACCENT_RED, fill_color='#FFFFFF', badge_text="TARGET ENV", badge_color=ACCENT_RED)

    # Connectors Layer 1
    draw_arrow(25, 83.2, 32, 83.2, label="Configures", label_side='right')
    draw_arrow(60, 83.2, 67, 83.2, label="Injects", label_side='right')

    # ==========================================
    # 2. AUDIO ROUTING & DSP CONDUIT (y: 59 to 72)
    # ==========================================
    layer2_box = FancyBboxPatch((4, 58.5), 92, 14.5,
                                boxstyle="round,pad=0.5,rounding_size=1.5",
                                facecolor='#F8FAFC', edgecolor='#E2E8F0', linewidth=1.0, zorder=1)
    ax.add_patch(layer2_box)
    ax.text(6, 71.7, "LAYER 2: HARDWARE AUDIO LOOPBACK & DSP CONDITIONING", 
            color='#64748B', fontsize=8.5, fontweight='bold', ha='left', va='center', fontfamily=FONT_FAMILY)

    # VB-Audio Cable
    draw_card(7, 60.5, 26, 10.0, "VB-Audio Virtual Cable", "Isolated Virtual Loopback",
              details=["Unidirectional conduit", "Eliminates Stereo Mix echo", "Zero microphone feedback"],
              border_color=ACCENT_AMBER, fill_color='#FFFFFF', badge_text="VIRTUAL CONDUIT", badge_color=ACCENT_AMBER)

    # PyAudio Capture
    draw_card(38, 60.5, 26, 10.0, "PyAudio Ingestion Stream", "Real-Time Audio Worker",
              details=["16 kHz 16-bit mono PCM", "4,000-byte chunks (125 ms)", "Non-blocking ring deque"],
              border_color=ACCENT_BLUE, fill_color='#FFFFFF', badge_text="STREAMING", badge_color=ACCENT_BLUE)

    # SciPy Butterworth Filter
    draw_card(69, 60.5, 24, 10.0, "SciPy DSP Filter", "4th-Order Butterworth",
              details=["Bandpass: 300 – 3,400 Hz", "Removes 50 Hz hum & hiss", "Isolates speech formants"],
              border_color=ACCENT_GREEN, fill_color='#FFFFFF', badge_text="DSP MODULE", badge_color=ACCENT_GREEN)

    # Connectors Meet -> Audio -> DSP
    draw_arrow(80, 78, 20, 70.5, label="Speaker Audio Stream", label_side='left', color=ACCENT_AMBER)
    draw_arrow(33, 65.5, 38, 65.5, label="PCM Data", label_side='right')
    draw_arrow(64, 65.5, 69, 65.5, label="Raw Chunks", label_side='right')

    # ==========================================
    # 3. REAL-TIME PERCEPTION ENGINE (y: 37 to 55)
    # ==========================================
    layer3_box = FancyBboxPatch((4, 37.5), 92, 17.5,
                                boxstyle="round,pad=0.5,rounding_size=1.5",
                                facecolor='#EFF6FF', edgecolor='#BFDBFE', linewidth=1.2, zorder=1)
    ax.add_patch(layer3_box)
    ax.text(6, 53.7, "LAYER 3: EDGE SPEECH RECOGNITION & TRIGGER DETECTION (100% HOST CPU — $0 CLOUD COST)", 
            color=ACCENT_BLUE, fontsize=8.5, fontweight='bold', ha='left', va='center', fontfamily=FONT_FAMILY)

    # VOSK Kaldi Engine
    draw_card(7, 39.5, 27, 13.0, "VOSK Kaldi Engine", "Offline Speech Recognition",
              details=["Factorized TDNN model (112 MB)", "Kaldi Online2 WFST decoder", "Sub-500ms decoding latency", "Runs locally on CPU ($0/mo)"],
              border_color=ACCENT_RED, fill_color='#FFFFFF', badge_text="OFFLINE ASR", badge_color=ACCENT_RED)

    # Sliding Context Deque
    draw_card(38, 39.5, 27, 13.0, "Sliding Context Deque", "Discourse Memory Buffer",
              details=["collections.deque(maxlen=50)", "Rolling conversation history", "Extracts recent 15-word context", "Prevents LLM hallucinations"],
              border_color=ACCENT_BLUE, fill_color='#FFFFFF', badge_text="RING BUFFER", badge_color=ACCENT_BLUE)

    # Fuzzy Trigger Detector
    draw_card(69, 39.5, 24, 13.0, "Trigger & Intent Detector", "Acoustic Confidence Gating",
              details=["Levenshtein similarity (>= 0.80)", "Confidence score filter (>= 0.70)", "Lookahead: 'just in' -> 'Justin'", "Attendance vs Technical Q&A"],
              border_color=ACCENT_GREEN, fill_color='#FFFFFF', badge_text="FUZZY MATCHER", badge_color=ACCENT_GREEN)

    # Connectors DSP -> Layer 3
    draw_arrow(81, 60.5, 20, 52.5, label="Filtered Audio Buffer (96KB = 3.0s)", label_side='left', color=ACCENT_BLUE)
    draw_arrow(34, 46.0, 38, 46.0, label="Tokens", label_side='right')
    draw_arrow(65, 46.0, 69, 46.0, label="15-Word Context", label_side='right')

    # ==========================================
    # 4. LOWER DUAL-PATH WORKFLOWS (y: 6 to 34)
    # ==========================================
    # Path A: Real-Time In-Call Execution (Left)
    path_a_box = FancyBboxPatch((4, 6), 44, 28,
                                boxstyle="round,pad=0.5,rounding_size=1.5",
                                facecolor='#FEF2F2', edgecolor='#FECACA', linewidth=1.2, zorder=1)
    ax.add_patch(path_a_box)
    ax.text(6, 32.5, "PATH A: REAL-TIME IN-CALL Q&A LOOP (< 3.5s)", 
            color=ACCENT_RED, fontsize=8.5, fontweight='bold', ha='left', va='center', fontfamily=FONT_FAMILY)

    # OpenAI GPT-4o-mini
    draw_card(6.5, 17.5, 39, 13.0, "OpenAI GPT-4o-mini", "Contextual Cognitive Reasoning",
              details=["Bounded prompt with 15-word lecture slice", "Attendance: 'Present, Professor.'", "Technical Query: 2-sentence grounded reply", "Sub-2.5s cloud latency ($0.001/query)"],
              border_color=ACCENT_RED, fill_color='#FFFFFF', badge_text="CLOUD LLM", badge_color=ACCENT_RED)

    # DOM Chat Injector
    draw_card(6.5, 8.0, 39, 8.0, "In-Call DOM Chat Injector", "Selenium Action Dispatcher",
              details=["Locates chat textarea via dynamic XPath", "Types answer & dispatches Keys.RETURN (0.8s)"],
              border_color=ACCENT_BLUE, fill_color='#FFFFFF', badge_text="CHAT DISPATCH", badge_color=ACCENT_BLUE)

    draw_arrow(26, 17.5, 26, 16.0, label="Synthesized Reply", label_side='right', color=ACCENT_RED)

    # Path B: Post-Session Knowledge Base (Right)
    path_b_box = FancyBboxPatch((52, 6), 44, 28,
                                boxstyle="round,pad=0.5,rounding_size=1.5",
                                facecolor='#F0FDF4', edgecolor='#BBF7D0', linewidth=1.2, zorder=1)
    ax.add_patch(path_b_box)
    ax.text(54, 32.5, "PATH B: POST-SESSION SYNTHESIS & NOTION SYNC", 
            color=ACCENT_GREEN, fontsize=8.5, fontweight='bold', ha='left', va='center', fontfamily=FONT_FAMILY)

    # Serialization & Whisper
    draw_card(54.5, 17.5, 39, 13.0, "OpenAI Whisper-1 API", "Verbatim Master Transcription",
              details=["Serialized lossless audio (debug_audio.wav)", "Timestamped full lecture master transcript", "Preserves complex technical terminology", "Dual-path fallback for extended files"],
              border_color=ACCENT_GREEN, fill_color='#FFFFFF', badge_text="POST ASR", badge_color=ACCENT_GREEN)

    # Notion Integration
    draw_card(54.5, 8.0, 39, 8.0, "Notion REST API Adapter", "Knowledge Base Cloud Sync",
              details=["Synthesizes hierarchical notes (# Topic, ## Concepts)", "Parses markdown to Notion block trees in student DB"],
              border_color='#0D9488', fill_color='#FFFFFF', badge_text="NOTION SYNC", badge_color='#0D9488')

    draw_arrow(74, 17.5, 74, 16.0, label="Lecture Summary", label_side='right', color=ACCENT_GREEN)

    # Connectors from Layer 3 to Paths
    # Trigger to GPT-4o-mini
    draw_arrow(75, 39.5, 35, 30.5, label="Verified Trigger Event", label_side='left', color=ACCENT_RED, lw=2.0)
    # PyAudio / Session End to Whisper
    draw_arrow(85, 39.5, 74, 30.5, label="Session End / Audio Dump", label_side='right', color=ACCENT_GREEN, lw=2.0)

    # Feedback loop from Chat Injector back up to Google Meet!
    # Curved path or multi-segment arrow back to Google Meet
    ax.annotate("", xy=(46, 78), xytext=(45.5, 12),
                arrowprops=dict(arrowstyle="->", color=ACCENT_RED, lw=2.0,
                                connectionstyle="arc3,rad=-0.35"))
    ax.text(49, 48, "Auto-Injected\nChat Message\n(< 3.5s total)", 
            color=ACCENT_RED, fontsize=8, fontweight='bold', ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.3", facecolor="#FFF1F2", edgecolor=ACCENT_RED, alpha=0.95))

    # Footer note
    ax.text(50, 2.5, "AI Student Agent Architecture | SVKM's NMIMS MPSTME | Dept. of Information Technology 2026-27",
            color='#94A3B8', fontsize=8.5, ha='center', va='center', fontfamily=FONT_FAMILY)

    plt.tight_layout()
    plt.savefig(output_path, dpi=300, facecolor='#FFFFFF', edgecolor='none')
    plt.close()
    print(f"Minimalist architecture diagram generated successfully at: {output_path}")

if __name__ == '__main__':
    create_minimalist_architecture()
