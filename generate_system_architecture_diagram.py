import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Rectangle, Ellipse, PathPatch
from matplotlib.path import Path
import numpy as np
import os

def generate_system_architecture(output_path="system_architecture_diagram.png"):
    # Set high-resolution figure (approx 16:9 ratio, matching reference)
    fig, ax = plt.subplots(figsize=(13, 7.5), dpi=300)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')
    
    # Coordinate space: 0 to 100 on X, 0 to 60 on Y
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 60)
    ax.axis('off')
    
    # Fonts and Styling
    SERIF_FONT = 'Times New Roman'
    SANS_FONT = 'Arial'
    
    FILL_GRAY = '#E6E6E6'
    BORDER_COLOR = '#000000'
    LINE_WIDTH = 1.4
    TEXT_BLACK = '#000000'
    
    # 1. Main Title at Top
    ax.text(50, 56.8, "System Architecture", fontfamily=SERIF_FONT, 
            fontsize=26, fontweight='bold', color=TEXT_BLACK, ha='center', va='center')
    
    # Helper: Draw rounded box with optional gray header
    def draw_box(x, y, w, h, title, subtitle=None, bullet_points=None, is_split=False):
        # Base Card
        box = FancyBboxPatch((x, y), w, h,
                             boxstyle="round,pad=0,rounding_size=0.8",
                             facecolor='#FFFFFF' if is_split else FILL_GRAY,
                             edgecolor=BORDER_COLOR,
                             linewidth=LINE_WIDTH, zorder=2)
        ax.add_patch(box)
        
        if is_split:
            header_h = h * 0.28
            header = Rectangle((x, y + h - header_h), w, header_h,
                               facecolor=FILL_GRAY, edgecolor='none', zorder=3)
            header.set_clip_path(box)
            ax.add_patch(header)
            
            # Divider line
            ax.plot([x, x + w], [y + h - header_h, y + h - header_h], 
                    color=BORDER_COLOR, linewidth=LINE_WIDTH, zorder=4)
            
            # Header Title
            ax.text(x + w/2, y + h - header_h/2, title,
                    fontfamily=SANS_FONT, fontsize=10.5, fontweight='bold',
                    color=TEXT_BLACK, ha='center', va='center', zorder=5)
            
            # Bullet items inside white body
            if bullet_points:
                n = len(bullet_points)
                spacing = (h - header_h) / (n + 1)
                for i, item in enumerate(bullet_points):
                    item_y = (y + h - header_h) - (i + 1) * spacing + 0.2
                    ax.text(x + 0.8, item_y, item,
                            fontfamily=SANS_FONT, fontsize=8.6,
                            color=TEXT_BLACK, ha='left', va='center', zorder=5)
        else:
            # Solid box with Title & Subtitle centered
            if subtitle:
                ax.text(x + w/2, y + h * 0.64, title,
                        fontfamily=SANS_FONT, fontsize=11.5, fontweight='bold',
                        color=TEXT_BLACK, ha='center', va='center', zorder=3)
                ax.text(x + w/2, y + h * 0.32, subtitle,
                        fontfamily=SANS_FONT, fontsize=9.2,
                        color=TEXT_BLACK, ha='center', va='center', zorder=3)
            else:
                ax.text(x + w/2, y + h/2, title,
                        fontfamily=SANS_FONT, fontsize=11.5, fontweight='bold',
                        color=TEXT_BLACK, ha='center', va='center', zorder=3)
        return box

    # Helper: Draw 3D Database Cylinder
    def draw_database(cx, cy, w, h, eh, title, subtitle):
        cy_top = cy + h/2 - eh/2
        cy_bot = cy - h/2 + eh/2
        
        # Bottom half-ellipse
        theta_bot = np.linspace(np.pi, 2*np.pi, 60)
        bot_x = cx + (w/2) * np.cos(theta_bot)
        bot_y = cy_bot + (eh/2) * np.sin(theta_bot)
        
        verts = [(cx - w/2, cy_top)]
        codes = [Path.MOVETO]
        for bx, by in zip(bot_x, bot_y):
            verts.append((bx, by))
            codes.append(Path.LINETO)
            
        verts.append((cx + w/2, cy_top))
        codes.append(Path.LINETO)
        
        # Top half-ellipse to close body
        theta_top = np.linspace(0, np.pi, 60)
        top_x = cx + (w/2) * np.cos(theta_top)
        top_y = cy_top + (eh/2) * np.sin(theta_top)
        for tx, ty in zip(top_x, top_y):
            verts.append((tx, ty))
            codes.append(Path.LINETO)
            
        verts.append((cx - w/2, cy_top))
        codes.append(Path.CLOSEPOLY)
        
        body_patch = PathPatch(Path(verts, codes), facecolor=FILL_GRAY,
                               edgecolor=BORDER_COLOR, linewidth=LINE_WIDTH, zorder=2)
        ax.add_patch(body_patch)
        
        # Top Rim Ellipse
        top_ellipse = Ellipse((cx, cy_top), w, eh, facecolor=FILL_GRAY,
                              edgecolor=BORDER_COLOR, linewidth=LINE_WIDTH, zorder=3)
        ax.add_patch(top_ellipse)
        
        # Text
        ax.text(cx, cy + 0.3, title,
                fontfamily=SANS_FONT, fontsize=11.5, fontweight='bold',
                color=TEXT_BLACK, ha='center', va='center', zorder=4)
        ax.text(cx, cy - 2.0, subtitle,
                fontfamily=SANS_FONT, fontsize=8.6,
                color=TEXT_BLACK, ha='center', va='center', zorder=4)

    # 2. Top Box: Frontend
    top_w = 56
    top_h = 6.2
    top_x = 50 - top_w/2  # 22
    top_y = 47.5
    draw_box(top_x, top_y, top_w, top_h, 
             title="Frontend (Web Interface)", 
             subtitle="User Interface for Google Meet & Configuration",
             is_split=False)
    
    # 3. Middle Box: Backend Server
    mid_w = 56
    mid_h = 6.2
    mid_x = 50 - mid_w/2  # 22
    mid_y = 36.8
    draw_box(mid_x, mid_y, mid_w, mid_h,
             title="Backend Server (Application Layer)",
             subtitle="Handles business logic and API requests",
             is_split=False)
    
    # Arrows between Frontend and Backend Server
    # Left arrow: Request (downwards)
    req_x = 48.0
    ax.annotate('', xy=(req_x, mid_y + mid_h), xytext=(req_x, top_y),
                arrowprops=dict(arrowstyle="-|>", color=BORDER_COLOR, 
                                lw=LINE_WIDTH, mutation_scale=14), zorder=6)
    ax.text(req_x - 1.2, (top_y + mid_y + mid_h)/2, "Request",
            fontfamily=SANS_FONT, fontsize=8.8, color=TEXT_BLACK, ha='right', va='center')
    
    # Right arrow: Response (upwards)
    resp_x = 52.0
    ax.annotate('', xy=(resp_x, top_y), xytext=(resp_x, mid_y + mid_h),
                arrowprops=dict(arrowstyle="-|>", color=BORDER_COLOR, 
                                lw=LINE_WIDTH, mutation_scale=14), zorder=6)
    ax.text(resp_x + 1.2, (top_y + mid_y + mid_h)/2, "Response",
            fontfamily=SANS_FONT, fontsize=8.8, color=TEXT_BLACK, ha='left', va='center')

    # 4. Four Module Boxes
    # Width and heights
    mod_w = 20.8
    mod_h = 13.8
    spacing = 2.4
    total_w = 4 * mod_w + 3 * spacing # 83.2 + 7.2 = 90.4
    start_x = 50 - total_w/2           # 4.8
    mod_y = 17.5
    
    modules_data = [
        {
            "title": "Audio Ingest Module",
            "bullets": [
                "VB-Cable audio loopback",
                "PyAudio 16 kHz mono stream",
                "4th-order Butterworth DSP",
                "3.0s rolling PCM buffer"
            ]
        },
        {
            "title": "Speech & Trigger Engine",
            "bullets": [
                "Offline VOSK speech engine",
                "Rolling context deque (50w)",
                "Levenshtein fuzzy matching",
                "Confidence threshold (≥ 0.70)"
            ]
        },
        {
            "title": "AI & Reasoning Module",
            "bullets": [
                "Intent routing (roll-call / QA)",
                "Bounded prompt constructor",
                "GPT-4o-mini response engine",
                "Whisper-1 full transcription"
            ]
        },
        {
            "title": "Automation & Export Module",
            "bullets": [
                "Headless Selenium Chrome",
                "DOM selector matching",
                "Automated chat typing",
                "Notion API notes sync"
            ]
        }
    ]
    
    mod_centers_x = []
    for idx, mod in enumerate(modules_data):
        mx = start_x + idx * (mod_w + spacing)
        cx = mx + mod_w/2
        mod_centers_x.append(cx)
        draw_box(mx, mod_y, mod_w, mod_h,
                 title=mod["title"],
                 bullet_points=mod["bullets"],
                 is_split=True)
        
    # Connect Backend Server to the 4 Modules
    # Line from backend bottom center down to bus
    bus_y = (mid_y + (mod_y + mod_h)) / 2 + 0.3 # around 33.8
    ax.plot([50, 50], [mid_y, bus_y], color=BORDER_COLOR, linewidth=LINE_WIDTH, zorder=5)
    
    # Horizontal bus line across 4 module centers
    ax.plot([mod_centers_x[0], mod_centers_x[3]], [bus_y, bus_y], 
            color=BORDER_COLOR, linewidth=LINE_WIDTH, zorder=5)
    
    # 4 drop arrows from bus to top of each module
    for cx in mod_centers_x:
        ax.annotate('', xy=(cx, mod_y + mod_h), xytext=(cx, bus_y),
                    arrowprops=dict(arrowstyle="-|>", color=BORDER_COLOR,
                                    lw=LINE_WIDTH, mutation_scale=14), zorder=6)

    # 5. Database Cylinder at Bottom
    db_w = 26.5
    db_h = 7.8
    db_eh = 2.0
    db_cx = 50.0
    db_cy = 6.2
    draw_database(db_cx, db_cy, db_w, db_h, db_eh,
                  title="Database",
                  subtitle="Stores audio, transcripts, notes and logs")
    
    # Connect 4 Modules to Database
    # Module 1 (outer left): drops down, turns right into database left wall
    m1_x = mod_centers_x[0]
    m1_bottom_y = mod_y
    turn_y = db_cy - 0.4
    db_left_x = db_cx - db_w/2
    ax.plot([m1_x, m1_x], [m1_bottom_y, turn_y], color=BORDER_COLOR, linewidth=LINE_WIDTH, zorder=5)
    ax.annotate('', xy=(db_left_x, turn_y), xytext=(m1_x, turn_y),
                arrowprops=dict(arrowstyle="-|>", color=BORDER_COLOR,
                                lw=LINE_WIDTH, mutation_scale=14), zorder=6)
    
    # Module 4 (outer right): drops down, turns left into database right wall
    m4_x = mod_centers_x[3]
    m4_bottom_y = mod_y
    db_right_x = db_cx + db_w/2
    ax.plot([m4_x, m4_x], [m4_bottom_y, turn_y], color=BORDER_COLOR, linewidth=LINE_WIDTH, zorder=5)
    ax.annotate('', xy=(db_right_x, turn_y), xytext=(m4_x, turn_y),
                arrowprops=dict(arrowstyle="-|>", color=BORDER_COLOR,
                                lw=LINE_WIDTH, mutation_scale=14), zorder=6)
    
    # Module 2 (inner left): drops straight down to database top ellipse
    m2_x = mod_centers_x[1]
    m2_bottom_y = mod_y
    # Top rim intersection
    db_top_rim_y = db_cy + db_h/2 - db_eh/2 + 0.3
    ax.annotate('', xy=(m2_x, db_top_rim_y + 0.5), xytext=(m2_x, m2_bottom_y),
                arrowprops=dict(arrowstyle="-|>", color=BORDER_COLOR,
                                lw=LINE_WIDTH, mutation_scale=14), zorder=6)

    # Module 3 (inner right): drops straight down to database top ellipse
    m3_x = mod_centers_x[2]
    m3_bottom_y = mod_y
    ax.annotate('', xy=(m3_x, db_top_rim_y + 0.5), xytext=(m3_x, m3_bottom_y),
                arrowprops=dict(arrowstyle="-|>", color=BORDER_COLOR,
                                lw=LINE_WIDTH, mutation_scale=14), zorder=6)

    plt.tight_layout()
    plt.savefig(output_path, dpi=300, facecolor='#FFFFFF', bbox_inches='tight')
    plt.close()
    print(f"Generated architecture diagram saved to: {output_path}")

if __name__ == "__main__":
    generate_system_architecture("system_architecture_diagram.png")
