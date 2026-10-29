import os
import sys
import time
import docx
from docx.shared import Inches, Pt
from docx.enum.section import WD_SECTION_START

from doc_styles import setup_document
from front_matter import (
    create_cover_page, create_certificate_page,
    create_acknowledgement_page, create_abstract_page,
    create_table_of_contents_page, create_list_of_figures_page,
    create_list_of_tables_page, create_abbreviations_page
)
from ch1_intro import build_chapter_1
from ch2_lit_survey import build_chapter_2
from ch3_methodology import build_chapter_3
from ch4_results import build_chapter_4
from ch5_adv_lim_app import build_chapter_5
from ch6_conclusion import build_chapter_6
from back_matter import build_references, build_appendices

def generate_black_book(docx_output="AI_Student_Agent_Black_Book_Final.docx"):
    """Assemble all preliminary pages, chapters, references, and appendices into a complete Black Book."""
    print("=" * 65, flush=True)
    print(" Compiling AI Student Agent Capstone Black Book...", flush=True)
    print("=" * 65, flush=True)

    # 1. Initialize Document (Section 1: Cover Page and Certificate)
    doc = setup_document()

    print(" -> Creating Section 1: Cover Page & Certificate...", flush=True)
    create_cover_page(doc, logo_path="nmims_logo.png")
    create_certificate_page(doc, logo_path="nmims_logo.png")

    # 2. Section 2: Preliminary Pages (Acknowledgement, Abstract, TOC, LOF, LOT, Abbreviations)
    print(" -> Creating Section 2: Preliminary Pages (Lowercase Roman: i, ii, iii...)...", flush=True)
    sec2 = doc.add_section(WD_SECTION_START.NEW_PAGE)

    create_acknowledgement_page(doc)
    create_abstract_page(doc)
    create_table_of_contents_page(doc)
    create_list_of_figures_page(doc)
    create_list_of_tables_page(doc)
    create_abbreviations_page(doc)

    # 3. Section 3: Main Chapters, References, and Appendices (Arabic Decimal Numbers starting at 1)
    print(" -> Creating Section 3: Main Chapters 1 to 6, References, Appendices (Arabic: 1, 2, 3...)...", flush=True)
    sec3 = doc.add_section(WD_SECTION_START.NEW_PAGE)

    print("    * Chapter 1: Introduction", flush=True)
    build_chapter_1(doc)

    print("    * Chapter 2: Literature Survey", flush=True)
    build_chapter_2(doc)

    print("    * Chapter 3: Methodology and Implementation", flush=True)
    build_chapter_3(doc)

    print("    * Chapter 4: Results and Analysis", flush=True)
    build_chapter_4(doc)

    print("    * Chapter 5: Advantages, Limitations and Applications", flush=True)
    build_chapter_5(doc)

    print("    * Chapter 6: Conclusion and Future Scope", flush=True)
    build_chapter_6(doc)

    print("    * References (IEEE Format)", flush=True)
    build_references(doc)

    print("    * Appendices I to IV", flush=True)
    build_appendices(doc)

    # Save DOCX
    t0 = time.time()
    doc.save(docx_output)
    print(f"\n [SUCCESS] Successfully compiled Black Book DOCX: '{docx_output}' in {time.time()-t0:.2f}s", flush=True)
    print(f" DOCX File Size: {os.path.getsize(docx_output):,} bytes", flush=True)
    print("=" * 65, flush=True)

if __name__ == "__main__":
    generate_black_book()
