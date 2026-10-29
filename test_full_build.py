import os
import time
import docx
from docx.shared import Inches
from docx.enum.section import WD_SECTION_START
import docx2pdf

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

def main():
    print("Building full docx...", flush=True)
    doc = setup_document()
    sec1 = doc.sections[0]
    sec1.top_margin = Inches(1.0)
    sec1.bottom_margin = Inches(1.0)
    sec1.left_margin = Inches(1.5)
    sec1.right_margin = Inches(1.0)

    create_cover_page(doc, logo_path="nmims_logo.png")
    create_certificate_page(doc, logo_path="nmims_logo.png")

    sec2 = doc.add_section(WD_SECTION_START.NEW_PAGE)
    sec2.top_margin = Inches(1.0)
    sec2.bottom_margin = Inches(1.0)
    sec2.left_margin = Inches(1.5)
    sec2.right_margin = Inches(1.0)

    create_acknowledgement_page(doc)
    create_abstract_page(doc)
    create_table_of_contents_page(doc)
    create_list_of_figures_page(doc)
    create_list_of_tables_page(doc)
    create_abbreviations_page(doc)

    sec3 = doc.add_section(WD_SECTION_START.NEW_PAGE)
    sec3.top_margin = Inches(1.0)
    sec3.bottom_margin = Inches(1.0)
    sec3.left_margin = Inches(1.5)
    sec3.right_margin = Inches(1.0)

    build_chapter_1(doc)
    build_chapter_2(doc)
    build_chapter_3(doc)
    build_chapter_4(doc)
    build_chapter_5(doc)
    build_chapter_6(doc)
    build_references(doc)
    build_appendices(doc)

    out_docx = "AI_Student_Agent_Black_Book_Final.docx"
    doc.save(out_docx)
    print(f"Saved {out_docx} ({os.path.getsize(out_docx):,} bytes)", flush=True)

    print("Converting with docx2pdf...", flush=True)
    t0 = time.time()
    out_pdf = "AI_Student_Agent_Black_Book_Final.pdf"
    docx2pdf.convert(out_docx, out_pdf)
    elapsed = time.time() - t0
    print(f"Exported {out_pdf} in {elapsed:.2f}s ({os.path.getsize(out_pdf):,} bytes)", flush=True)

if __name__ == "__main__":
    main()
