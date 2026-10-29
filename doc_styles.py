import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_margins(cell, top=40, bottom=40, left=50, right=50):
    """Set inner padding for a table cell (in dxa: 20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color="E8EEF5"):
    """Set background color of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def set_table_borders(table, color="7F7F7F", sz="4", val="single"):
    """Set clean, visible professional grid borders on a table (0.5 pt solid) compliant with OpenXML schema."""
    tblPr = table._tbl.tblPr
    existing_borders = tblPr.find(qn('w:tblBorders'))
    if existing_borders is not None:
        tblPr.remove(existing_borders)
        
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)  # 4 = 0.5 pt
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tblBorders.append(border)
        
    tblLook = tblPr.find(qn('w:tblLook'))
    if tblLook is not None:
        tblLook.addprevious(tblBorders)
    else:
        tblPr.append(tblBorders)

def format_table(table, header_bg="E8EEF5", font_size=8.0, header_font_size=8.5):
    """Format an entire table: schema-valid borders, cell padding, left alignment, and header shading."""
    set_table_borders(table, color="7F7F7F", sz="4")
    for r_idx, row in enumerate(table.rows):
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
            if is_header:
                set_cell_shading(cell, header_bg)
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_header else WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(header_font_size if is_header else font_size)
                    r.font.bold = is_header or (c_idx == 0 and not is_header)

def add_algorithm_box(doc, algo_id, title, input_desc, output_desc, steps):
    """Insert an academic IEEE-style bordered algorithm box with clean left-aligned text."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(5.77)
    
    cell = tbl.cell(0, 0)
    set_cell_margins(cell, top=30, bottom=30, left=50, right=50)
    set_cell_shading(cell, "F9FAFB")
    set_table_borders(tbl, color="888888", sz="6") # 0.75 pt border
    
    # Header paragraph
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(1)
    p0.paragraph_format.space_after = Pt(2)
    p0.paragraph_format.line_spacing = 1.15
    p0.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r0 = p0.add_run(f"Algorithm {algo_id}: {title}")
    r0.font.name = 'Times New Roman'
    r0.font.size = Pt(9.5)
    r0.font.bold = True
    
    # Input/Output
    p_io = cell.add_paragraph()
    p_io.paragraph_format.space_before = Pt(1)
    p_io.paragraph_format.space_after = Pt(2)
    p_io.paragraph_format.line_spacing = 1.15
    p_io.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_in = p_io.add_run("Input: ")
    r_in.font.bold = True
    r_in_txt = p_io.add_run(f"{input_desc}\n")
    r_out = p_io.add_run("Output: ")
    r_out.font.bold = True
    r_out_txt = p_io.add_run(f"{output_desc}")
    for r in [r_in, r_in_txt, r_out, r_out_txt]:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(8.5)
        
    # Steps
    for step in steps:
        p_s = cell.add_paragraph()
        p_s.paragraph_format.space_before = Pt(0)
        p_s.paragraph_format.space_after = Pt(1)
        p_s.paragraph_format.line_spacing = 1.15
        p_s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_s = p_s.add_run(step)
        r_s.font.name = 'Courier New' if ('=' in step or 'def ' in step or '(' in step) else 'Times New Roman'
        r_s.font.size = Pt(8.0)
        
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(3)


def setup_document():
    """Create and configure base document with Times New Roman and official margins."""
    doc = docx.Document()
    
    # Configure base section margins: Left = 1.5 in (108 pt), Top = 1.0 in (72 pt), Right = 1.0 in, Bottom = 1.0 in
    sec = doc.sections[0]
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.5)
    sec.right_margin = Inches(1.0)
    sec.page_width = Inches(8.27)   # A4
    sec.page_height = Inches(11.69) # A4
    
    # Configure Normal Style
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    return doc

def add_body_p(doc, text, bold_prefix="", italic=False):
    """Add standard justified body paragraph with 1.5 line spacing."""
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Times New Roman'
        r_pre.font.size = Pt(12)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(0, 0, 0)
        
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.italic = italic
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_bullet_p(doc, text, bold_prefix=""):
    """Add a bullet point paragraph formatted with Times New Roman 12pt, 1.5 line spacing."""
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Times New Roman'
        r_pre.font.size = Pt(12)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(0, 0, 0)
        
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_chapter_heading(doc, text):
    """Add Chapter Title: 18 pt, Bold, Centered."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_section_heading(doc, text):
    """Add Section Title: 14 pt, Bold, Left aligned."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.keep_with_next = True
    
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_subsection_heading(doc, text):
    """Add Subsection Title: 12 pt, Bold, Left aligned."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.keep_with_next = True
    
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_figure_caption(doc, text):
    """Add Figure caption: 10 pt, Centered, below figure."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_table_title(doc, text):
    """Add Table title: 10 pt, Bold, Centered, above table."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_image_centered(doc, image_path, width_in_inches=3.6):
    """Insert centered image into document."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    r.add_picture(image_path, width=Inches(width_in_inches))
    return p
