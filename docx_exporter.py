import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxcompose.composer import Composer
from parser_engine import strip_header_numbering

TEMPLATE_DOCX = "etc.docx"


def remove_table_borders(table):
    """Makes the top photo + personal details layout table borderless."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "none")
        borders.append(border)
    tblPr.append(borders)


def add_section_block(doc, title, content):
    """Adds a clean unnumbered section header and detailed body text."""
    clean_title = strip_header_numbering(title).upper()
    p_head = doc.add_paragraph()
    p_head.paragraph_format.space_before = Pt(10)
    p_head.paragraph_format.space_after = Pt(3)
    run_head = p_head.add_run(clean_title)
    run_head.bold = True
    run_head.font.size = Pt(11)
    run_head.font.color.rgb = RGBColor(20, 55, 100)

    pPr = p_head._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "143764")
    pBdr.append(bottom)
    pPr.append(pBdr)

    for line in content.split("\n"):
        clean_line = strip_header_numbering(line)
        if clean_line:
            p_body = doc.add_paragraph(clean_line)
            p_body.paragraph_format.space_after = Pt(2)
            p_body.paragraph_format.line_spacing = 1.15
            for r in p_body.runs:
                r.font.size = Pt(10)


def set_section_footer(section, position_text):
    """Inserts the applied job position into the Word document footer."""
    section.different_first_page_header_footer = False
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.text = f"Position Applied For: {position_text}"
    for run in p.runs:
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(80, 80, 80)


def build_single_candidate_doc(candidate):
    """Creates a single candidate document based on etc.docx template."""
    doc = Document(TEMPLATE_DOCX) if os.path.exists(TEMPLATE_DOCX) else Document()

    set_section_footer(doc.sections[-1], candidate.get("position", "Not Specified"))

    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    remove_table_borders(table)

    cell_photo = table.cell(0, 0)
    cell_details = table.cell(0, 1)
    cell_photo.width = Inches(1.7)
    cell_details.width = Inches(4.8)

    p_photo = cell_photo.paragraphs[0]
    photo_path = candidate.get("photo_path")
    if photo_path and os.path.exists(photo_path):
        run_img = p_photo.add_run()
        run_img.add_picture(photo_path, width=Inches(1.45))
    else:
        run_no_img = p_photo.add_run("[Candidate Photo]")
        run_no_img.italic = True
        run_no_img.font.size = Pt(9)

    p_name = cell_details.paragraphs[0]
    p_name.paragraph_format.space_after = Pt(4)
    r_name = p_name.add_run(strip_header_numbering(candidate["name"]).upper())
    r_name.bold = True
    r_name.font.size = Pt(14)
    r_name.font.color.rgb = RGBColor(20, 55, 100)

    details_fields = [
        ("Phone Number", candidate["phone"]),
        ("Email", candidate["email"]),
        ("Gender", candidate["gender"]),
        ("Age", candidate["age"]),
        ("Nationality", candidate["nationality"]),
        ("Marital Status", candidate["marital_status"]),
        ("Address", candidate["address"]),
    ]

    for label, val in details_fields:
        p = cell_details.add_paragraph()
        p.paragraph_format.space_after = Pt(1.5)
        r_lbl = p.add_run(f"{label}: ")
        r_lbl.bold = True
        r_lbl.font.size = Pt(9.5)
        r_val = p.add_run(str(val))
        r_val.font.size = Pt(9.5)

    add_section_block(doc, "Objective", candidate["objective"])
    add_section_block(doc, "Education", candidate["education"])
    add_section_block(doc, "Course & Training Attended", candidate["courses"])
    add_section_block(doc, "Working Experience", candidate["experience"])
    add_section_block(doc, "Skills & Capabilities", candidate["skills"])
    add_section_block(doc, "Certificates", candidate["certificates"])

    return doc


def export_all_to_master_docx(candidates, output_path="Compiled_PreScreened_Resumes.docx"):
    """Merges all standardized resumes into one master Word file using etc.docx."""
    if not candidates:
        return None

    master_doc = build_single_candidate_doc(candidates[0])
    composer = Composer(master_doc)

    for cand in candidates[1:]:
        master_doc.add_section()
        sub_doc = build_single_candidate_doc(cand)
        composer.append(sub_doc)

    composer.save(output_path)
    return output_path