import os
import re
import io
from PIL import Image
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


def is_subheader_line(line, section_type):
    """Detects job titles, company names, date ranges, or degrees so they are bolded like a professional CV."""
    stripped = line.lstrip("•-*▪➢➤ ").strip()
    if not stripped or len(stripped) > 110:
        return False

    # Never treat full sentences ending with a period as a subheader
    if stripped.endswith("."):
        return False

    # Lines containing date ranges (e.g., 2021 - 2024, Jan 2020 - Present)
    if re.search(r"\b(19|20)\d{2}\b", stripped) and len(stripped) < 95:
        return True

    # Explicit subheader endings (e.g., "Senior Welder:", "ABC Engineering Sdn Bhd")
    if stripped.endswith(":"):
        return True

    if section_type in ("experience", "education"):
        role_or_org_words = (
            "sdn bhd", "bhd", "ltd", "inc", "corp", "company", "engineering",
            "services", "offshore", "oil", "gas", "petronas", "shell",
            "university", "universiti", "college", "institute", "politeknik",
            "diploma", "degree", "bachelor", "master", "certificate", "spm", "stpm",
            "engineer", "supervisor", "technician", "welder", "fitter", "inspector",
            "manager", "officer", "coordinator", "lead", "rigger", "operator",
            "specialist", "advisor", "accountant", "controller", "storekeeper"
        )
        words = stripped.split()
        if len(words) <= 10 and any(w in stripped.lower() for w in role_or_org_words):
            return True

    return False


def add_section_block(doc, title, content, section_type="general"):
    """Adds a bold unnumbered header, bold subheaders, and bulleted professional details."""
    clean_title = strip_header_numbering(title).upper()
    p_head = doc.add_paragraph()
    p_head.paragraph_format.space_before = Pt(11)
    p_head.paragraph_format.space_after = Pt(4)
    p_head.paragraph_format.keep_with_next = True
    run_head = p_head.add_run(clean_title)
    run_head.bold = True
    run_head.font.size = Pt(11)
    run_head.font.color.rgb = RGBColor(20, 55, 100)

    # Clean bottom divider line under main section header
    pPr = p_head._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "143764")
    pBdr.append(bottom)
    pPr.append(pBdr)

    lines = [l.strip() for l in content.split("\n") if l.strip()]
    if not lines:
        return

    for line in lines:
        # Remove leading numbering beside subheaders/headers while preserving exact spelling
        clean_line = strip_header_numbering(line)
        has_existing_bullet = bool(re.match(r"^[•\-\*▪➢➤]\s*", clean_line))
        text_body = re.sub(r"^[•\-\*▪➢➤]\s*", "", clean_line).strip()

        if not text_body:
            continue

        # Objective is rendered as clean justified/left-aligned prose (unless explicitly bulleted)
        if section_type == "objective" and not has_existing_bullet:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(text_body)
            r.font.size = Pt(10)
            continue

        # Subheaders (Job Title, Company, Dates, Education Qualification) -> Bold, no bullet, no numbering
        if is_subheader_line(text_body, section_type) and not has_existing_bullet:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(5)
            p_sub.paragraph_format.space_after = Pt(1.5)
            p_sub.paragraph_format.keep_with_next = True
            r_sub = p_sub.add_run(text_body)
            r_sub.bold = True
            r_sub.font.size = Pt(10)
            r_sub.font.color.rgb = RGBColor(30, 30, 30)
        else:
            # Relevant details, duties, courses, skills, certificates -> Clean professional bullet points
            p_bul = doc.add_paragraph()
            p_bul.paragraph_format.left_indent = Inches(0.25)
            p_bul.paragraph_format.first_line_indent = Inches(-0.15)
            p_bul.paragraph_format.space_after = Pt(2)
            p_bul.paragraph_format.line_spacing = 1.15

            r_dot = p_bul.add_run("•  ")
            r_dot.bold = True
            r_dot.font.size = Pt(10)

            r_txt = p_bul.add_run(text_body)
            r_txt.font.size = Pt(10)


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


def get_compatible_image_stream(photo_path):
    """Converts any image into a python-docx compatible PNG stream."""
    if not photo_path or not os.path.exists(photo_path):
        return None
    try:
        with Image.open(photo_path) as img:
            img = img.convert("RGB")
            img_stream = io.BytesIO()
            img.save(img_stream, format="PNG")
            img_stream.seek(0)
            return img_stream
    except Exception:
        return None


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

    # Top-Left Photo
    p_photo = cell_photo.paragraphs[0]
    img_stream = get_compatible_image_stream(candidate.get("photo_path"))
    if img_stream:
        run_img = p_photo.add_run()
        run_img.add_picture(img_stream, width=Inches(1.45))
    else:
        run_no_img = p_photo.add_run("[Candidate Photo]")
        run_no_img.italic = True
        run_no_img.font.size = Pt(9)

    # Top-Right Standardized Personal Details
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

    # Standardized Resume Sections
    add_section_block(doc, "Objective", candidate["objective"], section_type="objective")
    add_section_block(doc, "Education", candidate["education"], section_type="education")
    add_section_block(doc, "Course & Training Attended", candidate["courses"], section_type="courses")
    add_section_block(doc, "Working Experience", candidate["experience"], section_type="experience")
    add_section_block(doc, "Skills & Capabilities", candidate["skills"], section_type="skills")
    add_section_block(doc, "Certificates", candidate["certificates"], section_type="certificates")

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