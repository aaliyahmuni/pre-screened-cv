import os
import re
import io
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxcompose.composer import Composer
from parser_engine import (
    strip_header_numbering,
    strip_all_bullet_shapes,
    is_noise_or_garbage,
    normalize_exact_text,
    ALL_BULLET_CHARS_PATTERN,
)

TEMPLATE_DOCX = "etc.docx"
PRIMARY_COLOR = RGBColor(20, 55, 100)     # Executive Navy
DARK_TEXT = RGBColor(30, 30, 30)          # Crisp Charcoal
MUTED_TEXT = RGBColor(90, 90, 90)         # Subtle Gray


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    """Sets custom borders on an individual Word table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side, cfg in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        if cfg:
            el.set(qn("w:val"), cfg.get("val", "single"))
            el.set(qn("w:sz"), str(cfg.get("sz", 4)))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), cfg.get("color", "CCCCCC"))
        else:
            el.set(qn("w:val"), "none")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    """Sets clean internal padding inside table cells (in dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def remove_table_borders(table):
    """Removes outer grid borders from layout tables."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "none")
        borders.append(border)
    tblPr.append(borders)


def is_subheader_line(line, section_type):
    """Detects job titles, company names, date ranges, or qualifications so they render in bold."""
    stripped = strip_all_bullet_shapes(line)
    if not stripped or len(stripped) > 110 or stripped.endswith("."):
        return False

    if re.search(r"\b(19|20)\d{2}\b", stripped) and len(stripped) < 95:
        return True

    if stripped.endswith(":"):
        return True

    if section_type in ("experience", "education", "courses"):
        role_or_org_words = (
            "sdn bhd", "bhd", "ltd", "inc", "corp", "company", "engineering",
            "services", "offshore", "oil", "gas", "petronas", "shell", "sarawak",
            "university", "universiti", "college", "institute", "politeknik",
            "diploma", "degree", "bachelor", "master", "certificate", "spm", "stpm",
            "engineer", "supervisor", "technician", "welder", "fitter", "inspector",
            "manager", "officer", "coordinator", "lead", "rigger", "operator",
            "specialist", "advisor", "accountant", "controller", "storekeeper",
            "client", "project", "contractor"
        )
        words = stripped.split()
        if len(words) <= 11 and any(w in stripped.lower() for w in role_or_org_words):
            return True

    return False


def add_section_header(doc, title):
    """Adds a uniform, unnumbered bold navy section header with a clean bottom rule."""
    clean_title = strip_header_numbering(title).upper()
    p_head = doc.add_paragraph()
    p_head.paragraph_format.space_before = Pt(11)
    p_head.paragraph_format.space_after = Pt(4)
    p_head.paragraph_format.keep_with_next = True

    run_head = p_head.add_run(clean_title)
    run_head.bold = True
    run_head.font.name = "Calibri"
    run_head.font.size = Pt(11)
    run_head.font.color.rgb = PRIMARY_COLOR

    pPr = p_head._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), "143764")
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_section_block(doc, title, content, section_type="general", cert_images=None):
    """Renders every standardized resume section uniformly for boss review."""
    add_section_header(doc, title)

    lines = [
        normalize_exact_text(l).strip()
        for l in (content or "").split("\n")
        if l.strip() and not is_noise_or_garbage(l)
    ]

    valid_cert_streams = []
    if cert_images:
        for img_path in cert_images:
            stream = get_compatible_image_stream(img_path)
            if stream:
                valid_cert_streams.append(stream)

    if not lines and not valid_cert_streams:
        p_empty = doc.add_paragraph()
        p_empty.paragraph_format.space_after = Pt(3)
        r_emp = p_empty.add_run("None listed.")
        r_emp.italic = True
        r_emp.font.name = "Calibri"
        r_emp.font.size = Pt(10)
        r_emp.font.color.rgb = MUTED_TEXT
        return

    for line in lines:
        clean_line = strip_header_numbering(line)
        has_existing_bullet = bool(re.match(ALL_BULLET_CHARS_PATTERN, clean_line))
        text_body = strip_all_bullet_shapes(clean_line)

        if not text_body or is_noise_or_garbage(text_body):
            continue

        # Objective paragraph formatting
        if section_type == "objective" and not has_existing_bullet:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(text_body)
            r.font.name = "Calibri"
            r.font.size = Pt(10)
            r.font.color.rgb = DARK_TEXT
            continue

        # Bold Subheaders (Companies, Job Titles, Dates, Degrees)
        if is_subheader_line(text_body, section_type) and not has_existing_bullet:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(5)
            p_sub.paragraph_format.space_after = Pt(1.5)
            p_sub.paragraph_format.keep_with_next = True
            r_sub = p_sub.add_run(text_body)
            r_sub.bold = True
            r_sub.font.name = "Calibri"
            r_sub.font.size = Pt(10)
            r_sub.font.color.rgb = DARK_TEXT
        else:
            # Standardized • Bullet Points for all detailed information
            p_bul = doc.add_paragraph()
            p_bul.paragraph_format.left_indent = Inches(0.25)
            p_bul.paragraph_format.first_line_indent = Inches(-0.15)
            p_bul.paragraph_format.space_after = Pt(2)
            p_bul.paragraph_format.line_spacing = 1.15

            r_dot = p_bul.add_run("•  ")
            r_dot.bold = True
            r_dot.font.name = "Calibri"
            r_dot.font.size = Pt(10)
            r_dot.font.color.rgb = PRIMARY_COLOR

            r_txt = p_bul.add_run(text_body)
            r_txt.font.name = "Calibri"
            r_txt.font.size = Pt(10)
            r_txt.font.color.rgb = DARK_TEXT

    # Insert crisp certificate attachments if provided
    for stream in valid_cert_streams:
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(6)
        run_img = p_img.add_run()
        run_img.add_picture(stream, width=Inches(5.8))


def set_section_footer(section, position_text):
    """Inserts the applied job position into the Word document footer."""
    section.different_first_page_header_footer = False
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.text = ""
    run = p.add_run(f"Position Applied For: {position_text}")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.font.color.rgb = PRIMARY_COLOR


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
    """Builds a 100% standardized executive resume inside etc.docx."""
    doc = Document(TEMPLATE_DOCX) if os.path.exists(TEMPLATE_DOCX) else Document()

    set_section_footer(doc.sections[-1], candidate.get("position", "Not Specified"))

    # =========================================================================
    # TOP HEADER BLOCK: Photo on Top-Left Corner | Personal Details on the Side
    # =========================================================================
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    remove_table_borders(header_table)

    cell_photo = header_table.cell(0, 0)
    cell_details = header_table.cell(0, 1)
    cell_photo.width = Inches(1.65)
    cell_details.width = Inches(4.85)
    cell_photo.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    cell_details.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # Clean subtle frame around the top-left photo box
    photo_border = {"val": "single", "sz": 6, "color": "143764"}
    set_cell_borders(cell_photo, top=photo_border, bottom=photo_border, left=photo_border, right=photo_border)
    set_cell_margins(cell_photo, top=60, bottom=60, left=60, right=60)
    set_cell_margins(cell_details, top=20, bottom=20, left=180, right=40)

    # Top-Left: Candidate Photo
    p_photo = cell_photo.paragraphs[0]
    p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_photo.paragraph_format.space_after = Pt(0)
    img_stream = get_compatible_image_stream(candidate.get("photo_path"))
    if img_stream:
        run_img = p_photo.add_run()
        run_img.add_picture(img_stream, width=Inches(1.45))
    else:
        p_photo.paragraph_format.space_before = Pt(35)
        p_photo.paragraph_format.space_after = Pt(35)
        run_no_img = p_photo.add_run("PHOTO\nNOT PROVIDED")
        run_no_img.bold = True
        run_no_img.font.name = "Calibri"
        run_no_img.font.size = Pt(8.5)
        run_no_img.font.color.rgb = MUTED_TEXT

    # Top-Right (Side of Photo): Candidate Full Name + Standardized Personal Details
    p_name = cell_details.paragraphs[0]
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(6)
    clean_name = strip_all_bullet_shapes(strip_header_numbering(candidate["name"])).upper()
    r_name = p_name.add_run(clean_name)
    r_name.bold = True
    r_name.font.name = "Calibri"
    r_name.font.size = Pt(15)
    r_name.font.color.rgb = PRIMARY_COLOR

    # Nested clean 2-column personal details grid beside the photo
    details_table = cell_details.add_table(rows=7, cols=2)
    details_table.autofit = False
    remove_table_borders(details_table)

    details_fields = [
        ("Phone Number", candidate.get("phone", "Not Provided")),
        ("Email", candidate.get("email", "Not Provided")),
        ("Gender", candidate.get("gender", "Not Provided")),
        ("Age", candidate.get("age", "Not Provided")),
        ("Nationality", candidate.get("nationality", "Not Provided")),
        ("Marital Status", candidate.get("marital_status", "Not Provided")),
        ("Address", candidate.get("address", "Not Provided")),
    ]

    for row_idx, (label, val) in enumerate(details_fields):
        row = details_table.rows[row_idx]
        c_lbl, c_val = row.cells[0], row.cells[1]
        c_lbl.width = Inches(1.25)
        c_val.width = Inches(3.45)
        set_cell_margins(c_lbl, top=15, bottom=15, left=0, right=40)
        set_cell_margins(c_val, top=15, bottom=15, left=20, right=0)

        p_l = c_lbl.paragraphs[0]
        p_l.paragraph_format.space_after = Pt(0)
        r_l = p_l.add_run(f"{label}:")
        r_l.bold = True
        r_l.font.name = "Calibri"
        r_l.font.size = Pt(9.5)
        r_l.font.color.rgb = DARK_TEXT

        p_v = c_val.paragraphs[0]
        p_v.paragraph_format.space_after = Pt(0)
        r_v = p_v.add_run(str(val or "Not Provided"))
        r_v.font.name = "Calibri"
        r_v.font.size = Pt(9.5)
        r_v.font.color.rgb = DARK_TEXT

    # =========================================================================
    # STANDARDIZED BODY SECTIONS (Exact Order for Every Candidate)
    # =========================================================================
    add_section_block(doc, "Objective", candidate.get("objective", ""), section_type="objective")
    add_section_block(doc, "Education", candidate.get("education", ""), section_type="education")
    add_section_block(doc, "Course & Training Attended", candidate.get("courses", ""), section_type="courses")
    add_section_block(doc, "Working Experience", candidate.get("experience", ""), section_type="experience")
    add_section_block(doc, "Skills & Capabilities", candidate.get("skills", ""), section_type="skills")
    add_section_block(
        doc,
        "Certificates",
        candidate.get("certificates", ""),
        section_type="certificates",
        cert_images=candidate.get("cert_images", []),
    )

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