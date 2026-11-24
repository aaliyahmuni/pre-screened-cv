import os
import re
import unicodedata
import fitz  # PyMuPDF
from docx import Document

PHOTO_DIR = os.path.join("saved_resumes", "photos")
CERT_IMG_DIR = os.path.join("saved_resumes", "cert_images")
os.makedirs(PHOTO_DIR, exist_ok=True)
os.makedirs(CERT_IMG_DIR, exist_ok=True)

JOB_LIST = [
    "Service Technician",
    "Signalman",
    "Lorry Driver",
    "Mason",
    "General Labor",
    "Welding Inspector (CSWIP)",
    "Welder",
    "Technical Executive - Material Coordinator",
    "Technical Executive - Scheduler",
    "Storekeeper",
    "Senior Welder",
    "Senior Pipe Fitter",
    "Senior Mechanical Fitter",
    "Scaffolding Inspectors",
    "Rigger",
    "Pipe Fitter",
    "Millwright",
    "Mechanical Fitter",
    "Machinist",
    "Lifting Supervisor",
    "HSE Supervisor",
    "Document Controller",
    "Chargeman B0/B4",
    "EHS Team Lead",
    "Logistic EHS Specialist",
    "Junior Accountant",
    "Senior Engineer",
    "Lead Drilling Advisor",
    "IT End User Support",
    "Technician",
    "Engineer (Specialist Telecom/OEM)",
    "Telecom Rigger / Abseilor",
    "Engineer",
]


def is_noise_or_garbage(text):
    """Filters out page numbers and garbled artifacts."""
    if not text:
        return True
    if re.match(r"^\s*page\s+\d+(?:\s+of\s+\d+)?\s*$", text, re.IGNORECASE):
        return True

    stripped = text.strip()
    if len(stripped) <= 1:
        return True

    alnum_count = sum(c.isalnum() for c in stripped)
    letter_words = re.findall(r"[A-Za-z]{2,}", stripped)

    if not letter_words and not re.match(r"^\+?\d[\d\s\-]{6,}$", stripped):
        return True

    if len(stripped) > 3 and (alnum_count / len(stripped)) < 0.65:
        return True

    if re.search(r"[~\\|]{2,}|(?:[,\.;:'\"`~•]{2,}\s*){2,}", stripped):
        return True

    return False


def normalize_exact_text(text):
    """Normalizes Unicode ligatures and removes PDF Wingdings/square bullet glyphs."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    text = re.sub(r"[\ue000-\uf8ff\u25a0-\u27bf]", "", text)
    return text.strip()


def strip_header_numbering(text):
    """Removes any numbering beside headers and subheaders while keeping exact wording."""
    cleaned = re.sub(
        r"^\s*(?:[\d]+(?:\.[\d]+)*[\.\)\-:]?|[A-Za-z][\.\)])\s+",
        "",
        text,
    )
    return cleaned.strip()


def extract_jobs_from_pdf():
    """Returns the standardized job position list for the dropdown."""
    return JOB_LIST


def render_pdf_pages_as_highres_images(pdf_path, prefix_id):
    """Renders PDF pages at crisp 300 DPI PNGs so certificates are sharp and readable."""
    rendered_paths = []
    try:
        doc = fitz.open(pdf_path)
        zoom_matrix = fitz.Matrix(300 / 72, 300 / 72)  # 300 DPI high-clarity render
        for idx, page in enumerate(doc):
            pix = page.get_pixmap(matrix=zoom_matrix, alpha=False)
            out_path = os.path.join(CERT_IMG_DIR, f"{prefix_id}_page{idx+1}.png")
            pix.save(out_path)
            rendered_paths.append(out_path)
    except Exception:
        pass
    return rendered_paths


def extract_images_from_pdf(pdf_path, candidate_id):
    """
    Extracts:
    1. Profile photo from page 1.
    2. High-resolution (300 DPI) page renders for attached certificate pages (without low-res thumbnail extraction).
    """
    profile_photo = None
    cert_images = []

    try:
        doc = fitz.open(pdf_path)
        zoom_matrix = fitz.Matrix(300 / 72, 300 / 72)  # 300 DPI for crisp certificates

        for page_index, page in enumerate(doc):
            images = page.get_images(full=True)
            page_text = page.get_text("text").strip()
            clean_words = re.findall(r"[A-Za-z]{3,}", page_text)

            # Page 1: Extract candidate headshot photo
            if page_index == 0:
                for img in images:
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    width, height = base_image["width"], base_image["height"]
                    if 80 <= width <= 1000 and 80 <= height <= 1200:
                        ext = base_image["ext"]
                        photo_path = os.path.join(PHOTO_DIR, f"{candidate_id}.{ext}")
                        with open(photo_path, "wb") as f:
                            f.write(base_image["image"])
                        profile_photo = photo_path
                        break

            # Subsequent pages that contain embedded certificate images/scans -> Render full page at 300 DPI
            if page_index >= 1 and len(images) > 0:
                # Avoid capturing plain text pages that only have a tiny header/footer logo
                has_large_image = False
                for img in images:
                    base_image = doc.extract_image(img[0])
                    if base_image["width"] >= 300 and base_image["height"] >= 300:
                        has_large_image = True
                        break

                if has_large_image or len(clean_words) < 35:
                    pix = page.get_pixmap(matrix=zoom_matrix, alpha=False)
                    cert_path = os.path.join(CERT_IMG_DIR, f"{candidate_id}_cert_p{page_index+1}.png")
                    pix.save(cert_path)
                    cert_images.append(cert_path)

    except Exception:
        pass

    return profile_photo, cert_images


def extract_images_from_docx(docx_path, candidate_id):
    """Extracts profile photo (first image) and original full-resolution certificate images from DOCX."""
    profile_photo = None
    cert_images = []
    try:
        doc = Document(docx_path)
        idx = 0
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                img_bytes = rel.target_part.blob
                ext = rel.target_ref.split(".")[-1]
                if idx == 0:
                    photo_path = os.path.join(PHOTO_DIR, f"{candidate_id}.{ext}")
                    with open(photo_path, "wb") as f:
                        f.write(img_bytes)
                    profile_photo = photo_path
                else:
                    cert_path = os.path.join(CERT_IMG_DIR, f"{candidate_id}_cert_{idx}.{ext}")
                    with open(cert_path, "wb") as f:
                        f.write(img_bytes)
                    cert_images.append(cert_path)
                idx += 1
    except Exception:
        pass
    return profile_photo, cert_images


def extract_verbatim_lines_from_pdf(pdf_path):
    """Extracts clean, readable lines in strict visual reading order."""
    doc = fitz.open(pdf_path)
    all_lines = []
    for page in doc:
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            if b[6] == 0:
                block_text = normalize_exact_text(b[4])
                for line in block_text.split("\n"):
                    clean = line.strip()
                    if clean and not is_noise_or_garbage(clean):
                        all_lines.append(clean)
    return all_lines


def extract_verbatim_lines_from_docx(docx_path):
    """Extracts clean lines from paragraphs and tables in a DOCX resume."""
    doc = Document(docx_path)
    all_lines = []
    for p in doc.paragraphs:
        clean = normalize_exact_text(p.text)
        if clean and not is_noise_or_garbage(clean):
            all_lines.append(clean)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    clean = normalize_exact_text(p.text)
                    if clean and not is_noise_or_garbage(clean) and clean not in all_lines:
                        all_lines.append(clean)
    return all_lines


def parse_resume_file(file_path, candidate_id, default_filename):
    """Extracts exact verbatim text and high-clarity images into standardized resume sections."""
    ext = os.path.splitext(file_path)[1].lower()
    lines = []
    photo_path = None
    cert_images = []

    if ext == ".pdf":
        photo_path, cert_images = extract_images_from_pdf(file_path, candidate_id)
        lines = extract_verbatim_lines_from_pdf(file_path)
    elif ext == ".docx":
        photo_path, cert_images = extract_images_from_docx(file_path, candidate_id)
        lines = extract_verbatim_lines_from_docx(file_path)

    raw_text = "\n".join(lines)

    email_match = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", raw_text)
    phone_match = re.search(r"(\+?\d[\d\s\-\(\)]{7,16}\d)", raw_text)
    gender_match = re.search(r"\b(Male|Female)\b", raw_text, re.IGNORECASE)
    age_match = re.search(
        r"\b(?:Age\s*[:\-]?\s*(\d{2})|(\d{2})\s*years\s*old)\b", raw_text, re.IGNORECASE
    )
    marital_match = re.search(
        r"\b(Single|Married|Divorced|Widowed)\b", raw_text, re.IGNORECASE
    )
    nationality_match = re.search(
        r"Nationality\s*[:\-]?\s*([A-Za-z\s]+)", raw_text, re.IGNORECASE
    )
    address_match = re.search(
        r"(?:Address|Location)\s*[:\-]?\s*(.+)", raw_text, re.IGNORECASE
    )

    candidate_name = (
        strip_header_numbering(lines[0])
        if lines
        else os.path.splitext(default_filename)[0]
    )

    section_headers = {
        "objective": [
            "objective",
            "career objective",
            "professional summary",
            "summary",
            "profile",
            "about me",
        ],
        "education": [
            "education",
            "academic",
            "academic qualification",
            "qualification",
            "educational background",
        ],
        "courses": [
            "course",
            "training",
            "courses & training",
            "course & training attended",
            "training attended",
            "workshops",
        ],
        "experience": [
            "experience",
            "working experience",
            "work experience",
            "employment",
            "employment history",
            "career history",
        ],
        "skills": [
            "skills",
            "capabilities",
            "skills & capabilities",
            "key skills",
            "competencies",
            "technical skills",
        ],
        "certificates": [
            "certificate",
            "certificates",
            "certification",
            "certifications",
            "licenses",
            "licenses & certifications",
        ],
        "ignore": [
            "reference",
            "references",
            "character reference",
            "character references",
            "referee",
            "referees",
            "declaration",
        ],
    }

    sections = {k: [] for k in section_headers}
    current_section = None

    for line in lines[1:]:
        cleaned = strip_header_numbering(line)
        lower_clean = cleaned.lower().rstrip(":").strip()

        matched_key = None
        if len(cleaned) < 50:
            for key, aliases in section_headers.items():
                if lower_clean in aliases or any(
                    lower_clean.startswith(a) and len(lower_clean) < 38 for a in aliases
                ):
                    matched_key = key
                    break

        if matched_key:
            current_section = matched_key
        elif current_section and current_section != "ignore":
            sections[current_section].append(cleaned)

    return {
        "id": candidate_id,
        "filename": default_filename,
        "file_path": file_path,
        "photo_path": photo_path,
        "cert_images": cert_images,
        "position": "",
        "name": candidate_name,
        "phone": phone_match.group(1).strip() if phone_match else "Not Provided",
        "email": email_match.group(0).strip() if email_match else "Not Provided",
        "gender": gender_match.group(1).capitalize() if gender_match else "Not Provided",
        "age": (age_match.group(1) or age_match.group(2)) if age_match else "Not Provided",
        "nationality": (
            nationality_match.group(1).strip().split("\n")[0]
            if nationality_match
            else "Not Provided"
        ),
        "marital_status": (
            marital_match.group(1).capitalize() if marital_match else "Not Provided"
        ),
        "address": address_match.group(1).strip() if address_match else "Not Provided",
        "objective": "\n".join(sections["objective"]) or "Not Provided",
        "education": "\n".join(sections["education"]) or "Not Provided",
        "courses": "\n".join(sections["courses"]) or "None listed.",
        "experience": "\n".join(sections["experience"]) or "",
        "skills": "\n".join(sections["skills"]) or "None listed.",
        "certificates": "\n".join(sections["certificates"]) or "",
    }