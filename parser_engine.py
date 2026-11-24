import os
import re
import unicodedata
import fitz  # PyMuPDF
from docx import Document

PHOTO_DIR = os.path.join("saved_resumes", "photos")
os.makedirs(PHOTO_DIR, exist_ok=True)

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


def normalize_exact_text(text):
    """Normalizes Unicode ligatures (fi, fl, etc.) without altering spelling."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    # Remove non-printable control characters while keeping exact words & punctuation
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
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


def extract_photo_from_pdf(pdf_path, candidate_id):
    """Extracts the first valid profile image from a PDF resume."""
    try:
        doc = fitz.open(pdf_path)
        for page_index in range(min(2, len(doc))):
            images = doc[page_index].get_images(full=True)
            for img in images:
                xref = img[0]
                base_image = doc.extract_image(xref)
                width, height = base_image["width"], base_image["height"]
                if 60 <= width <= 1500 and 60 <= height <= 1500:
                    ext = base_image["ext"]
                    photo_path = os.path.join(PHOTO_DIR, f"{candidate_id}.{ext}")
                    with open(photo_path, "wb") as f:
                        f.write(base_image["image"])
                    return photo_path
    except Exception:
        pass
    return None


def extract_photo_from_docx(docx_path, candidate_id):
    """Extracts embedded profile picture from a DOCX resume."""
    try:
        doc = Document(docx_path)
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                img_bytes = rel.target_part.blob
                ext = rel.target_ref.split(".")[-1]
                photo_path = os.path.join(PHOTO_DIR, f"{candidate_id}.{ext}")
                with open(photo_path, "wb") as f:
                    f.write(img_bytes)
                return photo_path
    except Exception:
        pass
    return None


def extract_verbatim_lines_from_pdf(pdf_path):
    """Extracts exact lines from PDF in strict top-to-bottom visual reading order."""
    doc = fitz.open(pdf_path)
    all_lines = []
    for page in doc:
        # sort=True ensures exact top-to-bottom, left-to-right reading order without scrambled words
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            if b[6] == 0:  # Text block only
                block_text = normalize_exact_text(b[4])
                for line in block_text.split("\n"):
                    clean = line.strip()
                    if clean:
                        all_lines.append(clean)
    return all_lines


def extract_verbatim_lines_from_docx(docx_path):
    """Extracts exact lines from paragraphs and tables in a DOCX resume."""
    doc = Document(docx_path)
    all_lines = []
    for p in doc.paragraphs:
        clean = normalize_exact_text(p.text)
        if clean:
            all_lines.append(clean)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    clean = normalize_exact_text(p.text)
                    if clean and clean not in all_lines:
                        all_lines.append(clean)
    return all_lines


def parse_resume_file(file_path, candidate_id, default_filename):
    """Extracts exact verbatim text and categorizes into standardized resume sections."""
    ext = os.path.splitext(file_path)[1].lower()
    lines = []
    photo_path = None

    if ext == ".pdf":
        photo_path = extract_photo_from_pdf(file_path, candidate_id)
        lines = extract_verbatim_lines_from_pdf(file_path)
    elif ext == ".docx":
        photo_path = extract_photo_from_docx(file_path, candidate_id)
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
        elif current_section:
            # Preserve exact verbatim wording of every line
            sections[current_section].append(cleaned)

    return {
        "id": candidate_id,
        "filename": default_filename,
        "file_path": file_path,
        "photo_path": photo_path,
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
        "experience": "\n".join(sections["experience"]) or raw_text,
        "skills": "\n".join(sections["skills"]) or "None listed.",
        "certificates": "\n".join(sections["certificates"]) or "None listed.",
    }