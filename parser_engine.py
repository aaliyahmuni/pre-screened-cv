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

ALL_BULLET_CHARS_PATTERN = r"^[\s•\-\*▪▫■□●○◦‣⁃➢➣➤➔→»›◆◇❖✓✔+·∙∘]+"


def strip_all_bullet_shapes(text):
    """Removes every bullet character shape so only clean • bullets are applied."""
    return re.sub(ALL_BULLET_CHARS_PATTERN, "", text).strip()


def is_noise_or_garbage(text):
    """Filters out page numbers and garbled OCR artifacts."""
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
    """Normalizes Unicode ligatures and removes private-use Wingdings/box glyphs."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    text = re.sub(r"[\ue000-\uf8ff\u25a0-\u27bf]", "", text)
    # Collapse multiple spaces within the same line without altering words
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


def strip_header_numbering(text):
    """Removes any numbering beside headers and subheaders while preserving exact wording."""
    cleaned = re.sub(
        r"^\s*(?:[\d]+(?:\.[\d]+)*[\.\)\-:]?|[A-Za-z][\.\)])\s+",
        "",
        text,
    )
    return cleaned.strip()


def is_standalone_subheader(line):
    """Identifies company names, job titles, date ranges, or degrees that must stay on their own line."""
    stripped = strip_all_bullet_shapes(line).strip()
    if not stripped or len(stripped) > 105:
        return False
    if stripped.endswith("."):
        return False
    if stripped.endswith(":"):
        return True
    # Date ranges (e.g., Jan 2021 - Present, 2019 - 2022)
    if re.search(r"\b(19|20)\d{2}\b", stripped) and len(stripped) < 85:
        return True
    # All-caps company/role/institution header (e.g., "PETRONAS CARIGALI SDN BHD")
    if stripped.isupper() and 4 <= len(stripped) <= 85:
        return True
    # Short title/company line containing role or organization keywords
    role_org_words = (
        "sdn bhd", "bhd", "ltd", "inc", "corp", "company", "engineering",
        "services", "offshore", "oil", "gas", "petronas", "shell",
        "university", "universiti", "college", "institute", "politeknik",
        "diploma", "degree", "bachelor", "master", "certificate", "spm", "stpm",
        "engineer", "supervisor", "technician", "welder", "fitter", "inspector",
        "manager", "officer", "coordinator", "lead", "rigger", "operator",
        "specialist", "advisor", "accountant", "controller", "storekeeper"
    )
    if len(stripped.split()) <= 10 and any(w in stripped.lower() for w in role_org_words):
        return True
    return False


def merge_wrapped_lines(lines, is_prose=False):
    """
    Rejoins PDF line-wrapped sentences so one full bullet or paragraph is never split
    into multiple separate bullets.
    """
    if not lines:
        return ""

    merged = []
    for raw_line in lines:
        clean = strip_header_numbering(raw_line).strip()
        if not clean or is_noise_or_garbage(clean):
            continue

        starts_with_bullet = bool(re.match(ALL_BULLET_CHARS_PATTERN, clean))
        text_no_bullet = strip_all_bullet_shapes(clean)
        if not text_no_bullet:
            continue

        if not merged:
            merged.append(f"• {text_no_bullet}" if starts_with_bullet else text_no_bullet)
            continue

        prev = merged[-1]
        prev_no_bullet = strip_all_bullet_shapes(prev)

        # Keep subheaders and explicit new bullets on separate lines
        if starts_with_bullet or is_standalone_subheader(text_no_bullet) or is_standalone_subheader(prev_no_bullet):
            merged.append(f"• {text_no_bullet}" if starts_with_bullet else text_no_bullet)
            continue

        # Join wrapped continuation lines into the same full sentence/bullet
        prev_ended_sentence = bool(re.search(r"[\.\!\?;:]$", prev_no_bullet))
        starts_continuation = bool(re.match(r"^[a-z0-9\(,&/\-]", text_no_bullet))

        if is_prose:
            # Objective paragraphs merge cleanly into cohesive sentences
            if not prev_ended_sentence or starts_continuation:
                merged[-1] = f"{prev} {text_no_bullet}"
            else:
                merged.append(text_no_bullet)
        else:
            # In bulleted sections, join if previous line did NOT finish a sentence
            # or if the next line starts with a lowercase/continuation word
            if not prev_ended_sentence or starts_continuation:
                merged[-1] = f"{prev} {text_no_bullet}"
            else:
                merged.append(text_no_bullet)

    return "\n".join(merged)


def extract_candidate_name(lines, raw_text, default_filename):
    """Extracts the candidate's actual full name, skipping cover-letter greetings."""
    explicit_name = re.search(
        r"^(?:Full\s+)?Name\s*[:\-]\s*([A-Za-z][A-Za-z\s\.'/@\-]{2,60})$",
        raw_text,
        re.MULTILINE | re.IGNORECASE,
    )
    if explicit_name:
        return strip_header_numbering(explicit_name.group(1)).strip()

    signoff_match = re.search(
        r"(?:Yours\s+(?:sincerely|faithfully)|Sincerely|Best\s+regards|Regards)[,\s]*\n+([A-Za-z][A-Za-z\s\.'/@\-]{2,55})",
        raw_text,
        re.IGNORECASE,
    )
    if signoff_match:
        sig_name = signoff_match.group(1).strip().split("\n")[0].strip()
        if 2 <= len(sig_name.split()) <= 6:
            return strip_header_numbering(sig_name)

    ignore_phrases = (
        "dear", "hiring manager", "to whom", "sir", "madam", "curriculum vitae",
        "resume", "application", "cover letter", "personal details",
        "personal particulars", "profile", "objective", "summary", "contact",
        "phone", "email", "address", "position", "date", "ref:", "subject:", "re:",
    )

    for line in lines[:25]:
        cleaned = strip_all_bullet_shapes(strip_header_numbering(line)).rstrip(",").strip()
        lower_c = cleaned.lower()

        if not cleaned or len(cleaned) < 4 or len(cleaned) > 55:
            continue
        if any(phrase in lower_c for phrase in ignore_phrases):
            continue
        if "@" in cleaned or re.search(r"\d{3,}", cleaned):
            continue
        if cleaned.endswith(".") or cleaned.endswith(":"):
            continue

        words = cleaned.split()
        if 2 <= len(words) <= 6 and re.match(r"^[A-Za-z\s\.'/@\-]+$", cleaned):
            return cleaned

    base = os.path.splitext(default_filename)[0]
    base = re.sub(r"(?i)\b(resume|cv|curriculum[\s_\-]*vitae|updated|latest|new|\d+)\b", " ", base)
    base = re.sub(r"[_\-]+", " ", base).strip()
    return base if len(base) >= 3 else os.path.splitext(default_filename)[0]


def extract_personal_details(lines, raw_text):
    """Extracts structured personal details accurately."""
    email_match = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", raw_text)
    phone_match = re.search(r"(\+?\d[\d\s\-\(\)]{7,16}\d)", raw_text)
    gender_match = re.search(r"\b(?:Gender|Sex)\s*[:\-]?\s*(Male|Female)\b|\b(Male|Female)\b", raw_text, re.IGNORECASE)
    age_match = re.search(r"\b(?:Age\s*[:\-]?\s*(\d{2})|(\d{2})\s*years\s*old)\b", raw_text, re.IGNORECASE)
    marital_match = re.search(r"\b(?:Marital\s*Status\s*[:\-]?\s*)?(Single|Married|Divorced|Widowed)\b", raw_text, re.IGNORECASE)
    nationality_match = re.search(r"Nationality\s*[:\-]?\s*([A-Za-z\s]{3,30})", raw_text, re.IGNORECASE)

    address_val = "N/A"
    for idx, line in enumerate(lines):
        addr_m = re.match(r"^(?:Permanent\s+|Current\s+|Home\s+)?(?:Address|Location)\s*[:\-]?\s*(.*)$", line, re.IGNORECASE)
        if addr_m:
            parts = [addr_m.group(1).strip()] if addr_m.group(1).strip() else []
            for nxt in lines[idx + 1 : idx + 3]:
                if re.search(r"\b\d{5}\b|sarawak|sabah|johor|selangor|kuala lumpur|penang|perak|pahang|kedah|kelantan|terengganu|melaka|negeri sembilan|perlis|malaysia", nxt, re.IGNORECASE) and ":" not in nxt:
                    parts.append(nxt.strip())
                else:
                    break
            if parts:
                address_val = ", ".join(parts)
            break

    gender_val = "N/A"
    if gender_match:
        gender_val = (gender_match.group(1) or gender_match.group(2)).capitalize()

    nat_val = "N/A"
    if nationality_match:
        nat_val = nationality_match.group(1).strip().split("\n")[0].strip()

    return {
        "phone": phone_match.group(1).strip() if phone_match else "N/A",
        "email": email_match.group(0).strip() if email_match else "N/A",
        "gender": gender_val,
        "age": (age_match.group(1) or age_match.group(2)) if age_match else "N/A",
        "nationality": nat_val,
        "marital_status": marital_match.group(1).capitalize() if marital_match else "N/A",
        "address": address_val,
    }


def extract_jobs_from_pdf():
    """Returns the standardized job position list for the dropdown."""
    return JOB_LIST


def render_pdf_pages_as_highres_images(pdf_path, prefix_id):
    """Renders PDF pages at crisp 300 DPI PNGs so certificates are sharp and readable."""
    rendered_paths = []
    try:
        doc = fitz.open(pdf_path)
        zoom_matrix = fitz.Matrix(300 / 72, 300 / 72)
        for idx, page in enumerate(doc):
            pix = page.get_pixmap(matrix=zoom_matrix, alpha=False)
            out_path = os.path.join(CERT_IMG_DIR, f"{prefix_id}_page{idx+1}.png")
            pix.save(out_path)
            rendered_paths.append(out_path)
    except Exception:
        pass
    return rendered_paths


def extract_images_from_pdf(pdf_path, candidate_id):
    """Extracts profile photo from page 1 and high-res (300 DPI) certificate pages."""
    profile_photo = None
    cert_images = []

    try:
        doc = fitz.open(pdf_path)
        zoom_matrix = fitz.Matrix(300 / 72, 300 / 72)

        for page_index, page in enumerate(doc):
            images = page.get_images(full=True)
            page_text = page.get_text("text").strip()
            clean_words = re.findall(r"[A-Za-z]{3,}", page_text)

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

            if page_index >= 1 and len(images) > 0:
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
    """Extracts profile photo and original full-resolution certificate images from DOCX."""
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


def match_section_header(line):
    """
    Strictly matches standalone resume section headers so normal sentences inside
    experience or education never accidentally trigger a section switch.
    """
    cleaned = strip_all_bullet_shapes(strip_header_numbering(line)).rstrip(":").strip()
    if not cleaned or len(cleaned) > 45 or cleaned.endswith("."):
        return None

    lower_clean = re.sub(r"\s+", " ", cleaned.lower())

    exact_section_map = {
        "objective": {
            "objective", "career objective", "job objective", "professional objective",
            "professional summary", "executive summary", "career summary", "summary",
            "profile", "personal profile", "career profile", "about me",
        },
        "education": {
            "education", "educational", "educational background", "educational qualification",
            "educational qualifications", "academic", "academic background",
            "academic qualification", "academic qualifications", "qualification", "qualifications",
        },
        "courses": {
            "course", "courses", "training", "trainings", "courses & training",
            "courses and training", "course & training", "course & training attended",
            "courses & training attended", "training attended", "trainings attended",
            "professional training", "workshops", "seminars", "workshops & seminars",
        },
        "experience": {
            "experience", "experiences", "working experience", "work experience",
            "employment", "employment history", "work history", "career history",
            "professional experience", "relevant experience", "industrial experience",
        },
        "skills": {
            "skills", "skill", "capabilities", "skills & capabilities",
            "skills and capabilities", "key skills", "core skills",
            "core competencies", "competencies", "technical skills", "areas of expertise",
        },
        "certificates": {
            "certificate", "certificates", "certification", "certifications",
            "licenses", "licences", "licenses & certifications",
            "certifications & licenses", "professional certificates", "professional certifications",
        },
        "ignore": {
            "personal details", "personal particulars", "personal information",
            "reference", "references", "character reference", "character references",
            "referee", "referees", "declaration", "language", "languages", "hobbies", "interests",
        },
    }

    for key, exact_names in exact_section_map.items():
        if lower_clean in exact_names:
            return key
    return None


def parse_resume_file(file_path, candidate_id, default_filename):
    """Extracts exact, unmixed, sentence-complete resume sections."""
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
    candidate_name = extract_candidate_name(lines, raw_text, default_filename)
    personal = extract_personal_details(lines, raw_text)

    sections = {
        "objective": [],
        "education": [],
        "courses": [],
        "experience": [],
        "skills": [],
        "certificates": [],
        "ignore": [],
    }
    current_section = None

    for line in lines:
        matched_key = match_section_header(line)
        if matched_key:
            current_section = matched_key
        elif current_section and current_section != "ignore":
            sections[current_section].append(strip_header_numbering(line))

    # Reconstruct full sentences/bullets cleanly per section; default to "N/A" if empty
    objective_text = merge_wrapped_lines(sections["objective"], is_prose=True) or "N/A"
    education_text = merge_wrapped_lines(sections["education"], is_prose=False) or "N/A"
    courses_text = merge_wrapped_lines(sections["courses"], is_prose=False) or "N/A"
    experience_text = merge_wrapped_lines(sections["experience"], is_prose=False) or "N/A"
    skills_text = merge_wrapped_lines(sections["skills"], is_prose=False) or "N/A"
    certificates_text = merge_wrapped_lines(sections["certificates"], is_prose=False) or "N/A"

    return {
        "id": candidate_id,
        "filename": default_filename,
        "file_path": file_path,
        "photo_path": photo_path,
        "cert_images": cert_images,
        "position": "",
        "name": candidate_name,
        "phone": personal["phone"],
        "email": personal["email"],
        "gender": personal["gender"],
        "age": personal["age"],
        "nationality": personal["nationality"],
        "marital_status": personal["marital_status"],
        "address": personal["address"],
        "objective": objective_text,
        "education": education_text,
        "courses": courses_text,
        "experience": experience_text,
        "skills": skills_text,
        "certificates": certificates_text,
    }