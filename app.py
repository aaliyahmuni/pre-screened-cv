import os
import json
import uuid
import streamlit as st
from parser_engine import (
    extract_jobs_from_pdf,
    parse_resume_file,
    render_pdf_pages_as_highres_images,
)
from docx_exporter import export_all_to_master_docx

st.set_page_config(page_title="Pre-Screened CV Portal", layout="wide")

DB_FILE = "database.json"
SAVE_DIR = "saved_resumes"
os.makedirs(SAVE_DIR, exist_ok=True)

VALID_USERS = {
    "admin": "admin123",
    "hr@company.com": "hr2026",
    "staff@company.com": "staff2026",
}


def load_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user" not in st.session_state:
    st.session_state.user = ""


def login_page():
    col1, col2, col3 = st.columns([1, 1.3, 1])
    with col2:
        st.write("")
        st.write("")
        if os.path.exists("logo 2.png"):
            st.image("logo 2.png", use_container_width=True)
        else:
            st.warning("logo 2.png not found in pre-screened-cv folder.")

        st.markdown("### Staff Login")
        username = st.text_input("Username or Email", placeholder="Enter username or email")

        pw_col, toggle_col = st.columns([4, 1])
        with toggle_col:
            st.write("")
            st.write("")
            show_pw = st.checkbox("View", value=False)
        with pw_col:
            password = st.text_input(
                "Password",
                type="default" if show_pw else "password",
                placeholder="Enter password",
            )

        if st.button("Log In", type="primary", use_container_width=True):
            if username.strip() in VALID_USERS and VALID_USERS[username.strip()] == password:
                st.session_state.authenticated = True
                st.session_state.user = username.strip()
                st.rerun()
            else:
                st.error("Invalid Username/Email or Password.")


def main_portal():
    with st.sidebar:
        if os.path.exists("logo.png"):
            st.image("logo.png", use_container_width=True)
        st.divider()
        st.write(f"**Logged in as:** {st.session_state.user}")
        if st.button("Log Out", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.user = ""
            st.rerun()

    st.title("Pre-Screened Candidate Resume Manager")

    job_list = extract_jobs_from_pdf()
    candidates = load_db()

    st.subheader("Upload Candidate Resumes")
    default_job = st.selectbox(
        "Default Position Applied For (Batch Upload)",
        options=job_list,
    )

    uploaded_files = st.file_uploader(
        "Select candidate resumes (PDF or DOCX — Unlimited size, auto-saved)",
        type=["pdf", "docx"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        existing_filenames = {c["filename"] for c in candidates}
        newly_added = 0

        for uf in uploaded_files:
            if uf.name not in existing_filenames:
                cid = str(uuid.uuid4())[:8]
                dest_path = os.path.join(SAVE_DIR, f"{cid}_{uf.name}")
                with open(dest_path, "wb") as f:
                    f.write(uf.getbuffer())

                parsed_data = parse_resume_file(dest_path, cid, uf.name)
                parsed_data["position"] = default_job
                candidates.append(parsed_data)
                newly_added += 1

        if newly_added > 0:
            save_db(candidates)
            st.success(f"Saved {newly_added} new resume(s) permanently.")
            st.rerun()

    st.divider()
    st.subheader(f"Saved Candidate Resumes ({len(candidates)})")

    if not candidates:
        st.info("No resumes uploaded yet. Upload files above to begin.")
        return

    updated = False
    to_delete = None

    for idx, cand in enumerate(candidates):
        with st.expander(f"{cand['name']}  |  Applied For: {cand.get('position', 'Unassigned')}", expanded=False):
            top_left, top_right = st.columns([1, 3])

            with top_left:
                if cand.get("photo_path") and os.path.exists(cand["photo_path"]):
                    st.image(cand["photo_path"], width=150)
                else:
                    st.caption("No photo in resume (optional)")

                new_photo = st.file_uploader(
                    "Upload/Replace Photo", type=["png", "jpg", "jpeg"], key=f"photo_{cand['id']}"
                )
                if new_photo:
                    p_path = os.path.join(SAVE_DIR, "photos", f"{cand['id']}_{new_photo.name}")
                    with open(p_path, "wb") as pf:
                        pf.write(new_photo.getbuffer())
                    cand["photo_path"] = p_path
                    updated = True

            with top_right:
                current_pos = cand.get("position", job_list[0])
                pos_index = job_list.index(current_pos) if current_pos in job_list else 0
                selected_pos = st.selectbox(
                    "Position Applied For (Appears in Document Footer)",
                    options=job_list,
                    index=pos_index,
                    key=f"pos_{cand['id']}",
                )
                if selected_pos != cand.get("position"):
                    cand["position"] = selected_pos
                    updated = True

                c1, c2 = st.columns(2)
                cand["name"] = c1.text_input("Full Name", cand["name"], key=f"name_{cand['id']}")
                cand["phone"] = c2.text_input("Phone Number", cand["phone"], key=f"phone_{cand['id']}")
                cand["email"] = c1.text_input("Email", cand["email"], key=f"email_{cand['id']}")
                cand["gender"] = c2.text_input("Gender", cand["gender"], key=f"gender_{cand['id']}")
                cand["age"] = c1.text_input("Age", cand["age"], key=f"age_{cand['id']}")
                cand["nationality"] = c2.text_input("Nationality", cand["nationality"], key=f"nat_{cand['id']}")
                cand["marital_status"] = c1.text_input("Marital Status", cand["marital_status"], key=f"mar_{cand['id']}")
                cand["address"] = c2.text_input("Address", cand["address"], key=f"addr_{cand['id']}")

            cand["objective"] = st.text_area("Objective", cand["objective"], height=90, key=f"obj_{cand['id']}")
            cand["education"] = st.text_area("Education", cand["education"], height=110, key=f"edu_{cand['id']}")
            cand["courses"] = st.text_area("Course & Training Attended", cand["courses"], height=110, key=f"crs_{cand['id']}")
            cand["experience"] = st.text_area("Working Experience", cand["experience"], height=160, key=f"exp_{cand['id']}")
            cand["skills"] = st.text_area("Skills & Capabilities", cand["skills"], height=110, key=f"skl_{cand['id']}")
            cand["certificates"] = st.text_area("Certificates (Text List)", cand["certificates"], height=110, key=f"crt_{cand['id']}")

            # Direct High-Resolution Certificate Insertion (PNG, JPG, or PDF)
            existing_cert_imgs = [p for p in cand.get("cert_images", []) if os.path.exists(p)]
            if existing_cert_imgs:
                st.caption(f"Attached Certificates ({len(existing_cert_imgs)}):")
                img_cols = st.columns(min(4, len(existing_cert_imgs)))
                for i, c_img in enumerate(existing_cert_imgs):
                    col = img_cols[i % 4]
                    col.image(c_img, use_container_width=True)
                    if col.button(f"Remove #{i+1}", key=f"rm_cert_{cand['id']}_{i}"):
                        cand["cert_images"].remove(c_img)
                        save_db(candidates)
                        st.rerun()

            extra_certs = st.file_uploader(
                "Insert Certificate Files (High-Res PNG, JPG, or PDF)",
                type=["png", "jpg", "jpeg", "pdf"],
                accept_multiple_files=True,
                key=f"cert_up_{cand['id']}",
            )
            if extra_certs:
                cert_list = cand.get("cert_images", [])
                for ec in extra_certs:
                    ec_raw_path = os.path.join(SAVE_DIR, "cert_images", f"{cand['id']}_{ec.name}")
                    if not os.path.exists(ec_raw_path):
                        with open(ec_raw_path, "wb") as ef:
                            ef.write(ec.getbuffer())

                        if ec.name.lower().endswith(".pdf"):
                            rendered_pages = render_pdf_pages_as_highres_images(
                                ec_raw_path, f"{cand['id']}_{os.path.splitext(ec.name)[0]}"
                            )
                            for rp in rendered_pages:
                                if rp not in cert_list:
                                    cert_list.append(rp)
                        else:
                            if ec_raw_path not in cert_list:
                                cert_list.append(ec_raw_path)
                        updated = True
                cand["cert_images"] = cert_list

            btn_save, btn_del = st.columns([1, 5])
            if btn_save.button("Save Changes", key=f"save_{cand['id']}"):
                save_db(candidates)
                st.success("Candidate details updated.")
            if btn_del.button("Remove Candidate", key=f"del_{cand['id']}"):
                to_delete = idx

    if updated:
        save_db(candidates)

    if to_delete is not None:
        candidates.pop(to_delete)
        save_db(candidates)
        st.rerun()

    st.divider()
    st.subheader("Export Compiled Document for Boss Review")
    if st.button("Generate Consolidated Word Document (etc.docx)", type="primary"):
        output_filename = "Compiled_PreScreened_Resumes.docx"
        export_all_to_master_docx(candidates, output_filename)
        with open(output_filename, "rb") as f:
            st.download_button(
                label="Download Consolidated Resumes (.docx)",
                data=f,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )


if st.session_state.authenticated:
    main_portal()
else:
    login_page()