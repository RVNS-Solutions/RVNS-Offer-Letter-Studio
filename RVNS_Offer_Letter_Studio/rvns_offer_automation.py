from __future__ import annotations
import argparse, base64, csv, json, re, shutil, subprocess, sys, tempfile, uuid
from datetime import datetime, date
from email.message import EmailMessage
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
TEMPLATE_DIR = ROOT / "templates"
OUTPUT_DIR = ROOT / "output"
ASSET_DIR = ROOT / "assets"
GMAIL_SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
]

ALIASES = {
    "student_name": ["student_name", "name", "full_name", "candidate_name", "intern_name"],
    "email": ["email", "email_id", "mail", "student_email", "gmail", "email_address"],
    "college_name": ["college_name", "college", "institution", "institute_name", "university_college"],
    "branch": ["branch", "department", "branch_department", "course_branch", "stream"],
    "internship_domain": ["internship_domain", "domain", "course", "internship", "technology", "program", "position", "role"],
    "internship_mode": ["internship_mode", "mode", "work_mode", "training_mode", "delivery_mode"],
    "start_date": ["start_date", "commencement_date", "joining_date", "from_date", "internship_start_date"],
    "end_date": ["end_date", "completion_date", "last_date", "to_date", "internship_end_date"],
    "duration": ["duration", "internship_duration", "period", "tenure"],
    "office_location": ["office_location", "location", "work_location", "internship_location", "address"],
    "offer_date": ["offer_date", "letter_date", "issue_date"],
    "mentor_name": ["mentor_name", "mentor", "supervisor"],
    "stipend": ["stipend", "salary", "monthly_stipend"],
    "additional_notes": ["additional_notes", "notes", "remarks", "comments"],
}

def norm(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(value or "").strip().lower()).strip("_")

def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d %B %Y")
    value = str(value).strip()
    return "" if value.lower() in {"none", "nan", "n/a", "na", "null", "-"} else value

def fmt_date(value: Any) -> str:
    value = clean(value)
    if not value:
        return ""
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%d %B %Y", "%d %b %Y"):
            try:
                return datetime.strptime(value, fmt).strftime("%d %B %Y")
            except ValueError:
                continue
    return value

def load_students(path: Path, sheet_name: str | None = None):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    headers = next(rows, None)
    if not headers:
        raise ValueError("The selected worksheet is empty.")
    header_map = {norm(h): i for i, h in enumerate(headers) if h is not None}
    columns = {}
    for field, aliases in ALIASES.items():
        for alias in aliases:
            if norm(alias) in header_map:
                columns[field] = header_map[norm(alias)]
                break
    if "student_name" not in columns:
        raise ValueError("Student Name column is required. Use Student_Name, Name, or Full_Name.")
    students = []
    for row_number, row in enumerate(rows, start=2):
        if not row or not any(v is not None and str(v).strip() for v in row):
            continue
        item = {field: clean(row[idx] if idx < len(row) else None) for field, idx in columns.items()}
        for field in ("start_date", "end_date", "offer_date"):
            item[field] = fmt_date(item.get(field, ""))
        if not item.get("offer_date"):
            item["offer_date"] = datetime.now().strftime("%d %B %Y")
        item["_row"] = row_number
        if not item.get("student_name"):
            item["_error"] = "Student name is blank."
        if not item.get("internship_domain"):
            item["internship_domain"] = "Internship"
        students.append(item)
    wb.close()
    return students

def mode_list(student, override="From Excel"):
    raw = (student.get("internship_mode") or "").strip().lower()
    if override != "From Excel":
        raw = override.lower()
    if raw in {"both", "online/offline", "offline/online", "online and offline"}:
        return ["Offline", "Online"]
    if raw in {"online", "remote", "virtual"}:
        return ["Online"]
    if raw in {"offline", "onsite", "on-site", "in-person", "in person"}:
        return ["Offline"]
    if raw in {"general", "not mentioned", "none", ""}:
        return ["General"]
    return ["General"]

def iter_paragraphs(container):
    for paragraph in container.paragraphs:
        yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from iter_paragraphs(cell)

def set_text_keep_style(paragraph, text):
    # Paragraphs containing pictures are deliberately not rewritten.
    if paragraph._p.xpath(".//w:drawing"):
        return
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)

def remove_paragraph(paragraph):
    element = paragraph._element
    element.getparent().remove(element)


def apply_authorized_signature(doc):
    """Place the supplied RVNS signature/seal neatly in the signatory block."""
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph

    signature_path = ASSET_DIR / "rvns_authorized_signature_seal.png"
    if not signature_path.exists():
        raise FileNotFoundError(f"Authorized signature image not found: {signature_path}")

    body_paragraphs = list(doc.paragraphs)
    for paragraph in body_paragraphs:
        for drawing in list(paragraph._p.xpath(".//w:drawing")):
            drawing.getparent().remove(drawing)

    md_index = next((i for i, p in enumerate(body_paragraphs)
                     if "Managing Director" in p.text), None)
    if md_index is None:
        raise RuntimeError("Managing Director line not found in offer-letter template.")

    # Remove empty template spacer paragraphs between Authorized Signatory and Managing Director.
    sign_index = next((i for i in range(md_index - 1, -1, -1)
                       if "Authorized Signatory" in body_paragraphs[i].text), None)
    if sign_index is not None:
        for paragraph in body_paragraphs[sign_index + 1:md_index]:
            if not paragraph.text.strip():
                paragraph._element.getparent().remove(paragraph._element)

    md_paragraph = next(p for p in doc.paragraphs if "Managing Director" in p.text)
    new_p = OxmlElement("w:p")
    md_paragraph._p.addprevious(new_p)
    signature_paragraph = Paragraph(new_p, md_paragraph._parent)
    signature_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    signature_paragraph.paragraph_format.space_before = Pt(0)
    signature_paragraph.paragraph_format.space_after = Pt(1)
    signature_paragraph.paragraph_format.line_spacing = 1.0
    signature_paragraph.add_run().add_picture(str(signature_path), width=Inches(1.15))

def personalize(student, mode, destination):
    from docx import Document
    from docx.shared import Pt
    template = TEMPLATE_DIR / ("online_template.docx" if mode == "Online" else "offline_template.docx")
    doc = Document(template)
    name = student.get("student_name", "")
    domain = student.get("internship_domain", "Internship")
    start, end = student.get("start_date", ""), student.get("end_date", "")
    duration = student.get("duration", "")
    location = student.get("office_location", "")
    replacements = {
        "[STUDENT NAME]": name,
        "[COLLEGE NAME]": student.get("college_name", ""),
        "[BRANCH / DEPARTMENT]": student.get("branch", ""),
        "[INTERNSHIP DOMAIN]": domain,
        "[OFFICE LOCATION]": location,
        "22 October 2026": start,
        "18 November 2026": end,
        "22 November 2026": start,
        "19 December 2026": end,
        "4 Weeks": duration,
    }
    containers = [doc]
    for section in doc.sections:
        containers.extend([section.header, section.footer, section.first_page_header, section.first_page_footer])
    visited = set()
    for container in containers:
        for p in iter_paragraphs(container):
            if id(p._p) in visited:
                continue
            visited.add(id(p._p))
            original = p.text
            if not original:
                continue
            # Dynamic issue date, independent of internship dates.
            if original.strip().startswith("Date:"):
                set_text_keep_style(p, "Date: " + student.get("offer_date", ""))
                continue
            new = original
            for old, value in replacements.items():
                new = new.replace(old, value)
            if mode == "General":
                new = new.replace(" (Offline Mode)", "").replace(" (Online Mode)", "")
                new = new.replace(" under our Offline Internship Program", " under our internship program")
                new = new.replace(" under our Online Internship Program", " under our internship program")
                new = new.replace("Offline Internship Program", "internship program")
                new = new.replace("Online Internship Program", "internship program")
                new = new.replace("The intern must attend the internship regularly at the assigned office location.",
                                  "The intern must participate regularly in the scheduled internship activities.")
            if new != original:
                set_text_keep_style(p, new)

    # Remove empty optional detail lines; retain all other original terms and clauses.
    for p in list(iter_paragraphs(doc)):
        t = p.text.strip()
        if not t:
            continue
        remove = False
        if t.startswith("College Name:") and not student.get("college_name"):
            remove = True
        elif t.startswith("Branch:") and not student.get("branch"):
            remove = True
        elif t.startswith("Internship Location:") and not location:
            remove = True
        elif t.startswith("Internship Duration:") and not duration:
            remove = True
        elif t.startswith("Commencement Date:") and not start:
            remove = True
        elif t.startswith("Completion Date:") and not end:
            remove = True
        elif t.startswith("Internship Mode:") and mode == "General":
            remove = True
        if remove:
            remove_paragraph(p)

    # Keep the letter clean and compact while preserving the existing two-page letterhead layout.
    for paragraph in iter_paragraphs(doc):
        if paragraph.text.strip():
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(3)
            paragraph.paragraph_format.line_spacing = 1.0

    # The actual supplied signature/seal is inserted once, in the signatory block.
    apply_authorized_signature(doc)

    destination.parent.mkdir(parents=True, exist_ok=True)
    doc.save(destination)

def convert_pdf(docx_path: Path):
    pdf_path = docx_path.with_suffix(".pdf")
    try:
        from docx2pdf import convert
        convert(str(docx_path), str(pdf_path))
        if pdf_path.exists():
            return pdf_path
    except Exception:
        pass
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if office:
        profile_dir = Path(tempfile.gettempdir()) / ("rvns_lo_profile_" + uuid.uuid4().hex)
        subprocess.run([office, f"-env:UserInstallation={profile_dir.as_uri()}", "--headless",
                        "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
                       check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180)
        if pdf_path.exists():
            return pdf_path
    raise RuntimeError("PDF conversion requires Microsoft Word (Windows) or LibreOffice. Install one and retry.")

def generate(excel_path, sheet_name=None, override="From Excel", output_dir=OUTPUT_DIR, progress=None):
    students = load_students(Path(excel_path), sheet_name)
    if not students:
        raise ValueError("No student records found.")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    results, errors = [], []
    for student in students:
        if student.get("_error"):
            errors.append({"row": student["_row"], "student": student.get("student_name", ""),
                           "email": student.get("email", ""), "error": student["_error"]})
            continue
        for mode in mode_list(student, override):
            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", student["student_name"]).strip("_") or "Student"
            safe_domain = re.sub(r"[^A-Za-z0-9_-]+", "_", student["internship_domain"]).strip("_") or "Internship"
            stem = f"{safe_name}_{safe_domain}_{mode}_Row{student['_row']}"
            docx = output_dir / f"{stem}.docx"
            try:
                personalize(student, mode, docx)
                pdf = convert_pdf(docx)
                results.append({"student": student, "mode": mode, "docx": docx, "pdf": pdf,
                                "email_status": "NOT_CREATED"})
                if progress:
                    progress(f"Generated {mode} letter: {student['student_name']}")
            except Exception as exc:
                errors.append({"row": student["_row"], "student": student["student_name"],
                               "email": student.get("email", ""), "error": str(exc)})
    write_report(results, errors, output_dir)
    return results, errors

def write_report(results, errors, output_dir):
    path = Path(output_dir) / "generation_report.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Student", "Email", "College", "Domain", "Mode", "Start Date", "End Date",
                         "Duration", "Word File", "PDF File", "Status"])
        for item in results:
            s = item["student"]
            writer.writerow([s.get("student_name"), s.get("email"), s.get("college_name"),
                             s.get("internship_domain"), item["mode"], s.get("start_date"),
                             s.get("end_date"), s.get("duration"), item["docx"].name,
                             item["pdf"].name, "PDF GENERATED"])
        for item in errors:
            writer.writerow([item["student"], item["email"], "", "", "", "", "", "", "", "", "ERROR: " + item["error"]])
    return path

def gmail_service():
    """Build Gmail API service using Streamlit Cloud Secrets or local OAuth files."""
    try:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Gmail libraries are missing. Check requirements.txt and redeploy."
        ) from exc

    # Streamlit Cloud: credentials are stored in App settings -> Secrets.
    # Expected TOML:
    # [gmail]
    # client_id = "..."
    # client_secret = "..."
    # refresh_token = "..."
    try:
        import streamlit as st
        gmail_cfg = st.secrets.get("gmail", {})
    except Exception:
        gmail_cfg = {}

    if gmail_cfg:
        client_id = str(gmail_cfg.get("client_id", "")).strip()
        client_secret = str(gmail_cfg.get("client_secret", "")).strip()
        refresh_token = str(gmail_cfg.get("refresh_token", "")).strip()
        token_uri = str(
            gmail_cfg.get("token_uri", "https://oauth2.googleapis.com/token")
        ).strip()

        missing = [
            key for key, value in (
                ("client_id", client_id),
                ("client_secret", client_secret),
                ("refresh_token", refresh_token),
            ) if not value
        ]
        if missing:
            raise RuntimeError(
                "Gmail Secrets are incomplete. Add client_id, client_secret, "
                "and refresh_token under the [gmail] section."
            )

        credentials = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri=token_uri,
            client_id=client_id,
            client_secret=client_secret,
            scopes=GMAIL_SCOPES,
        )
        try:
            credentials.refresh(Request())
        except Exception as exc:
            raise RuntimeError(
                "Gmail OAuth refresh failed. Check the client ID, client secret, "
                "refresh token, and Gmail API authorization scopes in Streamlit Secrets."
            ) from exc
        return build("gmail", "v1", credentials=credentials)

    # Local Windows: retain the existing credentials.json/token.json workflow.
    token = ROOT / "token.json"
    credentials = None
    if token.exists():
        try:
            token_data = json.loads(token.read_text(encoding="utf-8"))
            granted_scopes = set(token_data.get("scopes") or [])
            if not set(GMAIL_SCOPES).issubset(granted_scopes):
                token.unlink(missing_ok=True)
            else:
                credentials = Credentials.from_authorized_user_file(
                    str(token), GMAIL_SCOPES
                )
        except Exception:
            credentials = None
            token.unlink(missing_ok=True)

    if credentials and credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(Request())
            token.write_text(credentials.to_json(), encoding="utf-8")
        except Exception:
            credentials = None

    if not credentials or not credentials.valid:
        client_file = ROOT / "credentials.json"
        if not client_file.exists():
            raise FileNotFoundError(
                "For Streamlit Cloud, configure [gmail] secrets. "
                "For local use, add credentials.json and authorize Gmail."
            )
        flow = InstalledAppFlow.from_client_secrets_file(
            str(client_file), GMAIL_SCOPES
        )
        credentials = flow.run_local_server(
            port=0, access_type="offline", prompt="consent"
        )
        token.write_text(credentials.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=credentials)

def create_gmail_drafts(results, expected_sender="rvnssolutions@gmail.com"):
    service = gmail_service()
    sender = service.users().getProfile(userId="me").execute().get("emailAddress", "")
    if sender.lower() != expected_sender.lower():
        raise RuntimeError(f"Connected Gmail account is {sender}; expected {expected_sender}. No drafts created.")
    created, skipped = [], []
    for item in results:
        student = item["student"]
        recipient = (student.get("email") or "").strip()
        if not recipient or "@" not in recipient:
            skipped.append(f"{student['student_name']}: missing/invalid email")
            continue
        subject = f"Internship Offer Letter – {student.get('internship_domain', 'Internship')} | RVNS Solutions"
        body = (f"Dear {student['student_name']},\n\nGreetings from RVNS Solutions Private Limited!\n\n"
                f"Please find attached your personalized offer letter for the {student.get('internship_domain', 'internship')} program.\n\n"
                "Kindly review the attached letter and confirm your acceptance.\n\n"
                "Best regards,\nHR Department\nRVNS Solutions Private Limited\n"
                "rvnssolutions@gmail.com\nhttps://rvnssolutionspvtltd.in/")
        message = EmailMessage()
        message["To"] = recipient
        message["From"] = sender
        message["Subject"] = subject
        message.set_content(body)
        message.add_attachment(item["pdf"].read_bytes(), maintype="application", subtype="pdf",
                               filename=item["pdf"].name)
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")
        response = service.users().drafts().create(userId="me", body={"message": {"raw": raw}}).execute()
        created.append({"student": student["student_name"], "email": recipient,
                        "draft_id": response.get("id", ""), "pdf": item["pdf"].name})
    report = OUTPUT_DIR / "gmail_draft_report.csv"
    with report.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Student", "Email", "Gmail Draft ID", "PDF Attachment", "Status"])
        for item in created:
            writer.writerow([item["student"], item["email"], item["draft_id"], item["pdf"], "DRAFT CREATED"])
        for item in skipped:
            writer.writerow([item, "", "", "", "SKIPPED"])
    return created, skipped

def main():
    parser = argparse.ArgumentParser(description="RVNS Solutions Offer Letter Studio")
    parser.add_argument("--excel", required=True)
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--mode", default="From Excel",
                        choices=["From Excel", "Offline", "Online", "Both", "General"])
    args = parser.parse_args()
    results, errors = generate(args.excel, args.sheet, args.mode)
    print(f"Generated {len(results)} Word/PDF letters; {len(errors)} errors.")
    for item in results:
        print(f"OK | {item['student']['student_name']} | {item['mode']} | {item['pdf'].name}")
    for item in errors:
        print(f"ERROR | row {item['row']} | {item['student']} | {item['error']}")

if __name__ == "__main__":
    main()



def prepare_email_batches(results):
    """Group letters by student/email so Both mode sends one email with two PDF attachments."""
    batches = {}
    skipped = []
    for item in results:
        student = item["student"]
        name = (student.get("student_name") or "Student").strip()
        email = (student.get("email") or "").strip()
        if not email or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
            skipped.append(f"{name}: missing or invalid email address")
            continue
        key = (name.casefold(), email.casefold())
        if key not in batches:
            batches[key] = {
                "student": name,
                "email": email,
                "domain": student.get("internship_domain") or "Internship",
                "attachments": [],
            }
        batches[key]["attachments"].append(item["pdf"])
    return list(batches.values()), skipped


def send_gmail_messages(results, expected_sender="rvnssolutions@gmail.com"):
    """Send one email per student, attaching all generated PDF letters for that student."""
    batches, skipped = prepare_email_batches(results)
    if not batches:
        return [], skipped, []

    service = gmail_service()
    sender = service.users().getProfile(userId="me").execute().get("emailAddress", "")
    if sender.lower() != expected_sender.lower():
        raise RuntimeError(
            f"Connected Gmail account is {sender}; expected {expected_sender}. No emails sent."
        )

    sent, failed = [], []
    for batch in batches:
        message = EmailMessage()
        message["To"] = batch["email"]
        message["From"] = sender
        message["Subject"] = (
            f"Internship Offer Letter – {batch['domain']} | RVNS Solutions"
        )
        body = (
            f"Dear {batch['student']},\n\n"
            "Greetings from RVNS Solutions Private Limited!\n\n"
            "Please find attached your personalized internship offer letter(s). "
            "Kindly review the attached document(s) and confirm your acceptance.\n\n"
            "Best regards,\nHR Department\nRVNS Solutions Private Limited\n"
            "rvnssolutions@gmail.com\nhttps://rvnssolutionspvtltd.in/"
        )
        message.set_content(body)

        attachment_names = []
        missing_files = []
        for pdf_path in batch["attachments"]:
            pdf_path = Path(pdf_path)
            if not pdf_path.exists():
                missing_files.append(pdf_path.name)
                continue
            message.add_attachment(
                pdf_path.read_bytes(),
                maintype="application",
                subtype="pdf",
                filename=pdf_path.name,
            )
            attachment_names.append(pdf_path.name)

        if missing_files:
            failed.append({
                "student": batch["student"], "email": batch["email"],
                "attachments": "; ".join(missing_files),
                "error": "One or more PDF attachments are missing; no email sent for this student",
            })
            continue
        if not attachment_names:
            continue

        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")
        try:
            response = service.users().messages().send(
                userId="me", body={"raw": raw}
            ).execute()
            sent.append({
                "student": batch["student"],
                "email": batch["email"],
                "message_id": response.get("id", ""),
                "attachments": "; ".join(attachment_names),
                "status": "SENT",
            })
        except Exception as exc:
            failed.append({
                "student": batch["student"], "email": batch["email"],
                "attachments": "; ".join(attachment_names), "error": str(exc),
            })

    report_path = OUTPUT_DIR / "gmail_send_report.csv"
    with report_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Student", "Email", "Gmail Message ID", "PDF Attachments", "Status", "Error"])
        for item in sent:
            writer.writerow([
                item["student"], item["email"], item["message_id"],
                item["attachments"], item["status"], "",
            ])
        for item in failed:
            writer.writerow([
                item["student"], item["email"], "", item["attachments"],
                "FAILED", item["error"],
            ])
        for item in skipped:
            writer.writerow([item, "", "", "", "SKIPPED", "Missing/invalid email"])
    return sent, skipped, failed
