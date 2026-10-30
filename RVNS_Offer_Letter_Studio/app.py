from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from rvns_offer_automation import load_students, generate, create_gmail_drafts, prepare_email_batches, send_gmail_messages

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)

st.set_page_config(
    page_title="RVNS Offer Letter Studio",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {max-width:1500px;padding-top:1.1rem;padding-bottom:3rem}
.hero {padding:27px 32px;border-radius:19px;background:linear-gradient(115deg,#10243e 0%,#174d70 58%,#16849a 100%);color:#fff;margin-bottom:20px}
.hero h1 {color:#fff;font-size:2.15rem;font-weight:800;margin:0 0 7px 0}
.hero p {color:#d8e8f2;margin:0;font-size:1rem}
.eyebrow {font-size:.72rem;letter-spacing:2px;text-transform:uppercase;color:#f3c56d;font-weight:800;margin-bottom:9px}
.panel {border:1px solid rgba(100,130,160,.23);padding:17px 19px;border-radius:15px;background:rgba(125,155,185,.045);margin:8px 0}
.small-note {color:#6d7c8d;font-size:.86rem}
.action-card {min-height:185px;border:1px solid #d7e1ea;border-radius:16px;padding:18px;background:linear-gradient(180deg,#fff 0%,#f7fafc 100%);box-shadow:0 4px 14px rgba(16,36,62,.055)}
.action-card .num {font-size:.72rem;letter-spacing:1.5px;font-weight:800;color:#176b87;margin-bottom:7px}
.action-card h3 {font-size:1.08rem;margin:0 0 7px 0;color:#10243e}
.action-card p {font-size:.87rem;color:#66788a;margin-bottom:0;min-height:55px}
div[data-testid="stMetric"] {background:rgba(100,140,170,.07);border:1px solid rgba(100,140,170,.18);padding:12px 15px;border-radius:14px}
.stButton>button,.stDownloadButton>button {border-radius:10px;font-weight:700;min-height:2.7rem}
[data-testid="stSidebar"] {background:#f4f7fa}
</style>
<div class="hero">
  <div class="eyebrow">RVNS SOLUTIONS PRIVATE LIMITED · DOCUMENT AUTOMATION</div>
  <h1>Offer Letter Studio</h1>
  <p>Excel to personalized Word/PDF offer letters, Gmail drafts, and confirmed direct email sending.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## RVNS Solutions")
    st.caption("Offer Letter Studio · Local Admin Workspace")
    st.divider()
    st.markdown("**Workflow**")
    st.markdown("1. Import and validate Excel\n2. Generate Word + PDF\n3. Create Gmail drafts\n4. Review recipients\n5. Send approved emails")
    st.divider()
    st.markdown("**Company sender**")
    st.code("rvnssolutions@gmail.com")
    st.caption("Direct sending is available after recipient review and explicit confirmation.")
    st.divider()
    st.caption("Use this dashboard only on your trusted computer. Never share credentials.json or token.json.")

for key, default in {
    "results": [],
    "errors": [],
    "draft_report": None,
    "generated_for": None,
    "drafted_for": None,
    "send_report": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# Upload and mode rules
left, right = st.columns([1.2, .8], gap="large")
with left:
    st.markdown('<div class="panel"><b>01 · Student data</b><br><span class="small-note">Upload an Excel workbook (.xlsx), with one student per row.</span></div>', unsafe_allow_html=True)
    upload = st.file_uploader("Choose student Excel workbook", type=["xlsx"], label_visibility="collapsed")
with right:
    st.markdown('<div class="panel"><b>02 · Letter mode</b><br><span class="small-note">Choose the rule for all rows, or use each row’s Excel value.</span></div>', unsafe_allow_html=True)
    mode_choice = st.radio(
        "Select letter mode",
        ["From Excel", "Offline", "Online", "Both", "General"],
        horizontal=True,
        help="Both creates two separate letters per student. Blank/General in Excel creates a general letter.",
    )
    st.caption("Excel mode values: Online, Offline, Both, or blank/General.")

students, sheet, upload_path = [], None, None
bad = []
if upload is not None:
    upload_path = ROOT / "_uploaded_students.xlsx"
    upload_path.write_bytes(upload.getvalue())
    try:
        workbook = pd.ExcelFile(upload_path)
        sheet = st.selectbox("Worksheet", workbook.sheet_names, key="worksheet_select")
        students = load_students(upload_path, sheet)
        good = [s for s in students if not s.get("_error")]
        bad = [s for s in students if s.get("_error")]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Student rows", len(students))
        c2.metric("Ready", len(good))
        c3.metric("Needs correction", len(bad))
        c4.metric("Worksheets", len(workbook.sheet_names))

        st.markdown("### Student data preview")
        preview = [{
            "Excel Row": s.get("_row", ""),
            "Student Name": s.get("student_name", ""),
            "Email": s.get("email", ""),
            "College": s.get("college_name", ""),
            "Branch": s.get("branch", ""),
            "Internship Domain": s.get("internship_domain", ""),
            "Mode": s.get("internship_mode", ""),
            "Start Date": s.get("start_date", ""),
            "End Date": s.get("end_date", ""),
            "Duration": s.get("duration", ""),
            "Office Location": s.get("office_location", ""),
        } for s in students]
        st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)
        if bad:
            st.warning(f"{len(bad)} row(s) have a missing student name and will be skipped.")
        st.caption("Missing optional fields are omitted from letters. The app does not invent student details.")
    except Exception as exc:
        st.error(f"Could not read this workbook: {exc}")
        students, sheet = [], None
else:
    st.info("Upload an Excel workbook to begin. All four workflow actions remain visible below.")
    sample_path = ROOT / "students_sample.xlsx"
    if sample_path.exists():
        st.download_button(
            "Download sample Excel format",
            sample_path.read_bytes(),
            "RVNS_Students_Sample.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# Four workflow actions are displayed together in one row.
st.markdown("## 03 · Workflow actions")
st.caption("Create letters, prepare Gmail drafts, send approved emails, or download your documents.")
a1, a2, a3, a4 = st.columns(4, gap="medium")

with a1:
    st.markdown('<div class="action-card"><div class="num">ACTION 01</div><h3>📄 Create Offer Letters</h3><p>Generate personalized Word and PDF documents using Excel data and selected mode rules.</p></div>', unsafe_allow_html=True)
    can_generate = bool(upload_path and sheet and any(not s.get("_error") for s in students))
    if st.button("Create Word + PDF", key="create_letters", type="primary", use_container_width=True, disabled=not can_generate):
        try:
            with st.spinner("Generating and checking offer letters..."):
                results, errors = generate(upload_path, sheet, mode_choice, OUT)
            st.session_state.results = results
            st.session_state.errors = errors
            st.session_state.generated_for = f"{upload.name}|{sheet}|{mode_choice}"
            st.session_state.draft_report = None
            st.session_state.drafted_for = None
            st.session_state.send_report = None
            if errors:
                st.warning(f"Generated {len(results)} letter(s); {len(errors)} record(s) need attention.")
            else:
                st.success(f"Generated {len(results)} Word/PDF letter pair(s).")
        except Exception as exc:
            st.error(f"Generation failed: {exc}")

with a2:
    st.markdown('<div class="action-card"><div class="num">ACTION 02</div><h3>✉ Create Gmail Drafts</h3><p>Create Gmail drafts with the matching PDF attachments. Nothing is sent by this action.</p></div>', unsafe_allow_html=True)
    already_drafted = st.session_state.draft_report is not None
    if st.button("Create Gmail Drafts" if not already_drafted else "Drafts Already Created",
                 key="create_drafts", use_container_width=True,
                 disabled=not st.session_state.results or already_drafted):
        try:
            with st.spinner("Connecting to the authorized RVNS Gmail account..."):
                created, skipped = create_gmail_drafts(st.session_state.results)
            st.session_state.draft_report = created
            st.session_state.drafted_for = st.session_state.generated_for
            st.success(f"Created {len(created)} Gmail draft(s).")
            if skipped:
                st.warning("Skipped: " + "; ".join(skipped))
        except Exception as exc:
            st.error(str(exc))

with a3:
    st.markdown('<div class="action-card"><div class="num">ACTION 03</div><h3>➤ Send Emails Directly</h3><p>Send one email per student with their generated PDF letter(s) attached, after reviewing recipients.</p></div>', unsafe_allow_html=True)
    send_ready = bool(st.session_state.results)
    if send_ready:
        if st.checkbox("I reviewed the recipient list and authorize sending these emails.",
                       key="authorize_direct_send"):
            if st.button("Send Emails Now", key="send_emails_directly", type="primary",
                         use_container_width=True):
                try:
                    with st.spinner("Sending approved emails from rvnssolutions@gmail.com..."):
                        sent, skipped, failed = send_gmail_messages(st.session_state.results)
                    st.session_state.send_report = {
                        "sent": sent, "skipped": skipped, "failed": failed
                    }
                    if sent:
                        st.success(f"Successfully sent {len(sent)} email(s).")
                    if skipped:
                        st.warning("Skipped: " + "; ".join(skipped))
                    if failed:
                        st.error(f"{len(failed)} email(s) failed. See the sending report below.")
                except Exception as exc:
                    st.error(str(exc))
        else:
            st.button("Send Emails Now", key="send_emails_disabled",
                      use_container_width=True, disabled=True)
    else:
        st.button("Send Emails Now", key="send_emails_disabled_no_results",
                  use_container_width=True, disabled=True)

with a4:
    st.markdown('<div class="action-card"><div class="num">ACTION 04</div><h3>⬇ Review & Download</h3><p>Download all generated Word and PDF letters together with the generation report.</p></div>', unsafe_allow_html=True)
    if st.session_state.results:
        package = io.BytesIO()
        with zipfile.ZipFile(package, "w", zipfile.ZIP_DEFLATED) as archive:
            for item in st.session_state.results:
                archive.write(item["docx"], item["docx"].name)
                archive.write(item["pdf"], item["pdf"].name)
            report = OUT / "generation_report.csv"
            if report.exists():
                archive.write(report, report.name)
            for extra_report in [OUT / "gmail_draft_report.csv", OUT / "gmail_send_report.csv"]:
                if extra_report.exists():
                    archive.write(extra_report, extra_report.name)
        st.download_button(
            "Download all documents (ZIP)",
            package.getvalue(),
            "RVNS_Generated_Offer_Letters.zip",
            "application/zip",
            use_container_width=True,
            type="primary",
        )
    else:
        st.button("Download all documents (ZIP)", key="download_disabled",
                  disabled=True, use_container_width=True)

# Recipient and attachment review is visible before any direct-send action.
if st.session_state.results:
    email_batches, email_skips = prepare_email_batches(st.session_state.results)
    st.markdown("## 04 · Email recipient review")
    st.warning("Check every recipient and attachment below before authorizing direct sending. 'Both' mode is grouped into one email with both PDFs attached.")
    recipient_rows = [{
        "Student": batch["student"],
        "Recipient": batch["email"],
        "PDF attachments": ", ".join(Path(p).name for p in batch["attachments"]),
    } for batch in email_batches]
    if recipient_rows:
        st.dataframe(pd.DataFrame(recipient_rows), use_container_width=True, hide_index=True)
    if email_skips:
        st.warning("These records will not be emailed: " + "; ".join(email_skips))

if st.session_state.send_report is not None:
    report = st.session_state.send_report
    st.markdown("### Direct email sending report")
    if report["sent"]:
        st.success(f"Sent: {len(report['sent'])}")
        st.dataframe(pd.DataFrame(report["sent"]), use_container_width=True, hide_index=True)
    if report["failed"]:
        st.error(f"Failed: {len(report['failed'])}")
        st.dataframe(pd.DataFrame(report["failed"]), use_container_width=True, hide_index=True)
    if report["skipped"]:
        st.warning("Skipped: " + "; ".join(report["skipped"]))
    st.caption("The sending report is also saved as output/gmail_send_report.csv.")

# Results and review
if st.session_state.results:
    st.markdown("## 04 · Generated document review")
    m1, m2, m3 = st.columns(3)
    m1.metric("Word/PDF pairs", len(st.session_state.results))
    m2.metric("Generation errors", len(st.session_state.errors))
    m3.metric("Rows with email", sum(bool((x["student"].get("email") or "").strip()) for x in st.session_state.results))
    table = [{
        "Student": x["student"].get("student_name", ""),
        "Email": x["student"].get("email", ""),
        "Mode": x["mode"],
        "Word document": x["docx"].name,
        "PDF document": x["pdf"].name,
    } for x in st.session_state.results]
    st.dataframe(pd.DataFrame(table), use_container_width=True, hide_index=True)
    with st.expander("Preview and download individual letters"):
        for index, item in enumerate(st.session_state.results, start=1):
            student_label = item["student"].get("student_name", "Student")
            st.markdown(f"**{index}. {student_label} — {item['mode']}**")
            cdoc, cpdf = st.columns(2)
            cdoc.download_button("Download Word", item["docx"].read_bytes(), item["docx"].name,
                                 "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                 key=f"word_{index}")
            cpdf.download_button("Download PDF", item["pdf"].read_bytes(), item["pdf"].name,
                                 "application/pdf", key=f"pdf_{index}")

if st.session_state.draft_report is not None:
    st.markdown("### Gmail draft report")
    if st.session_state.draft_report:
        st.dataframe(pd.DataFrame(st.session_state.draft_report), use_container_width=True, hide_index=True)
    else:
        st.info("No drafts were created. Check student email addresses.")
    st.info("Open Gmail → Drafts to review each recipient and attachment. The app does not send emails.")

if st.session_state.errors:
    st.markdown("### Generation errors")
    st.dataframe(pd.DataFrame(st.session_state.errors), use_container_width=True, hide_index=True)

st.divider()
st.caption("RVNS SOLUTIONS PRIVATE LIMITED · https://rvnssolutionspvtltd.in/ · Local document automation workspace.")
