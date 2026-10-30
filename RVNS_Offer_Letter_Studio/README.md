# RVNS Solutions — Offer Letter Studio

A local Windows application for importing student Excel data, validating records, generating personalized two-page Word/PDF internship offer letters, reviewing documents, and creating Gmail drafts with matching PDF attachments.

**Email safety:** this application creates Gmail drafts only. It does not send email. Review and send drafts manually from Gmail.

## Project contents

- `app.py` — Streamlit dashboard with three side-by-side workflow actions.
- `rvns_offer_automation.py` — Excel reading, mode rules, document personalization, PDF conversion, Gmail draft creation.
- `templates/offline_template.docx` — RVNS Offline and General letter source template.
- `templates/online_template.docx` — RVNS Online letter source template.
- `assets/rvns_authorized_signature_seal.png` — supplied authorized signature/seal image.
- `students_sample.xlsx` — sample workbook with six example records.
- `install.bat` — creates the Python environment and installs packages.
- `run_dashboard.bat` — starts the local dashboard.
- `run_windows.bat` — installs missing dependencies and starts the dashboard.
- `output/` — generated Word/PDF letters and reports.

## Requirements

- Windows 10 or Windows 11.
- Python 3.10 or newer (Python 3.11/3.12 recommended).
- Microsoft Word installed for DOCX-to-PDF conversion, or LibreOffice installed and available in PATH.
- Internet connection for initial package installation and Gmail authorization.

## Installation and launch

1. Extract the complete ZIP file to a normal local folder, for example `C:\RVNS\OfferLetterStudio`.
2. Double-click `install.bat` once and wait until installation completes.
3. Double-click `run_dashboard.bat`.
4. Keep the terminal window open. The dashboard normally opens at `http://localhost:8501`.
5. If the browser does not open automatically, copy the local URL printed in the terminal into your browser.

Do not run the dashboard from a public/shared computer. Do not expose Streamlit to the public internet.

## Dashboard workflow

All three action cards are shown side by side:

1. **Create Offer Letters** — generate individualized `.docx` and `.pdf` files.
2. **Create Gmail Drafts** — create drafts with the matching PDF attached. Nothing is sent.
3. **Review & Download** — download all generated letters and the report as one ZIP, or download individual Word/PDF files.

Upload the Excel workbook, select its worksheet, inspect the data preview, and select the mode rule before generating. Fix rows with missing student names before issuing letters. Review all generated documents before using them.

## Excel columns

The required column is Student Name. Common header variations are recognized.

| Field | Required | Purpose |
|---|---|---|
| `Student_Name` | Yes | Student's full name |
| `Email` | For Gmail drafts | Student email address |
| `College_Name` | No | College/institution |
| `Branch` | No | Department/branch |
| `Internship_Domain` | No | Internship domain/title |
| `Internship_Mode` | No | Online, Offline, Both, or General |
| `Start_Date` | No | Commencement date |
| `End_Date` | No | Completion date |
| `Duration` | No | Duration exactly as supplied |
| `Office_Location` | No | Offline internship location |
| `Offer_Date` | No | Letter issue date; defaults to the current date if blank |

The application recognizes common alternatives such as `Name`, `Full_Name`, `Email_ID`, `College`, `Department`, `Course`, `Mode`, `Commencement_Date`, `Completion_Date`, and `Internship_Duration`.

Excel date cells are supported. Text dates in formats such as `2026-11-01`, `01-11-2026`, `01/11/2026`, or `01 November 2026` are recognized. Unrecognized date text is preserved instead of being guessed.

## Internship mode rules

- `Offline` — creates one Offline letter.
- `Online` — creates one Online letter.
- `Both` — creates two separate letters: one Offline and one Online.
- `General`, blank, or an unrecognized mode — creates a General letter without Online/Offline wording.
- The dashboard can use the value from each Excel row or apply Offline, Online, Both, or General to all students.

When `From Excel` is selected, each student's row controls the mode.

## Dynamic date and optional-field rules

- Start date only: show commencement date only.
- End date only: show completion date only.
- Both dates: show both dates.
- Duration only: show duration only.
- Start date plus duration: show start date and supplied duration; do not calculate an end date.
- No dates or duration: omit those detail lines.
- Blank college, branch, duration, start date, end date, and office location fields are omitted.
- No `None`, `NaN`, `N/A`, or invented student details are inserted.
- Offer Date is the issue date. If the Excel cell is blank, the system uses the current date.

## Signature and page layout

The supplied RVNS signature/seal image is inserted above the Managing Director designation. The company letterhead/header and orange footer are retained. The sample workbook was tested and generated two-page PDFs for all six example records. Different names or unusually long text may affect pagination, so review every generated PDF before issuing it.

The supplied signature image is 196 × 135 pixels. A higher-resolution transparent PNG is recommended for sharper print quality. The application uses only the supplied image; it does not fabricate a signature.

## Gmail setup — drafts only

1. Open Google Cloud Console and create/select a project.
2. Enable the Gmail API.
3. Configure the OAuth consent screen for the company account.
4. Create an OAuth Client ID with application type **Desktop app**.
5. Download the client JSON and save it in the project root as `credentials.json`.
6. Launch the dashboard and click **Create Gmail Drafts**.
7. Complete Google's authorization in the browser using `rvnssolutions@gmail.com`.
8. The app verifies the connected Gmail account before creating drafts.
9. Open Gmail → Drafts. Check every recipient, subject, message, and PDF attachment. Send manually only after review.

The OAuth scope is `gmail.modify`, which is required to create Gmail drafts. If an old token was authorized with a different scope, close the dashboard, delete the local `token.json`, and authorize again.

**Never share or upload `credentials.json` or `token.json`.** They grant access to the connected Gmail account. They are excluded from the project ZIP and ignored by Git. If exposed, revoke the OAuth grant in your Google account and create fresh credentials.

## Output files

Generated documents and reports are saved in `output/`:

- Individual Word documents (`.docx`).
- Individual PDF documents (`.pdf`).
- `generation_report.csv`.
- `gmail_draft_report.csv` after Gmail draft creation.

Use the dashboard's download action to download all generated Word/PDF files and the report as a ZIP.

## Troubleshooting

**`py` is not recognized:** install Python from python.org, enable “Add Python to PATH”, then reopen the terminal.

**PDF conversion fails:** install Microsoft Word or LibreOffice. Close any open/locked generated PDF and retry.

**Gmail authentication fails:** check that Gmail API is enabled, OAuth client type is Desktop app, the correct company account is authorized, and the local token is valid.

**Student Name column not detected:** rename the header to `Student_Name`, `Name`, or `Full_Name`.

**Some rows are skipped:** inspect the dashboard error table and fill in the missing Student Name.

**Dashboard does not open:** keep the terminal window open and use the local address printed there. If port 8501 is already occupied, close the other Streamlit process and restart.

## Final review checklist

- Verify every student's name, email, college, branch, domain, mode, dates, duration, and location.
- Confirm the selected mode matches the Excel data.
- Open each PDF and check it is exactly two pages, with the signature visible and not overlapping text/footer.
- Check the recipient and PDF attachment in every Gmail draft.
- Send manually from Gmail only after the final review.
