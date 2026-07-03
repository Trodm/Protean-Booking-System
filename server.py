from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from typing import List
from pathlib import Path
from datetime import datetime
import sqlite3
import csv
import os
import json
import urllib.request
from openpyxl import Workbook

application = FastAPI(title="Protean Booking System")

application.mount("/static", StaticFiles(directory="static"), name="static")

DB_FILE = "protean_bookings.db"
ADMIN_PASSWORD = "Protean123%"
RETENTION_DAYS = 365
TEAMS_WEBHOOK_URL = os.getenv("https://teams.microsoft.com/l/channel/19%3A34762dafb7fc427a9cf4b44d7a496a58%40thread.tacv2/Booking%20Tool?groupId=75aec17d-ec41-4041-a295-18a7f89a3c1d&tenantId=6a332dd2-de65-4fa5-92e6-dc155a4781f9", "")

CSS = """
<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;padding:20px;color:#111827}
.container{max-width:1700px;margin:auto;background:white;padding:22px;border-radius:12px;box-shadow:0 4px 18px rgba(0,0,0,.08)}
.login-box{max-width:480px;margin:60px auto;background:white;padding:28px;border-radius:12px;box-shadow:0 4px 18px rgba(0,0,0,.08)}
.brand-header{text-align:center;margin-bottom:25px}
.logo{width:360px;max-width:95%;margin:auto;display:block;margin-bottom:20px}
.nav{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:16px}
h1{text-align:center;margin-bottom:6px}.subtitle{text-align:center;color:#555;margin-bottom:20px}
.top-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin-bottom:20px}
.bottom-grid{display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-top:20px}
.table-wrap{overflow-x:auto}table{width:100%;border-collapse:collapse;font-size:12px;background:white}
th,td{border:1px solid #d1d5db;padding:6px;vertical-align:top}
th{background:#1f2937;color:white;position:sticky;top:0;z-index:1}
input,select,textarea{width:100%;box-sizing:border-box;padding:6px;border:1px solid #c7c7c7;border-radius:4px;font-size:12px}
textarea{min-height:70px}.badge{background:#e5e7eb;border-radius:999px;padding:4px 8px;font-size:12px}
button,.btn{background:#1f2937;color:white;padding:10px 16px;border:none;border-radius:6px;text-decoration:none;cursor:pointer;font-size:14px;display:inline-block}
.btn-green{background:#15803d}.btn-orange{background:#c2410c}.notice{background:#ecfeff;border-left:4px solid #0891b2;padding:10px;margin:12px 0}
.terms{background:#f9fafb;border:1px solid #d1d5db;border-radius:8px;padding:14px;margin-top:18px;line-height:1.5}
.error{background:#fee2e2;border-left:4px solid #b91c1c;padding:10px;margin:12px 0}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:15px}
.check-row{display:flex;gap:10px;align-items:flex-start;background:#fff7ed;border-left:4px solid #f97316;padding:12px;margin-top:15px}
.check-row input{width:auto;margin-top:3px}
@media(max-width:900px){.top-grid,.bottom-grid{grid-template-columns:1fr}}
</style>
"""

def retention_message():
    return f"Retention Policy: submitted bookings must be retained for at least {RETENTION_DAYS} days. The admin backend has no delete button before the 365-day retention period has passed."

def get_db():
    return sqlite3.connect(DB_FILE)

def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS bookings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        created_at TEXT,
        law_firm TEXT,
        law_firm_contact_person TEXT,
        law_firm_phone TEXT,
        law_firm_email TEXT,
        assessment_place TEXT,
        assessment_date TEXT,
        claimant_name TEXT,
        date_of_birth TEXT,
        gender TEXT,
        preferred_language TEXT,
        contact_number TEXT,
        occupation_status TEXT,
        claim_type TEXT,
        mandatory_documents_submitted TEXT,
        injuries_sustained TEXT,
        prescribing_date TEXT,
        protean_experts TEXT,
        additional_information TEXT,
        permission_to_contact TEXT,
        expert_affidavits TEXT,
        terms_accepted TEXT,
        retention_policy TEXT,
        external_experts TEXT,
        assigned_caller TEXT,
        call_attempted TEXT,
        date_of_call TEXT,
        contact_outcome TEXT,
        documents_requested TEXT,
        documents_expected_on_day TEXT,
        claimants_readiness_notes TEXT,
        booking_status TEXT,
        mandatory_los_documents TEXT
    )
    """)
    conn.commit()
    conn.close()

init_db()

def option_tags(items):
    return "".join(f"<option>{x}</option>" for x in items)

def is_admin(request: Request):
    return request.query_params.get("admin_key") == ADMIN_PASSWORD

def fetch_bookings():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings ORDER BY id DESC")
    rows = cur.fetchall()
    headers = [d[0] for d in cur.description]
    conn.close()
    return headers, rows

def notify_teams(law_firm, assessment_date, saved):
    if not TEAMS_WEBHOOK_URL:
        return
    try:
        payload = {
            "text": f"New Protean booking submitted. Law Firm: {law_firm} | Assessment Date: {assessment_date} | Claimants Saved: {saved}"
        }
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            TEAMS_WEBHOOK_URL,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        urllib.request.urlopen(req, timeout=5)
    except Exception:
        pass

GENDERS = ["Male", "Female", "Other"]
LANGUAGES = ["English", "IsiZulu", "IsiXhosa", "Sesotho", "Setswana", "Sepedi", "Tshivenda", "XiTsonga", "Afrikaans", "Shona", "Ndebele", "Other"]
OCCUPATIONS = ["Employed", "Self-employed", "Student/scholar", "Unemployed", "Minor child", "Pensioner", "Other"]
YES_NO = ["Yes", "No"]

@application.get("/")
async def root():
    return RedirectResponse("/client", status_code=302)

@application.get("/health")
async def health():
    return {"status": "ok"}

@application.get("/client", response_class=HTMLResponse)
async def client_interface():
    rows = ""
    for i in range(1, 11):
        rows += f"""
        <tr>
            <td>{i}</td>
            <td><input name="claimant_name" placeholder="Claimant Name"></td>
            <td><input type="date" name="date_of_birth"></td>
            <td><select name="gender"><option></option>{option_tags(GENDERS)}</select></td>
            <td><select name="preferred_language"><option></option>{option_tags(LANGUAGES)}</select></td>
            <td><input name="contact_number" placeholder="Phone / WhatsApp"></td>
            <td><select name="occupation_status"><option></option>{option_tags(OCCUPATIONS)}</select></td>
            <td><input name="claim_type" placeholder="LOS, LOE, Medical negligence, etc."></td>
            <td><input name="mandatory_documents_submitted" placeholder="LOI, ID, Hospital records, RAF 1, RAF 4, etc."></td>
            <td><input name="injuries_sustained" placeholder="Head, spinal, fracture, etc."></td>
            <td><input type="date" name="prescribing_date"></td>
            <td><input name="protean_experts" placeholder="Expert(s) scheduled"></td>
        </tr>
        """

    return f"""
    <html>
    <head>
        <title>Protean Booking System</title>
        {CSS}
    </head>

    <body>
    <div class="container">

        <div class="brand-header">
            <img src="/static/logo.png" alt="Protean Medico Legal" class="logo">
            <h1>Protean Booking System</h1>
            <p class="subtitle">Medico-Legal Assessment Booking Platform</p>
        </div>

        <form action="/submit-bulk" method="post">

            <div class="top-grid">
                <div>
                    <label><b>Law Firm</b></label>
                    <input name="law_firm" placeholder="Law Firm">
                </div>

                <div>
                    <label><b>Law Firm Contact Person</b></label>
                    <input name="law_firm_contact_person" placeholder="Contact Person">
                </div>

                <div>
                    <label><b>Law Firm Contact Number</b></label>
                    <input name="law_firm_phone" placeholder="Contact Number">
                </div>
            </div>

            <div class="top-grid">
                <div>
                    <label><b>Law Firm Email</b></label>
                    <input type="email" name="law_firm_email" placeholder="Email Address">
                </div>

                <div>
                    <label><b>Assessment Place</b></label>
                    <input name="assessment_place" placeholder="Assessment Place">
                </div>

                <div>
                    <label><b>Assessment Date</b></label>
                    <input type="date" name="assessment_date">
                </div>
            </div>

            <div class="table-wrap">
                <table>
                    <tr>
                        <th>#</th>
                        <th>Claimant Name</th>
                        <th>DOB</th>
                        <th>Gender</th>
                        <th>Preferred Language</th>
                        <th>Contact Number</th>
                        <th>Occupation Status</th>
                        <th>Type of Claim</th>
                        <th>Documents Submitted</th>
                        <th>Injuries Sustained</th>
                        <th>Prescribing Date</th>
                        <th>Protean Expert(s) Scheduled</th>
                    </tr>
                    {rows}
                </table>
            </div>

            <div class="bottom-grid">
                <div>
                    <label><b>Additional Information</b></label>
                    <textarea name="additional_information" placeholder="Please provide any other relevant details for any claimant."></textarea>
                </div>

                <div>
                    <label><b>Do you grant us permission to contact the claimant if additional information is required?</b></label>
                    <select name="permission_to_contact">
                        <option></option>
                        {option_tags(YES_NO)}
                    </select>

                    <br><br>

                    <label><b>Would you like expert affidavits submitted with the reports?</b></label>
                    <select name="expert_affidavits">
                        <option></option>
                        {option_tags(YES_NO)}
                    </select>
                </div>
            </div>

            <div class="terms">
                <h3>Terms and Conditions</h3>
                <ul>
                    <li>All mandatory documents, including the Letter of Instruction, ID document, hospital records, and RAF 1 and RAF 4 forms, must be submitted using the same email address through which the booking link was sent.</li>
                    <li>If the required documents cannot be submitted electronically before the assessment, the original documents or copies thereof must be brought on the day of the assessment.</li>
                    <li>It remains the responsibility of the attorney to ensure that all supporting documentation required to finalize the report is submitted timeously and accurately.</li>
                </ul>
            </div>

            <label class="check-row">
                <input type="checkbox" name="terms_accepted" value="Accepted" required>
                <span>I have read and accepted the Terms and Conditions.</span>
            </label>

            <div class="actions">
                <button type="submit">Submit Booking</button>
            </div>

        </form>
    </div>
    </body>
    </html>
    """

@application.post("/submit-bulk", response_class=HTMLResponse)
async def submit_bulk(
    law_firm: str = Form(""),
    law_firm_contact_person: str = Form(""),
    law_firm_phone: str = Form(""),
    law_firm_email: str = Form(""),
    assessment_place: str = Form(""),
    assessment_date: str = Form(""),
    additional_information: str = Form(""),
    permission_to_contact: str = Form(""),
    expert_affidavits: str = Form(""),
    terms_accepted: str = Form("Not accepted"),
    claimant_name: List[str] = Form([]),
    date_of_birth: List[str] = Form([]),
    gender: List[str] = Form([]),
    preferred_language: List[str] = Form([]),
    contact_number: List[str] = Form([]),
    occupation_status: List[str] = Form([]),
    claim_type: List[str] = Form([]),
    mandatory_documents_submitted: List[str] = Form([]),
    injuries_sustained: List[str] = Form([]),
    prescribing_date: List[str] = Form([]),
    protean_experts: List[str] = Form([]),
):
    conn = get_db()
    cur = conn.cursor()
    saved = 0

    for i in range(10):
        cn = (claimant_name[i] if i < len(claimant_name) else "").strip()
        if not cn:
            continue

        row = (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            law_firm,
            law_firm_contact_person,
            law_firm_phone,
            law_firm_email,
            assessment_place,
            assessment_date,
            cn,
            date_of_birth[i] if i < len(date_of_birth) else "",
            gender[i] if i < len(gender) else "",
            preferred_language[i] if i < len(preferred_language) else "",
            contact_number[i] if i < len(contact_number) else "",
            occupation_status[i] if i < len(occupation_status) else "",
            claim_type[i] if i < len(claim_type) else "",
            mandatory_documents_submitted[i] if i < len(mandatory_documents_submitted) else "",
            injuries_sustained[i] if i < len(injuries_sustained) else "",
            prescribing_date[i] if i < len(prescribing_date) else "",
            protean_experts[i] if i < len(protean_experts) else "",
            additional_information,
            permission_to_contact,
            expert_affidavits,
            terms_accepted,
            retention_message(),
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "New",
            ""
        )

        cur.execute("""
        INSERT INTO bookings (
            created_at,
            law_firm,
            law_firm_contact_person,
            law_firm_phone,
            law_firm_email,
            assessment_place,
            assessment_date,
            claimant_name,
            date_of_birth,
            gender,
            preferred_language,
            contact_number,
            occupation_status,
            claim_type,
            mandatory_documents_submitted,
            injuries_sustained,
            prescribing_date,
            protean_experts,
            additional_information,
            permission_to_contact,
            expert_affidavits,
            terms_accepted,
            retention_policy,
            external_experts,
            assigned_caller,
            call_attempted,
            date_of_call,
            contact_outcome,
            documents_requested,
            documents_expected_on_day,
            claimants_readiness_notes,
            booking_status,
            mandatory_los_documents
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, row)

        saved += 1

    conn.commit()
    conn.close()

    notify_teams(law_firm, assessment_date, saved)

    return f"""
    <html>
    <head>{CSS}</head>
    <body>
    <div class="container">
        <div class="brand-header">
            <img src="/static/logo.png" alt="Protean Medico Legal" class="logo">
            <h1>{saved} booking(s) submitted successfully.</h1>
            <p>Thank you. Your booking information has been received.</p>
        </div>

        <div class="actions">
            <a class="btn" href="/client">Submit Another Booking</a>
        </div>
    </div>
    </body>
    </html>
    """

@application.get("/admin-login", response_class=HTMLResponse)
async def admin_login(error: str = ""):
    message = "<div class='error'>Invalid admin password.</div>" if error else ""

    return f"""
    <html>
    <head>
        <title>Admin Login</title>
        {CSS}
    </head>

    <body>
    <div class="login-box">
        <div class="brand-header">
            <img src="/static/logo.png" alt="Protean Medico Legal" class="logo">
        </div>

        <h1>Admin Login</h1>

        {message}

        <form action="/admin-login" method="post">
            <label><b>Admin Password</b></label>
            <input type="password" name="password" required>

            <br><br>

            <button type="submit">Access Backend</button>
        </form>
    </div>
    </body>
    </html>
    """

@application.post("/admin-login")
async def admin_login_submit(password: str = Form("")):
    if password == ADMIN_PASSWORD:
        return RedirectResponse(f"/backend?admin_key={ADMIN_PASSWORD}", status_code=302)

    return RedirectResponse("/admin-login?error=1", status_code=302)

@application.get("/backend", response_class=HTMLResponse)
async def backend(request: Request):
    if not is_admin(request):
        return RedirectResponse("/admin-login", status_code=302)

    headers, rows = fetch_bookings()

    header_html = "".join(f"<th>{h}</th>" for h in headers)

    body_html = "".join(
        "<tr>" + "".join(f"<td>{v if v is not None else ''}</td>" for v in row) + "</tr>"
        for row in rows
    )

    teams_status = "Configured" if TEAMS_WEBHOOK_URL else "Not configured. Add TEAMS_WEBHOOK_URL in Render environment variables."

    return f"""
    <html>
    <head>
        <title>Protean Backend</title>
        {CSS}
    </head>

    <body>
    <div class="container">

        <div class="brand-header">
            <img src="/static/logo.png" alt="Protean Medico Legal" class="logo">
        </div>

        <div class="nav">
            <b>Protean Booking System <span class="badge">Admin Backend</span></b>

            <div>
                <a class="btn" href="/client">Client Interface</a>
                <a class="btn btn-green" href="/export?admin_key={ADMIN_PASSWORD}">Export CSV</a>
                <a class="btn btn-orange" href="/export-excel?admin_key={ADMIN_PASSWORD}">Export Excel</a>
            </div>
        </div>

        <h1>Admin Backend</h1>

        <div class="notice">
            Microsoft Teams Integration: {teams_status}
            <br><br>
            {retention_message()}
        </div>

        <div class="table-wrap">
            <table>
                <tr>{header_html}</tr>
                {body_html}
            </table>
        </div>
    </div>
    </body>
    </html>
    """

@application.get("/export")
async def export_csv(request: Request):
    if not is_admin(request):
        return RedirectResponse("/admin-login", status_code=302)

    headers, rows = fetch_bookings()
    output = Path("protean_booking_export.csv")

    with open(output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

    return FileResponse(output, media_type="text/csv", filename="protean_booking_export.csv")

@application.get("/export-excel")
async def export_excel(request: Request):
    if not is_admin(request):
        return RedirectResponse("/admin-login", status_code=302)

    headers, rows = fetch_bookings()
    output = Path("protean_booking_export.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Protean Bookings"

    ws.append(headers)

    for row in rows:
        ws.append(list(row))

    wb.save(output)

    return FileResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="protean_booking_export.xlsx"
    )
