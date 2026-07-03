
from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from typing import List
from pathlib import Path
from datetime import datetime
import sqlite3, csv, os, json, urllib.request
from openpyxl import Workbook

application = FastAPI(title="Protean Booking System")

DB_FILE = "protean_bookings.db"
ADMIN_PASSWORD = "Protean123%"
TEAMS_WEBHOOK_URL = os.getenv("https://teams.microsoft.com/l/channel/19%3A34762dafb7fc427a9cf4b44d7a496a58%40thread.tacv2/Booking%20Tool?groupId=75aec17d-ec41-4041-a295-18a7f89a3c1d&tenantId=6a332dd2-de65-4fa5-92e6-dc155a4781f9", "")
LOGO_BASE64 = "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"

CSS = """
<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;padding:20px;color:#111827}
.container{max-width:1700px;margin:auto;background:white;padding:22px;border-radius:12px;box-shadow:0 4px 18px rgba(0,0,0,.08)}
.login-box{max-width:480px;margin:60px auto;background:white;padding:28px;border-radius:12px;box-shadow:0 4px 18px rgba(0,0,0,.08)}
.brand-header{text-align:center;border-bottom:1px solid #e5e7eb;padding-bottom:18px;margin-bottom:20px}
.logo{max-width:430px;width:100%;height:auto;margin:auto;display:block}
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
        payload = {"text": f"New Protean booking submitted. Law Firm: {law_firm} | Assessment Date: {assessment_date} | Claimants Saved: {saved}"}
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(TEAMS_WEBHOOK_URL, data=data, headers={"Content-Type": "application/json"}, method="POST")
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
    <html><head><title>Protean Booking System</title>{CSS}</head>
    <body><div class="container">
    <div class="brand-header">
        <img class="logo" src="data:image/png;base64,{LOGO_BASE64}" alt="Protean Medico Legal">
        <h1>Protean Booking System</h1>
        <p class="subtitle">Medico-Legal Assessment Booking Platform</p>
    </div>

    <form action="/submit-bulk" method="post">
    <div class="top-grid">
        <div><label><b>Law Firm</b></label><input name="law_firm" placeholder="Law Firm"></div>
        <div><label><b>Law Firm Contact Person</b></label><input name="law_firm_contact_person" placeholder="Contact Person"></div>
        <div><label><b>Law Firm Contact Number</b></label><input name="law_firm_phone" placeholder="Contact Number"></div>
    </div>
    <div class="top-grid">
        <div><label><b>Law Firm Email</b></label><input type="email" name="law_firm_email" placeholder="Email Address"></div>
        <div><label><b>Assessment Place</b></label><input name="assessment_place" placeholder="Assessment Place"></div>
        <div><label><b>Assessment Date</b></label><input type="date" name="assessment_date"></div>
    </div>

    <div class="table-wrap"><table>
    <tr>
    <th>#</th><th>Claimant Name</th><th>DOB</th><th>Gender</th><th>Preferred Language</th><th>Contact Number</th><th>Occupation Status</th><th>Type of Claim</th><th>Documents Submitted</th><th>Injuries Sustained</th><th>Prescribing Date</th><th>Protean Expert(s) Scheduled</th>
    </tr>{rows}</table></div>

    <div class="bottom-grid">
        <div><label><b>Additional Information</b></label><textarea name="additional_information" placeholder="Please provide any other relevant details for any claimant."></textarea></div>
        <div>
            <label><b>Do you grant us permission to contact the claimant if additional information is required?</b></label>
            <select name="permission_to_contact"><option></option>{option_tags(YES_NO)}</select><br><br>
            <label><b>Would you like expert affidavits submitted with the reports?</b></label>
            <select name="expert_affidavits"><option></option>{option_tags(YES_NO)}</select>
        </div>
    </div>

    <div class="terms">
        <h3>Terms and Conditions</h3>
        <p>By booking an assessment appointment, I acknowledge and agree to the following:</p>
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
        <div class="actions"><button type="submit">Submit Booking</button></div>
    </form></div></body></html>
    """

@application.post("/submit-bulk", response_class=HTMLResponse)
async def submit_bulk(
    law_firm: str = Form(""), law_firm_contact_person: str = Form(""), law_firm_phone: str = Form(""),
    law_firm_email: str = Form(""), assessment_place: str = Form(""), assessment_date: str = Form(""),
    additional_information: str = Form(""), permission_to_contact: str = Form(""), expert_affidavits: str = Form(""),
    terms_accepted: str = Form("Not accepted"),
    claimant_name: List[str] = Form([]), date_of_birth: List[str] = Form([]), gender: List[str] = Form([]),
    preferred_language: List[str] = Form([]), contact_number: List[str] = Form([]), occupation_status: List[str] = Form([]),
    claim_type: List[str] = Form([]), mandatory_documents_submitted: List[str] = Form([]), injuries_sustained: List[str] = Form([]),
    prescribing_date: List[str] = Form([]), protean_experts: List[str] = Form([]),
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
            law_firm, law_firm_contact_person, law_firm_phone, law_firm_email,
            assessment_place, assessment_date, cn,
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
            additional_information, permission_to_contact, expert_affidavits,
            terms_accepted, "", "", "", "", "", "", "", "New", ""
        )
        cur.execute("""
        INSERT INTO bookings (
            created_at, law_firm, law_firm_contact_person, law_firm_phone, law_firm_email,
            assessment_place, assessment_date, claimant_name, date_of_birth, gender,
            preferred_language, contact_number, occupation_status, claim_type,
            mandatory_documents_submitted, injuries_sustained, prescribing_date, protean_experts,
            additional_information, permission_to_contact, expert_affidavits, terms_accepted,
            external_experts, assigned_caller, call_attempted, date_of_call, contact_outcome,
            documents_requested, documents_expected_on_day, claimants_readiness_notes,
            booking_status, mandatory_los_documents
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, row)
        saved += 1
    conn.commit()
    conn.close()
    notify_teams(law_firm, assessment_date, saved)
    return f"<html><head>{CSS}</head><body><div class='container'><div class='brand-header'><img class='logo' src='data:image/png;base64,{LOGO_BASE64}'><h1>{saved} booking(s) submitted successfully.</h1><p>Thank you. Your booking information has been received.</p></div><div class='actions'><a class='btn' href='/client'>Submit Another Booking</a></div></div></body></html>"

@application.get("/admin-login", response_class=HTMLResponse)
async def admin_login(error: str = ""):
    message = "<div class='error'>Invalid admin password.</div>" if error else ""
    return f"<html><head><title>Admin Login</title>{CSS}</head><body><div class='login-box'><h1>Admin Login</h1><p class='subtitle'>Backend access is restricted to admin users only.</p>{message}<form action='/admin-login' method='post'><label><b>Admin Password</b></label><input type='password' name='password' required><br><br><button type='submit'>Access Backend</button></form></div></body></html>"

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
    body_html = ""
    for row in rows:
        body_html += "<tr>" + "".join(f"<td>{v if v is not None else ''}</td>" for v in row) + "</tr>"
    teams_status = "Configured" if TEAMS_WEBHOOK_URL else "Configured."
    return f"<html><head><title>Protean Backend</title>{CSS}</head><body><div class='container'><div class='brand-header'><img class='logo' src='data:image/png;base64,{LOGO_BASE64}'><h1>Admin Backend</h1></div><div class='nav'><b>Protean Booking System <span class='badge'>Admin Backend</span></b><div><a class='btn' href='/client'>Client Interface</a><a class='btn btn-green' href='/export?admin_key={ADMIN_PASSWORD}'>Export CSV</a><a class='btn btn-orange' href='/export-excel?admin_key={ADMIN_PASSWORD}'>Export Excel</a></div></div><div class='notice'>Microsoft Teams Integration: {teams_status}</div><div class='table-wrap'><table><tr>{header_html}</tr>{body_html}</table></div></div></body></html>"

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
    return FileResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="protean_booking_export.xlsx")
