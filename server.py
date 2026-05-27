from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from typing import List
from pathlib import Path
from datetime import datetime
import sqlite3
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

application = FastAPI(title="Protean Booking System")

DB_FILE = "protean_bookings.db"
ADMIN_PASSWORD = "admin123"

CSS = """
<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;padding:20px;color:#111827}
.container{max-width:1700px;margin:auto;background:white;padding:22px;border-radius:12px;box-shadow:0 4px 18px rgba(0,0,0,.08)}
.login-box{max-width:480px;margin:60px auto;background:white;padding:28px;border-radius:12px;box-shadow:0 4px 18px rgba(0,0,0,.08)}
.nav{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:16px}
h1{text-align:center;margin-bottom:6px}.subtitle{text-align:center;color:#555;margin-bottom:20px}
.top-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin-bottom:20px}
.bottom-grid{display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-top:20px}
.table-wrap{overflow-x:auto}table{width:100%;border-collapse:collapse;font-size:12px;background:white}
th,td{border:1px solid #d1d5db;padding:6px;vertical-align:top}
th{background:#1f2937;color:white;position:sticky;top:0;z-index:1}
input,select,textarea{width:100%;box-sizing:border-box;padding:6px;border:1px solid #c7c7c7;border-radius:4px;font-size:12px}
textarea{min-height:70px}.badge{background:#e5e7eb;border-radius:999px;padding:4px 8px;font-size:12px}
button,.btn{background:#1f2937;color:white;padding:10px 16px;border:none;border-radius:6px;text-decoration:none;cursor:pointer;font-size:14px}
.btn-blue{background:#2563eb}.btn-green{background:#15803d}.btn-orange{background:#c2410c}.notice{background:#ecfeff;border-left:4px solid #0891b2;padding:10px;margin:12px 0}
.error{background:#fee2e2;border-left:4px solid #b91c1c;padding:10px;margin:12px 0}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:15px}
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
        reference TEXT,
        law_firm TEXT,
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
    return "".join([f"<option>{x}</option>" for x in items])

def clean_refs(raw):
    if not raw:
        return []
    output = []
    for value in raw.replace(";", ",").replace("\\n", ",").split(","):
        ref = value.strip().upper()
        if ref and ref not in output:
            output.append(ref)
    return output[:10]

def is_admin(request: Request):
    return request.query_params.get("admin_key") == ADMIN_PASSWORD

def fetch_bookings(references: str = ""):
    refs = clean_refs(references)
    conn = get_db()
    cur = conn.cursor()
    if refs:
        placeholders = ",".join(["?"] * len(refs))
        cur.execute(f"SELECT * FROM bookings WHERE reference IN ({placeholders}) ORDER BY reference, id DESC", refs)
    else:
        cur.execute("SELECT * FROM bookings ORDER BY reference, id DESC")
    rows = cur.fetchall()
    headers = [d[0] for d in cur.description]
    conn.close()
    return headers, rows

GENDERS = ["Male","Female","Other"]
LANGUAGES = ["English","IsiZulu","IsiXhosa","Sesotho","Setswana","Sepedi","Tshivenda","XiTsonga","Afrikaans","Shona","Ndebele","Other"]
OCCUPATIONS = ["Employed","Self-employed","Student/scholar","Unemployed","Minor child","Pensioner","Other"]
YES_NO = ["Yes","No"]

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
            <td><input name="mandatory_documents_submitted" placeholder="RAF 1, RAF 4, payslips, hospital records, etc."></td>
            <td><input name="injuries_sustained" placeholder="Head, spinal, fracture, etc."></td>
            <td><input type="date" name="prescribing_date"></td>
            <td><input name="protean_experts" placeholder="Expert(s) scheduled"></td>
        </tr>
        """
    return f"""
    <html><head><title>Protean Booking System</title>{CSS}</head>
    <body><div class="container">
    <div class="nav"><b>Protean Booking System <span class="badge">Client Interface</span></b></div>
    <h1>Client Booking Interface</h1>
    <p class="subtitle">Reference, Law Firm, Assessment Place and Assessment Date are captured once at the top. Claimant records are captured below.</p>
    <form action="/submit-bulk" method="post">
    <div class="top-grid">
        <div><label><b>Reference</b></label><input name="reference" placeholder="Reference Number"></div>
        <div><label><b>Law Firm</b></label><input name="law_firm" placeholder="Law Firm"></div>
        <div><label><b>Assessment Place</b></label><input name="assessment_place" placeholder="Assessment Place"></div>
        <div><label><b>Assessment Date</b></label><input type="date" name="assessment_date"></div>
    </div>
    <div class="table-wrap"><table>
    <tr>
    <th>#</th><th>Claimant Name</th><th>DOB</th><th>Gender</th><th>Preferred Language</th><th>Contact Number</th><th>Occupation Status</th><th>Type of Claim</th><th>Documents Submitted</th><th>Injuries Sustained</th><th>Prescribing Date</th><th>Protean Expert(s) the Claimant is Scheduled to See</th>
    </tr>
    {rows}
    </table></div>

    <div class="bottom-grid">
        <div>
            <label><b>Additional Information</b></label>
            <textarea name="additional_information" placeholder="Please provide any other relevant details for any of the claimant that we should be aware of."></textarea>
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

    <div class="actions"><button type="submit">Submit Booking</button></div>
    </form></div></body></html>
    """

@application.post("/submit-bulk", response_class=HTMLResponse)
async def submit_bulk(
    reference: str = Form(""),
    law_firm: str = Form(""),
    assessment_place: str = Form(""),
    assessment_date: str = Form(""),
    additional_information: str = Form(""),
    permission_to_contact: str = Form(""),
    expert_affidavits: str = Form(""),
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
            reference.strip().upper(),
            law_firm,
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
            "", "", "", "", "", "", "", "", "New", ""
        )
        cur.execute("""
        INSERT INTO bookings (
            created_at, reference, law_firm, assessment_place, assessment_date,
            claimant_name, date_of_birth, gender, preferred_language, contact_number,
            occupation_status, claim_type, mandatory_documents_submitted, injuries_sustained,
            prescribing_date, protean_experts, additional_information, permission_to_contact,
            expert_affidavits, external_experts, assigned_caller, call_attempted, date_of_call,
            contact_outcome, documents_requested, documents_expected_on_day,
            claimants_readiness_notes, booking_status, mandatory_los_documents
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, row)
        saved += 1
    conn.commit()
    conn.close()
    return f"<html><head>{CSS}</head><body><div class='container'><h1>{saved} booking(s) submitted successfully.</h1><p>Thank you. Your booking information has been received.</p><div class='actions'><a class='btn' href='/client'>Submit Another Booking</a></div></div></body></html>"

@application.get("/admin-login", response_class=HTMLResponse)
async def admin_login(error: str = ""):
    message = "<div class='error'>Invalid admin password.</div>" if error else ""
    return f"""
    <html><head><title>Admin Login</title>{CSS}</head>
    <body><div class="login-box">
    <h1>Admin Login</h1>
    <p class="subtitle">Backend access is restricted to admin users only.</p>
    {message}
    <form action="/admin-login" method="post">
        <label><b>Admin Password</b></label>
        <input type="password" name="password" required>
        <br><br>
        <button type="submit">Access Backend</button>
    </form>
    </div></body></html>
    """

@application.post("/admin-login")
async def admin_login_submit(password: str = Form("")):
    if password == ADMIN_PASSWORD:
        return RedirectResponse(f"/backend?admin_key={ADMIN_PASSWORD}", status_code=302)
    return RedirectResponse("/admin-login?error=1", status_code=302)

@application.get("/backend", response_class=HTMLResponse)
async def backend(request: Request, references: str = ""):
    if not is_admin(request):
        return RedirectResponse("/admin-login", status_code=302)

    headers, rows = fetch_bookings(references)
    header_html = "".join(f"<th>{h}</th>" for h in headers)
    body_html = ""
    for row in rows:
        body_html += "<tr>" + "".join(f"<td>{v if v is not None else ''}</td>" for v in row) + "</tr>"
    return f"""
    <html><head><title>Protean Backend</title>{CSS}</head>
    <body><div class="container">
    <div class="nav"><b>Protean Booking System <span class="badge">Admin Backend</span></b><div><a class="btn" href="/client">Client Interface</a><a class="btn btn-green" href="/export?admin_key={ADMIN_PASSWORD}&references={references}">Export CSV</a><a class="btn btn-orange" href="/export-excel?admin_key={ADMIN_PASSWORD}&references={references}">Export Excel</a></div></div>
    <h1>Backend Interface</h1>
    <div class="notice">Backend is locked. Only admin users can access booking records and export CSV or Excel.</div>
    <form action="/backend" method="get">
        <input type="hidden" name="admin_key" value="{ADMIN_PASSWORD}">
        <input name="references" placeholder="Filter by references, separated by commas" value="{references}">
        <br><br>
        <button type="submit">Load Selected References</button>
    </form><br>
    <div class="table-wrap"><table><tr>{header_html}</tr>{body_html}</table></div>
    </div></body></html>
    """

@application.get("/export")
async def export_csv(request: Request, references: str = ""):
    if not is_admin(request):
        return RedirectResponse("/admin-login", status_code=302)

    headers, rows = fetch_bookings(references)
    output = Path("protean_booking_export.csv")
    with open(output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)
    return FileResponse(output, media_type="text/csv", filename="protean_booking_export.csv")

@application.get("/export-excel")
async def export_excel(request: Request, references: str = ""):
    if not is_admin(request):
        return RedirectResponse("/admin-login", status_code=302)

    headers, rows = fetch_bookings(references)
    output = Path("protean_booking_export.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Protean Bookings"

    ws.append(headers)
    for row in rows:
        ws.append(list(row))

    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for col in ws.columns:
        max_length = 10
        col_letter = col[0].column_letter
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, min(len(value), 40))
        ws.column_dimensions[col_letter].width = max_length + 2

    ws.freeze_panes = "A2"
    wb.save(output)

    return FileResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="protean_booking_export.xlsx"
    )
