Protean Booking System V2.0 Complete

Included:
- Protean branding and logo retained.
- Client booking page.
- Admin backend with password.
- Up to 10 claimants per booking.
- Law Firm, Contact Person and Contact Number on first row.
- Law Firm Email, Assessment Place and Assessment Date on second row.
- Terms and Conditions with required acceptance.
- Document upload disabled.
- Microsoft Teams webhook support via TEAMS_WEBHOOK_URL environment variable.
- Excel and CSV exports.
- 365-day retention policy.
- No delete function before 365 days.

Render Build Command:
pip install -r requirements.txt

Render Start Command:
uvicorn server:application --host 0.0.0.0 --port $PORT

Default Admin Password: admin123

Optional Render Environment Variables:
ADMIN_PASSWORD = your secure password
TEAMS_WEBHOOK_URL = your Microsoft Teams incoming webhook URL
