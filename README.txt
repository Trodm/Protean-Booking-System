Protean Booking System - Backend Excel Export

Changes:
- Backend is locked for admin only.
- Admin can export to CSV.
- Admin can export to Excel (.xlsx).
- Admin login route: /admin-login
- Default admin password: admin123
- Client route: /client
- Backend route: /backend

Render:
Build Command: pip install -r requirements.txt
Start Command: uvicorn server:application --host 0.0.0.0 --port $PORT
