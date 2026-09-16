# Approtech Internship Certificate Management System

ISO 9001:2015 Certified Certificate Generation, Issuance & Verification Portal.

---

## Quick Access URLs

When the server is running locally:

| Portal | URL | Description |
|---|---|---|
| **Student Application Form (Default)** | **`http://localhost:3000/`** or **`/student-form/APP26-27`** | Default landing page. Public student form for certificate issuance. |
| **Admin Login & Portal** | **`http://localhost:3000/admin`** | Protected administrative portal. Requires authentication. |
| **Certificate Verification** | **`http://localhost:3000/verify/<certificate_id>`** | Public verification portal for scanned QR codes and verified certificates. |

---

## Admin Credentials

- **Username**: `admin`
- **Password**: `admin@approtech2026`

*(Can be configured via environment variables `ADMIN_USER` and `ADMIN_PASSWORD`).*

---

## How to Run Locally

### Prerequisites
- Python 3.9+
- Dependencies: Flask, python-docx, qrcode, Pillow

### Start the Application Server
```bash
python app.py
```
Or specify host and port:
```bash
python app.py --host 0.0.0.0 --port 3000
```

The system will start at:
👉 **Student Form (Direct Default)**: [http://localhost:3000/](http://localhost:3000/)  
👉 **Admin Portal (Protected)**: [http://localhost:3000/admin](http://localhost:3000/admin)
