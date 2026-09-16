# Approtech Internship Certificate Management System

ISO 9001:2015 Certified Certificate Generation, Issuance & Verification Portal.

---

## Quick Access URLs

When the server is running locally:

| Portal | URL | Description |
|---|---|---|
| **Admin Management Dashboard** | **`http://localhost:3000/`** or **`http://localhost:3000/admin`** | Manage student applications, approve requests, generate certificates, view QR codes, and manage batches. |
| **Student Application Form** | **`http://localhost:3000/student-form/APP26-27`** | Public student application form for certificate issuance. |
| **Certificate Verification** | **`http://localhost:3000/verify/<certificate_id>`** | Public verification portal for scanned QR codes and verified certificates. |

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
👉 **Admin Portal**: [http://localhost:3000/](http://localhost:3000/)  
👉 **Student Form**: [http://localhost:3000/student-form/APP26-27](http://localhost:3000/student-form/APP26-27)
