# Approtech Internship Certificate Management System

ISO 9001:2015 Certified Certificate Generation, Issuance & Verification Portal.

---

## 🚀 Portals & Quick Access

| Portal | URL Path | Access Level | Description |
|---|---|---|---|
| **Student Application Form** | `/` or `/student-form/APP26-27` | **Public** | Student form for internship certificate requests |
| **Certificate Verification** | `/verify/<certificate_id>` | **Public** | Instant QR code scan verification & document download |
| **Admin Management Portal** | `/admin` | **Protected** | Approval, certificate issuance, batch & student management |
| **Admin Login** | `/admin/login` | **Public** | Authenticate administrator session |
| **Health Check** | `/health` | **Public** | Monitoring & keep-alive ping endpoint |

---

## 🔐 Administrative Credentials

- **Username**: `admin`
- **Password**: `admin@approtech2026`

*(Configurable via environment variables `ADMIN_USER` and `ADMIN_PASSWORD`)*

---

## 🛡️ Security Features

- **Mandatory Endpoint Authentication**: All `/api/batches` and `/api/students` administration endpoints are guarded by `@admin_required` (HTTP 401 Unauthorized for unauthenticated calls).
- **Anti-Open-Redirect Defense**: Built-in `is_safe_url()` blocks protocol-relative (`//evil.com`), backslash (`\evil.com`), and off-domain redirection attacks.
- **Brute-Force Rate Limiting**: Automatic 10-minute lockout after 5 consecutive failed login attempts (HTTP 429).
- **Hardened Session Cookies**: Configured with `HttpOnly`, `SameSite=Lax`, and 12-hour session lifetime.
- **Security Response Headers**: HTTP headers include `X-Frame-Options: SAMEORIGIN`, `X-Content-Type-Options: nosniff`, and `Referrer-Policy: strict-origin-when-cross-origin`.
- **Permanent Verification (Soft Delete)**: Deleting a student from the admin dashboard archives them from the portal view while keeping their public certificate QR verification permanently active.

---

## 💻 Local Setup & Execution

### Prerequisites
- Python 3.9+
- Dependencies listed in `requirements.txt`

### Install Dependencies
```bash
pip install -r requirements.txt
```

### Run the Server
```bash
python app.py
```
*(Default port: `3000`)*

Or specify custom host and port:
```bash
python app.py --host 0.0.0.0 --port 3000
```
