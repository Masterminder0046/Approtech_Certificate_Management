#!/usr/bin/env python3
"""
Approtech Internship Certificate Management System
Backend: Python 3 + Flask + native SQLite3
"""

import os
import io
import sys
import random
import sqlite3
import argparse
from datetime import datetime, timezone
from flask import (
    Flask,
    request,
    jsonify,
    render_template,
    send_from_directory,
    send_file,
    redirect,
    abort
)
import qrcode
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'database.db')

app = Flask(
    __name__,
    static_folder='.',
    static_url_path='',
    template_folder='templates'
)
app.config['TEMPLATES_AUTO_RELOAD'] = True

# ---------------------------------------------------------------------------
# DATABASE HELPERS
# ---------------------------------------------------------------------------

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initializes the database without dropping existing tables or data."""
    conn = get_db()
    cur = conn.cursor()

    # Batches table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_code TEXT UNIQUE NOT NULL,
        batch_name TEXT NOT NULL,
        leader TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        status TEXT DEFAULT 'Active',
        email_enabled INTEGER DEFAULT 1
    )
    ''')

    # Students table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id INTEGER NOT NULL REFERENCES batches(id),
        full_name TEXT NOT NULL,
        register_number TEXT NOT NULL,
        college_name TEXT NOT NULL,
        degree_branch TEXT NOT NULL,
        state TEXT,
        email TEXT NOT NULL,
        domain TEXT NOT NULL,
        mode TEXT NOT NULL,
        internship_start_date TEXT NOT NULL,
        internship_end_date TEXT NOT NULL,
        project_title TEXT NOT NULL,
        phone_number TEXT NOT NULL,
        status TEXT DEFAULT 'Pending',
        certificate_id TEXT UNIQUE,
        certificate_generated INTEGER DEFAULT 0,
        submitted_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        approved_at TIMESTAMP,
        archived INTEGER DEFAULT 0
    )
    ''')

    # Ensure default Approtech batch exists (check before inserting to prevent duplicate)
    cur.execute("SELECT id FROM batches WHERE batch_code = ?", ('APP26-27',))
    row = cur.fetchone()
    if not row:
        cur.execute('''
            INSERT INTO batches (batch_code, batch_name, leader, created_at, status, email_enabled)
            VALUES (?, ?, ?, datetime('now'), 'Active', 1)
        ''', ('APP26-27', 'Approtech', 'Admin Leader'))

    conn.commit()
    conn.close()

# ---------------------------------------------------------------------------
# DURATION & CERTIFICATE ID UTILITIES
# ---------------------------------------------------------------------------

def calculate_duration(start_str, end_str):
    """Calculates duration in days and weeks from start and end dates."""
    if not start_str or not end_str:
        return "N/A"
    try:
        d1 = datetime.strptime(start_str.strip(), "%Y-%m-%d")
        d2 = datetime.strptime(end_str.strip(), "%Y-%m-%d")
        diff_days = (d2 - d1).days + 1
        if diff_days <= 0:
            return "0 Days"
        weeks = round(diff_days / 7)
        if weeks >= 1:
            return f"{diff_days} Days (~ {weeks} {'week' if weeks == 1 else 'weeks'})"
        return f"{diff_days} {'Day' if diff_days == 1 else 'Days'}"
    except Exception:
        return "N/A"

def generate_unique_certificate_id(batch_code, conn):
    """Generates INT:{BATCH_CODE}/{4-DIGIT}-{4-DIGIT} ensuring uniqueness."""
    cur = conn.cursor()
    # Sanitize batch code for cert ID format (strip spaces, uppercase)
    clean_code = batch_code.strip().upper()
    while True:
        p1 = random.randint(1000, 9999)
        p2 = random.randint(1000, 9999)
        candidate = f"INT:{clean_code}/{p1}-{p2}"
        cur.execute("SELECT id FROM students WHERE certificate_id = ?", (candidate,))
        if not cur.fetchone():
            return candidate

def format_certificate_date(date_str):
    """Converts YYYY-MM-DD into 10th August 2026 matching reference PDF format."""
    if not date_str:
        return ""
    try:
        dt = datetime.strptime(str(date_str).strip(), "%Y-%m-%d")
        day = dt.day
        if 11 <= day <= 13:
            suffix = "th"
        else:
            suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
        return f"{day}{suffix} {dt.strftime('%B %Y')}"
    except Exception:
        return str(date_str)

def generate_certificate_docx(student, base_url):
    """Generates a genuine .docx certificate matching the uploaded PDF reference."""
    doc = Document()

    # Page setup: A4 Portrait
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Certificate Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(18)
    r_title = p_title.add_run("CERTIFICATE OF COMPLETION")
    r_title.bold = True
    r_title.font.name = "Times New Roman"
    r_title.font.size = Pt(14)
    r_title.font.color.rgb = RGBColor(0, 0, 0)

    # Certificate ID
    p_id = doc.add_paragraph()
    p_id.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_id.paragraph_format.space_before = Pt(0)
    p_id.paragraph_format.space_after = Pt(20)
    r_id = p_id.add_run(student.get('certificate_id') or 'INT:APP26-27/0000-0000')
    r_id.bold = True
    r_id.font.name = "Times New Roman"
    r_id.font.size = Pt(13.5)

    # Subheading
    p_concern = doc.add_paragraph()
    p_concern.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_concern.paragraph_format.space_before = Pt(0)
    p_concern.paragraph_format.space_after = Pt(18)
    r_concern = p_concern.add_run("TO WHOMSOEVER IT MAY CONCERN")
    r_concern.bold = True
    r_concern.underline = True
    r_concern.font.name = "Times New Roman"
    r_concern.font.size = Pt(12)

    def add_run(p, text, bold=False):
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(13)
        r.bold = bold
        return r

    # Paragraph 1
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(12)
    p1.paragraph_format.line_spacing = 1.35

    add_run(p1, "This is to certify that ")
    add_run(p1, f"{student.get('full_name')}, Reg. No: {student.get('register_number')}", bold=True)
    add_run(p1, ", a student of ")
    add_run(p1, f"{student.get('college_name')}", bold=True)
    add_run(p1, " , pursuing ")
    add_run(p1, f"{student.get('degree_branch')}", bold=True)
    add_run(p1, ", has successfully completed an Internship Program Through ")
    mode_text = str(student.get('mode', 'Online')).strip().upper()
    add_run(p1, mode_text, bold=True)
    add_run(p1, " at our organization in the domain of ")
    add_run(p1, f"{student.get('domain')}", bold=True)
    add_run(p1, " ,")

    # Paragraph 2
    start_fmt = format_certificate_date(student.get('internship_start_date'))
    end_fmt = format_certificate_date(student.get('internship_end_date'))
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(12)
    p2.paragraph_format.line_spacing = 1.35
    add_run(p2, "The internship was undertaken from ")
    add_run(p2, f"{start_fmt} to {end_fmt}.", bold=True)

    # Paragraph 3
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(12)
    p3.paragraph_format.line_spacing = 1.35
    add_run(p3, "During the course of the internship, the student exhibited commendable professional behaviour and technical proficiency, particularly in the project titled ")
    add_run(p3, f"“{student.get('project_title')}”", bold=True)
    add_run(p3, ".")

    # Paragraph 4
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p4.paragraph_format.space_before = Pt(0)
    p4.paragraph_format.space_after = Pt(28)
    p4.paragraph_format.line_spacing = 1.35
    add_run(p4, "We extend our best wishes to continue success in all future endeavors.")

    # Generate QR Code image
    cert_id = student.get('certificate_id') or 'PENDING'
    verification_url = f"{base_url.rstrip('/')}/verify/{cert_id}"
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=6,
        border=1,
    )
    qr.add_data(verification_url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="#0F172A", back_color="#FFFFFF")
    qr_buf = io.BytesIO()
    qr_img.save(qr_buf, format='PNG')
    qr_buf.seek(0)

    # Table for Footer (QR code left, Authorized Signature right)
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    table.columns[0].width = Inches(2.8)
    table.columns[1].width = Inches(3.47)

    # Remove cell borders
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._element.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    # Left cell: QR
    cell_qr = table.cell(0, 0)
    p_qr = cell_qr.paragraphs[0]
    p_qr.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_qr.paragraph_format.space_before = Pt(0)
    p_qr.paragraph_format.space_after = Pt(0)
    run_qr = p_qr.add_run()
    run_qr.add_picture(qr_buf, width=Inches(1.25), height=Inches(1.25))

    # Right cell: Signatures
    cell_sig = table.cell(0, 1)
    p_for = cell_sig.paragraphs[0]
    p_for.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_for.paragraph_format.space_before = Pt(0)
    p_for.paragraph_format.space_after = Pt(45)
    r_for = p_for.add_run("For Approtech R&D Solutions Pvt. Ltd.,")
    r_for.font.name = "Times New Roman"
    r_for.font.size = Pt(13)

    p_sign = cell_sig.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.paragraph_format.space_before = Pt(0)
    p_sign.paragraph_format.space_after = Pt(0)
    r_sign = p_sign.add_run("Authorized Signature")
    r_sign.bold = True
    r_sign.font.name = "Times New Roman"
    r_sign.font.size = Pt(13)

    target_buf = io.BytesIO()
    doc.save(target_buf)
    target_buf.seek(0)
    return target_buf

# ---------------------------------------------------------------------------
# STATIC FILE SERVING
# ---------------------------------------------------------------------------

@app.route('/style.css')
def serve_style():
    return send_from_directory('.', 'style.css')

@app.route('/script.js')
def serve_script():
    return send_from_directory('.', 'script.js')

@app.route('/assets/<path:path>')
def serve_assets(path):
    return send_from_directory('assets', path)

@app.route('/public/<path:path>')
def serve_public(path):
    return send_from_directory('public', path)

# ---------------------------------------------------------------------------
# PUBLIC / PAGE ROUTES
# ---------------------------------------------------------------------------

@app.route('/')
@app.route('/admin')
@app.route('/admin/')
def admin_dashboard():
    """Renders the Admin Management Dashboard."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM batches ORDER BY id ASC")
    batches = [dict(b) for b in cur.fetchall()]
    conn.close()
    return render_template('admin.html', batches=batches)

@app.route('/index.html')
def index_html():
    return redirect('/student-form/APP26-27')

@app.route('/student-form')
@app.route('/student-form/')
def student_form_redirect():
    """Redirects to default Approtech batch form."""
    return redirect('/student-form/APP26-27')

@app.route('/student-form/<batch_code>')
def student_form(batch_code):
    """Renders the Student Internship Certificate Form pre-bound to batch_code."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM batches WHERE batch_code = ?", (batch_code.strip(),))
    batch = cur.fetchone()
    if not batch:
        # If batch doesn't exist, create it or default to APP26-27
        cur.execute("SELECT * FROM batches WHERE batch_code = 'APP26-27'")
        batch = cur.fetchone()
    batch_dict = dict(batch) if batch else {'batch_code': 'APP26-27', 'batch_name': 'Approtech'}
    conn.close()
    return render_template('student_form.html', batch=batch_dict)

@app.route('/verify/<path:certificate_id>')
def verify_certificate(certificate_id):
    """Public Certificate Verification Page - NO login required."""
    conn = get_db()
    cur = conn.cursor()

    cur.execute('''
        SELECT s.*, b.batch_code, b.batch_name
        FROM students s
        JOIN batches b ON s.batch_id = b.id
        WHERE s.certificate_id = ?
    ''', (certificate_id.strip(),))
    row = cur.fetchone()
    conn.close()

    if not row:
        # Invalid Certificate
        return render_template(
            'verification.html',
            status='invalid',
            certificate_id=certificate_id
        ), 404

    student = dict(row)

    # Check if archived
    if student.get('archived') == 1:
        return render_template(
            'verification.html',
            status='archived',
            certificate_id=certificate_id,
            student=student
        )

    # Check if not yet generated
    if student.get('certificate_generated') != 1:
        return render_template(
            'verification.html',
            status='pending',
            certificate_id=certificate_id,
            student=student
        )

    # Valid Verified Certificate
    duration = calculate_duration(student.get('internship_start_date'), student.get('internship_end_date'))
    student['formatted_start_date'] = format_certificate_date(student.get('internship_start_date'))
    student['formatted_end_date'] = format_certificate_date(student.get('internship_end_date'))
    return render_template(
        'verification.html',
        status='verified',
        certificate_id=certificate_id,
        student=student,
        duration=duration
    )

# ---------------------------------------------------------------------------
# API: QR CODE GENERATOR
# ---------------------------------------------------------------------------

@app.route('/api/qr/<path:certificate_id>')
def generate_qr(certificate_id):
    """Generates PNG QR code pointing to public verification URL."""
    # Build complete public verification URL
    base_url = request.host_url.rstrip('/')
    verification_url = f"{base_url}/verify/{certificate_id}"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=2,
    )
    qr.add_data(verification_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#0F172A", back_color="#FFFFFF")

    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return send_file(buf, mimetype='image/png')

# ---------------------------------------------------------------------------
# API: BATCHES
# ---------------------------------------------------------------------------

@app.route('/api/batches', methods=['GET'])
def list_batches():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT b.*,
            (SELECT COUNT(*) FROM students s WHERE s.batch_id = b.id AND s.archived = 0) AS student_count,
            (SELECT COUNT(*) FROM students s WHERE s.batch_id = b.id AND s.status = 'Pending' AND s.archived = 0) AS pending_count,
            (SELECT COUNT(*) FROM students s WHERE s.batch_id = b.id AND s.status = 'Approved' AND s.archived = 0) AS approved_count,
            (SELECT COUNT(*) FROM students s WHERE s.batch_id = b.id AND s.status = 'Certificate Generated' AND s.archived = 0) AS generated_count
        FROM batches b
        ORDER BY b.id ASC
    ''')
    batches = [dict(r) for r in cur.fetchall()]
    conn.close()
    return jsonify(batches)

@app.route('/api/batches', methods=['POST'])
def create_batch():
    data = request.get_json() or {}
    batch_name = (data.get('batch_name') or '').strip()
    batch_code = (data.get('batch_code') or '').strip().upper()
    leader = (data.get('leader') or 'Admin Leader').strip()

    if not batch_name or not batch_code:
        return jsonify({'error': 'Batch Name and Batch Code are required'}), 400

    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT id FROM batches WHERE batch_code = ?", (batch_code,))
    if cur.fetchone():
        conn.close()
        return jsonify({'error': f"Batch code '{batch_code}' already exists"}), 409

    cur.execute('''
        INSERT INTO batches (batch_code, batch_name, leader, created_at, status, email_enabled)
        VALUES (?, ?, ?, datetime('now'), 'Active', 1)
    ''', (batch_code, batch_name, leader))
    new_id = cur.lastrowid
    conn.commit()

    cur.execute("SELECT * FROM batches WHERE id = ?", (new_id,))
    new_batch = dict(cur.fetchone())
    conn.close()
    return jsonify({'success': True, 'batch': new_batch}), 201

@app.route('/api/batches/<int:batch_id>/toggle-email', methods=['POST'])
def toggle_batch_email(batch_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT email_enabled FROM batches WHERE id = ?", (batch_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Batch not found'}), 404

    new_val = 0 if row['email_enabled'] else 1
    cur.execute("UPDATE batches SET email_enabled = ? WHERE id = ?", (new_val, batch_id))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'email_enabled': new_val})

# ---------------------------------------------------------------------------
# API: STUDENTS
# ---------------------------------------------------------------------------

@app.route('/api/batches/<int:batch_id>/students', methods=['GET'])
def get_batch_students(batch_id):
    conn = get_db()
    cur = conn.cursor()

    status_filter = request.args.get('status', '').strip()
    search = request.args.get('search', '').strip().lower()

    query = '''
        SELECT s.*, b.batch_code, b.batch_name
        FROM students s
        JOIN batches b ON s.batch_id = b.id
        WHERE s.batch_id = ? AND s.archived = 0
    '''
    params = [batch_id]

    if status_filter and status_filter.lower() != 'all':
        query += ' AND s.status = ?'
        params.append(status_filter)

    if search:
        query += ''' AND (
            LOWER(s.full_name) LIKE ? OR
            LOWER(s.register_number) LIKE ? OR
            LOWER(s.college_name) LIKE ? OR
            LOWER(s.domain) LIKE ? OR
            LOWER(s.certificate_id) LIKE ?
        )'''
        s_param = f"%{search}%"
        params.extend([s_param, s_param, s_param, s_param, s_param])

    query += ' ORDER BY s.id DESC'
    cur.execute(query, params)
    students = []
    for row in cur.fetchall():
        st = dict(row)
        st['duration'] = calculate_duration(st.get('internship_start_date'), st.get('internship_end_date'))
        students.append(st)

    conn.close()
    return jsonify(students)

@app.route('/api/students', methods=['POST'])
def submit_student():
    """Handles student form submission from Vanilla JS fetch."""
    data = request.get_json() or {}

    batch_code = (data.get('batchCode') or data.get('batch_code') or 'APP26-27').strip()
    full_name = (data.get('fullName') or data.get('full_name') or '').strip()
    register_number = (data.get('registerNo') or data.get('register_no') or data.get('register_number') or '').strip()
    college_name = (data.get('collegeName') or data.get('college_name') or '').strip()
    degree_branch = (data.get('degreeBranch') or data.get('degree_branch') or '').strip()
    state = (data.get('state') or '').strip()
    email = (data.get('email') or '').strip()
    phone_number = (data.get('phone') or data.get('phone_number') or '').strip()
    domain = (data.get('domain') or '').strip()
    mode = (data.get('mode') or 'Online').strip()
    start_date = (data.get('startDate') or data.get('internship_start_date') or '').strip()
    end_date = (data.get('endDate') or data.get('internship_end_date') or '').strip()
    project_title = (data.get('projectTitle') or data.get('project_title') or '').strip()

    # Validation
    if not (full_name and register_number and college_name and degree_branch and email and domain and start_date and end_date and project_title and phone_number):
        return jsonify({'error': 'All required fields must be provided.'}), 400

    conn = get_db()
    cur = conn.cursor()

    # Find batch ID
    cur.execute("SELECT id FROM batches WHERE batch_code = ?", (batch_code,))
    batch_row = cur.fetchone()
    if not batch_row:
        # Fallback to default
        cur.execute("SELECT id FROM batches WHERE batch_code = 'APP26-27'")
        batch_row = cur.fetchone()
    batch_id = batch_row['id'] if batch_row else 1

    cur.execute('''
        INSERT INTO students (
            batch_id, full_name, register_number, college_name, degree_branch,
            state, email, domain, mode, internship_start_date, internship_end_date,
            project_title, phone_number, status, certificate_id,
            certificate_generated, submitted_at, approved_at, archived
        ) VALUES (
            ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?,
            ?, ?, 'Pending', NULL,
            0, datetime('now'), NULL, 0
        )
    ''', (
        batch_id, full_name, register_number, college_name, degree_branch,
        state, email, domain, mode, start_date, end_date,
        project_title, phone_number
    ))

    new_id = cur.lastrowid
    conn.commit()
    conn.close()

    ref_id = f"AP-2026-{new_id:04d}"
    return jsonify({
        'success': True,
        'student_id': new_id,
        'reference_id': ref_id,
        'message': 'Student submission successfully recorded as Pending.'
    }), 201

@app.route('/api/students/<int:student_id>', methods=['GET'])
def get_student(student_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT s.*, b.batch_code, b.batch_name
        FROM students s
        JOIN batches b ON s.batch_id = b.id
        WHERE s.id = ?
    ''', (student_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return jsonify({'error': 'Student not found'}), 404

    st = dict(row)
    st['duration'] = calculate_duration(st.get('internship_start_date'), st.get('internship_end_date'))
    return jsonify(st)

@app.route('/api/students/<int:student_id>', methods=['PUT'])
def update_student(student_id):
    """Admin edit student fields without losing Certificate ID or timestamps."""
    data = request.get_json() or {}
    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT * FROM students WHERE id = ?", (student_id,))
    current = cur.fetchone()
    if not current:
        conn.close()
        return jsonify({'error': 'Student not found'}), 404

    cur.execute('''
        UPDATE students SET
            full_name = ?,
            register_number = ?,
            college_name = ?,
            degree_branch = ?,
            state = ?,
            email = ?,
            domain = ?,
            mode = ?,
            internship_start_date = ?,
            internship_end_date = ?,
            project_title = ?,
            phone_number = ?
        WHERE id = ?
    ''', (
        data.get('full_name', current['full_name']).strip(),
        data.get('register_number', current['register_number']).strip(),
        data.get('college_name', current['college_name']).strip(),
        data.get('degree_branch', current['degree_branch']).strip(),
        data.get('state', current['state']),
        data.get('email', current['email']).strip(),
        data.get('domain', current['domain']).strip(),
        data.get('mode', current['mode']).strip(),
        data.get('internship_start_date', current['internship_start_date']).strip(),
        data.get('internship_end_date', current['internship_end_date']).strip(),
        data.get('project_title', current['project_title']).strip(),
        data.get('phone_number', current['phone_number']).strip(),
        student_id
    ))

    conn.commit()
    cur.execute('''
        SELECT s.*, b.batch_code, b.batch_name
        FROM students s
        JOIN batches b ON s.batch_id = b.id
        WHERE s.id = ?
    ''', (student_id,))
    updated = dict(cur.fetchone())
    conn.close()
    return jsonify({'success': True, 'student': updated})

@app.route('/api/students/<int:student_id>/approve', methods=['POST'])
def approve_student(student_id):
    """Admin approves student and generates permanent Certificate ID."""
    conn = get_db()
    cur = conn.cursor()

    cur.execute('''
        SELECT s.*, b.batch_code
        FROM students s
        JOIN batches b ON s.batch_id = b.id
        WHERE s.id = ?
    ''', (student_id,))
    student = cur.fetchone()
    if not student:
        conn.close()
        return jsonify({'error': 'Student not found'}), 404

    # Keep existing cert ID if already generated, else generate new
    cert_id = student['certificate_id']
    if not cert_id:
        cert_id = generate_unique_certificate_id(student['batch_code'], conn)

    now_str = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')

    cur.execute('''
        UPDATE students SET
            status = 'Approved',
            certificate_id = ?,
            approved_at = ?
        WHERE id = ?
    ''', (cert_id, now_str, student_id))

    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'status': 'Approved',
        'certificate_id': cert_id,
        'approved_at': now_str
    })

@app.route('/api/students/<int:student_id>/generate-certificate', methods=['POST'])
def generate_certificate(student_id):
    """Admin generates certificate after approval."""
    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT * FROM students WHERE id = ?", (student_id,))
    student = cur.fetchone()
    if not student:
        conn.close()
        return jsonify({'error': 'Student not found'}), 404

    if not student['certificate_id']:
        conn.close()
        return jsonify({'error': 'Student must be approved before certificate generation'}), 400

    cur.execute('''
        UPDATE students SET
            status = 'Certificate Generated',
            certificate_generated = 1
        WHERE id = ?
    ''', (student_id,))

    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'status': 'Certificate Generated',
        'certificate_id': student['certificate_id'],
        'certificate_generated': 1
    })

@app.route('/api/students/<int:student_id>/certificate', methods=['GET'])
def get_student_certificate(student_id):
    """Retrieves full certificate display model with verification URL."""
    conn = get_db()
    cur = conn.cursor()

    cur.execute('''
        SELECT s.*, b.batch_code, b.batch_name
        FROM students s
        JOIN batches b ON s.batch_id = b.id
        WHERE s.id = ?
    ''', (student_id,))
    student = cur.fetchone()
    conn.close()

    if not student:
        return jsonify({'error': 'Student not found'}), 404

    st = dict(student)
    st['duration'] = calculate_duration(st.get('internship_start_date'), st.get('internship_end_date'))
    st['formatted_start_date'] = format_certificate_date(st.get('internship_start_date'))
    st['formatted_end_date'] = format_certificate_date(st.get('internship_end_date'))
    base_url = request.host_url.rstrip('/')
    st['verification_url'] = f"{base_url}/verify/{st.get('certificate_id')}"
    st['qr_api_url'] = f"{base_url}/api/qr/{st.get('certificate_id')}"
    st['docx_url'] = f"{base_url}/api/students/{student_id}/download-docx"
    st['company_name'] = 'Approtech R&D Solutions Pvt. Ltd.'

    return jsonify(st)

@app.route('/api/students/<int:student_id>/download-docx', methods=['GET'])
def download_student_docx(student_id):
    """Generates and serves a real .docx certificate file for a student."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT s.*, b.batch_code, b.batch_name
        FROM students s
        JOIN batches b ON s.batch_id = b.id
        WHERE s.id = ?
    ''', (student_id,))
    student = cur.fetchone()
    conn.close()

    if not student:
        return jsonify({'error': 'Student not found'}), 404

    st = dict(student)
    if not st.get('certificate_id'):
        return jsonify({'error': 'Certificate has not been approved or generated yet'}), 400

    base_url = request.host_url.rstrip('/')
    buf = generate_certificate_docx(st, base_url)
    clean_id = (st.get('certificate_id') or 'CERT').replace(':', '_').replace('/', '_')
    filename = f"Approtech_Certificate_{clean_id}.docx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/download-docx/<path:certificate_id>', methods=['GET'])
def download_docx_by_cert_id(certificate_id):
    """Public direct .docx certificate download by Certificate ID."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT s.*, b.batch_code, b.batch_name
        FROM students s
        JOIN batches b ON s.batch_id = b.id
        WHERE s.certificate_id = ?
    ''', (certificate_id.strip(),))
    student = cur.fetchone()
    conn.close()

    if not student:
        return render_template('verification.html', status='invalid', certificate_id=certificate_id), 404

    st = dict(student)
    if st.get('certificate_generated') != 1:
        return render_template('verification.html', status='pending', certificate_id=certificate_id, student=st), 400

    base_url = request.host_url.rstrip('/')
    buf = generate_certificate_docx(st, base_url)
    clean_id = certificate_id.replace(':', '_').replace('/', '_')
    filename = f"Approtech_Certificate_{clean_id}.docx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/api/students/<int:student_id>/archive', methods=['POST'])
def archive_student(student_id):
    """Soft archives a student without destroying database data."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE students SET archived = 1 WHERE id = ?", (student_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'Student record archived.'})

@app.route('/api/students/<int:student_id>', methods=['DELETE'])
@app.route('/api/students/<int:student_id>/delete', methods=['POST'])
def delete_student(student_id):
    """Permanently deletes a student record from the database."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id FROM students WHERE id = ?", (student_id,))
    if not cur.fetchone():
        conn.close()
        return jsonify({'error': 'Student not found'}), 404

    cur.execute("DELETE FROM students WHERE id = ?", (student_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'Student record permanently deleted.'})

# ---------------------------------------------------------------------------
# CLI & ENTRY POINT
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Approtech Internship Certificate Management System')
    parser.add_argument('--port', type=int, default=3000, help='Port to bind (default: 3000)')
    parser.add_argument('--host', type=str, default='0.0.0.0', help='Host to bind (default: 0.0.0.0)')
    args, _ = parser.parse_known_args()

    # Always ensure DB is initialized safely
    init_db()

    print(f"Starting Approtech System on http://{args.host}:{args.port}")
    app.run(host=args.host, port=args.port, debug=False)
