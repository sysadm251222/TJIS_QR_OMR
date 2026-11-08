# TJIS-BPT QR + OMR Examination System

## Features
- Student bulk upload from Excel (.xlsx) or CSV; downloadable import template.
- Exam question bulk upload from Excel or CSV; downloadable template.
- Bulk ZIP download containing an individual QR question paper PDF and QR OMR sheet PDF for every active student in the exam grade.
- Admin can create multiple teacher logins and enable/disable teacher accounts.
- Passwords stored using Werkzeug password hashing.
- QR identification, fixed-layout OMR bubble detection prototype, MCQ answer-key scoring, teacher descriptive marking, results list.

## Install (Windows)
1. Install Python 3.10+.
2. Extract this ZIP and open Command Prompt in the folder.
3. Run:
   `py -m venv .venv`
   `.venv\Scripts\activate`
   `pip install -r requirements.txt`
   `python app.py`
4. Open http://127.0.0.1:5000

Initial admin login: admin / admin123. Change this before any live use. Set environment variable SECRET_KEY to a long random secret.

## Student bulk upload
Students > Download Excel Template. Required columns: admission_no, name, grade. Optional: section, roll_no. Upload .xlsx or .csv. Duplicates and incomplete rows are skipped.

## Question bulk upload
Open an exam > Download Question Excel Template. Columns: qno, qtype, prompt, max_marks, option_a, option_b, option_c, option_d, answer_key. qtype is MCQ or DESCRIPTIVE; MCQ answer_key is A/B/C/D.

## Bulk download
Open exam > Download ALL Papers + OMR ZIP. Contains Question_Papers and OMR_Sheets folders, with personalized PDFs for every active student in that exam grade. OMR template supports up to 50 MCQs and A-D options.

## Teacher logins
Admin > Teacher Logins > create full name, username and password (minimum 8 characters). Admin can disable/enable teacher accounts.

## OMR scanning limitations
Scan the entire A4 sheet at 300 DPI, grayscale, no cropping/rotation. This fixed-layout OMR prototype must be tested against the school's printer/scanner. Skew, faint pencil, lighting, scaling or poor image quality can cause errors. Teacher verification is required before publishing marks. Descriptive handwritten answers are not OCR-read; teachers enter marks manually.

This is a starter system, not a certified high-stakes OMR product. For production, add HTTPS, CSRF protection, role-specific permissions, audit logs, backups, password-change workflow, secure deployment and extensive pilot testing.

\n## Updated school branding\n- The navigation bar now displays the supplied JGI logo and The Jain International School, Bangarpet branding.\n- A school-campus banner crop is used as a subtle page background with a light overlay for readability.\n- Branding assets are stored in `static/images/`. Replace `jgi-logo.png` or `school-campus-bg.jpg` to use different images; keep the same filenames.\n