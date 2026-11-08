import os,io,csv,json,sqlite3,zipfile,re
from functools import wraps
from datetime import datetime
from flask import Flask,render_template,request,redirect,url_for,session,flash,send_file,abort
from werkzeug.security import generate_password_hash,check_password_hash
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font, PatternFill
import qrcode,cv2,numpy as np
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

HERE=os.path.dirname(os.path.abspath(__file__))
DB=os.path.join(HERE,"school_exam.db")
app=Flask(__name__); app.secret_key=os.environ.get("SECRET_KEY","CHANGE-ME-use-a-long-random-secret")
app.config["MAX_CONTENT_LENGTH"]=20*1024*1024

def db():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; c.execute("PRAGMA foreign_keys=ON"); return c
def init_db():
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,username TEXT UNIQUE NOT NULL,password_hash TEXT NOT NULL,full_name TEXT NOT NULL,role TEXT NOT NULL,active INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS students(id INTEGER PRIMARY KEY,admission_no TEXT UNIQUE NOT NULL,name TEXT NOT NULL,grade TEXT NOT NULL,section TEXT DEFAULT '',roll_no TEXT DEFAULT '',active INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS exams(id INTEGER PRIMARY KEY,title TEXT NOT NULL,grade TEXT NOT NULL,subject TEXT NOT NULL,exam_date TEXT,duration INTEGER DEFAULT 0,instructions TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS questions(id INTEGER PRIMARY KEY,exam_id INTEGER NOT NULL,qno INTEGER NOT NULL,qtype TEXT NOT NULL,prompt TEXT NOT NULL,max_marks REAL NOT NULL,options TEXT DEFAULT '[]',answer_key TEXT DEFAULT '',UNIQUE(exam_id,qno),FOREIGN KEY(exam_id) REFERENCES exams(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS scripts(id INTEGER PRIMARY KEY,exam_id INTEGER NOT NULL,student_id INTEGER NOT NULL,status TEXT DEFAULT 'Pending',UNIQUE(exam_id,student_id),FOREIGN KEY(exam_id) REFERENCES exams(id),FOREIGN KEY(student_id) REFERENCES students(id));
        CREATE TABLE IF NOT EXISTS answers(id INTEGER PRIMARY KEY,script_id INTEGER NOT NULL,question_id INTEGER NOT NULL,response TEXT DEFAULT '',marks REAL DEFAULT 0,teacher_note TEXT DEFAULT '',UNIQUE(script_id,question_id),FOREIGN KEY(script_id) REFERENCES scripts(id) ON DELETE CASCADE,FOREIGN KEY(question_id) REFERENCES questions(id));
        """)
        if not c.execute("SELECT 1 FROM users WHERE username='admin'").fetchone():
            c.execute("INSERT INTO users(username,password_hash,full_name,role) VALUES(?,?,?,'admin')",("admin",generate_password_hash("admin123"),"System Administrator"))
        if not c.execute("SELECT 1 FROM students LIMIT 1").fetchone():
            c.executemany("INSERT INTO students(admission_no,name,grade,section,roll_no) VALUES(?,?,?,?,?)",[("TJIS001","Aarav Kumar","8","A","1"),("TJIS002","Diya Reddy","8","A","2")])
        if not c.execute("SELECT 1 FROM exams LIMIT 1").fetchone():
            eid=c.execute("INSERT INTO exams(title,grade,subject,exam_date) VALUES(?,?,?,?)",("Sample Term 1 Examination","8","Mathematics","2026-11-15")).lastrowid
            c.executemany("INSERT INTO questions(exam_id,qno,qtype,prompt,max_marks,options,answer_key) VALUES(?,?,?,?,?,?,?)",[(eid,1,"MCQ","What is 7 × 8?",1,json.dumps(["Option A","Option B","Option C","Option D"]),"B"),(eid,2,"DESCRIPTIVE","Explain your answer.",5,"[]","")])
def logged(f):
    @wraps(f)
    def w(*a,**k):
        if not session.get("uid"): return redirect(url_for("login"))
        with db() as c: u=c.execute("SELECT active FROM users WHERE id=?",(session["uid"],)).fetchone()
        if not u or not u["active"]: session.clear(); return redirect(url_for("login"))
        return f(*a,**k)
    return w
def admin(f):
    @wraps(f)
    def w(*a,**k):
        if not session.get("uid"): return redirect(url_for("login"))
        if session.get("role")!="admin": abort(403)
        with db() as c: u=c.execute("SELECT active,role FROM users WHERE id=?",(session["uid"],)).fetchone()
        if not u or not u["active"] or u["role"]!="admin": session.clear(); return redirect(url_for("login"))
        return f(*a,**k)
    return w
def excel_rows(f):
    ext=f.filename.rsplit(".",1)[-1].lower()
    if ext=="csv": return list(csv.DictReader(io.StringIO(f.stream.read().decode("utf-8-sig"))))
    if ext!="xlsx": raise ValueError("Only XLSX or CSV files are supported.")
    ws=load_workbook(f,read_only=True,data_only=True).active
    rows=list(ws.iter_rows(values_only=True))
    if not rows:return []
    heads=[str(x).strip().lower() if x is not None else "" for x in rows[0]]
    return [dict(zip(heads,r)) for r in rows[1:] if any(x is not None for x in r)]
def normalize_header(value):
    """Normalize headings such as 'Admission No.' and 'Student Name'."""
    return re.sub(r"[^a-z0-9]", "", str(value or "").strip().lower())

def getv(r,*names):
    d={normalize_header(k):v for k,v in r.items()}
    for n in names:
        v=d.get(normalize_header(n))
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return ""

def normalize_grade(value):
    """Compare labels such as 'Grade 8', 'Class-8', 'VIII', and '8'."""
    raw=str(value or "").strip().lower()
    raw=re.sub(r"^(grade|class|standard|std)\s*", "", raw)
    raw=re.sub(r"\s+", "", raw)
    roman={"i":"1","ii":"2","iii":"3","iv":"4","v":"5","vi":"6",
           "vii":"7","viii":"8","ix":"9","x":"10","xi":"11","xii":"12"}
    return roman.get(raw, raw)
def qr_img(exam_id,student_id):
    im=qrcode.make(json.dumps({"exam_id":int(exam_id),"student_id":int(student_id)}))
    b=io.BytesIO(); im.save(b,format="PNG"); b.seek(0); return b
def exam_student(exam_id,student_id):
    with db() as c:
        e=c.execute("SELECT * FROM exams WHERE id=?",(exam_id,)).fetchone()
        s=c.execute("SELECT * FROM students WHERE id=? AND active=1",(student_id,)).fetchone()
        qs=c.execute("SELECT * FROM questions WHERE exam_id=? ORDER BY qno",(exam_id,)).fetchall()
    if not e or not s: raise ValueError("Exam or student not found.")
    if e["grade"]!=s["grade"]: raise ValueError("Student grade does not match exam grade.")
    return e,s,qs
def pdf_omr(e,s,qs):
    mcq=[q for q in qs if q["qtype"]=="MCQ"]
    if len(mcq)>50: raise ValueError("OMR supports up to 50 MCQs per exam.")
    b=io.BytesIO(); w,h=A4; p=canvas.Canvas(b,pagesize=A4)
    p.setFont("Helvetica-Bold",13); p.drawString(15*mm,h-16*mm,"THE JAIN INTERNATIONAL SCHOOL, BANGARPET")
    p.setFont("Helvetica",9); p.drawString(15*mm,h-23*mm,f"{e['title']} | {e['subject']} | Grade {e['grade']}")
    p.drawString(15*mm,h-30*mm,f"Student: {s['name']} | Admission: {s['admission_no']} | Section: {s['section']}")
    p.drawImage(ImageReader(qr_img(e["id"],s["id"])),w-34*mm,h-42*mm,24*mm,24*mm)
    p.setFont("Helvetica-Bold",9); p.drawString(15*mm,h-47*mm,"OMR GRID — Shade ONE bubble per question")
    p.setFont("Helvetica",7); p.drawString(15*mm,h-53*mm,"Use dark blue/black pen. Do not tick. Scan complete page at 300 DPI.")
    xs=[75*mm,105*mm,135*mm,165*mm]; top=h-66*mm; gap=9.2*mm
    p.setFont("Helvetica-Bold",8); p.drawString(18*mm,top+3*mm,"Q.NO.")
    for x,l in zip(xs,"ABCD"):p.drawCentredString(x,top+3*mm,l)
    for i,q in enumerate(mcq):
        y=top-(i+1)*gap;p.drawString(20*mm,y-2.2*mm,str(q["qno"]))
        for x in xs:p.circle(x,y,2.8*mm,stroke=1,fill=0)
    p.setFont("Helvetica",7);p.drawString(15*mm,10*mm,"Blank or multiple responses are flagged for teacher verification.")
    p.save();b.seek(0);return b
def pdf_paper(e,s,qs):
    b=io.BytesIO();w,h=A4;p=canvas.Canvas(b,pagesize=A4)
    p.setFont("Helvetica-Bold",13);p.drawCentredString(w/2,h-17*mm,"THE JAIN INTERNATIONAL SCHOOL, BANGARPET")
    p.setFont("Helvetica-Bold",12);p.drawCentredString(w/2,h-26*mm,e["title"])
    p.setFont("Helvetica",9);p.drawString(15*mm,h-36*mm,f"Subject: {e['subject']} | Grade: {e['grade']} | Date: {e['exam_date'] or ''}")
    p.drawString(15*mm,h-43*mm,f"Student: {s['name']} | Admission: {s['admission_no']} | Section: {s['section']}")
    p.drawImage(ImageReader(qr_img(e["id"],s["id"])),w-32*mm,h-42*mm,20*mm,20*mm)
    y=h-56*mm;p.setFont("Helvetica",8);p.drawString(15*mm,y,f"Duration: {e['duration']} minutes");y-=12*mm
    for q in qs:
        p.setFont("Helvetica-Bold",9);p.drawString(15*mm,y,f"Q{q['qno']}. {q['prompt']} [{q['max_marks']} marks]");y-=7*mm
        if q["qtype"]=="MCQ":
            p.setFont("Helvetica",8)
            for i,opt in enumerate(json.loads(q["options"] or "[]")):p.drawString(22*mm,y,f"{'ABCD'[i]}. {opt}");y-=5*mm
        else:
            for _ in range(3):p.line(18*mm,y,w-18*mm,y);y-=8*mm
        y-=5*mm
        if y<25*mm:p.showPage();y=h-20*mm
    p.save();b.seek(0);return b

@app.route("/login",methods=["GET","POST"])
def login():
    if request.method=="POST":
        with db() as c:u=c.execute("SELECT * FROM users WHERE username=?",(request.form["username"].strip(),)).fetchone()
        if u and u["active"] and check_password_hash(u["password_hash"],request.form["password"]):
            session.update(uid=u["id"],username=u["username"],full_name=u["full_name"],role=u["role"]);return redirect(url_for("home"))
        flash("Invalid login or account disabled.")
    return render_template("login.html")
@app.route("/logout")
def logout():session.clear();return redirect(url_for("login"))
@app.route("/")
@logged
def home():
    with db() as c:
        stats=[c.execute("SELECT COUNT(*) FROM students WHERE active=1").fetchone()[0],c.execute("SELECT COUNT(*) FROM exams").fetchone()[0],c.execute("SELECT COUNT(*) FROM users WHERE role='teacher' AND active=1").fetchone()[0]]
        exams=c.execute("SELECT * FROM exams ORDER BY id DESC").fetchall()
    return render_template("home.html",stats=stats,exams=exams)

@app.route("/teachers",methods=["GET","POST"])
@admin
def teachers():
    if request.method=="POST":
        pw=request.form["password"]
        if len(pw)<8:flash("Password must be at least 8 characters.")
        else:
            try:
                with db() as c:c.execute("INSERT INTO users(username,password_hash,full_name,role) VALUES(?,?,?,'teacher')",(request.form["username"].strip(),generate_password_hash(pw),request.form["full_name"].strip()))
                flash("Teacher login created.")
            except sqlite3.IntegrityError:flash("Username already exists.")
        return redirect(url_for("teachers"))
    with db() as c:users=c.execute("SELECT id,username,full_name,role,active FROM users ORDER BY role,full_name").fetchall()
    return render_template("teachers.html",users=users)
@app.route("/teachers/<int:uid>/toggle",methods=["POST"])
@admin
def teacher_toggle(uid):
    with db() as c:c.execute("UPDATE users SET active=1-active WHERE id=? AND role='teacher'",(uid,))
    flash("Teacher account status changed.");return redirect(url_for("teachers"))

@app.route("/teachers/<int:uid>/edit",methods=["GET","POST"])
@admin
def teacher_edit(uid):
    with db() as c:
        user=c.execute("SELECT * FROM users WHERE id=? AND role='teacher'",(uid,)).fetchone()
        if not user: abort(404)
        if request.method=="POST":
            username=request.form.get("username","").strip(); full_name=request.form.get("full_name","").strip()
            if not username or not full_name: flash("Name and username are required.")
            else:
                try:
                    c.execute("UPDATE users SET username=?,full_name=? WHERE id=? AND role='teacher'",(username,full_name,uid))
                    flash("Teacher profile updated."); return redirect(url_for("teachers"))
                except sqlite3.IntegrityError: flash("That username is already in use.")
    return render_template("edit_teacher.html",u=user)

@app.route("/teachers/bulk-delete",methods=["POST"])
@admin
def teacher_bulk_delete():
    ids=[int(x) for x in request.form.getlist("user_ids") if x.isdigit()]
    if not ids: flash("Select at least one teacher.")
    else:
        with db() as c:
            marks=','.join('?' for _ in ids)
            cur=c.execute(f"UPDATE users SET active=0 WHERE role='teacher' AND id IN ({marks})",ids)
            flash(f"{cur.rowcount} teacher account(s) disabled. Accounts are retained for audit and can be re-enabled.")
    return redirect(url_for("teachers"))

@app.route("/teachers/<int:uid>/reset-password",methods=["POST"])
@admin
def teacher_reset_password(uid):
    pw=request.form.get("new_password","")
    if len(pw)<8: flash("New password must be at least 8 characters.")
    else:
        with db() as c:
            cur=c.execute("UPDATE users SET password_hash=? WHERE id=? AND role IN ('admin','teacher')",(generate_password_hash(pw),uid))
            flash("Password reset successfully." if cur.rowcount else "Account not found.")
    return redirect(url_for("teachers"))

@app.route("/account/password",methods=["GET","POST"])
@logged
def account_password():
    if request.method=="POST":
        old=request.form.get("current_password",""); new=request.form.get("new_password",""); confirm=request.form.get("confirm_password","")
        with db() as c: user=c.execute("SELECT * FROM users WHERE id=?",(session["uid"],)).fetchone()
        if not user or not check_password_hash(user["password_hash"],old): flash("Current password is incorrect.")
        elif len(new)<8: flash("New password must be at least 8 characters.")
        elif new!=confirm: flash("New password and confirmation do not match.")
        else:
            with db() as c:c.execute("UPDATE users SET password_hash=? WHERE id=?",(generate_password_hash(new),session["uid"]))
            flash("Your password has been changed."); return redirect(url_for("home"))
    return render_template("change_password.html")

@app.route("/students",methods=["GET","POST"])
@admin
def students():
    if request.method=="POST":
        try:
            with db() as c:c.execute("INSERT INTO students(admission_no,name,grade,section,roll_no) VALUES(?,?,?,?,?)",(request.form["admission_no"].strip(),request.form["name"].strip(),request.form["grade"].strip(),request.form.get("section",""),request.form.get("roll_no","")))
            flash("Student added.")
        except sqlite3.IntegrityError:flash("Admission number already exists.")
        return redirect(url_for("students"))
    with db() as c:rows=c.execute("SELECT * FROM students WHERE active=1 ORDER BY grade,section,name").fetchall()
    return render_template("students.html",students=rows)
@app.route("/students/<int:sid>/edit",methods=["GET","POST"])
@admin
def student_edit(sid):
    with db() as c:
        student=c.execute("SELECT * FROM students WHERE id=?",(sid,)).fetchone()
        if not student: abort(404)
        if request.method=="POST":
            values=(request.form.get("admission_no","").strip(),request.form.get("name","").strip(),request.form.get("grade","").strip(),request.form.get("section","").strip(),request.form.get("roll_no","").strip(),sid)
            if not all(values[:3]): flash("Admission number, student name and grade are required.")
            else:
                try:
                    c.execute("UPDATE students SET admission_no=?,name=?,grade=?,section=?,roll_no=? WHERE id=?",values)
                    flash("Student profile updated."); return redirect(url_for("students"))
                except sqlite3.IntegrityError: flash("Admission number already exists.")
    return render_template("edit_student.html",s=student)

@app.route("/students/bulk-delete",methods=["POST"])
@admin
def student_bulk_delete():
    ids=[int(x) for x in request.form.getlist("student_ids") if x.isdigit()]
    if not ids: flash("Select at least one student.")
    else:
        with db() as c:
            marks=','.join('?' for _ in ids)
            cur=c.execute(f"UPDATE students SET active=0 WHERE id IN ({marks}) AND active=1",ids)
            flash(f"{cur.rowcount} student(s) archived. Exam scripts and results are preserved.")
    return redirect(url_for("students"))

@app.route("/students/template")
@admin
def student_template():
    wb=Workbook();ws=wb.active;ws.append(["admission_no","name","grade","section","roll_no"]);ws.append(["TJIS001","Aarav Kumar","8","A","1"])
    b=io.BytesIO();wb.save(b);b.seek(0);return send_file(b,as_attachment=True,download_name="student_bulk_upload_template.xlsx")
@app.route("/students/bulk",methods=["POST"])
@admin
def student_bulk():
    f=request.files.get("file")
    if not f or not f.filename:
        flash("Please choose an Excel (.xlsx) or CSV file.")
        return redirect(url_for("students"))
    try:
        rows=excel_rows(f)
        ok=skip=0
        with db() as c:
            for r in rows:
                adm=getv(r,"admission_no","admission number","admission no","admission_number",
                         "adm no","adm_no","student id","student_code","registration number")
                name=getv(r,"name","student_name","student name","full name","name of student")
                grade=getv(r,"grade","class","standard","std","class/grade","grade name")
                section=getv(r,"section","sec","division")
                roll=getv(r,"roll_no","roll no","roll number","rollnumber")
                if not adm or not name or not grade:
                    skip+=1
                    continue
                grade=normalize_grade(grade)
                try:
                    c.execute(
                        "INSERT INTO students(admission_no,name,grade,section,roll_no) VALUES(?,?,?,?,?)",
                        (adm.strip(),name.strip(),grade,section.strip(),roll.strip())
                    )
                    ok+=1
                except sqlite3.IntegrityError:
                    skip+=1
        flash(f"Student import completed: {ok} added; {skip} skipped. "
              f"Required columns: Admission Number, Student Name, and Grade/Class.")
    except Exception as e:
        flash(f"Import error: {e}")
    return redirect(url_for("students"))

@app.route("/exams",methods=["GET","POST"])
@logged
def exams():
    if request.method=="POST":
        with db() as c:eid=c.execute("INSERT INTO exams(title,grade,subject,exam_date,duration,instructions) VALUES(?,?,?,?,?,?)",(request.form["title"],request.form["grade"],request.form["subject"],request.form.get("exam_date",""),int(request.form.get("duration") or 0),request.form.get("instructions",""))).lastrowid
        return redirect(url_for("exam",eid=eid))
    with db() as c:rows=c.execute("SELECT * FROM exams ORDER BY id DESC").fetchall()
    return render_template("exams.html",exams=rows)
@app.route("/exam/<int:eid>")
@logged
def exam(eid):
    with db() as c:
        e=c.execute("SELECT * FROM exams WHERE id=?",(eid,)).fetchone()
        if not e:abort(404)
        qs=c.execute("SELECT * FROM questions WHERE exam_id=? ORDER BY qno",(eid,)).fetchall()
        active_students=c.execute("SELECT id,name,admission_no,grade,section,roll_no FROM students WHERE active=1").fetchall()
        eligible_students=[s for s in active_students
                           if normalize_grade(s["grade"]) == normalize_grade(e["grade"])]
        n=len(eligible_students)
    return render_template("exam.html",e=e,questions=qs,n=n,eligible_students=eligible_students)
@app.route("/exam/<int:eid>/question",methods=["POST"])
@logged
def addq(eid):
    try:
        typ=request.form["qtype"];opts=[x.strip() for x in request.form.get("options","").split("|") if x.strip()] if typ=="MCQ" else []
        with db() as c:c.execute("INSERT INTO questions(exam_id,qno,qtype,prompt,max_marks,options,answer_key) VALUES(?,?,?,?,?,?,?)",(eid,int(request.form["qno"]),typ,request.form["prompt"],float(request.form["max_marks"]),json.dumps(opts),request.form.get("answer_key","").upper() if typ=="MCQ" else ""))
        flash("Question added.")
    except Exception as e:flash(f"Question error: {e}")
    return redirect(url_for("exam",eid=eid))
@app.route("/exam/<int:eid>/questions-template")
@logged
def qtemplate(eid):
    wb=Workbook();ws=wb.active;ws.append(["qno","qtype","prompt","max_marks","option_a","option_b","option_c","option_d","answer_key"]);ws.append([1,"MCQ","Example question",1,"Red","Blue","Green","Yellow","B"]);ws.append([2,"DESCRIPTIVE","Explain your answer",5,"","","","",""])
    b=io.BytesIO();wb.save(b);b.seek(0);return send_file(b,as_attachment=True,download_name="exam_questions_template.xlsx")
@app.route("/exam/<int:eid>/questions-bulk",methods=["POST"])
@logged
def qbulk(eid):
    try:
        rows=excel_rows(request.files["file"]);ok=skip=0
        with db() as c:
            for r in rows:
                try:
                    qno=int(float(getv(r,"qno","question_no")));typ=getv(r,"qtype","type").upper();prompt=getv(r,"prompt","question");marks=float(getv(r,"max_marks","marks"))
                    if typ not in ("MCQ","DESCRIPTIVE") or not prompt or marks<=0:raise ValueError()
                    if typ=="MCQ":
                        opts=[getv(r,"option_"+x) for x in "abcd"];opts=[x for x in opts if x];key=getv(r,"answer_key","correct_answer").upper()
                        if len(opts)<2 or key not in "ABCD":raise ValueError()
                    else:opts=[];key=""
                    c.execute("INSERT INTO questions(exam_id,qno,qtype,prompt,max_marks,options,answer_key) VALUES(?,?,?,?,?,?,?)",(eid,qno,typ,prompt,marks,json.dumps(opts),key));ok+=1
                except Exception:skip+=1
        flash(f"Question import: {ok} added; {skip} skipped.")
    except Exception as e:flash(f"Import failed: {e}")
    return redirect(url_for("exam",eid=eid))
@app.route("/exam/<int:eid>/bulk-download")
@logged
def bulk_download(eid):
    with db() as c:
        e=c.execute("SELECT * FROM exams WHERE id=?",(eid,)).fetchone()
        if not e:abort(404)
        all_students=c.execute("SELECT * FROM students WHERE active=1 ORDER BY section,name").fetchall()
        sts=[s for s in all_students if normalize_grade(s["grade"]) == normalize_grade(e["grade"])]
        qs=c.execute("SELECT * FROM questions WHERE exam_id=? ORDER BY qno",(eid,)).fetchall()
    if not qs or not sts:flash("Add questions and eligible students before bulk download.");return redirect(url_for("exam",eid=eid))
    b=io.BytesIO()
    with zipfile.ZipFile(b,"w",zipfile.ZIP_DEFLATED) as z:
        for s in sts:
            stem=f"{s['section']}_{s['admission_no']}".replace("/","-")
            z.writestr(f"Question_Papers/{stem}_Question_Paper.pdf",pdf_paper(e,s,qs).getvalue())
            try:z.writestr(f"OMR_Sheets/{stem}_OMR_Sheet.pdf",pdf_omr(e,s,qs).getvalue())
            except Exception as ex:z.writestr(f"ERRORS/{stem}.txt",str(ex))
    b.seek(0);return send_file(b,mimetype="application/zip",as_attachment=True,download_name=f"Exam_{eid}_Student_Papers_and_OMR.zip")

def process_omr_bytes(file_bytes):
    """Decode one complete A4 OMR image and save its MCQ responses."""
    raw=np.frombuffer(file_bytes,np.uint8)
    im=cv2.imdecode(raw,cv2.IMREAD_COLOR)
    if im is None: raise ValueError("Could not read image. Use JPG or PNG.")
    if im.shape[1]<700 or im.shape[0]<900: raise ValueError("Scan the complete A4 sheet; avoid cropping.")
    im=cv2.resize(im,(2480,3508))
    payload,_,_=cv2.QRCodeDetector().detectAndDecode(im)
    if not payload: raise ValueError("QR code not detected. Ensure it is clear and not covered.")
    data=json.loads(payload); eid=int(data["exam_id"]); sid=int(data["student_id"])
    e,student,qs=exam_student(eid,sid)
    mcqs=[q for q in qs if q["qtype"]=="MCQ"]
    gray=cv2.cvtColor(im,cv2.COLOR_BGR2GRAY)
    xs=[75*mm,105*mm,135*mm,165*mm]; top=A4[1]-66*mm; gap=9.2*mm; flags=[]
    with db() as c:
        c.execute("INSERT OR IGNORE INTO scripts(exam_id,student_id,status) VALUES(?,?,?)",(eid,sid,"Pending"))
        script=c.execute("SELECT id FROM scripts WHERE exam_id=? AND student_id=?",(eid,sid)).fetchone()["id"]
        for q in qs:
            c.execute("INSERT OR IGNORE INTO answers(script_id,question_id) VALUES(?,?)",(script,q["id"]))
        for i,q in enumerate(mcqs):
            yp=int((A4[1]-(top-(i+1)*gap))*300/72); dark=[]
            for xppt in xs:
                xp=int(xppt*300/72); rad=22
                roi=gray[max(0,yp-rad):yp+rad+1,max(0,xp-rad):xp+rad+1]
                dark.append(255-float(np.mean(roi)))
            marked=[j for j,x in enumerate(dark) if x>48]
            response="ABCD"[marked[0]] if len(marked)==1 else ("MULTIPLE" if marked else "")
            marks=float(q["max_marks"]) if response and response==q["answer_key"].upper() else 0
            if len(marked)!=1 or not q["answer_key"]: flags.append(f"Q{q['qno']}")
            aid=c.execute("SELECT id FROM answers WHERE script_id=? AND question_id=?",(script,q["id"])).fetchone()["id"]
            c.execute("UPDATE answers SET response=?,marks=? WHERE id=?",(response,marks,aid))
        c.execute("UPDATE scripts SET status='Auto-Corrected - Review Required' WHERE id=?",(script,))
    return script,student["name"],flags

@app.route("/omr-scan",methods=["GET","POST"])
@logged
def omr_scan():
    if request.method=="POST":
        try:
            script,name,flags=process_omr_bytes(request.files["sheet"].read())
            flash(f"OMR read for {name}. Review flagged questions: {', '.join(flags) if flags else 'none'}.")
            return redirect(url_for("grade",sid=script))
        except Exception as ex: flash(f"OMR scan error: {ex}")
        return redirect(url_for("omr_scan"))
    return render_template("omr_scan.html")

@app.route("/omr-bulk-upload",methods=["POST"])
@logged
def omr_bulk_upload():
    files=request.files.getlist("sheets"); success=[]; errors=[]
    for f in files:
        if not f or not f.filename: continue
        try:
            script,name,flags=process_omr_bytes(f.read())
            success.append(f"{f.filename}: {name} (script #{script})" + (f"; review {', '.join(flags)}" if flags else ""))
        except Exception as ex: errors.append(f"{f.filename}: {ex}")
    if success: flash(f"Bulk OMR completed: {len(success)} sheet(s) processed. " + " | ".join(success[:8]))
    if errors: flash(f"Could not process {len(errors)} file(s): " + " | ".join(errors[:8]))
    return redirect(url_for("omr_scan"))

@app.route("/grade/<int:sid>",methods=["GET","POST"])
@logged
def grade(sid):
    with db() as c:
        script=c.execute("SELECT sc.*,st.name,st.admission_no,st.grade,st.section,e.title,e.subject FROM scripts sc JOIN students st ON st.id=sc.student_id JOIN exams e ON e.id=sc.exam_id WHERE sc.id=?",(sid,)).fetchone()
        if not script:abort(404)
        qs=c.execute("SELECT q.*,a.id aid,a.response,a.marks,a.teacher_note FROM questions q JOIN answers a ON a.question_id=q.id WHERE a.script_id=? ORDER BY q.qno",(sid,)).fetchall()
        if request.method=="POST":
            for q in qs:
                response=request.form.get(f"r{q['aid']}","").strip()
                if q["qtype"]=="MCQ":marks=float(q["max_marks"]) if response.upper()==q["answer_key"].upper() else 0
                else:
                    try:marks=float(request.form.get(f"m{q['aid']}","0"))
                    except:marks=0
                marks=max(0,min(float(q["max_marks"]),marks))
                c.execute("UPDATE answers SET response=?,marks=?,teacher_note=? WHERE id=?",(response,marks,request.form.get(f"n{q['aid']}",""),q["aid"]))
            c.execute("UPDATE scripts SET status='Evaluated' WHERE id=?",(sid,));flash("Evaluation saved; verify before publishing.");return redirect(url_for("grade",sid=sid))
        total=sum(float(q["marks"] or 0) for q in qs);maximum=sum(float(q["max_marks"]) for q in qs)
    return render_template("grade.html",s=script,qs=qs,total=total,maximum=maximum)
@app.route("/results")
@logged
def results():
    with db() as c:
        rows=c.execute("""
            SELECT
                sc.id AS sid, sc.status, st.name, st.admission_no, st.grade,
                e.title, e.subject,
                COALESCE((SELECT SUM(a.marks) FROM answers a WHERE a.script_id=sc.id),0) AS marks,
                COALESCE((SELECT SUM(q.max_marks) FROM questions q WHERE q.exam_id=sc.exam_id),0) AS maximum
            FROM scripts sc
            JOIN students st ON st.id=sc.student_id
            JOIN exams e ON e.id=sc.exam_id
            ORDER BY e.id DESC, st.grade, st.section, st.name
        """).fetchall()
    return render_template("results.html",rows=rows)

@app.route("/results/consolidated")
@logged
def consolidated_results():
    with db() as c:
        rows=c.execute("""
            SELECT st.id AS student_id, st.name, st.admission_no, st.grade, st.section,
                   e.subject, SUM(COALESCE(a.marks,0)) AS marks,
                   SUM(q.max_marks) AS maximum,
                   COUNT(DISTINCT sc.id) AS exam_count,
                   GROUP_CONCAT(DISTINCT e.title) AS exams
            FROM scripts sc
            JOIN students st ON st.id=sc.student_id
            JOIN exams e ON e.id=sc.exam_id
            JOIN answers a ON a.script_id=sc.id
            JOIN questions q ON q.id=a.question_id
            GROUP BY st.id, e.subject
            ORDER BY st.grade, st.section, st.name, e.subject
        """).fetchall()
    return render_template("consolidated_results.html",rows=rows)

if __name__=="__main__":
    init_db();print(f"Using student database: {DB}");app.run(host="0.0.0.0",port=int(os.environ.get("PORT",5000)),debug=False)
