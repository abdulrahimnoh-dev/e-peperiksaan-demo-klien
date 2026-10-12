from flask import Flask, render_template, request, redirect, url_for, session, flash, abort, send_file
import sqlite3
import os
from pathlib import Path
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE = Path(__file__).resolve().parent
DB = BASE / 'demo_data' / 'epeperiksaan_demo.db'
DB.parent.mkdir(parents=True, exist_ok=True)
app = Flask(__name__)
app.secret_key = os.environ.get('DEMO_SECRET_KEY', 'local-demo-only-change-before-public-deployment')
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Lax')
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024

STD_TYPES = {'UPSA','UASA','PSRA','PERCUBAAN PSRA','GERAK GEMPUR PSRA'}
UPKK_TYPES = {'UPKK','PERCUBAAN UPKK','GERAK GEMPUR UPKK'}

# Struktur rasmi yang telah dipersetujui untuk prototaip KISLB.
SUBJECT_DEFS = {
    'STD': [
        ('TAUHID_FEQAH','Tauhid & Feqah',100,[('TAUHID','Tauhid',50),('FEQAH','Feqah',50)],'STD100'),
        ('AKHLAK_SIRAH','Akhlak & Sirah',100,[('AKHLAK','Akhlak',50),('SIRAH','Sirah',50)],'STD100'),
        ('BARAB','Bahasa Arab',100,[('BARAB','Bahasa Arab',100)],'STD100'),
        ('TAJWID','Tajwid',100,[('TAJWID','Tajwid',100)],'STD100'),
        ('JAWI_IMLAK_KHAT','Jawi, Imlak & Khat',100,[('JAWI','Jawi',75),('IMLAK','Imlak',15),('KHAT','Khat',10)],'STD100'),
    ],
    'UPKK': [
        ('AKIDAH','Akidah',70,[('AKIDAH','Akidah',70)],'UPKK70'),
        ('IBADAH','Ibadah',70,[('IBADAH','Ibadah',70)],'UPKK70'),
        ('ADAB','Adab',70,[('ADAB','Adab',70)],'UPKK70'),
        ('SIRAH_UPKK','Sirah',70,[('SIRAH_UPKK','Sirah',70)],'UPKK70'),
        ('JAWI_KHAT_UPKK','Jawi & Khat',70,[('JAWI_KHAT_UPKK','Jawi & Khat',70)],'UPKK70'),
        ('BARAB_UPKK','Bahasa Arab',70,[('BARAB_UPKK','Bahasa Arab',70)],'UPKK70'),
        ('QURAN_HAFAZAN','Al-Quran & Hafazan',100,[('QURAN_HAFAZAN','Al-Quran & Hafazan',100)],'UPKK100'),
        ('PCHI','Penghayatan Cara Hidup Islam (PCHI)',100,[('PCHI','PCHI',100)],'UPKK100'),
        ('AMALI_SOLAT','Amali Solat',100,[('AMALI_SOLAT','Amali Solat',100)],'UPKK100'),
    ]
}
GRADE_SCHEMES = {
    'STD100': [(90,100,'MUMTAZ',1),(75,89.999,'JAYYID JIDDAN',2),(60,74.999,'JAYYID',3),(40,59.999,'MAQBUL',4),(0,39.999,'MUSAADAH',5)],
    'UPKK70':[(63,70,'A',1),(42,62,'B',2),(28,41,'C',3),(0,27,'D',4)],
    'UPKK100':[(85,100,'A',1),(60,84,'B',2),(40,59,'C',3),(0,39,'D',4)],
}

def db():
    con=sqlite3.connect(DB); con.row_factory=sqlite3.Row; con.execute('PRAGMA foreign_keys=ON'); return con

def init_db():
    con=db(); con.executescript('''
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE NOT NULL,password_hash TEXT NOT NULL,role TEXT NOT NULL,full_name TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS students(id INTEGER PRIMARY KEY,full_name TEXT NOT NULL,mykid TEXT UNIQUE NOT NULL,active INTEGER NOT NULL DEFAULT 1);
    CREATE TABLE IF NOT EXISTS classes(id INTEGER PRIMARY KEY,session_year INTEGER NOT NULL,level INTEGER NOT NULL,class_name TEXT NOT NULL,UNIQUE(session_year,class_name));
    CREATE TABLE IF NOT EXISTS student_class_history(id INTEGER PRIMARY KEY,student_id INTEGER NOT NULL,class_id INTEGER NOT NULL,start_date TEXT,end_date TEXT,is_current INTEGER NOT NULL DEFAULT 1,FOREIGN KEY(student_id) REFERENCES students(id),FOREIGN KEY(class_id) REFERENCES classes(id));
    CREATE TABLE IF NOT EXISTS subjects(id INTEGER PRIMARY KEY,code TEXT UNIQUE NOT NULL,name TEXT NOT NULL,max_score REAL NOT NULL DEFAULT 100,scheme_key TEXT NOT NULL DEFAULT 'STD100');
    CREATE TABLE IF NOT EXISTS subject_components(id INTEGER PRIMARY KEY,subject_id INTEGER NOT NULL,code TEXT NOT NULL,name TEXT NOT NULL,max_score REAL NOT NULL,sort_order INTEGER NOT NULL DEFAULT 0,UNIQUE(subject_id,code),FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE);
    CREATE TABLE IF NOT EXISTS exams(id INTEGER PRIMARY KEY,session_year INTEGER NOT NULL,exam_type TEXT NOT NULL,exam_name TEXT NOT NULL,series_no INTEGER,status TEXT NOT NULL DEFAULT 'Draf');
    CREATE TABLE IF NOT EXISTS exam_subjects(exam_id INTEGER NOT NULL,subject_id INTEGER NOT NULL,PRIMARY KEY(exam_id,subject_id),FOREIGN KEY(exam_id) REFERENCES exams(id),FOREIGN KEY(subject_id) REFERENCES subjects(id));
    CREATE TABLE IF NOT EXISTS teacher_assignments(teacher_user_id INTEGER NOT NULL,class_id INTEGER NOT NULL,subject_id INTEGER NOT NULL,PRIMARY KEY(teacher_user_id,class_id,subject_id));
    CREATE TABLE IF NOT EXISTS marks(id INTEGER PRIMARY KEY,exam_id INTEGER NOT NULL,student_id INTEGER NOT NULL,subject_id INTEGER NOT NULL,class_id_at_exam INTEGER NOT NULL,score REAL,updated_by INTEGER,updated_at TEXT DEFAULT CURRENT_TIMESTAMP,UNIQUE(exam_id,student_id,subject_id));
    CREATE TABLE IF NOT EXISTS component_marks(id INTEGER PRIMARY KEY,mark_id INTEGER NOT NULL,component_id INTEGER NOT NULL,score REAL,UNIQUE(mark_id,component_id),FOREIGN KEY(mark_id) REFERENCES marks(id) ON DELETE CASCADE,FOREIGN KEY(component_id) REFERENCES subject_components(id));
    CREATE TABLE IF NOT EXISTS parent_students(parent_user_id INTEGER NOT NULL,student_id INTEGER NOT NULL,PRIMARY KEY(parent_user_id,student_id));
    CREATE TABLE IF NOT EXISTS audit_log(id INTEGER PRIMARY KEY,user_id INTEGER,action TEXT NOT NULL,entity_type TEXT NOT NULL,entity_id INTEGER,detail TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS exam_grade_rules(id INTEGER PRIMARY KEY,exam_id INTEGER NOT NULL,scheme_key TEXT NOT NULL,grade TEXT NOT NULL,min_score REAL NOT NULL,max_score REAL NOT NULL,gps_point REAL NOT NULL,sort_order INTEGER NOT NULL DEFAULT 0,UNIQUE(exam_id,scheme_key,grade),FOREIGN KEY(exam_id) REFERENCES exams(id) ON DELETE CASCADE);
    ''')
    # migrations for v0.1 database if reused
    cols=[r['name'] for r in con.execute('PRAGMA table_info(subjects)').fetchall()]
    if 'max_score' not in cols: con.execute("ALTER TABLE subjects ADD COLUMN max_score REAL NOT NULL DEFAULT 100")
    if 'scheme_key' not in cols: con.execute("ALTER TABLE subjects ADD COLUMN scheme_key TEXT NOT NULL DEFAULT 'STD100'")
    ta_cols=[r['name'] for r in con.execute('PRAGMA table_info(teacher_assignments)').fetchall()]
    if 'exam_id' not in ta_cols:
        con.execute('ALTER TABLE teacher_assignments RENAME TO teacher_assignments_old')
        con.execute('''CREATE TABLE teacher_assignments(
            teacher_user_id INTEGER NOT NULL,class_id INTEGER NOT NULL,exam_id INTEGER NOT NULL,subject_id INTEGER NOT NULL,
            PRIMARY KEY(teacher_user_id,class_id,exam_id,subject_id))''')
        # Preserve old assignments by mapping them only to matching existing exams in the same session.
        con.execute('''INSERT OR IGNORE INTO teacher_assignments(teacher_user_id,class_id,exam_id,subject_id)
            SELECT o.teacher_user_id,o.class_id,e.id,o.subject_id
            FROM teacher_assignments_old o
            JOIN classes c ON c.id=o.class_id
            JOIN exams e ON e.session_year=c.session_year
            JOIN exam_subjects es ON es.exam_id=e.id AND es.subject_id=o.subject_id''')
        con.execute('DROP TABLE teacher_assignments_old')
    seed(con)
    for ex in con.execute("SELECT id FROM exams WHERE exam_type IN ('UPKK','PERCUBAAN UPKK','GERAK GEMPUR UPKK')").fetchall(): ensure_exam_grade_rules(con,ex['id'])
    # Bersihkan sempadan lalai lama (.999) kepada markah bulat tanpa mengubah tetapan custom Admin.
    for oldv,newv in [(62.999,62),(41.999,41),(27.999,27),(84.999,84),(59.999,59),(39.999,39)]:
        con.execute('UPDATE exam_grade_rules SET max_score=? WHERE ABS(max_score-?)<0.0001',(newv,oldv))
    con.commit(); con.close()

def seed_subjects(con):
    for family in SUBJECT_DEFS.values():
        for code,name,max_score,components,scheme in family:
            con.execute('INSERT OR IGNORE INTO subjects(code,name,max_score,scheme_key) VALUES(?,?,?,?)',(code,name,max_score,scheme))
            con.execute('UPDATE subjects SET name=?,max_score=?,scheme_key=? WHERE code=?',(name,max_score,scheme,code))
            sid=con.execute('SELECT id FROM subjects WHERE code=?',(code,)).fetchone()['id']
            for i,(cc,cn,cm) in enumerate(components):
                con.execute('INSERT OR IGNORE INTO subject_components(subject_id,code,name,max_score,sort_order) VALUES(?,?,?,?,?)',(sid,cc,cn,cm,i))
                con.execute('UPDATE subject_components SET name=?,max_score=?,sort_order=? WHERE subject_id=? AND code=?',(cn,cm,i,sid,cc))

def exam_family(exam_type):
    return 'UPKK' if exam_type in UPKK_TYPES else 'STD'

def configure_exam_subjects(con,eid,exam_type):
    family=exam_family(exam_type)
    codes=[x[0] for x in SUBJECT_DEFS[family]]
    con.execute('DELETE FROM exam_subjects WHERE exam_id=?',(eid,))
    for code in codes:
        sid=con.execute('SELECT id FROM subjects WHERE code=?',(code,)).fetchone()['id']
        con.execute('INSERT OR IGNORE INTO exam_subjects(exam_id,subject_id) VALUES(?,?)',(eid,sid))

def seed(con):
    # Pembetulan struktur PSRA: Gerak Gempur PSRA hanya Siri 1 dan Siri 2.
    # Jika DB lama masih ada demo Siri 3 dan belum ada Percubaan PSRA 2026,
    # tukar rekod itu kepada Percubaan PSRA supaya markah sedia ada kekal.
    old_s3=con.execute("SELECT id FROM exams WHERE exam_name='Gerak Gempur PSRA Siri 3'").fetchone()
    trial=con.execute("SELECT id FROM exams WHERE exam_type='PERCUBAAN PSRA' AND session_year=2026").fetchone()
    if old_s3 and not trial:
        con.execute("""UPDATE exams
            SET exam_type='PERCUBAAN PSRA', exam_name='Percubaan PSRA 2026', series_no=NULL
            WHERE id=?""",(old_s3['id'],))
        configure_exam_subjects(con,old_s3['id'],'PERCUBAAN PSRA')
    # Password demo diurus melalui Render Environment, bukan dalam GitHub.
    accounts = [
        ('admin', os.environ.get('DEMO_ADMIN_PASSWORD'), 'admin', 'Pentadbir Demo'),
        ('guru', os.environ.get('DEMO_GURU_PASSWORD'), 'teacher', 'Guru Demo'),
        ('ibu', os.environ.get('DEMO_IBU_PASSWORD'), 'parent', 'Ibu Bapa Demo'),
    ]
    for username, password, role, name in accounts:
        if not password:
            raise RuntimeError(f'Environment variable password untuk {username} belum ditetapkan')
        existing = con.execute('SELECT id FROM users WHERE username=?', (username,)).fetchone()
        if existing:
            con.execute('UPDATE users SET password_hash=?,role=?,full_name=? WHERE id=?',
                        (generate_password_hash(password), role, name, existing['id']))
        else:
            con.execute('INSERT INTO users(username,password_hash,role,full_name) VALUES(?,?,?,?)',
                        (username, generate_password_hash(password), role, name))
    seed_subjects(con)
    for yr,level,name in [(2026,5,'5 Al-Farabi'),(2026,5,'5 Al-Biruni'),(2027,6,'6 Al-Farabi')]: con.execute('INSERT OR IGNORE INTO classes(session_year,level,class_name) VALUES(?,?,?)',(yr,level,name))
    demo=[('Ahmad Demo','150101-10-0001'),('Aisyah Demo','150101-10-0002'),('Hafiz Demo','150101-10-0003'),('Maryam Demo','150101-10-0004'),('Umar Demo','150101-10-0005'),('Sofea Demo','150101-10-0006')]
    cid=con.execute("SELECT id FROM classes WHERE session_year=2026 AND class_name='5 Al-Farabi'").fetchone()['id']
    for name,mykid in demo:
        con.execute('INSERT OR IGNORE INTO students(full_name,mykid) VALUES(?,?)',(name,mykid)); sid=con.execute('SELECT id FROM students WHERE mykid=?',(mykid,)).fetchone()['id']
        if not con.execute('SELECT 1 FROM student_class_history WHERE student_id=? AND class_id=?',(sid,cid)).fetchone(): con.execute("INSERT INTO student_class_history(student_id,class_id,start_date,is_current) VALUES(?,?, '2026-01-01',1)",(sid,cid))
    parent=con.execute("SELECT id FROM users WHERE username='ibu'").fetchone()['id']; first=con.execute("SELECT id FROM students WHERE mykid='150101-10-0001'").fetchone()['id']; con.execute('INSERT OR IGNORE INTO parent_students VALUES(?,?)',(parent,first))
    # demo UPSA + Gerak Gempur PSRA Siri 1-2 + Percubaan PSRA
    demo_exams=[('UPSA','UPSA 2026',None,'Diterbitkan'),('GERAK GEMPUR PSRA','Gerak Gempur PSRA Siri 1',1,'Diterbitkan'),('PERCUBAAN PSRA','Percubaan PSRA 2026',None,'Sedang Key-in'),('GERAK GEMPUR PSRA','Gerak Gempur PSRA Siri 2',2,'Diterbitkan')]
    # UASA 2026 - struktur standard 5 subjek seperti UPSA/PSRA.
    uasa=con.execute("SELECT id FROM exams WHERE session_year=2026 AND exam_type='UASA'").fetchone()
    if not uasa:
        cur=con.execute("INSERT INTO exams(session_year,exam_type,exam_name,series_no,status) VALUES(?,?,?,?,?)",
                        (2026,'UASA','UASA 2026',None,'Draf'))
        configure_exam_subjects(con,cur.lastrowid,'UASA')

    # Laluan UPKK 2026: GG Siri 1 -> Percubaan -> GG Siri 2 -> UPKK sebenar.
    upkk_defaults=[
        ('GERAK GEMPUR UPKK','Gerak Gempur UPKK Siri 1',1,'Draf'),
        ('GERAK GEMPUR UPKK','Gerak Gempur UPKK Siri 2',2,'Draf'),
        ('UPKK','UPKK 2026',None,'Draf')
    ]
    for etype,ename,series,status in upkk_defaults:
        exists=con.execute('SELECT id FROM exams WHERE session_year=2026 AND exam_type=? AND COALESCE(series_no,0)=COALESCE(?,0)',(etype,series)).fetchone()
        if not exists:
            cur=con.execute('INSERT INTO exams(session_year,exam_type,exam_name,series_no,status) VALUES(?,?,?,?,?)',(2026,etype,ename,series,status))
            configure_exam_subjects(con,cur.lastrowid,etype)
            ensure_exam_grade_rules(con,cur.lastrowid)
    for et,name,series,status in demo_exams:
        if not con.execute('SELECT 1 FROM exams WHERE exam_name=?',(name,)).fetchone(): con.execute('INSERT INTO exams(session_year,exam_type,exam_name,series_no,status) VALUES(2026,?,?,?,?)',(et,name,series,status))
        e=con.execute('SELECT id FROM exams WHERE exam_name=?',(name,)).fetchone()['id']; configure_exam_subjects(con,e,et)
    teacher=con.execute("SELECT id FROM users WHERE username='guru'").fetchone()['id']
    for ex in con.execute('SELECT id FROM exams WHERE session_year=2026').fetchall():
        for s in con.execute('SELECT subject_id FROM exam_subjects WHERE exam_id=?',(ex['id'],)).fetchall():
            con.execute('INSERT OR IGNORE INTO teacher_assignments(teacher_user_id,class_id,exam_id,subject_id) VALUES(?,?,?,?)',(teacher,cid,ex['id'],s['subject_id']))
    # sample marks to make analysis immediately visible
    students=con.execute('SELECT s.id FROM students s JOIN student_class_history h ON h.student_id=s.id WHERE h.class_id=? ORDER BY s.id',(cid,)).fetchall()
    base_scores=[92,81,73,66,54,35]
    exams=con.execute("SELECT * FROM exams WHERE exam_name IN ('UPSA 2026','Gerak Gempur PSRA Siri 1','Gerak Gempur PSRA Siri 2') ORDER BY id").fetchall()
    for ex_index,e in enumerate(exams):
        subjects=con.execute('SELECT s.* FROM subjects s JOIN exam_subjects es ON es.subject_id=s.id WHERE es.exam_id=? ORDER BY s.id',(e['id'],)).fetchall()
        for sj_index,sj in enumerate(subjects):
            comps=con.execute('SELECT * FROM subject_components WHERE subject_id=? ORDER BY sort_order',(sj['id'],)).fetchall()
            for i,st in enumerate(students):
                target=min(100,max(0,base_scores[i] + ex_index*4 + (sj_index%3-1)*3))
                # distribute target percentage across components
                vals=[]; total=0
                for comp in comps:
                    val=round((target/100)*comp['max_score'],1); vals.append((comp,val)); total+=val
                total=round(total,1)
                cur=con.execute('INSERT OR IGNORE INTO marks(exam_id,student_id,subject_id,class_id_at_exam,score,updated_by) VALUES(?,?,?,?,?,?)',(e['id'],st['id'],sj['id'],cid,total,teacher))
                mark=con.execute('SELECT id FROM marks WHERE exam_id=? AND student_id=? AND subject_id=?',(e['id'],st['id'],sj['id'])).fetchone()
                if mark:
                    con.execute('UPDATE marks SET score=? WHERE id=?',(total,mark['id']))
                    for comp,val in vals: con.execute('INSERT OR REPLACE INTO component_marks(mark_id,component_id,score) VALUES(?,?,?)',(mark['id'],comp['id'],val))
    con.commit()

def login_required(role=None):
    def deco(fn):
        @wraps(fn)
        def inner(*a,**kw):
            if 'user_id' not in session:return redirect(url_for('login'))
            if role and session.get('role')!=role:abort(403)
            return fn(*a,**kw)
        return inner
    return deco

def log_action(con,action,entity_type,entity_id=None,detail=''):
    con.execute('INSERT INTO audit_log(user_id,action,entity_type,entity_id,detail) VALUES(?,?,?,?,?)',(session.get('user_id'),action,entity_type,entity_id,detail))

def grade_info(scheme_key,score):
    if score is None:return ('-',None)
    for mn,mx,g,p in GRADE_SCHEMES.get(scheme_key,[]):
        if mn<=score<=mx:return (g,p)
    return ('-',None)

def ensure_exam_grade_rules(con,eid):
    exam=con.execute('SELECT * FROM exams WHERE id=?',(eid,)).fetchone()
    if not exam or exam_family(exam['exam_type'])!='UPKK': return
    for scheme in ('UPKK70','UPKK100'):
        if not con.execute('SELECT 1 FROM exam_grade_rules WHERE exam_id=? AND scheme_key=?',(eid,scheme)).fetchone():
            for order,(mn,mx,g,p) in enumerate(GRADE_SCHEMES[scheme]):
                con.execute('INSERT INTO exam_grade_rules(exam_id,scheme_key,grade,min_score,max_score,gps_point,sort_order) VALUES(?,?,?,?,?,?,?)',(eid,scheme,g,mn,mx,p,order))

def exam_grade_info(con,eid,scheme_key,score):
    if score is None:return ('-',None)
    if scheme_key not in ('UPKK70','UPKK100'): return grade_info(scheme_key,score)
    ensure_exam_grade_rules(con,eid)
    for r in con.execute('SELECT * FROM exam_grade_rules WHERE exam_id=? AND scheme_key=? ORDER BY sort_order',(eid,scheme_key)).fetchall():
        if r['min_score']<=score<=r['max_score']: return (r['grade'],r['gps_point'])
    return ('-',None)

def exam_grade_rules_dict(con,eid,scheme_key):
    if scheme_key not in ('UPKK70','UPKK100'): return []
    ensure_exam_grade_rules(con,eid)
    return [dict(r) for r in con.execute('SELECT grade,min_score,max_score,gps_point FROM exam_grade_rules WHERE exam_id=? AND scheme_key=? ORDER BY sort_order',(eid,scheme_key)).fetchall()]


def max_grade_point(scheme_key):
    points=[p for _mn,_mx,_g,p in GRADE_SCHEMES.get(scheme_key,[]) if p is not None]
    return max(points) if points else 5

def clean_text(value):
    if value is None:return ''
    if isinstance(value,float) and value.is_integer():value=int(value)
    return str(value).strip()

def normalize_header(value):
    return ' '.join(clean_text(value).lower().replace('.', ' ').replace('_',' ').split())

def current_class_for_student(con,sid):
    return con.execute('''SELECT c.* FROM student_class_history h JOIN classes c ON c.id=h.class_id WHERE h.student_id=? AND h.is_current=1 ORDER BY h.id DESC LIMIT 1''',(sid,)).fetchone()

def move_student_to_class(con,sid,cid,start_date=None):
    current=current_class_for_student(con,sid)
    if current and current['id']==cid:return False
    con.execute('UPDATE student_class_history SET is_current=0,end_date=? WHERE student_id=? AND is_current=1',(start_date,sid))
    con.execute('INSERT INTO student_class_history(student_id,class_id,start_date,is_current) VALUES(?,?,?,1)',(sid,cid,start_date))
    return True

def exam_mark_progress(con,eid):
    exam=con.execute('SELECT * FROM exams WHERE id=?',(eid,)).fetchone()
    if not exam:return None,[],False
    classes=con.execute('SELECT * FROM classes WHERE session_year=? ORDER BY level,class_name',(exam['session_year'],)).fetchall()
    subjects=con.execute('''SELECT s.* FROM subjects s JOIN exam_subjects es ON es.subject_id=s.id WHERE es.exam_id=? ORDER BY s.id''',(eid,)).fetchall()
    rows=[]
    all_complete=True
    any_expected=False
    for cls in classes:
        total_students=con.execute('''SELECT COUNT(DISTINCT st.id) c FROM students st JOIN student_class_history h ON h.student_id=st.id WHERE h.class_id=? AND st.active=1''',(cls['id'],)).fetchone()['c']
        if not total_students: continue
        any_expected=True
        for sj in subjects:
            comp_count=con.execute('SELECT COUNT(*) c FROM subject_components WHERE subject_id=?',(sj['id'],)).fetchone()['c']
            complete=con.execute('''SELECT COUNT(*) c FROM (
                SELECT m.student_id FROM marks m
                JOIN component_marks cm ON cm.mark_id=m.id AND cm.score IS NOT NULL
                WHERE m.exam_id=? AND m.class_id_at_exam=? AND m.subject_id=?
                GROUP BY m.id,m.student_id HAVING COUNT(cm.id)=?
            )''',(eid,cls['id'],sj['id'],comp_count)).fetchone()['c'] if comp_count else 0
            teachers=con.execute('''SELECT u.full_name FROM teacher_assignments ta JOIN users u ON u.id=ta.teacher_user_id WHERE ta.class_id=? AND ta.exam_id=? AND ta.subject_id=? ORDER BY u.full_name''',(cls['id'],eid,sj['id'])).fetchall()
            teacher_names=', '.join(r['full_name'] for r in teachers) if teachers else 'Belum ditugaskan'
            if complete>=total_students: status='Lengkap'
            elif complete>0: status='Sedang Diisi'
            else: status='Belum Diisi'
            if status!='Lengkap': all_complete=False
            rows.append({'class_id':cls['id'],'class_name':cls['class_name'],'subject_id':sj['id'],'subject_name':sj['name'],'teacher_names':teacher_names,'complete':complete,'total':total_students,'status':status})
    if not any_expected: all_complete=False
    return exam,rows,all_complete

@app.route('/')
def home(): return redirect(url_for('dashboard' if 'user_id' in session else 'login'))
@app.route('/login',methods=['GET','POST'])
def login():
    if request.method=='POST':
        con=db();u=con.execute('SELECT * FROM users WHERE username=?',(request.form['username'],)).fetchone();con.close()
        if u and check_password_hash(u['password_hash'],request.form['password']): session.clear();session.update(user_id=u['id'],username=u['username'],role=u['role'],full_name=u['full_name']);return redirect(url_for('dashboard'))
        flash('Nama pengguna atau kata laluan tidak tepat.')
    return render_template('login.html')
@app.route('/logout')
def logout():session.clear();return redirect(url_for('login'))
@app.route('/dashboard')
@login_required()
def dashboard():
    con=db()
    counts={
        'students':con.execute('SELECT COUNT(*) c FROM students WHERE active=1').fetchone()['c'],
        'exams':con.execute('SELECT COUNT(*) c FROM exams').fetchone()['c'],
        'published':con.execute("SELECT COUNT(*) c FROM exams WHERE status='Diterbitkan'").fetchone()['c'],
        'teachers':con.execute("SELECT COUNT(*) c FROM users WHERE role='teacher'").fetchone()['c'],
        'parents':con.execute("SELECT COUNT(*) c FROM users WHERE role='parent'").fetchone()['c'],
    }
    recent=con.execute('SELECT * FROM exams ORDER BY session_year DESC,id DESC LIMIT 5').fetchall()
    open_exams=con.execute("SELECT COUNT(*) c FROM exams WHERE status IN ('Sedang Key-in','Semakan')").fetchone()['c']
    con.close()
    return render_template('dashboard.html',counts=counts,recent=recent,open_exams=open_exams)

@app.route('/admin/students',methods=['GET','POST'])
@login_required('admin')
def students():
    con=db()
    if request.method=='POST':
        try:
            cur=con.execute('INSERT INTO students(full_name,mykid) VALUES(?,?)',(request.form['full_name'].strip(),request.form['mykid'].strip()));sid=cur.lastrowid;cid=int(request.form['class_id']);con.execute('INSERT INTO student_class_history(student_id,class_id,start_date,is_current) VALUES(?,?,?,1)',(sid,cid,request.form.get('start_date') or None));log_action(con,'CREATE','student',sid,'Tambah murid');con.commit();flash('Murid berjaya ditambah.')
        except sqlite3.IntegrityError:flash('MyKid sudah wujud atau data tidak sah.')
    rows=con.execute('''SELECT s.*,c.class_name,c.session_year FROM students s LEFT JOIN student_class_history h ON h.student_id=s.id AND h.is_current=1 LEFT JOIN classes c ON c.id=h.class_id ORDER BY s.full_name''').fetchall();classes=con.execute('SELECT * FROM classes ORDER BY session_year DESC,level,class_name').fetchall();con.close();return render_template('students.html',students=rows,classes=classes)
@app.route('/admin/student/<int:sid>/edit',methods=['GET','POST'])
@login_required('admin')
def edit_student(sid):
    con=db()
    if request.method=='POST': con.execute('UPDATE students SET full_name=?,mykid=? WHERE id=?',(request.form['full_name'].strip(),request.form['mykid'].strip(),sid));log_action(con,'UPDATE','student',sid,'Kemaskini nama/MyKid');con.commit();flash('Maklumat murid dikemas kini.')
    student=con.execute('SELECT * FROM students WHERE id=?',(sid,)).fetchone();history=con.execute('''SELECT h.*,c.session_year,c.level,c.class_name FROM student_class_history h JOIN classes c ON c.id=h.class_id WHERE h.student_id=? ORDER BY c.session_year DESC,h.id DESC''',(sid,)).fetchall();classes=con.execute('SELECT * FROM classes ORDER BY session_year DESC,level,class_name').fetchall();con.close();return render_template('student_edit.html',student=student,history=history,classes=classes)
@app.route('/admin/student/<int:sid>/move-class',methods=['POST'])
@login_required('admin')
def move_class(sid):
    con=db();cid=int(request.form['class_id']);d=request.form.get('start_date') or None
    changed=move_student_to_class(con,sid,cid,d)
    if changed:log_action(con,'MOVE_CLASS','student',sid,f'class_id={cid}')
    con.commit();con.close();flash('Kelas semasa ditukar. Sejarah kelas lama kekal.' if changed else 'Murid sudah berada dalam kelas tersebut.');return redirect(url_for('edit_student',sid=sid))
@app.route('/admin/classes',methods=['GET','POST'])
@login_required('admin')
def classes():
    con=db()
    if request.method=='POST':con.execute('INSERT OR IGNORE INTO classes(session_year,level,class_name) VALUES(?,?,?)',(int(request.form['session_year']),int(request.form['level']),request.form['class_name'].strip()));con.commit();flash('Kelas disimpan.')
    rows=con.execute('SELECT * FROM classes ORDER BY session_year DESC,level,class_name').fetchall();con.close();return render_template('classes.html',classes=rows)
@app.route('/admin/exams',methods=['GET','POST'])
@login_required('admin')
def exams():
    con=db()
    if request.method=='POST':
        series=request.form.get('series_no') or None
        exam_type=request.form['exam_type']
        series_num=int(series) if series else None
        if exam_type in ('GERAK GEMPUR PSRA','GERAK GEMPUR UPKK') and series_num not in (1,2):
            label='Gerak Gempur PSRA' if exam_type=='GERAK GEMPUR PSRA' else 'Gerak Gempur UPKK'
            flash(f'{label} hanya dibenarkan Siri 1 atau Siri 2.')
            con.close()
            return redirect(url_for('exams'))
        cur=con.execute('INSERT INTO exams(session_year,exam_type,exam_name,series_no,status) VALUES(?,?,?,?,?)',(int(request.form['session_year']),exam_type,request.form['exam_name'].strip(),series_num,request.form['status']))
        configure_exam_subjects(con,cur.lastrowid,exam_type);ensure_exam_grade_rules(con,cur.lastrowid);con.commit();flash('Peperiksaan ditambah bersama struktur subjek automatik.')
    rows=con.execute('SELECT * FROM exams ORDER BY session_year DESC,id DESC').fetchall();con.close();return render_template('exams.html',exams=rows)

@app.route('/admin/exam/<int:eid>/grade-settings',methods=['GET','POST'])
@login_required('admin')
def exam_grade_settings(eid):
    con=db();exam=con.execute('SELECT * FROM exams WHERE id=?',(eid,)).fetchone()
    if not exam: con.close();abort(404)
    if exam_family(exam['exam_type'])!='UPKK':
        con.close();flash('Tetapan ini hanya untuk peperiksaan keluarga UPKK.');return redirect(url_for('exams'))
    ensure_exam_grade_rules(con,eid);con.commit()
    if request.method=='POST':
        if exam['status']=='Diterbitkan':
            con.close();flash('Tetapan gred dikunci kerana keputusan telah diterbitkan.');return redirect(url_for('exam_grade_settings',eid=eid))
        try:
            for scheme,full in [('UPKK70',70),('UPKK100',100)]:
                # Admin ubah minimum A/B/C sahaja. Maksimum disambung automatik
                # supaya setiap markah 0..full sentiasa mempunyai satu gred.
                amin=int(request.form[f'{scheme}_A_min'])
                bmin=int(request.form[f'{scheme}_B_min'])
                cmin=int(request.form[f'{scheme}_C_min'])
                if not (0 < cmin < bmin < amin <= full): raise ValueError
                ranges={'A':(amin,full),'B':(bmin,amin-1),'C':(cmin,bmin-1),'D':(0,cmin-1)}
                for order,g in enumerate(('A','B','C','D')):
                    mn,mx=ranges[g]
                    pt=float(request.form[f'{scheme}_{g}_point'])
                    con.execute('UPDATE exam_grade_rules SET min_score=?,max_score=?,gps_point=?,sort_order=? WHERE exam_id=? AND scheme_key=? AND grade=?',(mn,mx,pt,order,eid,scheme,g))
        except (ValueError,KeyError):
            con.close();flash('Julat gred tidak sah. Semak markah minimum dan maksimum.');return redirect(url_for('exam_grade_settings',eid=eid))
        log_action(con,'GRADE_SETTINGS','exam',eid,'Tetapan gred UPKK dikemas kini');con.commit();con.close()
        flash('Tetapan gred UPKK berjaya disimpan.');return redirect(url_for('exam_grade_settings',eid=eid))
    rules={}
    for scheme in ('UPKK70','UPKK100'):
        rules[scheme]={r['grade']:dict(r) for r in con.execute('SELECT * FROM exam_grade_rules WHERE exam_id=? AND scheme_key=?',(eid,scheme)).fetchall()}
    con.close();return render_template('exam_grade_settings.html',exam=exam,rules=rules)

@app.route('/admin/exam/<int:eid>/status',methods=['POST'])
@login_required('admin')
def exam_status(eid):
    con=db();new_status=request.form['status']
    if new_status=='Diterbitkan':
        exam,progress,ready=exam_mark_progress(con,eid)
        if not ready:
            con.close();flash('Belum boleh diterbitkan: masih ada kelas/subjek yang markahnya belum lengkap.');return redirect(request.referrer or url_for('mark_monitor'))
    con.execute('UPDATE exams SET status=? WHERE id=?',(new_status,eid));log_action(con,'STATUS','exam',eid,new_status);con.commit();con.close();flash('Status peperiksaan dikemas kini.');return redirect(request.referrer or url_for('exams'))

@app.route('/admin/mark-monitor')
@login_required('admin')
def mark_monitor():
    con=db();exams=con.execute('SELECT * FROM exams ORDER BY session_year DESC,id DESC').fetchall();summaries=[]
    for e in exams:
        _exam,rows,ready=exam_mark_progress(con,e['id'])
        total=len(rows);complete=sum(1 for r in rows if r['status']=='Lengkap');partial=sum(1 for r in rows if r['status']=='Sedang Diisi')
        summaries.append({'exam':e,'total':total,'complete':complete,'partial':partial,'ready':ready})
    con.close();return render_template('mark_monitor.html',summaries=summaries)

@app.route('/admin/mark-monitor/<int:eid>')
@login_required('admin')
def mark_monitor_exam(eid):
    con=db();exam,rows,ready=exam_mark_progress(con,eid)
    if not exam: con.close();abort(404)
    grouped={}
    for r in rows: grouped.setdefault(r['class_name'],[]).append(r)
    con.close();return render_template('mark_monitor_exam.html',exam=exam,grouped=grouped,ready=ready,rows=rows)


@app.route('/admin/teachers',methods=['GET','POST'])
@login_required('admin')
def teachers():
    con=db()
    if request.method=='POST':
        username=request.form['username'].strip();name=request.form['full_name'].strip();password=request.form.get('password','').strip() or 'guru123'
        try:
            con.execute("INSERT INTO users(username,password_hash,role,full_name) VALUES(?,?, 'teacher',?)",(username,generate_password_hash(password),name));con.commit();flash('Akaun guru berjaya ditambah.')
        except sqlite3.IntegrityError:flash('Nama pengguna guru sudah digunakan.')
    rows=con.execute("SELECT u.*,COUNT(ta.subject_id) assignment_count FROM users u LEFT JOIN teacher_assignments ta ON ta.teacher_user_id=u.id WHERE u.role='teacher' GROUP BY u.id ORDER BY u.full_name").fetchall();con.close();return render_template('teachers.html',teachers=rows)

@app.route('/admin/teacher/<int:tid>/reset-password',methods=['POST'])
@login_required('admin')
def reset_teacher_password(tid):
    password=request.form.get('password','').strip()
    if len(password)<6:flash('Kata laluan baharu mesti sekurang-kurangnya 6 aksara.');return redirect(url_for('teachers'))
    con=db();con.execute("UPDATE users SET password_hash=? WHERE id=? AND role='teacher'",(generate_password_hash(password),tid));log_action(con,'RESET_PASSWORD','teacher',tid,'Admin reset kata laluan');con.commit();con.close();flash('Kata laluan guru telah ditukar.');return redirect(url_for('teachers'))

@app.route('/admin/assignments',methods=['GET','POST'])
@login_required('admin')
def assignments():
    con=db()
    if request.method=='POST':
        tid=int(request.form['teacher_user_id']);cid=int(request.form['class_id']);eid=int(request.form['exam_id']);sid=int(request.form['subject_id'])
        valid=con.execute('''SELECT 1 FROM classes c JOIN exams e ON e.session_year=c.session_year
            JOIN exam_subjects es ON es.exam_id=e.id
            WHERE c.id=? AND e.id=? AND es.subject_id=?''',(cid,eid,sid)).fetchone()
        if not valid:
            flash('Tugasan tidak sah. Pastikan peperiksaan sepadan dengan sesi kelas dan mata pelajaran.')
        else:
            con.execute('INSERT OR IGNORE INTO teacher_assignments(teacher_user_id,class_id,exam_id,subject_id) VALUES(?,?,?,?)',(tid,cid,eid,sid))
            log_action(con,'ASSIGN','teacher',tid,f'class={cid},exam={eid},subject={sid}');con.commit();flash('Tugasan guru disimpan.')
    teachers=con.execute("SELECT * FROM users WHERE role='teacher' ORDER BY full_name").fetchall()
    classes=con.execute('SELECT * FROM classes ORDER BY session_year DESC,level,class_name').fetchall()
    exams=con.execute('SELECT * FROM exams ORDER BY session_year DESC,id DESC').fetchall()
    exam_subject_map={}
    for e in exams:
        exam_subject_map[str(e['id'])]=[dict(r) for r in con.execute('''SELECT s.id,s.name FROM subjects s JOIN exam_subjects es ON es.subject_id=s.id WHERE es.exam_id=? ORDER BY s.name''',(e['id'],)).fetchall()]
    rows=con.execute('''SELECT ta.teacher_user_id,ta.class_id,ta.exam_id,ta.subject_id,u.full_name,c.session_year,c.class_name,
        e.exam_name,e.exam_type,s.name subject_name
        FROM teacher_assignments ta JOIN users u ON u.id=ta.teacher_user_id JOIN classes c ON c.id=ta.class_id
        JOIN exams e ON e.id=ta.exam_id JOIN subjects s ON s.id=ta.subject_id
        ORDER BY c.session_year DESC,e.id DESC,c.class_name,u.full_name,s.name''').fetchall()
    con.close()
    return render_template('assignments.html',teachers=teachers,classes=classes,exams=exams,exam_subject_map=exam_subject_map,assignments=rows)

@app.route('/admin/assignments/delete',methods=['POST'])
@login_required('admin')
def assignment_delete():
    tid=int(request.form['teacher_user_id']);cid=int(request.form['class_id']);eid=int(request.form['exam_id']);sid=int(request.form['subject_id'])
    con=db();con.execute('DELETE FROM teacher_assignments WHERE teacher_user_id=? AND class_id=? AND exam_id=? AND subject_id=?',(tid,cid,eid,sid))
    log_action(con,'UNASSIGN','teacher',tid,f'class={cid},exam={eid},subject={sid}');con.commit();con.close()
    flash('Tugasan guru dibuang.');return redirect(url_for('assignments'))

@app.route('/admin/parents',methods=['GET','POST'])
@login_required('admin')
def parents():
    con=db()
    if request.method=='POST':
        username=request.form['username'].strip();name=request.form['full_name'].strip();password=request.form.get('password','').strip() or 'ibubapa123'
        try:
            cur=con.execute("INSERT INTO users(username,password_hash,role,full_name) VALUES(?,?, 'parent',?)",(username,generate_password_hash(password),name));pid=cur.lastrowid
            sid=request.form.get('student_id')
            if sid:con.execute('INSERT OR IGNORE INTO parent_students(parent_user_id,student_id) VALUES(?,?)',(pid,int(sid)))
            con.commit();flash('Akaun ibu bapa berjaya ditambah.')
        except sqlite3.IntegrityError:flash('Nama pengguna ibu bapa sudah digunakan.')
    rows=con.execute("SELECT * FROM users WHERE role='parent' ORDER BY full_name").fetchall();students=con.execute('SELECT * FROM students WHERE active=1 ORDER BY full_name').fetchall();links={}
    for p in rows:links[p['id']]=con.execute('''SELECT s.* FROM students s JOIN parent_students ps ON ps.student_id=s.id WHERE ps.parent_user_id=? ORDER BY s.full_name''',(p['id'],)).fetchall()
    con.close();return render_template('parents.html',parents=rows,students=students,links=links)

@app.route('/admin/parent/<int:pid>/link',methods=['POST'])
@login_required('admin')
def parent_link(pid):
    sid=int(request.form['student_id']);con=db();con.execute('INSERT OR IGNORE INTO parent_students(parent_user_id,student_id) VALUES(?,?)',(pid,sid));con.commit();con.close();flash('Murid dipautkan kepada akaun ibu bapa.');return redirect(url_for('parents'))

@app.route('/admin/parent/<int:pid>/unlink/<int:sid>',methods=['POST'])
@login_required('admin')
def parent_unlink(pid,sid):
    con=db();con.execute('DELETE FROM parent_students WHERE parent_user_id=? AND student_id=?',(pid,sid));con.commit();con.close();flash('Pautan murid dibuang.');return redirect(url_for('parents'))

@app.route('/admin/import-students',methods=['GET','POST'])
@login_required('admin')
def import_students():
    if request.method=='GET':return render_template('import_students.html')
    f=request.files.get('file')
    if not f or not f.filename.lower().endswith('.xlsx'):flash('Sila pilih fail Excel .xlsx.');return redirect(url_for('import_students'))
    try:
        wb=load_workbook(f,read_only=True,data_only=True);ws=wb.active
        headers={normalize_header(v):i for i,v in enumerate(next(ws.iter_rows(values_only=True)))}
        aliases={
            'name':['nama penuh','nama murid','nama'],
            'mykid':['no mykid','no. mykid','mykid','no my kid'],
            'session':['sesi','tahun sesi','session'],
            'level':['tahun','darjah','level'],
            'class':['kelas','nama kelas','class'],
            'start':['tarikh mula','mula','start date'],
        }
        idx={}
        for key,names in aliases.items():
            idx[key]=next((headers[n] for n in names if n in headers),None)
        if any(idx[k] is None for k in ('name','mykid','session','level','class')):raise ValueError('Header wajib: Nama Penuh, No MyKid, Sesi, Tahun, Kelas.')
        con=db();added=updated=moved=skipped=0
        for row in ws.iter_rows(values_only=True):
            name=clean_text(row[idx['name']]);mykid=clean_text(row[idx['mykid']]);ses=clean_text(row[idx['session']]);lev=clean_text(row[idx['level']]);cname=clean_text(row[idx['class']])
            start=clean_text(row[idx['start']]) if idx['start'] is not None and idx['start']<len(row) else ''
            if not name and not mykid:continue
            if not all((name,mykid,ses,lev,cname)):skipped+=1;continue
            try:session_year=int(float(ses));level=int(float(lev))
            except Exception:skipped+=1;continue
            con.execute('INSERT OR IGNORE INTO classes(session_year,level,class_name) VALUES(?,?,?)',(session_year,level,cname));cid=con.execute('SELECT id FROM classes WHERE session_year=? AND class_name=?',(session_year,cname)).fetchone()['id']
            st=con.execute('SELECT * FROM students WHERE mykid=?',(mykid,)).fetchone()
            if not st:
                cur=con.execute('INSERT INTO students(full_name,mykid) VALUES(?,?)',(name,mykid));sid=cur.lastrowid;con.execute('INSERT INTO student_class_history(student_id,class_id,start_date,is_current) VALUES(?,?,?,1)',(sid,cid,start or f'{session_year}-01-01'));added+=1
            else:
                sid=st['id'];con.execute('UPDATE students SET full_name=? WHERE id=?',(name,sid));updated+=1
                if move_student_to_class(con,sid,cid,start or f'{session_year}-01-01'):moved+=1
        log_action(con,'IMPORT','student',None,f'added={added},updated={updated},moved={moved},skipped={skipped}');con.commit();con.close();flash(f'Import siap: {added} murid baharu, {updated} dikemas kini, {moved} pertukaran kelas, {skipped} baris dilangkau.')
    except Exception as e:flash(f'Import gagal: {e}')
    return redirect(url_for('import_students'))

@app.route('/admin/import-template')
@login_required('admin')
def import_template():
    wb=Workbook();ws=wb.active;ws.title='Murid'
    headers=['Nama Penuh','No MyKid','Sesi','Tahun','Kelas','Tarikh Mula']
    ws.append(headers);ws.append(['CONTOH AHMAD','150101-10-1234',2026,5,'5 Al-Farabi','2026-01-01'])
    fill=PatternFill('solid',fgColor='176B4D')
    for cell in ws[1]:cell.font=Font(color='FFFFFF',bold=True);cell.fill=fill;cell.alignment=Alignment(horizontal='center')
    widths=[32,20,12,10,22,16]
    for i,w in enumerate(widths,1):ws.column_dimensions[chr(64+i)].width=w
    bio=BytesIO();wb.save(bio);bio.seek(0);return send_file(bio,as_attachment=True,download_name='template-import-murid-kislb.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/admin/promote',methods=['GET','POST'])
@login_required('admin')
def promote_students():
    con=db();classes=con.execute('SELECT * FROM classes ORDER BY session_year DESC,level,class_name').fetchall();source_id=request.args.get('source_class_id',type=int);source_students=[]
    if source_id:source_students=con.execute('''SELECT s.* FROM students s JOIN student_class_history h ON h.student_id=s.id WHERE h.class_id=? AND h.is_current=1 AND s.active=1 ORDER BY s.full_name''',(source_id,)).fetchall()
    if request.method=='POST':
        source_id=int(request.form['source_class_id']);target_id=int(request.form['target_class_id']);start=request.form.get('start_date') or None;ids=[int(x) for x in request.form.getlist('student_ids')]
        if source_id==target_id:flash('Kelas asal dan kelas baharu tidak boleh sama.')
        elif not ids:flash('Pilih sekurang-kurangnya seorang murid.')
        else:
            changed=0
            for sid in ids:
                ok=con.execute('SELECT 1 FROM student_class_history WHERE student_id=? AND class_id=? AND is_current=1',(sid,source_id)).fetchone()
                if ok and move_student_to_class(con,sid,target_id,start):changed+=1
            log_action(con,'BULK_PROMOTE','class',target_id,f'{changed} murid dari class={source_id}');con.commit();flash(f'{changed} murid berjaya dipindahkan/naik kelas.');con.close();return redirect(url_for('promote_students',source_class_id=target_id))
    con.close();return render_template('promote.html',classes=classes,source_id=source_id,students=source_students)

@app.route('/admin/audit')
@login_required('admin')
def audit():
    con=db();rows=con.execute('''SELECT a.*,u.full_name FROM audit_log a LEFT JOIN users u ON u.id=a.user_id ORDER BY a.id DESC LIMIT 300''').fetchall();con.close();return render_template('audit.html',rows=rows)

@app.route('/teacher/marks')
@login_required('teacher')
def teacher_marks():
    con=db()
    assignments=con.execute('''SELECT ta.*,c.class_name,c.session_year,s.name subject_name,s.code
        FROM teacher_assignments ta
        JOIN classes c ON c.id=ta.class_id
        JOIN subjects s ON s.id=ta.subject_id
        WHERE ta.teacher_user_id=?
        ORDER BY c.session_year DESC,ta.exam_id DESC,c.class_name,s.name''',(session['user_id'],)).fetchall()
    # Paparkan hanya peperiksaan yang benar-benar ditugaskan kepada guru ini.
    # Draf turut dipaparkan sebagai terkunci supaya guru/admin boleh nampak tugasan sudah sampai.
    exams=con.execute('''SELECT DISTINCT e.*
        FROM exams e
        JOIN teacher_assignments ta ON ta.exam_id=e.id
        JOIN classes c ON c.id=ta.class_id AND c.session_year=e.session_year
        JOIN exam_subjects es ON es.exam_id=e.id AND es.subject_id=ta.subject_id
        WHERE ta.teacher_user_id=?
        ORDER BY e.session_year DESC,e.id DESC''',(session['user_id'],)).fetchall()
    allowed={e['id']:{r['subject_id'] for r in con.execute('SELECT subject_id FROM exam_subjects WHERE exam_id=?',(e['id'],)).fetchall()} for e in exams}

    progress={}
    for e in exams:
        for a in assignments:
            if a['exam_id'] != e['id'] or a['session_year'] != e['session_year'] or a['subject_id'] not in allowed[e['id']]:
                continue
            total_students=con.execute('''SELECT COUNT(DISTINCT st.id) c
                FROM students st JOIN student_class_history h ON h.student_id=st.id
                WHERE h.class_id=? AND st.active=1''',(a['class_id'],)).fetchone()['c']
            components=con.execute('SELECT COUNT(*) c FROM subject_components WHERE subject_id=?',(a['subject_id'],)).fetchone()['c']
            filled=con.execute('''SELECT COUNT(DISTINCT m.student_id) c
                FROM marks m WHERE m.exam_id=? AND m.class_id_at_exam=? AND m.subject_id=? AND m.score IS NOT NULL''',
                (e['id'],a['class_id'],a['subject_id'])).fetchone()['c']
            complete_students=0
            if components:
                complete_students=con.execute('''SELECT COUNT(*) c FROM (
                    SELECT m.student_id
                    FROM marks m
                    JOIN component_marks cm ON cm.mark_id=m.id AND cm.score IS NOT NULL
                    WHERE m.exam_id=? AND m.class_id_at_exam=? AND m.subject_id=?
                    GROUP BY m.id,m.student_id
                    HAVING COUNT(cm.id)=?
                )''',(e['id'],a['class_id'],a['subject_id'],components)).fetchone()['c']
            if total_students and complete_students >= total_students:
                status='Lengkap'
            elif filled:
                status='Sedang Diisi'
            else:
                status='Belum Diisi'
            progress[(e['id'],a['class_id'],a['subject_id'])]={
                'status':status,'filled':complete_students,'total':total_students
            }
    con.close()
    return render_template('teacher_marks.html',assignments=assignments,exams=exams,allowed=allowed,progress=progress)

@app.route('/teacher/marks/<int:eid>/<int:cid>/<int:subid>',methods=['GET','POST'])
@login_required('teacher')
def mark_entry(eid,cid,subid):
    con=db()
    allowed=con.execute('SELECT 1 FROM teacher_assignments WHERE teacher_user_id=? AND class_id=? AND exam_id=? AND subject_id=?',
        (session['user_id'],cid,eid,subid)).fetchone() and con.execute(
        'SELECT 1 FROM exam_subjects WHERE exam_id=? AND subject_id=?',(eid,subid)).fetchone()
    if not allowed: abort(403)

    exam=con.execute('SELECT * FROM exams WHERE id=?',(eid,)).fetchone()
    subject=con.execute('SELECT * FROM subjects WHERE id=?',(subid,)).fetchone()
    cls=con.execute('SELECT * FROM classes WHERE id=?',(cid,)).fetchone()
    components=con.execute('SELECT * FROM subject_components WHERE subject_id=? ORDER BY sort_order',(subid,)).fetchall()
    if not exam or exam['status'] not in ('Sedang Key-in','Semakan'): abort(403)

    students=con.execute('''SELECT DISTINCT st.id,st.full_name,st.mykid,m.id mark_id,m.score
        FROM students st
        JOIN student_class_history h ON h.student_id=st.id
        LEFT JOIN marks m ON m.student_id=st.id AND m.exam_id=? AND m.subject_id=?
        WHERE h.class_id=? AND st.active=1 ORDER BY st.full_name''',(eid,subid,cid)).fetchall()

    if request.method=='POST':
        errors=[]
        prepared=[]
        for st in students:
            vals=[]; total=0; entered=0
            for comp in components:
                raw=request.form.get(f"c_{st['id']}_{comp['id']}",'').strip()
                try:
                    v=float(raw) if raw else None
                except ValueError:
                    errors.append(f"{st['full_name']}: markah {comp['name']} tidak sah.")
                    v=None
                if v is not None and not (0 <= v <= comp['max_score']):
                    errors.append(f"{st['full_name']}: {comp['name']} mesti 0 hingga {comp['max_score']}.")
                vals.append((comp['id'],v))
                if v is not None:
                    total += v; entered += 1
            # Jumlah hanya dianggap lengkap apabila semua komponen murid diisi.
            score=round(total,2) if entered == len(components) and components else None
            prepared.append((st,vals,score))

        if errors:
            for msg in errors[:8]: flash(msg)
            if len(errors)>8: flash(f"Dan {len(errors)-8} ralat lagi. Sila semak markah.")
            con.close()
            return redirect(request.url)

        for st,vals,score in prepared:
            # Jika semua ruang kosong dan tiada rekod lama, tidak perlu cipta markah kosong.
            if all(v is None for _,v in vals) and not st['mark_id']:
                continue
            con.execute('''INSERT INTO marks(exam_id,student_id,subject_id,class_id_at_exam,score,updated_by,updated_at)
                VALUES(?,?,?,?,?,?,CURRENT_TIMESTAMP)
                ON CONFLICT(exam_id,student_id,subject_id) DO UPDATE SET
                score=excluded.score,class_id_at_exam=excluded.class_id_at_exam,
                updated_by=excluded.updated_by,updated_at=CURRENT_TIMESTAMP''',
                (eid,st['id'],subid,cid,score,session['user_id']))
            mid=con.execute('SELECT id FROM marks WHERE exam_id=? AND student_id=? AND subject_id=?',
                (eid,st['id'],subid)).fetchone()['id']
            for compid,v in vals:
                con.execute('''INSERT INTO component_marks(mark_id,component_id,score) VALUES(?,?,?)
                    ON CONFLICT(mark_id,component_id) DO UPDATE SET score=excluded.score''',(mid,compid,v))

        log_action(con,'KEYIN_MARKS','exam',eid,f'class={cid},subject={subid}')
        con.commit()
        flash('Markah berjaya disimpan. Jumlah dan gred dikira automatik apabila semua komponen murid lengkap.')
        con.close()
        return redirect(request.url)

    values={}
    for st in students:
        if st['mark_id']:
            for r in con.execute('SELECT component_id,score FROM component_marks WHERE mark_id=?',(st['mark_id'],)).fetchall():
                values[(st['id'],r['component_id'])]=r['score']

    total_students=len(students)
    complete=0
    for st in students:
        if components and all(values.get((st['id'],c['id'])) is not None for c in components):
            complete += 1
    status='Lengkap' if total_students and complete==total_students else ('Sedang Diisi' if complete or values else 'Belum Diisi')
    grade_rules=exam_grade_rules_dict(con,exam['id'],subject['scheme_key'])
    def display_grade(scheme,score): return exam_grade_info(con,exam['id'],scheme,score)
    rendered=render_template('mark_entry.html',exam=exam,subject=subject,cls=cls,students=students,
        components=components,values=values,grade_info=display_grade,grade_rules=grade_rules,entry_status=status,
        complete_count=complete,total_students=total_students)
    con.close();return rendered

@app.route('/admin/analysis')
@login_required('admin')
def analysis_index():
    con=db();exams=con.execute('SELECT * FROM exams ORDER BY session_year DESC,id DESC').fetchall();con.close();return render_template('analysis_index.html',exams=exams)
@app.route('/admin/analysis/<int:eid>')
@login_required('admin')
def analysis_exam(eid):
    con=db();exam=con.execute('SELECT * FROM exams WHERE id=?',(eid,)).fetchone()
    if not exam:abort(404)
    subjects=con.execute('''SELECT s.* FROM subjects s JOIN exam_subjects es ON es.subject_id=s.id WHERE es.exam_id=? ORDER BY s.id''',(eid,)).fetchall();stats=[];all_points=[]
    for s in subjects:
        scores=[r['score'] for r in con.execute('SELECT score FROM marks WHERE exam_id=? AND subject_id=? AND score IS NOT NULL',(eid,s['id'])).fetchall()]
        infos=[exam_grade_info(con,eid,s['scheme_key'],x) for x in scores];points=[p for g,p in infos if p is not None];all_points.extend(points)
        dist={};
        for g,p in infos:dist[g]=dist.get(g,0)+1
        stats.append({'id':s['id'],'name':s['name'],'max_score':s['max_score'],'count':len(scores),'avg':round(sum(scores)/len(scores),2) if scores else None,'highest':max(scores) if scores else None,'lowest':min(scores) if scores else None,'gps':round(sum(points)/len(points),2) if points else None,'dist':dist})
    gps=round(sum(all_points)/len(all_points),2) if all_points else None
    # trend for same Gerak Gempur family/year, GPS overall per series
    trend=[]
    if exam['exam_type'].startswith('GERAK GEMPUR'):
        related=con.execute('SELECT * FROM exams WHERE session_year=? AND exam_type=? AND series_no IS NOT NULL ORDER BY series_no',(exam['session_year'],exam['exam_type'])).fetchall()
        for re in related:
            pts=[]
            for r in con.execute('''SELECT m.score,s.scheme_key FROM marks m JOIN subjects s ON s.id=m.subject_id WHERE m.exam_id=? AND m.score IS NOT NULL''',(re['id'],)).fetchall():
                p=exam_grade_info(con,re['id'],r['scheme_key'],r['score'])[1]
                if p is not None:pts.append(p)
            trend.append({'series':re['series_no'],'gps':round(sum(pts)/len(pts),2) if pts else None,'name':re['exam_name']})
    class_stats=[]
    classes=con.execute('''SELECT DISTINCT c.id,c.class_name,c.session_year FROM marks m JOIN classes c ON c.id=m.class_id_at_exam WHERE m.exam_id=? ORDER BY c.class_name''',(eid,)).fetchall()
    for c in classes:
        pts=[];student_ids=set()
        rows=con.execute('''SELECT m.student_id,m.score,s.scheme_key FROM marks m JOIN subjects s ON s.id=m.subject_id WHERE m.exam_id=? AND m.class_id_at_exam=? AND m.score IS NOT NULL''',(eid,c['id'])).fetchall()
        for r in rows:
            p=exam_grade_info(con,eid,r['scheme_key'],r['score'])[1]
            if p is not None:pts.append(p);student_ids.add(r['student_id'])
        class_stats.append({'name':c['class_name'],'candidates':len(student_ids),'gps':round(sum(pts)/len(pts),2) if pts else None})
    subject_trends=[]
    if trend:
        related=con.execute('SELECT * FROM exams WHERE session_year=? AND exam_type=? AND series_no IS NOT NULL ORDER BY series_no',(exam['session_year'],exam['exam_type'])).fetchall()
        for srow in subjects:
            values=[]
            for re in related:
                pts=[]
                for mr in con.execute('SELECT score FROM marks WHERE exam_id=? AND subject_id=? AND score IS NOT NULL',(re['id'],srow['id'])).fetchall():
                    p=exam_grade_info(con,re['id'],srow['scheme_key'],mr['score'])[1]
                    if p is not None:pts.append(p)
                values.append({'series':re['series_no'],'gps':round(sum(pts)/len(pts),2) if pts else None})
            subject_trends.append({'name':srow['name'],'values':values,'max_point':max_grade_point(srow['scheme_key'])})
    for item,srow in zip(stats,subjects):item['max_point']=max_grade_point(srow['scheme_key'])
    con.close();return render_template('analysis_exam.html',exam=exam,stats=stats,gps=gps,trend=trend,class_stats=class_stats,subject_trends=subject_trends)


@app.route('/admin/student/<int:sid>/profile')
@login_required('admin')
def student_profile(sid):
    con=db()
    student=con.execute('SELECT * FROM students WHERE id=?',(sid,)).fetchone()
    if not student:
        con.close();abort(404)
    current=con.execute("SELECT c.* FROM student_class_history h JOIN classes c ON c.id=h.class_id WHERE h.student_id=? AND h.is_current=1 ORDER BY h.id DESC LIMIT 1",(sid,)).fetchone()
    history=con.execute("SELECT c.session_year,c.level,c.class_name,h.start_date,h.end_date,h.is_current FROM student_class_history h JOIN classes c ON c.id=h.class_id WHERE h.student_id=? ORDER BY c.session_year DESC,h.id DESC",(sid,)).fetchall()
    exams=con.execute("""SELECT DISTINCT e.*
        FROM exams e JOIN marks m ON m.exam_id=e.id
        WHERE m.student_id=?
        ORDER BY e.session_year ASC,
        CASE
          WHEN e.exam_type='UPSA' THEN 10
          WHEN e.exam_type='GERAK GEMPUR PSRA' AND e.series_no=1 THEN 20
          WHEN e.exam_type='PERCUBAAN PSRA' THEN 30
          WHEN e.exam_type='GERAK GEMPUR PSRA' AND e.series_no=2 THEN 40
          WHEN e.exam_type='PSRA' THEN 50
          ELSE 60
        END,
        e.id ASC""",(sid,)).fetchall()
    results=[]
    for e in exams:
        rows=con.execute("SELECT s.id,s.name,s.scheme_key,s.max_score,m.score FROM marks m JOIN subjects s ON s.id=m.subject_id WHERE m.exam_id=? AND m.student_id=? ORDER BY s.id",(e['id'],sid)).fetchall()
        subjects=[];points=[];scores=[]
        for r in rows:
            g,p=exam_grade_info(con,e['id'],r['scheme_key'],r['score'])
            if p is not None: points.append(p)
            if r['score'] is not None and r['max_score']: scores.append((r['score']/r['max_score'])*100)
            subjects.append({'name':r['name'],'score':r['score'],'grade':g,'point':p})
        gps=round(sum(points)/len(points),2) if points else None
        avg=round(sum(scores)/len(scores),2) if scores else None
        results.append({'exam':e,'subjects':subjects,'gps':gps,'avg':avg})
    trend=[]
    gg=con.execute("SELECT DISTINCT e.* FROM exams e JOIN marks m ON m.exam_id=e.id WHERE m.student_id=? AND e.exam_type LIKE 'GERAK GEMPUR%' AND e.series_no IS NOT NULL ORDER BY e.session_year,e.exam_type,e.series_no",(sid,)).fetchall()
    previous={}
    for e in gg:
        rows=con.execute("SELECT m.score,s.scheme_key FROM marks m JOIN subjects s ON s.id=m.subject_id WHERE m.exam_id=? AND m.student_id=? AND m.score IS NOT NULL",(e['id'],sid)).fetchall()
        pts=[]
        for r in rows:
            pt=exam_grade_info(con,e['id'],r['scheme_key'],r['score'])[1]
            if pt is not None: pts.append(pt)
        gps=round(sum(pts)/len(pts),2) if pts else None
        key=(e['session_year'],e['exam_type'])
        prev=previous.get(key)
        change=None;status='-'
        if gps is not None and prev is not None:
            change=round(gps-prev,2)
            status='Meningkat' if gps < prev else ('Menurun' if gps > prev else 'Kekal')
        if gps is not None: previous[key]=gps
        trend.append({'exam':e,'gps':gps,'change':change,'status':status})
    con.close()
    return render_template('student_profile.html',student=student,current=current,history=history,results=results,trend=trend)

@app.route('/parent')
@login_required('parent')
def parent_home():
    con=db();students=con.execute('SELECT st.* FROM students st JOIN parent_students ps ON ps.student_id=st.id WHERE ps.parent_user_id=? ORDER BY st.full_name',(session['user_id'],)).fetchall();con.close();return render_template('parent_home.html',students=students)
@app.route('/parent/student/<int:sid>')
@login_required('parent')
def parent_student(sid):
    con=db();ok=con.execute('SELECT 1 FROM parent_students WHERE parent_user_id=? AND student_id=?',(session['user_id'],sid)).fetchone()
    if not ok:abort(403)
    student=con.execute('SELECT * FROM students WHERE id=?',(sid,)).fetchone();exams=con.execute("SELECT DISTINCT e.* FROM exams e JOIN marks m ON m.exam_id=e.id WHERE m.student_id=? AND e.status='Diterbitkan' ORDER BY e.session_year DESC,e.id DESC",(sid,)).fetchall();con.close();return render_template('parent_student.html',student=student,exams=exams)
@app.route('/slip/<int:eid>/<int:sid>')
@login_required()
def slip(eid,sid):
    con=db();exam=con.execute('SELECT * FROM exams WHERE id=?',(eid,)).fetchone();student=con.execute('SELECT * FROM students WHERE id=?',(sid,)).fetchone()
    if not exam or not student:abort(404)
    if session['role']=='parent':
        if not con.execute('SELECT 1 FROM parent_students WHERE parent_user_id=? AND student_id=?',(session['user_id'],sid)).fetchone() or exam['status']!='Diterbitkan':abort(403)
    rows=con.execute('''SELECT s.id,s.name,s.scheme_key,m.id mark_id,m.score,c.class_name FROM marks m JOIN subjects s ON s.id=m.subject_id JOIN classes c ON c.id=m.class_id_at_exam WHERE m.exam_id=? AND m.student_id=? ORDER BY s.id''',(eid,sid)).fetchall();result=[];points=[]
    for r in rows:
        g,p=exam_grade_info(con,eid,r['scheme_key'],r['score']);
        if p is not None:points.append(p)
        comps=con.execute('''SELECT sc.name,sc.max_score,cm.score FROM subject_components sc LEFT JOIN component_marks cm ON cm.component_id=sc.id AND cm.mark_id=? WHERE sc.subject_id=? ORDER BY sc.sort_order''',(r['mark_id'],r['id'])).fetchall()
        result.append({'name':r['name'],'score':r['score'],'grade':g,'point':p,'components':comps})
    # Ringkasan /500 untuk peperiksaan berstruktur 5 subjek PSRA/UASA/UPSA.
    # Termasuk Percubaan PSRA dan Gerak Gempur PSRA; UPKK kekal ikut subjek.
    exam_type_norm = (exam['exam_type'] or '').strip().upper()
    exam_name_norm = (exam['exam_name'] or '').strip().upper()
    overall_types = {'PSRA', 'UASA', 'UPSA', 'PERCUBAAN PSRA', 'GERAK GEMPUR PSRA'}
    show_overall = (
        exam_type_norm in overall_types
        or 'GERAK GEMPUR PSRA' in exam_name_norm
        or 'PERCUBAAN PSRA' in exam_name_norm
    )
    scored = [r['score'] for r in rows if r['score'] is not None]
    overall_total = round(sum(scored), 2) if show_overall and len(scored) == 5 else None
    overall_percentage = round(overall_total / 5, 2) if overall_total is not None else None
    overall_grade = grade_info('STD100', overall_percentage)[0] if overall_percentage is not None else None
    class_name=rows[0]['class_name'] if rows else '-'
    con.close()
    return render_template('slip.html',exam=exam,student=student,rows=result,class_name=class_name,show_overall=show_overall,overall_total=overall_total,overall_percentage=overall_percentage,overall_grade=overall_grade)

if __name__=='__main__':
    init_db()
    app.run(debug=False, host=os.environ.get('DEMO_HOST','127.0.0.1'), port=int(os.environ.get('PORT','5000')))
