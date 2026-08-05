from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, g
import sqlite3
import hashlib
import os
import random
import string
import json
from functools import wraps
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'spark_secret_key_2027'

DB_PATH = os.path.join(os.path.dirname(__file__), 'instance', 'spark.db')

# ─── Database ────────────────────────────────────────────────────────────────

def get_db():
    if 'db' not in g:
        conn = sqlite3.connect(DB_PATH, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute('PRAGMA journal_mode=WAL')
        conn.execute('PRAGMA busy_timeout=5000')
        g.db = conn
    return g.db

@app.teardown_appcontext
def close_db(error=None):
    db = g.pop('db', None)
    if db is not None:
        db.close()

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'teacher', 'student')),
            program TEXT,
            year_level TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS programs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_code TEXT UNIQUE NOT NULL,
            subject_code TEXT NOT NULL,
            subject_name TEXT NOT NULL,
            block_name TEXT NOT NULL,
            program TEXT NOT NULL,
            year_level TEXT NOT NULL,
            teacher_id INTEGER NOT NULL,
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (teacher_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS class_enrollments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            joined_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (class_id) REFERENCES classes(id),
            FOREIGN KEY (student_id) REFERENCES users(id),
            UNIQUE(class_id, student_id)
        );

        CREATE TABLE IF NOT EXISTS exams (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            class_id INTEGER NOT NULL,
            duration_minutes INTEGER NOT NULL,
            scheduled_at TEXT,
            activated_at TEXT,
            status TEXT DEFAULT 'upcoming' CHECK(status IN ('upcoming','active','completed')),
            show_results INTEGER DEFAULT 1,
            randomize_questions INTEGER DEFAULT 0,
            tab_switch_limit INTEGER DEFAULT 3,
            tab_switch_enabled INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (class_id) REFERENCES classes(id)
        );

        CREATE TABLE IF NOT EXISTS sections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            section_type TEXT NOT NULL CHECK(section_type IN ('multiple_choice','short_answer')),
            order_index INTEGER DEFAULT 0,
            FOREIGN KEY (exam_id) REFERENCES exams(id)
        );

        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL,
            section_id INTEGER,
            question_text TEXT NOT NULL,
            question_type TEXT NOT NULL CHECK(question_type IN ('multiple_choice','short_answer')),
            points INTEGER DEFAULT 1,
            correct_answer TEXT,
            order_index INTEGER DEFAULT 0,
            FOREIGN KEY (exam_id) REFERENCES exams(id),
            FOREIGN KEY (section_id) REFERENCES sections(id)
        );

        CREATE TABLE IF NOT EXISTS choices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question_id INTEGER NOT NULL,
            choice_label TEXT NOT NULL,
            choice_text TEXT NOT NULL,
            FOREIGN KEY (question_id) REFERENCES questions(id)
        );

        CREATE TABLE IF NOT EXISTS exam_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            started_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            submitted_at TIMESTAMP,
            status TEXT DEFAULT 'ongoing' CHECK(status IN ('ongoing','submitted','terminated')),
            score REAL,
            total_points INTEGER,
            tab_switch_count INTEGER DEFAULT 0,
            question_order TEXT,
            last_seen TIMESTAMP,
            FOREIGN KEY (exam_id) REFERENCES exams(id),
            FOREIGN KEY (student_id) REFERENCES users(id),
            UNIQUE(exam_id, student_id)
        );

        CREATE TABLE IF NOT EXISTS answers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            question_id INTEGER NOT NULL,
            answer_text TEXT,
            FOREIGN KEY (session_id) REFERENCES exam_sessions(id),
            FOREIGN KEY (question_id) REFERENCES questions(id),
            UNIQUE(session_id, question_id)
        );

        CREATE TABLE IF NOT EXISTS suspicious_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            exam_id INTEGER NOT NULL,
            event_type TEXT NOT NULL,
            logged_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (session_id) REFERENCES exam_sessions(id)
        );

        CREATE TABLE IF NOT EXISTS login_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            email TEXT,
            success INTEGER DEFAULT 1,
            ip_address TEXT,
            logged_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS question_bank_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            description TEXT,
            class_id INTEGER REFERENCES classes(id),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (teacher_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS allowed_student_emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            added_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS allowed_teacher_emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            added_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')

    # Default programs
    cursor.execute("INSERT OR IGNORE INTO programs (code, name) VALUES ('BSIT', 'Bachelor of Science in Information Technology')")
    cursor.execute("INSERT OR IGNORE INTO programs (code, name) VALUES ('BSCS', 'Bachelor of Science in Computer Science')")

    # Default admin
    admin_pw = hashlib.sha256('admin123'.encode()).hexdigest()
    cursor.execute('''
        INSERT OR IGNORE INTO users (full_name, email, password, role)
        VALUES (?, ?, ?, ?)
    ''', ('Administrator', 'admin@spark.edu', admin_pw, 'admin'))

    # Migration: add tab_switch_enabled if it doesn't exist yet
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN tab_switch_enabled INTEGER DEFAULT 1')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add activated_at if it doesn't exist yet
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN activated_at TEXT')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add manually_closed flag to prevent auto-scheduler from re-opening
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN manually_closed INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add bank_group_id to questions for question bank grouping
    try:
        conn.execute('ALTER TABLE questions ADD COLUMN bank_group_id INTEGER REFERENCES question_bank_groups(id)')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add bank_group_id to exams so each exam auto-links to its group
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN bank_group_id INTEGER REFERENCES question_bank_groups(id)')
    except: pass
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN passing_score INTEGER DEFAULT 75')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add exam_code for student code-based access
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN exam_code TEXT')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add fullscreen_required for fullscreen-mode enforcement
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN fullscreen_required INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Back-fill exam_code for any exams that don't have one yet
    try:
        import random, string
        exams_without_code = conn.execute(
            'SELECT id FROM exams WHERE exam_code IS NULL OR exam_code = ""'
        ).fetchall()
        for row in exams_without_code:
            code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
            conn.execute('UPDATE exams SET exam_code=? WHERE id=?', (code, row['id']))
        if exams_without_code:
            conn.commit()
    except Exception:
        pass

    # Migration: add class_id to question_bank_groups so groups can be tagged to a class
    try:
        conn.execute('ALTER TABLE question_bank_groups ADD COLUMN class_id INTEGER REFERENCES classes(id)')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add is_bank_only flag so questions can live in the bank without belonging to an exam
    try:
        conn.execute('ALTER TABLE questions ADD COLUMN is_bank_only INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add teacher_id to questions so bank-only questions are owned without an exam
    try:
        conn.execute('ALTER TABLE questions ADD COLUMN teacher_id INTEGER REFERENCES users(id)')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: make exam_id nullable on questions so bank-only questions can exist without an exam.
    # SQLite does not support ALTER COLUMN, so we recreate the table if exam_id is still NOT NULL.
    try:
        col_info = conn.execute("PRAGMA table_info(questions)").fetchall()
        exam_id_col = next((c for c in col_info if c['name'] == 'exam_id'), None)
        if exam_id_col and exam_id_col['notnull']:
            conn.execute('PRAGMA foreign_keys=OFF')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS questions_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    exam_id INTEGER,
                    section_id INTEGER,
                    question_text TEXT NOT NULL,
                    question_type TEXT NOT NULL CHECK(question_type IN ('multiple_choice','short_answer')),
                    points INTEGER DEFAULT 1,
                    correct_answer TEXT,
                    order_index INTEGER DEFAULT 0,
                    bank_group_id INTEGER REFERENCES question_bank_groups(id),
                    is_bank_only INTEGER DEFAULT 0,
                    teacher_id INTEGER REFERENCES users(id),
                    FOREIGN KEY (exam_id) REFERENCES exams(id),
                    FOREIGN KEY (section_id) REFERENCES sections(id)
                )
            ''')
            conn.execute('''
                INSERT INTO questions_new
                    (id, exam_id, section_id, question_text, question_type, points,
                     correct_answer, order_index, bank_group_id, is_bank_only, teacher_id)
                SELECT id, exam_id, section_id, question_text, question_type, points,
                       correct_answer, order_index,
                       bank_group_id,
                       COALESCE(is_bank_only, 0),
                       teacher_id
                FROM questions
            ''')
            conn.execute('DROP TABLE questions')
            conn.execute('ALTER TABLE questions_new RENAME TO questions')
            conn.execute('PRAGMA foreign_keys=ON')
            conn.commit()
    except Exception as e:
        conn.execute('PRAGMA foreign_keys=ON')
        pass  # Already nullable or migration failed gracefully

    conn.commit()

    # Migration: add question_order to exam_sessions for persistent shuffle per student
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN question_order TEXT')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add last_seen to exam_sessions for connection tracking
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN last_seen TIMESTAMP')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add consent tracking to exam_sessions (User Consent & Data Privacy)
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN consent_given INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN consent_at TIMESTAMP')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add case_sensitive toggle to questions (for short-answer grading)
    try:
        conn.execute('ALTER TABLE questions ADD COLUMN case_sensitive INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: ensure answers table has UNIQUE(session_id, question_id)
    # SQLite doesn't support ADD CONSTRAINT, so we recreate the table if needed.
    try:
        indexes = conn.execute("PRAGMA index_list(answers)").fetchall()
        has_unique = any(idx['unique'] == 1 for idx in indexes
                         if 'session_id' in (conn.execute(f"PRAGMA index_info({idx['name']})").fetchall() or []))
        # Simpler check: look at table SQL
        tbl_sql = conn.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='answers'").fetchone()
        if tbl_sql and 'UNIQUE' not in (tbl_sql['sql'] or '').upper():
            conn.execute('PRAGMA foreign_keys=OFF')
            # Keep only the latest answer per (session_id, question_id)
            conn.execute('''
                CREATE TABLE answers_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id INTEGER NOT NULL,
                    question_id INTEGER NOT NULL,
                    answer_text TEXT,
                    FOREIGN KEY (session_id) REFERENCES exam_sessions(id),
                    FOREIGN KEY (question_id) REFERENCES questions(id),
                    UNIQUE(session_id, question_id)
                )
            ''')
            # Insert only the last saved answer per (session_id, question_id)
            conn.execute('''
                INSERT INTO answers_new (session_id, question_id, answer_text)
                SELECT session_id, question_id, answer_text
                FROM answers
                GROUP BY session_id, question_id
                HAVING id = MAX(id)
            ''')
            conn.execute('DROP TABLE answers')
            conn.execute('ALTER TABLE answers_new RENAME TO answers')
            conn.execute('PRAGMA foreign_keys=ON')
            conn.commit()
    except Exception as e:
        try:
            conn.execute('PRAGMA foreign_keys=ON')
            conn.commit()
        except Exception:
            pass


def auto_activate_scheduled_exams():
    """Auto-open any exams whose scheduled_at has arrived and are still 'upcoming'.
    Only auto-opens if the scheduled time is within the SAME DAY (today).
    Past-day scheduled exams are left closed — teacher must open them manually.
    scheduled_at is stored as the teacher's local time (from datetime-local input),
    so we compare against local time, not UTC.
    """
    try:
        conn = get_db()
        now = datetime.now()
        now_str = now.strftime('%Y-%m-%d %H:%M')
        today_str = now.strftime('%Y-%m-%d')
        # Auto-open: only if scheduled_at is today or earlier today (same day),
        # not if it was a past date (yesterday or older).
        conn.execute("""
            UPDATE exams
            SET status = 'active',
                activated_at = strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')
            WHERE status = 'upcoming'
              AND (manually_closed = 0 OR manually_closed IS NULL)
              AND scheduled_at IS NOT NULL
              AND scheduled_at != ''
              AND substr(replace(scheduled_at, 'T', ' '), 1, 10) = ?
              AND substr(replace(scheduled_at, 'T', ' '), 1, 16) <= ?
        """, (today_str, now_str))
        conn.commit()
    except Exception:
        pass


@app.before_request
def check_scheduled_exams():
    auto_activate_scheduled_exams()

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# ─── Auth Decorators ─────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please log in first.', 'error')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if 'user_id' not in session:
                return redirect(url_for('login'))
            if session.get('role') not in roles:
                flash('Access denied.', 'error')
                return redirect(url_for('login'))
            return f(*args, **kwargs)
        return decorated
    return decorator

# ─── Auth Routes ─────────────────────────────────────────────────────────────

@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for(f"{session['role']}_home"))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for(f"{session['role']}_home"))
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        if not email or not password:
            flash('Please fill in all fields.', 'error')
            return render_template('login.html')
        conn = get_db()
        user = conn.execute(
            'SELECT * FROM users WHERE email = ? AND password = ?',
            (email, hash_password(password))
        ).fetchone()
        ip = request.remote_addr
        if user:
            conn.execute('INSERT INTO login_logs (user_id, email, success) VALUES (?,?,1)',
                         (user['id'], email))
            conn.commit()
            session['user_id'] = user['id']
            session['role'] = user['role']
            session['full_name'] = user['full_name']
            session['email'] = user['email']
            return redirect(url_for(f"{user['role']}_home"))
        else:
            conn.execute('INSERT INTO login_logs (email, success) VALUES (?,0)', (email,))
            conn.commit()
            flash('Invalid email or password.', 'error')
    return render_template('login.html')

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if 'user_id' in session:
        return redirect(url_for(f"{session['role']}_home"))
    conn = get_db()
    programs = conn.execute('SELECT * FROM programs ORDER BY code').fetchall()
    if request.method == 'POST':
        full_name = request.form.get('full_name', '').strip()
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        program = request.form.get('program', '')
        year_level = request.form.get('year_level', '')
        if not all([full_name, email, password, program, year_level]):
            flash('Please fill in all fields.', 'error')
            return render_template('signup.html', programs=programs)
        # ── Whitelist check: only enrolled student emails can register ──
        allowed = conn.execute(
            'SELECT id FROM allowed_student_emails WHERE LOWER(email) = LOWER(?)',
            (email,)
        ).fetchone()
        if not allowed:
            flash('Your email is not registered. Contact your admin.', 'error')
            return render_template('signup.html', programs=programs)
        try:
            conn = get_db()
            conn.execute('''
                INSERT INTO users (full_name, email, password, role, program, year_level)
                VALUES (?, ?, ?, 'student', ?, ?)
            ''', (full_name, email, hash_password(password), program, year_level))
            conn.commit()
            flash('Account created! You can now log in.', 'success')
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            flash('Email already exists.', 'error')
    return render_template('signup.html', programs=programs)

@app.route('/signup/teacher', methods=['GET', 'POST'])
def signup_teacher():
    if 'user_id' in session:
        return redirect(url_for(f"{session['role']}_home"))
    if request.method == 'POST':
        full_name = request.form.get('full_name', '').strip()
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        if not all([full_name, email, password]):
            flash('Please fill in all fields.', 'error')
            return render_template('signup_teacher.html')
        # ── Whitelist check: only pre-approved teacher emails can register ──
        conn = get_db()
        allowed = conn.execute(
            'SELECT id FROM allowed_teacher_emails WHERE LOWER(email) = LOWER(?)',
            (email,)
        ).fetchone()
        if not allowed:
            flash('Your email is not registered. Contact your admin.', 'error')
            return render_template('signup_teacher.html')
        try:
            conn = get_db()
            conn.execute('''
                INSERT INTO users (full_name, email, password, role)
                VALUES (?, ?, ?, 'teacher')
            ''', (full_name, email, hash_password(password)))
            conn.commit()
            flash('Teacher account created! You can now log in.', 'success')
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            flash('Email already exists.', 'error')
    return render_template('signup_teacher.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out.', 'success')
    return redirect(url_for('login'))

# ─── Student Routes ───────────────────────────────────────────────────────────

@app.route('/student')
@role_required('student')
def student_home():
    conn = get_db()
    classes = conn.execute('''
        SELECT c.*, u.full_name as teacher_name
        FROM classes c
        JOIN class_enrollments ce ON c.id = ce.class_id
        JOIN users u ON c.teacher_id = u.id
        WHERE ce.student_id = ?
    ''', (session['user_id'],)).fetchall()
    return render_template('student/classes.html', classes=classes)

@app.route('/student/join', methods=['GET', 'POST'])
@role_required('student')
def student_join_class():
    if request.method == 'POST':
        code = request.form.get('class_code', '').strip().upper()
        conn = get_db()
        cls = conn.execute(
            'SELECT * FROM classes WHERE class_code = ? AND is_active = 1', (code,)
        ).fetchone()
        if not cls:
            flash('Invalid or inactive class code.', 'error')
        else:
            try:
                conn.execute(
                    'INSERT INTO class_enrollments (class_id, student_id) VALUES (?, ?)',
                    (cls['id'], session['user_id'])
                )
                conn.commit()
                flash(f'Successfully joined {cls["subject_name"]}!', 'success')
                return redirect(url_for('student_home'))
            except sqlite3.IntegrityError:
                flash('You are already enrolled in this class.', 'error')
    return render_template('student/join_class.html')

@app.route('/student/class/<int:class_id>')
@role_required('student')
def student_class_detail(class_id):
    conn = get_db()
    cls = conn.execute('''
        SELECT c.*, u.full_name as teacher_name
        FROM classes c JOIN users u ON c.teacher_id = u.id
        WHERE c.id = ?
    ''', (class_id,)).fetchone()
    enrolled = conn.execute(
        'SELECT * FROM class_enrollments WHERE class_id=? AND student_id=?',
        (class_id, session['user_id'])
    ).fetchone()
    if not cls or not enrolled:
        flash('Class not found.', 'error')
        return redirect(url_for('student_home'))
    exams = conn.execute('''
        SELECT e.*, es.status as session_status, es.score, es.total_points, es.tab_switch_count
        FROM exams e
        LEFT JOIN exam_sessions es ON e.id = es.exam_id AND es.student_id = ?
        WHERE e.class_id = ?
        ORDER BY e.scheduled_at
    ''', (session['user_id'], class_id)).fetchall()
    return render_template('student/class_detail.html', cls=cls, exams=exams)

@app.route('/student/class/<int:class_id>/leave', methods=['POST'])
@role_required('student')
def student_leave_class(class_id):
    conn = get_db()
    enrollment = conn.execute(
        'SELECT * FROM class_enrollments WHERE class_id=? AND student_id=?',
        (class_id, session['user_id'])
    ).fetchone()
    if not enrollment:
        flash('You are not enrolled in this class.', 'error')
        return redirect(url_for('student_home'))
    cls = conn.execute('SELECT subject_name FROM classes WHERE id=?', (class_id,)).fetchone()
    conn.execute(
        'DELETE FROM class_enrollments WHERE class_id=? AND student_id=?',
        (class_id, session['user_id'])
    )
    conn.commit()
    flash(f'You have left "{cls["subject_name"]}".', 'success')
    return redirect(url_for('student_home'))

@app.route('/student/exam/<int:exam_id>/has-session')
@role_required('student')
def student_exam_has_session(exam_id):
    """Return whether this student already has an ongoing session for this exam.
    Used by the frontend to skip the code modal for returning students."""
    conn = get_db()
    sess = conn.execute(
        "SELECT status FROM exam_sessions WHERE exam_id=? AND student_id=?",
        (exam_id, session['user_id'])
    ).fetchone()
    # Any existing session (ongoing, submitted, terminated) means the student
    # has already been in this exam — skip the code modal entirely.
    return jsonify({'ongoing': sess is not None})

@app.route('/student/exam/verify-code', methods=['POST'])
@role_required('student')
def student_verify_exam_code():
    data = request.get_json(force=True, silent=True) or {}
    code = (data.get('code') or '').strip().upper()
    exam_id = data.get('exam_id')
    if not exam_id or not code:
        return jsonify({'ok': False, 'reason': 'Invalid request.'})
    conn = get_db()
    exam = conn.execute('SELECT * FROM exams WHERE id=?', (int(exam_id),)).fetchone()
    if not exam:
        return jsonify({'ok': False, 'reason': 'Exam not found.'})
    if (exam['exam_code'] or '').upper() != code:
        return jsonify({'ok': False, 'reason': 'Incorrect code. Please try again.'})
    return jsonify({'ok': True})


@app.route('/student/exam/join', methods=['GET', 'POST'])
@role_required('student')
def student_join_exam_by_code():
    if request.method == 'POST':
        code = request.form.get('exam_code', '').strip().upper()
        if not code:
            flash('Please enter an exam code.', 'error')
            return redirect(url_for('student_join_exam_by_code'))
        conn = get_db()
        exam = conn.execute(
            'SELECT * FROM exams WHERE UPPER(exam_code)=?', (code,)
        ).fetchone()
        if not exam:
            flash('Invalid exam code. Please check and try again.', 'error')
            return redirect(url_for('student_join_exam_by_code'))
        # Check student is enrolled in the exam's class
        enrolled = conn.execute(
            'SELECT 1 FROM class_enrollments WHERE class_id=? AND student_id=?',
            (exam['class_id'], session['user_id'])
        ).fetchone()
        if not enrolled:
            flash('You are not enrolled in the class for this exam.', 'error')
            return redirect(url_for('student_join_exam_by_code'))
        if exam['status'] != 'active':
            status_msg = {
                'upcoming': 'This exam has not opened yet.',
                'completed': 'This exam has already closed.'
            }.get(exam['status'], 'This exam is not currently active.')
            flash(status_msg, 'error')
            return redirect(url_for('student_join_exam_by_code'))
        # Check if already submitted/terminated
        existing = conn.execute(
            'SELECT * FROM exam_sessions WHERE exam_id=? AND student_id=?',
            (exam['id'], session['user_id'])
        ).fetchone()
        if existing and existing['status'] in ('submitted', 'terminated'):
            return redirect(url_for('student_exam_result', exam_id=exam['id']))
        return redirect(url_for('student_take_exam', exam_id=exam['id']))
    return render_template('student/join_exam_by_code.html')


@app.route('/student/exams')
@role_required('student')
def student_exams():
    conn = get_db()
    raw_exams = conn.execute('''
        SELECT e.*, c.subject_name, c.block_name, c.id as class_id,
               es.status as session_status, es.score, es.total_points
        FROM exams e
        JOIN classes c ON e.class_id = c.id
        JOIN class_enrollments ce ON c.id = ce.class_id
        LEFT JOIN exam_sessions es ON e.id = es.exam_id AND es.student_id = ?
        WHERE ce.student_id = ?
        ORDER BY e.scheduled_at
    ''', (session['user_id'], session['user_id'])).fetchall()

    # Build enriched exam list with a student-aware display_status.
    # If the student has already submitted or been terminated, the exam
    # is always shown as "completed" regardless of whether the teacher
    # later opens or closes it.
    exams = []
    for e in raw_exams:
        d = dict(e)
        if d.get('session_status') in ('submitted', 'terminated'):
            d['display_status'] = 'completed'
        else:
            d['display_status'] = d['status']  # upcoming / active
        exams.append(d)

    return render_template('student/exams.html', exams=exams)

@app.route('/student/exam/<int:exam_id>/take', methods=['GET', 'POST'])
@role_required('student')
def student_take_exam(exam_id):
    conn = get_db()
    exam = conn.execute('SELECT * FROM exams WHERE id=?', (exam_id,)).fetchone()
    if not exam or exam['status'] != 'active':
        flash('This exam is not currently active.', 'error')
        return redirect(url_for('student_exams'))

    # Check enrollment
    enrolled = conn.execute('''
        SELECT ce.* FROM class_enrollments ce
        WHERE ce.class_id = ? AND ce.student_id = ?
    ''', (exam['class_id'], session['user_id'])).fetchone()
    if not enrolled:
        flash('You are not enrolled in this class.', 'error')
        return redirect(url_for('student_exams'))

    # Check existing session
    existing = conn.execute(
        'SELECT * FROM exam_sessions WHERE exam_id=? AND student_id=?',
        (exam_id, session['user_id'])
    ).fetchone()
    if existing:
        if existing['status'] in ('submitted', 'terminated'):
            return redirect(url_for('student_exam_result', exam_id=exam_id))

    # Verify exam code on GET — skip if student already has an ongoing session
    # (they already passed the code check when they first entered, no need to re-enter)
    if request.method == 'GET' and not (existing and existing['status'] == 'ongoing'):
        submitted_code = (request.args.get('code') or '').strip().upper()
        correct_code = (exam['exam_code'] or '').upper()
        if correct_code and submitted_code != correct_code:
            flash('Incorrect exam code. Please ask your teacher for the correct code.', 'error')
            return redirect(url_for('student_exams'))

    if request.method == 'POST':
        sess_id = request.form.get('session_id', type=int)
        exam_sess = conn.execute('SELECT * FROM exam_sessions WHERE id=? AND student_id=?',
                                 (sess_id, session['user_id'])).fetchone()
        if not exam_sess or exam_sess['status'] != 'ongoing':
            flash('Invalid session.', 'error')
            return redirect(url_for('student_exams'))

        questions = conn.execute('SELECT * FROM questions WHERE exam_id=?', (exam_id,)).fetchall()
        total_points = sum(q['points'] for q in questions)
        score = 0

        for q in questions:
            ans = request.form.get(f'answer_{q["id"]}', '').strip()
            conn.execute('''
                INSERT OR REPLACE INTO answers (session_id, question_id, answer_text)
                VALUES (?, ?, ?)
            ''', (sess_id, q['id'], ans))
            if q['question_type'] == 'multiple_choice':
                if ans.upper() == (q['correct_answer'] or '').upper():
                    score += q['points']
            else:
                # Short answer: respect the per-question case-sensitivity toggle.
                # Default (case_sensitive not set / 0) keeps the original case-insensitive
                # comparison so existing questions/behavior are unaffected.
                try:
                    is_case_sensitive = bool(q['case_sensitive'])
                except (IndexError, KeyError):
                    is_case_sensitive = False
                if is_case_sensitive:
                    if ans.strip() == (q['correct_answer'] or '').strip():
                        score += q['points']
                elif ans.lower() == (q['correct_answer'] or '').lower():
                    score += q['points']

        conn.execute('''
            UPDATE exam_sessions SET status='submitted', submitted_at=CURRENT_TIMESTAMP,
            score=?, total_points=? WHERE id=?
        ''', (score, total_points, sess_id))
        conn.commit()
        return redirect(url_for('student_exam_result', exam_id=exam_id))

    # Create or get session
    if not existing:
        conn.execute(
            'INSERT INTO exam_sessions (exam_id, student_id) VALUES (?, ?)',
            (exam_id, session['user_id'])
        )
        conn.commit()
        existing = conn.execute(
            'SELECT * FROM exam_sessions WHERE exam_id=? AND student_id=?',
            (exam_id, session['user_id'])
        ).fetchone()

    sections = conn.execute(
        'SELECT * FROM sections WHERE exam_id=? ORDER BY order_index', (exam_id,)
    ).fetchall()

    sections_data = []
    all_questions = []

    # Load persisted question order (so resuming preserves the same shuffle)
    saved_order = None
    if existing and existing['question_order']:
        try:
            saved_order = json.loads(existing['question_order'])
        except Exception:
            saved_order = None

    def apply_order(q_list, order_ids):
        id_to_q = {q['id']: q for q in q_list}
        ordered = [id_to_q[qid] for qid in order_ids if qid in id_to_q]
        extras = [q for q in q_list if q['id'] not in set(order_ids)]
        return ordered + extras

    if sections:
        for sec in sections:
            qs = conn.execute(
                'SELECT * FROM questions WHERE section_id=? ORDER BY order_index', (sec['id'],)
            ).fetchall()
            q_list = [dict(q) for q in qs]
            if exam['randomize_questions']:
                if saved_order is not None:
                    q_list = apply_order(q_list, saved_order)
                else:
                    random.shuffle(q_list)
            for q in q_list:
                if q['question_type'] == 'multiple_choice':
                    choices = conn.execute('SELECT * FROM choices WHERE question_id=?', (q['id'],)).fetchall()
                    q['choices'] = [dict(c) for c in choices]
                else:
                    q['choices'] = []
                all_questions.append(q)
            sections_data.append({'title': sec['title'], 'questions': q_list})
    else:
        # No sections — treat all questions as one page
        qs = conn.execute(
            'SELECT q.* FROM questions q WHERE q.exam_id=? ORDER BY q.order_index', (exam_id,)
        ).fetchall()
        q_list = [dict(q) for q in qs]
        if exam['randomize_questions']:
            if saved_order is not None:
                q_list = apply_order(q_list, saved_order)
            else:
                random.shuffle(q_list)
        for q in q_list:
            if q['question_type'] == 'multiple_choice':
                choices = conn.execute('SELECT * FROM choices WHERE question_id=?', (q['id'],)).fetchall()
                q['choices'] = [dict(c) for c in choices]
            else:
                q['choices'] = []
            all_questions.append(q)
        sections_data.append({'title': None, 'questions': q_list})

    # Persist the question order for this session (first time only)
    if exam['randomize_questions'] and saved_order is None and existing:
        new_order = json.dumps([q['id'] for q in all_questions])
        conn.execute('UPDATE exam_sessions SET question_order=? WHERE id=?',
                     (new_order, existing['id']))
        conn.commit()

    # Get saved answers — UNIQUE(session_id, question_id) guarantees one row per question
    saved = conn.execute(
        'SELECT question_id, answer_text FROM answers WHERE session_id=?',
        (existing['id'],)
    ).fetchall()
    saved_map = {a['question_id']: a['answer_text'] for a in saved}

    # Timer is based on when the exam was opened (activated_at), not when the
    # individual student started. This gives all students the same shared clock.
    # activated_at is stored in local time, so compare with datetime.now().
    total_seconds = exam['duration_minutes'] * 60
    activated_at = exam['activated_at']
    if activated_at:
        if isinstance(activated_at, str):
            try:
                activated_at = datetime.strptime(activated_at, '%Y-%m-%d %H:%M:%S')
            except ValueError:
                activated_at = datetime.strptime(activated_at, '%Y-%m-%d %H:%M')
        elapsed = int((datetime.now() - activated_at).total_seconds())
    else:
        # Fallback: use the student's own started_at if activated_at is missing
        started_at = existing['started_at']
        if isinstance(started_at, str):
            started_at = datetime.strptime(started_at, '%Y-%m-%d %H:%M:%S')
        elapsed = int((datetime.now() - started_at).total_seconds())
    time_remaining_seconds = max(0, total_seconds - elapsed)

    return render_template('student/take_exam.html', exam=exam, sections_data=sections_data,
                           exam_session=existing, saved_map=saved_map,
                           time_remaining_seconds=time_remaining_seconds)

@app.route('/student/exam/<int:exam_id>/result')
@role_required('student')
def student_exam_result(exam_id):
    conn = get_db()
    exam = conn.execute('SELECT * FROM exams WHERE id=?', (exam_id,)).fetchone()
    exam_sess = conn.execute(
        'SELECT * FROM exam_sessions WHERE exam_id=? AND student_id=?',
        (exam_id, session['user_id'])
    ).fetchone()
    if not exam_sess:
        return redirect(url_for('student_exams'))

    # Terminated students cannot view the result page — redirect them away
    if exam_sess['status'] == 'terminated':
        return redirect(url_for('student_class_detail', class_id=exam['class_id']))

    questions = []
    # If teacher enabled "Show results to students", show full answer review immediately after submit.
    # Otherwise show score only.
    exam_closed = exam['show_results'] == 1 if exam['show_results'] is not None else False
    if exam_closed and exam_sess['status'] in ('submitted', 'terminated'):
        qs = conn.execute('''
            SELECT q.*, s.title as section_title
            FROM questions q LEFT JOIN sections s ON q.section_id = s.id
            WHERE q.exam_id = ? ORDER BY s.order_index, q.order_index
        ''', (exam_id,)).fetchall()
        # Reorder to match the student's randomized order if exam was shuffled
        if exam['randomize_questions'] and exam_sess['question_order']:
            try:
                saved_order = json.loads(exam_sess['question_order'])
                qs_map = {q['id']: q for q in qs}
                qs = [qs_map[qid] for qid in saved_order if qid in qs_map]
            except Exception:
                pass  # Fall back to default order on error
        for i, q in enumerate(qs, 1):
            qd = dict(q)
            qd['original_number'] = i
            ans = conn.execute('SELECT * FROM answers WHERE session_id=? AND question_id=?',
                               (exam_sess['id'], q['id'])).fetchone()
            qd['student_answer'] = ans['answer_text'] if ans else ''
            if q['question_type'] == 'multiple_choice':
                choices = conn.execute('SELECT * FROM choices WHERE question_id=?', (q['id'],)).fetchall()
                qd['choices'] = [dict(c) for c in choices]
            else:
                qd['choices'] = []
            is_correct = False
            if q['question_type'] == 'multiple_choice':
                is_correct = (qd['student_answer'] or '').upper() == (q['correct_answer'] or '').upper()
            else:
                is_correct = (qd['student_answer'] or '').lower() == (q['correct_answer'] or '').lower()
            qd['is_correct'] = is_correct
            questions.append(qd)

    return render_template('student/exam_result.html', exam=exam, exam_session=exam_sess, questions=questions, class_id=exam['class_id'], exam_closed=exam_closed)

@app.route('/student/profile')
@role_required('student')
def student_profile():
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
    exam_history = conn.execute('''
        SELECT es.*, e.title as exam_title, e.status as exam_status,
               c.subject_name, c.block_name, e.passing_score,
               CASE WHEN es.total_points > 0
                    THEN ROUND(es.score * 100.0 / es.total_points, 1)
                    ELSE 0 END as percentage
        FROM exam_sessions es
        JOIN exams e ON es.exam_id = e.id
        JOIN classes c ON e.class_id = c.id
        WHERE es.student_id = ?
        ORDER BY es.started_at DESC
    ''', (session['user_id'],)).fetchall()
    # Current enrolled classes
    enrolled_classes = conn.execute('''
        SELECT c.*, u.full_name as teacher_name,
               COUNT(DISTINCT ce2.student_id) as classmate_count,
               COUNT(DISTINCT e.id) as exam_count
        FROM class_enrollments ce
        JOIN classes c ON ce.class_id = c.id
        JOIN users u ON c.teacher_id = u.id
        LEFT JOIN class_enrollments ce2 ON c.id = ce2.class_id
        LEFT JOIN exams e ON c.id = e.class_id
        WHERE ce.student_id = ?
        GROUP BY c.id
        ORDER BY c.is_active DESC, c.created_at DESC
    ''', (session['user_id'],)).fetchall()
    return render_template('student/profile.html', user=user, exam_history=exam_history, enrolled_classes=enrolled_classes)

# ─── Teacher Routes ───────────────────────────────────────────────────────────

@app.route('/teacher')
@role_required('teacher')
def teacher_home():
    conn = get_db()
    classes = conn.execute('''
        SELECT c.*, COUNT(ce.student_id) as student_count
        FROM classes c
        LEFT JOIN class_enrollments ce ON c.id = ce.class_id
        WHERE c.teacher_id = ?
        GROUP BY c.id
    ''', (session['user_id'],)).fetchall()
    return render_template('teacher/classes.html', classes=classes)

@app.route('/teacher/class/create', methods=['GET', 'POST'])
@role_required('teacher')
def teacher_create_class():
    conn = get_db()
    programs = conn.execute('SELECT * FROM programs ORDER BY code').fetchall()
    if request.method == 'POST':
        subject_code = request.form.get('subject_code', '').strip().upper()
        subject_name = request.form.get('subject_name', '').strip()
        block_name   = request.form.get('block_name', '').strip().upper()
        program      = request.form.get('program', '').strip()
        year_level   = request.form.get('year_level', '').strip()
        if not all([subject_code, subject_name, block_name, program, year_level]):
            flash('Please fill in all fields.', 'error')
        else:
            code = f"{subject_code}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=5))}"
            conn = get_db()
            conn.execute('''
                INSERT INTO classes (class_code, subject_code, subject_name, block_name, program, year_level, teacher_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (code, subject_code, subject_name, block_name, program, year_level, session['user_id']))
            conn.commit()
            flash(f'Class created! Code: {code}', 'success')
            return redirect(url_for('teacher_home'))
    return render_template('teacher/create_class.html', programs=programs)

@app.route('/teacher/class/<int:class_id>')
@role_required('teacher')
def teacher_class_detail(class_id):
    conn = get_db()
    cls = conn.execute(
        'SELECT * FROM classes WHERE id = ? AND teacher_id = ?',
        (class_id, session['user_id'])
    ).fetchone()
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    students = conn.execute('''
        SELECT u.* FROM users u
        JOIN class_enrollments ce ON u.id = ce.student_id
        WHERE ce.class_id = ?
    ''', (class_id,)).fetchall()
    exams_raw = conn.execute(
        'SELECT * FROM exams WHERE class_id = ? ORDER BY scheduled_at',
        (class_id,)
    ).fetchall()

    total_students = len(students)
    exams = []
    for exam in exams_raw:
        ed = dict(exam)
        stats = conn.execute('''
            SELECT COUNT(*) as sub_count,
                   AVG(score) as avg_score,
                   MAX(total_points) as total_pts,
                   SUM(CASE WHEN score IS NOT NULL AND total_points > 0 AND (score * 100.0 / total_points) >= ? THEN 1 ELSE 0 END) as pass_count
            FROM exam_sessions
            WHERE exam_id = ? AND status = 'submitted'
        ''', (exam['passing_score'] if exam['passing_score'] is not None else 75, exam['id'],)).fetchone()
        ed['submission_count'] = stats['sub_count'] or 0
        ed['total_pts'] = stats['total_pts'] or 0
        if stats['avg_score'] is not None and stats['total_pts']:
            ed['avg_score'] = round(stats['avg_score'], 1)
            ed['avg_pct'] = int(stats['avg_score'] / stats['total_pts'] * 100)
            ed['pass_count'] = stats['pass_count'] or 0
            ed['fail_count'] = (stats['sub_count'] or 0) - (stats['pass_count'] or 0)
        else:
            ed['avg_score'] = None
            ed['avg_pct'] = 0
            ed['pass_count'] = 0
            ed['fail_count'] = 0
        exams.append(ed)
    return render_template('teacher/class_detail.html', cls=cls, students=students, exams=exams)

@app.route('/teacher/class/<int:class_id>/delete', methods=['POST'])
@role_required('teacher')
def teacher_delete_class(class_id):
    conn = get_db()
    cls = conn.execute(
        'SELECT * FROM classes WHERE id = ? AND teacher_id = ?',
        (class_id, session['user_id'])
    ).fetchone()
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    # Delete all related data
    exam_ids = [r['id'] for r in conn.execute('SELECT id FROM exams WHERE class_id=?', (class_id,)).fetchall()]
    for eid in exam_ids:
        section_ids = [r['id'] for r in conn.execute('SELECT id FROM sections WHERE exam_id=?', (eid,)).fetchall()]
        for sid in section_ids:
            conn.execute('DELETE FROM choices WHERE question_id IN (SELECT id FROM questions WHERE section_id=?)', (sid,))
            conn.execute('DELETE FROM questions WHERE section_id=?', (sid,))
        conn.execute('DELETE FROM sections WHERE exam_id=?', (eid,))
        conn.execute('DELETE FROM exam_sessions WHERE exam_id=?', (eid,))
        conn.execute('DELETE FROM exams WHERE id=?', (eid,))
    conn.execute('DELETE FROM class_enrollments WHERE class_id=?', (class_id,))
    conn.execute('DELETE FROM classes WHERE id=?', (class_id,))
    conn.commit()
    flash(f'Class "{cls["subject_name"]}" has been deleted.', 'success')
    return redirect(url_for('teacher_home'))

@app.route('/teacher/class/<int:class_id>/create-exam', methods=['GET', 'POST'])
@role_required('teacher')
def teacher_create_exam(class_id):
    conn = get_db()
    cls = conn.execute(
        'SELECT * FROM classes WHERE id=? AND teacher_id=?',
        (class_id, session['user_id'])
    ).fetchone()
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    if request.method == 'POST':
        title = request.form.get('title', '').strip()
        duration = request.form.get('duration_minutes', type=int)
        scheduled_at = request.form.get('scheduled_at', '').strip()
        if scheduled_at:
            scheduled_at = scheduled_at.replace('T', ' ')  # normalize datetime-local format
        show_results = 1 if request.form.get('show_results') else 0
        randomize = 1 if request.form.get('randomize_questions') else 0
        tab_switch_enabled = 1 if request.form.get('tab_switch_enabled') else 0
        tab_limit = request.form.get('tab_switch_limit', 3, type=int) if tab_switch_enabled else 0
        passing_score = request.form.get('passing_score', 75, type=int)
        if not title or not duration:
            flash('Title and duration are required.', 'error')
            return render_template('teacher/create_exam.html', cls=cls)
        exam_code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
        cur = conn.execute('''
            INSERT INTO exams (title, class_id, duration_minutes, scheduled_at, show_results, randomize_questions, tab_switch_limit, tab_switch_enabled, passing_score, exam_code)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (title, class_id, duration, scheduled_at or None, show_results, randomize, tab_limit, tab_switch_enabled, passing_score, exam_code))
        exam_id = cur.lastrowid

        # Auto-create a question bank group named after this exam so questions
        # added via the exam page are automatically organised by exam/class.
        grp_cur = conn.execute(
            'INSERT INTO question_bank_groups (teacher_id, name, description, class_id) VALUES (?,?,?,?)',
            (session['user_id'], title, f'Auto-created group for exam: {title}', class_id)
        )
        group_id = grp_cur.lastrowid
        conn.execute('UPDATE exams SET bank_group_id=? WHERE id=?', (group_id, exam_id))

        conn.commit()
        flash('Exam created!', 'success')
        return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    return render_template('teacher/create_exam.html', cls=cls)

@app.route('/teacher/exam/<int:exam_id>')
@role_required('teacher')
def teacher_exam_detail(exam_id):
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.subject_name, c.block_name, c.teacher_id
        FROM exams e JOIN classes c ON e.class_id = c.id
        WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam["teacher_id"]) != int(session["user_id"]):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))
    sections = conn.execute('SELECT * FROM sections WHERE exam_id=? ORDER BY order_index', (exam_id,)).fetchall()
    section_data = []
    for sec in sections:
        qs = conn.execute('SELECT * FROM questions WHERE section_id=? ORDER BY order_index', (sec['id'],)).fetchall()
        q_list = []
        for q in qs:
            qd = dict(q)
            if q['question_type'] == 'multiple_choice':
                choices = conn.execute('SELECT * FROM choices WHERE question_id=?', (q['id'],)).fetchall()
                qd['choices'] = [dict(c) for c in choices]
            else:
                qd['choices'] = []
            q_list.append(qd)
        section_data.append({'section': dict(sec), 'questions': q_list})
    raw_bank = conn.execute('''
        SELECT q.id, q.question_text, q.question_type, q.points, q.correct_answer,
               q.bank_group_id, q.is_bank_only,
               e.title as exam_title,
               e.bank_group_id as exam_bank_group_id,
               c.subject_name, c.block_name,
               g.name as group_name,
               gc.subject_name as group_subject, gc.block_name as group_block
        FROM questions q
        LEFT JOIN exams e ON q.exam_id = e.id
        LEFT JOIN classes c ON e.class_id = c.id
        LEFT JOIN sections s ON q.section_id = s.id
        LEFT JOIN question_bank_groups g ON q.bank_group_id = g.id
        LEFT JOIN classes gc ON g.class_id = gc.id
        WHERE (q.is_bank_only = 1 AND q.teacher_id = ?)
           OR ((q.is_bank_only IS NULL OR q.is_bank_only = 0) AND c.teacher_id = ? AND q.exam_id != ?)
        ORDER BY q.bank_group_id IS NULL ASC, COALESCE(g.name,''), COALESCE(c.subject_name,''), COALESCE(c.block_name,''), e.title
    ''', (session['user_id'], session['user_id'], exam_id)).fetchall()
    # Deduplicate: keep only the first occurrence of each (question_text, question_type) pair
    seen_bank = set()
    bank_questions = []
    for bq in raw_bank:
        key = (bq['question_text'].strip().lower(), bq['question_type'])
        if key not in seen_bank:
            seen_bank.add(key)
            bank_questions.append(dict(bq))
    # Get teacher's bank groups for the import filter (with class info)
    exam_bank_groups = [dict(g) for g in conn.execute('''
        SELECT g.*, c.subject_name, c.block_name
        FROM question_bank_groups g
        LEFT JOIN classes c ON g.class_id = c.id
        WHERE g.teacher_id=? ORDER BY g.name
    ''', (session['user_id'],)).fetchall()]
    return render_template('teacher/exam_detail.html', exam=exam, section_data=section_data,
                           bank_questions=bank_questions, exam_bank_groups=exam_bank_groups)

@app.route('/teacher/exam/<int:exam_id>/monitoring')
@role_required('teacher')
def teacher_exam_monitoring(exam_id):
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id, c.subject_name
        FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam["teacher_id"]) != int(session["user_id"]):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))
    total_q = conn.execute('SELECT COUNT(*) FROM questions WHERE exam_id=?', (exam_id,)).fetchone()[0]
    active_sessions = conn.execute('''
        SELECT es.*, u.full_name,
               COUNT(DISTINCT a.question_id) as answered,
               es.tab_switch_count
        FROM exam_sessions es
        JOIN users u ON es.student_id = u.id
        LEFT JOIN answers a ON es.id = a.session_id
        WHERE es.exam_id=? AND es.status='ongoing'
        GROUP BY es.id
    ''', (exam_id,)).fetchall()
    past_logs = conn.execute('''
        SELECT sl.*, u.full_name, sl.event_type, sl.logged_at
        FROM suspicious_logs sl
        JOIN users u ON sl.student_id = u.id
        WHERE sl.exam_id=?
        ORDER BY sl.logged_at DESC LIMIT 100
    ''', (exam_id,)).fetchall()

    # Results data for the Results tab
    results = conn.execute('''
        SELECT u.full_name, u.id as student_id, es.score, es.total_points, es.status, es.submitted_at, es.tab_switch_count
        FROM exam_sessions es JOIN users u ON es.student_id = u.id
        WHERE es.exam_id=?
        ORDER BY es.score DESC
    ''', (exam_id,)).fetchall()

    submitted_sessions = conn.execute(
        "SELECT id FROM exam_sessions WHERE exam_id=? AND status='submitted'",
        (exam_id,)).fetchall()
    total_submitted = len(submitted_sessions)
    session_ids = [r['id'] for r in submitted_sessions]

    questions_raw = conn.execute('''
        SELECT q.id, q.question_text, q.question_type, q.points, q.correct_answer,
               s.title as section_title, q.order_index, s.order_index as sec_order
        FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id = ?
        ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()

    question_stats = []
    for q in questions_raw:
        if session_ids:
            correct_count = conn.execute('''
                SELECT COUNT(*) FROM answers
                WHERE question_id=? AND session_id IN ({})
                AND LOWER(TRIM(answer_text)) = LOWER(TRIM(?))
            '''.format(','.join('?' * len(session_ids))),
            [q['id']] + session_ids + [q['correct_answer']]).fetchone()[0]
        else:
            correct_count = 0
        pct = round((correct_count / total_submitted * 100)) if total_submitted else 0
        question_stats.append({
            'question_text': q['question_text'],
            'question_type': q['question_type'],
            'section_title': q['section_title'],
            'correct_count': correct_count,
            'total': total_submitted,
            'pct': pct,
        })
    question_stats.sort(key=lambda x: x['pct'], reverse=True)
    for i, qs in enumerate(question_stats, 1):
        qs['number'] = i

    return render_template('teacher/exam_monitoring.html', exam=exam,
                           sessions=active_sessions, past_logs=past_logs, total_q=total_q,
                           results=results, question_stats=question_stats,
                           total_submitted=total_submitted)

@app.route('/teacher/exam/<int:exam_id>/results')
@role_required('teacher')
def teacher_exam_results(exam_id):
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id, c.subject_name, c.block_name
        FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam["teacher_id"]) != int(session["user_id"]):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))
    results = conn.execute('''
        SELECT u.full_name, u.id as student_id, es.score, es.total_points, es.status, es.submitted_at, es.tab_switch_count
        FROM exam_sessions es JOIN users u ON es.student_id = u.id
        WHERE es.exam_id=?
        ORDER BY es.score DESC
    ''', (exam_id,)).fetchall()

    # Question analysis — how many students answered each question correctly
    questions_raw = conn.execute('''
        SELECT q.id, q.question_text, q.question_type, q.points, q.correct_answer,
               s.title as section_title, q.order_index, s.order_index as sec_order
        FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id = ?
        ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()

    submitted_sessions = conn.execute('''
        SELECT id FROM exam_sessions WHERE exam_id=? AND status='submitted'
    ''', (exam_id,)).fetchall()
    total_submitted = len(submitted_sessions)
    session_ids = [r['id'] for r in submitted_sessions]

    question_stats = []
    for q in questions_raw:
        if session_ids:
            correct_count = conn.execute('''
                SELECT COUNT(*) FROM answers
                WHERE question_id=? AND session_id IN ({})
                AND LOWER(TRIM(answer_text)) = LOWER(TRIM(?))
            '''.format(','.join('?' * len(session_ids))),
            [q['id']] + session_ids + [q['correct_answer']]).fetchone()[0]
        else:
            correct_count = 0
        pct = round((correct_count / total_submitted * 100)) if total_submitted else 0
        question_stats.append({
            'question_text': q['question_text'],
            'question_type': q['question_type'],
            'section_title': q['section_title'],
            'correct_count': correct_count,
            'total': total_submitted,
            'pct': pct,
        })

    # Sort by correct rate descending: #1 = easiest (most got right), last = hardest
    question_stats.sort(key=lambda x: x['pct'], reverse=True)
    # Assign rank AFTER sorting so number reflects difficulty order
    for i, qs in enumerate(question_stats, 1):
        qs['number'] = i

    # Per-Section Analytics: aggregate the question stats above by exam section
    # (e.g. "Section A: Multiple Choice"), so a teacher can see which section
    # students struggled with overall, not just individual questions.
    section_order = []
    section_agg = {}
    for q in questions_raw:
        title = q['section_title'] or 'Untitled Section'
        if title not in section_agg:
            section_agg[title] = {'section_title': title, 'question_count': 0, 'pct_sum': 0}
            section_order.append(title)
        section_agg[title]['question_count'] += 1
    for qs in question_stats:
        title = qs['section_title'] or 'Untitled Section'
        section_agg[title]['pct_sum'] += qs['pct']
    section_stats = []
    for title in section_order:
        s = section_agg[title]
        avg_pct = round(s['pct_sum'] / s['question_count']) if s['question_count'] else 0
        section_stats.append({
            'section_title': title,
            'question_count': s['question_count'],
            'avg_pct': avg_pct,
        })

    return render_template('teacher/exam_results.html', exam=exam, results=results,
                           question_stats=question_stats, total_submitted=total_submitted,
                           section_stats=section_stats)

@app.route('/teacher/exam/<int:exam_id>/toggle-status', methods=['POST'])
@role_required('teacher')
def teacher_toggle_exam_status(exam_id):
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id, c.id as class_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam["teacher_id"]) != int(session["user_id"]):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))
    class_id = exam['class_id']
    # Teachers can always manually open or close, even if a schedule is set.
    # manually_closed flag prevents the auto-scheduler from re-opening a teacher-closed exam.
    new_status = 'active' if exam['status'] == 'upcoming' else 'upcoming'
    if new_status == 'active':
        # Reset activated_at fresh every time the exam is opened — timer always starts from now
        conn.execute("UPDATE exams SET status=?, manually_closed=0, activated_at=strftime('%Y-%m-%d %H:%M:%S','now','localtime') WHERE id=?", (new_status, exam_id))
    else:
        # Clear activated_at on close so the next open always gets a fresh timer
        conn.execute("UPDATE exams SET status=?, manually_closed=1, activated_at=NULL WHERE id=?", (new_status, exam_id))
    conn.commit()
    flash(f"Exam is now {'Open' if new_status == 'active' else 'Closed'}.", 'success')
    return redirect(url_for('teacher_class_detail', class_id=class_id))

@app.route('/teacher/class/<int:class_id>/monitoring')
@role_required('teacher')
def teacher_class_monitoring(class_id):
    conn = get_db()
    cls = conn.execute(
        'SELECT * FROM classes WHERE id=? AND teacher_id=?',
        (class_id, session['user_id'])
    ).fetchone()
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    # Get ALL exams for this class (active + upcoming so teacher can open/close)
    exams = conn.execute(
        "SELECT * FROM exams WHERE class_id=? ORDER BY scheduled_at", (class_id,)
    ).fetchall()
    monitoring_data = []
    for exam in exams:
        sessions = conn.execute('''
            SELECT es.id, u.full_name, es.tab_switch_count, es.status,
                   es.started_at, es.submitted_at, es.score, es.total_points
            FROM exam_sessions es JOIN users u ON es.student_id = u.id
            WHERE es.exam_id=?
            ORDER BY es.started_at DESC
        ''', (exam['id'],)).fetchall()
        monitoring_data.append({'exam': dict(exam), 'sessions': [dict(s) for s in sessions]})
    return render_template('teacher/class_monitoring.html', cls=cls, monitoring_data=monitoring_data)

@app.route('/teacher/exam/<int:exam_id>/regenerate-code', methods=['POST'])
@role_required('teacher')
def teacher_regenerate_exam_code(exam_id):
    conn = get_db()
    exam = conn.execute(
        'SELECT * FROM exams WHERE id=? AND class_id IN (SELECT id FROM classes WHERE teacher_id=?)',
        (exam_id, session['user_id'])
    ).fetchone()
    if not exam:
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))
    new_code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    conn.execute('UPDATE exams SET exam_code=? WHERE id=?', (new_code, exam_id))
    conn.commit()
    flash(f'New exam code generated: {new_code}', 'success')
    return redirect(url_for('teacher_exam_settings', exam_id=exam_id))


@app.route('/teacher/exam/<int:exam_id>/settings', methods=['GET', 'POST'])
@role_required('teacher')
def teacher_exam_settings(exam_id):
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam["teacher_id"]) != int(session["user_id"]):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'delete':
            conn.execute('DELETE FROM answers WHERE session_id IN (SELECT id FROM exam_sessions WHERE exam_id=?)', (exam_id,))
            conn.execute('DELETE FROM exam_sessions WHERE exam_id=?', (exam_id,))
            conn.execute('DELETE FROM suspicious_logs WHERE exam_id=?', (exam_id,))
            conn.execute('DELETE FROM choices WHERE question_id IN (SELECT id FROM questions WHERE exam_id=?)', (exam_id,))
            conn.execute('DELETE FROM questions WHERE exam_id=?', (exam_id,))
            conn.execute('DELETE FROM sections WHERE exam_id=?', (exam_id,))
            class_id = exam['class_id']
            conn.execute('DELETE FROM exams WHERE id=?', (exam_id,))
            conn.commit()
            flash('Exam deleted.', 'success')
            return redirect(url_for('teacher_class_detail', class_id=class_id))
        title = request.form.get('title', '').strip()
        duration = request.form.get('duration_minutes', type=int)
        scheduled_at = request.form.get('scheduled_at', '').strip()
        if scheduled_at:
            scheduled_at = scheduled_at.replace('T', ' ')  # normalize datetime-local format
        show_results = 1 if request.form.get('show_results') else 0
        randomize = 1 if request.form.get('randomize_questions') else 0
        tab_switch_enabled = 1 if request.form.get('tab_switch_enabled') else 0
        tab_limit = request.form.get('tab_switch_limit', 3, type=int) if tab_switch_enabled else 0
        fullscreen_required = 1 if request.form.get('fullscreen_required') else 0
        passing_score = request.form.get('passing_score', 75, type=int)
        status = request.form.get('status', 'upcoming')
        # Track manually_closed so the auto-scheduler doesn't re-open a teacher-closed exam
        existing_status = exam['status']
        # Determine manually_closed value based on schedule changes:
        # - Schedule cleared → reset to 0 (normal unscheduled exam)
        # - Schedule added or changed → reset to 0 (new schedule should auto-open)
        # - Schedule unchanged → keep existing value (respect teacher's last manual action)
        new_scheduled_at = scheduled_at or None
        old_scheduled_at = exam['scheduled_at'] or None
        if not new_scheduled_at:
            # Schedule cleared — reset flag
            manually_closed_val = 0
        elif new_scheduled_at != old_scheduled_at:
            # Schedule added or changed — reset flag so auto-scheduler can fire
            manually_closed_val = 0
        else:
            # Schedule unchanged — preserve existing flag
            manually_closed_val = exam['manually_closed'] if exam['manually_closed'] is not None else 0

        if status == 'active' and existing_status != 'active':
            conn.execute('''
                UPDATE exams SET title=?, duration_minutes=?, scheduled_at=?, show_results=?,
                randomize_questions=?, tab_switch_limit=?, tab_switch_enabled=?, fullscreen_required=?, status=?, passing_score=?,
                manually_closed=0, activated_at=strftime('%Y-%m-%d %H:%M:%S','now','localtime') WHERE id=?
            ''', (title, duration, new_scheduled_at, show_results, randomize, tab_limit, tab_switch_enabled, fullscreen_required, status, passing_score, exam_id))
        elif status == 'upcoming' and existing_status == 'active':
            # Clear activated_at on close so next open always gets a fresh timer
            conn.execute('''
                UPDATE exams SET title=?, duration_minutes=?, scheduled_at=?, show_results=?,
                randomize_questions=?, tab_switch_limit=?, tab_switch_enabled=?, fullscreen_required=?, status=?, passing_score=?,
                manually_closed=1, activated_at=NULL WHERE id=?
            ''', (title, duration, new_scheduled_at, show_results, randomize, tab_limit, tab_switch_enabled, fullscreen_required, status, passing_score, exam_id))
        else:
            # Status unchanged — still save all fields including manually_closed reset if schedule cleared
            conn.execute('''
                UPDATE exams SET title=?, duration_minutes=?, scheduled_at=?, show_results=?,
                randomize_questions=?, tab_switch_limit=?, tab_switch_enabled=?, fullscreen_required=?, status=?, passing_score=?,
                manually_closed=? WHERE id=?
            ''', (title, duration, new_scheduled_at, show_results, randomize, tab_limit, tab_switch_enabled, fullscreen_required, status, passing_score, manually_closed_val, exam_id))
        conn.commit()
        flash('Settings saved.', 'success')
        redirect_to = request.form.get('redirect_to')
        if redirect_to == 'detail':
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
        exam = conn.execute('''
            SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
        ''', (exam_id,)).fetchone()
    return render_template('teacher/exam_settings.html', exam=exam)

@app.route('/teacher/exam/<int:exam_id>/section/add', methods=['POST'])
@role_required('teacher')
def teacher_add_section(exam_id):
    title = request.form.get('title', '').strip()
    section_type = request.form.get('section_type', 'multiple_choice')
    if title:
        conn = get_db()
        count = conn.execute('SELECT COUNT(*) FROM sections WHERE exam_id=?', (exam_id,)).fetchone()[0]
        conn.execute('INSERT INTO sections (exam_id, title, section_type, order_index) VALUES (?,?,?,?)',
                     (exam_id, title, section_type, count))
        conn.commit()
        flash('Section added.', 'success')
    return redirect(url_for('teacher_exam_detail', exam_id=exam_id) + '#questions')

@app.route('/teacher/section/<int:section_id>/delete', methods=['POST'])
@role_required('teacher')
def teacher_delete_section(section_id):
    conn = get_db()
    sec = conn.execute('SELECT * FROM sections WHERE id=?', (section_id,)).fetchone()
    if sec:
        exam_id = sec['exam_id']
        conn.execute('DELETE FROM choices WHERE question_id IN (SELECT id FROM questions WHERE section_id=?)', (section_id,))
        conn.execute('DELETE FROM questions WHERE section_id=?', (section_id,))
        conn.execute('DELETE FROM sections WHERE id=?', (section_id,))
        conn.commit()
        return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    return redirect(url_for('teacher_home'))

@app.route('/teacher/section/<int:section_id>/import-from-bank', methods=['POST'])
@role_required('teacher')
def teacher_import_from_bank(section_id):
    conn = get_db()
    sec = conn.execute('SELECT * FROM sections WHERE id=?', (section_id,)).fetchone()
    if not sec:
        return redirect(url_for('teacher_home'))
    exam_id = sec['exam_id']
    question_ids = request.form.getlist('question_ids')
    # Get or create the exam's own bank group
    exam_row = conn.execute('SELECT e.*, c.subject_name FROM exams e JOIN classes c ON e.class_id=c.id WHERE e.id=?', (exam_id,)).fetchone()
    exam_bank_group_id = exam_row['bank_group_id'] if exam_row and exam_row['bank_group_id'] else None
    if not exam_bank_group_id and exam_row:
        grp_cur = conn.execute(
            'INSERT INTO question_bank_groups (teacher_id, name, description, class_id) VALUES (?,?,?,?)',
            (session['user_id'], exam_row['title'], f'Auto-created group for exam: {exam_row["title"]}', exam_row['class_id'])
        )
        exam_bank_group_id = grp_cur.lastrowid
        conn.execute('UPDATE exams SET bank_group_id=? WHERE id=?', (exam_bank_group_id, exam_id))
    imported = 0
    skipped = 0
    for qid in question_ids:
        src = conn.execute('SELECT * FROM questions WHERE id=?', (qid,)).fetchone()
        if not src:
            continue
        # Skip if same question text+type already exists in this section
        existing = conn.execute(
            'SELECT id FROM questions WHERE section_id=? AND question_text=? AND question_type=?',
            (section_id, src['question_text'], src['question_type'])
        ).fetchone()
        if existing:
            skipped += 1
            continue
        count = conn.execute('SELECT COUNT(*) FROM questions WHERE section_id=?', (section_id,)).fetchone()[0]
        cur = conn.execute('''
            INSERT INTO questions (exam_id, section_id, question_text, question_type, points, correct_answer, order_index, bank_group_id)
            VALUES (?,?,?,?,?,?,?,?)
        ''', (exam_id, section_id, src['question_text'], src['question_type'], src['points'], src['correct_answer'], count, exam_bank_group_id or src['bank_group_id']))
        new_qid = cur.lastrowid
        if src['question_type'] == 'multiple_choice':
            choices = conn.execute('SELECT * FROM choices WHERE question_id=?', (qid,)).fetchall()
            for c in choices:
                conn.execute('INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                             (new_qid, c['choice_label'], c['choice_text']))
        imported += 1
    conn.commit()
    msg = f'{imported} question(s) imported from bank.'
    if skipped:
        msg += f' {skipped} duplicate(s) skipped.'
    flash(msg, 'success')
    return redirect(url_for('teacher_exam_detail', exam_id=exam_id))

@app.route('/teacher/section/<int:section_id>/question/add', methods=['POST'])
@role_required('teacher')
def teacher_add_question(section_id):
    conn = get_db()
    sec = conn.execute('SELECT * FROM sections WHERE id=?', (section_id,)).fetchone()
    if not sec:
        return redirect(url_for('teacher_home'))
    exam_id = sec['exam_id']
    q_text = request.form.get('question_text', '').strip()
    q_type = sec['section_type']
    points = request.form.get('points', 1, type=int)
    correct = request.form.get('correct_answer', '').strip()
    case_sensitive = 1 if request.form.get('case_sensitive') else 0
    if q_text:
        count = conn.execute('SELECT COUNT(*) FROM questions WHERE section_id=?', (section_id,)).fetchone()[0]
        # Use form-provided group; if none, auto-assign the exam's own bank group
        # so questions added via the exam page are always grouped by exam.
        bank_group_id = request.form.get('bank_group_id') or None
        if not bank_group_id:
            exam_row = conn.execute('SELECT e.*, c.subject_name FROM exams e JOIN classes c ON e.class_id=c.id WHERE e.id=?', (exam_id,)).fetchone()
            if exam_row and exam_row['bank_group_id']:
                bank_group_id = exam_row['bank_group_id']
            elif exam_row:
                # Exam was created before auto-group migration — create one now
                grp_cur = conn.execute(
                    'INSERT INTO question_bank_groups (teacher_id, name, description, class_id) VALUES (?,?,?,?)',
                    (session['user_id'], exam_row['title'], f'Auto-created group for exam: {exam_row["title"]}', exam_row['class_id'])
                )
                bank_group_id = grp_cur.lastrowid
                conn.execute('UPDATE exams SET bank_group_id=? WHERE id=?', (bank_group_id, exam_id))
        cur = conn.execute('''
            INSERT INTO questions (exam_id, section_id, question_text, question_type, points, correct_answer, order_index, bank_group_id, case_sensitive)
            VALUES (?,?,?,?,?,?,?,?,?)
        ''', (exam_id, section_id, q_text, q_type, points, correct, count, bank_group_id, case_sensitive))
        q_id = cur.lastrowid
        if q_type == 'multiple_choice':
            for label in ['A', 'B', 'C', 'D']:
                ct = request.form.get(f'choice_{label}', '').strip()
                if ct:
                    conn.execute('INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                                 (q_id, label, ct))
        conn.commit()
        flash('Question added.', 'success')
    redirect_to = request.args.get('redirect_to') or request.form.get('redirect_to', '')
    if redirect_to == 'bank':
        return redirect(url_for('teacher_question_bank'))
    return redirect(url_for('teacher_exam_detail', exam_id=exam_id))

@app.route('/teacher/question/<int:question_id>/delete', methods=['POST'])
@role_required('teacher')
def teacher_delete_question(question_id):
    conn = get_db()
    q = conn.execute('''
        SELECT q.* FROM questions q
        LEFT JOIN exams e ON q.exam_id = e.id
        LEFT JOIN classes c ON e.class_id = c.id
        WHERE q.id=?
          AND ((q.is_bank_only=1 AND q.teacher_id=?) OR c.teacher_id=?)
    ''', (question_id, session['user_id'], session['user_id'])).fetchone()
    if q:
        exam_id = q['exam_id']
        conn.execute('DELETE FROM choices WHERE question_id=?', (question_id,))
        conn.execute('DELETE FROM questions WHERE id=?', (question_id,))
        conn.commit()
        if request.form.get('redirect_to') == 'bank':
            return redirect(url_for('teacher_question_bank'))
        if exam_id:
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
        return redirect(url_for('teacher_question_bank'))
    return redirect(url_for('teacher_home'))

@app.route('/teacher/question/<int:question_id>/edit', methods=['POST'])
@role_required('teacher')
def teacher_edit_question(question_id):
    conn = get_db()
    q = conn.execute('''
        SELECT q.* FROM questions q
        LEFT JOIN exams e ON q.exam_id = e.id
        LEFT JOIN classes c ON e.class_id = c.id
        WHERE q.id=?
          AND ((q.is_bank_only=1 AND q.teacher_id=?) OR c.teacher_id=?)
    ''', (question_id, session['user_id'], session['user_id'])).fetchone()
    if not q:
        return redirect(url_for('teacher_home'))
    exam_id = q['exam_id']
    q_text = request.form.get('question_text', '').strip()
    correct = request.form.get('correct_answer', '').strip()
    points = request.form.get('points', 1, type=int)
    bank_group_id = request.form.get('bank_group_id') or None
    case_sensitive = 1 if request.form.get('case_sensitive') else 0
    conn.execute('UPDATE questions SET question_text=?, correct_answer=?, points=?, bank_group_id=?, case_sensitive=? WHERE id=?',
                 (q_text, correct, points, bank_group_id, case_sensitive, question_id))
    if q['question_type'] == 'multiple_choice':
        conn.execute('DELETE FROM choices WHERE question_id=?', (question_id,))
        for label in ['A', 'B', 'C', 'D']:
            ct = request.form.get(f'choice_{label}', '').strip()
            if ct:
                conn.execute('INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                             (question_id, label, ct))
    conn.commit()
    flash('Question updated.', 'success')
    redirect_to = request.form.get('redirect_to', '')
    if redirect_to == 'bank' or not exam_id:
        return redirect(url_for('teacher_question_bank'))
    return redirect(url_for('teacher_exam_detail', exam_id=exam_id))

@app.route('/teacher/exams')
@role_required('teacher')
def teacher_my_exams():
    conn = get_db()
    exams = conn.execute('''
        SELECT e.*, c.subject_name, c.block_name, c.id as class_id,
               COUNT(DISTINCT es.id) as session_count
        FROM exams e
        JOIN classes c ON e.class_id = c.id
        LEFT JOIN exam_sessions es ON e.id = es.exam_id
        WHERE c.teacher_id = ?
        GROUP BY e.id
        ORDER BY e.created_at DESC
    ''', (session['user_id'],)).fetchall()
    return render_template('teacher/my_exams.html', exams=exams)

@app.route('/teacher/question-bank/groups', methods=['GET'])
@role_required('teacher')
def teacher_bank_groups():
    conn = get_db()
    groups = conn.execute('''
        SELECT g.*, c.subject_name, c.block_name, c.year_level
        FROM question_bank_groups g
        LEFT JOIN classes c ON g.class_id = c.id
        WHERE g.teacher_id=? ORDER BY g.name
    ''', (session['user_id'],)).fetchall()
    return jsonify([dict(g) for g in groups])

@app.route('/teacher/question-bank/groups/create', methods=['POST'])
@role_required('teacher')
def teacher_create_bank_group():
    conn = get_db()
    name = request.form.get('name', '').strip()
    description = request.form.get('description', '').strip()
    class_id = request.form.get('class_id') or None
    if not name:
        flash('Group name is required.', 'error')
        return redirect(url_for('teacher_question_bank'))
    # Allow same name if it belongs to a different class
    existing = conn.execute(
        'SELECT id FROM question_bank_groups WHERE teacher_id=? AND name=? AND (class_id=? OR (class_id IS NULL AND ? IS NULL))',
        (session['user_id'], name, class_id, class_id)
    ).fetchone()
    if existing:
        flash(f'A group named "{name}" already exists for that class.', 'error')
        return redirect(url_for('teacher_question_bank'))
    conn.execute(
        'INSERT INTO question_bank_groups (teacher_id, name, description, class_id) VALUES (?,?,?,?)',
        (session['user_id'], name, description, class_id)
    )
    conn.commit()
    flash(f'Group "{name}" created successfully.', 'success')
    return redirect(url_for('teacher_question_bank'))

@app.route('/teacher/question-bank/groups/<int:group_id>/delete', methods=['POST'])
@role_required('teacher')
def teacher_delete_bank_group(group_id):
    conn = get_db()
    grp = conn.execute(
        'SELECT * FROM question_bank_groups WHERE id=? AND teacher_id=?',
        (group_id, session['user_id'])
    ).fetchone()
    if not grp:
        flash('Group not found.', 'error')
        return redirect(url_for('teacher_question_bank'))
    # Unassign questions from this group
    conn.execute('UPDATE questions SET bank_group_id=NULL WHERE bank_group_id=?', (group_id,))
    conn.execute('DELETE FROM question_bank_groups WHERE id=?', (group_id,))
    conn.commit()
    flash(f'Group "{grp["name"]}" deleted. Questions are now ungrouped.', 'success')
    return redirect(url_for('teacher_question_bank'))

@app.route('/teacher/question-bank/groups/<int:group_id>/rename', methods=['POST'])
@role_required('teacher')
def teacher_rename_bank_group(group_id):
    conn = get_db()
    grp = conn.execute(
        'SELECT * FROM question_bank_groups WHERE id=? AND teacher_id=?',
        (group_id, session['user_id'])
    ).fetchone()
    if not grp:
        flash('Group not found.', 'error')
        return redirect(url_for('teacher_question_bank'))
    new_name = request.form.get('name', '').strip()
    if not new_name:
        flash('Group name cannot be empty.', 'error')
        return redirect(url_for('teacher_question_bank'))
    new_class_id = request.form.get('class_id') or None
    conn.execute('UPDATE question_bank_groups SET name=?, class_id=? WHERE id=?', (new_name, new_class_id, group_id))
    conn.commit()
    flash(f'Group updated to "{new_name}".', 'success')
    return redirect(url_for('teacher_question_bank'))

@app.route('/teacher/question/<int:question_id>/assign-group', methods=['POST'])
@role_required('teacher')
def teacher_assign_question_group(question_id):
    conn = get_db()
    group_id = request.form.get('bank_group_id') or None
    if group_id:
        grp = conn.execute(
            'SELECT id FROM question_bank_groups WHERE id=? AND teacher_id=?',
            (group_id, session['user_id'])
        ).fetchone()
        if not grp:
            flash('Invalid group.', 'error')
            return redirect(url_for('teacher_question_bank'))
    conn.execute('UPDATE questions SET bank_group_id=? WHERE id=?', (group_id, question_id))
    conn.commit()
    flash('Question group updated.', 'success')
    return redirect(url_for('teacher_question_bank'))

@app.route('/teacher/question-bank/add', methods=['POST'])
@role_required('teacher')
def teacher_bank_add_question():
    """Add a standalone bank-only question (not tied to any exam or section)."""
    conn = get_db()
    q_text = request.form.get('question_text', '').strip()
    q_type = request.form.get('question_type', '').strip()
    points = request.form.get('points', 1, type=int)
    correct = request.form.get('correct_answer', '').strip()
    bank_group_id = request.form.get('bank_group_id') or None

    if not q_text or q_type not in ('multiple_choice', 'short_answer'):
        flash('Question text and type are required.', 'error')
        return redirect(url_for('teacher_question_bank'))

    if bank_group_id:
        grp = conn.execute(
            'SELECT id FROM question_bank_groups WHERE id=? AND teacher_id=?',
            (bank_group_id, session['user_id'])
        ).fetchone()
        if not grp:
            bank_group_id = None

    cur = conn.execute(
        '''INSERT INTO questions
           (exam_id, section_id, question_text, question_type, points, correct_answer,
            order_index, bank_group_id, is_bank_only, teacher_id)
           VALUES (NULL, NULL, ?, ?, ?, ?, 0, ?, 1, ?)''',
        (q_text, q_type, points, correct, bank_group_id, session['user_id'])
    )
    q_id = cur.lastrowid
    if q_type == 'multiple_choice':
        for label in ['A', 'B', 'C', 'D']:
            ct = request.form.get(f'choice_{label}', '').strip()
            if ct:
                conn.execute(
                    'INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                    (q_id, label, ct)
                )
    conn.commit()
    flash('Question added to bank.', 'success')
    return redirect(url_for('teacher_question_bank'))

@app.route('/teacher/question-bank/import', methods=['POST'])
@role_required('teacher')
def teacher_bank_import_file():
    """Parse an uploaded .json file and bulk-add questions to the bank."""
    import re as _re
    conn = get_db()
    uploaded = request.files.get('import_file')
    if not uploaded or uploaded.filename == '':
        flash('No file selected.', 'error')
        return redirect(url_for('teacher_question_bank'))

    # Derive group name from filename (strip extension)
    raw_name = uploaded.filename
    group_name = _re.sub(r'\.[^.]+$', '', raw_name).strip()
    # Replace underscores/dashes with spaces for readability
    group_name = _re.sub(r'[_\-]+', ' ', group_name).strip() or 'Imported Questions'

    # Read file content
    try:
        content = uploaded.read().decode('utf-8', errors='replace')
    except Exception:
        flash('Could not read file. Make sure it is a valid (.json) file.', 'error')
        return redirect(url_for('teacher_question_bank'))

    # ── Parse questions ──────────────────────────────────────────────────────
    # Supported formats:
    #
    # Multiple choice:
    #   1. Which loop repeats a block of code?
    #   a. if
    #   b. for
    #   c. print
    #   d. input
    #   Answer: b
    #
    # Short answer:
    #   1. What is the brain of the computer?
    #   Answer: CPU
    # ─────────────────────────────────────────────────────────────────────────
    lines = [l.rstrip() for l in content.splitlines()]

    # ── Clean parser: split file into question blocks first, then parse each ──
    # A new block starts whenever we see a numbered line: "1.", "2.", "3)" etc.
    blocks = []
    current_block = []
    for line in lines:
        stripped = line.strip()
        if _re.match(r'^\d+[\.\)]\s+', stripped) and current_block:
            blocks.append(current_block)
            current_block = [stripped]
        elif stripped or current_block:
            current_block.append(stripped)
    if current_block:
        blocks.append(current_block)

    parsed = []
    for block in blocks:
        if not block:
            continue
        # First line is the question (strip leading number)
        q_line = block[0]
        q_text = _re.sub(r'^\d+[\.\)]\s+', '', q_line).strip()
        if not q_text:
            continue

        choices = {}
        answer_raw = ''
        for bline in block[1:]:
            c_match = _re.match(r'^([a-dA-D])[\.\)]\s+(.+)', bline)
            a_match = _re.match(r'^[Aa]nswer\s*:\s*(.+)', bline)
            if c_match:
                choices[c_match.group(1).upper()] = c_match.group(2).strip()
            elif a_match:
                answer_raw = a_match.group(1).strip()

        if choices:
            ans_label = answer_raw.upper().strip('.')[:1] if answer_raw else 'A'
            parsed.append({
                'type': 'multiple_choice',
                'text': q_text,
                'choices': choices,
                'answer': ans_label,
            })
        else:
            parsed.append({
                'type': 'short_answer',
                'text': q_text,
                'answer': answer_raw,
            })

    if not parsed:
        flash('No questions found. Make sure your file uses the required format.', 'error')
        return redirect(url_for('teacher_question_bank'))

    # ── Create group (or reuse existing) ────────────────────────────────────
    existing_grp = conn.execute(
        'SELECT id FROM question_bank_groups WHERE teacher_id=? AND name=?',
        (session['user_id'], group_name)
    ).fetchone()
    if existing_grp:
        group_id = existing_grp['id']
    else:
        cur = conn.execute(
            'INSERT INTO question_bank_groups (teacher_id, name, description, class_id) VALUES (?,?,?,NULL)',
            (session['user_id'], group_name, f'Imported from {raw_name}')
        )
        group_id = cur.lastrowid

    # ── Insert questions ─────────────────────────────────────────────────────
    added = 0
    for q in parsed:
        cur = conn.execute(
            '''INSERT INTO questions
               (exam_id, section_id, question_text, question_type, points, correct_answer,
                order_index, bank_group_id, is_bank_only, teacher_id)
               VALUES (NULL, NULL, ?, ?, 1, ?, 0, ?, 1, ?)''',
            (q['text'], q['type'], q['answer'], group_id, session['user_id'])
        )
        q_id = cur.lastrowid
        if q['type'] == 'multiple_choice':
            label_map = {'A': 0, 'B': 1, 'C': 2, 'D': 3}
            for lbl, txt in q.get('choices', {}).items():
                conn.execute(
                    'INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                    (q_id, lbl, txt)
                )
        added += 1

    conn.commit()
    flash(f'Imported {added} question{"s" if added != 1 else ""} into group "{group_name}".', 'success')
    return redirect(url_for('teacher_question_bank'))


@app.route('/teacher/question-bank')
@role_required('teacher')
def teacher_question_bank():
    conn = get_db()
    # Fetch all questions owned by this teacher:
    #   - bank-only questions (is_bank_only=1, teacher_id matches)
    #   - exam questions (via exam -> class -> teacher)
    # Grouped ones come first for dedup (richest info wins)
    raw = conn.execute('''
        SELECT q.id, q.question_text, q.question_type, q.points, q.correct_answer,
               q.bank_group_id, q.is_bank_only,
               e.title as exam_title,
               c.subject_name as exam_subject, c.block_name as exam_block,
               s.title as section_title,
               g.name as group_name,
               gc.subject_name as group_subject, gc.block_name as group_block, gc.id as group_class_id
        FROM questions q
        LEFT JOIN exams e ON q.exam_id = e.id
        LEFT JOIN classes c ON e.class_id = c.id
        LEFT JOIN sections s ON q.section_id = s.id
        LEFT JOIN question_bank_groups g ON q.bank_group_id = g.id
        LEFT JOIN classes gc ON g.class_id = gc.id
        WHERE (q.is_bank_only = 1 AND q.teacher_id = ?)
           OR ((q.is_bank_only IS NULL OR q.is_bank_only = 0) AND c.teacher_id = ?)
        ORDER BY q.bank_group_id IS NULL ASC, COALESCE(g.name,''), COALESCE(c.subject_name,''), e.title
    ''', (session['user_id'], session['user_id'])).fetchall()
    questions = []
    seen_texts = set()
    for q in raw:
        # Always keep bank-only questions (they were explicitly added to the bank)
        # For exam questions, deduplicate by text to avoid showing the same question
        # from multiple exams
        if q['is_bank_only']:
            qd = dict(q)
            if q['question_type'] == 'multiple_choice':
                qd['choices'] = [dict(c) for c in conn.execute('SELECT * FROM choices WHERE question_id=?', (q['id'],)).fetchall()]
            else:
                qd['choices'] = []
            questions.append(qd)
        else:
            key = (q['question_text'].strip().lower(), q['question_type'])
            if key in seen_texts:
                continue
            seen_texts.add(key)
            qd = dict(q)
            if q['question_type'] == 'multiple_choice':
                qd['choices'] = [dict(c) for c in conn.execute('SELECT * FROM choices WHERE question_id=?', (q['id'],)).fetchall()]
            else:
                qd['choices'] = []
            questions.append(qd)
    # Get all sections grouped by exam for the "Add Question" form
    sections_raw = conn.execute('''
        SELECT s.id as section_id, s.title as section_title, s.section_type,
               e.id as exam_id, e.title as exam_title
        FROM sections s
        JOIN exams e ON s.exam_id = e.id
        JOIN classes c ON e.class_id = c.id
        WHERE c.teacher_id = ?
        ORDER BY e.title, s.order_index
    ''', (session['user_id'],)).fetchall()
    bank_sections = [dict(s) for s in sections_raw]

    # Get all bank groups for this teacher, joining class info
    groups = conn.execute('''
        SELECT g.*, c.subject_name, c.block_name, c.year_level
        FROM question_bank_groups g
        LEFT JOIN classes c ON g.class_id = c.id
        WHERE g.teacher_id=?
        ORDER BY g.name
    ''', (session['user_id'],)).fetchall()
    bank_groups = [dict(g) for g in groups]

    # Get teacher's classes for the group create/edit form
    teacher_classes = conn.execute(
        'SELECT id, subject_name, block_name, year_level FROM classes WHERE teacher_id=? AND is_active=1 ORDER BY subject_name, block_name',
        (session['user_id'],)
    ).fetchall()
    teacher_classes = [dict(c) for c in teacher_classes]

    return render_template('teacher/question_bank.html', questions=questions, bank_sections=bank_sections, bank_groups=bank_groups, teacher_classes=teacher_classes)

@app.route('/teacher/profile')
@role_required('teacher')
def teacher_profile():
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
    exam_history = conn.execute('''
        SELECT e.*, c.subject_name, c.block_name,
               COUNT(DISTINCT es.id) as total_takers,
               ROUND(AVG(CASE WHEN es.total_points > 0
                    THEN es.score * 100.0 / es.total_points ELSE NULL END), 1) as avg_score
        FROM exams e
        JOIN classes c ON e.class_id = c.id
        LEFT JOIN exam_sessions es ON e.id = es.exam_id AND es.status = 'submitted'
        WHERE c.teacher_id = ?
        GROUP BY e.id
        ORDER BY e.created_at DESC
    ''', (session['user_id'],)).fetchall()
    # Per-question correct count for bar graph
    question_stats = conn.execute('''
        SELECT e.id as exam_id, e.title as exam_title,
               q.id as q_id, q.question_text, q.question_type, q.correct_answer,
               COUNT(DISTINCT es.id) as total_answered,
               SUM(CASE
                   WHEN q.question_type = 'multiple_choice'
                        AND UPPER(TRIM(a.answer_text)) = UPPER(TRIM(q.correct_answer)) THEN 1
                   WHEN q.question_type = 'short_answer'
                        AND LOWER(TRIM(a.answer_text)) = LOWER(TRIM(q.correct_answer)) THEN 1
                   ELSE 0
               END) as correct_count
        FROM exams e
        JOIN classes c ON e.class_id = c.id
        JOIN questions q ON q.exam_id = e.id
        LEFT JOIN exam_sessions es ON e.id = es.exam_id AND es.status = 'submitted'
        LEFT JOIN answers a ON a.session_id = es.id AND a.question_id = q.id
        WHERE c.teacher_id = ?
        GROUP BY e.id, q.id
        ORDER BY e.created_at DESC, q.order_index
    ''', (session['user_id'],)).fetchall()
    exam_questions = {}
    for row in question_stats:
        eid = row['exam_id']
        if eid not in exam_questions:
            exam_questions[eid] = {'title': row['exam_title'], 'questions': []}
        exam_questions[eid]['questions'].append({
            'text': row['question_text'][:60],
            'total': row['total_answered'] or 0,
            'correct': row['correct_count'] or 0,
        })
    return render_template('teacher/profile.html', user=user, exam_history=exam_history, exam_questions=exam_questions)

# ─── Admin Routes ─────────────────────────────────────────────────────────────

@app.route('/admin')
@role_required('admin')
def admin_home():
    conn = get_db()
    stats = {
        'teachers': conn.execute("SELECT COUNT(*) FROM users WHERE role='teacher'").fetchone()[0],
        'students': conn.execute("SELECT COUNT(*) FROM users WHERE role='student'").fetchone()[0],
        'classes':  conn.execute("SELECT COUNT(*) FROM classes").fetchone()[0],
        'exams':    conn.execute("SELECT COUNT(*) FROM exams").fetchone()[0],
        'programs': conn.execute("SELECT COUNT(*) FROM programs").fetchone()[0],
        'active_exams': conn.execute("SELECT COUNT(*) FROM exams WHERE status='active'").fetchone()[0],
        'completed_exams': conn.execute("SELECT COUNT(*) FROM exams WHERE status='completed'").fetchone()[0],
        'upcoming_exams': conn.execute("SELECT COUNT(*) FROM exams WHERE status='upcoming'").fetchone()[0],
    }
    recent_logins = conn.execute("""
        SELECT ll.email, ll.success, ll.logged_at, u.full_name
        FROM login_logs ll
        LEFT JOIN users u ON u.id = ll.user_id
        ORDER BY ll.logged_at DESC LIMIT 8
    """).fetchall()
    active_exams = conn.execute("""
        SELECT e.id, e.title, e.duration_minutes, e.exam_code,
               c.subject_name, c.block_name,
               u.full_name AS teacher_name,
               COUNT(es.id) AS session_count
        FROM exams e
        JOIN classes c ON c.id = e.class_id
        JOIN users u ON u.id = c.teacher_id
        LEFT JOIN exam_sessions es ON es.exam_id = e.id AND es.status = 'ongoing'
        WHERE e.status = 'active'
        GROUP BY e.id
        ORDER BY e.activated_at DESC
        LIMIT 5
    """).fetchall()
    admin_name = session.get('full_name', 'Admin')
    recent_users = conn.execute("SELECT full_name, email, role FROM users ORDER BY created_at DESC LIMIT 5").fetchall()
    return render_template('admin/dashboard.html',
        stats=stats,
        recent_logins=recent_logins,
        active_exams=active_exams,
        admin_name=admin_name,
        recent_users=recent_users
    )

@app.route('/admin/users')
@role_required('admin')
def admin_users():
    conn = get_db()
    users = conn.execute('SELECT * FROM users ORDER BY role, full_name').fetchall()
    return render_template('admin/users.html', users=users)

@app.route('/admin/users/create', methods=['GET', 'POST'])
@role_required('admin')
def admin_create_user():
    conn = get_db()
    programs = conn.execute('SELECT * FROM programs ORDER BY code').fetchall()
    if request.method == 'POST':
        full_name  = request.form.get('full_name', '').strip()
        email      = request.form.get('email', '').strip()
        password   = request.form.get('password', '')
        role       = request.form.get('role', '')
        program    = request.form.get('program', '')
        year_level = request.form.get('year_level', '')
        if not all([full_name, email, password, role]):
            flash('Please fill in all required fields.', 'error')
        else:
            try:
                conn = get_db()
                conn.execute('''
                    INSERT INTO users (full_name, email, password, role, program, year_level)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (full_name, email, hash_password(password), role, program, year_level))
                conn.commit()
                flash(f'Account created for {full_name}!', 'success')
                return redirect(url_for('admin_users'))
            except sqlite3.IntegrityError:
                flash('Email already exists.', 'error')
    return render_template('admin/create_user.html', programs=programs)

@app.route('/admin/users/delete/<int:user_id>', methods=['POST'])
@role_required('admin')
def admin_delete_user(user_id):
    conn = get_db()
    conn.execute('DELETE FROM users WHERE id = ? AND role != "admin"', (user_id,))
    conn.commit()
    flash('User deleted.', 'success')
    return redirect(url_for('admin_users'))

@app.route('/admin/programs', methods=['GET', 'POST'])
@role_required('admin')
def admin_programs():
    conn = get_db()
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            code = request.form.get('code', '').strip().upper()
            name = request.form.get('name', '').strip()
            if code and name:
                try:
                    conn.execute('INSERT INTO programs (code, name) VALUES (?,?)', (code, name))
                    conn.commit()
                    flash('Program added.', 'success')
                except sqlite3.IntegrityError:
                    flash('Program code already exists.', 'error')
        elif action == 'delete':
            prog_id = request.form.get('program_id', type=int)
            conn.execute('DELETE FROM programs WHERE id=?', (prog_id,))
            conn.commit()
            flash('Program deleted.', 'success')
    programs = conn.execute('SELECT * FROM programs ORDER BY code').fetchall()
    return render_template('admin/programs.html', programs=programs)

@app.route('/admin/exams')
@role_required('admin')
def admin_exam_overview():
    conn = get_db()
    # Separation of Results: optional filters by program, year level, and section
    # (block). Defaults are empty, so with no query params the result set is
    # identical to the original unfiltered list.
    f_program = request.args.get('program', '').strip()
    f_year_level = request.args.get('year_level', '').strip()
    f_section = request.args.get('section', '').strip()

    query = '''
        SELECT e.*, c.subject_name, c.block_name, c.program, c.year_level, u.full_name as teacher_name,
               COUNT(DISTINCT es.id) as session_count
        FROM exams e
        JOIN classes c ON e.class_id = c.id
        JOIN users u ON c.teacher_id = u.id
        LEFT JOIN exam_sessions es ON e.id = es.exam_id
        WHERE 1=1
    '''
    params = []
    if f_program:
        query += ' AND c.program = ?'
        params.append(f_program)
    if f_year_level:
        query += ' AND c.year_level = ?'
        params.append(f_year_level)
    if f_section:
        query += ' AND c.block_name = ?'
        params.append(f_section)
    query += ' GROUP BY e.id ORDER BY e.created_at DESC'

    exams = conn.execute(query, params).fetchall()

    # Distinct filter options, pulled from classes so the dropdowns only ever
    # show values that actually exist in the system.
    programs = [r['program'] for r in conn.execute('SELECT DISTINCT program FROM classes ORDER BY program').fetchall()]
    year_levels = [r['year_level'] for r in conn.execute('SELECT DISTINCT year_level FROM classes ORDER BY year_level').fetchall()]
    sections = [r['block_name'] for r in conn.execute('SELECT DISTINCT block_name FROM classes ORDER BY block_name').fetchall()]

    return render_template('admin/exam_overview.html', exams=exams,
                           programs=programs, year_levels=year_levels, sections=sections,
                           f_program=f_program, f_year_level=f_year_level, f_section=f_section)

@app.route('/admin/logs')
@role_required('admin')
def admin_logs():
    conn = get_db()
    login_logs = conn.execute('''
        SELECT ll.*, u.full_name FROM login_logs ll
        LEFT JOIN users u ON ll.user_id = u.id
        ORDER BY ll.logged_at DESC LIMIT 100
    ''').fetchall()
    suspicious = conn.execute('''
        SELECT sl.*, u.full_name, e.title as exam_title
        FROM suspicious_logs sl
        JOIN users u ON sl.student_id = u.id
        JOIN exams e ON sl.exam_id = e.id
        ORDER BY sl.logged_at DESC LIMIT 100
    ''').fetchall()
    return render_template('admin/logs.html', login_logs=login_logs, suspicious=suspicious)

@app.route('/admin/settings')
@role_required('admin')
def admin_settings():
    conn = get_db()
    db_size = '—'
    try:
        size_bytes = os.path.getsize(DB_PATH)
        if size_bytes < 1024:
            db_size = f'{size_bytes} B'
        elif size_bytes < 1048576:
            db_size = f'{size_bytes / 1024:.1f} KB'
        else:
            db_size = f'{size_bytes / 1048576:.2f} MB'
    except:
        pass
    record_counts = {
        'users': conn.execute("SELECT COUNT(*) FROM users").fetchone()[0],
        'programs': conn.execute("SELECT COUNT(*) FROM programs").fetchone()[0],
        'classes': conn.execute("SELECT COUNT(*) FROM classes").fetchone()[0],
        'exams': conn.execute("SELECT COUNT(*) FROM exams").fetchone()[0],
        'questions': conn.execute("SELECT COUNT(*) FROM questions").fetchone()[0],
        'sessions': conn.execute("SELECT COUNT(*) FROM exam_sessions").fetchone()[0],
        'logins': conn.execute("SELECT COUNT(*) FROM login_logs").fetchone()[0],
    }
    return render_template('admin/settings.html', db_size=db_size, record_counts=record_counts)

@app.route('/admin/profile')
@role_required('admin')
def admin_profile():
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
    return render_template('admin/profile.html', user=user)


# ─── Allowed Email Whitelist Management ─────────────────────────────────────

@app.route('/admin/allowed-emails/students')
@role_required('admin')
def admin_allowed_students():
    conn = get_db()
    search = request.args.get('search', '').strip()
    if search:
        emails = conn.execute(
            'SELECT * FROM allowed_student_emails WHERE email LIKE ? ORDER BY email',
            (f'%{search}%',)
        ).fetchall()
    else:
        emails = conn.execute('SELECT * FROM allowed_student_emails ORDER BY email').fetchall()
    total = conn.execute('SELECT COUNT(*) FROM allowed_student_emails').fetchone()[0]
    return render_template('admin/allowed_students.html', emails=emails, total=total, search=search)

@app.route('/admin/allowed-emails/students/add', methods=['POST'])
@role_required('admin')
def admin_add_student_email():
    email = request.form.get('email', '').strip().lower()
    if not email:
        flash('Please enter an email address.', 'error')
    else:
        conn = get_db()
        try:
            conn.execute('INSERT INTO allowed_student_emails (email) VALUES (?)', (email,))
            conn.commit()
            flash(f'Added {email} to allowed student emails.', 'success')
        except Exception:
            flash(f'Email {email} is already in the list.', 'error')
    return redirect(url_for('admin_allowed_students'))

@app.route('/admin/allowed-emails/students/bulk', methods=['POST'])
@role_required('admin')
def admin_bulk_add_student_emails():
    raw = request.form.get('emails_bulk', '')
    import re
    emails = [e.strip().lower() for e in re.split(r'[,;\n\r\s]+', raw) if e.strip()]
    valid_emails = [e for e in emails if '@' in e and '.' in e.split('@')[-1]]
    conn = get_db()
    added = 0
    skipped = 0
    for email in valid_emails:
        try:
            conn.execute('INSERT INTO allowed_student_emails (email) VALUES (?)', (email,))
            added += 1
        except Exception:
            skipped += 1
    conn.commit()
    if added:
        flash(f'Successfully added {added} student email(s).', 'success')
    if skipped:
        flash(f'{skipped} email(s) were already in the list (skipped).', 'error')
    if not valid_emails:
        flash('No valid emails found. Make sure each email contains @ and a domain.', 'error')
    return redirect(url_for('admin_allowed_students'))

@app.route('/admin/allowed-emails/students/delete/<int:email_id>', methods=['POST'])
@role_required('admin')
def admin_delete_student_email(email_id):
    conn = get_db()
    conn.execute('DELETE FROM allowed_student_emails WHERE id = ?', (email_id,))
    conn.commit()
    flash('Email removed from allowed list.', 'success')
    return redirect(url_for('admin_allowed_students'))

@app.route('/admin/allowed-emails/teachers')
@role_required('admin')
def admin_allowed_teachers():
    conn = get_db()
    search = request.args.get('search', '').strip()
    if search:
        emails = conn.execute(
            'SELECT * FROM allowed_teacher_emails WHERE email LIKE ? ORDER BY email',
            (f'%{search}%',)
        ).fetchall()
    else:
        emails = conn.execute('SELECT * FROM allowed_teacher_emails ORDER BY email').fetchall()
    total = conn.execute('SELECT COUNT(*) FROM allowed_teacher_emails').fetchone()[0]
    return render_template('admin/allowed_teachers.html', emails=emails, total=total, search=search)

@app.route('/admin/allowed-emails/teachers/add', methods=['POST'])
@role_required('admin')
def admin_add_teacher_email():
    email = request.form.get('email', '').strip().lower()
    if not email:
        flash('Please enter an email address.', 'error')
    else:
        conn = get_db()
        try:
            conn.execute('INSERT INTO allowed_teacher_emails (email) VALUES (?)', (email,))
            conn.commit()
            flash(f'Added {email} to allowed teacher emails.', 'success')
        except Exception:
            flash(f'Email {email} is already in the list.', 'error')
    return redirect(url_for('admin_allowed_teachers'))

@app.route('/admin/allowed-emails/teachers/bulk', methods=['POST'])
@role_required('admin')
def admin_bulk_add_teacher_emails():
    raw = request.form.get('emails_bulk', '')
    import re
    emails = [e.strip().lower() for e in re.split(r'[,;\n\r\s]+', raw) if e.strip()]
    valid_emails = [e for e in emails if '@' in e and '.' in e.split('@')[-1]]
    conn = get_db()
    added = 0
    skipped = 0
    for email in valid_emails:
        try:
            conn.execute('INSERT INTO allowed_teacher_emails (email) VALUES (?)', (email,))
            added += 1
        except Exception:
            skipped += 1
    conn.commit()
    if added:
        flash(f'Successfully added {added} teacher email(s).', 'success')
    if skipped:
        flash(f'{skipped} email(s) were already in the list (skipped).', 'error')
    if not valid_emails:
        flash('No valid emails found. Make sure each email contains @ and a domain.', 'error')
    return redirect(url_for('admin_allowed_teachers'))

@app.route('/admin/allowed-emails/teachers/delete/<int:email_id>', methods=['POST'])
@role_required('admin')
def admin_delete_teacher_email(email_id):
    conn = get_db()
    conn.execute('DELETE FROM allowed_teacher_emails WHERE id = ?', (email_id,))
    conn.commit()
    flash('Email removed from allowed list.', 'success')
    return redirect(url_for('admin_allowed_teachers'))

# ─── Enhanced Admin Routes ────────────────────────────────────────────────────

# ── Reports ──
@app.route('/admin/reports')
@role_required('admin')
def admin_reports():
    conn = get_db()
    role_rows = conn.execute("SELECT role, COUNT(*) as cnt FROM users GROUP BY role").fetchall()
    role_counts = {r['role']: r['cnt'] for r in role_rows}
    total_users = sum(role_counts.values())
    program_enrollment = conn.execute("""
        SELECT p.code, p.name, COUNT(u.id) as count
        FROM programs p LEFT JOIN users u ON u.program = p.code AND u.role = 'student'
        GROUP BY p.code, p.name ORDER BY count DESC
    """).fetchall()
    exam_rows = conn.execute("SELECT status, COUNT(*) as cnt FROM exams GROUP BY status").fetchall()
    exam_status = {'upcoming': 0, 'active': 0, 'completed': 0}
    total_exams = 0
    for r in exam_rows:
        exam_status[r['status']] = r['cnt']
        total_exams += r['cnt']
    total_logins = conn.execute("SELECT COUNT(*) FROM login_logs").fetchone()[0]
    failed_logins = conn.execute("SELECT COUNT(*) FROM login_logs WHERE success = 0").fetchone()[0]
    login_activity = conn.execute("""
        SELECT DATE(logged_at) as date,
               SUM(CASE WHEN success = 1 THEN 1 ELSE 0 END) as success,
               SUM(CASE WHEN success = 0 THEN 1 ELSE 0 END) as failed
        FROM login_logs WHERE logged_at >= DATE('now', '-7 days')
        GROUP BY DATE(logged_at) ORDER BY date DESC
    """).fetchall()
    sus_rows = conn.execute("SELECT event_type, COUNT(*) as cnt FROM suspicious_logs GROUP BY event_type").fetchall()
    suspicious_summary = {'tab_switch': 0, 'lost_focus': 0, 'other': 0}
    for r in sus_rows:
        if r['event_type'] in suspicious_summary:
            suspicious_summary[r['event_type']] = r['cnt']
        else:
            suspicious_summary['other'] += r['cnt']
    report = {
        'total_users': total_users, 'total_exams': total_exams,
        'total_logins': total_logins, 'failed_logins': failed_logins,
        'role_counts': role_counts, 'program_enrollment': program_enrollment,
        'exam_status': exam_status, 'login_activity': login_activity,
        'suspicious_summary': suspicious_summary,
    }
    return render_template('admin/reports.html', report=report)

# ── Edit User ──
@app.route('/admin/users/edit/<int:user_id>', methods=['GET', 'POST'])
@role_required('admin')
def admin_edit_user(user_id):
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (user_id,)).fetchone()
    if not user:
        flash('User not found.', 'error')
        return redirect(url_for('admin_users'))
    programs = conn.execute('SELECT * FROM programs ORDER BY code').fetchall()
    if request.method == 'POST':
        full_name  = request.form.get('full_name', '').strip()
        email      = request.form.get('email', '').strip()
        role       = request.form.get('role', '')
        program    = request.form.get('program', '')
        year_level = request.form.get('year_level', '')
        if not all([full_name, email, role]):
            flash('Please fill in all required fields.', 'error')
        else:
            try:
                conn.execute("""
                    UPDATE users SET full_name=?, email=?, role=?, program=?, year_level=?
                    WHERE id=?
                """, (full_name, email, role, program, year_level, user_id))
                conn.commit()
                flash(f'User {full_name} updated successfully!', 'success')
                return redirect(url_for('admin_users'))
            except sqlite3.IntegrityError:
                flash('Email already exists for another user.', 'error')
    is_self = (user['id'] == session['user_id'])
    return render_template('admin/edit_user.html', user=user, programs=programs, is_self=is_self)

# ── Reset Password ──
@app.route('/admin/users/reset-password/<int:user_id>', methods=['POST'])
@role_required('admin')
def admin_reset_password(user_id):
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (user_id,)).fetchone()
    if not user:
        flash('User not found.', 'error')
        return redirect(url_for('admin_users'))
    new_password = request.form.get('new_password', '').strip()
    if len(new_password) < 6:
        flash('Password must be at least 6 characters.', 'error')
        return redirect(url_for('admin_edit_user', user_id=user_id))
    conn.execute('UPDATE users SET password=? WHERE id=?', (hash_password(new_password), user_id))
    conn.commit()
    flash(f'Password reset for {user["full_name"]}. New password: {new_password}', 'success')
    return redirect(url_for('admin_edit_user', user_id=user_id))

# ── Exam Analytics ──
@app.route('/admin/exams/<int:exam_id>/analytics')
@role_required('admin')
def admin_exam_analytics(exam_id):
    conn = get_db()
    exam = conn.execute('SELECT * FROM exams WHERE id=?', (exam_id,)).fetchone()
    if not exam:
        flash('Exam not found.', 'error')
        return redirect(url_for('admin_exam_overview'))
    class_info = conn.execute('SELECT * FROM classes WHERE id=?', (exam['class_id'],)).fetchone()
    teacher = conn.execute('SELECT full_name FROM users WHERE id=?', (class_info['teacher_id'],)).fetchone()
    teacher_name = teacher['full_name'] if teacher else '—'
    sessions = conn.execute("""
        SELECT es.*, u.full_name FROM exam_sessions es
        JOIN users u ON es.student_id = u.id WHERE es.exam_id=? ORDER BY es.score DESC
    """, (exam_id,)).fetchall()
    total_sessions = len(sessions)
    submitted = sum(1 for s in sessions if s['status'] == 'submitted')
    terminated = sum(1 for s in sessions if s['status'] == 'terminated')
    scored = [s for s in sessions if s['score'] is not None and s['total_points'] and s['total_points'] > 0]
    avg_score = round(sum((s['score'] / s['total_points']) * 100 for s in scored) / len(scored), 1) if scored else 0
    score_ranges = {'90-100%': 0, '80-89%': 0, '70-79%': 0, '60-69%': 0, 'Below 60%': 0}
    for s in scored:
        pct = (s['score'] / s['total_points']) * 100
        if pct >= 90: score_ranges['90-100%'] += 1
        elif pct >= 80: score_ranges['80-89%'] += 1
        elif pct >= 70: score_ranges['70-79%'] += 1
        elif pct >= 60: score_ranges['60-69%'] += 1
        else: score_ranges['Below 60%'] += 1
    total_questions = conn.execute('SELECT COUNT(*) FROM questions WHERE exam_id=?', (exam_id,)).fetchone()[0]
    stats = {
        'total_sessions': total_sessions, 'submitted': submitted,
        'terminated': terminated, 'avg_score': avg_score,
        'score_ranges': score_ranges, 'total_questions': total_questions,
    }
    hard_questions = conn.execute("""
        SELECT q.id, q.question_text, q.question_type, q.correct_answer,
               COUNT(a.id) as total_answers,
               SUM(CASE WHEN LOWER(TRIM(a.answer_text)) = LOWER(TRIM(q.correct_answer)) THEN 1 ELSE 0 END) as correct_count
        FROM questions q LEFT JOIN answers a ON q.id = a.question_id
        WHERE q.exam_id=? GROUP BY q.id HAVING total_answers > 0
        ORDER BY (correct_count * 1.0 / total_answers) ASC LIMIT 10
    """, (exam_id,)).fetchall()
    hard_questions_list = []
    for q in hard_questions:
        rate = round((q['correct_count'] / q['total_answers']) * 100, 1) if q['total_answers'] > 0 else 0
        hard_questions_list.append({
            'question_text': q['question_text'], 'question_type': q['question_type'],
            'correct_count': q['correct_count'], 'total_answers': q['total_answers'], 'rate': rate,
        })
    suspicious = conn.execute("""
        SELECT sl.*, u.full_name FROM suspicious_logs sl
        JOIN users u ON sl.student_id = u.id WHERE sl.exam_id=? ORDER BY sl.logged_at DESC
    """, (exam_id,)).fetchall()

    # Per-Section Analytics: same idea as the teacher-facing results page —
    # average correctness grouped by exam section, for a quick "which part of
    # the exam was hardest overall" view.
    section_rows = conn.execute("""
        SELECT q.id, s.title as section_title, s.order_index as sec_order, q.correct_answer,
               COUNT(a.id) as total_answers,
               SUM(CASE WHEN LOWER(TRIM(a.answer_text)) = LOWER(TRIM(q.correct_answer)) THEN 1 ELSE 0 END) as correct_count
        FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        LEFT JOIN answers a ON q.id = a.question_id
        WHERE q.exam_id=?
        GROUP BY q.id
        ORDER BY s.order_index
    """, (exam_id,)).fetchall()
    section_order = []
    section_agg = {}
    for r in section_rows:
        title = r['section_title'] or 'Untitled Section'
        if title not in section_agg:
            section_agg[title] = {'section_title': title, 'question_count': 0, 'pct_sum': 0}
            section_order.append(title)
        pct = round((r['correct_count'] / r['total_answers']) * 100) if r['total_answers'] else 0
        section_agg[title]['question_count'] += 1
        section_agg[title]['pct_sum'] += pct
    section_stats = []
    for title in section_order:
        s = section_agg[title]
        avg_pct = round(s['pct_sum'] / s['question_count']) if s['question_count'] else 0
        section_stats.append({'section_title': title, 'question_count': s['question_count'], 'avg_pct': avg_pct})

    return render_template('admin/exam_analytics.html',
        exam=exam, class_info=class_info, teacher_name=teacher_name,
        stats=stats, sessions=sessions, hard_questions=hard_questions_list, suspicious=suspicious,
        section_stats=section_stats)

# ── Class Management ──
@app.route('/admin/classes')
@role_required('admin')
def admin_classes():
    conn = get_db()
    classes = conn.execute("""
        SELECT c.*, u.full_name as teacher_name,
               COUNT(DISTINCT ce.student_id) as enrollment_count
        FROM classes c JOIN users u ON c.teacher_id = u.id
        LEFT JOIN class_enrollments ce ON c.id = ce.class_id
        GROUP BY c.id ORDER BY c.is_active DESC, c.subject_name
    """).fetchall()
    return render_template('admin/classes.html', classes=classes)

# ── Edit Class ──
@app.route('/admin/classes/edit/<int:class_id>', methods=['GET', 'POST'])
@role_required('admin')
def admin_edit_class(class_id):
    conn = get_db()
    cls = conn.execute('SELECT * FROM classes WHERE id=?', (class_id,)).fetchone()
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('admin_classes'))
    teachers = conn.execute("SELECT id, full_name, email FROM users WHERE role='teacher' ORDER BY full_name").fetchall()
    enrolled = conn.execute("""
        SELECT u.full_name, u.email, u.program, u.year_level
        FROM class_enrollments ce JOIN users u ON ce.student_id = u.id
        WHERE ce.class_id=? ORDER BY u.full_name
    """, (class_id,)).fetchall()
    if request.method == 'POST':
        subject_name = request.form.get('subject_name', '').strip()
        block_name = request.form.get('block_name', '').strip()
        teacher_id = request.form.get('teacher_id', type=int)
        is_active = request.form.get('is_active', type=int)
        if subject_name and block_name and teacher_id is not None:
            conn.execute("""
                UPDATE classes SET subject_name=?, block_name=?, teacher_id=?, is_active=? WHERE id=?
            """, (subject_name, block_name, teacher_id, is_active, class_id))
            conn.commit()
            flash('Class updated successfully!', 'success')
            return redirect(url_for('admin_classes'))
        else:
            flash('Please fill in all fields.', 'error')
    return render_template('admin/edit_class.html', cls=cls, teachers=teachers, enrolled=enrolled)

# ── Database Backup ──
@app.route('/admin/backup')
@role_required('admin')
def admin_backup_db():
    import shutil
    from flask import send_file
    backup_path = os.path.join(os.path.dirname(__file__), 'instance', 'spark_backup.db')
    shutil.copy2(DB_PATH, backup_path)
    return send_file(backup_path, as_attachment=True, download_name='spark_backup.db')


# ─── API Endpoints ────────────────────────────────────────────────────────────

@app.route('/api/log-suspicious', methods=['POST'])
@login_required
def log_suspicious():
    data = request.get_json(force=True, silent=True) or {}
    session_id = data.get('session_id')
    event_type = data.get('event_type', 'tab_switch')
    conn = get_db()
    sess = conn.execute('SELECT * FROM exam_sessions WHERE id=?', (session_id,)).fetchone()
    if sess and sess['status'] == 'ongoing':
        conn.execute('''
            INSERT INTO suspicious_logs (session_id, student_id, exam_id, event_type)
            VALUES (?, ?, ?, ?)
        ''', (session_id, sess['student_id'], sess['exam_id'], event_type))
        # window_minimize and screenshot are logged for teacher visibility but do NOT count toward tab_switch_limit
        if event_type in ('window_minimize', 'screenshot'):
            conn.commit()
            return jsonify({'status': 'logged', 'count': sess['tab_switch_count'], 'terminated': False})
        # Always count tab switches so teacher can monitor
        new_count = sess['tab_switch_count'] + 1
        conn.execute('UPDATE exam_sessions SET tab_switch_count=? WHERE id=?', (new_count, session_id))
        conn.commit()
        exam = conn.execute('SELECT * FROM exams WHERE id=?', (sess['exam_id'],)).fetchone()
        terminated = False
        # Auto-terminate only when tab_switch_enabled is ON and limit is reached
        if exam and exam['tab_switch_enabled'] and exam['tab_switch_limit'] and new_count >= exam['tab_switch_limit']:
            total_points = conn.execute('SELECT COALESCE(SUM(points),0) FROM questions WHERE exam_id=?', (sess['exam_id'],)).fetchone()[0]
            conn.execute("UPDATE exam_sessions SET status='terminated', submitted_at=CURRENT_TIMESTAMP, score=0, total_points=? WHERE id=?",
                         (total_points, session_id,))
            conn.commit()
            terminated = True
        return jsonify({'status': 'logged', 'count': new_count, 'terminated': terminated})
    return jsonify({'status': 'error'})

@app.route('/api/exam-consent', methods=['POST'])
@login_required
def api_exam_consent():
    """Records explicit student consent to monitoring + the data privacy statement.
    Must be called (and succeed) before any monitoring/logging is allowed to start
    on the client — see exam.js, which gates fullscreen/tab/blur tracking on this."""
    data = request.get_json(force=True, silent=True) or {}
    session_id = data.get('session_id')
    conn = get_db()
    sess = conn.execute('SELECT * FROM exam_sessions WHERE id=?', (session_id,)).fetchone()
    if not sess or sess['student_id'] != session['user_id']:
        return jsonify({'status': 'error', 'message': 'Invalid session.'}), 403
    if sess['status'] != 'ongoing':
        return jsonify({'status': 'error', 'message': 'Exam session is not active.'}), 400

    conn.execute(
        "UPDATE exam_sessions SET consent_given=1, consent_at=CURRENT_TIMESTAMP WHERE id=?",
        (session_id,)
    )
    # Audit trail: consent is itself a suspicious_logs-style event so teachers/admins
    # can see exactly when each student agreed, alongside the rest of the session log.
    conn.execute('''
        INSERT INTO suspicious_logs (session_id, student_id, exam_id, event_type)
        VALUES (?, ?, ?, ?)
    ''', (session_id, sess['student_id'], sess['exam_id'], 'consent_given'))
    conn.commit()
    return jsonify({'status': 'ok'})

@app.route('/privacy-policy')
def privacy_policy():
    """Public data privacy statement. Linked from the exam consent modal and
    can also be linked from signup/footer. No login required so students can
    review it even before creating an account."""
    return render_template('privacy_policy.html')

@app.route('/api/exam-status/<int:exam_id>')
@login_required
def api_exam_status(exam_id):
    conn = get_db()
    exam = conn.execute('SELECT status, activated_at, duration_minutes FROM exams WHERE id=?', (exam_id,)).fetchone()
    sess = conn.execute(
        'SELECT status, tab_switch_count FROM exam_sessions WHERE exam_id=? AND student_id=?',
        (exam_id, session['user_id'])
    ).fetchone()

    # Compute server-side time remaining so the student timer stays in sync
    time_remaining = None
    if exam and exam['activated_at'] and exam['status'] == 'active':
        try:
            activated_at = datetime.strptime(exam['activated_at'], '%Y-%m-%d %H:%M:%S')
        except ValueError:
            activated_at = datetime.strptime(exam['activated_at'], '%Y-%m-%d %H:%M')
        elapsed = int((datetime.now() - activated_at).total_seconds())
        time_remaining = max(0, exam['duration_minutes'] * 60 - elapsed)

    return jsonify({
        'exam_status': exam['status'] if exam else None,
        'session_status': sess['status'] if sess else None,
        'tab_switch_count': sess['tab_switch_count'] if sess else 0,
        'time_remaining_seconds': time_remaining,
    })

@app.route('/api/heartbeat', methods=['POST'])
@login_required
def api_heartbeat():
    """Student pings this every 4s. Gap > 8s = disconnected on teacher side."""
    data = request.get_json(force=True, silent=True) or {}
    session_id = data.get('session_id')
    conn = get_db()

    # Ensure column exists on old DBs
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN last_seen TIMESTAMP')
        conn.commit()
    except Exception:
        pass

    sess = conn.execute('SELECT * FROM exam_sessions WHERE id=?', (session_id,)).fetchone()
    if not sess or sess['student_id'] != session['user_id']:
        return jsonify({'status': 'error'}), 403
    if sess['status'] != 'ongoing':
        return jsonify({'status': 'ok'})

    now = datetime.now()
    now_str = now.strftime('%Y-%m-%d %H:%M:%S')

    # Detect reconnect: if last_seen gap > 8s, student was offline — log reconnection
    try:
        was_seen = sess['last_seen']
    except (IndexError, KeyError):
        was_seen = None

    if was_seen:
        try:
            last = datetime.strptime(was_seen, '%Y-%m-%d %H:%M:%S')
            gap = (now - last).total_seconds()
            if gap > 8:
                # Was offline, now back — log connected
                conn.execute(
                    'INSERT INTO suspicious_logs (session_id, student_id, exam_id, event_type) VALUES (?,?,?,?)',
                    (session_id, sess['student_id'], sess['exam_id'], 'connected')
                )
        except Exception:
            pass
    else:
        # Very first heartbeat — log initial connection
        conn.execute(
            'INSERT INTO suspicious_logs (session_id, student_id, exam_id, event_type) VALUES (?,?,?,?)',
            (session_id, sess['student_id'], sess['exam_id'], 'connected')
        )

    conn.execute('UPDATE exam_sessions SET last_seen=? WHERE id=?', (now_str, session_id))
    conn.commit()
    return jsonify({'status': 'ok'})

@app.route('/api/monitoring/<int:exam_id>')
@login_required
def api_monitoring(exam_id):
    if session.get('role') not in ('teacher', 'admin'):
        return jsonify({'error': 'Not authorized'}), 403
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or (session.get('role') == 'teacher' and int(exam["teacher_id"]) != int(session["user_id"])):
        return jsonify({'error': 'Not authorized'}), 403

    total_q = conn.execute('SELECT COUNT(*) FROM questions WHERE exam_id=?', (exam_id,)).fetchone()[0]
    # Ensure last_seen column exists (handles existing DBs before migration)
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN last_seen TIMESTAMP')
        conn.commit()
    except Exception:
        pass  # Already exists

    active_sessions = conn.execute('''
        SELECT es.id, u.full_name, es.tab_switch_count, es.status,
               COUNT(DISTINCT a.question_id) as answered,
               es.started_at, es.last_seen
        FROM exam_sessions es
        JOIN users u ON es.student_id = u.id
        LEFT JOIN answers a ON es.id = a.session_id
        WHERE es.exam_id=? AND es.status='ongoing'
        GROUP BY es.id
    ''', (exam_id,)).fetchall()

    # Determine connection status: connected if last_seen within 8 seconds
    # Also log disconnection exactly once when gap first crosses 8s
    from datetime import datetime as dt
    now_ts = dt.now()
    sessions_data = []
    for s in active_sessions:
        row = dict(s)
        try:
            was_seen = s['last_seen']
        except (IndexError, KeyError):
            was_seen = None

        if was_seen:
            try:
                ls = dt.strptime(was_seen, '%Y-%m-%d %H:%M:%S')
                gap = (now_ts - ls).total_seconds()
                is_connected = gap <= 8
                row['connected'] = is_connected

                if not is_connected:
                    # Log disconnect once: only if the last connection-state log is 'connected'
                    last_state = conn.execute(
                        """SELECT event_type FROM suspicious_logs
                           WHERE session_id=? AND event_type IN ('connected','disconnected')
                           ORDER BY logged_at DESC LIMIT 1""",
                        (s['id'],)
                    ).fetchone()
                    # Log if: no prior state log at all, OR last state was 'connected'
                    if last_state is None or last_state['event_type'] == 'connected':
                        stu_id = conn.execute(
                            'SELECT student_id FROM exam_sessions WHERE id=?', (s['id'],)
                        ).fetchone()[0]
                        conn.execute(
                            'INSERT INTO suspicious_logs (session_id, student_id, exam_id, event_type) VALUES (?,?,?,?)',
                            (s['id'], stu_id, exam_id, 'disconnected')
                        )
                        conn.commit()
            except Exception:
                row['connected'] = True
        else:
            # No heartbeat yet — treat as connected (just joined, waiting for first ping)
            row['connected'] = True
        sessions_data.append(row)

    recent_logs = conn.execute('''
        SELECT sl.*, u.full_name FROM suspicious_logs sl
        JOIN users u ON sl.student_id = u.id
        WHERE sl.exam_id=?
        ORDER BY sl.logged_at DESC LIMIT 50
    ''', (exam_id,)).fetchall()

    # Results data
    passing_score = exam['passing_score'] if exam['passing_score'] is not None else 75
    results_rows = conn.execute('''
        SELECT u.full_name, es.score, es.total_points, es.status, es.submitted_at
        FROM exam_sessions es JOIN users u ON es.student_id = u.id
        WHERE es.exam_id=?
        ORDER BY
            CASE es.status WHEN 'ongoing' THEN 0 WHEN 'submitted' THEN 1 ELSE 2 END,
            es.score DESC
    ''', (exam_id,)).fetchall()
    results_list = []
    for r in results_rows:
        row = dict(r)
        if row['score'] is not None and row['total_points']:
            row['pct'] = round((row['score'] / row['total_points']) * 100)
        else:
            row['pct'] = None
        results_list.append(row)

    submitted_sessions = conn.execute(
        "SELECT id FROM exam_sessions WHERE exam_id=? AND status='submitted'", (exam_id,)
    ).fetchall()
    total_submitted = len(submitted_sessions)
    session_ids = [r['id'] for r in submitted_sessions]

    questions_raw = conn.execute('''
        SELECT q.id, q.question_text, q.correct_answer,
               s.title as section_title, q.order_index, s.order_index as sec_order
        FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id = ?
        ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()

    question_stats = []
    for q in questions_raw:
        if session_ids:
            correct_count = conn.execute('''
                SELECT COUNT(*) FROM answers
                WHERE question_id=? AND session_id IN ({})
                AND LOWER(TRIM(answer_text)) = LOWER(TRIM(?))
            '''.format(','.join('?' * len(session_ids))),
            [q['id']] + session_ids + [q['correct_answer']]).fetchone()[0]
        else:
            correct_count = 0
        pct = round((correct_count / total_submitted * 100)) if total_submitted else 0
        question_stats.append({
            'question_text': q['question_text'],
            'section_title': q['section_title'],
            'correct_count': correct_count,
            'total': total_submitted,
            'pct': pct,
        })
    question_stats.sort(key=lambda x: x['pct'], reverse=True)
    for i, qs in enumerate(question_stats, 1):
        qs['number'] = i

    return jsonify({
        'sessions': sessions_data,
        'total_q': total_q,
        'tab_limit': exam['tab_switch_limit'],
        'tab_switch_enabled': bool(exam['tab_switch_enabled']),
        'logs': [dict(l) for l in recent_logs],
        'results': results_list,
        'question_stats': question_stats,
        'total_submitted': total_submitted,
        'passing_score': passing_score,
    })

@app.route('/api/results/<int:exam_id>')
@login_required
def api_results(exam_id):
    if session.get('role') not in ('teacher', 'admin'):
        return jsonify({'error': 'Not authorized'}), 403
    conn = get_db()
    exam = conn.execute(
        'SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?',
        (exam_id,)
    ).fetchone()
    if not exam or (session.get('role') == 'teacher' and int(exam['teacher_id']) != int(session['user_id'])):
        return jsonify({'error': 'Not authorized'}), 403

    passing_score = exam['passing_score'] if exam['passing_score'] is not None else 75

    results = conn.execute('''
        SELECT u.full_name, u.id as student_id, es.score, es.total_points,
               es.status, es.submitted_at, es.tab_switch_count
        FROM exam_sessions es JOIN users u ON es.student_id = u.id
        WHERE es.exam_id=?
        ORDER BY es.score DESC
    ''', (exam_id,)).fetchall()

    results_list = []
    for r in results:
        row = dict(r)
        if row['score'] is not None and row['total_points']:
            pct = round((row['score'] / row['total_points']) * 100)
        else:
            pct = None
        row['pct'] = pct
        results_list.append(row)

    submitted_sessions = conn.execute(
        "SELECT id FROM exam_sessions WHERE exam_id=? AND status='submitted'",
        (exam_id,)
    ).fetchall()
    total_submitted = len(submitted_sessions)
    session_ids = [r['id'] for r in submitted_sessions]

    questions_raw = conn.execute('''
        SELECT q.id, q.question_text, q.question_type, q.points, q.correct_answer,
               s.title as section_title, q.order_index, s.order_index as sec_order
        FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id = ?
        ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()

    question_stats = []
    for q in questions_raw:
        if session_ids:
            correct_count = conn.execute('''
                SELECT COUNT(*) FROM answers
                WHERE question_id=? AND session_id IN ({})
                AND LOWER(TRIM(answer_text)) = LOWER(TRIM(?))
            '''.format(','.join('?' * len(session_ids))),
            [q['id']] + session_ids + [q['correct_answer']]).fetchone()[0]
        else:
            correct_count = 0
        pct = round((correct_count / total_submitted * 100)) if total_submitted else 0
        question_stats.append({
            'question_text': q['question_text'],
            'section_title': q['section_title'],
            'correct_count': correct_count,
            'total': total_submitted,
            'pct': pct,
        })

    question_stats.sort(key=lambda x: x['pct'], reverse=True)
    for i, qs in enumerate(question_stats, 1):
        qs['number'] = i

    return jsonify({
        'results': results_list,
        'question_stats': question_stats,
        'total_submitted': total_submitted,
        'passing_score': passing_score,
    })


@app.route('/api/terminate-session/<int:session_id>', methods=['POST'])
@login_required
def api_terminate_session(session_id):
    conn = get_db()
    sess = conn.execute('SELECT * FROM exam_sessions WHERE id=?', (session_id,)).fetchone()
    if sess:
        # Allow teacher of that exam OR admin
        exam = conn.execute('''
            SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id
            WHERE e.id=?
        ''', (sess['exam_id'],)).fetchone()
        if exam and (exam['teacher_id'] == session['user_id'] or session.get('role') == 'admin'):
            total_points = conn.execute('SELECT COALESCE(SUM(points),0) FROM questions WHERE exam_id=?', (exam['id'],)).fetchone()[0]
            conn.execute(
                "UPDATE exam_sessions SET status='terminated', submitted_at=CURRENT_TIMESTAMP, score=0, total_points=? WHERE id=?",
                (total_points, session_id,)
            )
            conn.commit()
            return jsonify({'status': 'terminated'})
    return jsonify({'status': 'error', 'message': 'Not authorized or session not found'}), 403

@app.route('/api/save-answer', methods=['POST'])
@login_required
def api_save_answer():
    data = request.get_json(force=True, silent=True) or {}
    session_id = data.get('session_id')
    question_id = data.get('question_id')
    answer_text = data.get('answer_text', '').strip()
    if not session_id or question_id is None:
        return jsonify({'status': 'error', 'reason': 'missing fields'})
    conn = get_db()
    sess = conn.execute('SELECT * FROM exam_sessions WHERE id=? AND student_id=?',
                        (session_id, session['user_id'])).fetchone()
    if sess and sess['status'] == 'ongoing':
        conn.execute('''
            INSERT OR REPLACE INTO answers (session_id, question_id, answer_text)
            VALUES (?,?,?)
        ''', (session_id, question_id, answer_text))
        conn.commit()
        return jsonify({'status': 'saved'})
    return jsonify({'status': 'error', 'reason': 'session not found or not ongoing'})

if __name__ == '__main__':
    os.makedirs(os.path.join(os.path.dirname(__file__), 'instance'), exist_ok=True)
    with app.app_context():
        init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
