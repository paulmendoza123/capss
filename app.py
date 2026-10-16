from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, g, send_file
import sqlite3
import hashlib
import hmac
import secrets
import os
import random
import string
import json
import re
import io
from functools import wraps
from datetime import datetime, timedelta
import threading
import time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)

# ── Upload limit: 10 MB per imported file ──
MAX_UPLOAD_MB = 10
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024
# Hard cap on the whole request (file + form fields overhead) so oversized
# uploads are rejected before they are fully read into memory.
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_BYTES + 512 * 1024


def _upload_too_large(file_storage):
    """True if an uploaded file is bigger than MAX_UPLOAD_MB."""
    stream = file_storage.stream
    pos = stream.tell()
    stream.seek(0, os.SEEK_END)
    size = stream.tell()
    stream.seek(pos)
    return size > MAX_UPLOAD_BYTES


@app.errorhandler(413)
def _file_too_large(_e):
    flash(f'File is too large. The maximum allowed size is {MAX_UPLOAD_MB} MB.', 'error')
    return redirect(request.referrer or '/')

_PW_ALPHABET = string.ascii_letters + string.digits
INSTANCE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'instance')


def _write_private_file(path, text):
    """Write text to a file readable only by the current user (0600)."""
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, 'w') as f:
        f.write(text)


def _load_secret_key():
    """Session signing key. Never hardcoded in source.

    Order: SPARK_SECRET_KEY env var -> instance/secret_key file -> generate a
    new random key and save it to that file (so sessions survive restarts).
    """
    key = os.environ.get('SPARK_SECRET_KEY', '').strip()
    if key:
        return key
    os.makedirs(INSTANCE_DIR, exist_ok=True)
    path = os.path.join(INSTANCE_DIR, 'secret_key')
    if os.path.exists(path):
        with open(path) as f:
            key = f.read().strip()
        if key:
            return key
    key = secrets.token_hex(32)
    _write_private_file(path, key)
    return key


app.secret_key = _load_secret_key()


# Multiple-choice questions support up to 26 options, labeled A through Z.
MC_LABELS = list(string.ascii_uppercase)

# Minimum number of choices a multiple-choice question must have.
MC_MIN_CHOICES = 2


def rubric_collect_criteria(form):
    """Return [(criterion_text, max_points), ...] for every non-empty
    rubric_text_<n> field submitted alongside its rubric_points_<n> field."""
    out = []
    texts = form.getlist('rubric_text[]')
    pts = form.getlist('rubric_points[]')
    for t, p in zip(texts, pts):
        t = (t or '').strip()
        if not t:
            continue
        try:
            p = int(p)
        except (TypeError, ValueError):
            p = 0
        if p > 0:
            out.append((t, p))
    return out


def rubric_validate(criteria):
    """Essay questions require at least one rubric criterion worth points."""
    if not criteria:
        return False, 'Essay questions require at least one rubric criterion with points greater than 0.'
    return True, None


def mc_collect_choices(form):
    """Return [(label, text), ...] for every non-empty choice_<label> field."""
    out = []
    for label in MC_LABELS:
        ct = (form.get(f'choice_{label}') or '').strip()
        if ct:
            out.append((label, ct))
    return out


def mc_validate(choices, correct):
    """Validate a multiple-choice submission.

    Returns (ok, correct_label, error_message). The correct answer must be one
    of the labels that actually has choice text, and there must be at least
    MC_MIN_CHOICES choices.
    """
    if len(choices) < MC_MIN_CHOICES:
        return False, correct, (
            f'Multiple choice questions need at least {MC_MIN_CHOICES} choices. '
            f'Only {len(choices)} was provided.' if len(choices) == 1 else
            f'Multiple choice questions need at least {MC_MIN_CHOICES} choices. '
            f'Only {len(choices)} were provided.'
        )
    labels = [lbl for lbl, _ in choices]
    correct = (correct or '').strip().upper()
    if correct not in labels:
        return False, correct, (
            'The correct answer must be one of the choices you filled in '
            f'({", ".join(labels)}).'
        )
    return True, correct, None

# ─── Fill in the Blank helpers ──────────────────────────────────────────────
# Storage format for a fill_blank question's correct_answer column:
#   "ans1a/ans1b | ans2a | ans3a/ans3b/ans3c"
# '|' separates the blanks (in the order the '___' markers appear in the
# question text), and '/' separates alternate acceptable answers for one blank.

_FIB_BLANK_RE = re.compile(r'_{3,}')

def fib_count_blanks(question_text):
    """Number of blank markers in a question's text. A marker is any run of
    3 or more underscores, not just exactly '___' — a teacher typing '____'
    or '_____' by habit still counts as one blank rather than being missed
    or (worse) splitting into a mangled display with a stray underscore."""
    return len(_FIB_BLANK_RE.findall(question_text or ''))

def fib_split_text(question_text):
    """Split a fill_blank question's text on its blank markers (any run of
    3+ underscores), for rendering the blanks as input boxes. Kept in sync
    with fib_count_blanks so the number of pieces here always matches the
    count used for grading/validation."""
    return _FIB_BLANK_RE.split(question_text or '')

def fib_parse_answer(correct_answer):
    """Parse a fill_blank correct_answer string into a list of blanks, each a
    list of acceptable alternate answers, e.g.
    'CPU/processor | RAM' -> [['CPU', 'processor'], ['RAM']]"""
    if not correct_answer:
        return []
    blanks = []
    for part in correct_answer.split('|'):
        alts = [a.strip() for a in part.split('/') if a.strip()]
        blanks.append(alts)
    return blanks

def fib_grade(student_answer, correct_answer, case_sensitive=False):
    """Grade a fill_blank submission per blank.
    student_answer: '|'-separated student inputs, one per blank, in order.
    Returns (correct_blank_count, total_blank_count)."""
    blanks = fib_parse_answer(correct_answer)
    total = len(blanks)
    if total == 0:
        return 0, 0
    student_parts = (student_answer or '').split('|')
    correct_count = 0
    for i, alts in enumerate(blanks):
        given = student_parts[i].strip() if i < len(student_parts) else ''
        if not given:
            continue
        if case_sensitive:
            match = any(given == alt for alt in alts)
        else:
            match = any(given.lower() == alt.lower() for alt in alts)
        if match:
            correct_count += 1
    return correct_count, total

# ─── Short Answer helpers ───────────────────────────────────────────────────
# A short_answer question's correct_answer column may hold several acceptable
# (alternative) answers, separated by '/' — the same way Fill in the Blank
# does it:   "CPU/Processor/Central Processing Unit"
# A student is correct if their answer matches ANY of them. A single answer with
# no '/' behaves exactly as before. (An answer that itself contains a slash,
# e.g. "TCP/IP", is split into alternatives too, but the whole text "TCP/IP" is
# still accepted as a match so existing questions keep working.)

def sa_parse_answer(correct_answer):
    """'CPU/Processor' -> ['CPU', 'Processor'] (blank pieces dropped)."""
    return [a.strip() for a in (correct_answer or '').split('/') if a.strip()]

def sa_join_answers(values):
    """Join one or more answer strings into the stored 'a/b/c' form,
    dropping blanks and exact duplicates while keeping the order."""
    out = []
    for v in values:
        for piece in (v or '').split('/'):
            piece = piece.strip()
            if piece and piece not in out:
                out.append(piece)
    return '/'.join(out)

def sa_collect_form(form, name):
    """Read every input called `name` (primary + alternatives) from a form."""
    return sa_join_answers(form.getlist(name))

@app.template_filter('time12')
def time12_filter(value):
    """Display 'YYYY-MM-DD HH:MM[:SS]' as 12-hour time with AM/PM.
    Display only: stored values are unchanged. Anything that is not a
    date-time string (None, '—', plain text) is returned untouched."""
    if not isinstance(value, str):
        return value
    m = re.match(r'^(\d{4}-\d{2}-\d{2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?', value.strip())
    if not m:
        return value
    date_part, hh, mm, ss = m.group(1), int(m.group(2)), m.group(3), m.group(4)
    suffix = 'PM' if hh >= 12 else 'AM'
    h12 = hh % 12 or 12
    return f"{date_part} {h12}:{mm}{':' + ss if ss else ''} {suffix}"


@app.template_filter('sa_alts')
def _sa_alts_filter(correct_answer):
    return sa_parse_answer(correct_answer)

@app.template_filter('sa_display')
def _sa_display_filter(correct_answer, question_type=''):
    """Readable form of a correct answer: short answers show 'A or B'."""
    if question_type == 'short_answer':
        return ' or '.join(sa_parse_answer(correct_answer)) or (correct_answer or '')
    return correct_answer


def _q_get(q, key, default=None):
    """Read a column from a sqlite3.Row or dict, tolerating a missing column."""
    try:
        v = q[key]
    except (IndexError, KeyError):
        return default
    return default if v is None else v

def question_credit(q, answer, manual_score=None):
    """Fraction of credit (0.0 - 1.0) that `answer` earns on question `q`.
    Single source of truth used by submit grading, the student answer review
    and the per-question analytics, so they can never disagree.

    Essay questions cannot be auto-graded from the raw text — they are always
    manually scored by the teacher against a rubric. `manual_score` (the
    points already awarded, if any) is optional and defaults to None, meaning
    "not graded yet" → 0 credit until a teacher scores it.
    """
    ans = (answer or '').strip()
    correct = (_q_get(q, 'correct_answer', '') or '').strip()
    qtype = _q_get(q, 'question_type', '')
    case_sensitive = bool(_q_get(q, 'case_sensitive', 0))
    if qtype == 'essay':
        if manual_score is None:
            return 0.0
        pts = _q_get(q, 'points', 0) or 0
        return (manual_score / pts) if pts else 0.0
    if qtype == 'fill_blank':
        got, total = fib_grade(ans, correct, case_sensitive)
        return (got / total) if total else 0.0
    if not ans:
        return 0.0  # an unanswered question never earns credit
    if qtype == 'short_answer':
        alts = sa_parse_answer(correct) or [correct]
        if correct not in alts:
            alts.append(correct)  # the whole text also counts, e.g. "TCP/IP"
        if case_sensitive:
            return 1.0 if any(ans == a for a in alts) else 0.0
        return 1.0 if any(ans.lower() == a.lower() for a in alts) else 0.0
    if qtype == 'multiple_choice' or qtype == 'true_false':
        return 1.0 if ans.upper() == correct.upper() else 0.0
    if case_sensitive:
        return 1.0 if ans == correct else 0.0
    return 1.0 if ans.lower() == correct.lower() else 0.0

def count_fully_correct(conn, exam_id, questions):
    """{question_id: number of SUBMITTED sessions whose answer is fully correct}.
    Uses question_credit so fill-in-the-blank alternates ('/'), multiple blanks
    ('|') and case-sensitivity are respected (a raw string compare is not).
    Essay answers pass their manual_score through so a fully-graded, full-marks
    essay counts as correct too (ungraded essays correctly count as 0 here)."""
    by_id = {q['id']: q for q in questions}
    counts = {qid: 0 for qid in by_id}
    if not by_id:
        return counts
    rows = conn.execute('''
        SELECT question_id, answer_text, manual_score FROM answers
        WHERE session_id IN (SELECT id FROM exam_sessions WHERE exam_id=? AND status='submitted')
    ''', (exam_id,)).fetchall()
    for r in rows:
        q = by_id.get(r['question_id'])
        if q is not None and question_credit(q, r['answer_text'], r['manual_score']) >= 1.0:
            counts[q['id']] += 1
    return counts


def avg_essay_pct(conn, exam_id, question_id, points):
    """Average score percentage for one essay question, across every
    SUBMITTED session that has been graded so far. Ungraded submissions are
    excluded (not counted as 0%) so the average isn't dragged down just
    because grading isn't finished yet. Returns None if nothing is graded."""
    if not points:
        return None
    rows = conn.execute('''
        SELECT manual_score FROM answers
        WHERE question_id=? AND manual_score IS NOT NULL
          AND session_id IN (SELECT id FROM exam_sessions WHERE exam_id=? AND status='submitted')
    ''', (question_id, exam_id)).fetchall()
    if not rows:
        return None
    total_pct = sum((r['manual_score'] / points) * 100 for r in rows)
    return round(total_pct / len(rows))

def build_question_stats(conn, exam_id, questions_raw, total_submitted):
    """Question Analysis rows, easiest first: for each question, how many of the
    students who submitted got it RIGHT vs. not.
    - A student counts as right only when the whole answer is right (for a
      fill-in-the-blank, every blank must be right).
    - Essay questions are left out: they are graded by hand, so there is no
      simple right/wrong tally for them."""
    gradable = [q for q in questions_raw if q['question_type'] != 'essay']
    correct_counts = count_fully_correct(conn, exam_id, gradable)
    stats = []
    for q in gradable:
        correct_count = correct_counts.get(q['id'], 0)
        pct = round(correct_count / total_submitted * 100) if total_submitted else 0
        stats.append({
            'question_text': q['question_text'],
            'question_type': q['question_type'],
            'section_title': q['section_title'],
            'correct_count': correct_count,
            'total': total_submitted,
            'pct': pct,
        })
    stats.sort(key=lambda x: x['pct'], reverse=True)
    for i, qs in enumerate(stats, 1):
        qs['number'] = i
    return stats


def pct_floor(score, total):
    """Whole-number percentage, multiplying BEFORE dividing so exact scores
    are not floored by float error (e.g. 58/100 must be 58, not 57)."""
    if score is None or not total:
        return None
    return int(score * 100 / total)

@app.template_filter('score_fmt')
def score_fmt(v):
    """Show a score without truncating partial credit: 7 -> '7', 7.5 -> '7.5'."""
    if v is None:
        return '—'
    v = float(v)
    if v == int(v):
        return str(int(v))
    return f'{v:.2f}'.rstrip('0').rstrip('.')

DB_PATH = os.path.join(os.path.dirname(__file__), 'instance', 'spark.db')

# ─── Database ────────────────────────────────────────────────────────────────

def get_db():
    if 'db' not in g:
        conn = sqlite3.connect(DB_PATH, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute('PRAGMA journal_mode=WAL')
        conn.execute('PRAGMA synchronous=NORMAL')
        conn.execute('PRAGMA busy_timeout=5000')
        g.db = conn
    return g.db

@app.teardown_appcontext
def close_db(error=None):
    db = g.pop('db', None)
    if db is not None:
        db.close()

def _migrate_local_time_defaults(conn):
    """One-time, idempotent. SQLite's CURRENT_TIMESTAMP is UTC, but the rest of
    the app (exam schedule, timer, activated_at) uses local time. Switch the
    column defaults of an existing database to local time. Only the schema
    default changes: existing rows are NOT modified."""
    try:
        conn.commit()
        conn.execute('BEGIN IMMEDIATE')
        found = conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' "
            "AND sql LIKE '%DEFAULT CURRENT_TIMESTAMP%'").fetchone()[0]
        if not found:
            conn.execute('ROLLBACK')
            return
        ver = conn.execute('PRAGMA schema_version').fetchone()[0]
        conn.execute('PRAGMA writable_schema=ON')
        conn.execute(
            "UPDATE sqlite_master SET sql = replace(sql, 'DEFAULT CURRENT_TIMESTAMP', "
            "\"DEFAULT (datetime('now','localtime'))\") "
            "WHERE type='table' AND sql LIKE '%DEFAULT CURRENT_TIMESTAMP%'")
        conn.execute('PRAGMA schema_version=%d' % (ver + 1))
        conn.execute('PRAGMA writable_schema=OFF')
        conn.execute('COMMIT')
    except Exception as e:
        try:
            conn.execute('ROLLBACK')
        except Exception:
            pass
        print('[init_db] local-time default migration skipped:', e)


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
            created_at TIMESTAMP DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS programs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT (datetime('now','localtime'))
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
            created_at TIMESTAMP DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (teacher_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS class_enrollments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            joined_at TIMESTAMP DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (class_id) REFERENCES classes(id),
            FOREIGN KEY (student_id) REFERENCES users(id),
            UNIQUE(class_id, student_id)
        );

        CREATE TABLE IF NOT EXISTS class_allowed_emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER NOT NULL,
            email TEXT NOT NULL,
            added_at TIMESTAMP DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (class_id) REFERENCES classes(id),
            UNIQUE(class_id, email)
        );
        CREATE INDEX IF NOT EXISTS idx_class_allowed_emails_class ON class_allowed_emails(class_id);

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
            created_at TIMESTAMP DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (class_id) REFERENCES classes(id)
        );

        CREATE TABLE IF NOT EXISTS sections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            section_type TEXT NOT NULL CHECK(section_type IN ('multiple_choice','short_answer','fill_blank')),
            order_index INTEGER DEFAULT 0,
            FOREIGN KEY (exam_id) REFERENCES exams(id)
        );

        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL,
            section_id INTEGER,
            question_text TEXT NOT NULL,
            question_type TEXT NOT NULL CHECK(question_type IN ('multiple_choice','short_answer','fill_blank','essay','true_false')),
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
            started_at TIMESTAMP DEFAULT (datetime('now','localtime')),
            submitted_at TIMESTAMP,
            status TEXT DEFAULT 'ongoing' CHECK(status IN ('ongoing','submitted','terminated')),
            score REAL,
            total_points INTEGER,
            tab_switch_count INTEGER DEFAULT 0,
            fullscreen_exit_count INTEGER DEFAULT 0,
            lost_focus_count INTEGER DEFAULT 0,
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
            logged_at TIMESTAMP DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (session_id) REFERENCES exam_sessions(id)
        );

        CREATE TABLE IF NOT EXISTS password_reset_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            teacher_id INTEGER NOT NULL,
            class_id INTEGER,
            reset_at TIMESTAMP DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS login_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            email TEXT,
            success INTEGER DEFAULT 1,
            ip_address TEXT,
            logged_at TIMESTAMP DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS question_bank_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            description TEXT,
            class_id INTEGER REFERENCES classes(id),
            created_at TIMESTAMP DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (teacher_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );
    ''')

    # Indexes on the columns hit hardest by exam-time traffic (heartbeat, answer
    # save, and status polling all filter on these every few seconds per
    # student). Without them, sqlite does a full table scan per lookup, which
    # gets noticeably slow under 50 concurrent students on modest hardware.
    cursor.executescript('''
        CREATE INDEX IF NOT EXISTS idx_exam_sessions_exam ON exam_sessions(exam_id);
        CREATE INDEX IF NOT EXISTS idx_exam_sessions_student ON exam_sessions(student_id);
        CREATE INDEX IF NOT EXISTS idx_answers_session ON answers(session_id);
        CREATE INDEX IF NOT EXISTS idx_answers_question ON answers(question_id);
        CREATE INDEX IF NOT EXISTS idx_suspicious_logs_session ON suspicious_logs(session_id);
        CREATE INDEX IF NOT EXISTS idx_suspicious_logs_exam ON suspicious_logs(exam_id);
        CREATE INDEX IF NOT EXISTS idx_questions_exam ON questions(exam_id);
        CREATE INDEX IF NOT EXISTS idx_questions_section ON questions(section_id);
        CREATE INDEX IF NOT EXISTS idx_sections_exam ON sections(exam_id);
        CREATE INDEX IF NOT EXISTS idx_exams_class ON exams(class_id);
        CREATE INDEX IF NOT EXISTS idx_class_enrollments_class ON class_enrollments(class_id);
        CREATE INDEX IF NOT EXISTS idx_class_enrollments_student ON class_enrollments(student_id);
        CREATE INDEX IF NOT EXISTS idx_users_email ON users(email);
    ''')

    # Default programs
    cursor.execute("INSERT OR IGNORE INTO programs (code, name) VALUES ('BSIT', 'Bachelor of Science in Information Technology')")
    cursor.execute("INSERT OR IGNORE INTO programs (code, name) VALUES ('BSCS', 'Bachelor of Science in Computer Science')")

    # Migration: users.must_change_password (forces a password change at next login)
    try:
        conn.execute('ALTER TABLE users ADD COLUMN must_change_password INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: account lockout bookkeeping
    try:
        conn.execute('ALTER TABLE users ADD COLUMN failed_login_attempts INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists
    try:
        conn.execute('ALTER TABLE users ADD COLUMN locked_until TIMESTAMP')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # First-run admin: only created when NO admin exists. It starts with the
    # default password (admin123, or SPARK_ADMIN_PASSWORD if set) but is FLAGGED so
    # the admin is forced to choose a new password right after the first login.
    has_admin = cursor.execute("SELECT 1 FROM users WHERE role='admin' LIMIT 1").fetchone()
    if not has_admin:
        admin_plain = os.environ.get('SPARK_ADMIN_PASSWORD', '').strip() or 'admin123'
        cursor.execute('''
            INSERT INTO users (full_name, email, password, role, must_change_password)
            VALUES (?, ?, ?, ?, 1)
        ''', ('Administrator', 'admin@spark.edu', generate_password_hash(admin_plain), 'admin'))
        conn.commit()
        print("First run: admin account created -> admin@spark.edu "
              "(password must be changed at first login).")
    else:
        # Existing DB: any admin still on the default password is forced to change it.
        for a in cursor.execute("SELECT id, email, password FROM users WHERE role='admin'").fetchall():
            if verify_password('admin123', a['password']):
                cursor.execute('UPDATE users SET must_change_password=1 WHERE id=?', (a['id'],))
                conn.commit()
                print(f"NOTICE: admin '{a['email']}' still uses the default password; "
                      "a password change will be required at next login.")

    # Migration: drop the unused allowed_*_emails tables (no code ever used them).
    # Only dropped when empty, so no data can be lost.
    for _t in ('allowed_student_emails', 'allowed_teacher_emails'):
        try:
            if conn.execute(f'SELECT COUNT(*) FROM {_t}').fetchone()[0] == 0:
                conn.execute(f'DROP TABLE {_t}')
                conn.commit()
        except Exception:
            pass  # Table already gone

    # Migration: short-answer alternatives are now separated by '/' (like Fill in
    # the Blank). Convert the ' | ' form written by an earlier build. Idempotent.
    try:
        conn.execute("UPDATE questions SET correct_answer = REPLACE(correct_answer, ' | ', '/') "
                     "WHERE question_type = 'short_answer' AND correct_answer LIKE '% | %'")
        conn.commit()
    except Exception:
        pass

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

    # Migration: add is_ready flag. New exams start as drafts (is_ready=0) and are
    # never auto-opened by the scheduler until the teacher clicks "Finish editing".
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN is_ready INTEGER DEFAULT 0')
        conn.execute('UPDATE exams SET is_ready=1')  # existing exams keep current behavior
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

    # Migration: teacher-set "time left" warning (minutes before the end of the
    # exam). Defaults to 5; 0 turns the popup off.
    try:
        conn.execute('ALTER TABLE exams ADD COLUMN time_warning_minutes INTEGER DEFAULT 5')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: add fullscreen_exit_count / lost_focus_count as their own
    # counters, separate from tab_switch_count — only tab_switch_count drives
    # auto-termination; these two are informational-only for teacher monitoring.
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN fullscreen_exit_count INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN lost_focus_count INTEGER DEFAULT 0')
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
                    question_type TEXT NOT NULL CHECK(question_type IN ('multiple_choice','short_answer','fill_blank')),
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

    # Migration: timestamp defaults use local time instead of UTC
    _migrate_local_time_defaults(conn)

    # Migration: add case_sensitive toggle to questions (for short-answer grading)
    try:
        conn.execute('ALTER TABLE questions ADD COLUMN case_sensitive INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: track who created each account (Admin Registration audit trail —
    # controlled/invite-based admin creation). NULL for pre-existing/self-signup accounts.
    try:
        conn.execute('ALTER TABLE users ADD COLUMN created_by INTEGER')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # Migration: widen questions.question_type CHECK constraint to allow 'fill_blank'.
    # SQLite can't ALTER a CHECK constraint in place, so recreate the table (same
    # pattern used for the exam_id-nullable migration above) if the old constraint
    # is still in effect.
    try:
        tbl_sql = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='questions'"
        ).fetchone()
        if tbl_sql and 'fill_blank' not in (tbl_sql['sql'] or ''):
            conn.execute('PRAGMA foreign_keys=OFF')
            conn.execute('''
                CREATE TABLE questions_fib_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    exam_id INTEGER,
                    section_id INTEGER,
                    question_text TEXT NOT NULL,
                    question_type TEXT NOT NULL CHECK(question_type IN ('multiple_choice','short_answer','fill_blank')),
                    points INTEGER DEFAULT 1,
                    correct_answer TEXT,
                    order_index INTEGER DEFAULT 0,
                    bank_group_id INTEGER REFERENCES question_bank_groups(id),
                    is_bank_only INTEGER DEFAULT 0,
                    teacher_id INTEGER REFERENCES users(id),
                    case_sensitive INTEGER DEFAULT 0,
                    FOREIGN KEY (exam_id) REFERENCES exams(id),
                    FOREIGN KEY (section_id) REFERENCES sections(id)
                )
            ''')
            conn.execute('''
                INSERT INTO questions_fib_new
                    (id, exam_id, section_id, question_text, question_type, points,
                     correct_answer, order_index, bank_group_id, is_bank_only, teacher_id, case_sensitive)
                SELECT id, exam_id, section_id, question_text, question_type, points,
                       correct_answer, order_index, bank_group_id,
                       COALESCE(is_bank_only, 0), teacher_id, COALESCE(case_sensitive, 0)
                FROM questions
            ''')
            conn.execute('DROP TABLE questions')
            conn.execute('ALTER TABLE questions_fib_new RENAME TO questions')
            conn.execute('PRAGMA foreign_keys=ON')
            conn.commit()
    except Exception:
        conn.execute('PRAGMA foreign_keys=ON')
        pass  # Already widened or migration failed gracefully

    # Migration: add description to sections (sections no longer have a fixed question type)
    try:
        conn.execute('ALTER TABLE sections ADD COLUMN description TEXT')
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

    # Migration: widen questions.question_type CHECK constraint to allow 'essay'.
    # Same recreate-table pattern used for the 'fill_blank' widening above.
    try:
        tbl_sql = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='questions'"
        ).fetchone()
        if tbl_sql and 'essay' not in (tbl_sql['sql'] or ''):
            conn.execute('PRAGMA foreign_keys=OFF')
            conn.execute('''
                CREATE TABLE questions_essay_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    exam_id INTEGER,
                    section_id INTEGER,
                    question_text TEXT NOT NULL,
                    question_type TEXT NOT NULL CHECK(question_type IN ('multiple_choice','short_answer','fill_blank','essay')),
                    points INTEGER DEFAULT 1,
                    correct_answer TEXT,
                    order_index INTEGER DEFAULT 0,
                    bank_group_id INTEGER REFERENCES question_bank_groups(id),
                    is_bank_only INTEGER DEFAULT 0,
                    teacher_id INTEGER REFERENCES users(id),
                    case_sensitive INTEGER DEFAULT 0,
                    FOREIGN KEY (exam_id) REFERENCES exams(id),
                    FOREIGN KEY (section_id) REFERENCES sections(id)
                )
            ''')
            conn.execute('''
                INSERT INTO questions_essay_new
                    (id, exam_id, section_id, question_text, question_type, points,
                     correct_answer, order_index, bank_group_id, is_bank_only, teacher_id, case_sensitive)
                SELECT id, exam_id, section_id, question_text, question_type, points,
                       correct_answer, order_index, bank_group_id,
                       COALESCE(is_bank_only, 0), teacher_id, COALESCE(case_sensitive, 0)
                FROM questions
            ''')
            conn.execute('DROP TABLE questions')
            conn.execute('ALTER TABLE questions_essay_new RENAME TO questions')
            conn.execute('PRAGMA foreign_keys=ON')
            conn.commit()
    except Exception:
        try:
            conn.execute('PRAGMA foreign_keys=ON')
        except Exception:
            pass  # Already widened or migration failed gracefully

    # Migration: widen questions.question_type CHECK constraint to allow 'true_false'.
    # Same recreate-table pattern used above.
    try:
        tbl_sql = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='questions'"
        ).fetchone()
        if tbl_sql and 'true_false' not in (tbl_sql['sql'] or ''):
            conn.execute('PRAGMA foreign_keys=OFF')
            conn.execute('''
                CREATE TABLE questions_tf_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    exam_id INTEGER,
                    section_id INTEGER,
                    question_text TEXT NOT NULL,
                    question_type TEXT NOT NULL CHECK(question_type IN ('multiple_choice','short_answer','fill_blank','essay','true_false')),
                    points INTEGER DEFAULT 1,
                    correct_answer TEXT,
                    order_index INTEGER DEFAULT 0,
                    bank_group_id INTEGER REFERENCES question_bank_groups(id),
                    is_bank_only INTEGER DEFAULT 0,
                    teacher_id INTEGER REFERENCES users(id),
                    case_sensitive INTEGER DEFAULT 0,
                    FOREIGN KEY (exam_id) REFERENCES exams(id),
                    FOREIGN KEY (section_id) REFERENCES sections(id)
                )
            ''')
            conn.execute('''
                INSERT INTO questions_tf_new
                    (id, exam_id, section_id, question_text, question_type, points,
                     correct_answer, order_index, bank_group_id, is_bank_only, teacher_id, case_sensitive)
                SELECT id, exam_id, section_id, question_text, question_type, points,
                       correct_answer, order_index, bank_group_id,
                       COALESCE(is_bank_only, 0), teacher_id, COALESCE(case_sensitive, 0)
                FROM questions
            ''')
            conn.execute('DROP TABLE questions')
            conn.execute('ALTER TABLE questions_tf_new RENAME TO questions')
            conn.execute('PRAGMA foreign_keys=ON')
            conn.commit()
    except Exception:
        try:
            conn.execute('PRAGMA foreign_keys=ON')
        except Exception:
            pass  # Already widened or migration failed gracefully

    # Migration: Essay questions + Rubrics + Manual Grading
    # ─────────────────────────────────────────────────────
    # Rubric criteria live on the question itself (defined once by the teacher).
    conn.execute('''
        CREATE TABLE IF NOT EXISTS rubric_criteria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question_id INTEGER NOT NULL,
            criterion_text TEXT NOT NULL,
            max_points INTEGER NOT NULL DEFAULT 1,
            order_index INTEGER DEFAULT 0,
            FOREIGN KEY (question_id) REFERENCES questions(id)
        )
    ''')
    # Per-criterion score awarded to one specific student answer.
    conn.execute('''
        CREATE TABLE IF NOT EXISTS rubric_scores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            answer_id INTEGER NOT NULL,
            criterion_id INTEGER NOT NULL,
            score REAL NOT NULL DEFAULT 0,
            FOREIGN KEY (answer_id) REFERENCES answers(id),
            FOREIGN KEY (criterion_id) REFERENCES rubric_criteria(id),
            UNIQUE(answer_id, criterion_id)
        )
    ''')
    conn.commit()

    # answers: fields needed to record a manually-graded essay score/feedback
    for col_sql in (
        'ALTER TABLE answers ADD COLUMN manual_score REAL',
        'ALTER TABLE answers ADD COLUMN feedback TEXT',
        'ALTER TABLE answers ADD COLUMN graded_by INTEGER',
        'ALTER TABLE answers ADD COLUMN graded_at TIMESTAMP',
    ):
        try:
            conn.execute(col_sql)
            conn.commit()
        except Exception:
            pass  # Column already exists

    # exam_sessions: tracks whether this submission still has ungraded essay
    # answers, so the student sees "pending" instead of a final score, and the
    # teacher's results list can flag which submissions need attention.
    try:
        conn.execute('ALTER TABLE exam_sessions ADD COLUMN needs_grading INTEGER DEFAULT 0')
        conn.commit()
    except Exception:
        pass  # Column already exists

    # classes.restrict_join: explicit "Open to Anyone" vs "Restricted to
    # Allowed List" switch for joining by class code. Previously this was
    # only ever *implied* by whether class_allowed_emails had any rows,
    # which meant a teacher couldn't prepare an allow-list in advance (or
    # temporarily open the class) without permanently losing/clearing it.
    try:
        conn.execute('ALTER TABLE classes ADD COLUMN restrict_join INTEGER DEFAULT 0')
        conn.commit()
        # One-time backfill for existing classes: anything that already has
        # allowed emails was, under the old implicit rule, already
        # restricted — keep that behavior unchanged for those classes.
        conn.execute('''
            UPDATE classes SET restrict_join = 1
            WHERE id IN (SELECT DISTINCT class_id FROM class_allowed_emails)
        ''')
        conn.commit()
    except Exception:
        pass  # Column already exists


def get_maintenance_db():
    """Standalone sqlite connection for the background maintenance thread, which
    runs outside any Flask request context and therefore can't use get_db()/g."""
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA synchronous=NORMAL')
    conn.execute('PRAGMA busy_timeout=5000')
    return conn


def auto_activate_scheduled_exams(conn):
    """Auto-open any exams whose scheduled_at has arrived and are still 'upcoming'.
    Only auto-opens if the scheduled time is within the SAME DAY (today).
    Past-day scheduled exams are left closed — teacher must open them manually.
    scheduled_at is stored as the teacher's local time (from datetime-local input),
    so we compare against local time, not UTC.
    """
    try:
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
              AND is_ready = 1
              AND EXISTS (SELECT 1 FROM questions q WHERE q.exam_id = exams.id)
              AND substr(replace(scheduled_at, 'T', ' '), 1, 10) = ?
              AND substr(replace(scheduled_at, 'T', ' '), 1, 16) <= ?
        """, (today_str, now_str))
        conn.commit()
    except Exception:
        pass


def auto_complete_finished_exams(conn):
    """Mark an 'active' exam as 'completed' once its scheduled window (duration
    plus a 15-minute grace period for stragglers) has fully elapsed and no
    student is still mid-exam. Exams the teacher closed manually (status flips
    to 'upcoming' via the toggle button) are untouched by this — it only ever
    moves 'active' -> 'completed', never touches 'upcoming'.
    """
    try:
        conn.execute("""
            UPDATE exams
            SET status = 'completed'
            WHERE status = 'active'
              AND activated_at IS NOT NULL
              AND datetime(activated_at, '+' || duration_minutes || ' minutes', '+15 minutes')
                  <= datetime('now', 'localtime')
              AND id NOT IN (SELECT DISTINCT exam_id FROM exam_sessions WHERE status = 'ongoing')
        """)
        conn.commit()
    except Exception:
        pass


def background_maintenance_loop(interval_seconds=20):
    """Runs exam scheduling/completion checks on a timer instead of on every
    single HTTP request. Uses its own sqlite connection since it runs outside
    any Flask request context."""
    while True:
        try:
            conn = get_maintenance_db()
            try:
                auto_activate_scheduled_exams(conn)
                auto_complete_finished_exams(conn)
            finally:
                conn.close()
        except Exception:
            pass
        time.sleep(interval_seconds)

def hash_password(password):
    """Salted, slow hash (werkzeug scrypt by default)."""
    return generate_password_hash(password)


def _is_legacy_hash(stored):
    """Old accounts used a bare unsalted SHA-256 hex digest (64 hex chars)."""
    return bool(re.fullmatch(r'[0-9a-f]{64}', stored or ''))


def verify_password(password, stored):
    """Check a password against a stored hash. Still accepts legacy SHA-256
    hashes so existing users can log in; login() upgrades them automatically."""
    if not stored:
        return False
    if _is_legacy_hash(stored):
        legacy = hashlib.sha256(password.encode()).hexdigest()
        return hmac.compare_digest(legacy, stored)
    return check_password_hash(stored, password)


def validate_password_strength(password):
    """Returns None if the password is strong enough, or an error message
    otherwise. Rule: at least 8 characters, at least 1 special character,
    AND at least 3 of the following: lowercase letters, uppercase letters,
    numbers (i.e. all three) — matches the live checklist shown next to
    every "set a new password" field in the UI."""
    if len(password) < 8:
        return 'Password must be at least 8 characters long.'
    if not re.search(r'[^A-Za-z0-9]', password):
        return 'Password must contain at least 1 special character.'
    classes_met = sum([
        bool(re.search(r'[a-z]', password)),
        bool(re.search(r'[A-Z]', password)),
        bool(re.search(r'[0-9]', password)),
    ])
    if classes_met < 3:
        return ('Password must contain at least 3 of the following: lowercase letters, '
                'uppercase letters, numbers.')
    return None

DEFAULT_TIME_WARNING_MINUTES = 5

def clean_time_warning(raw, duration):
    """Minutes before the end at which students get the "time left" popup.

    Blank/invalid falls back to the default (5). 0 turns the popup off. It can't
    be longer than the exam itself (otherwise it would pop up the moment the
    exam starts), so it is capped at duration - 1.
    """
    try:
        minutes = int(raw)
    except (TypeError, ValueError):
        minutes = DEFAULT_TIME_WARNING_MINUTES
    return max(0, min(minutes, max((duration or 0) - 1, 0)))

def handle_profile_update(conn, user, with_year_level=False):
    """Update the editable parts of a profile (name, and year level for students).

    Returns (ok, profile_errors). Flashes the outcome.
    """
    full_name = ' '.join(request.form.get('full_name', '').split())
    errs = {}
    if not full_name:
        errs['full_name'] = True
    elif len(full_name) > 100:
        errs['full_name'] = True
        flash('Name must be 100 characters or fewer.', 'error')
        return False, errs
    year_level = None
    if with_year_level:
        year_level = request.form.get('year_level', '').strip()
        if year_level not in ('1st Year', '2nd Year', '3rd Year', '4th Year'):
            errs['year_level'] = True
    if errs:
        if 'Name must' not in ''.join(m for _, m in session.get('_flashes', [])):
            flash('Please fill in your name' + (' and select a valid year level.' if with_year_level else '.'), 'error')
        return False, errs
    if with_year_level:
        conn.execute('UPDATE users SET full_name=?, year_level=? WHERE id=?',
                     (full_name, year_level, user['id']))
    else:
        conn.execute('UPDATE users SET full_name=? WHERE id=?', (full_name, user['id']))
    conn.commit()
    session['full_name'] = full_name
    flash('Profile updated successfully.', 'success')
    return True, {}

def handle_password_change(conn, user):
    """Shared "Change Password" form handler for the teacher / student profile
    pages (same rules as the admin profile page).

    Returns (form, errors, success). On success the password is updated and
    `success` is True; otherwise an error is flashed and `errors` marks the
    offending fields so the template can highlight them.
    """
    current_password = request.form.get('current_password', '')
    new_password = request.form.get('new_password', '')
    confirm_password = request.form.get('confirm_password', '')
    # Keep what was typed so a failed attempt doesn't wipe the form
    form = {'current_password': current_password, 'new_password': new_password,
            'confirm_password': confirm_password}
    errors = {}

    if not all([current_password, new_password, confirm_password]):
        flash('Please fill in all password fields.', 'error')
        for key, val in form.items():
            if not val:
                errors[key] = True
    elif not verify_password(current_password, user['password']):
        flash('Current password is incorrect.', 'error')
        errors['current_password'] = True
    elif new_password != confirm_password:
        flash('New passwords do not match.', 'error')
        errors['confirm_password'] = True
    elif new_password == current_password:
        flash('New password must be different from the current password.', 'error')
        errors['new_password'] = True
    else:
        pw_error = validate_password_strength(new_password)
        if pw_error:
            flash(pw_error, 'error')
            errors['new_password'] = True
        else:
            conn.execute('UPDATE users SET password=?, must_change_password=0 WHERE id=?',
                         (hash_password(new_password), user['id']))
            conn.commit()
            session.pop('must_change_pw', None)  # lift the forced-change lock
            flash('Password updated successfully.', 'success')
            return {}, {}, True
    return form, errors, False

# ─── System Settings (key/value store, admin-editable) ──────────────────────

SETTINGS_DEFAULTS = {
    'allow_signup': '1',
    'allow_teacher_signup': '1',
    'lockout_max_attempts': '5',
    'lockout_duration_minutes': '3',
}


def get_settings():
    """Current system settings as a dict, with defaults filled in for any
    key that hasn't been saved yet. Booleans/ints are already converted."""
    conn = get_db()
    rows = conn.execute('SELECT key, value FROM settings').fetchall()
    raw = {**SETTINGS_DEFAULTS, **{r['key']: r['value'] for r in rows}}
    return {
        'allow_signup': raw['allow_signup'] == '1',
        'allow_teacher_signup': raw['allow_teacher_signup'] == '1',
        'lockout_max_attempts': int(raw['lockout_max_attempts']),
        'lockout_duration_minutes': int(raw['lockout_duration_minutes']),
    }


def set_settings(values):
    """values: {key: value} — saves/overwrites each key."""
    conn = get_db()
    for key, value in values.items():
        conn.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?) '
            'ON CONFLICT(key) DO UPDATE SET value=excluded.value',
            (key, str(value))
        )
    conn.commit()

# ─── Auth Decorators ─────────────────────────────────────────────────────────

@app.before_request
def force_password_change():
    """While the account is flagged, the only pages allowed are the profile
    page (to change the password), logout, and static files."""
    if session.get('user_id') and session.get('must_change_pw'):
        profile_endpoint = f"{session.get('role')}_profile"
        if request.endpoint not in (profile_endpoint, 'logout', 'static'):
            flash('For security, you must change your temporary password before continuing.', 'error')
            return redirect(url_for(profile_endpoint))

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
    page_settings = get_settings()
    form = {}
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        # Keep what the person typed so a failed attempt doesn't wipe the form.
        # (The password itself is deliberately not echoed back.)
        form = {'email': email}
        if not email or not password:
            flash('Please fill in all fields.', 'error')
            return render_template('login.html', form=form, settings=page_settings)
        conn = get_db()
        account = conn.execute('SELECT * FROM users WHERE email = ?', (email,)).fetchone()

        # Account lockout: if this account is currently locked out, refuse
        # the attempt before even checking the password, and don't restart
        # the clock — that would let someone keep the account locked forever
        # just by repeatedly trying.
        if account and account['locked_until']:
            locked_until = datetime.strptime(account['locked_until'], '%Y-%m-%d %H:%M:%S.%f') \
                if '.' in account['locked_until'] else datetime.strptime(account['locked_until'], '%Y-%m-%d %H:%M:%S')
            if locked_until > datetime.now():
                minutes_left = max(1, int((locked_until - datetime.now()).total_seconds() // 60) + 1)
                conn.execute('INSERT INTO login_logs (user_id, email, success) VALUES (?,?,0)',
                             (account['id'], email))
                conn.commit()
                flash(f'This account is locked due to too many failed login attempts. '
                      f'Try again in {minutes_left} minute(s).', 'error')
                return render_template('login.html', form=form, settings=page_settings)
            else:
                # The lockout has run out: start fresh with a full set of attempts
                # (otherwise the old failed-attempt count would re-lock the account
                # on the very next wrong password).
                conn.execute('UPDATE users SET failed_login_attempts=0, locked_until=NULL WHERE id=?',
                             (account['id'],))
                conn.commit()
                account = conn.execute('SELECT * FROM users WHERE email = ?', (email,)).fetchone()

        user = account
        if user and not verify_password(password, user['password']):
            user = None

        if user:
            # Silently upgrade old unsalted SHA-256 hashes to the new format.
            if _is_legacy_hash(user['password']):
                conn.execute('UPDATE users SET password=? WHERE id=?',
                             (hash_password(password), user['id']))
            # Successful login clears any lockout bookkeeping.
            conn.execute('UPDATE users SET failed_login_attempts=0, locked_until=NULL WHERE id=?',
                         (user['id'],))
            conn.execute('INSERT INTO login_logs (user_id, email, success) VALUES (?,?,1)',
                         (user['id'], email))
            conn.commit()
            session['user_id'] = user['id']
            session['role'] = user['role']
            session['full_name'] = user['full_name']
            session['email'] = user['email']
            session['must_change_pw'] = bool(user['must_change_password'])
            return redirect(url_for(f"{user['role']}_home"))
        else:
            conn.execute('INSERT INTO login_logs (email, success) VALUES (?,0)', (email,))
            if account:
                settings = get_settings()
                attempts = (account['failed_login_attempts'] or 0) + 1
                if attempts >= settings['lockout_max_attempts']:
                    locked_until = datetime.now() + timedelta(minutes=settings['lockout_duration_minutes'])
                    conn.execute('UPDATE users SET failed_login_attempts=?, locked_until=? WHERE id=?',
                                 (attempts, locked_until.strftime('%Y-%m-%d %H:%M:%S'), account['id']))
                    conn.commit()
                    flash(f'Too many failed attempts. This account is now locked for '
                          f'{settings["lockout_duration_minutes"]} minute(s).', 'error')
                else:
                    conn.execute('UPDATE users SET failed_login_attempts=? WHERE id=?',
                                 (attempts, account['id']))
                    conn.commit()
                    remaining = settings['lockout_max_attempts'] - attempts
                    flash(f'Invalid email or password. {remaining} attempt(s) remaining before '
                          f'this account is locked.', 'error')
            else:
                conn.commit()
                flash('Invalid email or password.', 'error')
    return render_template('login.html', form=form, settings=page_settings)

ALLOWED_SIGNUP_EMAIL_DOMAIN = '@psu.palawan.edu.ph'


def _is_allowed_signup_email(email):
    return email.strip().lower().endswith(ALLOWED_SIGNUP_EMAIL_DOMAIN)


SIGNUP_REQUIRED_STUDENT = ['full_name', 'email', 'password', 'confirm_password', 'program', 'year_level']
SIGNUP_REQUIRED_TEACHER = ['full_name', 'email', 'password', 'confirm_password']


def _validate_signup_form(form, required):
    """Returns {field_name: True} for every field that failed validation
    (empty, bad email domain, weak password, passwords not matching).
    The signup pages turn exactly those fields/requirements red instead of
    showing a message at the top of the page."""
    errors = {}
    for key in required:
        if not form.get(key):
            errors[key] = True
    email = form.get('email', '')
    if email and not _is_allowed_signup_email(email):
        errors['email'] = True
    password = form.get('password', '')
    if password and validate_password_strength(password):
        errors['password'] = True
    confirm = form.get('confirm_password', '')
    if password and confirm and password != confirm:
        errors['confirm_password'] = True
    return errors


@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if 'user_id' in session:
        return redirect(url_for(f"{session['role']}_home"))
    if not get_settings()['allow_signup']:
        flash('Student self-registration is currently disabled. Please contact your admin to have an account created.', 'error')
        return redirect(url_for('login'))
    conn = get_db()
    programs = conn.execute('SELECT * FROM programs ORDER BY code').fetchall()
    form, errors = {}, {}
    if request.method == 'POST':
        form = {
            'full_name': request.form.get('full_name', '').strip(),
            'email': request.form.get('email', '').strip(),
            'password': request.form.get('password', ''),
            'confirm_password': request.form.get('confirm_password', ''),
            'program': request.form.get('program', ''),
            'year_level': request.form.get('year_level', ''),
        }
        errors = _validate_signup_form(form, SIGNUP_REQUIRED_STUDENT)
        if not errors:
            try:
                conn = get_db()
                conn.execute('''
                    INSERT INTO users (full_name, email, password, role, program, year_level)
                    VALUES (?, ?, ?, 'student', ?, ?)
                ''', (form['full_name'], form['email'], hash_password(form['password']),
                      form['program'], form['year_level']))
                conn.commit()
                flash('Account created! You can now log in.', 'success')
                return redirect(url_for('login'))
            except sqlite3.IntegrityError:
                errors['email'] = True
                errors['email_msg'] = 'Email already exists.'
    return render_template('signup.html', programs=programs, form=form, errors=errors)

@app.route('/signup/teacher', methods=['GET', 'POST'])
def signup_teacher():
    if 'user_id' in session:
        return redirect(url_for(f"{session['role']}_home"))
    if not get_settings()['allow_teacher_signup']:
        flash('Teacher self-registration is currently disabled. Please contact your admin to have an account created.', 'error')
        return redirect(url_for('login'))
    form, errors = {}, {}
    if request.method == 'POST':
        form = {
            'full_name': request.form.get('full_name', '').strip(),
            'email': request.form.get('email', '').strip(),
            'password': request.form.get('password', ''),
            'confirm_password': request.form.get('confirm_password', ''),
        }
        errors = _validate_signup_form(form, SIGNUP_REQUIRED_TEACHER)
        if not errors:
            try:
                conn = get_db()
                conn.execute('''
                    INSERT INTO users (full_name, email, password, role)
                    VALUES (?, ?, ?, 'teacher')
                ''', (form['full_name'], form['email'], hash_password(form['password'])))
                conn.commit()
                flash('Teacher account created! You can now log in.', 'success')
                return redirect(url_for('login'))
            except sqlite3.IntegrityError:
                errors['email'] = True
                errors['email_msg'] = 'Email already exists.'
    return render_template('signup_teacher.html', form=form, errors=errors)

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
        ORDER BY ce.joined_at DESC, ce.id DESC
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
            # The teacher controls this explicitly via the "Open to Anyone" /
            # "Restricted to Allowed List" switch on the class's Allowed
            # Emails tab. When restricted, only students whose email is on
            # the list may join (an empty list while restricted means no one
            # can join yet — correct, if unhelpful, until emails are added).
            if cls['restrict_join']:
                student_email = (session.get('email') or '').strip().lower()
                is_allowed = conn.execute(
                    'SELECT 1 FROM class_allowed_emails WHERE class_id=? AND email=?',
                    (cls['id'], student_email)
                ).fetchone()
                if not is_allowed:
                    flash('Your email is not on this class\'s allowed list. Please check with your teacher.', 'error')
                    return render_template('student/join_class.html')
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
        # Carry the code through as a query param: student_take_exam re-validates
        # it on GET for any brand-new session, so forgetting it here would bounce
        # the student straight back out with "Incorrect exam code" even though
        # they just entered the right one.
        return redirect(url_for('student_take_exam', exam_id=exam['id'], code=exam['exam_code']))
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
    # A closed exam blocks NEW entries (GET), but a student who is already
    # mid-exam must still be able to submit (manually or via the timer's
    # auto-submit). Otherwise their session stays 'ongoing' with no score.
    if not exam or (exam['status'] != 'active' and request.method != 'POST'):
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
        exam_sess = conn.execute('SELECT * FROM exam_sessions WHERE id=? AND student_id=? AND exam_id=?',
                                 (sess_id, session['user_id'], exam_id)).fetchone()
        if not exam_sess or exam_sess['status'] != 'ongoing':
            flash('Invalid session.', 'error')
            return redirect(url_for('student_exams'))

        questions = conn.execute('SELECT * FROM questions WHERE exam_id=?', (exam_id,)).fetchall()
        total_points = sum(q['points'] for q in questions)
        score = 0
        has_essay = any(q['question_type'] == 'essay' for q in questions)

        for q in questions:
            ans = request.form.get(f'answer_{q["id"]}', '').strip()
            conn.execute('''
                INSERT OR REPLACE INTO answers (session_id, question_id, answer_text)
                VALUES (?, ?, ?)
            ''', (sess_id, q['id'], ans))
            # Fill-in-the-blank earns partial credit per blank; everything
            # else is all-or-nothing. Case-sensitivity is honoured. Essay
            # questions always contribute 0 here — they earn no credit until
            # a teacher manually grades them against the rubric.
            score += q['points'] * question_credit(q, ans)

        # Essay questions mean this submission's score isn't final yet — the
        # student sees "pending" until every essay answer has been graded.
        conn.execute('''
            UPDATE exam_sessions SET status='submitted', submitted_at=datetime('now','localtime'),
            score=?, total_points=?, needs_grading=? WHERE id=?
        ''', (score, total_points, 1 if has_essay else 0, sess_id))
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
                if q['question_type'] == 'fill_blank':
                    q['text_parts'] = fib_split_text(q['question_text'])
                all_questions.append(q)
            sections_data.append({'title': sec['title'], 'description': sec['description'], 'questions': q_list})
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
            if q['question_type'] == 'fill_blank':
                q['text_parts'] = fib_split_text(q['question_text'])
            all_questions.append(q)
        sections_data.append({'title': None, 'description': None, 'questions': q_list})

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

    # Still mid-exam: there is no result yet, send them back to the exam
    # instead of showing "0 / None — FAILED".
    if exam_sess['status'] == 'ongoing':
        return redirect(url_for('student_take_exam', exam_id=exam_id))

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
            is_case_sensitive = bool(_q_get(q, 'case_sensitive', 0))
            is_correct = False
            if q['question_type'] == 'multiple_choice':
                is_correct = question_credit(q, qd['student_answer']) >= 1.0
            elif q['question_type'] == 'fill_blank':
                correct_blanks, total_blanks = fib_grade(qd['student_answer'], q['correct_answer'], is_case_sensitive)
                is_correct = total_blanks > 0 and correct_blanks == total_blanks
                qd['fib_correct_blanks'] = correct_blanks
                qd['fib_total_blanks'] = total_blanks
                qd['fib_points_earned'] = round(q['points'] * (correct_blanks / total_blanks), 2) if total_blanks else 0
                # Per-blank breakdown for the review UI
                blanks = fib_parse_answer(q['correct_answer'])
                student_parts = (qd['student_answer'] or '').split('|')
                blank_results = []
                for i, alts in enumerate(blanks):
                    given = student_parts[i].strip() if i < len(student_parts) else ''
                    if is_case_sensitive:
                        b_correct = any(given == alt for alt in alts) if given else False
                    else:
                        b_correct = any(given.lower() == alt.lower() for alt in alts) if given else False
                    blank_results.append({
                        'index': i + 1,
                        'student': given,
                        'accepted': ' / '.join(alts),
                        'is_correct': b_correct,
                    })
                qd['fib_blanks'] = blank_results
                qd['text_parts'] = fib_split_text(q['question_text'])
            elif q['question_type'] == 'essay':
                qd['is_graded'] = bool(ans and ans['manual_score'] is not None)
                qd['manual_score'] = ans['manual_score'] if ans else None
                qd['feedback'] = ans['feedback'] if ans else ''
                is_correct = None  # not applicable for essay — shown as graded/pending instead
            else:
                # Same grader as submit-time scoring, so the review always
                # agrees with the score (incl. case-sensitive questions).
                is_correct = question_credit(q, qd['student_answer']) >= 1.0
            qd['is_correct'] = is_correct
            questions.append(qd)

    return render_template('student/exam_result.html', exam=exam, exam_session=exam_sess, questions=questions, class_id=exam['class_id'], exam_closed=exam_closed)

@app.route('/student/profile', methods=['GET', 'POST'])
@role_required('student')
def student_profile():
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
    form, errors, profile_errors = {}, {}, {}
    if request.method == 'POST':
        if request.form.get('action') == 'update_profile':
            ok, profile_errors = handle_profile_update(conn, user, with_year_level=True)
            if ok:
                return redirect(url_for('student_profile'))
            user = dict(user)
            user['full_name'] = request.form.get('full_name', '')
            user['year_level'] = request.form.get('year_level', '')
        else:
            form, errors, ok = handle_password_change(conn, user)
            if ok:
                return redirect(url_for('student_profile'))
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
    return render_template('student/profile.html', user=user, exam_history=exam_history,
                           form=form, errors=errors, profile_errors=profile_errors)

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
        ORDER BY c.id DESC
    ''', (session['user_id'],)).fetchall()
    programs = conn.execute('SELECT * FROM programs ORDER BY code').fetchall()
    return render_template('teacher/classes.html', classes=classes, programs=programs)

@app.route('/teacher/class/<int:class_id>/edit', methods=['POST'])
@role_required('teacher')
def teacher_edit_class(class_id):
    conn = get_db()
    cls = conn.execute(
        'SELECT * FROM classes WHERE id = ? AND teacher_id = ?',
        (class_id, session['user_id'])
    ).fetchone()
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    subject_code = request.form.get('subject_code', '').strip().upper()
    subject_name = request.form.get('subject_name', '').strip()
    block_name   = request.form.get('block_name', '').strip().upper()
    program      = request.form.get('program', '').strip()
    year_level   = request.form.get('year_level', '').strip()
    if not all([subject_code, subject_name, block_name, program, year_level]):
        flash('Please fill in all fields.', 'error')
        return redirect(url_for('teacher_home'))
    conn.execute('''
        UPDATE classes SET subject_code=?, subject_name=?, block_name=?, program=?, year_level=?
        WHERE id=?
    ''', (subject_code, subject_name, block_name, program, year_level, class_id))
    conn.commit()
    flash(f'"{subject_name}" has been updated.', 'success')
    return redirect(url_for('teacher_home'))

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
        ORDER BY LOWER(u.full_name)
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

    allowed_emails = conn.execute(
        'SELECT * FROM class_allowed_emails WHERE class_id=? ORDER BY email', (class_id,)
    ).fetchall()
    # Mark which allowed emails have already joined, so the teacher can see
    # who from their list hasn't enrolled yet.
    enrolled_emails = {s['email'].lower() for s in students}
    allowed_emails_view = []
    for a in allowed_emails:
        allowed_emails_view.append({'email': a['email'], 'joined': a['email'].lower() in enrolled_emails})

    # One-time notice carrying a freshly generated temporary password. Kept
    # out of flash() because flash alerts auto-dismiss after a few seconds,
    # and the teacher needs time to read it out to the student.
    reset_notice = session.pop('reset_notice', None)
    if reset_notice and reset_notice.get('class_id') != class_id:
        reset_notice = None

    return render_template('teacher/class_detail.html', cls=cls, students=students, exams=exams,
                           allowed_emails=allowed_emails_view, reset_notice=reset_notice)

def _get_owned_class(conn, class_id):
    """Fetch a class only if it belongs to the currently logged-in teacher."""
    return conn.execute(
        'SELECT * FROM classes WHERE id = ? AND teacher_id = ?',
        (class_id, session['user_id'])
    ).fetchone()


_EMAIL_RE = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')


@app.route('/teacher/class/<int:class_id>/allowed-emails/add', methods=['POST'])
@role_required('teacher')
def teacher_add_allowed_emails(class_id):
    """Lets a teacher add one or more student emails to their class's
    allow-list (typed/pasted directly, one per line or comma-separated).
    Once a class has at least one allowed email, only students whose email
    is on this list can join it with the class code."""
    conn = get_db()
    cls = _get_owned_class(conn, class_id)
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))

    raw = request.form.get('emails', '')
    found = sorted(set(m.lower() for m in _EMAIL_RE.findall(raw)))
    if not found:
        flash('No valid email addresses found. Separate multiple emails with a comma or a new line.', 'error')
        return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')

    added = 0
    for email in found:
        try:
            conn.execute('INSERT INTO class_allowed_emails (class_id, email) VALUES (?,?)', (class_id, email))
            added += 1
        except sqlite3.IntegrityError:
            pass  # already on the list
    conn.commit()
    skipped = len(found) - added
    msg = f'Added {added} email{"s" if added != 1 else ""} to the allowed list.'
    if skipped:
        msg += f' {skipped} were already on it.'
    flash(msg, 'success' if added else 'info')
    return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')


@app.route('/teacher/class/<int:class_id>/allowed-emails/import', methods=['POST'])
@role_required('teacher')
def teacher_import_allowed_emails(class_id):
    """Bulk-adds allowed emails from an uploaded .txt, .csv, .docx or .xlsx
    file — scans every line/cell/paragraph for anything that looks like an
    email address, so it works with whatever list format the teacher
    already has on hand (e.g. a class list exported from the registrar)."""
    conn = get_db()
    cls = _get_owned_class(conn, class_id)
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))

    uploaded = request.files.get('emails_file')
    if not uploaded or uploaded.filename == '':
        flash('No file selected.', 'error')
        return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')

    if _upload_too_large(uploaded):
        flash(f'File is too large. The maximum allowed size is {MAX_UPLOAD_MB} MB.', 'error')
        return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')

    raw_name = uploaded.filename
    ext = raw_name.rsplit('.', 1)[-1].lower() if '.' in raw_name else ''
    if ext not in ('txt', 'csv', 'docx', 'xlsx'):
        flash('Unsupported file type. Please upload a .txt, .csv, .docx, or .xlsx file.', 'error')
        return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')

    text_blob = ''
    try:
        if ext in ('txt', 'csv'):
            text_blob = uploaded.read().decode('utf-8', errors='replace')
        elif ext == 'docx':
            import docx
            document = docx.Document(uploaded)
            text_blob = '\n'.join(p.text for p in document.paragraphs)
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text_blob += '\n' + cell.text
        elif ext == 'xlsx':
            wb = load_workbook(uploaded, data_only=True, read_only=True)
            parts = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for val in row:
                        if val:
                            parts.append(str(val))
            text_blob = '\n'.join(parts)
    except Exception:
        flash('Could not read that file. Make sure it is a valid, uncorrupted file of that type.', 'error')
        return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')

    found = sorted(set(m.lower() for m in _EMAIL_RE.findall(text_blob)))
    if not found:
        flash('No email addresses were found in that file.', 'error')
        return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')

    added = 0
    for email in found:
        try:
            conn.execute('INSERT INTO class_allowed_emails (class_id, email) VALUES (?,?)', (class_id, email))
            added += 1
        except sqlite3.IntegrityError:
            pass
    conn.commit()
    skipped = len(found) - added
    msg = f'Found {len(found)} email{"s" if len(found) != 1 else ""} in the file — added {added} new.'
    if skipped:
        msg += f' {skipped} were already on the list.'
    flash(msg, 'success' if added else 'info')
    return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')


@app.route('/teacher/class/<int:class_id>/allowed-emails/remove', methods=['POST'])
@role_required('teacher')
def teacher_remove_allowed_email(class_id):
    conn = get_db()
    cls = _get_owned_class(conn, class_id)
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    email = request.form.get('email', '').strip().lower()
    conn.execute('DELETE FROM class_allowed_emails WHERE class_id=? AND email=?', (class_id, email))
    conn.commit()
    flash(f'Removed {email} from the allowed list.', 'success')
    return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')


@app.route('/teacher/class/<int:class_id>/allowed-emails/clear', methods=['POST'])
@role_required('teacher')
def teacher_clear_allowed_emails(class_id):
    """Empties the allow-list (just the emails — does not change whether the
    class is currently set to Open or Restricted). Handy for starting a
    fresh list without re-uploading everyone, or clearing out a stale one."""
    conn = get_db()
    cls = _get_owned_class(conn, class_id)
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    conn.execute('DELETE FROM class_allowed_emails WHERE class_id=?', (class_id,))
    conn.commit()
    flash('Cleared the allowed-emails list.', 'success')
    return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')


@app.route('/teacher/class/<int:class_id>/allowed-emails/toggle', methods=['POST'])
@role_required('teacher')
def teacher_toggle_class_restriction(class_id):
    """Explicit switch between 'Open to Anyone with the code' and
    'Restricted to the allowed-emails list'. Kept separate from the list
    itself, so a teacher can prep an allow-list in advance, or temporarily
    reopen a class, without losing it."""
    conn = get_db()
    cls = _get_owned_class(conn, class_id)
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    new_state = 0 if cls['restrict_join'] else 1
    conn.execute('UPDATE classes SET restrict_join=? WHERE id=?', (new_state, class_id))
    conn.commit()
    if new_state:
        flash('This class is now restricted — only emails on the allowed list can join.', 'success')
    else:
        flash('This class is now open — anyone with the class code can join.', 'success')
    return redirect(url_for('teacher_class_detail', class_id=class_id) + '#allowed-emails')


@app.route('/teacher/class/<int:class_id>/students/remove/<int:student_id>', methods=['POST'])
@role_required('teacher')
def teacher_remove_student(class_id, student_id):
    """Removes one student's enrollment from this class (does not touch the
    allowed-emails list). The student can rejoin later with the class code —
    immediately if the class is open, or only if their email is still on the
    allowed list if it's restricted."""
    conn = get_db()
    cls = _get_owned_class(conn, class_id)
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    student = conn.execute('SELECT full_name FROM users WHERE id=?', (student_id,)).fetchone()
    conn.execute('DELETE FROM class_enrollments WHERE class_id=? AND student_id=?', (class_id, student_id))
    conn.commit()
    name = student['full_name'] if student else 'Student'
    flash(f'{name} was removed from the class.', 'success')
    return redirect(url_for('teacher_class_detail', class_id=class_id) + '#students')


def _generate_temp_password():
    """Random temporary password that always satisfies the app's strength
    rules (upper + lower + number + special). Avoids look-alike characters
    (0/O, 1/l/I) since the teacher may read it out loud to the student."""
    upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    lower = 'abcdefghijkmnpqrstuvwxyz'
    digits = '23456789'
    special = '!@#$%&*?'
    chars = [secrets.choice(upper), secrets.choice(lower),
             secrets.choice(digits), secrets.choice(special)]
    pool = upper + lower + digits
    chars += [secrets.choice(pool) for _ in range(6)]
    secrets.SystemRandom().shuffle(chars)
    return ''.join(chars)


@app.route('/teacher/class/<int:class_id>/students/reset-password/<int:student_id>', methods=['POST'])
@role_required('teacher')
def teacher_reset_student_password(class_id, student_id):
    """A teacher can reset the password of a student enrolled in one of THEIR
    classes (so a student who forgot their password mid-semester doesn't need
    the admin). Generates a one-time temporary password, forces the student to
    change it at next login, clears any lockout, and writes an audit row."""
    conn = get_db()
    cls = _get_owned_class(conn, class_id)
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))
    student = conn.execute(
        '''SELECT u.* FROM users u
           JOIN class_enrollments ce ON ce.student_id = u.id
           WHERE ce.class_id=? AND u.id=? AND u.role='student' ''',
        (class_id, student_id)
    ).fetchone()
    if not student:
        flash('That student is not enrolled in this class.', 'error')
        return redirect(url_for('teacher_class_detail', class_id=class_id) + '#students')
    temp_password = _generate_temp_password()
    conn.execute(
        '''UPDATE users SET password=?, must_change_password=1,
           failed_login_attempts=0, locked_until=NULL WHERE id=?''',
        (hash_password(temp_password), student_id)
    )
    conn.execute(
        'INSERT INTO password_reset_logs (student_id, teacher_id, class_id) VALUES (?,?,?)',
        (student_id, session['user_id'], class_id)
    )
    conn.commit()
    session['reset_notice'] = {'class_id': class_id, 'name': student['full_name'],
                               'email': student['email'], 'password': temp_password}
    return redirect(url_for('teacher_class_detail', class_id=class_id) + '#students')


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

@app.route('/teacher/class/<int:class_id>/copy', methods=['POST'])
@role_required('teacher')
def teacher_copy_class(class_id):
    conn = get_db()
    cls = conn.execute(
        'SELECT * FROM classes WHERE id = ? AND teacher_id = ?',
        (class_id, session['user_id'])
    ).fetchone()
    if not cls:
        flash('Class not found.', 'error')
        return redirect(url_for('teacher_home'))

    # Generate a fresh, unique class code for the copy
    new_code = f"{cls['subject_code']}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=5))}"
    new_name = f"{cls['subject_name']} (Copy)"

    cur = conn.execute('''
        INSERT INTO classes (class_code, subject_code, subject_name, block_name, program, year_level, teacher_id, is_active)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (new_code, cls['subject_code'], new_name, cls['block_name'], cls['program'], cls['year_level'],
          session['user_id'], cls['is_active']))
    new_class_id = cur.lastrowid

    # Duplicate every exam that belongs to the class, along with its sections,
    # questions and choices. Student-specific data (enrollments, exam sessions,
    # answers, logs) is intentionally NOT copied — the new class starts fresh.
    exams = conn.execute('SELECT * FROM exams WHERE class_id=?', (class_id,)).fetchall()
    for exam in exams:
        new_exam_code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
        ecur = conn.execute('''
            INSERT INTO exams (title, class_id, duration_minutes, scheduled_at, activated_at, status,
                                show_results, randomize_questions, tab_switch_limit, tab_switch_enabled,
                                manually_closed, passing_score, exam_code, fullscreen_required,
                                time_warning_minutes)
            VALUES (?, ?, ?, NULL, NULL, 'upcoming', ?, ?, ?, ?, 0, ?, ?, ?, ?)
        ''', (exam['title'], new_class_id, exam['duration_minutes'],
              exam['show_results'], exam['randomize_questions'], exam['tab_switch_limit'], exam['tab_switch_enabled'],
              exam['passing_score'], new_exam_code, exam['fullscreen_required'],
              exam['time_warning_minutes'] if exam['time_warning_minutes'] is not None else DEFAULT_TIME_WARNING_MINUTES))
        new_exam_id = ecur.lastrowid

        section_id_map = {}
        sections = conn.execute('SELECT * FROM sections WHERE exam_id=?', (exam['id'],)).fetchall()
        for sec in sections:
            scur = conn.execute('''
                INSERT INTO sections (exam_id, title, description, section_type, order_index)
                VALUES (?, ?, ?, ?, ?)
            ''', (new_exam_id, sec['title'], sec['description'] if 'description' in sec.keys() else None,
                  sec['section_type'], sec['order_index']))
            section_id_map[sec['id']] = scur.lastrowid

        questions = conn.execute('SELECT * FROM questions WHERE exam_id=?', (exam['id'],)).fetchall()
        for q in questions:
            new_section_id = section_id_map.get(q['section_id']) if q['section_id'] else None
            qcur = conn.execute('''
                INSERT INTO questions (exam_id, section_id, question_text, question_type, points, correct_answer, order_index, case_sensitive)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (new_exam_id, new_section_id, q['question_text'], q['question_type'], q['points'],
                  q['correct_answer'], q['order_index'], _q_get(q, 'case_sensitive', 0)))
            new_question_id = qcur.lastrowid

            choices = conn.execute('SELECT * FROM choices WHERE question_id=?', (q['id'],)).fetchall()
            for c in choices:
                conn.execute('''
                    INSERT INTO choices (question_id, choice_label, choice_text)
                    VALUES (?, ?, ?)
                ''', (new_question_id, c['choice_label'], c['choice_text']))

    conn.commit()
    flash(f'"{cls["subject_name"]}" has been copied as "{new_name}" (code: {new_code}).', 'success')
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
        fullscreen_required = 1 if request.form.get('fullscreen_required') else 0
        passing_score = request.form.get('passing_score', 50, type=int)
        time_warning = clean_time_warning(request.form.get('time_warning_minutes'), duration)
        if not title or not duration:
            flash('Title and duration are required.', 'error')
            return render_template('teacher/create_exam.html', cls=cls)
        exam_code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
        cur = conn.execute('''
            INSERT INTO exams (title, class_id, duration_minutes, scheduled_at, show_results, randomize_questions, tab_switch_limit, tab_switch_enabled, fullscreen_required, passing_score, exam_code, time_warning_minutes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (title, class_id, duration, scheduled_at or None, show_results, randomize, tab_limit, tab_switch_enabled, fullscreen_required, passing_score, exam_code, time_warning))
        exam_id = cur.lastrowid

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
    locked = _exam_locked_response(exam_id)
    if locked:
        return locked
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
            if q['question_type'] == 'essay':
                criteria = conn.execute(
                    'SELECT * FROM rubric_criteria WHERE question_id=? ORDER BY order_index', (q['id'],)
                ).fetchall()
                qd['rubric_criteria'] = [dict(c) for c in criteria]
            else:
                qd['rubric_criteria'] = []
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
        WHERE q.is_bank_only = 1 AND q.teacher_id = ?
        ORDER BY q.bank_group_id IS NULL ASC, COALESCE(g.name,''), COALESCE(c.subject_name,''), COALESCE(c.block_name,''), e.title
    ''', (session['user_id'],)).fetchall()
    bank_questions = [dict(bq) for bq in raw_bank]
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
               COUNT(DISTINCT CASE WHEN TRIM(COALESCE(a.answer_text,'')) != '' THEN a.question_id END) as answered,
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
        SELECT u.full_name, u.id as student_id, es.score, es.total_points, es.status, es.submitted_at, es.tab_switch_count,
               COALESCE(es.needs_grading, 0) as needs_grading
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
               COALESCE(q.case_sensitive, 0) AS case_sensitive,
               s.title as section_title, q.order_index, s.order_index as sec_order
        FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id = ?
        ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()

    question_stats = build_question_stats(conn, exam_id, questions_raw, total_submitted)

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
        SELECT u.full_name, u.id as student_id, es.score, es.total_points, es.status, es.submitted_at, es.tab_switch_count,
               COALESCE(es.needs_grading, 0) as needs_grading
        FROM exam_sessions es JOIN users u ON es.student_id = u.id
        WHERE es.exam_id=?
        ORDER BY es.score DESC
    ''', (exam_id,)).fetchall()
    # (student_id already selected above — used to link to the full result review)

    # Question analysis — how many students answered each question correctly
    questions_raw = conn.execute('''
        SELECT q.id, q.question_text, q.question_type, q.points, q.correct_answer,
               COALESCE(q.case_sensitive, 0) AS case_sensitive,
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

    question_stats = build_question_stats(conn, exam_id, questions_raw, total_submitted)

    # Per-Section Analytics: aggregate the question stats above by exam section
    # (e.g. "Section A: Multiple Choice"), so a teacher can see which section
    # students struggled with overall, not just individual questions.
    section_order = []
    section_agg = {}
    for q in questions_raw:
        if q['question_type'] == 'essay':
            continue  # manually graded — not part of the auto-graded analytics
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

@app.route('/teacher/exam/<int:exam_id>/student/<int:student_id>/result')
@role_required('teacher')
def teacher_view_student_result(exam_id, student_id):
    # Lets a teacher open a specific student's full answer review from the
    # monitoring/results pages. This ignores the exam's "Show Results to
    # students" setting on purpose — that setting only controls what
    # students themselves can see, not what their teacher can see.
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id, c.subject_name, c.id as class_id
        FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam["teacher_id"]) != int(session["user_id"]):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))

    student = conn.execute('SELECT * FROM users WHERE id=?', (student_id,)).fetchone()
    if not student:
        flash('Student not found.', 'error')
        return redirect(url_for('teacher_exam_monitoring', exam_id=exam_id) + '#results')

    exam_sess = conn.execute(
        'SELECT * FROM exam_sessions WHERE exam_id=? AND student_id=?',
        (exam_id, student_id)
    ).fetchone()
    if not exam_sess:
        flash('This student has not taken this exam yet.', 'error')
        return redirect(url_for('teacher_exam_monitoring', exam_id=exam_id) + '#results')

    if exam_sess['status'] == 'ongoing':
        flash('This student is still taking the exam.', 'error')
        return redirect(url_for('teacher_exam_monitoring', exam_id=exam_id) + '#results')

    qs = conn.execute('''
        SELECT q.*, s.title as section_title
        FROM questions q LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id = ? ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()
    # Reorder to match the student's randomized order if the exam was shuffled
    if exam['randomize_questions'] and exam_sess['question_order']:
        try:
            saved_order = json.loads(exam_sess['question_order'])
            qs_map = {q['id']: q for q in qs}
            qs = [qs_map[qid] for qid in saved_order if qid in qs_map]
        except Exception:
            pass  # Fall back to default order on error

    questions = []
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
        is_case_sensitive = bool(_q_get(q, 'case_sensitive', 0))
        is_correct = False
        if q['question_type'] == 'multiple_choice':
            is_correct = question_credit(q, qd['student_answer']) >= 1.0
        elif q['question_type'] == 'fill_blank':
            correct_blanks, total_blanks = fib_grade(qd['student_answer'], q['correct_answer'], is_case_sensitive)
            is_correct = total_blanks > 0 and correct_blanks == total_blanks
            qd['fib_correct_blanks'] = correct_blanks
            qd['fib_total_blanks'] = total_blanks
            qd['fib_points_earned'] = round(q['points'] * (correct_blanks / total_blanks), 2) if total_blanks else 0
            # Per-blank breakdown for the review UI
            blanks = fib_parse_answer(q['correct_answer'])
            student_parts = (qd['student_answer'] or '').split('|')
            blank_results = []
            for bi, alts in enumerate(blanks):
                given = student_parts[bi].strip() if bi < len(student_parts) else ''
                if is_case_sensitive:
                    b_correct = any(given == alt for alt in alts) if given else False
                else:
                    b_correct = any(given.lower() == alt.lower() for alt in alts) if given else False
                blank_results.append({
                    'index': bi + 1,
                    'student': given,
                    'accepted': ' / '.join(alts),
                    'is_correct': b_correct,
                })
            qd['fib_blanks'] = blank_results
            qd['text_parts'] = fib_split_text(q['question_text'])
        elif q['question_type'] == 'essay':
            criteria = conn.execute(
                'SELECT * FROM rubric_criteria WHERE question_id=? ORDER BY order_index', (q['id'],)
            ).fetchall()
            qd['rubric_criteria'] = [dict(c) for c in criteria]
            qd['answer_id'] = ans['id'] if ans else None
            qd['is_graded'] = bool(ans and ans['manual_score'] is not None)
            qd['manual_score'] = ans['manual_score'] if ans else None
            qd['feedback'] = ans['feedback'] if ans else ''
            scores_map = {}
            if ans:
                rows = conn.execute('SELECT * FROM rubric_scores WHERE answer_id=?', (ans['id'],)).fetchall()
                scores_map = {r['criterion_id']: r['score'] for r in rows}
            qd['rubric_scores_map'] = scores_map
            is_correct = None  # essay has no simple correct/incorrect — shown as graded/pending instead
        else:
            is_correct = question_credit(q, qd['student_answer']) >= 1.0
        qd['is_correct'] = is_correct
        questions.append(qd)

    return render_template('teacher/student_exam_result.html', exam=exam, student=student,
                           exam_session=exam_sess, questions=questions)

def apply_essay_grade(conn, answer_id, criteria, criterion_scores, feedback, teacher_id):
    """Saves rubric scores + feedback for ONE essay answer.

    `criteria` is the list of rubric_criteria rows for the question.
    `criterion_scores` is a dict {criterion_id: raw_score} (raw_score may be
    a string, float, int, or missing/invalid -> treated as 0).
    Returns the clamped total score that was saved.
    """
    total = 0.0
    for crit in criteria:
        raw = criterion_scores.get(crit['id'], 0)
        try:
            given = float(raw)
        except (TypeError, ValueError):
            given = 0.0
        given = max(0.0, min(given, crit['max_points']))  # clamp to the criterion's range
        total += given
        conn.execute('''
            INSERT INTO rubric_scores (answer_id, criterion_id, score) VALUES (?,?,?)
            ON CONFLICT(answer_id, criterion_id) DO UPDATE SET score=excluded.score
        ''', (answer_id, crit['id'], given))

    conn.execute('''
        UPDATE answers SET manual_score=?, feedback=?, graded_by=?, graded_at=datetime('now','localtime')
        WHERE id=?
    ''', (total, feedback, teacher_id, answer_id))
    return total


def recompute_session_score(conn, exam_id, session_id):
    """Recomputes one exam session's total score from scratch across every
    question in the exam, and updates whether it still has essay answers
    waiting to be graded (needs_grading). Returns (new_score, still_pending)."""
    all_qs = conn.execute('SELECT * FROM questions WHERE exam_id=?', (exam_id,)).fetchall()
    new_score = 0.0
    still_pending = False
    for aq in all_qs:
        aq_ans = conn.execute(
            'SELECT * FROM answers WHERE session_id=? AND question_id=?', (session_id, aq['id'])
        ).fetchone()
        a_text = aq_ans['answer_text'] if aq_ans else ''
        a_manual = aq_ans['manual_score'] if aq_ans else None
        new_score += aq['points'] * question_credit(aq, a_text, a_manual)
        if aq['question_type'] == 'essay' and a_manual is None:
            still_pending = True

    conn.execute('UPDATE exam_sessions SET score=?, needs_grading=? WHERE id=?',
                 (new_score, 1 if still_pending else 0, session_id))
    return new_score, still_pending


@app.route('/teacher/exam/<int:exam_id>/student/<int:student_id>/grade-essay/<int:question_id>', methods=['POST'])
@role_required('teacher')
def teacher_grade_essay(exam_id, student_id, question_id):
    """Saves a teacher's rubric scores + feedback for one student's essay
    answer, then recomputes that student's overall exam score and whether
    the submission still has other ungraded essay questions pending."""
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam['teacher_id']) != int(session['user_id']):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))

    q = conn.execute('SELECT * FROM questions WHERE id=? AND exam_id=?', (question_id, exam_id)).fetchone()
    if not q or q['question_type'] != 'essay':
        flash('Question not found.', 'error')
        return redirect(url_for('teacher_view_student_result', exam_id=exam_id, student_id=student_id))

    exam_sess = conn.execute(
        'SELECT * FROM exam_sessions WHERE exam_id=? AND student_id=?', (exam_id, student_id)
    ).fetchone()
    if not exam_sess:
        flash('Submission not found.', 'error')
        return redirect(url_for('teacher_exam_monitoring', exam_id=exam_id) + '#results')

    ans = conn.execute(
        'SELECT * FROM answers WHERE session_id=? AND question_id=?', (exam_sess['id'], question_id)
    ).fetchone()
    if not ans:
        flash('This student did not answer that question.', 'error')
        return redirect(url_for('teacher_view_student_result', exam_id=exam_id, student_id=student_id))

    criteria = conn.execute('SELECT * FROM rubric_criteria WHERE question_id=?', (question_id,)).fetchall()
    feedback = request.form.get('feedback', '').strip()
    criterion_scores = {crit['id']: request.form.get(f'rubric_score_{crit["id"]}', '0').strip() for crit in criteria}

    apply_essay_grade(conn, ans['id'], criteria, criterion_scores, feedback, session['user_id'])
    conn.commit()

    # Recompute this student's overall score now that one more essay is graded,
    # and check whether any OTHER essay answers in this submission are still
    # waiting to be graded.
    recompute_session_score(conn, exam_id, exam_sess['id'])
    conn.commit()

    flash('Grade saved.', 'success')
    return redirect(url_for('teacher_view_student_result', exam_id=exam_id, student_id=student_id) + f'#q-{question_id}')


def _safe_sheet_title(base, used_titles):
    """Excel sheet titles must be <=31 chars and unique within the workbook."""
    base = re.sub(r'[\\/*\[\]:?]', ' ', base).strip() or 'Question'
    title = base[:31]
    n = 2
    while title in used_titles:
        suffix = f' ({n})'
        title = base[:31 - len(suffix)] + suffix
        n += 1
    used_titles.add(title)
    return title


@app.route('/teacher/exam/<int:exam_id>/essays/export')
@role_required('teacher')
def teacher_export_essays(exam_id):
    """Exports every essay question's student answers to an .xlsx workbook —
    one sheet per essay question — with blank/editable rubric-score columns
    and a Feedback column, so a teacher can grade fully offline (no need to
    stay connected to the Raspberry Pi) and re-upload it afterward."""
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id, c.subject_name, c.block_name FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam['teacher_id']) != int(session['user_id']):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))

    essay_questions = conn.execute('''
        SELECT q.*, s.title as section_title FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id=? AND q.question_type='essay'
        ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()

    if not essay_questions:
        flash('This exam has no essay questions to export.', 'error')
        return redirect(url_for('teacher_exam_results', exam_id=exam_id))

    wb = Workbook()
    wb.remove(wb.active)
    used_titles = set()

    info_ws = wb.create_sheet('Instructions')
    used_titles.add('Instructions')
    info_lines = [
        ['SPARK — Essay Grading Export'],
        [f'Exam: {exam["title"]}'],
        [f'Class: {exam["subject_name"]} · {exam["block_name"]}'],
        [''],
        ['How to use this file:'],
        ['1. Each essay question has its own sheet (tab at the bottom).'],
        ['2. Fill in a score for each rubric criterion column (do not exceed the "max" shown in the header).'],
        ['3. You may also fill in the Feedback column — this will show to the student.'],
        ['4. Do NOT edit or delete the "AnswerID" column — it is used to match each row back to the right student.'],
        ['5. Do NOT rename the sheet tabs or add/remove columns.'],
        ['6. Save the file, then upload it back on the "Import Graded Essays" button on the exam results page.'],
        ['7. Rows left blank for a criterion will be scored as 0 for that criterion.'],
    ]
    for line in info_lines:
        info_ws.append(line)
    info_ws.column_dimensions['A'].width = 90
    info_ws['A1'].font = Font(bold=True, size=14)

    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill('solid', fgColor='4472C4')

    for q in essay_questions:
        criteria = conn.execute(
            'SELECT * FROM rubric_criteria WHERE question_id=? ORDER BY order_index', (q['id'],)
        ).fetchall()

        label = q['question_text'][:25].strip() or f'Q{q["id"]}'
        title = _safe_sheet_title(f'Q{q["order_index"] + 1 if q["order_index"] is not None else q["id"]} {label}', used_titles)
        ws = wb.create_sheet(title)

        headers = ['AnswerID (do not edit)', 'Student Name', 'Student Answer']
        for crit in criteria:
            headers.append(f'Score [C{crit["id"]}]: {crit["criterion_text"]} (max {crit["max_points"]:g})')
        headers += ['Feedback', 'Status']
        ws.append(headers)
        for col_idx, _ in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(wrap_text=True, vertical='center')
        ws.freeze_panes = 'D2'
        ws.column_dimensions['A'].hidden = True
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 22
        ws.column_dimensions['C'].width = 60
        for i, crit in enumerate(criteria):
            ws.column_dimensions[get_column_letter(4 + i)].width = 22
        ws.column_dimensions[get_column_letter(4 + len(criteria))].width = 40
        ws.column_dimensions[get_column_letter(5 + len(criteria))].width = 12

        rows = conn.execute('''
            SELECT a.id as answer_id, a.answer_text, a.manual_score, a.feedback
            FROM answers a
            JOIN exam_sessions es ON a.session_id = es.id
            JOIN users u ON es.student_id = u.id
            WHERE a.question_id=? AND es.status='submitted'
            ORDER BY u.full_name
        ''', (q['id'],)).fetchall()
        # Join student name separately (SQLite column name collision safety)
        names = conn.execute('''
            SELECT a.id as answer_id, u.full_name
            FROM answers a JOIN exam_sessions es ON a.session_id = es.id JOIN users u ON es.student_id = u.id
            WHERE a.question_id=? AND es.status='submitted'
        ''', (q['id'],)).fetchall()
        name_by_answer = {r['answer_id']: r['full_name'] for r in names}

        for r in rows:
            existing_scores = {
                sc['criterion_id']: sc['score'] for sc in
                conn.execute('SELECT criterion_id, score FROM rubric_scores WHERE answer_id=?', (r['answer_id'],)).fetchall()
            }
            row_vals = [r['answer_id'], name_by_answer.get(r['answer_id'], ''), r['answer_text'] or '']
            for crit in criteria:
                row_vals.append(existing_scores.get(crit['id'], ''))
            row_vals.append(r['feedback'] or '')
            row_vals.append('Graded' if r['manual_score'] is not None else 'Pending')
            ws.append(row_vals)
            ws.cell(row=ws.max_row, column=3).alignment = Alignment(wrap_text=True, vertical='top')

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_title = re.sub(r'[^A-Za-z0-9_-]+', '_', exam['title'])[:40] or 'exam'
    filename = f'essays_{safe_title}_{exam_id}.xlsx'
    return send_file(buf, as_attachment=True, download_name=filename,
                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/teacher/exam/<int:exam_id>/essays/import', methods=['POST'])
@role_required('teacher')
def teacher_import_essays(exam_id):
    """Accepts a filled-in essay export back, applies every rubric score and
    feedback it contains, then recomputes the affected students' scores."""
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam['teacher_id']) != int(session['user_id']):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))

    file = request.files.get('essay_file')
    if not file or not file.filename:
        flash('Please choose the filled-in Excel file to upload.', 'error')
        return redirect(url_for('teacher_exam_results', exam_id=exam_id))
    if _upload_too_large(file):
        flash(f'File is too large. The maximum allowed size is {MAX_UPLOAD_MB} MB.', 'error')
        return redirect(url_for('teacher_exam_results', exam_id=exam_id))
    if not file.filename.lower().endswith('.xlsx'):
        flash('Please upload the .xlsx file exported from this exam — other formats are not supported.', 'error')
        return redirect(url_for('teacher_exam_results', exam_id=exam_id))

    try:
        wb = load_workbook(file, data_only=True)
    except Exception:
        flash('Could not read that file. Make sure it is the unmodified .xlsx file exported from this exam.', 'error')
        return redirect(url_for('teacher_exam_results', exam_id=exam_id))

    exam_essay_question_ids = {
        row['id'] for row in conn.execute(
            "SELECT id FROM questions WHERE exam_id=? AND question_type='essay'", (exam_id,)
        ).fetchall()
    }

    graded_count = 0
    affected_sessions = set()
    skipped = []

    for ws in wb.worksheets:
        if ws.title.strip().lower() == 'instructions':
            continue
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header_row:
            continue

        col_answer_id = None
        col_feedback = None
        col_criterion = {}
        for idx, h in enumerate(header_row):
            if not h:
                continue
            h = str(h)
            if h.startswith('AnswerID'):
                col_answer_id = idx
            elif h.strip().lower() == 'feedback':
                col_feedback = idx
            else:
                m = re.search(r'\[C(\d+)\]', h)
                if m:
                    col_criterion[idx] = int(m.group(1))

        if col_answer_id is None:
            skipped.append(f'Sheet "{ws.title}": missing the AnswerID column, skipped.')
            continue

        for row in ws.iter_rows(min_row=2, values_only=True):
            if col_answer_id >= len(row) or row[col_answer_id] in (None, ''):
                continue
            try:
                answer_id = int(row[col_answer_id])
            except (TypeError, ValueError):
                skipped.append(f'Sheet "{ws.title}": a row had an invalid AnswerID, skipped.')
                continue

            ans = conn.execute('SELECT * FROM answers WHERE id=?', (answer_id,)).fetchone()
            if not ans or ans['question_id'] not in exam_essay_question_ids:
                skipped.append(f'Sheet "{ws.title}": AnswerID {answer_id} does not belong to this exam, skipped.')
                continue

            criteria = conn.execute(
                'SELECT * FROM rubric_criteria WHERE question_id=?', (ans['question_id'],)
            ).fetchall()
            criterion_scores = {}
            for idx, crit_id in col_criterion.items():
                if idx < len(row) and row[idx] not in (None, ''):
                    criterion_scores[crit_id] = row[idx]
            feedback = ''
            if col_feedback is not None and col_feedback < len(row) and row[col_feedback]:
                feedback = str(row[col_feedback]).strip()

            apply_essay_grade(conn, answer_id, criteria, criterion_scores, feedback, session['user_id'])
            graded_count += 1
            affected_sessions.add(ans['session_id'])

    conn.commit()

    for sid in affected_sessions:
        recompute_session_score(conn, exam_id, sid)
    conn.commit()

    if graded_count:
        flash(f'Imported grades for {graded_count} essay answer(s) across {len(affected_sessions)} student submission(s).', 'success')
    else:
        flash('No gradable rows were found in that file.', 'error')
    for s in skipped[:8]:
        flash(s, 'error')
    if len(skipped) > 8:
        flash(f'...and {len(skipped) - 8} more row(s) skipped.', 'error')

    return redirect(url_for('teacher_exam_results', exam_id=exam_id))


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
        conn.execute("UPDATE exams SET status=?, is_ready=1, manually_closed=0, activated_at=strftime('%Y-%m-%d %H:%M:%S','now','localtime') WHERE id=?", (new_status, exam_id))
    else:
        # Clear activated_at on close so the next open always gets a fresh timer
        conn.execute("UPDATE exams SET status=?, manually_closed=1, activated_at=NULL WHERE id=?", (new_status, exam_id))
    conn.commit()
    flash(f"Exam is now {'Open' if new_status == 'active' else 'Closed'}.", 'success')
    if request.form.get('next') == 'monitoring':
        return redirect(url_for('teacher_exam_monitoring', exam_id=exam_id))
    return redirect(url_for('teacher_class_detail', class_id=class_id))

@app.route('/teacher/exam/<int:exam_id>/finish-editing', methods=['POST'])
@role_required('teacher')
def teacher_finish_editing(exam_id):
    """Teacher is done building the exam: allow the scheduler to open it."""
    conn = get_db()
    exam = conn.execute('''
        SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
    ''', (exam_id,)).fetchone()
    if not exam or int(exam['teacher_id']) != int(session['user_id']):
        flash('Exam not found.', 'error')
        return redirect(url_for('teacher_home'))
    has_q = conn.execute('SELECT 1 FROM questions WHERE exam_id=? LIMIT 1', (exam_id,)).fetchone()
    if not has_q:
        flash('Add at least one question before finishing.', 'error')
        return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    # If the schedule is current or already past, open the exam immediately
    # (timer starts now). A future schedule stays Upcoming and the scheduler opens it.
    schedule_arrived = False
    if exam['scheduled_at']:
        raw = str(exam['scheduled_at']).replace('T', ' ')[:16]
        try:
            schedule_arrived = datetime.strptime(raw, '%Y-%m-%d %H:%M') <= datetime.now()
        except ValueError:
            schedule_arrived = False
    if schedule_arrived and exam['status'] == 'upcoming':
        conn.execute("""UPDATE exams SET status='active', is_ready=1, manually_closed=0,
                        activated_at=strftime('%Y-%m-%d %H:%M:%S','now','localtime') WHERE id=?""", (exam_id,))
        conn.commit()
        flash('Exam is ready and now Open for students.', 'success')
        return redirect(url_for('teacher_class_detail', class_id=exam['class_id']))
    conn.execute('UPDATE exams SET is_ready=1, manually_closed=0 WHERE id=?', (exam_id,))
    conn.commit()
    if exam['scheduled_at']:
        flash('Exam is ready. It will open automatically at its scheduled time.', 'success')
    else:
        flash('Exam is ready. Open it from the class page when you want students to start.', 'success')
    return redirect(url_for('teacher_class_detail', class_id=exam['class_id']))

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
    locked = _exam_locked_response(exam_id)
    if locked:
        return locked
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
        passing_score = request.form.get('passing_score', 50, type=int)
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
                manually_closed=0, is_ready=1, activated_at=strftime('%Y-%m-%d %H:%M:%S','now','localtime') WHERE id=?
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
        conn.execute('UPDATE exams SET time_warning_minutes=? WHERE id=?',
                     (clean_time_warning(request.form.get('time_warning_minutes'),
                                         duration or exam['duration_minutes']), exam_id))
        conn.commit()
        flash('Settings saved.', 'success')
        redirect_to = request.form.get('redirect_to')
        if redirect_to == 'detail':
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
        exam = conn.execute('''
            SELECT e.*, c.teacher_id FROM exams e JOIN classes c ON e.class_id = c.id WHERE e.id=?
        ''', (exam_id,)).fetchone()
    return render_template('teacher/exam_settings.html', exam=exam)

def _exam_locked_response(exam_id):
    """While an exam is open (status 'active') students may be taking it, so
    it can't be edited. Returns a ready-made response if the exam is locked,
    otherwise None. Close the exam first to edit it."""
    if not exam_id:
        return None
    row = get_db().execute('SELECT status, class_id FROM exams WHERE id=?', (exam_id,)).fetchone()
    if not row or row['status'] != 'active':
        return None
    # Only lock once a student has actually started/taken the exam. An open exam
    # nobody has entered yet (e.g. auto-opened by its schedule while the teacher
    # is still building it) is safe to keep editing.
    has_sessions = get_db().execute(
        'SELECT 1 FROM exam_sessions WHERE exam_id=? LIMIT 1', (exam_id,)).fetchone()
    if not has_sessions:
        return None
    msg = 'A student has already started this exam, so it can\'t be edited. Close it first.'
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.is_json:
        return jsonify({'ok': False, 'error': msg}), 409
    flash(msg, 'error')
    return redirect(url_for('teacher_class_detail', class_id=row['class_id']))

def _wants_json():
    return request.headers.get('X-Requested-With') == 'XMLHttpRequest'

@app.route('/teacher/exam/<int:exam_id>/section/add', methods=['POST'])
@role_required('teacher')
def teacher_add_section(exam_id):
    title = request.form.get('title', '').strip()
    description = request.form.get('description', '').strip()
    # section_type is a legacy column kept for schema compatibility only —
    # sections no longer restrict their questions to a single type.
    if not title:
        if _wants_json():
            return jsonify({'ok': False, 'error': 'Section title is required.'}), 400
        return redirect(url_for('teacher_exam_detail', exam_id=exam_id) + '#add-section-card')
    conn = get_db()
    if not conn.execute('SELECT 1 FROM exams e JOIN classes c ON e.class_id=c.id WHERE e.id=? AND c.teacher_id=?',
                        (exam_id, session['user_id'])).fetchone():
        if _wants_json():
            return jsonify({'ok': False, 'error': 'Exam not found.'}), 404
        return redirect(url_for('teacher_home'))
    locked = _exam_locked_response(exam_id)
    if locked:
        return locked
    count = conn.execute('SELECT COUNT(*) FROM sections WHERE exam_id=?', (exam_id,)).fetchone()[0]
    cur = conn.execute('INSERT INTO sections (exam_id, title, description, section_type, order_index) VALUES (?,?,?,?,?)',
                 (exam_id, title, description or None, 'multiple_choice', count))
    conn.commit()
    section_id = cur.lastrowid
    if _wants_json():
        return jsonify({'ok': True, 'section': {
            'id': section_id,
            'title': title,
            'description': description or '',
            'order_index': count,
            'position': count + 1
        }})
    flash('Section added.', 'success')
    return redirect(url_for('teacher_exam_detail', exam_id=exam_id) + '#add-section-card')

@app.route('/teacher/section/<int:section_id>/edit', methods=['POST'])
@role_required('teacher')
def teacher_edit_section(section_id):
    conn = get_db()
    sec = conn.execute('''
        SELECT s.*, e.id as exam_id FROM sections s
        JOIN exams e ON s.exam_id = e.id
        JOIN classes c ON e.class_id = c.id
        WHERE s.id=? AND c.teacher_id=?
    ''', (section_id, session['user_id'])).fetchone()
    if not sec:
        if _wants_json():
            return jsonify({'ok': False, 'error': 'Section not found.'}), 404
        return redirect(url_for('teacher_home'))
    exam_id = sec['exam_id']
    locked = _exam_locked_response(exam_id)
    if locked:
        return locked
    title = request.form.get('title', '').strip()
    description = request.form.get('description', '').strip()
    if not title:
        if _wants_json():
            return jsonify({'ok': False, 'error': 'Section title is required.'}), 400
        flash('Section title is required.', 'error')
        return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    conn.execute('UPDATE sections SET title=?, description=? WHERE id=?',
                 (title, description or None, section_id))
    conn.commit()
    if _wants_json():
        return jsonify({'ok': True, 'section': {
            'id': section_id,
            'title': title,
            'description': description or ''
        }})
    flash('Section updated.', 'success')
    return redirect(url_for('teacher_exam_detail', exam_id=exam_id))

@app.route('/teacher/section/<int:section_id>/delete', methods=['POST'])
@role_required('teacher')
def teacher_delete_section(section_id):
    conn = get_db()
    sec = conn.execute('''
        SELECT s.* FROM sections s
        JOIN exams e ON s.exam_id = e.id
        JOIN classes c ON e.class_id = c.id
        WHERE s.id=? AND c.teacher_id=?
    ''', (section_id, session['user_id'])).fetchone()
    if sec:
        exam_id = sec['exam_id']
        locked = _exam_locked_response(sec['exam_id'])
        if locked:
            return locked
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
    sec = conn.execute('''
        SELECT s.* FROM sections s
        JOIN exams e ON s.exam_id = e.id
        JOIN classes c ON e.class_id = c.id
        WHERE s.id=? AND c.teacher_id=?
    ''', (section_id, session['user_id'])).fetchone()
    if not sec:
        if _wants_json():
            return jsonify({'ok': False, 'error': 'Section not found.'}), 404
        return redirect(url_for('teacher_home'))
    exam_id = sec['exam_id']
    locked = _exam_locked_response(exam_id)
    if locked:
        return locked
    question_ids = request.form.getlist('question_ids')
    imported = 0
    skipped = 0
    for qid in question_ids:
        # Only this teacher's own bank questions may be imported.
        src = conn.execute('SELECT * FROM questions WHERE id=? AND is_bank_only=1 AND teacher_id=?',
                           (qid, session['user_id'])).fetchone()
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
        # Imported copy becomes a regular exam question — it is NOT part of
        # the bank (is_bank_only left at its default 0) and is not linked to
        # any bank group. Only the Question Bank page can create groups.
        cur = conn.execute('''
            INSERT INTO questions (exam_id, section_id, question_text, question_type, points, correct_answer, order_index, case_sensitive)
            VALUES (?,?,?,?,?,?,?,?)
        ''', (exam_id, section_id, src['question_text'], src['question_type'], src['points'], src['correct_answer'], count,
              _q_get(src, 'case_sensitive', 0)))
        new_qid = cur.lastrowid
        if src['question_type'] == 'multiple_choice':
            choices = conn.execute('SELECT * FROM choices WHERE question_id=?', (qid,)).fetchall()
            for c in choices:
                conn.execute('INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                             (new_qid, c['choice_label'], c['choice_text']))
        if src['question_type'] == 'essay':
            criteria = conn.execute(
                'SELECT * FROM rubric_criteria WHERE question_id=? ORDER BY order_index', (qid,)
            ).fetchall()
            for c in criteria:
                conn.execute('''
                    INSERT INTO rubric_criteria (question_id, criterion_text, max_points, order_index)
                    VALUES (?,?,?,?)
                ''', (new_qid, c['criterion_text'], c['max_points'], c['order_index']))
        imported += 1
    conn.commit()
    msg = f'{imported} question(s) imported from bank.'
    if skipped:
        msg += f' {skipped} duplicate(s) skipped.'
    flash(msg, 'success')
    if _wants_json():
        return jsonify({'ok': True, 'imported': imported, 'skipped': skipped, 'section_id': section_id})
    return redirect(url_for('teacher_exam_detail', exam_id=exam_id))

@app.route('/teacher/section/<int:section_id>/question/add', methods=['POST'])
@role_required('teacher')
def teacher_add_question(section_id):
    conn = get_db()
    sec = conn.execute('''
        SELECT s.* FROM sections s
        JOIN exams e ON s.exam_id = e.id
        JOIN classes c ON e.class_id = c.id
        WHERE s.id=? AND c.teacher_id=?
    ''', (section_id, session['user_id'])).fetchone()
    if not sec:
        if _wants_json():
            return jsonify({'ok': False, 'error': 'Section not found.'}), 404
        return redirect(url_for('teacher_home'))
    exam_id = sec['exam_id']
    locked = _exam_locked_response(exam_id)
    if locked:
        return locked
    q_text = request.form.get('question_text', '').strip()
    q_type = request.form.get('question_type', '').strip()
    if q_type not in ('multiple_choice', 'short_answer', 'fill_blank', 'essay', 'true_false'):
        q_type = sec['section_type']
    points = request.form.get('points', 1, type=int)
    if q_type == 'multiple_choice':
        correct = request.form.get('correct_answer_mc', '').strip()
    elif q_type == 'true_false':
        correct = request.form.get('correct_answer_tf', '').strip()  # "True" or "False"
    elif q_type == 'fill_blank':
        correct = request.form.get('correct_answer_fib', '').strip()
    elif q_type == 'essay':
        correct = None  # essay questions have no single "correct answer"
    else:
        correct = sa_collect_form(request.form, 'correct_answer_sa')
    case_sensitive = 1 if request.form.get('case_sensitive') else 0
    if not q_text:
        if _wants_json():
            return jsonify({'ok': False, 'error': 'Question text is required.'}), 400
    mc_choices = None
    if q_type == 'multiple_choice' and q_text:
        mc_choices = mc_collect_choices(request.form)
        ok, correct, err = mc_validate(mc_choices, correct)
        if not ok:
            flash(err, 'error')
            if _wants_json():
                return jsonify({'ok': False, 'error': err}), 400
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    if q_type == 'fill_blank' and q_text:
        blank_count = fib_count_blanks(q_text)
        answer_count = len(correct.split('|')) if correct else 0
        if blank_count == 0 or blank_count != answer_count:
            msg = f'Fill in the Blank questions need one "___" per answer. Found {blank_count} blank(s) but {answer_count} answer group(s).'
            flash(msg, 'error')
            if _wants_json():
                return jsonify({'ok': False, 'error': msg}), 400
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    if q_type == 'true_false' and q_text and correct not in ('True', 'False'):
        msg = 'Please select whether True or False is the correct answer.'
        flash(msg, 'error')
        if _wants_json():
            return jsonify({'ok': False, 'error': msg}), 400
        return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    rubric_criteria = None
    if q_type == 'essay' and q_text:
        rubric_criteria = rubric_collect_criteria(request.form)
        ok, err = rubric_validate(rubric_criteria)
        if not ok:
            flash(err, 'error')
            if _wants_json():
                return jsonify({'ok': False, 'error': err}), 400
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
        points = sum(p for _, p in rubric_criteria)  # question's max points = sum of rubric points
    if q_text:
        count = conn.execute('SELECT COUNT(*) FROM questions WHERE section_id=?', (section_id,)).fetchone()[0]
        # Questions added directly to an exam section are NOT part of the
        # question bank and are never linked to a bank group — only the
        # Question Bank page can create/assign bank groups.
        cur = conn.execute('''
            INSERT INTO questions (exam_id, section_id, question_text, question_type, points, correct_answer, order_index, case_sensitive)
            VALUES (?,?,?,?,?,?,?,?)
        ''', (exam_id, section_id, q_text, q_type, points, correct, count, case_sensitive))
        q_id = cur.lastrowid
        if mc_choices is not None:
            for label, ct in mc_choices:
                conn.execute('INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                             (q_id, label, ct))
        if rubric_criteria is not None:
            for idx, (crit_text, max_pts) in enumerate(rubric_criteria):
                conn.execute('''
                    INSERT INTO rubric_criteria (question_id, criterion_text, max_points, order_index)
                    VALUES (?,?,?,?)
                ''', (q_id, crit_text, max_pts, idx))
        conn.commit()
        flash('Question added.', 'success')
        if _wants_json():
            return jsonify({'ok': True, 'question': {'id': q_id, 'section_id': section_id}})
    redirect_to = request.args.get('redirect_to') or request.form.get('redirect_to', '')
    if redirect_to == 'bank':
        return redirect(url_for('teacher_question_bank'))
    return redirect(url_for('teacher_exam_detail', exam_id=exam_id))

@app.route('/teacher/questions/reorder', methods=['POST'])
@role_required('teacher')
def teacher_reorder_questions():
    """Persist drag-and-drop moves of questions — within a section (reorder)
    or across sections (move). Expects JSON: {"sections": {section_id: [question_id, ...], ...}}
    covering every section currently rendered on the exam page."""
    conn = get_db()
    data = request.get_json(silent=True) or {}
    sections = data.get('sections', {})
    if not isinstance(sections, dict) or not sections:
        return jsonify({'ok': False, 'error': 'No data provided.'}), 400
    try:
        for section_id, qids in sections.items():
            sec = conn.execute('''
                SELECT s.*, e.id as exam_id FROM sections s
                JOIN exams e ON s.exam_id = e.id
                JOIN classes c ON e.class_id = c.id
                WHERE s.id=? AND c.teacher_id=?
            ''', (section_id, session['user_id'])).fetchone()
            if not sec:
                continue  # skip sections that don't belong to this teacher
            locked = _exam_locked_response(sec['exam_id'])
            if locked:
                conn.rollback()
                return locked
            if not isinstance(qids, list):
                continue
            for idx, qid in enumerate(qids):
                conn.execute(
                    'UPDATE questions SET section_id=?, order_index=? WHERE id=? AND exam_id=?',
                    (sec['id'], idx, qid, sec['exam_id'])
                )
        conn.commit()
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400

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
        locked = _exam_locked_response(q['exam_id'])
        if locked:
            return locked
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
    locked = _exam_locked_response(exam_id)
    if locked:
        return locked
    q_text = request.form.get('question_text', '').strip()
    correct = request.form.get('correct_answer', '').strip()
    if q['question_type'] == 'short_answer':
        correct = sa_collect_form(request.form, 'correct_answer')  # primary + alternatives
    points = request.form.get('points', 1, type=int)
    bank_group_id = request.form.get('bank_group_id') or None
    case_sensitive = 1 if request.form.get('case_sensitive') else 0
    if q['question_type'] == 'fill_blank':
        blank_count = fib_count_blanks(q_text)
        answer_count = len(correct.split('|')) if correct else 0
        if blank_count == 0 or blank_count != answer_count:
            flash(f'Fill in the Blank questions need one "___" per answer. Found {blank_count} blank(s) but {answer_count} answer group(s).', 'error')
            redirect_to = request.form.get('redirect_to', '')
            if redirect_to == 'bank' or not exam_id:
                return redirect(url_for('teacher_question_bank'))
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    mc_choices = None
    if q['question_type'] == 'multiple_choice':
        mc_choices = mc_collect_choices(request.form)
        ok, correct, err = mc_validate(mc_choices, correct)
        if not ok:
            flash(err, 'error')
            redirect_to = request.form.get('redirect_to', '')
            if redirect_to == 'bank' or not exam_id:
                return redirect(url_for('teacher_question_bank'))
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    if q['question_type'] == 'true_false' and correct not in ('True', 'False'):
        flash('Please select whether True or False is the correct answer.', 'error')
        redirect_to = request.form.get('redirect_to', '')
        if redirect_to == 'bank' or not exam_id:
            return redirect(url_for('teacher_question_bank'))
        return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
    rubric_criteria = None
    if q['question_type'] == 'essay':
        rubric_criteria = rubric_collect_criteria(request.form)
        ok, err = rubric_validate(rubric_criteria)
        if not ok:
            flash(err, 'error')
            redirect_to = request.form.get('redirect_to', '')
            if redirect_to == 'bank' or not exam_id:
                return redirect(url_for('teacher_question_bank'))
            return redirect(url_for('teacher_exam_detail', exam_id=exam_id))
        correct = None
        points = sum(p for _, p in rubric_criteria)
    conn.execute('UPDATE questions SET question_text=?, correct_answer=?, points=?, bank_group_id=?, case_sensitive=? WHERE id=?',
                 (q_text, correct, points, bank_group_id, case_sensitive, question_id))
    if mc_choices is not None:
        conn.execute('DELETE FROM choices WHERE question_id=?', (question_id,))
        for label, ct in mc_choices:
            conn.execute('INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                         (question_id, label, ct))
    if rubric_criteria is not None:
        conn.execute('DELETE FROM rubric_criteria WHERE question_id=?', (question_id,))
        for idx, (crit_text, max_pts) in enumerate(rubric_criteria):
            conn.execute('''
                INSERT INTO rubric_criteria (question_id, criterion_text, max_points, order_index)
                VALUES (?,?,?,?)
            ''', (question_id, crit_text, max_pts, idx))
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

    delete_questions = request.form.get('delete_questions') == '1'
    if delete_questions:
        # Only remove this teacher's own bank-only questions in the group —
        # copies already imported into a real exam are separate rows created
        # at import time, so this never touches a live/past exam.
        conn.execute(
            'DELETE FROM questions WHERE bank_group_id=? AND is_bank_only=1 AND teacher_id=?',
            (group_id, session['user_id'])
        )
        conn.execute('DELETE FROM question_bank_groups WHERE id=?', (group_id,))
        conn.commit()
        flash(f'Group "{grp["name"]}" and its questions were deleted.', 'success')
    else:
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
    if q_type == 'short_answer':
        correct = sa_join_answers([correct])  # normalise 'A/B' alternatives
    bank_group_id = request.form.get('bank_group_id') or None
    case_sensitive = 1 if request.form.get('case_sensitive') else 0

    if not q_text or q_type not in ('multiple_choice', 'short_answer', 'fill_blank', 'essay', 'true_false'):
        flash('Question text and type are required.', 'error')
        return redirect(url_for('teacher_question_bank'))

    mc_choices = None
    if q_type == 'multiple_choice':
        mc_choices = mc_collect_choices(request.form)
        ok, correct, err = mc_validate(mc_choices, correct)
        if not ok:
            flash(err, 'error')
            return redirect(url_for('teacher_question_bank'))

    if q_type == 'fill_blank':
        blank_count = fib_count_blanks(q_text)
        answer_count = len(correct.split('|')) if correct else 0
        if blank_count == 0 or blank_count != answer_count:
            flash(f'Fill in the Blank questions need one "___" per answer. Found {blank_count} blank(s) but {answer_count} answer group(s).', 'error')
            return redirect(url_for('teacher_question_bank'))

    if q_type == 'true_false' and correct not in ('True', 'False'):
        flash('Please select whether True or False is the correct answer.', 'error')
        return redirect(url_for('teacher_question_bank'))

    rubric_criteria = None
    if q_type == 'essay':
        correct = None  # essay questions have no single "correct answer"
        rubric_criteria = rubric_collect_criteria(request.form)
        ok, err = rubric_validate(rubric_criteria)
        if not ok:
            flash(err, 'error')
            return redirect(url_for('teacher_question_bank'))
        points = sum(p for _, p in rubric_criteria)  # points = sum of rubric criteria

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
            order_index, bank_group_id, is_bank_only, teacher_id, case_sensitive)
           VALUES (NULL, NULL, ?, ?, ?, ?, 0, ?, 1, ?, ?)''',
        (q_text, q_type, points, correct, bank_group_id, session['user_id'], case_sensitive)
    )
    q_id = cur.lastrowid
    if mc_choices is not None:
        for label, ct in mc_choices:
            conn.execute(
                'INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                (q_id, label, ct)
            )
    if rubric_criteria is not None:
        for idx, (crit_text, max_pts) in enumerate(rubric_criteria):
            conn.execute('''
                INSERT INTO rubric_criteria (question_id, criterion_text, max_points, order_index)
                VALUES (?,?,?,?)
            ''', (q_id, crit_text, max_pts, idx))
    conn.commit()
    flash('Question added to bank.', 'success')
    return redirect(url_for('teacher_question_bank'))

@app.route('/teacher/question-bank/import', methods=['POST'])
@role_required('teacher')
def teacher_bank_import_file():
    """Parse an uploaded .txt or .docx file and bulk-add questions to the bank."""
    import re as _re
    conn = get_db()
    uploaded = request.files.get('import_file')
    if not uploaded or uploaded.filename == '':
        flash('No file selected.', 'error')
        return redirect(url_for('teacher_question_bank'))

    if _upload_too_large(uploaded):
        flash(f'File is too large. The maximum allowed size is {MAX_UPLOAD_MB} MB.', 'error')
        return redirect(url_for('teacher_question_bank'))

    raw_name = uploaded.filename
    ext = raw_name.rsplit('.', 1)[-1].lower() if '.' in raw_name else ''

    if ext not in ('txt', 'docx'):
        flash('Unsupported file type. Please upload a .txt or .docx file.', 'error')
        return redirect(url_for('teacher_question_bank'))

    # Derive group name from filename (strip extension)
    group_name = _re.sub(r'\.[^.]+$', '', raw_name).strip()
    # Replace underscores/dashes with spaces for readability
    group_name = _re.sub(r'[_\-]+', ' ', group_name).strip() or 'Imported Questions'

    # ── Extract raw text depending on file type ─────────────────────────────
    try:
        if ext == 'txt':
            content = uploaded.read().decode('utf-8', errors='replace')
        elif ext == 'docx':
            import docx
            document = docx.Document(uploaded)
            content = '\n'.join(p.text for p in document.paragraphs)
    except Exception:
        flash('Could not read file. Make sure it is a valid (.txt or .docx) file.', 'error')
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
    # Multiple choice questions may have up to 26 choices, labeled a-z.
    #
    # Short answer:
    #   1. What is the brain of the computer?
    #   Answer: CPU
    #
    # Fill in the Blank (use "___" to mark each blank; separate multiple
    # blanks in the answer with "|", and alternate acceptable answers for
    # the same blank with "/"):
    #   1. The ___ is the powerhouse of the cell, and water is H2O and ___.
    #   Answer: mitochondria/mitochondrion | oxygen/O2
    #
    # True/False:
    #   1. The sky is blue.
    #   Answer: True
    #
    # Essay (always needs a "Rubric:" section — essay questions are always
    # manually graded against a rubric, never auto-graded, so there is no
    # "Answer:" line; points are the sum of the rubric's max points):
    #   1. Explain the water cycle in your own words.
    #   Rubric:
    #   - Clarity: 5
    #   - Accuracy: 5
    #   - Completeness: 5
    #
    # Any question type may wrap a code sample in a fenced block — use
    # ``` ... ``` or ''' ... ''' (either marker; they don't have to match on
    # both ends). Indentation inside the fence is preserved, and on the
    # student-facing exam/results pages a question with a fenced block
    # (i.e. any question_text containing a newline) is shown in a
    # monospaced code box instead of plain text. Example with a blank
    # inside the code:
    #   3. What will this code output?
    #   '''
    #   x = 10
    #   if x > 5:
    #       print("___")
    #   else:
    #       print("small")
    #   '''
    #   Answer: big
    # ─────────────────────────────────────────────────────────────────────────
    # Only trim trailing newline artifacts here — leading whitespace is kept
    # so indentation inside a fenced ``` or ''' code block survives.
    lines = [l.rstrip('\r\n') for l in content.splitlines()]

    _HEADER_RE = _re.compile(
        r'^(multiple\s*choice|short\s*answer|fill\s*in\s*the\s*blank|true\s*/?\s*false|essay)\s*:?\s*$',
        _re.IGNORECASE
    )
    # Code-block markers are auto-detected: ```  '''  or """ (Word's curly-quote
    # versions of ''' and """ count too). A block closes only with the SAME
    # marker that opened it, so e.g. a """ docstring inside a ``` block is
    # kept as code instead of ending the block early.
    _FENCE_KINDS = (
        ('backticks', _re.compile(r"^```")),
        ('single', _re.compile(r"^(?:'''|[\u2018\u2019]{3})")),
        ('double', _re.compile(r'^(?:"""|[\u201c\u201d]{3})')),
    )

    def _fence_step(line, open_kind):
        """Returns (is_fence_line, new_open_kind)."""
        for kind, rx in _FENCE_KINDS:
            if rx.match(line):
                if not open_kind:
                    return True, kind        # opens a block
                if open_kind == kind:
                    return True, None        # closes it
                return False, open_kind      # other marker inside code: just text
        return False, open_kind

    _RUBRIC_HEADER_RE = _re.compile(r'^rubric\s*:?\s*$', _re.IGNORECASE)
    _RUBRIC_ITEM_RE = _re.compile(r'^-?\s*(.+?)\s*:\s*(\d+(?:\.\d+)?)\s*(?:pts?|points?)?\s*$', _re.IGNORECASE)

    # Drop section header lines (e.g. "Multiple Choice:", "Short Answer:",
    # "Fill in the Blank:") so they don't get mistaken for a stray question
    # with no answer. Skipped while inside a ``` fence so a code line that
    # happens to look like a header (rare, but possible) isn't eaten.
    filtered = []
    in_fence = False
    for l in lines:
        s = l.strip()
        _is_fence, in_fence = _fence_step(s, in_fence)
        if _is_fence:
            filtered.append(l)
            continue
        if in_fence or not _HEADER_RE.match(s):
            filtered.append(l)
    lines = filtered

    # ── Clean parser: split file into question blocks first, then parse each ──
    # A new block starts whenever we see a numbered line: "1.", "2.", "3)" etc,
    # except while inside a ``` fence (so a stray "1)" inside a code sample
    # doesn't get mistaken for the start of the next question).
    blocks = []
    current_block = []
    in_fence = False
    for line in lines:
        stripped = line.strip()
        _is_fence, in_fence = _fence_step(stripped, in_fence)
        if _is_fence:
            current_block.append(line)
            continue
        if in_fence:
            current_block.append(line)
            continue
        if _re.match(r'^\d+[\.\)]\s+', stripped) and current_block:
            blocks.append(current_block)
            current_block = [line]
        elif stripped or current_block:
            current_block.append(line)
    if current_block:
        blocks.append(current_block)

    parsed = []
    skipped_no_answer = 0
    skipped_blank_mismatch = 0
    skipped_essay_no_rubric = 0
    for block in blocks:
        if not block:
            continue
        # First line is the question (strip leading number)
        q_line = block[0]
        q_first = _re.sub(r'^\d+[\.\)]\s+', '', q_line.strip()).strip()
        if not q_first:
            continue

        choices = {}
        answer_raw = ''
        rubric_items = []  # [(criterion_text, max_points), ...] — essay only
        in_rubric = False
        # Everything up to the first choice line becomes part of the question
        # text (joined with real newlines), so a ``` fenced code block right
        # after the question prompt is kept intact — indentation and all.
        question_lines = [q_first]
        in_fence = False
        for bline in block[1:]:
            stripped_b = bline.strip()
            _is_fence, in_fence = _fence_step(stripped_b, in_fence)
            if _is_fence:
                continue  # don't keep the ``` markers themselves
            if in_fence:
                question_lines.append(bline)  # preserve original indentation
                continue
            if _RUBRIC_HEADER_RE.match(stripped_b):
                in_rubric = True
                continue  # the "Rubric:" header line itself isn't kept
            if in_rubric:
                # Once a "Rubric:" header is seen, every remaining line in
                # this block is treated as a rubric item (or ignored if it
                # doesn't match "- Criterion: points") — a rubric section is
                # expected to be the last thing in a question block.
                r_match = _RUBRIC_ITEM_RE.match(stripped_b)
                if r_match:
                    crit_text = r_match.group(1).strip()
                    try:
                        max_pts = float(r_match.group(2))
                    except ValueError:
                        continue
                    if crit_text and max_pts > 0:
                        rubric_items.append((crit_text, max_pts))
                continue
            c_match = _re.match(r'^([a-zA-Z])[\.\)]\s+(.+)', stripped_b)
            a_match = _re.match(r'^[Aa]nswer\s*:\s*(.+)', stripped_b)
            if c_match:
                choices[c_match.group(1).upper()] = c_match.group(2).strip()
            elif a_match:
                answer_raw = a_match.group(1).strip()
            elif not choices:
                # Still part of the question body (plain extra line, no fence)
                question_lines.append(stripped_b)

        q_text = '\n'.join(question_lines).strip('\n')
        if not q_text:
            continue

        if in_rubric and not rubric_items:
            # Had a "Rubric:" header but no valid "- Criterion: points" lines
            # under it. Essay questions must always have a rubric, so skip
            # rather than import an ungradeable essay.
            skipped_essay_no_rubric += 1
            continue

        if rubric_items:
            # A "Rubric:" section was found — this is always an Essay
            # question (mandatory rubric, manually graded, no single
            # correct answer), regardless of anything in choices/answer_raw.
            parsed.append({
                'type': 'essay',
                'text': q_text,
                'rubric': rubric_items,
                'points': sum(pts for _, pts in rubric_items),
            })
        elif choices:
            # A multiple-choice question with no "Answer:" line has no way to
            # know the correct choice. Previously this silently defaulted to
            # "A", which could ship a wrong answer key with no indication
            # anything was wrong. Skip it instead and tell the teacher how
            # many were skipped so they can fix the source file.
            if not answer_raw:
                skipped_no_answer += 1
                continue
            ans_label = answer_raw.upper().strip('.')[:1]
            if ans_label not in choices:
                # Answer line doesn't match any of this question's choice
                # labels (e.g. "Answer: e" but only a-d exist) — also not
                # safe to guess, so skip rather than silently mis-key it.
                skipped_no_answer += 1
                continue
            parsed.append({
                'type': 'multiple_choice',
                'text': q_text,
                'choices': choices,
                'answer': ans_label,
            })
        elif '___' in q_text:
            # Fill in the Blank: skip if the blank count doesn't match the
            # number of '|'-separated answer groups, rather than importing
            # a broken question.
            blank_count = q_text.count('___')
            answer_count = len([p for p in answer_raw.split('|')]) if answer_raw else 0
            if blank_count == 0 or blank_count != answer_count:
                skipped_blank_mismatch += 1
                continue
            parsed.append({
                'type': 'fill_blank',
                'text': q_text,
                'answer': answer_raw,
                'points': blank_count,
            })
        elif answer_raw.strip().lower() in ('true', 'false'):
            # True/False: inferred when the answer is literally "True" or
            # "False" and the question has no choices/blanks/rubric. (If a
            # short-answer question's expected text genuinely is the word
            # "true"/"false", it will import as True/False instead — this is
            # intentional since the two behave identically for the student
            # either way, and True/False is almost always what's meant.)
            parsed.append({
                'type': 'true_false',
                'text': q_text,
                'answer': 'True' if answer_raw.strip().lower() == 'true' else 'False',
            })
        else:
            parsed.append({
                'type': 'short_answer',
                'text': q_text,
                'answer': sa_join_answers([answer_raw]),  # 'A/B' = alternative answers
            })

    if not parsed:
        total_skipped = skipped_no_answer + skipped_blank_mismatch + skipped_essay_no_rubric
        if total_skipped:
            flash(
                f'No questions could be imported: {total_skipped} '
                f'question(s) were skipped because they were missing a valid "Answer:" line, '
                f'had a blank/answer count mismatch, or an essay "Rubric:" section with no valid '
                f'criteria. Please check the file and try again.',
                'error'
            )
        else:
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

    # ── Insert questions (skip ones already in this group — text+type match,
    # same rule the "import from bank" feature already uses — so re-uploading
    # the same file, or a file with overlapping questions, doesn't pile up
    # duplicates in the bank every time) ─────────────────────────────────────
    added = 0
    duplicates = 0
    for q in parsed:
        dup = conn.execute(
            'SELECT 1 FROM questions WHERE bank_group_id=? AND question_type=? AND question_text=?',
            (group_id, q['type'], q['text'])
        ).fetchone()
        if dup:
            duplicates += 1
            continue
        # Fill-in-the-blank points scale with how many blanks the question has
        # (one point per blank); every other type defaults to 1 point.
        points = q.get('points', 1)
        cur = conn.execute(
            '''INSERT INTO questions
               (exam_id, section_id, question_text, question_type, points, correct_answer,
                order_index, bank_group_id, is_bank_only, teacher_id)
               VALUES (NULL, NULL, ?, ?, ?, ?, 0, ?, 1, ?)''',
            (q['text'], q['type'], points, q.get('answer'), group_id, session['user_id'])
        )
        q_id = cur.lastrowid
        if q['type'] == 'multiple_choice':
            label_map = {'A': 0, 'B': 1, 'C': 2, 'D': 3}
            for lbl, txt in q.get('choices', {}).items():
                conn.execute(
                    'INSERT INTO choices (question_id, choice_label, choice_text) VALUES (?,?,?)',
                    (q_id, lbl, txt)
                )
        elif q['type'] == 'essay':
            for idx, (crit_text, max_pts) in enumerate(q.get('rubric', [])):
                conn.execute('''
                    INSERT INTO rubric_criteria (question_id, criterion_text, max_points, order_index)
                    VALUES (?,?,?,?)
                ''', (q_id, crit_text, max_pts, idx))
        added += 1

    conn.commit()
    msg = f'Imported {added} question{"s" if added != 1 else ""} into group "{group_name}".'
    if duplicates:
        msg += f' Skipped {duplicates} already in this group (same question text).'
    skipped_total = skipped_no_answer + skipped_blank_mismatch + skipped_essay_no_rubric
    if skipped_total:
        parts = []
        if skipped_no_answer or skipped_blank_mismatch:
            parts.append('missing a valid "Answer:" line or had a blank/answer mismatch')
        if skipped_essay_no_rubric:
            parts.append('had a "Rubric:" section with no valid criteria')
        msg += (
            f' Skipped {skipped_total} question{"s" if skipped_total != 1 else ""} '
            f'that {" / ".join(parts)} — please review the source file.'
        )
        flash(msg, 'info')
    else:
        flash(msg, 'success')
    return redirect(url_for('teacher_question_bank'))


@app.route('/teacher/question-bank')
@role_required('teacher')
def teacher_question_bank():
    conn = get_db()
    # Fetch only questions explicitly added to the bank (is_bank_only=1).
    # Regular exam-section questions are NOT part of the bank and must not
    # appear here — the bank only contains what the teacher added via
    # "Add Question" / "Import File" on this page.
    raw = conn.execute('''
        SELECT q.id, q.question_text, q.question_type, q.points, q.correct_answer,
               q.bank_group_id, q.is_bank_only, COALESCE(q.case_sensitive, 0) AS case_sensitive,
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
        WHERE q.is_bank_only = 1 AND q.teacher_id = ?
        ORDER BY q.bank_group_id IS NULL ASC, COALESCE(g.name,''), COALESCE(c.subject_name,''), e.title
    ''', (session['user_id'],)).fetchall()
    questions = []
    for q in raw:
        qd = dict(q)
        if q['question_type'] == 'multiple_choice':
            qd['choices'] = [dict(c) for c in conn.execute('SELECT * FROM choices WHERE question_id=?', (q['id'],)).fetchall()]
        else:
            qd['choices'] = []
        if q['question_type'] == 'essay':
            qd['rubric_criteria'] = [dict(c) for c in conn.execute(
                'SELECT * FROM rubric_criteria WHERE question_id=? ORDER BY order_index', (q['id'],)
            ).fetchall()]
        else:
            qd['rubric_criteria'] = []
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

@app.route('/teacher/profile', methods=['GET', 'POST'])
@role_required('teacher')
def teacher_profile():
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
    form, errors, profile_errors = {}, {}, {}
    if request.method == 'POST':
        if request.form.get('action') == 'update_profile':
            ok, profile_errors = handle_profile_update(conn, user)
            if ok:
                return redirect(url_for('teacher_profile'))
            user = dict(user)
            user['full_name'] = request.form.get('full_name', '')
        else:
            form, errors, ok = handle_password_change(conn, user)
            if ok:
                return redirect(url_for('teacher_profile'))
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
    # Short answers can have alternative answers / case-sensitivity, which plain
    # SQL can't evaluate, so count those with the shared grading helper instead.
    sa_correct = {}
    _sa_by_exam = {}
    for _r in question_stats:
        if _r['question_type'] == 'short_answer':
            _sa_by_exam.setdefault(_r['exam_id'], []).append(_r['q_id'])
    for _eid, _qids in _sa_by_exam.items():
        _qs = conn.execute(
            'SELECT id, question_type, correct_answer, case_sensitive, points FROM questions WHERE id IN (%s)'
            % ','.join('?' * len(_qids)), _qids).fetchall()
        sa_correct.update(count_fully_correct(conn, _eid, _qs))
    exam_questions = {}
    for row in question_stats:
        eid = row['exam_id']
        if eid not in exam_questions:
            exam_questions[eid] = {'title': row['exam_title'], 'questions': []}
        exam_questions[eid]['questions'].append({
            'text': row['question_text'][:60],
            'total': row['total_answered'] or 0,
            'correct': (sa_correct.get(row['q_id'], 0) if row['question_type'] == 'short_answer'
                        else (row['correct_count'] or 0)),
        })
    return render_template('teacher/profile.html', user=user, exam_history=exam_history, exam_questions=exam_questions,
                           form=form, errors=errors, profile_errors=profile_errors)

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
    users = conn.execute('''
        SELECT u.*, creator.full_name as created_by_name
        FROM users u
        LEFT JOIN users creator ON u.created_by = creator.id
        ORDER BY u.role, u.full_name
    ''').fetchall()
    admin_count = conn.execute("SELECT COUNT(*) FROM users WHERE role='admin'").fetchone()[0]
    return render_template('admin/users.html', users=users, admin_count=admin_count)

@app.route('/admin/users/create', methods=['GET', 'POST'])
@role_required('admin')
def admin_create_user():
    conn = get_db()
    programs = conn.execute('SELECT * FROM programs ORDER BY code').fetchall()
    form, errors = {}, {}
    if request.method == 'POST':
        full_name  = request.form.get('full_name', '').strip()
        email      = request.form.get('email', '').strip()
        password   = request.form.get('password', '')
        role       = request.form.get('role', '')
        program    = request.form.get('program', '')
        year_level = request.form.get('year_level', '')
        # Keep what was typed so a failed attempt doesn't wipe the form
        # (the password itself is not sent back).
        form = {'full_name': full_name, 'email': email, 'role': role,
                'program': program, 'year_level': year_level}
        if not all([full_name, email, password, role]):
            flash('Please fill in all required fields.', 'error')
        else:
            # Collect every problem at once; the page turns exactly those fields red
            if not _is_allowed_signup_email(email):
                errors['email'] = True
            if role == 'student':
                if not program:
                    errors['program'] = True
                if not year_level:
                    errors['year_level'] = True
            if not errors:
                pw_error = validate_password_strength(password)
                if pw_error:
                    flash(pw_error, 'error')
                    return render_template('admin/create_user.html', programs=programs, form=form, errors=errors)
                try:
                    conn = get_db()
                    conn.execute('''
                        INSERT INTO users (full_name, email, password, role, program, year_level, created_by)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (full_name, email, hash_password(password), role, program, year_level, session['user_id']))
                    conn.commit()
                    flash(f'Account created for {full_name}!', 'success')
                    return redirect(url_for('admin_users'))
                except sqlite3.IntegrityError:
                    flash('Email already exists.', 'error')
    return render_template('admin/create_user.html', programs=programs, form=form, errors=errors)

@app.route('/admin/users/delete/<int:user_id>', methods=['POST'])
@role_required('admin')
def admin_delete_user(user_id):
    conn = get_db()
    target = conn.execute('SELECT * FROM users WHERE id=?', (user_id,)).fetchone()
    if not target:
        flash('User not found.', 'error')
        return redirect(url_for('admin_users'))

    # Safety rule 1: an admin can never delete their own account (avoids lockout)
    if user_id == session['user_id']:
        flash('You cannot delete your own account.', 'error')
        return redirect(url_for('admin_users'))

    # Safety rule 2: never allow deleting the last remaining admin in the system
    if target['role'] == 'admin':
        admin_count = conn.execute("SELECT COUNT(*) FROM users WHERE role='admin'").fetchone()[0]
        if admin_count <= 1:
            flash('Cannot delete the last remaining admin account.', 'error')
            return redirect(url_for('admin_users'))

    conn.execute('DELETE FROM users WHERE id = ?', (user_id,))
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
        WHERE sl.event_type IN ('tab_switch', 'lost_focus', 'window_minimize',
                                'window_blur', 'fullscreen_exit')
        ORDER BY sl.logged_at DESC LIMIT 100
    ''').fetchall()
    pw_resets = conn.execute('''
        SELECT pr.reset_at, s.full_name AS student_name, s.email AS student_email,
               t.full_name AS teacher_name, c.subject_name, c.block_name
        FROM password_reset_logs pr
        LEFT JOIN users s ON pr.student_id = s.id
        LEFT JOIN users t ON pr.teacher_id = t.id
        LEFT JOIN classes c ON pr.class_id = c.id
        ORDER BY pr.reset_at DESC, pr.id DESC LIMIT 100
    ''').fetchall()
    return render_template('admin/logs.html', login_logs=login_logs, suspicious=suspicious,
                           pw_resets=pw_resets)

@app.route('/admin/settings', methods=['GET', 'POST'])
@role_required('admin')
def admin_settings():
    conn = get_db()

    if request.method == 'POST':
        # One form saves both the self-registration toggles and the lockout policy.
        try:
            max_attempts = int(request.form.get('lockout_max_attempts', 5))
        except ValueError:
            max_attempts = 5
        try:
            duration = int(request.form.get('lockout_duration_minutes', 3))
        except ValueError:
            duration = 3
        max_attempts = min(max(max_attempts, 3), 20)
        duration = min(max(duration, 1), 1440)
        set_settings({
            'allow_signup': '1' if request.form.get('allow_signup') == '1' else '0',
            'allow_teacher_signup': '1' if request.form.get('allow_teacher_signup') == '1' else '0',
            'lockout_max_attempts': max_attempts,
            'lockout_duration_minutes': duration,
        })
        flash('Settings saved.', 'success')
        return redirect(url_for('admin_settings'))

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
    return render_template('admin/settings.html', db_size=db_size, record_counts=record_counts,
                            settings=get_settings())

@app.route('/admin/profile', methods=['GET', 'POST'])
@role_required('admin')
def admin_profile():
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
    form, errors, profile_errors = {}, {}, {}

    if request.method == 'POST' and request.form.get('action') == 'update_profile':
        ok, profile_errors = handle_profile_update(conn, user)
        if ok:
            return redirect(url_for('admin_profile'))
        user = dict(user)
        user['full_name'] = request.form.get('full_name', '')
    elif request.method == 'POST':
        current_password = request.form.get('current_password', '')
        new_password = request.form.get('new_password', '')
        confirm_password = request.form.get('confirm_password', '')
        # Keep what was typed so a failed attempt doesn't wipe the form
        form = {'current_password': current_password, 'new_password': new_password,
                'confirm_password': confirm_password}

        if not all([current_password, new_password, confirm_password]):
            flash('Please fill in all password fields.', 'error')
            for key, val in form.items():
                if not val:
                    errors[key] = True
        elif not verify_password(current_password, user['password']):
            flash('Current password is incorrect.', 'error')
            errors['current_password'] = True
        elif new_password != confirm_password:
            flash('New passwords do not match.', 'error')
            errors['confirm_password'] = True
        elif new_password == current_password:
            flash('New password must be different from the current password.', 'error')
            errors['new_password'] = True
        else:
            pw_error = validate_password_strength(new_password)
            if pw_error:
                flash(pw_error, 'error')
                errors['new_password'] = True
            else:
                was_forced = bool(session.get('must_change_pw'))
                conn.execute('UPDATE users SET password=?, must_change_password=0 WHERE id=?',
                             (hash_password(new_password), session['user_id']))
                conn.commit()
                session.pop('must_change_pw', None)
                flash('Password updated successfully.', 'success')
                return redirect(url_for('admin_home' if was_forced else 'admin_profile'))

    return render_template('admin/profile.html', user=user, form=form, errors=errors, profile_errors=profile_errors)



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
        FROM login_logs WHERE logged_at >= DATE('now', 'localtime', '-7 days')
        GROUP BY DATE(logged_at) ORDER BY date DESC
    """).fetchall()
    sus_rows = conn.execute("SELECT event_type, COUNT(*) as cnt FROM suspicious_logs GROUP BY event_type").fetchall()
    # Connectivity events (connected/disconnected) aren't cheating signals —
    # they're just network status — so they're excluded from this summary
    # rather than dumped into "other" alongside real anti-cheat events.
    _CONNECTIVITY_EVENTS = {'connected', 'disconnected'}
    suspicious_summary = {'tab_switch': 0, 'fullscreen_exit': 0, 'window_minimize': 0, 'other': 0}
    for r in sus_rows:
        et = r['event_type']
        if et in _CONNECTIVITY_EVENTS:
            continue
        if et in suspicious_summary:
            suspicious_summary[et] = r['cnt']
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
    original_email = user['email']
    errors = {}
    if request.method == 'POST':
        full_name  = request.form.get('full_name', '').strip()
        email      = request.form.get('email', '').strip()
        role       = request.form.get('role', '')
        program    = request.form.get('program', '')
        year_level = request.form.get('year_level', '')
        email_changed = email.lower() != (user['email'] or '').lower()
        if not all([full_name, email, role]):
            flash('Please fill in all required fields.', 'error')
        else:
            # Only checked when the email is actually being changed, so existing
            # accounts with an older address (e.g. the default admin) can still
            # be edited without being forced onto the school domain.
            if email_changed and not _is_allowed_signup_email(email):
                errors['email'] = True   # the page turns the email field + hint red
            if role == 'student':
                if not program:
                    errors['program'] = True
                if not year_level:
                    errors['year_level'] = True
            if not errors:
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
        # Failed: show the form again with what was typed, not the old saved values
        user = {**dict(user), 'full_name': full_name, 'email': email, 'role': role,
                'program': program, 'year_level': year_level}
    is_self = (user['id'] == session['user_id'])
    locked_until_raw = user['locked_until']
    is_locked = False
    if locked_until_raw:
        try:
            fmt = '%Y-%m-%d %H:%M:%S.%f' if '.' in locked_until_raw else '%Y-%m-%d %H:%M:%S'
            is_locked = datetime.strptime(locked_until_raw, fmt) > datetime.now()
        except ValueError:
            is_locked = False
    return render_template('admin/edit_user.html', user=user, programs=programs, is_self=is_self,
                           original_email=original_email, errors=errors, is_locked=is_locked)

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
    pw_error = validate_password_strength(new_password)
    if pw_error:
        flash(pw_error, 'error')
        return redirect(url_for('admin_edit_user', user_id=user_id))
    conn.execute('UPDATE users SET password=? WHERE id=?', (hash_password(new_password), user_id))
    conn.commit()
    flash(f'Password reset for {user["full_name"]}. New password: {new_password}', 'success')
    return redirect(url_for('admin_edit_user', user_id=user_id))

# ── Unlock Account (clears a lockout from too many failed login attempts) ──
@app.route('/admin/users/unlock/<int:user_id>', methods=['POST'])
@role_required('admin')
def admin_unlock_user(user_id):
    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE id=?', (user_id,)).fetchone()
    if not user:
        flash('User not found.', 'error')
        return redirect(url_for('admin_users'))
    conn.execute('UPDATE users SET failed_login_attempts=0, locked_until=NULL WHERE id=?', (user_id,))
    conn.commit()
    flash(f'{user["full_name"]}\'s account has been unlocked.', 'success')
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
    # Per-question correctness (submitted sessions only), graded with the same
    # rules as scoring so fill-in-the-blank / case-sensitive questions are right.
    q_rows = conn.execute("""
        SELECT q.id, q.question_text, q.question_type, q.correct_answer, q.points,
               COALESCE(q.case_sensitive, 0) AS case_sensitive,
               s.title as section_title, s.order_index as sec_order, q.order_index
        FROM questions q LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id=? ORDER BY s.order_index, q.order_index
    """, (exam_id,)).fetchall()
    correct_counts = count_fully_correct(conn, exam_id, [q for q in q_rows if q['question_type'] != 'essay'])
    answer_totals = {r['question_id']: r['n'] for r in conn.execute("""
        SELECT question_id, COUNT(*) as n FROM answers
        WHERE session_id IN (SELECT id FROM exam_sessions WHERE exam_id=? AND status='submitted')
        GROUP BY question_id
    """, (exam_id,)).fetchall()}
    hard_questions_list = []
    for q in q_rows:
        if q['question_type'] == 'essay':
            continue  # manually graded — no meaningful correct-rate
        total_ans = answer_totals.get(q['id'], 0)
        if total_ans > 0:
            rate = round((correct_counts[q['id']] / total_ans) * 100, 1)
            hard_questions_list.append({
                'question_text': q['question_text'], 'question_type': q['question_type'],
                'correct_count': correct_counts[q['id']], 'total_answers': total_ans,
                'rate': rate,
            })
    hard_questions_list.sort(key=lambda x: x['rate'])
    hard_questions_list = hard_questions_list[:10]
    suspicious = conn.execute("""
        SELECT sl.*, u.full_name FROM suspicious_logs sl
        JOIN users u ON sl.student_id = u.id WHERE sl.exam_id=? ORDER BY sl.logged_at DESC
    """, (exam_id,)).fetchall()

    # Per-Section Analytics: same idea as the teacher-facing results page —
    # average correctness grouped by exam section, for a quick "which part of
    # the exam was hardest overall" view.
    section_order = []
    section_agg = {}
    for r in q_rows:
        if r['question_type'] == 'essay':
            continue  # manually graded — not part of the auto-graded analytics
        title = r['section_title'] or 'Untitled Section'
        if title not in section_agg:
            section_agg[title] = {'section_title': title, 'question_count': 0, 'pct_sum': 0}
            section_order.append(title)
        _tot = answer_totals.get(r['id'], 0)
        pct = round((correct_counts[r['id']] / _tot) * 100) if _tot else 0
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
    # Only the student who owns the session may log events against it —
    # otherwise anyone logged in could push another student to auto-termination.
    sess = conn.execute('SELECT * FROM exam_sessions WHERE id=? AND student_id=?',
                        (session_id, session['user_id'])).fetchone()
    if sess and sess['status'] == 'ongoing':
        conn.execute('''
            INSERT INTO suspicious_logs (session_id, student_id, exam_id, event_type)
            VALUES (?, ?, ?, ?)
        ''', (session_id, sess['student_id'], sess['exam_id'], event_type))

        # Each event type has its own independent counter for teacher
        # monitoring. Only tab_switch can ever auto-terminate the session —
        # fullscreen_exit and window_minimize (lost focus) are counted for
        # visibility only and never end the exam on their own.
        if event_type == 'fullscreen_exit':
            new_count = sess['fullscreen_exit_count'] + 1
            conn.execute('UPDATE exam_sessions SET fullscreen_exit_count=? WHERE id=?', (new_count, session_id))
            conn.commit()
            return jsonify({'status': 'logged', 'count': new_count, 'terminated': False})

        if event_type == 'window_minimize':
            new_count = sess['lost_focus_count'] + 1
            conn.execute('UPDATE exam_sessions SET lost_focus_count=? WHERE id=?', (new_count, session_id))
            conn.commit()
            return jsonify({'status': 'logged', 'count': new_count, 'terminated': False})

        # tab_switch (default/fallback event type) — the only event that can
        # auto-terminate the exam once it hits the teacher-set limit.
        new_count = sess['tab_switch_count'] + 1
        conn.execute('UPDATE exam_sessions SET tab_switch_count=? WHERE id=?', (new_count, session_id))
        conn.commit()
        exam = conn.execute('SELECT * FROM exams WHERE id=?', (sess['exam_id'],)).fetchone()
        terminated = False
        # Auto-terminate only when tab_switch_enabled is ON and limit is reached
        if exam and exam['tab_switch_enabled'] and exam['tab_switch_limit'] and new_count >= exam['tab_switch_limit']:
            total_points = conn.execute('SELECT COALESCE(SUM(points),0) FROM questions WHERE exam_id=?', (sess['exam_id'],)).fetchone()[0]
            conn.execute("UPDATE exam_sessions SET status='terminated', submitted_at=datetime('now','localtime'), score=0, total_points=? WHERE id=?",
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
        "UPDATE exam_sessions SET consent_given=1, consent_at=datetime('now','localtime') WHERE id=?",
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

    # ── Activity-log filters (applied in SQL so they search the WHOLE log, not
    #    just the most recent rows the page happens to have loaded) ──
    #   ?log_student=<name fragment>  ?log_event=<event_type>  ?log_limit=<1..500>
    log_student = (request.args.get('log_student') or '').strip().lower()[:100]
    log_event = (request.args.get('log_event') or '').strip()[:50]
    try:
        log_limit = int(request.args.get('log_limit', 100))
    except (TypeError, ValueError):
        log_limit = 100
    log_limit = max(1, min(log_limit, 500))
    log_where = 'sl.exam_id=?'
    log_params = [exam_id]
    if log_student:
        like = log_student.replace('\\', '\\\\').replace('%', '\\%').replace('_', '\\_')
        log_where += " AND LOWER(u.full_name) LIKE ? ESCAPE '\\'"
        log_params.append('%' + like + '%')
    if log_event:
        log_where += ' AND sl.event_type=?'
        log_params.append(log_event)

    active_sessions = conn.execute('''
        SELECT es.id, u.full_name, es.tab_switch_count, es.fullscreen_exit_count, es.lost_focus_count, es.status,
               COUNT(DISTINCT CASE WHEN TRIM(COALESCE(a.answer_text,'')) != '' THEN a.question_id END) as answered,
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

    logs_total = conn.execute(
        'SELECT COUNT(*) FROM suspicious_logs sl JOIN users u ON sl.student_id = u.id WHERE ' + log_where,
        log_params).fetchone()[0]
    recent_logs = conn.execute(
        'SELECT sl.*, u.full_name FROM suspicious_logs sl '
        'JOIN users u ON sl.student_id = u.id '
        'WHERE ' + log_where + ' '
        'ORDER BY sl.logged_at DESC, sl.id DESC LIMIT ?',
        log_params + [log_limit]).fetchall()

    # Results data
    passing_score = exam['passing_score'] if exam['passing_score'] is not None else 75
    results_rows = conn.execute('''
        SELECT u.id as student_id, u.full_name, es.score, es.total_points, es.status, es.submitted_at, es.started_at
        FROM exam_sessions es JOIN users u ON es.student_id = u.id
        WHERE es.exam_id=?
        ORDER BY es.started_at ASC, es.id ASC
    ''', (exam_id,)).fetchall()
    results_list = []
    for r in results_rows:
        row = dict(r)
        if row['score'] is not None and row['total_points']:
            row['pct'] = pct_floor(row['score'], row['total_points'])
        else:
            row['pct'] = None
        results_list.append(row)

    submitted_sessions = conn.execute(
        "SELECT id FROM exam_sessions WHERE exam_id=? AND status='submitted'", (exam_id,)
    ).fetchall()
    total_submitted = len(submitted_sessions)
    session_ids = [r['id'] for r in submitted_sessions]

    questions_raw = conn.execute('''
        SELECT q.id, q.question_text, q.question_type, q.correct_answer, q.points,
               COALESCE(q.case_sensitive, 0) AS case_sensitive,
               s.title as section_title, q.order_index, s.order_index as sec_order
        FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id = ?
        ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()

    question_stats = build_question_stats(conn, exam_id, questions_raw, total_submitted)

    return jsonify({
        'sessions': sessions_data,
        'total_q': total_q,
        'tab_limit': exam['tab_switch_limit'],
        'tab_switch_enabled': bool(exam['tab_switch_enabled']),
        'logs': [dict(l) for l in recent_logs],
        'logs_total': logs_total,
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
            pct = pct_floor(row['score'], row['total_points'])
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
               COALESCE(q.case_sensitive, 0) AS case_sensitive,
               s.title as section_title, q.order_index, s.order_index as sec_order
        FROM questions q
        LEFT JOIN sections s ON q.section_id = s.id
        WHERE q.exam_id = ?
        ORDER BY s.order_index, q.order_index
    ''', (exam_id,)).fetchall()

    question_stats = build_question_stats(conn, exam_id, questions_raw, total_submitted)

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
                "UPDATE exam_sessions SET status='terminated', submitted_at=datetime('now','localtime'), score=0, total_points=? WHERE id=?",
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
    answer_text = str(data.get('answer_text') or '').strip()
    if not session_id or question_id is None:
        return jsonify({'status': 'error', 'reason': 'missing fields'})
    conn = get_db()
    sess = conn.execute('SELECT * FROM exam_sessions WHERE id=? AND student_id=?',
                        (session_id, session['user_id'])).fetchone()
    if sess and sess['status'] == 'ongoing':
        # The question must belong to the exam this session is for.
        if not conn.execute('SELECT 1 FROM questions WHERE id=? AND exam_id=?',
                            (question_id, sess['exam_id'])).fetchone():
            return jsonify({'status': 'error', 'reason': 'invalid question'})
        if answer_text == '':
            # Nothing to save — and if a prior (now-cleared) answer exists, remove it
            # so the question no longer counts as "answered" in monitoring/progress.
            conn.execute('DELETE FROM answers WHERE session_id=? AND question_id=?',
                        (session_id, question_id))
            conn.commit()
            return jsonify({'status': 'cleared'})
        conn.execute('''
            INSERT OR REPLACE INTO answers (session_id, question_id, answer_text)
            VALUES (?,?,?)
        ''', (session_id, question_id, answer_text))
        conn.commit()
        return jsonify({'status': 'saved'})
    return jsonify({'status': 'error', 'reason': 'session not found or not ongoing'})

# ─── App startup: DB init + background maintenance thread ───────────────────
# This runs on *import*, not just when the script is executed directly, so it
# fires whether the app is started as `python app.py` (local/Windows testing,
# via waitress/Flask's dev server below) OR loaded by a production WSGI server
# like gunicorn on the Raspberry Pi (`gunicorn app:app`), which imports this
# module as `app:app` and never hits the `__name__ == '__main__'` block.
os.makedirs(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'instance'), exist_ok=True)
with app.app_context():
    init_db()

# Scheduling/auto-close checks run on their own timer instead of on every
# request (see background_maintenance_loop / removed before_request hook).
# Guarded against the Werkzeug reloader (which re-imports this module in its
# child process) so only one maintenance thread ever runs per process.
if not os.environ.get('WERKZEUG_RUN_MAIN') or __name__ == '__main__':
    maintenance_thread = threading.Thread(target=background_maintenance_loop, daemon=True)
    maintenance_thread.start()

if __name__ == '__main__':
    # Local/manual run: `python app.py`. Used for Windows dev/testing, and as
    # a quick fallback on the Pi if gunicorn isn't set up yet. NOT what the
    # Raspberry Pi should run for real exams — see README for the gunicorn
    # (+ optional nginx) production setup, which handles concurrent students
    # far better than either of these.
    try:
        from waitress import serve
        print("Starting local server (waitress) on http://0.0.0.0:5000 ...")
        serve(app, host='0.0.0.0', port=5000, threads=16)
    except ImportError:
        print("WARNING: 'waitress' is not installed (pip install waitress).")
        print("Falling back to Flask's built-in server with threading enabled.")
        app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
