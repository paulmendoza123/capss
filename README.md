# ⚡ SPARK — Exam Monitoring System v2.0

A Flask-based online exam system with role-based access for **Admin**, **Teacher**, and **Student**.

---

## 📁 Project Structure

```
spark/
├── app.py                        # Main Flask app — all routes, DB, API
├── requirements.txt
├── instance/
│   └── spark.db                  # Auto-created SQLite database
├── static/
│   ├── css/style.css             # Full dark UI theme
│   └── js/
│       ├── main.js               # Alerts, tabs, modal helpers
│       └── exam.js               # Timer, tab detection, auto-save
└── templates/
    ├── base.html                 # Sidebar layout with role-based nav
    ├── login.html
    ├── signup.html               # Student self-registration
    ├── student/
    │   ├── classes.html
    │   ├── join_class.html
    │   ├── class_detail.html     # Exams + Results tabs
    │   ├── exams.html
    │   ├── take_exam.html        # Live exam with timer + tab detection
    │   ├── exam_result.html      # Score + answer review
    │   ├── exam_terminated.html
    │   └── profile.html
    ├── teacher/
    │   ├── classes.html
    │   ├── create_class.html
    │   ├── class_detail.html     # Exams + Results + Students tabs
    │   ├── create_exam.html
    │   ├── exam_detail.html      # Questions + Monitoring + Results + Settings tabs
    │   ├── exam_monitoring.html  # Live student monitoring
    │   ├── exam_results.html
    │   ├── exam_settings.html
    │   ├── my_exams.html
    │   ├── question_bank.html
    │   └── profile.html
    └── admin/
        ├── dashboard.html
        ├── users.html
        ├── create_user.html
        ├── programs.html
        ├── exam_overview.html
        ├── logs.html
        ├── settings.html
        └── profile.html
```

---

## 🚀 Setup & Run

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app (DB auto-initializes on first run)
python app.py
```

Open: **http://localhost:5000**

---

## 🔑 Default Accounts

| Role    | Email              | Password   |
|---------|--------------------|------------|
| Admin   | admin@spark.edu    | admin123   |

> Teachers and students are created via Admin → User Management,
> or students can self-register at `/signup`.

---

## ✨ Features

### Student
- Self-registration (no admin approval needed)
- Join class by code
- Take exams with countdown timer
- Tab switch detection → auto-terminate on limit
- View results with per-question answer review (if teacher enables it)

### Teacher
- Create classes with auto-generated codes
- Create exams with sections (Multiple Choice / Short Answer)
- Add, edit, delete questions per section
- Live monitoring: tab switch count, progress, risk level, manual terminate
- View results with pass/fail breakdown
- Question bank across all exams

### Admin
- Dashboard with system stats
- User management (create / delete teachers & students)
- Program management (add / delete programs like BSIT, BSCS)
- System-wide exam overview
- Login history + suspicious activity logs

---

## 🔒 Security Notes

- Passwords hashed with SHA-256 (no salt — upgrade to `bcrypt` for production)
- Secret key should be changed before deploying
- SQLite is fine for development; use PostgreSQL for production
