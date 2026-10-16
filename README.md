# ⚡ SPARK — Exam Monitoring System v2.0

A Flask-based offline exam system with role-based access for **Admin**, **Teacher**, and **Student**.

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

This local run uses waitress (or Flask's dev server as a fallback) — fine for
development and for Windows machines, where gunicorn below doesn't run.

---

## 🖥️ Production Deployment on the Raspberry Pi

For real exam day, run SPARK behind **gunicorn** (production WSGI server) with
**nginx** in front of it as a reverse proxy, instead of `python app.py`. This
is Linux-only, so it's meant for the Pi, not for Windows dev machines.

```bash
# 1. On the Pi, inside the project folder:
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 2. Quick manual test — confirm gunicorn itself runs the app correctly
#    before wiring it into systemd:
gunicorn --worker-class gthread --workers 1 --threads 8 --bind 0.0.0.0:5000 app:app
# Visit http://<pi-ip>:5000 from another device on the network, then Ctrl+C.

# 3. Install it as a systemd service so it auto-starts and restarts on crash:
#    (edit deploy/spark.service first — paths/username must match your Pi)
sudo cp deploy/spark.service /etc/systemd/system/spark.service
sudo systemctl daemon-reload
sudo systemctl enable spark
sudo systemctl start spark
sudo systemctl status spark          # should say "active (running)"

# 4. (Optional but recommended) put nginx in front of it on port 80, so
#    students just hit the Pi's IP with no ":5000":
sudo apt install nginx -y            # skip if already installed
#    (edit deploy/nginx_spark.conf first — static files path must match your Pi)
sudo cp deploy/nginx_spark.conf /etc/nginx/sites-available/spark
sudo ln -s /etc/nginx/sites-available/spark /etc/nginx/sites-enabled/spark
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl reload nginx
```

See `deploy/spark.service` and `deploy/nginx_spark.conf` for the full,
commented config files (why single-worker-multi-thread, why `/static/` is
served by nginx directly, etc.).

**After pulling new code**, restart the service so the change takes effect:
```bash
git pull
sudo systemctl restart spark
```

---

## 🔑 First-Run Admin Account

On the first run the app creates `admin@spark.edu` with the default password `admin123`
(or the value of `SPARK_ADMIN_PASSWORD`, if set). The admin is **forced to change the
password immediately after the first login**; nothing else in the app is accessible until
it is changed. Log in and set the new password **before connecting students to the network.**

Lost the admin password? Run `python reset_admin.py`.

> Teachers and students are created via Admin → User Management,
> or students can self-register at `/signup`.

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

- Session secret key: read from `SPARK_SECRET_KEY`, otherwise auto-generated once and
  stored in `instance/secret_key` (never commit `instance/`)
- Passwords use salted scrypt hashes (werkzeug). Old SHA-256 hashes are upgraded
  automatically the next time that user logs in
- SQLite is fine for a single-server local classroom network
