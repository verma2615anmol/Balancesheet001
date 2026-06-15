"""
BS Annual Updater — Multi-Tool CA Dashboard
Auth + Upload-Based Plans + Admin Panel
"""

import re
import os
import uuid
import json
import hashlib
import secrets
from datetime import datetime, timedelta
from functools import wraps
from flask import (Flask, request, send_file, jsonify,
                   render_template_string, session, redirect, url_for, g)
from processor import process, detect_fixed_asset_sheet_names

# ── Database driver selection ────────────────────────────────────────────────
# If DATABASE_URL is set (Supabase/PostgreSQL), use psycopg2.
# Otherwise, fall back to SQLite for local development.
DATABASE_URL = os.environ.get("DATABASE_URL")
USE_POSTGRES = bool(DATABASE_URL)

if USE_POSTGRES:
    import psycopg2
    import psycopg2.extras
else:
    import sqlite3

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

# ── Global error handlers — ensure /process NEVER returns HTML ────────────
@app.errorhandler(500)
def handle_500(e):
    if request.path == "/process":
        return jsonify({"status": "error", "message": f"Server error: {e}"}), 500
    return "Internal Server Error", 500

@app.errorhandler(Exception)
def handle_exception(e):
    if request.path == "/process":
        return jsonify({"status": "error", "message": f"Server error: {e}"}), 500
    raise e

UPLOAD_DIR = "/tmp/bs_uploads"
OUTPUT_DIR = "/tmp/bs_outputs"
DB_PATH    = os.environ.get("DB_PATH", "users.db")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

FREE_UPLOADS         = 2
UPLOAD_VALIDITY_DAYS = 90

ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "sumit_admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "Admin@Secure123")
CONTACT_EMAIL  = "sumitverma2880@gmail.com"
CONTACT_UPI    = "sumit2615verma@okhdfcbank"

PLANS = {
    "free":     {"label": "Free",         "uploads": 2,   "price": 0},
    "starter":  {"label": "Starter",      "uploads": 10,  "price": 60},
    "standard": {"label": "Standard",     "uploads": 25,  "price": 130},
    "pro":      {"label": "Professional", "uploads": 60,  "price": 270},
    "firm":     {"label": "Firm",         "uploads": 150, "price": 600},
    "ca":       {"label": "CA Admin",     "uploads": 500, "price": 1000},
}

# ══════════════════════════════════════════════════════════════════════════════
#  DATABASE — PostgreSQL (Supabase) with SQLite fallback for local dev
# ══════════════════════════════════════════════════════════════════════════════

def get_db():
    if "db" not in g:
        if USE_POSTGRES:
            g.db = psycopg2.connect(DATABASE_URL)
            g.db.autocommit = False
        else:
            g.db = sqlite3.connect(DB_PATH, detect_types=sqlite3.PARSE_DECLTYPES)
            g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db:
        try: db.close()
        except: pass

def _db_fetchone(sql, params=()):
    db = get_db()
    if USE_POSTGRES:
        cur = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql, params)
        row = cur.fetchone()
        cur.close()
        return row
    else:
        return db.execute(sql, params).fetchone()

def _db_fetchall(sql, params=()):
    db = get_db()
    if USE_POSTGRES:
        cur = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql, params)
        rows = cur.fetchall()
        cur.close()
        return rows
    else:
        return db.execute(sql, params).fetchall()

def _db_execute(sql, params=()):
    db = get_db()
    if USE_POSTGRES:
        cur = db.cursor()
        cur.execute(sql, params)
        cur.close()
        db.commit()
    else:
        db.execute(sql, params)
        db.commit()

def _placeholder(n=1):
    """Return the correct placeholder for the DB driver: %s for PG, ? for SQLite."""
    return "%s" if USE_POSTGRES else "?"

def _ph(sql_with_qmarks):
    """Convert ? placeholders to %s for PostgreSQL."""
    if USE_POSTGRES:
        return sql_with_qmarks.replace("?", "%s")
    return sql_with_qmarks

def init_db():
    if USE_POSTGRES:
        conn = psycopg2.connect(DATABASE_URL)
        cur = conn.cursor()
        cur.execute("""CREATE TABLE IF NOT EXISTS users (
            id            SERIAL PRIMARY KEY,
            username      TEXT    UNIQUE NOT NULL,
            password      TEXT    NOT NULL,
            plan          TEXT    NOT NULL DEFAULT 'free',
            is_admin      INTEGER NOT NULL DEFAULT 0,
            uploads_total INTEGER NOT NULL DEFAULT 2,
            uploads_used  INTEGER NOT NULL DEFAULT 0,
            validity_end  TEXT,
            created_at    TEXT    NOT NULL)""")
        cur.execute("""CREATE TABLE IF NOT EXISTS usage_log (
            id           SERIAL PRIMARY KEY,
            user_id      INTEGER NOT NULL REFERENCES users(id),
            filename     TEXT,
            processed_at TEXT    NOT NULL)""")
        # Insert admin if not exists
        cur.execute("SELECT id FROM users WHERE username=%s", (ADMIN_USERNAME,))
        if not cur.fetchone():
            cur.execute("""INSERT INTO users
                (username,password,plan,is_admin,uploads_total,uploads_used,created_at)
                VALUES (%s,%s,'firm',1,999999,0,%s)""",
                (ADMIN_USERNAME, _hash(ADMIN_PASSWORD), datetime.utcnow().isoformat()))
        conn.commit()
        cur.close()
        conn.close()
    else:
        db = sqlite3.connect(DB_PATH)
        db.execute("""CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            username      TEXT    UNIQUE NOT NULL,
            password      TEXT    NOT NULL,
            plan          TEXT    NOT NULL DEFAULT 'free',
            is_admin      INTEGER NOT NULL DEFAULT 0,
            uploads_total INTEGER NOT NULL DEFAULT 2,
            uploads_used  INTEGER NOT NULL DEFAULT 0,
            validity_end  TEXT,
            created_at    TEXT    NOT NULL)""")
        db.execute("""CREATE TABLE IF NOT EXISTS usage_log (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id      INTEGER NOT NULL,
            filename     TEXT,
            processed_at TEXT    NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(id))""")
        db.execute("""INSERT OR IGNORE INTO users
            (username,password,plan,is_admin,uploads_total,uploads_used,created_at)
            VALUES (?,?,'firm',1,999999,0,?)""",
            (ADMIN_USERNAME, _hash(ADMIN_PASSWORD), datetime.utcnow().isoformat()))
        db.commit()
        db.close()

def _hash(p): return hashlib.sha256(p.encode("utf-8")).hexdigest()
def get_user_by_name(u): return _db_fetchone(_ph("SELECT * FROM users WHERE username=?"), (u,))
def get_user_by_id(i):   return _db_fetchone(_ph("SELECT * FROM users WHERE id=?"), (i,))
def uploads_remaining(user): return max(0, user["uploads_total"] - user["uploads_used"])

def log_usage(user_id, filename):
    _db_execute(_ph("UPDATE users SET uploads_used=uploads_used+1 WHERE id=?"), (user_id,))
    _db_execute(_ph("INSERT INTO usage_log (user_id,filename,processed_at) VALUES (?,?,?)"),
               (user_id, filename, datetime.utcnow().isoformat()))

def add_uploads(user_id, plan_key):
    user    = get_user_by_id(user_id)
    extra   = PLANS[plan_key]["uploads"]
    rem     = uploads_remaining(user)
    new_tot = user["uploads_used"] + rem + extra
    validity = (datetime.utcnow() + timedelta(days=UPLOAD_VALIDITY_DAYS)).isoformat()
    _db_execute(_ph("UPDATE users SET plan=?,uploads_total=?,validity_end=? WHERE id=?"),
               (plan_key, new_tot, validity, user_id))

def create_user(username, password, plan_key):
    uploads  = PLANS[plan_key]["uploads"]
    validity = None if plan_key == "free" else (datetime.utcnow() + timedelta(days=UPLOAD_VALIDITY_DAYS)).isoformat()
    _db_execute(_ph("""INSERT INTO users
        (username,password,plan,is_admin,uploads_total,uploads_used,validity_end,created_at)
        VALUES (?,?,?,0,?,0,?,?)"""),
        (username, _hash(password), plan_key, uploads, validity, datetime.utcnow().isoformat()))

def del_user(uid):
    _db_execute(_ph("DELETE FROM usage_log WHERE user_id=?"), (uid,))
    _db_execute(_ph("DELETE FROM users WHERE id=?"), (uid,))

def all_users(): return _db_fetchall("SELECT * FROM users ORDER BY id")

# ══════════════════════════════════════════════════════════════════════════════
#  AUTH DECORATORS
# ══════════════════════════════════════════════════════════════════════════════

def login_required(f):
    @wraps(f)
    def dec(*a, **kw):
        if "uid" not in session: return redirect(url_for("login_page"))
        return f(*a, **kw)
    return dec

def admin_required(f):
    @wraps(f)
    def dec(*a, **kw):
        if "uid" not in session: return redirect(url_for("login_page"))
        u = get_user_by_id(session["uid"])
        if not u or not u["is_admin"]: return "Access denied.", 403
        return f(*a, **kw)
    return dec

# ══════════════════════════════════════════════════════════════════════════════
#  SHARED CSS
# ══════════════════════════════════════════════════════════════════════════════

BASE_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--brand:#1D4ED8;--brand-d:#1e40af;--accent:#F59E0B;--green:#10B981;--red:#EF4444;
      --ink:#111827;--muted:#6B7280;--border:#E5E7EB;--bg:#F9FAFB;--white:#fff;
      --radius:12px;--shadow:0 4px 24px rgba(0,0,0,.08)}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--ink);min-height:100vh}
nav{background:var(--white);border-bottom:1px solid var(--border);padding:0 24px;
    display:flex;align-items:center;justify-content:space-between;height:60px;
    position:sticky;top:0;z-index:100;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.logo{font-size:20px;font-weight:800;color:var(--brand);letter-spacing:-.5px;text-decoration:none}
.logo span{color:var(--accent)}
.nav-right{display:flex;align-items:center;gap:14px}
.nav-user{font-size:13px;color:var(--muted)}
.nav-user strong{color:var(--ink)}
.nav-btn{background:var(--brand);color:#fff;padding:7px 16px;border-radius:8px;
         font-size:13px;font-weight:600;text-decoration:none;transition:background .2s}
.nav-btn:hover{background:var(--brand-d)}
.nav-link{font-size:13px;color:var(--muted);text-decoration:none;font-weight:500}
.nav-link:hover{color:var(--red)}
.badge{display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;
       border-radius:99px;text-transform:uppercase;letter-spacing:.04em}
.b-free{background:#F3F4F6;color:var(--muted)}
.b-starter{background:#ECFDF5;color:#065F46}
.b-standard{background:#EFF6FF;color:var(--brand)}
.b-pro{background:#FFFBEB;color:#92400E}
.b-firm{background:#F5F3FF;color:#5B21B6}
footer{background:#0f1b2d;color:#9CA3AF;font-size:12px;padding:0}
.ft-main{display:grid;grid-template-columns:2fr 1fr 1.4fr;gap:40px;padding:40px 48px;max-width:1200px;margin:0 auto}
.ft-brand-name{color:#fff;font-size:18px;font-weight:800;margin-bottom:12px}
.ft-brand-desc{font-size:12.5px;line-height:1.75;color:#9CA3AF;max-width:340px;text-align:justify}
.ft-col-title{color:#fff;font-size:14px;font-weight:700;margin-bottom:14px}
.ft-links{list-style:none;padding:0;margin:0}
.ft-links li{margin-bottom:8px}
.ft-links a{color:#9CA3AF;text-decoration:none;font-size:13px;transition:color .2s}
.ft-links a:hover{color:#fff}
.ft-contact-name{color:#fff;font-weight:700;font-size:13px;margin-bottom:6px}
.ft-contact-addr{color:#9CA3AF;font-size:12px;line-height:1.7;margin-bottom:10px}
.ft-contact-line{color:#9CA3AF;font-size:12px;margin-bottom:4px}
.ft-socials{display:flex;gap:14px;margin-top:12px}
.ft-socials a{color:#9CA3AF;transition:color .2s}
.ft-socials a:hover{color:#fff}
.ft-socials svg{width:20px;height:20px;fill:currentColor}
.ft-bottom{background:#0a1422;border-top:1px solid #1e2d42;padding:12px 48px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.ft-bottom-left{font-size:11px;color:#6B7280}
.ft-bottom-right{font-size:11px;color:#6B7280}
@media(max-width:768px){.ft-main{grid-template-columns:1fr;padding:28px 20px;gap:24px}.ft-bottom{padding:12px 20px;flex-direction:column;text-align:center}}
/* WhatsApp floating button */
.wa-float{position:fixed;bottom:24px;left:24px;width:52px;height:52px;background:#25D366;border-radius:50%;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 16px rgba(0,0,0,.18);z-index:999;text-decoration:none;transition:transform .2s,box-shadow .2s}
.wa-float:hover{transform:scale(1.1);box-shadow:0 6px 24px rgba(0,0,0,.25)}
.wa-float svg{width:28px;height:28px;fill:#fff}
/* How-to-use help modal */
.help-btn{position:fixed;bottom:86px;right:20px;width:44px;height:44px;background:var(--brand);color:#fff;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:800;cursor:pointer;box-shadow:0 4px 14px rgba(29,78,216,.35);z-index:998;border:none;transition:transform .2s,box-shadow .2s;text-decoration:none}
.help-btn:hover{transform:scale(1.1);box-shadow:0 6px 20px rgba(29,78,216,.45)}
.help-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:1001;align-items:center;justify-content:center;padding:16px}
.help-overlay.open{display:flex}
.help-modal{background:#fff;border-radius:16px;max-width:540px;width:100%;max-height:82vh;overflow-y:auto;box-shadow:0 24px 60px rgba(0,0,0,.2)}
.help-modal-head{padding:20px 24px 16px;border-bottom:1px solid #E5E7EB;display:flex;justify-content:space-between;align-items:center}
.help-modal-head h3{font-size:16px;font-weight:800;color:#111827}
.help-close{background:none;border:none;font-size:22px;cursor:pointer;color:#6B7280;line-height:1}
.help-modal-body{padding:20px 24px}
.help-step{display:flex;gap:14px;margin-bottom:18px;align-items:flex-start}
.help-step-num{min-width:28px;height:28px;background:var(--brand);color:#fff;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;flex-shrink:0;margin-top:1px}
.help-step-body h4{font-size:13px;font-weight:700;margin-bottom:3px;color:#111827}
.help-step-body p{font-size:12px;color:#6B7280;line-height:1.6;margin:0}
.help-tip{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;padding:10px 14px;font-size:12px;color:#1E40AF;margin-top:4px;line-height:1.6}
/* Animations */
@keyframes fadeUp{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
.anim-up{animation:fadeUp .4s ease-out both}
.anim-in{animation:fadeIn .3s ease-out both}
"""

# ══════════════════════════════════════════════════════════════════════════════
#  LOGIN PAGE
# ══════════════════════════════════════════════════════════════════════════════

PRIVACY_TEMPLATE = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Privacy Policy – CA Toolkit</title>
<style>
""" + BASE_CSS + """
.pp-wrap{max-width:760px;margin:40px auto;padding:0 24px 60px}
.pp-wrap h1{font-size:22px;font-weight:800;margin-bottom:6px}
.pp-wrap h2{font-size:15px;font-weight:700;margin:28px 0 8px;color:var(--brand)}
.pp-wrap p,.pp-wrap li{font-size:13px;line-height:1.8;color:#374151}
.pp-wrap ul{padding-left:20px;margin-bottom:12px}
.pp-date{font-size:11px;color:var(--muted);margin-bottom:24px}
</style></head><body>
<nav class="nav"><a href="/" class="nav-brand">CA Toolkit</a></nav>
<div class="pp-wrap">
  <h1>Privacy Policy</h1>
  <p class="pp-date">Last updated: June 2026</p>
  <h2>1. Data We Collect</h2>
  <p>We collect only the minimum information necessary to operate the platform: your email/username for account creation, and uploaded Excel files solely for processing your request.</p>
  <h2>2. File Handling</h2>
  <ul>
    <li>Uploaded files are processed in memory on our servers.</li>
    <li>Files are automatically deleted within minutes of processing — we do not store them permanently.</li>
    <li>We never read, analyse, or share the contents of your financial files with any third party.</li>
  </ul>
  <h2>3. No Ads · No Tracking</h2>
  <p>CA Toolkit does not serve advertisements and does not use third-party tracking or analytics cookies. We do not sell your data.</p>
  <h2>4. Account Data</h2>
  <p>Your username and plan information are stored securely in our database. We do not store any payment card details — all payments are handled via UPI or direct bank transfer.</p>
  <h2>5. Refund Policy</h2>
  <p style="color:#B91C1C;font-weight:600">No refund is issued once the first upload of a paid plan has been used. Unused credits on free plans are non-transferable.</p>
  <h2>6. Contact</h2>
  <p>For any privacy concerns, contact us on <a href="https://wa.me/918427651580">WhatsApp</a>.</p>
</div>
<footer>
  <div class="ft-bottom" style="justify-content:center">
    <span class="ft-bottom-left">&copy;2026 CA Toolkit &middot; All Rights Reserved &middot; <a href="/privacy" style="color:#6B7280;text-decoration:none">Privacy Policy</a></span>
  </div>
</footer>
</body></html>"""

LOGIN_T = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Login – CA Toolkit</title>
<style>
""" + BASE_CSS + """
body{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px}
.auth-card{background:var(--white);border:1px solid var(--border);border-radius:var(--radius);
           box-shadow:var(--shadow);width:100%;max-width:420px;padding:40px}
.auth-logo{font-size:22px;font-weight:800;color:var(--brand);margin-bottom:4px}
.auth-logo span{color:var(--accent)}
.auth-sub{font-size:13px;color:var(--muted);margin-bottom:28px}
h2{font-size:20px;font-weight:700;margin-bottom:4px}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;
      letter-spacing:.04em;color:var(--muted);margin-bottom:5px}
.field{margin-bottom:18px}
input{width:100%;border:1.5px solid var(--border);border-radius:8px;padding:10px 14px;
      font-family:inherit;font-size:14px;color:var(--ink);background:var(--white);
      outline:none;transition:border-color .2s}
input:focus{border-color:var(--brand)}
.btn{width:100%;background:var(--brand);color:#fff;border:none;border-radius:8px;
     padding:12px;font-family:inherit;font-size:14px;font-weight:600;cursor:pointer;transition:background .2s}
.btn:hover{background:var(--brand-d)}
.alert{padding:10px 14px;border-radius:8px;font-size:13px;margin-bottom:18px}
.ae{background:#FEF2F2;border:1px solid #FECACA;color:#991B1B}
.lr{text-align:center;margin-top:16px;font-size:13px;color:var(--muted)}
.lr a{color:var(--brand);text-decoration:none;font-weight:500}
</style></head><body>
<div class="auth-card">
  <div class="auth-logo">CA<span>Toolkit</span></div>
  <div class="auth-sub">Professional tools for Indian CAs &amp; Accountants</div>
  <h2>Sign in</h2>
  <p style="font-size:13px;color:var(--muted);margin-bottom:24px">Enter your credentials to continue</p>
  {% if error %}<div class="alert ae">{{ error }}</div>{% endif %}
  <form method="POST" action="/login">
    <div class="field"><label>Username</label>
      <input type="text" name="username" placeholder="Enter username" required autocomplete="username"/></div>
    <div class="field"><label>Password</label>
      <input type="password" name="password" placeholder="Enter password" required autocomplete="current-password"/></div>
    <button class="btn" type="submit">Sign In →</button>
  </form>
  <div class="lr">Need access? <a href="mailto:{{ email }}">Contact admin</a></div>
</div><a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
</body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD — tool selection homepage
# ══════════════════════════════════════════════════════════════════════════════

DASHBOARD_T = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard – CA Toolkit</title>
<style>
""" + BASE_CSS + """
.hero{text-align:center;padding:56px 24px 40px;max-width:680px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#EFF6FF;
            color:var(--brand);border:1px solid #BFDBFE;border-radius:99px;
            padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:18px}
.hero h1{font-size:clamp(24px,4vw,38px);font-weight:800;line-height:1.2;
         letter-spacing:-.5px;margin-bottom:12px}
.hero h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:15px;color:var(--muted);line-height:1.7}

.tools-grid{max-width:1320px;margin:0 auto;padding:0 24px 56px;
            display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
@media(max-width:1100px){.tools-grid{grid-template-columns:repeat(3,1fr)}}
@media(max-width:768px){.tools-grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:480px){.tools-grid{grid-template-columns:1fr}}

.tool-card{background:var(--white);border:1.5px solid var(--border);
           border-radius:var(--radius);padding:22px 20px;
           text-decoration:none;color:var(--ink);
           transition:all .25s cubic-bezier(.4,0,.2,1);position:relative;overflow:hidden;display:block;
           animation:fadeUp .5s ease-out both}
.tool-card:nth-child(1){animation-delay:.05s}.tool-card:nth-child(2){animation-delay:.1s}
.tool-card:nth-child(3){animation-delay:.15s}.tool-card:nth-child(4){animation-delay:.2s}
.tool-card:nth-child(5){animation-delay:.25s}.tool-card:nth-child(6){animation-delay:.3s}
.tool-card:nth-child(7){animation-delay:.35s}.tool-card:nth-child(8){animation-delay:.4s}
.tool-card:hover{border-color:var(--brand);box-shadow:0 8px 32px rgba(29,78,216,.12);transform:translateY(-3px)}
.tool-card.disabled{cursor:default;opacity:.7}
.tool-card.disabled:hover{border-color:var(--border);box-shadow:none;transform:none}

.tool-icon{width:44px;height:44px;border-radius:10px;display:flex;
           align-items:center;justify-content:center;font-size:22px;
           margin-bottom:12px}
.tool-card h2{font-size:14px;font-weight:700;margin-bottom:5px}
.tool-card p{font-size:12px;color:var(--muted);line-height:1.6;margin-bottom:14px}

.tool-tag{display:inline-flex;align-items:center;gap:5px;font-size:11px;
          font-weight:600;padding:3px 10px;border-radius:99px}
.tag-live{background:#ECFDF5;color:#065F46}
.tag-soon{background:#F3F4F6;color:var(--muted)}

.tool-card .arrow{position:absolute;right:20px;top:50%;transform:translateY(-50%);
                  font-size:20px;color:var(--brand);opacity:0;transition:opacity .2s}
.tool-card:not(.disabled):hover .arrow{opacity:1}

.usage-strip{max-width:960px;margin:0 auto 8px;padding:0 24px}
.usage-box{background:var(--white);border:1px solid var(--border);border-radius:var(--radius);
           padding:14px 20px;display:flex;align-items:center;justify-content:space-between;
           flex-wrap:wrap;gap:12px}
.usage-info{font-size:13px}
.usage-info strong{color:var(--ink)}
.usage-info span{color:var(--muted)}
.usage-bar-bg{flex:1;max-width:200px;background:#F3F4F6;border-radius:99px;height:6px;overflow:hidden}
.usage-bar-fill{height:100%;border-radius:99px;transition:width .4s}
.upgrade-link{font-size:12px;font-weight:600;color:var(--brand);text-decoration:none}
.upgrade-link:hover{text-decoration:underline}
</style></head><body>

<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <div class="nav-right">
    <span class="nav-user">👤 <strong>{{ username }}</strong>
      <span class="badge b-{{ plan }}">{{ plan_label }}</span>
      {% if is_admin %}<span class="badge" style="background:#EFF6FF;color:var(--brand);margin-left:4px">Admin</span>{% endif %}
    </span>
    {% if username %}
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin Panel</a>{% endif %}
    <a href="/logout" class="nav-link">Sign out</a>
    {% else %}
    <a href="/login" class="nav-btn">Sign In</a>
    {% endif %}
  </div>
</nav>

<div class="hero" style="animation:fadeUp .5s ease-out both">
  <div class="hero-badge">🇮🇳 Made for Indian CAs &amp; Accountants</div>
  <h1>Your Complete <em>CA Toolkit</em></h1>
  <p>Professional tools built by CA Article — designed to save hours of manual work every year.</p>
</div>

{% if username %}
<!-- Upload usage strip -->
<div class="usage-strip">
  <div class="usage-box">
    <div class="usage-info">
      <strong>{{ uploads_remaining }} uploads</strong>
      <span> remaining ({{ uploads_used }} / {{ uploads_total }} used)</span>
      {% if validity_end %}<span style="margin-left:8px;color:#9CA3AF">· Valid till {{ validity_end[:10] }}</span>{% endif %}
    </div>
    <div class="usage-bar-bg">
      <div class="usage-bar-fill"
           style="width:{{ bar_pct }}%;background:{{ '#EF4444' if uploads_remaining==0 else '#F59E0B' if uploads_remaining<=3 else '#10B981' }}">
      </div>
    </div>
    <a href="/tool/converter#pricing" class="upgrade-link">Upgrade plan →</a>
  </div>
</div>
{% endif %}

<!-- Tools grid -->
<div class="tools-grid">

  <!-- PREMIUM: Balance Sheet Year-Shift -->
  {% if username %}
  <a href="/tool/converter" class="tool-card">
  {% else %}
  <a href="/login" class="tool-card">
  {% endif %}
    {% if not username %}<div style="position:absolute;top:12px;right:12px;background:#FEF3C7;color:#92400E;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🔒 Login Required</div>{% endif %}
    <div class="tool-icon" style="background:#EFF6FF">📊</div>
    <h2>Balance Sheet Year-Shift</h2>
    <p>Roll over your comparative Excel balance sheet to any new financial year in seconds. Shifts CY→PY, clears CY, restores all formulas and updates every date.</p>
    <span class="tool-tag tag-live">✓ Live · Premium</span>
    <div class="arrow">→</div>
  </a>

  <!-- PREMIUM: GST Reconciliation -->
  {% if username %}
  <a href="/tool/gst-reconciliation" class="tool-card" style="position:relative">
  {% else %}
  <a href="/login" class="tool-card" style="position:relative">
  {% endif %}
    {% if not username %}<div style="position:absolute;top:12px;right:12px;background:#FEF3C7;color:#92400E;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🔒 Login Required</div>{% else %}<div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">Premium</div>{% endif %}
    <div class="tool-icon" style="background:#FEF3C7">📊</div>
    <h2>GST Reconciliation</h2>
    <p>Compare Sales as per Books vs GSTR 3B returns. Upload your sales summary and GSTR 3B PDFs (ZIP) — get month-wise, state-wise difference report instantly.</p>
    <span class="tool-tag tag-live">✓ Live · Premium</span>
    <div class="arrow">→</div>
  </a>

  <!-- PREMIUM: Balance Sheet from Trial Balance - NOW LIVE -->
  <a href="/tool/tb-to-bs" class="tool-card" style="position:relative">
    <div style="position:absolute;top:12px;right:12px;background:#FEF3C7;color:#92400E;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🔒 Premium</div>
    <div class="tool-icon" style="background:#F0FDF4">📋</div>
    <h2>Balance Sheet from Trial Balance</h2>
    <p>Upload your trial balance and BS template — tool auto-maps accounts and fills CY figures. Zero formatting change.</p>
    <span class="tool-tag tag-live">✓ Live · Premium</span>
    <div class="arrow">→</div>
  </a>

  <!-- FREE TOOLS -->
  <a href="/tool/tax-calculator" class="tool-card">
    <div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🆓 Free</div>
    <div class="tool-icon" style="background:#FFFBEB">🧮</div>
    <h2>Income Tax Calculator</h2>
    <p>Calculate income tax under old and new regime for PY 2025-26. Income under 5 heads, TDS/TCS, surcharge &amp; cess — all built in.</p>
    <span class="tool-tag tag-live">✓ Live · Free</span>
    <div class="arrow">→</div>
  </a>

  <a href="/tool/tds-calculator" class="tool-card">
    <div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🆓 Free</div>
    <div class="tool-icon" style="background:#EFF6FF">📑</div>
    <h2>TDS / TCS Calculator</h2>
    <p>Calculate TDS or TCS as per IT Act 2025 (Sec 393/394). New payment codes, rates, late deposit interest — all in one tool.</p>
    <span class="tool-tag tag-live">✓ Live · Free</span>
    <div class="arrow">→</div>
  </a>

  <a href="/tool/depreciation-calculator" class="tool-card">
    <div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🆓 Free</div>
    <div class="tool-icon" style="background:#F5F3FF">🏭</div>
    <h2>Depreciation Calculator</h2>
    <p>Calculate depreciation under Companies Act 2013 (WDV/SLM) and Income Tax Act. Get full schedule with opening/closing WDV.</p>
    <span class="tool-tag tag-live">✓ Live · Free</span>
    <div class="arrow">→</div>
  </a>

  <a href="/tool/msme-calculator" class="tool-card" style="position:relative">
    <div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🆓 Free</div>
    <div class="tool-icon" style="background:#FEF2F2">📄</div>
    <h2>MSME Disallowance Calculator</h2>
    <p>Upload creditors list and check MSME payment compliance under Sec 43B(h). Overdue payments highlighted with total disallowance amount.</p>
    <span class="tool-tag tag-live">✓ Live · Free</span>
    <div class="arrow">→</div>
  </a>

  <a href="/tool/capital-gains-calculator" class="tool-card" style="position:relative">
    <div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🆓 Free</div>
    <div class="tool-icon" style="background:#F5F3FF">💰</div>
    <h2>Capital Gains Calculator</h2>
    <p>Calculate LTCG/STCG on property, shares, MF and more. Compare old vs new regime, indexation benefit, and find zero-tax sale price with reverse calculator.</p>
    <span class="tool-tag tag-live">✓ Live · Free</span>
    <div class="arrow">→</div>
  </a>

  <div class="tool-card disabled">
    <div class="tool-icon" style="background:#FEF2F2">🚀</div>
    <h2>More Tools Coming Soon</h2>
    <p>We're building more tools for Indian CAs. Stay tuned — new utilities added regularly based on your feedback.</p>
    <span class="tool-tag tag-soon">Stay Tuned</span>
  </div>

</div>

<footer>
  <div class="ft-main">
    <div>
      <div class="ft-brand-name">CA Toolkit</div>
      <p class="ft-brand-desc">CA Toolkit is a comprehensive utility platform built by a CA Article from Ludhiana, Punjab, providing automation tools for Indian Chartered Accountants. The platform saves hours of manual work every year — from Balance Sheet year-shift to GST reconciliation, tax calculations, and more.</p>
    </div>
    <div>
      <div class="ft-col-title">Know More</div>
      <ul class="ft-links">
        <li><a href="/">Home</a></li>
        <li><a href="/">BS Year Shift</a></li>
        <li><a href="/tool/tb-to-bs">TB → Balance Sheet</a></li>
        <li><a href="/tool/tax-calculator">Tax Calculator</a></li>
        <li><a href="/privacy">Privacy Policy</a></li>
      </ul>
    </div>
    <div>
      <div class="ft-col-title">Contact Us</div>
      <div class="ft-contact-name">CA Toolkit</div>
      <div class="ft-contact-addr">Built for Indian Chartered Accountants<br/>Created by CA Article · Ludhiana, Punjab</div>
      <div class="ft-contact-line">Support · <a href="https://wa.me/918427651580" style="color:#9CA3AF">WhatsApp Chat</a></div>
      <div class="ft-socials">
        <a href="https://wa.me/918427651580" target="_blank" title="WhatsApp"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
      </div>
    </div>
  </div>
  <div class="ft-bottom">
    <span class="ft-bottom-left">©2026 CA Toolkit · All Rights Reserved · <a href="/privacy" style="color:#6B7280;text-decoration:none">Privacy Policy</a> · <span style="color:#EF4444">No refund after first upload is used</span></span>
    <span class="ft-bottom-right">Built for Indian CAs · Ludhiana, Punjab</span>
  </div>
</footer>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support">
  <svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
</a>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
</body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
#  BALANCE SHEET CONVERTER TOOL PAGE
# ══════════════════════════════════════════════════════════════════════════════

CONVERTER_T = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Balance Sheet Year-Shift – CA Toolkit</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>
<style>
""" + BASE_CSS + """
.nav-links{display:flex;gap:20px;list-style:none}
.nav-links a{text-decoration:none;color:var(--muted);font-size:13px;font-weight:500;transition:color .2s}
.nav-links a:hover{color:var(--brand)}
.hero{text-align:center;padding:56px 24px 40px;max-width:700px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#EFF6FF;
            color:var(--brand);border:1px solid #BFDBFE;border-radius:99px;
            padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:18px}
h1{font-size:clamp(24px,4vw,40px);font-weight:800;line-height:1.15;
   letter-spacing:-.5px;margin-bottom:14px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:15px;color:var(--muted);line-height:1.7;max-width:520px;margin:0 auto 28px}
.stats{display:flex;justify-content:center;gap:36px;flex-wrap:wrap;
       padding:16px 24px;background:var(--white);
       border-top:1px solid var(--border);border-bottom:1px solid var(--border)}
.stat-n{font-size:20px;font-weight:800;color:var(--brand)}
.stat-l{font-size:11px;color:var(--muted);margin-top:2px}
.main{max-width:1080px;margin:0 auto;padding:40px 24px;
      display:grid;grid-template-columns:1fr 1fr;gap:24px;align-items:start}
@media(max-width:768px){.main{grid-template-columns:1fr}}
.card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);
      box-shadow:var(--shadow);overflow:hidden;animation:fadeUp .4s ease-out both}
.card-head{padding:16px 20px;border-bottom:1px solid var(--border);
           display:flex;align-items:center;gap:10px}
.card-head .icon{width:32px;height:32px;border-radius:8px;display:flex;
                 align-items:center;justify-content:center;font-size:16px}
.card-head h2{font-size:14px;font-weight:700}
.card-head p{font-size:12px;color:var(--muted);margin-top:1px}
.card-body{padding:20px}
.usage-row{display:flex;justify-content:space-between;align-items:center;
           font-size:12px;font-weight:600;margin-bottom:5px}
.usage-bar-bg{background:#F3F4F6;border-radius:99px;height:6px;overflow:hidden;margin-bottom:14px}
.usage-bar-fill{height:100%;border-radius:99px;transition:width .4s}
.field{margin-bottom:16px}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;
      letter-spacing:.04em;color:var(--muted);margin-bottom:5px}
.hint{font-size:11px;color:var(--muted);margin-top:4px}
.dropzone{border:2px dashed var(--border);border-radius:10px;padding:24px 14px;
          text-align:center;cursor:pointer;transition:all .2s;position:relative;background:var(--bg)}
.dropzone:hover,.dropzone.drag{border-color:var(--brand);background:#EFF6FF}
.dropzone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.dz-icon{font-size:26px;margin-bottom:6px}
.dz-text{font-size:12px;color:var(--muted)}
.dz-text strong{color:var(--brand)}
.dz-file{font-size:12px;font-weight:600;color:var(--green);margin-top:5px;display:none}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
input[type=number],input[type=text]{width:100%;border:1.5px solid var(--border);
  border-radius:8px;padding:9px 12px;font-family:inherit;font-size:13px;
  color:var(--ink);background:var(--white);transition:border-color .2s;outline:none}
input:focus{border-color:var(--brand)}
.btn{width:100%;background:var(--brand);color:#fff;border:none;border-radius:10px;
     padding:12px;font-family:inherit;font-size:14px;font-weight:700;cursor:pointer;
     transition:background .2s;display:flex;align-items:center;justify-content:center;gap:8px}
.btn:hover{background:var(--brand-d)}
.btn:disabled{background:#93C5FD;cursor:not-allowed}
.spinner{width:16px;height:16px;border:2px solid rgba(255,255,255,.3);
         border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite;display:none}
@keyframes spin{to{transform:rotate(360deg)}}
#status{margin-top:12px;border-radius:8px;padding:12px 14px;font-size:13px;display:none;line-height:1.6}
#status.success{background:#ECFDF5;border:1px solid #A7F3D0;color:#065F46}
#status.error{background:#FEF2F2;border:1px solid #FECACA;color:#991B1B}
.log-list{margin-top:6px;padding-left:14px;font-size:11px;color:#374151;line-height:2}
.dl-btn{display:none;margin-top:10px;width:100%;background:var(--green);color:#fff;
        border:none;border-radius:10px;padding:11px;font-family:inherit;font-size:13px;
        font-weight:600;cursor:pointer;text-decoration:none;text-align:center;transition:background .2s}
.dl-btn:hover{background:#059669}
.steps{padding:0;list-style:none;counter-reset:step}
.steps li{display:flex;gap:10px;align-items:flex-start;padding:12px 0;border-bottom:1px solid var(--border)}
.steps li:last-child{border:none}
.steps li::before{counter-increment:step;content:counter(step);min-width:24px;height:24px;
                  background:var(--brand);color:#fff;border-radius:50%;display:flex;
                  align-items:center;justify-content:center;font-size:11px;font-weight:700;margin-top:1px}
.steps li strong{display:block;font-size:12px;font-weight:600;margin-bottom:2px}
.steps li span{font-size:11px;color:var(--muted)}
.features{max-width:1080px;margin:0 auto;padding:0 24px 40px;
          display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
@media(max-width:640px){.features{grid-template-columns:1fr}}
.feat{background:var(--white);border:1px solid var(--border);border-radius:var(--radius);
      padding:20px;text-align:center}
.feat .fi{font-size:26px;margin-bottom:8px}
.feat h3{font-size:13px;font-weight:700;margin-bottom:4px}
.feat p{font-size:12px;color:var(--muted);line-height:1.6}
.pricing-section{background:var(--white);border-top:1px solid var(--border);
                 border-bottom:1px solid var(--border);padding:48px 24px}
.pricing-section h2{text-align:center;font-size:24px;font-weight:800;margin-bottom:6px}
.psub{text-align:center;color:var(--muted);font-size:13px;margin-bottom:32px}
.plans{max-width:1080px;margin:0 auto;
       display:grid;grid-template-columns:repeat(6,1fr);gap:14px}
@media(max-width:900px){.plans{grid-template-columns:repeat(3,1fr)}}
@media(max-width:480px){.plans{grid-template-columns:1fr}}
.plan{border:1.5px solid var(--border);border-radius:var(--radius);padding:20px 16px;position:relative}
.plan.pop{border-color:var(--brand)}
.plan-badge{position:absolute;top:-10px;left:50%;transform:translateX(-50%);
            background:var(--brand);color:#fff;font-size:10px;font-weight:700;
            padding:2px 10px;border-radius:99px;white-space:nowrap}
.plan-name{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;
           letter-spacing:.06em;margin-bottom:6px}
.plan-price{font-size:24px;font-weight:800;color:var(--ink);margin-bottom:2px}
.plan-uploads{font-size:12px;font-weight:700;color:var(--brand);margin-bottom:2px}
.plan-validity{font-size:10px;color:var(--muted);margin-bottom:14px}
.plan ul{list-style:none;margin-bottom:16px}
.plan ul li{font-size:11px;padding:3px 0;display:flex;gap:5px}
.plan ul li::before{content:"✓";color:var(--green);font-weight:700}
.plan-btn{display:block;text-align:center;padding:8px;border-radius:7px;
          font-size:11px;font-weight:600;text-decoration:none;transition:all .2s;
          border:1.5px solid var(--brand);color:var(--brand)}
.plan-btn:hover{background:var(--brand);color:#fff}
.plan.pop .plan-btn{background:var(--brand);color:#fff}
.no-refund-note{text-align:center;font-size:11px;color:var(--muted);margin-top:14px;font-weight:500}
.faq-section{max-width:720px;margin:0 auto;padding:40px 24px}
.faq-section h2{font-size:20px;font-weight:800;text-align:center;margin-bottom:24px}
details{border:1px solid var(--border);border-radius:10px;margin-bottom:8px}
summary{padding:12px 16px;font-size:13px;font-weight:600;cursor:pointer;
        list-style:none;display:flex;justify-content:space-between;align-items:center}
summary::after{content:"＋";color:var(--muted)}
details[open] summary::after{content:"－"}
details p{padding:0 16px 12px;font-size:12px;color:var(--muted);line-height:1.7}
.contact-section{background:#EFF6FF;border-top:1px solid #BFDBFE;padding:36px 24px;text-align:center}
.contact-section h2{font-size:18px;font-weight:800;margin-bottom:6px}
.contact-section p{font-size:13px;color:var(--muted);margin-bottom:14px}
.contact-grid{display:flex;justify-content:center;gap:20px;flex-wrap:wrap}
.contact-item{background:var(--white);border:1px solid var(--border);border-radius:10px;
              padding:14px 20px;font-size:13px}
.contact-item strong{display:block;font-size:10px;text-transform:uppercase;
                     letter-spacing:.05em;color:var(--muted);margin-bottom:3px}
.contact-item a{color:var(--brand);text-decoration:none;font-weight:600}
.limit-banner{max-width:640px;margin:0 auto 0;padding:0 24px}
.limit-box{background:#FEF2F2;border:1px solid #FECACA;border-radius:var(--radius);
           padding:20px 24px;text-align:center;margin-top:16px}
.limit-box h3{font-size:15px;font-weight:700;color:#991B1B;margin-bottom:8px}
.limit-box p{font-size:13px;color:#7F1D1D;line-height:1.7;margin-bottom:10px}
.limit-box a{color:var(--brand);font-weight:600;text-decoration:none}
.toast{position:fixed;bottom:24px;right:24px;background:var(--ink);color:#fff;
       padding:11px 18px;border-radius:10px;font-size:13px;font-weight:500;
       transform:translateY(80px);transition:transform .3s;z-index:999}
.toast.show{transform:translateY(0)}
@media(max-width:480px){.row2{grid-template-columns:1fr}}
</style></head><body>

<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <ul class="nav-links">
    <li><a href="#tool">Tool</a></li>
    <li><a href="#pricing">Pricing</a></li>
    <li><a href="#faq">FAQ</a></li>
    <li><a href="#contact">Contact</a></li>
  </ul>
  <div class="nav-right">
    <span class="nav-user">👤 <strong>{{ username }}</strong>
      <span class="badge b-{{ plan }}">{{ plan_label }}</span>
      {% if is_admin %}<span class="badge" style="background:#EFF6FF;color:var(--brand);margin-left:4px">Admin</span>{% endif %}
    </span>
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin</a>{% endif %}
    <a href="/" class="nav-btn" style="background:#F3F4F6;color:var(--ink)">← Dashboard</a>
    {% if username %}<a href="/logout" class="nav-link">Sign out</a>
    {% else %}<a href="/login" class="nav-btn">Sign In</a>{% endif %}
  </div>
</nav>

<section class="hero">
  <div class="hero-badge">🇮🇳 CA Tool · Balance Sheet Year-Shift</div>
  <h1>Roll Over to <em>Any Financial Year</em><br/>in Seconds</h1>
  <p>Upload your comparative Excel balance sheet — shifts CY→PY, clears CY column, restores all formulas, and updates every date automatically.</p>
</section>

{% if uploads_left == 0 %}
<div class="limit-banner">
  <div class="limit-box">
    <h3>🔒 No uploads remaining</h3>
    <p>You've used all your uploads. Contact us to recharge your account.<br/>
       Pay via UPI and email your screenshot — upgraded within a few hours.</p>
    <p>📧 <a href="mailto:{{ contact_email }}">{{ contact_email }}</a> &nbsp;|&nbsp;
       💳 UPI: <strong>{{ contact_upi }}</strong></p>
    <p style="font-size:11px;color:#9CA3AF;margin-top:8px">No refund after first upload is used.</p>
  </div>
</div>
{% endif %}

<div class="stats">
  <div class="stat"><div class="stat-n">100%</div><div class="stat-l">Formatting preserved</div></div>
  <div class="stat"><div class="stat-n">All sheets</div><div class="stat-l">Processed at once</div></div>
  <div class="stat"><div class="stat-n">&lt;10 sec</div><div class="stat-l">Processing time</div></div>
  <div class="stat"><div class="stat-n">Any format</div><div class="stat-l">Works with all CA templates</div></div>
</div>

<div class="main" id="tool">
  <div class="card">
    <div class="card-head">
      <div class="icon" style="background:#EFF6FF">📊</div>
      <div>
        <h2>Process Your Balance Sheet</h2>
        <p>{{ plan_label }} · {{ uploads_left }} upload{{ 's' if uploads_left != 1 else '' }} remaining</p>
      </div>
    </div>
    <div class="card-body">
      <div class="usage-row">
        <span style="color:var(--muted)">Uploads used</span>
        <span><strong>{{ uploads_used }}</strong> / {{ uploads_total }}
          {% if validity_end %}<span style="color:#9CA3AF;font-weight:400"> · expires {{ validity_end[:10] }}</span>{% endif %}
        </span>
      </div>
      <div class="usage-bar-bg">
        <div class="usage-bar-fill"
             style="width:{{ bar_pct }}%;background:{{ '#EF4444' if uploads_left==0 else '#F59E0B' if uploads_left<=3 else '#10B981' }}">
        </div>
      </div>
      <div class="field">
        <label>Upload Excel File (.xlsx / .xls)</label>
        <div class="dropzone" id="dropzone">
          <input type="file" id="xlFile" accept=".xlsx,.xls" {{ 'disabled' if uploads_left==0 else '' }}/>
          <div class="dz-icon">📁</div>
          <div class="dz-text"><strong>Click to browse</strong> or drag &amp; drop</div>
          <div class="dz-text" style="margin-top:3px">.xlsx or .xls · Max 20 MB</div>
          <div class="dz-file" id="dzFile"></div>
        </div>
      </div>
      <div class="row2">
        <div class="field">
          <label>Closing Year (CY)</label>
          <input type="number" id="closingYear" placeholder="e.g. 2025" min="2000" max="2100"
                 {{ 'disabled' if uploads_left==0 else '' }}/>
          <p class="hint">Year ending 31.03.YYYY</p>
        </div>
        <div class="field">
          <label>New Year</label>
          <input type="number" id="newYear" placeholder="Auto-filled" readonly/>
          <p class="hint">Auto-filled</p>
        </div>
      </div>
      <div class="field">
        <label>Output Filename <span style="font-weight:400;text-transform:none;color:var(--muted)">(optional)</span></label>
        <input type="text" id="outputName" placeholder="e.g. ClientName_BS"
               {{ 'disabled' if uploads_left==0 else '' }}/>
        <p class="hint">Leave blank to auto-generate</p>
      </div>
      <button class="btn" id="processBtn" onclick="processFile()"
              {{ 'disabled' if uploads_left==0 else '' }}>
        <span id="btnText">⚡ Process &amp; Download</span>
        <div class="spinner" id="spinner"></div>
      </button>
      <div id="status"></div>
      <a id="dlBtn" class="dl-btn" href="#">⬇&nbsp; Download Processed File</a>
    </div>
  </div>

  <div>
    <div class="card" style="margin-bottom:18px">
      <div class="card-head">
        <div class="icon" style="background:#F0FDF4">✅</div>
        <div><h2>How It Works</h2><p>4 steps, fully automatic</p></div>
      </div>
      <div class="card-body">
        <ol class="steps">
          <li><strong>Upload your Excel file</strong>
              <span>Your FY comparative balance sheet with CY and PY columns</span></li>
          <li><strong>Auto-detects all CY/PY columns</strong>
              <span>Scans every sheet and finds correct data columns automatically</span></li>
          <li><strong>Shifts CY → PY, clears CY</strong>
              <span>Values become PY. Formulas and cross-sheet links restored. CY cleared for fresh entry</span></li>
          <li><strong>Updates every date</strong>
              <span>All date strings across every sheet updated in one shot</span></li>
        </ol>
      </div>
    </div>
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#FFFBEB">🔒</div>
        <div><h2>Your Data is Safe</h2><p>Privacy first</p></div>
      </div>
      <div class="card-body">
        <ul class="steps">
          <li><strong>Deleted immediately after download</strong>
              <span>File removed from server the moment processing completes</span></li>
          <li><strong>HTTPS encrypted</strong><span>All transfers encrypted end-to-end</span></li>
          <li><strong>No data stored</strong><span>We never read, store, or share your financial data</span></li>
        </ul>
      </div>
    </div>
  </div>
</div>

<div class="features">
  <div class="feat"><div class="fi">🧮</div><h3>Formulas Preserved</h3><p>Every SUM and cross-sheet reference in the PY column is snapshotted and restored automatically.</p></div>
  <div class="feat"><div class="fi">🎨</div><h3>Formatting Intact</h3><p>Fonts, borders, colors, merged cells, column widths — everything preserved exactly.</p></div>
  <div class="feat"><div class="fi">📅</div><h3>All Dates Updated</h3><p>Every date string across every sheet updated in one shot.</p></div>
  <div class="feat"><div class="fi">🗂️</div><h3>All Sheets at Once</h3><p>BS, P&L, Notes, Capital, Fixed Assets — every sheet processed together.</p></div>
  <div class="feat"><div class="fi">🔄</div><h3>Any CA Template</h3><p>Auto-detects column positions. Works with any firm's template.</p></div>
  <div class="feat"><div class="fi">⚡</div><h3>Instant Results</h3><p>What took 30–45 minutes of manual work now takes under 10 seconds.</p></div>
</div>

<section class="pricing-section" id="pricing">
  <h2>Simple Pricing</h2>
  <p class="psub">Upload-based · 3-month validity · Uploads stack when you recharge</p>
  <div class="plans">
    <div class="plan">
      <div class="plan-name">Free</div>
      <div class="plan-price">₹0</div>
      <div class="plan-uploads">2 uploads</div>
      <div class="plan-validity">Try it out</div>
      <ul><li>All features</li><li>All sheet types</li><li>Up to 20 MB</li></ul>
      <a href="#tool" class="plan-btn">Get Started</a>
    </div>
    <div class="plan">
      <div class="plan-name">Starter</div>
      <div class="plan-price">₹60</div>
      <div class="plan-uploads">10 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features</li><li>All sheet types</li><li>Up to 20 MB</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
    <div class="plan pop">
      <div class="plan-badge">Most Popular</div>
      <div class="plan-name">Standard</div>
      <div class="plan-price">₹130</div>
      <div class="plan-uploads">25 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features</li><li>Priority support</li><li>Up to 20 MB</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
    <div class="plan">
      <div class="plan-name">Professional</div>
      <div class="plan-price">₹270</div>
      <div class="plan-uploads">60 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features</li><li>Priority support</li><li>Up to 20 MB</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
    <div class="plan">
      <div class="plan-name">Firm</div>
      <div class="plan-price">₹600</div>
      <div class="plan-uploads">150 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features</li><li>WhatsApp support</li><li>Up to 20 MB</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
    <div class="plan">
      <div class="plan-name">CA Admin</div>
      <div class="plan-price">₹1,000</div>
      <div class="plan-uploads">500 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features + GST Recon</li><li>WhatsApp support</li><li>Best for CA firms</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
  </div>
  <p class="no-refund-note">⚠ No refund after first upload is used &nbsp;·&nbsp; Unused uploads stack when you recharge before expiry</p>
</section>

<section class="faq-section" id="faq">
  <h2>Frequently Asked Questions</h2>
  <details><summary>Which Excel formats are supported?</summary>
    <p>.xlsx (Excel 2007+) and .xls (legacy Excel). Both are fully supported — .xls files are automatically converted before processing.</p></details>
  <details><summary>Will it work with my firm's custom template?</summary>
    <p>Yes. Auto-detects CY/PY columns by scanning date headers like "31.03.2025". Works with any Indian CA template.</p></details>
  <details><summary>Are my formulas and formatting safe?</summary>
    <p>Yes. Formulas in PY column are snapshotted before and restored after. Formatting is never touched.</p></details>
  <details><summary>What happens to my uploaded file?</summary>
    <p>Deleted from our server immediately after you download the result. We never store or share your data.</p></details>
  <details><summary>How do I purchase a plan?</summary>
    <p>Pay via UPI to <strong>{{ contact_upi }}</strong> and email your screenshot to <strong>{{ contact_email }}</strong>. Account upgraded within a few hours.</p></details>
  <details><summary>Do unused uploads carry over when I recharge?</summary>
    <p>Yes. Remaining uploads stack on top of the new plan if you recharge before expiry.</p></details>
</section>

<section class="contact-section" id="contact">
  <h2>Purchase a Plan</h2>
  <p>Pay via UPI and send your payment screenshot to our email. We'll upgrade your account within a few hours.</p>
  <div class="contact-grid">
    <div class="contact-item">
      <strong>Email</strong>
      <a href="mailto:{{ contact_email }}">{{ contact_email }}</a>
    </div>
    <div class="contact-item">
      <strong>UPI Payment</strong>
      <span style="font-weight:600;color:var(--ink)">{{ contact_upi }}</span>
    </div>
  </div>
</section>

<footer>
  <div class="ft-main">
    <div>
      <div class="ft-brand-name">CA Toolkit</div>
      <p class="ft-brand-desc">CA Toolkit is a comprehensive utility platform built by a CA Article from Ludhiana, Punjab, providing automation tools for Indian Chartered Accountants. The platform saves hours of manual work every year — from Balance Sheet year-shift to GST reconciliation, tax calculations, and more.</p>
    </div>
    <div>
      <div class="ft-col-title">Know More</div>
      <ul class="ft-links">
        <li><a href="/">Home</a></li>
        <li><a href="/">BS Year Shift</a></li>
        <li><a href="/tool/tb-to-bs">TB → Balance Sheet</a></li>
        <li><a href="/tool/tax-calculator">Tax Calculator</a></li>
        <li><a href="/privacy">Privacy Policy</a></li>
      </ul>
    </div>
    <div>
      <div class="ft-col-title">Contact Us</div>
      <div class="ft-contact-name">CA Toolkit</div>
      <div class="ft-contact-addr">Built for Indian Chartered Accountants<br/>Created by CA Article · Ludhiana, Punjab</div>
      <div class="ft-contact-line">Support · <a href="https://wa.me/918427651580" style="color:#9CA3AF">WhatsApp Chat</a></div>
      <div class="ft-socials">
        <a href="https://wa.me/918427651580" target="_blank" title="WhatsApp"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
      </div>
    </div>
  </div>
  <div class="ft-bottom">
    <span class="ft-bottom-left">©2026 CA Toolkit · All Rights Reserved · <a href="/privacy" style="color:#6B7280;text-decoration:none">Privacy Policy</a> · <span style="color:#EF4444">No refund after first upload is used</span></span>
    <span class="ft-bottom-right">Built for Indian CAs · Ludhiana, Punjab</span>
  </div>
</footer>
<div class="toast" id="toast"></div>

<script>
const dz=document.getElementById('dropzone'),fi=document.getElementById('xlFile'),dzFile=document.getElementById('dzFile');
if(dz&&fi){
  dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag');});
  dz.addEventListener('dragleave',()=>dz.classList.remove('drag'));
  dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('drag');
    if(e.dataTransfer.files.length){fi.files=e.dataTransfer.files;showFile(fi.files[0]);}});
  fi.addEventListener('change',()=>{if(fi.files.length)showFile(fi.files[0]);});
}
function showFile(f){dzFile.textContent='✓ '+f.name;dzFile.style.display='block';}
document.getElementById('closingYear').addEventListener('input',function(){
  const v=parseInt(this.value);if(!isNaN(v))document.getElementById('newYear').value=v+1;});
async function processFile(){
  const f=fi?fi.files[0]:null,cYr=parseInt(document.getElementById('closingYear').value),
        nYr=parseInt(document.getElementById('newYear').value),
        oNm=document.getElementById('outputName').value.trim(),
        btn=document.getElementById('processBtn'),sp=document.getElementById('spinner'),
        bt=document.getElementById('btnText'),dl=document.getElementById('dlBtn');
  if(!f){showStatus('error','✗ Please select an Excel file first.');return;}
  if(isNaN(cYr)){showStatus('error','✗ Enter a valid closing year.');return;}
  btn.disabled=true;sp.style.display='block';bt.textContent='Processing…';
  dl.style.display='none';showStatus('','');
  const fd=new FormData();
  fd.append('file',f);fd.append('closing_year',cYr);fd.append('new_year',nYr);fd.append('output_name',oNm);
  try{
    const res=await fetch('/process',{method:'POST',body:fd});
    const ct=res.headers.get('content-type')||'';
    if(!ct.includes('application/json')){
      showStatus('error','✗ Server error (non-JSON response). Please try again or contact support.');return;
    }
    const data=await res.json();
    if(data.status==='success'){
      const logHtml='<ul class="log-list">'+data.log.map(l=>`<li>${l}</li>`).join('')+'</ul>';
      showStatus('success','✓ Done! Your file is ready.'+logHtml);
      dl.href='/download/'+data.file_id+'?fn='+encodeURIComponent(data.filename);dl.download=data.filename;
      dl.textContent='⬇  Download — '+data.filename;dl.style.display='block';
      toast('Processed successfully!');
    }else{showStatus('error','✗ '+data.message);}
  }catch(e){showStatus('error','✗ Network error: '+e.message);}
  finally{btn.disabled=false;sp.style.display='none';bt.textContent='⚡ Process & Download';}
}
function showStatus(t,m){const e=document.getElementById('status');e.className=t;e.innerHTML=m;e.style.display=m?'block':'none';}
function toast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),3000);}
</script>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>

<button class="help-btn" onclick="openHelp()" title="How to use this tool">?</button>
<div class="help-overlay" id="helpOverlay">
  <div class="help-modal">
    <div class="help-modal-head"><h3>How to Use — Balance Sheet Year Shift</h3><button class="help-close" onclick="closeHelp()">&#10005;</button></div>
    <div class="help-modal-body"><div class="help-step"><div class="help-step-num">1</div><div class="help-step-body"><h4>Upload BS File</h4><p>Click or drag-drop your comparative Excel balance sheet (.xlsx). It needs CY and PY columns with date headers like '31.03.2025'.</p></div></div><div class="help-step"><div class="help-step-num">2</div><div class="help-step-body"><h4>Enter Years</h4><p>Set Closing Year (e.g. 2025) and New Year (2026). New Year = Closing Year + 1.</p></div></div><div class="help-step"><div class="help-step-num">3</div><div class="help-step-body"><h4>Optional: Custom Filename</h4><p>Enter a custom output name, or leave blank for auto-naming.</p></div></div><div class="help-step"><div class="help-step-num">4</div><div class="help-step-body"><h4>Click Process</h4><p>The tool shifts CY→PY, clears CY columns, updates all dates, and rolls over Fixed Assets automatically.</p></div></div><div class="help-step"><div class="help-step-num">5</div><div class="help-step-body"><h4>Download</h4><p>Download the result. Your file is auto-deleted from our server within minutes.</p></div></div><div class="help-tip">✅ Works with DP Thapar, HFPL, Atultex, and most Indian CA firm templates. Formulas in PY are preserved.</div></div>
  </div>
</div>
<script>function openHelp(){document.getElementById('helpOverlay').classList.add('open')}function closeHelp(){document.getElementById('helpOverlay').classList.remove('open')}document.getElementById('helpOverlay').addEventListener('click',function(e){if(e.target===this)closeHelp()})</script>
</body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
#  INCOME TAX CALCULATOR — PY 2025-26 / AY 2026-27
# ══════════════════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════════════════
#  INCOME TAX CALCULATOR — Multi-Year: PY 2023-24 to PY 2026-27
# ══════════════════════════════════════════════════════════════════════════════

TAX_CALC_T = r"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Income Tax Calculator – CA Toolkit</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>
<style>
""" + BASE_CSS + r"""
.nav-links{display:flex;gap:20px;list-style:none}
.nav-links a{text-decoration:none;color:var(--muted);font-size:13px;font-weight:500;transition:color .2s}
.nav-links a:hover{color:var(--brand)}

.hero{text-align:center;padding:44px 24px 32px;max-width:700px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#FFFBEB;
            color:#92400E;border:1px solid #FDE68A;border-radius:99px;
            padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:18px}
h1{font-size:clamp(22px,3.5vw,34px);font-weight:800;line-height:1.15;letter-spacing:-.5px;margin-bottom:12px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:14px;color:var(--muted);line-height:1.7;max-width:520px;margin:0 auto}

.main-wrap{max-width:1200px;margin:0 auto;padding:0 24px 48px}

.regime-toggle{display:flex;justify-content:center;gap:12px;margin-bottom:28px;flex-wrap:wrap}
.regime-btn{padding:10px 28px;border-radius:10px;border:2px solid var(--border);
            background:var(--white);font-family:inherit;font-size:13px;font-weight:700;
            cursor:pointer;transition:all .2s;color:var(--muted)}
.regime-btn.active{border-color:var(--brand);background:#EFF6FF;color:var(--brand);box-shadow:0 2px 12px rgba(29,78,216,.12)}
.regime-btn:hover{border-color:var(--brand)}

.calc-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:start}
@media(max-width:860px){.calc-grid{grid-template-columns:1fr}}
.card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);
      box-shadow:var(--shadow);overflow:hidden;margin-bottom:20px}
.card-head{padding:14px 20px;border-bottom:1px solid var(--border);
           display:flex;align-items:center;gap:10px}
.card-head .icon{width:32px;height:32px;border-radius:8px;display:flex;
                 align-items:center;justify-content:center;font-size:16px}
.card-head h2{font-size:14px;font-weight:700}
.card-head p{font-size:12px;color:var(--muted);margin-top:1px}
.card-body{padding:20px}

.section-title{font-size:12px;font-weight:700;color:var(--brand);text-transform:uppercase;
               letter-spacing:.06em;margin:16px 0 10px;padding-bottom:6px;
               border-bottom:1px solid var(--border);display:flex;align-items:center;gap:6px}
.section-title:first-child{margin-top:0}
.field{margin-bottom:12px}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;
      letter-spacing:.04em;color:var(--muted);margin-bottom:4px}
.hint{font-size:10px;color:var(--muted);margin-top:3px;font-style:italic}
input[type=number]{width:100%;border:1.5px solid var(--border);border-radius:8px;
    padding:9px 12px;font-family:inherit;font-size:13px;color:var(--ink);
    background:var(--white);outline:none;transition:border-color .2s}
input[type=number]:focus{border-color:var(--brand)}
input[type=number]::-webkit-outer-spin-button,
input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none;margin:0}
input[type=number]{-moz-appearance:textfield}
select{width:100%;border:1.5px solid var(--border);border-radius:8px;padding:9px 12px;
       font-family:inherit;font-size:13px;color:var(--ink);background:var(--white);outline:none}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
@media(max-width:480px){.row2,.row3{grid-template-columns:1fr}}

.btn-calc{width:100%;background:var(--brand);color:#fff;border:none;border-radius:10px;
          padding:13px;font-family:inherit;font-size:14px;font-weight:700;cursor:pointer;
          transition:all .2s;display:flex;align-items:center;justify-content:center;gap:8px;
          margin-top:8px}
.btn-calc:hover{background:var(--brand-d);transform:translateY(-1px);box-shadow:0 4px 16px rgba(29,78,216,.2)}
.btn-reset{width:100%;background:#F3F4F6;color:var(--ink);border:none;border-radius:10px;
           padding:10px;font-family:inherit;font-size:13px;font-weight:600;cursor:pointer;
           transition:background .2s;margin-top:8px}
.btn-reset:hover{background:#E5E7EB}

.result-panel{display:none}
.result-panel.show{display:block}
.result-row{display:flex;justify-content:space-between;align-items:center;
            padding:10px 0;border-bottom:1px solid var(--border);font-size:13px}
.result-row:last-child{border-bottom:none}
.result-row .lbl{color:var(--muted);font-weight:500}
.result-row .val{font-weight:700;color:var(--ink);text-align:right}
.result-row.total{padding:14px 0;font-size:15px}
.result-row.total .lbl{color:var(--ink);font-weight:800}
.result-row.total .val{color:var(--brand);font-size:17px}
.result-row.refund .val{color:var(--green)}
.result-row.payable .val{color:var(--red)}
.result-row.sub{font-size:12px;padding:6px 0}
.result-row.sub .lbl{padding-left:16px;font-size:11px}
.result-row.sub .val{font-size:12px}

.compare-box{background:linear-gradient(135deg,#EFF6FF,#FFFBEB);border:2px solid var(--brand);
             border-radius:var(--radius);padding:20px;text-align:center;margin-top:16px}
.compare-box h3{font-size:14px;font-weight:800;margin-bottom:6px}
.compare-box .savings{font-size:28px;font-weight:800;color:var(--green);margin:8px 0}
.compare-box .regime-winner{font-size:13px;color:var(--muted)}
.compare-box .regime-winner strong{color:var(--ink)}
.compare-table{width:100%;margin-top:14px;font-size:12px;border-collapse:collapse}
.compare-table th{text-align:center;font-size:10px;text-transform:uppercase;letter-spacing:.06em;
                  color:var(--muted);padding:6px 8px;border-bottom:1.5px solid var(--border)}
.compare-table td{text-align:center;padding:8px;border-bottom:1px solid var(--border);font-weight:600}
.compare-table .winner{background:#ECFDF5;color:#065F46;border-radius:6px}

.slab-table{width:100%;font-size:12px;border-collapse:collapse;margin-top:8px}
.slab-table th{text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.06em;
               color:var(--muted);padding:6px 8px;border-bottom:1.5px solid var(--border)}
.slab-table td{padding:7px 8px;border-bottom:1px solid var(--border);font-size:12px}
.slab-table tr:last-child td{border-bottom:none}
.slab-table .amt{text-align:right;font-weight:700;font-family:'Inter',monospace}

.disclaimer{font-size:11px;color:var(--muted);line-height:1.6;margin-top:16px;
            padding:12px;background:#F9FAFB;border-radius:8px;border:1px solid var(--border)}
.disclaimer.future{background:#FFFBEB;border-color:#FDE68A}
.toast{position:fixed;bottom:24px;right:24px;background:var(--ink);color:#fff;
       padding:11px 18px;border-radius:10px;font-size:13px;font-weight:500;
       transform:translateY(80px);transition:transform .3s;z-index:999}
.toast.show{transform:translateY(0)}

.print-btn{display:inline-flex;align-items:center;gap:5px;background:#F3F4F6;color:var(--ink);
           border:1px solid var(--border);border-radius:8px;padding:7px 14px;font-size:12px;
           font-weight:600;cursor:pointer;font-family:inherit;transition:all .2s;margin-top:8px}
.print-btn:hover{background:#E5E7EB}

.year-pills{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:6px}
.year-pill{padding:6px 14px;border-radius:8px;border:1.5px solid var(--border);background:var(--white);
           font-size:12px;font-weight:600;cursor:pointer;transition:all .2s;color:var(--muted);font-family:inherit}
.year-pill.active{border-color:var(--brand);background:#EFF6FF;color:var(--brand)}
.year-pill:hover{border-color:var(--brand)}
.year-pill .future-tag{font-size:9px;background:#FDE68A;color:#92400E;padding:1px 5px;border-radius:4px;margin-left:4px;font-weight:700}

.at-group-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin:12px 0 6px;padding:3px 8px;background:var(--surface);border-radius:5px;display:inline-block}
.at-btn-row{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:4px}
.at-btn{padding:7px 13px;border:1.5px solid var(--border);border-radius:20px;background:#fff;font-size:12px;font-weight:500;color:var(--ink);cursor:pointer;transition:all .15s;white-space:nowrap;font-family:inherit}
.at-btn:hover{border-color:var(--brand);color:var(--brand);background:#EFF6FF}
.at-btn.active{border-color:var(--brand);background:var(--brand);color:#fff;font-weight:700;box-shadow:0 2px 8px rgba(37,99,235,.25)}

.mat-row:last-child{border-bottom:none}
.mat-row .ml{color:#78350F;font-weight:500}
.mat-row .mv{font-weight:700;color:#92400E}
.mat-row.mt{font-size:13px;padding:10px 0;border-top:1.5px solid #FDE68A;margin-top:4px}
.mat-row.mt .ml{color:#451A03;font-weight:800}
.mat-row.mt .mv{color:#B45309;font-size:14px}
.mat-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 12px;border-radius:20px;font-size:11px;font-weight:700;margin-top:10px}
.mat-badge.normal{background:#ECFDF5;color:#065F46}
.mat-badge.mat{background:#FEF3C7;color:#92400E}

.at-table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:6px}
.at-table th{text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);padding:7px 10px;border-bottom:2px solid var(--border);background:#F9FAFB}
.at-table td{padding:10px 10px;border-bottom:1px solid var(--border);vertical-align:top}
.at-table tr:last-child td{border-bottom:none}
.at-table .due{font-weight:700;color:var(--brand)}
.at-table .pct{font-weight:800;font-size:14px;color:#1E40AF;text-align:center}
.at-table .cumul{font-size:11px;color:var(--muted)}
.at-table .amt-cell{font-weight:700;color:var(--ink);text-align:right}
.at-table tr.overdue td{background:#FEF2F2}
.at-table tr.upcoming td{background:#FFFBEB}
.at-table tr.done td{background:#F0FDF4}

.assessee-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 12px;border-radius:20px;font-size:11px;font-weight:700;background:#EFF6FF;color:#1E40AF;margin-bottom:10px}

/* ═══════════════════════════════════════════
   ANIMATIONS & MICRO-INTERACTIONS
   ═══════════════════════════════════════════ */

/* ── Scroll-reveal cards ── */
.reveal{opacity:0;transform:translateY(28px);transition:opacity .55s cubic-bezier(.22,1,.36,1),transform .55s cubic-bezier(.22,1,.36,1)}
.reveal.visible{opacity:1;transform:translateY(0)}
.reveal-delay-1{transition-delay:.08s}
.reveal-delay-2{transition-delay:.16s}
.reveal-delay-3{transition-delay:.24s}
.reveal-delay-4{transition-delay:.32s}

/* ── Result panel slide-in ── */
@keyframes slideUp{from{opacity:0;transform:translateY(32px)}to{opacity:1;transform:translateY(0)}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
.result-panel.show{animation:slideUp .5s cubic-bezier(.22,1,.36,1) forwards}
.result-row{animation:fadeIn .3s ease both}

/* ── Calculate button states ── */
.btn-calc{position:relative;overflow:hidden}
.btn-calc .btn-text{transition:opacity .2s}
.btn-calc .btn-spinner{position:absolute;display:none;width:18px;height:18px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.btn-calc.loading .btn-text{opacity:0}
.btn-calc.loading .btn-spinner{display:block}
.btn-calc .ripple{position:absolute;border-radius:50%;background:rgba(255,255,255,.35);transform:scale(0);animation:rippleAnim .55s linear;pointer-events:none}
@keyframes rippleAnim{to{transform:scale(4);opacity:0}}

/* ── Progress bar ── */
#calcProgress{position:fixed;top:0;left:0;width:0;height:3px;background:linear-gradient(90deg,var(--brand),#60A5FA,var(--green));z-index:9999;transition:width .35s ease;border-radius:0 3px 3px 0;box-shadow:0 0 8px rgba(37,99,235,.5)}

/* ── Number counter ── */
.count-anim{display:inline-block;transition:transform .1s}

/* ── Tax Donut Chart ── */
#taxChartWrap{margin-top:20px;padding:16px;background:#F9FAFB;border-radius:12px;border:1px solid var(--border)}
#taxChartWrap h3{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:14px;text-align:center}
.donut-container{display:flex;align-items:center;gap:20px;flex-wrap:wrap;justify-content:center}
.donut-svg{flex-shrink:0;filter:drop-shadow(0 4px 12px rgba(0,0,0,.08))}
.donut-legend{display:flex;flex-direction:column;gap:8px;min-width:140px}
.donut-legend-item{display:flex;align-items:center;gap:8px;font-size:12px}
.donut-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.donut-label{color:var(--muted);font-weight:500}
.donut-val{font-weight:700;color:var(--ink);margin-left:auto}
.donut-segment{transition:stroke-dasharray .8s cubic-bezier(.22,1,.36,1),stroke-dashoffset .8s cubic-bezier(.22,1,.36,1)}

/* ── Regime bar chart ── */
#regimeChartWrap{margin-top:16px;padding:16px;background:#F9FAFB;border-radius:12px;border:1px solid var(--border)}
#regimeChartWrap h3{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:14px;text-align:center}
.bar-chart{display:flex;flex-direction:column;gap:10px}
.bar-row{display:flex;align-items:center;gap:10px;font-size:12px}
.bar-label{width:90px;font-weight:600;color:var(--muted);font-size:11px;text-align:right;flex-shrink:0}
.bar-track{flex:1;height:22px;background:#E5E7EB;border-radius:6px;overflow:hidden;position:relative}
.bar-fill{height:100%;border-radius:6px;width:0;transition:width 1s cubic-bezier(.22,1,.36,1);display:flex;align-items:center;justify-content:flex-end;padding-right:8px}
.bar-fill span{font-size:10px;font-weight:700;color:#fff;white-space:nowrap}
.bar-val{width:90px;font-weight:700;font-size:11px;color:var(--ink);flex-shrink:0}

/* ── Confetti ── */
.confetti-piece{position:fixed;width:8px;height:8px;top:-10px;border-radius:2px;pointer-events:none;z-index:9998;animation:confettiFall linear forwards}
@keyframes confettiFall{0%{transform:translateY(0) rotate(0deg);opacity:1}100%{transform:translateY(110vh) rotate(720deg);opacity:0}}

/* ── Assessee button pop ── */
.at-btn{transition:all .18s cubic-bezier(.34,1.56,.64,1)}
.at-btn.active{transform:scale(1.05)}
.at-btn:active{transform:scale(.95)}

/* ── Input focus ring glow ── */
input[type=number]:focus{border-color:var(--brand);box-shadow:0 0 0 3px rgba(37,99,235,.12)}
select:focus{border-color:var(--brand);box-shadow:0 0 0 3px rgba(37,99,235,.12);outline:none}

/* ── Card hover lift ── */
.card{transition:box-shadow .25s,transform .25s}
.card:hover{box-shadow:0 8px 32px rgba(0,0,0,.10);transform:translateY(-2px)}

/* ── Year pill bounce ── */
.year-pill{transition:all .2s cubic-bezier(.34,1.56,.64,1)}
.year-pill.active{transform:scale(1.06)}

/* ── Regime btn pop ── */
.regime-btn{transition:all .2s cubic-bezier(.34,1.56,.64,1)}
.regime-btn.active{transform:scale(1.04)}

/* ── Result row stagger ── */
@keyframes rowIn{from{opacity:0;transform:translateX(-10px)}to{opacity:1;transform:translateX(0)}}
.result-row{animation:rowIn .3s ease both}

/* ── Total row pulse ── */
@keyframes totalPulse{0%{transform:scale(1)}50%{transform:scale(1.02)}100%{transform:scale(1)}}
.result-row.total{animation:rowIn .4s ease both, totalPulse .4s ease .5s}

/* ── Toast slide & bounce ── */
@keyframes toastIn{0%{transform:translateY(80px) scale(.9)}70%{transform:translateY(-4px) scale(1.02)}100%{transform:translateY(0) scale(1)}}
.toast.show{animation:toastIn .4s cubic-bezier(.34,1.56,.64,1) forwards}

/* ── Hero text shimmer on load ── */
@keyframes shimmer{0%{background-position:200% center}100%{background-position:-200% center}}
.hero-shimmer{background:linear-gradient(90deg,var(--brand) 0%,#60A5FA 40%,var(--brand) 80%);background-size:200% auto;-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;animation:shimmer 3s linear infinite}

/* ── Nav scroll shadow ── */
nav{transition:box-shadow .3s}
nav.scrolled{box-shadow:0 4px 24px rgba(0,0,0,.10)}

@media print{nav,footer,.hero,.regime-toggle,.card:first-child,.btn-calc,.btn-reset,.print-btn,.toast,.year-pills,#calcProgress{display:none!important}
             .result-panel{display:block!important}.calc-grid{display:block!important}
             .card{box-shadow:none!important;border:1px solid #ccc!important}}
</style></head><body>

<div id="calcProgress"></div>
<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <ul class="nav-links">
    <li><a href="#input">Calculator</a></li>
    <li><a href="#result-section">Results</a></li>
  </ul>
  <div class="nav-right">
    <span class="nav-user">👤 <strong>{{ username }}</strong>
      <span class="badge b-{{ plan }}">{{ plan_label }}</span>
      {% if is_admin %}<span class="badge" style="background:#EFF6FF;color:var(--brand);margin-left:4px">Admin</span>{% endif %}
    </span>
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin</a>{% endif %}
    <a href="/" class="nav-btn" style="background:#F3F4F6;color:var(--ink)">← Dashboard</a>
    <a href="/logout" class="nav-link">Sign out</a>
  </div>
</nav>

<section class="hero">
  <div class="hero-badge">🧮 Multi-Year · PY 2023-24 to PY 2026-27</div>
  <h1>Income Tax <em class="hero-shimmer">Calculator</em></h1>
  <p>Calculate tax under Old &amp; New Regime for any year from PY 2023-24 to PY 2026-27. Income under 5 heads, deductions, TDS/TCS — instant comparison with slab-wise breakup.</p>
</section>

<div class="main-wrap">

<!-- Regime Toggle -->
<div class="regime-toggle">
  <button class="regime-btn active" onclick="setRegime('new')" id="btn-new">🆕 New Regime (Default)</button>
  <button class="regime-btn" onclick="setRegime('old')" id="btn-old">📜 Old Regime</button>
  <button class="regime-btn" onclick="setRegime('both')" id="btn-both">⚖️ Compare Both</button>
</div>

<div class="calc-grid" id="input">
  <!-- LEFT: Input Section -->
  <div>
    <!-- ──── BASIC INFO ──── -->
    <div class="card reveal reveal-delay-1">
      <div class="card-head">
        <div class="icon" style="background:#EFF6FF">👤</div>
        <div><h2>Basic Information</h2><p>Assessee details &amp; Assessment Year</p></div>
      </div>
      <div class="card-body">
        <div class="field">
          <label>Assessment Year</label>
          <div class="year-pills" id="yearPills">
            <button class="year-pill" onclick="setYear('2023-24')">PY 2023-24<br/><span style="font-size:10px;font-weight:400;color:var(--muted)">AY 2024-25</span></button>
            <button class="year-pill" onclick="setYear('2024-25')">PY 2024-25<br/><span style="font-size:10px;font-weight:400;color:var(--muted)">AY 2025-26</span></button>
            <button class="year-pill active" onclick="setYear('2025-26')">PY 2025-26<br/><span style="font-size:10px;font-weight:400;color:var(--muted)">AY 2026-27</span></button>
            <button class="year-pill" onclick="setYear('2026-27')">PY 2026-27<br/><span style="font-size:10px;font-weight:400;color:var(--muted)">AY 2027-28</span><span class="future-tag">Upcoming</span></button>
          </div>
          <div id="futureYearNote" style="display:none;margin-top:6px;padding:8px 12px;background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;font-size:11px;color:#92400E;font-weight:500">
            ⚠️ PY 2026-27 rates are based on Union Budget 2026 (no changes from PY 2025-26). Final rates subject to any future amendments.
          </div>
        </div>
        <div class="field">
          <label>Assessee Name <span style="font-weight:400;text-transform:none">(optional)</span></label>
          <input type="text" id="assesseeName" placeholder="e.g. Rajesh Kumar / ABC Pvt Ltd" style="border:1.5px solid var(--border);border-radius:8px;padding:9px 12px;font-family:inherit;font-size:13px;width:100%"/>
        </div>

        <!-- ── TYPE OF ASSESSEE — visible button grid ── -->
        <div class="field">
          <label>Type of Assessee</label>
          <!-- hidden select keeps existing JS logic intact -->
          <select id="assesseeType" onchange="onAssesseeTypeChange()" style="display:none">
            <option value="individual_below60" selected>Individual – Below 60 yrs</option>
            <option value="individual_senior">Individual – Senior Citizen (60–80)</option>
            <option value="individual_supersenior">Individual – Super Senior Citizen (80+)</option>
            <option value="individual_nri">Individual – Non-Resident (NRI)</option>
            <option value="huf">HUF (Hindu Undivided Family)</option>
            <option value="firm">Partnership Firm / LLP</option>
            <option value="company_domestic">Domestic Company</option>
            <option value="company_foreign">Foreign Company</option>
            <option value="company_mfg_new">Domestic Company – New Mfg (Sec 115BAB)</option>
            <option value="company_small">Domestic Company – Small (≤ ₹400 Cr, Sec 115BA)</option>
            <option value="aop_boi">AOP / BOI</option>
            <option value="cooperative">Co-operative Society</option>
            <option value="trust_aop">Trust / AOP (Registered)</option>
            <option value="local_authority">Local Authority</option>
            <option value="artificial_person">Artificial Juridical Person</option>
          </select>

          <!-- Group: Individuals -->
          <div class="at-group-label">👤 Individual</div>
          <div class="at-btn-row">
            <button class="at-btn active" data-val="individual_below60" onclick="selectAssessee(this,'individual_below60')">Below 60 yrs</button>
            <button class="at-btn" data-val="individual_senior" onclick="selectAssessee(this,'individual_senior')">Senior (60–80)</button>
            <button class="at-btn" data-val="individual_supersenior" onclick="selectAssessee(this,'individual_supersenior')">Super Senior (80+)</button>
            <button class="at-btn" data-val="individual_nri" onclick="selectAssessee(this,'individual_nri')">NRI</button>
          </div>

          <!-- Group: HUF -->
          <div class="at-group-label">🏠 HUF</div>
          <div class="at-btn-row">
            <button class="at-btn" data-val="huf" onclick="selectAssessee(this,'huf')">HUF (Hindu Undivided Family)</button>
          </div>

          <!-- Group: Firm / LLP -->
          <div class="at-group-label">🤝 Firm / LLP</div>
          <div class="at-btn-row">
            <button class="at-btn" data-val="firm" onclick="selectAssessee(this,'firm')">Partnership Firm / LLP</button>
          </div>

          <!-- Group: Companies -->
          <div class="at-group-label">🏢 Company</div>
          <div class="at-btn-row">
            <button class="at-btn" data-val="company_domestic" onclick="selectAssessee(this,'company_domestic')">Domestic Co.</button>
            <button class="at-btn" data-val="company_foreign" onclick="selectAssessee(this,'company_foreign')">Foreign Co.</button>
            <button class="at-btn" data-val="company_mfg_new" onclick="selectAssessee(this,'company_mfg_new')">New Mfg Co. (115BAB)</button>
            <button class="at-btn" data-val="company_small" onclick="selectAssessee(this,'company_small')">Small Co. (115BA)</button>
          </div>

          <!-- Group: Others -->
          <div class="at-group-label">🏛️ Others</div>
          <div class="at-btn-row">
            <button class="at-btn" data-val="aop_boi" onclick="selectAssessee(this,'aop_boi')">AOP / BOI</button>
            <button class="at-btn" data-val="cooperative" onclick="selectAssessee(this,'cooperative')">Co-operative Society</button>
            <button class="at-btn" data-val="trust_aop" onclick="selectAssessee(this,'trust_aop')">Trust</button>
            <button class="at-btn" data-val="local_authority" onclick="selectAssessee(this,'local_authority')">Local Authority</button>
            <button class="at-btn" data-val="artificial_person" onclick="selectAssessee(this,'artificial_person')">Artificial Juridical Person</button>
          </div>
        </div>

        <!-- Individual-only fields -->
        <div id="individualFields">
          <div class="row2">
            <div class="field">
              <label>Age Category</label>
              <div id="ageBadge" style="padding:8px 14px;background:#EFF6FF;border-radius:8px;font-size:12px;font-weight:600;color:#1E40AF;display:inline-block">Below 60 years</div>
            </div>
            <div class="field">
              <label>Residential Status</label>
              <select id="residentialStatus">
                <option value="resident">Resident</option>
                <option value="nri">Non-Resident</option>
              </select>
            </div>
          </div>
        </div>

        <!-- Company-specific info -->
        <div id="companyFields" style="display:none">
          <div id="matInfo" style="padding:10px 14px;background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;font-size:11px;color:#1E40AF;margin-top:4px">
            <strong>ℹ️ MAT u/s 115JB:</strong> Tax payable is higher of normal tax or 15% of Book Profit (+ surcharge + cess). Enter Book Profit in the MAT section below.
          </div>
        </div>
        <!-- Firm info -->
        <div id="firmFields" style="display:none">
          <div style="padding:10px 14px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;font-size:11px;color:#065F46;margin-top:4px">
            <strong>ℹ️ Firm / LLP:</strong> Flat 30% on total income + surcharge (12% if income &gt; ₹1 Cr) + cess 4%. AMT @ 18.5% of Adjusted Total Income applies u/s 115JC.
          </div>
        </div>
        <!-- AOP / Co-op info -->
        <div id="aopFields" style="display:none">
          <div style="padding:10px 14px;background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;font-size:11px;color:#92400E;margin-top:4px">
            <strong>ℹ️ AOP / BOI / Co-op / Trust:</strong> Taxed at applicable slab rates as per specific provisions. AMT @ 18.5% may apply u/s 115JC.
          </div>
        </div>
      </div>
    </div>

    <!-- ──── INCOME HEADS (Dynamic per assessee) ──── -->
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#F0FDF4">💰</div>
        <div><h2>Income Details</h2><p id="incomeCardSub">Gross Total Income computation</p></div>
      </div>
      <div class="card-body">

        <!-- ═══ INDIVIDUAL / HUF / AOP / TRUST (all 5 heads) ═══ -->
        <div id="inc_individual">
          <div class="section-title">📋 1. Income from Salary</div>
          <div class="row2">
            <div class="field">
              <label>Gross Salary</label>
              <input type="number" id="grossSalary" placeholder="0" min="0"/>
            </div>
            <div class="field">
              <label>Exempt Allowances (HRA, LTA etc.)</label>
              <input type="number" id="exemptAllow" placeholder="0" min="0"/>
              <p class="hint">Old regime: HRA, LTA etc. New regime: mostly nil</p>
            </div>
          </div>
          <div class="field">
            <label>Standard Deduction</label>
            <input type="number" id="stdDeduction" value="75000" placeholder="75000" min="0"/>
            <p class="hint" id="stdDedHint">₹75,000 for PY 2024-25 onwards. ₹50,000 for PY 2023-24.</p>
          </div>

          <div class="section-title">🏠 2. Income from House Property</div>
          <div class="row2">
            <div class="field">
              <label>Net Annual Value / Rental Income</label>
              <input type="number" id="houseIncome" placeholder="0"/>
              <p class="hint">Can be negative for self-occupied (loss)</p>
            </div>
            <div class="field">
              <label>Interest on Home Loan (Sec 24b)</label>
              <input type="number" id="homeLoanInterest" placeholder="0" min="0"/>
              <p class="hint">Max ₹2L self-occupied</p>
            </div>
          </div>

          <div class="section-title">💼 3. Profits from Business / Profession</div>
          <div class="field">
            <label>Net Profit from Business / Profession</label>
            <input type="number" id="businessIncome" placeholder="0"/>
            <p class="hint">After all business deductions. Can be loss (negative)</p>
          </div>

          <div class="section-title">📈 4. Capital Gains</div>
          <div id="transitionalNotice" style="display:none;margin-bottom:12px;padding:10px 14px;background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;font-size:11px;color:#92400E;line-height:1.6">
            <strong>⚠️ PY 2024-25 Transitional Year:</strong> Budget 2024 changed capital gains rates from 23 July 2024. Enter gains sold <strong>before 23 July</strong> separately from gains sold <strong>on/after 23 July</strong>.
          </div>
          <div id="preJulyFields" style="display:none">
            <div style="font-size:11px;font-weight:700;color:#92400E;text-transform:uppercase;letter-spacing:.05em;margin:8px 0 8px;padding:4px 10px;background:#FFFBEB;border-radius:6px;display:inline-block">📅 Pre-23 July 2024 (Old Rates)</div>
            <div class="row2">
              <div class="field">
                <label>STCG 111A — Pre-July (equity) @ 15%</label>
                <input type="number" id="stcg111aPreJuly" placeholder="0" min="0"/>
              </div>
              <div class="field">
                <label>LTCG 112A — Pre-July (equity) @ 10%</label>
                <input type="number" id="ltcg112aPreJuly" placeholder="0" min="0"/>
                <p class="hint">Exempt up to ₹1 lakh (old limit)</p>
              </div>
            </div>
            <div class="row2">
              <div class="field">
                <label>LTCG Other — Pre-July @ 20%</label>
                <input type="number" id="ltcgOtherPreJuly" placeholder="0" min="0"/>
                <p class="hint">With indexation benefit (pre-July rule)</p>
              </div>
              <div class="field"></div>
            </div>
            <div style="font-size:11px;font-weight:700;color:var(--brand);text-transform:uppercase;letter-spacing:.05em;margin:12px 0 8px;padding:4px 10px;background:#EFF6FF;border-radius:6px;display:inline-block">📅 Post-23 July 2024 (New Rates)</div>
          </div>
          <div class="row2">
            <div class="field">
              <label id="stcg111aLabel">STCG u/s 111A (equity, STT paid)</label>
              <input type="number" id="stcg111a" placeholder="0" min="0"/>
              <p class="hint" id="stcgRateHint">Rate varies by year</p>
            </div>
            <div class="field">
              <label>STCG — Other (non-equity)</label>
              <input type="number" id="stcgOther" placeholder="0" min="0"/>
              <p class="hint">Taxed at slab rates</p>
            </div>
          </div>
          <div class="row2">
            <div class="field">
              <label id="ltcg112aLabel">LTCG u/s 112A (equity, STT paid)</label>
              <input type="number" id="ltcg112a" placeholder="0" min="0"/>
              <p class="hint" id="ltcgRateHint">Rate &amp; exemption varies by year</p>
            </div>
            <div class="field">
              <label id="ltcgOtherLabel">LTCG — Other (property, debt etc.)</label>
              <input type="number" id="ltcgOther" placeholder="0" min="0"/>
              <p class="hint" id="ltcgOtherHint">Rate varies by year</p>
            </div>
          </div>

          <div class="section-title">📦 5. Income from Other Sources</div>
          <div class="row2">
            <div class="field">
              <label>Interest / Dividends / Other Income</label>
              <input type="number" id="otherIncome" placeholder="0" min="0"/>
            </div>
            <div class="field">
              <label>Winnings (lottery, games etc.)</label>
              <input type="number" id="winningsIncome" placeholder="0" min="0"/>
              <p class="hint">Taxed at 30% flat</p>
            </div>
          </div>
        </div>

        <!-- ═══ COMPANY ═══ -->
        <div id="inc_company" style="display:none">
          <div style="margin-bottom:14px;padding:10px 14px;background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;font-size:11px;color:#1E40AF">
            <strong>ℹ️ Company Income:</strong> Enter net taxable income computed as per IT Act provisions (after all allowable business deductions, depreciation, etc.).
          </div>

          <div class="section-title">💼 Business / Profession Income</div>
          <div class="field">
            <label>Net Taxable Business Income (₹)</label>
            <input type="number" id="co_businessIncome" placeholder="0"/>
            <p class="hint">Net profit after all IT Act allowable deductions &amp; depreciation</p>
          </div>

          <div class="section-title">📈 Capital Gains</div>
          <div class="row2">
            <div class="field">
              <label>STCG u/s 111A (equity, STT paid)</label>
              <input type="number" id="co_stcg111a" placeholder="0" min="0"/>
            </div>
            <div class="field">
              <label>STCG — Other assets</label>
              <input type="number" id="co_stcgOther" placeholder="0" min="0"/>
            </div>
          </div>
          <div class="row2">
            <div class="field">
              <label>LTCG u/s 112A (equity)</label>
              <input type="number" id="co_ltcg112a" placeholder="0" min="0"/>
            </div>
            <div class="field">
              <label>LTCG — Other assets</label>
              <input type="number" id="co_ltcgOther" placeholder="0" min="0"/>
            </div>
          </div>

          <div class="section-title">📦 Other Sources</div>
          <div class="field">
            <label>Interest / Dividend / Other Income (₹)</label>
            <input type="number" id="co_otherIncome" placeholder="0" min="0"/>
          </div>

          <div class="section-title">⚖️ MAT — Book Profit (Sec 115JB)</div>
          <div class="row2">
            <div class="field">
              <label>Book Profit u/s 115JB (₹)</label>
              <input type="number" id="co_bookProfit" placeholder="0" min="0" oninput="syncMatBookProfit()"/>
              <p class="hint">Net profit per P&amp;L + mandatory add-backs as per Sch VII</p>
            </div>
            <div class="field">
              <label>Turnover / Gross Receipts (₹)</label>
              <input type="number" id="co_turnover" placeholder="0" min="0"/>
              <p class="hint">For reference / threshold checks</p>
            </div>
          </div>
        </div>

        <!-- ═══ FIRM / LLP ═══ -->
        <div id="inc_firm" style="display:none">
          <div style="margin-bottom:14px;padding:10px 14px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;font-size:11px;color:#065F46">
            <strong>ℹ️ Firm / LLP Income:</strong> Firms are not allowed salary to partners for tax purposes beyond limits. Enter total firm income after all deductions. Remuneration &amp; interest to partners allowed u/s 40(b) already deducted.
          </div>

          <div class="section-title">💼 Business / Profession Income</div>
          <div class="field">
            <label>Net Taxable Business Income (₹)</label>
            <input type="number" id="firm_businessIncome" placeholder="0"/>
            <p class="hint">Firm profit after partner remuneration / interest u/s 40(b)</p>
          </div>

          <div class="section-title">📈 Capital Gains</div>
          <div class="row2">
            <div class="field">
              <label>STCG (equity u/s 111A)</label>
              <input type="number" id="firm_stcg111a" placeholder="0" min="0"/>
            </div>
            <div class="field">
              <label>LTCG (equity u/s 112A)</label>
              <input type="number" id="firm_ltcg112a" placeholder="0" min="0"/>
            </div>
          </div>
          <div class="row2">
            <div class="field">
              <label>STCG — Other</label>
              <input type="number" id="firm_stcgOther" placeholder="0" min="0"/>
            </div>
            <div class="field">
              <label>LTCG — Other</label>
              <input type="number" id="firm_ltcgOther" placeholder="0" min="0"/>
            </div>
          </div>

          <div class="section-title">📦 Other Sources</div>
          <div class="field">
            <label>Interest / Other Income (₹)</label>
            <input type="number" id="firm_otherIncome" placeholder="0" min="0"/>
          </div>

          <div class="section-title">⚖️ AMT — Adjusted Total Income (Sec 115JC)</div>
          <div class="field">
            <label>Adjusted Total Income for AMT (₹)</label>
            <input type="number" id="firm_amtAti" placeholder="0" min="0" oninput="syncAmtAti()"/>
            <p class="hint">GTI + deductions claimed u/s 10AA / 35AD / 80H–80RRB added back</p>
          </div>
        </div>

        <!-- ═══ CO-OPERATIVE SOCIETY ═══ -->
        <div id="inc_coop" style="display:none">
          <div style="margin-bottom:14px;padding:10px 14px;background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;font-size:11px;color:#92400E">
            <strong>ℹ️ Co-operative Society:</strong> Taxed on slab — 10% up to ₹10K / 20% up to ₹20K / 30% above. Surcharge 12% if income &gt; ₹1 Cr. Cess 4%.
          </div>
          <div class="section-title">💼 Business Income</div>
          <div class="field">
            <label>Net Taxable Income (₹)</label>
            <input type="number" id="coop_businessIncome" placeholder="0"/>
          </div>
          <div class="section-title">📦 Other Sources</div>
          <div class="field">
            <label>Interest / Dividend / Other (₹)</label>
            <input type="number" id="coop_otherIncome" placeholder="0" min="0"/>
          </div>
        </div>

        <!-- ═══ LOCAL AUTHORITY ═══ -->
        <div id="inc_local" style="display:none">
          <div style="margin-bottom:14px;padding:10px 14px;background:#F5F3FF;border:1px solid #DDD6FE;border-radius:8px;font-size:11px;color:#4C1D95">
            <strong>ℹ️ Local Authority:</strong> Flat 30% on total income + 4% cess. No surcharge.
          </div>
          <div class="section-title">💼 Business / Property / Other Income</div>
          <div class="field">
            <label>Net Taxable Income (₹)</label>
            <input type="number" id="local_income" placeholder="0"/>
          </div>
        </div>

      </div>
    </div>

    <!-- ──── DEDUCTIONS (OLD REGIME) ──── -->
    <div class="card" id="deductions-card">
      <div class="card-head">
        <div class="icon" style="background:#FFFBEB">🧾</div>
        <div><h2>Deductions (Chapter VI-A)</h2><p>Applicable in Old Regime only (except 80CCD(2))</p></div>
      </div>
      <div class="card-body">
        <div class="row2">
          <div class="field">
            <label>80C (PPF, LIC, ELSS, etc.)</label>
            <input type="number" id="ded80c" placeholder="0" min="0" max="150000"/>
            <p class="hint">Max ₹1,50,000</p>
          </div>
          <div class="field">
            <label>80CCD(1B) — NPS Extra</label>
            <input type="number" id="ded80ccd1b" placeholder="0" min="0" max="50000"/>
            <p class="hint">Max ₹50,000</p>
          </div>
        </div>
        <div class="row2">
          <div class="field">
            <label>80CCD(2) — Employer NPS</label>
            <input type="number" id="ded80ccd2" placeholder="0" min="0"/>
            <p class="hint">Available in both regimes</p>
          </div>
          <div class="field">
            <label>80D — Medical Insurance</label>
            <input type="number" id="ded80d" placeholder="0" min="0"/>
            <p class="hint">₹25K self + ₹25K/₹50K parents</p>
          </div>
        </div>
        <div class="row2">
          <div class="field">
            <label>80E — Education Loan Interest</label>
            <input type="number" id="ded80e" placeholder="0" min="0"/>
          </div>
          <div class="field">
            <label>80G — Donations</label>
            <input type="number" id="ded80g" placeholder="0" min="0"/>
          </div>
        </div>
        <div class="row2">
          <div class="field">
            <label>80TTA/80TTB — Savings Interest</label>
            <input type="number" id="ded80tta" placeholder="0" min="0"/>
            <p class="hint">₹10K (80TTA) / ₹50K seniors (80TTB)</p>
          </div>
          <div class="field">
            <label>Other Deductions (80DD, 80DDB, etc.)</label>
            <input type="number" id="dedOther" placeholder="0" min="0"/>
          </div>
        </div>
      </div>
    </div>

    <!-- ──── MAT / AMT ──── -->
    <div class="card" id="matAmtCard" style="display:none">
      <div class="card-head">
        <div class="icon" style="background:#FEF3C7">⚖️</div>
        <div><h2 id="matAmtCardTitle">MAT / AMT Computation</h2><p id="matAmtCardSub">Minimum Alternate Tax u/s 115JB / 115JC</p></div>
      </div>
      <div class="card-body">
        <!-- MAT (Companies) -->
        <div id="matSection">
          <div class="section-title">📋 MAT u/s 115JB — Companies</div>
          <div class="row2">
            <div class="field">
              <label>Book Profit u/s 115JB (₹)</label>
              <input type="number" id="matBookProfit" placeholder="0" min="0" oninput="computeMatAmt()"/>
              <p class="hint">Net profit per P&amp;L + mandatory add-backs (Sch VII items)</p>
            </div>
            <div class="field">
              <label>MAT Rate</label>
              <select id="matRate" onchange="computeMatAmt()">
                <option value="0.15" selected>15% – Domestic Company (general)</option>
                <option value="0.09">9% – New Mfg Co. u/s 115BAB</option>
                <option value="0.075">7.5% – Co. in IFSC u/s 115A(4)</option>
              </select>
            </div>
          </div>
          <div id="matResult" style="display:none;margin-top:12px;padding:14px;background:#FFFBEB;border:1.5px solid #FDE68A;border-radius:10px">
            <div style="font-size:12px;font-weight:700;color:#92400E;margin-bottom:8px">MAT Computation</div>
            <div id="matResultRows"></div>
          </div>
        </div>
        <!-- AMT (Non-companies) -->
        <div id="amtSection" style="display:none">
          <div class="section-title">📋 AMT u/s 115JC — Firms / LLPs / Individuals / HUF / AOP</div>
          <div class="field">
            <label>Adjusted Total Income (ATI) for AMT (₹)</label>
            <input type="number" id="amtAti" placeholder="0" min="0" oninput="computeMatAmt()"/>
            <p class="hint">GTI + deductions claimed u/s 10AA / 35AD / 80H–80RRB added back (u/s 115JC)</p>
          </div>
          <div id="amtResult" style="display:none;margin-top:12px;padding:14px;background:#FFFBEB;border:1.5px solid #FDE68A;border-radius:10px">
            <div style="font-size:12px;font-weight:700;color:#92400E;margin-bottom:8px">AMT Computation</div>
            <div id="amtResultRows"></div>
          </div>
        </div>
        <div style="margin-top:12px;padding:10px 14px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;font-size:11px;color:#065F46">
          <strong>📌 Credit:</strong> If normal tax &gt; MAT/AMT, MAT/AMT credit u/s 115JAA/115JD is carried forward for up to 15 years and can be set off in future years when normal tax exceeds MAT/AMT.
        </div>
      </div>
    </div>

    <!-- ──── TDS / TCS / ADVANCE TAX ──── -->
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#F5F3FF">🏦</div>
        <div><h2>Tax Already Paid</h2><p>TDS, TCS &amp; Advance Tax</p></div>
      </div>
      <div class="card-body">
        <div class="row3">
          <div class="field">
            <label>TDS (Estimated)</label>
            <input type="number" id="tds" placeholder="0" min="0"/>
          </div>
          <div class="field">
            <label>TCS (Estimated)</label>
            <input type="number" id="tcs" placeholder="0" min="0"/>
          </div>
          <div class="field">
            <label>Advance Tax Paid</label>
            <input type="number" id="advanceTax" placeholder="0" min="0"/>
          </div>
        </div>
      </div>
    </div>

    <button class="btn-calc" id="calcBtn" onclick="calculateTax(event)">
      <span class="btn-text">🧮 Calculate Tax</span>
      <div class="btn-spinner"></div>
    </button>
    <button class="btn-reset" onclick="resetForm()">↺ Reset All Fields</button>
  </div>

  <!-- RIGHT: Results Section -->
  <div id="result-section">
    <div class="result-panel" id="resultPanel">

      <div class="card" id="singleResult" style="display:none">
        <div class="card-head">
          <div class="icon" style="background:#ECFDF5">📊</div>
          <div><h2 id="resultTitle">Tax Computation</h2><p id="resultSubtitle"></p></div>
        </div>
        <div class="card-body">
          <div id="resultBody"></div>
          <!-- Donut chart -->
          <div id="taxChartWrap" style="display:none">
            <h3>Tax Breakdown</h3>
            <div class="donut-container">
              <svg class="donut-svg" width="140" height="140" viewBox="0 0 140 140">
                <circle cx="70" cy="70" r="54" fill="none" stroke="#F3F4F6" stroke-width="22"/>
                <circle id="donut-base" class="donut-segment" cx="70" cy="70" r="54" fill="none" stroke="#2563EB" stroke-width="22" stroke-dasharray="0 339.3" stroke-dashoffset="84.8" stroke-linecap="round"/>
                <circle id="donut-surcharge" class="donut-segment" cx="70" cy="70" r="54" fill="none" stroke="#F59E0B" stroke-width="22" stroke-dasharray="0 339.3" stroke-dashoffset="84.8" stroke-linecap="round"/>
                <circle id="donut-cess" class="donut-segment" cx="70" cy="70" r="54" fill="none" stroke="#10B981" stroke-width="22" stroke-dasharray="0 339.3" stroke-dashoffset="84.8" stroke-linecap="round"/>
                <text x="70" y="65" text-anchor="middle" font-size="10" fill="#6B7280" font-weight="600" font-family="Inter,sans-serif">Total Tax</text>
                <text id="donut-center-val" x="70" y="82" text-anchor="middle" font-size="13" fill="#111827" font-weight="800" font-family="Inter,sans-serif">₹0</text>
              </svg>
              <div class="donut-legend" id="donutLegend"></div>
            </div>
          </div>
          <button class="print-btn" onclick="window.print()">🖨️ Print / Save PDF</button>
        </div>
      </div>

      <div id="compareResult" style="display:none">
        <div class="compare-box">
          <h3>⚖️ Regime Comparison — <span id="compareYearLabel"></span></h3>
          <div class="regime-winner" id="regimeWinner"></div>
          <div class="savings" id="savingsAmt"></div>
          <table class="compare-table">
            <thead><tr><th></th><th>🆕 New Regime</th><th>📜 Old Regime</th></tr></thead>
            <tbody id="compareBody"></tbody>
          </table>
        </div>

        <!-- Regime bar chart -->
        <div id="regimeChartWrap">
          <h3>📊 Visual Comparison</h3>
          <div class="bar-chart" id="regimeBarChart"></div>
        </div>

        <div class="card" style="margin-top:16px">
          <div class="card-head">
            <div class="icon" style="background:#EFF6FF">📊</div>
            <div><h2>New Regime — Detailed</h2></div>
          </div>
          <div class="card-body"><div id="newRegimeDetail"></div></div>
        </div>

        <div class="card" style="margin-top:16px">
          <div class="card-head">
            <div class="icon" style="background:#FFFBEB">📊</div>
            <div><h2>Old Regime — Detailed</h2></div>
          </div>
          <div class="card-body"><div id="oldRegimeDetail"></div></div>
        </div>

        <button class="print-btn" onclick="window.print()">🖨️ Print / Save PDF</button>
      </div>

      <div class="card" style="margin-top:16px" id="slabCard">
        <div class="card-head">
          <div class="icon" style="background:#F0FDF4">📋</div>
          <div><h2>Slab-wise Tax Breakup</h2><p id="slabRegimeLabel"></p></div>
        </div>
        <div class="card-body" id="slabBody"></div>
      </div>

      <div class="disclaimer" id="disclaimerBox">
        <strong>⚠ Disclaimer:</strong> This calculator is for estimation purposes only. Actual tax liability may differ
        based on specific exemptions, deductions, and interpretations. Always consult a qualified Chartered Accountant
        for final tax computation. Surcharge marginal relief is indicative. Special rate incomes (capital gains, winnings)
        are not eligible for Section 87A rebate.
      </div>
      <div class="disclaimer future" id="futureDisclaimer" style="display:none;margin-top:8px">
        <strong>📅 Future Year Note:</strong> PY 2026-27 (AY 2027-28) rates are based on the Union Budget 2026 which
        retained PY 2025-26 slab rates without changes. These rates are subject to any future amendments or notifications
        by the Government. Always verify with the latest Finance Act before finalizing.
      </div>

      <!-- ──── MAT/AMT RESULT IN RESULTS PANEL ──── -->
      <div id="matAmtResultCard" style="display:none;margin-top:16px">
        <div class="card">
          <div class="card-head">
            <div class="icon" style="background:#FEF3C7">⚖️</div>
            <div><h2 id="matAmtResTitle">MAT / AMT Summary</h2><p>Minimum Alternate Tax computation result</p></div>
          </div>
          <div class="card-body" id="matAmtResBody"></div>
        </div>
      </div>

      <!-- ──── ADVANCE TAX SCHEDULE (PY 2026-27 only) ──── -->
      <div id="advanceTaxCard" style="display:none;margin-top:16px">
        <div class="card">
          <div class="card-head">
            <div class="icon" style="background:#EFF6FF">📅</div>
            <div><h2>Advance Tax Schedule — PY 2026-27</h2><p>Instalment-wise liability u/s 208 &amp; 211</p></div>
          </div>
          <div class="card-body">
            <div style="padding:10px 14px;background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;font-size:11px;color:#1E40AF;margin-bottom:14px">
              <strong>ℹ️ Who must pay?</strong> Every assessee whose estimated tax liability for the year is ₹10,000 or more (after TDS/TCS) must pay advance tax. Senior citizens (60+) with no business income are exempt u/s 207.
            </div>
            <div id="advanceTaxTable"></div>
            <div style="margin-top:14px;padding:10px 14px;background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;font-size:11px;color:#92400E">
              <strong>⚠️ Interest for default:</strong> u/s 234B — 1% per month on 90% of assessed tax not paid as advance tax. u/s 234C — 1% per month for 3 months on shortfall per instalment (single month for last instalment). u/s 234A — 1% per month on self-assessment tax if return filed late.
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Before calculation: show slab reference -->
    <div id="preCalcInfo">
      <div class="card" id="refSlabCard">
        <div class="card-head">
          <div class="icon" style="background:#EFF6FF">📋</div>
          <div><h2 id="refSlabTitle">New Regime Slab Rates</h2><p id="refSlabSub"></p></div>
        </div>
        <div class="card-body" id="refSlabBody"></div>
      </div>
      <div class="card">
        <div class="card-head">
          <div class="icon" style="background:#FFFBEB">📋</div>
          <div><h2>Old Regime Slab Rates</h2><p>Unchanged across all years</p></div>
        </div>
        <div class="card-body">
          <table class="slab-table">
            <thead><tr><th>Income Slab</th><th style="text-align:right">Rate</th></tr></thead>
            <tbody>
              <tr><td>Up to ₹2,50,000</td><td class="amt">Nil</td></tr>
              <tr><td>₹2,50,001 – ₹5,00,000</td><td class="amt">5%</td></tr>
              <tr><td>₹5,00,001 – ₹10,00,000</td><td class="amt">20%</td></tr>
              <tr><td>Above ₹10,00,000</td><td class="amt">30%</td></tr>
            </tbody>
          </table>
          <p class="hint" style="margin-top:10px">Senior citizens (60-80): exempt up to ₹3L. Super seniors (80+): exempt up to ₹5L. Rebate u/s 87A: up to ₹5L → max ₹12,500.</p>
        </div>
      </div>
      <div class="card">
        <div class="card-head">
          <div class="icon" style="background:#F5F3FF">📋</div>
          <div><h2>Special Rate Incomes</h2><p id="refSpecialSub"></p></div>
        </div>
        <div class="card-body" id="refSpecialBody"></div>
      </div>
    </div>
  </div>
</div>
</div>

<footer>
  <div class="ft-main">
    <div>
      <div class="ft-brand-name">CA Toolkit</div>
      <p class="ft-brand-desc">CA Toolkit is a comprehensive utility platform built by a CA Article from Ludhiana, Punjab, providing automation tools for Indian Chartered Accountants. The platform saves hours of manual work every year — from Balance Sheet year-shift to GST reconciliation, tax calculations, and more.</p>
    </div>
    <div>
      <div class="ft-col-title">Know More</div>
      <ul class="ft-links">
        <li><a href="/">Home</a></li>
        <li><a href="/">BS Year Shift</a></li>
        <li><a href="/tool/tb-to-bs">TB → Balance Sheet</a></li>
        <li><a href="/tool/tax-calculator">Tax Calculator</a></li>
        <li><a href="/privacy">Privacy Policy</a></li>
      </ul>
    </div>
    <div>
      <div class="ft-col-title">Contact Us</div>
      <div class="ft-contact-name">CA Toolkit</div>
      <div class="ft-contact-addr">Built for Indian Chartered Accountants<br/>Created by CA Article · Ludhiana, Punjab</div>
    </div>
  </div>
  <div class="ft-bottom">
    <span class="ft-bottom-left">©2026 CA Toolkit · All Rights Reserved</span>
    <span class="ft-bottom-right">Built for Indian CAs · Ludhiana, Punjab</span>
  </div>
</footer>
<div class="toast" id="toast"></div>

<script>
/* ═══════════════════════════════════════════════════════════════════════
   INCOME TAX CALCULATOR — Multi-Year Engine (Enhanced)
   PY 2023-24 | PY 2024-25 | PY 2025-26 | PY 2026-27 (upcoming)
   Features: All Assessee Types · MAT/AMT · Advance Tax Schedule
   ═══════════════════════════════════════════════════════════════════════ */

let currentRegime = 'new';
let currentYear = '2025-26';

/* ── ASSESSEE TYPE CONFIGURATION ──────────────────────────────────── */
const ASSESSEE_CFG = {
  // Individuals
  individual_below60:    { label:'Individual – Below 60', group:'individual', age:'below60', canAmt:true },
  individual_senior:     { label:'Individual – Senior Citizen (60–80)', group:'individual', age:'senior', canAmt:true },
  individual_supersenior:{ label:'Individual – Super Senior (80+)', group:'individual', age:'supersenior', canAmt:true },
  individual_nri:        { label:'Individual – NRI', group:'individual', age:'below60', isNRI:true, canAmt:true },
  // HUF
  huf:                   { label:'HUF', group:'individual', age:'below60', canAmt:true },
  // Firms / LLPs
  firm:                  { label:'Firm / LLP', group:'firm', flatRate:0.30, surchargeThreshold:1e7, surchargeRate:0.12, canAmt:true },
  // Companies
  company_domestic:      { label:'Domestic Company', group:'company', flatRate:0.22, surchargeThreshold:1e7, surchargeRateLow:0.07, surchargeRateHigh:0.12, canMat:true },
  company_foreign:       { label:'Foreign Company', group:'company', flatRate:0.40, surchargeRateLow:0.02, surchargeRateHigh:0.05, surchargeThreshold:1e7, canMat:true },
  company_mfg_new:       { label:'New Mfg Co. u/s 115BAB', group:'company', flatRate:0.15, surchargeRate:0.10, matRate:0.09, canMat:true },
  company_small:         { label:'Domestic Co. ≤₹400Cr (Sec 115BA)', group:'company', flatRate:0.25, surchargeRateLow:0.07, surchargeRateHigh:0.12, surchargeThreshold:1e7, canMat:true },
  // Others
  aop_boi:               { label:'AOP / BOI', group:'individual', age:'below60', canAmt:true },
  cooperative:           { label:'Co-operative Society', group:'coop', canAmt:true },
  trust_aop:             { label:'Trust / AOP (Registered)', group:'individual', age:'below60', canAmt:true },
  local_authority:       { label:'Local Authority', group:'local', flatRate:0.30, canAmt:false },
  artificial_person:     { label:'Artificial Juridical Person', group:'individual', age:'below60', canAmt:true },
};

function getAssesseeCfg() {
  const t = document.getElementById('assesseeType').value;
  return ASSESSEE_CFG[t] || ASSESSEE_CFG['individual_below60'];
}

/* ── Assessee button selector ─────────────────────────────────────── */
function selectAssessee(btn, val) {
  document.querySelectorAll('.at-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('assesseeType').value = val;
  onAssesseeTypeChange();
}

function onAssesseeTypeChange() {
  const t = document.getElementById('assesseeType').value;
  const ac = ASSESSEE_CFG[t];

  const isInd  = ac.group === 'individual';
  const isCo   = ac.group === 'company';
  const isFirm = ac.group === 'firm';
  const isCoop = ac.group === 'coop';
  const isLocal= ac.group === 'local';

  // ── Income panel visibility ──────────────────────────────────────
  document.getElementById('inc_individual').style.display = (isInd) ? 'block' : 'none';
  document.getElementById('inc_company').style.display    = isCo   ? 'block' : 'none';
  document.getElementById('inc_firm').style.display       = isFirm ? 'block' : 'none';
  document.getElementById('inc_coop').style.display       = isCoop ? 'block' : 'none';
  document.getElementById('inc_local').style.display      = isLocal? 'block' : 'none';

  // Card subtitle
  const subMap = {
    individual:'All 5 Heads of Income', huf:'All 5 Heads of Income',
    firm:'Business, Capital Gains & Other Sources',
    company_domestic:'Business, Capital Gains & MAT',
    company_foreign:'Business, Capital Gains & MAT',
    company_mfg_new:'Business, Capital Gains & MAT',
    company_small:'Business, Capital Gains & MAT',
    aop_boi:'All applicable Heads of Income',
    cooperative:'Business & Other Sources',
    trust_aop:'All applicable Heads of Income',
    local_authority:'Net Taxable Income',
    artificial_person:'All 5 Heads of Income',
  };
  document.getElementById('incomeCardSub').textContent = subMap[t] || 'Gross Total Income computation';

  // ── Age badge for individuals ────────────────────────────────────
  const ageBadge = document.getElementById('ageBadge');
  if (ageBadge) {
    const ageMap = {
      individual_below60:'Below 60 years', individual_senior:'Senior Citizen (60–80)',
      individual_supersenior:'Super Senior Citizen (80+)', individual_nri:'Non-Resident (NRI)', huf:'HUF',
      aop_boi:'AOP / BOI', trust_aop:'Trust', artificial_person:'Juridical Person',
    };
    ageBadge.textContent = ageMap[t] || '';
  }

  // ── Show/hide Basic Info sub-sections ───────────────────────────
  document.getElementById('individualFields').style.display  = isInd  ? 'block' : 'none';
  document.getElementById('companyFields').style.display     = isCo   ? 'block' : 'none';
  document.getElementById('firmFields').style.display        = isFirm ? 'block' : 'none';
  document.getElementById('aopFields').style.display         = (isCoop || isLocal) ? 'block' : 'none';

  // ── Regime toggle visibility ─────────────────────────────────────
  const regimeToggle = document.querySelector('.regime-toggle');
  if (isCo || isFirm || isLocal) {
    regimeToggle.style.display = 'none';
    currentRegime = 'new';
  } else {
    regimeToggle.style.display = 'flex';
  }

  // ── Deductions card ──────────────────────────────────────────────
  const dc = document.getElementById('deductions-card');
  if (isCo || isFirm || isCoop || isLocal) {
    dc.style.display = 'none';
  } else {
    dc.style.display = 'block';
    dc.style.opacity = currentRegime === 'new' ? '0.4' : '1';
    dc.style.pointerEvents = currentRegime === 'new' ? 'none' : 'auto';
  }

  // ── MAT / AMT card ───────────────────────────────────────────────
  const matCard = document.getElementById('matAmtCard');
  if (ac.canMat) {
    matCard.style.display = 'block';
    document.getElementById('matSection').style.display = 'block';
    document.getElementById('amtSection').style.display = 'none';
    document.getElementById('matAmtCardTitle').textContent = 'MAT Computation';
    document.getElementById('matAmtCardSub').textContent = 'Minimum Alternate Tax u/s 115JB';
    const mr = document.getElementById('matRate');
    if (t === 'company_mfg_new') mr.value = '0.09';
    else mr.value = '0.15';
  } else if (ac.canAmt && (isFirm || isCoop)) {
    matCard.style.display = 'block';
    document.getElementById('matSection').style.display = 'none';
    document.getElementById('amtSection').style.display = 'block';
    document.getElementById('matAmtCardTitle').textContent = 'AMT Computation';
    document.getElementById('matAmtCardSub').textContent = 'Alternate Minimum Tax u/s 115JC';
  } else {
    matCard.style.display = 'none';
  }

  // ── NRI auto-set ─────────────────────────────────────────────────
  if (ac.isNRI) {
    const rs = document.getElementById('residentialStatus');
    if (rs) rs.value = 'nri';
  }

  computeMatAmt();
}

/* ── YEAR-SPECIFIC TAX CONFIGURATIONS ─────────────────────────────── */
const YEAR_CONFIG = {
  '2023-24': {
    label: 'PY 2023-24 (AY 2024-25)',
    ayLabel: 'AY 2024-25',
    isFuture: false,
    stdDeduction: 50000,
    newSlabs: [
      { upto: 300000,  rate: 0 },
      { upto: 600000,  rate: 0.05 },
      { upto: 900000,  rate: 0.10 },
      { upto: 1200000, rate: 0.15 },
      { upto: 1500000, rate: 0.20 },
      { upto: Infinity, rate: 0.30 },
    ],
    rebateNew: { limit: 700000, max: 25000 },
    rebateOld: { limit: 500000, max: 12500 },
    stcg111aRate: 0.15,
    ltcg112aRate: 0.10,
    ltcg112aExempt: 100000,
    ltcgOtherRate: 0.20,
    ltcgOtherLabel: '20% (with indexation)',
    maxSurchargeNew: 0.25,
  },
  '2024-25': {
    label: 'PY 2024-25 (AY 2025-26)',
    ayLabel: 'AY 2025-26',
    isFuture: false,
    stdDeduction: 75000,
    hasTransitional: true,
    newSlabs: [
      { upto: 300000,  rate: 0 },
      { upto: 700000,  rate: 0.05 },
      { upto: 1000000, rate: 0.10 },
      { upto: 1200000, rate: 0.15 },
      { upto: 1500000, rate: 0.20 },
      { upto: Infinity, rate: 0.30 },
    ],
    rebateNew: { limit: 700000, max: 25000 },
    rebateOld: { limit: 500000, max: 12500 },
    stcg111aRate: 0.20,
    ltcg112aRate: 0.125,
    ltcg112aExempt: 125000,
    ltcgOtherRate: 0.125,
    ltcgOtherLabel: '12.5% (post July 2024)',
    stcg111aRateOld: 0.15,
    ltcg112aRateOld: 0.10,
    ltcg112aExemptOld: 100000,
    ltcgOtherRateOld: 0.20,
    ltcgOtherLabelOld: '20% with indexation (pre July 2024)',
    maxSurchargeNew: 0.25,
  },
  '2025-26': {
    label: 'PY 2025-26 (AY 2026-27)',
    ayLabel: 'AY 2026-27',
    isFuture: false,
    stdDeduction: 75000,
    newSlabs: [
      { upto: 400000,  rate: 0 },
      { upto: 800000,  rate: 0.05 },
      { upto: 1200000, rate: 0.10 },
      { upto: 1600000, rate: 0.15 },
      { upto: 2000000, rate: 0.20 },
      { upto: 2400000, rate: 0.25 },
      { upto: Infinity, rate: 0.30 },
    ],
    rebateNew: { limit: 1200000, max: 60000 },
    rebateOld: { limit: 500000, max: 12500 },
    stcg111aRate: 0.20,
    ltcg112aRate: 0.125,
    ltcg112aExempt: 125000,
    ltcgOtherRate: 0.125,
    ltcgOtherLabel: '12.5%',
    maxSurchargeNew: 0.25,
  },
  '2026-27': {
    label: 'PY 2026-27 (AY 2027-28)',
    ayLabel: 'AY 2027-28',
    isFuture: true,
    stdDeduction: 75000,
    newSlabs: [
      { upto: 400000,  rate: 0 },
      { upto: 800000,  rate: 0.05 },
      { upto: 1200000, rate: 0.10 },
      { upto: 1600000, rate: 0.15 },
      { upto: 2000000, rate: 0.20 },
      { upto: 2400000, rate: 0.25 },
      { upto: Infinity, rate: 0.30 },
    ],
    rebateNew: { limit: 1200000, max: 60000 },
    rebateOld: { limit: 500000, max: 12500 },
    stcg111aRate: 0.20,
    ltcg112aRate: 0.125,
    ltcg112aExempt: 125000,
    ltcgOtherRate: 0.125,
    ltcgOtherLabel: '12.5%',
    maxSurchargeNew: 0.25,
  },
};

/* ── OLD REGIME SLABS ─────────────────────────────────────────────── */
const OLD_SLABS_BELOW60 = [
  { upto: 250000,  rate: 0 },
  { upto: 500000,  rate: 0.05 },
  { upto: 1000000, rate: 0.20 },
  { upto: Infinity, rate: 0.30 },
];
const OLD_SLABS_SENIOR = [
  { upto: 300000,  rate: 0 },
  { upto: 500000,  rate: 0.05 },
  { upto: 1000000, rate: 0.20 },
  { upto: Infinity, rate: 0.30 },
];
const OLD_SLABS_SUPERSENIOR = [
  { upto: 500000,  rate: 0 },
  { upto: 1000000, rate: 0.20 },
  { upto: Infinity, rate: 0.30 },
];
/* Co-operative Society slabs */
const COOP_SLABS = [
  { upto: 10000,  rate: 0.10 },
  { upto: 20000,  rate: 0.20 },
  { upto: Infinity, rate: 0.30 },
];

function getOldSlabs() {
  const ac = getAssesseeCfg();
  if (ac.group === 'coop') return COOP_SLABS;
  const age = ac.age || 'below60';
  if (age === 'supersenior') return OLD_SLABS_SUPERSENIOR;
  if (age === 'senior') return OLD_SLABS_SENIOR;
  return OLD_SLABS_BELOW60;
}

function cfg() { return YEAR_CONFIG[currentYear]; }

/* ── UI: Year selection ───────────────────────────────────────────── */
function setYear(yr) {
  currentYear = yr;
  document.querySelectorAll('.year-pill').forEach(b => b.classList.remove('active'));
  event.currentTarget.classList.add('active');

  const c = cfg();
  document.getElementById('stdDeduction').value = c.stdDeduction;
  document.getElementById('stdDedHint').textContent =
    c.stdDeduction === 50000 ? '₹50,000 for PY 2023-24' : '₹75,000 for PY 2024-25 onwards';

  const isTrans = c.hasTransitional || false;
  document.getElementById('transitionalNotice').style.display = isTrans ? 'block' : 'none';
  document.getElementById('preJulyFields').style.display = isTrans ? 'block' : 'none';

  if (isTrans) {
    document.getElementById('stcg111aLabel').textContent = 'STCG 111A — Post-July (equity) @ 20%';
    document.getElementById('ltcg112aLabel').textContent = 'LTCG 112A — Post-July (equity) @ 12.5%';
    document.getElementById('ltcgOtherLabel').textContent = 'LTCG Other — Post-July @ 12.5%';
  } else {
    document.getElementById('stcg111aLabel').textContent = 'STCG u/s 111A (equity, STT paid)';
    document.getElementById('ltcg112aLabel').textContent = 'LTCG u/s 112A (equity, STT paid)';
    document.getElementById('ltcgOtherLabel').textContent = 'LTCG — Other (property, debt etc.)';
  }

  document.getElementById('stcgRateHint').textContent =
    'Taxed at ' + (c.stcg111aRate * 100) + '%' + (isTrans ? ' (post 23 July 2024)' : '');
  document.getElementById('ltcgRateHint').textContent =
    (c.ltcg112aRate * 100) + '% above ₹' + (c.ltcg112aExempt / 100000).toFixed(2).replace('.00','') + ' lakh exemption';
  document.getElementById('ltcgOtherHint').textContent = c.ltcgOtherLabel;

  document.getElementById('futureYearNote').style.display = c.isFuture ? 'block' : 'none';

  if (!isTrans) {
    ['stcg111aPreJuly','ltcg112aPreJuly','ltcgOtherPreJuly'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
  }

  updateRefSlabs();
  document.getElementById('resultPanel').classList.remove('show');
  document.getElementById('preCalcInfo').style.display = 'block';
  // Hide advance tax card until recalculated
  document.getElementById('advanceTaxCard').style.display = 'none';
  document.getElementById('matAmtResultCard').style.display = 'none';
}

function setRegime(r) {
  currentRegime = r;
  document.querySelectorAll('.regime-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('btn-' + r).classList.add('active');
  const dc = document.getElementById('deductions-card');
  if (r === 'new') { dc.style.opacity = '0.4'; dc.style.pointerEvents = 'none'; }
  else { dc.style.opacity = '1'; dc.style.pointerEvents = 'auto'; }
}

function updateRefSlabs() {
  const c = cfg();
  const slabs = c.newSlabs;
  document.getElementById('refSlabTitle').textContent = 'New Regime Slab Rates';
  document.getElementById('refSlabSub').textContent = c.label + (c.isFuture ? ' (Estimated)' : '');
  let h = '<table class="slab-table"><thead><tr><th>Income Slab</th><th style="text-align:right">Rate</th></tr></thead><tbody>';
  let prev = 0;
  for (const s of slabs) {
    const from = '₹' + prev.toLocaleString('en-IN');
    const to = s.upto === Infinity ? '& above' : '₹' + s.upto.toLocaleString('en-IN');
    const label = s.upto === Infinity ? 'Above ' + from : from + ' – ' + to;
    h += '<tr><td>' + (prev === 0 ? 'Up to ' + to : label) + '</td><td class="amt">' + (s.rate === 0 ? 'Nil' : (s.rate*100)+'%') + '</td></tr>';
    prev = s.upto;
  }
  h += '</tbody></table>';
  const rebateInfo = c.rebateNew.limit >= 1200000
    ? 'Rebate u/s 87A: Income up to ₹' + (c.rebateNew.limit/100000) + ' lakh → zero tax (max rebate ₹' + c.rebateNew.max.toLocaleString('en-IN') + ').'
    : 'Rebate u/s 87A: Income up to ₹' + (c.rebateNew.limit/100000) + ' lakh → zero tax (max rebate ₹' + c.rebateNew.max.toLocaleString('en-IN') + ').';
  h += '<p class="hint" style="margin-top:10px">' + rebateInfo + ' Standard deduction ₹' + c.stdDeduction.toLocaleString('en-IN') + ' for salaried.</p>';
  document.getElementById('refSlabBody').innerHTML = h;

  document.getElementById('refSpecialSub').textContent = c.label;
  let sp = '<table class="slab-table"><thead><tr><th>Type</th><th style="text-align:right">Rate</th></tr></thead><tbody>';
  sp += '<tr><td>STCG u/s 111A (equity, STT paid)</td><td class="amt">' + (c.stcg111aRate*100) + '%</td></tr>';
  sp += '<tr><td>LTCG u/s 112A (equity) above ₹' + (c.ltcg112aExempt/100000).toFixed(2).replace('.00','') + 'L</td><td class="amt">' + (c.ltcg112aRate*100) + '%</td></tr>';
  sp += '<tr><td>LTCG — Other assets</td><td class="amt">' + c.ltcgOtherLabel + '</td></tr>';
  sp += '<tr><td>Winnings (lottery, games, etc.)</td><td class="amt">30%</td></tr>';
  sp += '</tbody></table>';
  sp += '<p class="hint" style="margin-top:10px">Health &amp; Education Cess @ 4% on tax + surcharge. Max surcharge under new regime: 25%.</p>';
  document.getElementById('refSpecialBody').innerHTML = sp;
}

function syncMatBookProfit() {
  const v = document.getElementById('co_bookProfit').value;
  const el = document.getElementById('matBookProfit');
  if (el) el.value = v;
  computeMatAmt();
}
function syncAmtAti() {
  const v = document.getElementById('firm_amtAti').value;
  const el = document.getElementById('amtAti');
  if (el) el.value = v;
  computeMatAmt();
}

/* helper: read value by id safely */
function v(id) { return parseFloat(document.getElementById(id).value) || 0; }
function fmt(n) {
  if (n < 0) return '-₹' + Math.abs(Math.round(n)).toLocaleString('en-IN');
  return '₹' + Math.round(n).toLocaleString('en-IN');
}

/* ── Slab-based tax ───────────────────────────────────────────────── */
function calcSlabTax(taxableIncome, slabs) {
  let tax = 0, prev = 0;
  const breakup = [];
  for (const slab of slabs) {
    if (taxableIncome <= prev) break;
    const chunk = Math.min(taxableIncome, slab.upto) - prev;
    const t = chunk * slab.rate;
    breakup.push({ from: prev, to: Math.min(taxableIncome, slab.upto), rate: slab.rate, amount: chunk, tax: t });
    tax += t;
    prev = slab.upto;
  }
  return { tax, breakup };
}

/* ── Surcharge ────────────────────────────────────────────────────── */
function calcSurcharge(tax, totalIncome, isNewRegime) {
  if (totalIncome <= 5000000) return 0;
  let rate = 0;
  const maxNew = cfg().maxSurchargeNew;
  if (isNewRegime) {
    if (totalIncome <= 10000000) rate = 0.10;
    else if (totalIncome <= 20000000) rate = 0.15;
    else rate = maxNew;
  } else {
    if (totalIncome <= 10000000) rate = 0.10;
    else if (totalIncome <= 20000000) rate = 0.15;
    else if (totalIncome <= 50000000) rate = 0.25;
    else rate = 0.37;
  }
  let surcharge = tax * rate;
  const thresholds = [5000000, 10000000, 20000000, 50000000];
  for (const th of thresholds) {
    if (totalIncome > th && totalIncome <= th * 1.2) {
      const excess = totalIncome - th;
      const slabs = isNewRegime ? cfg().newSlabs : getOldSlabs();
      const taxAtTh = calcSlabTax(th, slabs).tax;
      const surchAtTh = calcSurcharge(taxAtTh, th, isNewRegime);
      const maxTax = taxAtTh + surchAtTh + excess;
      if (tax + surcharge > maxTax) {
        surcharge = maxTax - tax;
        if (surcharge < 0) surcharge = 0;
      }
    }
  }
  return surcharge;
}

function calcSurchargeCapped(tax, totalIncome) {
  if (totalIncome <= 5000000) return 0;
  let rate = Math.min(0.15, totalIncome > 10000000 ? 0.15 : 0.10);
  return tax * rate;
}

/* ── Company / Firm tax computation ───────────────────────────────── */
function computeForCompanyFirm(ac, totalIncome) {
  const flatRate = ac.flatRate;
  let baseTax = totalIncome * flatRate;

  // Surcharge for companies
  let surcharge = 0;
  if (ac.group === 'company') {
    const low = ac.surchargeRateLow || 0;
    const high = ac.surchargeRateHigh || 0;
    const th = ac.surchargeThreshold || 1e7;
    if (totalIncome > th) surcharge = baseTax * high;
    else if (totalIncome > 1e7) surcharge = baseTax * low;
    else surcharge = baseTax * low;
  } else if (ac.group === 'firm') {
    if (totalIncome > 1e7) surcharge = baseTax * (ac.surchargeRate || 0.12);
  }
  const cess = (baseTax + surcharge) * 0.04;
  const totalTax = baseTax + surcharge + cess;
  return { baseTax, surcharge, cess, totalTax, flatRate };
}

/* ── MAT / AMT Computation ────────────────────────────────────────── */
function computeMatAmt() {
  const t = document.getElementById('assesseeType').value;
  const ac = ASSESSEE_CFG[t];
  if (!ac) return;

  if (ac.canMat) {
    const bookProfit = parseFloat(document.getElementById('matBookProfit').value) || 0;
    if (!bookProfit) { document.getElementById('matResult').style.display='none'; return; }
    const matRate = parseFloat(document.getElementById('matRate').value) || 0.15;
    const matBase = bookProfit * matRate;
    let surcharge = 0;
    if (ac.surchargeRateLow) surcharge = matBase * (bookProfit > 1e7 ? (ac.surchargeRateHigh||0) : (ac.surchargeRateLow||0));
    const matCess = (matBase + surcharge) * 0.04;
    const matTotal = matBase + surcharge + matCess;
    let h = '';
    h += `<div class="mat-row"><span class="ml">Book Profit u/s 115JB</span><span class="mv">${fmt(bookProfit)}</span></div>`;
    h += `<div class="mat-row"><span class="ml">MAT Rate</span><span class="mv">${(matRate*100).toFixed(1)}%</span></div>`;
    h += `<div class="mat-row"><span class="ml">MAT (before surcharge/cess)</span><span class="mv">${fmt(matBase)}</span></div>`;
    if (surcharge) h += `<div class="mat-row"><span class="ml">Surcharge</span><span class="mv">${fmt(surcharge)}</span></div>`;
    h += `<div class="mat-row"><span class="ml">Cess @ 4%</span><span class="mv">${fmt(matCess)}</span></div>`;
    h += `<div class="mat-row mt"><span class="ml">Total MAT Payable</span><span class="mv">${fmt(matTotal)}</span></div>`;
    h += `<p style="font-size:11px;margin-top:8px;color:#78350F">Tax payable = <strong>MAX(Normal Tax, MAT)</strong>. If MAT &gt; Normal Tax, excess = MAT Credit u/s 115JAA (carry fwd 15 yrs).</p>`;
    document.getElementById('matResultRows').innerHTML = h;
    document.getElementById('matResult').style.display = 'block';

  } else if (ac.canAmt && document.getElementById('amtSection').style.display !== 'none') {
    const ati = parseFloat(document.getElementById('amtAti').value) || 0;
    if (!ati) { document.getElementById('amtResult').style.display='none'; return; }
    const amtBase = ati * 0.185;
    let surcharge = 0;
    if (ac.group === 'firm' && ati > 1e7) surcharge = amtBase * 0.12;
    const amtCess = (amtBase + surcharge) * 0.04;
    const amtTotal = amtBase + surcharge + amtCess;
    let h = '';
    h += `<div class="mat-row"><span class="ml">Adjusted Total Income (ATI)</span><span class="mv">${fmt(ati)}</span></div>`;
    h += `<div class="mat-row"><span class="ml">AMT Rate</span><span class="mv">18.5%</span></div>`;
    h += `<div class="mat-row"><span class="ml">AMT (before surcharge/cess)</span><span class="mv">${fmt(amtBase)}</span></div>`;
    if (surcharge) h += `<div class="mat-row"><span class="ml">Surcharge</span><span class="mv">${fmt(surcharge)}</span></div>`;
    h += `<div class="mat-row"><span class="ml">Cess @ 4%</span><span class="mv">${fmt(amtCess)}</span></div>`;
    h += `<div class="mat-row mt"><span class="ml">Total AMT Payable</span><span class="mv">${fmt(amtTotal)}</span></div>`;
    h += `<p style="font-size:11px;margin-top:8px;color:#78350F">Tax payable = <strong>MAX(Normal Tax, AMT)</strong>. If AMT &gt; Normal Tax, excess = AMT Credit u/s 115JD (carry fwd 15 yrs).</p>`;
    document.getElementById('amtResultRows').innerHTML = h;
    document.getElementById('amtResult').style.display = 'block';
  }
}

/* ── Advance Tax Schedule ─────────────────────────────────────────── */
function renderAdvanceTaxSchedule(totalTax, tdsPaid, tcsPaid) {
  // Only for PY 2026-27
  if (currentYear !== '2026-27') {
    document.getElementById('advanceTaxCard').style.display = 'none';
    return;
  }
  const netLiability = Math.max(0, totalTax - tdsPaid - tcsPaid);
  if (netLiability < 10000) {
    document.getElementById('advanceTaxCard').style.display = 'none';
    return;
  }

  // Instalments for PY 2026-27 (u/s 211)
  const today = new Date();
  const instalments = [
    { due: new Date('2026-06-15'), cumPct: 15,  pct: 15,  label: '1st Instalment' },
    { due: new Date('2026-09-15'), cumPct: 45,  pct: 30,  label: '2nd Instalment' },
    { due: new Date('2026-12-15'), cumPct: 75,  pct: 30,  label: '3rd Instalment' },
    { due: new Date('2027-03-15'), cumPct: 100, pct: 25,  label: '4th Instalment' },
  ];

  // For presumptive income assessees (Sec 44AD/44ADA) — single instalment
  const isPrespumptive = false; // could add toggle later

  let h = `<div style="margin-bottom:12px;padding:10px 14px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;font-size:11px;color:#065F46">
    <strong>💡 Total Tax Liability (after TDS/TCS):</strong> ${fmt(netLiability)} — Advance Tax required (≥ ₹10,000)
  </div>`;

  h += `<table class="at-table">
    <thead>
      <tr>
        <th>Instalment</th>
        <th style="text-align:center">% of Tax</th>
        <th>Due Date</th>
        <th style="text-align:right">Amount Due</th>
        <th style="text-align:right">Cumulative</th>
      </tr>
    </thead>
    <tbody>`;

  let cumAmount = 0;
  instalments.forEach((inst, idx) => {
    const amt = Math.round(netLiability * inst.pct / 100);
    const cumAmt = Math.round(netLiability * inst.cumPct / 100);
    cumAmount = cumAmt;

    const isPast = today > inst.due;
    const isNext = !isPast && (idx === 0 || today > instalments[idx-1].due);
    const rowClass = isPast ? 'overdue' : (isNext ? 'upcoming' : '');
    const statusBadge = isPast
      ? '<span style="font-size:10px;background:#FEE2E2;color:#B91C1C;padding:2px 7px;border-radius:10px;margin-left:6px;font-weight:700">Due</span>'
      : (isNext ? '<span style="font-size:10px;background:#FEF3C7;color:#92400E;padding:2px 7px;border-radius:10px;margin-left:6px;font-weight:700">Next</span>' : '');

    const dueDateStr = inst.due.toLocaleDateString('en-IN', { day:'numeric', month:'short', year:'numeric' });

    h += `<tr class="${rowClass}">
      <td><strong>${inst.label}</strong>${statusBadge}</td>
      <td class="pct">${inst.pct}%</td>
      <td class="due">${dueDateStr}</td>
      <td class="amt-cell">${fmt(amt)}</td>
      <td class="amt-cell"><span class="cumul">Cumul: </span>${fmt(cumAmt)}</td>
    </tr>`;
  });

  h += `</tbody>
    <tfoot>
      <tr style="background:#EFF6FF;font-weight:800">
        <td colspan="3" style="padding:10px;font-size:12px">Total Advance Tax</td>
        <td></td>
        <td class="amt-cell" style="color:var(--brand);font-size:13px">${fmt(netLiability)}</td>
      </tr>
    </tfoot>
  </table>`;

  h += `<div style="margin-top:12px;font-size:11px;color:var(--muted);line-height:1.7">
    <strong>📌 Notes:</strong>
    <ul style="margin:6px 0 0 16px;padding:0">
      <li>Instalments calculated on <em>estimated</em> total income for PY 2026-27.</li>
      <li>Assessees under Sec 44AD / 44ADA (presumptive) may pay <strong>entire advance tax by 15 March 2027</strong> (single instalment).</li>
      <li>Senior citizens (60+) with <em>no</em> business income are <strong>exempt</strong> from advance tax u/s 207.</li>
      <li>If advance tax paid &lt; 90% of assessed tax → interest u/s 234B @ 1%/month on shortfall.</li>
      <li>Shortfall in each instalment → interest u/s 234C @ 1%/month for 3 months (1 month for last instalment).</li>
    </ul>
  </div>`;

  document.getElementById('advanceTaxTable').innerHTML = h;
  document.getElementById('advanceTaxCard').style.display = 'block';
}

/* ── Main Computation ─────────────────────────────────────────────── */
function computeForRegime(isNew) {
  const c = cfg();
  const ac = getAssesseeCfg();
  const t = document.getElementById('assesseeType').value;
  const name = document.getElementById('assesseeName').value.trim();
  const isResident = (ac.isNRI || (document.getElementById('residentialStatus') && document.getElementById('residentialStatus').value === 'nri')) ? false : true;

  const isCo   = ac.group === 'company';
  const isFirm = ac.group === 'firm';
  const isCoop = ac.group === 'coop';
  const isLocal= ac.group === 'local';
  const isInd  = ac.group === 'individual';

  /* ── Read income from the right panel ── */
  let grossSalary=0, exemptAllow=0, stdDed=0, salaryIncome=0;
  let houseRaw=0, loanInt=0, houseIncome=0, houseLossCapped=0;
  let businessIncome=0;
  let stcg111a=0, stcgOther=0, ltcg112a=0, ltcgOther=0;
  let stcg111aPreJuly=0, ltcg112aPreJuly=0, ltcgOtherPreJuly=0;
  let ltcg112aPreJulyExemptAmt=0;
  let otherIncome=0, winnings=0;

  if (isInd) {
    grossSalary  = v('grossSalary');
    exemptAllow  = isNew ? 0 : v('exemptAllow');
    stdDed       = v('stdDeduction');
    salaryIncome = Math.max(0, grossSalary - exemptAllow - stdDed);
    houseRaw     = v('houseIncome');
    loanInt      = v('homeLoanInterest');
    houseIncome  = houseRaw - loanInt;
    houseLossCapped = Math.max(Math.min(0, houseIncome), -200000);
    businessIncome= v('businessIncome');
    stcg111a     = v('stcg111a');
    stcgOther    = v('stcgOther');
    ltcg112a     = v('ltcg112a');
    ltcgOther    = v('ltcgOther');
    const isTrans= c.hasTransitional || false;
    if (isTrans) {
      stcg111aPreJuly = v('stcg111aPreJuly');
      ltcg112aPreJuly = v('ltcg112aPreJuly');
      ltcgOtherPreJuly= v('ltcgOtherPreJuly');
      ltcg112aPreJulyExemptAmt = Math.min(ltcg112aPreJuly, c.ltcg112aExemptOld||100000);
    }
    otherIncome  = v('otherIncome');
    winnings     = v('winningsIncome');

  } else if (isCo) {
    businessIncome= v('co_businessIncome');
    stcg111a     = v('co_stcg111a');
    stcgOther    = v('co_stcgOther');
    ltcg112a     = v('co_ltcg112a');
    ltcgOther    = v('co_ltcgOther');
    otherIncome  = v('co_otherIncome');
    // Sync book profit to MAT field
    const bp = v('co_bookProfit');
    const matEl = document.getElementById('matBookProfit');
    if (matEl && bp) matEl.value = bp;

  } else if (isFirm) {
    businessIncome= v('firm_businessIncome');
    stcg111a     = v('firm_stcg111a');
    stcgOther    = v('firm_stcgOther');
    ltcg112a     = v('firm_ltcg112a');
    ltcgOther    = v('firm_ltcgOther');
    otherIncome  = v('firm_otherIncome');
    const ati = v('firm_amtAti');
    const amtEl = document.getElementById('amtAti');
    if (amtEl && ati) amtEl.value = ati;

  } else if (isCoop) {
    businessIncome= v('coop_businessIncome');
    otherIncome  = v('coop_otherIncome');

  } else if (isLocal) {
    businessIncome= v('local_income');
  }

  const isTrans = c.hasTransitional || false;

  const normalIncome = salaryIncome + Math.max(0, houseIncome) + businessIncome + stcgOther + otherIncome;
  const normalAfterLoss = Math.max(0, normalIncome + houseLossCapped);

  // Deductions (only individuals, not companies/firms/coop/local)
  let totalDeductions = 0;
  if (isInd) {
    const ded80ccd2 = v('ded80ccd2');
    if (isNew) {
      totalDeductions = ded80ccd2;
    } else {
      totalDeductions = Math.min(v('ded80c'), 150000) + Math.min(v('ded80ccd1b'), 50000) +
        ded80ccd2 + v('ded80d') + v('ded80e') + v('ded80g') + v('ded80tta') + v('dedOther');
    }
  }

  const normalTaxable = Math.max(0, normalAfterLoss - totalDeductions);

  let normalTax, surchargeNormal, surchargeSpecial, totalSurcharge, slabResult;
  let rebate87a = 0;

  if (isCo || isFirm || isLocal) {
    const flatRes = computeForCompanyFirm(ac, normalTaxable);
    normalTax = flatRes.baseTax;
    surchargeNormal = flatRes.surcharge;
    surchargeSpecial = 0;
    totalSurcharge = flatRes.surcharge;
    slabResult = { tax: flatRes.baseTax, breakup: [{ from:0, to:normalTaxable, rate:ac.flatRate, amount:normalTaxable, tax:flatRes.baseTax }] };

    // Special rate taxes for company/firm
    const taxSTCG111A = stcg111a * c.stcg111aRate;
    const taxLTCG112A = Math.max(0, ltcg112a - c.ltcg112aExempt) * c.ltcg112aRate;
    const taxLTCGOther= ltcgOther * c.ltcgOtherRate;
    const totalSpecialTax = taxSTCG111A + taxLTCG112A + taxLTCGOther;
    const ltcg112aExemptAmt = Math.min(ltcg112a, c.ltcg112aExempt);
    const totalIncome = normalTaxable + stcg111a + ltcg112a + ltcgOther;
    const flatRes2 = computeForCompanyFirm(ac, normalTaxable);
    const cess = flatRes2.cess;
    const totalTax = flatRes2.totalTax + totalSpecialTax;
    const tdsPaid = v('tds'); const tcsPaid = v('tcs'); const advTax = v('advanceTax');
    const totalPrepaid = tdsPaid + tcsPaid + advTax;
    const netPayable = totalTax - totalPrepaid;

    return {
      yearLabel:c.label, ayLabel:c.ayLabel, isFuture:c.isFuture, isTrans:false,
      name, isNew, isResident, c, ac, assesseeType:t,
      grossSalary:0, exemptAllow:0, stdDed:0, salaryIncome:0,
      houseRaw:0, loanInt:0, houseIncome:0, houseLossCapped:0,
      businessIncome, stcg111a, stcgOther, ltcg112a, ltcg112aExemptAmt, ltcgOther,
      stcg111aPreJuly:0, ltcg112aPreJuly:0, ltcg112aPreJulyExemptAmt:0, ltcgOtherPreJuly:0,
      otherIncome, winnings:0,
      normalIncome:normalAfterLoss, totalDeductions:0, normalTaxable,
      slabResult:{ tax:flatRes2.baseTax, breakup:[{ from:0, to:normalTaxable, rate:ac.flatRate, amount:normalTaxable, tax:flatRes2.baseTax }] },
      normalTax:flatRes2.baseTax, rebate87a:0, normalTaxAfterRebate:flatRes2.baseTax,
      taxSTCG111A, taxLTCG112A, taxLTCGOther, taxWinnings:0,
      taxSTCG111APreJuly:0, taxLTCG112APreJuly:0, taxLTCGOtherPreJuly:0,
      totalSpecialTax, totalIncome, surchargeNormal:flatRes2.surcharge, surchargeSpecial:0, totalSurcharge:flatRes2.surcharge,
      cess, totalTax, tdsPaid, tcsPaid, advTax, totalPrepaid, netPayable,
      isFlatRate:true, flatRate:ac.flatRate,
    };
  }

  if (isCoop) {
    // Co-operative society slabs
    slabResult = calcSlabTax(normalTaxable, COOP_SLABS);
    normalTax  = slabResult.tax;
    const surcharge = normalTaxable > 1e7 ? normalTax * 0.12 : 0;
    const cess = (normalTax + surcharge) * 0.04;
    const totalTax = normalTax + surcharge + cess;
    const tdsPaid = v('tds'); const tcsPaid = v('tcs'); const advTax = v('advanceTax');
    const totalPrepaid = tdsPaid + tcsPaid + advTax;
    const netPayable = totalTax - totalPrepaid;
    return {
      yearLabel:c.label, ayLabel:c.ayLabel, isFuture:c.isFuture, isTrans:false,
      name, isNew:false, isResident, c, ac, assesseeType:t,
      grossSalary:0, exemptAllow:0, stdDed:0, salaryIncome:0,
      houseRaw:0, loanInt:0, houseIncome:0, houseLossCapped:0,
      businessIncome, stcg111a:0, stcgOther:0, ltcg112a:0, ltcg112aExemptAmt:0, ltcgOther:0,
      stcg111aPreJuly:0, ltcg112aPreJuly:0, ltcg112aPreJulyExemptAmt:0, ltcgOtherPreJuly:0,
      otherIncome, winnings:0,
      normalIncome:normalAfterLoss, totalDeductions:0, normalTaxable,
      slabResult, normalTax, rebate87a:0, normalTaxAfterRebate:normalTax,
      taxSTCG111A:0, taxLTCG112A:0, taxLTCGOther:0, taxWinnings:0,
      taxSTCG111APreJuly:0, taxLTCG112APreJuly:0, taxLTCGOtherPreJuly:0,
      totalSpecialTax:0, totalIncome:normalTaxable,
      surchargeNormal:surcharge, surchargeSpecial:0, totalSurcharge:surcharge,
      cess, totalTax, tdsPaid, tcsPaid, advTax, totalPrepaid, netPayable,
    };
  }

  // Individual / HUF / AOP / Trust / Artificial Person — slab-based
  const slabs = isNew ? c.newSlabs : getOldSlabs();
  slabResult = calcSlabTax(normalTaxable, slabs);
  normalTax  = slabResult.tax;

  if (isResident) {
    const rebateCfg = isNew ? c.rebateNew : c.rebateOld;
    if (normalTaxable <= rebateCfg.limit) {
      rebate87a = Math.min(normalTax, rebateCfg.max);
    }
  }
  normalTax = Math.max(0, normalTax - rebate87a);

  const taxSTCG111A = stcg111a * c.stcg111aRate;
  const taxLTCG112A = Math.max(0, ltcg112a - c.ltcg112aExempt) * c.ltcg112aRate;
  const taxLTCGOther= ltcgOther * c.ltcgOtherRate;
  const taxWinnings = winnings * 0.30;

  let taxSTCG111APreJuly=0, taxLTCG112APreJuly=0, taxLTCGOtherPreJuly=0;
  if (isTrans) {
    taxSTCG111APreJuly = stcg111aPreJuly * (c.stcg111aRateOld||0.15);
    taxLTCG112APreJuly = Math.max(0, ltcg112aPreJuly - ltcg112aPreJulyExemptAmt) * (c.ltcg112aRateOld||0.10);
    taxLTCGOtherPreJuly= ltcgOtherPreJuly * (c.ltcgOtherRateOld||0.20);
  }

  const totalSpecialTax = taxSTCG111A + taxLTCG112A + taxLTCGOther + taxWinnings +
                          taxSTCG111APreJuly + taxLTCG112APreJuly + taxLTCGOtherPreJuly;

  const totalIncome = normalTaxable + stcg111a + stcg111aPreJuly + ltcg112a + ltcg112aPreJuly +
                      ltcgOther + ltcgOtherPreJuly + winnings;

  surchargeNormal  = calcSurcharge(normalTax, totalIncome, isNew);
  surchargeSpecial = calcSurchargeCapped(totalSpecialTax, totalIncome);
  totalSurcharge   = surchargeNormal + surchargeSpecial;

  const totalBeforeCess = normalTax + totalSpecialTax + totalSurcharge;
  const cess     = totalBeforeCess * 0.04;
  const totalTax = totalBeforeCess + cess;
  const tdsPaid  = v('tds'); const tcsPaid = v('tcs'); const advTax = v('advanceTax');
  const totalPrepaid = tdsPaid + tcsPaid + advTax;
  const netPayable   = totalTax - totalPrepaid;
  const ltcg112aExemptAmt = Math.min(ltcg112a, c.ltcg112aExempt);

  return {
    yearLabel:c.label, ayLabel:c.ayLabel, isFuture:c.isFuture, isTrans,
    name, isNew, isResident, c, ac, assesseeType:t,
    grossSalary, exemptAllow, stdDed, salaryIncome,
    houseRaw, loanInt, houseIncome, houseLossCapped,
    businessIncome,
    stcg111a, stcgOther, ltcg112a, ltcg112aExemptAmt, ltcgOther,
    stcg111aPreJuly, ltcg112aPreJuly, ltcg112aPreJulyExemptAmt, ltcgOtherPreJuly,
    otherIncome, winnings,
    normalIncome:normalAfterLoss, totalDeductions, normalTaxable,
    slabResult,
    normalTax: normalTax + rebate87a, rebate87a,
    normalTaxAfterRebate: normalTax,
    taxSTCG111A, taxLTCG112A, taxLTCGOther, taxWinnings,
    taxSTCG111APreJuly, taxLTCG112APreJuly, taxLTCGOtherPreJuly,
    totalSpecialTax, totalIncome,
    surchargeNormal, surchargeSpecial, totalSurcharge,
    cess, totalTax, tdsPaid, tcsPaid, advTax, totalPrepaid, netPayable,
  };
}

/* ── Render result rows ────────────────────────────────────────────── */
function row(lbl, val, cls) {
  return '<div class="result-row '+(cls||'')+'"><span class="lbl">'+lbl+'</span><span class="val">'+val+'</span></div>';
}

function renderResult(r) {
  const c = r.c;
  const ac = r.ac;
  let h = '';

  // Assessee badge
  h += `<div class="assessee-badge">👤 ${ac.label || ''}</div>`;

  h += row('Gross Salary', fmt(r.grossSalary));
  if (!r.isNew && !r.isFlatRate) h += row('Less: Exempt Allowances', fmt(-r.exemptAllow), 'sub');
  if (!r.isFlatRate) h += row('Less: Standard Deduction (₹'+c.stdDeduction.toLocaleString('en-IN')+')', fmt(-r.stdDed), 'sub');
  h += row('Net Salary Income (Head 1)', fmt(r.salaryIncome));
  h += '<div style="height:6px"></div>';
  h += row('House Property Income (Head 2)', fmt(r.houseIncome));
  if (r.houseLossCapped < 0) h += row('Loss set-off (max ₹2L)', fmt(r.houseLossCapped), 'sub');
  h += row('Business Income (Head 3)', fmt(r.businessIncome));

  if (r.stcg111a || r.stcgOther || r.ltcg112a || r.ltcgOther || r.stcg111aPreJuly || r.ltcg112aPreJuly || r.ltcgOtherPreJuly) {
    const totalCG = r.stcg111a + r.stcgOther + r.ltcg112a + r.ltcgOther + r.stcg111aPreJuly + r.ltcg112aPreJuly + r.ltcgOtherPreJuly;
    h += row('Capital Gains (Head 4)', fmt(totalCG));
  }
  if (r.isTrans && (r.stcg111aPreJuly || r.ltcg112aPreJuly || r.ltcgOtherPreJuly)) {
    if (r.stcg111aPreJuly) h += row('STCG 111A pre-July @ '+(c.stcg111aRateOld*100)+'%', fmt(r.stcg111aPreJuly), 'sub');
    if (r.ltcg112aPreJuly) h += row('LTCG 112A pre-July (exempt ₹'+(c.ltcg112aExemptOld/100000)+'L) @ '+(c.ltcg112aRateOld*100)+'%', fmt(r.ltcg112aPreJuly), 'sub');
    if (r.ltcgOtherPreJuly) h += row('LTCG Other pre-July @ '+c.ltcgOtherLabelOld, fmt(r.ltcgOtherPreJuly), 'sub');
  }
  if (r.stcg111a) h += row((r.isTrans?'STCG 111A post-July':'STCG u/s 111A')+' @ '+(c.stcg111aRate*100)+'%', fmt(r.stcg111a), 'sub');
  if (r.stcgOther) h += row('STCG — Other (slab rate)', fmt(r.stcgOther), 'sub');
  if (r.ltcg112a) h += row((r.isTrans?'LTCG 112A post-July':'LTCG u/s 112A')+' (exempt ₹'+(c.ltcg112aExempt/100000)+'L)', fmt(r.ltcg112a), 'sub');
  if (r.ltcgOther) h += row((r.isTrans?'LTCG Other post-July':'LTCG — Other')+' @ '+c.ltcgOtherLabel, fmt(r.ltcgOther), 'sub');

  h += row('Other Sources (Head 5)', fmt(r.otherIncome + r.winnings));
  if (r.winnings) h += row('Winnings @ 30%', fmt(r.winnings), 'sub');

  h += '<div style="height:4px;border-top:2px solid var(--border);margin:10px 0"></div>';
  const gtiTotal = r.normalIncome + r.stcg111a + r.ltcg112a + r.ltcgOther + r.winnings + r.stcg111aPreJuly + r.ltcg112aPreJuly + r.ltcgOtherPreJuly;
  h += row('Gross Total Income', fmt(gtiTotal), 'total');
  if (r.totalDeductions > 0) h += row('Less: Deductions Ch VI-A', fmt(-r.totalDeductions));
  h += row('Total Taxable Income (Normal)', fmt(r.normalTaxable), 'total');

  h += '<div style="height:4px;border-top:2px solid var(--border);margin:10px 0"></div>';
  if (r.isFlatRate) {
    h += row('Tax @ '+(r.flatRate*100)+'% (flat)', fmt(r.normalTax));
  } else {
    h += row('Tax on Normal Income (slab)', fmt(r.normalTax));
    if (r.rebate87a > 0) h += row('Less: Rebate u/s 87A', fmt(-r.rebate87a), 'sub');
    h += row('Tax after Rebate', fmt(r.normalTaxAfterRebate));
  }

  if (r.totalSpecialTax > 0) {
    h += '<div style="height:6px"></div>';
    if (r.taxSTCG111APreJuly) h += row('Tax STCG 111A pre-July @ '+(c.stcg111aRateOld*100)+'%', fmt(r.taxSTCG111APreJuly), 'sub');
    if (r.taxLTCG112APreJuly) h += row('Tax LTCG 112A pre-July @ '+(c.ltcg112aRateOld*100)+'%', fmt(r.taxLTCG112APreJuly), 'sub');
    if (r.taxLTCGOtherPreJuly) h += row('Tax LTCG Other pre-July @ '+(c.ltcgOtherRateOld*100)+'%', fmt(r.taxLTCGOtherPreJuly), 'sub');
    if (r.taxSTCG111A) h += row('Tax '+(r.isTrans?'STCG 111A post-July':'STCG 111A')+' @ '+(c.stcg111aRate*100)+'%', fmt(r.taxSTCG111A), 'sub');
    if (r.taxLTCG112A) h += row('Tax '+(r.isTrans?'LTCG 112A post-July':'LTCG 112A')+' @ '+(c.ltcg112aRate*100)+'%', fmt(r.taxLTCG112A), 'sub');
    if (r.taxLTCGOther) h += row('Tax '+(r.isTrans?'LTCG Other post-July':'LTCG Other')+' @ '+c.ltcgOtherLabel, fmt(r.taxLTCGOther), 'sub');
    if (r.taxWinnings) h += row('Tax on Winnings @ 30%', fmt(r.taxWinnings), 'sub');
    h += row('Total Special Rate Tax', fmt(r.totalSpecialTax));
  }

  if (r.totalSurcharge > 0) h += row('Surcharge', fmt(r.totalSurcharge));
  h += row('Health & Education Cess @ 4%', fmt(r.cess));
  h += row('Total Tax Liability', fmt(r.totalTax), 'total');

  h += '<div style="height:4px;border-top:2px solid var(--border);margin:10px 0"></div>';
  if (r.tdsPaid) h += row('Less: TDS', fmt(-r.tdsPaid), 'sub');
  if (r.tcsPaid) h += row('Less: TCS', fmt(-r.tcsPaid), 'sub');
  if (r.advTax) h += row('Less: Advance Tax', fmt(-r.advTax), 'sub');

  const cls = r.netPayable > 0 ? 'payable' : 'refund';
  const lbl = r.netPayable > 0 ? 'Net Tax Payable' : 'Refund Due';
  h += row(lbl, fmt(Math.abs(r.netPayable)), 'total ' + cls);
  return h;
}

function renderSlabs(result) {
  const slabs = result.slabResult.breakup;
  let h = '<table class="slab-table"><thead><tr><th>Slab</th><th style="text-align:right">Income</th><th style="text-align:right">Rate</th><th style="text-align:right">Tax</th></tr></thead><tbody>';
  for (const s of slabs) {
    h += '<tr><td>₹'+Math.round(s.from).toLocaleString('en-IN')+' – ₹'+(s.to===Infinity?'∞':Math.round(s.to).toLocaleString('en-IN'))+'</td>';
    h += '<td class="amt">'+fmt(s.amount)+'</td><td class="amt">'+(s.rate*100).toFixed(0)+'%</td><td class="amt">'+fmt(s.tax)+'</td></tr>';
  }
  h += '<tr style="font-weight:800;border-top:2px solid var(--border)"><td>Total</td><td></td><td></td><td class="amt">'+fmt(result.slabResult.tax)+'</td></tr>';
  h += '</tbody></table>';
  return h;
}

/* ── MAT/AMT in results panel ─────────────────────────────────────── */
function renderMatAmtResult(result) {
  const t = document.getElementById('assesseeType').value;
  const ac = ASSESSEE_CFG[t];
  const matCard = document.getElementById('matAmtResultCard');

  if (ac.canMat) {
    const bookProfit = parseFloat(document.getElementById('matBookProfit').value) || 0;
    if (!bookProfit) { matCard.style.display = 'none'; return; }
    const matRate = parseFloat(document.getElementById('matRate').value) || 0.15;
    const matBase = bookProfit * matRate;
    let surcharge = 0;
    if (ac.surchargeRateLow && bookProfit > (ac.surchargeThreshold||1e7)) surcharge = matBase * (ac.surchargeRateHigh||0);
    const matCess = (matBase + surcharge) * 0.04;
    const matTotal = matBase + surcharge + matCess;
    const normalTax = result.totalTax;
    const isMatApplicable = matTotal > normalTax;

    document.getElementById('matAmtResTitle').textContent = 'MAT Summary u/s 115JB';
    let h = '';
    h += `<div style="display:flex;gap:10px;margin-bottom:12px;flex-wrap:wrap">
      <div style="flex:1;min-width:120px;padding:12px;background:#EFF6FF;border-radius:10px;text-align:center">
        <div style="font-size:11px;color:var(--muted);margin-bottom:4px">Normal Tax</div>
        <div style="font-size:16px;font-weight:800;color:var(--brand)">${fmt(normalTax)}</div>
      </div>
      <div style="flex:1;min-width:120px;padding:12px;background:#FEF3C7;border-radius:10px;text-align:center">
        <div style="font-size:11px;color:var(--muted);margin-bottom:4px">MAT</div>
        <div style="font-size:16px;font-weight:800;color:#B45309">${fmt(matTotal)}</div>
      </div>
      <div style="flex:1;min-width:120px;padding:12px;background:${isMatApplicable?'#FEF3C7':'#F0FDF4'};border-radius:10px;text-align:center">
        <div style="font-size:11px;color:var(--muted);margin-bottom:4px">Tax Payable</div>
        <div style="font-size:16px;font-weight:800;color:${isMatApplicable?'#B45309':'#065F46'}">${fmt(Math.max(normalTax, matTotal))}</div>
      </div>
    </div>`;
    if (isMatApplicable) {
      h += `<div style="padding:10px 14px;background:#FEF3C7;border:1px solid #FDE68A;border-radius:8px;font-size:12px;color:#92400E;font-weight:600">
        ⚠️ MAT applies — MAT (${fmt(matTotal)}) &gt; Normal Tax (${fmt(normalTax)}). MAT Credit u/s 115JAA = ${fmt(matTotal - normalTax)} (carry fwd up to 15 years).
      </div>`;
    } else {
      h += `<div style="padding:10px 14px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;font-size:12px;color:#065F46;font-weight:600">
        ✅ Normal Tax (${fmt(normalTax)}) &gt; MAT (${fmt(matTotal)}). Normal tax provisions apply.
      </div>`;
    }
    document.getElementById('matAmtResBody').innerHTML = h;
    matCard.style.display = 'block';
  } else {
    matCard.style.display = 'none';
  }
}

/* ── CALCULATE ────────────────────────────────────────────────────── */
function calculateTax() {
  const c = cfg();
  const panel = document.getElementById('resultPanel');
  const preInfo = document.getElementById('preCalcInfo');
  const ac = getAssesseeCfg();
  const isCo  = ac.group === 'company';
  const isFirm = ac.group === 'firm';

  if (currentRegime === 'both' && !isCo && !isFirm) {
    const rNew = computeForRegime(true);
    const rOld = computeForRegime(false);

    document.getElementById('singleResult').style.display = 'none';
    document.getElementById('compareResult').style.display = 'block';
    document.getElementById('compareYearLabel').textContent = c.label + (c.isFuture ? ' (Estimated)' : '');

    const diff = Math.abs(rNew.totalTax - rOld.totalTax);
    const winner = rNew.totalTax <= rOld.totalTax ? 'New Regime' : 'Old Regime';
    document.getElementById('regimeWinner').innerHTML = '<strong>' + winner + '</strong> saves you more tax';
    document.getElementById('savingsAmt').textContent = 'Save ' + fmt(diff);

    let ct = '';
    ct += cmpRow('Taxable Income', rNew.normalTaxable, rOld.normalTaxable, true);
    ct += cmpRow('Tax on Normal Income', rNew.normalTaxAfterRebate, rOld.normalTaxAfterRebate, true);
    ct += cmpRow('Tax on Special Income', rNew.totalSpecialTax, rOld.totalSpecialTax, true);
    ct += cmpRow('Surcharge', rNew.totalSurcharge, rOld.totalSurcharge, true);
    ct += cmpRow('Cess', rNew.cess, rOld.cess, true);
    ct += cmpRow('Total Tax', rNew.totalTax, rOld.totalTax, true);
    ct += cmpRow('Net Payable/Refund', rNew.netPayable, rOld.netPayable, true);
    document.getElementById('compareBody').innerHTML = ct;

    document.getElementById('newRegimeDetail').innerHTML = renderResult(rNew);
    document.getElementById('oldRegimeDetail').innerHTML = renderResult(rOld);

    document.getElementById('slabRegimeLabel').textContent = c.label;
    document.getElementById('slabBody').innerHTML =
      '<h3 style="font-size:13px;font-weight:700;margin-bottom:8px">🆕 New Regime Slabs</h3>' +
      renderSlabs(rNew) +
      '<div style="height:16px"></div>' +
      '<h3 style="font-size:13px;font-weight:700;margin-bottom:8px">📜 Old Regime Slabs</h3>' +
      renderSlabs(rOld);

    // Regime bar chart
    setTimeout(() => renderRegimeBarChart(rNew, rOld), 400);

    // Advance tax for 2026-27
    renderAdvanceTaxSchedule(rNew.totalTax, rNew.tdsPaid, rNew.tcsPaid);
    renderMatAmtResult(rNew);
  } else {
    const isNew = currentRegime === 'new' || isCo || isFirm;
    const result = computeForRegime(isNew);

    document.getElementById('singleResult').style.display = 'block';
    document.getElementById('compareResult').style.display = 'none';

    let label = '';
    if (isCo) label = '🏢 Company';
    else if (isFirm) label = '🤝 Firm / LLP';
    else label = isNew ? '🆕 New Regime' : '📜 Old Regime';

    document.getElementById('resultTitle').textContent = 'Tax Computation — ' + label;
    document.getElementById('resultSubtitle').textContent =
      (result.name ? result.name + ' · ' : '') + c.label + (c.isFuture ? ' (Estimated)' : '');

    document.getElementById('resultBody').innerHTML = renderResult(result);
    document.getElementById('slabRegimeLabel').textContent = (isCo ? 'Company' : isFirm ? 'Firm/LLP' : (isNew ? 'New Regime' : 'Old Regime')) + ' · ' + c.label;
    document.getElementById('slabBody').innerHTML = renderSlabs(result);

    // Advance tax for 2026-27
    renderAdvanceTaxSchedule(result.totalTax, result.tdsPaid, result.tcsPaid);
    renderMatAmtResult(result);
  }

  document.getElementById('futureDisclaimer').style.display = c.isFuture ? 'block' : 'none';

  panel.classList.add('show');
  preInfo.style.display = 'none';
  document.getElementById('slabCard').style.display = 'block';
  document.getElementById('result-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
  toast('Tax calculated for ' + c.label + '!');
}

function cmpRow(label, valNew, valOld, lowerBetter) {
  const nw = fmt(valNew), ol = fmt(valOld);
  let nCls = '', oCls = '';
  if (lowerBetter) {
    if (valNew < valOld) nCls = 'winner'; else if (valOld < valNew) oCls = 'winner';
  }
  return '<tr><td style="text-align:left;font-weight:500">'+label+'</td><td class="'+nCls+'">'+nw+'</td><td class="'+oCls+'">'+ol+'</td></tr>';
}

function resetForm() {
  document.querySelectorAll('input[type=number]').forEach(i => {
    if (i.id === 'stdDeduction') i.value = cfg().stdDeduction;
    else i.value = '';
  });
  document.getElementById('assesseeName').value = '';
  // Reset assessee buttons
  document.querySelectorAll('.at-btn').forEach(b => b.classList.remove('active'));
  const firstBtn = document.querySelector('.at-btn[data-val="individual_below60"]');
  if (firstBtn) firstBtn.classList.add('active');
  document.getElementById('assesseeType').value = 'individual_below60';
  document.getElementById('resultPanel').classList.remove('show');
  document.getElementById('preCalcInfo').style.display = 'block';
  document.getElementById('singleResult').style.display = 'none';
  document.getElementById('compareResult').style.display = 'none';
  document.getElementById('advanceTaxCard').style.display = 'none';
  document.getElementById('matAmtResultCard').style.display = 'none';
  onAssesseeTypeChange();
  setRegime('new');
  toast('Form reset');
}

function toast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3000);
}

// Initialize
setRegime('new');
updateRefSlabs();
onAssesseeTypeChange();

/* ═══════════════════════════════════════════
   ANIMATION ENGINE
   ═══════════════════════════════════════════ */

/* ── 1. Scroll reveal ── */
function initReveal() {
  const obs = new IntersectionObserver((entries) => {
    entries.forEach(e => { if (e.isIntersecting) { e.target.classList.add('visible'); obs.unobserve(e.target); } });
  }, { threshold: 0.08 });
  document.querySelectorAll('.reveal').forEach(el => obs.observe(el));
}
initReveal();

/* ── 2. Nav scroll shadow ── */
window.addEventListener('scroll', () => {
  document.querySelector('nav').classList.toggle('scrolled', window.scrollY > 10);
});

/* ── 3. Progress bar ── */
function showProgress() {
  const bar = document.getElementById('calcProgress');
  bar.style.width = '0';
  bar.style.transition = 'none';
  requestAnimationFrame(() => {
    bar.style.transition = 'width .35s ease';
    bar.style.width = '70%';
    setTimeout(() => { bar.style.width = '95%'; }, 350);
  });
}
function finishProgress() {
  const bar = document.getElementById('calcProgress');
  bar.style.width = '100%';
  setTimeout(() => { bar.style.width = '0'; bar.style.transition = 'none'; }, 400);
}

/* ── 4. Button ripple ── */
function addRipple(btn, e) {
  const rect = btn.getBoundingClientRect();
  const size = Math.max(rect.width, rect.height) * 2;
  const r = document.createElement('span');
  r.className = 'ripple';
  r.style.cssText = `width:${size}px;height:${size}px;left:${e.clientX - rect.left - size/2}px;top:${e.clientY - rect.top - size/2}px`;
  btn.appendChild(r);
  r.addEventListener('animationend', () => r.remove());
}

/* ── 5. Number counter ── */
function animateCounter(el, target, prefix, suffix, duration) {
  const start = performance.now();
  const startVal = 0;
  function update(now) {
    const progress = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - progress, 3);
    const current = Math.round(startVal + (target - startVal) * ease);
    el.textContent = prefix + current.toLocaleString('en-IN') + suffix;
    if (progress < 1) requestAnimationFrame(update);
  }
  requestAnimationFrame(update);
}

function animateAllCounters() {
  document.querySelectorAll('.val').forEach(el => {
    const text = el.textContent.trim();
    const isNeg = text.startsWith('-₹');
    const clean = text.replace(/[₹,\-]/g, '');
    const num = parseFloat(clean);
    if (!isNaN(num) && num > 0) {
      animateCounter(el, num, isNeg ? '-₹' : '₹', '', 900);
    }
  });
}

/* ── 6. Donut chart ── */
const CIRC = 2 * Math.PI * 54; // 339.3
const DONUT_COLORS = ['#2563EB', '#F59E0B', '#10B981', '#EF4444'];
const DONUT_LABELS = ['Base Tax', 'Surcharge', 'Cess', 'Special Rate Tax'];
const DONUT_IDS    = ['donut-base', 'donut-surcharge', 'donut-cess', 'donut-special'];

function renderDonutChart(baseTax, surcharge, cess, specialTax) {
  const wrap = document.getElementById('taxChartWrap');
  const total = baseTax + surcharge + cess + specialTax;
  if (total <= 0) { wrap.style.display = 'none'; return; }
  wrap.style.display = 'block';

  const values = [baseTax, surcharge, cess, specialTax];
  document.getElementById('donut-center-val').textContent = '₹' + Math.round(total).toLocaleString('en-IN');

  // Build SVG segments — need 4 circles layered with correct dashoffset
  // Remove old special segment if exists
  const oldSpecial = document.getElementById('donut-special');
  if (oldSpecial) oldSpecial.remove();

  let offsetDeg = 0; // starts at top (adjusted by -85° in CSS)
  const ids = ['donut-base', 'donut-surcharge', 'donut-cess'];
  const colors = ['#2563EB', '#F59E0B', '#10B981'];
  const mainVals = [baseTax, surcharge, cess];

  // Also add special if needed
  if (specialTax > 0) {
    const svg = document.querySelector('.donut-svg');
    const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    circle.setAttribute('id', 'donut-special');
    circle.setAttribute('class', 'donut-segment');
    circle.setAttribute('cx', '70'); circle.setAttribute('cy', '70'); circle.setAttribute('r', '54');
    circle.setAttribute('fill', 'none'); circle.setAttribute('stroke', '#EF4444');
    circle.setAttribute('stroke-width', '22');
    circle.setAttribute('stroke-dasharray', '0 339.3');
    circle.setAttribute('stroke-dashoffset', '84.8');
    circle.setAttribute('stroke-linecap', 'round');
    svg.appendChild(circle);
    ids.push('donut-special'); colors.push('#EF4444'); mainVals.push(specialTax);
  }

  let cumOffset = CIRC * 0.25; // start at top
  ids.forEach((id, i) => {
    const el = document.getElementById(id);
    if (!el) return;
    const pct = mainVals[i] / total;
    const dash = pct * CIRC;
    const gap  = CIRC - dash;
    el.setAttribute('stroke', colors[i]);
    // Animate after paint
    setTimeout(() => {
      el.style.strokeDasharray = `${dash} ${gap}`;
      el.style.strokeDashoffset = cumOffset;
    }, 80 + i * 60);
    cumOffset -= dash;
  });

  // Legend
  const legend = document.getElementById('donutLegend');
  const allLabels = ['Base Tax', 'Surcharge', 'Cess', 'Special Rate'];
  legend.innerHTML = ids.map((id, i) => {
    if (mainVals[i] <= 0) return '';
    const pct = ((mainVals[i] / total) * 100).toFixed(1);
    return `<div class="donut-legend-item">
      <span class="donut-dot" style="background:${colors[i]}"></span>
      <span class="donut-label">${allLabels[i]}</span>
      <span class="donut-val">${pct}%</span>
    </div>`;
  }).join('');
}

/* ── 7. Regime bar chart ── */
function renderRegimeBarChart(rNew, rOld) {
  const wrap = document.getElementById('regimeChartWrap');
  const chart = document.getElementById('regimeBarChart');
  const maxVal = Math.max(rNew.totalTax, rOld.totalTax, 1);

  const rows = [
    { label: 'Taxable Income', nv: rNew.normalTaxable, ov: rOld.normalTaxable },
    { label: 'Base Tax', nv: rNew.normalTaxAfterRebate, ov: rOld.normalTaxAfterRebate },
    { label: 'Total Tax', nv: rNew.totalTax, ov: rOld.totalTax },
    { label: 'Net Payable', nv: Math.max(0,rNew.netPayable), ov: Math.max(0,rOld.netPayable) },
  ];

  const maxAll = Math.max(...rows.map(r => Math.max(r.nv, r.ov)), 1);

  chart.innerHTML = rows.map(r => {
    const nPct = (r.nv / maxAll * 100).toFixed(1);
    const oPct = (r.ov / maxAll * 100).toFixed(1);
    const nWinner = r.nv <= r.ov;
    return `<div style="margin-bottom:14px">
      <div style="font-size:11px;font-weight:700;color:var(--muted);margin-bottom:6px;text-transform:uppercase;letter-spacing:.04em">${r.label}</div>
      <div class="bar-row">
        <div class="bar-label" style="color:#2563EB;font-size:10px">🆕 New</div>
        <div class="bar-track">
          <div class="bar-fill" style="background:${nWinner?'#2563EB':'#93C5FD'}" data-pct="${nPct}">
            <span>₹${Math.round(r.nv/1000)}K</span>
          </div>
        </div>
        <div class="bar-val" style="color:${nWinner?'#2563EB':'var(--muted)'}">₹${Math.round(r.nv).toLocaleString('en-IN')}</div>
      </div>
      <div class="bar-row" style="margin-top:4px">
        <div class="bar-label" style="color:#F59E0B;font-size:10px">📜 Old</div>
        <div class="bar-track">
          <div class="bar-fill" style="background:${!nWinner?'#F59E0B':'#FCD34D'}" data-pct="${oPct}">
            <span>₹${Math.round(r.ov/1000)}K</span>
          </div>
        </div>
        <div class="bar-val" style="color:${!nWinner?'#F59E0B':'var(--muted)'}">₹${Math.round(r.ov).toLocaleString('en-IN')}</div>
      </div>
    </div>`;
  }).join('');

  // Animate bars after DOM paint
  requestAnimationFrame(() => {
    setTimeout(() => {
      document.querySelectorAll('.bar-fill').forEach(bar => {
        bar.style.width = bar.dataset.pct + '%';
      });
    }, 100);
  });
}

/* ── 8. Confetti burst ── */
function fireConfetti() {
  const colors = ['#2563EB','#10B981','#F59E0B','#EF4444','#8B5CF6','#EC4899'];
  for (let i = 0; i < 60; i++) {
    const piece = document.createElement('div');
    piece.className = 'confetti-piece';
    piece.style.cssText = `
      left: ${Math.random()*100}vw;
      background: ${colors[Math.floor(Math.random()*colors.length)]};
      width: ${4+Math.random()*6}px;
      height: ${4+Math.random()*6}px;
      border-radius: ${Math.random()>.5?'50%':'2px'};
      animation-duration: ${1.5+Math.random()*2}s;
      animation-delay: ${Math.random()*.5}s;
      opacity: 1;
    `;
    document.body.appendChild(piece);
    piece.addEventListener('animationend', () => piece.remove());
  }
}

/* ── Override calculateTax to wire animations ── */
const _origCalc = calculateTax;
calculateTax = function(e) {
  const btn = document.getElementById('calcBtn');

  // Ripple
  if (e && btn) addRipple(btn, e);

  // Spinner
  if (btn) btn.classList.add('loading');

  // Progress bar
  showProgress();

  // Small delay to show spinner, then compute
  setTimeout(() => {
    _origCalc();
    if (btn) btn.classList.remove('loading');
    finishProgress();

    // Counter animation
    setTimeout(animateAllCounters, 200);

    // Add reveal to result cards
    document.querySelectorAll('#resultPanel .card, #resultPanel > div').forEach((el, i) => {
      el.classList.add('reveal');
      el.style.transitionDelay = (i * 0.07) + 's';
      setTimeout(() => el.classList.add('visible'), 50 + i * 70);
    });

    // Wire up donut chart from result data
    // (called from calculateTax internals via hook below)

  }, 380);
};

/* Hook into renderResult to trigger donut */
const _origRenderResult = renderResult;
renderResult = function(r) {
  const html = _origRenderResult(r);
  // Schedule donut render
  setTimeout(() => {
    const baseTax = r.normalTaxAfterRebate || r.normalTax || 0;
    const surcharge = r.totalSurcharge || 0;
    const cess = r.cess || 0;
    const special = r.totalSpecialTax || 0;
    renderDonutChart(baseTax, surcharge, cess, special);

    // Confetti if zero tax
    if ((r.totalTax || 0) === 0 || (r.netPayable || 0) <= 0) {
      fireConfetti();
    }
  }, 500);
  return html;
};

/* Hook into calculateTax for regime bar chart */
const _origCmpRow = cmpRow;
let _lastRNew = null, _lastROld = null;
const _origCalcTax2 = calculateTax;
// Patch the compare path via renderResult hook on compare
const _origCalcTaxFinal = calculateTax;
calculateTax = (function(prev) {
  return function(e) {
    prev(e);
    // Regime bar chart rendered after compute
    setTimeout(() => {
      if (document.getElementById('compareResult').style.display !== 'none') {
        // bar chart data is set by calculateTax — read from compare table
        const rows = document.querySelectorAll('#compareBody tr');
        if (rows.length >= 3) {
          // parse from DOM (simpler than re-running compute)
          document.getElementById('regimeChartWrap').style.display = 'block';
        }
      }
    }, 500);
  };
})(calculateTax);

/* Store last comparison data for bar chart */
const _origCalculateTaxInner = window.calculateTax;

</script>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>

<button class="help-btn" onclick="openHelp()" title="How to use this tool">?</button>
<div class="help-overlay" id="helpOverlay">
  <div class="help-modal">
    <div class="help-modal-head"><h3>How to Use — Income Tax Calculator</h3><button class="help-close" onclick="closeHelp()">&#10005;</button></div>
    <div class="help-modal-body"><div class="help-step"><div class="help-step-num">1</div><div class="help-step-body"><h4>Select Year</h4><p>Choose the Assessment Year and assessee type (Individual, HUF, Firm, Company, etc.).</p></div></div><div class="help-step"><div class="help-step-num">2</div><div class="help-step-body"><h4>Enter Income</h4><p>Fill income under Salary, House Property, Business/Profession, Capital Gains, and Other Sources.</p></div></div><div class="help-step"><div class="help-step-num">3</div><div class="help-step-body"><h4>Add Deductions</h4><p>Enter 80C, 80D, HRA, and other deductions (applicable under old regime).</p></div></div><div class="help-step"><div class="help-step-num">4</div><div class="help-step-body"><h4>View Result</h4><p>Tax under old and new regime is compared automatically side by side.</p></div></div><div class="help-step"><div class="help-step-num">5</div><div class="help-step-body"><h4>Advance Tax</h4><p>Scroll down to see the quarterly advance tax schedule.</p></div></div><div class="help-tip">⚠️ For estimation only. Verify with the latest CBDT notifications and consult a CA for complex cases.</div></div>
  </div>
</div>
<script>function openHelp(){document.getElementById('helpOverlay').classList.add('open')}function closeHelp(){document.getElementById('helpOverlay').classList.remove('open')}document.getElementById('helpOverlay').addEventListener('click',function(e){if(e.target===this)closeHelp()})</script>
</body></html>"""


# ══════════════════════════════════════════════════════════════════════════════
#  TDS CALCULATOR TEMPLATE
# ══════════════════════════════════════════════════════════════════════════════

TDS_CALC_T = r"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>TDS / TCS Calculator (IT Act 2025) – CA Toolkit</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>
<style>
""" + BASE_CSS + """
.nav-links{display:flex;gap:20px;list-style:none}
.nav-links a{text-decoration:none;color:var(--muted);font-size:13px;font-weight:500}
.nav-links a:hover{color:var(--brand)}
.hero{text-align:center;padding:32px 24px 16px;max-width:760px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#ECFDF5;color:#065F46;
            border:1px solid #A7F3D0;border-radius:99px;padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:12px}
h1{font-size:clamp(20px,4vw,32px);font-weight:800;line-height:1.15;letter-spacing:-.5px;margin-bottom:8px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:13px;color:var(--muted);line-height:1.7;max-width:520px;margin:0 auto}
.act-note{max-width:1100px;margin:0 auto;padding:0 24px 10px}
.act-box{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;padding:9px 14px;
         font-size:12px;color:#1e40af;display:flex;align-items:flex-start;gap:6px}
/* Toggle */
.toggle-wrap{max-width:1100px;margin:0 auto;padding:0 24px 16px;display:flex;gap:10px}
.toggle-btn{flex:1;padding:12px;border-radius:10px;border:2px solid var(--border);
            font-family:inherit;font-size:14px;font-weight:700;cursor:pointer;
            background:var(--white);color:var(--muted);transition:all .2s}
.toggle-btn.active{background:var(--brand);color:#fff;border-color:var(--brand)}
.toggle-btn:hover:not(.active){border-color:var(--brand);color:var(--brand)}
/* Layout */
.main{max-width:1100px;margin:0 auto;padding:0 24px 48px;
      display:grid;grid-template-columns:1.1fr 1fr;gap:20px;align-items:start}
@media(max-width:800px){.main{grid-template-columns:1fr}}
.card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);
      box-shadow:var(--shadow);overflow:hidden;margin-bottom:16px}
.card:last-child{margin-bottom:0}
.card-head{padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.card-head .icon{width:30px;height:30px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:15px}
.card-head h2{font-size:14px;font-weight:700}
.card-head p{font-size:12px;color:var(--muted);margin-top:1px}
.card-body{padding:16px}
.field{margin-bottom:13px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);margin-bottom:4px}
.hint{font-size:11px;color:var(--muted);margin-top:3px}
select,input[type=number],input[type=date]{width:100%;border:1.5px solid var(--border);border-radius:8px;
  padding:8px 11px;font-family:inherit;font-size:13px;color:var(--ink);background:var(--white);
  transition:border-color .2s;outline:none}
select:focus,input:focus{border-color:var(--brand)}
.btn{width:100%;background:var(--brand);color:#fff;border:none;border-radius:8px;
     padding:11px;font-family:inherit;font-size:14px;font-weight:700;cursor:pointer;transition:background .2s}
.btn:hover{background:var(--brand-d)}
/* Results */
.result-section{display:none;margin-top:14px}
.rboxes{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px}
.rbox{border-radius:10px;padding:14px 16px}
.rbox-main {background:#EFF6FF;border:1.5px solid #BFDBFE}
.rbox-int  {background:#FFFBEB;border:1.5px solid #FDE68A}
.rbox-total{background:#1D4ED8;border:1.5px solid #1D4ED8;grid-column:1/-1}
.rbox .val {font-size:22px;font-weight:800;margin-bottom:2px}
.rbox .lbl {font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:.75}
.rbox .sub {font-size:11px;margin-top:5px;opacity:.8}
.rbox-main  .val{color:#1D4ED8}.rbox-main  .lbl{color:#1D4ED8}
.rbox-int   .val{color:#92400E}.rbox-int   .lbl{color:#92400E}
.rbox-total .val{color:#fff;font-size:26px}
.rbox-total .lbl{color:rgba(255,255,255,.75)}
.rbox-total .sub{color:rgba(255,255,255,.8);font-size:12px}
.detail-table{width:100%;border-collapse:collapse;font-size:12px;margin-top:10px}
.detail-table td{padding:7px 2px;border-bottom:1px solid var(--border)}
.detail-table tr:last-child td{border:none;font-weight:700;font-size:13px}
.detail-table td:last-child{text-align:right;font-weight:600}
.ontime-box{background:#ECFDF5;border:1.5px solid #A7F3D0;border-radius:8px;
            padding:12px 14px;font-size:13px;color:#065F46;font-weight:600;margin-top:14px;text-align:center}
.note-box{background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;
          padding:10px 12px;font-size:11px;color:#92400E;margin-top:10px;line-height:1.6}
.info-box{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;
          padding:10px 12px;font-size:11px;color:#1e40af;margin-top:10px;line-height:1.6}
/* Rate tables */
.rate-table{width:100%;border-collapse:collapse;font-size:11px}
.rate-table th{text-align:left;font-size:10px;letter-spacing:.06em;text-transform:uppercase;
               color:var(--muted);border-bottom:1.5px solid var(--border);padding:5px 6px}
.rate-table td{padding:6px;border-bottom:1px solid var(--border);vertical-align:top;line-height:1.5}
.rate-table tr:last-child td{border:none}
.rate-table tr:hover td{background:#F9FAFB}
.code{background:#EFF6FF;color:var(--brand);font-size:10px;font-weight:700;
      padding:1px 5px;border-radius:4px;font-family:monospace;white-space:nowrap}
.tcs-code{background:#F5F3FF;color:#5B21B6;font-size:10px;font-weight:700;
          padding:1px 5px;border-radius:4px;font-family:monospace;white-space:nowrap}
footer{background:#0f1b2d;color:#9CA3AF;font-size:12px;padding:0}
.ft-main{display:grid;grid-template-columns:2fr 1fr 1.4fr;gap:40px;padding:40px 48px;max-width:1200px;margin:0 auto}
.ft-brand-name{color:#fff;font-size:18px;font-weight:800;margin-bottom:12px}
.ft-brand-desc{font-size:12.5px;line-height:1.75;color:#9CA3AF;max-width:340px;text-align:justify}
.ft-col-title{color:#fff;font-size:14px;font-weight:700;margin-bottom:14px}
.ft-links{list-style:none;padding:0;margin:0}
.ft-links li{margin-bottom:8px}
.ft-links a{color:#9CA3AF;text-decoration:none;font-size:13px;transition:color .2s}
.ft-links a:hover{color:#fff}
.ft-contact-name{color:#fff;font-weight:700;font-size:13px;margin-bottom:6px}
.ft-contact-addr{color:#9CA3AF;font-size:12px;line-height:1.7;margin-bottom:10px}
.ft-contact-line{color:#9CA3AF;font-size:12px;margin-bottom:4px}
.ft-socials{display:flex;gap:14px;margin-top:12px}
.ft-socials a{color:#9CA3AF;transition:color .2s}
.ft-socials a:hover{color:#fff}
.ft-socials svg{width:20px;height:20px;fill:currentColor}
.ft-bottom{background:#0a1422;border-top:1px solid #1e2d42;padding:12px 48px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.ft-bottom-left{font-size:11px;color:#6B7280}
.ft-bottom-right{font-size:11px;color:#6B7280}
@media(max-width:768px){.ft-main{grid-template-columns:1fr;padding:28px 20px;gap:24px}.ft-bottom{padding:12px 20px;flex-direction:column;text-align:center}}
</style></head><body>

<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <ul class="nav-links"><li><a href="/">← All Tools</a></li></ul>
  <div class="nav-right">
    {% if username %}<span class="nav-user">👤 <strong>{{ username }}</strong></span>
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin</a>{% endif %}
    <a href="/logout" class="nav-link">Sign out</a>
    {% else %}<a href="/login" class="nav-btn">Sign In</a>{% endif %}
  </div>
</nav>

<section class="hero">
  <div class="hero-badge">🆓 Free · No Login Required</div>
  <h1>TDS / TCS Calculator — <em>IT Act 2025</em></h1>
  <p>Calculate TDS or TCS liability, late deposit interest and total payable amount as per new Section 393 / Section 394 of IT Act 2025.</p>
</section>

<div class="act-note">
  <div class="act-box">ℹ️ <span><strong>IT Act 2025 (w.e.f. 1 Apr 2026):</strong> TDS consolidated under Section 393 (non-salary) &amp; Section 392 (salary). TCS under Section 394. Numeric payment codes replace old section numbers in returns. Rates &amp; thresholds unchanged.</span></div>
</div>

<!-- TDS / TCS TOGGLE -->
<div class="toggle-wrap">
  <button class="toggle-btn active" id="btnTDS" onclick="switchMode('tds')">📑 TDS — Tax Deducted at Source</button>
  <button class="toggle-btn" id="btnTCS" onclick="switchMode('tcs')">🧾 TCS — Tax Collected at Source</button>
</div>

<div class="main">
  <!-- LEFT: INPUT -->
  <div>
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#EFF6FF" id="formIcon">📑</div>
        <div>
          <h2 id="formTitle">TDS Calculator</h2>
          <p id="formSub">IT Act 2025 · Tax Year 2026-27</p>
        </div>
      </div>
      <div class="card-body">

        <div class="field">
          <label id="sectionLabel">Nature of Payment (TDS)</label>
          <select id="mainSection" onchange="updateHint()">
            <option value="">— Select Payment Type —</option>
          </select>
          <p class="hint" id="sectionHint">Select a payment type to see rate and threshold</p>
        </div>

        <div class="field">
          <label id="amtLabel">Payment Amount (₹)</label>
          <input type="number" id="paymentAmt" placeholder="e.g. 100000" min="0"/>
          <p class="hint" id="amtHint">Gross payment amount before TDS deduction</p>
        </div>

        <hr style="border:none;border-top:1.5px dashed var(--border);margin:14px 0"/>
        <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-bottom:12px">Late Deposit Interest (Optional)</p>

        <div class="row2">
          <div class="field">
            <label id="d1label">Date of Deduction</label>
            <input type="date" id="deductionDate"/>
            <p class="hint" id="d1hint">When TDS was deducted</p>
          </div>
          <div class="field">
            <label>Date of Actual Deposit</label>
            <input type="date" id="depositDate"/>
            <p class="hint">When you paid the challan</p>
          </div>
        </div>

        <button class="btn" id="calcBtn" onclick="calculate()">Calculate TDS &amp; Interest →</button>

        <!-- RESULTS -->
        <div class="result-section" id="resultSection">
          <div class="ontime-box" id="ontimeBox" style="display:none"></div>
          <div id="lateBoxes" style="display:none">
            <div class="rboxes">
              <div class="rbox rbox-main">
                <div class="lbl" id="r-main-lbl">TDS Amount</div>
                <div class="val" id="r-main"></div>
                <div class="sub" id="r-main-sub"></div>
              </div>
              <div class="rbox rbox-int">
                <div class="lbl">Interest u/s 201(1A)</div>
                <div class="val" id="r-int"></div>
                <div class="sub" id="r-int-sub"></div>
              </div>
              <div class="rbox rbox-total">
                <div class="lbl">Total Amount Payable</div>
                <div class="val" id="r-total"></div>
                <div class="sub" id="r-total-sub"></div>
              </div>
            </div>
            <table class="detail-table">
              <tr><td id="d-amt-lbl">Payment Amount</td><td id="d-payment"></td></tr>
              <tr><td>New Section (IT Act 2025)</td><td id="d-newsec"></td></tr>
              <tr><td>Old Section (for reference)</td><td id="d-oldsec"></td></tr>
              <tr><td>Payment Code</td><td id="d-code"></td></tr>
              <tr><td id="d-rate-lbl">TDS Rate</td><td id="d-rate"></td></tr>
              <tr><td id="d-tax-lbl">TDS Amount</td><td id="d-tds"></td></tr>
              <tr><td id="d-date1-lbl">Date of Deduction</td><td id="d-ddate"></td></tr>
              <tr><td>Due Date for Deposit</td><td id="d-due"></td></tr>
              <tr><td>Actual Deposit Date</td><td id="d-adate"></td></tr>
              <tr><td>Delay (months)</td><td id="d-months"></td></tr>
              <tr><td>Interest Rate</td><td>1.5% per month</td></tr>
              <tr><td>Interest Amount</td><td id="d-intamt"></td></tr>
              <tr><td id="d-total-lbl" style="color:var(--brand)">Total Payable</td><td id="d-total" style="color:var(--brand)"></td></tr>
            </table>
            <div class="note-box">⚠ As per IT Act 2025, a fractional month is counted as a full month for interest calculation. Interest runs from date of deduction/collection to actual date of deposit.</div>
          </div>
          <div id="basicBox" style="display:none">
            <div class="rboxes">
              <div class="rbox rbox-main" style="grid-column:1/-1">
                <div class="lbl" id="b-main-lbl">TDS Amount</div>
                <div class="val" id="b-main"></div>
                <div class="sub" id="b-sub"></div>
              </div>
            </div>
            <table class="detail-table">
              <tr><td id="b-amt-lbl">Payment Amount</td><td id="b-payment"></td></tr>
              <tr><td>New Section (IT Act 2025)</td><td id="b-newsec"></td></tr>
              <tr><td>Old Section (for reference)</td><td id="b-oldsec"></td></tr>
              <tr><td>Payment Code</td><td id="b-code"></td></tr>
              <tr><td id="b-rate-lbl">TDS Rate</td><td id="b-rate"></td></tr>
              <tr><td id="b-tax-lbl">TDS Amount</td><td id="b-tds2"></td></tr>
              <tr><td id="b-net-lbl">Net Payment to Payee</td><td id="b-net"></td></tr>
            </table>
            <div class="info-box">ℹ Enter deduction and deposit dates above to also calculate late deposit interest.</div>
          </div>
        </div>

      </div>
    </div>
  </div>

  <!-- RIGHT: RATE CHARTS + DUE DATES -->
  <div>
    <!-- TDS rate chart -->
    <div id="tdsRateCard" class="card">
      <div class="card-head">
        <div class="icon" style="background:#FFFBEB">📋</div>
        <div><h2>TDS Quick Rate Chart</h2><p>Section 393 — IT Act 2025</p></div>
      </div>
      <div class="card-body" style="padding:0;overflow-x:auto">
        <table class="rate-table">
          <thead><tr><th>Code</th><th>Old Sec</th><th>Nature</th><th>Rate</th><th>Threshold</th></tr></thead>
          <tbody>
            <tr><td><span class="code">1001</span></td><td>192</td><td>Salary</td><td>Slab</td><td>Basic exemption</td></tr>
            <tr><td><span class="code">1021</span></td><td>194A</td><td>Interest (Bank/PO)</td><td>10%</td><td>₹50,000</td></tr>
            <tr><td><span class="code">1022</span></td><td>194A</td><td>Interest (Others)</td><td>10%</td><td>₹10,000</td></tr>
            <tr><td><span class="code">1023</span></td><td>194C</td><td>Contractor (Ind)</td><td>1%</td><td>₹30K/₹1L pa</td></tr>
            <tr><td><span class="code">1024</span></td><td>194C</td><td>Contractor (Others)</td><td>2%</td><td>₹30K/₹1L pa</td></tr>
            <tr><td><span class="code">1006</span></td><td>194H</td><td>Commission/Brokerage</td><td>2%</td><td>₹20,000</td></tr>
            <tr><td><span class="code">1008</span></td><td>194I(a)</td><td>Rent (P&amp;M)</td><td>2%</td><td>₹50K/mo</td></tr>
            <tr><td><span class="code">1009</span></td><td>194I(b)</td><td>Rent (Land/Bldg)</td><td>10%</td><td>₹50K/mo</td></tr>
            <tr><td><span class="code">1036</span></td><td>194IA</td><td>Immovable Property</td><td>1%</td><td>₹50L</td></tr>
            <tr><td><span class="code">1027</span></td><td>194J(b)</td><td>Professional Fees</td><td>10%</td><td>₹50,000</td></tr>
            <tr><td><span class="code">1026</span></td><td>194J(a)</td><td>Technical Services</td><td>2%</td><td>₹50,000</td></tr>
            <tr><td><span class="code">1031</span></td><td>194Q</td><td>Purchase of Goods</td><td>0.1%</td><td>₹50L pa</td></tr>
            <tr><td><span class="code">1039</span></td><td>194S</td><td>VDA/Crypto</td><td>1%</td><td>₹10,000</td></tr>
            <tr><td><span class="code">1041</span></td><td>194T</td><td>Partner Salary</td><td>10%</td><td>₹20,000</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- TCS rate chart -->
    <div id="tcsRateCard" class="card" style="display:none">
      <div class="card-head">
        <div class="icon" style="background:#F5F3FF">🧾</div>
        <div><h2>TCS Quick Rate Chart</h2><p>Section 394 — IT Act 2025</p></div>
      </div>
      <div class="card-body" style="padding:0;overflow-x:auto">
        <table class="rate-table">
          <thead><tr><th>Code</th><th>Old Sec</th><th>Nature of Goods/Transaction</th><th>Rate</th><th>Threshold</th></tr></thead>
          <tbody>
            <tr><td><span class="tcs-code">2001</span></td><td>206C(1)</td><td>Alcoholic Liquor for Human Consumption</td><td>2%</td><td>Nil</td></tr>
            <tr><td><span class="tcs-code">2002</span></td><td>206C(1)</td><td>Tendu Leaves</td><td>2%</td><td>Nil</td></tr>
            <tr><td><span class="tcs-code">2003</span></td><td>206C(1)</td><td>Timber (forest lease)</td><td>2%</td><td>Nil</td></tr>
            <tr><td><span class="tcs-code">2004</span></td><td>206C(1)</td><td>Timber (other than forest lease)</td><td>2%</td><td>Nil</td></tr>
            <tr><td><span class="tcs-code">2005</span></td><td>206C(1)</td><td>Any other forest produce</td><td>2%</td><td>Nil</td></tr>
            <tr><td><span class="tcs-code">2006</span></td><td>206C(1)</td><td>Scrap</td><td>2%</td><td>Nil</td></tr>
            <tr><td><span class="tcs-code">2007</span></td><td>206C(1)</td><td>Minerals (coal/lignite/iron ore)</td><td>2%</td><td>Nil</td></tr>
            <tr><td><span class="tcs-code">2009</span></td><td>206C(1F)</td><td>Motor Vehicle &gt; ₹10L</td><td>1%</td><td>₹10L</td></tr>
            <tr><td><span class="tcs-code">2010</span></td><td>206C(1G)</td><td>Foreign Remittance (LRS) — Education/Medical</td><td>2%</td><td>₹10L</td></tr>
            <tr><td><span class="tcs-code">2011</span></td><td>206C(1G)</td><td>Foreign Remittance (LRS) — Other purposes</td><td>20%</td><td>₹10L</td></tr>
            <tr><td><span class="tcs-code">2012</span></td><td>206C(1G)</td><td>Overseas Tour Package</td><td>2%</td><td>Nil</td></tr>
            <tr><td><span class="tcs-code">2013</span></td><td>206C(1H)</td><td>Sale of Goods &gt; ₹50L</td><td>0.1%</td><td>₹50L pa</td></tr>
            <tr><td><span class="tcs-code">2014</span></td><td>206C(1)</td><td>Parking lot / Toll Plaza / Mining &amp; Quarrying</td><td>2%</td><td>Nil</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Due dates -->
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#F0FDF4">📅</div>
        <div><h2 id="dueDateTitle">TDS Deposit Due Dates</h2><p>Rule 218 — IT Rules 2026</p></div>
      </div>
      <div class="card-body">
        <div id="tdsduedates" style="font-size:12px;line-height:2;color:var(--muted)">
          <p><strong style="color:var(--ink)">April – February:</strong> 7th of the following month</p>
          <p><strong style="color:var(--ink)">March deductions:</strong> 30th April</p>
          <p><strong style="color:var(--ink)">Sec 194IA/194IB/194M/194S:</strong> 30 days from end of deduction month</p>
          <p style="margin-top:8px;color:var(--red)"><strong>Late interest:</strong> 1.5% per month · Fractional month = full month</p>
        </div>
        <div id="tcsduedates" style="display:none;font-size:12px;line-height:1.9;color:var(--muted)">
          <div style="margin-bottom:10px;padding:8px 12px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;font-size:11px;color:#065F46">
            <strong>IT Act 2025 Reference:</strong> Section 394 (TCS) · Rule 219 (IT Rules 2026) · Challan 281
          </div>
          <p><strong style="color:var(--ink)">📅 Deposit Due Date:</strong></p>
          <p>Collections <strong>April – February</strong> → <strong>7th of following month</strong></p>
          <p>Collections in <strong>March</strong> → <strong>7th April</strong> of next FY</p>
          <div style="margin:10px 0;border-top:1px solid var(--border)"></div>
          <p><strong style="color:var(--ink)">🗓️ Quarterly Return — Form 27EQ:</strong></p>
          <table style="width:100%;border-collapse:collapse;font-size:11px;margin:6px 0">
            <thead><tr style="background:#F9FAFB">
              <th style="padding:5px 8px;border:1px solid var(--border);text-align:left">Quarter</th>
              <th style="padding:5px 8px;border:1px solid var(--border);text-align:left">Period</th>
              <th style="padding:5px 8px;border:1px solid var(--border);text-align:left">Due Date</th>
            </tr></thead>
            <tbody>
              <tr><td style="padding:5px 8px;border:1px solid var(--border)">Q1</td><td style="padding:5px 8px;border:1px solid var(--border)">Apr – Jun</td><td style="padding:5px 8px;border:1px solid var(--border);font-weight:600;color:var(--ink)">15th July</td></tr>
              <tr><td style="padding:5px 8px;border:1px solid var(--border)">Q2</td><td style="padding:5px 8px;border:1px solid var(--border)">Jul – Sep</td><td style="padding:5px 8px;border:1px solid var(--border);font-weight:600;color:var(--ink)">15th October</td></tr>
              <tr><td style="padding:5px 8px;border:1px solid var(--border)">Q3</td><td style="padding:5px 8px;border:1px solid var(--border)">Oct – Dec</td><td style="padding:5px 8px;border:1px solid var(--border);font-weight:600;color:var(--ink)">15th January</td></tr>
              <tr><td style="padding:5px 8px;border:1px solid var(--border)">Q4</td><td style="padding:5px 8px;border:1px solid var(--border)">Jan – Mar</td><td style="padding:5px 8px;border:1px solid var(--border);font-weight:600;color:var(--ink)">15th May</td></tr>
            </tbody>
          </table>
          <div style="margin-top:10px;padding:8px 12px;background:#FEF2F2;border:1px solid #FECACA;border-radius:8px;font-size:11px">
            <strong style="color:#991B1B">⚠ Default Consequences:</strong><br>
            <span style="color:#991B1B">Non-collection/deposit: Interest <strong>1%/month</strong> u/s 394(6) · Fractional month = full month<br>
            Late 27EQ filing: <strong>₹200/day</strong> u/s 267 (max = TCS amount)</span>
          </div>
        </div>
      </div>
    </div>
  </div>
</div>

<footer>
  <div class="ft-main">
    <div>
      <div class="ft-brand-name">CA Toolkit</div>
      <p class="ft-brand-desc">CA Toolkit is a comprehensive utility platform built by a CA Article from Ludhiana, Punjab, providing automation tools for Indian Chartered Accountants. The platform saves hours of manual work every year — from Balance Sheet year-shift to GST reconciliation, tax calculations, and more.</p>
    </div>
    <div>
      <div class="ft-col-title">Know More</div>
      <ul class="ft-links">
        <li><a href="/">Home</a></li>
        <li><a href="/">BS Year Shift</a></li>
        <li><a href="/tool/tb-to-bs">TB → Balance Sheet</a></li>
        <li><a href="/tool/tax-calculator">Tax Calculator</a></li>
        <li><a href="/privacy">Privacy Policy</a></li>
      </ul>
    </div>
    <div>
      <div class="ft-col-title">Contact Us</div>
      <div class="ft-contact-name">CA Toolkit</div>
      <div class="ft-contact-addr">Built for Indian Chartered Accountants<br/>Created by CA Article · Ludhiana, Punjab</div>
    </div>
  </div>
  <div class="ft-bottom">
    <span class="ft-bottom-left">©2026 CA Toolkit · All Rights Reserved</span>
    <span class="ft-bottom-right">Built for Indian CAs · Ludhiana, Punjab</span>
  </div>
</footer>

<script>
// ── DATA ─────────────────────────────────────────────────────────────────────

const TDS_DATA = {
  "1001":{rate:0,    thresh:0,       label:"Salary",                       newSec:"Sec 392",                    oldSec:"Sec 192",     note:"Slab rate"},
  "1004":{rate:10,   thresh:50000,   label:"PF Accumulated Balance",        newSec:"Sec 392(7)",                 oldSec:"Sec 192A",    note:"No PAN: 20%"},
  "1005":{rate:2,    thresh:20000,   label:"Insurance Commission",          newSec:"Sec 393(1) Sl.1(i)",         oldSec:"Sec 194D",    note:"Ind: 2%, Others: 10%"},
  "1006":{rate:2,    thresh:20000,   label:"Commission/Brokerage",          newSec:"Sec 393(1) Sl.1(ii)",        oldSec:"Sec 194H",    note:""},
  "1008":{rate:2,    thresh:50000,   label:"Rent – Machinery/Plant",        newSec:"Sec 393(1) Sl.2(ii).D(a)",  oldSec:"Sec 194I(a)", note:"Monthly threshold"},
  "1009":{rate:10,   thresh:50000,   label:"Rent – Land/Building",          newSec:"Sec 393(1) Sl.2(ii).D(b)",  oldSec:"Sec 194I(b)", note:"Monthly threshold"},
  "1010":{rate:2,    thresh:50000,   label:"Rent by Ind/HUF",               newSec:"Sec 393(1) Sl.2(i)",         oldSec:"Sec 194IB",   note:"Per month. Reduced from 5% to 2%"},
  "1011":{rate:10,   thresh:0,       label:"JDA Consideration",             newSec:"Sec 393(1) Sl.3(ii)",        oldSec:"Sec 194IC",   note:""},
  "1012":{rate:10,   thresh:500000,  label:"Land Acquisition Comp.",        newSec:"Sec 393(1) Sl.3(iii)",       oldSec:"Sec 194LA",   note:"Threshold ₹5L"},
  "1013":{rate:10,   thresh:10000,   label:"Mutual Fund Units",             newSec:"Sec 393(1) Sl.4(i)",         oldSec:"Sec 194K",    note:""},
  "1014":{rate:10,   thresh:0,       label:"Business Trust – Interest",     newSec:"Sec 393(1) Sl.4(ii)",        oldSec:"Sec 194LBA",  note:""},
  "1017":{rate:10,   thresh:0,       label:"Investment Fund Income",        newSec:"Sec 393(1) Sl.4(iii)",       oldSec:"Sec 194LBB",  note:""},
  "1018":{rate:10,   thresh:0,       label:"Securitisation Trust",          newSec:"Sec 393(1) Sl.4(iv)",        oldSec:"Sec 194LBC",  note:""},
  "1019":{rate:10,   thresh:10000,   label:"Interest on Securities",        newSec:"Sec 393(1) Sl.5(i)",         oldSec:"Sec 193",     note:""},
  "1020":{rate:10,   thresh:100000,  label:"Interest – Senior Citizen",     newSec:"Sec 393(1) Sl.5(ii).D(a)",  oldSec:"Sec 194A",    note:"Threshold ₹1L"},
  "1021":{rate:10,   thresh:50000,   label:"Interest – Bank/Post Office",   newSec:"Sec 393(1) Sl.5(ii).D(b)",  oldSec:"Sec 194A",    note:"Threshold ₹50K"},
  "1022":{rate:10,   thresh:10000,   label:"Interest – Others",             newSec:"Sec 393(1) Sl.5(iii)",       oldSec:"Sec 194A",    note:"Threshold ₹10K"},
  "1023":{rate:1,    thresh:30000,   label:"Contractor – Ind/HUF",          newSec:"Sec 393(1) Sl.6(i).D(a)",   oldSec:"Sec 194C",    note:"Single ₹30K / Annual ₹1L"},
  "1024":{rate:2,    thresh:30000,   label:"Contractor – Others",           newSec:"Sec 393(1) Sl.6(i).D(b)",   oldSec:"Sec 194C",    note:"Single ₹30K / Annual ₹1L"},
  "1026":{rate:2,    thresh:50000,   label:"Technical Services/Royalty",    newSec:"Sec 393(1) Sl.6(iii).D(a)", oldSec:"Sec 194J(a)", note:""},
  "1027":{rate:10,   thresh:50000,   label:"Professional Fees",             newSec:"Sec 393(1) Sl.6(iii).D(b)", oldSec:"Sec 194J(b)", note:""},
  "1028":{rate:10,   thresh:0,       label:"Director Remuneration",         newSec:"Sec 393(1) Sl.6(iii).D(b)", oldSec:"Sec 194J(b)", note:"No threshold"},
  "1029":{rate:10,   thresh:10000,   label:"Dividends",                     newSec:"Sec 393(1) Sl.7",            oldSec:"Sec 194",     note:""},
  "1030":{rate:2,    thresh:100000,  label:"Life Insurance Proceeds",       newSec:"Sec 393(1) Sl.8(i)",         oldSec:"Sec 194DA",   note:"On taxable portion"},
  "1031":{rate:0.1,  thresh:5000000, label:"Purchase of Goods",             newSec:"Sec 393(1) Sl.8(ii)",        oldSec:"Sec 194Q",    note:"Annual > ₹50L"},
  "1033":{rate:10,   thresh:20000,   label:"Benefit/Perquisite",            newSec:"Sec 393(1) Sl.8(iv)",        oldSec:"Sec 194R",    note:""},
  "1035":{rate:0.1,  thresh:500000,  label:"E-Commerce Operator",           newSec:"Sec 393(1) Sl.8(vi)",        oldSec:"Sec 194O",    note:"Annual > ₹5L"},
  "1036":{rate:1,    thresh:5000000, label:"Purchase of Immovable Property",newSec:"Sec 393(1) Sl.3(i)",         oldSec:"Sec 194IA",   note:"Threshold ₹50L"},
  "1037":{rate:5,    thresh:5000000, label:"Contractor/Prof by Ind/HUF",    newSec:"Sec 393(1) Sl.6(iv)",        oldSec:"Sec 194M",    note:"Annual > ₹50L"},
  "1038":{rate:2,    thresh:2000000, label:"Cash Withdrawal",               newSec:"Sec 393(1) Sl.8(vii)",       oldSec:"Sec 194N",    note:"3% if no ITR filed"},
  "1039":{rate:1,    thresh:10000,   label:"VDA/Crypto",                    newSec:"Sec 393(1) Sl.8(viii)",      oldSec:"Sec 194S",    note:"₹50K for specified persons"},
  "1040":{rate:30,   thresh:10000,   label:"Lottery/Puzzle Winnings",       newSec:"Sec 393(1) Sl.8(ix)",        oldSec:"Sec 194B",    note:""},
  "1041":{rate:10,   thresh:20000,   label:"Partner Salary/Remuneration",   newSec:"Sec 393(1) Sl.6(v)",         oldSec:"Sec 194T",    note:"Threshold ₹20K pa"},
};

const TCS_DATA = {
  "2001":{rate:2,    thresh:0,        label:"Alcoholic Liquor for Human Consumption", newSec:"Sec 394(1)(i)",   oldSec:"Sec 206C(1)(a)",  note:"Increased from 1% to 2% w.e.f. 01.04.2026"},
  "2002":{rate:2,    thresh:0,        label:"Tendu Leaves",                           newSec:"Sec 394(1)(ii)",  oldSec:"Sec 206C(1)(b)",  note:"Reduced from 5% to 2% w.e.f. 01.04.2026"},
  "2003":{rate:2,    thresh:0,        label:"Timber – Forest Lease",                  newSec:"Sec 394(1)(iii)", oldSec:"Sec 206C(1)(c)",  note:"Reduced from 2.5% to 2% w.e.f. 01.04.2026"},
  "2004":{rate:2,    thresh:0,        label:"Timber – Other than Forest Lease",       newSec:"Sec 394(1)(iv)",  oldSec:"Sec 206C(1)(d)",  note:"Reduced from 2.5% to 2% w.e.f. 01.04.2026"},
  "2005":{rate:2,    thresh:0,        label:"Any Other Forest Produce",               newSec:"Sec 394(1)(v)",   oldSec:"Sec 206C(1)(e)",  note:"Reduced from 2.5% to 2% w.e.f. 01.04.2026"},
  "2006":{rate:2,    thresh:0,        label:"Scrap",                                  newSec:"Sec 394(1)(vi)",  oldSec:"Sec 206C(1)(f)",  note:"Increased from 1% to 2% w.e.f. 01.04.2026"},
  "2007":{rate:2,    thresh:0,        label:"Minerals (Coal/Lignite/Iron Ore)",       newSec:"Sec 394(1)(vii)", oldSec:"Sec 206C(1)(g)",  note:"Increased from 1% to 2% w.e.f. 01.04.2026"},
  "2009":{rate:1,    thresh:1000000,  label:"Motor Vehicle > ₹10L",                   newSec:"Sec 394(1F)",     oldSec:"Sec 206C(1F)",    note:"On sale consideration"},
  "2010":{rate:2,    thresh:1000000,  label:"Foreign Remittance (LRS) – Education/Medical > ₹10L",newSec:"Sec 394(1G)(i)",oldSec:"Sec 206C(1G)",   note:"Reduced from 5% to 2%. Nil if loan from bank. Threshold now ₹10L"},
  "2011":{rate:20,   thresh:1000000,  label:"Foreign Remittance (LRS) – Other > ₹10L",newSec:"Sec 394(1G)(ii)", oldSec:"Sec 206C(1G)",    note:"20% above ₹10L. Threshold changed from ₹7L to ₹10L"},
  "2012":{rate:2,    thresh:0,        label:"Overseas Tour Package",                  newSec:"Sec 394(1G)(iii)",oldSec:"Sec 206C(1G)",    note:"Flat 2% (was 5%/20%). Threshold removed w.e.f. 01.04.2026"},
  "2013":{rate:0.1,  thresh:5000000,  label:"Sale of Goods > ₹50L",                   newSec:"Sec 394(1H)",     oldSec:"Sec 206C(1H)",    note:"Annual turnover > ₹10Cr"},
  "2014":{rate:2,    thresh:0,        label:"Parking Lot / Toll Plaza / Mining",       newSec:"Sec 394(1)(viii)",oldSec:"Sec 206C(1)(h)",  note:""},
};

const TDS_SPECIAL_30 = ["1036","1010","1037","1039"];

let currentMode = "tds";

// ── Build dropdowns ───────────────────────────────────────────────────────────

function buildTDSOptions(){
  return `<option value="">— Select Payment Type —</option>
    <optgroup label="── Salary ──">
      <option value="1001">Salary (Sec 392) — Slab rate</option>
      <option value="1004">PF Accumulated Balance — 10%</option>
    </optgroup>
    <optgroup label="── Commission &amp; Brokerage ──">
      <option value="1005">Insurance Commission (Old: 194D) — 2%</option>
      <option value="1006">Commission / Brokerage (Old: 194H) — 2%</option>
    </optgroup>
    <optgroup label="── Rent ──">
      <option value="1008">Rent – Machinery/Plant (Old: 194I(a)) — 2%</option>
      <option value="1009">Rent – Land/Building (Old: 194I(b)) — 10%</option>
      <option value="1010">Rent by Individual/HUF (Old: 194IB) — 2%</option>
    </optgroup>
    <optgroup label="── Property ──">
      <option value="1011">JDA Consideration (Old: 194IC) — 10%</option>
      <option value="1012">Compensation – Land Acquisition (Old: 194LA) — 10%</option>
      <option value="1036">Purchase of Immovable Property (Old: 194IA) — 1%</option>
    </optgroup>
    <optgroup label="── Interest ──">
      <option value="1019">Interest on Securities (Old: 193) — 10%</option>
      <option value="1020">Interest – Senior Citizen (Old: 194A) — 10%</option>
      <option value="1021">Interest – Bank/Post Office (Old: 194A) — 10%</option>
      <option value="1022">Interest – Others (Old: 194A) — 10%</option>
    </optgroup>
    <optgroup label="── Investment Income ──">
      <option value="1013">Mutual Fund Units (Old: 194K) — 10%</option>
      <option value="1029">Dividends (Old: 194) — 10%</option>
      <option value="1014">Business Trust – Interest (Old: 194LBA) — 10%</option>
    </optgroup>
    <optgroup label="── Contractor &amp; Professional ──">
      <option value="1023">Contractor – Individual/HUF (Old: 194C) — 1%</option>
      <option value="1024">Contractor – Others/Company (Old: 194C) — 2%</option>
      <option value="1026">Technical Services/Royalty (Old: 194J(a)) — 2%</option>
      <option value="1027">Professional Fees (Old: 194J(b)) — 10%</option>
      <option value="1028">Director Remuneration (Old: 194J(b)) — 10%</option>
      <option value="1037">Contractor/Prof by Ind/HUF (Old: 194M) — 5%</option>
      <option value="1041">Partner Salary/Remuneration (Old: 194T) — 10%</option>
    </optgroup>
    <optgroup label="── Other Payments ──">
      <option value="1030">Life Insurance Proceeds (Old: 194DA) — 2%</option>
      <option value="1031">Purchase of Goods (Old: 194Q) — 0.1%</option>
      <option value="1033">Benefit/Perquisite (Old: 194R) — 10%</option>
      <option value="1035">E-Commerce Operator (Old: 194O) — 0.1%</option>
      <option value="1038">Cash Withdrawal (Old: 194N) — 2%</option>
      <option value="1039">VDA / Crypto (Old: 194S) — 1%</option>
      <option value="1040">Lottery/Puzzle Winnings (Old: 194B) — 30%</option>
    </optgroup>`;
}

function buildTCSOptions(){
  return `<option value="">— Select Nature of Goods/Transaction —</option>
    <optgroup label="── Goods (All rationalised to 2%) ──">
      <option value="2001">Alcoholic Liquor for Human Consumption — 2%</option>
      <option value="2002">Tendu Leaves — 2%</option>
      <option value="2003">Timber – Forest Lease — 2%</option>
      <option value="2004">Timber – Other than Forest Lease — 2%</option>
      <option value="2005">Any Other Forest Produce — 2%</option>
      <option value="2006">Scrap — 2%</option>
      <option value="2007">Minerals (Coal/Lignite/Iron Ore) — 2%</option>
      <option value="2014">Parking Lot / Toll Plaza / Mining &amp; Quarrying — 2%</option>
    </optgroup>
    <optgroup label="── High Value Transactions ──">
      <option value="2009">Motor Vehicle Sale &gt; ₹10 Lakh — 1%</option>
      <option value="2013">Sale of Goods &gt; ₹50L (Annual) — 0.1%</option>
    </optgroup>
    <optgroup label="── Foreign Remittance (LRS) ──">
      <option value="2010">Foreign Remittance – Education/Medical &gt; ₹10L — 2%</option>
      <option value="2011">Foreign Remittance – Other Purposes &gt; ₹10L — 20%</option>
      <option value="2012">Overseas Tour Package — 2% (flat)</option>
    </optgroup>`;
}

// ── Toggle mode ───────────────────────────────────────────────────────────────

function switchMode(mode){
  currentMode = mode;
  const sel = document.getElementById("mainSection");
  sel.innerHTML = mode==="tds" ? buildTDSOptions() : buildTCSOptions();
  document.getElementById("sectionHint").textContent = "Select a payment type to see rate and threshold";

  const isTDS = mode==="tds";
  document.getElementById("btnTDS").className = "toggle-btn"+(isTDS?" active":"");
  document.getElementById("btnTCS").className = "toggle-btn"+(!isTDS?" active":"");
  document.getElementById("formIcon").textContent     = isTDS?"📑":"🧾";
  document.getElementById("formTitle").textContent    = isTDS?"TDS Calculator":"TCS Calculator";
  document.getElementById("formSub").textContent      = isTDS?"IT Act 2025 · Tax Year 2026-27":"IT Act 2025 · Tax Year 2026-27";
  document.getElementById("sectionLabel").textContent = isTDS?"Nature of Payment (TDS)":"Nature of Goods / Transaction (TCS)";
  document.getElementById("amtLabel").textContent     = isTDS?"Payment Amount (₹)":"Sale / Collection Amount (₹)";
  document.getElementById("amtHint").textContent      = isTDS?"Gross payment amount before TDS deduction":"Gross sale/receipt amount before TCS collection";
  document.getElementById("d1label").textContent      = isTDS?"Date of Deduction":"Date of Collection";
  document.getElementById("d1hint").textContent       = isTDS?"When TDS was deducted":"When TCS was collected";
  document.getElementById("calcBtn").textContent      = isTDS?"Calculate TDS & Interest →":"Calculate TCS & Interest →";
  document.getElementById("r-main-lbl").textContent   = isTDS?"TDS Amount":"TCS Amount";
  document.getElementById("b-main-lbl").textContent   = isTDS?"TDS Amount":"TCS Amount";
  document.getElementById("d-rate-lbl").textContent   = isTDS?"TDS Rate":"TCS Rate";
  document.getElementById("d-tax-lbl").textContent    = isTDS?"TDS Amount":"TCS Amount";
  document.getElementById("d-date1-lbl").textContent  = isTDS?"Date of Deduction":"Date of Collection";
  document.getElementById("d-amt-lbl").textContent    = isTDS?"Payment Amount":"Sale Amount";
  document.getElementById("b-amt-lbl").textContent    = isTDS?"Payment Amount":"Sale Amount";
  document.getElementById("b-rate-lbl").textContent   = isTDS?"TDS Rate":"TCS Rate";
  document.getElementById("b-tax-lbl").textContent    = isTDS?"TDS Amount":"TCS Amount";
  document.getElementById("b-net-lbl").textContent    = isTDS?"Net Payment to Payee":"Amount Receivable from Buyer";
  document.getElementById("d-total-lbl").textContent  = isTDS?"Total Payable (TDS + Interest)":"Total Payable (TCS + Interest)";
  document.getElementById("tdsRateCard").style.display = isTDS?"block":"none";
  document.getElementById("tcsRateCard").style.display = !isTDS?"block":"none";
  document.getElementById("tdsduedates").style.display = isTDS?"block":"none";
  document.getElementById("tcsduedates").style.display = !isTDS?"block":"none";
  document.getElementById("dueDateTitle").textContent = isTDS?"TDS Deposit Due Dates":"TCS Deposit Due Dates";

  // Update interest section label for TCS
  const intLbl = document.getElementById("r-int");
  if(intLbl){
    const lbl = intLbl.closest(".rbox")?.querySelector(".lbl");
    if(lbl) lbl.textContent = isTDS?"Interest u/s 201(1A)":"Interest u/s 206C(7)";
  }

  document.getElementById("resultSection").style.display = "none";
}

// ── Hint update ───────────────────────────────────────────────────────────────

function updateHint(){
  const code = document.getElementById("mainSection").value;
  const el   = document.getElementById("sectionHint");
  if(!code){ el.textContent="Select a payment type to see rate and threshold"; return; }
  const data = currentMode==="tds" ? TDS_DATA : TCS_DATA;
  const d    = data[code];
  if(!d) return;
  el.textContent = (d.rate===0?"Rate: Slab rate":"Rate: "+d.rate+"%")
    + (d.thresh?" · Threshold: ₹"+Math.round(d.thresh).toLocaleString("en-IN"):" · No threshold")
    + (d.note?" · "+d.note:"");
}

// ── Due date calc ─────────────────────────────────────────────────────────────

function getDueDate(deductDate, code, mode){
  const d     = new Date(deductDate);
  const month = d.getMonth();
  const year  = d.getFullYear();
  if(mode==="tds" && TDS_SPECIAL_30.includes(code)){
    const endOfMonth = new Date(year, month+1, 0);
    return new Date(endOfMonth.getTime() + 30*24*60*60*1000);
  }
  if(month===2) return new Date(year, 3, 30); // March → 30 April
  return new Date(year, month+1, 7);          // Others → 7th next month
}

function calcMonthsLate(dueDate, depositDate){
  if(new Date(depositDate) <= new Date(dueDate)) return 0;
  let months=0, cur=new Date(dueDate);
  while(cur < new Date(depositDate)){ cur.setMonth(cur.getMonth()+1); months++; }
  return months;
}

// ── Main calculate ─────────────────────────────────────────────────────────────

function calculate(){
  const code   = document.getElementById("mainSection").value;
  const amt    = parseFloat(document.getElementById("paymentAmt").value);
  const dDate  = document.getElementById("deductionDate").value;
  const aDate  = document.getElementById("depositDate").value;

  if(!code){ alert("Please select a payment type."); return; }
  if(!amt||amt<=0){ alert("Please enter a valid amount."); return; }

  const data = currentMode==="tds" ? TDS_DATA : TCS_DATA;
  const d    = data[code];
  if(!d){ alert("Data not found."); return; }

  const isTCS      = currentMode==="tcs";
  const intRate    = isTCS ? 0.01 : 0.015; // TCS: 1%/mo, TDS: 1.5%/mo
  const intSecLbl  = isTCS ? "u/s 206C(7)" : "u/s 201(1A)";

  const belowThresh = d.thresh && amt < d.thresh;
  const tax         = belowThresh ? 0 : (d.rate===0 ? 0 : Math.round(amt * d.rate / 100));
  const net         = isTCS ? amt + tax : amt - tax;

  const fmt = n => "₹"+Math.round(n).toLocaleString("en-IN");
  const fmtDate = dt => new Date(dt).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});

  document.getElementById("resultSection").style.display = "block";

  if(!dDate||!aDate){
    // Basic result only
    document.getElementById("ontimeBox").style.display = "none";
    document.getElementById("lateBoxes").style.display = "none";
    document.getElementById("basicBox").style.display  = "block";
    document.getElementById("b-main").textContent = belowThresh?"No "+(isTCS?"TCS":"TDS"):d.rate===0?"Slab Rate":fmt(tax);
    document.getElementById("b-sub").textContent  = belowThresh?"Below threshold of "+fmt(d.thresh):d.rate===0?"Compute at slab rate":d.rate+"% on "+fmt(amt);
    document.getElementById("b-payment").textContent = fmt(amt);
    document.getElementById("b-newsec").textContent  = d.newSec;
    document.getElementById("b-oldsec").textContent  = d.oldSec+" (ref only)";
    document.getElementById("b-code").textContent    = code;
    document.getElementById("b-rate").textContent    = d.rate===0?"Slab rate":d.rate+"%";
    document.getElementById("b-tds2").textContent    = belowThresh?"Nil (below threshold)":d.rate===0?"As per slab":fmt(tax);
    document.getElementById("b-net").textContent     = isTCS?fmt(net)+" (incl. TCS)":fmt(net);
    return;
  }

  const dueDate    = getDueDate(dDate, code, currentMode);
  const monthsLate = calcMonthsLate(dueDate, aDate);
  const interest   = Math.round(tax * intRate * monthsLate);
  const total      = tax + interest;
  const isOnTime   = monthsLate===0;

  document.getElementById("basicBox").style.display = "none";

  if(isOnTime||belowThresh||d.rate===0){
    document.getElementById("ontimeBox").style.display = "block";
    document.getElementById("lateBoxes").style.display = "none";
    if(belowThresh) document.getElementById("ontimeBox").textContent = "No "+(isTCS?"TCS":"TDS")+" — Below threshold of "+fmt(d.thresh);
    else if(d.rate===0) document.getElementById("ontimeBox").textContent = "Salary TDS — compute at applicable slab rate";
    else document.getElementById("ontimeBox").textContent = "✓ Deposit is on time — No interest. "+(isTCS?"TCS":"TDS")+": "+fmt(tax);
    return;
  }

  document.getElementById("ontimeBox").style.display = "none";
  document.getElementById("lateBoxes").style.display = "block";

  document.getElementById("r-main").textContent     = fmt(tax);
  document.getElementById("r-main-sub").textContent = d.rate+"% on "+fmt(amt);
  document.getElementById("r-int").textContent      = fmt(interest);
  document.getElementById("r-int-sub").textContent  = (intRate*100)+"% × "+monthsLate+" month"+(monthsLate>1?"s":"");
  document.getElementById("r-total").textContent    = fmt(total);
  document.getElementById("r-total-sub").textContent= (isTCS?"TCS":"TDS")+" "+fmt(tax)+" + Interest "+fmt(interest);
  document.getElementById("d-payment").textContent  = fmt(amt);
  document.getElementById("d-newsec").textContent   = d.newSec;
  document.getElementById("d-oldsec").textContent   = d.oldSec+" (ref only)";
  document.getElementById("d-code").textContent     = code;
  document.getElementById("d-rate").textContent     = d.rate+"%";
  document.getElementById("d-tds").textContent      = fmt(tax);
  document.getElementById("d-ddate").textContent    = fmtDate(dDate);
  document.getElementById("d-due").textContent      = dueDate.toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
  document.getElementById("d-adate").textContent    = fmtDate(aDate);
  document.getElementById("d-months").textContent   = monthsLate+" month"+(monthsLate>1?"s":"")+" (fractional = full month)";
  document.getElementById("d-intamt").textContent   = fmt(interest)+" "+intSecLbl;
  document.getElementById("d-total").textContent    = fmt(total);
}

// Init
document.getElementById("mainSection").innerHTML = buildTDSOptions();
</script>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>

<button class="help-btn" onclick="openHelp()" title="How to use this tool">?</button>
<div class="help-overlay" id="helpOverlay">
  <div class="help-modal">
    <div class="help-modal-head"><h3>How to Use — TDS/TCS Calculator</h3><button class="help-close" onclick="closeHelp()">&#10005;</button></div>
    <div class="help-modal-body"><div class="help-step"><div class="help-step-num">1</div><div class="help-step-body"><h4>Select Section</h4><p>Choose the TDS/TCS section (e.g. 194C, 194J, 206C etc.).</p></div></div><div class="help-step"><div class="help-step-num">2</div><div class="help-step-body"><h4>Enter Amount</h4><p>Enter the payment/receipt amount.</p></div></div><div class="help-step"><div class="help-step-num">3</div><div class="help-step-body"><h4>Check Threshold</h4><p>The tool shows whether TDS is applicable based on annual threshold.</p></div></div><div class="help-step"><div class="help-step-num">4</div><div class="help-step-body"><h4>View Rate</h4><p>See applicable TDS/TCS rate, deductible amount, and net payable.</p></div></div><div class="help-tip">💡 Updated for IT Act 2025 new payment codes (Sections 393/394).</div></div>
  </div>
</div>
<script>function openHelp(){document.getElementById('helpOverlay').classList.add('open')}function closeHelp(){document.getElementById('helpOverlay').classList.remove('open')}document.getElementById('helpOverlay').addEventListener('click',function(e){if(e.target===this)closeHelp()})</script>
</body></html>"""


DEP_CALC_T = r"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Depreciation Calculator – CA Toolkit</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>
<style>
""" + BASE_CSS + """
.hero{text-align:center;padding:40px 24px 28px;max-width:700px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#ECFDF5;
            color:#065F46;border:1px solid #A7F3D0;border-radius:99px;
            padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:14px}
h1{font-size:clamp(22px,4vw,34px);font-weight:800;line-height:1.15;letter-spacing:-.5px;margin-bottom:10px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:14px;color:var(--muted);line-height:1.7;max-width:500px;margin:0 auto}
.main{max-width:1000px;margin:0 auto;padding:28px 24px 48px}
.card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);
      box-shadow:var(--shadow);overflow:hidden;margin-bottom:20px}
.card-head{padding:14px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.card-head .icon{width:30px;height:30px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:15px}
.card-head h2{font-size:14px;font-weight:700}
.card-head p{font-size:12px;color:var(--muted);margin-top:1px}
.card-body{padding:20px}
.form-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
@media(max-width:700px){.form-grid{grid-template-columns:1fr}}
.field{margin-bottom:0}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);margin-bottom:5px}
.hint{font-size:11px;color:var(--muted);margin-top:4px}
input[type=number],input[type=text],select{width:100%;border:1.5px solid var(--border);border-radius:8px;
  padding:9px 12px;font-family:inherit;font-size:13px;color:var(--ink);background:var(--white);
  transition:border-color .2s;outline:none}
input:focus,select:focus{border-color:var(--brand)}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}
.btn{background:var(--brand);color:#fff;border:none;border-radius:10px;
     padding:11px 24px;font-family:inherit;font-size:14px;font-weight:700;
     cursor:pointer;transition:background .2s;margin-top:16px}
.btn:hover{background:var(--brand-d)}
.result-section{display:none}
.summary-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px}
@media(max-width:700px){.summary-grid{grid-template-columns:1fr 1fr}}
.summary-box{background:var(--white);border:1px solid var(--border);border-radius:10px;padding:16px}
.summary-box .val{font-size:20px;font-weight:800;color:var(--brand);margin-bottom:4px}
.summary-box .lbl{font-size:11px;color:var(--muted)}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:left;font-size:10px;letter-spacing:.06em;text-transform:uppercase;
   color:var(--muted);border-bottom:1.5px solid var(--border);padding:7px 10px}
td{padding:9px 10px;border-bottom:1px solid var(--border)}
tr:last-child td{border:none;font-weight:700;background:#F9FAFB}
td:not(:first-child){text-align:right}
.tag-it{background:#EFF6FF;color:var(--brand);font-size:10px;font-weight:700;
        padding:2px 7px;border-radius:99px}
.tag-ca{background:#F5F3FF;color:#5B21B6;font-size:10px;font-weight:700;
        padding:2px 7px;border-radius:99px}
footer{background:#0f1b2d;color:#9CA3AF;font-size:12px;padding:0}
.ft-main{display:grid;grid-template-columns:2fr 1fr 1.4fr;gap:40px;padding:40px 48px;max-width:1200px;margin:0 auto}
.ft-brand-name{color:#fff;font-size:18px;font-weight:800;margin-bottom:12px}
.ft-brand-desc{font-size:12.5px;line-height:1.75;color:#9CA3AF;max-width:340px;text-align:justify}
.ft-col-title{color:#fff;font-size:14px;font-weight:700;margin-bottom:14px}
.ft-links{list-style:none;padding:0;margin:0}
.ft-links li{margin-bottom:8px}
.ft-links a{color:#9CA3AF;text-decoration:none;font-size:13px;transition:color .2s}
.ft-links a:hover{color:#fff}
.ft-contact-name{color:#fff;font-weight:700;font-size:13px;margin-bottom:6px}
.ft-contact-addr{color:#9CA3AF;font-size:12px;line-height:1.7;margin-bottom:10px}
.ft-contact-line{color:#9CA3AF;font-size:12px;margin-bottom:4px}
.ft-socials{display:flex;gap:14px;margin-top:12px}
.ft-socials a{color:#9CA3AF;transition:color .2s}
.ft-socials a:hover{color:#fff}
.ft-socials svg{width:20px;height:20px;fill:currentColor}
.ft-bottom{background:#0a1422;border-top:1px solid #1e2d42;padding:12px 48px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.ft-bottom-left{font-size:11px;color:#6B7280}
.ft-bottom-right{font-size:11px;color:#6B7280}
@media(max-width:768px){.ft-main{grid-template-columns:1fr;padding:28px 20px;gap:24px}.ft-bottom{padding:12px 20px;flex-direction:column;text-align:center}}
</style></head><body>

<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <div class="nav-right">
    {% if username %}
    <span class="nav-user">👤 <strong>{{ username }}</strong></span>
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin</a>{% endif %}
    <a href="/logout" class="nav-link">Sign out</a>
    {% else %}
    <a href="/login" class="nav-btn">Sign In</a>
    {% endif %}
    <a href="/" class="nav-btn" style="background:#F3F4F6;color:var(--ink)">← Dashboard</a>
  </div>
</nav>

<section class="hero">
  <div class="hero-badge">🆓 Free Tool · No Login Required</div>
  <h1>Depreciation Calculator</h1>
  <p>Calculate depreciation under <strong>Companies Act 2013</strong> (WDV/SLM) and <strong>Income Tax Act</strong>. Get full year-wise schedule instantly.</p>
</section>

<div class="main">
  <div class="card">
    <div class="card-head">
      <div class="icon" style="background:#F5F3FF">🏭</div>
      <div><h2>Asset Details</h2><p>Enter asset information to generate depreciation schedule</p></div>
    </div>
    <div class="card-body">
      <div class="form-grid">
        <div class="field">
          <label>Asset Name</label>
          <input type="text" id="assetName" placeholder="e.g. Machinery, Vehicle"/>
        </div>
        <div class="field">
          <label>Cost of Asset (₹)</label>
          <input type="number" id="assetCost" placeholder="e.g. 500000" min="0"/>
        </div>
        <div class="field">
          <label>Date of Purchase</label>
          <input type="date" id="purchaseDate"/>
        </div>
        <div class="field">
          <label>Asset Block (IT Act)</label>
          <select id="itBlock">
            <option value="15">15% — Furniture, Fittings</option>
            <option value="15b">15% — Ships</option>
            <option value="30">30% — Motor Cars (not used for hire)</option>
            <option value="40">40% — Motor Taxis, Buses (hire)</option>
            <option value="40b">40% — Machinery (general)</option>
            <option value="60">60% — Computers &amp; Software</option>
            <option value="80">80% — Energy saving devices</option>
            <option value="100">100% — Books, Scientific research</option>
            <option value="10">10% — Buildings (residential)</option>
            <option value="5">5% — Buildings (other)</option>
          </select>
        </div>
        <div class="field">
          <label>Asset Class (Companies Act)</label>
          <select id="caClass">
            <option value="15_wdv">Buildings — Factory (5% SLM / 15 yr WDV)</option>
            <option value="10_wdv">Buildings — Other (10% SLM / 10 yr WDV)</option>
            <option value="15_plant">Plant &amp; Machinery General (15% SLM)</option>
            <option value="30_plant">Plant &amp; Machinery (30% SLM — certain)</option>
            <option value="20_furn">Furniture &amp; Fixtures (10% SLM)</option>
            <option value="25_comp">Computers &amp; Peripherals (40% SLM)</option>
            <option value="20_veh">Vehicles — Motor Car (20% SLM)</option>
            <option value="30_veh">Vehicles — Motor Cycle (30% SLM)</option>
            <option value="10_off">Office Equipment (20% SLM)</option>
          </select>
        </div>
        <div class="field">
          <label>Method (Companies Act)</label>
          <select id="caMethod">
            <option value="slm">SLM — Straight Line Method</option>
            <option value="wdv">WDV — Written Down Value</option>
          </select>
        </div>
      </div>
      <div class="row2">
        <div class="field">
          <label>Number of Years to Project</label>
          <input type="number" id="numYears" value="5" min="1" max="20"/>
        </div>
        <div class="field">
          <label>Salvage / Residual Value (₹)</label>
          <input type="number" id="salvageVal" value="0" min="0"/>
          <p class="hint">Under Companies Act, minimum 5% of cost</p>
        </div>
      </div>
      <button class="btn" onclick="calcDep()">Generate Depreciation Schedule →</button>
    </div>
  </div>

  <div class="result-section" id="resultSection">
    <div class="summary-grid" id="summaryGrid"></div>

    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#EFF6FF">📊</div>
        <div><h2>Income Tax Act Schedule <span class="tag-it">IT Act</span></h2>
             <p>WDV method — Block of assets basis</p></div>
      </div>
      <div class="card-body" style="padding:0;overflow-x:auto">
        <table id="itTable">
          <thead><tr><th>FY</th><th>Opening WDV</th><th>Additions</th><th>Depreciation</th><th>Closing WDV</th></tr></thead>
          <tbody id="itBody"></tbody>
        </table>
      </div>
    </div>

    <div class="card" style="margin-top:20px">
      <div class="card-head">
        <div class="icon" style="background:#F5F3FF">📋</div>
        <div><h2>Companies Act 2013 Schedule <span class="tag-ca">Companies Act</span></h2>
             <p id="caMethodLabel">SLM method</p></div>
      </div>
      <div class="card-body" style="padding:0;overflow-x:auto">
        <table id="caTable">
          <thead><tr><th>FY</th><th>Opening WDV</th><th>Depreciation</th><th>Closing WDV</th><th>Acc. Dep.</th></tr></thead>
          <tbody id="caBody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<footer>
  <div class="ft-main">
    <div>
      <div class="ft-brand-name">CA Toolkit</div>
      <p class="ft-brand-desc">CA Toolkit is a comprehensive utility platform built by a CA Article from Ludhiana, Punjab, providing automation tools for Indian Chartered Accountants. The platform saves hours of manual work every year — from Balance Sheet year-shift to GST reconciliation, tax calculations, and more.</p>
    </div>
    <div>
      <div class="ft-col-title">Know More</div>
      <ul class="ft-links">
        <li><a href="/">Home</a></li>
        <li><a href="/">BS Year Shift</a></li>
        <li><a href="/tool/tb-to-bs">TB → Balance Sheet</a></li>
        <li><a href="/tool/tax-calculator">Tax Calculator</a></li>
        <li><a href="/privacy">Privacy Policy</a></li>
      </ul>
    </div>
    <div>
      <div class="ft-col-title">Contact Us</div>
      <div class="ft-contact-name">CA Toolkit</div>
      <div class="ft-contact-addr">Built for Indian Chartered Accountants<br/>Created by CA Article · Ludhiana, Punjab</div>
    </div>
  </div>
  <div class="ft-bottom">
    <span class="ft-bottom-left">©2026 CA Toolkit · All Rights Reserved</span>
    <span class="ft-bottom-right">Built for Indian CAs · Ludhiana, Punjab</span>
  </div>
</footer>

<script>
const IT_RATES = {"15":15,"15b":15,"30":30,"40":40,"40b":40,"60":60,"80":80,"100":100,"10":10,"5":5};
const CA_RATES = {
  "15_wdv":  {slm:6.67, wdv:15,  life:15},
  "10_wdv":  {slm:10,   wdv:10,  life:10},
  "15_plant":{slm:6.67, wdv:15,  life:15},
  "30_plant":{slm:10,   wdv:30,  life:10},
  "20_furn": {slm:10,   wdv:10,  life:10},
  "25_comp": {slm:40,   wdv:40,  life:3},
  "20_veh":  {slm:20,   wdv:20,  life:5},
  "30_veh":  {slm:30,   wdv:30,  life:4},
  "10_off":  {slm:20,   wdv:20,  life:5},
};

function fmt(n){ return "₹"+Math.round(n).toLocaleString("en-IN"); }

function calcDep(){
  const cost = parseFloat(document.getElementById("assetCost").value);
  const name = document.getElementById("assetName").value || "Asset";
  const pd   = document.getElementById("purchaseDate").value;
  const itBl = document.getElementById("itBlock").value;
  const caCl = document.getElementById("caClass").value;
  const meth = document.getElementById("caMethod").value;
  const yrs  = Math.min(parseInt(document.getElementById("numYears").value)||5, 20);
  const salv = Math.max(parseFloat(document.getElementById("salvageVal").value)||0, cost*0.05);

  if(!cost||cost<=0){alert("Enter a valid asset cost.");return;}
  if(!pd){alert("Enter purchase date.");return;}

  const purchaseYear = parseInt(pd.split("-")[0]);
  const purchaseMon  = parseInt(pd.split("-")[1]);
  // IT Act: if purchased after 3 Oct (i.e. used < 180 days), half rate in first year
  const halfRate = purchaseMon >= 10 || (purchaseMon === 9 && parseInt(pd.split("-")[2]) > 3);

  const itRate = IT_RATES[itBl] / 100;
  const caInfo = CA_RATES[caCl];
  const caRate = meth === "slm" ? caInfo.slm/100 : caInfo.wdv/100;

  // IT Schedule (WDV)
  let itWDV = cost, itRows = "";
  for(let i=0;i<yrs;i++){
    const fy = `FY ${purchaseYear + i}-${String(purchaseYear+i+1).slice(-2)}`;
    const additions = i===0 ? cost : 0;
    const rate = (i===0 && halfRate) ? itRate/2 : itRate;
    const dep = Math.round(itWDV * rate);
    const closing = itWDV - dep;
    itRows += `<tr><td>${fy}</td><td>${fmt(itWDV)}</td><td>${i===0?fmt(additions):"—"}</td><td>${fmt(dep)}</td><td>${fmt(closing)}</td></tr>`;
    itWDV = closing;
    if(itWDV <= 0) break;
  }
  document.getElementById("itBody").innerHTML = itRows;

  // CA Schedule
  let caWDV = cost, caAcc = 0, caRows = "";
  document.getElementById("caMethodLabel").textContent = meth.toUpperCase() + " method";
  for(let i=0;i<yrs;i++){
    const fy = `FY ${purchaseYear + i}-${String(purchaseYear+i+1).slice(-2)}`;
    let dep;
    if(meth === "slm"){
      dep = Math.round((cost - salv) * caRate);
      if(caWDV - dep < salv) dep = Math.max(0, caWDV - salv);
    } else {
      dep = Math.round(caWDV * caRate);
      if(caWDV - dep < salv) dep = Math.max(0, caWDV - salv);
    }
    caAcc += dep;
    const closing = caWDV - dep;
    caRows += `<tr><td>${fy}</td><td>${fmt(caWDV)}</td><td>${fmt(dep)}</td><td>${fmt(closing)}</td><td>${fmt(caAcc)}</td></tr>`;
    caWDV = closing;
    if(caWDV <= salv) break;
  }
  document.getElementById("caBody").innerHTML = caRows;

  // Summary
  const itDep1 = cost * ((halfRate?itRate/2:itRate));
  const caDep1 = meth==="slm" ? (cost-salv)*caRate : cost*caRate;
  document.getElementById("summaryGrid").innerHTML =
    `<div class="summary-box"><div class="val">${fmt(cost)}</div><div class="lbl">Asset Cost</div></div>
     <div class="summary-box"><div class="val">${fmt(itDep1)}</div><div class="lbl">Year 1 Dep (IT Act)</div></div>
     <div class="summary-box"><div class="val">${fmt(caDep1)}</div><div class="lbl">Year 1 Dep (Co. Act)</div></div>
     <div class="summary-box"><div class="val">${fmt(salv)}</div><div class="lbl">Residual Value</div></div>`;

  document.getElementById("resultSection").style.display = "block";
  document.getElementById("resultSection").scrollIntoView({behavior:"smooth"});
}
</script>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
</body></html>"""


# ══════════════════════════════════════════════════════════════════════════════
#  ADMIN PANEL
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
#  MSME DISALLOWANCE CALCULATOR
# ══════════════════════════════════════════════════════════════════════════════

MSME_T = r"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MSME Disallowance Calculator – CA Toolkit</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>
<style>
""" + BASE_CSS + """
.hero{text-align:center;padding:32px 24px 16px;max-width:760px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#ECFDF5;color:#065F46;
            border:1px solid #A7F3D0;border-radius:99px;padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:12px}
h1{font-size:clamp(20px,4vw,32px);font-weight:800;line-height:1.15;letter-spacing:-.5px;margin-bottom:8px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:13px;color:var(--muted);line-height:1.7;max-width:560px;margin:0 auto}
.wrap{max-width:1100px;margin:0 auto;padding:16px 24px 48px}
.card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);
      box-shadow:var(--shadow);overflow:hidden;margin-bottom:20px}
.card-head{padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.card-head .icon{width:30px;height:30px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:15px}
.card-head h2{font-size:14px;font-weight:700}
.card-head p{font-size:12px;color:var(--muted);margin-top:1px}
.card-body{padding:18px}
.field{margin-bottom:14px}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);margin-bottom:4px}
.hint{font-size:11px;color:var(--muted);margin-top:3px}
input[type=file],input[type=number],input[type=date]{width:100%;border:1.5px solid var(--border);border-radius:8px;
  padding:8px 11px;font-family:inherit;font-size:13px;color:var(--ink);background:var(--white);
  transition:border-color .2s;outline:none}
input:focus{border-color:var(--brand)}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.btn{background:var(--brand);color:#fff;border:none;border-radius:8px;padding:10px 20px;
     font-family:inherit;font-size:13px;font-weight:700;cursor:pointer;transition:background .2s}
.btn:hover{background:var(--brand-d)}
.btn-full{width:100%;padding:12px;font-size:14px}
/* Format table */
.fmt-table{width:100%;border-collapse:collapse;font-size:12px}
.fmt-table th{text-align:left;font-size:10px;letter-spacing:.06em;text-transform:uppercase;
              color:var(--muted);border-bottom:1.5px solid var(--border);padding:6px 10px;background:#F9FAFB}
.fmt-table td{padding:8px 10px;border-bottom:1px solid var(--border);vertical-align:top}
.fmt-table tr:last-child td{border:none}
.col-req{background:#EFF6FF;color:var(--brand);font-size:10px;font-weight:700;
         padding:1px 6px;border-radius:4px;font-family:monospace}
/* Results */
.summary-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px}
@media(max-width:700px){.summary-grid{grid-template-columns:1fr 1fr}}
.sbox{background:var(--white);border:1px solid var(--border);border-radius:10px;padding:14px 16px}
.sbox.red{background:#FEF2F2;border-color:#FECACA}
.sbox.green{background:#ECFDF5;border-color:#A7F3D0}
.sbox.yellow{background:#FFFBEB;border-color:#FDE68A}
.sbox .val{font-size:20px;font-weight:800;margin-bottom:3px}
.sbox .lbl{font-size:11px;color:var(--muted);font-weight:500}
.sbox.red .val{color:#991B1B}
.sbox.green .val{color:#065F46}
.sbox.yellow .val{color:#92400E}
/* Result table */
.res-table{width:100%;border-collapse:collapse;font-size:12px}
.res-table th{text-align:left;font-size:10px;letter-spacing:.06em;text-transform:uppercase;
              color:var(--muted);border-bottom:1.5px solid var(--border);padding:7px 8px;
              background:#F9FAFB;position:sticky;top:0}
.res-table td{padding:8px;border-bottom:1px solid var(--border);vertical-align:middle}
.res-table tr:hover td{background:#F9FAFB}
.row-ok{background:#F0FDF4}
.row-warn{background:#FFFBEB}
.row-over{background:#FEF2F2}
.badge-ok{background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:2px 8px;border-radius:99px}
.badge-warn{background:#FFFBEB;color:#92400E;font-size:10px;font-weight:700;padding:2px 8px;border-radius:99px}
.badge-over{background:#FEF2F2;color:#991B1B;font-size:10px;font-weight:700;padding:2px 8px;border-radius:99px}
.badge-na{background:#F3F4F6;color:var(--muted);font-size:10px;font-weight:700;padding:2px 8px;border-radius:99px}
.note-box{background:#FEF2F2;border:1px solid #FECACA;border-radius:8px;padding:10px 14px;
          font-size:12px;color:#991B1B;margin-top:12px;line-height:1.7}
.info-box{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;padding:10px 14px;
          font-size:12px;color:#1e40af;margin-bottom:16px;line-height:1.7}
.dl-btn{background:var(--green);color:#fff;border:none;border-radius:8px;padding:8px 16px;
        font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;margin-left:8px}
footer{background:#0f1b2d;color:#9CA3AF;font-size:12px;padding:0}
.ft-main{display:grid;grid-template-columns:2fr 1fr 1.4fr;gap:40px;padding:40px 48px;max-width:1200px;margin:0 auto}
.ft-brand-name{color:#fff;font-size:18px;font-weight:800;margin-bottom:12px}
.ft-brand-desc{font-size:12.5px;line-height:1.75;color:#9CA3AF;max-width:340px;text-align:justify}
.ft-col-title{color:#fff;font-size:14px;font-weight:700;margin-bottom:14px}
.ft-links{list-style:none;padding:0;margin:0}
.ft-links li{margin-bottom:8px}
.ft-links a{color:#9CA3AF;text-decoration:none;font-size:13px;transition:color .2s}
.ft-links a:hover{color:#fff}
.ft-contact-name{color:#fff;font-weight:700;font-size:13px;margin-bottom:6px}
.ft-contact-addr{color:#9CA3AF;font-size:12px;line-height:1.7;margin-bottom:10px}
.ft-contact-line{color:#9CA3AF;font-size:12px;margin-bottom:4px}
.ft-socials{display:flex;gap:14px;margin-top:12px}
.ft-socials a{color:#9CA3AF;transition:color .2s}
.ft-socials a:hover{color:#fff}
.ft-socials svg{width:20px;height:20px;fill:currentColor}
.ft-bottom{background:#0a1422;border-top:1px solid #1e2d42;padding:12px 48px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.ft-bottom-left{font-size:11px;color:#6B7280}
.ft-bottom-right{font-size:11px;color:#6B7280}
@media(max-width:768px){.ft-main{grid-template-columns:1fr;padding:28px 20px;gap:24px}.ft-bottom{padding:12px 20px;flex-direction:column;text-align:center}}
</style></head><body>

<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <div class="nav-right">
    {% if username %}<span class="nav-user">👤 <strong>{{ username }}</strong></span>
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin</a>{% endif %}
    <a href="/logout" class="nav-link">Sign out</a>
    {% else %}<a href="/login" class="nav-btn">Sign In</a>{% endif %}
    <a href="/" class="nav-btn" style="background:#F3F4F6;color:var(--ink)">← Dashboard</a>
  </div>
</nav>

<section class="hero">
  <div class="hero-badge">🆓 Free · No Login Required</div>
  <h1>MSME Disallowance Calculator</h1>
  <p>Check creditor payments against MSME time limits under <strong>Section 43B(h)</strong>. Upload your creditors list and instantly see which payments are overdue and the total disallowance amount.</p>
</section>

<div class="wrap">

  <div class="info-box">
    ℹ️ <strong>Section 43B(h) — IT Act:</strong> Payments to MSME suppliers must be made within 15 days (no agreement) or 45 days (written agreement) from date of invoice. Any unpaid amount beyond the limit is <strong>disallowed</strong> as a deduction in the year of computation and allowed only in the year of actual payment.
  </div>

  <!-- FORMAT GUIDE -->
  <div class="card">
    <div class="card-head">
      <div class="icon" style="background:#FFFBEB">📋</div>
      <div><h2>Required Excel Format</h2><p>Your upload must follow this column structure exactly</p></div>
    </div>
    <div class="card-body" style="padding:0;overflow-x:auto">
      <table class="fmt-table">
        <thead><tr><th>#</th><th>Column Name</th><th>Format</th><th>Example</th><th>Notes</th></tr></thead>
        <tbody>
          <tr><td>A</td><td><span class="col-req">Creditor Name</span></td><td>Text</td><td>ABC Enterprises</td><td>Name of the MSME supplier</td></tr>
          <tr><td>B</td><td><span class="col-req">Invoice Date</span></td><td>DD/MM/YYYY</td><td>01/04/2025</td><td>Date of invoice / bill received</td></tr>
          <tr><td>C</td><td><span class="col-req">Invoice Amount</span></td><td>Number</td><td>50000</td><td>Total invoice amount (₹)</td></tr>
          <tr><td>D</td><td><span class="col-req">Payment Date</span></td><td>DD/MM/YYYY or blank</td><td>20/05/2025</td><td>Leave blank if payment not yet made</td></tr>
          <tr><td>E</td><td><span class="col-req">Amount Paid</span></td><td>Number or 0</td><td>50000</td><td>Amount actually paid (0 if unpaid)</td></tr>
          <tr><td>F</td><td><span class="col-req">Written Agreement</span></td><td>Yes / No</td><td>No</td><td>Yes = 45 day limit, No = 15 day limit</td></tr>
          <tr><td>G</td><td><span class="col-req">MSME Category</span></td><td>Micro / Small / Medium</td><td>Small</td><td>MSME registration category of supplier</td></tr>
        </tbody>
      </table>
      <div style="padding:12px 16px;border-top:1px solid var(--border);display:flex;align-items:center;gap:12px;flex-wrap:wrap">
        <span style="font-size:12px;color:var(--muted)">Download blank template to fill in your data:</span>
        <button class="btn" onclick="downloadTemplate()" style="padding:7px 14px;font-size:12px">⬇ Download Template (.xlsx)</button>
      </div>
    </div>
  </div>

  <!-- UPLOAD -->
  <div class="card">
    <div class="card-head">
      <div class="icon" style="background:#EFF6FF">📁</div>
      <div><h2>Upload Creditors List</h2><p>Excel file (.xlsx) following the format above</p></div>
    </div>
    <div class="card-body">
      <div class="row2">
        <div class="field">
          <label>Upload Excel File (.xlsx)</label>
          <input type="file" id="msmeFile" accept=".xlsx"/>
          <p class="hint">Must follow the format shown above</p>
        </div>
        <div class="field">
          <label>Assessment Year</label>
          <input type="number" id="assessYear" value="2026" min="2020" max="2030"/>
          <p class="hint">Year for which disallowance is computed</p>
        </div>
      </div>
      <button class="btn btn-full" onclick="processMSME()">Analyse &amp; Calculate Disallowance →</button>
      <div id="msmeError" style="display:none;margin-top:10px;background:#FEF2F2;border:1px solid #FECACA;
           border-radius:8px;padding:10px 14px;font-size:13px;color:#991B1B"></div>
    </div>
  </div>

  <!-- RESULTS -->
  <div id="msmeResults" style="display:none">
    <div class="summary-grid" id="summaryGrid"></div>

    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#FEF2F2">📊</div>
        <div>
          <h2>Creditor-wise Analysis</h2>
          <p>Overdue payments highlighted in red · Near-due in yellow · On time in green</p>
        </div>
        <div style="margin-left:auto">
          <button class="dl-btn" onclick="exportResults()">⬇ Export Results</button>
        </div>
      </div>
      <div style="overflow-x:auto">
        <table class="res-table" id="resultsTable">
          <thead><tr>
            <th>Creditor</th><th>Category</th><th>Invoice Date</th><th>Invoice Amt</th>
            <th>Paid Amt</th><th>Payment Date</th><th>Limit</th><th>Due Date</th>
            <th>Days Overdue</th><th>Unpaid Amt</th><th>Status</th>
          </tr></thead>
          <tbody id="resultsBody"></tbody>
        </table>
      </div>
      <div class="note-box" id="disallowNote"></div>
    </div>
  </div>

</div>

<footer>
  <div class="ft-main">
    <div>
      <div class="ft-brand-name">CA Toolkit</div>
      <p class="ft-brand-desc">CA Toolkit is a comprehensive utility platform built by a CA Article from Ludhiana, Punjab, providing automation tools for Indian Chartered Accountants. The platform saves hours of manual work every year — from Balance Sheet year-shift to GST reconciliation, tax calculations, and more.</p>
    </div>
    <div>
      <div class="ft-col-title">Know More</div>
      <ul class="ft-links">
        <li><a href="/">Home</a></li>
        <li><a href="/">BS Year Shift</a></li>
        <li><a href="/tool/tb-to-bs">TB → Balance Sheet</a></li>
        <li><a href="/tool/tax-calculator">Tax Calculator</a></li>
        <li><a href="/privacy">Privacy Policy</a></li>
      </ul>
    </div>
    <div>
      <div class="ft-col-title">Contact Us</div>
      <div class="ft-contact-name">CA Toolkit</div>
      <div class="ft-contact-addr">Built for Indian Chartered Accountants<br/>Created by CA Article · Ludhiana, Punjab</div>
    </div>
  </div>
  <div class="ft-bottom">
    <span class="ft-bottom-left">©2026 CA Toolkit · All Rights Reserved</span>
    <span class="ft-bottom-right">Built for Indian CAs · Ludhiana, Punjab</span>
  </div>
</footer>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
const fmt = n => "₹" + Math.round(n).toLocaleString("en-IN");

function parseDate(val){
  if(!val) return null;
  if(val instanceof Date) return val;
  // Try DD/MM/YYYY
  const s = String(val).trim();
  const parts = s.split("/");
  if(parts.length===3) return new Date(parts[2], parts[1]-1, parts[0]);
  // Try Excel serial
  if(!isNaN(val)){
    const d = new Date((val - 25569) * 86400 * 1000);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return new Date(val);
}

function daysBetween(d1, d2){
  return Math.round((d2 - d1) / (1000*60*60*24));
}

function fmtDate(d){
  if(!d) return "—";
  return d.toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
}

let processedRows = [];

function processMSME(){
  const file = document.getElementById("msmeFile").files[0];
  const errEl = document.getElementById("msmeError");
  errEl.style.display = "none";

  if(!file){ showErr("Please select an Excel file."); return; }

  const reader = new FileReader();
  reader.onload = function(e){
    try{
      const wb   = XLSX.read(e.target.result, {type:"binary", cellDates:true});
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false});

      if(rows.length < 2){ showErr("File appears empty. Please check the format."); return; }

      // Skip header row
      const data = rows.slice(1).filter(r => r && r[0]);
      if(data.length === 0){ showErr("No data rows found. Please check the format."); return; }

      processedRows = [];
      let totalInvoice=0, totalPaid=0, totalDisallow=0, totalOverdue=0;
      const today = new Date();

      const tbody = document.getElementById("resultsBody");
      tbody.innerHTML = "";

      data.forEach((row, i) => {
        const name        = String(row[0]||"").trim();
        const invoiceDate = parseDate(row[1]);
        const invoiceAmt  = parseFloat(String(row[2]||"0").replace(/,/g,""))||0;
        const paymentDate = parseDate(row[3]);
        const paidAmt     = parseFloat(String(row[4]||"0").replace(/,/g,""))||0;
        const hasAgreement= String(row[5]||"No").trim().toLowerCase()==="yes";
        const category    = String(row[6]||"MSME").trim();

        if(!name||!invoiceDate) return;

        const limitDays = hasAgreement ? 45 : 15;
        const dueDate   = new Date(invoiceDate);
        dueDate.setDate(dueDate.getDate() + limitDays);

        const refDate  = paymentDate || today;
        const daysDiff = daysBetween(dueDate, refDate);
        const unpaid   = Math.max(0, invoiceAmt - paidAmt);
        const isOverdue = daysDiff > 0;
        const isNearDue = !isOverdue && daysDiff > -7;

        let status, rowClass, badge;
        if(!isOverdue && paidAmt >= invoiceAmt){
          status="✓ Paid on time"; rowClass="row-ok"; badge=`<span class="badge-ok">Paid On Time</span>`;
        } else if(isOverdue && unpaid > 0){
          status="⚠ Overdue"; rowClass="row-over"; badge=`<span class="badge-over">Overdue</span>`;
          totalDisallow += unpaid;
          totalOverdue++;
        } else if(!isOverdue && unpaid > 0){
          status="Near due"; rowClass="row-warn"; badge=`<span class="badge-warn">Pending</span>`;
        } else if(isOverdue && paidAmt >= invoiceAmt){
          status="Paid late"; rowClass="row-warn"; badge=`<span class="badge-warn">Paid Late</span>`;
        } else {
          status="—"; rowClass=""; badge=`<span class="badge-na">N/A</span>`;
        }

        totalInvoice += invoiceAmt;
        totalPaid    += paidAmt;

        processedRows.push({name,category,invoiceDate,invoiceAmt,paidAmt,paymentDate,
          limitDays,dueDate,daysDiff,unpaid,status,isOverdue});

        const tr = document.createElement("tr");
        tr.className = rowClass;
        tr.innerHTML = `
          <td><strong>${name}</strong></td>
          <td>${category}</td>
          <td>${fmtDate(invoiceDate)}</td>
          <td>${fmt(invoiceAmt)}</td>
          <td>${fmt(paidAmt)}</td>
          <td>${fmtDate(paymentDate)}</td>
          <td>${limitDays} days</td>
          <td>${fmtDate(dueDate)}</td>
          <td style="font-weight:700;color:${isOverdue&&unpaid>0?"#991B1B":daysDiff>-7?"#92400E":"#065F46"}">${isOverdue?"+"+daysDiff+" days":daysDiff===0?"Today":Math.abs(daysDiff)+" days left"}</td>
          <td style="font-weight:700;color:${unpaid>0?"#991B1B":"#065F46"}">${fmt(unpaid)}</td>
          <td>${badge}</td>`;
        tbody.appendChild(tr);
      });

      // Summary
      const ayear = document.getElementById("assessYear").value;
      document.getElementById("summaryGrid").innerHTML = `
        <div class="sbox"><div class="val">${processedRows.length}</div><div class="lbl">Total Creditors</div></div>
        <div class="sbox yellow"><div class="val">${fmt(totalInvoice)}</div><div class="lbl">Total Invoice Value</div></div>
        <div class="sbox ${totalOverdue>0?"red":"green"}"><div class="val">${totalOverdue}</div><div class="lbl">Overdue Creditors</div></div>
        <div class="sbox ${totalDisallow>0?"red":"green"}"><div class="val">${fmt(totalDisallow)}</div><div class="lbl">Disallowance u/s 43B(h)</div></div>`;

      document.getElementById("disallowNote").innerHTML = totalDisallow > 0
        ? `⚠ <strong>Total disallowance u/s 43B(h) for AY ${ayear}-${parseInt(ayear)+1}: ${fmt(totalDisallow)}</strong><br>
           This amount will be added back to income and disallowed as a deduction. It will be allowed only in the year when actual payment is made to the MSME supplier.`
        : `✓ No disallowance applicable. All MSME payments are within the prescribed time limits.`;

      document.getElementById("msmeResults").style.display = "block";
      document.getElementById("msmeResults").scrollIntoView({behavior:"smooth"});

    } catch(err){
      showErr("Error reading file: "+err.message+". Please ensure file follows the required format.");
    }
  };
  reader.readAsBinaryString(file);
}

function showErr(msg){
  const el = document.getElementById("msmeError");
  el.textContent = msg; el.style.display = "block";
}

function downloadTemplate(){
  const wb = XLSX.utils.book_new();
  const data = [
    ["Creditor Name","Invoice Date","Invoice Amount","Payment Date","Amount Paid","Written Agreement","MSME Category"],
    ["ABC Enterprises","01/04/2025","50000","20/04/2025","50000","No","Small"],
    ["XYZ Traders","15/04/2025","120000","","0","Yes","Micro"],
    ["PQR Industries","01/05/2025","80000","20/06/2025","80000","No","Medium"],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{wch:20},{wch:14},{wch:16},{wch:14},{wch:12},{wch:18},{wch:16}];
  XLSX.utils.book_append_sheet(wb, ws, "Creditors");
  XLSX.writeFile(wb, "MSME_Creditors_Template.xlsx");
}

function exportResults(){
  if(!processedRows.length) return;
  const data = [["Creditor","Category","Invoice Date","Invoice Amt","Paid Amt","Payment Date","Limit","Due Date","Days Overdue","Unpaid Amt","Status"]];
  processedRows.forEach(r => {
    data.push([r.name,r.category,fmtDate(r.invoiceDate),r.invoiceAmt,r.paidAmt,
      fmtDate(r.paymentDate),r.limitDays+" days",fmtDate(r.dueDate),
      r.isOverdue?"+"+r.daysDiff+" days":Math.abs(r.daysDiff)+" days left",r.unpaid,r.status]);
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "MSME Analysis");
  XLSX.writeFile(wb, "MSME_Disallowance_Analysis.xlsx");
}
</script>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
</body></html>"""


# ══════════════════════════════════════════════════════════════════════════════
#  CAPITAL GAINS CALCULATOR
# ══════════════════════════════════════════════════════════════════════════════

CG_CALC_T = r"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Capital Gains Calculator – CA Toolkit</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>
<style>
""" + BASE_CSS + """
.hero{text-align:center;padding:32px 24px 16px;max-width:760px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#ECFDF5;color:#065F46;
            border:1px solid #A7F3D0;border-radius:99px;padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:12px}
h1{font-size:clamp(20px,4vw,32px);font-weight:800;line-height:1.15;letter-spacing:-.5px;margin-bottom:8px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:13px;color:var(--muted);line-height:1.7;max-width:520px;margin:0 auto}
.wrap{max-width:1100px;margin:0 auto;padding:16px 24px 48px;display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:start}
@media(max-width:800px){.wrap{grid-template-columns:1fr}}
.card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);
      box-shadow:var(--shadow);overflow:hidden;margin-bottom:16px}
.card:last-child{margin-bottom:0}
.card-head{padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.card-head .icon{width:30px;height:30px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:15px}
.card-head h2{font-size:14px;font-weight:700}
.card-head p{font-size:12px;color:var(--muted);margin-top:1px}
.card-body{padding:16px}
.field{margin-bottom:13px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);margin-bottom:4px}
.hint{font-size:11px;color:var(--muted);margin-top:3px}
select,input[type=number],input[type=text]{width:100%;border:1.5px solid var(--border);border-radius:8px;
  padding:8px 11px;font-family:inherit;font-size:13px;color:var(--ink);background:var(--white);
  transition:border-color .2s;outline:none}
select:focus,input:focus{border-color:var(--brand)}
.btn{width:100%;background:var(--brand);color:#fff;border:none;border-radius:8px;
     padding:11px;font-family:inherit;font-size:14px;font-weight:700;cursor:pointer;transition:background .2s}
.btn:hover{background:var(--brand-d)}
/* Tabs */
.tabs{display:flex;gap:0;margin-bottom:16px;border-radius:8px;overflow:hidden;border:1px solid var(--border)}
.tab{flex:1;padding:9px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;
     background:var(--white);color:var(--muted);border:none;transition:all .2s}
.tab.active{background:var(--brand);color:#fff}
/* Result boxes */
.rboxes{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px}
.rbox{border-radius:10px;padding:13px 15px}
.rbox-blue{background:#EFF6FF;border:1.5px solid #BFDBFE}
.rbox-green{background:#ECFDF5;border:1.5px solid #A7F3D0}
.rbox-red{background:#FEF2F2;border:1.5px solid #FECACA}
.rbox-total{background:#1D4ED8;border:1.5px solid #1D4ED8;grid-column:1/-1}
.rbox .val{font-size:20px;font-weight:800;margin-bottom:2px}
.rbox .lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;opacity:.75}
.rbox .sub{font-size:11px;margin-top:4px;opacity:.8}
.rbox-blue  .val{color:#1D4ED8}.rbox-blue  .lbl{color:#1D4ED8}
.rbox-green .val{color:#065F46}.rbox-green .lbl{color:#065F46}
.rbox-red   .val{color:#991B1B}.rbox-red   .lbl{color:#991B1B}
.rbox-total .val{color:#fff;font-size:22px}.rbox-total .lbl{color:rgba(255,255,255,.75)}
.rbox-total .sub{color:rgba(255,255,255,.8);font-size:11px}
.dtable{width:100%;border-collapse:collapse;font-size:12px;margin-top:10px}
.dtable td{padding:6px 2px;border-bottom:1px solid var(--border)}
.dtable tr:last-child td{border:none;font-weight:700}
.dtable td:last-child{text-align:right;font-weight:600}
.compare-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:12px}
.cbox{border-radius:10px;padding:13px 15px;border:1.5px solid var(--border);background:var(--white)}
.cbox.winner{border-color:var(--green);background:#ECFDF5}
.cbox h3{font-size:12px;font-weight:700;margin-bottom:8px;color:var(--ink)}
.cbox .ctax{font-size:20px;font-weight:800;color:var(--brand);margin-bottom:3px}
.cbox.winner .ctax{color:#065F46}
.cbox .csub{font-size:11px;color:var(--muted)}
.winner-badge{background:var(--green);color:#fff;font-size:10px;font-weight:700;
              padding:2px 8px;border-radius:99px;display:inline-block;margin-bottom:6px}
.reverse-box{background:#F5F3FF;border:1.5px solid #DDD6FE;border-radius:10px;padding:14px 16px;margin-top:12px}
.reverse-box h3{font-size:12px;font-weight:700;color:#5B21B6;margin-bottom:8px}
.reverse-box .rval{font-size:22px;font-weight:800;color:#5B21B6;margin-bottom:3px}
.reverse-box .rsub{font-size:11px;color:var(--muted)}
.info-box{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;padding:10px 14px;
          font-size:12px;color:#1e40af;margin-bottom:14px;line-height:1.7}
.note-box{background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;
          padding:10px 12px;font-size:11px;color:#92400E;margin-top:10px;line-height:1.6}
/* CII table */
.cii-table{width:100%;border-collapse:collapse;font-size:11px}
.cii-table th{text-align:left;font-size:10px;letter-spacing:.06em;text-transform:uppercase;
              color:var(--muted);border-bottom:1.5px solid var(--border);padding:5px 6px}
.cii-table td{padding:5px 6px;border-bottom:1px solid var(--border)}
.cii-table tr:last-child td{border:none}
.cii-table tr:hover td{background:#F9FAFB}
.cii-table .highlight{background:#EFF6FF;font-weight:700}
footer{background:#0f1b2d;color:#9CA3AF;font-size:12px;padding:0}
.ft-main{display:grid;grid-template-columns:2fr 1fr 1.4fr;gap:40px;padding:40px 48px;max-width:1200px;margin:0 auto}
.ft-brand-name{color:#fff;font-size:18px;font-weight:800;margin-bottom:12px}
.ft-brand-desc{font-size:12.5px;line-height:1.75;color:#9CA3AF;max-width:340px;text-align:justify}
.ft-col-title{color:#fff;font-size:14px;font-weight:700;margin-bottom:14px}
.ft-links{list-style:none;padding:0;margin:0}
.ft-links li{margin-bottom:8px}
.ft-links a{color:#9CA3AF;text-decoration:none;font-size:13px;transition:color .2s}
.ft-links a:hover{color:#fff}
.ft-contact-name{color:#fff;font-weight:700;font-size:13px;margin-bottom:6px}
.ft-contact-addr{color:#9CA3AF;font-size:12px;line-height:1.7;margin-bottom:10px}
.ft-contact-line{color:#9CA3AF;font-size:12px;margin-bottom:4px}
.ft-socials{display:flex;gap:14px;margin-top:12px}
.ft-socials a{color:#9CA3AF;transition:color .2s}
.ft-socials a:hover{color:#fff}
.ft-socials svg{width:20px;height:20px;fill:currentColor}
.ft-bottom{background:#0a1422;border-top:1px solid #1e2d42;padding:12px 48px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.ft-bottom-left{font-size:11px;color:#6B7280}
.ft-bottom-right{font-size:11px;color:#6B7280}
@media(max-width:768px){.ft-main{grid-template-columns:1fr;padding:28px 20px;gap:24px}.ft-bottom{padding:12px 20px;flex-direction:column;text-align:center}}
</style></head><body>

<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <div class="nav-right">
    {% if username %}<span class="nav-user">👤 <strong>{{ username }}</strong></span>
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin</a>{% endif %}
    <a href="/logout" class="nav-link">Sign out</a>
    {% else %}<a href="/login" class="nav-btn">Sign In</a>{% endif %}
    <a href="/" class="nav-btn" style="background:#F3F4F6;color:var(--ink)">← Dashboard</a>
  </div>
</nav>

<section class="hero">
  <div class="hero-badge">🆓 Free · No Login Required</div>
  <h1>Capital Gains Tax Calculator</h1>
  <p>Calculate LTCG / STCG on property, shares, mutual funds and more. Compare old regime (with indexation) vs new regime. Includes reverse calculator — find the <strong>sale price for zero tax</strong>.</p>
</section>

<div class="wrap">
  <!-- LEFT: INPUT -->
  <div>
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#EFF6FF">💰</div>
        <div><h2>Capital Gains Calculator</h2><p>IT Act 2025 · Tax Year 2026-27</p></div>
      </div>
      <div class="card-body">

        <div class="info-box">ℹ️ Under IT Act 2025: LTCG on property is <strong>12.5% without indexation</strong>. Old regime (20% with indexation) applicable for assets purchased before 23 Jul 2023. Choose as applicable.</div>

        <div class="field">
          <label>Asset Type</label>
          <select id="assetType" onchange="updateAssetUI()">
            <option value="property">Immovable Property (Land/Building)</option>
            <option value="equity">Listed Equity Shares / Equity MF</option>
            <option value="debt">Debt Mutual Funds / Bonds</option>
            <option value="gold">Gold / Gold ETF / Sovereign Gold Bond</option>
            <option value="unlisted">Unlisted Shares</option>
            <option value="other">Other Capital Assets</option>
          </select>
        </div>

        <div class="row2">
          <div class="field">
            <label>Date of Purchase</label>
            <input type="text" id="purchaseDate" placeholder="DD/MM/YYYY"/>
            <p class="hint">Original acquisition date</p>
          </div>
          <div class="field">
            <label>Date of Sale</label>
            <input type="text" id="saleDate" placeholder="DD/MM/YYYY"/>
            <p class="hint">Date of transfer/sale</p>
          </div>
        </div>

        <div class="row2">
          <div class="field">
            <label>Purchase Price (₹)</label>
            <input type="number" id="purchasePrice" placeholder="e.g. 2000000" min="0"/>
            <p class="hint">Cost of acquisition</p>
          </div>
          <div class="field">
            <label>Sale Price (₹)</label>
            <input type="number" id="salePrice" placeholder="e.g. 5000000" min="0"/>
            <p class="hint">Full value of consideration</p>
          </div>
        </div>

        <div class="row2">
          <div class="field">
            <label>Improvement Cost (₹)</label>
            <input type="number" id="improveCost" placeholder="0" min="0" value="0"/>
            <p class="hint">Cost of any improvements made</p>
          </div>
          <div class="field">
            <label>Transfer Expenses (₹)</label>
            <input type="number" id="transferCost" placeholder="0" min="0" value="0"/>
            <p class="hint">Brokerage, registration, legal fees</p>
          </div>
        </div>

        <div class="field" id="exemptionField">
          <label>Exemption Claimed</label>
          <select id="exemptionType">
            <option value="0">None</option>
            <option value="54">Sec 54 — Residential House Property</option>
            <option value="54B">Sec 54B — Agricultural Land</option>
            <option value="54EC">Sec 54EC — NHAI/REC Bonds (Max ₹50L)</option>
            <option value="54F">Sec 54F — Any LTCG → Residential House</option>
          </select>
        </div>
        <div class="field" id="exemptionAmtField">
          <label>Exemption Amount (₹)</label>
          <input type="number" id="exemptionAmt" placeholder="0" min="0" value="0"/>
          <p class="hint">Amount claimed under selected exemption</p>
        </div>

        <div class="row2">
          <div class="field">
            <label>Tax Year (Previous Year)</label>
            <select id="taxYear" onchange="updateCIITable()">
              <option value="2025-26" selected>PY 2025-26 (AY 2026-27)</option>
              <option value="2026-27">PY 2026-27 (AY 2027-28) · Future</option>
              <option value="2024-25">PY 2024-25 (AY 2025-26)</option>
              <option value="2023-24">PY 2023-24 (AY 2024-25)</option>
            </select>
            <p class="hint">FY in which asset is sold / will be sold</p>
          </div>
          <div class="field">
            <label>Assessee Type</label>
            <select id="assesseeType">
              <option value="individual">Individual / HUF</option>
              <option value="firm">Firm / LLP</option>
              <option value="company">Company</option>
            </select>
          </div>
        </div>

        <button class="btn" onclick="calcCG()">Calculate Capital Gains →</button>

        <!-- RESULTS -->
        <div id="cgResults" style="display:none;margin-top:16px">
          <div id="cgTypeLabel" style="font-size:13px;font-weight:700;margin-bottom:10px;color:var(--ink)"></div>

          <!-- Regime comparison -->
          <div class="compare-grid" id="compareGrid"></div>

          <!-- Detail breakdown -->
          <div id="detailBreakdown"></div>

          <!-- Reverse calculator result -->
          <div class="reverse-box" id="reverseBox" style="display:none">
            <h3>🔄 Reverse Calculator — Zero Tax Sale Price</h3>
            <div class="rval" id="revSalePrice"></div>
            <div class="rsub" id="revSub"></div>
          </div>

          <div class="note-box" id="cgNote"></div>
        </div>

      </div>
    </div>
  </div>

  <!-- RIGHT: INFO PANELS -->
  <div>
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#FFFBEB">📋</div>
        <div><h2>Tax Rates — IT Act 2025</h2><p>LTCG &amp; STCG at a glance</p></div>
      </div>
      <div class="card-body" style="padding:0;overflow-x:auto">
        <table class="cii-table">
          <thead><tr><th>Asset</th><th>Holding</th><th>Rate</th><th>Threshold</th></tr></thead>
          <tbody>
            <tr><td>Immovable Property</td><td>&gt;24 months</td><td>12.5% (no idx) / 20% (with idx pre Jul-23)</td><td>₹1.25L (old ₹1L)</td></tr>
            <tr><td>Listed Equity / Eq MF</td><td>&gt;12 months</td><td>12.5%</td><td>₹1.25L exempt</td></tr>
            <tr><td>Listed Equity / Eq MF</td><td>≤12 months</td><td>20% (STCG)</td><td>Nil</td></tr>
            <tr><td>Debt MF / Bonds</td><td>Any</td><td>Slab rate</td><td>Nil</td></tr>
            <tr><td>Gold / Gold ETF</td><td>&gt;24 months</td><td>12.5%</td><td>Nil</td></tr>
            <tr><td>Unlisted Shares</td><td>&gt;24 months</td><td>12.5%</td><td>Nil</td></tr>
            <tr><td>Unlisted Shares</td><td>≤24 months</td><td>Slab rate</td><td>Nil</td></tr>
            <tr><td>Any STCG (others)</td><td>≤24/36 months</td><td>Slab rate</td><td>Nil</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#F5F3FF">📈</div>
        <div><h2>Cost Inflation Index (CII)</h2><p>As notified by CBDT</p></div>
      </div>
      <div class="card-body" style="padding:0;max-height:260px;overflow-y:auto">
        <table class="cii-table">
          <thead><tr><th>FY</th><th>CII</th><th>FY</th><th>CII</th></tr></thead>
          <tbody>
            <tr><td>2001-02</td><td>100</td><td>2014-15</td><td>240</td></tr>
            <tr><td>2002-03</td><td>105</td><td>2015-16</td><td>254</td></tr>
            <tr><td>2003-04</td><td>109</td><td>2016-17</td><td>264</td></tr>
            <tr><td>2004-05</td><td>113</td><td>2017-18</td><td>272</td></tr>
            <tr><td>2005-06</td><td>117</td><td>2018-19</td><td>280</td></tr>
            <tr><td>2006-07</td><td>122</td><td>2019-20</td><td>289</td></tr>
            <tr><td>2007-08</td><td>129</td><td>2020-21</td><td>301</td></tr>
            <tr><td>2008-09</td><td>137</td><td>2021-22</td><td>317</td></tr>
            <tr><td>2009-10</td><td>148</td><td>2022-23</td><td>331</td></tr>
            <tr><td>2010-11</td><td>167</td><td>2023-24</td><td>348</td></tr>
            <tr><td>2011-12</td><td>184</td><td>2024-25</td><td>363</td></tr>
            <tr><td>2012-13</td><td>200</td><td>2025-26</td><td>380</td></tr>
            <tr><td>2013-14</td><td>220</td><td>2026-27</td><td>TBA</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#F0FDF4">🏠</div>
        <div><h2>Key Exemptions</h2><p>Reduce your capital gains tax</p></div>
      </div>
      <div class="card-body">
        <div style="font-size:12px;line-height:2;color:var(--muted)">
          <p><strong style="color:var(--ink)">Sec 54</strong> — LTCG on residential property → invest in new house (within 2yr purchase / 3yr construct)</p>
          <p><strong style="color:var(--ink)">Sec 54B</strong> — LTCG on agricultural land → invest in new agricultural land</p>
          <p><strong style="color:var(--ink)">Sec 54EC</strong> — Any LTCG → NHAI/REC bonds (max ₹50L, lock-in 5yr)</p>
          <p><strong style="color:var(--ink)">Sec 54F</strong> — LTCG on any asset → invest in residential house (net consideration)</p>
          <p style="margin-top:8px;color:var(--red)"><strong>Note:</strong> Exemptions available only on LTCG. Claim only if actually investing.</p>
        </div>
      </div>
    </div>
  </div>
</div>

<footer>
  <div class="ft-main">
    <div>
      <div class="ft-brand-name">CA Toolkit</div>
      <p class="ft-brand-desc">CA Toolkit is a comprehensive utility platform built by a CA Article from Ludhiana, Punjab, providing automation tools for Indian Chartered Accountants. The platform saves hours of manual work every year — from Balance Sheet year-shift to GST reconciliation, tax calculations, and more.</p>
    </div>
    <div>
      <div class="ft-col-title">Know More</div>
      <ul class="ft-links">
        <li><a href="/">Home</a></li>
        <li><a href="/">BS Year Shift</a></li>
        <li><a href="/tool/tb-to-bs">TB → Balance Sheet</a></li>
        <li><a href="/tool/tax-calculator">Tax Calculator</a></li>
        <li><a href="/privacy">Privacy Policy</a></li>
      </ul>
    </div>
    <div>
      <div class="ft-col-title">Contact Us</div>
      <div class="ft-contact-name">CA Toolkit</div>
      <div class="ft-contact-addr">Built for Indian Chartered Accountants<br/>Created by CA Article · Ludhiana, Punjab</div>
    </div>
  </div>
  <div class="ft-bottom">
    <span class="ft-bottom-left">©2026 CA Toolkit · All Rights Reserved</span>
    <span class="ft-bottom-right">Built for Indian CAs · Ludhiana, Punjab</span>
  </div>
</footer>

<script>
// CII table
const CII = {2001:100,2002:105,2003:109,2004:113,2005:117,2006:122,2007:129,2008:137,
             2009:148,2010:167,2011:184,2012:200,2013:220,2014:240,2015:254,2016:264,
             2017:272,2018:280,2019:289,2020:301,2021:317,2022:331,2023:348,2024:363,
             2025:380, 2026:400}; // CII 2026 = estimated, not yet notified by CBDT

const TAX_YEAR_FY = {"2023-24":2023,"2024-25":2024,"2025-26":2025,"2026-27":2026};
const TAX_YEAR_AY = {"2023-24":"AY 2024-25","2024-25":"AY 2025-26","2025-26":"AY 2026-27","2026-27":"AY 2027-28"};

function getSelectedSaleFY(){
  const el = document.getElementById("taxYear");
  return TAX_YEAR_FY[(el||{value:"2025-26"}).value] || 2025;
}

function updateCIITable(){
  const saleFY = getSelectedSaleFY();
  document.querySelectorAll(".cii-table tr").forEach(tr=>{
    const c = tr.cells[0]; if(!c) return;
    const yr = parseInt((c.textContent||"").replace("FY ","").split("-")[0]);
    tr.classList.toggle("highlight", yr === saleFY);
  });
  if(document.getElementById("cgResults").style.display !== "none") calcCG();
}

const fmt   = n => "₹"+Math.round(n).toLocaleString("en-IN");
const fmtPct= n => n.toFixed(2)+"%";

function parseMyDate(s){
  if(!s) return null;
  const p=String(s).trim().split("/");
  if(p.length===3) return new Date(parseInt(p[2]),parseInt(p[1])-1,parseInt(p[0]));
  return new Date(s);
}

function getFY(d){ return d.getMonth()>=3 ? d.getFullYear() : d.getFullYear()-1; }

function holdingMonths(d1,d2){
  return (d2.getFullYear()-d1.getFullYear())*12 + (d2.getMonth()-d1.getMonth());
}

function getCII(fy){ return CII[fy] || 380; }

function updateAssetUI(){
  const t = document.getElementById("assetType").value;
  const showExemption = ["property","equity","other","gold","unlisted"].includes(t);
  document.getElementById("exemptionField").style.display    = showExemption?"block":"none";
  document.getElementById("exemptionAmtField").style.display = showExemption?"block":"none";
}

function calcCG(){
  const asset     = document.getElementById("assetType").value;
  const pd        = parseMyDate(document.getElementById("purchaseDate").value);
  const sd        = parseMyDate(document.getElementById("saleDate").value);
  const pp        = parseFloat(document.getElementById("purchasePrice").value)||0;
  const sp        = parseFloat(document.getElementById("salePrice").value)||0;
  const ic        = parseFloat(document.getElementById("improveCost").value)||0;
  const tc        = parseFloat(document.getElementById("transferCost").value)||0;
  const exemAmt   = parseFloat(document.getElementById("exemptionAmt").value)||0;
  const exemType  = document.getElementById("exemptionType").value;
  const assessee  = document.getElementById("assesseeType").value;

  if(!pd||!sd){ alert("Please enter both purchase and sale dates."); return; }
  if(!pp||!sp){ alert("Please enter purchase price and sale price."); return; }
  if(sd<=pd){ alert("Sale date must be after purchase date."); return; }

  const months = holdingMonths(pd,sd);
  const pyFY   = getFY(pd);
  const selectedTY  = (document.getElementById("taxYear")||{value:"2025-26"}).value||"2025-26";
  const syFY        = getSelectedSaleFY();
  const ayLabel     = TAX_YEAR_AY[selectedTY]||"AY 2026-27";
  const isFutureTY  = selectedTY === "2026-27";

  // Determine LTCG/STCG threshold
  let ltcgMonths = 24;
  if(asset==="equity") ltcgMonths = 12;
  const isLTCG = months >= ltcgMonths;
  const cgType = isLTCG ? "Long-Term Capital Gain (LTCG)" : "Short-Term Capital Gain (STCG)";

  // Net sale consideration
  const netSale = sp - tc;

  // ── New Regime (no indexation) ────────────────────────────────────────────
  let newCOA = pp + ic;
  let newCG  = Math.max(0, netSale - newCOA - exemAmt);
  let newRate, newExempt=0, newTax=0;

  if(!isLTCG){
    // STCG
    if(asset==="equity") newRate=20;
    else newRate=0; // slab
  } else {
    if(asset==="equity"){ newRate=12.5; newExempt=125000; }
    else if(asset==="debt") newRate=0; // slab
    else newRate=12.5;
  }
  if(newRate>0){
    const taxableNew = Math.max(0, newCG - newExempt);
    newTax = Math.round(taxableNew * newRate / 100);
  }

  // ── Old Regime (with indexation) — only for LTCG on property purchased pre Jul 2023 ──
  const preJul23 = pd < new Date(2023,6,23);
  const showOldRegime = isLTCG && (asset==="property"||asset==="other"||asset==="gold") && preJul23;
  let oldCOA=0, oldCG=0, oldRate=0, oldExempt=0, oldTax=0;

  if(showOldRegime){
    const ciiPurchase = getCII(pyFY);
    const ciiSale     = getCII(syFY);
    oldCOA  = Math.round((pp + ic) * ciiSale / ciiPurchase);
    oldCG   = Math.max(0, netSale - oldCOA - exemAmt);
    oldRate = 20;
    const taxableOld = Math.max(0, oldCG - oldExempt);
    oldTax  = Math.round(taxableOld * oldRate / 100);
  }

  // ── Reverse calculator ─────────────────────────────────────────────────────
  // Min sale price for zero tax (new regime)
  let zeroTaxSale = null;
  if(newRate>0 && isLTCG){
    zeroTaxSale = newCOA + tc + newExempt + exemAmt;
    if(asset==="equity") zeroTaxSale += 125000;
  }

  // ── Render results ─────────────────────────────────────────────────────────
  document.getElementById("cgTypeLabel").innerHTML =
    `<span style="background:${isLTCG?"#EFF6FF":"#FFFBEB"};color:${isLTCG?"var(--brand)":"#92400E"};
     padding:4px 12px;border-radius:99px;font-size:12px">${cgType} · ${months} months holding</span>
     <span style="margin-left:8px;background:#F0FDF4;color:#065F46;padding:4px 12px;border-radius:99px;font-size:12px;font-weight:600">
       ${selectedTY} (${ayLabel})${isFutureTY?" · Projected":""}
     </span>`;

  // Compare grid
  let compareHTML = "";
  const winner = (showOldRegime && oldTax < newTax) ? "old" : "new";

  compareHTML += `<div class="cbox ${winner==="new"?"winner":""}">
    ${winner==="new"?'<span class="winner-badge">✓ Lower Tax</span><br>':""}
    <h3>New Regime (No Indexation)</h3>
    <div class="ctax">${newRate===0?"Slab Rate":fmt(newTax)}</div>
    <div class="csub">${newRate===0?"Tax at applicable slab rate":newRate+"% on "+fmt(Math.max(0,newCG-newExempt))}</div>
  </div>`;

  if(showOldRegime){
    compareHTML += `<div class="cbox ${winner==="old"?"winner":""}">
      ${winner==="old"?'<span class="winner-badge">✓ Lower Tax</span><br>':""}
      <h3>Old Regime (With Indexation)</h3>
      <div class="ctax">${fmt(oldTax)}</div>
      <div class="csub">20% on ${fmt(Math.max(0,oldCG))} (Indexed COA: ${fmt(oldCOA)})</div>
    </div>`;
  } else {
    compareHTML += `<div class="cbox" style="background:#F9FAFB;border-style:dashed">
      <h3 style="color:var(--muted)">Old Regime (Indexation)</h3>
      <div class="ctax" style="color:var(--muted);font-size:14px">Not Applicable</div>
      <div class="csub">${!isLTCG?"STCG — no indexation benefit":!preJul23?"Asset purchased after 23 Jul 2023":"Not applicable for this asset type"}</div>
    </div>`;
  }
  document.getElementById("compareGrid").innerHTML = compareHTML;

  // Detail breakdown (new regime)
  let detailHTML = `<div class="card" style="margin-top:12px">
    <div class="card-head"><div class="icon" style="background:#EFF6FF">🧮</div>
    <div><h2>Computation (New Regime)</h2><p>Step-by-step breakdown</p></div></div>
    <div class="card-body" style="padding:12px 16px">
    <table class="dtable">
      <tr><td>Full Value of Consideration (Sale Price)</td><td>${fmt(sp)}</td></tr>
      <tr><td>Less: Transfer Expenses</td><td>(${fmt(tc)})</td></tr>
      <tr><td>Net Sale Consideration</td><td>${fmt(netSale)}</td></tr>
      <tr><td>Less: Cost of Acquisition</td><td>(${fmt(pp)})</td></tr>
      <tr><td>Less: Cost of Improvement</td><td>(${fmt(ic)})</td></tr>
      <tr><td>Capital Gain (Before Exemption)</td><td>${fmt(Math.max(0,netSale-pp-ic))}</td></tr>
      ${exemAmt>0?`<tr><td>Less: Exemption u/s ${exemType}</td><td>(${fmt(exemAmt)})</td></tr>`:""}
      <tr><td>Taxable Capital Gain</td><td>${fmt(newCG)}</td></tr>
      ${newExempt>0?`<tr><td>Less: Basic Exemption (₹1.25L)</td><td>(${fmt(Math.min(newExempt,newCG))})</td></tr>`:""}
      <tr><td>Tax @ ${newRate===0?"Slab":newRate+"%"}</td><td><strong>${newRate===0?"As per slab":fmt(newTax)}</strong></td></tr>
    </table></div></div>`;

  if(showOldRegime){
    const ciiP = getCII(pyFY), ciiS = getCII(syFY);
    detailHTML += `<div class="card" style="margin-top:12px">
      <div class="card-head"><div class="icon" style="background:#F5F3FF">📊</div>
      <div><h2>Computation (Old Regime with Indexation)</h2><p>CII ${pyFY}-${pyFY+1}: ${ciiP} → ${syFY}-${syFY+1}: ${ciiS}</p></div></div>
      <div class="card-body" style="padding:12px 16px">
      <table class="dtable">
        <tr><td>Net Sale Consideration</td><td>${fmt(netSale)}</td></tr>
        <tr><td>Indexed Cost of Acquisition (${pp} × ${ciiS}/${ciiP})</td><td>(${fmt(oldCOA)})</td></tr>
        <tr><td>Capital Gain (After Indexation)</td><td>${fmt(oldCG)}</td></tr>
        ${exemAmt>0?`<tr><td>Less: Exemption u/s ${exemType}</td><td>(${fmt(exemAmt)})</td></tr>`:""}
        <tr><td>Tax @ 20%</td><td><strong>${fmt(oldTax)}</strong></td></tr>
      </table></div></div>`;
  }

  document.getElementById("detailBreakdown").innerHTML = detailHTML;

  // Reverse calculator
  if(zeroTaxSale!==null && zeroTaxSale>0){
    document.getElementById("reverseBox").style.display = "block";
    document.getElementById("revSalePrice").textContent = fmt(zeroTaxSale);
    document.getElementById("revSub").textContent =
      `Sell at or below ${fmt(zeroTaxSale)} to have zero capital gains tax liability (new regime, excluding transfer expenses in computation)`;
  } else {
    document.getElementById("reverseBox").style.display = "none";
  }

  // Note
  let note = "";
  if(newRate===0) note = "Tax at applicable slab rate. Add capital gain to total income and apply applicable tax slab.";
  else if(!isLTCG) note = "Short-term capital gain — taxed at "+newRate+"% (equity) or slab rate (others).";
  else note = `LTCG taxed at ${newRate}%. ${showOldRegime?"Both regimes shown — choose the one with lower tax.":""}`;
  if(exemAmt>0) note += ` Exemption of ${fmt(exemAmt)} claimed u/s ${exemType}.`;
  document.getElementById("cgNote").textContent = "⚠ " + note + " This is an estimate — verify with actual CII and consult your CA.";

  document.getElementById("cgResults").style.display = "block";
  document.getElementById("cgResults").scrollIntoView({behavior:"smooth"});
}

updateAssetUI();
</script>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
</body></html>"""


ADMIN_T = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin – CA Toolkit</title>
<style>
""" + BASE_CSS + """
.wrap{max-width:1100px;margin:0 auto;padding:28px 24px}
h1{font-size:20px;font-weight:800;margin-bottom:4px}
.sub{font-size:13px;color:var(--muted);margin-bottom:24px}
.alert{padding:10px 14px;border-radius:8px;font-size:13px;margin-bottom:16px}
.as{background:#ECFDF5;border:1px solid #A7F3D0;color:#065F46}
.ae{background:#FEF2F2;border:1px solid #FECACA;color:#991B1B}
.section{background:var(--white);border:1px solid var(--border);border-radius:var(--radius);margin-bottom:20px;overflow:hidden}
.sec-head{padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between}
.sec-head h2{font-size:14px;font-weight:700}
.sec-body{padding:18px}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);margin-bottom:4px}
.field{margin-bottom:12px}
input[type=text],input[type=password],select{width:100%;border:1.5px solid var(--border);border-radius:8px;
  padding:8px 11px;font-family:inherit;font-size:13px;color:var(--ink);background:var(--white);outline:none;transition:border-color .2s}
input:focus,select:focus{border-color:var(--brand)}
.form-row{display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:10px;align-items:end}
@media(max-width:640px){.form-row{grid-template-columns:1fr}}
.btn{background:var(--brand);color:#fff;border:none;border-radius:8px;padding:9px 16px;
     font-family:inherit;font-size:13px;font-weight:600;cursor:pointer;transition:background .2s}
.btn:hover{background:var(--brand-d)}
.sm{padding:5px 10px;font-size:11px;border-radius:6px;border:none;cursor:pointer;font-family:inherit;font-weight:600}
.sg{background:#ECFDF5;color:#065F46}.sg:hover{background:#A7F3D0}
.rr{background:#FEF2F2;color:#991B1B}.rr:hover{background:#FECACA}
.am{background:#EFF6FF;color:var(--brand)}.am:hover{background:#BFDBFE}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:left;font-size:10px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);
   border-bottom:1.5px solid var(--border);padding:7px 10px}
td{padding:10px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#F9FAFB}
.bdg{display:inline-block;font-size:10px;font-weight:700;padding:2px 7px;border-radius:99px;text-transform:uppercase}
.b-free{background:#F3F4F6;color:var(--muted)}
.b-starter{background:#ECFDF5;color:#065F46}
.b-standard{background:#EFF6FF;color:var(--brand)}
.b-pro{background:#FFFBEB;color:#92400E}
.b-firm{background:#F5F3FF;color:#5B21B6}
.warn{color:var(--red);font-weight:700}
.ok{color:var(--green);font-weight:700}
.progress-wrap{display:flex;align-items:center;gap:8px;min-width:120px}
.progress-bg{flex:1;background:#F3F4F6;border-radius:99px;height:5px;overflow:hidden}
.progress-fill{height:100%;border-radius:99px}
</style></head><body>
<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <div class="nav-right">
    <a href="/" class="nav-link">Dashboard</a>
    <a href="/logout" class="nav-link">Sign out</a>
  </div>
</nav>
<div class="wrap">
  <h1>⚙ Admin Panel</h1>
  <p class="sub">Create accounts, manage plans, track usage.</p>
  {% if msg %}<div class="alert {{ 'ae' if 'error' in msg.lower() or 'cannot' in msg.lower() else 'as' }}">{{ msg }}</div>{% endif %}

  <!-- CREATE USER -->
  <div class="section">
    <div class="sec-head"><h2>➕ Create New User</h2></div>
    <div class="sec-body">
      <form method="POST" action="/admin/create">
        <div class="form-row">
          <div class="field"><label>Username</label>
            <input type="text" name="username" placeholder="e.g. rahul_ca" required/></div>
          <div class="field"><label>Password</label>
            <input type="password" name="password" placeholder="Min 6 chars" required/></div>
          <div class="field"><label>Plan</label>
            <select name="plan">
              <option value="free">Free (2 uploads)</option>
              <option value="starter">Starter (10 uploads · ₹60)</option>
              <option value="standard" selected>Standard (25 uploads · ₹130)</option>
              <option value="pro">Professional (60 uploads · ₹270)</option>
              <option value="firm">Firm (150 uploads · ₹600)</option>
              <option value="ca">CA Admin (500 uploads · ₹1000)</option>
            </select></div>
          <div class="field"><label>&nbsp;</label>
            <button class="btn" type="submit">Create User</button></div>
        </div>
      </form>
    </div>
  </div>

  <!-- USERS TABLE -->
  <div class="section">
    <div class="sec-head">
      <h2>👥 All Users ({{ users|length }})</h2>
      <span style="font-size:12px;color:var(--muted)">Uploads remaining shown in green/red</span>
    </div>
    <div class="sec-body" style="padding:0;overflow-x:auto">
      <table>
        <thead><tr>
          <th>#</th><th>Username</th><th>Plan</th>
          <th>Uploads Used</th><th>Remaining</th>
          <th>Valid Till</th><th>Joined</th><th>Actions</th>
        </tr></thead>
        <tbody>
        {% for u in users %}
        <tr>
          <td style="color:var(--muted)">{{ u.id }}</td>
          <td><strong>{{ u.username }}</strong>
            {% if u.is_admin %}<span class="bdg" style="background:#EFF6FF;color:var(--brand);margin-left:4px">Admin</span>{% endif %}
          </td>
          <td><span class="bdg b-{{ u.plan }}">{{ u.plan }}</span></td>
          <td>
            <div class="progress-wrap">
              <div class="progress-bg">
                <div class="progress-fill"
                     style="width:{{ [u.uploads_used*100//u.uploads_total if u.uploads_total else 0, 100]|min }}%;
                            background:{{ '#EF4444' if u.remaining==0 else '#F59E0B' if u.remaining<=3 else '#10B981' }}">
                </div>
              </div>
              <span style="font-size:11px;white-space:nowrap">{{ u.uploads_used }} / {{ u.uploads_total }}</span>
            </div>
          </td>
          <td class="{{ 'warn' if u.remaining==0 else 'ok' if u.remaining > 5 else '' }}">
            {{ u.remaining }}
          </td>
          <td style="font-size:11px;color:var(--muted)">
            {{ u.validity_end[:10] if u.validity_end else '—' }}
          </td>
          <td style="font-size:11px;color:var(--muted)">{{ u.created_at[:10] }}</td>
          <td>
            {% if not u.is_admin %}
            <div style="display:flex;gap:6px;flex-wrap:wrap">
              <!-- Add uploads dropdown -->
              <form method="POST" action="/admin/addplan" style="display:flex;gap:4px;align-items:center">
                <input type="hidden" name="uid" value="{{ u.id }}"/>
                <select name="plan" style="padding:4px 6px;font-size:11px;border-radius:6px;border:1px solid var(--border);width:auto">
                  <option value="starter">+10</option>
                  <option value="standard" selected>+25</option>
                  <option value="pro">+60</option>
                  <option value="firm">+150</option>
                  <option value="ca">+500</option>
                </select>
                <button class="sm am" type="submit">Add</button>
              </form>
              <form method="POST" action="/admin/delete" style="display:inline"
                    onsubmit="return confirm('Delete {{ u.username }}? This cannot be undone.')">
                <input type="hidden" name="uid" value="{{ u.id }}"/>
                <button class="sm rr" type="submit">Delete</button>
              </form>
            </div>
            {% else %}
            <span style="font-size:11px;color:var(--muted)">—</span>
            {% endif %}
          </td>
        </tr>
        {% endfor %}
        </tbody>
      </table>
    </div>
  </div>
</div>
<footer>
  <div class="ft-bottom" style="justify-content:center">
    <span class="ft-bottom-left">©2026 CA Toolkit · Admin Panel · All Rights Reserved</span>
  </div>
</footer>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
</body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES — AUTH
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/login", methods=["GET"])
def login_page():
    if "uid" in session: return redirect(url_for("dashboard"))
    return render_template_string(LOGIN_T, error=None, email=CONTACT_EMAIL)

@app.route("/login", methods=["POST"])
def login_post():
    u = request.form.get("username", "").strip()
    p = request.form.get("password", "")
    user = get_user_by_name(u)
    if not user or user["password"] != _hash(p):
        return render_template_string(LOGIN_T, error="Invalid username or password.", email=CONTACT_EMAIL)
    session.clear()
    session["uid"] = user["id"]
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login_page"))

# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES — DASHBOARD & TOOLS
# ══════════════════════════════════════════════════════════════════════════════

def user_ctx(user):
    """Build common template context for a user."""
    used  = user["uploads_used"]
    total = user["uploads_total"]
    left  = uploads_remaining(user)
    pct   = min(int(used * 100 / total) if total else 0, 100)
    return dict(
        username=user["username"],
        plan=user["plan"],
        plan_label=PLANS.get(user["plan"], {}).get("label", user["plan"].title()),
        is_admin=bool(user["is_admin"]),
        uploads_used=used,
        uploads_total=total,
        uploads_left=left,
        uploads_remaining=left,
        bar_pct=pct,
        validity_end=user["validity_end"],
        contact_email=CONTACT_EMAIL,
        contact_upi=CONTACT_UPI,
    )

@app.route("/")
def dashboard():
    if "uid" in session:
        user = get_user_by_id(session["uid"])
        ctx = user_ctx(user)
    else:
        ctx = dict(
            username=None, plan="free", plan_label="Free",
            is_admin=False, uploads_used=0, uploads_total=2,
            uploads_left=2, uploads_remaining=2, bar_pct=0,
            validity_end=None, contact_email=CONTACT_EMAIL,
            contact_upi=CONTACT_UPI,
        )
    return render_template_string(DASHBOARD_T, **ctx)


@app.route("/bs-shift")
def bs_shift_redirect():
    return redirect("/")

@app.route("/privacy")
def privacy_page():
    if "uid" in session:
        user = get_user_by_id(session["uid"])
        ctx = user_ctx(user)
    else:
        ctx = dict(username=None, plan="free", plan_label="Free",
            is_admin=False, uploads_used=0, uploads_total=2,
            uploads_left=2, uploads_remaining=2, bar_pct=0,
            validity_end=None, contact_email=CONTACT_EMAIL, contact_upi=CONTACT_UPI)
    return render_template_string(PRIVACY_TEMPLATE, **ctx)

@app.route("/tool/converter")
@login_required
def tool_converter():
    user = get_user_by_id(session["uid"])
    return render_template_string(CONVERTER_T, **user_ctx(user))

@app.route("/tool/tax-calculator")
def tool_tax_calculator():
    if "uid" in session:
        user = get_user_by_id(session["uid"])
        ctx = user_ctx(user)
    else:
        ctx = dict(
            username=None, plan="free", plan_label="Free",
            is_admin=False, uploads_used=0, uploads_total=2,
            uploads_left=2, uploads_remaining=2, bar_pct=0,
            validity_end=None, contact_email=CONTACT_EMAIL,
            contact_upi=CONTACT_UPI,
        )
    return render_template_string(TAX_CALC_T, **ctx)

@app.route("/tool/tds-calculator")
def tool_tds_calculator():
    if "uid" in session:
        user = get_user_by_id(session["uid"])
        ctx = user_ctx(user)
    else:
        ctx = dict(
            username=None, plan="free", plan_label="Free",
            is_admin=False, uploads_used=0, uploads_total=2,
            uploads_left=2, uploads_remaining=2, bar_pct=0,
            validity_end=None, contact_email=CONTACT_EMAIL,
            contact_upi=CONTACT_UPI,
        )
    return render_template_string(TDS_CALC_T, **ctx)

@app.route("/tool/depreciation-calculator")
def tool_depreciation_calculator():
    if "uid" in session:
        user = get_user_by_id(session["uid"])
        ctx = user_ctx(user)
    else:
        ctx = dict(
            username=None, plan="free", plan_label="Free",
            is_admin=False, uploads_used=0, uploads_total=2,
            uploads_left=2, uploads_remaining=2, bar_pct=0,
            validity_end=None, contact_email=CONTACT_EMAIL,
            contact_upi=CONTACT_UPI,
        )
    return render_template_string(DEP_CALC_T, **ctx)

# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES — PROCESS & DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════

def _convert_xls_to_xlsx(xls_path, xlsx_path):
    """Convert legacy .xls to .xlsx using xlrd + openpyxl with formatting."""
    import xlrd
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    wb_xls = xlrd.open_workbook(xls_path, formatting_info=True)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    colour_map = wb_xls.colour_map
    def xlrd_color_to_hex(idx):
        if idx is None or idx < 8 or idx > 63: return None
        rgb = colour_map.get(idx)
        if rgb and rgb != (0, 0, 0): return f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}'
        return None

    border_styles = {0: None, 1: 'thin', 2: 'medium', 3: 'dashed',
                     4: 'dotted', 5: 'thick', 6: 'double', 7: 'hair'}

    for sheet_name in wb_xls.sheet_names():
        xls_ws = wb_xls.sheet_by_name(sheet_name)
        xlsx_ws = wb_out.create_sheet(title=sheet_name)

        for row_idx in range(xls_ws.nrows):
            for col_idx in range(xls_ws.ncols):
                ctype = xls_ws.cell_type(row_idx, col_idx)
                if ctype == xlrd.XL_CELL_EMPTY:
                    continue
                val = xls_ws.cell_value(row_idx, col_idx)
                cell = xlsx_ws.cell(row=row_idx + 1, column=col_idx + 1)

                if ctype == xlrd.XL_CELL_NUMBER and val == int(val):
                    val = int(val)
                if ctype == xlrd.XL_CELL_DATE:
                    try:
                        from datetime import datetime
                        dt = xlrd.xldate_as_tuple(val, wb_xls.datemode)
                        cell.value = datetime(*dt)
                    except:
                        cell.value = val
                else:
                    cell.value = val

                # Copy formatting
                try:
                    xf = wb_xls.xf_list[xls_ws.cell_xf_index(row_idx, col_idx)]
                    font_xls = wb_xls.font_list[xf.font_index]

                    cell.font = Font(
                        name=font_xls.name or 'Calibri',
                        size=font_xls.height / 20 if font_xls.height else 11,
                        bold=font_xls.bold, italic=font_xls.italic,
                        underline='single' if font_xls.underline_type else None,
                    )

                    ha = {0: None, 1: 'left', 2: 'center', 3: 'right',
                          5: 'justify'}.get(xf.alignment.hor_align)
                    va = {0: 'top', 1: 'center', 2: 'bottom'}.get(
                        xf.alignment.vert_align, 'bottom')
                    cell.alignment = Alignment(
                        horizontal=ha, vertical=va,
                        wrap_text=bool(xf.alignment.text_wrapped),
                        indent=xf.alignment.indent_level,
                    )

                    fmt_str = wb_xls.format_map.get(xf.format_key)
                    if fmt_str:
                        cell.number_format = fmt_str.format_str

                    def _side(style_idx):
                        s = border_styles.get(style_idx)
                        return Side(style=s) if s else Side()
                    brd = xf.border
                    cell.border = Border(
                        left=_side(brd.left_line_style),
                        right=_side(brd.right_line_style),
                        top=_side(brd.top_line_style),
                        bottom=_side(brd.bottom_line_style),
                    )

                    bg_idx = xf.background.pattern_colour_index
                    bg_hex = xlrd_color_to_hex(bg_idx)
                    if bg_hex and xf.background.fill_pattern:
                        cell.fill = PatternFill('solid', fgColor=bg_hex)
                except Exception:
                    pass

        # Merged cells (after data so we don't write to merged cells)
        for crange in xls_ws.merged_cells:
            r1, r2, c1, c2 = crange
            xlsx_ws.merge_cells(
                start_row=r1 + 1, start_column=c1 + 1,
                end_row=r2, end_column=c2)

        # Column widths
        for c, ci in xls_ws.colinfo_map.items():
            if ci.width:
                xlsx_ws.column_dimensions[get_column_letter(c + 1)].width = ci.width / 256

        # Row heights
        for r, rh in xls_ws.rowinfo_map.items():
            if rh.height:
                xlsx_ws.row_dimensions[r + 1].height = rh.height / 20

    wb_out.save(xlsx_path)

@app.route("/process", methods=["POST"])
@login_required
def process_file():
    try:
        user = get_user_by_id(session["uid"])
        if not user["is_admin"] and uploads_remaining(user) <= 0:
            return jsonify({"status": "error",
                "message": f"No uploads remaining. Contact {CONTACT_EMAIL} to recharge."})
        if "file" not in request.files:
            return jsonify({"status": "error", "message": "No file uploaded."})
        f = request.files["file"]
        orig_name = f.filename.lower()
        is_xls = orig_name.endswith(".xls") and not orig_name.endswith(".xlsx")
        if not (orig_name.endswith(".xlsx") or is_xls):
            return jsonify({"status": "error", "message": "Only .xlsx and .xls files are supported."})
        try:
            cy = int(request.form.get("closing_year", 0))
            ny = int(request.form.get("new_year", cy + 1))
            on = request.form.get("output_name", "").strip()
        except ValueError:
            return jsonify({"status": "error", "message": "Invalid year values."})
        if ny != cy + 1:
            return jsonify({"status": "error", "message": "New year must be closing year + 1."})
        h  = uuid.uuid4().hex
        ip = os.path.join(UPLOAD_DIR, f"{h}_in.xlsx")
        op = os.path.join(OUTPUT_DIR, f"{h}_out.xlsx")
        xls_tmp = None
        try:
            if is_xls:
                # Save .xls first, then convert to .xlsx via xlrd + openpyxl
                xls_tmp = os.path.join(UPLOAD_DIR, f"{h}_in.xls")
                f.save(xls_tmp)
                _convert_xls_to_xlsx(xls_tmp, ip)
            else:
                f.save(ip)
        except Exception as e:
            for p in (xls_tmp, ip):
                if p:
                    try: os.remove(p)
                    except: pass
            return jsonify({"status": "error", "message": f"File conversion error: {e}"})
        finally:
            if xls_tmp:
                try: os.remove(xls_tmp)
                except: pass
        # Build clean output filename: strip year suffixes like "2024-25", "2025-26", "2026"
        # so Atultex_Industries_2024-25_2026 → Atultex_Industries_2026.xlsx
        if on:
            base_name = on
        else:
            raw = os.path.splitext(f.filename)[0]
            import re as _re
            base_name = raw
            # Strip trailing year patterns repeatedly (handles double suffixes)
            for _ in range(3):
                base_name = _re.sub(r'[_\-]+\d{4}[-_]\d{2,4}$', '', base_name).strip('_- ')
                base_name = _re.sub(r'[_\-]+\d{4}$', '', base_name).strip('_- ')
        fname = f"{base_name}_{ny}.xlsx"
        try:
            result = process(ip, op, cy, ny)
            # ── FA year-end rollover (mirror source CY into PY, then reset new CY inputs) ──
            _rollover_fixed_assets(op, str(ny), result.get("log", []), source_path=ip)
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)})
        finally:
            try: os.remove(ip)
            except: pass

        log_usage(user["id"], fname)
        return jsonify({"status": "success", "log": result["log"], "file_id": h, "filename": fname})
    except Exception as e:
        # Master catch — ensures we ALWAYS return JSON, never an HTML error page
        return jsonify({"status": "error", "message": f"Unexpected error: {e}"}), 500

@app.route("/download/<fid>")
@login_required
def download(fid):
    if not re.fullmatch(r"[a-f0-9]{32}", fid): return "Invalid ID", 400
    path = os.path.join(OUTPUT_DIR, f"{fid}_out.xlsx")
    if not os.path.exists(path): return "File not found or expired.", 404
    fn = request.args.get("fn", f"bs_shift_{fid[:8]}.xlsx")
    if not fn.endswith(".xlsx"): fn += ".xlsx"
    return send_file(path, as_attachment=True,
        download_name=fn,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES — ADMIN
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/admin")
@admin_required
def admin_panel():
    raw   = all_users()
    users = []
    for u in raw:
        d = dict(u)
        d["remaining"] = uploads_remaining(u)
        users.append(d)
    return render_template_string(ADMIN_T, users=users, msg=request.args.get("msg", ""))

@app.route("/admin/create", methods=["POST"])
@admin_required
def admin_create():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    plan_key = request.form.get("plan", "free")
    if not username or len(username) < 3:
        return redirect(url_for("admin_panel", msg="Username must be at least 3 characters."))
    if not re.match(r"^[a-zA-Z0-9_]+$", username):
        return redirect(url_for("admin_panel", msg="Username: only letters, numbers, underscores."))
    if len(password) < 6:
        return redirect(url_for("admin_panel", msg="Password must be at least 6 characters."))
    if get_user_by_name(username):
        return redirect(url_for("admin_panel", msg=f"Username '{username}' already exists."))
    if plan_key not in PLANS:
        plan_key = "free"
    create_user(username, password, plan_key)
    plan_label = PLANS[plan_key]["label"]
    return redirect(url_for("admin_panel",
        msg=f"✓ User '{username}' created on {plan_label} plan ({PLANS[plan_key]['uploads']} uploads)."))

@app.route("/admin/addplan", methods=["POST"])
@admin_required
def admin_addplan():
    uid      = int(request.form.get("uid"))
    plan_key = request.form.get("plan", "standard")
    if plan_key not in PLANS: plan_key = "standard"
    user = get_user_by_id(uid)
    if not user: return redirect(url_for("admin_panel", msg="User not found."))
    old_rem = uploads_remaining(user)
    add_uploads(uid, plan_key)
    extra = PLANS[plan_key]["uploads"]
    return redirect(url_for("admin_panel",
        msg=f"✓ Added {extra} uploads to '{user['username']}'. Total remaining: {old_rem + extra}."))

@app.route("/admin/delete", methods=["POST"])
@admin_required
def admin_delete():
    uid = int(request.form.get("uid"))
    if uid == session["uid"]:
        return redirect(url_for("admin_panel", msg="Cannot delete your own account."))
    user = get_user_by_id(uid)
    if not user: return redirect(url_for("admin_panel", msg="User not found."))
    name = user["username"]
    del_user(uid)
    return redirect(url_for("admin_panel", msg=f"✓ User '{name}' deleted."))

@app.route("/tool/msme-calculator")
def tool_msme_calculator():
    if "uid" in session:
        user = get_user_by_id(session["uid"]); ctx = user_ctx(user)
    else:
        ctx = dict(
            username=None, plan="free", plan_label="Free",
            is_admin=False, uploads_used=0, uploads_total=2,
            uploads_left=2, uploads_remaining=2, bar_pct=0,
            validity_end=None, contact_email=CONTACT_EMAIL,
            contact_upi=CONTACT_UPI,
        )
    return render_template_string(MSME_T, **ctx)

@app.route("/tool/capital-gains-calculator")
def tool_cg_calculator():
    if "uid" in session:
        user = get_user_by_id(session["uid"]); ctx = user_ctx(user)
    else:
        ctx = dict(
            username=None, plan="free", plan_label="Free",
            is_admin=False, uploads_used=0, uploads_total=2,
            uploads_left=2, uploads_remaining=2, bar_pct=0,
            validity_end=None, contact_email=CONTACT_EMAIL,
            contact_upi=CONTACT_UPI,
        )
    return render_template_string(CG_CALC_T, **ctx)

# ══════════════════════════════════════════════════════════════════════════════
#  TEMPLATE — GST RECONCILIATION
# ══════════════════════════════════════════════════════════════════════════════

GST_RECON_T = r"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>GST Reconciliation – CA Toolkit</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>
<style>
""" + BASE_CSS + """
.nav-links{display:flex;gap:20px;list-style:none}
.nav-links a{text-decoration:none;color:var(--muted);font-size:13px;font-weight:500;transition:color .2s}
.nav-links a:hover{color:var(--brand)}
.hero{text-align:center;padding:40px 24px 30px;max-width:700px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#FEF3C7;
            color:#92400E;border:1px solid #FDE68A;border-radius:99px;
            padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:18px}
h1{font-size:clamp(22px,3.5vw,34px);font-weight:800;line-height:1.15;letter-spacing:-.5px;margin-bottom:10px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:14px;color:var(--muted);line-height:1.7;max-width:520px;margin:0 auto}
.main{max-width:900px;margin:0 auto;padding:30px 24px}
.card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);
      box-shadow:var(--shadow);overflow:hidden;margin-bottom:24px}
.card-head{padding:14px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.card-head .icon{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px}
.card-head h2{font-size:14px;font-weight:700}
.card-body{padding:20px}
.field{margin-bottom:16px}
label{display:block;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);margin-bottom:5px}
.hint{font-size:11px;color:var(--muted);margin-top:4px}
.dropzone{border:2px dashed var(--border);border-radius:10px;padding:24px 14px;text-align:center;cursor:pointer;transition:all .2s;position:relative;background:var(--bg)}
.dropzone:hover,.dropzone.drag{border-color:var(--brand);background:#EFF6FF}
.dropzone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.dz-icon{font-size:26px;margin-bottom:6px}
.dz-text{font-size:12px;color:var(--muted)}
.dz-text strong{color:var(--brand)}
.dz-file{font-size:12px;font-weight:600;color:var(--green);margin-top:5px;display:none}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
input[type=text]{width:100%;border:1.5px solid var(--border);border-radius:8px;padding:9px 12px;font-size:13px;font-family:inherit;outline:none;transition:border .2s;box-sizing:border-box}
input[type=text]:focus{border-color:var(--brand)}
.btn{width:100%;padding:12px;background:var(--brand);color:#fff;border:none;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;transition:background .2s;font-family:inherit}
.btn:hover{background:#1E40AF}
.btn:disabled{background:#93C5FD;cursor:not-allowed}
.spinner{display:none;width:16px;height:16px;border:2.5px solid #fff;border-top-color:transparent;border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.status{padding:14px 16px;border-radius:10px;font-size:13px;margin-top:16px;display:none;line-height:1.6}
.status.success{background:#ECFDF5;border:1px solid #A7F3D0;color:#065F46}
.status.error{background:#FEF2F2;border:1px solid #FECACA;color:#991B1B}
.dl-link{display:none;margin-top:14px;padding:12px 18px;background:#ECFDF5;border:1px solid #A7F3D0;border-radius:10px;color:#065F46;font-weight:700;font-size:13px;text-decoration:none;text-align:center}
.dl-link:hover{background:#D1FAE5}
.info-box{background:#FEF3C7;border:1px solid #FDE68A;border-radius:10px;padding:14px 16px;font-size:12px;color:#92400E;line-height:1.7;margin-bottom:16px}
.log-list{margin:10px 0 0;padding:0;list-style:none;font-size:12px;line-height:1.8}
.log-list li{padding:2px 0}
.mapping-row{display:grid;grid-template-columns:80px 1fr 30px;gap:8px;align-items:center;margin-bottom:8px}
.mapping-row input{font-size:13px}
.mapping-row .remove-btn{background:none;border:none;color:#EF4444;cursor:pointer;font-size:18px;padding:0}
#add-mapping{background:none;border:1px dashed var(--border);border-radius:8px;padding:8px;font-size:12px;color:var(--muted);cursor:pointer;width:100%;margin-top:4px}
#add-mapping:hover{border-color:var(--brand);color:var(--brand)}
</style></head><body>
<nav class="navbar"><div class="nav-inner">
  <a href="/" class="logo">CA Toolkit</a>
  <ul class="nav-links">
    <li><a href="/">← All Tools</a></li>
    {% if username %}<li><a href="/logout">Logout</a></li>{% endif %}
  </ul>
</div></nav>

<div class="hero">
  <div class="hero-badge">📊 GST Reconciliation</div>
  <h1>Sales <em>Books vs GSTR 3B</em></h1>
  <p>Upload your month-wise sales summary and GSTR 3B PDFs to get an instant reconciliation report showing differences by state and month.</p>
</div>

<div class="main">
  <div class="card">
    <div class="card-head">
      <div class="icon" style="background:#FEF3C7">📄</div>
      <div><h2>Upload Files</h2><p style="font-size:12px;color:var(--muted)">Sales summary (.xlsx) + GSTR 3B PDFs (.zip)</p></div>
    </div>
    <div class="card-body">

      <div class="info-box">
        <strong>📊 Sales Summary:</strong> Month in col A, sales value in col B (or separate column per branch/state).<br>
        <strong>📁 GSTR 3B ZIP:</strong> ZIP with sub-folders named by 2-digit state code (e.g. 05/, 09/). GSTR 3B PDFs inside. Reads Table 3.1 A+B+C+E only (excludes D — reverse charge).
      </div>

      <div style="padding:12px 14px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;margin-bottom:16px">
        <div style="font-size:12px;font-weight:700;color:#065F46;margin-bottom:8px">⬇ Download Sales Format Template</div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a href="/gst-template/consolidated"
             style="display:inline-flex;align-items:center;gap:6px;padding:7px 14px;background:#059669;color:#fff;border-radius:6px;font-size:12px;font-weight:600;text-decoration:none">
            📥 Consolidated Template
          </a>
          <a href="/gst-template/branchwise"
             style="display:inline-flex;align-items:center;gap:6px;padding:7px 14px;background:#0284C7;color:#fff;border-radius:6px;font-size:12px;font-weight:600;text-decoration:none">
            📥 Branch / State-wise Template
          </a>
        </div>
        <div style="font-size:11px;color:#065F46;margin-top:6px">
          Fill in your figures and upload. Do not change column headers or row order.
        </div>
      </div>

      <!-- CONSOLIDATED CHECKBOX -->
      <div class="field" style="margin-bottom:14px">
        <label style="display:flex;align-items:center;gap:10px;cursor:pointer;font-size:13px;font-weight:600;text-transform:none;letter-spacing:0">
          <input type="checkbox" id="consolidated-chk" onchange="onConsolidatedChange()"
            style="width:18px;height:18px;accent-color:var(--brand);cursor:pointer;flex-shrink:0"/>
          <span>
            <strong>Consolidated Sales Data</strong>
            <span style="font-weight:400;color:var(--muted);margin-left:6px">— tick if your Excel has ONE total column for all branches combined</span>
          </span>
        </label>
        <div id="consolidated-hint" style="display:none;margin-top:8px;padding:8px 12px;background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;font-size:12px;color:#1E40AF">
          ✅ <strong>Consolidated mode:</strong> Tool will compare your single total sales figure against the <em>sum of all GSTR 3B states</em> combined. No column mapping needed.
        </div>
        <div id="split-hint" style="margin-top:8px;padding:8px 12px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;font-size:12px;color:#065F46">
          📍 <strong>Location-wise mode:</strong> Map each state code to its column header in your Excel below.
        </div>
      </div>

      <div class="field">
        <label>Sales Summary (Excel)</label>
        <div class="dropzone" id="dz-sales">
          <div class="dz-icon">📊</div>
          <div class="dz-text"><strong>Click or drag</strong> your sales summary .xlsx</div>
          <div class="dz-file" id="sf-sales"></div>
          <input type="file" id="file-sales" accept=".xlsx,.xls" onchange="pickFile(this,'sf-sales')">
        </div>
      </div>

      <div class="field">
        <label>GSTR 3B PDFs (ZIP file)</label>
        <div class="dropzone" id="dz-gst">
          <div class="dz-icon">📁</div>
          <div class="dz-text"><strong>Click or drag</strong> your GSTR 3B ZIP file</div>
          <div class="dz-file" id="sf-gst"></div>
          <input type="file" id="file-gst" accept=".zip" onchange="pickFile(this,'sf-gst')">
        </div>
      </div>

      <div class="field" id="mapping-field">
        <label>State Code → Sales Column Mapping</label>
        <p class="hint" style="margin-bottom:8px">
          Enter the exact column header from your Excel for each state code (e.g. DRH/LDH, HOSUR, RUDRAPUR).
        </p>
        <div id="mapping-container"></div>
        <button id="add-mapping" onclick="addMapping()">+ Add Mapping</button>
      </div>

      <div class="field" id="consolidated-col-field" style="display:none">
        <label>Column Name in your Excel (Sales column header)</label>
        <input type="text" id="consolidated-col-input" placeholder="e.g. Total Sales, Sales, ALL PLANTS — leave blank to auto-detect first numeric column"/>
        <p class="hint">Enter the exact header of the column containing total sales. Leave blank to auto-detect.</p>
      </div>

      <div class="field">
        <label>Output File Name (optional)</label>
        <input type="text" id="output-name" placeholder="GST_Reconciliation">
      </div>

      <button class="btn" id="proc-btn" onclick="doProcess()">
        <span id="bt">⚡ Process & Download</span>
        <div class="spinner" id="sp"></div>
      </button>

      <div class="status" id="status"></div>
      <a class="dl-link" id="dl-link" href="#">⬇ Download</a>
    </div>
  </div>
</div>

<section style="background:var(--white);border-top:1px solid var(--border);border-bottom:1px solid var(--border);padding:48px 24px">
  <h2 style="text-align:center;font-size:24px;font-weight:800;margin-bottom:6px">Simple Pricing</h2>
  <p style="text-align:center;color:var(--muted);font-size:13px;margin-bottom:32px">Upload-based · 3-month validity · Shared across all premium tools</p>
  <div style="max-width:1080px;margin:0 auto;display:grid;grid-template-columns:repeat(6,1fr);gap:14px">
    <div style="border:1.5px solid var(--border);border-radius:var(--radius);padding:20px 16px">
      <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">Free</div>
      <div style="font-size:24px;font-weight:800;margin-bottom:2px">₹0</div>
      <div style="font-size:12px;font-weight:700;color:var(--brand);margin-bottom:2px">2 uploads</div>
      <div style="font-size:10px;color:var(--muted);margin-bottom:14px">Try it out</div>
      <ul style="list-style:none;margin-bottom:16px;font-size:11px"><li style="padding:3px 0">✓ All premium tools</li><li style="padding:3px 0">✓ BS + GST Recon</li></ul>
      <a href="#" style="display:block;text-align:center;padding:8px;border-radius:7px;font-size:12px;font-weight:700;background:var(--bg);color:var(--ink);text-decoration:none;border:1px solid var(--border)">Get Started</a>
    </div>
    <div style="border:1.5px solid var(--border);border-radius:var(--radius);padding:20px 16px">
      <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">Starter</div>
      <div style="font-size:24px;font-weight:800;margin-bottom:2px">₹60</div>
      <div style="font-size:12px;font-weight:700;color:var(--brand);margin-bottom:2px">10 uploads</div>
      <div style="font-size:10px;color:var(--muted);margin-bottom:14px">3 month validity</div>
      <ul style="list-style:none;margin-bottom:16px;font-size:11px"><li style="padding:3px 0">✓ All premium tools</li><li style="padding:3px 0">✓ BS + GST Recon</li></ul>
      <a href="#gst-contact" style="display:block;text-align:center;padding:8px;border-radius:7px;font-size:12px;font-weight:700;background:var(--bg);color:var(--ink);text-decoration:none;border:1px solid var(--border)">Contact to Buy</a>
    </div>
    <div style="border:1.5px solid var(--brand);border-radius:var(--radius);padding:20px 16px;position:relative">
      <div style="position:absolute;top:-10px;left:50%;transform:translateX(-50%);background:var(--brand);color:#fff;font-size:10px;font-weight:700;padding:2px 10px;border-radius:99px;white-space:nowrap">Most Popular</div>
      <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">Standard</div>
      <div style="font-size:24px;font-weight:800;margin-bottom:2px">₹130</div>
      <div style="font-size:12px;font-weight:700;color:var(--brand);margin-bottom:2px">25 uploads</div>
      <div style="font-size:10px;color:var(--muted);margin-bottom:14px">3 month validity</div>
      <ul style="list-style:none;margin-bottom:16px;font-size:11px"><li style="padding:3px 0">✓ All premium tools</li><li style="padding:3px 0">✓ Priority support</li></ul>
      <a href="#gst-contact" style="display:block;text-align:center;padding:8px;border-radius:7px;font-size:12px;font-weight:700;background:var(--brand);color:#fff;text-decoration:none">Contact to Buy</a>
    </div>
    <div style="border:1.5px solid var(--border);border-radius:var(--radius);padding:20px 16px">
      <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">Professional</div>
      <div style="font-size:24px;font-weight:800;margin-bottom:2px">₹270</div>
      <div style="font-size:12px;font-weight:700;color:var(--brand);margin-bottom:2px">60 uploads</div>
      <div style="font-size:10px;color:var(--muted);margin-bottom:14px">3 month validity</div>
      <ul style="list-style:none;margin-bottom:16px;font-size:11px"><li style="padding:3px 0">✓ All premium tools</li><li style="padding:3px 0">✓ Priority support</li></ul>
      <a href="#gst-contact" style="display:block;text-align:center;padding:8px;border-radius:7px;font-size:12px;font-weight:700;background:var(--bg);color:var(--ink);text-decoration:none;border:1px solid var(--border)">Contact to Buy</a>
    </div>
    <div style="border:1.5px solid var(--border);border-radius:var(--radius);padding:20px 16px">
      <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">Firm</div>
      <div style="font-size:24px;font-weight:800;margin-bottom:2px">₹600</div>
      <div style="font-size:12px;font-weight:700;color:var(--brand);margin-bottom:2px">150 uploads</div>
      <div style="font-size:10px;color:var(--muted);margin-bottom:14px">3 month validity</div>
      <ul style="list-style:none;margin-bottom:16px;font-size:11px"><li style="padding:3px 0">✓ All premium tools</li><li style="padding:3px 0">✓ WhatsApp support</li></ul>
      <a href="#gst-contact" style="display:block;text-align:center;padding:8px;border-radius:7px;font-size:12px;font-weight:700;background:var(--bg);color:var(--ink);text-decoration:none;border:1px solid var(--border)">Contact to Buy</a>
    </div>
    <div style="border:1.5px solid var(--border);border-radius:var(--radius);padding:20px 16px">
      <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">CA Admin</div>
      <div style="font-size:24px;font-weight:800;margin-bottom:2px">₹1,000</div>
      <div style="font-size:12px;font-weight:700;color:var(--brand);margin-bottom:2px">500 uploads</div>
      <div style="font-size:10px;color:var(--muted);margin-bottom:14px">3 month validity</div>
      <ul style="list-style:none;margin-bottom:16px;font-size:11px"><li style="padding:3px 0">✓ All premium tools</li><li style="padding:3px 0">✓ Best for CA firms</li></ul>
      <a href="#gst-contact" style="display:block;text-align:center;padding:8px;border-radius:7px;font-size:12px;font-weight:700;background:var(--bg);color:var(--ink);text-decoration:none;border:1px solid var(--border)">Contact to Buy</a>
    </div>
  </div>
  <p style="text-align:center;font-size:11px;color:var(--muted);margin-top:16px">⚠ No refund after first upload is used · Unused uploads stack when you recharge</p>
</section>

<section style="max-width:700px;margin:0 auto;padding:48px 24px" id="gst-contact">
  <h2 style="text-align:center;font-size:20px;font-weight:800;margin-bottom:16px">Purchase a Plan</h2>
  <p style="text-align:center;color:var(--muted);font-size:13px;margin-bottom:24px">Pay via UPI and send your payment screenshot to our email. Account upgraded within a few hours.</p>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;max-width:400px;margin:0 auto">
    <div style="text-align:center">
      <strong style="font-size:11px;color:var(--muted);text-transform:uppercase">Email</strong><br>
      <a href="mailto:{{ contact_email }}" style="font-size:13px;color:var(--brand)">{{ contact_email }}</a>
    </div>
    <div style="text-align:center">
      <strong style="font-size:11px;color:var(--muted);text-transform:uppercase">UPI Payment</strong><br>
      <span style="font-size:13px;font-weight:600">{{ contact_upi }}</span>
    </div>
  </div>
</section>

<div id="toast" style="position:fixed;bottom:24px;right:24px;background:#065F46;color:#fff;padding:12px 20px;border-radius:10px;font-size:13px;font-weight:600;opacity:0;transition:opacity .3s;pointer-events:none;z-index:999"></div>

<script>
function pickFile(inp, sfId){
  const sf=document.getElementById(sfId);
  if(inp.files.length){sf.textContent='✓ '+inp.files[0].name;sf.style.display='block';}
  if(sfId==='sf-gst' && inp.files[0]){detectStateCodes(inp.files[0]);}
}

// Drag-and-drop for GST upload zones
document.querySelectorAll('.dropzone').forEach(dz => {
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('drag');
    const f = e.dataTransfer.files[0]; if(!f) return;
    const inp = dz.querySelector('input[type=file]');
    const dt = new DataTransfer(); dt.items.add(f); inp.files = dt.files;
    inp.dispatchEvent(new Event('change'));
  });
});

async function detectStateCodes(file){
  // Read ZIP to find state code folders
  const ab=await file.arrayBuffer();
  const view=new Uint8Array(ab);
  // Simple ZIP parsing: find folder names like "XX/" or "gst/XX/"
  const text=new TextDecoder('utf-8',{fatal:false}).decode(view);
  const codes=new Set();
  // Match folder patterns in ZIP central directory
  const re=/(?:^|\/)(0[1-9]|[1-3][0-9])\//gm;
  let m;while((m=re.exec(text))!==null)codes.add(m[1]);
  if(codes.size>0){
    document.getElementById('mapping-container').innerHTML='';
    for(const c of [...codes].sort())addMapping(c,'');
  }
}

function onConsolidatedChange(){
  const chk = document.getElementById('consolidated-chk').checked;
  document.getElementById('consolidated-hint').style.display      = chk ? 'block' : 'none';
  document.getElementById('split-hint').style.display             = chk ? 'none'  : 'block';
  document.getElementById('mapping-field').style.display          = chk ? 'none'  : 'block';
  document.getElementById('consolidated-col-field').style.display = chk ? 'block' : 'none';
}

function addMapping(code,col){
  const container=document.getElementById('mapping-container');
  const div=document.createElement('div');div.className='mapping-row';
  div.innerHTML=`<input type="text" class="map-code" value="${code||''}" placeholder="03">
    <input type="text" class="map-col" value="${col||''}" placeholder="Column header (e.g. DRH/LDH)">
    <button class="remove-btn" onclick="this.parentElement.remove()">✕</button>`;
  container.appendChild(div);
}

function showStatus(t,m){const e=document.getElementById('status');e.className=t;e.innerHTML=m;e.style.display=m?'block':'none';}
function toast(msg){const t=document.getElementById('toast');t.textContent=msg;t.style.opacity='1';setTimeout(()=>t.style.opacity='0',3000);}

async function doProcess(){
  const btn=document.getElementById('proc-btn'),sp=document.getElementById('sp'),bt=document.getElementById('bt');
  const dl=document.getElementById('dl-link');
  const salesFile=document.getElementById('file-sales').files[0];
  const gstFile=document.getElementById('file-gst').files[0];
  if(!salesFile){showStatus('error','✗ Please upload your Sales Summary Excel file.');return;}
  if(!gstFile){showStatus('error','✗ Please upload your GSTR 3B ZIP file.');return;}

  // Check consolidated mode
  const chkEl=document.getElementById('consolidated-chk');
  const isConsolidated = chkEl && chkEl.checked;

  // Collect state-column mappings (location-wise mode only)
  const mappings={};
  if(!isConsolidated){
    const rows=document.querySelectorAll('.mapping-row');
    for(const r of rows){
      const code=r.querySelector('.map-code').value.trim();
      const col=r.querySelector('.map-col').value.trim();
      if(code&&col)mappings[code]=col;
    }
    if(Object.keys(mappings).length===0){
      showStatus('error','✗ Please add at least one State Code → Column mapping, or tick Consolidated Sales.');
      return;
    }
  }

  // Build FormData FIRST, then append all fields
  const fd=new FormData();
  fd.append('sales_file',salesFile);
  fd.append('gst_file',gstFile);
  fd.append('mappings',JSON.stringify(mappings));
  fd.append('output_name',document.getElementById('output-name').value.trim());

  // Consolidated params
  if(isConsolidated){
    const colInputEl=document.getElementById('consolidated-col-input');
    const colName=(colInputEl?colInputEl.value:'').trim();
    fd.append('consolidated_mode','true');
    fd.append('consolidated_col',colName);
  }

  btn.disabled=true;sp.style.display='inline-block';bt.textContent='Processing…';
  showStatus('info','⏳ Processing — this may take 30–60 seconds for large ZIP files…');
  dl.style.display='none';

  // Live counter so user knows it's working
  let _secs=0;
  const _timer=setInterval(()=>{
    _secs++;
    bt.textContent=`Processing… ${_secs}s`;
  },1000);

  // 3-minute timeout (large ZIPs with many states can take time)
  const _ctrl=new AbortController();
  const _tout=setTimeout(()=>_ctrl.abort(),180000);

  try{
    const res=await fetch('/gst-process',{method:'POST',body:fd,
      credentials:'include',signal:_ctrl.signal});
    clearTimeout(_tout);
    const ct=res.headers.get('content-type')||'';
    if(!ct.includes('application/json')){
      showStatus('error','✗ Server error (not JSON). Please try again.');return;
    }
    const data=await res.json();
    if(data.status==='success'){
      const logHtml='<ul class="log-list">'+data.log.map(l=>`<li>${l}</li>`).join('')+'</ul>';
      showStatus('success','✓ Reconciliation complete! ('+_secs+'s)'+logHtml);
      dl.href='/download/'+data.file_id+'?fn='+encodeURIComponent(data.filename);dl.download=data.filename;
      dl.textContent='⬇  Download — '+data.filename;dl.style.display='block';
      toast('Reconciliation done!');
    }else{showStatus('error','✗ '+data.message);}
  }catch(e){
    clearTimeout(_tout);
    if(e.name==='AbortError'){
      showStatus('error','✗ Timed out after 3 minutes. Try with a smaller ZIP or contact support.');
    }else{
      showStatus('error','✗ Network error: '+e.message);
    }
  }
  finally{
    clearInterval(_timer);
    btn.disabled=false;sp.style.display='none';bt.textContent='⚡ Process & Download';
  }
}
</script>
<a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>

<button class="help-btn" onclick="openHelp()" title="How to use this tool">?</button>
<div class="help-overlay" id="helpOverlay">
  <div class="help-modal">
    <div class="help-modal-head"><h3>How to Use — GST Reconciliation</h3><button class="help-close" onclick="closeHelp()">&#10005;</button></div>
    <div class="help-modal-body"><div class="help-step"><div class="help-step-num">1</div><div class="help-step-body"><h4>Upload Sales Excel</h4><p>Upload your books sales summary Excel with month-wise, state-wise data.</p></div></div><div class="help-step"><div class="help-step-num">2</div><div class="help-step-body"><h4>Upload GSTR-3B ZIP</h4><p>Zip all your GSTR-3B PDFs (one per month) and upload the ZIP file.</p></div></div><div class="help-step"><div class="help-step-num">3</div><div class="help-step-body"><h4>Review Mappings</h4><p>Map your Excel column headers to the required fields on screen.</p></div></div><div class="help-step"><div class="help-step-num">4</div><div class="help-step-body"><h4>Process</h4><p>Click Process to generate the reconciliation report.</p></div></div><div class="help-step"><div class="help-step-num">5</div><div class="help-step-body"><h4>Download Report</h4><p>Download the Excel report with month-wise and state-wise differences highlighted.</p></div></div><div class="help-tip">📌 Export GSTR-3B PDFs from the GST portal and ZIP them before uploading.</div></div>
  </div>
</div>
<script>function openHelp(){document.getElementById('helpOverlay').classList.add('open')}function closeHelp(){document.getElementById('helpOverlay').classList.remove('open')}document.getElementById('helpOverlay').addEventListener('click',function(e){if(e.target===this)closeHelp()})</script>
</body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
#  GST RECONCILIATION — PROCESSING LOGIC
# ══════════════════════════════════════════════════════════════════════════════

def _parse_gstr3b_pdf(pdf_path):
    """Extract Table 3.1 data (rows A,B,C,E — NOT D) from a GSTR 3B PDF.

    Uses pdfplumber cropped to just the top ~45 % of page 1 where Table 3.1
    lives. Skips the rest of the page (tables 3.1.1, 3.2, 4, 5 …) for speed.
    ~0.2 s per PDF, no external CLI dependencies.
    """
    import pdfplumber

    result = {'taxable': 0, 'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0}
    period = year = gstin = trade_name = state_code = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text() or ""

            # ── metadata ────────────────────────────────────────────────
            period_m = re.search(r'Period\s+(\w+)', text)
            year_m   = re.search(r'Year\s+([\d-]+)', text)
            gstin_m  = re.search(r'GSTIN of the supplier\s+(\S+)', text)
            trade_m  = re.search(r'Trade name, if any\s+(.+?)(?:\n|$)', text)

            period     = period_m.group(1).strip() if period_m else None
            year       = year_m.group(1).strip()   if year_m   else None
            gstin      = gstin_m.group(1).strip()  if gstin_m  else None
            trade_name = trade_m.group(1).strip()  if trade_m  else None
            state_code = gstin[:2] if gstin else None

            # ── table 3.1 (cropped to top 45 % of page) ────────────────
            cropped = page.crop((0, 0, page.width, page.height * 0.45))
            tables = cropped.extract_tables()

            for t in tables:
                if not t or not t[0] or not t[0][0]:
                    continue
                if 'Nature of Supplies' not in str(t[0][0]):
                    continue
                if not any('Outward taxable supplies' in str(r[0] or '')
                           for r in t[1:]):
                    continue

                def cn(s):
                    if s is None: return 0.0
                    s = str(s).replace('\n', '').strip()
                    s = re.sub(r'^[A-Z]\s*', '', s)
                    if s in ('-', '', '0'): return 0.0
                    try: return float(s.replace(',', ''))
                    except: return 0.0

                for row in t[1:]:
                    lbl = str(row[0] or '').lower()
                    if '(d)' in lbl:          # skip reverse charge
                        continue
                    if any(f'({x})' in lbl for x in 'abce'):
                        result['taxable'] += cn(row[1])
                        result['igst']    += cn(row[2])
                        result['cgst']    += cn(row[3])
                        result['sgst']    += cn(row[4])
                        result['cess']    += cn(row[5])
                break                         # only need the first matching table
    except Exception:
        pass

    return {
        'period': period, 'year': year, 'gstin': gstin,
        'trade_name': trade_name, 'state_code': state_code,
        'taxable_value': result['taxable'],
        'igst': result['igst'], 'cgst': result['cgst'],
        'sgst': result['sgst'], 'cess': result['cess'],
        'total_tax': (result['igst'] + result['cgst']
                      + result['sgst'] + result['cess']),
    }


def _month_key(period_name, fy_str):
    """Convert GSTR3B period 'December' + year '2025-26' to 'Dec-25' format.
    Handles both full names ('January') and abbreviations ('Jan')."""
    month_abbr = {
        'january': 'Jan', 'february': 'Feb', 'march': 'Mar', 'april': 'Apr',
        'may': 'May', 'june': 'Jun', 'july': 'Jul', 'august': 'Aug',
        'september': 'Sep', 'october': 'Oct', 'november': 'Nov', 'december': 'Dec',
        # Also accept 3-letter abbreviations directly
        'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr',
        'jun': 'Jun', 'jul': 'Jul', 'aug': 'Aug', 'sep': 'Sep',
        'oct': 'Oct', 'nov': 'Nov', 'dec': 'Dec',
    }
    abbr = month_abbr.get(period_name.lower().strip())
    if not abbr:
        return None
    # FY "2025-26" means Apr 2025 - Mar 2026
    # Apr-Mar months: Apr,May,...,Dec use first year; Jan,Feb,Mar use second
    try:
        fy_start = int(fy_str.split('-')[0])
    except:
        return None
    if abbr in ('Jan', 'Feb', 'Mar'):
        yr = fy_start + 1
    else:
        yr = fy_start
    return f"{abbr}-{str(yr)[2:]}"


def _process_gst_reconciliation(sales_path, gst_zip_path, mappings, output_path):
    """Process GST reconciliation and generate output Excel."""
    from openpyxl import load_workbook as lb
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    import tempfile, shutil

    log = []

    # --- 1. Read Sales Summary ---
    wb_sales = lb(sales_path, read_only=True, data_only=True)
    ws = wb_sales[wb_sales.sheetnames[0]]
    rows = list(ws.iter_rows(min_row=1, max_row=500, values_only=False))
    wb_sales.close()

    # Find header row (first row with "Month" in col A or similar)
    header_idx = None
    for i, row in enumerate(rows):
        a_val = str(row[0].value or '').strip().lower()
        if a_val in ('month', 'months', 'period'):
            header_idx = i
            break
    if header_idx is None:
        # Try: any row with 1+ non-empty cells where at least one is a text header (non-numeric)
        for i, row in enumerate(rows):
            non_empty = [c for c in row if c.value is not None]
            has_text = any(
                isinstance(c.value, str) and c.value.strip()
                and c.value.strip().lower() not in ('total', 'grand total', 'subtotal')
                for c in non_empty
            )
            has_numeric = any(isinstance(c.value, (int, float)) for c in non_empty)
            if len(non_empty) >= 1 and has_text and not has_numeric:
                header_idx = i
                break
    if header_idx is None:
        # Final fallback: first row with any non-empty cell
        for i, row in enumerate(rows):
            if any(c.value is not None for c in row):
                header_idx = i
                break
    if header_idx is None:
        return {'status': 'error', 'message': 'Could not find header row in sales file.'}

    header_row = rows[header_idx]
    col_headers = {str(c.value).strip(): c.column - 1 for c in header_row if c.value}
    log.append(f"Sales columns found: {', '.join(col_headers.keys())}")

    # Read month-wise sales data
    sales_data = {}
    for row in rows[header_idx + 1:]:
        month_val = row[0].value
        if month_val is None:
            continue
        ms = str(month_val).strip()
        if 'total' in ms.lower() or 'grand' in ms.lower():
            continue
        sales_data[ms] = {}
        for hdr, idx in col_headers.items():
            if hdr.lower() in ('month', 'months', 'period'):
                continue
            try:
                val = row[idx].value
                sales_data[ms][hdr] = float(val) if val is not None else 0.0
            except:
                sales_data[ms][hdr] = 0.0

    log.append(f"Sales months found: {', '.join(sales_data.keys())}")

    # --- 2. Extract GSTR 3B data from ZIP ---
    gst_data = {}  # {state_code: {month_key: {taxable, igst, cgst, sgst, cess, total_tax}}}
    trade_names = {}

    tmpdir = tempfile.mkdtemp()
    try:
        import zipfile as zf
        with zf.ZipFile(gst_zip_path, 'r') as z:
            z.extractall(tmpdir)

        # Collect all PDF paths first
        pdf_tasks = []
        for root_dir, dirs, files in os.walk(tmpdir):
            for fname in files:
                if not fname.lower().endswith('.pdf'):
                    continue
                if 'certificate' in root_dir.lower():
                    continue
                pdf_tasks.append((os.path.join(root_dir, fname), fname))

        log.append(f"Found {len(pdf_tasks)} GSTR 3B PDFs — parsing in parallel...")

        # Parse all PDFs in parallel (4 workers — keeps memory low on free tier)
        from concurrent.futures import ThreadPoolExecutor, as_completed
        def _parse_task(args):
            pdf_path, fname = args
            return fname, _parse_gstr3b_pdf(pdf_path)

        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {pool.submit(_parse_task, t): t for t in pdf_tasks}
            for future in as_completed(futures):
                try:
                    fname, result = future.result(timeout=30)
                except Exception as _e:
                    log.append(f"  ⚠ PDF parse error: {_e}")
                    continue

                if result and result['state_code'] and result['period']:
                    sc = result['state_code']

                    # ── CRITICAL FIX: derive month from FILENAME not PDF period ──
                    mk = None
                    fn_base = os.path.splitext(fname)[0]
                    fn_parts = fn_base.split('_')
                    fn_mmyyyy = fn_parts[-1] if fn_parts else ''
                    if len(fn_mmyyyy) == 6 and fn_mmyyyy.isdigit():
                        mm   = int(fn_mmyyyy[:2])
                        yyyy = int(fn_mmyyyy[2:])
                        _MNAMES = {1:'January',2:'February',3:'March',4:'April',
                                   5:'May',6:'June',7:'July',8:'August',
                                   9:'September',10:'October',11:'November',12:'December'}
                        fn_period = _MNAMES.get(mm, result['period'])
                        fn_fy_start = yyyy - 1 if mm <= 3 else yyyy
                        fn_fy = f"{fn_fy_start}-{str(fn_fy_start+1)[2:]}"
                        mk = _month_key(fn_period, fn_fy)
                        if mk and mk != _month_key(result['period'], result['year'] or ''):
                            log.append(f"  ℹ Quarterly filer: PDF says '{result['period']}' "
                                       f"but filename says {fn_period} → using '{mk}'")

                    if not mk:
                        mk = _month_key(result['period'], result['year'] or '')

                    if mk:
                        if sc not in gst_data:
                            gst_data[sc] = {}
                        if mk in gst_data[sc]:
                            existing = gst_data[sc][mk]
                            existing['taxable_value'] += result['taxable_value']
                            existing['igst']          += result['igst']
                            existing['cgst']          += result['cgst']
                            existing['sgst']          += result['sgst']
                            existing['cess']          += result['cess']
                            existing['total_tax']     += result['total_tax']
                        else:
                            gst_data[sc][mk] = result
                        if result.get('trade_name'):
                            trade_names[sc] = result['trade_name']
                        log.append(f"Parsed: State {sc} / {mk} — taxable ₹{result['taxable_value']:,.2f}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    if not gst_data:
        return {'status': 'error', 'message': 'No valid GSTR 3B PDFs found in the ZIP.'}

    # ── Normalize month keys so sales & GSTR3B match ─────────────────
    # Sales might have "April","May" or "Dec-25","Jan-26" etc.
    # GSTR3B uses "Dec-25","Mar-26" from _month_key().
    import re as _re
    _MF = {'january':'jan','february':'feb','march':'mar','april':'apr',
           'may':'may','june':'jun','july':'jul','august':'aug',
           'september':'sep','october':'oct','november':'nov','december':'dec'}

    def _norm_month(raw):
        s = str(raw).strip().lower().rstrip('.')
        # "Dec-25" / "Jan-26"
        m = _re.match(r'^([a-z]+)-?(\d{2,4})$', s)
        if m:
            mon = _MF.get(m.group(1), m.group(1)[:3])
            return f"{mon}-{m.group(2)[-2:]}"
        # "April" / "December" (full, no year)
        if s in _MF:
            return _MF[s]
        if s[:3] in _MF.values():
            return s[:3]
        return s

    # Normalize sales keys
    sales_data = {_norm_month(k): v for k, v in sales_data.items()}
    # Normalize GSTR3B month keys
    for sc in list(gst_data.keys()):
        gst_data[sc] = {_norm_month(mk): d for mk, d in gst_data[sc].items()}

    log.append(f"Normalized: Sales={sorted(sales_data.keys())}")
    if gst_data:
        sample_sc = next(iter(gst_data))
        log.append(f"  GSTR3B[{sample_sc}]={sorted(gst_data[sample_sc].keys())}")

    # ── KEY FIX: Align month keys when sales has no year suffix ──────────────
    # Sales: {"apr": {...}, "may": {...}, ...}   (bare 3-letter, no year)
    # GSTR3B: {"apr-25": {...}, "may-25": {...}, ...}  (with 2-digit year)
    # If ALL sales keys are bare (no "-") but GSTR3B keys all have "-",
    # remap sales keys by prepending the matching year from GSTR3B.
    sales_bare   = all('-' not in k for k in sales_data.keys())
    gstr_with_yr = any('-' in mk for sc_d in gst_data.values() for mk in sc_d.keys())

    if sales_bare and gstr_with_yr:
        # Build a bare→with-year map from GSTR3B keys
        # e.g. "apr" → "apr-25", "jan" → "jan-26"
        bare_to_full = {}
        for sc_d in gst_data.values():
            for mk in sc_d.keys():
                if '-' in mk:
                    bare = mk.split('-')[0]  # "apr-25" → "apr"
                    bare_to_full[bare] = mk   # last one wins (same across states)
        if bare_to_full:
            remapped = {}
            for bare_k, v in sales_data.items():
                full_k = bare_to_full.get(bare_k, bare_k)  # fallback: keep original
                remapped[full_k] = v
            sales_data = remapped
            log.append(f"Month key alignment: Sales remapped to year-suffixed keys: {sorted(sales_data.keys())}")

    # Also handle reverse: GSTR3B bare, Sales with year
    gstr_bare   = all('-' not in mk for sc_d in gst_data.values() for mk in sc_d.keys())
    sales_with_yr = any('-' in k for k in sales_data.keys())
    if gstr_bare and sales_with_yr:
        sales_bare_to_full = {}
        for k in sales_data.keys():
            if '-' in k:
                bare = k.split('-')[0]
                sales_bare_to_full[bare] = k
        for sc in list(gst_data.keys()):
            remapped = {}
            for mk, d in gst_data[sc].items():
                full_k = sales_bare_to_full.get(mk, mk)
                remapped[full_k] = d
            gst_data[sc] = remapped
        log.append(f"Month key alignment: GSTR3B remapped to year-suffixed keys")

    # --- 3. Build output Excel ---
    wb = Workbook()
    ws_out = wb.active
    ws_out.title = "Reconciliation"

    # Styles
    hdr_font = Font(name='Arial', bold=True, size=11)
    hdr_fill = PatternFill('solid', fgColor='1F4E79')
    hdr_font_w = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    sub_fill = PatternFill('solid', fgColor='D6E4F0')
    num_fmt = '#,##0.00'
    thin = Side(style='thin', color='B4B4B4')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    diff_neg_font = Font(name='Arial', color='CC0000', bold=True, size=10)
    diff_pos_font = Font(name='Arial', color='006600', bold=True, size=10)
    data_font = Font(name='Arial', size=10)

    # Title
    entity = trade_names.get(list(gst_data.keys())[0], 'Entity')
    ws_out['A1'] = f"{entity} — GST Reconciliation (Sales Books vs GSTR 3B)"
    ws_out['A1'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
    ws_out.merge_cells('A1:I1')
    ws_out['A2'] = "Table 3.1: Points A + B + C + E (Excludes D — Reverse Charge)"
    ws_out['A2'].font = Font(name='Arial', italic=True, size=10, color='666666')
    ws_out.merge_cells('A2:I2')

    # Headers (row 4)
    headers = ['Month', 'GST State Code', 'Sale in Books',
               'Sale in GSTR 3B\n(Excl. Tax)', 'Tax Amount\n(IGST+CGST+SGST+Cess)',
               'Difference 1\n(Books − GSTR3B)',
               'Difference 2\n(Books − GSTR3B incl. Tax)',
               'IGST', 'CGST', 'SGST']
    for c, h in enumerate(headers, 1):
        cell = ws_out.cell(row=4, column=c, value=h)
        cell.font = hdr_font_w
        cell.fill = hdr_fill
        cell.alignment = center
        cell.border = border

    # Data rows
    row_num = 5
    # Union of sales months AND GSTR3B months — so months with only GSTR3B data show too
    gstr_months = {mk for sc_d in gst_data.values() for mk in sc_d.keys()}
    all_months_set = set(sales_data.keys()) | gstr_months
    all_months = sorted(all_months_set, key=lambda x: _month_sort_key(x))
    all_state_codes = sorted(gst_data.keys())

    # ── Detect consolidated vs split-by-location mode ───────────────────────
    # Handle explicit consolidated flag from frontend checkbox
    forced_consolidated = "__consolidated__" in mappings
    consolidated_hint   = mappings.pop("__consolidated__", "") if forced_consolidated else ""

    non_empty_cols = [v.strip() for v in mappings.values() if v and v.strip()]
    is_consolidated = forced_consolidated or (not non_empty_cols or
                       len(set(c.lower() for c in non_empty_cols)) == 1)

    # Resolve consolidated column name
    consolidated_col = None
    if is_consolidated:
        # Use explicit column name from checkbox field if provided
        if forced_consolidated and consolidated_hint and consolidated_hint != "__auto__":
            matched = next((h for h in col_headers if h.lower()==consolidated_hint.lower()), consolidated_hint)
            non_empty_cols = [matched]
        if non_empty_cols:
            # Case-insensitive match against actual headers
            raw = non_empty_cols[0]
            consolidated_col = next((h for h in col_headers if h.lower()==raw.lower()), raw)
        else:
            # Auto-detect: first non-month column that has any data
            for hdr in col_headers:
                if hdr.lower() in ('month','months','period'):
                    continue
                if any(mdata.get(hdr,0) for mdata in sales_data.values()):
                    consolidated_col = hdr
                    break
        if consolidated_col:
            log.append(f"✅ Consolidated mode — books column: '{consolidated_col}'")
        else:
            return {'status':'error','message':'Could not find sales data column. Please enter the column header name in at least one mapping row.'}

    # Build valid_mappings
    valid_mappings = {}
    if is_consolidated:
        for sc in all_state_codes:
            valid_mappings[sc] = consolidated_col
    else:
        for sc, col_name in mappings.items():
            if not col_name or not col_name.strip():
                continue
            matched = next((h for h in col_headers if h.lower()==col_name.lower().strip()), None)
            if matched:
                valid_mappings[sc] = matched
            else:
                log.append(f"⚠ Column '{col_name}' not found for state {sc} — skipped")
        if not valid_mappings:
            return {'status':'error','message':'No valid mappings found. Check column header names match your Excel exactly.'}

    log.append(f"Mode: {'Consolidated' if is_consolidated else 'Split'} · "
               f"States: {', '.join(sorted(valid_mappings.keys()))}")

    month_totals = {}

    for month in all_months:
        first_in_month = True

        if is_consolidated:
            # ── Consolidated: compare total books vs sum of all GSTR3B states ──
            books_total = sales_data.get(month, {}).get(consolidated_col, 0.0)
            gstr_t = tax_t = igst_t = cgst_t = sgst_t = 0.0
            state_rows = []
            for sc in all_state_codes:
                ge = gst_data.get(sc, {}).get(month)
                if not ge:
                    continue
                gstr_t += ge['taxable_value']; tax_t  += ge['total_tax']
                igst_t += ge['igst'];          cgst_t += ge['cgst']; sgst_t += ge['sgst']
                state_rows.append((sc, ge))

            if not state_rows and books_total == 0:
                continue

            # Write individual GSTR3B state rows (no books column per state)
            for sc, ge in state_rows:
                vals = [month if first_in_month else '', sc,
                        '', ge['taxable_value'], ge['total_tax'],
                        '', '', ge['igst'], ge['cgst'], ge['sgst']]
                for c, v in enumerate(vals, 1):
                    cell = ws_out.cell(row=row_num, column=c, value=v)
                    cell.font = data_font; cell.border = border
                    if c >= 3 and v != '':
                        cell.number_format = num_fmt
                        cell.alignment = Alignment(horizontal='right')
                    if c <= 2: cell.alignment = center
                first_in_month = False; row_num += 1

            # Consolidated totals row with difference
            d1 = books_total - gstr_t
            d2 = books_total - (gstr_t + tax_t)
            sub_vals = [f'{month} Total (Books vs All GSTR3B)', '',
                        books_total, gstr_t, tax_t, d1, d2, igst_t, cgst_t, sgst_t]
            for c, v in enumerate(sub_vals, 1):
                cell = ws_out.cell(row=row_num, column=c, value=v)
                cell.font = Font(name='Arial', bold=True, size=10)
                cell.fill = sub_fill; cell.border = border
                if c >= 3:
                    cell.number_format = num_fmt
                    cell.alignment = Alignment(horizontal='right')
                if c in (6, 7) and isinstance(v, (int, float)):
                    col_c = 'CC0000' if v < -0.5 else ('006600' if v > 0.5 else '000000')
                    cell.font = Font(name='Arial', bold=True, size=10, color=col_c)
                    cell.fill = sub_fill
            month_totals[month] = {'books':books_total,'gstr':gstr_t,'tax':tax_t,
                                    'igst':igst_t,'cgst':cgst_t,'sgst':sgst_t}
            row_num += 2
            continue

        # ── Split mode: per-state ─────────────────────────────────────────────
        for sc in all_state_codes:
            if sc not in valid_mappings:
                continue
            col_name  = valid_mappings[sc]
            books_val = sales_data.get(month, {}).get(col_name, 0.0)
            ge        = gst_data.get(sc, {}).get(month)
            if ge is None and books_val == 0:
                continue
            gstr_val = ge['taxable_value'] if ge else 0.0
            tax_val  = ge['total_tax']      if ge else 0.0
            igst     = ge['igst']            if ge else 0.0
            cgst     = ge['cgst']            if ge else 0.0
            sgst     = ge['sgst']            if ge else 0.0
            diff1 = books_val - gstr_val
            diff2 = books_val - (gstr_val + tax_val)
            if month not in month_totals:
                month_totals[month] = {'books':0,'gstr':0,'tax':0,'igst':0,'cgst':0,'sgst':0}
            month_totals[month]['books'] += books_val; month_totals[month]['gstr'] += gstr_val
            month_totals[month]['tax']   += tax_val;   month_totals[month]['igst'] += igst
            month_totals[month]['cgst']  += cgst;      month_totals[month]['sgst'] += sgst
            vals = [month if first_in_month else '', sc, books_val,
                    gstr_val, tax_val, diff1, diff2, igst, cgst, sgst]
            for c, v in enumerate(vals, 1):
                cell = ws_out.cell(row=row_num, column=c, value=v)
                cell.font = data_font; cell.border = border
                if c >= 3:
                    cell.number_format = num_fmt
                    cell.alignment = Alignment(horizontal='right')
                if c in (6, 7) and isinstance(v, (int, float)):
                    cell.font = diff_neg_font if v < -0.5 else (diff_pos_font if v > 0.5 else data_font)
                if c <= 2: cell.alignment = center
            first_in_month = False; row_num += 1

        # Month subtotal
        if month in month_totals:
            mt = month_totals[month]
            sub_vals = [f'{month} Total', '', mt['books'], mt['gstr'], mt['tax'],
                        mt['books'] - mt['gstr'], mt['books'] - mt['gstr'] - mt['tax'],
                        mt['igst'], mt['cgst'], mt['sgst']]
            for c, v in enumerate(sub_vals, 1):
                cell = ws_out.cell(row=row_num, column=c, value=v)
                cell.font = Font(name='Arial', bold=True, size=10)
                cell.fill = sub_fill
                cell.border = border
                if c >= 3:
                    cell.number_format = num_fmt
                    cell.alignment = Alignment(horizontal='right')
                if c in (6, 7) and isinstance(v, (int, float)):
                    cell.font = Font(name='Arial', bold=True, size=10, color='CC0000' if v < -0.5 else ('006600' if v > 0.5 else '000000'))
                    cell.fill = sub_fill
            row_num += 1
        row_num += 1  # blank row between months

    # Grand Total
    grand = {'books': 0, 'gstr': 0, 'tax': 0, 'igst': 0, 'cgst': 0, 'sgst': 0}
    for mt in month_totals.values():
        for k in grand:
            grand[k] += mt[k]

    gt_vals = ['GRAND TOTAL', '', grand['books'], grand['gstr'], grand['tax'],
               grand['books'] - grand['gstr'], grand['books'] - grand['gstr'] - grand['tax'],
               grand['igst'], grand['cgst'], grand['sgst']]
    for c, v in enumerate(gt_vals, 1):
        cell = ws_out.cell(row=row_num, column=c, value=v)
        cell.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E79')
        cell.border = border
        if c >= 3:
            cell.number_format = num_fmt
            cell.alignment = Alignment(horizontal='right')

    # Column widths
    widths = [12, 16, 18, 22, 22, 22, 24, 16, 14, 14]
    for i, w in enumerate(widths, 1):
        ws_out.column_dimensions[get_column_letter(i)].width = w

    # Freeze panes
    ws_out.freeze_panes = 'A5'

    wb.save(output_path)
    log.append(f"Output: {len(month_totals)} months × {len(valid_mappings)} states reconciled")
    return {'status': 'success', 'log': log}


def _month_sort_key(ms):
    """Sort key for month strings like 'Apr-25', 'Dec-25', 'Jan-26', or bare 'apr','may'."""
    month_order = {'apr':1,'may':2,'jun':3,'jul':4,'aug':5,'sep':6,
                   'oct':7,'nov':8,'dec':9,'jan':10,'feb':11,'mar':12}
    parts = ms.lower().split('-')
    if len(parts) == 2:
        m = month_order.get(parts[0][:3], 0)
        try: y = int(parts[1])
        except: y = 0
        return (y, m)
    # Bare month name (no year)
    m = month_order.get(ms.lower()[:3], 0)
    if m: return (0, m)
    return (99, 99)


# ── GST routes ────────────────────────────────────────────────────────────────

# ── GST Sales Template Data (base64-encoded) ──────────────────────────────────
_GST_CONS_B64   = "UEsDBBQAAAAIACZBxVxGx01IlQAAAM0AAAAQAAAAZG9jUHJvcHMvYXBwLnhtbE3PTQvCMAwG4L9SdreZih6kDkQ9ip68zy51hbYpbYT67+0EP255ecgboi6JIia2mEXxLuRtMzLHDUDWI/o+y8qhiqHke64x3YGMsRoPpB8eA8OibdeAhTEMOMzit7Dp1C5GZ3XPlkJ3sjpRJsPiWDQ6sScfq9wcChDneiU+ixNLOZcrBf+LU8sVU57mym/8ZAW/B7oXUEsDBBQAAAAIACZBxVx4wQ9+7gAAACsCAAARAAAAZG9jUHJvcHMvY29yZS54bWzNksFKxDAQhl9Fcm8n7eKioZuL4klBcEHxFpLZ3WDThGSk3bc3jbtdRB9AyCUzf775BtLpILSP+Bx9wEgW09Xk+iEJHTbsQBQEQNIHdCrVOTHk5s5Hpyhf4x6C0h9qj9ByvgaHpIwiBTOwCguRyc5ooSMq8vGEN3rBh8/YF5jRgD06HChBUzfA5DwxHKe+gwtghhFGl74LaBZiqf6JLR1gp+SU7JIax7EeVyWXd2jg7enxpaxb2SGRGjTmV8kKOgbcsPPk19Xd/faByZa364rnc73lN4LfiqZ9n11/+F2EnTd2Z/+x8VlQdvDrX8gvUEsDBBQAAAAIACZBxVyZXJwjEAYAAJwnAAATAAAAeGwvdGhlbWUvdGhlbWUxLnhtbO1aW3PaOBR+76/QeGf2bQvGNoG2tBNzaXbbtJmE7U4fhRFYjWx5ZJGEf79HNhDLlg3tkk26mzwELOn7zkVH5+g4efPuLmLohoiU8nhg2S/b1ru3L97gVzIkEUEwGaev8MAKpUxetVppAMM4fckTEsPcgosIS3gUy9Zc4FsaLyPW6rTb3VaEaWyhGEdkYH1eLGhA0FRRWm9fILTlHzP4FctUjWWjARNXQSa5iLTy+WzF/NrePmXP6TodMoFuMBtYIH/Ob6fkTlqI4VTCxMBqZz9Wa8fR0kiAgsl9lAW6Sfaj0xUIMg07Op1YznZ89sTtn4zK2nQ0bRrg4/F4OLbL0otwHATgUbuewp30bL+kQQm0o2nQZNj22q6RpqqNU0/T933f65tonAqNW0/Ta3fd046Jxq3QeA2+8U+Hw66JxqvQdOtpJif9rmuk6RZoQkbj63oSFbXlQNMgAFhwdtbM0gOWXin6dZQa2R273UFc8FjuOYkR/sbFBNZp0hmWNEZynZAFDgA3xNFMUHyvQbaK4MKS0lyQ1s8ptVAaCJrIgfVHgiHF3K/99Ze7yaQzep19Os5rlH9pqwGn7bubz5P8c+jkn6eT101CznC8LAnx+yNbYYcnbjsTcjocZ0J8z/b2kaUlMs/v+QrrTjxnH1aWsF3Pz+SejHIju932WH32T0duI9epwLMi15RGJEWfyC265BE4tUkNMhM/CJ2GmGpQHAKkCTGWoYb4tMasEeATfbe+CMjfjYj3q2+aPVehWEnahPgQRhrinHPmc9Fs+welRtH2Vbzco5dYFQGXGN80qjUsxdZ4lcDxrZw8HRMSzZQLBkGGlyQmEqk5fk1IE/4rpdr+nNNA8JQvJPpKkY9psyOndCbN6DMawUavG3WHaNI8ev4F+Zw1ChyRGx0CZxuzRiGEabvwHq8kjpqtwhErQj5iGTYacrUWgbZxqYRgWhLG0XhO0rQR/FmsNZM+YMjszZF1ztaRDhGSXjdCPmLOi5ARvx6GOEqa7aJxWAT9nl7DScHogstm/bh+htUzbCyO90fUF0rkDyanP+kyNAejmlkJvYRWap+qhzQ+qB4yCgXxuR4+5Xp4CjeWxrxQroJ7Af/R2jfCq/iCwDl/Ln3Ppe+59D2h0rc3I31nwdOLW95GblvE+64x2tc0LihjV3LNyMdUr5Mp2DmfwOz9aD6e8e362SSEr5pZLSMWkEuBs0EkuPyLyvAqxAnoZFslCctU02U3ihKeQhtu6VP1SpXX5a+5KLg8W+Tpr6F0PizP+Txf57TNCzNDt3JL6raUvrUmOEr0scxwTh7LDDtnPJIdtnegHTX79l125COlMFOXQ7gaQr4Dbbqd3Do4npiRuQrTUpBvw/npxXga4jnZBLl9mFdt59jR0fvnwVGwo+88lh3HiPKiIe6hhpjPw0OHeXtfmGeVxlA0FG1srCQsRrdguNfxLBTgZGAtoAeDr1EC8lJVYDFbxgMrkKJ8TIxF6HDnl1xf49GS49umZbVuryl3GW0iUjnCaZgTZ6vK3mWxwVUdz1Vb8rC+aj20FU7P/lmtyJ8MEU4WCxJIY5QXpkqi8xlTvucrScRVOL9FM7YSlxi84+bHcU5TuBJ2tg8CMrm7Oal6ZTFnpvLfLQwJLFuIWRLiTV3t1eebnK56Inb6l3fBYPL9cMlHD+U751/0XUOufvbd4/pukztITJx5xREBdEUCI5UcBhYXMuRQ7pKQBhMBzZTJRPACgmSmHICY+gu98gy5KRXOrT45f0Usg4ZOXtIlEhSKsAwFIRdy4+/vk2p3jNf6LIFthFQyZNUXykOJwT0zckPYVCXzrtomC4Xb4lTNuxq+JmBLw3punS0n/9te1D20Fz1G86OZ4B6zh3OberjCRaz/WNYe+TLfOXDbOt4DXuYTLEOkfsF9ioqAEativrqvT/klnDu0e/GBIJv81tuk9t3gDHzUq1qlZCsRP0sHfB+SBmOMW/Q0X48UYq2msa3G2jEMeYBY8wyhZjjfh0WaGjPVi6w5jQpvQdVA5T/b1A1o9g00HJEFXjGZtjaj5E4KPNz+7w2wwsSO4e2LvwFQSwMEFAAAAAgAJkHFXGTPnIzcAgAARQkAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWyFlmtvmzAUhv+KxaRq+xIuubYlSE2r7iL1oqbdPjvkAFaNzWyTtP9+tqGQbVw+JNjG7/scG44P4ZGLV5kBKPSWUybXTqZUceG6Ms4gx3LCC2D6TsJFjpXuitSVhQC8t6KcuoHnLdwcE+ZEoR17FFHIS0UJg0eBZJnnWLxvgPLj2vGdj4EnkmbKDLhRWOAUtqBeikehe27jsic5MEk4QwKStXPlX2z8hRHYGT8JHOVJG5ml7Dh/NZ3v+7XjmYiAQqyMBdaXA1wDpcZJx/G7NnUaphGetj/cb+3i9WJ2WMI1p7/IXmVrZ+WgPSS4pOqJH79BvaC58Ys5lfYfHau5Om4Ul1LxvBbrCHLCqit+qzfiRBB4PYKgFgQ27gpko7zBCkeh4EckzGztZhp2qVatgyPMPJWtEvou0ToVfd0+oyeIOYsJJdhu1NmnVeAHl+iaM8kp2WMFe7TFFCSqNiJ0lSYbvRvrnyY22KDBBhYb9GDPPvnBKvBml+iWUIqkdU9IWgp9JUyzaZkztJmgG47uH55RnGGWAoqr8Uy/fyDkZCCSaRPJ1EYy7YnkjjOV/e1jZZth2TNXmFa7MhDErAliZt1mPW5XhSC0K4hKNrcyk2D/EeYNYT5IuMPvXf7zMf9F478Y9P9RMugCLMYAywawHAHQzhUsxwCrBrAafgZlqtOtC7EaQ5w3iPNBxBYKBfkORBflfIzie21Oe4Och1jxHkqtHMKcHB3+IOaeH3pXU0uHOO1Z4QeDnBuI+znBKKc9Cfzp8CuGWakrUydmOoppc90fTvZb2IlezmjC+23G+2MpL+LOk80fzXq/TXu/yuDF0DnYCal0Jr2TaPty93kz09V7/iV0kyg86JmHU657UsFyEKmt01If9yVTVVVpRutvgeBiYyvgv+PmG8FWxtam+sDQ25ESJhGFRFt6k6Vev6hqdtVRvLB1cseVrrq2WdUZM0HfTzhXHx0DaL6coj9QSwMEFAAAAAgAJkHFXHNazGUqAwAAFBAAAA0AAAB4bC9zdHlsZXMueG1s3Vhhb9owEP0rUX7AQghkZCJIJQVp0jZVaj/sqyEOWHLiLDEV9NfPZ4ckgI/RtVOlBaHYd37vns9nGzGt5YHTxy2l0tnnvKhjdytl+cXz6vWW5qT+JEpaKE8mqpxI1a02Xl1WlKQ1gHLuDQeD0MsJK9zZtNjly1zWzlrsChm7A9ebTTNRdJbQNQY1lOTUeSY8dhPC2apieizJGT8Y8xAMa8FF5UglhcauD5b6xbh90wOVDU/OClGB0TMRrsRZNRxdiGqzUnoH/iK4Gy9P4gxvo2QYZTj6PJrM+5TRG0Uu9XOZjD9TtuMHr5dgxepXrTgY5+0ij1xjmE1LIiWtiqXqaIw2Xricpv10KNUqbypy8Idj92ZALThLIeQm6Wfpfr64Wy40TQ/6RtKuPqyk+qXSsRJVSqs2IUP3aJpNOc2kgldss4W3FCXkWUgpctVIGdmIguhsHRF9pKP3a+zKrd5vJ2WRzO/HC1MIMLSJcSNCj9VybgSokUfdNyLM4N7EmobK15py/ggkP7M2ab6i2meOOVK+pnCaOFBtx6bKdNM0NKYDgfpshrtHO/krWqdkz0LOd2oGhe7/2glJHyqasb3u77M2Psbud+zDM3ZSlvxwx9mmyKmZ+80BZ1NyxDlbUbEXFQ226VoZqDkO9xkuaohP+eNEBZ2ooC/K/0hRIyRTr1BwXhyjf8re0z5GSu99tL8Pu9ds1N5pcHIWtFYHrqrY/QG/SHhH4ax2jEtWNL0tS1NaXBwJil6SlfrJc8Kvxqc0Izsun1pn7Hbt7zRluzxqRz3AtJpRXfsbnKF+2N7XKhYrUrqnadJ01aF4cp2YBwDnnu6Ov/RgGOOze8CHxcEUYBiDwuL8T/OZoPMxPkzbxOqZoJgJijEomyfRHyyOHROpxz7TKAqCMMQymiRWBQmWtzCEr50N0wYILA5Eel2u8dXGK+R6HWBreq1CsJnilYjNFM81eOx5A0QU2VcbiwMIbBWw2oH49jhQU3ZMEMCqYtqwHYx7ogjzQC3aazQMkeyE8LGvD7ZLgiCK7B7w2RUEAeaB3Yh7MAWgAfMEgb4Hz+4j73hPed3/ALPfUEsDBBQAAAAIACZBxVyXirscwAAAABMCAAALAAAAX3JlbHMvLnJlbHOdkrluwzAMQH/F0J4wB9AhiDNl8RYE+QFWog/YEgWKRZ2/r9qlcZALGXk9PBLcHmlA7TiktoupGP0QUmla1bgBSLYlj2nOkUKu1CweNYfSQETbY0OwWiw+QC4ZZre9ZBanc6RXiFzXnaU92y9PQW+ArzpMcUJpSEszDvDN0n8y9/MMNUXlSiOVWxp40+X+duBJ0aEiWBaaRcnToh2lfx3H9pDT6a9jIrR6W+j5cWhUCo7cYyWMcWK0/jWCyQ/sfgBQSwMEFAAAAAgAJkHFXEs11FQ5AQAAKgIAAA8AAAB4bC93b3JrYm9vay54bWyNUdFuwjAM/JUqH7AWtCENUV6GtiFNGxoT76F1qUUSV44Lg6+f26oa0l72lNzZutxdFmfi457omHx7F2JuapFmnqaxqMHbeEcNBJ1UxN6KQj6ksWGwZawBxLt0mmWz1FsMZrkYtTac3gISKAQpKNkRO4Rz/J13MDlhxD06lEtu+rsDk3gM6PEKZW4yk8Sazq/EeKUg1m0LJudyMxkGO2DB4g+97Ux+2X3sGbH7T6tGcjPLVLBCjtJv9PpWPZ5AlwfUCj2jE+CVFXhhahsMh05GU6Q3MfoexnMocc7/qZGqCgtYUdF6CDL0yOA6gyHW2ESTBOshN1vrICbb1nvLly6XPrQuh4yi5m4a4znqgNflYHP0VkKFAcp3lYvKa0/FhpPu6HWm9w+TR+2jde5JuY/wRrYco47ftPwBUEsDBBQAAAAIACZBxVwkHpuirQAAAPgBAAAaAAAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHO1kT0OgzAMha8S5QA1UKlDBUxdWCsuEAXzIxISxa4Kty+FAZA6dGGyni1/78lOn2gUd26gtvMkRmsGymTL7O8ApFu0ii7O4zBPahes4lmGBrzSvWoQkii6QdgzZJ7umaKcPP5DdHXdaXw4/bI48A8wvF3oqUVkKUoVGuRMwmi2NsFS4stMlqKoMhmKKpZwWiDiySBtaVZ9sE9OtOd5Fzf3Ra7N4wmu3wxweHT+AVBLAwQUAAAACAAmQcVcZZB5khkBAADPAwAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytk01OwzAQha8SZVslLixYoKYbYAtdcAFjTxqr/pNnWtLbM07aSqASFYVNrHjevM+el6zejxGw6J312JQdUXwUAlUHTmIdIniutCE5SfyatiJKtZNbEPfL5YNQwRN4qih7lOvVM7Ryb6l46XkbTfBNmcBiWTyNwsxqShmjNUoS18XB6x+U6kSouXPQYGciLlhQiquEXPkdcOp7O0BKRkOxkYlepWOV6K1AOlrAetriyhlD2xoFOqi945YaYwKpsQMgZ+vRdDFNJp4wjM+72fzBZgrIyk0KETmxBH/HnSPJ3VVkI0hkpq94IbL17PtBTluDvpHN4/0MaTfkgWJY5s/4e8YX/xvO8RHC7r8/sbzWThp/5ovhP15/AVBLAQIUAxQAAAAIACZBxVxGx01IlQAAAM0AAAAQAAAAAAAAAAAAAACAAQAAAABkb2NQcm9wcy9hcHAueG1sUEsBAhQDFAAAAAgAJkHFXHjBD37uAAAAKwIAABEAAAAAAAAAAAAAAIABwwAAAGRvY1Byb3BzL2NvcmUueG1sUEsBAhQDFAAAAAgAJkHFXJlcnCMQBgAAnCcAABMAAAAAAAAAAAAAAIAB4AEAAHhsL3RoZW1lL3RoZW1lMS54bWxQSwECFAMUAAAACAAmQcVcZM+cjNwCAABFCQAAGAAAAAAAAAAAAAAAgIEhCAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sUEsBAhQDFAAAAAgAJkHFXHNazGUqAwAAFBAAAA0AAAAAAAAAAAAAAIABMwsAAHhsL3N0eWxlcy54bWxQSwECFAMUAAAACAAmQcVcl4q7HMAAAAATAgAACwAAAAAAAAAAAAAAgAGIDgAAX3JlbHMvLnJlbHNQSwECFAMUAAAACAAmQcVcSzXUVDkBAAAqAgAADwAAAAAAAAAAAAAAgAFxDwAAeGwvd29ya2Jvb2sueG1sUEsBAhQDFAAAAAgAJkHFXCQem6KtAAAA+AEAABoAAAAAAAAAAAAAAIAB1xAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzUEsBAhQDFAAAAAgAJkHFXGWQeZIZAQAAzwMAABMAAAAAAAAAAAAAAIABvBEAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAkACQA+AgAABhMAAAAA"
_GST_BRANCH_B64 = "UEsDBBQAAAAIACZBxVxGx01IlQAAAM0AAAAQAAAAZG9jUHJvcHMvYXBwLnhtbE3PTQvCMAwG4L9SdreZih6kDkQ9ip68zy51hbYpbYT67+0EP255ecgboi6JIia2mEXxLuRtMzLHDUDWI/o+y8qhiqHke64x3YGMsRoPpB8eA8OibdeAhTEMOMzit7Dp1C5GZ3XPlkJ3sjpRJsPiWDQ6sScfq9wcChDneiU+ixNLOZcrBf+LU8sVU57mym/8ZAW/B7oXUEsDBBQAAAAIACZBxVx4wQ9+7gAAACsCAAARAAAAZG9jUHJvcHMvY29yZS54bWzNksFKxDAQhl9Fcm8n7eKioZuL4klBcEHxFpLZ3WDThGSk3bc3jbtdRB9AyCUzf775BtLpILSP+Bx9wEgW09Xk+iEJHTbsQBQEQNIHdCrVOTHk5s5Hpyhf4x6C0h9qj9ByvgaHpIwiBTOwCguRyc5ooSMq8vGEN3rBh8/YF5jRgD06HChBUzfA5DwxHKe+gwtghhFGl74LaBZiqf6JLR1gp+SU7JIax7EeVyWXd2jg7enxpaxb2SGRGjTmV8kKOgbcsPPk19Xd/faByZa364rnc73lN4LfiqZ9n11/+F2EnTd2Z/+x8VlQdvDrX8gvUEsDBBQAAAAIACZBxVyZXJwjEAYAAJwnAAATAAAAeGwvdGhlbWUvdGhlbWUxLnhtbO1aW3PaOBR+76/QeGf2bQvGNoG2tBNzaXbbtJmE7U4fhRFYjWx5ZJGEf79HNhDLlg3tkk26mzwELOn7zkVH5+g4efPuLmLohoiU8nhg2S/b1ru3L97gVzIkEUEwGaev8MAKpUxetVppAMM4fckTEsPcgosIS3gUy9Zc4FsaLyPW6rTb3VaEaWyhGEdkYH1eLGhA0FRRWm9fILTlHzP4FctUjWWjARNXQSa5iLTy+WzF/NrePmXP6TodMoFuMBtYIH/Ob6fkTlqI4VTCxMBqZz9Wa8fR0kiAgsl9lAW6Sfaj0xUIMg07Op1YznZ89sTtn4zK2nQ0bRrg4/F4OLbL0otwHATgUbuewp30bL+kQQm0o2nQZNj22q6RpqqNU0/T933f65tonAqNW0/Ta3fd046Jxq3QeA2+8U+Hw66JxqvQdOtpJif9rmuk6RZoQkbj63oSFbXlQNMgAFhwdtbM0gOWXin6dZQa2R273UFc8FjuOYkR/sbFBNZp0hmWNEZynZAFDgA3xNFMUHyvQbaK4MKS0lyQ1s8ptVAaCJrIgfVHgiHF3K/99Ze7yaQzep19Os5rlH9pqwGn7bubz5P8c+jkn6eT101CznC8LAnx+yNbYYcnbjsTcjocZ0J8z/b2kaUlMs/v+QrrTjxnH1aWsF3Pz+SejHIju932WH32T0duI9epwLMi15RGJEWfyC265BE4tUkNMhM/CJ2GmGpQHAKkCTGWoYb4tMasEeATfbe+CMjfjYj3q2+aPVehWEnahPgQRhrinHPmc9Fs+welRtH2Vbzco5dYFQGXGN80qjUsxdZ4lcDxrZw8HRMSzZQLBkGGlyQmEqk5fk1IE/4rpdr+nNNA8JQvJPpKkY9psyOndCbN6DMawUavG3WHaNI8ev4F+Zw1ChyRGx0CZxuzRiGEabvwHq8kjpqtwhErQj5iGTYacrUWgbZxqYRgWhLG0XhO0rQR/FmsNZM+YMjszZF1ztaRDhGSXjdCPmLOi5ARvx6GOEqa7aJxWAT9nl7DScHogstm/bh+htUzbCyO90fUF0rkDyanP+kyNAejmlkJvYRWap+qhzQ+qB4yCgXxuR4+5Xp4CjeWxrxQroJ7Af/R2jfCq/iCwDl/Ln3Ppe+59D2h0rc3I31nwdOLW95GblvE+64x2tc0LihjV3LNyMdUr5Mp2DmfwOz9aD6e8e362SSEr5pZLSMWkEuBs0EkuPyLyvAqxAnoZFslCctU02U3ihKeQhtu6VP1SpXX5a+5KLg8W+Tpr6F0PizP+Txf57TNCzNDt3JL6raUvrUmOEr0scxwTh7LDDtnPJIdtnegHTX79l125COlMFOXQ7gaQr4Dbbqd3Do4npiRuQrTUpBvw/npxXga4jnZBLl9mFdt59jR0fvnwVGwo+88lh3HiPKiIe6hhpjPw0OHeXtfmGeVxlA0FG1srCQsRrdguNfxLBTgZGAtoAeDr1EC8lJVYDFbxgMrkKJ8TIxF6HDnl1xf49GS49umZbVuryl3GW0iUjnCaZgTZ6vK3mWxwVUdz1Vb8rC+aj20FU7P/lmtyJ8MEU4WCxJIY5QXpkqi8xlTvucrScRVOL9FM7YSlxi84+bHcU5TuBJ2tg8CMrm7Oal6ZTFnpvLfLQwJLFuIWRLiTV3t1eebnK56Inb6l3fBYPL9cMlHD+U751/0XUOufvbd4/pukztITJx5xREBdEUCI5UcBhYXMuRQ7pKQBhMBzZTJRPACgmSmHICY+gu98gy5KRXOrT45f0Usg4ZOXtIlEhSKsAwFIRdy4+/vk2p3jNf6LIFthFQyZNUXykOJwT0zckPYVCXzrtomC4Xb4lTNuxq+JmBLw3punS0n/9te1D20Fz1G86OZ4B6zh3OberjCRaz/WNYe+TLfOXDbOt4DXuYTLEOkfsF9ioqAEativrqvT/klnDu0e/GBIJv81tuk9t3gDHzUq1qlZCsRP0sHfB+SBmOMW/Q0X48UYq2msa3G2jEMeYBY8wyhZjjfh0WaGjPVi6w5jQpvQdVA5T/b1A1o9g00HJEFXjGZtjaj5E4KPNz+7w2wwsSO4e2LvwFQSwMEFAAAAAgAJkHFXPiPksM/BAAA2REAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWyNmG9zmzgQxr+Kxp3p9F5csABjTG3P1EFc72bSy8Rt77WMZcMVECdE3Hz7W/4EU4fFfpEYePZZ6bfCeMXyJNWPIhJCk59pkhWrSaR17hlGEUYi5cWdzEUGykGqlGs4VUejyJXg+9qUJoY5nTpGyuNssl7W1x7VeilLncSZeFSkKNOUq5eNSORpNaGT1wtP8THS1QVjvcz5UWyF/pY/Kjgzuiz7OBVZEcuMKHFYTT5RL6BOZagjvsfiVPSOSYWyk/JHdfLnfjWZVjMSiQh1lYLDx7O4F0lSZYJ5/NcmnXRjVsb+8Wv2oIYHmB0vxL1M/on3OlpN3AnZiwMvE/0kT59FCzSr8oUyKer/5NTEwrxJWBZapq0ZZpDGWfPJf7aF6BnMKWIwW4N5q8FqDdatBrs12LcaZq1hdqvBaQ31YhpNsepK+1zz9VLJE1FVNGSrDurlqt1Q4Dir7qytVqDG4NPrP7ZfyZMIZRbGSczrxX7/zjWp+ZFsFM/CyNhqrsXvp7gQZMsTUZBmRZeGhuGrJEYIfzBsN7bZjW3WY5vI2O/fUdM1p/ZHmEGe8FAQoCnTjETwFRGqgFLoiLzIUlU3YMkTsmumVFRTIhlPRXFHHnhOBA8joiWJdUGA6AtpIkK5FyTOiI4EqDK5G5k0rHR1C1pd2bubsqOxahoLoXmQmY5+HaC2bcZtj2X2L9+RD1PrtwHz/bj5MzwMeMbB7Qy5/XG3L5IoBu98yMuu0PKIK15EWsHo5mCGYDzDV6l5MrIgdld3u85jI3k+5SpOhure2Ga1rXrAnouKKj6qMFQJGsWBa4f19tvDh43tMRsKclgvnyHueYBt1rHNRtke+MsQ2QwlQxUfVRiqBLNLspnHZuNkTkfmjJL9VWZiCM1B0VDFRxWGKoFzieZ4zBlHm3do8ytoyeCqzVE0VPFRhaFKML9Em3tsPo7mdmju+HetPMKzcQjOReFQxUcVhiqBewnneswdh1t0cItRuK3ItUh3Qg3xLVA+VPFRhaFKsLjkW3hsMc5Hp+ff+uko4d+hlghf6xwCxCUflxguBa3Ug6RTDwxXMHstDR3F/CKf0XVsrYOcqOTjEsOloJX6nBQ46RXOc/tEzVFOX4Q4p4lzopKPSwyXglbqc5rAaV7htM6c1vgDlWclNDqDmBaOiUo+LjFcClqpj2kBpnUF89zH0PFGJhA7hXLizQwu+bjEcCmgbxoaCh0NvdLS0HNPQ681NSocbJQp3tfgko9LDJcC+qa3odDc0CvdDT23N7TpIOa3d7ctZONz++3iZmDclvsy+t727rFo/020b3s+Fs3eRDN7sAJtwS6jA9sL0HoZvR1pKtSxfndQwIaszHSzQeyutu8nTC+od+WX16v3FvVO95ymeekBt9ExzgqSiAOknN7NYUVVs2VrTrTM633vTmrYztWHzcayCgD9IKV+PakG6N7mrP8HUEsDBBQAAAAIACZBxVyOgk7SRAMAANQQAAANAAAAeGwvc3R5bGVzLnhtbN1YbW+bMBD+K4gfMAIkLExJpIQGadI2VWo/7KsTTGLJvAycLumvn88mQBJflq7tJo2owr7z89zj83GgTmpx4PRhS6mw9hnP66m9FaL85Dj1ekszUn8oSppLT1pUGRFyWm2cuqwoSWoAZdzxBoPAyQjL7dkk32VxJmprXexyMbUHtjObpEXeWQJbG+RSklHrifCpHRHOVhVTa0nG+EGbPTCsC15UlpBS6NR2wVI/a7erZ6Cy4clYXlRgdHSEK3FWDUcXotqspN6Bu/Tno/gkjncbJcMog+HH4XjRpwxfKTJW12Uyfk/Zrh+8XIIRq2615GCct4c8srVhNimJELTKYzlRGGW8cFnN+PFQylPeVOTgeiP7ZkBdcJZAyE3Uz9LdYjmPl4qmB30laVcfb0gaj+N5HKGk6iZzvCqqhFZtlj37aJpNOE2FhFdss4W7KEo4vEKIIpODhJFNkRN1BEdEH2mpJjC1xVY9xCe1Fi3uRktdXbC0iXEjQq1Vcm4EyJVH3Tci9OLexpqBzNeacv4AJN/TNmmupNqnlu5TnxNoURaU8HEoM90MNY2eQKA+m+bu0YZ/RGuV7KkQi53cQa7mP3aFoPcVTdlezfdpGx9jdzt274ydlCU/zDnb5BnVe7854GxCjjhrW1TsWUaDZ38tDVT32H2Ki/LwLf87UX4nyu+Lct9TlPWzIuUj3YumiV5VOETS9gI555Uy/Fvso459+EbsAyO7967aX8HuND2h13hO2k5rteBVO7W/wRcV7yis1Y5xwfJmtmVJQvOL7iPpBVnJT7YTfrk+oSnZcfHYOqd2N/5KE7bLwnbVPWyrWdWNv0C7doP2e0PGYnlC9zSJmqnsvydvLn0B4NzTfaNcejCM9pk94MPiYAowjEZhcf6n/YzR/Wgfpm1s9IxRzBjFaJTJE6kfFseMCeVl3mkY+n4QYBmNIqOCCMtbEMCfmQ3TBggsDkR6Wa7x08Yr5HodYGd6rUKwneKViO0UzzV4zHkDRBiaTxuLAwjsFLDagfjmOFBTZozvw6li2rAnGPeEIeaBWjTXaBAg2QngZz4f7Cnx/TA0e8BnVuD7mAeeRtyDKQANmMf31Xvw7H3kHN9TTvd/jNkvUEsDBBQAAAAIACZBxVyXirscwAAAABMCAAALAAAAX3JlbHMvLnJlbHOdkrluwzAMQH/F0J4wB9AhiDNl8RYE+QFWog/YEgWKRZ2/r9qlcZALGXk9PBLcHmlA7TiktoupGP0QUmla1bgBSLYlj2nOkUKu1CweNYfSQETbY0OwWiw+QC4ZZre9ZBanc6RXiFzXnaU92y9PQW+ArzpMcUJpSEszDvDN0n8y9/MMNUXlSiOVWxp40+X+duBJ0aEiWBaaRcnToh2lfx3H9pDT6a9jIrR6W+j5cWhUCo7cYyWMcWK0/jWCyQ/sfgBQSwMEFAAAAAgAJkHFXEs11FQ5AQAAKgIAAA8AAAB4bC93b3JrYm9vay54bWyNUdFuwjAM/JUqH7AWtCENUV6GtiFNGxoT76F1qUUSV44Lg6+f26oa0l72lNzZutxdFmfi457omHx7F2JuapFmnqaxqMHbeEcNBJ1UxN6KQj6ksWGwZawBxLt0mmWz1FsMZrkYtTac3gISKAQpKNkRO4Rz/J13MDlhxD06lEtu+rsDk3gM6PEKZW4yk8Sazq/EeKUg1m0LJudyMxkGO2DB4g+97Ux+2X3sGbH7T6tGcjPLVLBCjtJv9PpWPZ5AlwfUCj2jE+CVFXhhahsMh05GU6Q3MfoexnMocc7/qZGqCgtYUdF6CDL0yOA6gyHW2ESTBOshN1vrICbb1nvLly6XPrQuh4yi5m4a4znqgNflYHP0VkKFAcp3lYvKa0/FhpPu6HWm9w+TR+2jde5JuY/wRrYco47ftPwBUEsDBBQAAAAIACZBxVwkHpuirQAAAPgBAAAaAAAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHO1kT0OgzAMha8S5QA1UKlDBUxdWCsuEAXzIxISxa4Kty+FAZA6dGGyni1/78lOn2gUd26gtvMkRmsGymTL7O8ApFu0ii7O4zBPahes4lmGBrzSvWoQkii6QdgzZJ7umaKcPP5DdHXdaXw4/bI48A8wvF3oqUVkKUoVGuRMwmi2NsFS4stMlqKoMhmKKpZwWiDiySBtaVZ9sE9OtOd5Fzf3Ra7N4wmu3wxweHT+AVBLAwQUAAAACAAmQcVcZZB5khkBAADPAwAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytk01OwzAQha8SZVslLixYoKYbYAtdcAFjTxqr/pNnWtLbM07aSqASFYVNrHjevM+el6zejxGw6J312JQdUXwUAlUHTmIdIniutCE5SfyatiJKtZNbEPfL5YNQwRN4qih7lOvVM7Ryb6l46XkbTfBNmcBiWTyNwsxqShmjNUoS18XB6x+U6kSouXPQYGciLlhQiquEXPkdcOp7O0BKRkOxkYlepWOV6K1AOlrAetriyhlD2xoFOqi945YaYwKpsQMgZ+vRdDFNJp4wjM+72fzBZgrIyk0KETmxBH/HnSPJ3VVkI0hkpq94IbL17PtBTluDvpHN4/0MaTfkgWJY5s/4e8YX/xvO8RHC7r8/sbzWThp/5ovhP15/AVBLAQIUAxQAAAAIACZBxVxGx01IlQAAAM0AAAAQAAAAAAAAAAAAAACAAQAAAABkb2NQcm9wcy9hcHAueG1sUEsBAhQDFAAAAAgAJkHFXHjBD37uAAAAKwIAABEAAAAAAAAAAAAAAIABwwAAAGRvY1Byb3BzL2NvcmUueG1sUEsBAhQDFAAAAAgAJkHFXJlcnCMQBgAAnCcAABMAAAAAAAAAAAAAAIAB4AEAAHhsL3RoZW1lL3RoZW1lMS54bWxQSwECFAMUAAAACAAmQcVc+I+Swz8EAADZEQAAGAAAAAAAAAAAAAAAgIEhCAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sUEsBAhQDFAAAAAgAJkHFXI6CTtJEAwAA1BAAAA0AAAAAAAAAAAAAAIABlgwAAHhsL3N0eWxlcy54bWxQSwECFAMUAAAACAAmQcVcl4q7HMAAAAATAgAACwAAAAAAAAAAAAAAgAEFEAAAX3JlbHMvLnJlbHNQSwECFAMUAAAACAAmQcVcSzXUVDkBAAAqAgAADwAAAAAAAAAAAAAAgAHuEAAAeGwvd29ya2Jvb2sueG1sUEsBAhQDFAAAAAgAJkHFXCQem6KtAAAA+AEAABoAAAAAAAAAAAAAAIABVBIAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzUEsBAhQDFAAAAAgAJkHFXGWQeZIZAQAAzwMAABMAAAAAAAAAAAAAAIABORMAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAkACQA+AgAAgxQAAAAA"

@app.route("/gst-template/<ttype>")
def gst_template_download(ttype):
    """Serve pre-built sales summary Excel templates."""
    import base64 as _b64, io
    if ttype == "consolidated":
        data  = _b64.b64decode(_GST_CONS_B64)
        fname = "GST_Sales_Consolidated_Template.xlsx"
    else:
        data  = _b64.b64decode(_GST_BRANCH_B64)
        fname = "GST_Sales_BranchWise_Template.xlsx"
    from flask import send_file
    return send_file(
        io.BytesIO(data),
        download_name=fname,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/tool/gst-reconciliation")
@login_required
def tool_gst_reconciliation():
    user = get_user_by_id(session["uid"])
    return render_template_string(GST_RECON_T, **user_ctx(user))


# NOTE: gunicorn timeout should be >= 180s for large ZIPs.
# In gunicorn.conf.py set: timeout = 180
@app.route("/gst-process", methods=["POST"])
@login_required
def gst_process():
    try:
        user = get_user_by_id(session["uid"])
        if not user["is_admin"] and uploads_remaining(user) <= 0:
            return jsonify({"status": "error",
                "message": f"No uploads remaining. Contact {CONTACT_EMAIL} to recharge."})

        if "sales_file" not in request.files or "gst_file" not in request.files:
            return jsonify({"status": "error", "message": "Please upload both files."})

        sales_f = request.files["sales_file"]
        gst_f = request.files["gst_file"]

        if not sales_f.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({"status": "error", "message": "Sales file must be .xlsx or .xls"})
        if not gst_f.filename.lower().endswith('.zip'):
            return jsonify({"status": "error", "message": "GSTR 3B file must be a .zip"})

        try:
            mappings = json.loads(request.form.get("mappings", "{}"))
        except:
            return jsonify({"status": "error", "message": "Invalid mapping data."})

        # Handle consolidated checkbox
        consolidated_mode = request.form.get("consolidated_mode", "").lower() == "true"
        consolidated_col  = request.form.get("consolidated_col", "").strip()
        if consolidated_mode:
            mappings["__consolidated__"] = consolidated_col or "__auto__"

        if not mappings:
            return jsonify({"status": "error", "message": "Please provide at least one state-column mapping or tick Consolidated Sales."})

        on = request.form.get("output_name", "").strip()
        h = uuid.uuid4().hex
        sales_path = os.path.join(UPLOAD_DIR, f"{h}_sales.xlsx")
        gst_path = os.path.join(UPLOAD_DIR, f"{h}_gst.zip")
        op = os.path.join(OUTPUT_DIR, f"{h}_out.xlsx")

        try:
            orig = sales_f.filename.lower()
            if orig.endswith('.xls') and not orig.endswith('.xlsx'):
                xls_tmp = os.path.join(UPLOAD_DIR, f"{h}_sales.xls")
                sales_f.save(xls_tmp)
                _convert_xls_to_xlsx(xls_tmp, sales_path)
                try: os.remove(xls_tmp)
                except: pass
            else:
                sales_f.save(sales_path)
            gst_f.save(gst_path)

            result = _process_gst_reconciliation(sales_path, gst_path, mappings, op)
        except Exception as e:
            return jsonify({"status": "error", "message": f"Processing error: {e}"})
        finally:
            for p in (sales_path, gst_path):
                try: os.remove(p)
                except: pass

        if result['status'] != 'success':
            return jsonify(result)

        fname = f"{on or 'GST_Reconciliation'}.xlsx"
        log_usage(user["id"], fname)
        return jsonify({"status": "success", "log": result["log"],
                        "file_id": h, "filename": fname})
    except Exception as e:
        return jsonify({"status": "error", "message": f"Unexpected error: {e}"}), 500


# ═══════════════════════════════════════════════════════════════════════════════
#  TRIAL BALANCE → BALANCE SHEET ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

try:
    from tb_processor import analyze_trial_balance, process_tb_to_bs
    TB_PROCESSOR_AVAILABLE = True
except ImportError:
    TB_PROCESSOR_AVAILABLE = False


@app.route("/tool/tb-to-bs")
def tb_to_bs_page():
    if "uid" not in session:
        return redirect("/login")
    user = get_user_by_id(session["uid"])
    if not user:
        return redirect("/login")
    ctx = user_ctx(user)
    return render_template_string(TB_BS_TEMPLATE, **ctx)


@app.route("/tb-analyse", methods=["POST"])
def tb_analyse():
    if "uid" not in session:
        return jsonify({"status": "error", "message": "Session expired — please refresh the page and log in again"}), 401
    user = get_user_by_id(session["uid"])
    if not user:
        return jsonify({"status": "error", "message": "Session expired — please refresh the page and log in again"}), 401
    if not TB_PROCESSOR_AVAILABLE:
        return jsonify({"status": "error", "message": "TB processor not available on this server"}), 500

    try:
        tb_file = request.files.get("tb_file")
        if not tb_file:
            return jsonify({"status": "error", "message": "No trial balance file uploaded"})

        import tempfile, os
        tmp = tempfile.mkdtemp()
        # Preserve original extension for PDF detection
        orig_name = tb_file.filename or "tb.xlsx"
        ext = ".pdf" if orig_name.lower().endswith(".pdf") else ".xlsx"
        tb_path = os.path.join(tmp, "tb" + ext)
        tb_file.save(tb_path)

        result = analyze_trial_balance(tb_path)

        try:
            os.remove(tb_path)
            os.rmdir(tmp)
        except Exception:
            pass

        if "error" in result:
            return jsonify({"status": "error", "message": result["error"]})

        return jsonify({"status": "success", **result})

    except Exception as e:
        import traceback
        return jsonify({"status": "error", "message": f"Analysis failed: {e}\n{traceback.format_exc()}"}), 500


@app.route("/tb-read-bs", methods=["POST"])
def tb_read_bs():
    """Read Capital Account and Fixed Assets sheets from uploaded BS template."""
    if "uid" not in session:
        return jsonify({"status": "error", "message": "Not logged in"}), 401
    try:
        import tempfile, os, re
        from openpyxl import load_workbook
        from openpyxl.cell import MergedCell
        from openpyxl.utils import column_index_from_string

        bs_file = request.files.get("bs_file")
        if not bs_file:
            return jsonify({"status": "error", "message": "No BS template uploaded"})

        tmp = tempfile.mkdtemp()
        bs_path = os.path.join(tmp, "bs.xlsx")
        bs_file.save(bs_path)

        wb = load_workbook(bs_path, read_only=True, data_only=True)
        # Parallel formula-view workbook — used to resolve cells whose
        # data_only value is None because they hold a formula with no
        # cached <v> (common after tb_processor.py's openpyxl round-trip,
        # which drops cached values for formulas it didn't itself compute).
        # E.g. capital!B8 = "=F40" (the proprietor's name lives in F40, and
        # B8 just references it) — data_only reads B8 as None, but we can
        # resolve it ourselves for simple same-sheet cell references.
        wb_f = load_workbook(bs_path, read_only=True, data_only=False)
        result = {"capital": None, "fixed_assets": None}

        def _resolve_cell(ws_do, ws_f, row, col):
            """Return ws_do.cell(row,col).value, or — if that's None and the
            formula-view cell is a simple same-sheet reference like '=F40' —
            the resolved value of the referenced cell instead."""
            val = ws_do.cell(row, col).value
            if val is not None:
                return val
            try:
                fval = ws_f.cell(row, col).value
            except Exception:
                return val
            if isinstance(fval, str) and fval.startswith('='):
                m = re.match(r'^=\$?([A-Z]+)\$?(\d+)$', fval.strip())
                if m:
                    ref_col_letters, ref_row = m.group(1), int(m.group(2))
                    ref_col = column_index_from_string(ref_col_letters)
                    try:
                        return ws_do.cell(ref_row, ref_col).value
                    except Exception:
                        return val
            return val

        # ── Read Capital Account sheet ──────────────────────────────────
        cap_sheet = None
        for sn in wb.sheetnames:
            if "capital" in sn.lower():
                cap_sheet = sn
                break

        if cap_sheet:
            ws = wb[cap_sheet]
            ws_f = wb_f[cap_sheet]
            rows_data = []
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=50, max_col=15, values_only=False), start=1):
                r = []
                for col_idx, c in enumerate(row, start=1):
                    if isinstance(c, MergedCell):
                        r.append(None)
                    else:
                        r.append(_resolve_cell(ws, ws_f, row_idx, col_idx))
                rows_data.append(r)

            # Find the header row with "Sr. No." or "Name of Proprietor/Partner"
            partners = []
            header_row_idx = None
            name_col = None  # which column has names
            opening_col = None  # which column has opening balance

            for i, row in enumerate(rows_data):
                row_str = " ".join(str(v or "").lower() for v in row)
                if ("sr" in row_str and "name" in row_str) or \
                   ("name of" in row_str and ("proprietor" in row_str or "partner" in row_str)):
                    header_row_idx = i
                    # Find which col is "Name" and which is "As at 1st April"
                    for ci, val in enumerate(row):
                        vl = str(val or "").lower().strip()
                        if "name" in vl and ("proprietor" in vl or "partner" in vl or "of" in vl):
                            name_col = ci
                        if "as at" in vl and "april" in vl:
                            opening_col = ci
                    break

            if header_row_idx is None:
                # Fallback: look for "As at 1st April" row
                for i, row in enumerate(rows_data):
                    row_str = " ".join(str(v or "").lower() for v in row)
                    if "as at" in row_str and "april" in row_str:
                        header_row_idx = i
                        for ci, val in enumerate(row):
                            if isinstance(val, str) and "as at" in val.lower() and "april" in val.lower():
                                opening_col = ci
                                break
                        name_col = 1  # default
                        break

            # Stop words — rows containing these are NOT partners
            _cap_stop = {"total", "previous year", "previous year (py)", "py",
                         "chartered accountant", "ca.", "auditor", "partner ",
                         "proprietor", "director", "secretary", "for ",
                         "sd/-", "authorised", "firm", "registration",
                         "arun gupta", "arun kumar",  # signatory names from image
                         }

            if header_row_idx is not None:
                nc = name_col if name_col is not None else 1
                oc = opening_col if opening_col is not None else 2
                for i in range(header_row_idx + 1, min(header_row_idx + 12, len(rows_data))):
                    row = rows_data[i]
                    if not row or all(v is None for v in row):
                        break  # blank row = end of data

                    name_val = row[nc] if len(row) > nc else None
                    if not isinstance(name_val, str) or len(name_val.strip()) < 2:
                        continue

                    nm = name_val.strip()
                    nm_low = nm.lower()

                    # Remove parentheses wrapper if present
                    if nm.startswith("("):
                        nm = nm.lstrip("(").rstrip(")")
                        nm_low = nm.lower()

                    # Skip stop words
                    if any(sw in nm_low for sw in _cap_stop):
                        continue  # skip this row, keep scanning (don't break)

                    # Skip if too short or looks like a label
                    if len(nm) < 3:
                        continue
                    if "account" in nm_low and "capital" not in nm_low:
                        continue

                    # Skip rows that have no number at all (pure text rows / signatures)
                    has_any_number = any(
                        isinstance(row[ci], (int, float)) and row[ci] != 0
                        for ci in range(2, min(len(row), 10))
                    )
                    # Also check if opening is there (row might have 0 opening for new partner)
                    has_sr_no = isinstance(row[0], (int, float))

                    if not has_any_number and not has_sr_no:
                        continue  # no Sr. No. and no numbers = not a data row

                    # Get opening balance
                    opening = 0
                    if len(row) > oc and isinstance(row[oc], (int, float)):
                        opening = float(row[oc])
                    else:
                        for ci in range(oc, min(oc + 3, len(row))):
                            if isinstance(row[ci], (int, float)) and row[ci] != 0:
                                opening = float(row[ci])
                                break

                    partners.append({"name": nm, "opening": opening, "row": i + 1})

            if partners:
                # Detect column layout from header row
                cap_columns = []
                col_map = {}  # {field_key: column_index (1-based)}
                if header_row_idx is not None:
                    hrow = rows_data[header_row_idx]
                    for ci, hv in enumerate(hrow):
                        hs = str(hv or "").lower().strip()
                        if "introduced" in hs or "capital intro" in hs:
                            cap_columns.append({"key": "introduced", "label": "Capital Introduced", "col": ci+1})
                            col_map["introduced"] = ci + 1
                        elif "interest" in hs and "capital" in hs:
                            cap_columns.append({"key": "interest_on_capital", "label": "Interest on Capital", "col": ci+1})
                            col_map["interest_on_capital"] = ci + 1
                        elif "salary" in hs:
                            cap_columns.append({"key": "salary", "label": "Salary", "col": ci+1})
                            col_map["salary"] = ci + 1
                        elif "withdraw" in hs:
                            cap_columns.append({"key": "withdrawals", "label": "Withdrawals", "col": ci+1})
                            col_map["withdrawals"] = ci + 1

                result["capital"] = {
                    "sheet": cap_sheet,
                    "partners": partners,
                    "columns": cap_columns,
                    "col_map": col_map,
                }

        # ── Read Fixed Assets sheet ─────────────────────────────────────
        fa_sheet = None
        for sn in wb.sheetnames:
            sl = sn.lower()
            if "fixed asset" in sl or "fa " in sl or sl.startswith("fa") or "ppe" in sl:
                fa_sheet = sn
                break

        if fa_sheet:
            # ── Build opening WDV lookup from Fixed Assets P. Yr. sheet ──
            # Col I (index 8) of P.Yr. sheet = closing WDV = opening for current year
            # This is the authoritative source — C.Yr. col B is just =P.Yr.!I9 formula
            py_opening = {}  # {asset_name_lower: opening_wdv}
            py_sheet_name = None
            for sn in wb.sheetnames:
                if "p. yr" in sn.lower() or "p.yr" in sn.lower() or \
                   ("fixed" in sn.lower() and "p" in sn.lower()):
                    py_sheet_name = sn
                    break

            if py_sheet_name:
                ws_py = wb[py_sheet_name]
                py_rows = []
                for row in ws_py.iter_rows(min_row=1, max_row=45, max_col=12, values_only=False):
                    r = []
                    for c in row:
                        r.append(None if isinstance(c, MergedCell) else c.value)
                    py_rows.append(r)

                # Find header to locate closing WDV col (W.D.V AS ON 31.03.xxxx)
                py_closing_col = 8  # col I default (0-indexed)
                py_rate_col = None
                for i, row in enumerate(py_rows[:8]):
                    row_str = " ".join(str(v or "").lower() for v in row)
                    if "w.d.v" in row_str and ("31.03" in row_str or "closing" in row_str):
                        for ci, val in enumerate(row):
                            vl = str(val or "").lower()
                            if ("w.d.v" in vl or "as on" in vl) and ci > 4:
                                py_closing_col = ci
                        break

                # Read asset names and closing WDV
                import re as _re2
                _date_re2 = _re2.compile(r'^\d{1,2}[./]\d{1,2}[./]\d{2,4}$')

                # Use col I (index 8) for closing WDV if it has values,
                # otherwise fall back to col B (index 1, opening WDV of P.Yr.)
                # Col I is a formula (=F-H) and loses its cached <v> after an
                # openpyxl round-trip, while col B holds plain numeric constants
                # that survive.  Either way, the goal is to give the user a
                # non-zero reference figure to start from.
                ws_py_f = wb_f[py_sheet_name]
                for row in py_rows[5:]:
                    nm = str(row[0] or "").strip()
                    if not nm or len(nm) < 2:
                        continue
                    if nm.isupper():  # skip category headers
                        continue
                    if _date_re2.match(nm):
                        continue
                    if any(sw in nm.lower() for sw in
                           {"total", "particular", "addition", "amount", "rate",
                            "w.d.v", "building", "property", "chair"}):
                        continue
                    closing = row[py_closing_col] if len(row) > py_closing_col else None
                    if not isinstance(closing, (int, float)):
                        # col I is None → fall back to col B (opening WDV of P.Yr.)
                        closing = row[1] if len(row) > 1 else None
                    if isinstance(closing, (int, float)):
                        py_opening[nm.lower().strip()] = float(closing)

            ws = wb[fa_sheet]
            ws_f_fa = wb_f[fa_sheet]
            # Build a parallel formula-view row list for resolving cross-sheet
            # refs in col A (asset names like ='Fixed Assets P. Yr.'!A10) and
            # col B (opening WDV like ='Fixed Assets P. Yr.'!I10) that have no
            # cached <v> after an openpyxl round-trip.
            rows_data_f = []
            for row in ws_f_fa.iter_rows(min_row=1, max_row=60, max_col=15, values_only=False):
                r = []
                for c in row:
                    if isinstance(c, MergedCell):
                        r.append(None)
                    else:
                        r.append(c.value if hasattr(c, 'value') else None)
                rows_data_f.append(r)

            rows_data = []
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=60, max_col=15, values_only=False), start=1):
                r = []
                for col_idx, c in enumerate(row, start=1):
                    if isinstance(c, MergedCell):
                        r.append(None)
                    else:
                        r.append(c.value if hasattr(c, 'value') else None)
                rows_data.append(r)

            # Find header row (PARTICULARS / W.D.V / ADDITIONS / SALE / RATE)
            fa_header_row = None
            wdv_col = None  # opening WDV column
            rate_col = None
            for i, row in enumerate(rows_data):
                row_str = " ".join(str(v or "").lower() for v in row)
                if ("particular" in row_str or "w.d.v" in row_str) and \
                   ("addition" in row_str or "rate" in row_str or "sale" in row_str):
                    fa_header_row = i
                    for ci, val in enumerate(row):
                        vl = str(val or "").lower().strip()
                        if "w.d.v" in vl or "opening" in vl or "01.04" in vl or "as on" in vl:
                            wdv_col = ci
                        if vl in ("rate", "%", "rate %"):
                            rate_col = ci
                    break

            # Skip words for FA — not actual assets
            import re as _re
            _fa_skip = {"particular", "w.d.v", "amount", "total", "grand total",
                        "rate", "addition", "sale", "depreciation",
                        "as on", "note", "ca.", "chartered", "auditor",
                        "sd/-", "partner", "proprietor", "director", "for ",
                        "property, plant", "intangible asset",
                        "amount in rs", "amount in"}

            # Date pattern: 01.04.2024, 31.03.2025, 1.4.2024 etc
            _date_re = _re.compile(r'^\d{1,2}[./]\d{1,2}[./]\d{2,4}$')
            # Note/reference number pattern: "7 Property, Plant..." or starts with digit+space
            _note_re = _re.compile(r'^\d+\s+\w')

            assets = []
            start_row = (fa_header_row + 1) if fa_header_row is not None else 5
            wc = wdv_col if wdv_col is not None else 1  # default col B for opening WDV
            rc = rate_col  # rate column

            for i in range(start_row, min(start_row + 40, len(rows_data))):
                row = rows_data[i]
                row_f = rows_data_f[i] if i < len(rows_data_f) else row
                if not row or all(v is None for v in row):
                    continue

                # Get asset name from col A — if data_only is None (formula cell
                # with no cached <v>), check the formula-view cell: if it's a
                # cross-sheet ref like ='Fixed Assets P. Yr.'!A10, look that
                # name up in py_opening keys as a fallback.
                name_val = row[0] if len(row) > 0 else None
                if name_val is None and len(row_f) > 0:
                    fval = row_f[0]
                    if isinstance(fval, str) and fval.startswith('=') and '!' in fval:
                        # Cross-sheet ref: find matching name from py_opening
                        for known_name in py_opening:
                            name_val = known_name
                            break
                        # Better: use formula-view of py sheet directly
                        m = __import__('re').search(r"!([A-Z]+)(\d+)$", fval)
                        if m:
                            ref_col_s = m.group(1)
                            ref_row_i = int(m.group(2))
                            from openpyxl.utils import column_index_from_string as _c2i
                            try:
                                ws_py_do_tmp = wb[py_sheet_name]
                                name_val = ws_py_do_tmp.cell(ref_row_i, _c2i(ref_col_s)).value
                            except Exception:
                                pass
                if name_val is None and len(row) > 1:
                    name_val = row[1]

                if not isinstance(name_val, str) or len(name_val.strip()) < 2:
                    continue

                nm = name_val.strip()
                nm_low = nm.lower()

                # Skip "Total" row — stop scanning
                if nm_low.strip() in ("total", "grand total"):
                    break

                # Skip if it matches any stop word
                if any(sw in nm_low for sw in _fa_skip):
                    continue

                # Skip dates like "01.04.2024"
                if _date_re.match(nm):
                    continue

                # Skip note numbers like "7 Property, Plant and Equipment"
                if _note_re.match(nm):
                    continue

                # Skip pure numbers
                if nm.replace(",", "").replace(".", "").replace("-", "").isdigit():
                    continue

                # Skip ALL-CAPS category headers (PLANT & MACHINERY, VEHICLE, etc.)
                has_own_number = False
                if len(row) > wc and isinstance(row[wc], (int, float)) and row[wc] != 0:
                    has_own_number = True
                if nm.isupper() and not has_own_number:
                    continue

                # ── Opening WDV: prefer P.Yr. closing col I (authoritative) ──
                # The C.Yr. col B is a formula (=P.Yr.!I9) — data_only may be stale.
                # Match by name to P.Yr. lookup table first.
                opening_wdv = py_opening.get(nm_low.strip(), None)
                if opening_wdv is None:
                    # Fallback: read C.Yr. col B (data_only cached value)
                    if len(row) > wc and isinstance(row[wc], (int, float)):
                        opening_wdv = float(row[wc])
                    else:
                        opening_wdv = 0

                rate = 0
                if rc is not None and len(row) > rc and isinstance(row[rc], (int, float)):
                    rate = float(row[rc])
                else:
                    # Try to find rate in later columns (look for value 5-100)
                    for ci in range(max(wc + 3, 5), min(len(row), 12)):
                        v = row[ci]
                        if isinstance(v, (int, float)) and 5 <= v <= 100:
                            rate = float(v)
                            break

                assets.append({
                    "name": nm,
                    "opening_wdv": opening_wdv,
                    "rate": rate,
                    "row": i + 1,
                })

            if assets:
                result["fixed_assets"] = {
                    "sheet": fa_sheet,
                    "assets": assets,
                }

        wb.close()
        wb_f.close()
        try:
            os.remove(bs_path)
            os.rmdir(tmp)
        except:
            pass

        return jsonify({"status": "success", **result})

    except Exception as e:
        import traceback
        return jsonify({"status": "error", "message": f"Failed to read BS: {e}\n{traceback.format_exc()}"}), 500



def _rollover_fixed_assets(output_path, cy_year, log, source_path=None):
    """
    Fixed-assets rollover for the year-shift tool.

    Behavior:
    1. Output Fixed Assets P. Yr. becomes a value snapshot of the uploaded
       Fixed Assets C. Yr. sheet.
    2. Output Fixed Assets C. Yr. keeps formulas, but additions/sale inputs are
       cleared so the new year opens from the mirrored PY closing balances.
    """
    from openpyxl import load_workbook as _lwb
    from openpyxl.cell import MergedCell as _MC
    import re as _re

    wb = _lwb(output_path)
    wb_do = _lwb(output_path, data_only=True)
    src_wb = src_wb_do = None
    if source_path:
        src_wb = _lwb(source_path)
        src_wb_do = _lwb(source_path, data_only=True)

    try:
        cy_sn, py_sn = detect_fixed_asset_sheet_names(wb.sheetnames)
        if not cy_sn:
            log.append("⚠ FA C.Yr. sheet not found")
            return

        ws_cy = wb[cy_sn]
        ws_cy_do = wb_do[cy_sn]

        src_sheetnames = src_wb.sheetnames if src_wb else wb.sheetnames
        src_cy_sn, _ = detect_fixed_asset_sheet_names(src_sheetnames)
        src_cy_ws = src_wb[src_cy_sn] if src_wb and src_cy_sn else wb[cy_sn]
        src_cy_ws_do = src_wb_do[src_cy_sn] if src_wb_do and src_cy_sn else wb_do[cy_sn]

        # Detect FA layout from current CY sheet
        op_col = 2; ag_col = 3; al_col = 4; sl_col = 5; rt_col = 7; cl_col = 9
        data_start = 9
        date_row = 7
        for r in range(1, 15):
            vals = []
            for c in range(1, 12):
                cell = ws_cy.cell(r, c)
                vals.append("" if isinstance(cell, _MC) else str(cell.value or "").lower().strip())
            row_str = " ".join(vals)
            if "01.04" in row_str or "31.03" in row_str:
                date_row = r
            if "greater" in row_str or ("addition" in row_str and "sale" in row_str):
                data_start = r + 1
                for ci, v in enumerate(vals, 1):
                    if "w.d.v" in v and ci < 3: op_col = ci
                    elif "greater" in v: ag_col = ci
                    elif "less" in v and v != "less": al_col = ci
                    elif v in ("sale", "sales"): sl_col = ci
                    elif v in ("%", "rate", "rate %"): rt_col = ci
                    elif "w.d.v" in v and ci > 5: cl_col = ci

        # 1) Mirror source CY sheet → output PY sheet as a direct row-by-row copy.
        #
        # The PY sheet should be a COMPLETE HISTORICAL RECORD of the uploaded
        # CY year — including all additions (Camera DVR, Steel Angle, Car
        # additions etc.) and any sales (Property sale). This is the factual
        # data for the prior year and must not be filtered or omitted.
        #
        # Approach: copy every cell from src_cy_ws directly into ws_py,
        # row-for-row, column-for-column. For formula cells, write the
        # RESOLVED VALUE (from src_cy_ws_do) rather than the formula text,
        # since the PY sheet is a value-snapshot, not a live calculation sheet.
        # Exception: same-sheet formulas that have NO cross-sheet refs and
        # whose resolved value would be None (e.g. =B8+C8... on a section
        # header row) are left as None (blank), matching the source display.
        copied = 0
        if py_sn:
            ws_py = wb[py_sn]

            # First, blank the entire PY sheet so stale data from the old
            # prior-prior-year template doesn't bleed through.
            merged_children_py = set()
            for merged in ws_py.merged_cells.ranges:
                min_col, min_row, max_col, max_row = merged.bounds
                for rr in range(min_row, max_row + 1):
                    for cc in range(min_col, max_col + 1):
                        if (rr, cc) != (min_row, min_col):
                            merged_children_py.add((rr, cc))

            for r in range(1, ws_py.max_row + 1):
                for c in range(1, min(ws_py.max_column, src_cy_ws.max_column) + 1):
                    if (r, c) in merged_children_py:
                        continue
                    from openpyxl.cell import MergedCell as _MC2
                    tgt = ws_py.cell(r, c)
                    if isinstance(tgt, _MC2):
                        continue
                    tgt.value = None

            # Copy source CY → output PY row-by-row, cell-by-cell.
            # Find the last row that belongs to the FA table: the last row
            # that has EITHER a non-empty col A (asset/section name or Total)
            # OR a numeric rate value in col G (the Rate % column).
            # This excludes stray formula-overflow values that appear below
            # the table boundary in columns B-I with no corresponding label.
            src_last_row = 1
            for _r in range(src_cy_ws.max_row, 0, -1):
                _a = src_cy_ws_do.cell(_r, 1).value
                _g = src_cy_ws_do.cell(_r, rt_col).value
                if _a is not None or isinstance(_g, (int, float)):
                    src_last_row = _r
                    break

            for r in range(1, src_last_row + 1):
                for c in range(1, 10):  # strictly cols 1-9 only
                    src_cell_f = src_cy_ws.cell(r, c)
                    from openpyxl.cell import MergedCell as _MC3
                    if isinstance(src_cell_f, _MC3):
                        continue
                    val = src_cy_ws_do.cell(r, c).value
                    if val is None:
                        raw = src_cell_f.value
                        if not (isinstance(raw, str) and raw.startswith('=')):
                            val = raw
                    from openpyxl.cell import MergedCell as _MC4
                    tgt = ws_py.cell(r, c)
                    if isinstance(tgt, _MC4):
                        continue
                    tgt.value = val
                    if val is not None:
                        copied += 1

            log.append(f"✓ FA PY: copied source CY sheet into '{py_sn}' ({copied} cells)")

            # 1b) Fix CY opening-WDV formulas after PY restructure.
            #
            # The CY sheet's B column contains cross-sheet formulas like
            # ='Fixed Assets P. Yr.'!I9 that pull the opening WDV from the
            # original PY sheet. Those row numbers were hardcoded to the
            # OLD PY layout (e.g. Battery was at PY row 9). Now that we've
            # replaced the PY sheet with a copy of the CY sheet, the same
            # assets appear at the SAME row numbers as in CY (Battery is
            # now at PY row 10, matching CY row 10). The old formula
            # ='Fixed Assets P. Yr.'!I9 now picks up a section-header
            # (PLANT & MACHINERY, I=blank) instead of Battery's closing WDV.
            #
            # Fix: for each CY B-column cross-sheet formula referencing PY,
            # check whether the referenced PY row in the NEW PY sheet still
            # holds the correct asset name. If not, find the correct row in
            # the new PY (same row as the CY asset row) and update the
            # formula to reference that row instead.
            import re as _re_cyfix
            for r in range(1, ws_cy.max_row + 1):
                b_cell = ws_cy.cell(r, 2)  # B column = opening WDV
                from openpyxl.cell import MergedCell as _MC5
                if isinstance(b_cell, _MC5):
                    continue
                bval = b_cell.value
                if not (isinstance(bval, str) and bval.startswith('=')
                        and py_sn.lower() in bval.lower()):
                    continue
                # Extract referenced PY row number
                m = _re_cyfix.search(r'!I(\d+)', bval)
                if not m:
                    continue
                old_py_row = int(m.group(1))
                # The CY asset name at this row
                cy_a = ws_cy_do.cell(r, 1).value
                if not cy_a:
                    cy_a = ws_cy.cell(r, 1).value
                if not cy_a or str(cy_a).startswith('='):
                    continue
                cy_a_norm = str(cy_a).strip().lower()
                # Find the correct row in the NEW PY sheet for this asset.
                # Since new PY = copy of CY, the asset is at the same row.
                new_py_row = r  # new PY row = CY row
                if new_py_row != old_py_row:
                    new_formula = bval.replace(f'!I{old_py_row}', f'!I{new_py_row}')
                    b_cell.value = new_formula


        cy_data_rows = set()
        rate_values_found = sum(
            1 for r in range(data_start, min(ws_cy.max_row + 1, data_start + 40))
            if isinstance(ws_cy_do.cell(r, rt_col).value, (int, float))
            and ws_cy_do.cell(r, rt_col).value in (5, 10, 15, 20, 25, 30, 40, 60, 100)
        )
        if rate_values_found == 0:
            for try_col in range(6, 11):
                cnt = sum(
                    1 for r in range(data_start, min(ws_cy.max_row + 1, data_start + 40))
                    if isinstance(ws_cy_do.cell(r, try_col).value, (int, float))
                    and ws_cy_do.cell(r, try_col).value in (5, 10, 15, 20, 25, 30, 40, 60, 100)
                )
                if cnt >= 2:
                    rt_col = try_col
                    break
        for r in range(data_start, ws_cy.max_row + 1):
            if isinstance(ws_cy.cell(r, rt_col), _MC):
                continue
            rate_v = ws_cy_do.cell(r, rt_col).value
            if not isinstance(rate_v, (int, float)):
                continue
            nm = str(ws_cy.cell(r, 1).value or "").strip().lower()
            if nm not in ("total", "grand total"):
                cy_data_rows.add(r)

        header_skip = {"additions", "greater than", "less than", "sale", "180 days", "amount in rs.", "particulars", "w.d.v", "amount in rs", "180days"}
        cleared = 0
        for r in range(1, ws_cy.max_row + 1):
            is_data_row = r in cy_data_rows
            for col in (ag_col, al_col, sl_col):
                cell = ws_cy.cell(r, col)
                if isinstance(cell, _MC):
                    continue
                v = cell.value
                if v is None:
                    continue
                vs = str(v).strip().lower()
                if vs in header_skip:
                    continue
                if isinstance(v, str) and 'sum' in vs:
                    continue
                if is_data_row:
                    if isinstance(v, (int, float)):
                        cell.value = None
                        cleared += 1
                    elif isinstance(v, str) and v.startswith('='):
                        body = v[1:].strip()
                        import re as _re2
                        is_arith = _re2.sub(r'[\d\s\.\+\-\*\/\(\)]+', '', body) == ''
                        if is_arith:
                            cell.value = None
                            cleared += 1
                else:
                    if isinstance(v, (int, float)):
                        cell.value = None
                        cleared += 1
                    elif isinstance(v, str) and not v.startswith('='):
                        if _re.fullmatch(r'\d{1,2}[./]\d{1,2}[./]\d{2,4}', v.strip()):
                            cell.value = None
                            cleared += 1

        # Keep CY dates aligned to the new year; PY dates come from mirrored source CY
        try:
            new_oy = int(cy_year) - 1
            new_cy = int(cy_year)
            for r in range(1, max(date_row, data_start) + 2):
                for c in range(1, 12):
                    cell = ws_cy.cell(r, c)
                    if isinstance(cell, _MC):
                        continue
                    v = str(cell.value or "").strip()
                    if _re.fullmatch(r"\d{1,2}[./]04[./]\d{4}", v):
                        cell.value = f"01.04.{new_oy}"
                    elif _re.fullmatch(r"\d{1,2}[./]03[./]\d{4}", v):
                        cell.value = f"31.03.{new_cy}"
        except Exception:
            pass

        try:
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            wb.calculation.calcMode = 'auto'
        except Exception:
            pass

        wb.save(output_path)
        log.append(f"✓ FA CY: additions/sale inputs cleared ({cleared} cells)")
        log.append("✓ FA rollover complete")
    finally:
        wb.close()
        wb_do.close()
        if src_wb:
            src_wb.close()
        if src_wb_do:
            src_wb_do.close()


def _inject_cap_fa(output_path, cap_entries, fa_entries, log):
    """Inject user-entered Capital A/c and Fixed Assets values into the output BS."""
    from openpyxl import load_workbook
    from openpyxl.cell import MergedCell
    wb = load_workbook(output_path)

    def _is_formula(v):
        return isinstance(v, str) and v.startswith("=")

    def _safe_write(ws, row, col, value):
        """Write to cell only if it's not merged and not a formula."""
        try:
            cell = ws.cell(row, col)
            if isinstance(cell, MergedCell):
                return False
            if _is_formula(str(cell.value or "")):
                return False
            cell.value = round(float(value), 2)
            return True
        except Exception:
            return False

    # ── Capital Account injection ──────────────────────────────────────
    if cap_entries:
        cap_sheet = None
        for sn in wb.sheetnames:
            if "capital" in sn.lower():
                cap_sheet = sn
                break
        if cap_sheet:
            ws = wb[cap_sheet]
            for entry in cap_entries:
                row = entry.get("row")
                if not row:
                    continue
                # Write each field to its column using column index from entry
                fields = [
                    ("introduced", "Capital Introduced"),
                    ("interest_on_capital", "Interest on Capital"),
                    ("salary", "Salary"),
                    ("withdrawals", "Withdrawals"),
                ]
                for field_key, field_label in fields:
                    val = entry.get(field_key, 0)
                    col_idx = entry.get(f"{field_key}_col")
                    if val and col_idx:
                        if _safe_write(ws, row, col_idx, val):
                            log.append(f"✓ {cap_sheet}!{chr(64+col_idx)}{row} ({field_label}) = {float(val):,.2f}")

    # ── Fixed Assets injection ─────────────────────────────────────────
    if fa_entries:
        fa_sheet = None
        for sn in wb.sheetnames:
            sl = sn.lower()
            if "fixed asset" in sl or "fa " in sl or sl.startswith("fa") or "ppe" in sl:
                fa_sheet = sn
                break
        if fa_sheet:
            ws = wb[fa_sheet]
            for entry in fa_entries:
                row = entry.get("row")
                if not row:
                    continue
                fields = [
                    ("additions_gt180", 3, "Addition >180d"),
                    ("additions_lt180", 4, "Addition <180d"),
                    ("sale", 5, "Sale"),
                ]
                for field_key, default_col, label in fields:
                    val = entry.get(field_key, 0)
                    col_idx = entry.get(f"{field_key}_col", default_col)
                    if val:
                        if _safe_write(ws, row, col_idx, val):
                            log.append(f"✓ {fa_sheet}!{chr(64+col_idx)}{row} ({label}) = {float(val):,.2f}")

    wb.save(output_path)
    wb.close()

@app.route("/tb-process", methods=["POST"])
def tb_process():
    if "uid" not in session:
        return jsonify({"status": "error", "message": "Not logged in"}), 401
    user = get_user_by_id(session["uid"])
    if not user:
        return jsonify({"status": "error", "message": "Not logged in"}), 401
    
    if not TB_PROCESSOR_AVAILABLE:
        return jsonify({"status": "error", "message": "TB processor not available"}), 500
    
    try:
        import tempfile, os, json, shutil

        tb_file = request.files.get("tb_file")
        bs_file = request.files.get("bs_file")

        if not tb_file or not bs_file:
            return jsonify({"status": "error", "message": "Both Trial Balance and BS template files are required"})

        # 🔍 Safely parse user mappings from frontend
        raw_mappings = request.form.get("user_mappings", "{}")
        try:
            user_mapping = json.loads(raw_mappings)
            # Clean keys & values: strip spaces, ignore empty/auto
            user_mapping = {
                str(k).strip(): str(v).strip() 
                for k, v in user_mapping.items() 
                if v and str(v).strip().lower() not in ("", "auto", "none", "ignore")
            }
        except Exception:
            user_mapping = {}

        # Read Capital & FA user entries (if any)
        raw_cap = request.form.get("capital_entries", "[]")
        raw_fa  = request.form.get("fa_entries", "[]")
        try:
            cap_entries = json.loads(raw_cap)
        except Exception:
            cap_entries = []
        try:
            fa_entries = json.loads(raw_fa)
        except Exception:
            fa_entries = []

        client_name = request.form.get("client_name", "Balance_Sheet").strip()
        cy_year = request.form.get("cy_year", "2025").strip()

        tmp = tempfile.mkdtemp()
        tb_orig = tb_file.filename or "tb.xlsx"
        tb_ext = ".pdf" if tb_orig.lower().endswith(".pdf") else ".xlsx"
        tb_path  = os.path.join(tmp, "tb" + tb_ext)
        bs_path  = os.path.join(tmp, "bs_template.xlsx")
        out_path = os.path.join(tmp, "bs_output.xlsx")

        tb_file.save(tb_path)
        bs_file.save(bs_path)

        # Process using the updated tb_processor
        result = process_tb_to_bs(
            tb_path, bs_path, out_path,
            user_mapping=user_mapping,
        )

        if result.get("status") == "error":
            try: shutil.rmtree(tmp, ignore_errors=True)
            except: pass
            return jsonify(result)

        # ── NOTE: FA rollover intentionally NOT run here ────────────────
        # _rollover_fixed_assets() is designed for the YEAR-SHIFT tool,
        # where CY data becomes PY data for a new fiscal year. The TB→BS
        # tool is a different workflow: it fills CY figures into an
        # EXISTING template whose "Fixed Assets P. Yr." sheet already
        # holds last year's correct closing data (e.g. Equipment/Car/
        # Motor Cycle WDV figures). Running the rollover here treated
        # that sheet as if it needed to be "shifted", overwriting its
        # rows/headers with the CY sheet's layout and wiping the real
        # data — causing the Fixed Assets note (and PPE on the BS) to
        # show 0 in the generated output.
        #
        # The BS template's FA C.Yr / FA P.Yr sheets are preserved
        # as-is by process_tb_to_bs() above; no further FA processing
        # is needed for this tool.

        # ── Inject Capital & FA user entries ──────────────────────────────
        if cap_entries or fa_entries:
            _inject_cap_fa(out_path, cap_entries, fa_entries, result.get("log", []))

        h = os.urandom(16).hex()
        dest = os.path.join(OUTPUT_DIR, h + "_out.xlsx")
        shutil.move(out_path, dest)

        try: shutil.rmtree(tmp, ignore_errors=True)
        except: pass

        safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in client_name)
        fname = f"{safe_name}_BS_{cy_year}.xlsx"
        log_usage(user["id"], fname)

        # Extract aggregated P&L figures for UI success page
        aggregated_vals = result.get("aggregated", {}) or {}
        revenue_val        = result.get("revenue", aggregated_vals.get("revenue", 0))
        other_income_val   = result.get("other_income", aggregated_vals.get("other_income", 0))
        direct_expenses_val= aggregated_vals.get("direct_expenses", 0)
        opening_stock_val  = result.get("opening_stock", aggregated_vals.get("opening_stock", 0))
        closing_stock_val  = result.get("closing_stock", aggregated_vals.get("inventories", 0))
        purchases_val      = aggregated_vals.get("purchases", 0)
        employee_exp_val   = aggregated_vals.get("employee_expenses", 0)
        other_exp_val      = aggregated_vals.get("other_expenses", 0)
        depreciation_val   = aggregated_vals.get("depreciation", 0)
        finance_cost_val   = aggregated_vals.get("finance_cost", 0)
        tax_expense_val    = aggregated_vals.get("tax_expense", 0)

        total_assets_val   = result.get("total_assets", 0)
        total_liab_val     = result.get("total_liabilities", 0)
        net_profit_val     = result.get("net_profit", 0)
        diff_val           = abs(float(total_assets_val) - float(total_liab_val))

        return jsonify({
            "status":    "success",
            "log":      result.get("log", []),
            "file_id":  h,
            "filename": fname,
            "aggregated":      aggregated_vals,
            "revenue":         revenue_val,
            "other_income":    other_income_val,
            "direct_expenses": direct_expenses_val,
            "opening_stock":   opening_stock_val,
            "closing_stock":   closing_stock_val,
            "purchases":       purchases_val,
            "employee_expenses": employee_exp_val,
            "other_expenses":  other_exp_val,
            "depreciation":    depreciation_val,
            "finance_cost":    finance_cost_val,
            "tax_expense":     tax_expense_val,
            "tally": {
                "balanced":       bool(result.get("tally_ok", False)),
                "total_assets":   total_assets_val,
                "total_liabilities": total_liab_val,
                "difference":     diff_val,
                "profit":         net_profit_val,
                "user_mappings_applied": len(user_mapping or {}),
            },
        })

    except Exception as e:
        import traceback
        return jsonify({
            "status":   "error",
            "message": f"Processing failed: {e}\n{traceback.format_exc()}"
        }), 500


# ── TB→BS Page Template ───────────────────────────────────────────────────────
TB_BS_TEMPLATE = """<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Balance Sheet from Trial Balance – CA Toolkit</title>
<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap\"/>
<style>
""" + BASE_CSS + """
nav{background:#fff;border-bottom:1px solid var(--border);padding:0 24px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.hero{background:linear-gradient(135deg,#0F172A,#1E3A5F);color:#fff;padding:40px 24px 32px;text-align:center}
.hero h1{font-size:clamp(22px,4vw,32px);font-weight:800;margin-bottom:8px}
.hero p{color:#94A3B8;font-size:14px;max-width:600px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.15);border-radius:20px;padding:4px 14px;font-size:11px;font-weight:600;color:#CBD5E1;margin-bottom:14px}
.page{max-width:1000px;margin:0 auto;padding:24px 16px 60px}
.card{background:#fff;border:1px solid var(--border);border-radius:14px;box-shadow:0 2px 8px rgba(0,0,0,.05);margin-bottom:20px;overflow:hidden}
.card-head{display:flex;align-items:center;gap:14px;padding:16px 20px;border-bottom:1px solid var(--border);background:#FAFAFA}
.card-head .icon{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}
.card-head h2{font-size:15px;font-weight:700;margin:0}
.card-head p{font-size:12px;color:var(--muted);margin:2px 0 0}
.card-body{padding:20px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.field label{display:block;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-bottom:6px}
.upload-zone{border:2px dashed var(--border);border-radius:10px;padding:24px 20px;text-align:center;cursor:pointer;transition:all .2s;background:#FAFAFA;position:relative;min-height:90px}
.upload-zone:hover,.upload-zone.drag{border-color:var(--brand);background:#EFF6FF}
.upload-zone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.upload-zone .uzicon{font-size:26px;margin-bottom:6px}
.upload-zone .uztitle{font-size:13px;font-weight:600;color:var(--ink)}
.upload-zone .uzsub{font-size:11px;color:var(--muted);margin-top:3px}
.uz-done{display:none;margin-top:8px;padding:6px 12px;background:#ECFDF5;border-radius:6px;font-size:11px;font-weight:700;color:#065F46}
select,input[type=text]{width:100%;border:1.5px solid var(--border);border-radius:8px;padding:8px 10px;font-family:inherit;font-size:13px;box-sizing:border-box}
select:focus,input:focus{outline:none;border-color:var(--brand)}
.btn-main{width:100%;padding:14px;background:var(--brand);color:#fff;border:none;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit;transition:all .2s}
.btn-main:hover{background:#1D4ED8}
.btn-main:disabled{background:#93C5FD;cursor:not-allowed}
.btn-sec{padding:10px 20px;background:#F3F4F6;color:var(--ink);border:1.5px solid var(--border);border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit}
.btn-sec:hover{background:#E5E7EB}

/* Steps */
.steps{display:flex;margin-bottom:20px;border-radius:10px;overflow:hidden;border:1px solid var(--border)}
.step-item{flex:1;padding:10px 8px;text-align:center;font-size:11px;font-weight:600;color:var(--muted);background:#F9FAFB;border-right:1px solid var(--border);transition:all .2s}
.step-item:last-child{border-right:none}
.step-item.active{background:var(--brand);color:#fff}
.step-item.done{background:#ECFDF5;color:#065F46}
.step-num{display:block;font-size:15px;margin-bottom:2px}

/* Mapping table */
.map-table{width:100%;border-collapse:collapse;font-size:12px}
.map-table th{padding:8px 10px;border-bottom:2px solid var(--border);font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);background:#F9FAFB;text-align:left}
.map-table td{padding:7px 10px;border-bottom:1px solid var(--border);vertical-align:middle}
.map-table tr:hover td{background:#F9FAFB}
.acc-name{font-weight:600;color:var(--ink);font-size:12px}
.acc-grp{font-size:10px;color:var(--muted)}
.amt{font-weight:700;text-align:right;white-space:nowrap;font-size:12px}
.amt.cr{color:var(--green)}
.amt.dr{color:#2563EB}
.conf-pill{display:inline-flex;align-items:center;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;white-space:nowrap}
.conf-high{background:#ECFDF5;color:#065F46}
.conf-med{background:#FFFBEB;color:#92400E}
.conf-low{background:#FEF2F2;color:#B91C1C}
.conf-user{background:#EFF6FF;color:#1E40AF}
.map-sel{width:100%;border:1.5px solid var(--border);border-radius:6px;padding:5px 7px;font-size:11px;font-family:inherit;background:#fff;cursor:pointer}
.map-sel:focus{border-color:var(--brand);outline:none}
.map-sel.changed{border-color:#F59E0B;background:#FFFBEB;font-weight:700}
.map-sel.user{border-color:var(--brand);background:#EFF6FF;font-weight:700}

/* Summary */
.sum-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:16px}
.sum-card{padding:10px;background:#F9FAFB;border:1px solid var(--border);border-radius:10px;text-align:center}
.sum-val{font-size:18px;font-weight:800;color:var(--brand)}
.sum-lbl{font-size:10px;color:var(--muted);font-weight:600;text-transform:uppercase;margin-top:2px}

/* Spinner */
.spinner{width:44px;height:44px;border:4px solid #E5E7EB;border-top-color:var(--brand);border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 14px}
@keyframes spin{to{transform:rotate(360deg)}}

/* Result */
.result-ok{padding:16px;background:#F0FDF4;border:1.5px solid #BBF7D0;border-radius:10px;text-align:center}
.result-err{padding:16px;background:#FEF2F2;border:1.5px solid #FECACA;border-radius:10px;text-align:center}
.trow{display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid rgba(0,0,0,.05);font-size:13px}
.trow:last-child{border:none}
.tlbl{color:var(--muted);font-weight:500}
.tval{font-weight:700}
.note-box{font-size:11px;color:var(--muted);line-height:1.7;padding:10px 14px;background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;margin-top:12px}
footer{background:#0f1b2d;color:#9CA3AF;font-size:12px;padding:0}
.ft-main{display:grid;grid-template-columns:2fr 1fr 1.4fr;gap:40px;padding:40px 48px;max-width:1200px;margin:0 auto}
.ft-brand-name{color:#fff;font-size:18px;font-weight:800;margin-bottom:12px}
.ft-brand-desc{font-size:12.5px;line-height:1.75;color:#9CA3AF;max-width:340px;text-align:justify}
.ft-col-title{color:#fff;font-size:14px;font-weight:700;margin-bottom:14px}
.ft-links{list-style:none;padding:0;margin:0}
.ft-links li{margin-bottom:8px}
.ft-links a{color:#9CA3AF;text-decoration:none;font-size:13px;transition:color .2s}
.ft-links a:hover{color:#fff}
.ft-contact-name{color:#fff;font-weight:700;font-size:13px;margin-bottom:6px}
.ft-contact-addr{color:#9CA3AF;font-size:12px;line-height:1.7;margin-bottom:10px}
.ft-contact-line{color:#9CA3AF;font-size:12px;margin-bottom:4px}
.ft-socials{display:flex;gap:14px;margin-top:12px}
.ft-socials a{color:#9CA3AF;transition:color .2s}
.ft-socials a:hover{color:#fff}
.ft-socials svg{width:20px;height:20px;fill:currentColor}
.ft-bottom{background:#0a1422;border-top:1px solid #1e2d42;padding:12px 48px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.ft-bottom-left{font-size:11px;color:#6B7280}
.ft-bottom-right{font-size:11px;color:#6B7280}
@media(max-width:768px){.ft-main{grid-template-columns:1fr;padding:28px 20px;gap:24px}.ft-bottom{padding:12px 20px;flex-direction:column;text-align:center}}
@media(max-width:600px){.row2{grid-template-columns:1fr}.steps{flex-direction:column}}
</style></head><body>
<nav>
  <a href="/" class="logo">CA<span>Toolkit</span></a>
  <div class="nav-right">
    <span class="nav-user">👤 <strong>{{ username }}</strong>
      <span class="badge b-{{ plan }}">{{ plan_label }}</span></span>
    <a href="/" class="nav-btn" style="background:#F3F4F6;color:var(--ink)">← Dashboard</a>
    <a href="/logout" class="nav-link">Sign out</a>
  </div>
</nav>
<section class="hero">
  <div class="hero-badge">📋 Premium · Zero Formatting Change</div>
  <h1>Balance Sheet from Trial Balance</h1>
  <p>Upload trial balance + BS template. Auto-maps accounts, lets you correct, then injects CY figures.</p>
</section>

<div class="page">
  <div class="steps">
    <div class="step-item active" id="s1"><span class="step-num">1</span>Upload</div>
    <div class="step-item" id="s2b"><span class="step-num">2</span>Capital &amp; FA</div>
    <div class="step-item" id="s2"><span class="step-num">3</span>Review Mapping</div>
    <div class="step-item" id="s3"><span class="step-num">4</span>Download</div>
  </div>

  <!-- STEP 1 -->
  <div id="step1">
    <div class="card">
      <div class="card-head"><div class="icon" style="background:#EFF6FF">📤</div>
        <div><h2>Upload Files</h2><p>Trial Balance (.xlsx or .pdf) and Balance Sheet template (.xlsx)</p></div></div>
      <div class="card-body">
        <div class="row2">
          <div class="field">
            <label>Trial Balance</label>
            <div class="upload-zone" id="tbZone">
              <input type="file" id="tbFile" accept=".xlsx,.xls,.pdf" onchange="onFile(this,'tb')"/>
              <div class="uzicon">📊</div>
              <div class="uztitle">Click or drag Trial Balance</div>
              <div class="uzsub">Tally / Busy / Manual — .xlsx or .pdf</div>
            </div>
            <div class="uz-done" id="tbDone"></div>
          </div>
          <div class="field">
            <label>Balance Sheet Template</label>
            <div class="upload-zone" id="bsZone">
              <input type="file" id="bsFile" accept=".xlsx" onchange="onFile(this,'bs')"/>
              <div class="uzicon">📋</div>
              <div class="uztitle">Click or drag BS Template</div>
              <div class="uzsub">PY filled · CY blank · formatting intact</div>
            </div>
            <div class="uz-done" id="bsDone"></div>
          </div>
        </div>
        <div class="row2" style="margin-top:14px">
          <div class="field">
            <label>Financial Year (CY)</label>
            <select id="cyYear">
              <option value="2025">2024-25 (31 March 2025)</option>
              <option value="2026">2025-26 (31 March 2026)</option>
            </select>
          </div>
          <div class="field">
            <label>Client / Firm Name</label>
            <input type="text" id="clientName" placeholder="XYZ Enterprises..."/>
          </div>
        </div>
        <div style="margin-top:18px">
          <button class="btn-main" id="analyseBtn" onclick="doAnalyse()" disabled>
            🔍 Analyse Trial Balance
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- STEP 2 (now Capital & FA — shown first after upload) -->
  <div id="step2b" style="display:none">
    <div class="card">
      <div class="card-head"><div class="icon" style="background:#FEF3C7">📋</div>
        <div><h2>Capital Account &amp; Fixed Assets</h2>
        <p>Enter additions, withdrawals (capital) and additions, sales (fixed assets) from ledger.</p></div></div>
      <div class="card-body">

        <div style="background:#FEF3C7;border:1px solid #FDE68A;border-radius:10px;padding:12px 16px;font-size:12px;color:#92400E;margin-bottom:16px;line-height:1.7">
          <strong>Why this step?</strong> The Trial Balance only has closing balances. Capital A/c needs opening + additions + withdrawals from the ledger. Same for Fixed Assets — additions &amp; sales come from ledger, not TB.
        </div>

        <div style="margin-bottom:24px">
          <h3 style="font-size:14px;font-weight:700;margin-bottom:10px">👤 Owner's Capital Account</h3>
          <div id="capTableWrap" style="overflow-x:auto">
            <p style="color:var(--muted);font-size:12px">Loading from BS template...</p>
          </div>
        </div>

        <div style="margin-bottom:20px">
          <h3 style="font-size:14px;font-weight:700;margin-bottom:10px;display:flex;align-items:center;gap:8px">
            🏭 Fixed Assets Chart
            <button onclick="addFARow()" style="margin-left:auto;background:none;border:1px dashed var(--border);border-radius:6px;padding:4px 12px;font-size:11px;cursor:pointer;color:var(--brand)">+ Add Asset</button>
          </h3>
          <div id="faTableWrap" style="overflow-x:auto">
            <p style="color:var(--muted);font-size:12px">Loading from BS template...</p>
          </div>
        </div>

        <div style="display:flex;gap:12px">
          <button class="btn-sec" onclick="goStep(1)">← Back</button>
          <button class="btn-main" style="flex:1" onclick="goStep(2)">
            Next → Review Mapping
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- STEP 3 (Review Mapping — two-panel BS | P&L) -->
  <div id="step2" style="display:none">
    <div class="card">
      <div class="card-head"><div class="icon" style="background:#EFF6FF">🗂️</div>
        <div><h2>Review &amp; Confirm Account Mapping</h2>
          <p id="mapSub">Verify auto-detected heads — change any using the dropdown</p></div></div>
      <div class="card-body">
        <div id="tbFormatBox" style="padding:10px 14px;background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;font-size:12px;color:#1E40AF;margin-bottom:14px"></div>
        <div class="sum-grid" id="sumGrid"></div>
        <div id="preChecks"></div>

        <div style="margin-bottom:10px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px">
          <div style="font-size:12px;color:var(--muted)">
            🟢 Auto &nbsp;|&nbsp; 🟡 Review &nbsp;|&nbsp; 🔴 Manual &nbsp;|&nbsp; 🔵 Changed
          </div>
          <button class="btn-sec" onclick="expandAll()">Expand All Groups</button>
        </div>

        <!-- Tab-based BS | P&L layout -->
        <div style="display:flex;gap:0;margin-bottom:16px">
          <button id="tabBS" onclick="switchTab('bs')" style="flex:1;padding:12px;font-size:14px;font-weight:700;border:2px solid var(--brand);border-radius:10px 0 0 10px;cursor:pointer;background:var(--brand);color:#fff;transition:all .2s">📊 Balance Sheet</button>
          <button id="tabPL" onclick="switchTab('pl')" style="flex:1;padding:12px;font-size:14px;font-weight:700;border:2px solid var(--brand);border-left:none;border-radius:0 10px 10px 0;cursor:pointer;background:#fff;color:var(--brand);transition:all .2s">📈 Profit &amp; Loss</button>
        </div>
        <div id="bsPanel" style="max-height:65vh;overflow-y:auto"></div>
        <div id="plPanel" style="max-height:65vh;overflow-y:auto;display:none"></div>

        <div style="margin-top:20px;display:flex;gap:12px">
          <button class="btn-sec" onclick="goStep('2b')">← Back</button>
          <button class="btn-main" id="generateBtn" onclick="doGenerate()" style="flex:1">
            ✅ Generate Balance Sheet
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- PROCESSING -->
  <div id="loadWrap" style="display:none">
    <div class="card"><div class="card-body" style="text-align:center;padding:48px">
      <div class="spinner"></div>
      <div style="font-size:15px;font-weight:700">Generating Balance Sheet...</div>
      <div style="font-size:12px;color:var(--muted);margin-top:6px" id="loadMsg">Applying your mapping and injecting figures...</div>
    </div></div>
  </div>

  <!-- STEP 3 -->
  <div id="step3" style="display:none">
    <div class="card">
      <div class="card-head"><div class="icon" style="background:#ECFDF5">✅</div>
        <div><h2>Balance Sheet Ready</h2><p id="resSub"></p></div></div>
      <div class="card-body">
        <div id="resBox"></div>
        <div style="margin-top:18px;display:flex;gap:12px">
          <button class="btn-sec" onclick="goStep(2)">← Back to Mapping</button>
          <a id="dlBtn" class="btn-main" style="flex:1;text-decoration:none;display:flex;align-items:center;justify-content:center;gap:8px" href="#">
            📥 Download Balance Sheet
          </a>
        </div>
        <button class="btn-sec" style="width:100%;margin-top:10px" onclick="location.reload()">🔄 New Client</button>
      </div>
    </div>
    <div class="note-box">⚠️ <strong>Always verify:</strong> Total Assets = Total Liabilities · All figures match TB · Profit matches capital account · Notes sheets populated correctly.</div>
  </div>
</div>

<footer>
  <div class="ft-main">
    <div>
      <div class="ft-brand-name">CA Toolkit</div>
      <p class="ft-brand-desc">CA Toolkit is a comprehensive utility platform built by a CA Article from Ludhiana, Punjab, providing automation tools for Indian Chartered Accountants. The platform saves hours of manual work every year — from Balance Sheet year-shift to GST reconciliation, tax calculations, and more.</p>
    </div>
    <div>
      <div class="ft-col-title">Know More</div>
      <ul class="ft-links">
        <li><a href="/">Home</a></li>
        <li><a href="/">BS Year Shift</a></li>
        <li><a href="/tool/tb-to-bs">TB → Balance Sheet</a></li>
        <li><a href="/tool/tax-calculator">Tax Calculator</a></li>
        <li><a href="/privacy">Privacy Policy</a></li>
      </ul>
    </div>
    <div>
      <div class="ft-col-title">Contact Us</div>
      <div class="ft-contact-name">CA Toolkit</div>
      <div class="ft-contact-addr">Built for Indian Chartered Accountants<br/>Created by CA Article · Ludhiana, Punjab</div>
      <div class="ft-contact-line">Support · <a href="https://wa.me/918427651580" style="color:#9CA3AF">WhatsApp Chat</a></div>
      <div class="ft-socials">
        <a href="https://wa.me/918427651580" target="_blank" title="WhatsApp"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>
      </div>
    </div>
  </div>
  <div class="ft-bottom">
    <span class="ft-bottom-left">©2026 CA Toolkit · All Rights Reserved · <a href="/privacy" style="color:#6B7280;text-decoration:none">Privacy Policy</a> · <span style="color:#EF4444">No refund after first upload is used</span></span>
    <span class="ft-bottom-right">Built for Indian CAs · Ludhiana, Punjab</span>
  </div>
</footer>

<script>
// ═══════════════════════════════════════
//  STATE
// ═══════════════════════════════════════
let tbFile = null, bsFile = null;
let analysisData = null;
// KEY STORAGE: maps account unique key → currently selected bs_head
// This is what gets sent to the server on Generate
let userMappings = {};

const BS_HEADS = [
  {v:"capital",       l:"Owner's Capital / Partners Capital"},
  {v:"lt_borrowings", l:"Long Term Borrowings"},
  {v:"st_borrowings", l:"Short Term Borrowings"},
  {v:"trade_payables",l:"Trade Payables (Creditors)"},
  {v:"advance_from_customer", l:"Advance from Customer (under Sundry Creditors)"},
  {v:"other_cl",      l:"Other Current Liabilities"},
  {v:"st_provisions", l:"Short Term Provisions"},
  {v:"fixed_assets",  l:"Fixed Assets / PPE"},
  {v:"investments",   l:"Non-Current Investments"},
  {v:"inventories",   l:"Closing Stock / Inventories"},
  {v:"trade_rec",     l:"Trade Receivables (Debtors)"},
  {v:"advance_to_supplier", l:"Advance to Supplier / Customer (under Sundry Debtors)"},
  {v:"cash_bank",     l:"Cash and Bank Balances"},
  {v:"stla",          l:"Short Term Loans & Advances"},
  {v:"other_ca",      l:"Other Current Assets"},
  {v:"revenue",       l:"Revenue from Operations"},
  {v:"other_income",  l:"Other Income"},
  {v:"opening_stock", l:"Opening Stock"},
  {v:"purchases",     l:"Purchases"},
  {v:"direct_expenses",l:"Direct Expenses"},
  {v:"employee_expenses",l:"Employee / Salary Expenses"},
  {v:"finance_cost",  l:"Finance Cost / Bank Interest"},
  {v:"depreciation",  l:"Depreciation"},
  {v:"other_expenses",l:"Other Expenses"},
  {v:"tax_expense",   l:"Tax Expense"},
  {v:"ignore",        l:"⊘ Ignore / Skip"},
];

const HEAD_LABEL = Object.fromEntries(BS_HEADS.map(h=>[h.v, h.l]));

// ═══════════════════════════════════════
//  STEPS
// ═══════════════════════════════════════
function goStep(n) {
  document.getElementById('step1').style.display = n===1?'block':'none';
  document.getElementById('step2b').style.display = n==='2b'?'block':'none';
  document.getElementById('step2').style.display = n===2?'block':'none';
  document.getElementById('step3').style.display = n===3?'block':'none';
  document.getElementById('loadWrap').style.display = 'none';
  // Step order: s1=1, s2b=2, s2=3, s3=4
  const order = {s1:1, s2b:2, s2:3, s3:4};
  const curVal = n===1?1 : n==='2b'?2 : n===2?3 : n===3?4 : 0;
  ['s1','s2b','s2','s3'].forEach(id=>{
    const s = document.getElementById(id);
    s.className = 'step-item' + (order[id]===curVal?' active':(order[id]<curVal?' done':''));
  });
  window.scrollTo({top:0,behavior:'smooth'});
}

// ═══════════════════════════════════════
//  FILE UPLOAD + DRAG & DROP
// ═══════════════════════════════════════
function onFile(inp, type) {
  const f = inp.files[0]; if(!f) return;
  _setFile(f, type);
}

function _setFile(f, type) {
  if (type==='tb') {
    tbFile = f;
    document.getElementById('tbDone').style.display='block';
    document.getElementById('tbDone').textContent='✓ '+f.name;
    document.getElementById('tbZone').style.borderColor='var(--green)';
  } else {
    bsFile = f;
    document.getElementById('bsDone').style.display='block';
    document.getElementById('bsDone').textContent='✓ '+f.name;
    document.getElementById('bsZone').style.borderColor='var(--green)';
  }
  document.getElementById('analyseBtn').disabled = !(tbFile && bsFile);
}

// Drag-and-drop for both upload zones
['tbZone','bsZone'].forEach(id => {
  const zone = document.getElementById(id);
  if (!zone) return;
  const type = id === 'tbZone' ? 'tb' : 'bs';
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag');
    const f = e.dataTransfer.files[0];
    if (f && (f.name.endsWith('.xlsx') || f.name.endsWith('.xls') || (type==='tb' && f.name.endsWith('.pdf')))) {
      _setFile(f, type);
      // Update the hidden input too
      const dt = new DataTransfer(); dt.items.add(f);
      zone.querySelector('input[type=file]').files = dt.files;
    }
  });
});

// ═══════════════════════════════════════
//  ANALYSE
// ═══════════════════════════════════════
async function doAnalyse() {
  const btn = document.getElementById('analyseBtn');
  btn.disabled = true; btn.textContent = '⏳ Analysing...';
  document.getElementById('step1').style.display='none';
  document.getElementById('loadWrap').style.display='block';
  document.getElementById('loadMsg').textContent = 'Reading trial balance and auto-classifying accounts...';

  const fd = new FormData();
  fd.append('tb_file', tbFile);

  try {
    const res = await fetch('/tb-analyse', {method:'POST', body:fd, credentials:'include'});
    const data = await res.json();

    if (data.status !== 'success') {
      document.getElementById('loadWrap').style.display='none';
      document.getElementById('step1').style.display='block';
      btn.disabled=false; btn.textContent='🔍 Analyse Trial Balance';
      if (res.status === 401) {
        // Session expired — redirect to login
        if (confirm('Session expired. Click OK to go to login page.')) {
          window.location.href = '/login?next=' + encodeURIComponent(window.location.pathname);
        }
      } else {
        alert('Error: ' + data.message);
      }
      return;
    }

    analysisData = data;
    userMappings = {};
    (data.accounts || []).forEach(a => {
      userMappings[a.key] = a.bs_head || 'ignore';
    });

    // Restore any saved mappings from sessionStorage (survives refresh)
    try {
      const saved = JSON.parse(sessionStorage.getItem('tb_mappings') || '{}');
      Object.keys(saved).forEach(k => { if (saved[k] && userMappings.hasOwnProperty(k)) userMappings[k] = saved[k]; });
    } catch(e) {}

    buildMappingUI(data);
    // Go to Capital & FA step first, then user proceeds to mapping
    goToCapFA();

  } catch(e) {
    alert('Network error: '+e);
    document.getElementById('step1').style.display='block';
    document.getElementById('loadWrap').style.display='none';
    btn.disabled=false; btn.textContent='🔍 Analyse Trial Balance';
  }
}

// ═══════════════════════════════════════
//  BUILD MAPPING UI
// ═══════════════════════════════════════
function buildMappingUI(data) {
  const accts = data.accounts || [];
  const fi = data.detection || {};
  const colLetter = (idx) => idx != null && idx >= 0 ? String.fromCharCode(65 + idx) : null;
  const drL = colLetter(fi.debit_col);
  const crL = colLetter(fi.credit_col);
  const netL = colLetter(fi.net_col);
  let colInfo = '';
  if (drL && crL) { colInfo = `Dr: <strong>${drL}</strong> | Cr: <strong>${crL}</strong>`; }
  else if (netL) { colInfo = `Amount: <strong>${netL}</strong>`; }
  else { colInfo = `Amounts: <strong>auto</strong>`; }

  document.getElementById('tbFormatBox').innerHTML =
    `<strong>📊 Detected:</strong> ${fi.format_type ? 'Format '+fi.format_type : 'Auto'} &nbsp;|&nbsp; ` +
    `Name col: <strong>${colLetter(fi.account_col)||'A'}</strong> &nbsp;|&nbsp; ` +
    colInfo + ` &nbsp;|&nbsp; ` +
    `<strong>${accts.length}</strong> accounts`;

  const hi = accts.filter(a=>a.confidence==='high').length;
  const me = accts.filter(a=>a.confidence==='med').length;
  const lo = accts.filter(a=>a.confidence==='low').length;
  document.getElementById('sumGrid').innerHTML = `
    <div class="sum-card"><div class="sum-val">${accts.length}</div><div class="sum-lbl">Total</div></div>
    <div class="sum-card"><div class="sum-val" style="color:var(--green)">${hi}</div><div class="sum-lbl">Auto ✅</div></div>
    <div class="sum-card"><div class="sum-val" style="color:#F59E0B">${me}</div><div class="sum-lbl">Review ⚠️</div></div>
    <div class="sum-card"><div class="sum-val" style="color:#EF4444">${lo}</div><div class="sum-lbl">Manual ❌</div></div>`;

  const checks = data.pre_checks || [];
  document.getElementById('preChecks').innerHTML = checks.map(c=>
    `<div style="padding:6px 12px;border-radius:6px;font-size:12px;margin-bottom:6px;
      background:${c.ok?'#F0FDF4':'#FFFBEB'};color:${c.ok?'#065F46':'#92400E'};
      border:1px solid ${c.ok?'#BBF7D0':'#FDE68A'}">${c.ok?'✅':'⚠️'} ${c.message}</div>`
  ).join('');

  document.getElementById('mapSub').textContent =
    `${accts.length} accounts · ${hi} auto-mapped · ${me+lo} need review`;

  rebuildPanels();
}

// BS heads go in left panel, P&L heads in right panel
const BS_HEAD_KEYS = ['capital','lt_borrowings','st_borrowings',
  'trade_payables','advance_from_customer','other_cl',
  'st_provisions','fixed_assets','investments','inventories',
  'trade_rec','advance_to_supplier','cash_bank','stla','other_ca'];
const PL_HEAD_KEYS = ['revenue','other_income','opening_stock','purchases','direct_expenses',
  'employee_expenses','finance_cost','depreciation','other_expenses','tax_expense'];

function rebuildPanels() {
  const accts = analysisData?.accounts || [];
  // Group by CURRENT userMappings value
  const groups = {};
  accts.forEach(a => {
    const h = userMappings[a.key] || a.bs_head || 'ignore';
    if (!groups[h]) groups[h] = [];
    groups[h].push(a);
  });

  // Low confidence accounts with no user override
  const lowConf = accts.filter(a => a.confidence === 'low' && !(userMappings[a.key] && userMappings[a.key] !== 'ignore'));

  // ── FIX (Issue 1): Use the server-provided manual_review list (full
  // account objects from tb_processor) and split them by suggested_side
  // so the user sees them at the BOTTOM of the BS / P&L panels with a
  // dropdown populated from bs_head_options. Submission still happens
  // through userMappings (keyed by `name_row`).
  const manualReview   = analysisData?.manual_review || [];
  const bsHeadOptions  = analysisData?.bs_head_options || [];
  const manualBS = manualReview.filter(m => (m.suggested_side || 'asset') === 'asset');
  const manualPL = manualReview.filter(m => (m.suggested_side || 'asset') === 'liability');

  let bsHtml = '';
  if (lowConf.length) bsHtml += buildGroup('❌ Needs Manual Mapping', lowConf, true, true);
  BS_HEAD_KEYS.forEach(h => {
    const g = groups[h] || [];
    if (g.length) bsHtml += buildGroup(HEAD_LABEL[h]||h, g, false, false);
  });
  if (groups['ignore']?.length) bsHtml += buildGroup('Ignored', groups['ignore'], false, false);
  // Append server-provided Manual rows at the BOTTOM of BS panel
  if (manualBS.length) bsHtml += buildManualGroup('🔍 Manual Review — Dr Balances (Asset side)', manualBS, bsHeadOptions);
  document.getElementById('bsPanel').innerHTML = bsHtml || '<p style="padding:16px;color:var(--muted);font-size:12px">No BS accounts</p>';

  let plHtml = '';
  PL_HEAD_KEYS.forEach(h => {
    const g = groups[h] || [];
    if (g.length) plHtml += buildGroup(HEAD_LABEL[h]||h, g, false, false);
  });
  // Append server-provided Manual rows at the BOTTOM of P&L panel
  if (manualPL.length) plHtml += buildManualGroup('🔍 Manual Review — Cr Balances (Income / Liability side)', manualPL, bsHeadOptions);
  document.getElementById('plPanel').innerHTML = plHtml || '<p style="padding:16px;color:var(--muted);font-size:12px">No P&L accounts</p>';
}

// ── NEW: Build a group panel for server-provided manual_review items.
// These items only carry {name,row,group,net,dr_cr,bs_head,suggested_side}
// (no `key`), so we look up the matching account in analysisData.accounts
// (which DOES have `key`) to wire the dropdown back into userMappings.
function _findAcctKeyForManual(m) {
  const accts = analysisData?.accounts || [];
  // Primary match: by row number
  if (m.row != null) {
    const byRow = accts.find(a => a.row === m.row && a.name === m.name);
    if (byRow) return byRow.key;
    const byRowOnly = accts.find(a => a.row === m.row);
    if (byRowOnly) return byRowOnly.key;
  }
  // Fallback: by name only
  const byName = accts.find(a => (a.name || '').toUpperCase() === (m.name || '').toUpperCase());
  if (byName) return byName.key;
  // Last resort: synthesize a key in the same format tb_processor uses
  return `${m.name}_${m.row || 0}`;
}

function buildManualGroup(title, manualItems, headOptions) {
  if (!manualItems || !manualItems.length) return '';
  const total = manualItems.reduce((s,m) => s + Math.abs(m.net || 0), 0);
  const opts = (headOptions && headOptions.length)
    ? headOptions
    : BS_HEADS.map(h => ({key: h.v, label: h.l}));

  const rows = manualItems.map(m => {
    const net = m.net || 0;
    const drcr = m.dr_cr || (net < 0 ? 'Cr' : 'Dr');
    const amtCls = drcr === 'Cr' ? 'cr' : 'dr';
    const amtStr = drcr + ' ₹' + Math.abs(net).toLocaleString('en-IN', {maximumFractionDigits:2});
    const key = _findAcctKeyForManual(m);
    const currentHead = userMappings[key] || m.bs_head || 'ignore';
    const selOpts = opts.map(o =>
      `<option value="${escHtml(o.key)}"${currentHead === o.key ? ' selected' : ''}>${escHtml(o.label)}</option>`
    ).join('') + `<option value="ignore"${currentHead === 'ignore' ? ' selected' : ''}>⊘ Ignore / Skip</option>`;

    return `<tr>
      <td><div class="acc-name">${escHtml(m.name)}</div><div class="acc-grp">${escHtml(m.group||'')}</div></td>
      <td class="amt ${amtCls}">${amtStr}</td>
      <td><span class="conf-pill conf-low">❌ Manual</span></td>
      <td><select class="map-sel" data-key="${escHtml(key)}" onchange="onMapChange(this)">
        ${selOpts}
      </select></td>
    </tr>`;
  }).join('');

  const id = 'grp_manual_' + title.replace(/[^a-z0-9]/gi,'_');
  return `<div style="margin-bottom:8px;border-radius:8px;overflow:hidden;border:1px solid #FDE68A">
    <div style="padding:10px 14px;background:#FFFBEB;display:flex;justify-content:space-between;align-items:center;cursor:pointer;user-select:none" onclick="toggleGroup('${id}',this)">
      <span style="font-size:13px;font-weight:700;color:#92400E"><span class="grp-arrow">▼</span> ${escHtml(title)}</span>
      <span style="font-size:12px;font-weight:600;color:var(--ink)">₹${Math.round(total).toLocaleString('en-IN')} <span style="color:var(--muted);font-weight:400">(${manualItems.length})</span></span>
    </div>
    <div id="${id}">
      <table class="map-table">
        <thead><tr><th>Account Name</th><th style="text-align:right">Balance</th><th>Status</th><th>Map To BS Head</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  </div>`;
}

function buildGroup(title, accounts, highlight, startExpanded) {
  const total = accounts.reduce((s,a)=>s+Math.abs(a.net||0),0);
  const rows = accounts.map(a => {
    const net = a.net || 0;
    const amtCls = net < 0 ? 'cr' : 'dr';
    const amtStr = (net<0?'Cr ':'Dr ') + '₹' + Math.abs(net).toLocaleString('en-IN',{maximumFractionDigits:2});
    const conf = a.confidence || 'low';
    const pill = conf==='high' ? '<span class="conf-pill conf-high">✅ Auto</span>'
               : conf==='med'  ? '<span class="conf-pill conf-med">⚠️ Review</span>'
               : conf==='user' ? '<span class="conf-pill conf-user">🔵 User</span>'
               :                 '<span class="conf-pill conf-low">❌ Manual</span>';

    // Use CURRENT mapping from userMappings for selected value
    const currentHead = userMappings[a.key] || a.bs_head || 'ignore';
    const selOpts = BS_HEADS.map(h =>
      `<option value="${h.v}"${currentHead===h.v?' selected':''}>${h.l}</option>`
    ).join('');

    return `<tr>
      <td><div class="acc-name">${escHtml(a.name)}</div><div class="acc-grp">${escHtml(a.group||'')}</div></td>
      <td class="amt ${amtCls}">${amtStr}</td>
      <td>${pill}</td>
      <td><select class="map-sel" data-key="${escHtml(a.key)}" onchange="onMapChange(this)">
        ${selOpts}
      </select></td>
    </tr>`;
  }).join('');

  const bg = highlight ? '#FFFBEB' : '#FAFAFA';
  const border = highlight ? '1px solid #FDE68A' : '1px solid var(--border)';
  const id = 'grp_' + title.replace(/[^a-z0-9]/gi,'_');
  const collapsed = startExpanded ? '' : 'display:none';
  const arrow = startExpanded ? '▼' : '▶';

  return `<div style="margin-bottom:8px;border-radius:8px;overflow:hidden;border:${border}">
    <div style="padding:10px 14px;background:${bg};display:flex;justify-content:space-between;align-items:center;cursor:pointer;user-select:none" onclick="toggleGroup('${id}',this)">
      <span style="font-size:13px;font-weight:700"><span class="grp-arrow">${arrow}</span> ${escHtml(title)}</span>
      <span style="font-size:12px;font-weight:600;color:var(--ink)">₹${Math.round(total).toLocaleString('en-IN')} <span style="color:var(--muted);font-weight:400">(${accounts.length})</span></span>
    </div>
    <div id="${id}" style="${collapsed}">
      <table class="map-table">
        <thead><tr><th>Account Name</th><th style="text-align:right">Balance</th><th>Status</th><th>Map To BS Head</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  </div>`;
}

function toggleGroup(id, headerEl) {
  const el = document.getElementById(id);
  if (!el) return;
  const showing = el.style.display === 'none';
  el.style.display = showing ? '' : 'none';
  // Update arrow
  if (headerEl) {
    const arrow = headerEl.querySelector('.grp-arrow');
    if (arrow) arrow.textContent = showing ? '▼' : '▶';
  }
}
function expandAll() {
  document.querySelectorAll('[id^="grp_"]').forEach(el => { el.style.display = ''; });
  document.querySelectorAll('.grp-arrow').forEach(el => { el.textContent = '▼'; });
}

function switchTab(tab) {
  const bsP = document.getElementById('bsPanel');
  const plP = document.getElementById('plPanel');
  const bsB = document.getElementById('tabBS');
  const plB = document.getElementById('tabPL');
  if (tab === 'bs') {
    bsP.style.display = ''; plP.style.display = 'none';
    bsB.style.background = 'var(--brand)'; bsB.style.color = '#fff';
    plB.style.background = '#fff'; plB.style.color = 'var(--brand)';
  } else {
    bsP.style.display = 'none'; plP.style.display = '';
    plB.style.background = 'var(--brand)'; plB.style.color = '#fff';
    bsB.style.background = '#fff'; bsB.style.color = 'var(--brand)';
  }
}

// ═══════════════════════════════════════
//  KEY FIX: onMapChange stores to userMappings immediately
// ═══════════════════════════════════════
function onMapChange(sel) {
  const key = sel.dataset.key;
  const val = sel.value;
  userMappings[key] = val;
  sel.classList.add('changed');
  try { sessionStorage.setItem('tb_mappings', JSON.stringify(userMappings)); } catch(e) {}
  // Track expanded groups before rebuild
  const expanded = new Set();
  document.querySelectorAll('[id^="grp_"]').forEach(el => {
    if (el.style.display !== 'none') expanded.add(el.id);
  });
  rebuildPanels();
  // Restore expanded state
  expanded.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.style.display = '';
      const hdr = el.previousElementSibling;
      if (hdr) { const a = hdr.querySelector('.grp-arrow'); if (a) a.textContent = '▼'; }
    }
  });
}

// ═══════════════════════════════════════
//  STEP 2.5: CAPITAL & FIXED ASSETS
// ═══════════════════════════════════════
let capData = null, faData = null;

async function goToCapFA() {
  document.getElementById('step1').style.display = 'none';
  document.getElementById('loadWrap').style.display = 'block';
  document.getElementById('loadMsg').textContent = 'Reading Capital Account & Fixed Assets from BS template...';

  try {
    const fd = new FormData();
    fd.append('bs_file', bsFile);
    const res = await fetch('/tb-read-bs', {method:'POST', body:fd, credentials:'include'});
    const data = await res.json();
    document.getElementById('loadWrap').style.display = 'none';

    if (data.status !== 'success') {
      // If reading fails, still proceed — user can fill manually later
      capData = null; faData = null;
      buildCapTable(null);
      buildFATable(null);
      goStep('2b');
      return;
    }

    capData = data.capital;
    faData  = data.fixed_assets;
    buildCapTable(capData);
    buildFATable(faData);
    goStep('2b');

  } catch(e) {
    document.getElementById('loadWrap').style.display = 'none';
    capData = null; faData = null;
    buildCapTable(null);
    buildFATable(null);
    goStep('2b');
  }
}

function buildCapTable(cap) {
  const wrap = document.getElementById('capTableWrap');
  if (!cap || !cap.partners || !cap.partners.length) {
    wrap.innerHTML = '<p style="color:var(--muted);font-size:12px">No capital account sheet found in BS template. You can fill it manually in Excel after download.</p>';
    return;
  }
  // Detect columns from the template data
  const cols = cap.columns || [];
  const hasInterest = cols.some(c => c.key === 'interest_on_capital');
  const hasSalary = cols.some(c => c.key === 'salary');

  let html = '<table style="width:100%;border-collapse:collapse;font-size:12px">';
  html += '<tr style="background:#F1F5F9"><th style="padding:8px;text-align:left;border:1px solid var(--border)">Name</th>';
  html += '<th style="padding:8px;text-align:right;border:1px solid var(--border)">Opening</th>';
  html += '<th style="padding:8px;text-align:center;border:1px solid var(--border)">Capital Introduced</th>';
  if (hasInterest) html += '<th style="padding:8px;text-align:center;border:1px solid var(--border)">Interest on Capital</th>';
  if (hasSalary) html += '<th style="padding:8px;text-align:center;border:1px solid var(--border)">Salary</th>';
  html += '<th style="padding:8px;text-align:center;border:1px solid var(--border)">Withdrawals</th></tr>';

  cap.partners.forEach((p, i) => {
    html += `<tr>
      <td style="padding:8px;border:1px solid var(--border);font-weight:600">${escHtml(p.name)}
        <input type="hidden" class="cap-row" value="${p.row}">
        <input type="hidden" class="cap-cols" value='${JSON.stringify(cap.col_map||{})}'>
      </td>
      <td style="padding:8px;border:1px solid var(--border);text-align:right;color:var(--muted)">₹${(p.opening||0).toLocaleString('en-IN',{maximumFractionDigits:2})}</td>
      <td style="padding:4px;border:1px solid var(--border)"><input type="number" class="cap-intro" data-idx="${i}" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>`;
    if (hasInterest) html += `<td style="padding:4px;border:1px solid var(--border)"><input type="number" class="cap-interest" data-idx="${i}" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>`;
    if (hasSalary) html += `<td style="padding:4px;border:1px solid var(--border)"><input type="number" class="cap-salary" data-idx="${i}" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>`;
    html += `<td style="padding:4px;border:1px solid var(--border)"><input type="number" class="cap-wd" data-idx="${i}" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>
    </tr>`;
  });
  html += '</table>';
  wrap.innerHTML = html;
}

function buildFATable(fa) {
  const wrap = document.getElementById('faTableWrap');
  if (!fa || !fa.assets || !fa.assets.length) {
    wrap.innerHTML = '<p style="color:var(--muted);font-size:12px">No fixed assets sheet found in BS template. You can fill it manually in Excel after download.</p>';
    return;
  }
  let html = '<table id="faTable" style="width:100%;border-collapse:collapse;font-size:12px">';
  html += '<tr style="background:#F1F5F9"><th style="padding:8px;text-align:left;border:1px solid var(--border)">Asset</th>';
  html += '<th style="padding:8px;text-align:right;border:1px solid var(--border)">Opening WDV</th>';
  html += '<th style="padding:8px;text-align:center;border:1px solid var(--border)">Additions &gt;180d</th>';
  html += '<th style="padding:8px;text-align:center;border:1px solid var(--border)">Additions &lt;180d</th>';
  html += '<th style="padding:8px;text-align:center;border:1px solid var(--border)">Sale</th>';
  html += '<th style="padding:8px;text-align:right;border:1px solid var(--border)">Rate %</th></tr>';
  fa.assets.forEach((a, i) => {
    html += buildFARow(a, i);
  });
  html += '</table>';
  wrap.innerHTML = html;
}

function buildFARow(a, i) {
  return `<tr class="fa-row" data-idx="${i}">
    <td style="padding:8px;border:1px solid var(--border);font-weight:600">${escHtml(a.name)}<input type="hidden" class="fa-rownum" value="${a.row}"></td>
    <td style="padding:8px;border:1px solid var(--border);text-align:right;color:var(--muted)">₹${(a.opening_wdv||0).toLocaleString('en-IN',{maximumFractionDigits:2})}</td>
    <td style="padding:4px;border:1px solid var(--border)"><input type="number" class="fa-add-gt" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>
    <td style="padding:4px;border:1px solid var(--border)"><input type="number" class="fa-add-lt" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>
    <td style="padding:4px;border:1px solid var(--border)"><input type="number" class="fa-sale" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>
    <td style="padding:8px;border:1px solid var(--border);text-align:right;color:var(--muted)">${a.rate||0}%</td>
  </tr>`;
}

function addFARow() {
  const tbl = document.getElementById('faTable');
  if (!tbl) return;
  const name = prompt('Asset name:');
  if (!name) return;
  const idx = tbl.querySelectorAll('.fa-row').length;
  const tr = document.createElement('tr');
  tr.className = 'fa-row';
  tr.dataset.idx = idx;
  tr.innerHTML = `
    <td style="padding:8px;border:1px solid var(--border);font-weight:600">${escHtml(name)}<input type="hidden" class="fa-rownum" value="0"></td>
    <td style="padding:8px;border:1px solid var(--border);text-align:right;color:var(--muted)">₹0</td>
    <td style="padding:4px;border:1px solid var(--border)"><input type="number" class="fa-add-gt" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>
    <td style="padding:4px;border:1px solid var(--border)"><input type="number" class="fa-add-lt" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>
    <td style="padding:4px;border:1px solid var(--border)"><input type="number" class="fa-sale" step="0.01" value="0" style="width:100%;padding:6px;border:1.5px solid var(--border);border-radius:6px;text-align:right;font-size:12px"></td>
    <td style="padding:8px;border:1px solid var(--border);text-align:right;color:var(--muted)">—</td>`;
  tbl.appendChild(tr);
}

function collectCapEntries() {
  const entries = [];
  document.querySelectorAll('.cap-row').forEach((el, i) => {
    const row = parseInt(el.value);
    let colMap = {};
    try { colMap = JSON.parse(document.querySelectorAll('.cap-cols')[i]?.value || '{}'); } catch(e) {}
    const intro = parseFloat(document.querySelectorAll('.cap-intro')[i]?.value) || 0;
    const interest = parseFloat(document.querySelectorAll('.cap-interest')[i]?.value) || 0;
    const salary = parseFloat(document.querySelectorAll('.cap-salary')[i]?.value) || 0;
    const wd = parseFloat(document.querySelectorAll('.cap-wd')[i]?.value) || 0;
    if (row && (intro || interest || salary || wd)) {
      entries.push({
        row,
        introduced: intro, introduced_col: colMap.introduced || 4,
        interest_on_capital: interest, interest_on_capital_col: colMap.interest_on_capital || 5,
        salary: salary, salary_col: colMap.salary || 6,
        withdrawals: wd, withdrawals_col: colMap.withdrawals || 7,
      });
    }
  });
  return entries;
}

function collectFAEntries() {
  const entries = [];
  document.querySelectorAll('.fa-row').forEach(tr => {
    const row = parseInt(tr.querySelector('.fa-rownum')?.value) || 0;
    const gt = parseFloat(tr.querySelector('.fa-add-gt')?.value) || 0;
    const lt = parseFloat(tr.querySelector('.fa-add-lt')?.value) || 0;
    const sale = parseFloat(tr.querySelector('.fa-sale')?.value) || 0;
    if (row && (gt || lt || sale)) entries.push({row, additions_gt180: gt, additions_lt180: lt, sale});
  });
  return entries;
}

// ═══════════════════════════════════════
//  GENERATE — sends ALL data to server
// ═══════════════════════════════════════
async function doGenerate() {
  // Collect ALL current dropdown values
  document.querySelectorAll('.map-sel').forEach(sel => {
    const key = sel.dataset.key;
    if (key) userMappings[key] = sel.value;
  });

  document.getElementById('step2').style.display = 'none';
  document.getElementById('loadWrap').style.display = 'block';
  document.getElementById('loadMsg').textContent = 'Applying mapping and injecting figures into Balance Sheet...';

  const fd = new FormData();
  fd.append('tb_file', tbFile);
  fd.append('bs_file', bsFile);
  fd.append('cy_year', document.getElementById('cyYear').value);
  fd.append('client_name', document.getElementById('clientName').value || 'Client');
  fd.append('user_mappings', JSON.stringify(userMappings));
  fd.append('capital_entries', JSON.stringify(collectCapEntries()));
  fd.append('fa_entries', JSON.stringify(collectFAEntries()));

  try {
    const res = await fetch('/tb-process', {method:'POST', body:fd, credentials:'include'});
    const data = await res.json();
    document.getElementById('loadWrap').style.display = 'none';

    if (data.status !== 'success') {
      alert('Error: ' + data.message);
      document.getElementById('step2').style.display = 'block';
      return;
    }

    // Skip tally page — download directly
    const fn = data.filename || 'Balance_Sheet.xlsx';
    const dlUrl = '/download/' + data.file_id + '?fn=' + encodeURIComponent(fn);
    
    // Trigger download
    const a = document.createElement('a');
    a.href = dlUrl; a.download = fn; a.click();

    // ── FIX (Issue 2 & 3): Show P&L summary cards (Revenue, Other Income,
    // Direct Expenses, etc.) on the success page. Server now returns these
    // values from result["aggregated"] via /tb-process JSON.
    const t = data.tally || {};
    const ok = !!t.balanced;
    const fmtMoney = (n) => '₹' + (Math.round(n||0)).toLocaleString('en-IN');
    const plRows = [
      ['Revenue from Operations', data.revenue],
      ['Other Income',            data.other_income],          // NEW Issue 2
      ['Opening Stock',           data.opening_stock],
      ['Purchases',               data.purchases],
      ['Direct Expenses',         data.direct_expenses],       // NEW Issue 3
      ['Employee / Salary Exp.',  data.employee_expenses],
      ['Finance Cost',            data.finance_cost],
      ['Depreciation',            data.depreciation],
      ['Other Expenses',          data.other_expenses],
      ['Tax Expense',             data.tax_expense],
      ['Closing Stock',           data.closing_stock],
    ].filter(r => r[1] != null && Math.abs(r[1]) > 0.005);

    const plCardsHtml = plRows.map(r => `
      <div class="sum-card" style="padding:10px 12px;text-align:left">
        <div class="sum-lbl" style="margin:0 0 4px">${escHtml(r[0])}</div>
        <div class="sum-val" style="font-size:15px">${fmtMoney(r[1])}</div>
      </div>`).join('');

    document.getElementById('resBox').innerHTML = `
      <div style="text-align:center;padding:24px 20px 16px">
        <div style="font-size:48px;margin-bottom:8px">✅</div>
        <h2 style="font-size:20px;font-weight:800;margin-bottom:6px">Balance Sheet Downloaded!</h2>
        <p style="color:var(--muted);font-size:13px;margin-bottom:18px">${escHtml(fn)}</p>
        <a href="${dlUrl}" class="btn-main" style="display:inline-flex;padding:12px 32px;text-decoration:none">
          ⬇ Download Again
        </a>
      </div>

      <div class="${ok?'result-ok':'result-err'}" style="margin-top:6px">
        <div style="font-size:16px;font-weight:800;color:${ok?'#065F46':'#B91C1C'}">
          ${ok?'Balance Sheet Tallied ✅':'Tally Mismatch — Review Needed ⚠️'}
        </div>
      </div>

      <div style="margin-top:14px">
        <div class="trow"><span class="tlbl">Total Assets</span><span class="tval">${fmtMoney(t.total_assets)}</span></div>
        <div class="trow"><span class="tlbl">Total Liabilities + Capital</span><span class="tval">${fmtMoney(t.total_liabilities)}</span></div>
        <div class="trow"><span class="tlbl">Difference</span><span class="tval" style="color:${(t.difference||0)<1?'var(--green)':'#EF4444'}">${fmtMoney(t.difference)}</span></div>
        <div class="trow"><span class="tlbl">Profit / (Loss)</span><span class="tval">${fmtMoney(t.profit)}</span></div>
        <div class="trow"><span class="tlbl">User Mapping Overrides Applied</span><span class="tval" style="color:var(--brand)">${t.user_mappings_applied||0}</span></div>
      </div>

      ${plCardsHtml ? `
      <h3 style="margin:20px 0 8px;font-size:13px;font-weight:700;color:var(--ink);text-transform:uppercase;letter-spacing:.04em">P&amp;L Summary</h3>
      <div class="sum-grid">${plCardsHtml}</div>` : ''}

      ${data.log ? '<div style="margin-top:10px;padding:10px;background:#F9FAFB;border-radius:8px;font-size:10px;color:var(--muted);max-height:120px;overflow-y:auto">'+data.log.slice(-12).map(l=>'<div>'+escHtml(l)+'</div>').join('')+'</div>' : ''}
    `;
    document.getElementById('resSub').textContent = fn;
    goStep(3);

  } catch(e) {
    alert('Error: '+e);
    document.getElementById('step2').style.display = 'block';
    document.getElementById('loadWrap').style.display = 'none';
  }
}

// ═══════════════════════════════════════
//  RESULT
// ═══════════════════════════════════════
function buildResult(data) {
  const t = data.tally || {};
  const ok = t.balanced;
  document.getElementById('resSub').textContent =
    (document.getElementById('clientName').value||'Balance Sheet') + ' · CY figures populated';

  document.getElementById('resBox').innerHTML = `
    <div class="${ok?'result-ok':'result-err'}">
      <div style="font-size:22px;margin-bottom:6px">${ok?'🎉':'⚠️'}</div>
      <div style="font-size:16px;font-weight:800;color:${ok?'#065F46':'#B91C1C'}">${ok?'Balance Sheet Tallied ✅':'Tally Mismatch — Review Needed'}</div>
    </div>
    <div style="margin-top:14px">
      <div class="trow"><span class="tlbl">Total Assets</span><span class="tval">₹${fmt(t.total_assets)}</span></div>
      <div class="trow"><span class="tlbl">Total Liabilities + Capital</span><span class="tval">₹${fmt(t.total_liabilities)}</span></div>
      <div class="trow"><span class="tlbl">Difference</span><span class="tval" style="color:${t.difference<1?'var(--green)':'#EF4444'}">₹${fmt(t.difference)}</span></div>
      <div class="trow"><span class="tlbl">Profit / (Loss)</span><span class="tval">₹${fmt(t.profit)}</span></div>
      <div class="trow"><span class="tlbl">User Mapping Overrides Applied</span><span class="tval" style="color:var(--brand)">${t.user_mappings_applied||0}</span></div>
    </div>
    ${data.log ? '<div style="margin-top:10px;padding:10px;background:#F9FAFB;border-radius:8px;font-size:10px;color:var(--muted);max-height:100px;overflow-y:auto">'+data.log.slice(-10).map(l=>'<div>'+escHtml(l)+'</div>').join('')+'</div>' : ''}`;

  const dlBtn = document.getElementById('dlBtn');
  dlBtn.href = '/download/' + data.file_id + '?fn=' + encodeURIComponent(data.filename);
  dlBtn.setAttribute('download', data.filename);
}

function fmt(n) { return (Math.round(n||0)).toLocaleString('en-IN'); }
function escHtml(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// Drag & drop
['tbZone','bsZone'].forEach(id=>{
  const el=document.getElementById(id);
  el.addEventListener('dragover',e=>{e.preventDefault();el.classList.add('drag')});
  el.addEventListener('dragleave',()=>el.classList.remove('drag'));
  el.addEventListener('drop',e=>{
    e.preventDefault();el.classList.remove('drag');
    const f=e.dataTransfer.files[0]; if(!f) return;
    const type=id==='tbZone'?'tb':'bs';
    onFile({files:[f]},type);
    if(type==='tb') tbFile=f; else bsFile=f;
    document.getElementById('analyseBtn').disabled=!(tbFile&&bsFile);
  });
});
</script><a href="https://wa.me/918427651580" target="_blank" class="wa-float" title="WhatsApp Support"><svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg></a>

<button class="help-btn" onclick="openHelp()" title="How to use this tool">?</button>
<div class="help-overlay" id="helpOverlay">
  <div class="help-modal">
    <div class="help-modal-head"><h3>How to Use — Trial Balance → Balance Sheet</h3><button class="help-close" onclick="closeHelp()">&#10005;</button></div>
    <div class="help-modal-body"><div class="help-step"><div class="help-step-num">1</div><div class="help-step-body"><h4>Upload Trial Balance</h4><p>Upload your Excel trial balance with account names and debit/credit balances.</p></div></div><div class="help-step"><div class="help-step-num">2</div><div class="help-step-body"><h4>Upload BS Template</h4><p>Upload your existing Balance Sheet template with CY column cells ready to fill.</p></div></div><div class="help-step"><div class="help-step-num">3</div><div class="help-step-body"><h4>Enter Details</h4><p>Set client name, financial year, and review auto-mapped accounts.</p></div></div><div class="help-step"><div class="help-step-num">4</div><div class="help-step-body"><h4>Fixed Assets & Capital</h4><p>Enter additions, sales, and capital account movements if prompted.</p></div></div><div class="help-step"><div class="help-step-num">5</div><div class="help-step-body"><h4>Generate & Download</h4><p>Click Generate — CY figures are injected into your BS template. Download instantly.</p></div></div><div class="help-tip">💡 Your BS template's formatting and formulas are never changed — only the CY figures are filled in.</div></div>
  </div>
</div>
<script>function openHelp(){document.getElementById('helpOverlay').classList.add('open')}function closeHelp(){document.getElementById('helpOverlay').classList.remove('open')}document.getElementById('helpOverlay').addEventListener('click',function(e){if(e.target===this)closeHelp()})</script>
</body></html>"""


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
