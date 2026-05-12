"""
BS Annual Updater — Multi-Tool CA Dashboard
Auth + Upload-Based Plans + Admin Panel
"""

import re
import os
import uuid
import sqlite3
import hashlib
import secrets
from datetime import datetime, timedelta
from functools import wraps
from flask import (Flask, request, send_file, jsonify,
                   render_template_string, session, redirect, url_for, g)
from processor import process

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

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
    "starter":  {"label": "Starter",      "uploads": 10,  "price": 399},
    "standard": {"label": "Standard",     "uploads": 25,  "price": 899},
    "pro":      {"label": "Professional", "uploads": 60,  "price": 1799},
    "firm":     {"label": "Firm",         "uploads": 150, "price": 3499},
}

# ══════════════════════════════════════════════════════════════════════════════
#  DATABASE
# ══════════════════════════════════════════════════════════════════════════════

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH, detect_types=sqlite3.PARSE_DECLTYPES)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db: db.close()

def init_db():
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
def get_user_by_name(u): return get_db().execute("SELECT * FROM users WHERE username=?", (u,)).fetchone()
def get_user_by_id(i):   return get_db().execute("SELECT * FROM users WHERE id=?", (i,)).fetchone()
def uploads_remaining(user): return max(0, user["uploads_total"] - user["uploads_used"])

def log_usage(user_id, filename):
    db = get_db()
    db.execute("UPDATE users SET uploads_used=uploads_used+1 WHERE id=?", (user_id,))
    db.execute("INSERT INTO usage_log (user_id,filename,processed_at) VALUES (?,?,?)",
               (user_id, filename, datetime.utcnow().isoformat()))
    db.commit()

def add_uploads(user_id, plan_key):
    user    = get_user_by_id(user_id)
    extra   = PLANS[plan_key]["uploads"]
    rem     = uploads_remaining(user)
    new_tot = user["uploads_used"] + rem + extra
    validity = (datetime.utcnow() + timedelta(days=UPLOAD_VALIDITY_DAYS)).isoformat()
    db = get_db()
    db.execute("UPDATE users SET plan=?,uploads_total=?,validity_end=? WHERE id=?",
               (plan_key, new_tot, validity, user_id))
    db.commit()

def create_user(username, password, plan_key):
    uploads  = PLANS[plan_key]["uploads"]
    validity = None if plan_key == "free" else (datetime.utcnow() + timedelta(days=UPLOAD_VALIDITY_DAYS)).isoformat()
    db = get_db()
    db.execute("""INSERT INTO users
        (username,password,plan,is_admin,uploads_total,uploads_used,validity_end,created_at)
        VALUES (?,?,?,0,?,0,?,?)""",
        (username, _hash(password), plan_key, uploads, validity, datetime.utcnow().isoformat()))
    db.commit()

def del_user(uid):
    db = get_db()
    db.execute("DELETE FROM usage_log WHERE user_id=?", (uid,))
    db.execute("DELETE FROM users WHERE id=?", (uid,))
    db.commit()

def all_users(): return get_db().execute("SELECT * FROM users ORDER BY id").fetchall()

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
footer{background:var(--ink);color:#9CA3AF;text-align:center;padding:24px;font-size:12px}
footer a{color:#6B7280;text-decoration:none}
.footer-brand{color:#D1D5DB;font-weight:700;font-size:14px;margin-bottom:6px}
"""

# ══════════════════════════════════════════════════════════════════════════════
#  LOGIN PAGE
# ══════════════════════════════════════════════════════════════════════════════

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
</div></body></html>"""

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

.tools-grid{max-width:960px;margin:0 auto;padding:0 24px 56px;
            display:grid;grid-template-columns:repeat(2,1fr);gap:20px}
@media(max-width:600px){.tools-grid{grid-template-columns:1fr}}

.tool-card{background:var(--white);border:1.5px solid var(--border);
           border-radius:var(--radius);padding:28px 24px;
           text-decoration:none;color:var(--ink);
           transition:all .2s;position:relative;overflow:hidden;display:block}
.tool-card:hover{border-color:var(--brand);box-shadow:0 8px 32px rgba(29,78,216,.12);transform:translateY(-2px)}
.tool-card.disabled{cursor:default;opacity:.7}
.tool-card.disabled:hover{border-color:var(--border);box-shadow:none;transform:none}

.tool-icon{width:52px;height:52px;border-radius:12px;display:flex;
           align-items:center;justify-content:center;font-size:26px;
           margin-bottom:16px}
.tool-card h2{font-size:16px;font-weight:700;margin-bottom:6px}
.tool-card p{font-size:13px;color:var(--muted);line-height:1.6;margin-bottom:16px}

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
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin Panel</a>{% endif %}
    <a href="/logout" class="nav-link">Sign out</a>
  </div>
</nav>

<div class="hero">
  <div class="hero-badge">🇮🇳 Made for Indian CAs &amp; Accountants</div>
  <h1>Your Complete <em>CA Toolkit</em></h1>
  <p>Professional tools built by CA Articles of GD Singla &amp; Co. — designed to save hours of manual work every year.</p>
</div>

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

<!-- Tools grid -->
<div class="tools-grid">

  <a href="/tool/converter" class="tool-card">
    <div class="tool-icon" style="background:#EFF6FF">📊</div>
    <h2>Balance Sheet Year-Shift</h2>
    <p>Roll over your comparative Excel balance sheet to any new financial year in seconds. Shifts CY→PY, clears CY, restores all formulas and updates every date.</p>
    <span class="tool-tag tag-live">✓ Live</span>
    <div class="arrow">→</div>
  </a>

  <div class="tool-card disabled">
    <div class="tool-icon" style="background:#F0FDF4">📋</div>
    <h2>Balance Sheet from Trial Balance</h2>
    <p>Generate a formatted comparative balance sheet directly from your trial balance data. No manual formatting required.</p>
    <span class="tool-tag tag-soon">🔜 Coming Soon</span>
  </div>

  <a href="/tool/tax-calculator" class="tool-card">
    <div class="tool-icon" style="background:#FFFBEB">🧮</div>
    <h2>Income Tax Calculator</h2>
    <p>Calculate income tax liability under old and new regime for PY 2025-26 (AY 2026-27). Income under 5 heads, TDS/TCS, surcharge &amp; cess — all built in.</p>
    <span class="tool-tag tag-live">✓ Live</span>
    <div class="arrow">→</div>
  </a>

  <div class="tool-card disabled">
    <div class="tool-icon" style="background:#F5F3FF">🚀</div>
    <h2>More Features Coming Soon</h2>
    <p>We're building more tools for Indian CAs. Stay tuned — new utilities will be added regularly based on your feedback.</p>
    <span class="tool-tag tag-soon">Stay Tuned</span>
  </div>

</div>

<footer>
  <p class="footer-brand">CA Toolkit</p>
  <p>Built for Indian Chartered Accountants · Saves hours every year</p>
  <p style="margin-top:6px">Created by CA Articles of GD Singla &amp; Co.</p>
  <p style="margin-top:12px;font-size:11px">© 2026 CA Toolkit · Your data is never stored · <span style="color:#EF4444">No refund after first upload is used</span></p>
</footer>
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
      box-shadow:var(--shadow);overflow:hidden}
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
.plans{max-width:980px;margin:0 auto;
       display:grid;grid-template-columns:repeat(5,1fr);gap:14px}
@media(max-width:900px){.plans{grid-template-columns:repeat(2,1fr)}}
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
    <a href="/logout" class="nav-link">Sign out</a>
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
        <label>Upload Excel File (.xlsx)</label>
        <div class="dropzone" id="dropzone">
          <input type="file" id="xlFile" accept=".xlsx" {{ 'disabled' if uploads_left==0 else '' }}/>
          <div class="dz-icon">📁</div>
          <div class="dz-text"><strong>Click to browse</strong> or drag &amp; drop</div>
          <div class="dz-text" style="margin-top:3px">Only .xlsx · Max 20 MB</div>
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
      <div class="plan-price">₹399</div>
      <div class="plan-uploads">10 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features</li><li>All sheet types</li><li>Up to 20 MB</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
    <div class="plan pop">
      <div class="plan-badge">Most Popular</div>
      <div class="plan-name">Standard</div>
      <div class="plan-price">₹899</div>
      <div class="plan-uploads">25 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features</li><li>Priority support</li><li>Up to 20 MB</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
    <div class="plan">
      <div class="plan-name">Professional</div>
      <div class="plan-price">₹1,799</div>
      <div class="plan-uploads">60 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features</li><li>Priority support</li><li>Up to 20 MB</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
    <div class="plan">
      <div class="plan-name">Firm</div>
      <div class="plan-price">₹3,499</div>
      <div class="plan-uploads">150 uploads</div>
      <div class="plan-validity">3 month validity</div>
      <ul><li>All features</li><li>WhatsApp support</li><li>Up to 20 MB</li></ul>
      <a href="#contact" class="plan-btn">Contact to Buy</a>
    </div>
  </div>
  <p class="no-refund-note">⚠ No refund after first upload is used &nbsp;·&nbsp; Unused uploads stack when you recharge before expiry</p>
</section>

<section class="faq-section" id="faq">
  <h2>Frequently Asked Questions</h2>
  <details><summary>Which Excel formats are supported?</summary>
    <p>.xlsx only (Excel 2007+). Save .xls files as .xlsx first.</p></details>
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
  <p class="footer-brand">CA Toolkit</p>
  <p>Built for Indian Chartered Accountants · Saves hours every year</p>
  <p style="margin-top:6px">Created by CA Articles of GD Singla &amp; Co.</p>
  <p style="margin-top:12px;font-size:11px">© 2026 CA Toolkit · Your data is never stored · <span style="color:#EF4444">No refund after first upload is used</span></p>
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
    const res=await fetch('/process',{method:'POST',body:fd}),data=await res.json();
    if(data.status==='success'){
      const logHtml='<ul class="log-list">'+data.log.map(l=>`<li>${l}</li>`).join('')+'</ul>';
      showStatus('success','✓ Done! Your file is ready.'+logHtml);
      dl.href='/download/'+data.file_id;dl.download=data.filename;
      dl.textContent='⬇  Download — '+data.filename;dl.style.display='block';
      toast('Processed successfully!');
      setTimeout(()=>location.reload(),3000);
    }else{showStatus('error','✗ '+data.message);}
  }catch(e){showStatus('error','✗ Network error: '+e.message);}
  finally{btn.disabled=false;sp.style.display='none';bt.textContent='⚡ Process & Download';}
}
function showStatus(t,m){const e=document.getElementById('status');e.className=t;e.innerHTML=m;e.style.display=m?'block':'none';}
function toast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),3000);}
</script>
</body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
#  INCOME TAX CALCULATOR — PY 2025-26 / AY 2026-27
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

/* ── Regime Toggle ────────────────────────────────────────── */
.regime-toggle{display:flex;justify-content:center;gap:12px;margin-bottom:28px;flex-wrap:wrap}
.regime-btn{padding:10px 28px;border-radius:10px;border:2px solid var(--border);
            background:var(--white);font-family:inherit;font-size:13px;font-weight:700;
            cursor:pointer;transition:all .2s;color:var(--muted)}
.regime-btn.active{border-color:var(--brand);background:#EFF6FF;color:var(--brand);box-shadow:0 2px 12px rgba(29,78,216,.12)}
.regime-btn:hover{border-color:var(--brand)}

/* ── Cards ────────────────────────────────────────────────── */
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

/* ── Form Fields ──────────────────────────────────────────── */
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

/* ── Result Panel ─────────────────────────────────────────── */
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
.toast{position:fixed;bottom:24px;right:24px;background:var(--ink);color:#fff;
       padding:11px 18px;border-radius:10px;font-size:13px;font-weight:500;
       transform:translateY(80px);transition:transform .3s;z-index:999}
.toast.show{transform:translateY(0)}

.print-btn{display:inline-flex;align-items:center;gap:5px;background:#F3F4F6;color:var(--ink);
           border:1px solid var(--border);border-radius:8px;padding:7px 14px;font-size:12px;
           font-weight:600;cursor:pointer;font-family:inherit;transition:all .2s;margin-top:8px}
.print-btn:hover{background:#E5E7EB}
@media print{nav,footer,.hero,.regime-toggle,.card:first-child,.btn-calc,.btn-reset,.print-btn,.toast{display:none!important}
             .result-panel{display:block!important}.calc-grid{display:block!important}
             .card{box-shadow:none!important;border:1px solid #ccc!important}}
</style></head><body>

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
  <div class="hero-badge">🧮 PY 2025-26 · AY 2026-27</div>
  <h1>Income Tax <em>Calculator</em></h1>
  <p>Calculate tax under Old &amp; New Regime. Enter income under 5 heads, deductions, TDS/TCS — get instant comparison with slab-wise breakup.</p>
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
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#EFF6FF">👤</div>
        <div><h2>Basic Information</h2><p>Assessee details</p></div>
      </div>
      <div class="card-body">
        <div class="row2">
          <div class="field">
            <label>Assessee Name <span style="font-weight:400;text-transform:none">(optional)</span></label>
            <input type="text" id="assesseeName" placeholder="e.g. Rajesh Kumar" style="border:1.5px solid var(--border);border-radius:8px;padding:9px 12px;font-family:inherit;font-size:13px;width:100%"/>
          </div>
          <div class="field">
            <label>Assessment Year</label>
            <select id="assessYear">
              <option value="2026-27" selected>AY 2026-27 (PY 2025-26)</option>
            </select>
          </div>
        </div>
        <div class="row2">
          <div class="field">
            <label>Age Category</label>
            <select id="ageCategory">
              <option value="below60">Below 60 years</option>
              <option value="senior">Senior Citizen (60-80)</option>
              <option value="supersenior">Super Senior Citizen (80+)</option>
            </select>
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
    </div>

    <!-- ──── 5 HEADS OF INCOME ──── -->
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#F0FDF4">💰</div>
        <div><h2>Income Under 5 Heads</h2><p>Gross Total Income computation</p></div>
      </div>
      <div class="card-body">
        <!-- HEAD 1: SALARY -->
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
          <p class="hint">₹75,000 for both regimes from FY 2024-25 onwards</p>
        </div>

        <!-- HEAD 2: HOUSE PROPERTY -->
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
            <p class="hint">Max ₹2L self-occupied (old regime). ₹2L in new regime too</p>
          </div>
        </div>

        <!-- HEAD 3: BUSINESS/PROFESSION -->
        <div class="section-title">💼 3. Profits from Business / Profession</div>
        <div class="field">
          <label>Net Profit from Business / Profession</label>
          <input type="number" id="businessIncome" placeholder="0"/>
          <p class="hint">After all business deductions. Can be loss (negative)</p>
        </div>

        <!-- HEAD 4: CAPITAL GAINS -->
        <div class="section-title">📈 4. Capital Gains</div>
        <div class="row2">
          <div class="field">
            <label>STCG u/s 111A (equity, STT paid)</label>
            <input type="number" id="stcg111a" placeholder="0" min="0"/>
            <p class="hint">Taxed at 20%</p>
          </div>
          <div class="field">
            <label>STCG — Other (non-equity)</label>
            <input type="number" id="stcgOther" placeholder="0" min="0"/>
            <p class="hint">Taxed at slab rates</p>
          </div>
        </div>
        <div class="row2">
          <div class="field">
            <label>LTCG u/s 112A (equity, STT paid)</label>
            <input type="number" id="ltcg112a" placeholder="0" min="0"/>
            <p class="hint">12.5% above ₹1.25 lakh exemption</p>
          </div>
          <div class="field">
            <label>LTCG — Other (property, debt etc.)</label>
            <input type="number" id="ltcgOther" placeholder="0" min="0"/>
            <p class="hint">12.5% flat (post July 2024 rule)</p>
          </div>
        </div>

        <!-- HEAD 5: OTHER SOURCES -->
        <div class="section-title">📦 5. Income from Other Sources</div>
        <div class="row2">
          <div class="field">
            <label>Interest Income / Dividends / Others</label>
            <input type="number" id="otherIncome" placeholder="0" min="0"/>
          </div>
          <div class="field">
            <label>Winnings (lottery, games etc.)</label>
            <input type="number" id="winningsIncome" placeholder="0" min="0"/>
            <p class="hint">Taxed at 30% flat</p>
          </div>
        </div>
      </div>
    </div>

    <!-- ──── DEDUCTIONS (OLD REGIME) ──── -->
    <div class="card" id="deductions-card">
      <div class="card-head">
        <div class="icon" style="background:#FFFBEB">🧾</div>
        <div><h2>Deductions (Chapter VI-A)</h2><p>Applicable in Old Regime only</p></div>
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
            <p class="hint">14% of salary (available in both regimes)</p>
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

    <!-- ──── CALCULATE BUTTON ──── -->
    <button class="btn-calc" onclick="calculateTax()">🧮 Calculate Tax</button>
    <button class="btn-reset" onclick="resetForm()">↺ Reset All Fields</button>
  </div>

  <!-- RIGHT: Results Section -->
  <div id="result-section">
    <div class="result-panel" id="resultPanel">

      <!-- Single regime result -->
      <div class="card" id="singleResult" style="display:none">
        <div class="card-head">
          <div class="icon" style="background:#ECFDF5">📊</div>
          <div><h2 id="resultTitle">Tax Computation</h2><p id="resultSubtitle">PY 2025-26 · AY 2026-27</p></div>
        </div>
        <div class="card-body">
          <div id="resultBody"></div>
          <button class="print-btn" onclick="window.print()">🖨️ Print / Save PDF</button>
        </div>
      </div>

      <!-- Comparison result -->
      <div id="compareResult" style="display:none">
        <div class="compare-box">
          <h3>⚖️ Regime Comparison</h3>
          <div class="regime-winner" id="regimeWinner"></div>
          <div class="savings" id="savingsAmt"></div>
          <table class="compare-table">
            <thead><tr><th></th><th>🆕 New Regime</th><th>📜 Old Regime</th></tr></thead>
            <tbody id="compareBody"></tbody>
          </table>
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

      <!-- Slab Breakup -->
      <div class="card" style="margin-top:16px" id="slabCard">
        <div class="card-head">
          <div class="icon" style="background:#F0FDF4">📋</div>
          <div><h2>Slab-wise Tax Breakup</h2><p id="slabRegimeLabel"></p></div>
        </div>
        <div class="card-body" id="slabBody"></div>
      </div>

      <div class="disclaimer">
        <strong>⚠ Disclaimer:</strong> This calculator is for estimation purposes only based on Income Tax provisions
        for PY 2025-26 (AY 2026-27). Actual tax liability may differ based on specific exemptions, deductions, and
        interpretations. Always consult a qualified Chartered Accountant for final tax computation.
        Surcharge and marginal relief calculations are indicative. Special rate incomes (capital gains, winnings)
        are not eligible for Section 87A rebate.
      </div>
    </div>

    <!-- Before calculation: show slab reference -->
    <div id="preCalcInfo">
      <div class="card">
        <div class="card-head">
          <div class="icon" style="background:#EFF6FF">📋</div>
          <div><h2>New Regime Slab Rates</h2><p>AY 2026-27 (Default)</p></div>
        </div>
        <div class="card-body">
          <table class="slab-table">
            <thead><tr><th>Income Slab</th><th style="text-align:right">Rate</th></tr></thead>
            <tbody>
              <tr><td>Up to ₹4,00,000</td><td class="amt">Nil</td></tr>
              <tr><td>₹4,00,001 – ₹8,00,000</td><td class="amt">5%</td></tr>
              <tr><td>₹8,00,001 – ₹12,00,000</td><td class="amt">10%</td></tr>
              <tr><td>₹12,00,001 – ₹16,00,000</td><td class="amt">15%</td></tr>
              <tr><td>₹16,00,001 – ₹20,00,000</td><td class="amt">20%</td></tr>
              <tr><td>₹20,00,001 – ₹24,00,000</td><td class="amt">25%</td></tr>
              <tr><td>Above ₹24,00,000</td><td class="amt">30%</td></tr>
            </tbody>
          </table>
          <p class="hint" style="margin-top:10px">Rebate u/s 87A: Income up to ₹12 lakh → zero tax (max rebate ₹60,000). Standard deduction ₹75,000 for salaried → effective ₹12.75L tax-free.</p>
        </div>
      </div>
      <div class="card">
        <div class="card-head">
          <div class="icon" style="background:#FFFBEB">📋</div>
          <div><h2>Old Regime Slab Rates</h2><p>AY 2026-27</p></div>
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
          <p class="hint" style="margin-top:10px">Senior citizens (60-80): exempt up to ₹3L. Super seniors (80+): exempt up to ₹5L. Rebate u/s 87A: up to ₹5L → max ₹12,500 rebate.</p>
        </div>
      </div>
      <div class="card">
        <div class="card-head">
          <div class="icon" style="background:#F5F3FF">📋</div>
          <div><h2>Special Rate Incomes</h2><p>AY 2026-27</p></div>
        </div>
        <div class="card-body">
          <table class="slab-table">
            <thead><tr><th>Type</th><th style="text-align:right">Rate</th></tr></thead>
            <tbody>
              <tr><td>STCG u/s 111A (equity, STT paid)</td><td class="amt">20%</td></tr>
              <tr><td>LTCG u/s 112A (equity, STT paid) above ₹1.25L</td><td class="amt">12.5%</td></tr>
              <tr><td>LTCG — Other assets</td><td class="amt">12.5%</td></tr>
              <tr><td>Winnings (lottery, games, etc.)</td><td class="amt">30%</td></tr>
            </tbody>
          </table>
          <p class="hint" style="margin-top:10px">Health &amp; Education Cess @ 4% applicable on tax + surcharge. Surcharge: 10% (50L–1Cr), 15% (1Cr–2Cr), 25% (2Cr–5Cr), 37% (above 5Cr — old regime only). New regime max surcharge 25%.</p>
        </div>
      </div>
    </div>
  </div>
</div>
</div>

<footer>
  <p class="footer-brand">CA Toolkit</p>
  <p>Built for Indian Chartered Accountants · Saves hours every year</p>
  <p style="margin-top:6px">Created by CA Articles of GD Singla &amp; Co.</p>
  <p style="margin-top:12px;font-size:11px">© 2026 CA Toolkit · <span style="color:var(--accent)">Income Tax Calculator — PY 2025-26 / AY 2026-27</span></p>
</footer>
<div class="toast" id="toast"></div>

<script>
/* ═══════════════════════════════════════════════════════════════════════
   INCOME TAX CALCULATOR — PY 2025-26 / AY 2026-27
   ═══════════════════════════════════════════════════════════════════════ */

let currentRegime = 'new';

function setRegime(r) {
  currentRegime = r;
  document.querySelectorAll('.regime-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('btn-' + r).classList.add('active');
  // Show/hide deductions card
  const dc = document.getElementById('deductions-card');
  if (r === 'new') { dc.style.opacity = '0.4'; dc.style.pointerEvents = 'none'; }
  else { dc.style.opacity = '1'; dc.style.pointerEvents = 'auto'; }
}

// Init
setRegime('new');

function v(id) { return parseFloat(document.getElementById(id).value) || 0; }
function fmt(n) {
  if (n < 0) return '-₹' + Math.abs(Math.round(n)).toLocaleString('en-IN');
  return '₹' + Math.round(n).toLocaleString('en-IN');
}

/* ── NEW REGIME SLABS ─────────────────────────────────────────────── */
const NEW_SLABS = [
  { upto: 400000,  rate: 0    },
  { upto: 800000,  rate: 0.05 },
  { upto: 1200000, rate: 0.10 },
  { upto: 1600000, rate: 0.15 },
  { upto: 2000000, rate: 0.20 },
  { upto: 2400000, rate: 0.25 },
  { upto: Infinity, rate: 0.30 },
];

/* ── OLD REGIME SLABS (below 60) ──────────────────────────────────── */
const OLD_SLABS_BELOW60 = [
  { upto: 250000,  rate: 0    },
  { upto: 500000,  rate: 0.05 },
  { upto: 1000000, rate: 0.20 },
  { upto: Infinity, rate: 0.30 },
];
const OLD_SLABS_SENIOR = [
  { upto: 300000,  rate: 0    },
  { upto: 500000,  rate: 0.05 },
  { upto: 1000000, rate: 0.20 },
  { upto: Infinity, rate: 0.30 },
];
const OLD_SLABS_SUPERSENIOR = [
  { upto: 500000,  rate: 0    },
  { upto: 1000000, rate: 0.20 },
  { upto: Infinity, rate: 0.30 },
];

function getOldSlabs() {
  const age = document.getElementById('ageCategory').value;
  if (age === 'supersenior') return OLD_SLABS_SUPERSENIOR;
  if (age === 'senior') return OLD_SLABS_SENIOR;
  return OLD_SLABS_BELOW60;
}

/* ── Slab-based tax calculation ───────────────────────────────────── */
function calcSlabTax(taxableIncome, slabs) {
  let tax = 0;
  let prev = 0;
  const breakup = [];
  for (const slab of slabs) {
    if (taxableIncome <= prev) break;
    const chunk = Math.min(taxableIncome, slab.upto) - prev;
    const t = chunk * slab.rate;
    breakup.push({
      from: prev, to: Math.min(taxableIncome, slab.upto),
      rate: slab.rate, amount: chunk, tax: t
    });
    tax += t;
    prev = slab.upto;
  }
  return { tax, breakup };
}

/* ── Surcharge ────────────────────────────────────────────────────── */
function calcSurcharge(tax, totalIncome, isNewRegime) {
  if (totalIncome <= 5000000) return 0;
  let rate = 0;
  if (isNewRegime) {
    // New regime: max surcharge 25%
    if (totalIncome <= 10000000) rate = 0.10;
    else if (totalIncome <= 20000000) rate = 0.15;
    else rate = 0.25;
  } else {
    // Old regime
    if (totalIncome <= 10000000) rate = 0.10;
    else if (totalIncome <= 20000000) rate = 0.15;
    else if (totalIncome <= 50000000) rate = 0.25;
    else rate = 0.37;
  }
  let surcharge = tax * rate;

  // Marginal relief
  const thresholds = [5000000, 10000000, 20000000, 50000000];
  for (const th of thresholds) {
    if (totalIncome > th && totalIncome <= th * 1.2) {
      const excess = totalIncome - th;
      const taxAtThreshold = calcSlabTax(th, isNewRegime ? NEW_SLABS : getOldSlabs()).tax;
      const surchargeAtThreshold = calcSurcharge(taxAtThreshold, th, isNewRegime);
      const maxTax = taxAtThreshold + surchargeAtThreshold + excess;
      if (tax + surcharge > maxTax) {
        surcharge = maxTax - tax;
        if (surcharge < 0) surcharge = 0;
      }
    }
  }
  return surcharge;
}

/* ── Surcharge for special rate incomes (capped at 15%) ───────── */
function calcSurchargeCapped(tax, totalIncome) {
  if (totalIncome <= 5000000) return 0;
  let rate = Math.min(0.15, totalIncome > 10000000 ? 0.15 : 0.10);
  return tax * rate;
}

/* ── Main Computation ─────────────────────────────────────────── */
function computeForRegime(isNew) {
  const name = document.getElementById('assesseeName').value.trim();
  const age = document.getElementById('ageCategory').value;
  const isResident = document.getElementById('residentialStatus').value === 'resident';

  // ── Head 1: Salary
  const grossSalary = v('grossSalary');
  const exemptAllow = isNew ? 0 : v('exemptAllow');
  const stdDed = v('stdDeduction');
  const salaryIncome = Math.max(0, grossSalary - exemptAllow - stdDed);

  // ── Head 2: House Property
  let houseRaw = v('houseIncome');
  const loanInt = v('homeLoanInterest');
  const houseIncome = houseRaw - loanInt; // can be negative (loss)
  const houseLoss = houseIncome < 0 ? houseIncome : 0;
  // Set-off of house property loss capped at ₹2L against other heads
  const houseLossCapped = Math.max(houseLoss, -200000);

  // ── Head 3: Business
  const businessIncome = v('businessIncome');

  // ── Head 4: Capital Gains
  const stcg111a = v('stcg111a');
  const stcgOther = v('stcgOther');
  const ltcg112a = v('ltcg112a');
  const ltcgOther = v('ltcgOther');

  // ── Head 5: Other Sources
  const otherIncome = v('otherIncome');
  const winnings = v('winningsIncome');

  // ── Gross Total Income
  const normalIncome = salaryIncome + Math.max(0, houseIncome) + businessIncome + stcgOther + otherIncome;
  // Apply house loss set-off
  const normalAfterLoss = Math.max(0, normalIncome + houseLossCapped);

  // ── Deductions (Old Regime only, except 80CCD(2) available in both)
  let totalDeductions = 0;
  let ded80ccd2 = v('ded80ccd2');
  if (isNew) {
    totalDeductions = ded80ccd2; // Only employer NPS in new regime
  } else {
    const ded80c = Math.min(v('ded80c'), 150000);
    const ded80ccd1b = Math.min(v('ded80ccd1b'), 50000);
    const ded80d = v('ded80d');
    const ded80e = v('ded80e');
    const ded80g = v('ded80g');
    const ded80tta = v('ded80tta');
    const dedOth = v('dedOther');
    totalDeductions = ded80c + ded80ccd1b + ded80ccd2 + ded80d + ded80e + ded80g + ded80tta + dedOth;
  }

  // ── Taxable Income (normal slab portion)
  const normalTaxable = Math.max(0, normalAfterLoss - totalDeductions);

  // ── Tax on Normal Income at slab rates
  const slabs = isNew ? NEW_SLABS : getOldSlabs();
  const slabResult = calcSlabTax(normalTaxable, slabs);
  let normalTax = slabResult.tax;

  // ── Section 87A Rebate on normal slab income
  let rebate87a = 0;
  if (isResident) {
    if (isNew && normalTaxable <= 1200000) {
      rebate87a = Math.min(normalTax, 60000);
    } else if (!isNew && normalTaxable <= 500000) {
      rebate87a = Math.min(normalTax, 12500);
    }
  }
  normalTax = Math.max(0, normalTax - rebate87a);

  // ── Special rate taxes
  const taxSTCG111A = stcg111a * 0.20;
  const ltcg112aExempt = Math.min(ltcg112a, 125000);
  const taxLTCG112A = Math.max(0, ltcg112a - 125000) * 0.125;
  const taxLTCGOther = ltcgOther * 0.125;
  const taxWinnings = winnings * 0.30;

  const totalSpecialTax = taxSTCG111A + taxLTCG112A + taxLTCGOther + taxWinnings;

  // ── Total income for surcharge calculation
  const totalIncome = normalTaxable + stcg111a + ltcg112a + ltcgOther + winnings;

  // ── Surcharge
  const surchargeNormal = calcSurcharge(normalTax, totalIncome, isNew);
  const surchargeSpecial = calcSurchargeCapped(totalSpecialTax, totalIncome);
  const totalSurcharge = surchargeNormal + surchargeSpecial;

  // ── Health & Education Cess @ 4%
  const totalBeforeCess = normalTax + totalSpecialTax + totalSurcharge;
  const cess = totalBeforeCess * 0.04;

  // ── Total Tax
  const totalTax = totalBeforeCess + cess;

  // ── Less: TDS, TCS, Advance Tax
  const tdsPaid = v('tds');
  const tcsPaid = v('tcs');
  const advTax = v('advanceTax');
  const totalPrepaid = tdsPaid + tcsPaid + advTax;

  const netPayable = totalTax - totalPrepaid;

  return {
    name, age, isNew, isResident,
    grossSalary, exemptAllow, stdDed, salaryIncome,
    houseRaw, loanInt, houseIncome, houseLossCapped,
    businessIncome,
    stcg111a, stcgOther, ltcg112a, ltcg112aExempt, ltcgOther,
    otherIncome, winnings,
    normalIncome: normalAfterLoss,
    totalDeductions,
    normalTaxable,
    slabResult,
    normalTax: normalTax + rebate87a, // before rebate for display
    rebate87a,
    normalTaxAfterRebate: normalTax,
    taxSTCG111A, taxLTCG112A, taxLTCGOther, taxWinnings,
    totalSpecialTax,
    totalIncome,
    surchargeNormal, surchargeSpecial, totalSurcharge,
    cess,
    totalTax,
    tdsPaid, tcsPaid, advTax, totalPrepaid,
    netPayable,
  };
}

/* ── Render single regime result ───────────────────────────────── */
function renderResult(r) {
  const regLabel = r.isNew ? 'New Regime' : 'Old Regime';
  let h = '';

  h += row('Gross Salary', fmt(r.grossSalary));
  if (!r.isNew) h += row('Less: Exempt Allowances', fmt(-r.exemptAllow), 'sub');
  h += row('Less: Standard Deduction', fmt(-r.stdDed), 'sub');
  h += row('Net Salary Income (Head 1)', fmt(r.salaryIncome));
  h += '<div style="height:6px"></div>';

  h += row('House Property Income (Head 2)', fmt(r.houseIncome));
  if (r.houseLossCapped < 0) h += row('Loss set-off (max ₹2L)', fmt(r.houseLossCapped), 'sub');
  h += row('Business Income (Head 3)', fmt(r.businessIncome));

  if (r.stcg111a || r.stcgOther || r.ltcg112a || r.ltcgOther)
    h += row('Capital Gains (Head 4)', fmt(r.stcg111a + r.stcgOther + r.ltcg112a + r.ltcgOther));
  if (r.stcg111a) h += row('STCG u/s 111A @ 20%', fmt(r.stcg111a), 'sub');
  if (r.stcgOther) h += row('STCG — Other (slab rate)', fmt(r.stcgOther), 'sub');
  if (r.ltcg112a) h += row('LTCG u/s 112A (exempt ₹1.25L)', fmt(r.ltcg112a), 'sub');
  if (r.ltcgOther) h += row('LTCG — Other @ 12.5%', fmt(r.ltcgOther), 'sub');

  h += row('Other Sources (Head 5)', fmt(r.otherIncome + r.winnings));
  if (r.winnings) h += row('Winnings @ 30%', fmt(r.winnings), 'sub');

  h += '<div style="height:4px;border-top:2px solid var(--border);margin:10px 0"></div>';
  h += row('Gross Total Income', fmt(r.normalIncome + r.stcg111a + r.ltcg112a + r.ltcgOther + r.winnings), 'total');

  if (r.totalDeductions > 0) {
    h += row('Less: Deductions Ch VI-A', fmt(-r.totalDeductions));
  }
  h += row('Total Taxable Income (Normal)', fmt(r.normalTaxable), 'total');

  h += '<div style="height:4px;border-top:2px solid var(--border);margin:10px 0"></div>';
  h += row('Tax on Normal Income (slab)', fmt(r.normalTax));
  if (r.rebate87a > 0) h += row('Less: Rebate u/s 87A', fmt(-r.rebate87a), 'sub');
  h += row('Tax after Rebate', fmt(r.normalTaxAfterRebate));

  if (r.totalSpecialTax > 0) {
    h += '<div style="height:6px"></div>';
    if (r.taxSTCG111A) h += row('Tax on STCG 111A @ 20%', fmt(r.taxSTCG111A), 'sub');
    if (r.taxLTCG112A) h += row('Tax on LTCG 112A @ 12.5%', fmt(r.taxLTCG112A), 'sub');
    if (r.taxLTCGOther) h += row('Tax on LTCG Other @ 12.5%', fmt(r.taxLTCGOther), 'sub');
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

function row(lbl, val, cls) {
  return `<div class="result-row ${cls||''}""><span class="lbl">${lbl}</span><span class="val">${val}</span></div>`;
}

/* ── Render slab breakup table ─────────────────────────────────── */
function renderSlabs(result) {
  const slabs = result.slabResult.breakup;
  let h = '<table class="slab-table"><thead><tr><th>Slab</th><th style="text-align:right">Income</th><th style="text-align:right">Rate</th><th style="text-align:right">Tax</th></tr></thead><tbody>';
  for (const s of slabs) {
    h += `<tr><td>₹${Math.round(s.from).toLocaleString('en-IN')} – ₹${s.to===Infinity?'∞':Math.round(s.to).toLocaleString('en-IN')}</td>
          <td class="amt">${fmt(s.amount)}</td>
          <td class="amt">${(s.rate*100).toFixed(0)}%</td>
          <td class="amt">${fmt(s.tax)}</td></tr>`;
  }
  h += `<tr style="font-weight:800;border-top:2px solid var(--border)"><td>Total</td><td></td><td></td><td class="amt">${fmt(result.slabResult.tax)}</td></tr>`;
  h += '</tbody></table>';
  return h;
}

/* ── CALCULATE ────────────────────────────────────────────────── */
function calculateTax() {
  const panel = document.getElementById('resultPanel');
  const preInfo = document.getElementById('preCalcInfo');

  if (currentRegime === 'both') {
    // Compare both
    const rNew = computeForRegime(true);
    const rOld = computeForRegime(false);

    document.getElementById('singleResult').style.display = 'none';
    document.getElementById('compareResult').style.display = 'block';

    // Winner
    const diff = Math.abs(rNew.totalTax - rOld.totalTax);
    const winner = rNew.totalTax <= rOld.totalTax ? 'New Regime' : 'Old Regime';
    document.getElementById('regimeWinner').innerHTML =
      `<strong>${winner}</strong> saves you more tax`;
    document.getElementById('savingsAmt').textContent = 'Save ' + fmt(diff);

    // Compare table
    let ct = '';
    ct += cmpRow('Taxable Income', rNew.normalTaxable, rOld.normalTaxable, true);
    ct += cmpRow('Tax on Normal Income', rNew.normalTaxAfterRebate, rOld.normalTaxAfterRebate, true);
    ct += cmpRow('Tax on Special Income', rNew.totalSpecialTax, rOld.totalSpecialTax, true);
    ct += cmpRow('Surcharge', rNew.totalSurcharge, rOld.totalSurcharge, true);
    ct += cmpRow('Cess', rNew.cess, rOld.cess, true);
    ct += cmpRow('Total Tax', rNew.totalTax, rOld.totalTax, true);
    ct += cmpRow('Net Payable/Refund', rNew.netPayable, rOld.netPayable, true);
    document.getElementById('compareBody').innerHTML = ct;

    // Detailed
    document.getElementById('newRegimeDetail').innerHTML = renderResult(rNew);
    document.getElementById('oldRegimeDetail').innerHTML = renderResult(rOld);

    // Slab cards for both
    document.getElementById('slabRegimeLabel').textContent = 'New Regime';
    document.getElementById('slabBody').innerHTML =
      '<h3 style="font-size:13px;font-weight:700;margin-bottom:8px">🆕 New Regime Slabs</h3>' +
      renderSlabs(rNew) +
      '<div style="height:16px"></div>' +
      '<h3 style="font-size:13px;font-weight:700;margin-bottom:8px">📜 Old Regime Slabs</h3>' +
      renderSlabs(rOld);

  } else {
    // Single regime
    const isNew = currentRegime === 'new';
    const result = computeForRegime(isNew);

    document.getElementById('singleResult').style.display = 'block';
    document.getElementById('compareResult').style.display = 'none';

    const label = isNew ? '🆕 New Regime' : '📜 Old Regime';
    document.getElementById('resultTitle').textContent = 'Tax Computation — ' + label;
    if (result.name) {
      document.getElementById('resultSubtitle').textContent = result.name + ' · PY 2025-26 · AY 2026-27';
    }

    document.getElementById('resultBody').innerHTML = renderResult(result);
    document.getElementById('slabRegimeLabel').textContent = isNew ? 'New Regime' : 'Old Regime';
    document.getElementById('slabBody').innerHTML = renderSlabs(result);
  }

  panel.classList.add('show');
  preInfo.style.display = 'none';
  document.getElementById('slabCard').style.display = 'block';

  // Scroll to results
  document.getElementById('result-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
  toast('Tax calculated successfully!');
}

function cmpRow(label, valNew, valOld, lowerBetter) {
  const nw = fmt(valNew), ol = fmt(valOld);
  let nCls = '', oCls = '';
  if (lowerBetter) {
    if (valNew < valOld) nCls = 'winner'; else if (valOld < valNew) oCls = 'winner';
  }
  return `<tr><td style="text-align:left;font-weight:500">${label}</td><td class="${nCls}">${nw}</td><td class="${oCls}">${ol}</td></tr>`;
}

function resetForm() {
  document.querySelectorAll('input[type=number]').forEach(i => {
    if (i.id === 'stdDeduction') i.value = 75000;
    else i.value = '';
  });
  document.getElementById('assesseeName').value = '';
  document.getElementById('resultPanel').classList.remove('show');
  document.getElementById('preCalcInfo').style.display = 'block';
  document.getElementById('singleResult').style.display = 'none';
  document.getElementById('compareResult').style.display = 'none';
  setRegime('new');
  toast('Form reset');
}

function toast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3000);
}
</script>
</body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
#  ADMIN PANEL
# ══════════════════════════════════════════════════════════════════════════════

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
              <option value="starter">Starter (10 uploads · ₹399)</option>
              <option value="standard" selected>Standard (25 uploads · ₹899)</option>
              <option value="pro">Professional (60 uploads · ₹1,799)</option>
              <option value="firm">Firm (150 uploads · ₹3,499)</option>
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
  <p class="footer-brand">CA Toolkit — Admin</p>
  <p>Created by CA Articles of GD Singla &amp; Co.</p>
</footer>
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
@login_required
def dashboard():
    user = get_user_by_id(session["uid"])
    return render_template_string(DASHBOARD_T, **user_ctx(user))

@app.route("/tool/converter")
@login_required
def tool_converter():
    user = get_user_by_id(session["uid"])
    return render_template_string(CONVERTER_T, **user_ctx(user))

@app.route("/tool/tax-calculator")
@login_required
def tool_tax_calculator():
    user = get_user_by_id(session["uid"])
    return render_template_string(TAX_CALC_T, **user_ctx(user))

# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES — PROCESS & DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/process", methods=["POST"])
@login_required
def process_file():
    user = get_user_by_id(session["uid"])
    if not user["is_admin"] and uploads_remaining(user) <= 0:
        return jsonify({"status": "error",
            "message": f"No uploads remaining. Contact {CONTACT_EMAIL} to recharge."})
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "No file uploaded."})
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"status": "error", "message": "Only .xlsx files are supported."})
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
    f.save(ip)
    fname = f"{on or os.path.splitext(f.filename)[0]}_{ny}.xlsx"
    try:
        result = process(ip, op, cy, ny)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})
    finally:
        try: os.remove(ip)
        except: pass
    log_usage(user["id"], fname)
    return jsonify({"status": "success", "log": result["log"], "file_id": h, "filename": fname})

@app.route("/download/<fid>")
@login_required
def download(fid):
    if not re.fullmatch(r"[a-f0-9]{32}", fid): return "Invalid ID", 400
    path = os.path.join(OUTPUT_DIR, f"{fid}_out.xlsx")
    if not os.path.exists(path): return "File not found or expired.", 404
    return send_file(path, as_attachment=True,
        download_name=f"bs_shift_{fid[:8]}.xlsx",
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

# ══════════════════════════════════════════════════════════════════════════════
#  STARTUP
# ══════════════════════════════════════════════════════════════════════════════

init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
