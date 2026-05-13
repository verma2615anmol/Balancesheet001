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
    "starter":  {"label": "Starter",      "uploads": 10,  "price": 60},
    "standard": {"label": "Standard",     "uploads": 25,  "price": 130},
    "pro":      {"label": "Professional", "uploads": 60,  "price": 270},
    "firm":     {"label": "Firm",         "uploads": 150, "price": 600},
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
    {% if username %}
    {% if is_admin %}<a href="/admin" class="nav-btn">Admin Panel</a>{% endif %}
    <a href="/logout" class="nav-link">Sign out</a>
    {% else %}
    <a href="/login" class="nav-btn">Sign In</a>
    {% endif %}
  </div>
</nav>

<div class="hero">
  <div class="hero-badge">🇮🇳 Made for Indian CAs &amp; Accountants</div>
  <h1>Your Complete <em>CA Toolkit</em></h1>
  <p>Professional tools built by CA Articles of GD Singla &amp; Co. — designed to save hours of manual work every year.</p>
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

  <!-- PREMIUM: Login required -->
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

  <!-- PREMIUM: Login required -->
  <div class="tool-card disabled" style="position:relative">
    <div style="position:absolute;top:12px;right:12px;background:#FEF3C7;color:#92400E;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🔒 Premium</div>
    <div class="tool-icon" style="background:#F0FDF4">📋</div>
    <h2>Balance Sheet from Trial Balance</h2>
    <p>Generate a formatted comparative balance sheet directly from your trial balance data. No manual formatting required.</p>
    <span class="tool-tag tag-soon">🔜 Coming Soon</span>
  </div>

  <!-- FREE: No login needed -->
  <a href="/tool/tax-calculator" class="tool-card">
    <div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🆓 Free</div>
    <div class="tool-icon" style="background:#FFFBEB">🧮</div>
    <h2>Income Tax Calculator</h2>
    <p>Calculate income tax under old and new regime for PY 2025-26. Income under 5 heads, TDS/TCS, surcharge &amp; cess — all built in.</p>
    <span class="tool-tag tag-live">✓ Live · Free</span>
    <div class="arrow">→</div>
  </a>

  <!-- FREE: No login needed -->
  <a href="/tool/tds-calculator" class="tool-card">
    <div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🆓 Free</div>
    <div class="tool-icon" style="background:#EFF6FF">📑</div>
    <h2>TDS Calculator</h2>
    <p>Calculate TDS amount, rate and due date for any payment type. Covers all major sections — 194C, 194J, 194H, 192 and more.</p>
    <span class="tool-tag tag-live">✓ Live · Free</span>
    <div class="arrow">→</div>
  </a>

  <!-- FREE: No login needed -->
  <a href="/tool/depreciation-calculator" class="tool-card">
    <div style="position:absolute;top:12px;right:12px;background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:3px 8px;border-radius:99px">🆓 Free</div>
    <div class="tool-icon" style="background:#F5F3FF">🏭</div>
    <h2>Depreciation Calculator</h2>
    <p>Calculate depreciation under Companies Act 2013 (WDV/SLM) and Income Tax Act. Get full schedule with opening/closing WDV.</p>
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

@media print{nav,footer,.hero,.regime-toggle,.card:first-child,.btn-calc,.btn-reset,.print-btn,.toast,.year-pills{display:none!important}
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
  <div class="hero-badge">🧮 Multi-Year · PY 2023-24 to PY 2026-27</div>
  <h1>Income Tax <em>Calculator</em></h1>
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
    <div class="card">
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
        <div class="row2">
          <div class="field">
            <label>Assessee Name <span style="font-weight:400;text-transform:none">(optional)</span></label>
            <input type="text" id="assesseeName" placeholder="e.g. Rajesh Kumar" style="border:1.5px solid var(--border);border-radius:8px;padding:9px 12px;font-family:inherit;font-size:13px;width:100%"/>
          </div>
          <div class="field">
            <label>Age Category</label>
            <select id="ageCategory">
              <option value="below60">Below 60 years</option>
              <option value="senior">Senior Citizen (60-80)</option>
              <option value="supersenior">Super Senior Citizen (80+)</option>
            </select>
          </div>
        </div>
        <div class="field">
          <label>Residential Status</label>
          <select id="residentialStatus" style="max-width:280px">
            <option value="resident">Resident</option>
            <option value="nri">Non-Resident</option>
          </select>
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
        <div class="row2">
          <div class="field">
            <label>STCG u/s 111A (equity, STT paid)</label>
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
            <label>LTCG u/s 112A (equity, STT paid)</label>
            <input type="number" id="ltcg112a" placeholder="0" min="0"/>
            <p class="hint" id="ltcgRateHint">Rate &amp; exemption varies by year</p>
          </div>
          <div class="field">
            <label>LTCG — Other (property, debt etc.)</label>
            <input type="number" id="ltcgOther" placeholder="0" min="0"/>
            <p class="hint" id="ltcgOtherHint">Rate varies by year</p>
          </div>
        </div>

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

    <button class="btn-calc" onclick="calculateTax()">🧮 Calculate Tax</button>
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
  <p class="footer-brand">CA Toolkit</p>
  <p>Built for Indian Chartered Accountants · Saves hours every year</p>
  <p style="margin-top:6px">Created by CA Articles of GD Singla &amp; Co.</p>
  <p style="margin-top:12px;font-size:11px">© 2026 CA Toolkit · <span style="color:var(--accent)">Income Tax Calculator — PY 2023-24 to PY 2026-27</span></p>
</footer>
<div class="toast" id="toast"></div>

<script>
/* ═══════════════════════════════════════════════════════════════════════
   INCOME TAX CALCULATOR — Multi-Year Engine
   PY 2023-24 | PY 2024-25 | PY 2025-26 | PY 2026-27 (upcoming)
   ═══════════════════════════════════════════════════════════════════════ */

let currentRegime = 'new';
let currentYear = '2025-26';

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
    ltcgOtherRate: 0.20, // with indexation
    ltcgOtherLabel: '20% (with indexation)',
    maxSurchargeNew: 0.25,
  },
  '2024-25': {
    label: 'PY 2024-25 (AY 2025-26)',
    ayLabel: 'AY 2025-26',
    isFuture: false,
    stdDeduction: 75000,
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
    stcg111aRate: 0.20, // changed from July 2024
    ltcg112aRate: 0.125, // changed from July 2024
    ltcg112aExempt: 125000, // changed from July 2024
    ltcgOtherRate: 0.125, // changed from July 2024
    ltcgOtherLabel: '12.5% (no indexation)',
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

/* ── OLD REGIME SLABS (unchanged across all years) ────────────────── */
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

function getOldSlabs() {
  const age = document.getElementById('ageCategory').value;
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
  // Update std deduction default
  document.getElementById('stdDeduction').value = c.stdDeduction;
  document.getElementById('stdDedHint').textContent =
    c.stdDeduction === 50000 ? '₹50,000 for PY 2023-24' : '₹75,000 for PY 2024-25 onwards';

  // Update hints for capital gains rates
  document.getElementById('stcgRateHint').textContent =
    'Taxed at ' + (c.stcg111aRate * 100) + '%' + (currentYear === '2024-25' ? ' (post July 2024 budget)' : '');
  document.getElementById('ltcgRateHint').textContent =
    (c.ltcg112aRate * 100) + '% above ₹' + (c.ltcg112aExempt / 100000).toFixed(2).replace('.00','') + ' lakh exemption';
  document.getElementById('ltcgOtherHint').textContent = c.ltcgOtherLabel;

  // Future year note
  document.getElementById('futureYearNote').style.display = c.isFuture ? 'block' : 'none';

  // Update reference slab tables
  updateRefSlabs();

  // Hide results on year change
  document.getElementById('resultPanel').classList.remove('show');
  document.getElementById('preCalcInfo').style.display = 'block';
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

  // Special rates reference
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

/* ── Helpers ───────────────────────────────────────────────────────── */
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

  // Marginal relief
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

/* ── Main Computation ─────────────────────────────────────────────── */
function computeForRegime(isNew) {
  const c = cfg();
  const name = document.getElementById('assesseeName').value.trim();
  const isResident = document.getElementById('residentialStatus').value === 'resident';

  // Head 1: Salary
  const grossSalary = v('grossSalary');
  const exemptAllow = isNew ? 0 : v('exemptAllow');
  const stdDed = v('stdDeduction');
  const salaryIncome = Math.max(0, grossSalary - exemptAllow - stdDed);

  // Head 2: House Property
  let houseRaw = v('houseIncome');
  const loanInt = v('homeLoanInterest');
  const houseIncome = houseRaw - loanInt;
  const houseLoss = houseIncome < 0 ? houseIncome : 0;
  const houseLossCapped = Math.max(houseLoss, -200000);

  // Head 3: Business
  const businessIncome = v('businessIncome');

  // Head 4: Capital Gains
  const stcg111a = v('stcg111a');
  const stcgOther = v('stcgOther');
  const ltcg112a = v('ltcg112a');
  const ltcgOther = v('ltcgOther');

  // Head 5: Other Sources
  const otherIncome = v('otherIncome');
  const winnings = v('winningsIncome');

  // Gross Total Income (normal portion)
  const normalIncome = salaryIncome + Math.max(0, houseIncome) + businessIncome + stcgOther + otherIncome;
  const normalAfterLoss = Math.max(0, normalIncome + houseLossCapped);

  // Deductions
  let totalDeductions = 0;
  let ded80ccd2 = v('ded80ccd2');
  if (isNew) {
    totalDeductions = ded80ccd2;
  } else {
    totalDeductions = Math.min(v('ded80c'), 150000) + Math.min(v('ded80ccd1b'), 50000) +
      ded80ccd2 + v('ded80d') + v('ded80e') + v('ded80g') + v('ded80tta') + v('dedOther');
  }

  const normalTaxable = Math.max(0, normalAfterLoss - totalDeductions);

  // Tax on normal income
  const slabs = isNew ? c.newSlabs : getOldSlabs();
  const slabResult = calcSlabTax(normalTaxable, slabs);
  let normalTax = slabResult.tax;

  // Rebate u/s 87A
  let rebate87a = 0;
  if (isResident) {
    const rebateCfg = isNew ? c.rebateNew : c.rebateOld;
    if (normalTaxable <= rebateCfg.limit) {
      rebate87a = Math.min(normalTax, rebateCfg.max);
    }
  }
  normalTax = Math.max(0, normalTax - rebate87a);

  // Special rate taxes (year-specific rates)
  const taxSTCG111A = stcg111a * c.stcg111aRate;
  const ltcg112aExempt = Math.min(ltcg112a, c.ltcg112aExempt);
  const taxLTCG112A = Math.max(0, ltcg112a - c.ltcg112aExempt) * c.ltcg112aRate;
  const taxLTCGOther = ltcgOther * c.ltcgOtherRate;
  const taxWinnings = winnings * 0.30;
  const totalSpecialTax = taxSTCG111A + taxLTCG112A + taxLTCGOther + taxWinnings;

  const totalIncome = normalTaxable + stcg111a + ltcg112a + ltcgOther + winnings;

  // Surcharge
  const surchargeNormal = calcSurcharge(normalTax, totalIncome, isNew);
  const surchargeSpecial = calcSurchargeCapped(totalSpecialTax, totalIncome);
  const totalSurcharge = surchargeNormal + surchargeSpecial;

  // Cess
  const totalBeforeCess = normalTax + totalSpecialTax + totalSurcharge;
  const cess = totalBeforeCess * 0.04;
  const totalTax = totalBeforeCess + cess;

  // Prepaid taxes
  const tdsPaid = v('tds');
  const tcsPaid = v('tcs');
  const advTax = v('advanceTax');
  const totalPrepaid = tdsPaid + tcsPaid + advTax;
  const netPayable = totalTax - totalPrepaid;

  return {
    yearLabel: c.label, ayLabel: c.ayLabel, isFuture: c.isFuture,
    name, isNew, isResident, c,
    grossSalary, exemptAllow, stdDed, salaryIncome,
    houseRaw, loanInt, houseIncome, houseLossCapped,
    businessIncome,
    stcg111a, stcgOther, ltcg112a, ltcg112aExempt, ltcgOther,
    otherIncome, winnings,
    normalIncome: normalAfterLoss, totalDeductions, normalTaxable,
    slabResult,
    normalTax: normalTax + rebate87a, rebate87a,
    normalTaxAfterRebate: normalTax,
    taxSTCG111A, taxLTCG112A, taxLTCGOther, taxWinnings, totalSpecialTax,
    totalIncome, surchargeNormal, surchargeSpecial, totalSurcharge,
    cess, totalTax,
    tdsPaid, tcsPaid, advTax, totalPrepaid, netPayable,
  };
}

/* ── Render result rows ────────────────────────────────────────────── */
function row(lbl, val, cls) {
  return '<div class="result-row '+(cls||'')+'"><span class="lbl">'+lbl+'</span><span class="val">'+val+'</span></div>';
}

function renderResult(r) {
  const c = r.c;
  let h = '';
  h += row('Gross Salary', fmt(r.grossSalary));
  if (!r.isNew) h += row('Less: Exempt Allowances', fmt(-r.exemptAllow), 'sub');
  h += row('Less: Standard Deduction (₹'+c.stdDeduction.toLocaleString('en-IN')+')', fmt(-r.stdDed), 'sub');
  h += row('Net Salary Income (Head 1)', fmt(r.salaryIncome));
  h += '<div style="height:6px"></div>';
  h += row('House Property Income (Head 2)', fmt(r.houseIncome));
  if (r.houseLossCapped < 0) h += row('Loss set-off (max ₹2L)', fmt(r.houseLossCapped), 'sub');
  h += row('Business Income (Head 3)', fmt(r.businessIncome));

  if (r.stcg111a || r.stcgOther || r.ltcg112a || r.ltcgOther)
    h += row('Capital Gains (Head 4)', fmt(r.stcg111a + r.stcgOther + r.ltcg112a + r.ltcgOther));
  if (r.stcg111a) h += row('STCG u/s 111A @ '+(c.stcg111aRate*100)+'%', fmt(r.stcg111a), 'sub');
  if (r.stcgOther) h += row('STCG — Other (slab rate)', fmt(r.stcgOther), 'sub');
  if (r.ltcg112a) h += row('LTCG u/s 112A (exempt ₹'+(c.ltcg112aExempt/100000)+'L)', fmt(r.ltcg112a), 'sub');
  if (r.ltcgOther) h += row('LTCG — Other @ '+c.ltcgOtherLabel, fmt(r.ltcgOther), 'sub');

  h += row('Other Sources (Head 5)', fmt(r.otherIncome + r.winnings));
  if (r.winnings) h += row('Winnings @ 30%', fmt(r.winnings), 'sub');

  h += '<div style="height:4px;border-top:2px solid var(--border);margin:10px 0"></div>';
  h += row('Gross Total Income', fmt(r.normalIncome + r.stcg111a + r.ltcg112a + r.ltcgOther + r.winnings), 'total');
  if (r.totalDeductions > 0) h += row('Less: Deductions Ch VI-A', fmt(-r.totalDeductions));
  h += row('Total Taxable Income (Normal)', fmt(r.normalTaxable), 'total');

  h += '<div style="height:4px;border-top:2px solid var(--border);margin:10px 0"></div>';
  h += row('Tax on Normal Income (slab)', fmt(r.normalTax));
  if (r.rebate87a > 0) h += row('Less: Rebate u/s 87A', fmt(-r.rebate87a), 'sub');
  h += row('Tax after Rebate', fmt(r.normalTaxAfterRebate));

  if (r.totalSpecialTax > 0) {
    h += '<div style="height:6px"></div>';
    if (r.taxSTCG111A) h += row('Tax on STCG 111A @ '+(c.stcg111aRate*100)+'%', fmt(r.taxSTCG111A), 'sub');
    if (r.taxLTCG112A) h += row('Tax on LTCG 112A @ '+(c.ltcg112aRate*100)+'%', fmt(r.taxLTCG112A), 'sub');
    if (r.taxLTCGOther) h += row('Tax on LTCG Other @ '+c.ltcgOtherLabel, fmt(r.taxLTCGOther), 'sub');
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

/* ── CALCULATE ────────────────────────────────────────────────────── */
function calculateTax() {
  const c = cfg();
  const panel = document.getElementById('resultPanel');
  const preInfo = document.getElementById('preCalcInfo');

  if (currentRegime === 'both') {
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
  } else {
    const isNew = currentRegime === 'new';
    const result = computeForRegime(isNew);

    document.getElementById('singleResult').style.display = 'block';
    document.getElementById('compareResult').style.display = 'none';

    const label = isNew ? '🆕 New Regime' : '📜 Old Regime';
    document.getElementById('resultTitle').textContent = 'Tax Computation — ' + label;
    document.getElementById('resultSubtitle').textContent =
      (result.name ? result.name + ' · ' : '') + c.label + (c.isFuture ? ' (Estimated)' : '');

    document.getElementById('resultBody').innerHTML = renderResult(result);
    document.getElementById('slabRegimeLabel').textContent = (isNew ? 'New Regime' : 'Old Regime') + ' · ' + c.label;
    document.getElementById('slabBody').innerHTML = renderSlabs(result);
  }

  // Show/hide future disclaimer
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

// Initialize
setRegime('new');
updateRefSlabs();
</script>
</body></html>"""


# ══════════════════════════════════════════════════════════════════════════════
#  TDS CALCULATOR TEMPLATE
# ══════════════════════════════════════════════════════════════════════════════

TDS_CALC_T = r"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>TDS Calculator (IT Act 2025) – CA Toolkit</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>
<style>
""" + BASE_CSS + """
.nav-links{display:flex;gap:20px;list-style:none}
.nav-links a{text-decoration:none;color:var(--muted);font-size:13px;font-weight:500}
.nav-links a:hover{color:var(--brand)}
.hero{text-align:center;padding:36px 24px 20px;max-width:760px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#ECFDF5;color:#065F46;
            border:1px solid #A7F3D0;border-radius:99px;padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:12px}
h1{font-size:clamp(20px,4vw,32px);font-weight:800;line-height:1.15;letter-spacing:-.5px;margin-bottom:8px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:13px;color:var(--muted);line-height:1.7;max-width:520px;margin:0 auto}
.act-note{max-width:1100px;margin:0 auto;padding:0 24px 12px}
.act-box{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;padding:10px 14px;
         font-size:12px;color:#1e40af;display:flex;align-items:flex-start;gap:6px}
.main{max-width:1100px;margin:0 auto;padding:12px 24px 48px;
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
.result-section{display:none;margin-top:14px}
.rboxes{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px}
.rbox{border-radius:10px;padding:14px 16px}
.rbox-tds  {background:#EFF6FF;border:1.5px solid #BFDBFE}
.rbox-int  {background:#FFFBEB;border:1.5px solid #FDE68A}
.rbox-total{background:#1D4ED8;border:1.5px solid #1D4ED8;grid-column:1/-1}
.rbox .val {font-size:22px;font-weight:800;margin-bottom:2px}
.rbox .lbl {font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:.75}
.rbox .sub {font-size:11px;margin-top:5px;opacity:.8}
.rbox-tds   .val{color:#1D4ED8}.rbox-tds   .lbl{color:#1D4ED8}
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
.rate-table{width:100%;border-collapse:collapse;font-size:11px}
.rate-table th{text-align:left;font-size:10px;letter-spacing:.06em;text-transform:uppercase;
               color:var(--muted);border-bottom:1.5px solid var(--border);padding:5px 6px}
.rate-table td{padding:6px;border-bottom:1px solid var(--border);vertical-align:top;line-height:1.5}
.rate-table tr:last-child td{border:none}
.rate-table tr:hover td{background:#F9FAFB}
.code{background:#EFF6FF;color:var(--brand);font-size:10px;font-weight:700;
      padding:1px 5px;border-radius:4px;font-family:monospace;white-space:nowrap}
footer{background:var(--ink);color:#9CA3AF;text-align:center;padding:20px;font-size:12px}
.footer-brand{color:#D1D5DB;font-weight:700;font-size:14px;margin-bottom:4px}
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
  <h1>TDS Calculator — <em>New IT Act 2025</em></h1>
  <p>Updated for <strong>Section 393 / Section 392</strong> of IT Act 2025 (Tax Year 2026-27). Calculate TDS liability, late deposit interest and total payable amount.</p>
</section>

<div class="act-note">
  <div class="act-box">ℹ️ <span><strong>IT Act 2025 (w.e.f. 1 Apr 2026):</strong> TDS consolidated into Section 393 (non-salary) and Section 392 (salary). Numeric payment codes replace old section numbers in returns. Rates &amp; thresholds unchanged.</span></div>
</div>

<div class="main">

  <!-- LEFT: INPUT FORM -->
  <div>
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#EFF6FF">📑</div>
        <div><h2>TDS Liability Calculator</h2><p>IT Act 2025 · Tax Year 2026-27</p></div>
      </div>
      <div class="card-body">

        <div class="field">
          <label>Nature of Payment</label>
          <select id="tdsSection" onchange="updateHint()">
            <option value="">— Select Payment Type —</option>
            <optgroup label="── Salary ──">
              <option value="1001">Salary (Sec 392) — Slab rate</option>
              <option value="1004">PF Accumulated Balance (Sec 392(7)) — 10%</option>
            </optgroup>
            <optgroup label="── Commission &amp; Brokerage ──">
              <option value="1005">Insurance Commission (Old: 194D) — 2%/10%</option>
              <option value="1006">Commission / Brokerage (Old: 194H) — 2%</option>
            </optgroup>
            <optgroup label="── Rent ──">
              <option value="1008">Rent – Machinery/Plant/Equipment (Old: 194I(a)) — 2%</option>
              <option value="1009">Rent – Land/Building/Furniture (Old: 194I(b)) — 10%</option>
              <option value="1010">Rent by Individual/HUF (Old: 194IB) — 5%</option>
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
              <option value="1017">Investment Fund Income (Old: 194LBB) — 10%</option>
              <option value="1018">Securitisation Trust (Old: 194LBC) — 10%</option>
            </optgroup>
            <optgroup label="── Contractor &amp; Professional ──">
              <option value="1023">Contractor – Individual/HUF (Old: 194C) — 1%</option>
              <option value="1024">Contractor – Others/Company (Old: 194C) — 2%</option>
              <option value="1026">Technical Services/Royalty (Old: 194J(a)) — 2%</option>
              <option value="1027">Professional Fees (Old: 194J(b)) — 10%</option>
              <option value="1028">Director Remuneration (Old: 194J(b)) — 10%</option>
              <option value="1037">Contractor/Prof by Individual/HUF (Old: 194M) — 5%</option>
            </optgroup>
            <optgroup label="── Other Payments ──">
              <option value="1030">Life Insurance Policy Proceeds (Old: 194DA) — 2%</option>
              <option value="1031">Purchase of Goods (Old: 194Q) — 0.1%</option>
              <option value="1033">Benefit/Perquisite – Business (Old: 194R) — 10%</option>
              <option value="1035">E-Commerce Operator (Old: 194O) — 0.1%</option>
              <option value="1038">Cash Withdrawal (Old: 194N) — 2%</option>
              <option value="1039">VDA / Crypto (Old: 194S) — 1%</option>
              <option value="1040">Lottery/Puzzle Winnings (Old: 194B) — 30%</option>
              <option value="1041">Partner Salary/Remuneration (Old: 194T) — 10%</option>
            </optgroup>
          </select>
          <p class="hint" id="sectionHint">Select a payment type to see rate and threshold</p>
        </div>

        <div class="field">
          <label>Payment Amount (₹)</label>
          <input type="number" id="paymentAmt" placeholder="e.g. 100000" min="0"/>
          <p class="hint">Gross payment amount before TDS deduction</p>
        </div>

        <hr style="border:none;border-top:1.5px dashed var(--border);margin:16px 0"/>
        <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-bottom:12px">Late Deposit Interest (Optional)</p>

        <div class="row2">
          <div class="field">
            <label>Date of Deduction</label>
            <input type="date" id="deductionDate"/>
            <p class="hint">When TDS was deducted</p>
          </div>
          <div class="field">
            <label>Date of Actual Deposit</label>
            <input type="date" id="depositDate"/>
            <p class="hint">When you paid the challan</p>
          </div>
        </div>

        <button class="btn" onclick="calcTDS()">Calculate TDS &amp; Interest →</button>

        <!-- RESULT -->
        <div class="result-section" id="resultSection">

          <!-- On-time deposit -->
          <div class="ontime-box" id="ontimeBox" style="display:none">
            ✓ Deposit is on time — No interest applicable
          </div>

          <!-- Late deposit boxes -->
          <div id="lateBoxes" style="display:none">
            <div class="rboxes">
              <div class="rbox rbox-tds">
                <div class="lbl">TDS Amount</div>
                <div class="val" id="r-tds"></div>
                <div class="sub" id="r-tds-sub"></div>
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
              <tr><td>Payment Amount</td><td id="d-payment"></td></tr>
              <tr><td>New Section (IT Act 2025)</td><td id="d-newsec"></td></tr>
              <tr><td>Old Section (for reference)</td><td id="d-oldsec"></td></tr>
              <tr><td>Payment Code</td><td id="d-code"></td></tr>
              <tr><td>TDS Rate</td><td id="d-rate"></td></tr>
              <tr><td>TDS Amount</td><td id="d-tds"></td></tr>
              <tr><td>Date of Deduction</td><td id="d-ddate"></td></tr>
              <tr><td>Due Date for Deposit</td><td id="d-due"></td></tr>
              <tr><td>Actual Deposit Date</td><td id="d-adate"></td></tr>
              <tr><td>Delay (months for interest)</td><td id="d-months"></td></tr>
              <tr><td>Interest Rate</td><td>1.5% per month</td></tr>
              <tr><td>Interest Amount</td><td id="d-intamt"></td></tr>
              <tr><td style="color:var(--brand)">Total Payable (TDS + Interest)</td><td id="d-total" style="color:var(--brand)"></td></tr>
            </table>

            <div class="note-box">
              ⚠ As per IT Act 2025, a fractional month is counted as a full month for interest calculation. Interest runs from date of deduction to actual date of deposit.
            </div>
          </div>

          <!-- No date entered — show basic result -->
          <div id="basicBox" style="display:none">
            <div class="rboxes">
              <div class="rbox rbox-tds" style="grid-column:1/-1">
                <div class="lbl">TDS Amount</div>
                <div class="val" id="b-tds"></div>
                <div class="sub" id="b-sub"></div>
              </div>
            </div>
            <table class="detail-table">
              <tr><td>Payment Amount</td><td id="b-payment"></td></tr>
              <tr><td>New Section (IT Act 2025)</td><td id="b-newsec"></td></tr>
              <tr><td>Old Section (for reference)</td><td id="b-oldsec"></td></tr>
              <tr><td>Payment Code</td><td id="b-code"></td></tr>
              <tr><td>TDS Rate</td><td id="b-rate"></td></tr>
              <tr><td>TDS Amount</td><td id="b-tds2"></td></tr>
              <tr><td>Net Payment to Payee</td><td id="b-net"></td></tr>
            </table>
            <div class="note-box">
              ℹ Enter deduction and deposit dates above to also calculate late deposit interest.
            </div>
          </div>

        </div>
      </div>
    </div>
  </div>

  <!-- RIGHT: RATE CHART + DUE DATE RULES -->
  <div>
    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#FFFBEB">📋</div>
        <div><h2>Quick Rate Chart</h2><p>Section 393 — IT Act 2025</p></div>
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
            <tr><td><span class="code">1008</span></td><td>194I(a)</td><td>Rent (P&amp;M)</td><td>2%</td><td>₹50,000/mo</td></tr>
            <tr><td><span class="code">1009</span></td><td>194I(b)</td><td>Rent (Land/Bldg)</td><td>10%</td><td>₹50,000/mo</td></tr>
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

    <div class="card">
      <div class="card-head">
        <div class="icon" style="background:#F0FDF4">📅</div>
        <div><h2>Deposit Due Dates</h2><p>Rule 218 — IT Rules 2026</p></div>
      </div>
      <div class="card-body">
        <div style="font-size:12px;line-height:2;color:var(--muted)">
          <p><strong style="color:var(--ink)">April – February:</strong> 7th of the following month</p>
          <p><strong style="color:var(--ink)">March deductions:</strong> 30th April</p>
          <p><strong style="color:var(--ink)">Sec 194IA / 194IB / 194M / 194S:</strong> 30 days from end of deduction month</p>
          <p style="margin-top:8px;color:var(--red)"><strong>Late interest:</strong> 1.5% per month · Fractional month = full month</p>
        </div>
      </div>
    </div>
  </div>

</div>

<footer>
  <p class="footer-brand">CA Toolkit</p>
  <p>Built for Indian CAs · Created by CA Articles of GD Singla &amp; Co.</p>
  <p style="margin-top:8px;font-size:11px">© 2026 CA Toolkit · For reference only — verify with latest CBDT circulars</p>
</footer>

<script>
const TDS = {
  "1001":{rate:0,    thresh:0,       label:"Salary",                      newSec:"Sec 392",                    oldSec:"Sec 192",    note:"Slab rate"},
  "1004":{rate:10,   thresh:50000,   label:"PF Accumulated Balance",       newSec:"Sec 392(7)",                 oldSec:"Sec 192A",   note:"No PAN: 20%"},
  "1005":{rate:2,    thresh:20000,   label:"Insurance Commission",         newSec:"Sec 393(1) Sl.1(i)",         oldSec:"Sec 194D",   note:"Ind: 2%, Others: 10%"},
  "1006":{rate:2,    thresh:20000,   label:"Commission/Brokerage",         newSec:"Sec 393(1) Sl.1(ii)",        oldSec:"Sec 194H",   note:""},
  "1008":{rate:2,    thresh:50000,   label:"Rent – Machinery/Plant",       newSec:"Sec 393(1) Sl.2(ii).D(a)",  oldSec:"Sec 194I(a)",note:"Monthly threshold"},
  "1009":{rate:10,   thresh:50000,   label:"Rent – Land/Building",         newSec:"Sec 393(1) Sl.2(ii).D(b)",  oldSec:"Sec 194I(b)",note:"Monthly threshold"},
  "1010":{rate:5,    thresh:50000,   label:"Rent by Ind/HUF",              newSec:"Sec 393(1) Sl.2(i)",         oldSec:"Sec 194IB",  note:"Per month"},
  "1011":{rate:10,   thresh:0,       label:"JDA Consideration",            newSec:"Sec 393(1) Sl.3(ii)",        oldSec:"Sec 194IC",  note:""},
  "1012":{rate:10,   thresh:500000,  label:"Land Acquisition Comp.",       newSec:"Sec 393(1) Sl.3(iii)",       oldSec:"Sec 194LA",  note:"Threshold ₹5L"},
  "1013":{rate:10,   thresh:10000,   label:"Mutual Fund Units",            newSec:"Sec 393(1) Sl.4(i)",         oldSec:"Sec 194K",   note:""},
  "1014":{rate:10,   thresh:0,       label:"Business Trust – Interest",    newSec:"Sec 393(1) Sl.4(ii)",        oldSec:"Sec 194LBA", note:""},
  "1017":{rate:10,   thresh:0,       label:"Investment Fund Income",       newSec:"Sec 393(1) Sl.4(iii)",       oldSec:"Sec 194LBB", note:""},
  "1018":{rate:10,   thresh:0,       label:"Securitisation Trust",         newSec:"Sec 393(1) Sl.4(iv)",        oldSec:"Sec 194LBC", note:""},
  "1019":{rate:10,   thresh:10000,   label:"Interest on Securities",       newSec:"Sec 393(1) Sl.5(i)",         oldSec:"Sec 193",    note:""},
  "1020":{rate:10,   thresh:100000,  label:"Interest – Senior Citizen",    newSec:"Sec 393(1) Sl.5(ii).D(a)",  oldSec:"Sec 194A",   note:"Threshold ₹1L"},
  "1021":{rate:10,   thresh:50000,   label:"Interest – Bank/Post Office",  newSec:"Sec 393(1) Sl.5(ii).D(b)",  oldSec:"Sec 194A",   note:"Threshold ₹50K"},
  "1022":{rate:10,   thresh:10000,   label:"Interest – Others",            newSec:"Sec 393(1) Sl.5(iii)",       oldSec:"Sec 194A",   note:"Threshold ₹10K"},
  "1023":{rate:1,    thresh:30000,   label:"Contractor – Ind/HUF",         newSec:"Sec 393(1) Sl.6(i).D(a)",   oldSec:"Sec 194C",   note:"Single ₹30K / Annual ₹1L"},
  "1024":{rate:2,    thresh:30000,   label:"Contractor – Others",          newSec:"Sec 393(1) Sl.6(i).D(b)",   oldSec:"Sec 194C",   note:"Single ₹30K / Annual ₹1L"},
  "1026":{rate:2,    thresh:50000,   label:"Technical Services/Royalty",   newSec:"Sec 393(1) Sl.6(iii).D(a)", oldSec:"Sec 194J(a)",note:""},
  "1027":{rate:10,   thresh:50000,   label:"Professional Fees",            newSec:"Sec 393(1) Sl.6(iii).D(b)", oldSec:"Sec 194J(b)",note:""},
  "1028":{rate:10,   thresh:0,       label:"Director Remuneration",        newSec:"Sec 393(1) Sl.6(iii).D(b)", oldSec:"Sec 194J(b)",note:"No threshold"},
  "1029":{rate:10,   thresh:10000,   label:"Dividends",                    newSec:"Sec 393(1) Sl.7",            oldSec:"Sec 194",    note:""},
  "1030":{rate:2,    thresh:100000,  label:"Life Insurance Proceeds",      newSec:"Sec 393(1) Sl.8(i)",         oldSec:"Sec 194DA",  note:"On taxable portion"},
  "1031":{rate:0.1,  thresh:5000000, label:"Purchase of Goods",            newSec:"Sec 393(1) Sl.8(ii)",        oldSec:"Sec 194Q",   note:"Annual > ₹50L"},
  "1033":{rate:10,   thresh:20000,   label:"Benefit/Perquisite",           newSec:"Sec 393(1) Sl.8(iv)",        oldSec:"Sec 194R",   note:""},
  "1035":{rate:0.1,  thresh:500000,  label:"E-Commerce Operator",          newSec:"Sec 393(1) Sl.8(vi)",        oldSec:"Sec 194O",   note:"Annual > ₹5L"},
  "1036":{rate:1,    thresh:5000000, label:"Purchase of Immovable Property",newSec:"Sec 393(1) Sl.3(i)",        oldSec:"Sec 194IA",  note:"Threshold ₹50L"},
  "1037":{rate:5,    thresh:5000000, label:"Contractor/Prof by Ind/HUF",   newSec:"Sec 393(1) Sl.6(iv)",        oldSec:"Sec 194M",   note:"Annual > ₹50L"},
  "1038":{rate:2,    thresh:2000000, label:"Cash Withdrawal",              newSec:"Sec 393(1) Sl.8(vii)",       oldSec:"Sec 194N",   note:"3% if no ITR filed"},
  "1039":{rate:1,    thresh:10000,   label:"VDA/Crypto",                   newSec:"Sec 393(1) Sl.8(viii)",      oldSec:"Sec 194S",   note:"₹50K for specified persons"},
  "1040":{rate:30,   thresh:10000,   label:"Lottery/Puzzle Winnings",      newSec:"Sec 393(1) Sl.8(ix)",        oldSec:"Sec 194B",   note:""},
  "1041":{rate:10,   thresh:20000,   label:"Partner Salary/Remuneration",  newSec:"Sec 393(1) Sl.6(v)",         oldSec:"Sec 194T",   note:"Threshold ₹20K pa"},
};

// Special due date sections (30 days from end of month)
const SPECIAL_30 = ["1036","1010","1037","1039"];

function fmt(n){ return "₹" + Math.round(n).toLocaleString("en-IN"); }

function updateHint(){
  const code = document.getElementById("tdsSection").value;
  const el   = document.getElementById("sectionHint");
  if(!code){ el.textContent="Select a payment type to see rate and threshold"; return; }
  const d = TDS[code];
  if(!d) return;
  el.textContent = (d.rate===0?"Rate: Slab rate":"Rate: "+d.rate+"%")
    + (d.thresh?" · Threshold: "+fmt(d.thresh):" · No threshold")
    + (d.note?" · "+d.note:"");
}

function getDueDate(deductionDate, code){
  const d = new Date(deductionDate);
  const month = d.getMonth(); // 0-indexed
  const year  = d.getFullYear();

  if(SPECIAL_30.includes(code)){
    // 30 days from end of deduction month
    const endOfMonth = new Date(year, month+1, 0);
    return new Date(endOfMonth.getTime() + 30*24*60*60*1000);
  }
  // March → 30 April
  if(month === 2){
    return new Date(year, 3, 30); // April 30
  }
  // Otherwise → 7th of next month
  return new Date(year, month+1, 7);
}

function calcMonthsLate(dueDate, depositDate){
  // Count months (fractional = full month)
  const due     = new Date(dueDate);
  const deposit = new Date(depositDate);
  if(deposit <= due) return 0;

  // From date of deduction to deposit date (as per sec 201(1A))
  // months = ceil of difference in days / 30 approximately
  // Exact: count month boundaries crossed
  let months = 0;
  let cur = new Date(due);
  while(cur < deposit){
    cur.setMonth(cur.getMonth()+1);
    months++;
  }
  return months;
}

function calcTDS(){
  const code    = document.getElementById("tdsSection").value;
  const amt     = parseFloat(document.getElementById("paymentAmt").value);
  const dDate   = document.getElementById("deductionDate").value;
  const aDate   = document.getElementById("depositDate").value;

  if(!code){ alert("Please select a payment type."); return; }
  if(!amt || amt<=0){ alert("Please enter a valid payment amount."); return; }

  const d = TDS[code];
  if(!d){ alert("Section data not found."); return; }

  // Check threshold
  const belowThresh = d.thresh && amt < d.thresh;
  const tds   = belowThresh ? 0 : (d.rate===0 ? 0 : Math.round(amt * d.rate / 100));
  const net   = amt - tds;

  document.getElementById("resultSection").style.display = "block";

  // No dates entered — show basic result only
  if(!dDate || !aDate){
    document.getElementById("ontimeBox").style.display  = "none";
    document.getElementById("lateBoxes").style.display  = "none";
    document.getElementById("basicBox").style.display   = "block";

    document.getElementById("b-tds").textContent     = belowThresh?"No TDS":d.rate===0?"Slab Rate":fmt(tds);
    document.getElementById("b-sub").textContent     = belowThresh?"Below threshold of "+fmt(d.thresh):d.rate===0?"Compute at applicable slab":d.rate+"% on "+fmt(amt);
    document.getElementById("b-payment").textContent  = fmt(amt);
    document.getElementById("b-newsec").textContent   = d.newSec;
    document.getElementById("b-oldsec").textContent   = d.oldSec+" (ref only)";
    document.getElementById("b-code").textContent     = code;
    document.getElementById("b-rate").textContent     = d.rate===0?"Slab rate":d.rate+"%";
    document.getElementById("b-tds2").textContent     = belowThresh?"Nil (below threshold)":d.rate===0?"As per slab":fmt(tds);
    document.getElementById("b-net").textContent      = fmt(net);
    return;
  }

  // Both dates entered — calculate interest
  const dueDate    = getDueDate(dDate, code);
  const depositDt  = new Date(aDate);
  const deductDt   = new Date(dDate);

  const monthsLate = calcMonthsLate(dueDate, depositDt);
  const interest   = Math.round(tds * 0.015 * monthsLate);
  const total      = tds + interest;
  const isOnTime   = monthsLate === 0;

  const fmtDate = dt => dt.toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
  const dueDateStr = fmtDate(dueDate);

  document.getElementById("basicBox").style.display = "none";

  if(isOnTime || belowThresh || d.rate===0){
    document.getElementById("ontimeBox").style.display = "block";
    document.getElementById("lateBoxes").style.display = "none";
    if(belowThresh) document.getElementById("ontimeBox").textContent = "No TDS applicable — Payment below threshold of "+fmt(d.thresh);
    else if(d.rate===0) document.getElementById("ontimeBox").textContent = "Salary TDS — compute at applicable slab rate on estimated annual income";
    else document.getElementById("ontimeBox").textContent = "✓ Deposit is on time — No interest applicable. TDS: "+fmt(tds);
    return;
  }

  document.getElementById("ontimeBox").style.display = "none";
  document.getElementById("lateBoxes").style.display = "block";

  document.getElementById("r-tds").textContent      = fmt(tds);
  document.getElementById("r-tds-sub").textContent  = d.rate+"% on "+fmt(amt);
  document.getElementById("r-int").textContent      = fmt(interest);
  document.getElementById("r-int-sub").textContent  = "1.5% × "+monthsLate+" month"+(monthsLate>1?"s":"");
  document.getElementById("r-total").textContent    = fmt(total);
  document.getElementById("r-total-sub").textContent= "TDS "+fmt(tds)+" + Interest "+fmt(interest);

  document.getElementById("d-payment").textContent  = fmt(amt);
  document.getElementById("d-newsec").textContent   = d.newSec;
  document.getElementById("d-oldsec").textContent   = d.oldSec+" (ref only)";
  document.getElementById("d-code").textContent     = code;
  document.getElementById("d-rate").textContent     = d.rate+"%";
  document.getElementById("d-tds").textContent      = fmt(tds);
  document.getElementById("d-ddate").textContent    = fmtDate(deductDt);
  document.getElementById("d-due").textContent      = dueDateStr;
  document.getElementById("d-adate").textContent    = fmtDate(depositDt);
  document.getElementById("d-months").textContent   = monthsLate+" month"+(monthsLate>1?"s":"")+" (fractional months counted as full)";
  document.getElementById("d-intamt").textContent   = fmt(interest);
  document.getElementById("d-total").textContent    = fmt(total);
}
</script>
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
footer{background:var(--ink);color:#9CA3AF;text-align:center;padding:20px;font-size:12px}
.footer-brand{color:#D1D5DB;font-weight:700;font-size:14px;margin-bottom:4px}
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
  <p class="footer-brand">CA Toolkit</p>
  <p>Built for Indian Chartered Accountants · Created by CA Articles of GD Singla &amp; Co.</p>
  <p style="margin-top:10px;font-size:11px">© 2026 CA Toolkit · For reference only — verify with latest MCA/CBDT notifications</p>
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
              <option value="starter">Starter (10 uploads · ₹60)</option>
              <option value="standard" selected>Standard (25 uploads · ₹130)</option>
              <option value="pro">Professional (60 uploads · ₹270)</option>
              <option value="firm">Firm (150 uploads · ₹600)</option>
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
