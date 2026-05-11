"""
BS Annual Updater  –  with Auth & Plan Enforcement
Run:  python app.py  |  Then open http://localhost:5000
"""

import re
import os
import uuid
import sqlite3
import hashlib
import secrets
from datetime import datetime
from functools import wraps
from flask import (Flask, request, send_file, jsonify,
                   render_template_string, session, redirect, url_for, g)
from processor import process

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

UPLOAD_FOLDER = "/tmp/bs_uploads"
OUTPUT_FOLDER = "/tmp/bs_outputs"
DB_PATH       = os.environ.get("DB_PATH", "users.db")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

FREE_LIMIT     = 2
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "sumit_admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "Admin@Secure123")

# ─── Database ─────────────────────────────────────────────────────────────────

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH, detect_types=sqlite3.PARSE_DECLTYPES)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db:
        db.close()

def init_db():
    db = sqlite3.connect(DB_PATH)
    db.execute("""CREATE TABLE IF NOT EXISTS users (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        username   TEXT    UNIQUE NOT NULL,
        password   TEXT    NOT NULL,
        plan       TEXT    NOT NULL DEFAULT 'free',
        is_admin   INTEGER NOT NULL DEFAULT 0,
        created_at TEXT    NOT NULL)""")
    db.execute("""CREATE TABLE IF NOT EXISTS usage_log (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id      INTEGER NOT NULL,
        processed_at TEXT    NOT NULL,
        FOREIGN KEY(user_id) REFERENCES users(id))""")
    db.execute("""INSERT OR IGNORE INTO users
        (username, password, plan, is_admin, created_at) VALUES (?,?,'pro',1,?)""",
        (ADMIN_USERNAME, _hash(ADMIN_PASSWORD), datetime.utcnow().isoformat()))
    db.commit()
    db.close()

def _hash(p): return hashlib.sha256(p.encode("utf-8")).hexdigest()

def get_user_by_name(u): return get_db().execute("SELECT * FROM users WHERE username=?", (u,)).fetchone()
def get_user_by_id(i):   return get_db().execute("SELECT * FROM users WHERE id=?", (i,)).fetchone()

def files_this_month(uid):
    start = datetime.utcnow().replace(day=1,hour=0,minute=0,second=0,microsecond=0).isoformat()
    r = get_db().execute("SELECT COUNT(*) AS c FROM usage_log WHERE user_id=? AND processed_at>=?", (uid, start)).fetchone()
    return r["c"] if r else 0

def log_usage(uid):
    db = get_db()
    db.execute("INSERT INTO usage_log (user_id, processed_at) VALUES (?,?)", (uid, datetime.utcnow().isoformat()))
    db.commit()

def all_users(): return get_db().execute("SELECT * FROM users ORDER BY id").fetchall()

def set_plan(uid, plan):
    db = get_db(); db.execute("UPDATE users SET plan=? WHERE id=?", (plan, uid)); db.commit()

def del_user(uid):
    db = get_db()
    db.execute("DELETE FROM usage_log WHERE user_id=?", (uid,))
    db.execute("DELETE FROM users WHERE id=?", (uid,))
    db.commit()

# ─── Auth decorators ──────────────────────────────────────────────────────────

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

# ─── CSS ──────────────────────────────────────────────────────────────────────

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=DM+Sans:wght@300;400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--ink:#1a1a2e;--paper:#f5f0e8;--cream:#ede8dc;--gold:#c9a84c;--rust:#b5451b;--sage:#4a7c6f;--sh:rgba(26,26,46,.12)}
body{font-family:'DM Sans',sans-serif;background:var(--paper);color:var(--ink);min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:48px 20px 80px}
body::before{content:'';position:fixed;inset:0;background:repeating-linear-gradient(to bottom,transparent 0px,transparent 31px,rgba(201,168,76,.15) 32px);pointer-events:none;z-index:0}
.card{position:relative;z-index:1;background:#fff;border:1.5px solid var(--cream);border-radius:4px;box-shadow:0 2px 0 var(--gold),0 8px 32px var(--sh);width:100%;max-width:480px;padding:48px 52px}
.wide{max-width:760px}
.stamp{position:absolute;top:-18px;right:36px;background:var(--rust);color:#fff;font-size:10px;font-weight:500;letter-spacing:.12em;text-transform:uppercase;padding:4px 14px;border-radius:2px}
.stamp-pro{background:var(--gold);color:var(--ink)}
.stamp-adm{background:var(--sage)}
h1{font-family:'Playfair Display',serif;font-size:26px;font-weight:900;color:var(--ink);margin-bottom:6px}
.sub{font-size:13px;color:#888;margin-bottom:32px}
label{display:block;font-size:11px;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:var(--sage);margin-bottom:8px}
.field{margin-bottom:20px}
input[type=text],input[type=password],input[type=number],input[type=file],select{width:100%;border:1.5px solid var(--cream);border-radius:3px;padding:11px 14px;font-family:'DM Sans',sans-serif;font-size:14px;color:var(--ink);background:var(--paper);transition:border-color .2s;outline:none}
input:focus,select:focus{border-color:var(--gold);background:#fff}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.hint{font-size:11px;color:#aaa;margin-top:5px}
hr.div{border:none;border-top:1.5px dashed var(--cream);margin:28px 0}
.btn{width:100%;background:var(--ink);color:#fff;border:none;border-radius:3px;padding:13px;font-family:'Playfair Display',serif;font-size:15px;font-weight:700;cursor:pointer;transition:background .2s;position:relative;overflow:hidden}
.btn:hover{background:var(--rust)}
.btn:disabled{background:#ccc;cursor:not-allowed}
.sm{width:auto;padding:5px 12px;font-family:'DM Sans',sans-serif;font-size:12px;font-weight:500;border-radius:3px;border:none;cursor:pointer}
.sage{background:var(--sage);color:#fff}.sage:hover{background:#3a6559}
.rr{background:var(--rust);color:#fff}.rr:hover{background:#8a3010}
.spinner{display:none;width:18px;height:18px;border:2.5px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite;position:absolute;right:18px;top:50%;transform:translateY(-50%)}
@keyframes spin{to{transform:translateY(-50%) rotate(360deg)}}
#status{margin-top:20px;border-radius:3px;padding:14px 16px;font-size:13px;display:none}
#status.ok{background:#eef7f2;border:1.5px solid #a8d5bc;color:var(--sage)}
#status.er{background:#fdf0ed;border:1.5px solid #f0b8a8;color:var(--rust)}
.log-list{margin-top:10px;padding-left:16px;font-size:12px;line-height:1.8;color:#555}
.dlbtn{display:none;margin-top:14px;width:100%;background:var(--sage);color:#fff;border:none;border-radius:3px;padding:13px;font-family:'DM Sans',sans-serif;font-size:14px;font-weight:500;cursor:pointer;text-decoration:none;text-align:center;transition:background .2s}
.dlbtn:hover{background:#3a6559}
.alert{padding:10px 14px;border-radius:3px;font-size:13px;margin-bottom:20px}
.ae{background:#fdf0ed;border:1.5px solid #f0b8a8;color:var(--rust)}
.as{background:#eef7f2;border:1.5px solid #a8d5bc;color:var(--sage)}
.lr{text-align:center;margin-top:20px;font-size:13px;color:#888}
.lr a{color:var(--sage);text-decoration:none;font-weight:500}
.lr a:hover{text-decoration:underline}
.nav{position:relative;z-index:2;width:100%;max-width:760px;display:flex;justify-content:space-between;align-items:center;margin-bottom:24px;font-size:13px;color:#888}
.nav a{color:var(--sage);text-decoration:none;font-weight:500}
.nav a:hover{text-decoration:underline}
.badge{display:inline-block;font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;padding:3px 10px;border-radius:20px;background:var(--cream);color:var(--ink)}
.badge.pro{background:var(--gold);color:var(--ink)}
.ub-wrap{margin-bottom:24px}
.ub-bg{background:var(--cream);border-radius:3px;height:6px;margin-top:6px;overflow:hidden}
.ub-fill{background:var(--sage);height:100%;border-radius:3px;transition:width .4s}
.ub-full{background:var(--rust)}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--sage);border-bottom:1.5px solid var(--cream);padding:8px 10px}
td{padding:10px;border-bottom:1px solid var(--cream);vertical-align:middle}
tr:last-child td{border-bottom:none}
.wm{position:absolute;bottom:14px;right:20px;font-size:10px;color:var(--gold);opacity:.5;font-family:'Playfair Display',serif}
@media(max-width:540px){.card{padding:36px 24px}.row2{grid-template-columns:1fr}}
"""

# ─── Routes — Auth ────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET"])
def login_page():
    if "uid" in session: return redirect(url_for("index"))
    return render_template_string(LOGIN_T, css=CSS, error=None)

@app.route("/login", methods=["POST"])
def login_post():
    u = request.form.get("username","").strip()
    p = request.form.get("password","")
    user = get_user_by_name(u)
    if not user or user["password"] != _hash(p):
        return render_template_string(LOGIN_T, css=CSS, error="Invalid username or password.")
    session.clear(); session["uid"] = user["id"]
    return redirect(url_for("index"))

@app.route("/register", methods=["GET"])
def register_page():
    if "uid" in session: return redirect(url_for("index"))
    return render_template_string(REG_T, css=CSS, fl=FREE_LIMIT, error=None, success=None)

@app.route("/register", methods=["POST"])
def register_post():
    u  = request.form.get("username","").strip()
    p  = request.form.get("password","")
    c  = request.form.get("confirm","")
    def err(m): return render_template_string(REG_T, css=CSS, fl=FREE_LIMIT, error=m, success=None)
    if len(u) < 3:                              return err("Username must be at least 3 characters.")
    if not re.match(r"^[a-zA-Z0-9_]+$", u):    return err("Only letters, numbers, underscores allowed.")
    if len(p) < 8:                              return err("Password must be at least 8 characters.")
    if p != c:                                  return err("Passwords do not match.")
    if get_user_by_name(u):                     return err("Username already taken.")
    db = get_db()
    db.execute("INSERT INTO users (username,password,plan,is_admin,created_at) VALUES (?,?,'free',0,?)",
               (u, _hash(p), datetime.utcnow().isoformat()))
    db.commit()
    return render_template_string(REG_T, css=CSS, fl=FREE_LIMIT, error=None,
                                  success="Account created! You can now sign in.")

@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("login_page"))

# ─── Routes — Tool ────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    user = get_user_by_id(session["uid"])
    used = files_this_month(user["id"])
    return render_template_string(TOOL_T, css=CSS,
        username=user["username"], plan=user["plan"],
        is_admin=bool(user["is_admin"]), used=used, fl=FREE_LIMIT)

@app.route("/process", methods=["POST"])
@login_required
def process_file():
    user = get_user_by_id(session["uid"])
    if user["plan"] == "free" and not user["is_admin"]:
        if files_this_month(user["id"]) >= FREE_LIMIT:
            return jsonify({"status":"error",
                "message":f"Free plan limit ({FREE_LIMIT} files/month) reached. Contact admin to upgrade."})
    if "file" not in request.files:
        return jsonify({"status":"error","message":"No file uploaded."})
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"status":"error","message":"Only .xlsx files are supported."})
    try:
        cy = int(request.form.get("closing_year", 0))
        ny = int(request.form.get("new_year", 0))
        on = request.form.get("output_name","").strip()
    except ValueError:
        return jsonify({"status":"error","message":"Invalid year values."})
    if ny != cy + 1:
        return jsonify({"status":"error","message":"New year must be closing year + 1."})
    h = uuid.uuid4().hex
    ip = os.path.join(UPLOAD_FOLDER, f"{h}_input.xlsx")
    op = os.path.join(OUTPUT_FOLDER, f"{h}_output.xlsx")
    f.save(ip)
    fname = f"{on or os.path.splitext(f.filename)[0]}_{ny}.xlsx"
    try:
        result = process(ip, op, cy, ny)
    except Exception as e:
        return jsonify({"status":"error","message":str(e)})
    finally:
        try: os.remove(ip)
        except: pass
    log_usage(user["id"])
    return jsonify({"status":"success","log":result["log"],"file_id":h,"filename":fname})

@app.route("/download/<fid>")
@login_required
def download(fid):
    if not re.fullmatch(r"[a-f0-9]{32}", fid): return "Invalid ID", 400
    path = os.path.join(OUTPUT_FOLDER, f"{fid}_output.xlsx")
    if not os.path.exists(path): return "File not found or expired.", 404
    return send_file(path, as_attachment=True,
        download_name=f"balance_sheet_{fid[:8]}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ─── Routes — Admin ───────────────────────────────────────────────────────────

@app.route("/admin")
@admin_required
def admin_panel():
    users = [dict(u) | {"used": files_this_month(u["id"])} for u in all_users()]
    return render_template_string(ADMIN_T, css=CSS, users=users, msg=request.args.get("msg"))

@app.route("/admin/upgrade", methods=["POST"])
@admin_required
def admin_upgrade():
    uid  = int(request.form.get("uid"))
    plan = request.form.get("plan")
    if plan not in ("free","pro"): return "Invalid plan.", 400
    set_plan(uid, plan)
    return redirect(url_for("admin_panel", msg=f"User #{uid} set to {plan}."))

@app.route("/admin/delete", methods=["POST"])
@admin_required
def admin_delete():
    uid = int(request.form.get("uid"))
    if uid == session["uid"]:
        return redirect(url_for("admin_panel", msg="Cannot delete your own account."))
    del_user(uid)
    return redirect(url_for("admin_panel", msg=f"User #{uid} deleted."))

# ─── HTML Templates ───────────────────────────────────────────────────────────

LOGIN_T = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Login – Balance Sheet Tool</title><style>{{ css }}</style></head><body>
<div class="card">
  <span class="stamp">Accounts Utility</span>
  <h1>Welcome Back</h1>
  <p class="sub">Sign in to access the BS Annual Updater.</p>
  {% if error %}<div class="alert ae">{{ error }}</div>{% endif %}
  <form method="POST" action="/login">
    <div class="field"><label>Username</label>
      <input type="text" name="username" placeholder="Enter username" required autocomplete="username"/></div>
    <div class="field"><label>Password</label>
      <input type="password" name="password" placeholder="Enter password" required autocomplete="current-password"/></div>
    <button class="btn" type="submit">Sign In</button>
  </form>
  <div class="lr">No account? <a href="/register">Register free</a></div>
  <span class="wm">BS Annual Updater v2.0</span>
</div></body></html>"""

REG_T = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Register – Balance Sheet Tool</title><style>{{ css }}</style></head><body>
<div class="card">
  <span class="stamp">New Account</span>
  <h1>Create Account</h1>
  <p class="sub">Free plan: {{ fl }} files per month. Upgrade anytime.</p>
  {% if error %}<div class="alert ae">{{ error }}</div>{% endif %}
  {% if success %}<div class="alert as">{{ success }}</div>{% endif %}
  <form method="POST" action="/register">
    <div class="field"><label>Username</label>
      <input type="text" name="username" placeholder="Choose a username" required autocomplete="username"/></div>
    <div class="field"><label>Password</label>
      <input type="password" name="password" placeholder="Min 8 characters" required autocomplete="new-password"/></div>
    <div class="field"><label>Confirm Password</label>
      <input type="password" name="confirm" placeholder="Repeat password" required autocomplete="new-password"/></div>
    <button class="btn" type="submit">Create Account</button>
  </form>
  <div class="lr">Already have an account? <a href="/login">Sign in</a></div>
  <span class="wm">BS Annual Updater v2.0</span>
</div></body></html>"""

TOOL_T = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>BS Annual Updater</title><style>{{ css }}</style></head><body>
<div class="nav">
  <span>👤 <strong>{{ username }}</strong>
    <span class="badge {% if plan=='pro' %}pro{% endif %}">{{ plan }}</span>
    {% if is_admin %}&nbsp;<span class="badge" style="background:var(--sage);color:#fff;">admin</span>{% endif %}
  </span>
  <span>{% if is_admin %}<a href="/admin" style="margin-right:16px;">Admin Panel</a>{% endif %}<a href="/logout">Sign Out</a></span>
</div>
<div class="card wide" style="max-width:640px;">
  <span class="stamp {% if plan=='pro' %}stamp-pro{% endif %}">{% if plan=='pro' %}Pro Plan{% else %}Free Plan{% endif %}</span>
  <h1>BS Annual<br>Updater</h1>
  <p class="sub">Rolls over CY → PY and prepares a blank CY column. Works for any financial year.</p>
  {% if plan=='free' %}
  <div class="ub-wrap">
    <div style="display:flex;justify-content:space-between;font-size:12px;color:#888;">
      <span>Files used this month</span><span><strong>{{ used }}</strong> / {{ fl }}</span>
    </div>
    <div class="ub-bg"><div class="ub-fill {% if used>=fl %}ub-full{% endif %}"
      style="width:{{ [used*100//fl,100]|min }}%"></div></div>
    {% if used>=fl %}
    <p style="font-size:12px;color:var(--rust);margin-top:6px;">⚠ Monthly limit reached. Contact admin to upgrade to Pro.</p>
    {% endif %}
  </div>
  {% endif %}
  <div class="field"><label>Excel File (.xlsx)</label>
    <input type="file" id="xlFile" accept=".xlsx" {% if plan=='free' and used>=fl %}disabled{% endif %}/>
    <p class="hint">Upload the comparative balance sheet workbook</p></div>
  <div class="row2">
    <div class="field"><label>Closing Year (CY ending)</label>
      <input type="number" id="closingYear"  min="2000" max="2100" {% if plan=='free' and used>=fl %}disabled{% endif %}/>
      <p class="hint">e.g. 2024 for data as at 31.03.2024</p></div>
    <div class="field"><label>New Year (new CY)</label>
      <input type="number" id="newYear"  min="2000" max="2100" {% if plan=='free' and used>=fl %}disabled{% endif %}/>
      <p class="hint">e.g. 2025 → new blank CY column</p></div>
  </div>
  <div class="field"><label>Output Filename</label>
    <input type="text" id="outputName" placeholder="e.g. ClientName_BalanceSheet" {% if plan=='free' and used>=fl %}disabled{% endif %}/>
    <p class="hint">Do not include .xlsx extension</p></div>
  <hr class="div">
  <button class="btn" id="processBtn" onclick="processFile()" {% if plan=='free' and used>=fl %}disabled{% endif %}>
    Process &amp; Download<div class="spinner" id="spinner"></div></button>
  <div id="status"></div>
  <a id="dlLink" class="dlbtn">⬇ Download Processed File</a>
  <span class="wm">BS Annual Updater v2.0</span>
</div>
<script>
async function processFile(){
  const fi=document.getElementById('xlFile'),
        cy=parseInt(document.getElementById('closingYear').value),
        ny=parseInt(document.getElementById('newYear').value),
        on=document.getElementById('outputName').value.trim(),
        btn=document.getElementById('processBtn'),
        sp=document.getElementById('spinner'),
        dl=document.getElementById('dlLink');
  if(!fi.files.length){show('er','Please select an Excel file first.');return;}
  if(isNaN(cy)||isNaN(ny)||ny!==cy+1){show('er','New Year must be exactly one year after Closing Year.');return;}
  btn.disabled=true;sp.style.display='block';dl.style.display='none';show('','');
  const fd=new FormData();
  fd.append('file',fi.files[0]);fd.append('closing_year',cy);fd.append('new_year',ny);fd.append('output_name',on);
  try{
    const r=await fetch('/process',{method:'POST',body:fd}),d=await r.json();
    if(d.status==='success'){
      show('ok','✓ File processed successfully.<ul class="log-list">'+d.log.map(l=>`<li>${l}</li>`).join('')+'</ul>');
      dl.href='/download/'+d.file_id;dl.download=d.filename;
      dl.textContent='⬇  Download — '+d.filename;dl.style.display='block';
      setTimeout(()=>location.reload(),3000);
    }else{show('er','✗ '+d.message);}
  }catch(e){show('er','✗ Network error: '+e.message);}
  finally{btn.disabled=false;sp.style.display='none';}
}
function show(t,m){const e=document.getElementById('status');e.className=t;e.innerHTML=m;e.style.display=m?'block':'none';}
document.getElementById('closingYear').addEventListener('input',function(){
  const v=parseInt(this.value);if(!isNaN(v))document.getElementById('newYear').value=v+1;});
</script></body></html>"""

ADMIN_T = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Admin – Balance Sheet Tool</title><style>{{ css }}</style></head><body>
<div class="nav" style="max-width:760px;">
  <span>⚙ Admin Panel</span>
  <span><a href="/">Tool</a> &nbsp;|&nbsp; <a href="/logout">Sign Out</a></span>
</div>
<div class="card wide">
  <span class="stamp stamp-adm">Admin</span>
  <h1>User Management</h1>
  <p class="sub">Manage accounts and plan assignments.</p>
  {% if msg %}<div class="alert as">{{ msg }}</div>{% endif %}
  <table>
    <thead><tr><th>#</th><th>Username</th><th>Plan</th><th>Role</th><th>Joined</th><th>Used/mo</th><th>Actions</th></tr></thead>
    <tbody>
    {% for u in users %}
    <tr>
      <td style="color:#aaa">{{ u.id }}</td>
      <td><strong>{{ u.username }}</strong></td>
      <td><span class="badge {% if u.plan=='pro' %}pro{% endif %}">{{ u.plan }}</span></td>
      <td>{{ 'Admin' if u.is_admin else 'User' }}</td>
      <td style="color:#aaa;font-size:12px">{{ u.created_at[:10] }}</td>
      <td style="text-align:center">{{ u.used }}</td>
      <td>
        {% if not u.is_admin %}
          {% if u.plan=='free' %}
          <form method="POST" action="/admin/upgrade" style="display:inline">
            <input type="hidden" name="uid" value="{{ u.id }}"><input type="hidden" name="plan" value="pro">
            <button class="sm sage" type="submit">→ Pro</button></form>
          {% else %}
          <form method="POST" action="/admin/upgrade" style="display:inline">
            <input type="hidden" name="uid" value="{{ u.id }}"><input type="hidden" name="plan" value="free">
            <button class="sm" type="submit" style="background:var(--cream);color:var(--ink)">→ Free</button></form>
          {% endif %}
          &nbsp;
          <form method="POST" action="/admin/delete" style="display:inline"
                onsubmit="return confirm('Delete {{ u.username }}?')">
            <input type="hidden" name="uid" value="{{ u.id }}">
            <button class="sm rr" type="submit">Delete</button></form>
        {% else %}<span style="font-size:12px;color:#aaa">—</span>{% endif %}
      </td>
    </tr>
    {% endfor %}
    </tbody>
  </table>
  <span class="wm">BS Annual Updater v2.0</span>
</div></body></html>"""

# ─── Init & Run ───────────────────────────────────────────────────────────────

init_db()

if __name__ == "__main__":
    app.run(debug=True, port=5000)
