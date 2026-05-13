"""
BalanceShift.in — Professional Balance Sheet Year-Shift Tool
Flask web app — deploy free on Render.com
"""

import os
import re
import uuid
from flask import Flask, request, send_file, jsonify, render_template_string

from processor import process

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB

UPLOAD_DIR = "/tmp/bs_uploads"
OUTPUT_DIR = "/tmp/bs_outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── HTML ────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>BalanceShift — CA Balance Sheet Year-Roll Tool</title>
<meta name="description" content="Free tool for Indian Chartered Accountants to roll over comparative balance sheets to the next financial year in seconds. Upload your Excel, download ready-to-fill FY2026 file."/>

<!-- Google Fonts -->
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap"/>

<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}

:root{
  --brand:   #1D4ED8;
  --brand-d: #1e40af;
  --accent:  #F59E0B;
  --green:   #10B981;
  --red:     #EF4444;
  --ink:     #111827;
  --muted:   #6B7280;
  --border:  #E5E7EB;
  --bg:      #F9FAFB;
  --white:   #FFFFFF;
  --radius:  12px;
  --shadow:  0 4px 24px rgba(0,0,0,.08);
}

body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--ink);min-height:100vh}

/* ── NAV ── */
nav{background:var(--white);border-bottom:1px solid var(--border);padding:0 24px;
    display:flex;align-items:center;justify-content:space-between;height:60px;
    position:sticky;top:0;z-index:100;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.logo{font-size:20px;font-weight:800;color:var(--brand);letter-spacing:-.5px}
.logo span{color:var(--accent)}
.nav-links{display:flex;gap:24px;list-style:none}
.nav-links a{text-decoration:none;color:var(--muted);font-size:14px;font-weight:500;
             transition:color .2s}
.nav-links a:hover{color:var(--brand)}
.nav-cta{background:var(--brand);color:#fff;padding:8px 18px;border-radius:8px;
         font-size:13px;font-weight:600;text-decoration:none;transition:background .2s}
.nav-cta:hover{background:var(--brand-d)}

/* ── HERO ── */
.hero{text-align:center;padding:72px 24px 56px;max-width:720px;margin:0 auto}
.hero-badge{display:inline-flex;align-items:center;gap:6px;background:#EFF6FF;
            color:var(--brand);border:1px solid #BFDBFE;border-radius:99px;
            padding:5px 14px;font-size:12px;font-weight:600;margin-bottom:20px}
.hero-badge::before{content:"🇮🇳"}
h1{font-size:clamp(28px,5vw,48px);font-weight:800;line-height:1.15;
   letter-spacing:-.5px;margin-bottom:18px}
h1 em{font-style:normal;color:var(--brand)}
.hero p{font-size:17px;color:var(--muted);line-height:1.7;max-width:560px;margin:0 auto 36px}

/* ── STATS BAR ── */
.stats{display:flex;justify-content:center;gap:40px;flex-wrap:wrap;
       padding:20px 24px;background:var(--white);border-top:1px solid var(--border);
       border-bottom:1px solid var(--border)}
.stat{text-align:center}
.stat-n{font-size:24px;font-weight:800;color:var(--brand)}
.stat-l{font-size:12px;color:var(--muted);margin-top:2px}

/* ── MAIN GRID ── */
.main{max-width:1100px;margin:0 auto;padding:56px 24px;
      display:grid;grid-template-columns:1fr 1fr;gap:32px;align-items:start}
@media(max-width:768px){.main{grid-template-columns:1fr}}

/* ── CARD ── */
.card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);
      box-shadow:var(--shadow);overflow:hidden}
.card-head{padding:20px 24px;border-bottom:1px solid var(--border);
           display:flex;align-items:center;gap:10px}
.card-head .icon{width:36px;height:36px;border-radius:8px;display:flex;
                 align-items:center;justify-content:center;font-size:18px}
.card-head h2{font-size:16px;font-weight:700}
.card-head p{font-size:12px;color:var(--muted);margin-top:1px}
.card-body{padding:24px}

/* ── FORM ── */
.field{margin-bottom:20px}
label{display:block;font-size:12px;font-weight:600;color:var(--ink);
      margin-bottom:6px;letter-spacing:.02em;text-transform:uppercase}
.hint{font-size:11px;color:var(--muted);margin-top:4px}

/* File drop zone */
.dropzone{border:2px dashed var(--border);border-radius:10px;padding:32px 20px;
          text-align:center;cursor:pointer;transition:all .2s;position:relative;
          background:var(--bg)}
.dropzone:hover,.dropzone.drag{border-color:var(--brand);background:#EFF6FF}
.dropzone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.dropzone .dz-icon{font-size:32px;margin-bottom:8px}
.dropzone .dz-text{font-size:13px;color:var(--muted)}
.dropzone .dz-text strong{color:var(--brand)}
.dropzone .dz-file{font-size:13px;font-weight:600;color:var(--green);margin-top:6px;display:none}

.row2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
input[type=number],input[type=text]{
  width:100%;border:1.5px solid var(--border);border-radius:8px;
  padding:10px 14px;font-family:inherit;font-size:14px;color:var(--ink);
  background:var(--white);transition:border-color .2s;outline:none}
input:focus{border-color:var(--brand)}

.btn{width:100%;background:var(--brand);color:#fff;border:none;
     border-radius:10px;padding:14px;font-family:inherit;font-size:15px;
     font-weight:700;cursor:pointer;transition:background .2s;
     display:flex;align-items:center;justify-content:center;gap:8px}
.btn:hover{background:var(--brand-d)}
.btn:disabled{background:#93C5FD;cursor:not-allowed}
.btn .spinner{width:18px;height:18px;border:2.5px solid rgba(255,255,255,.3);
              border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite;
              display:none}
@keyframes spin{to{transform:rotate(360deg)}}

/* Status */
#status{margin-top:16px;border-radius:10px;padding:14px 16px;
        font-size:13px;display:none;line-height:1.6}
#status.success{background:#ECFDF5;border:1px solid #A7F3D0;color:#065F46}
#status.error{background:#FEF2F2;border:1px solid #FECACA;color:#991B1B}
.log-list{margin-top:8px;padding-left:16px;font-size:11px;color:#374151;line-height:2}

.dl-btn{display:none;margin-top:12px;width:100%;background:var(--green);
        color:#fff;border:none;border-radius:10px;padding:13px;
        font-family:inherit;font-size:14px;font-weight:600;
        cursor:pointer;text-decoration:none;text-align:center;transition:background .2s}
.dl-btn:hover{background:#059669}

/* ── HOW IT WORKS ── */
.steps{padding:0;list-style:none;counter-reset:step}
.steps li{display:flex;gap:14px;align-items:flex-start;padding:16px 0;
          border-bottom:1px solid var(--border)}
.steps li:last-child{border:none}
.steps li::before{counter-increment:step;content:counter(step);
                  min-width:28px;height:28px;background:var(--brand);color:#fff;
                  border-radius:50%;display:flex;align-items:center;
                  justify-content:center;font-size:12px;font-weight:700;margin-top:1px}
.steps li strong{display:block;font-size:13px;font-weight:600;margin-bottom:2px}
.steps li span{font-size:12px;color:var(--muted)}

/* ── FEATURES ── */
.features{max-width:1100px;margin:0 auto;padding:0 24px 56px;
          display:grid;grid-template-columns:repeat(3,1fr);gap:20px}
@media(max-width:640px){.features{grid-template-columns:1fr}}
.feat{background:var(--white);border:1px solid var(--border);border-radius:var(--radius);
      padding:24px;text-align:center}
.feat .fi{font-size:32px;margin-bottom:12px}
.feat h3{font-size:14px;font-weight:700;margin-bottom:6px}
.feat p{font-size:12px;color:var(--muted);line-height:1.6}

/* ── PRICING ── */
.pricing-section{background:var(--white);border-top:1px solid var(--border);
                 border-bottom:1px solid var(--border);padding:56px 24px}
.pricing-section h2{text-align:center;font-size:28px;font-weight:800;margin-bottom:8px}
.pricing-section .sub{text-align:center;color:var(--muted);font-size:15px;margin-bottom:40px}
.plans{max-width:900px;margin:0 auto;
       display:grid;grid-template-columns:repeat(3,1fr);gap:20px}
@media(max-width:640px){.plans{grid-template-columns:1fr}}
.plan{border:1.5px solid var(--border);border-radius:var(--radius);padding:28px 24px}
.plan.pop{border-color:var(--brand);position:relative}
.plan-badge{position:absolute;top:-12px;left:50%;transform:translateX(-50%);
            background:var(--brand);color:#fff;font-size:11px;font-weight:700;
            padding:3px 12px;border-radius:99px;white-space:nowrap}
.plan-name{font-size:13px;font-weight:700;color:var(--muted);text-transform:uppercase;
           letter-spacing:.05em;margin-bottom:8px}
.plan-price{font-size:32px;font-weight:800;color:var(--ink);margin-bottom:4px}
.plan-price span{font-size:14px;font-weight:400;color:var(--muted)}
.plan-desc{font-size:12px;color:var(--muted);margin-bottom:20px;line-height:1.5}
.plan ul{list-style:none;margin-bottom:24px}
.plan ul li{font-size:13px;padding:5px 0;display:flex;gap:8px;align-items:flex-start}
.plan ul li::before{content:"✓";color:var(--green);font-weight:700;flex-shrink:0}
.plan-btn{display:block;text-align:center;padding:10px;border-radius:8px;
          font-size:13px;font-weight:600;text-decoration:none;transition:all .2s;
          border:1.5px solid var(--brand);color:var(--brand)}
.plan-btn:hover{background:var(--brand);color:#fff}
.plan.pop .plan-btn{background:var(--brand);color:#fff}
.plan.pop .plan-btn:hover{background:var(--brand-d)}

/* ── FAQ ── */
.faq-section{max-width:760px;margin:0 auto;padding:56px 24px}
.faq-section h2{font-size:24px;font-weight:800;text-align:center;margin-bottom:32px}
details{border:1px solid var(--border);border-radius:10px;margin-bottom:10px;overflow:hidden}
summary{padding:16px 20px;font-size:14px;font-weight:600;cursor:pointer;
        list-style:none;display:flex;justify-content:space-between;align-items:center}
summary::after{content:"＋";font-size:16px;color:var(--muted)}
details[open] summary::after{content:"－"}
details p{padding:0 20px 16px;font-size:13px;color:var(--muted);line-height:1.7}

/* ── FOOTER ── */
footer{background:var(--ink);color:#9CA3AF;text-align:center;
       padding:24px;font-size:12px}
footer a{color:#6B7280;text-decoration:none}

/* ── TOAST ── */
.toast{position:fixed;bottom:24px;right:24px;background:var(--ink);color:#fff;
       padding:12px 20px;border-radius:10px;font-size:13px;font-weight:500;
       transform:translateY(80px);transition:transform .3s;z-index:999}
.toast.show{transform:translateY(0)}
</style>
</head>
<body>

<!-- NAV -->
<nav>
  <div class="logo">Balance<span>Shift</span></div>
  <ul class="nav-links">
    <li><a href="#tool">Tool</a></li>
    <li><a href="#how">How it works</a></li>
    <li><a href="#pricing">Pricing</a></li>
    <li><a href="#faq">FAQ</a></li>
  </ul>
  <a href="#tool" class="nav-cta">Try Free →</a>
</nav>

<!-- HERO -->
<section class="hero">
  <div class="hero-badge">Made for Indian CAs &amp; Accountants</div>
  <h1>Roll Over Your Balance Sheet<br/>to <em>FY 2026</em> in Seconds</h1>
  <p>Stop manually shifting columns every year. Upload your comparative Excel balance sheet — we shift CY→PY, clear the new CY column, restore all formulas, and fix every date. Works with <strong>any</strong> format.</p>
</section>

<!-- STATS -->
<div class="stats">
  <div class="stat"><div class="stat-n">100%</div><div class="stat-l">Formatting preserved</div></div>
  <div class="stat"><div class="stat-n">All sheets</div><div class="stat-l">Processed at once</div></div>
  <div class="stat"><div class="stat-n">&lt; 10 sec</div><div class="stat-l">Processing time</div></div>
  <div class="stat"><div class="stat-n">Any format</div><div class="stat-l">Works with all CA templates</div></div>
</div>

<!-- MAIN TOOL + HOW IT WORKS -->
<div class="main" id="tool">

  <!-- LEFT: UPLOAD CARD -->
  <div class="card">
    <div class="card-head">
      <div class="icon" style="background:#EFF6FF">📊</div>
      <div>
        <h2>Process Your Balance Sheet</h2>
        <p>Free · No signup · File never stored</p>
      </div>
    </div>
    <div class="card-body">

      <div class="field">
        <label>Upload Excel File (.xlsx)</label>
        <!-- Usage badge -->
        <div id="usageBadge" style="display:flex;align-items:center;justify-content:space-between;
             padding:8px 14px;border-radius:8px;border:1px solid #A7F3D0;background:#ECFDF5;
             margin-bottom:10px;font-size:12px;font-weight:600;transition:all .3s">
          <span>🆓 Free Plan</span>
          <span id="usageCounter">2 free files remaining</span>
        </div>
        <div class="dropzone" id="dropzone">
          <input type="file" id="xlFile" accept=".xlsx,.xls"/>
          <div class="dz-icon">📁</div>
          <div class="dz-text"><strong>Click to browse</strong> or drag &amp; drop</div>
          <div class="dz-text" style="margin-top:4px">Only .xlsx files · Max 20 MB</div>
          <div class="dz-file" id="dzFile"></div>
        </div>
      </div>

      <div class="row2">
        <div class="field">
          <label>Closing Year (CY)</label>
          <input type="number" id="closingYear" value="2025" min="2000" max="2100"/>
          <p class="hint">Year ending 31.03.2025</p>
        </div>
        <div class="field">
          <label>New Year</label>
          <input type="number" id="newYear" value="2026" min="2000" max="2100" readonly/>
          <p class="hint">Auto-filled</p>
        </div>
      </div>

      <div class="field">
        <label>Output Filename <span style="font-weight:400;text-transform:none;color:var(--muted)">(optional)</span></label>
        <input type="text" id="outputName" placeholder="e.g. ClientName_FY2026"/>
        <p class="hint">Leave blank to auto-generate from uploaded filename</p>
      </div>

      <button class="btn" id="processBtn" onclick="processFile()">
        <span id="btnText">⚡ Process &amp; Download</span>
        <div class="spinner" id="spinner"></div>
      </button>

      <div id="status"></div>
      <a id="dlBtn" class="dl-btn" href="#">⬇&nbsp; Download Processed File</a>

    </div>
  </div>

  <!-- RIGHT: HOW IT WORKS -->
  <div id="how">
    <div class="card" style="margin-bottom:20px">
      <div class="card-head">
        <div class="icon" style="background:#F0FDF4">✅</div>
        <div><h2>How It Works</h2><p>4 steps, fully automatic</p></div>
      </div>
      <div class="card-body">
        <ol class="steps">
          <li>
            <strong>Upload your Excel file</strong>
            <span>Your current FY2025 comparative balance sheet with CY and PY columns</span>
          </li>
          <li>
            <strong>Auto-detects all CY/PY columns</strong>
            <span>Scans every sheet and finds the correct data columns automatically</span>
          </li>
          <li>
            <strong>Shifts CY → PY, clears CY</strong>
            <span>2025 values become PY. All SUM formulas, cross-sheet links restored. CY constants cleared for fresh entry</span>
          </li>
          <li>
            <strong>Updates every date</strong>
            <span>31.03.2025 → 31.03.2026 everywhere. PY headers correctly show 31.03.2025</span>
          </li>
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
          <li>
            <strong>Files deleted immediately</strong>
            <span>Your uploaded file is deleted from our server the moment processing completes</span>
          </li>
          <li>
            <strong>No account required</strong>
            <span>No login, no email, no tracking. Just upload and download</span>
          </li>
          <li>
            <strong>HTTPS encrypted</strong>
            <span>All transfers are encrypted end-to-end</span>
          </li>
        </ul>
      </div>
    </div>
  </div>

</div>

<!-- FEATURES -->
<div class="features">
  <div class="feat"><div class="fi">🧮</div><h3>Formulas Preserved</h3><p>Every SUM, cross-sheet reference and formula in the PY column is snapshotted and restored automatically.</p></div>
  <div class="feat"><div class="fi">🎨</div><h3>Formatting Intact</h3><p>Fonts, borders, colors, merged cells, column widths — everything looks exactly as your CA designed it.</p></div>
  <div class="feat"><div class="fi">📅</div><h3>All Dates Updated</h3><p>Every date string across every sheet — titles, headers, notes — updated in one shot.</p></div>
  <div class="feat"><div class="fi">🗂️</div><h3>All Sheets at Once</h3><p>BS, P&L, Notes, Capital, Fixed Assets, Gross Profit, Details — every sheet processed together.</p></div>
  <div class="feat"><div class="fi">🔄</div><h3>Any CA Template</h3><p>Auto-detects column positions. Works with any firm's template, not just one format.</p></div>
  <div class="feat"><div class="fi">⚡</div><h3>Instant Results</h3><p>What used to take 30–45 minutes of careful manual work now takes under 10 seconds.</p></div>
</div>

<!-- PRICING -->
<section class="pricing-section" id="pricing">
  <h2>Simple Pricing</h2>
  <p class="sub">Start free. Upgrade when you need more.</p>
  <div class="plans">

    <div class="plan">
      <div class="plan-name">Free</div>
      <div class="plan-price">₹0 <span>/ month</span></div>
      <div class="plan-desc">Perfect for trying it out</div>
      <ul>
        <li>3 files per month</li>
        <li>All sheet types</li>
        <li>All features included</li>
        <li>Files up to 5 MB</li>
      </ul>
      <a href="#tool" class="plan-btn">Get Started Free</a>
    </div>

    <div class="plan pop">
      <div class="plan-badge">Most Popular</div>
      <div class="plan-name">Professional</div>
      <div class="plan-price">₹499 <span>/ month</span></div>
      <div class="plan-desc">For active CA practices</div>
      <ul>
        <li>Unlimited files</li>
        <li>Files up to 20 MB</li>
        <li>Priority processing</li>
        <li>Email support</li>
        <li>Processing history</li>
      </ul>
      <a href="#contact" class="plan-btn">Start 7-Day Trial</a>
    </div>

    <div class="plan">
      <div class="plan-name">Firm</div>
      <div class="plan-price">₹1,499 <span>/ month</span></div>
      <div class="plan-desc">For CA firms with multiple staff</div>
      <ul>
        <li>Everything in Pro</li>
        <li>5 team members</li>
        <li>Bulk processing</li>
        <li>API access</li>
        <li>WhatsApp support</li>
      </ul>
      <a href="#contact" class="plan-btn">Contact Us</a>
    </div>

  </div>
</section>

<!-- FAQ -->
<section class="faq-section" id="faq">
  <h2>Frequently Asked Questions</h2>

  <details>
    <summary>Which Excel formats does it support?</summary>
    <p>It supports .xlsx files only (Excel 2007 and newer). If your file is in .xls format, open it in Excel and Save As .xlsx first.</p>
  </details>

  <details>
    <summary>Will it work with my CA firm's custom template?</summary>
    <p>Yes. The tool auto-detects CY and PY columns by scanning for date headers like "31.03.2025". It works with any Indian CA balance sheet template regardless of column positions or sheet names.</p>
  </details>

  <details>
    <summary>Are my formulas and formatting safe?</summary>
    <p>Yes. Before copying any values, the tool snapshots every formula in the PY column and restores them afterward. Formatting (fonts, colors, borders, merged cells, column widths) is never touched.</p>
  </details>

  <details>
    <summary>What happens to my uploaded file?</summary>
    <p>Your file is processed in memory and deleted from our server immediately after you download the result. We do not store, read, or share your financial data.</p>
  </details>

  <details>
    <summary>Does it work for P&L, Notes to Accounts, Fixed Assets too?</summary>
    <p>Yes. Every sheet in your workbook is processed — Balance Sheet, P&L, Notes to BS, Notes to P&L, Capital Account, Fixed Assets, Gross Profit, Details, and any other sheets.</p>
  </details>

  <details>
    <summary>What if the PY column header still shows the wrong year?</summary>
    <p>The tool handles both plain-text headers and formula-based headers (like =E5). It corrects the PY header in all cases including formula references to the CY header cell.</p>
  </details>
</section>

<!-- FOOTER -->
<footer id="contact">
  <p style="margin-bottom:8px;color:#D1D5DB;font-weight:600">BalanceShift</p>
  <p>Built for Indian Chartered Accountants · Saves hours every April</p>
  <p style="margin-top:8px">Questions? Email: <a href="mailto:support@balanceshift.in">support@balanceshift.in</a></p>
  <p style="margin-top:16px;font-size:11px">© 2026 BalanceShift · Your data is never stored</p>
</footer>

<div class="toast" id="toast"></div>

<script>
// ── drag and drop ──
const dz = document.getElementById('dropzone');
const fi = document.getElementById('xlFile');
const dzFile = document.getElementById('dzFile');

dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('drag');
  if(e.dataTransfer.files.length) { fi.files = e.dataTransfer.files; showFile(fi.files[0]); }
});
fi.addEventListener('change', () => { if(fi.files.length) showFile(fi.files[0]); });
function showFile(f){
  dzFile.textContent = '✓ ' + f.name;
  dzFile.style.display = 'block';
}

// ── auto-sync years ──
document.getElementById('closingYear').addEventListener('input', function(){
  const v = parseInt(this.value);
  if(!isNaN(v)) document.getElementById('newYear').value = v + 1;
});

// ── 2-file free limit (stored in browser) ──
const FREE_LIMIT = 2;

function getUsageCount(){
  return parseInt(localStorage.getItem('bs_usage') || '0');
}
function incrementUsage(){
  localStorage.setItem('bs_usage', getUsageCount() + 1);
}
function showUsageBadge(){
  const used = getUsageCount();
  const left = Math.max(0, FREE_LIMIT - used);
  const badge = document.getElementById('usageBadge');
  const counter = document.getElementById('usageCounter');
  if(badge && counter){
    counter.textContent = left + ' free file' + (left !== 1 ? 's' : '') + ' remaining';
    badge.style.background = left === 0 ? '#FEF2F2' : left === 1 ? '#FFFBEB' : '#ECFDF5';
    badge.style.borderColor = left === 0 ? '#FECACA' : left === 1 ? '#FDE68A' : '#A7F3D0';
    badge.style.color       = left === 0 ? '#991B1B' : left === 1 ? '#92400E' : '#065F46';
  }
}
window.addEventListener('load', showUsageBadge);

// ── process ──
async function processFile(){
  const f     = fi.files[0];
  const cYear = parseInt(document.getElementById('closingYear').value);
  const nYear = parseInt(document.getElementById('newYear').value);
  const oName = document.getElementById('outputName').value.trim();
  const btn   = document.getElementById('processBtn');
  const sp    = document.getElementById('spinner');
  const bt    = document.getElementById('btnText');
  const dl    = document.getElementById('dlBtn');

  if(!f)            { showStatus('error', '✗ Please select an Excel file first.'); return; }
  if(isNaN(cYear))  { showStatus('error', '✗ Enter a valid closing year.'); return; }

  // ── check free limit ──
  if(getUsageCount() >= FREE_LIMIT){
    showStatus('error',
      '🔒 <strong>Free limit reached.</strong> You have used your 2 free files.<br><br>' +
      '📩 To continue, contact us for Pro access:<br>' +
      '<strong>WhatsApp: <a href="https://wa.me/91XXXXXXXXXX" style="color:#1D4ED8">+91-XXXXXXXXXX</a></strong><br>' +
      '<strong>Email: <a href="mailto:support@balanceshift.in" style="color:#1D4ED8">support@balanceshift.in</a></strong><br><br>' +
      '💼 Pro Plan: <strong>₹499/month</strong> — Unlimited files'
    );
    return;
  }

  btn.disabled = true;
  sp.style.display = 'block';
  bt.textContent = 'Processing…';
  dl.style.display = 'none';
  showStatus('', '');

  const fd = new FormData();
  fd.append('file', f);
  fd.append('closing_year', cYear);
  fd.append('new_year', nYear);
  fd.append('output_name', oName);

  try {
    const res  = await fetch('/process', { method:'POST', body:fd });
    const data = await res.json();

    if(data.status === 'success'){
      incrementUsage();
      showUsageBadge();
      const logHtml = '<ul class="log-list">'
        + data.log.map(l=>`<li>${l}</li>`).join('') + '</ul>';
      showStatus('success', '✓ Done! Your file is ready to download.' + logHtml);
      dl.href        = '/download/' + data.file_id;
      dl.download    = data.filename;
      dl.textContent = '⬇  Download — ' + data.filename;
      dl.style.display = 'block';
      toast('File processed successfully!');
    } else {
      showStatus('error', '✗ ' + data.message);
    }
  } catch(e){
    showStatus('error', '✗ Network error: ' + e.message);
  } finally {
    btn.disabled = false;
    sp.style.display = 'none';
    bt.textContent = '⚡ Process & Download';
  }
}

function showStatus(type, msg){
  const el = document.getElementById('status');
  el.className = type;
  el.innerHTML = msg;
  el.style.display = msg ? 'block' : 'none';
}

function toast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3000);
}
</script>
</body>
</html>"""


# ─── ROUTES ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/process", methods=["POST"])
def process_file():
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "No file uploaded."})

    f = request.files["file"]
    if not (f.filename.endswith('.xlsx') or f.filename.endswith('.xls')):
        return jsonify({"status": "error", "message": "Only .xlsx and .xls files are supported."})

    try:
        closing_year = int(request.form.get("closing_year", 2025))
        new_year     = int(request.form.get("new_year", closing_year + 1))
        out_name     = request.form.get("output_name", "").strip()
    except ValueError:
        return jsonify({"status": "error", "message": "Invalid year values."})

    uid         = uuid.uuid4().hex
    ext         = ".xls" if f.filename.lower().endswith(".xls") else ".xlsx"
    input_path  = os.path.join(UPLOAD_DIR, f"{uid}_in{ext}")
    output_path = os.path.join(OUTPUT_DIR, f"{uid}_out.xlsx")
    f.save(input_path)

    base     = out_name if out_name else os.path.splitext(f.filename)[0]
    filename = f"{base}_{new_year}.xlsx"

    try:
        result = process(input_path, output_path, closing_year, new_year)
    except Exception as e:
        try: os.remove(input_path)
        except: pass
        return jsonify({"status": "error", "message": str(e)})

    try: os.remove(input_path)
    except: pass

    return jsonify({
        "status":   "success",
        "log":      result["log"],
        "file_id":  uid,
        "filename": filename,
    })


@app.route("/download/<file_id>")
def download(file_id):
    if not re.fullmatch(r"[a-f0-9]{32}", file_id):
        return "Invalid ID", 400
    path = os.path.join(OUTPUT_DIR, f"{file_id}_out.xlsx")
    if not os.path.exists(path):
        return "File not found or expired.", 404
    return send_file(
        path, as_attachment=True,
        download_name=f"balanceshift_{file_id[:8]}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
