"""
Balance Sheet Year-Shift Web App
Run:  python app.py
Then open http://localhost:5000
"""

import re
import os
import uuid
from flask import Flask, request, send_file, jsonify, render_template_string
from processor import process

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024   # 20 MB upload limit

UPLOAD_FOLDER = "/tmp/bs_uploads"
OUTPUT_FOLDER = "/tmp/bs_outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Balance Sheet Year-Shift Tool</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=DM+Sans:wght@300;400;500&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --ink:     #1a1a2e;
    --paper:   #f5f0e8;
    --cream:   #ede8dc;
    --gold:    #c9a84c;
    --rust:    #b5451b;
    --sage:    #4a7c6f;
    --shadow:  rgba(26,26,46,.12);
  }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--paper);
    color: var(--ink);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 48px 20px 80px;
  }

  /* ── ledger-line texture ── */
  body::before {
    content: '';
    position: fixed; inset: 0;
    background: repeating-linear-gradient(
      to bottom, transparent 0px,
      transparent 31px, rgba(201,168,76,.15) 32px
    );
    pointer-events: none;
    z-index: 0;
  }

  .card {
    position: relative; z-index: 1;
    background: #fff;
    border: 1.5px solid var(--cream);
    border-radius: 4px;
    box-shadow: 0 2px 0 var(--gold), 0 8px 32px var(--shadow);
    width: 100%; max-width: 640px;
    padding: 48px 52px;
  }

  .stamp {
    position: absolute;
    top: -18px; right: 36px;
    background: var(--rust);
    color: #fff;
    font-family: 'DM Sans', sans-serif;
    font-size: 10px;
    font-weight: 500;
    letter-spacing: .12em;
    text-transform: uppercase;
    padding: 4px 14px;
    border-radius: 2px;
  }

  h1 {
    font-family: 'Playfair Display', serif;
    font-size: 28px;
    font-weight: 900;
    line-height: 1.2;
    color: var(--ink);
    margin-bottom: 6px;
  }

  .sub {
    font-size: 13px;
    color: #888;
    margin-bottom: 36px;
    letter-spacing: .02em;
  }

  label {
    display: block;
    font-size: 11px;
    font-weight: 500;
    letter-spacing: .1em;
    text-transform: uppercase;
    color: var(--sage);
    margin-bottom: 8px;
  }

  .field { margin-bottom: 24px; }

  input[type="file"],
  input[type="number"],
  input[type="text"] {
    width: 100%;
    border: 1.5px solid var(--cream);
    border-radius: 3px;
    padding: 11px 14px;
    font-family: 'DM Sans', sans-serif;
    font-size: 14px;
    color: var(--ink);
    background: var(--paper);
    transition: border-color .2s;
    outline: none;
  }

  input:focus { border-color: var(--gold); background: #fff; }

  .row-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }

  .hint {
    font-size: 11px;
    color: #aaa;
    margin-top: 5px;
  }

  .divider {
    border: none;
    border-top: 1.5px dashed var(--cream);
    margin: 28px 0;
  }

  button {
    width: 100%;
    background: var(--ink);
    color: #fff;
    border: none;
    border-radius: 3px;
    padding: 14px;
    font-family: 'Playfair Display', serif;
    font-size: 16px;
    font-weight: 700;
    letter-spacing: .04em;
    cursor: pointer;
    transition: background .2s, transform .1s;
    position: relative;
    overflow: hidden;
  }
  button:hover { background: var(--rust); }
  button:active { transform: scale(.99); }
  button:disabled { background: #ccc; cursor: not-allowed; }

  /* spinner inside button */
  .spinner {
    display: none;
    width: 18px; height: 18px;
    border: 2.5px solid rgba(255,255,255,.3);
    border-top-color: #fff;
    border-radius: 50%;
    animation: spin .7s linear infinite;
    position: absolute; right: 18px; top: 50%; transform: translateY(-50%);
  }
  @keyframes spin { to { transform: translateY(-50%) rotate(360deg); } }

  #status {
    margin-top: 20px;
    border-radius: 3px;
    padding: 14px 16px;
    font-size: 13px;
    display: none;
  }
  #status.success {
    background: #eef7f2;
    border: 1.5px solid #a8d5bc;
    color: var(--sage);
  }
  #status.error {
    background: #fdf0ed;
    border: 1.5px solid #f0b8a8;
    color: var(--rust);
  }

  .log-list {
    margin-top: 10px;
    padding-left: 16px;
    font-size: 12px;
    line-height: 1.8;
    color: #555;
  }

  .download-btn {
    display: none;
    margin-top: 14px;
    width: 100%;
    background: var(--sage);
    color: #fff;
    border: none;
    border-radius: 3px;
    padding: 13px;
    font-family: 'DM Sans', sans-serif;
    font-size: 14px;
    font-weight: 500;
    cursor: pointer;
    text-decoration: none;
    text-align: center;
    transition: background .2s;
  }
  .download-btn:hover { background: #3a6559; }

  .watermark {
    position: absolute;
    bottom: 18px; right: 24px;
    font-size: 10px;
    color: var(--gold);
    opacity: .5;
    font-family: 'Playfair Display', serif;
    letter-spacing: .05em;
  }

  @media (max-width: 480px) {
    .card { padding: 36px 24px; }
    .row-2 { grid-template-columns: 1fr; }
  }
</style>
</head>
<body>

<div class="card">
  <span class="stamp">Accounts Utility</span>

  <h1>Balance Sheet<br>Year-Shift Tool</h1>
  <p class="sub">Rolls over CY → PY and prepares a blank CY column for the new financial year.</p>

  <div class="field">
    <label>Excel File (.xlsx)</label>
    <input type="file" id="xlFile" accept=".xlsx" />
    <p class="hint">Upload the comparative balance sheet workbook</p>
  </div>

  <div class="row-2">
    <div class="field">
      <label>Closing Year (CY ending)</label>
      <input type="number" id="closingYear" value="2025" min="2000" max="2100" />
      <p class="hint">e.g. 2025 → data as at 31.03.2025</p>
    </div>
    <div class="field">
      <label>New Year (new CY)</label>
      <input type="number" id="newYear" value="2026" min="2000" max="2100" />
      <p class="hint">e.g. 2026 → will show 31.03.2026</p>
    </div>
  </div>

  <div class="field">
    <label>Output Filename</label>
    <input type="text" id="outputName" placeholder="Leave blank to auto-generate" />
    <p class="hint">Do not include .xlsx extension</p>
  </div>

  <hr class="divider">

  <button id="processBtn" onclick="processFile()">
    Process &amp; Download
    <div class="spinner" id="spinner"></div>
  </button>

  <div id="status"></div>
  <a id="downloadLink" class="download-btn">⬇ Download Processed File</a>

  <span class="watermark">BS Shifter v1.0</span>
</div>

<script>
async function processFile() {
  const fileInput  = document.getElementById('xlFile');
  const closingYr  = parseInt(document.getElementById('closingYear').value);
  const newYr      = parseInt(document.getElementById('newYear').value);
  const outName    = document.getElementById('outputName').value.trim();
  const btn        = document.getElementById('processBtn');
  const spinner    = document.getElementById('spinner');
  const statusDiv  = document.getElementById('status');
  const dlLink     = document.getElementById('downloadLink');

  // Validation
  if (!fileInput.files.length) {
    showStatus('error', 'Please select an Excel file first.');
    return;
  }
  if (isNaN(closingYr) || isNaN(newYr) || newYr !== closingYr + 1) {
    showStatus('error', 'New Year must be exactly one year after Closing Year (e.g. 2025 → 2026).');
    return;
  }

  btn.disabled = true;
  spinner.style.display = 'block';
  dlLink.style.display  = 'none';
  showStatus('', '');

  const fd = new FormData();
  fd.append('file', fileInput.files[0]);
  fd.append('closing_year', closingYr);
  fd.append('new_year', newYr);
  fd.append('output_name', outName);

  try {
    const res  = await fetch('/process', { method: 'POST', body: fd });
    const data = await res.json();

    if (data.status === 'success') {
      let logHtml = '<ul class="log-list">' +
        data.log.map(l => `<li>${l}</li>`).join('') + '</ul>';
      showStatus('success', '✓ File processed successfully.' + logHtml);

      dlLink.href        = '/download/' + data.file_id;
      dlLink.download    = data.filename;
      dlLink.textContent = '⬇  Download — ' + data.filename;
      dlLink.style.display = 'block';
    } else {
      showStatus('error', '✗ Error: ' + data.message);
    }
  } catch (e) {
    showStatus('error', '✗ Network error: ' + e.message);
  } finally {
    btn.disabled        = false;
    spinner.style.display = 'none';
  }
}

function showStatus(type, msg) {
  const el = document.getElementById('status');
  el.className = type;
  el.innerHTML = msg;
  el.style.display = msg ? 'block' : 'none';
}

// Auto-sync closing year → new year
document.getElementById('closingYear').addEventListener('input', function() {
  const v = parseInt(this.value);
  if (!isNaN(v)) document.getElementById('newYear').value = v + 1;
});
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/process", methods=["POST"])
def process_file():
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "No file uploaded."})

    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"status": "error", "message": "Only .xlsx files are supported."})

    try:
        closing_year = int(request.form.get("closing_year", 2025))
        new_year     = int(request.form.get("new_year", 2026))
        out_name     = request.form.get("output_name", "").strip()
    except ValueError:
        return jsonify({"status": "error", "message": "Invalid year values."})

    if new_year != closing_year + 1:
        return jsonify({"status": "error",
                        "message": "New year must be closing year + 1."})

    # Save upload
    uid        = uuid.uuid4().hex
    input_path = os.path.join(UPLOAD_FOLDER, f"{uid}_input.xlsx")
    f.save(input_path)

    # Build output filename
    base = out_name if out_name else os.path.splitext(f.filename)[0]
    filename    = f"{base}_{new_year}.xlsx"
    output_path = os.path.join(OUTPUT_FOLDER, f"{uid}_output.xlsx")

    try:
        result = process(input_path, output_path, closing_year, new_year)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})
    finally:
        try:
            os.remove(input_path)
        except Exception:
            pass

    return jsonify({
        "status":   "success",
        "log":      result["log"],
        "file_id":  uid,
        "filename": filename,
    })


@app.route("/download/<file_id>")
def download(file_id):
    # Safety: only hex chars allowed
    if not re.fullmatch(r"[a-f0-9]{32}", file_id):
        return "Invalid ID", 400
    path = os.path.join(OUTPUT_FOLDER, f"{file_id}_output.xlsx")
    if not os.path.exists(path):
        return "File not found or expired.", 404
    return send_file(
        path,
        as_attachment=True,
        download_name=f"balance_sheet_{file_id[:8]}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
