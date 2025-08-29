import os, tempfile, shutil, subprocess, uuid
from flask import Flask, request, send_file, render_template_string, after_this_request
from werkzeug.utils import secure_filename

ALLOWED = {".doc", ".docx", ".rtf", ".odt"}  # जो चाहें रखें
MAX_MB_FREE = 10
TITLE = "Mahamaya Stationery — Word → PDF Converter"

app = Flask(__name__)

INDEX_HTML = r"""
<!doctype html>
<html lang="hi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{title}}</title>
<style>
  body{margin:0;background:#0b1220;color:#e7eaf1;font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial}
  .wrap{min-height:100svh;display:grid;place-items:center;padding:22px}
  .card{width:min(760px,100%);background:linear-gradient(180deg,#0f172a,#0b1220);
        border:1px solid #203054;border-radius:18px;box-shadow:0 10px 40px rgba(0,0,0,.35);padding:22px}
  .top{display:flex;justify-content:space-between;gap:12px;align-items:center}
  .brand{display:flex;gap:10px;align-items:center;font-weight:800}
  .badge{font-size:12px;padding:2px 8px;border:1px solid #203054;border-radius:999px;color:#93a2bd}
  h1{font-size:22px;margin:10px 0}
  p.muted{color:#93a2bd;margin:4px 0 14px}
  .box{border:2px dashed #203054;border-radius:14px;padding:16px;text-align:center;background:#0d162a}
  input[type=file]{width:100%;margin-top:10px}
  .row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:14px}
  button{padding:10px 14px;border-radius:12px;border:1px solid #203054;background:#4f8cff;color:#fff;font-weight:700;cursor:pointer}
  button.ghost{background:#17233f}
  .note{font-size:12px;color:#93a2bd}
  .alert{margin-top:10px;padding:10px 12px;border-radius:12px;font-weight:600;display:none}
  .ok{background:rgba(34,197,94,.1);color:#22c55e;border:1px solid rgba(34,197,94,.25)}
  .err{background:rgba(239,68,68,.1);color:#ef4444;border:1px solid rgba(239,68,68,.25)}
  .spinner{width:18px;height:18px;border:3px solid rgba(255,255,255,.25);border-top-color:white;border-radius:50%;animation:spin 1s linear infinite;display:none}
  @keyframes spin{to{transform:rotate(360deg)}}
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <div class="top">
      <div class="brand">
        <div style="width:30px;height:30px;border-radius:8px;background:linear-gradient(135deg,#4f8cff,#22c55e)"></div>
        <div>Mahamaya Stationery</div>
      </div>
      <div class="badge">WORD → PDF (High Fidelity)</div>
    </div>

    <h1>Word से PDF — वही फ़ॉर्मैट, वही लेआउट ✅</h1>
    <p class="muted">यह कन्वर्टर LibreOffice का इस्तेमाल करता है ताकि MS Word जैसा लेआउट बना रहे (फ़ॉन्ट्स भी शामिल)।</p>

    <form id="f" method="post" action="/convert" enctype="multipart/form-data">
      <div class="box">
        <strong>अपने DOC/DOCX/RTF/ODT फ़ाइल चुनें</strong><br/>
        <input name="doc_file" type="file" accept=".doc,.docx,.rtf,.odt" required>
        <div class="note" style="margin-top:8px">टिप: अगर फ़ाइल पासवर्ड प्रोटेक्टेड है, कन्वर्ज़न संभव नहीं है।</div>
      </div>
      <div class="row">
        <button id="btn" type="submit">
          <span id="spin" class="spinner"></span>
          Convert & Download PDF
        </button>
        <span id="status" class="note"></span>
      </div>
      <div id="ok" class="alert ok">हो गया! डाउनलोड शुरू हो गया।</div>
      <div id="err" class="alert err">Error</div>
    </form>

    <p class="note" style="margin-top:12px">Free trial: 1 फ़ाइल प्रति कन्वर्ज़न, 10 MB तक। अधिक की ज़रूरत हो तो हम Razorpay जोड़ देंगे।</p>
  </div>
</div>

<script>
const f = document.getElementById('f');
const btn = document.getElementById('btn');
const spin = document.getElementById('spin');
const ok = document.getElementById('ok');
const err = document.getElementById('err');
const statusEl = document.getElementById('status');

function show(el,msg){ el.textContent=msg; el.style.display='block'; }
function hide(el){ el.style.display='none'; }

f.addEventListener('submit', ()=>{
  hide(ok); hide(err);
  btn.disabled = true; spin.style.display = 'inline-block';
  show(statusEl,'Converting… कृपया प्रतीक्षा करें');
});
</script>
</body>
</html>
"""

def _soffice_convert(in_path: str, out_dir: str):
    """
    LibreOffice (soffice) से हाई-फिडेलिटी PDF एक्सपोर्ट.
    """
    # --nologo/--headless/--convert-to pdf --outdir <dir> <file>
    cmd = [
        "soffice", "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
        "--convert-to", "pdf", "--outdir", out_dir, in_path
    ]
    # 120 sec timeout (बड़े docs के लिए बढ़ा सकते हैं)
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)

@app.route("/")
def home():
    return render_template_string(INDEX_HTML, title=TITLE)

@app.route("/healthz")
def health():
    return "OK"

@app.route("/convert", methods=["POST"])
def convert():
    f = request.files.get("doc_file")
    if not f:
        return ("No file uploaded", 400)

    name = secure_filename(f.filename or "")
    ext = os.path.splitext(name)[1].lower()
    if ext not in ALLOWED:
        return ("Unsupported file type", 400)

    # 10 MB free cap (फिलहाल hard cap; आगे Razorpay जोड़ सकते हैं)
    f.stream.seek(0, os.SEEK_END)
    size = f.stream.tell()
    f.stream.seek(0)
    if size > 10 * 1024 * 1024:
        return ("File too large for free (max 10 MB).", 400)

    tmp = tempfile.mkdtemp(prefix="word2pdf_")
    try:
        in_path = os.path.join(tmp, name or f"input{ext}")
        f.save(in_path)

        # पासवर्ड-प्रोटेक्टेड DOC/DOCX सपोर्टेड नहीं (LibreOffice non-interactive prompt)
        # quick detection: अगर फ़ाइल एन्क्रिप्टेड है तो LO विफल हो जाएगा
        try:
            _soffice_convert(in_path, tmp)
        except subprocess.CalledProcessError as e:
            return (f"Conversion failed (maybe password-protected or missing fonts).", 500)
        except subprocess.TimeoutExpired:
            return ("Conversion timed out. Try a smaller file.", 500)

        # आउट PDF ढूँढें
        base = os.path.splitext(os.path.basename(in_path))[0]
        out_path = os.path.join(tmp, base + ".pdf")
        # कुछ मामलों में LO filename sanitize करता है; फॉलबैक: पहला .pdf ढूँढें
        if not os.path.exists(out_path):
            for fn in os.listdir(tmp):
                if fn.lower().endswith(".pdf"):
                    out_path = os.path.join(tmp, fn); break
        if not os.path.exists(out_path):
            return ("Converted PDF not found.", 500)

        @after_this_request
        def _cleanup(resp):
            try: shutil.rmtree(tmp, ignore_errors=True)
            except: pass
            return resp

        return send_file(out_path, as_attachment=True, download_name=f"{base}.pdf")
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        return (f"Error: {e}", 500)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
