import os, tempfile, shutil, zipfile, subprocess, secrets, time
from threading import Timer
from datetime import datetime, timedelta

from flask import Flask, request, send_file, render_template_string, jsonify, after_this_request
from werkzeug.utils import secure_filename

# Razorpay (server-side)
import razorpay

# PDF page counting (no heavy deps)
from pypdf import PdfReader

app = Flask(__name__)

# -------- Business Rules --------
FREE_MAX_PAGES = 25          # प्रति फ़ाइल
FREE_MAX_MB = 10             # प्रति फ़ाइल
FREE_MAX_FILES = 2           # एक बार में
PAID_AMOUNT_RUPEES = 10      # ₹10
PAID_AMOUNT_PAISE = PAID_AMOUNT_RUPEES * 100

# Limits & security
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # कुल रिक्वेस्ट 50MB
ALLOWED_EXTS = {".doc", ".docx", ".odt", ".rtf"}

# Razorpay keys from ENV
RZP_KEY_ID = os.getenv("RAZORPAY_KEY_ID", "")
RZP_KEY_SECRET = os.getenv("RAZORPAY_KEY_SECRET", "")
RZP_ENABLED = bool(RZP_KEY_ID and RZP_KEY_SECRET)

# Razorpay client
client = razorpay.Client(auth=(RZP_KEY_ID, RZP_KEY_SECRET)) if RZP_ENABLED else None

# In-memory store: order_id/token -> session info
SESSION_STORE = {}   # token -> {"dir": str, "files": [pdfpaths], "created": ts}
ORDER_MAP = {}       # order_id -> token

# ---------- UI ----------
INDEX_HTML = r"""
<!doctype html>
<html lang="hi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mahamaya Stationery — Word → PDF Converter</title>
<link rel="preconnect" href="https://checkout.razorpay.com">
<style>
  :root{ --bg:#0b1220; --fg:#e7eaf1; --muted:#93a2bd; --card:#10182b; --stroke:#203054; --accent:#22c55e; --brand:#4f8cff; --warn:#f59e0b; --danger:#ef4444; }
  *{box-sizing:border-box} body{margin:0; background:var(--bg); color:var(--fg); font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial}
  .wrap{min-height:100svh; display:grid; place-items:center; padding:20px}
  .card{width:min(920px,100%); background:linear-gradient(180deg,#0f172a 0,#0b1220 100%); border:1px solid var(--stroke); border-radius:20px; padding:24px; box-shadow:0 10px 40px rgba(0,0,0,.35)}
  .top{display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap}
  .brand{display:flex; gap:10px; align-items:center; font-weight:800}
  .brand-badge{width:30px;height:30px;border-radius:8px;background:linear-gradient(135deg,var(--brand), var(--accent))}
  h1{margin:12px 0 6px; font-size:24px}
  p.muted{margin:0 0 14px; color:var(--muted)}
  .drop{border:2px dashed var(--stroke); background:#0d162a; border-radius:16px; padding:18px; text-align:center}
  .drop.drag{border-color:var(--brand); background:#10203f}
  .note{color:var(--muted); font-size:12px}
  input[type="file"]{display:none}
  .row{display:flex; gap:10px; align-items:center; flex-wrap:wrap}
  button.btn{padding:10px 14px; border-radius:12px; border:1px solid var(--stroke); background:var(--brand); color:#fff; font-weight:700; cursor:pointer}
  button.ghost{background:#17233f}
  button:disabled{opacity:.6; cursor:not-allowed}
  .alert{margin-top:10px; padding:10px 12px; border-radius:12px; display:none; font-weight:600}
  .ok{background:rgba(34,197,94,.1); color:#22c55e; border:1px solid rgba(34,197,94,.25)}
  .err{background:rgba(239,68,68,.1); color:#ef4444; border:1px solid rgba(239,68,68,.25)}
  .warn{background:rgba(245,158,11,.1); color:#f59e0b; border:1px solid rgba(245,158,11,.25)}
  .filelist{margin-top:8px; font-size:13px; line-height:1.4}
  .grid{display:grid; grid-template-columns: 1fr 1fr; gap:12px}
  @media (max-width:720px){ .grid{grid-template-columns:1fr} }
  /* Loading overlay */
  .overlay{position:fixed; inset:0; background:rgba(4,8,18,.6); display:none; place-items:center; z-index:50}
  .panel{background:#0f172a; border:1px solid var(--stroke); border-radius:16px; padding:18px; width:min(420px,90%)}
  .loader{width:48px;height:48px; border-radius:50%; border:6px solid rgba(255,255,255,.1); border-top-color:#fff; animation:spin 1s linear infinite; margin:12px auto}
  @keyframes spin {to { transform: rotate(360deg);}}
  .tip{font-size:12px; color:var(--muted)}
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <div class="top">
      <div class="brand"><div class="brand-badge"></div><div>Mahamaya Stationery</div></div>
      <div class="note">Word (.doc/.docx) → PDF • फॉर्मैटिंग Safe</div>
    </div>

    <h1>Word → PDF (Free ≤ 25 pages or 10 MB; ≤ 2 files)</h1>
    <p class="muted">सीमाओं से ऊपर होने पर Razorpay से ₹10 लगेगा।</p>

    <div id="drop" class="drop">
      <b>Drag & Drop files here</b> <span class="note">या क्लिक करें</span>
      <input id="files" type="file" accept=".doc,.docx,.odt,.rtf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" multiple>
      <div id="chosen" class="filelist note"></div>
    </div>

    <div style="height:12px"></div>

    <div class="row">
      <button id="convertBtn" class="btn">Convert</button>
      <button id="chooseBtn" class="btn ghost">Choose Files</button>
      <div id="status" class="note"></div>
    </div>

    <div id="ok" class="alert ok"></div>
    <div id="warn" class="alert warn"></div>
    <div id="err" class="alert err"></div>

    <div class="note" style="margin-top:10px">
      टिप्स: पासवर्ड-प्रोटेक्टेड Word फाइल सपोर्टेड नहीं। बड़ी फाइलों पर 10–20 सेकंड लग सकते हैं।
    </div>
  </div>
</div>

<!-- Overlay -->
<div id="overlay" class="overlay">
  <div class="panel">
    <div class="loader"></div>
    <div style="text-align:center; font-weight:700; margin-bottom:6px" id="overlayTitle">Processing…</div>
    <div class="tip" id="overlayTip">हम आपकी फाइलों का विश्लेषण कर रहे हैं। पेज गिनती और साइज चेक…</div>
  </div>
</div>

<script src="https://checkout.razorpay.com/v1/checkout.js"></script>
<script>
const drop = document.getElementById('drop');
const filesEl = document.getElementById('files');
const chooseBtn = document.getElementById('chooseBtn');
const convertBtn = document.getElementById('convertBtn');
const chosen = document.getElementById('chosen');
const ok = document.getElementById('ok');
const warn = document.getElementById('warn');
const err = document.getElementById('err');
const statusEl = document.getElementById('status');
const overlay = document.getElementById('overlay');
const overlayTitle = document.getElementById('overlayTitle');
const overlayTip = document.getElementById('overlayTip');

let selected = [];

function show(div,msg){ div.textContent = msg; div.style.display = 'block'; }
function hideAll(){ [ok,warn,err].forEach(d=>d.style.display='none'); statusEl.textContent=''; }
function setOverlay(on, title, tip){
  overlay.style.display = on ? 'grid' : 'none';
  if(title) overlayTitle.textContent = title;
  if(tip) overlayTip.textContent = tip;
}

drop.addEventListener('click', ()=> filesEl.click());
chooseBtn.addEventListener('click', ()=> filesEl.click());
['dragenter','dragover'].forEach(ev=>{
  drop.addEventListener(ev, e=>{ e.preventDefault(); drop.classList.add('drag'); });
});
['dragleave','drop'].forEach(ev=>{
  drop.addEventListener(ev, e=>{ e.preventDefault(); drop.classList.remove('drag'); });
});
drop.addEventListener('drop', e=>{
  e.preventDefault();
  const flist = Array.from(e.dataTransfer.files || []);
  applyFiles(flist);
});
filesEl.addEventListener('change', ()=>{
  applyFiles(Array.from(filesEl.files || []));
});

function applyFiles(list){
  hideAll();
  selected = list.filter(f=>/\.(docx?|odt|rtf)$/i.test(f.name));
  if(!selected.length){ show(err,"कृपया Word (.doc/.docx) फाइलें चुनें।"); chosen.textContent=""; return; }
  chosen.innerHTML = selected.map(f=>`• ${f.name} · ${(f.size/1024/1024).toFixed(2)} MB`).join('<br>');
}

convertBtn.addEventListener('click', async ()=>{
  hideAll();
  if(!selected.length){ show(err,"पहले Word फाइलें चुनें।"); return; }
  try{
    setOverlay(true,"Analyzing…","पेज और साइज चेक, फिर कन्वर्ज़न/पेमेंट फ़्लो।");
    convertBtn.disabled = true;

    const fd = new FormData();
    selected.forEach(f=>fd.append('files', f));

    const res = await fetch('/precheck', { method:'POST', body:fd });
    if(!res.ok){ throw new Error(await res.text()); }
    const data = await res.json();

    if(data.error){ throw new Error(data.error); }

    if(!data.payment_required){
      setOverlay(true,"Converting…","PDF तैयार हो रहे हैं।");
      window.location.href = `/download/${data.token}`;
      return;
    }

    // Payment needed
    const options = {
      key: data.key_id,
      amount: data.amount_paise,
      currency: "INR",
      name: "Mahamaya Stationery",
      description: data.message || "Word → PDF conversion",
      order_id: data.order_id,
      handler: async function (rsp) {
        // verify on server
        setOverlay(true,"Verifying Payment…","कृपया इंतज़ार करें, पेमेंट वेरीफाई हो रहा है।");
        const vr = await fetch('/verify', {
          method:'POST',
          headers:{'Content-Type':'application/json'},
          body: JSON.stringify({
            token: data.token,
            razorpay_order_id: rsp.razorpay_order_id,
            razorpay_payment_id: rsp.razorpay_payment_id,
            razorpay_signature: rsp.razorpay_signature
          })
        });
        if(!vr.ok){ show(err, await vr.text()); setOverlay(false); return; }
        const vj = await vr.json();
        if(vj.ok && vj.download_url){
          window.location.href = vj.download_url;
        }else{
          show(err, vj.error || "Verification failed.");
          setOverlay(false);
        }
      },
      theme:{ color:"#4f8cff" },
      modal: {
        ondismiss: function(){ setOverlay(false); show(warn,"पेमेंट रद्द कर दिया गया।"); }
      },
      prefill: {}
    };

    const rz = new Razorpay(options);
    setOverlay(false);
    rz.open();
  }catch(e){
    show(err, e.message || "Server error");
    setOverlay(false);
  }finally{
    convertBtn.disabled = false;
  }
});
</script>
</body>
</html>
"""

# ---------- Helpers ----------
def is_allowed(filename: str) -> bool:
    ext = os.path.splitext(filename or "")[1].lower()
    return ext in ALLOWED_EXTS

def mb(nbytes: int) -> float:
    return nbytes / 1024.0 / 1024.0

def run_soffice_to_pdf(src_path: str, out_dir: str):
    """
    Use LibreOffice headless to convert Word -> PDF
    """
    cmd = [
        "soffice", "--headless",
        "--convert-to", "pdf",
        "--outdir", out_dir,
        src_path
    ]
    r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if r.returncode != 0:
        raise RuntimeError("LibreOffice conversion failed (maybe password-protected).")
    base = os.path.splitext(os.path.basename(src_path))[0]
    pdf_path = os.path.join(out_dir, base + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError("PDF output missing (conversion error).")
    return pdf_path

def pdf_page_count(pdf_path: str) -> int:
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        return len(reader.pages)

def make_zip(pdf_paths, zip_path):
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in pdf_paths:
            arc = os.path.basename(p)
            zf.write(p, arc)

def cleanup_later(tmpdir: str, seconds: float = 120.0):
    Timer(seconds, shutil.rmtree, args=[tmpdir], kwargs={"ignore_errors": True}).start()

# ---------- Routes ----------
@app.route("/")
def index():
    return render_template_string(INDEX_HTML)

@app.route("/healthz")
def healthz():
    return "OK", 200

@app.route("/precheck", methods=["POST"])
def precheck():
    # Save uploads
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="No files uploaded."), 400
    if len(files) > 10:
        return jsonify(error="Too many files (max 10 at once)."), 400

    tmpdir = tempfile.mkdtemp(prefix="w2p_")
    srcdir = os.path.join(tmpdir, "src")
    pdfdir = os.path.join(tmpdir, "pdf")
    os.makedirs(srcdir, exist_ok=True)
    os.makedirs(pdfdir, exist_ok=True)

    saved = []
    reasons = []
    try:
        for f in files:
            fname = secure_filename(f.filename or "")
            if not fname or not is_allowed(fname):
                raise ValueError("Only .doc/.docx/.odt/.rtf files are allowed.")
            path = os.path.join(srcdir, fname)
            f.save(path)
            size_b = os.path.getsize(path)
            saved.append((path, size_b))

        # Convert all to PDF & evaluate
        pdf_paths = []
        need_payment = False

        for path, size_b in saved:
            # size rule
            if mb(size_b) > FREE_MAX_MB:
                reasons.append(f"{os.path.basename(path)} > {FREE_MAX_MB}MB")
                need_payment = True

            # convert & page count
            try:
                pdf_path = run_soffice_to_pdf(path, pdfdir)
            except Exception as e:
                shutil.rmtree(tmpdir, ignore_errors=True)
                return jsonify(error=str(e)), 400

            pdf_paths.append(pdf_path)
            try:
                pages = pdf_page_count(pdf_path)
            except Exception:
                pages = 0
            if pages > FREE_MAX_PAGES:
                reasons.append(f"{os.path.basename(path)} has {pages} pages (> {FREE_MAX_PAGES})")
                need_payment = True

        # file-count rule
        if len(saved) > FREE_MAX_FILES:
            reasons.append(f"More than {FREE_MAX_FILES} files in one go")
            need_payment = True

        # Store session
        token = secrets.token_urlsafe(16)
        SESSION_STORE[token] = {
            "dir": tmpdir,
            "files": pdf_paths,
            "created": time.time()
        }

        if not need_payment or not RZP_ENABLED:
            # Free OR Razorpay not configured => allow free
            return jsonify(payment_required=False, token=token)

        # Create Razorpay order
        receipt = f"w2p_{token}"
        order = client.order.create(dict(
            amount=PAID_AMOUNT_PAISE,
            currency="INR",
            receipt=receipt,
            payment_capture=1
        ))
        order_id = order["id"]
        ORDER_MAP[order_id] = token

        msg = "Payment required (₹10)"
        if reasons:
            msg += ": " + "; ".join(reasons)

        return jsonify(
            payment_required=True,
            amount_rupees=PAID_AMOUNT_RUPEES,
            amount_paise=PAID_AMOUNT_PAISE,
            key_id=RZP_KEY_ID,
            order_id=order_id,
            token=token,
            message=msg
        )
    except Exception as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        return jsonify(error=str(e)), 500

@app.route("/verify", methods=["POST"])
def verify():
    data = request.get_json(silent=True) or {}
    token = data.get("token")
    order_id = data.get("razorpay_order_id")
    payment_id = data.get("razorpay_payment_id")
    signature = data.get("razorpay_signature")

    if not (token and order_id and payment_id and signature):
        return jsonify(error="Missing verification fields."), 400
    if ORDER_MAP.get(order_id) != token:
        return jsonify(error="Order-token mismatch."), 400
    if token not in SESSION_STORE:
        return jsonify(error="Session expired."), 400

    # Verify signature
    try:
        params = {
            "razorpay_order_id": order_id,
            "razorpay_payment_id": payment_id,
            "razorpay_signature": signature
        }
        client.utility.verify_payment_signature(params)
    except Exception as e:
        return jsonify(error="Signature verification failed."), 400

    # OK
    return jsonify(ok=True, download_url=f"/download/{token}")

@app.route("/download/<token>")
def download(token):
    info = SESSION_STORE.get(token)
    if not info:
        return "Session expired or invalid.", 410

    tmpdir = info["dir"]
    pdf_paths = info["files"]
    zip_path = os.path.join(tmpdir, "converted_pdfs.zip")
    if not os.path.exists(zip_path):
        try:
            make_zip(pdf_paths, zip_path)
        except Exception as e:
            return (f"Zip error: {e}", 500)

    @after_this_request
    def later(resp):
        # cleanup in 2 minutes
        cleanup_later(tmpdir, 120)
        # remove from memory stores
        try:
            del SESSION_STORE[token]
            # remove any order mapping pointing to token
            for k, v in list(ORDER_MAP.items()):
                if v == token:
                    del ORDER_MAP[k]
        except Exception:
            pass
        return resp

    return send_file(zip_path, as_attachment=True, download_name="converted_pdfs.zip")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
