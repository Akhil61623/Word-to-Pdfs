import io
import os
import json
import hmac
import base64
import hashlib
import tempfile
from datetime import datetime
from typing import List, Tuple

from flask import Flask, request, render_template_string, send_file, jsonify, after_this_request
from werkzeug.utils import secure_filename

import requests           # Razorpay Orders API (SDK नहीं)
import mammoth            # DOCX -> HTML
from xhtml2pdf import pisa  # HTML -> PDF
import fitz               # PyMuPDF (PDF page count के लिए)

app = Flask(__name__)

# ====== Config ======
FREE_MAX_FILES = 2
FREE_MAX_MB = 10
FREE_MAX_PAGES = 25
PAID_AMOUNT_INR = 10
RAZORPAY_KEY_ID = os.environ.get("RAZORPAY_KEY_ID", "")
RAZORPAY_KEY_SECRET = os.environ.get("RAZORPAY_KEY_SECRET", "")

# Razorpay in-memory “paid orders” store (सरल डेमो के लिए)
PAID_ORDERS = set()

# ====== HTML (Jinja-safe) ======
INDEX_HTML = r"""
{% raw %}
<!doctype html>
<html lang="hi">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Mahamaya Stationery — Word → PDF Converter</title>
<style>
:root{
  --bg:#0b1220; --fg:#e7eaf1; --muted:#93a2bd; --card:#10182b;
  --accent:#4f8cff; --acc2:#22c55e; --warn:#f59e0b; --danger:#ef4444; --stroke:#223052;
}
*{box-sizing:border-box}
body{margin:0; font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial; background:var(--bg); color:var(--fg)}
.wrap{min-height:100svh; display:grid; place-items:center; padding:20px}
.card{width:min(950px,100%); background:linear-gradient(180deg,#0f172a,#0b1220);
  border:1px solid var(--stroke); border-radius:18px; padding:20px; box-shadow:0 10px 30px rgba(0,0,0,.35)}
.header{display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap}
.brand{display:flex; align-items:center; gap:10px; font-weight:900}
.brand-badge{width:32px;height:32px;border-radius:8px;background:linear-gradient(135deg,#4f8cff,#22c55e)}
h1{font-size:22px; margin:8px 0}
p.muted{color:var(--muted); margin:0 0 10px}

.drop{border:2px dashed var(--stroke); background:#0f1830; border-radius:16px; padding:16px; text-align:center}
.drop.drag{border-color:var(--accent); background:#0f2146}
.note{font-size:12px; color:var(--muted)}
.row{display:flex; align-items:center; gap:8px; flex-wrap:wrap}

.btn{display:inline-flex; align-items:center; gap:8px; padding:10px 14px; border-radius:12px;
  border:1px solid var(--stroke); background:var(--accent); color:#fff; font-weight:700; cursor:pointer}
.btn.ghost{background:#17243f}
.btn.warn{background:var(--warn); color:#111}
.btn.ok{background:var(--acc2)}
.btn:disabled{opacity:.6; cursor:not-allowed}
.badge{font-size:12px; padding:2px 8px; border:1px solid var(--stroke); border-radius:999px; color:var(--muted)}

.alert{display:none; margin-top:8px; padding:10px 12px; border-radius:12px; font-weight:600}
.alert.ok{background:rgba(34,197,94,.08); color:var(--acc2); border:1px solid rgba(34,197,94,.25)}
.alert.err{background:rgba(239,68,68,.08); color:var(--danger); border:1px solid rgba(239,68,68,.25)}
.alert.warn{background:rgba(245,158,11,.08); color:var(--warn); border:1px solid rgba(245,158,11,.25)}

.loader{
  --s: 60px;
  width:var(--s); height:var(--s); border-radius:50%;
  border:6px solid rgba(255,255,255,.15); border-top-color:#fff;
  animation:spin 1s linear infinite; margin-left:6px
}
@keyframes spin{to{transform:rotate(360deg)}}
.progress{display:none; align-items:center; gap:10px; margin-top:8px}

small.tos{display:block; margin-top:10px; color:#8aa0c7}
ul{margin:6px 0 0 18px; color:#cdd7ee; font-size:14px}
li{margin:2px 0}
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <div class="header">
      <div class="brand">
        <div class="brand-badge"></div>
        <div>Mahamaya Stationery</div>
      </div>
      <div class="badge">Word (.docx) → PDF</div>
    </div>

    <h1>Word → PDF (Free ≤ 2 files · 10 MB · 25 pages; ऊपर ₹10)</h1>
    <p class="muted">DOCX चुनें; हम HTML→PDF में convert करते हैं ताकि LibreOffice की ज़रूरत न पड़े। Images/tables/text styles सपोर्टेड हैं।</p>

    <div id="drop" class="drop" tabindex="0">
      <strong>Drag & Drop</strong> <span class="note">या क्लिक करके DOCX चुनें</span>
      <input id="file" type="file" accept=".docx" multiple style="display:none" />
      <div id="chosen" class="note" style="margin-top:8px"></div>
    </div>

    <div style="height:10px"></div>

    <div class="row">
      <button class="btn" id="choose">Choose Files</button>
      <button class="btn ghost" id="convert">Convert to PDF</button>
      <div id="limits" class="badge">Free: ≤2 files, ≤10 MB, ≤25 pages</div>
    </div>

    <div class="progress" id="progress">
      <div class="loader"></div>
      <div class="note" id="ptext">Converting… please wait</div>
    </div>

    <div id="ok" class="alert ok">Done! Download will start.</div>
    <div id="warn" class="alert warn"></div>
    <div id="err" class="alert err"></div>

    <small class="tos">
      नोट: Pure-Python कन्वर्ज़न हर layout को 100% नहीं दोहराता, लेकिन सामान्य use-cases (text + images + tables) अच्छे से चलते हैं।
    </small>

    <ul>
      <li>Free: कुल 2 फ़ाइलें, 10 MB और 25 pages तक।</li>
      <li>ऊपर जाएँ तो Razorpay से ₹10 पेमेंट; उसके बाद convert अनलॉक।</li>
    </ul>
  </div>
</div>

<script>
const fileInput = document.getElementById('file');
const drop = document.getElementById('drop');
const choose = document.getElementById('choose');
const convertBtn = document.getElementById('convert');
const chosen = document.getElementById('chosen');
const ok = document.getElementById('ok');
const warn = document.getElementById('warn');
const err = document.getElementById('err');
const progress = document.getElementById('progress');
const ptext = document.getElementById('ptext');

let paidOrder = null;

function show(el, msg){ el.textContent = msg; el.style.display='block'; }
function hide(el){ el.style.display='none'; }

function resetAlerts(){
  hide(ok); hide(warn); hide(err);
}

function fmtMB(bytes){ return (bytes/1024/1024).toFixed(2) + ' MB'; }

drop.addEventListener('click', ()=> fileInput.click());
choose.addEventListener('click', ()=> fileInput.click());

['dragenter','dragover'].forEach(ev=>{
  drop.addEventListener(ev, e=>{ e.preventDefault(); drop.classList.add('drag'); });
});
['dragleave','drop'].forEach(ev=>{
  drop.addEventListener(ev, e=>{ e.preventDefault(); drop.classList.remove('drag'); });
});
drop.addEventListener('drop', e=>{
  e.preventDefault();
  if (e.dataTransfer.files?.length){
    fileInput.files = e.dataTransfer.files;
    listChosen();
  }
});
fileInput.addEventListener('change', listChosen);

function listChosen(){
  resetAlerts();
  if(!fileInput.files.length){ chosen.textContent = ''; return; }
  let total = 0, names=[];
  [...fileInput.files].forEach(f=>{ total += f.size; names.push(f.name); });
  chosen.textContent = `Selected: ${names.join(', ')} · Total ${fmtMB(total)}`;
}

async function createOrder(){
  const res = await fetch('/create_order', {method:'POST'});
  if(!res.ok) throw new Error(await res.text());
  return await res.json(); // {order_id, amount, key_id, currency}
}

function openRazorpay({order_id, amount, key_id, currency}){
  return new Promise((resolve,reject)=>{
    const s = document.createElement('script');
    s.src = 'https://checkout.razorpay.com/v1/checkout.js';
    s.onload = ()=>{
      const rzp = new window.Razorpay({
        key: key_id,
        amount: amount,
        currency: currency,
        name: "Mahamaya Stationery",
        description: "Word→PDF unlock",
        order_id: order_id,
        handler: function (resp) {
          // verify on server
          fetch('/verify_payment', {
            method:'POST',
            headers:{'Content-Type':'application/json'},
            body: JSON.stringify({
              razorpay_order_id: resp.razorpay_order_id,
              razorpay_payment_id: resp.razorpay_payment_id,
              razorpay_signature: resp.razorpay_signature
            })
          }).then(r=>r.json()).then(js=>{
            if(js.ok){
              paidOrder = resp.razorpay_order_id;
              resolve(true);
            }else{
              reject(new Error(js.error || 'Verification failed'));
            }
          }).catch(reject);
        },
        theme: { color: "#4f8cff" }
      });
      rzp.open();
    };
    s.onerror = ()=>reject(new Error('Razorpay load failed'));
    document.body.appendChild(s);
  });
}

convertBtn.addEventListener('click', async ()=>{
  resetAlerts();
  if(!fileInput.files.length) { show(err, "पहले DOCX फाइलें चुनें."); return; }

  // Build FormData
  const fd = new FormData();
  [...fileInput.files].forEach(f=> fd.append('files', f));
  if (paidOrder) fd.append('paid_order', paidOrder);

  try{
    progress.style.display='flex';
    ptext.textContent = "Converting… please wait";
    convertBtn.disabled = true;

    let res = await fetch('/convert', { method:'POST', body: fd });
    if(res.status === 402){
      // needs payment
      const data = await res.json(); // {error, reason}
      show(warn, (data.error || "Payment required") + " — ₹10 लगेगा।");
      ptext.textContent = "Waiting for payment…";
      const ord = await createOrder();
      await openRazorpay(ord);
      // retry convert after payment
      const fd2 = new FormData();
      [...fileInput.files].forEach(f=> fd2.append('files', f));
      fd2.append('paid_order', paidOrder);
      res = await fetch('/convert', { method:'POST', body: fd2 });
    }

    if(!res.ok){
      const t = await res.text();
      throw new Error(t || ('HTTP '+res.status));
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'converted.zip';
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    show(ok, "हो गया! डाउनलोड शुरू हो गया।");
  }catch(e){
    show(err, e.message || "Conversion failed");
  }finally{
    progress.style.display='none';
    convertBtn.disabled = false;
  }
});
</script>
</body>
</html>
{% endraw %}
"""

# ====== Helpers ======

def bytes_to_mb(n: int) -> float:
    return n / (1024.0 * 1024.0)

def docx_to_html_bytes(docx_path: str) -> bytes:
    """
    DOCX -> HTML (images inline data URIs) using mammoth.
    """
    with open(docx_path, "rb") as f:
        res = mammoth.convert_to_html(
            f,
            convert_image=mammoth.images.inline(
                mammoth.images.img_element(lambda image: {"src": image.read("base64")})
            )
        )
        html = res.value  # HTML string
    # Basic CSS to keep structure readable
    style = """
    <style>
      body{font-family: DejaVu Sans, Arial, sans-serif; font-size:11pt}
      h1,h2,h3{margin:8px 0}
      p{margin:6px 0}
      table{border-collapse:collapse; width:100%}
      td,th{border:1px solid #888; padding:4px; vertical-align:top}
      img{max-width:100%}
    </style>
    """
    return (style + html).encode("utf-8")

def html_to_pdf_bytes(html_bytes: bytes) -> bytes:
    """
    HTML -> PDF using xhtml2pdf (pisa). Returns PDF bytes.
    """
    src = io.BytesIO(html_bytes)
    out = io.BytesIO()
    pisa.CreatePDF(src, dest=out, encoding='utf-8')
    return out.getvalue()

def count_pdf_pages(pdf_bytes: bytes) -> int:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        return doc.page_count

def enforce_free_limits(num_files: int, total_mb: float, total_pages: int, paid_order: str|None) -> Tuple[bool,str]:
    if paid_order and paid_order in PAID_ORDERS:
        return True, ""
    if num_files > FREE_MAX_FILES:
        return False, f"Free limit: max {FREE_MAX_FILES} files"
    if total_mb > FREE_MAX_MB:
        return False, f"Free limit: max {FREE_MAX_MB} MB"
    if total_pages > FREE_MAX_PAGES:
        return False, f"Free limit: max {FREE_MAX_PAGES} pages"
    return True, ""

def make_zip(files: List[Tuple[str, bytes]]) -> bytes:
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files:
            zf.writestr(name, data)
    buf.seek(0)
    return buf.getvalue()

# ====== Routes ======

@app.route("/")
def home():
    return render_template_string(INDEX_HTML)

@app.route("/healthz")
def health():
    return "OK"

@app.route("/create_order", methods=["POST"])
def create_order():
    """
    Create a Razorpay order for ₹10 using REST API (no SDK).
    """
    if not (RAZORPAY_KEY_ID and RAZORPAY_KEY_SECRET):
        return jsonify({"error":"Razorpay keys not configured"}), 500
    amount = PAID_AMOUNT_INR * 100  # paise
    data = {
        "amount": amount,
        "currency": "INR",
        "receipt": f"mm-w2p-{int(datetime.utcnow().timestamp())}",
        "payment_capture": 1
    }
    resp = requests.post(
        "https://api.razorpay.com/v1/orders",
        auth=(RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET),
        headers={"Content-Type":"application/json"},
        data=json.dumps(data),
        timeout=30
    )
    if resp.status_code >= 300:
        return jsonify({"error": f"order_failed {resp.text}"}), 500
    rj = resp.json()
    return jsonify({
        "order_id": rj["id"],
        "amount": rj["amount"],
        "currency": rj["currency"],
        "key_id": RAZORPAY_KEY_ID
    })

@app.route("/verify_payment", methods=["POST"])
def verify_payment():
    """
    Verify Razorpay signature without SDK.
    """
    try:
        body = request.get_json(force=True)
        oid = body["razorpay_order_id"]
        pid = body["razorpay_payment_id"]
        sig = body["razorpay_signature"]
    except Exception:
        return jsonify({"ok": False, "error":"bad request"}), 400

    if not (RAZORPAY_KEY_SECRET and oid and pid and sig):
        return jsonify({"ok": False, "error":"missing"}), 400

    message = f"{oid}|{pid}".encode()
    digest = hmac.new(RAZORPAY_KEY_SECRET.encode(), msg=message, digestmod=hashlib.sha256).hexdigest()
    if hmac.compare_digest(digest, sig):
        PAID_ORDERS.add(oid)
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error":"invalid signature"}), 400

@app.route("/convert", methods=["POST"])
def convert():
    """
    Accept multiple DOCX files, convert each to PDF, enforce free limits, return ZIP.
    """
    files = request.files.getlist("files")
    if not files:
        return "No files", 400

    paid_order = request.form.get("paid_order") or None

    # basic size check
    total_bytes = 0
    saved_paths = []
    tmpdir = tempfile.mkdtemp(prefix="w2p_")

    @after_this_request
    def cleanup(resp):
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)
        return resp

    for f in files:
        if not f.filename.lower().endswith(".docx"):
            return "Only .docx allowed", 400
        total_bytes += f.content_length or 0
        path = os.path.join(tmpdir, secure_filename(f.filename))
        f.save(path)
        saved_paths.append(path)

    total_mb = bytes_to_mb(total_bytes)

    # First pass convert to PDF bytes & count pages
    pdf_results: List[Tuple[str, bytes]] = []
    total_pages = 0
    try:
        for idx, path in enumerate(saved_paths, 1):
            html_bytes = docx_to_html_bytes(path)
            pdf_bytes = html_to_pdf_bytes(html_bytes)
            pages = count_pdf_pages(pdf_bytes)
            total_pages += pages
            base = os.path.splitext(os.path.basename(path))[0]
            out_name = f"{base}.pdf"
            pdf_results.append((out_name, pdf_bytes))
    except Exception as e:
        return (f"Conversion error: {e}", 500)

    # Enforce limits (unless paid)
    ok, reason = enforce_free_limits(len(files), total_mb, total_pages, paid_order)
    if not ok:
        return jsonify({"error":"payment_required", "reason":reason}), 402

    # Pack into ZIP
    zip_bytes = make_zip(pdf_results)
    return send_file(
        io.BytesIO(zip_bytes),
        mimetype="application/zip",
        as_attachment=True,
        download_name="converted.zip"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
