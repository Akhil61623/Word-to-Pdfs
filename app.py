import os, io, tempfile, zipfile, shutil, uuid, re
from flask import Flask, request, send_file, render_template_string, jsonify, after_this_request
import mammoth               # DOCX -> HTML
from xhtml2pdf import pisa   # HTML -> PDF
import fitz                  # PyMuPDF, count PDF pages
import razorpay              # payments
from werkzeug.utils import secure_filename

app = Flask(__name__)

# ---------------- Config ----------------
BRAND = os.environ.get("APP_BRAND", "Mahamaya Stationery")

# Free rules
FREE_MAX_FILES  = 2
FREE_MAX_MB     = 10
FREE_MAX_PAGES  = 25

# Pricing
PRICE_INR       = int(os.environ.get("PRICE_INR", "10"))           # ₹10
RAZORPAY_KEY_ID = os.environ.get("RAZORPAY_KEY_ID", "")
RAZORPAY_SECRET = os.environ.get("RAZORPAY_KEY_SECRET", "")

# Razorpay client (only if keys present)
rz_client = None
if RAZORPAY_KEY_ID and RAZORPAY_SECRET:
    rz_client = razorpay.Client(auth=(RAZORPAY_KEY_ID, RAZORPAY_SECRET))

# ---------------- UI ----------------
INDEX_HTML = r"""
<!doctype html>
<html lang="hi">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>{{brand}} — Word → PDF</title>
<style>
  :root{
    --bg:#0b1220; --fg:#e7eaf1; --muted:#93a2bd; --card:#10182b;
    --accent:#4f8cff; --accent2:#22c55e; --danger:#ef4444; --stroke:#203054;
  }
  *{box-sizing:border-box}
  body{margin:0; font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial; background:var(--bg); color:var(--fg)}
  .shell{min-height:100svh; display:grid; place-items:center; padding:24px}
  .card{width:min(900px,100%); background:linear-gradient(180deg,#0f172a,#0b1220);
        border:1px solid var(--stroke); border-radius:20px; padding:24px; box-shadow:0 10px 40px rgba(0,0,0,.35)}
  .top{display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap}
  .brand{display:flex; align-items:center; gap:10px; font-weight:800; letter-spacing:.2px}
  .badge{font-size:12px; padding:2px 8px; border:1px solid var(--stroke); border-radius:999px; color:var(--muted)}
  h1{margin:6px 0 8px; font-size:24px}
  p.muted{color:var(--muted); margin:0 0 18px}

  .drop{border:2px dashed var(--stroke); border-radius:16px; padding:18px; background:#0e1830; text-align:center; transition:.2s}
  .drop.drag{border-color:var(--accent); background:#112042}
  .note{font-size:12px; color:var(--muted)}
  input[type=file]{display:none}

  .row{display:flex; gap:10px; align-items:center; flex-wrap:wrap; margin-top:12px}
  button.btn{display:inline-flex; align-items:center; gap:10px; padding:10px 14px; border-radius:12px;
    border:1px solid var(--stroke); background:var(--accent); color:#fff; font-weight:700; cursor:pointer}
  button.ghost{background:#17233f}
  button:disabled{opacity:.6; cursor:not-allowed}

  .alert{margin-top:10px; padding:10px 12px; border-radius:12px; font-weight:600; display:none}
  .alert.ok{background:rgba(34,197,94,.1); color:var(--accent2); border:1px solid rgba(34,197,94,.25)}
  .alert.err{background:rgba(239,68,68,.1); color:var(--danger); border:1px solid rgba(239,68,68,.25)}

  /* fancy loader */
  .overlay{position:fixed; inset:0; display:none; place-items:center; background:rgba(11,18,32,.85); z-index:50}
  .loader{width:72px; aspect-ratio:1; border-radius:50%;
    background:conic-gradient(from 180deg, #4f8cff, #22c55e);
    -webkit-mask:radial-gradient(farthest-side,#0000 52%,#000 53%);
    animation:spin 1s linear infinite}
  @keyframes spin{to{transform:rotate(1turn)}}
  .quotes{margin-top:14px; text-align:center; color:#c9d6f2; font-size:14px}
  .dim{color:#9fb0d3; font-size:12px}

  .small{font-size:12px; color:var(--muted)}
</style>
</head>
<body>
<div class="shell">
  <div class="card">
    <div class="top">
      <div class="brand">
        <div style="width:30px;height:30px;border-radius:8px;background:linear-gradient(135deg,#4f8cff, #22c55e)"></div>
        <div>{{brand}}</div>
      </div>
      <div class="badge">Word (DOCX) → PDF</div>
    </div>

    <h1>तेज़ और क्लीन Word → PDF कन्वर्टर</h1>
    <p class="muted">
      फ्री: अधिकतम <b>2 DOCX</b>, हर फाइल ≤ <b>{{free_mb}} MB</b>, और आउटपुट ≤ <b>{{free_pages}}</b> पेज।
      इससे ज़्यादा पर <b>₹{{price}}</b> Razorpay पेमेंट लगेगा।
    </p>

    <div id="drop" class="drop" tabindex="0">
      <strong>Drag & Drop</strong> <span class="note">या क्लिक करके DOCX चुनें (एक साथ कई)</span>
      <input id="files" type="file" accept=".docx" multiple />
      <div id="chosen" class="note" style="margin-top:8px"></div>
    </div>

    <div class="row">
      <button id="convertBtn" class="btn">Convert & Download ZIP</button>
      <button id="chooseBtn" class="btn ghost">Choose Files</button>
      <div id="status" class="note"></div>
    </div>

    <div id="ok" class="alert ok">डाउनलोड शुरू हो गया।</div>
    <div id="err" class="alert err">Error</div>

    <p class="small">टिप: पासवर्ड-प्रोटेक्टेड DOCX सपोर्टेड नहीं हैं—ऐसे केस में साफ़ एरर दिखेगा।</p>
  </div>
</div>

<!-- full-screen loader -->
<div id="overlay" class="overlay">
  <div>
    <div class="loader" style="margin-inline:auto"></div>
    <div class="quotes" id="quote">Preparing your PDFs…</div>
    <div class="dim" id="sub">Quality conversion in progress</div>
  </div>
</div>

<script src="https://checkout.razorpay.com/v1/checkout.js"></script>
<script>
const FREE_MAX_FILES = {{free_files}};
const FREE_MAX_MB    = {{free_mb}};
const PRICE_PAISE    = {{price}} * 100; // just for display if needed

const drop     = document.getElementById('drop');
const filesEl  = document.getElementById('files');
const chooseBtn= document.getElementById('chooseBtn');
const convertBtn= document.getElementById('convertBtn');
const chosen   = document.getElementById('chosen');
const ok       = document.getElementById('ok');
const err      = document.getElementById('err');
const statusEl = document.getElementById('status');
const overlay  = document.getElementById('overlay');
const quoteEl  = document.getElementById('quote');
const subEl    = document.getElementById('sub');

const QUOTES = [
  "Formatting tables…",
  "Embedding images…",
  "Balancing fonts & layout…",
  "Almost there… prepping your ZIP…",
];

function show(el,msg){ el.textContent=msg; el.style.display='block'; }
function hide(el){ el.style.display='none'; }
function clearAlerts(){ hide(ok); hide(err); statusEl.textContent=''; }

function startOverlay(){
  overlay.style.display='grid';
  let i=0;
  quoteEl.textContent = QUOTES[i%QUOTES.length];
  const id = setInterval(()=>{
    i++; quoteEl.textContent = QUOTES[i%QUOTES.length];
  }, 1200);
  overlay.dataset.timer = id;
}
function stopOverlay(){
  const id = overlay.dataset.timer;
  if(id) clearInterval(id);
  overlay.style.display='none';
}

drop.addEventListener('click', ()=> filesEl.click());
chooseBtn.addEventListener('click', ()=> filesEl.click());

['dragenter','dragover'].forEach(ev=>{
  drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('drag'); });
});
['dragleave','drop'].forEach(ev=>{
  drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('drag'); });
});
drop.addEventListener('drop', e=>{
  e.preventDefault();
  filesEl.files = e.dataTransfer.files;
  listFiles();
});
filesEl.addEventListener('change', listFiles);

function listFiles(){
  clearAlerts();
  if(!filesEl.files.length){ chosen.textContent=''; return; }
  const names = [...filesEl.files].map(f=>`${f.name} · ${(f.size/1024/1024).toFixed(2)} MB`);
  chosen.textContent = names.join(' | ');
}

async function doConvert(extra = {}){
  if(!filesEl.files.length) { show(err,"कृपया DOCX फाइलें चुनें।"); return; }

  const fd = new FormData();
  for(const f of filesEl.files) fd.append('docs', f);
  if(extra.payment_id) fd.append('payment_id', extra.payment_id);
  if(extra.order_id)   fd.append('order_id', extra.order_id);

  convertBtn.disabled = true;
  startOverlay();

  try{
    const res = await fetch('/convert', { method:'POST', body: fd });
    if(res.status === 402){
      const data = await res.json();
      if(data.error === 'payment_required' && data.order_id){
        await launchRazorpay(data.order_id);
        // payment success triggers re-submit with payment details from handler
        return;
      }
    }
    if(!res.ok){
      const t = await res.text(); throw new Error(t || ('HTTP '+res.status));
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'converted_pdfs.zip';
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    show(ok, "डाउनलोड शुरू हो गया।");
  }catch(e){
    show(err, e.message || "Conversion failed.");
  }finally{
    stopOverlay();
    convertBtn.disabled = false;
  }
}

convertBtn.addEventListener('click', ()=> doConvert());

// Razorpay
async function launchRazorpay(order_id){
  const options = {
    key: "{{rz_key}}",
    amount: {{price}} * 100,
    currency: "INR",
    name: "{{brand}}",
    description: "Word → PDF (Premium)",
    order_id,
    handler: function (response) {
      // re-submit with payment proof
      doConvert({ payment_id: response.razorpay_payment_id, order_id: order_id });
    },
    modal: { ondismiss: function(){ show(err, "पेमेंट कैंसल हुआ।"); } },
    theme: { color: "#4f8cff" }
  };
  const rzp = new Razorpay(options);
  rzp.open();
}
</script>
</body>
</html>
"""

# ---------------- Helpers ----------------
def is_docx_secure(file_bytes: bytes) -> bool:
    # very basic encrypted DOCX check: encrypted files contain "EncryptedPackage"
    # to avoid heavy parsing; mammoth will fail anyway, but we show nicer error.
    return b"EncryptedPackage" in file_bytes[:4096] or b"drs:encryption" in file_bytes

def html_to_pdf_bytes(html: str) -> bytes:
    out = io.BytesIO()
    # xhtml2pdf expects a file-like
    pisa.CreatePDF(src=io.StringIO(html), dest=out)
    return out.getvalue()

def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes:
    # DOCX -> HTML
    try:
        result = mammoth.convert_to_html(io.BytesIO(docx_bytes))
        html = result.value
    except Exception as e:
        raise RuntimeError(f"DOCX पढ़ने में समस्या: {e}")
    # HTML -> PDF
    pdf_bytes = html_to_pdf_bytes(html)
    if not pdf_bytes or len(pdf_bytes) < 1000:
        raise RuntimeError("PDF आउटपुट नहीं बन पाया (unsupported content).")
    return pdf_bytes

def count_pdf_pages(pdf_bytes: bytes) -> int:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        return doc.page_count

def mb(nbytes: int) -> float:
    return nbytes / 1024.0 / 1024.0

def need_payment_precheck(file_objs) -> (bool, str):
    # file-count and size checks before conversion
    if len(file_objs) > FREE_MAX_FILES:
        return True, f"Free में केवल {FREE_MAX_FILES} DOCX तक।"
    for f in file_objs:
        f.stream.seek(0, os.SEEK_END)
        size = f.stream.tell()
        f.stream.seek(0)
        if size > FREE_MAX_MB * 1024 * 1024:
            return True, f"{secure_filename(f.filename)} फाइल {FREE_MAX_MB}MB से बड़ी है।"
    return False, ""

def require_payment_response():
    if not rz_client:
        return ("Payment required but Razorpay keys not configured.", 500)
    # amount in paise
    amount_paise = PRICE_INR * 100
    order = rz_client.order.create(dict(amount=amount_paise, currency="INR", payment_capture=1, notes={"purpose":"word2pdf"}))
    return jsonify({"error":"payment_required", "order_id":order["id"], "amount":amount_paise}), 402

def has_payment(request) -> bool:
    return (request.form.get("payment_id") and request.form.get("order_id"))

# ---------------- Routes ----------------
@app.route("/")
def home():
    return render_template_string(
        INDEX_HTML,
        brand=BRAND, free_mb=FREE_MAX_MB, free_pages=FREE_MAX_PAGES, free_files=FREE_MAX_FILES,
        price=PRICE_INR, rz_key=RAZORPAY_KEY_ID
    )

@app.route("/healthz")
def health():
    return "OK"

@app.route("/convert", methods=["POST"])
def convert():
    files = request.files.getlist("docs")
    if not files:
        return ("कोई DOCX अपलोड नहीं किया।", 400)

    # Pre-check (count & size)
    need_pay, reason = need_payment_precheck(files)
    paid = has_payment(request)
    if need_pay and not paid:
        return require_payment_response()

    tmpdir = tempfile.mkdtemp(prefix="doc2pdf_")
    @after_this_request
    def cleanup(resp):
        shutil.rmtree(tmpdir, ignore_errors=True)
        return resp

    # Convert all, and track pages
    outputs = []
    total_pages_exceeded = False
    err_names = []

    for f in files:
        name = secure_filename(f.filename or "input.docx")
        try:
            data = f.read()
            if is_docx_secure(data):
                raise RuntimeError("यह DOCX पासवर्ड-प्रोटेक्टेड/एन्क्रिप्टेड लगता है। पहले अनलॉक करें।")
            pdf_bytes = convert_docx_to_pdf_bytes(data)
            pages = count_pdf_pages(pdf_bytes)
            if pages > FREE_MAX_PAGES:
                total_pages_exceeded = True
            out_name = os.path.splitext(name)[0] + ".pdf"
            out_path = os.path.join(tmpdir, out_name)
            with open(out_path, "wb") as w:
                w.write(pdf_bytes)
            outputs.append((out_name, out_path))
        except Exception as e:
            err_names.append(f"{name}: {e}")

    if err_names:
        return ("कन्वर्ज़न एरर:\n" + "\n".join(err_names), 400)

    # If any page-limit exceeded and not paid -> ask payment
    if total_pages_exceeded and not paid:
        return require_payment_response()

    # Make ZIP
    zpath = os.path.join(tmpdir, "converted_pdfs.zip")
    with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for nm, pth in outputs:
            zf.write(pth, nm)

    return send_file(zpath, as_attachment=True, download_name="converted_pdfs.zip")
    

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
