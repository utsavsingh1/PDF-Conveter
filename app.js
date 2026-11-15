// ── Config ──
const API = "http://127.0.0.1:5000";

// ── Typewriter ──
const words = ["Any Format", "Word (.docx)", "Excel (.xlsx)", "Hindi Text", "Handwritten", "Searchable PDF", "JSON / CSV"];
let wi = 0, ci = 0, deleting = false;
const tw = document.getElementById("typewriter");

function typewrite() {
  ci = Math.max(0, ci);
  tw.textContent = deleting ? words[wi].slice(0, ci--) : words[wi].slice(0, ci++);
  if (!deleting && ci > words[wi].length) { deleting = true; setTimeout(typewrite, 1800); return; }
  if (deleting && ci <= 0) { ci = 0; deleting = false; wi = (wi + 1) % words.length; }
  setTimeout(typewrite, deleting ? 60 : 100);
}
typewrite();

// ── DOM Refs ──
const zone       = document.getElementById("uploadZone");
const fileInput  = document.getElementById("fileInput");
const fileInfo   = document.getElementById("fileInfo");
const fileNameEl = document.getElementById("fileName");
const fileSizeEl = document.getElementById("fileSize");
const convertBtn = document.getElementById("convertBtn");
const progBar    = document.getElementById("progressBar");
const progLabel  = document.getElementById("progressLabel");

let selectedFile = null;
let progInterval = null;
let lastBlobUrl  = null;

// ── Upload Zone ──
zone.addEventListener("click", () => fileInput.click());
zone.addEventListener("dragover",  e => { e.preventDefault(); zone.classList.add("drag-over"); });
zone.addEventListener("dragleave", () => zone.classList.remove("drag-over"));
zone.addEventListener("drop", e => {
  e.preventDefault();
  zone.classList.remove("drag-over");
  handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => handleFile(fileInput.files[0]));

function handleFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".pdf")) {
    showError("Please upload a valid PDF file (.pdf).");
    return;
  }
  selectedFile = file;
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = `(${(file.size / 1024 / 1024).toFixed(2)} MB)`;
  fileInfo.style.display = "flex";
  convertBtn.disabled = false;
  hideError();
  hideResult();
}

// ── Advanced Panel ──
const advToggle = document.getElementById("advToggle");
const advPanel  = document.getElementById("advPanel");
advToggle.addEventListener("click", () => {
  advToggle.classList.toggle("open");
  advPanel.classList.toggle("open");
});

// ── Convert ──
convertBtn.addEventListener("click", async () => {
  if (!selectedFile) return;

  if (lastBlobUrl) { URL.revokeObjectURL(lastBlobUrl); lastBlobUrl = null; }

  setConverting(true);
  showProgress();

  const fmt      = document.getElementById("formatSelect").value;
  const seg      = document.getElementById("segmentation").checked;
  const rules    = document.getElementById("rulesEngine").checked;

  const formData = new FormData();
  formData.append("file",         selectedFile);
  formData.append("lang",         document.getElementById("langSelect").value);
  formData.append("format",       fmt);
  formData.append("doc_type",     document.getElementById("docType").value);
  formData.append("segmentation", seg);
  formData.append("rules_engine", rules);
  formData.append("table_detect", document.getElementById("tableDetect").checked);
  formData.append("denoising",    document.getElementById("denoising").checked);

  const controller = new AbortController();
  const timer      = setTimeout(() => controller.abort(), 5 * 60 * 1000);

  try {
    const res = await fetch(`${API}/convert`, {
      method: "POST",
      body:   formData,
      signal: controller.signal,
    });

    clearTimeout(timer);

    if (!res.ok) {
      let msg = `Server error (${res.status})`;
      try { const e = await res.json(); msg = e.error || msg; } catch (_) {}
      throw new Error(msg);
    }

    const blob  = await res.blob();
    const cd    = res.headers.get("Content-Disposition") || "";
    const match = cd.match(/filename[^;=\n]*=\s*(?:UTF-8''|"?)([^";\n]+)"?/i);
    const name  = match
      ? decodeURIComponent(match[1].trim())
      : `converted.${fmt === "pdf_searchable" ? "pdf" : fmt}`;

    lastBlobUrl = URL.createObjectURL(blob);
    finishProgress();
    showResult(name, fmt, lastBlobUrl, seg, rules);

  } catch (err) {
    clearTimeout(timer);
    hideProgress();
    showError(
      err.name === "AbortError"
        ? "Request timed out. The PDF may be too large or the server is busy."
        : err.message || "Conversion failed. Make sure python app.py is running."
    );
  } finally {
    setConverting(false);
  }
});

// ── Button State ──
function setConverting(active) {
  convertBtn.disabled = active;
  convertBtn.classList.toggle("loading", active);
  convertBtn.innerHTML = active
    ? '<i class="fa-solid fa-spinner fa-spin"></i> Processing...'
    : '<i class="fa-solid fa-rocket"></i> Convert Now';
}

// ── Progress ──
const STEPS  = ["step1", "step2", "step3", "step4", "step5"];
const LABELS = [
  "Uploading PDF...",
  "Running PaddleOCR...",
  "Segmenting Layout...",
  "Applying Rules Engine...",
  "Generating Output File...",
];

function showProgress() {
  document.getElementById("progressSection").style.display = "block";
  hideResult(); hideError();
  if (progInterval) { clearInterval(progInterval); progInterval = null; }
  STEPS.forEach(s => { document.getElementById(s).className = "step"; });
  document.querySelectorAll(".step-line").forEach(l => l.classList.remove("done"));
  progBar.style.width = "0%";
  let i = 0;
  progInterval = setInterval(() => {
    if (i >= STEPS.length) { clearInterval(progInterval); progInterval = null; return; }
    if (i > 0) {
      document.getElementById(STEPS[i - 1]).classList.replace("active", "done");
      document.querySelectorAll(".step-line")[i - 1]?.classList.add("done");
    }
    document.getElementById(STEPS[i]).classList.add("active");
    progLabel.textContent = LABELS[i];
    progBar.style.width   = `${((i + 1) / STEPS.length) * 90}%`;
    i++;
  }, 1200);
}

function finishProgress() {
  if (progInterval) { clearInterval(progInterval); progInterval = null; }
  STEPS.forEach(s => {
    document.getElementById(s).classList.remove("active");
    document.getElementById(s).classList.add("done");
  });
  document.querySelectorAll(".step-line").forEach(l => l.classList.add("done"));
  progBar.style.width   = "100%";
  progLabel.textContent = "Done! \u2713";
}

function hideProgress() {
  if (progInterval) { clearInterval(progInterval); progInterval = null; }
  document.getElementById("progressSection").style.display = "none";
}

// ── Result ──
function showResult(name, fmt, url, seg, rules) {
  const acc = (seg && rules) ? "90\u201395%" : (seg || rules) ? "85\u201390%" : "80\u201388%";
  document.getElementById("resultTitle").textContent   = "Conversion Complete!";
  document.getElementById("resultMeta").textContent    = `${name} \u00b7 ${fmt.replace("pdf_searchable", "pdf").toUpperCase()}`;
  document.getElementById("accuracyBadge").textContent = `\uD83C\uDFAF Estimated Accuracy: ${acc}`;
  const dlBtn    = document.getElementById("downloadBtn");
  dlBtn.href     = url;
  dlBtn.download = name;
  dlBtn.onclick  = () => setTimeout(() => {
    if (lastBlobUrl) { URL.revokeObjectURL(lastBlobUrl); lastBlobUrl = null; }
  }, 10000);
  document.getElementById("resultSection").style.display = "block";
}

function hideResult() { document.getElementById("resultSection").style.display = "none"; }

// ── Error ──
function showError(msg) {
  document.getElementById("errorMsg").textContent = msg;
  document.getElementById("errorSection").style.display = "flex";
}
function hideError() { document.getElementById("errorSection").style.display = "none"; }

// ── Accuracy bar animation on scroll ──
const observer = new IntersectionObserver(entries => {
  entries.forEach(entry => {
    if (!entry.isIntersecting) return;
    entry.target.querySelectorAll(".acc-bar").forEach(bar => {
      const w = bar.style.width;
      bar.style.width = "0%";
      requestAnimationFrame(() => { bar.style.width = w; });
    });
    observer.unobserve(entry.target);
  });
}, { threshold: 0.3 });
document.querySelectorAll(".accuracy-grid").forEach(el => observer.observe(el));

// ── Health check on page load ──
window.addEventListener("load", () => {
  const ctrl = new AbortController();
  setTimeout(() => ctrl.abort(), 4000);
  fetch(`${API}/health`, { signal: ctrl.signal })
    .then(r => {
      if (!r.ok) throw new Error("not ok");
      return r.json();
    })
    .then(() => hideError())
    .catch(() => showError(
      "\u26A0\uFE0F Backend not reachable. Open a terminal and run:  python app.py  \u2014 then refresh this page."
    ));
});
