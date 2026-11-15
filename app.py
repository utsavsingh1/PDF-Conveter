import os, re, uuid, json, tempfile, csv
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import fitz  # PyMuPDF
import cv2
import numpy as np
from paddleocr import PaddleOCR, PPStructure
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPTpt, Emu
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.text import PP_ALIGN
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors

# Flask serves both frontend (static files) and backend API from http://127.0.0.1:5000
app = Flask(__name__, static_folder=".", static_url_path="")
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=False)

UPLOAD_FOLDER  = tempfile.gettempdir()
LANG_MAP       = {"en": "en", "hi": "hi", "en+hi": "ch"}
ALLOWED_FORMATS = {"txt", "docx", "xlsx", "pptx", "json", "html", "csv", "pdf_searchable"}

# ── OCR Engine (lazy init) ──
_ocr_cache       = {}
_structure_cache = {}

def get_ocr(lang):
    if lang not in _ocr_cache:
        _ocr_cache[lang] = PaddleOCR(use_angle_cls=True, lang=lang, use_gpu=False, show_log=False)
    return _ocr_cache[lang]

def get_structure(lang):
    if lang not in _structure_cache:
        _structure_cache[lang] = PPStructure(show_log=False, lang=lang)
    return _structure_cache[lang]

# ── Image Preprocessing ──
def preprocess_image(img_array, denoising=True):
    gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
    if denoising:
        gray = cv2.fastNlMeansDenoising(gray, h=10)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    thresh = cv2.dilate(thresh, np.ones((1, 1), np.uint8), iterations=1)
    return cv2.cvtColor(thresh, cv2.COLOR_GRAY2BGR)

# ── Rules Engine ──
HINDI_CORRECTIONS   = {r'।।': '।', r'\s{2,}': ' '}
ENGLISH_CORRECTIONS = {r'\bl\b(?=\s+[a-z])': 'I', r'(?<=[a-z])0(?=[a-z])': 'o', r'\s{2,}': ' '}

def apply_rules(text, lang):
    if not text or not text.strip():
        return ""
    corrections = {}
    if 'hi' in lang:
        corrections.update(HINDI_CORRECTIONS)
    if 'en' in lang or lang == 'en':
        corrections.update(ENGLISH_CORRECTIONS)
    for pattern, replacement in corrections.items():
        text = re.sub(pattern, replacement, text)
    text = re.sub(r'(?<!\n)\n(?!\n)', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

# ── PDF → Images ──
def pdf_to_images(pdf_path, dpi=200):
    doc = fitz.open(pdf_path)
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72))
        if pix.n == 1:
            pix = fitz.Pixmap(fitz.csRGB, pix)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
        elif pix.n == 3:
            img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        images.append(img)
    doc.close()
    return images

# ── OCR Pipeline ──
def run_ocr_pipeline(images, lang, segmentation, rules_engine, table_detect, denoising, doc_type):
    paddle_lang = LANG_MAP.get(lang, "en")
    ocr         = get_ocr(paddle_lang)
    structure   = get_structure(paddle_lang) if segmentation else None
    all_pages   = []

    for img in images:
        processed = preprocess_image(img, denoising)
        page_text = ""

        if segmentation and structure:
            blocks = []
            for region in structure(processed):
                rtype = region.get("type", "text").lower()
                if rtype in ("title", "text", "header", "footer"):
                    for line in region.get("res", []):
                        if isinstance(line, (list, tuple)) and len(line) == 2:
                            blocks.append(line[1][0])
                elif rtype == "table" and table_detect:
                    cells = re.findall(r'<td[^>]*>(.*?)</td>',
                                       region.get("res", {}).get("html", ""), re.DOTALL)
                    row = " | ".join(c.strip() for c in cells if c.strip())
                    if row:
                        blocks.append(row)
            page_text = "\n".join(blocks)
        else:
            result = ocr.ocr(processed, cls=True)
            if result and result[0]:
                page_text = "\n".join(line[1][0] for line in result[0] if line and line[1])

        if rules_engine:
            page_text = apply_rules(page_text, lang)
        all_pages.append(page_text or "")

    return all_pages

# ── Helpers ──
def _set_run_color(run, r, g, b):
    run.font.color.rgb = RGBColor(r, g, b)

# ── Output Generators ──

def to_txt(pages, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        for i, page in enumerate(pages):
            f.write(f"{'='*60}\n  PAGE {i+1}\n{'='*60}\n\n{page.strip()}\n\n")


def to_docx(pages, out_path):
    doc = Document()
    title_para = doc.add_heading("Converted Document", 0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title_para.runs:
        _set_run_color(title_para.runs[0], 0x6C, 0x63, 0xFF)

    for i, page in enumerate(pages):
        if i > 0:
            doc.add_page_break()
        h = doc.add_heading(f"Page {i+1}", level=1)
        if h.runs:
            _set_run_color(h.runs[0], 0x6C, 0x63, 0xFF)
        for line in page.split("\n"):
            line = line.strip()
            if not line:
                continue
            if " | " in line:
                cells = [c.strip() for c in line.split(" | ")]
                t = doc.add_table(rows=1, cols=len(cells))
                t.style = "Table Grid"
                for j, cell_text in enumerate(cells):
                    cell_obj = t.rows[0].cells[j]
                    cell_obj.text = cell_text
                    if cell_obj.paragraphs[0].runs:
                        cell_obj.paragraphs[0].runs[0].font.size = Pt(10)
            else:
                p = doc.add_paragraph(line)
                if p.runs:
                    p.runs[0].font.size = Pt(11)
    doc.save(out_path)


def to_xlsx(pages, out_path):
    wb          = openpyxl.Workbook()
    HEADER_FILL = PatternFill("solid", fgColor="6C63FF")
    ALT_FILL    = PatternFill("solid", fgColor="F0EFFF")
    WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")
    TITLE_FILL  = PatternFill("solid", fgColor="1A1A2E")
    HEADER_FONT = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    TITLE_FONT  = Font(bold=True, color="6C63FF", size=14, name="Calibri")
    BODY_FONT   = Font(size=10, name="Calibri")
    BOLD_FONT   = Font(bold=True, size=10, name="Calibri")
    CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT        = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin        = Side(style="thin", color="CCCCCC")
    BORDER      = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, page in enumerate(pages):
        ws = wb.create_sheet(title=f"Page {i+1}") if i > 0 else wb.active
        ws.title = f"Page {i+1}"
        ws.sheet_view.showGridLines = False

        # Title row
        ws.merge_cells("A1:E1")
        tc = ws["A1"]
        tc.value     = f"PDF Conversion \u2014 Page {i+1}"
        tc.font      = TITLE_FONT
        tc.fill      = TITLE_FILL
        tc.alignment = CENTER
        ws.row_dimensions[1].height = 32

        # Header row
        for col, h in enumerate(["#", "Line", "Content", "Word Count", "Type"], 1):
            cell            = ws.cell(row=2, column=col, value=h)
            cell.font       = HEADER_FONT
            cell.fill       = HEADER_FILL
            cell.alignment  = CENTER
            cell.border     = BORDER
        ws.row_dimensions[2].height = 22

        # Data rows
        data_row = 3
        line_num = 0
        for line in page.split("\n"):
            line = line.strip()
            if not line:
                continue
            line_num += 1
            row_type = "Table" if " | " in line else "Text"
            fill     = ALT_FILL if data_row % 2 == 0 else WHITE_FILL
            for col, val in enumerate([i+1, line_num, line, len(line.split()), row_type], 1):
                cell           = ws.cell(row=data_row, column=col, value=val)
                cell.font      = BODY_FONT
                cell.fill      = fill
                cell.border    = BORDER
                cell.alignment = LEFT if col == 3 else CENTER
            ws.row_dimensions[data_row].height = 18
            data_row += 1

        # Column widths + freeze
        for col, width in {1: 6, 2: 8, 3: 60, 4: 12, 5: 10}.items():
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.freeze_panes = "A3"

        # Summary
        ws.cell(row=data_row + 1, column=1, value="Total Lines:").font = BOLD_FONT
        ws.cell(row=data_row + 1, column=2, value=line_num).font = Font(bold=True, color="6C63FF", size=10)

    wb.save(out_path)


def to_pptx(pages, out_path):
    prs    = Presentation()
    W, H   = prs.slide_width, prs.slide_height
    ACCENT = PRGBColor(0x6C, 0x63, 0xFF)
    DARK   = PRGBColor(0x0A, 0x0A, 0x0F)
    WHITE  = PRGBColor(0xFF, 0xFF, 0xFF)
    blank  = prs.slide_layouts[6]

    # Title slide
    ts = prs.slides.add_slide(blank)
    ts.background.fill.solid()
    ts.background.fill.fore_color.rgb = DARK
    txb = ts.shapes.add_textbox(PInches(1), PInches(2.5), W - PInches(2), PInches(1.5))
    tf  = txb.text_frame
    tf.word_wrap = True
    p0  = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    r0  = p0.add_run()
    r0.text           = "PDF Converted Document"
    r0.font.size      = PPTpt(36)
    r0.font.bold      = True
    r0.font.color.rgb = ACCENT
    sub = tf.add_paragraph()
    sub.alignment = PP_ALIGN.CENTER
    rs  = sub.add_run()
    rs.text           = f"{len(pages)} page(s) \u00b7 PaddleOCR Pipeline"
    rs.font.size      = PPTpt(16)
    rs.font.color.rgb = WHITE

    for i, page in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = PRGBColor(0x12, 0x12, 0x1A)

        badge = slide.shapes.add_textbox(PInches(0.3), PInches(0.2), PInches(1.2), PInches(0.4))
        br    = badge.text_frame.paragraphs[0].add_run()
        br.text           = f"Page {i+1}"
        br.font.size      = PPTpt(11)
        br.font.bold      = True
        br.font.color.rgb = ACCENT

        ls = slide.shapes.add_shape(1, PInches(0.3), PInches(0.7), PInches(1.5), Emu(36000))
        ls.fill.solid()
        ls.fill.fore_color.rgb = ACCENT
        ls.line.fill.background()

        txb = slide.shapes.add_textbox(PInches(0.4), PInches(0.9), W - PInches(0.8), H - PInches(1.2))
        tf  = txb.text_frame
        tf.word_wrap = True
        first = True
        for line in page.split("\n"):
            line = line.strip()
            if not line:
                continue
            para  = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            run   = para.add_run()
            run.text           = line
            run.font.size      = PPTpt(13)
            run.font.color.rgb = WHITE

    prs.save(out_path)


def to_json(pages, out_path):
    data = {
        "meta": {"engine": "PaddleOCR", "pages": len(pages), "accuracy": "90-95%"},
        "pages": [
            {
                "page": i + 1,
                "lines": [l.strip() for l in p.split("\n") if l.strip()],
                "word_count": len(p.split()),
                "char_count": len(p),
            }
            for i, p in enumerate(pages)
        ],
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def to_html(pages, out_path):
    page_parts = []
    for i, page in enumerate(pages):
        line_parts = []
        for line in page.split("\n"):
            line = line.strip()
            if not line:
                continue
            if " | " in line:
                tds = "".join(f"<td>{c.strip()}</td>" for c in line.split(" | "))
                line_parts.append(f"<tr>{tds}</tr>")
            else:
                line_parts.append(f"<p>{line}</p>")
        inner = "".join(line_parts)
        inner = re.sub(r'(<tr>.*?</tr>)+',
                       lambda m: f'<table>{m.group(0)}</table>', inner, flags=re.DOTALL)
        page_parts.append(
            f'<section class="page"><div class="page-badge">Page {i+1}</div>{inner}</section>'
        )
    body = "".join(page_parts)
    html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Converted Document</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:'Segoe UI',sans-serif;background:#0a0a0f;color:#e8e8f0;padding:40px 24px}}
  .container{{max-width:860px;margin:auto}}
  h1{{font-size:2rem;font-weight:900;color:#6c63ff;margin-bottom:8px}}
  .subtitle{{color:#8888aa;margin-bottom:40px;font-size:.9rem}}
  .page{{background:#12121a;border:1px solid rgba(255,255,255,.08);border-radius:16px;padding:32px;margin-bottom:28px}}
  .page-badge{{display:inline-block;background:rgba(108,99,255,.15);color:#6c63ff;border:1px solid rgba(108,99,255,.3);padding:4px 14px;border-radius:100px;font-size:.8rem;font-weight:700;margin-bottom:16px}}
  p{{line-height:1.7;margin-bottom:8px;color:#d0d0e8}}
  table{{width:100%;border-collapse:collapse;margin:12px 0}}
  th,td{{padding:10px 14px;border:1px solid rgba(255,255,255,.1);font-size:.88rem}}
  tr:nth-child(even){{background:rgba(108,99,255,.06)}}
  tr:first-child td{{background:rgba(108,99,255,.2);font-weight:700;color:#6c63ff}}
  footer{{text-align:center;color:#555;font-size:.8rem;margin-top:40px}}
</style></head>
<body><div class="container">
  <h1>&#128196; Converted Document</h1>
  <p class="subtitle">Extracted via PaddleOCR &middot; 90&ndash;95% Accuracy Pipeline</p>
  {body}
  <footer>Generated by PDFAny &middot; PaddleOCR Powered</footer>
</div></body></html>"""
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)


def to_csv(pages, out_path):
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Page", "Line No", "Text", "Word Count", "Has Table"])
        for pi, page in enumerate(pages):
            li = 0
            for line in page.split("\n"):
                line = line.strip()
                if not line:
                    continue
                li += 1
                w.writerow([pi + 1, li, line, len(line.split()), "Yes" if " | " in line else "No"])


# Pre-register ReportLab styles once to avoid duplicate-name crash
_rl_styles   = getSampleStyleSheet()
_TITLE_STYLE = ParagraphStyle("pdf_title", parent=_rl_styles["Title"],
                               textColor=colors.HexColor("#6C63FF"), fontSize=20, spaceAfter=6)
_BODY_STYLE  = ParagraphStyle("pdf_body",  parent=_rl_styles["Normal"],
                               fontSize=10, leading=16, spaceAfter=4)
_PAGE_STYLE  = ParagraphStyle("pdf_page",  parent=_rl_styles["Heading2"],
                               textColor=colors.HexColor("#6C63FF"), fontSize=13,
                               spaceBefore=14, spaceAfter=6)
_TABLE_STYLE = TableStyle([
    ("BACKGROUND",     (0, 0), (-1, 0),  colors.HexColor("#6C63FF")),
    ("TEXTCOLOR",      (0, 0), (-1, 0),  colors.white),
    ("FONTSIZE",       (0, 0), (-1, -1), 9),
    ("GRID",           (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F0EFFF")]),
    ("PADDING",        (0, 0), (-1, -1), 6),
])


def _flush_table(story, table_rows):
    if not table_rows:
        return
    max_cols = max(len(r) for r in table_rows)
    padded   = [r + [""] * (max_cols - len(r)) for r in table_rows]
    t = Table(padded, hAlign="LEFT")
    t.setStyle(_TABLE_STYLE)
    story.append(t)
    story.append(Spacer(1, 8))


def to_searchable_pdf(pages, out_path):
    doc   = SimpleDocTemplate(out_path, pagesize=A4,
                               leftMargin=40, rightMargin=40, topMargin=50, bottomMargin=40)
    story = [Paragraph("Converted Document", _TITLE_STYLE), Spacer(1, 20)]
    for i, page in enumerate(pages):
        story.append(Paragraph(f"Page {i+1}", _PAGE_STYLE))
        table_rows = []
        for line in page.split("\n"):
            line = line.strip()
            if not line:
                continue
            if " | " in line:
                table_rows.append([c.strip() for c in line.split(" | ")])
            else:
                _flush_table(story, table_rows)
                table_rows = []
                story.append(Paragraph(line, _BODY_STYLE))
        _flush_table(story, table_rows)
        story.append(Spacer(1, 16))
    doc.build(story)


FORMAT_HANDLERS = {
    "txt":            (to_txt,            "text/plain"),
    "docx":           (to_docx,           "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    "xlsx":           (to_xlsx,           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    "pptx":           (to_pptx,           "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
    "json":           (to_json,           "application/json"),
    "html":           (to_html,           "text/html"),
    "csv":            (to_csv,            "text/csv"),
    "pdf_searchable": (to_searchable_pdf, "application/pdf"),
}


# ── Routes ──

@app.route("/")
def index():
    return send_file("index.html")


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "engine": "PaddleOCR"})


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported"}), 400

    fmt = request.form.get("format", "docx")
    if fmt not in ALLOWED_FORMATS:
        return jsonify({"error": "Invalid output format"}), 400

    lang         = request.form.get("lang", "en")
    doc_type     = request.form.get("doc_type", "printed")
    segmentation = request.form.get("segmentation", "true") == "true"
    rules_engine = request.form.get("rules_engine", "true") == "true"
    table_detect = request.form.get("table_detect", "true") == "true"
    denoising    = request.form.get("denoising", "true") == "true"

    uid      = uuid.uuid4().hex
    pdf_path = os.path.join(UPLOAD_FOLDER, f"{uid}.pdf")
    ext      = "pdf" if fmt == "pdf_searchable" else fmt
    out_path = os.path.join(UPLOAD_FOLDER, f"{uid}_out.{ext}")

    f.save(pdf_path)

    try:
        images  = pdf_to_images(pdf_path)
        pages   = run_ocr_pipeline(images, lang, segmentation, rules_engine,
                                   table_detect, denoising, doc_type)
        handler, mime = FORMAT_HANDLERS[fmt]
        handler(pages, out_path)

        base     = os.path.splitext(secure_filename(f.filename))[0]
        dl_name  = f"{base}_converted.{ext}"
        response = send_file(out_path, mimetype=mime, as_attachment=True, download_name=dl_name)
        return response

    except Exception as e:
        return jsonify({"error": str(e)}), 500

    finally:
        # Always clean up both temp files after response is sent
        for p in (pdf_path, out_path):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except OSError:
                pass


if __name__ == "__main__":
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    # host=0.0.0.0 so it's reachable from both 127.0.0.1 and localhost
    app.run(debug=debug, host="0.0.0.0", port=5000)
