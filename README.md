# PDF to Any — PaddleOCR Powered Converter

## Accuracy Pipeline
| Engine                           | Accuracy      |
|----------------------------------|---------------|
| Tesseract                        | 60–70%        |
| PaddleOCR Raw                    | 80–88%        |
| PaddleOCR + Segmentation + Rules | **90–95%**   |

## Quick Start (Recommended)
```
Double-click start.bat
```
Opens http://localhost:5000 automatically.

## Manual Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Start Flask (serves frontend + backend)
```bash
python app.py
```

### 3. Open in browser
```
http://localhost:5000
```

> ✅ Only ONE server needed. Flask serves both the UI and the /convert API.

## Features
- Hindi + Handwritten PDF support
- 8 output formats: DOCX, TXT, XLSX, PPTX, Searchable PDF, JSON, HTML, CSV
- Layout Segmentation via PPStructure
- Rules Engine for post-OCR correction
- Table Detection
- Image Denoising & Enhancement
- Auto backend health check on page load

## Notes
- For GPU: install `paddlepaddle-gpu` instead of `paddlepaddle`
- First run downloads PaddleOCR models (~200MB)
- Files deleted from server after conversion
- Debug mode: set `FLASK_DEBUG=1` environment variable
