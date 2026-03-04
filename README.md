# worship_lead_assistant

## 1) Current MVP: PDF Worship Sheet → Lyrics PowerPoint

This workspace now includes a working web app with a modern UI where a user can:

- upload one or multiple PDF worship sheets
- extract likely lyrics text from each PDF
- filter out likely chord-only lines and common metadata lines
- generate simple PowerPoint lyrics slides
- use OCR fallback for scanned/image-based PDFs when normal text extraction is insufficient
- output slides with black background and white lyric text
- download:
	- one `.pptx` file when one PDF is uploaded, or
	- one `.zip` containing multiple `.pptx` files when multiple PDFs are uploaded

---

## 2) Tech Stack

- Backend: FastAPI
- PDF text extraction: pypdf
- PowerPoint generation: python-pptx
- Frontend: HTML/CSS/vanilla JS (modern glassmorphism-style UI)

---

## 3) Project Structure

- `app/main.py` — API + conversion logic
- `app/static/index.html` — frontend
- `requirements.txt` — dependencies

---

## 4) Run Locally

### Prerequisites

- Python 3.10+

### Install dependencies

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Install Tesseract OCR (required for OCR fallback)

On macOS:

```bash
brew install tesseract
```

### Start server

```bash
uvicorn app.main:app --reload
```

Then open:

- http://127.0.0.1:8000

---

## 5) Notes / Limitations

- PDF parsing quality depends on the PDF text layer.
- Scanned-image PDFs use OCR fallback (requires Tesseract installed), but quality may vary by scan quality.
- Chord removal uses heuristics; edge cases may remain and can be refined.

---

## 6) Original Vision (Future Scope)

The broader assistant can still be extended to:

- recommend worship songs by sermon/preaching topic
- find sheet music from SharePoint
- generate templated slide decks automatically