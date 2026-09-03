# worship_lead_assistant

## 1) Current MVP: PDF Worship Sheet → Lyrics PowerPoint

This workspace now includes a working web app with a modern UI where a user can:

- upload one or multiple PDF worship sheets
- extract likely lyrics text from each PDF
- filter out likely chord-only lines and common metadata lines
- generate simple PowerPoint lyrics slides
- use OCR fallback for scanned/image-based PDFs when normal text extraction is insufficient
- output slides with style-guide defaults:
	- dark charcoal background (`#1A1A1A`) and soft white lyric text (`#F5F5F5`)
	- title + slide counter (`N/total`) in vocal-part colors (male/female/everyone)
	- optional red footer instruction note on the first slide of each song
- download:
	- one `.pptx` file when one PDF is uploaded, or
	- one `.pptx` with all songs combined when multiple PDFs are uploaded

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

## 6) Regression Test Scripts

Two one-command regression scripts are included under `scripts/`:

- `scripts/regression_pdf.py`
- `scripts/regression_youtube.py`

Use the configured Python environment command prefix:

```bash
/Users/zeliangxu/opt/anaconda3/envs/py39/bin/python scripts/regression_pdf.py
/Users/zeliangxu/opt/anaconda3/envs/py39/bin/python scripts/regression_youtube.py
```

Optional examples:

```bash
# Custom PDF input/output
/Users/zeliangxu/opt/anaconda3/envs/py39/bin/python scripts/regression_pdf.py \
	--input "samples/有一位神（D）.pdf" \
	--output "samples/regression_pdf_custom.pptx"

# Test only YouTube transcript mode for a specific URL
/Users/zeliangxu/opt/anaconda3/envs/py39/bin/python scripts/regression_youtube.py \
	--url "https://youtu.be/n0FBb6hnwTo?si=vlEZ8asfdAp1S9oW" \
	--source transcript
```

Each script prints `PASS/FAIL` and key output metadata (status code, output size,
basic first-slide style checks).

---

## 7) Original Vision (Future Scope)

The broader assistant can still be extended to:

- recommend worship songs by sermon/preaching topic
- find sheet music from SharePoint
- generate templated slide decks automatically