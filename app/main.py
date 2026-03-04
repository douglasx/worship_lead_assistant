from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from typing import Iterable

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pypdf import PdfReader
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

app = FastAPI(title="Worship Sheet to Slides")

BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

_CHORD_TOKEN = re.compile(
    r"^[A-G](?:#|b)?(?:m|maj|min|sus|dim|aug|add)?\d*(?:/[A-G](?:#|b)?)?$",
    re.IGNORECASE,
)
_SECTION_HEADER = re.compile(
    r"^(verse|chorus|bridge|pre-chorus|interlude|tag|ending|outro|intro)\s*\d*[:.]?$",
    re.IGNORECASE,
)
_DROP_LINE = re.compile(
    r"(copyright|all rights reserved|CCLI|\bpage\s*\d+\b)",
    re.IGNORECASE,
)
_MUSIC_SYMBOLS = re.compile(
    r"[\u2669\u266a\u266b\u266c\u266d\u266e\u266f\ud834\udd1e\ud834\udd22\ud834\udd21\ud834\udd2a\ud834\udd2b\ud834\udd5d\ud834\udd57\ud834\udd65\ud834\udd58\ud834\udd65\ud834\udd6e\ud834\udd6f\ud834\udd70\ud834\udd71\ud834\udd72]",
)


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


@app.post("/api/convert")
async def convert(
    files: list[UploadFile] = File(...),
    lines_per_slide: int = Form(6),
    font_size: int = Form(38),
) -> Response:
    if not files:
        raise HTTPException(status_code=400, detail="Please upload at least one PDF file.")

    lines_per_slide = max(2, min(lines_per_slide, 12))
    font_size = max(20, min(font_size, 72))

    output_files: list[tuple[str, bytes]] = []

    for upload in files:
        if not upload.filename or not upload.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail=f"{upload.filename or 'Unknown file'} is not a PDF.")

        file_bytes = await upload.read()
        if not file_bytes:
            raise HTTPException(status_code=400, detail=f"{upload.filename} is empty.")

        try:
            ppt_bytes = build_pptx_from_pdf(file_bytes, lines_per_slide=lines_per_slide, font_size=font_size)
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(
                status_code=422,
                detail=f"Failed to process {upload.filename}: {str(exc)}",
            ) from exc

        output_name = f"{_sanitize_basename(Path(upload.filename).stem)}_lyrics.pptx"
        output_files.append((output_name, ppt_bytes))

    if len(output_files) == 1:
        filename, data = output_files[0]
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )

    archive = io.BytesIO()
    with zipfile.ZipFile(archive, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, data in output_files:
            zf.writestr(filename, data)

    return Response(
        content=archive.getvalue(),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=worship_lyrics_slides.zip"},
    )


def build_pptx_from_pdf(pdf_bytes: bytes, *, lines_per_slide: int, font_size: int) -> bytes:
    raw_text = _extract_text_from_pdf(pdf_bytes)
    cleaned_lines = _clean_lines(raw_text.splitlines())
    slide_chunks = _build_slide_chunks(cleaned_lines, lines_per_slide=lines_per_slide)

    if not slide_chunks:
        raise ValueError("No usable lyrics were found in the uploaded PDF.")

    presentation = Presentation()
    blank_layout = presentation.slide_layouts[6]

    for chunk in slide_chunks:
        slide = presentation.slides.add_slide(blank_layout)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.8), Inches(5.8))
        text_frame = box.text_frame
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        for idx, line in enumerate(chunk):
            paragraph = text_frame.paragraphs[0] if idx == 0 else text_frame.add_paragraph()
            paragraph.text = line
            paragraph.alignment = PP_ALIGN.CENTER

            run = paragraph.runs[0]
            run.font.name = "Calibri"
            run.font.size = Pt(font_size + 4 if _is_section_header(line) else font_size)
            run.font.bold = _is_section_header(line)
            run.font.color.rgb = RGBColor(255, 255, 255)

    output = io.BytesIO()
    presentation.save(output)
    return output.getvalue()


def _extract_text_from_pdf(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = [page.extract_text() or "" for page in reader.pages]
    text = "\n".join(pages).strip()

    if _has_meaningful_text(text):
        return text

    ocr_text = _extract_text_with_ocr(pdf_bytes)
    if _has_meaningful_text(ocr_text):
        return ocr_text

    if text:
        return text

    raise ValueError("Could not extract readable text from this PDF, even with OCR fallback.")


def _extract_text_with_ocr(pdf_bytes: bytes) -> str:
    try:
        import fitz  # PyMuPDF
        import pytesseract
        from PIL import Image
    except Exception as exc:  # noqa: BLE001
        raise ValueError(
            "OCR fallback requires PyMuPDF, Pillow, pytesseract, and Tesseract OCR installed."
        ) from exc

    texts: list[str] = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        page_text = pytesseract.image_to_string(image)
        if page_text:
            texts.append(page_text)

    return "\n".join(texts).strip()


def _has_meaningful_text(text: str) -> bool:
    if not text or len(text.strip()) < 40:
        return False

    alpha_chars = sum(1 for ch in text if ch.isalpha())
    return alpha_chars >= 20


def _clean_lines(lines: Iterable[str]) -> list[str]:
    cleaned: list[str] = []

    for raw in lines:
        line = " ".join(raw.replace("\u00a0", " ").split()).strip()
        if not line:
            if cleaned and cleaned[-1] != "":
                cleaned.append("")
            continue

        if _DROP_LINE.search(line):
            continue

        if _is_probable_chord_line(line):
            continue

        # Remove music symbols and special characters
        line = _MUSIC_SYMBOLS.sub("", line)
        line = _remove_special_chars(line)
        
        # Skip if line becomes empty or too short after cleaning
        line = line.strip()
        if not line or len(line) < 2:
            continue
            
        # Skip lines that are mostly punctuation or numbers
        if _is_mostly_non_alpha(line):
            continue

        cleaned.append(line)

    while cleaned and cleaned[-1] == "":
        cleaned.pop()

    return cleaned


def _build_slide_chunks(lines: list[str], *, lines_per_slide: int) -> list[list[str]]:
    groups: list[list[str]] = []
    current: list[str] = []

    for line in lines:
        if line == "":
            if current:
                groups.append(current)
                current = []
            continue

        if _is_section_header(line) and current:
            groups.append(current)
            current = [line]
            continue

        current.append(line)

    if current:
        groups.append(current)

    chunks: list[list[str]] = []
    for group in groups:
        if len(group) <= lines_per_slide:
            chunks.append(group)
            continue

        start = 0
        while start < len(group):
            end = start + lines_per_slide
            piece = group[start:end]
            chunks.append(piece)
            start = end

    return chunks


def _is_probable_chord_line(line: str) -> bool:
    normalized = line.replace("|", " ").replace("-", " ")
    tokens = [t for t in normalized.split() if t]

    if not tokens:
        return False

    chord_like = sum(1 for token in tokens if _CHORD_TOKEN.match(token))
    return chord_like / len(tokens) >= 0.6 and len(tokens) <= 16


def _is_section_header(line: str) -> bool:
    return bool(_SECTION_HEADER.match(line.strip()))


def _remove_special_chars(text: str) -> str:
    """Remove common special characters that appear in sheet music but not in lyrics."""
    # Remove box drawing characters
    text = re.sub(r"[\u2500\u2502\u250c\u2510\u2514\u2518\u251c\u2524\u252c\u2534\u253c\u2550\u2551\u2554\u2557\u255a\u255d\u2560\u2563\u2566\u2569\u256c]", "", text)
    # Remove bullet points and arrows
    text = re.sub(r"[\u2022\u25cf\u25cb\u25e6\u25a0\u25a1\u25aa\u25ab\u25ba\u25b6\u25c4\u25c0\u2191\u2193\u2192\u2190]", "", text)
    # Remove excessive punctuation patterns (e.g., "----", "....", "||||")
    text = re.sub(r"([.\-_|=]){3,}", "", text)
    # Remove tab and other whitespace characters
    text = re.sub(r"[\t\r\f\v]", " ", text)
    return text


def _is_mostly_non_alpha(text: str) -> bool:
    """Check if a line is mostly non-alphabetic characters (numbers, punctuation, symbols)."""
    if not text:
        return True
    alpha_count = sum(1 for ch in text if ch.isalpha())
    total_count = len(text.replace(" ", ""))
    return total_count > 0 and (alpha_count / total_count) < 0.4


def _sanitize_basename(name: str) -> str:
    safe = re.sub(r"[^a-zA-Z0-9_-]+", "_", name).strip("_")
    return safe or "lyrics"
