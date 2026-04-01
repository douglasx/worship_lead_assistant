from __future__ import annotations

import base64
import io
import json
import re
import zipfile
from pathlib import Path
from urllib.parse import quote

from dotenv import load_dotenv
load_dotenv(Path(__file__).parents[1] / ".env")

import anthropic
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

app = FastAPI(title="Worship Sheet to Slides")

BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

_claude = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env

# Fonts: YaHei supports both CJK and Latin; Calibri for English-only
_FONT_EN = "Calibri"
_FONT_ZH = "Microsoft YaHei"

_PDF_EXTRACTION_PROMPT = """\
You are a worship music expert. Analyze this worship song sheet image and extract the lyrics.

Rules:
- Remove all guitar/piano chords (e.g. G, Am, F/C, Dsus4, Capo 3, etc.)
- Remove copyright notices, CCLI numbers, page numbers, tempo markings, and any metadata
- Remove music notation symbols and repeated section cues like "(repeat chorus)"
- Keep only the sung lyric lines
- Identify and label each section: Verse 1, Verse 2, Chorus, Bridge, Pre-Chorus, Tag, Outro, Intro, etc.
- The lyrics may be in English, Chinese (Traditional or Simplified), or both

Return ONLY valid JSON, no other text:
{
  "title": "Song Title or null",
  "language": "en",
  "sections": [
    {"type": "Verse", "number": 1, "lines": ["line 1", "line 2"]},
    {"type": "Chorus", "number": null, "lines": ["line 1", "line 2"]}
  ]
}

For language: use "en" for English, "zh" for Chinese, "en-zh" for mixed.\
"""

_TRANSCRIPT_STRUCTURING_PROMPT = """\
You are a worship music expert. Below is a raw transcript from a worship song video.

Your task:
1. Extract only the sung lyrics (ignore spoken parts, timestamps, [Music], [Applause], etc.)
2. Identify the song structure: Verse 1, Verse 2, Chorus, Bridge, Pre-Chorus, Tag, Outro, etc.
3. Each unique section should appear once with its distinct lines (deduplicate repeated choruses)
4. Fix obvious transcript errors (split/merged words, broken punctuation)
5. The lyrics may be in English, Chinese, or both

Return ONLY valid JSON, no other text:
{
  "title": "Song Title or null",
  "language": "en",
  "sections": [
    {"type": "Verse", "number": 1, "lines": ["line 1", "line 2"]},
    {"type": "Chorus", "number": null, "lines": ["line 1", "line 2"]}
  ]
}

For language: use "en" for English, "zh" for Chinese, "en-zh" for mixed.

Transcript:
"""


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


@app.post("/api/convert")
async def convert(
    files: list[UploadFile] = File(...),
    lines_per_slide: int = Form(4),
    font_size: int = Form(56),
) -> Response:
    if not files:
        raise HTTPException(status_code=400, detail="Please upload at least one PDF file.")

    lines_per_slide = max(2, min(lines_per_slide, 12))
    font_size = max(20, min(font_size, 72))

    output_files: list[tuple[str, bytes]] = []

    for upload in files:
        if not upload.filename or not upload.filename.lower().endswith(".pdf"):
            raise HTTPException(
                status_code=400,
                detail=f"{upload.filename or 'Unknown file'} is not a PDF.",
            )

        file_bytes = await upload.read()
        if not file_bytes:
            raise HTTPException(status_code=400, detail=f"{upload.filename} is empty.")

        try:
            song_data = _extract_lyrics_from_pdf(file_bytes)
            ppt_bytes = _build_pptx(song_data, lines_per_slide=lines_per_slide, font_size=font_size)
        except Exception as exc:
            raise HTTPException(
                status_code=422,
                detail=f"Failed to process {upload.filename}: {exc}",
            ) from exc

        stem = _sanitize_basename(Path(upload.filename).stem)
        output_files.append((f"{stem}_lyrics.pptx", ppt_bytes))

    if len(output_files) == 1:
        filename, data = output_files[0]
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": _content_disposition(filename)},
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


@app.post("/api/convert-youtube")
async def convert_youtube(
    url: str = Form(...),
    source: str = Form("transcript"),  # "transcript" or "description"
    lines_per_slide: int = Form(4),
    font_size: int = Form(56),
) -> Response:
    lines_per_slide = max(2, min(lines_per_slide, 12))
    font_size = max(20, min(font_size, 72))

    try:
        if source == "description":
            raw_text = _get_youtube_description(url)
        else:
            raw_text = _get_youtube_transcript(url)
        song_data = _structure_transcript_with_claude(raw_text)
        ppt_bytes = _build_pptx(song_data, lines_per_slide=lines_per_slide, font_size=font_size)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc

    title = song_data.get("title") or "worship_song"
    filename = f"{_sanitize_basename(title)}_lyrics.pptx"

    return Response(
        content=ppt_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": _content_disposition(filename)},
    )


# ---------------------------------------------------------------------------
# Extraction helpers
# ---------------------------------------------------------------------------

def _extract_lyrics_from_pdf(pdf_bytes: bytes) -> dict:
    """Render each PDF page as an image and send to Claude Vision for extraction."""
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:
        raise ValueError("PyMuPDF is required for PDF processing.") from exc

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    content: list[dict] = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
        img_b64 = base64.standard_b64encode(pix.tobytes("png")).decode()
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": img_b64},
        })

    content.append({"type": "text", "text": _PDF_EXTRACTION_PROMPT})

    response = _claude.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": content}],
    )

    return _parse_claude_json(response.content[0].text)


def _get_youtube_transcript(url: str) -> str:
    """Pull transcript text from a YouTube video URL."""
    try:
        from youtube_transcript_api import YouTubeTranscriptApi
    except ImportError as exc:
        raise ValueError("youtube-transcript-api package is required for YouTube support.") from exc

    match = re.search(r"(?:v=|youtu\.be/)([a-zA-Z0-9_-]{11})", url)
    if not match:
        raise ValueError(f"Could not find a valid video ID in: {url}")
    video_id = match.group(1)

    # Prefer manual captions; fall back to auto-generated in priority order
    for languages in [["zh-TW", "zh-Hant", "zh-Hans", "zh"], ["en"], None]:
        try:
            entries = (
                YouTubeTranscriptApi.get_transcript(video_id, languages=languages)
                if languages
                else YouTubeTranscriptApi.get_transcript(video_id)
            )
            return " ".join(e["text"] for e in entries)
        except Exception:
            continue

    raise ValueError(f"No transcript is available for this YouTube video: {url}")


def _get_youtube_description(url: str) -> str:
    """Extract the video description from a YouTube URL using yt-dlp."""
    try:
        import yt_dlp
    except ImportError as exc:
        raise ValueError("yt-dlp package is required for description extraction.") from exc

    ydl_opts = {"quiet": True, "no_warnings": True, "skip_download": True}
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(url, download=False)

    description = (info.get("description") or "").strip()
    if not description:
        raise ValueError("This video has no description.")
    return description


def _structure_transcript_with_claude(transcript_text: str) -> dict:
    """Ask Claude to clean and structure raw transcript text into labeled song sections."""
    response = _claude.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": _TRANSCRIPT_STRUCTURING_PROMPT + transcript_text,
        }],
    )
    return _parse_claude_json(response.content[0].text)


def _parse_claude_json(text: str) -> dict:
    """Extract JSON from Claude's response, stripping any markdown code fences."""
    text = text.strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if fence:
        text = fence.group(1).strip()
    return json.loads(text)


# ---------------------------------------------------------------------------
# Slide builder
# ---------------------------------------------------------------------------

def _build_pptx(song_data: dict, *, lines_per_slide: int, font_size: int) -> bytes:
    sections = song_data.get("sections", [])
    language = song_data.get("language", "en")
    # YaHei supports both CJK and Latin; use it whenever Chinese is present
    font_name = _FONT_ZH if "zh" in language else _FONT_EN

    if not sections:
        raise ValueError("No lyrics sections were found in the song data.")

    presentation = Presentation()
    # 16:9 widescreen (matches modern projectors and screens)
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank_layout = presentation.slide_layouts[6]

    SLIDE_W = Inches(13.333)
    SLIDE_H = Inches(7.5)
    MARGIN_H = Inches(0.5)
    MARGIN_V = Inches(0.5)

    for section in sections:
        section_type = section.get("type", "")
        section_num = section.get("number")
        lines = [ln for ln in section.get("lines", []) if ln.strip()]

        if not lines:
            continue

        header = f"{section_type} {section_num}" if section_num else section_type
        chunks = [lines[i: i + lines_per_slide] for i in range(0, len(lines), lines_per_slide)]

        for chunk_idx, chunk in enumerate(chunks):
            slide = presentation.slides.add_slide(blank_layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

            box = slide.shapes.add_textbox(
                MARGIN_H, MARGIN_V,
                SLIDE_W - 2 * MARGIN_H, SLIDE_H - 2 * MARGIN_V,
            )
            tf = box.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE

            # Show section header only on the first chunk of each section
            display_lines = ([header] if chunk_idx == 0 and header else []) + chunk

            for idx, line in enumerate(display_lines):
                is_header = idx == 0 and chunk_idx == 0 and bool(header)
                para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                para.text = line
                para.alignment = PP_ALIGN.CENTER

                run = para.runs[0]
                run.font.name = font_name
                run.font.size = Pt(font_size + 4 if is_header else font_size)
                run.font.bold = is_header
                run.font.color.rgb = RGBColor(255, 255, 255)

    output = io.BytesIO()
    presentation.save(output)
    return output.getvalue()


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _sanitize_basename(name: str) -> str:
    # Allow ASCII word chars, hyphens, and CJK characters
    safe = re.sub(r"[^\w\-\u4e00-\u9fff]+", "_", name).strip("_")
    return safe or "lyrics"


def _content_disposition(filename: str) -> str:
    """Return a Content-Disposition header value that safely handles non-ASCII filenames."""
    try:
        filename.encode("latin-1")
        return f'attachment; filename="{filename}"'
    except UnicodeEncodeError:
        encoded = quote(filename, safe="")
        return f"attachment; filename*=UTF-8''{encoded}"
