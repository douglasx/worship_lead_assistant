from __future__ import annotations

import base64
import io
import json
import os
import re
from functools import lru_cache
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from dotenv import load_dotenv
load_dotenv(Path(__file__).parents[1] / ".env")

import anthropic
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

app = FastAPI(title="Worship Sheet to Slides")

BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

_claude = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env

# Slide style defaults (based on worship_slide_style_guide.md)
_FONT_EN = "Open Sans"
_FONT_ZH = "Noto Sans SC"

_COLOR_BG = RGBColor(0x1A, 0x1A, 0x1A)
_COLOR_LYRIC = RGBColor(0xF5, 0xF5, 0xF5)
_COLOR_TITLE_MALE = RGBColor(0x4D, 0xA3, 0xFF)
_COLOR_TITLE_FEMALE = RGBColor(0xFF, 0x4D, 0x4D)
_COLOR_TITLE_EVERYONE = RGBColor(0x4A, 0xC7, 0x6D)
_COLOR_FOOTER = RGBColor(0xFF, 0x4D, 0x4D)

_VOCAL_PART_COLORS = {
    "male": _COLOR_TITLE_MALE,
    "female": _COLOR_TITLE_FEMALE,
    "everyone": _COLOR_TITLE_EVERYONE,
}

_PDF_EXTRACTION_PROMPT = """\
You are a worship music expert. Analyze this worship song sheet image and extract the lyrics.

Rules:
- Remove all guitar/piano chords (e.g. G, Am, F/C, Dsus4, Capo 3, etc.)
- Remove copyright notices, CCLI numbers, page numbers, tempo markings, and any metadata
- Remove music notation symbols and repeated section cues like "(repeat chorus)"
- Keep only the sung lyric lines
- Preserve inline repeat markers when present (e.g. x2, X2, (x2))
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
6. Preserve inline repeat markers when present (e.g. x2, X2, (x2))

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
# Song recommender – data + helpers
# ---------------------------------------------------------------------------

_SONGS_DB_PATH = Path(__file__).parents[1] / "songs_database.xlsx"


@lru_cache(maxsize=1)
def _load_songs() -> list[dict]:
    """Load songs from the Excel database. Cached after first call."""
    try:
        import pandas as pd
    except ImportError as exc:
        raise RuntimeError("pandas is required for song recommendations.") from exc

    df = pd.read_excel(_SONGS_DB_PATH, sheet_name="Sheet1")
    df = df.where(pd.notna(df), None)  # replace NaN with None

    songs = []
    for _, row in df.iterrows():
        songs.append({
            "title": str(row.get("Songs") or "").strip(),
            "link": str(row.get("Link") or "").strip(),
            "theme": str(row.get("Theme") or "").strip(),
            "theme_translated": str(row.get("Theme_translated") or "").strip(),
            "lyrics_snippet": str(row.get("lyrics") or "")[:300].strip(),
        })
    return [s for s in songs if s["title"]]


def _build_song_catalog(songs: list[dict]) -> str:
    lines = []
    for i, s in enumerate(songs, 1):
        theme_part = s["theme"] or s["theme_translated"]
        lines.append(f"{i}. {s['title']} | {theme_part}")
    return "\n".join(lines)


class RecommendRequest(BaseModel):
    message: str
    count: int = 7


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
    vocal_part: str = Form("everyone"),
    footer_note: str = Form(""),
) -> Response:
    if not files:
        raise HTTPException(status_code=400, detail="Please upload at least one PDF file.")

    lines_per_slide = max(2, min(lines_per_slide, 12))
    font_size = max(20, min(font_size, 72))
    vocal_part = _normalize_vocal_part(vocal_part)
    footer_note = footer_note.strip()

    songs: list[dict] = []
    first_stem: str = "lyrics"

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
        except Exception as exc:
            raise HTTPException(
                status_code=422,
                detail=f"Failed to process {upload.filename}: {exc}",
            ) from exc

        songs.append(song_data)
        if len(songs) == 1:
            first_stem = _sanitize_basename(Path(upload.filename).stem)

    ppt_bytes = _build_combined_pptx(
        songs,
        lines_per_slide=lines_per_slide,
        font_size=font_size,
        vocal_part=vocal_part,
        footer_note=footer_note,
    )

    if len(songs) == 1:
        filename = f"{first_stem}_lyrics.pptx"
    else:
        filename = "worship_songs_lyrics.pptx"

    return Response(
        content=ppt_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": _content_disposition(filename)},
    )


@app.post("/api/convert-youtube")
async def convert_youtube(
    urls: str = Form(...),
    source: str = Form("transcript"),  # "transcript", "description", or "auto"
    lines_per_slide: int = Form(4),
    font_size: int = Form(56),
    vocal_part: str = Form("everyone"),
    footer_note: str = Form(""),
) -> Response:
    lines_per_slide = max(2, min(lines_per_slide, 12))
    font_size = max(20, min(font_size, 72))
    vocal_part = _normalize_vocal_part(vocal_part)
    footer_note = footer_note.strip()

    url_list = [u.strip() for u in urls.splitlines() if u.strip()]
    if not url_list:
        raise HTTPException(status_code=400, detail="Please enter at least one YouTube URL.")

    songs: list[dict] = []
    for url in url_list:
        try:
            if source == "description":
                raw_text = _get_youtube_description(url)
            elif source == "auto":
                raw_text = _get_youtube_text_auto(url)
            else:
                # For transcript mode, transparently fall back to description.
                try:
                    raw_text = _get_youtube_transcript(url)
                except Exception as transcript_exc:
                    try:
                        raw_text = _get_youtube_description(url)
                    except Exception as description_exc:
                        raise ValueError(
                            "Transcript extraction failed "
                            f"({transcript_exc}); description fallback failed ({description_exc})"
                        ) from description_exc
            songs.append(_structure_transcript_with_claude(raw_text))
        except Exception as exc:
            raise HTTPException(status_code=422, detail=f"{url}: {exc}") from exc

    ppt_bytes = _build_combined_pptx(
        songs,
        lines_per_slide=lines_per_slide,
        font_size=font_size,
        vocal_part=vocal_part,
        footer_note=footer_note,
    )

    if len(songs) == 1:
        title = songs[0].get("title") or "worship_song"
        filename = f"{_sanitize_basename(title)}_lyrics.pptx"
    else:
        filename = "worship_songs_lyrics.pptx"

    return Response(
        content=ppt_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": _content_disposition(filename)},
    )


@app.post("/api/recommend-songs")
async def recommend_songs(req: RecommendRequest) -> dict:
    """Return a ranked list of worship songs relevant to the given message or Bible verses."""
    if not req.message.strip():
        raise HTTPException(status_code=400, detail="Please provide a message or Bible verses.")

    count = max(1, min(req.count, 15))

    try:
        songs = _load_songs()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Could not load song database: {exc}") from exc

    catalog = _build_song_catalog(songs)
    song_index = {s["title"]: s for s in songs}

    prompt = f"""You are a worship leader assistant. Given a sermon message or Bible passage, \
select the {count} most spiritually relevant worship songs from the catalog below.

CATALOG (index. Title | Theme):
{catalog}

SERMON / BIBLE PASSAGE:
{req.message}

Instructions:
- Choose songs whose theme and spirit align with the message or passage.
- Rank them from most to least relevant.
- For each song provide a brief reason (1-2 sentences) explaining why it fits.
- Respond ONLY with valid JSON — no markdown fences, no extra text:
{{
  "recommendations": [
    {{"title": "exact song title from catalog", "reason": "..."}}
  ]
}}"""

    response = _claude.messages.create(
        model="claude-opus-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    )

    try:
        result = _parse_claude_json(response.content[0].text)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to parse Claude response: {exc}") from exc

    # Enrich each recommendation with link and theme from the database
    enriched = []
    for rec in result.get("recommendations", []):
        title = rec.get("title", "")
        meta = song_index.get(title, {})
        enriched.append({
            "title": title,
            "reason": rec.get("reason", ""),
            "theme": meta.get("theme", ""),
            "theme_translated": meta.get("theme_translated", ""),
            "link": meta.get("link", ""),
        })

    return {"recommendations": enriched}


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

    video_id = _extract_youtube_video_id(url)
    if not video_id:
        raise ValueError(f"Could not find a valid video ID in: {url}")

    api = YouTubeTranscriptApi()
    transcript_errors: list[str] = []

    language_preference = [
        ["zh-TW", "zh-Hant", "zh-Hans", "zh"],
        ["en", "en-US", "en-GB"],
    ]

    # Fast path: direct fetch with language preferences.
    for languages in language_preference:
        try:
            transcript = api.fetch(video_id, languages=languages)
            text = _transcript_entries_to_text(transcript)
            if text:
                return text
        except Exception as exc:
            transcript_errors.append(f"fetch({languages}): {exc}")
            continue

    # Fallback: use transcript list API and try preferred languages, then any generated/manual transcript.
    try:
        transcript_list = api.list(video_id)
    except Exception as exc:
        transcript_errors.append(f"list(): {exc}")
        transcript_list = None

    if transcript_list is not None:
        for languages in language_preference:
            try:
                transcript = transcript_list.find_transcript(languages).fetch()
                text = _transcript_entries_to_text(transcript)
                if text:
                    return text
            except Exception as exc:
                transcript_errors.append(f"find_transcript({languages}): {exc}")

        for languages in language_preference:
            try:
                transcript = transcript_list.find_generated_transcript(languages).fetch()
                text = _transcript_entries_to_text(transcript)
                if text:
                    return text
            except Exception as exc:
                transcript_errors.append(f"find_generated_transcript({languages}): {exc}")

        try:
            for candidate in transcript_list:
                transcript = candidate.fetch()
                text = _transcript_entries_to_text(transcript)
                if text:
                    return text
        except Exception as exc:
            transcript_errors.append(f"iterate_transcripts(): {exc}")

    reason = transcript_errors[-1] if transcript_errors else "no transcript data returned"
    raise ValueError(f"No transcript is available for this YouTube video ({video_id}): {reason}")


def _get_youtube_description(url: str) -> str:
    """Extract the video description from a YouTube URL using yt-dlp."""
    try:
        import yt_dlp
    except ImportError as exc:
        raise ValueError("yt-dlp package is required for description extraction.") from exc

    video_id = _extract_youtube_video_id(url)
    if not video_id:
        raise ValueError(f"Could not find a valid video ID in: {url}")

    normalized_url = f"https://www.youtube.com/watch?v={video_id}"
    errors: list[str] = []

    for ydl_opts in _yt_dlp_option_candidates():
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(normalized_url, download=False)
            description = str(info.get("description") or "").strip()
            if description:
                return description
            errors.append("empty description")
        except Exception as exc:
            errors.append(str(exc))

    details = "; ".join(errors[-3:]) if errors else "unknown extraction error"
    raise ValueError(f"Could not extract YouTube description for {video_id}: {details}")


def _get_youtube_text_auto(url: str) -> str:
    """Try transcript first, then description, returning the first usable lyric source."""
    try:
        return _get_youtube_transcript(url)
    except Exception as transcript_exc:
        try:
            return _get_youtube_description(url)
        except Exception as description_exc:
            raise ValueError(
                "Auto mode failed. "
                f"Transcript error: {transcript_exc}. "
                f"Description error: {description_exc}."
            ) from description_exc


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


def _extract_youtube_video_id(url: str) -> str | None:
    """Extract a stable 11-char YouTube video id from common URL forms."""
    raw = (url or "").strip()
    if re.fullmatch(r"[A-Za-z0-9_-]{11}", raw):
        return raw

    parsed = urlparse(raw)
    host = (parsed.netloc or "").lower()
    path = parsed.path or ""

    if "youtu.be" in host:
        candidate = path.lstrip("/").split("/")[0]
        if re.fullmatch(r"[A-Za-z0-9_-]{11}", candidate or ""):
            return candidate

    if "youtube.com" in host:
        query_id = parse_qs(parsed.query).get("v", [None])[0]
        if re.fullmatch(r"[A-Za-z0-9_-]{11}", query_id or ""):
            return query_id

        parts = [p for p in path.split("/") if p]
        for idx, part in enumerate(parts):
            if part in {"shorts", "live", "embed", "v"} and idx + 1 < len(parts):
                candidate = parts[idx + 1]
                if re.fullmatch(r"[A-Za-z0-9_-]{11}", candidate or ""):
                    return candidate

    match = re.search(r"(?:v=|youtu\.be/|embed/|shorts/|live/)([A-Za-z0-9_-]{11})", raw)
    return match.group(1) if match else None


def _transcript_entries_to_text(entries) -> str:
    lines: list[str] = []
    for entry in entries:
        text = ""
        if isinstance(entry, dict):
            text = str(entry.get("text") or "")
        else:
            text = str(getattr(entry, "text", "") or "")
        text = text.strip()
        if text:
            lines.append(text)
    return " ".join(lines).strip()


def _yt_dlp_option_candidates() -> list[dict]:
    """Build yt-dlp option candidates with retries and opt-in cookie strategies."""
    base = {
        "quiet": True,
        "no_warnings": True,
        "skip_download": True,
        "extractor_retries": 3,
        "retries": 3,
        "socket_timeout": 20,
        "http_headers": {
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/128.0.0.0 Safari/537.36"
            )
        },
        "extractor_args": {
            "youtube": {
                "player_client": ["android", "web", "tv_embedded"],
            }
        },
    }

    candidates: list[dict] = [base]

    cookie_file = os.getenv("YTDLP_COOKIE_FILE", "").strip()
    if cookie_file and Path(cookie_file).exists():
        opts = dict(base)
        opts["cookiefile"] = cookie_file
        candidates.append(opts)

    raw_browser = os.getenv("YTDLP_COOKIES_FROM_BROWSER", "").strip()
    if raw_browser and raw_browser.lower() not in {"0", "false", "off", "none"}:
        for browser in [b.strip() for b in raw_browser.split(",") if b.strip()]:
            opts = dict(base)
            opts["cookiesfrombrowser"] = (browser,)
            candidates.append(opts)

    return candidates


# ---------------------------------------------------------------------------
# Slide builder
# ---------------------------------------------------------------------------

def _build_combined_pptx(
    songs: list[dict],
    *,
    lines_per_slide: int,
    font_size: int,
    vocal_part: str,
    footer_note: str,
) -> bytes:
    """Build a style-guide-compliant PPTX from one or more song dicts."""
    if not songs:
        raise ValueError("No songs to build presentation from.")

    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank_layout = presentation.slide_layouts[6]

    SLIDE_W = Inches(13.333)
    SLIDE_H = Inches(7.5)
    MARGIN_H = Inches(0.5)
    MARGIN_V = Inches(0.5)

    title_color = _VOCAL_PART_COLORS[vocal_part]

    for song_data in songs:
        sections = song_data.get("sections", [])
        language = song_data.get("language", "en")
        song_title = (song_data.get("title") or "").strip()

        if not sections:
            continue

        song_slides: list[dict] = []

        for section in sections:
            section_type = section.get("type", "")
            section_num = section.get("number")
            lines = [ln for ln in section.get("lines", []) if ln.strip()]

            if not lines:
                continue

            header = f"{section_type} {section_num}".strip() if section_num else str(section_type).strip()
            chunks = [lines[i: i + lines_per_slide] for i in range(0, len(lines), lines_per_slide)]

            for chunk in chunks:
                title_left = _build_title_left(song_title, header)
                song_slides.append({
                    "title_left": title_left,
                    "lyrics": chunk,
                })

        total = len(song_slides)
        if total == 0:
            continue

        for slide_idx, slide_data in enumerate(song_slides, start=1):
            slide = presentation.slides.add_slide(blank_layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = _COLOR_BG

            _add_title_row(
                slide,
                left_text=slide_data["title_left"],
                counter_text=f"{slide_idx}/{total}",
                title_color=title_color,
                language=language,
                left=MARGIN_H,
                top=Inches(0.35),
                width=SLIDE_W - 2 * MARGIN_H,
                height=Inches(0.7),
            )
            _add_lyric_block(
                slide,
                lines=slide_data["lyrics"],
                language=language,
                font_size=font_size,
                left=MARGIN_H,
                top=Inches(1.2),
                width=SLIDE_W - 2 * MARGIN_H,
                height=SLIDE_H - Inches(2.1),
            )

            if footer_note and slide_idx == 1:
                _add_footer_note(
                    slide,
                    footer_note,
                    language=language,
                    left=MARGIN_H,
                    top=SLIDE_H - Inches(0.52),
                    width=SLIDE_W - 2 * MARGIN_H,
                    height=Inches(0.3),
                )

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


def _normalize_vocal_part(vocal_part: str) -> str:
    part = (vocal_part or "").strip().lower()
    if part not in _VOCAL_PART_COLORS:
        return "everyone"
    return part


def _contains_cjk(text: str) -> bool:
    return bool(re.search(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]", text or ""))


def _font_for_text(text: str, language: str) -> str:
    if _contains_cjk(text) or "zh" in (language or ""):
        return _FONT_ZH
    return _FONT_EN


def _build_title_left(song_title: str, section_header: str) -> str:
    if song_title and section_header:
        return f"{song_title} · {section_header}"
    return song_title or section_header or "Worship"


def _add_title_row(
    slide,
    *,
    left_text: str,
    counter_text: str,
    title_color: RGBColor,
    language: str,
    left,
    top,
    width,
    height,
) -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.TOP

    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    para.space_after = Pt(0)

    title_run = para.add_run()
    title_run.text = left_text
    title_run.font.name = _font_for_text(left_text, language)
    title_run.font.size = Pt(32)
    title_run.font.bold = False
    title_run.font.color.rgb = title_color

    counter_run = para.add_run()
    counter_run.text = f"  {counter_text}"
    counter_run.font.name = _FONT_EN
    counter_run.font.size = Pt(20)
    counter_run.font.bold = False
    counter_run.font.color.rgb = title_color


def _add_lyric_block(slide, *, lines: list[str], language: str, font_size: int, left, top, width, height) -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    for idx, line in enumerate(lines):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        para.text = line
        para.alignment = PP_ALIGN.CENTER
        para.line_spacing = 1.25
        para.space_after = Pt(6)

        run = para.runs[0]
        run.font.name = _font_for_text(line, language)
        run.font.size = Pt(font_size)
        run.font.bold = True
        run.font.color.rgb = _COLOR_LYRIC


def _add_footer_note(slide, note: str, *, language: str, left, top, width, height) -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.BOTTOM

    para = tf.paragraphs[0]
    para.text = note
    para.alignment = PP_ALIGN.LEFT
    para.space_after = Pt(0)

    run = para.runs[0]
    run.font.name = _font_for_text(note, language)
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = _COLOR_FOOTER
