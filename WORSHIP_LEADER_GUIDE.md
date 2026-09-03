# Worship Lead Assistant — User Guide

**What it does:** Helps worship leaders prepare services in two ways: (1) automatically converts worship song sheets (PDF) or YouTube videos into ready-to-project PowerPoint slides using your church style guide (dark charcoal background, soft white lyrics, title/counter color by vocal part, 16:9 widescreen layout); and (2) recommends relevant songs from your church's song library based on your sermon or message.

---

## Getting Started

Open the tool in your browser. You will see three tabs at the top:

- **PDF Upload** — for song sheets saved as PDF files
- **YouTube URL** — for songs found on YouTube
- **Song Recommender** — for finding songs that match your sermon or service theme

The interface is available in **English** and **Traditional Chinese (繁中)**. Click the language button in the top-right corner to switch.

---

## Option A: PDF Upload

Use this when you have chord charts or lyric sheets saved as PDFs (e.g. exported from Planning Center, SongSelect, or scanned sheet music).

1. **Add your PDFs**
   - Drag and drop one or more PDF files onto the upload area, **or**
   - Click **Choose PDF files** and select them from your computer.
   - You can add multiple songs at once — they will all be combined into one PowerPoint file.

2. **Adjust the settings** *(optional)*
   | Setting | Default | What it does |
   |---|---|---|
   | Lines per slide | 4 | How many lyric lines appear on each slide |
   | Font size (pt) | 56 | Text size — increase for larger screens |
   | Title color / Vocal part | Everyone (Green) | Sets title + counter color based on karaoke convention (Male=Blue, Female=Red, Everyone=Green) |
   | Footer note (optional) | empty | Adds a small red instruction note at the bottom-left of the first slide of each song |

3. **Click "Convert to PowerPoint"**
   - The tool sends each PDF to Claude AI, which reads the sheet, removes chords and metadata, and organises the lyrics by section (Verse 1, Chorus, Bridge, etc.).
   - A `.pptx` file will download automatically when it is ready.
   - Processing time is typically **10–30 seconds per song**.

---

## Option B: YouTube URL

Use this when lyrics are available through a YouTube video (e.g. official lyric videos, live worship recordings with captions).

1. **Paste YouTube URLs**
   - Type or paste one URL per line in the text box.
   - You can add multiple songs — they will all be combined into one PowerPoint file.

2. **Choose the lyrics source**

   | Source | When to use |
   |---|---|
   | **Transcript (captions)** | The video has auto-generated or manual subtitles. Try this first. |
   | **Description (lyrics in description)** | The full lyrics are pasted in the video's description field. Use this if captions are inaccurate or missing. |

3. **Adjust the settings** *(same as PDF tab)*

4. **Click "Convert to PowerPoint"**
   - The tool fetches the lyrics from YouTube and sends them to Claude AI for structuring.
   - A `.pptx` file will download automatically.
   - Processing time is typically **15–45 seconds per song**.

---

## Option C: Song Recommender

Use this when you want to find songs from your church's song library that fit the theme or message of an upcoming service.

1. **Paste your sermon or service content**
   - Type or paste sermon notes, a Bible passage, or a description of the service theme into the text area.
   - Both English and Chinese content are supported.

2. **Choose how many songs to suggest**
   - Set the **Number of songs** field (1–15, default is 7).

3. **Click "Find Songs"**
   - Claude AI reads your message and your church's entire song catalog, then ranks the most spiritually relevant songs.
   - Results typically appear within a few seconds.

4. **Review the recommendations**
   - Each result card shows:
     - A **rank badge** (#1, #2, …)
     - The **song title** — click it to open the song on YouTube (if a link is available)
     - The song's **theme tags** (original language · translated)
     - A **1–2 sentence explanation** of why this song fits your message
   - You can use these results to plan your set list, then convert the chosen songs to slides using the PDF Upload or YouTube URL tabs.

---

## The Output PowerPoint

- **Slide size:** 16:9 widescreen (matches most projectors and screens)
- **Background:** Soft dark charcoal (`#1A1A1A`)
- **Lyrics:** Soft white (`#F5F5F5`), bold, centered, large type
- **Title row:** Song/section title at the top with slide counter (`N/total`) immediately after it
- **Title + counter color:** Blue (male), Red (female), Green (everyone)
- **Footer instruction:** Optional red bold note at bottom-left (first slide of each song)
- **Language support:** English and Chinese lyrics are both handled with projector-friendly font mapping (`Open Sans` for English, `Noto Sans SC` for Chinese/CJK).

You can open the `.pptx` in Microsoft PowerPoint or Google Slides and edit it further before your service.

---

## Tips

- **Plan your set list first:** Use the Song Recommender to pick songs, then convert them to slides with PDF Upload or YouTube URL — all in one session.
- **Multiple songs in one deck:** Add all the songs for your service in one go. The output will be a single PowerPoint with all songs in order, ready to present.
- **PDF quality matters:** Clearer, text-based PDFs (not blurry scans) give the best results.
- **No captions on YouTube?** Switch to the **Description** source if the song's lyrics are pasted in the video description.
- **Chinese songs:** The tool supports Traditional Chinese, Simplified Chinese, and bilingual (English + Chinese) sheets.
- **Too many lines per slide?** Reduce "Lines per slide" to 2–3 for songs with long lines or small screen situations.
- **Song Recommender gives no results?** Try rephrasing with more specific theological keywords (e.g. "forgiveness", "Holy Spirit", "resurrection") or paste a fuller passage of Scripture.

---

## Troubleshooting

| Problem | What to try |
|---|---|
| "No transcript available" error | Switch from **Transcript** to **Description**, or use a PDF instead |
| Chords or metadata appear in slides | The AI usually removes these automatically; if not, edit the slides manually in PowerPoint |
| Download did not start | Check that your browser is not blocking pop-ups or downloads |
| Conversion takes a long time | Large PDFs (many pages) or multiple songs take longer — please wait up to 2 minutes |
| Song Recommender returns no results | Try a more descriptive message with specific themes or Scripture references |
| Song Recommender results seem off | Paste more of the sermon content so Claude has fuller context for matching |

---

*Powered by Claude AI (Anthropic) · Built for New Hope Christian Church*
