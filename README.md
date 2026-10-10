# PPTX Splitter

Turn PowerPoint decks into upload-ready PDFs — without losing embedded videos.

- **Deck without video** → one PDF.
- **Deck with video** → split at the video slides into PDFs plus standalone video files, numbered in reading order (`Content piece 1 - …`, `Content piece 2 - …`).
- **Animated GIFs** that fill a slide (at least 15% of its area) count as videos too, since a PDF would freeze them on their first frame. They're converted to `.mp4` (looped to at least 10 seconds) if [ffmpeg](https://ffmpeg.org) is installed (`brew install ffmpeg`), otherwise the `.gif` itself is output. Small animated icons stay in the PDF.
- Speaker/teacher notes are kept and renumbered to each PDF's own pages.
- Decks with hidden slides are refused, since hidden slides would shift every later page.

## Quick start

```bash
cd pptx-splitter
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python3 app.py
```

Then open http://127.0.0.1:5000, drop in decks (or paste local paths for large files), and download the results individually or as a zip.

## Dependencies

### 1. Python

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.9 or newer | Tested on 3.14. Check with `python3 --version`. Get it from https://www.python.org/downloads/ or `brew install python`. |

### 2. Python libraries (installed by `pip install -r requirements.txt`)

| Library | Tested version | Used for |
|---|---|---|
| [Flask](https://pypi.org/project/Flask/) | 3.1.3 | The local web page (`app.py`). Not needed for the command-line tool. |
| [pypdf](https://pypi.org/project/pypdf/) | 6.16.2 | Cutting the converted PDF into parts and counting pages. **Required for both web and CLI.** |

Flask pulls in these automatically — no need to install them separately:

| Library | Tested version |
|---|---|
| Werkzeug | 3.1.8 |
| Jinja2 | 3.1.6 |
| MarkupSafe | 3.0.3 |
| itsdangerous | 2.2.0 |
| click | 8.5.0 |
| blinker | 1.9.0 |

Everything else the code uses (`zipfile`, `subprocess`, `argparse`, `json`, `re`, `tempfile`, `threading`, …) is part of Python's standard library.

### 3. A slide converter (one of these)

The PPTX → PDF rendering is done by an office app, not a Python library. You need **at least one**:

| Converter | Platform | Notes |
|---|---|---|
| **Microsoft PowerPoint** | macOS | Preferred — used automatically if `/Applications/Microsoft PowerPoint.app` exists. Shapes render exactly as authored. Driven via `osascript` (built into macOS). The first run may ask you to allow Terminal to control PowerPoint — click **OK**. |
| **LibreOffice** | macOS / Windows / Linux | Fallback. Install with `brew install --cask libreoffice` (macOS), `sudo apt install libreoffice` (Linux), or from https://www.libreoffice.org/download/. The `soffice` command must be on your `PATH`. |

If neither is found, conversion stops with: *"No converter available. Install Microsoft PowerPoint or LibreOffice."*

## Putting it online for other people

The `Dockerfile` builds a hosted version anyone can use from a browser, on any computer. In hosted mode:

- Everyone shares one **password** (`APP_PASSWORD`). The server won't start without one.
- Visitors can only **upload** decks. Pasting a path and choosing an output folder are switched off, because on a server both would expose the server's own files.
- Results are kept for **6 hours** (`JOB_TTL_HOURS`) and then deleted. Visitors download them from the results page.
- Decks convert **one at a time** (`MAX_PARALLEL_JOBS`). Later uploads wait their turn.
- Slides are rendered by **LibreOffice**, because servers have no PowerPoint. Most slides look the same, but check one converted deck against the original.

### Deploy on Railway (recommended: pay for what you use, roughly $5/month)

1. Sign up at https://railway.com with the GitHub account that owns this repository.
2. **New Project → Deploy from GitHub repo**, then pick `pptx-splitter`. Railway finds the `Dockerfile` by itself.
3. Open the service's **Variables** tab and add `APP_PASSWORD` = a password of your choice.
4. Under **Settings → Networking**, click **Generate Domain**.
5. Send that URL and the password to whoever needs it. When the browser asks, any username works; only the password matters.

Each `git push` to `main` redeploys automatically.

### Or Render

**New → Web Service**, pick the repository, and choose the Docker runtime. Add `APP_PASSWORD` as an environment variable. Choose a plan with at least **2 GB of memory**: LibreOffice runs out of memory on large decks with 512 MB.

### Settings

| Variable | Default | What it does |
|---|---|---|
| `APP_PASSWORD` | *(required)* | Shared password for the whole site |
| `JOB_TTL_HOURS` | `6` | How long results stay downloadable |
| `MAX_UPLOAD_MB` | `4096` | Largest upload accepted |
| `MAX_PARALLEL_JOBS` | `1` | Conversions running at the same time |

## Command line

The same conversion without the web page:

```bash
python3 split_pptx.py deck.pptx
python3 split_pptx.py *.pptx -o converted/
python3 split_pptx.py ~/Decks --recursive
python3 split_pptx.py ~/Decks --dry-run   # show the plan, write nothing
python3 split_pptx.py                     # prompts; drag files or a folder in
```

## Files

| File | What it is |
|---|---|
| `app.py` | Local web app (Flask) |
| `split_pptx.py` | All conversion logic; also the CLI |
| `templates/`, `static/` | Web page HTML and CSS |
| `requirements.txt` | Python libraries to install |
| `Dockerfile` | Hosted version (LibreOffice + ffmpeg) for putting it online |
| `output/` | Created on first run; converted files land here by default |
