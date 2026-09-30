# PPTX Splitter

Turn PowerPoint decks into upload-ready PDFs — without losing embedded videos.

- **Deck without video** → one PDF.
- **Deck with video** → split at the video slides into PDFs plus standalone video files, numbered in reading order (`Content piece 1 - …`, `Content piece 2 - …`).
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
| `output/` | Created on first run; converted files land here by default |
