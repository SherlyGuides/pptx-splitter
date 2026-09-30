# PPTX Splitter

Turn PowerPoint decks into upload-ready PDFs — without losing embedded videos.

- **Deck without video** → one PDF.
- **Deck with video** → split at the video slides into PDFs plus standalone video files, numbered in reading order (`Content piece 1 - …`, `Content piece 2 - …`).
- Speaker/teacher notes are kept and renumbered to each PDF's own pages.
- Decks with hidden slides are refused, since hidden slides would shift every later page.

Conversion uses Microsoft PowerPoint when installed, otherwise LibreOffice.

## Web app

```bash
pip install -r requirements.txt
python3 app.py
```

Then open http://127.0.0.1:5000, drop in decks (or paste local paths for large files), and download the results individually or as a zip.

## Command line

```bash
python3 split_pptx.py deck.pptx
python3 split_pptx.py *.pptx -o converted/
python3 split_pptx.py ~/Decks --recursive
python3 split_pptx.py ~/Decks --dry-run   # show the plan, write nothing
```

## Requirements

- Python 3.9+
- Microsoft PowerPoint (macOS) or [LibreOffice](https://www.libreoffice.org/)
