#!/usr/bin/env python3
"""Split a PowerPoint deck containing videos into PDFs and standalone video files.

A deck that embeds video cannot be published as a single PDF without losing the
video, and cannot be published as a PPTX without dragging the whole deck's
weight along with it. This splits it at the video slides instead:

    slides 1-6 -> PDF,  slide 7 video -> .mp4,  slide 8 -> PDF,  ...

Each video slide is dropped from the PDFs (its video becomes its own piece) and
its title is reused as the video's name, so nothing is silently lost. The pieces
come out numbered in reading order, ready to upload as content pieces.

Speaker notes travel with the slides: every PDF part gets a sidecar JSON with
its notes renumbered to that part's own page numbers, matching the format the
player already reads.

Conversion uses Microsoft PowerPoint when it is installed (its own renderer, so
shapes look exactly as authored) and falls back to LibreOffice.

Decks without video are simply converted to one PDF, so a mixed batch can be
handed over in one go.

Usage:
    python3 split_pptx.py                          # prompts; drag files or a folder in
    python3 split_pptx.py deck.pptx
    python3 split_pptx.py *.pptx -o converted/
    python3 split_pptx.py ~/Decks --recursive
    python3 split_pptx.py ~/Decks --dry-run        # show the plan, write nothing
"""

from __future__ import annotations

import argparse
import glob
import html
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field

VIDEO_EXTENSIONS = {".mp4", ".m4v", ".mov", ".avi", ".wmv", ".mpeg", ".mpg", ".webm"}

# Mirrors the uploader's own content-name rule, so a generated name is never
# rejected at upload time: 30 chars, no characters illegal in a filename, and
# no trailing period or space.
MAX_NAME = 30
UNSAFE_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_name(value: str, fallback: str, limit: int = MAX_NAME) -> str:
    cleaned = UNSAFE_NAME.sub("", value or "").replace("\n", " ").replace("\xa0", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if len(cleaned) > limit:
        # Prefer a word boundary so a name ends "blink two" rather than "blink tw",
        # but only when that still leaves something recognisable.
        clipped = cleaned[:limit]
        spaced = clipped.rsplit(" ", 1)[0]
        cleaned = spaced if len(spaced) >= 12 else clipped
    cleaned = cleaned.rstrip(". -")
    return cleaned or fallback


def part_name(base: str, number: int) -> str:
    """"<base> Part N", with the base shortened so "Part N" always survives.

    Truncating the whole string instead would cut the suffix off and leave every
    part sharing one indistinguishable name.
    """
    suffix = f" Part {number}"
    head = sanitize_name(base, "Slides", limit=MAX_NAME - len(suffix))
    return f"{head}{suffix}"


def unquote_path(raw: str) -> str:
    """Normalise a path typed, pasted, or dragged into a terminal."""
    value = (raw or "").strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
        value = value[1:-1]
    else:
        # Terminal escapes spaces and other shell-special characters on drop.
        value = re.sub(r"\\(.)", r"\1", value)
    return os.path.expanduser(value.strip())


def xml_attr(fragment: str, name: str) -> str | None:
    match = re.search(r'\b%s\s*=\s*["\']([^"\']*)["\']' % re.escape(name), fragment, re.I)
    return html.unescape(match.group(1)) if match else None


@dataclass
class Slide:
    index: int                       # 1-based position in the deck
    part: str                        # ppt/slides/slideN.xml
    title: str = ""
    notes: str = ""
    video_part: str | None = None    # ppt/media/mediaN.mp4 when this slide holds a video

    @property
    def has_video(self) -> bool:
        return self.video_part is not None


@dataclass
class Piece:
    kind: str                        # "pdf" | "video"
    name: str
    pages: list[int] = field(default_factory=list)   # 1-based source page numbers
    video_part: str | None = None
    notes: list[str] = field(default_factory=list)   # parallel to pages


class DeckError(RuntimeError):
    pass


def read_deck(path: str) -> list[Slide]:
    """Read slides in presentation order, with titles, notes and video links."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        def text(part: str) -> str:
            return z.read(part).decode("utf-8", errors="replace")

        # Presentation order is the r:id order in <p:sldIdLst>, not the numeric
        # filename order — decks that have been reordered differ between the two.
        pres, rels = text("ppt/presentation.xml"), text("ppt/_rels/presentation.xml.rels")
        targets: dict[str, str] = {}
        for rel in re.findall(r"<Relationship\b[^>]*/?>", rels, re.I):
            rid, typ, tgt = xml_attr(rel, "Id"), xml_attr(rel, "Type"), xml_attr(rel, "Target")
            if rid and tgt and typ and typ.lower().endswith("/slide"):
                targets[rid] = posixpath.normpath(posixpath.join("ppt", tgt))

        ordered: list[str] = []
        for entry in re.findall(r"<p:sldId\b[^>]*>", pres, re.I):
            rid = xml_attr(entry, "r:id") or xml_attr(entry, "id")
            if rid and targets.get(rid) in names:
                ordered.append(targets[rid])
        if not ordered:
            raise DeckError("No slides found in the presentation.")

        slides: list[Slide] = []
        for position, part in enumerate(ordered, start=1):
            slide = Slide(index=position, part=part)
            body = text(part)

            # A hidden slide is skipped by the PDF export, which would silently
            # shift every page after it. Refuse rather than mis-map.
            opening = re.search(r"<p:sld\b[^>]*>", body)
            if opening and re.search(r'show\s*=\s*["\']0["\']', opening.group(0)):
                raise DeckError(
                    f"{posixpath.basename(part)} is hidden. Hidden slides are dropped during PDF "
                    "export, so page numbers would no longer line up. Unhide it and re-run."
                )

            slide.title = extract_title(body)

            rel_part = posixpath.join(
                posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
            )
            if rel_part in names:
                rel_body = text(rel_part)
                for rel in re.findall(r"<Relationship\b[^>]*/?>", rel_body, re.I):
                    typ, tgt = xml_attr(rel, "Type"), xml_attr(rel, "Target")
                    if not typ or not tgt:
                        continue
                    resolved = posixpath.normpath(posixpath.join(posixpath.dirname(part), tgt))
                    if (
                        typ.lower().endswith("/video")
                        and os.path.splitext(tgt)[1].lower() in VIDEO_EXTENSIONS
                        and resolved in names
                    ):
                        slide.video_part = resolved
                    if typ.lower().endswith("/notesslide") and resolved in names:
                        slide.notes = extract_notes(text(resolved))
            slides.append(slide)
        return slides


def extract_title(slide_xml: str) -> str:
    """The title placeholder's text, used to name a video piece."""
    for shape in re.findall(r"<p:sp>.*?</p:sp>", slide_xml, re.S):
        placeholder = re.search(r"<p:ph\b[^>]*>", shape)
        if not placeholder:
            continue
        kind = (xml_attr(placeholder.group(0), "type") or "").lower()
        if kind in ("title", "ctrtitle"):
            runs = re.findall(r"<a:t>([^<]*)</a:t>", shape)
            return html.unescape(" ".join(runs)).strip()
    return ""


def extract_notes(notes_xml: str) -> str:
    """Only the body placeholder — the rest is the slide-image and page number."""
    for shape in re.findall(r"<p:sp>.*?</p:sp>", notes_xml, re.S):
        placeholder = re.search(r"<p:ph\b[^>]*>", shape)
        if not placeholder or (xml_attr(placeholder.group(0), "type") or "").lower() != "body":
            continue
        lines = []
        for para in re.findall(r"<a:p>.*?</a:p>", shape, re.S):
            buf = ""
            for token in re.finditer(r"<a:t>(.*?)</a:t>|<a:br\s*/?>", para, re.S):
                buf += html.unescape(token.group(1)) if token.group(1) is not None else "\n"
            lines.append(buf)
        return "\n".join(lines).replace("\r\n", "\n").strip()
    return ""


def plan(slides: list[Slide], base_name: str) -> list[Piece]:
    """Split the deck at every video slide, keeping reading order."""
    pieces: list[Piece] = []
    run: list[Slide] = []
    part_number = 1

    def flush() -> None:
        nonlocal run, part_number
        if not run:
            return
        pieces.append(
            Piece(
                kind="pdf",
                name=part_name(base_name, part_number),
                pages=[s.index for s in run],
                notes=[s.notes for s in run],
            )
        )
        part_number += 1
        run = []

    for slide in slides:
        if slide.has_video:
            # Consecutive videos leave no slides between them; flush() no-ops.
            flush()
            pieces.append(
                Piece(
                    kind="video",
                    name=sanitize_name(slide.title, f"Video {slide.index}"),
                    video_part=slide.video_part,
                )
            )
        else:
            run.append(slide)
    flush()

    # A deck with no video is just one PDF; calling it "Part 1" would imply a
    # part 2 that does not exist.
    if len(pieces) == 1 and pieces[0].kind == "pdf":
        pieces[0].name = sanitize_name(base_name, "Slides")
    return pieces


def convert_to_pdf(source: str, destination: str) -> str:
    """PowerPoint first (authored fidelity), LibreOffice second."""
    source, destination = os.path.abspath(source), os.path.abspath(destination)

    if sys.platform == "darwin" and os.path.isdir("/Applications/Microsoft PowerPoint.app"):
        script = (
            'on run argv\n'
            '  set src to item 1 of argv\n'
            '  set dst to item 2 of argv\n'
            '  tell application "Microsoft PowerPoint"\n'
            '    open (POSIX file src)\n'
            '    delay 2\n'
            '    set thePres to active presentation\n'
            '    save thePres in (POSIX file dst) as save as PDF\n'
            '    delay 1\n'
            '    close thePres saving no\n'
            '  end tell\n'
            'end run\n'
        )
        with tempfile.NamedTemporaryFile("w", suffix=".scpt", delete=False) as handle:
            handle.write(script)
            script_path = handle.name
        try:
            subprocess.run(
                ["osascript", script_path, source, destination],
                check=True, capture_output=True, text=True, timeout=1800,
            )
            if os.path.exists(destination) and os.path.getsize(destination) > 0:
                return "Microsoft PowerPoint"
        except subprocess.CalledProcessError as exc:
            print(f"  PowerPoint conversion failed ({exc.stderr.strip()}), trying LibreOffice…")
        finally:
            os.unlink(script_path)

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        outdir = os.path.dirname(destination)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, source],
            check=True, capture_output=True, text=True, timeout=1800,
        )
        produced = os.path.join(outdir, os.path.splitext(os.path.basename(source))[0] + ".pdf")
        if produced != destination and os.path.exists(produced):
            shutil.move(produced, destination)
        if os.path.exists(destination) and os.path.getsize(destination) > 0:
            return "LibreOffice"

    raise DeckError(
        "No converter available. Install Microsoft PowerPoint or LibreOffice "
        "(brew install --cask libreoffice)."
    )


def write_outputs(source: str, pdf_path: str, pieces: list[Piece], outdir: str) -> list[str]:
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(pdf_path)
    written: list[str] = []
    os.makedirs(outdir, exist_ok=True)

    with zipfile.ZipFile(source) as z:
        for number, piece in enumerate(pieces, start=1):
            stem = f"Content piece {number} - {piece.name}"

            if piece.kind == "video":
                suffix = os.path.splitext(piece.video_part)[1].lower()
                target = os.path.join(outdir, stem + suffix)
                with z.open(piece.video_part) as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst)
                written.append(target)
                continue

            writer = PdfWriter()
            for page in piece.pages:
                writer.add_page(reader.pages[page - 1])
            target = os.path.join(outdir, stem + ".pdf")
            with open(target, "wb") as handle:
                writer.write(handle)
            written.append(target)

            # Notes are renumbered to this PDF's own pages, since the player
            # looks them up by the page it is currently showing.
            if any(note.strip() for note in piece.notes):
                sidecar = {
                    "version": 1,
                    "source": "powerpoint-speaker-notes",
                    "pages": [
                        {"page": i, "notes": note}
                        for i, note in enumerate(piece.notes, start=1)
                    ],
                }
                notes_path = os.path.join(outdir, f".speaker-notes-piece-{number}.json")
                with open(notes_path, "w", encoding="utf-8") as handle:
                    json.dump(sidecar, handle, indent=2, ensure_ascii=False)
                written.append(notes_path)
    return written


def collect_decks(inputs: list[str], recursive: bool) -> list[str]:
    """Expand files and folders into a de-duplicated, ordered list of decks."""
    found: list[str] = []
    for raw in inputs:
        path = unquote_path(raw)
        if not path:
            continue
        if os.path.isdir(path):
            pattern = "**/*" if recursive else "*"
            for entry in sorted(glob.glob(os.path.join(path, pattern), recursive=recursive)):
                if entry.lower().endswith((".pptx", ".ppt")) and os.path.isfile(entry):
                    found.append(entry)
        elif os.path.isfile(path):
            found.append(path)
        else:
            print(f"  skipped (not found): {path}", file=sys.stderr)

    ordered, seen = [], set()
    for path in found:
        key = os.path.abspath(path)
        # PowerPoint's lock files start with "~$" and are not real decks.
        if key in seen or os.path.basename(path).startswith("~$"):
            continue
        seen.add(key)
        ordered.append(path)
    return ordered


def describe(pieces: list[Piece], slides: list[Slide]) -> None:
    by_media = {s.video_part: s.index for s in slides if s.has_video}
    for number, piece in enumerate(pieces, start=1):
        if piece.kind == "video":
            print(f"    {number}. [video] {piece.name}  (slide {by_media.get(piece.video_part, '?')})")
        else:
            span = (
                f"slides {piece.pages[0]}-{piece.pages[-1]}"
                if len(piece.pages) > 1
                else f"slide {piece.pages[0]}"
            )
            noted = sum(1 for note in piece.notes if note.strip())
            print(
                f"    {number}. [pdf]   {piece.name}  "
                f"({span}, {len(piece.pages)} page(s), {noted} with notes)"
            )


def convert_deck(
    deck: str,
    outroot: str,
    name_override: str | None = None,
    dry_run: bool = False,
) -> dict:
    """Convert one deck and return a structured result. Prints nothing.

    Shared by the CLI and the web app so the two cannot drift apart. Expected
    failures (a hidden slide, a missing converter) come back as `error` rather
    than raising, so one bad deck never aborts a batch.
    """
    stem = os.path.splitext(os.path.basename(deck))[0]
    result: dict = {
        "deck": stem,
        "path": deck,
        "ok": False,
        "slides": 0,
        "videos": 0,
        "pieces": [],
        "files": [],
        "outdir": None,
        "error": None,
    }

    try:
        slides = read_deck(deck)
    except (DeckError, zipfile.BadZipFile, KeyError, OSError) as exc:
        result["error"] = str(exc)
        return result

    result["slides"] = len(slides)
    result["videos"] = sum(1 for s in slides if s.has_video)

    pieces = plan(slides, name_override or stem)
    by_media = {s.video_part: s.index for s in slides if s.has_video}
    result["pieces"] = [
        {
            "number": number,
            "kind": piece.kind,
            "name": piece.name,
            "slide": by_media.get(piece.video_part) if piece.kind == "video" else None,
            "pages": len(piece.pages),
            "noted": sum(1 for note in piece.notes if note.strip()),
        }
        for number, piece in enumerate(pieces, start=1)
    ]

    if dry_run:
        result["ok"] = True
        return result

    # Each deck gets its own folder so pieces from different decks cannot collide.
    outdir = os.path.join(outroot, stem)
    result["outdir"] = outdir
    with tempfile.TemporaryDirectory() as tmp:
        full_pdf = os.path.join(tmp, "full.pdf")
        try:
            convert_to_pdf(deck, full_pdf)
        except (DeckError, subprocess.SubprocessError, OSError) as exc:
            result["error"] = f"conversion failed: {exc}"
            return result

        from pypdf import PdfReader

        pages = len(PdfReader(full_pdf).pages)
        if pages != len(slides):
            result["error"] = (
                f"PDF has {pages} pages but the deck has {len(slides)} slides; "
                "page numbers would not line up, so nothing was written"
            )
            return result

        try:
            result["files"] = write_outputs(deck, full_pdf, pieces, outdir)
        except OSError as exc:
            result["error"] = f"could not write output: {exc}"
            return result

    result["ok"] = True
    return result


def process_deck(deck: str, outroot: str, name_override: str | None, dry_run: bool) -> dict:
    """CLI wrapper around convert_deck: same work, with progress printed."""
    result = convert_deck(deck, outroot, name_override, dry_run)

    if result["error"] and not result["slides"]:
        print(f"    error: {result['error']}", file=sys.stderr)
        return result

    kind = f"{result['videos']} video(s) -> split" if result["videos"] else "no video -> single PDF"
    print(f"    {result['slides']} slides, {kind}")
    for piece in result["pieces"]:
        if piece["kind"] == "video":
            print(f"    {piece['number']}. [video] {piece['name']}  (slide {piece['slide']})")
        else:
            print(
                f"    {piece['number']}. [pdf]   {piece['name']}  "
                f"({piece['pages']} page(s), {piece['noted']} with notes)"
            )

    if result["error"]:
        print(f"    error: {result['error']}", file=sys.stderr)
    elif not dry_run:
        print(f"    wrote {len(result['files'])} file(s) -> {result['outdir']}")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Batch-convert PowerPoint decks. A deck with no video becomes one PDF; "
            "a deck with video is split into PDFs plus standalone video files."
        )
    )
    parser.add_argument(
        "decks", nargs="*",
        help="one or more .pptx files and/or folders (omit to be prompted)",
    )
    parser.add_argument("-o", "--outdir", default=None,
                        help="output folder (default: 'converted' beside the first deck)")
    parser.add_argument("--name", default=None,
                        help="base name for the PDF parts; only sensible for a single deck")
    parser.add_argument("--recursive", action="store_true",
                        help="search folders recursively")
    parser.add_argument("--dry-run", action="store_true",
                        help="show what would be produced, writing nothing")
    args = parser.parse_args()

    inputs = list(args.decks)
    if not inputs:
        prompt = (
            "Drag in the .pptx files or a folder (space-separated) and press Enter:\n> "
        )
        try:
            typed = input(prompt if sys.stdin.isatty() else "").strip()
        except EOFError:
            typed = ""
        if typed:
            # A drop of several files arrives space-separated with escaped spaces,
            # so split on unescaped whitespace only.
            inputs = [p for p in re.split(r"(?<!\\)\s+", typed) if p]

    if not inputs:
        print("error: no decks given.", file=sys.stderr)
        return 1

    decks = collect_decks(inputs, args.recursive)
    if not decks:
        print("error: no PowerPoint files found.", file=sys.stderr)
        return 1

    if args.name and len(decks) > 1:
        print("note: --name applies to every deck; ignoring it for a batch.", file=sys.stderr)
        args.name = None

    outroot = args.outdir or os.path.join(os.path.dirname(os.path.abspath(decks[0])), "converted")
    print(f"{len(decks)} deck(s) to convert -> {outroot}\n")

    results = []
    for position, deck in enumerate(decks, start=1):
        print(f"[{position}/{len(decks)}] {os.path.basename(deck)}")
        results.append(process_deck(deck, outroot, args.name, args.dry_run))
        print()

    ok = [r for r in results if r["ok"]]
    failed = [r for r in results if not r["ok"]]
    total_files = sum(len(r["files"]) for r in ok)
    split_decks = [r for r in ok if r["videos"]]

    print("=" * 60)
    print(f"converted {len(ok)}/{len(results)} deck(s)"
          + (f", {total_files} file(s) written" if not args.dry_run else " (dry run)"))
    if split_decks:
        print(f"  {len(split_decks)} deck(s) contained video and were split")
    for r in failed:
        print(f"  FAILED  {r['deck']}: {r['error']}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
