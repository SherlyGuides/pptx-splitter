#!/usr/bin/env python3
"""Local web page for turning PowerPoint decks into PDFs.

Run it, open the page, hand it decks, get PDFs back:

    python3 app.py            then open http://127.0.0.1:5000

A deck with no video becomes one PDF. A deck with video is split at the video
slides into PDFs plus standalone video files, so nothing is lost.

All the real work lives in split_pptx.py, which the command-line version uses
too — the page is only a front door, so the two can never disagree about what
a deck should turn into.

Two ways to hand over decks, because a 250 MB deck is slow to push through a
browser: pick files to upload, or paste the path of a file or folder already on
this machine, which is read in place.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import threading
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import datetime

from flask import Flask, abort, jsonify, redirect, render_template, request, send_file, url_for

import split_pptx

app = Flask(__name__)
# Decks are routinely over 100 MB; the default 16 MB cap would reject them.
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024 * 1024

DEFAULT_OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")


@dataclass
class Job:
    id: str
    outroot: str
    decks: list[str]
    upload_dir: str | None = None
    status: str = "queued"          # queued | running | done
    current: int = 0
    results: list[dict] = field(default_factory=list)
    error: str | None = None
    started: str = field(default_factory=lambda: datetime.now().strftime("%H:%M:%S"))

    def as_dict(self) -> dict:
        return {
            "id": self.id,
            "status": self.status,
            "current": self.current,
            "total": len(self.decks),
            "outroot": self.outroot,
            "error": self.error,
            "started": self.started,
            "results": [
                {
                    **r,
                    "files": [
                        {"name": os.path.basename(f), "size": file_size(f), "path": f}
                        for f in r.get("files", [])
                    ],
                }
                for r in self.results
            ],
        }


JOBS: dict[str, Job] = {}
LOCK = threading.Lock()


def file_size(path: str) -> str:
    try:
        mb = os.path.getsize(path) / 1048576
    except OSError:
        return "?"
    return f"{mb:.1f} MB" if mb >= 0.1 else f"{os.path.getsize(path) / 1024:.0f} KB"


def run_job(job: Job) -> None:
    job.status = "running"
    try:
        for index, deck in enumerate(job.decks, start=1):
            job.current = index
            job.results.append(split_pptx.convert_deck(deck, job.outroot))
    except Exception as exc:  # a crash here must not leave the page spinning forever
        job.error = f"{type(exc).__name__}: {exc}"
    finally:
        job.status = "done"
        job.current = len(job.decks)
        # Uploaded copies are only needed during conversion; the outputs stay.
        if job.upload_dir:
            shutil.rmtree(job.upload_dir, ignore_errors=True)


@app.get("/")
def index():
    return render_template("index.html", default_output=DEFAULT_OUTPUT)


@app.post("/convert")
def convert():
    outroot = (request.form.get("outdir") or "").strip() or DEFAULT_OUTPUT
    outroot = split_pptx.unquote_path(outroot)

    decks: list[str] = []
    upload_dir: str | None = None

    # Paths already on this machine are read in place — no copying.
    typed = (request.form.get("paths") or "").strip()
    if typed:
        decks.extend(split_pptx.collect_decks(
            [line for line in typed.splitlines() if line.strip()],
            recursive=bool(request.form.get("recursive")),
        ))

    uploads = [f for f in request.files.getlist("decks") if f and f.filename]
    if uploads:
        upload_dir = tempfile.mkdtemp(prefix="pptx-upload-")
        for item in uploads:
            name = os.path.basename(item.filename)
            if not name.lower().endswith((".pptx", ".ppt")) or name.startswith("~$"):
                continue
            target = os.path.join(upload_dir, name)
            item.save(target)
            decks.append(target)

    if not decks:
        return render_template(
            "index.html",
            default_output=DEFAULT_OUTPUT,
            error="No PowerPoint files found. Choose files, or paste a path to a deck or folder.",
        ), 400

    job = Job(id=uuid.uuid4().hex[:12], outroot=outroot, decks=decks, upload_dir=upload_dir)
    with LOCK:
        JOBS[job.id] = job
    threading.Thread(target=run_job, args=(job,), daemon=True).start()
    return redirect(url_for("job_page", job_id=job.id))


@app.get("/job/<job_id>")
def job_page(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        abort(404)
    return render_template("job.html", job=job.as_dict())


@app.get("/api/job/<job_id>")
def job_api(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        abort(404)
    return jsonify(job.as_dict())


@app.get("/download/<job_id>")
def download(job_id: str):
    """Serve one produced file, identified by its path within this job."""
    job = JOBS.get(job_id)
    if not job:
        abort(404)
    wanted = request.args.get("path", "")
    produced = {f for result in job.results for f in result.get("files", [])}
    # Only ever serve a file this job actually produced.
    if wanted not in produced or not os.path.isfile(wanted):
        abort(404)
    return send_file(wanted, as_attachment=True)


@app.get("/download-all/<job_id>")
def download_all(job_id: str):
    job = JOBS.get(job_id)
    if not job or job.status != "done":
        abort(404)
    produced = [f for result in job.results for f in result.get("files", []) if os.path.isfile(f)]
    if not produced:
        abort(404)

    bundle = os.path.join(tempfile.gettempdir(), f"pptx-splitter-{job_id}.zip")
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in produced:
            # Keep each deck's folder, so pieces stay grouped inside the zip.
            archive.write(path, os.path.relpath(path, job.outroot))
    return send_file(bundle, as_attachment=True, download_name="converted.zip")


if __name__ == "__main__":
    os.makedirs(DEFAULT_OUTPUT, exist_ok=True)
    print("\n  PPTX → PDF converter")
    print("  open http://127.0.0.1:5000")
    print(f"  output folder: {DEFAULT_OUTPUT}\n")
    # Local tool, single user: the reloader would restart mid-conversion.
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
