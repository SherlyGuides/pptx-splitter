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

Hosted mode (PPTX_SPLITTER_HOSTED=1, set by the Dockerfile) is for running it
on a server other people reach over the internet. Pasted paths and the output
folder are switched off there, since both would hand visitors the server's own
filesystem; every page sits behind the APP_PASSWORD shared password; results
live in a private folder per job and are deleted after JOB_TTL_HOURS.

Behind a site that already asks for a password (nginx basic auth), set
TRUST_PROXY_AUTH=1 instead of APP_PASSWORD, and have the proxy send
X-Forwarded-Prefix when the app is mounted under a path such as /splitter.
"""

from __future__ import annotations

import hmac
import os
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import datetime

from flask import (
    Flask, Response, abort, jsonify, redirect, render_template, request, send_file, url_for,
)

from werkzeug.middleware.proxy_fix import ProxyFix

import split_pptx

HOSTED = os.environ.get("PPTX_SPLITTER_HOSTED", "").strip().lower() in ("1", "true", "yes")
PASSWORD = os.environ.get("APP_PASSWORD", "")
# The proxy in front already asks for a password; asking again would clash,
# since both use the same Authorization header.
TRUST_PROXY_AUTH = os.environ.get("TRUST_PROXY_AUTH", "").strip().lower() in ("1", "true", "yes")
JOB_TTL_SECONDS = float(os.environ.get("JOB_TTL_HOURS", "6")) * 3600

if HOSTED and not PASSWORD and not TRUST_PROXY_AUTH:
    # An open converter on the internet is free compute for anyone who finds it.
    raise SystemExit(
        "Hosted mode needs a password: set APP_PASSWORD, or TRUST_PROXY_AUTH=1 "
        "when a proxy in front already asks for one."
    )

app = Flask(__name__)
if HOSTED:
    # Behind a proxy: honour its scheme, host and path prefix (/splitter) so
    # generated links point where the visitor actually is.
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)
# Decks are routinely over 100 MB; the default 16 MB cap would reject them.
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "4096")) * 1024 * 1024

DEFAULT_OUTPUT = (
    os.path.join(tempfile.gettempdir(), "pptx-splitter-jobs")
    if HOSTED
    else os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
)

# A server has no business converting several 250 MB decks at once: LibreOffice
# is memory-hungry, and parallel runs would only make every job slower. Jobs
# beyond this wait their turn.
UPLOAD_ROOT = os.path.join(tempfile.gettempdir(), "pptx-splitter-uploads")

CONVERSIONS = threading.Semaphore(int(os.environ.get("MAX_PARALLEL_JOBS", "1")))


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
    created: float = field(default_factory=time.time)

    def as_dict(self) -> dict:
        return {
            "id": self.id,
            "status": self.status,
            "current": self.current,
            "total": len(self.decks),
            # A server's internal folder means nothing to a visitor.
            "outroot": None if HOSTED else self.outroot,
            "ttl_hours": round(JOB_TTL_SECONDS / 3600) if HOSTED else None,
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
    try:
        with CONVERSIONS:
            job.status = "running"
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


def bundle_path(job_id: str) -> str:
    return os.path.join(tempfile.gettempdir(), f"pptx-splitter-{job_id}.zip")


def expire_old_jobs() -> None:
    """Hosted only: a server's disk would otherwise fill with everyone's decks."""
    if not HOSTED:
        return
    cutoff = time.time() - JOB_TTL_SECONDS
    with LOCK:
        stale = [j for j in JOBS.values() if j.status == "done" and j.created < cutoff]
        for job in stale:
            del JOBS[job.id]
    for job in stale:
        shutil.rmtree(job.outroot, ignore_errors=True)
        try:
            os.remove(bundle_path(job.id))
        except OSError:
            pass
    # Uploads that were started but never handed to a job.
    if os.path.isdir(UPLOAD_ROOT):
        for name in os.listdir(UPLOAD_ROOT):
            path = os.path.join(UPLOAD_ROOT, name)
            try:
                if os.path.getmtime(path) < cutoff:
                    shutil.rmtree(path, ignore_errors=True)
            except OSError:
                pass


@app.before_request
def require_password():
    if TRUST_PROXY_AUTH or not PASSWORD or request.endpoint in ("healthz", "static"):
        return None
    auth = request.authorization
    supplied = (auth.password or "") if auth else ""
    if hmac.compare_digest(supplied.encode(), PASSWORD.encode()):
        return None
    return Response(
        "Password required.", 401,
        {"WWW-Authenticate": 'Basic realm="PPTX Splitter", charset="UTF-8"'},
    )


def upload_dir_for(upload_id: str) -> str:
    # The id is ours (hex from uuid4); anything else is someone poking at paths.
    if not upload_id.isalnum() or len(upload_id) != 32:
        abort(404)
    path = os.path.join(UPLOAD_ROOT, upload_id)
    if not os.path.isdir(path):
        abort(404)
    return path


@app.post("/api/upload/start")
def upload_start():
    upload_id = uuid.uuid4().hex
    os.makedirs(os.path.join(UPLOAD_ROOT, upload_id))
    return jsonify(upload_id=upload_id)


@app.post("/api/upload/<upload_id>")
def upload_chunk(upload_id: str):
    """Append one chunk of one deck. Chunks arrive in order; the offset is
    checked so a retried or out-of-order chunk can never corrupt a deck."""
    folder = upload_dir_for(upload_id)
    name = os.path.basename(request.form.get("name") or "")
    chunk = request.files.get("chunk")
    if not name.lower().endswith((".pptx", ".ppt")) or name.startswith("~$") or not chunk:
        return "Only .pptx and .ppt files can be uploaded.", 400
    target = os.path.join(folder, name)
    offset = int(request.form.get("offset") or 0)
    current = os.path.getsize(target) if os.path.exists(target) else 0
    if offset != current:
        return f"Chunk out of order (expected offset {current}, got {offset}).", 409
    with open(target, "ab") as handle:
        shutil.copyfileobj(chunk.stream, handle)
    return jsonify(received=os.path.getsize(target))


@app.get("/healthz")
def healthz():
    return "ok"


@app.get("/")
def index():
    return render_template("index.html", default_output=DEFAULT_OUTPUT, hosted=HOSTED)


@app.post("/convert")
def convert():
    expire_old_jobs()
    job_id = uuid.uuid4().hex[:12]
    if HOSTED:
        # Each job gets a folder of its own; the visitor never chooses one.
        outroot = os.path.join(DEFAULT_OUTPUT, job_id)
    else:
        outroot = (request.form.get("outdir") or "").strip() or DEFAULT_OUTPUT
        outroot = split_pptx.unquote_path(outroot)

    decks: list[str] = []
    upload_dir: str | None = None

    # Paths already on this machine are read in place — no copying. Never on a
    # server: they would let a visitor convert, and download, any file on it.
    typed = "" if HOSTED else (request.form.get("paths") or "").strip()
    if typed:
        decks.extend(split_pptx.collect_decks(
            [line for line in typed.splitlines() if line.strip()],
            recursive=bool(request.form.get("recursive")),
        ))

    # Decks already sent in chunks by the page's script.
    upload_id = (request.form.get("upload_id") or "").strip()
    if upload_id:
        upload_dir = upload_dir_for(upload_id)
        decks.extend(
            os.path.join(upload_dir, name) for name in sorted(os.listdir(upload_dir))
        )

    # A plain form post (no script) still carries the files themselves.
    uploads = [f for f in request.files.getlist("decks") if f and f.filename]
    if uploads and not upload_dir:
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
            hosted=HOSTED,
            error=(
                "No PowerPoint files found. Choose one or more .pptx files."
                if HOSTED
                else "No PowerPoint files found. Choose files, or paste a path to a deck or folder."
            ),
        ), 400

    job = Job(id=job_id, outroot=outroot, decks=decks, upload_dir=upload_dir)
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

    bundle = bundle_path(job_id)
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in produced:
            # Keep each deck's folder, so pieces stay grouped inside the zip.
            archive.write(path, os.path.relpath(path, job.outroot))
    return send_file(bundle, as_attachment=True, download_name="converted.zip")


if __name__ == "__main__":
    os.makedirs(DEFAULT_OUTPUT, exist_ok=True)
    port = int(os.environ.get("PORT", "5000"))
    print("\n  PPTX → PDF converter")
    print(f"  open http://127.0.0.1:{port}")
    print(f"  output folder: {DEFAULT_OUTPUT}\n")
    # The reloader would restart mid-conversion. A server must listen on every
    # interface to be reachable; locally, only this machine should reach it.
    app.run(
        host="0.0.0.0" if HOSTED else "127.0.0.1",
        port=port, debug=False, use_reloader=False,
    )
