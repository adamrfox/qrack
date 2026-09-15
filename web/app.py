"""FastAPI front-end for qrack.

Two endpoints, per CLAUDE.md: parse (so the user can confirm/edit the
new-node guess) and render (so edits actually take effect). Uploads are
untrusted: capped size, PDF-only, parsed against a time box, and every
render happens in its own temp file that gets deleted after streaming --
never a shared/static path.
"""

from __future__ import annotations

import os
import tempfile
import time
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

from qumulo_rack.parser import ClusterReport, parse_report
from qumulo_rack.renderer import available_stats, render_rack

MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB -- sizing reports are a few hundred KB
PARSE_TIMEOUT_SECONDS = 20

STATIC_DIR = Path(__file__).parent / "static"

app = FastAPI(title="qrack")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)


class RenderRequest(BaseModel):
    report: dict = Field(..., description="ClusterReport.as_dict() output, possibly edited by the user")
    rack_label: str | None = None
    visible_stats: list[str] | None = Field(
        None, description="Stat keys to show in the stats panel (see /api/parse's available_stats). Omit or null shows everything."
    )


def _reject_if_not_a_sizing_report(report: ClusterReport) -> None:
    if not report.models:
        raise HTTPException(422, "Couldn't find an 'All Nodes' section -- this doesn't look like a Qumulo sizing report.")
    if report.usable_tb is None:
        raise HTTPException(422, "Couldn't find usable capacity -- this doesn't look like a Qumulo sizing report.")


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    # This page is under active development and small; always revalidate
    # rather than let a browser silently keep serving an old tab's copy
    # indefinitely (bit us once already -- a stale tab looked like a
    # localStorage persistence bug that wasn't actually one).
    return HTMLResponse((STATIC_DIR / "index.html").read_text(), headers={"Cache-Control": "no-cache"})


@app.post("/api/parse")
async def api_parse(file: UploadFile = File(...)):
    if file.content_type not in ("application/pdf", "application/octet-stream"):
        raise HTTPException(422, f"Expected a PDF upload, got content-type {file.content_type!r}.")

    data = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, f"File too large (max {MAX_UPLOAD_BYTES // (1024 * 1024)} MB).")
    if not data.startswith(b"%PDF-"):
        raise HTTPException(422, "That doesn't look like a PDF file.")

    # Never trust the upload filename for anything but a display label.
    display_name = os.path.basename(file.filename or "upload.pdf")

    start = time.monotonic()
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf") as tmp:
            tmp.write(data)
            tmp.flush()
            report = parse_report(tmp.name, name=display_name)
    except Exception as exc:
        raise HTTPException(422, f"Couldn't parse this PDF: {exc}") from exc
    if time.monotonic() - start > PARSE_TIMEOUT_SECONDS:
        raise HTTPException(422, "Parsing this report took too long.")

    _reject_if_not_a_sizing_report(report)
    return {"report": report.as_dict(), "available_stats": available_stats(report)}


@app.post("/api/render")
def api_render(req: RenderRequest):
    try:
        report = ClusterReport.from_dict(req.report)
    except TypeError as exc:
        raise HTTPException(422, f"Malformed report config: {exc}") from exc
    _reject_if_not_a_sizing_report(report)

    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    tmp.close()
    try:
        render_rack(report, tmp.name, rack_label=req.rack_label or None, visible_stats=req.visible_stats)
    except Exception:
        os.unlink(tmp.name)
        raise

    base = (report.source_file or "rack").rsplit(".", 1)[0]
    filename = f"{base}.rack.pptx"
    return FileResponse(
        tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
        background=BackgroundTask(os.unlink, tmp.name),
    )
