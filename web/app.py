"""FastAPI front-end for qrack.

Two endpoints, per CLAUDE.md: parse (so the user can confirm/edit the
new-node guess) and render (so edits actually take effect). Uploads are
untrusted: capped size, PDF-only, parsed against a time box, and every
render happens in its own temp file that gets deleted after streaming --
never a shared/static path.
"""

from __future__ import annotations

import base64
import binascii
import contextlib
import os
import subprocess
import tempfile
import time
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, Response
from pptx import Presentation
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

from qumulo_rack.parser import ClusterReport, parse_report
from qumulo_rack.renderer import available_stats, render_rack

MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB -- sizing reports are a few hundred KB
MAX_TEMPLATE_BYTES = 15 * 1024 * 1024  # 15 MB -- a branded deck with embedded images/logos
PARSE_TIMEOUT_SECONDS = 20
PREVIEW_TIMEOUT_SECONDS = 25
PREVIEW_WIDTH_PX = 2400
PREVIEW_HEIGHT_PX = 1350

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
    template_base64: str | None = Field(
        None, description="An existing .pptx, base64-encoded, to append the rack slide to and pick up its theme colors from. Omit or null for no template."
    )


def _reject_if_not_a_sizing_report(report: ClusterReport) -> None:
    if not report.models:
        raise HTTPException(422, "Couldn't find an 'All Nodes' section -- this doesn't look like a Qumulo sizing report.")
    if report.usable_tb is None:
        raise HTTPException(422, "Couldn't find usable capacity -- this doesn't look like a Qumulo sizing report.")


def _build_report(req: "RenderRequest") -> ClusterReport:
    try:
        report = ClusterReport.from_dict(req.report)
    except TypeError as exc:
        raise HTTPException(422, f"Malformed report config: {exc}") from exc
    _reject_if_not_a_sizing_report(report)
    return report


@contextlib.contextmanager
def _resolve_template(template_base64: str | None):
    """Decodes a base64 .pptx into a temp file for the duration of the
    `with` block, or yields None if no template was given. A .pptx is a
    zip archive, so a real one always starts with the zip magic bytes --
    cheap way to reject garbage before handing it to python-pptx."""
    if not template_base64:
        yield None
        return

    try:
        data = base64.b64decode(template_base64, validate=True)
    except binascii.Error as exc:
        raise HTTPException(422, f"Template isn't valid base64: {exc}") from exc
    if len(data) > MAX_TEMPLATE_BYTES:
        raise HTTPException(413, f"Template too large (max {MAX_TEMPLATE_BYTES // (1024 * 1024)} MB).")
    if not data.startswith(b"PK\x03\x04"):
        raise HTTPException(422, "That doesn't look like a .pptx file.")

    with tempfile.NamedTemporaryFile(suffix=".pptx") as tmp:
        tmp.write(data)
        tmp.flush()
        yield tmp.name


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
    report = _build_report(req)

    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    tmp.close()
    try:
        with _resolve_template(req.template_base64) as template_path:
            render_rack(report, tmp.name, rack_label=req.rack_label or None,
                        visible_stats=req.visible_stats, template_path=template_path)
    except HTTPException:
        os.unlink(tmp.name)
        raise
    except Exception as exc:
        os.unlink(tmp.name)
        if req.template_base64:
            raise HTTPException(422, f"Couldn't use that template: {exc}") from exc
        raise

    base = (report.source_file or "rack").rsplit(".", 1)[0]
    filename = f"{base}.rack.pptx"
    return FileResponse(
        tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
        background=BackgroundTask(os.unlink, tmp.name),
    )


@app.post("/api/preview")
def api_preview(req: RenderRequest):
    """Renders the exact same .pptx /api/render would produce, then
    rasterizes it with LibreOffice -- so the preview can never drift from
    what a download actually looks like, at the cost of a few seconds and
    an external process per request."""
    report = _build_report(req)

    with tempfile.TemporaryDirectory() as tmp_dir:
        pptx_path = os.path.join(tmp_dir, "slide.pptx")
        try:
            with _resolve_template(req.template_base64) as template_path:
                render_rack(report, pptx_path, rack_label=req.rack_label or None,
                            visible_stats=req.visible_stats, template_path=template_path)
        except HTTPException:
            raise
        except Exception as exc:
            if req.template_base64:
                raise HTTPException(422, f"Couldn't use that template: {exc}") from exc
            raise

        # Our slide is always the last one in the deck (appended after any
        # template slides). soffice's PNG export only ever renders slide 1 of
        # a multi-slide deck with no way to target another one, so we convert
        # to PDF (which renders every page) and then extract the one page we
        # want with pdftoppm.
        target_page = len(Presentation(pptx_path).slides)

        # A dedicated profile dir per request avoids soffice's user-profile
        # lock contention under concurrent preview requests.
        profile_dir = os.path.join(tmp_dir, "lo_profile")
        try:
            result = subprocess.run(
                ["soffice", "--headless", f"-env:UserInstallation=file://{profile_dir}",
                 "--convert-to", "pdf", "--outdir", tmp_dir, pptx_path],
                capture_output=True, text=True, timeout=PREVIEW_TIMEOUT_SECONDS,
            )
        except FileNotFoundError:
            raise HTTPException(501, "Preview isn't available -- LibreOffice isn't installed on this server.")
        except subprocess.TimeoutExpired:
            raise HTTPException(504, "Preview rendering took too long.")

        pdf_path = os.path.join(tmp_dir, "slide.pdf")
        if result.returncode != 0 or not os.path.exists(pdf_path):
            detail = (result.stderr or "").strip()[-500:]
            raise HTTPException(500, f"Preview rendering failed: {detail or 'unknown error'}")

        # 180 DPI on our fixed 13.333x7.5in slide yields exactly
        # PREVIEW_WIDTH_PX x PREVIEW_HEIGHT_PX (2400x1350).
        png_prefix = os.path.join(tmp_dir, "page")
        try:
            result = subprocess.run(
                ["pdftoppm", "-png", "-r", "180", "-f", str(target_page), "-l", str(target_page),
                 "-singlefile", pdf_path, png_prefix],
                capture_output=True, text=True, timeout=PREVIEW_TIMEOUT_SECONDS,
            )
        except FileNotFoundError:
            raise HTTPException(501, "Preview isn't available -- poppler-utils (pdftoppm) isn't installed on this server.")
        except subprocess.TimeoutExpired:
            raise HTTPException(504, "Preview rendering took too long.")

        png_path = png_prefix + ".png"
        if result.returncode != 0 or not os.path.exists(png_path):
            detail = (result.stderr or "").strip()[-500:]
            raise HTTPException(500, f"Preview rendering failed: {detail or 'unknown error'}")

        png_bytes = Path(png_path).read_bytes()

    return Response(content=png_bytes, media_type="image/png")
