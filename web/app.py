"""FastAPI front-end for qrack.

/api/parse (confirm/edit the new-node guess), /api/rack-split (suggest an
auto-split for the confirm step's editable rack-count fields),
/api/render and /api/preview (render, possibly against a template and/or
a manual rack split), and /api/derive-template (strip a template down to
just its theme) -- see CLAUDE.md. Uploads are untrusted: capped size,
magic-byte checked, parsed/rendered against a time box, and every render
happens in its own temp file that gets deleted after streaming -- never a
shared/static path.
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
from qumulo_rack.renderer import auto_rack_split, available_stats, derive_template, render_rack

MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB -- sizing reports are a few hundred KB
MAX_TEMPLATE_BYTES = 80 * 1024 * 1024  # 80 MB -- a real internal Qumulo template ran ~40MB with embedded video/images
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
        None, description="An existing .pptx, base64-encoded, to append the rack slide to and match colors sampled from its real content. Omit or null for no template."
    )
    template_slide: int | None = Field(
        None, description="1-based slide number in the template to sample colors from instead of the whole deck. Only meaningful alongside a template_base64 that still has its original slides."
    )
    rack_sizes: list[int] | None = Field(
        None, description="Explicit node count per rack (must sum to the report's total node count). Omit to auto-split by RU whenever the cluster needs more than one rack; up to 2 racks share one slide, more spill onto additional slides with the aggregated stats on a dedicated final slide."
    )


class DeriveTemplateRequest(BaseModel):
    template_base64: str = Field(..., description="An existing .pptx, base64-encoded, to strip down to just its theme/layouts.")
    template_slide: int | None = Field(
        None, description="1-based slide number to sample colors from instead of the whole deck."
    )


class RackSplitRequest(BaseModel):
    report: dict = Field(..., description="ClusterReport.as_dict() output, possibly edited by the user")


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


@app.post("/api/rack-split")
def api_rack_split(req: RackSplitRequest):
    """The node count per rack `render_rack` would use by default (no
    `rack_sizes` override) -- for the confirm step to show a suggested
    split and let the user edit it before rendering, without duplicating
    the auto-split logic in JS. Always `{"rack_sizes": [total node count]}`
    (one rack) for a cluster that fits in one."""
    report = _build_report(req)
    return {"rack_sizes": auto_rack_split(report)}


@app.post("/api/derive-template")
def api_derive_template(req: DeriveTemplateRequest):
    """Web equivalent of derive_template.py: strips every slide out of an
    uploaded .pptx, keeping only its masters/layouts/theme, and hands the
    result back as base64 -- for when the user wants the rack slide to
    match a company deck's style without literally inserting it into that
    whole deck. The browser swaps the response straight in as the active
    template (and persists that smaller file instead of the original)."""
    with _resolve_template(req.template_base64) as template_path:
        if template_path is None:
            raise HTTPException(422, "No template provided.")
        with tempfile.NamedTemporaryFile(suffix=".pptx") as out_tmp:
            try:
                derive_template(template_path, out_tmp.name, slide_index=req.template_slide)
            except ValueError as exc:
                raise HTTPException(422, str(exc)) from exc
            except Exception as exc:
                raise HTTPException(422, f"Couldn't use that file as a template: {exc}") from exc
            derived_bytes = Path(out_tmp.name).read_bytes()

    return {"template_base64": base64.b64encode(derived_bytes).decode("ascii")}


@app.post("/api/render")
def api_render(req: RenderRequest):
    report = _build_report(req)

    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    tmp.close()
    try:
        with _resolve_template(req.template_base64) as template_path:
            render_rack(report, tmp.name, rack_label=req.rack_label or None,
                        visible_stats=req.visible_stats, template_path=template_path,
                        template_slide=req.template_slide, rack_sizes=req.rack_sizes)
    except HTTPException:
        os.unlink(tmp.name)
        raise
    except ValueError as exc:
        os.unlink(tmp.name)
        raise HTTPException(422, str(exc)) from exc
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
        template_slide_count = 0
        try:
            with _resolve_template(req.template_base64) as template_path:
                if template_path is not None:
                    template_slide_count = len(Presentation(template_path).slides)
                render_rack(report, pptx_path, rack_label=req.rack_label or None,
                            visible_stats=req.visible_stats, template_path=template_path,
                            template_slide=req.template_slide, rack_sizes=req.rack_sizes)
        except HTTPException:
            raise
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from exc
        except Exception as exc:
            if req.template_base64:
                raise HTTPException(422, f"Couldn't use that template: {exc}") from exc
            raise

        # Our slide(s) are always appended after any template slides, so
        # the first one we added is at this fixed position regardless of
        # how many we ended up adding (1, for the common case; more for a
        # cluster split across multiple slides -- see render_rack's
        # rack_groups). Deliberately the *first* of ours, not the last: a
        # multi-slide render's last slide is the aggregated stats slide,
        # but the rack diagrams -- the part actually worth a visual sanity
        # check -- are the earlier ones. soffice's PNG export only ever
        # renders slide 1 of a multi-slide deck with no way to target
        # another one, so we convert to PDF (which renders every page) and
        # then extract the one page we want with pdftoppm.
        target_page = template_slide_count + 1

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
