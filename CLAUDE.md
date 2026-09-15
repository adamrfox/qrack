# CLAUDE.md — qrack

Context for coding agents working in this repo. Read this before changing code.

## What this is

`qrack` turns a Qumulo sizing-tool PDF (the report from `sizing.qumulo.com`)
into a single **editable** PowerPoint slide: a rack elevation (2 top-of-rack
switches + nodes stacked by rack-unit height, newly added nodes highlighted,
colour-coded cabling) next to a stats panel (capacity, performance, rack &
power). The point is to replace a tedious manual drawing task with a
one-command, standardized slide.

## The one rule that must not break: determinism

The whole value proposition is **"same input → same slide, every time."**

- **Extraction is deterministic.** The sizing report format is stable, so
  `parser.py` uses `pdfplumber` + targeted regex. Do **not** add an LLM, a
  vision model, or any nondeterministic step to the core parse or render path.
- **Rendering is deterministic geometry.** `renderer.py` computes positions from
  the parsed config. Never make the layout depend on anything random or
  time/locale-dependent.

If a future report format drifts and the regex parser misses fields, the fix is
to extend the format-targeted parser (or add a clearly-separated fallback), not
to bolt fuzzy reading into the core.

## Repo layout

```
qumulo_rack/
  __init__.py
  parser.py      # PDF -> ClusterReport   (deterministic, format-targeted)
  renderer.py    # ClusterReport -> .pptx (deterministic geometry, native shapes)
qrack.py         # CLI wrapper
web/             # FastAPI front-end (parse/confirm/render) — see "Web front-end" below
  app.py
  static/index.html
Dockerfile
requirements.txt
README.md
CLAUDE.md
samples/         # example sizing PDFs for regression testing (gitignored -- these
                 # tend to be real customer/prospect data; keep local only)
derive_template.py  # CLI: strip an existing .pptx down to just its theme/layouts
```

Keep `parser.py` and `renderer.py` free of any web/CLI/framework imports so
they stay reusable behind every front-end.

## Module API (current, authoritative)

### `qumulo_rack.parser`

- `parse_report(pdf_path: str) -> ClusterReport` — the only entry point.
- `ClusterReport` (dataclass) fields: `title`, `is_expansion`, `node_count`,
  `added_nodes`, `models: list[NodeModel]`, `usable_tb`, `raw_tb`,
  `added_capacity_tb`, `encoding`, `efficiency`, `scale_to_tb`, `license`,
  `perf: dict` (keys: `cached_read`, `uncached_read`, `sustained_write`,
  `burst_write`, `single_stream_write`, `ss_cached_read`, `ss_uncached_read`,
  `iops`), `write_volume_max`, `cluster_overwrite_cadence_max` (both strings —
  the PDF's own units/ranges are inconsistent enough across reports that
  these aren't worth normalizing into numbers), `drive_outage_tolerance`,
  `node_outage_tolerance`, `frontend_ports_total`, `backend_ports_total`,
  `frontend_networking`, `backend_networking`, `rack_u`, `additional_rack_u`,
  `weight_lbs`, `added_weight_lbs`, `watts`, `added_watts`, `amps_240`,
  `amps_110`, `thermal_btu`, `source_file`, `extra_stats: dict[str, str]`
  (best-effort catch-all — see below).
  - `.as_dict()` → dict plus derived `ru` per model, `total_ru_from_models`,
    and `new_node_summary`.
- `NodeModel` (dataclass) fields: `code`, `description`, `raw_tb_per_node`,
  `count`, `height_in`, `frontend_ports`, `backend_ports`, `is_new`,
  `new_count`; property `.ru` = `max(1, ceil(height_in / 1.75))`.

Unparsed fields come back `None` rather than raising — callers must handle that
(see web validation below).

### `qumulo_rack.renderer`

- `render_rack(report: ClusterReport, out_path: str, rack_label: str|None=None,
  visible_stats=None, template_path: str|None=None) -> str` writes the
  `.pptx` to `out_path` and returns it.
- Palette is module-level constants at the top of the file (node / switch /
  cable colours). Centralize any theming there.
- Slide is fixed 13.333" × 7.5" (16:9). All shapes are native `python-pptx`
  rectangles / connectors / text boxes, so the deck is hand-editable after
  generation — do not flatten anything to an image.
- **`template_path`**: when given, `Presentation(template_path)` is used as
  the base deck instead of a blank one — the rack slide is *appended* after
  whatever slides the template already has, and picks up the template's
  theme colors (title text = theme `dk1`, slide background = theme `lt1`,
  stat-card headers = theme `accent1`, via `_theme_colors()`), falling back
  to the normal palette constants for anything the theme doesn't define.
  Functional colours (new-node green, Switch A/B cable colours) stay fixed
  regardless of template, since they carry meaning. `_find_blank_layout()`
  picks the emptiest layout in the template (preferring one literally named
  "blank") by counting non-content placeholders (date/footer/slide-number
  don't count) — there's no schema flag for "this is the blank layout," so
  this is a heuristic, not a guarantee, on an unusual template. Colors are
  resolved into local variables and threaded through as function params,
  never mutated on shared module state, since the web app can render
  concurrently.
- **Known limitation:** the slide is always fixed to 13.333"×7.5" (16:9), so
  a 4:3 template gets its aspect ratio silently overridden. Not worth
  handling until someone actually hits it.
- `derive_template(source_path: str, out_path: str) -> str` strips every
  slide out of an existing `.pptx` via the standard python-pptx recipe
  (remove each `<p:sldId>` from `prs.slides._sldIdLst` and `drop_rel` its
  relationship — there's no public "delete slide" API), leaving only the
  masters/layouts/theme that `template_path` above actually reads. Meant as
  a one-time distillation step: point it at someone's real 40-slide branded
  deck once, get back a small style-only file, use *that* as `template_path`
  from then on instead of carrying the original deck's slides along on
  every render. Verified against a deck with images and speaker notes, not
  just a bare theme file — those get dropped along with their slides.

### `qrack.py` (CLI)

`python qrack.py cluster.pdf [-o out.pptx] [--json] [--new CODE:N] [--label STR]
[--hide-stat KEY] [--list-stats] [--template FILE.pptx]`
— `--json` dumps parsed config; `--new AH-96T:2` overrides the new-node guess;
`--template` appends the rack slide to an existing deck and picks up its
theme colors (see `render_rack`'s `template_path` above); a bad/missing
template path is caught and re-raised as a clean `SystemExit`, not a raw
`pptx` traceback.

### `derive_template.py` (CLI)

`python derive_template.py corp-deck.pptx [-o corp-template.pptx]` — thin
wrapper around `renderer.derive_template`; same clean-`SystemExit` handling
for a bad input path.

## Domain facts the code encodes (don't rederive these wrong)

- **Rack units:** `ceil(node_height_in / 1.75)`, min 1. Qumulo nodes here are
  1.7 in = 1U, but the field is per-model so mixed-height clusters work. Sanity
  check: `sum(m.ru * m.count) == report.rack_u` (the report's "Rack Space
  Required"). Treat a mismatch as a parse bug.
- **Heterogeneous clusters are normal.** The "All Nodes" section lists one block
  per model, each with its own count/raw/height. Never assume a single model.
- **Cabling:** each node has 2 front-end ports → dual-homed, one link to each of
  the two ToR switches. `backend_ports` is 0 on current configs; if it ever is
  non-zero that implies a *second* switch pair (back-end fabric) — not yet
  drawn.
- **New-vs-existing is a GUESS.** On an expansion the PDF gives the added-node
  count and the model breakdown, but not definitively which model is new
  (arithmetic is ambiguous). `_flag_new_nodes` infers it from the report's
  **internal** document name (footer, e.g. `…-QV-1UHG2-L2-96TB-…` → the 96 TB
  model). This is surfaced, not hidden — the CLI prints the guess and `--new`
  overrides it. **The web UI must let the user confirm/override this too.**
- **The stats panel is not a fixed list, by design.** `parser._extract_catchall_stats`
  scans the page-1 summary tables (bounded by the stable "Capacity
  Performance" / "Capacity Planning" section headers — not by field names)
  for anything shaped like `Label Value+Unit` that isn't already one of the
  named fields, and stores it in `report.extra_stats`. `renderer._stat_rows`
  surfaces those under an "Other" section alongside the named ones, so
  `available_stats()` (what the web checklist and `qrack.py --list-stats`
  both read from) reflects whatever Qumulo's tool actually put in *this*
  report, not just what this codebase has a name for. This is explicitly a
  hedge against the sizing tool adding/renaming metrics over time — see the
  reasoning and its limits in `_CATCHALL_RE`'s docstring-comment: it catches
  a new *numeric* metric added to the *existing* tables; it won't catch a
  restructured layout or a new non-numeric status field. When you do add a
  proper named field for something (recommended once a catch-all entry
  looks permanent, since a named field gets a clean key/formatting instead
  of the PDF's raw text), add its label text to `_KNOWN_STAT_LABELS` too, or
  it'll double up in `extra_stats`.
- Use the report's *internal* name for inference, never the upload filename
  (users rename uploads).

## Dev workflow

```bash
pip install pdfplumber python-pptx
python qrack.py samples/some-cluster.pdf --json      # inspect the parse
python qrack.py samples/some-cluster.pdf             # produce the slide
```

Validating a generated deck (no PowerPoint needed):

```bash
python -c "from pptx import Presentation; Presentation('out.pptx')"   # opens clean?
# visual check, if LibreOffice is available:
soffice --headless --convert-to pdf out.pptx && pdftoppm -jpeg -r 130 out.pdf slide
```

When you touch the parser or renderer, re-run **both** a homogeneous report and
a heterogeneous *expansion* report — they exercise different code paths (single
vs multiple model blocks, new-node highlighting). Add any regression PDFs to
`samples/`.

## Web front-end (built)

Upload a sizing PDF → confirm the new-vs-existing split → download the
`.pptx`. `web/app.py` (FastAPI) reuses `parse_report` / `render_rack`
unchanged; `web/static/index.html` is a single-file, no-build-step vanilla
HTML/JS page (no SPA framework).

Four endpoints, because the new-node ambiguity makes a confirm step worth
having, previewing shouldn't require a download first, and template
distillation shouldn't require a round-trip through the CLI:

- `POST /api/parse` (multipart PDF) → `200` with `{report: report.as_dict(),
  available_stats: [...]}` (see `renderer.available_stats` — every stat this
  report could show, for a selection checklist), or `422` with a clear
  message when it isn't a parseable Qumulo report.
- `POST /api/render` and `POST /api/preview` take the same body — `{report`
  (possibly edited by the user), `rack_label`, `visible_stats`,
  `template_base64}` (the last is optional — a `.pptx`, base64-encoded, to
  append the rack slide to; omit or `null` for the default styling) — and
  differ only in what they stream back: `/api/render` streams the `.pptx`
  with `Content-Disposition: attachment`; `/api/preview` renders that *same*
  `.pptx` to a PNG and streams that instead, so the preview can never drift
  from the real output the way a from-scratch HTML/CSS redraw of the layout
  could. Costs a few seconds and one-to-two external processes per request,
  and pulls `libreoffice-impress` + `poppler-utils` into the Docker image
  (~500MB) — worth it for the fidelity guarantee on a low-traffic internal
  tool; reconsider if that trade-off ever stops making sense (e.g. a
  from-scratch canvas redraw if preview latency/image size become the
  actual complaint).
  - `_resolve_template()` decodes/validates the base64 (size cap, zip magic
    bytes `PK\x03\x04`) into a temp file and yields its path (or `None`);
    template-caused render failures come back as `422` rather than `500`.
- `POST /api/derive-template` takes `{template_base64}` and returns
  `{template_base64: <stripped>}` — the web equivalent of
  `derive_template.py` (same underlying function), for the UI's "style
  only" checkbox (see below): strip an uploaded deck down to just its
  theme/layouts before it's used or persisted, so the browser never has to
  hold onto (or `localStorage`-persist) the original full deck.
  - **`/api/preview`'s rasterization is a two-step pipeline, not a single
    `soffice --convert-to png`:** our rack slide is always the *last* slide
    in the deck (appended after any template slides), but `soffice`'s PNG
    export filter only ever rasterizes slide/page 1 of a multi-page
    conversion — there's no filter option to target another page. So preview
    converts to PDF first (which renders every page), then uses
    `pdftoppm -png -r 180 -f N -l N -singlefile` to pull out page `N =
    len(Presentation(pptx_path).slides)` specifically. 180 DPI on the fixed
    13.333"×7.5" slide is what yields exactly `PREVIEW_WIDTH_PX` ×
    `PREVIEW_HEIGHT_PX` (2400×1350). This runs unconditionally (not just
    when a template is given) for simplicity — it happens to be a no-op
    difference when there's no template, since page 1 already is our slide.

The UI shows the parsed config, lets the user fix the highlighted-as-new
selection (and optionally the rack label, stat selection, and a template),
then calls preview and/or render. The template picker has a "style only"
checkbox (checked by default) that, when a file is chosen, first round-trips
it through `/api/derive-template` before storing/using it — so by default
the browser only ever persists the small distilled file, not the original
branded deck. The checkbox is read at file-selection time only (toggling it
afterward needs re-choosing the file, since the raw upload isn't kept
around once processed).

`ClusterReport.from_dict()` / `NodeModel.from_dict()` (parser.py) rebuild the
dataclasses from that edited JSON; `parse_report(source, *, name=None)`
accepts a path or a file-like object, and `name` (when given) always wins for
`source_file` — even when `source` is itself a path, since the web app saves
uploads to a randomly-named temp file and needs the *original* upload name to
survive for the download filename.

Implementation notes (current state):

- FastAPI + uvicorn; single container serves both the JSON API and the static
  page (see README for the current `docker run` invocation — HTTP + HTTPS
  ports, cert mounts). No auth — see "Not yet built" if that changes.
- Each render writes to its own `NamedTemporaryFile`, streams it back via
  `BackgroundTask(os.unlink, ...)` so it's deleted right after the response
  completes; each preview does the same with a `TemporaryDirectory` (needs a
  directory, not just a file, for its own LibreOffice profile — see
  `api_preview`). Never a shared/static path.
- Upload validation: 20 MB cap, `%PDF-` magic-byte check, content-type check;
  after parse, 422 if `report.models` is empty or `report.usable_tb is None`.
  The upload filename is never trusted for anything but the display label
  passed to `parse_report(..., name=...)`.

## Conventions & gotchas

- **`python-pptx`, not `pptxgenjs`.** Chosen deliberately: Python matches the
  rest of the toolchain and the output is native editable shapes. Stay in
  `python-pptx`.
- **Colours are `RGBColor(r, g, b)`** in the renderer — no `#`, no strings.
- **Turn off shape shadow inheritance** (`shape.shadow.inherit = False`) on new
  shapes, as the existing helpers do, or you'll get stray drop shadows.
- Text boxes have built-in padding; the helpers zero the margins where things
  must align. Reuse `_rect`, `_text`, `_label_in_shape`, `_connector` rather
  than hand-rolling shapes.
- Node ordering: `_node_sequence` floats a model's new nodes to the top of its
  group. Change ordering there, in one place.
- **Never do float-producing arithmetic on `Emu`/`Length` values headed for
  `_connector()`.** `Emu`/`Inches`/`Pt` are plain `int` subclasses with no
  arithmetic overrides, so `/2`, `* 0.5`, or any division on one produces a
  bare Python `float` (true division always returns `float` in Python 3,
  even `int / int`). `add_shape`/`add_textbox` silently sanitize a float
  position back to `int` internally, but `python-pptx`'s `add_connector()`
  computes its bounding box via raw `min()`/`abs()` on the endpoints with
  **no such sanitization** — a stray float lands in the XML as
  `<a:off y="1435608.0"/>` instead of `"1435608"`. That's schema-valid and
  LibreOffice/`python-pptx`/even ISO-29500 XSD validation all accept it
  fine, but real PowerPoint's parser rejects it outright as corrupt
  ("needs repair"), and Google Slides refuses to open it at all. `_connector()`
  already guards against this by coercing all four coordinates through
  `Emu(round(v))` before calling `add_connector()` — keep that guard if you
  ever touch that helper, and don't reintroduce raw float math on the call
  sites as a "cleanup."

## Not yet built (good next tasks, roughly in order)

- Multi-rack splitting when `total_node_ru` exceeds one rack's height (spill
  into a second rack column on the same slide).
- Back-end switch pair when `backend_ports > 0`.
- Auth, if the web app ever needs to leave a trusted network (currently none
  by design — see the "Docker shape" decision in project history).
