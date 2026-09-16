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
  visible_stats=None, template_path: str|None=None,
  template_slide: int|None=None, rack_sizes: list[int]|None=None) -> str`
  writes the `.pptx` to `out_path` and returns it.
- Palette is module-level constants at the top of the file (node / switch /
  cable colours). Centralize any theming there.
- Slide is fixed 13.333" × 7.5" (16:9). All shapes are native `python-pptx`
  rectangles / connectors / text boxes, so the deck is hand-editable after
  generation — do not flatten anything to an image.
- **`template_path`**: when given, `Presentation(template_path)` is used as
  the base deck instead of a blank one — the rack slide is *appended* after
  whatever slides the template already has, and picks up colors sampled
  from that template's *real content* (`_sample_deck_colors()`) rather
  than its declared theme scheme. This was a deliberate pivot away from
  reading `_theme_colors()` (the theme's abstract `<a:clrScheme>`) as the
  primary source: a real branded deck, especially a Google Slides export,
  does not reliably paint its actual slides with its own declared theme
  colors. Confirmed firsthand on a real customer deck whose theme scheme
  said "navy text on white," while 6 of its 8 slides were actually solid
  black backgrounds with white text and a teal/amber accent — the theme
  colors were essentially decorative metadata nobody used. `_theme_colors()`
  still exists and is used as the *fallback*, in two ways: (1) within
  `_sample_deck_colors()` itself, to resolve a `schemeClr` reference found
  on a real shape/run to its actual RGB, and (2) in `render_rack()`, for
  whichever of background/text/accent sampling couldn't determine (e.g. a
  template already stripped down to zero slides by `derive_template`,
  which has nothing left to sample — see below for how that case is
  handled). Functional colours (new-node green, Switch A/B cable colours)
  stay fixed regardless of template, since they carry meaning.
  Secondary/muted text (rack label, node model-code labels, subtitle, cable
  legend) is *not* read from the sample directly — it's derived by
  blending the sampled text color toward the sampled background
  (`_blend(title_color, bg_color, 0.35)`), since the module's own default
  colors for those (`SUBTITLE_TEXT`, `RACK_LABEL_TEXT`, `NODE_LABEL_TEXT`)
  assume a white background and would be nearly invisible against a
  legitimately dark one (also caught firsthand, from the same deck above).
  `_find_blank_layout()`
  picks the emptiest layout in the template (preferring one literally named
  "blank") by counting non-content placeholders (date/footer/slide-number
  don't count) — there's no schema flag for "this is the blank layout," so
  this is a heuristic, not a guarantee, on an unusual template. A real
  branded deck (as opposed to python-pptx's own default template, which has
  a genuine zero-placeholder "Blank" layout) commonly has *no* truly empty
  layout at all, so even the least-bad one picked here can still carry a
  few content placeholders (title/subtitle/body). `add_slide(layout)`
  clones those onto our new slide — normal python-pptx behavior for
  someone about to type into them, but we never do; every field on this
  slide is our own explicit shape. Left alone, those empty placeholders
  don't show up in a flattened render (PDF/PNG conversion, or `/api/preview`)
  since they carry no visible content, but they absolutely show up the
  moment the file is opened for editing in PowerPoint or Google Slides —
  empty "click to add title/text" boxes sitting on top of or behind the
  rack diagram and stats, right where the layout positioned them. Fixed by
  stripping every inherited placeholder off the new slide immediately
  after `add_slide()` (`for ph in list(slide.placeholders):
  ph._element.getparent().remove(ph._element)`) — keep that in place if you
  ever touch the slide-creation code, and don't assume a LibreOffice/PDF
  screenshot proves a template is clean, the same lesson as the earlier
  EMU-float saga above. Colors are resolved into local variables and
  threaded through as function params, never mutated on shared module state,
  since the web app can render
  concurrently.
- **Known limitation:** the slide is always fixed to 13.333"×7.5" (16:9), so
  a 4:3 template gets its aspect ratio silently overridden. Not worth
  handling until someone actually hits it.
- **Multi-rack layout** (`rack_sizes`): the sizing PDF gives a cluster's
  total node count and total rack-U, never how they're physically split
  across racks — that's purely a rendering decision, made by
  `_split_into_racks(seq, rack_sizes=None)`:
  - **Auto-split** (`rack_sizes=None`, the default): fills a rack by
    cumulative RU up to `RACK_NODE_CAPACITY_U` (`RACK_TOTAL_U` minus
    `RACK_SWITCH_RESERVED_U` — the report's own node/rack-U numbers never
    include the ToR switches, but they take real physical space; assumes
    1U each, a common ToR form factor) before spilling into the next rack.
    The overwhelmingly common case — everything fits in one rack — yields
    exactly one, so this is a no-op for every report that predates
    multi-rack support.
  - **Manual override** (`rack_sizes=[n, n, ...]`): an explicit node count
    per rack, in the same new-nodes-first order `_node_sequence` already
    produces. Meant to be the user's edit of a previous auto-split result,
    so it's strict: raises `ValueError` if any count is negative or the
    counts don't sum to the report's total node count exactly, rather than
    silently dropping nodes or absorbing a mismatch into the last rack.
  - **Up to 2 racks share one slide** (`_rack_geometry`), scaled down on
    both the rack frame and its label strip (`RACK_W_2UP`/
    `NODE_LABEL_W_2UP`, vs. the single-rack `RACK_W`/`NODE_LABEL_W`) to
    leave the stats panel a usable width. The stats panel always covers
    the *whole* cluster regardless of rack count (matching the source
    report, which has no per-rack breakdown either) — `_draw_stats` gets
    `allow_pair=False` (no side-by-side top pair, and no per-card internal
    two-column row layout either — both are the same "is there width for
    2 side-by-side label/value groups" question) and `compact=True`
    (tighter header height, card spacing, and row-height floor, plus a
    smaller stat font) once there's more than one rack. Forcing every
    card to one column of rows roughly doubles how many lines a report's
    full stat set needs, and without `compact`'s reclaimed overhead a
    *realistic* stat count (not even an extreme everything-selected case)
    overflowed past the slide's bottom edge at the narrower width —
    verified by deliberately constructing that failure before adding
    `compact`, then confirming it's gone after.
  - `_draw_rack` takes an explicit `seq`/`rack_x`/`rack_w`/`label_w`
    instead of a `ClusterReport` and the module's single-rack constants,
    so the same function draws every column regardless of how many racks
    share the slide; it no longer draws the legend (see next). At
    `rack_w` below the single-rack `RACK_W`, "ToR Switch A/B" at the
    normal font wraps to two lines (the switch's own label sub-box is a
    fixed fraction of `rack_w`) — `_draw_rack` drops to a smaller font
    once `rack_w < RACK_W`, not tied specifically to 2-rack mode, so a
    future narrower preset doesn't have to remember to handle this too.
  - `_draw_legend` (New node / Switch A / Switch B) is drawn once per
    slide, not once per rack — every rack shares the same frame height
    (the frame always spans the fixed 42U envelope regardless of content),
    so it doesn't matter which rack's returned frame-bottom the caller
    measures it from.
  - Multiple racks get labeled `"{base_label} — Rack {n}"` (or just
    `"Rack {n}"` with no base label) via `_rack_labels` — imperfect if
    `base_label` was itself already rack-specific (e.g. "Row 3 / Rack 12"
    from single-rack usage), but a free-text single label can't
    unambiguously name multiple racks without the user splitting it
    themselves. At the narrower 2-up `rack_w`, this longer combined label
    can itself be too wide for one line at the normal size — same problem
    as the switch labels above, just caught later because it only bites
    with a longer `base_label` than the "Rack 1"/"Rack 2" defaults used
    while first building this feature. `_draw_rack` drops to a smaller
    font and widens/bottom-anchors the label box whenever `rack_w <
    RACK_W`, so a wrapped second line grows up into the gap under the
    header instead of down into the switch frame.
  - `_rack_geometry` itself only ever has two presets — 1-up (unscaled,
    the original single-rack constants) and 2-up (`RACK_W_2UP` etc.) — and
    raises a clean `ValueError` for any other count; **more than 2 racks
    is handled one level up, in `render_rack`, by repeating the 2-up
    preset across additional slides** rather than teaching `_rack_geometry`
    a third size:
    - `render_rack` chunks `_split_into_racks`'s result into groups of ≤2
      racks (`rack_groups`) and calls `_draw_rack`/`_draw_legend` once per
      group on its own new slide, always requesting `_rack_geometry(2)`
      for every rack-only slide once there's more than one group —
      including a trailing group with only 1 rack — so every rack-only
      slide in a multi-slide deck looks visually consistent rather than
      having the last one suddenly render at the wider 1-up scale.
    - Because the stats panel no longer has to compete with rack columns
      for width once there's more than one slide, the aggregated stats
      (still whole-cluster, matching the source report) move to one
      dedicated, additional final slide instead of squeezing into
      `compact`/single-column form next to the last rack group —
      `allow_pair=True, compact=False` there (the original spacious
      layout), using the slide's full width margin-to-margin.
    - Each rack slide's subtitle gets a page-context suffix —
      `"Racks {start}–{end} of {total}"`, or `"Rack {n} of {total}"` when
      a trailing group has only one rack (avoids the awkward
      "Racks 3–3 of 3") — and the stats slide's is just `"Stats"`; omitted
      entirely for the common single-slide case, so a report that fits in
      one or two racks renders byte-for-byte the same subtitle as before
      this feature existed.
    - Background/text/accent colors and the muted-text blend are all
      computed once, outside this loop, from the template (or defaults),
      and threaded into every slide the same way — a multi-slide render
      against a template looks like the same deck throughout, not a style
      shift partway through.
- `_sample_deck_colors(prs) -> dict` returns up to
  `{'background':, 'text':, 'accent':}` as `RGBColor`, sourced from a
  template's *real* slides, not its theme scheme:
  - `background`: the most common effective background across all slides
    — explicit `<p:bg>`, or (very common on a Google Slides export, which
    routinely fakes a background this way instead of using the real OOXML
    mechanism) a plain shape sized to exactly cover the slide
    (`_full_bleed_fill_color`) — checked **per level**, slide then layout
    then master, each level's full-bleed shape before falling through to
    the *next* level's `<p:bg>`, not all three `<p:bg>`s before any
    full-bleed shape. A full-bleed shape at a level visually covers
    whatever `<p:bg>` that same level declares or inherits, so it has to
    win at that level, not lose to a `<p:bg>` one level up. Got this
    backwards originally and it mattered in practice: a real customer
    deck's master had an unused, literally-never-visible white `<p:bg>`
    that still out-prioritized the *layout's* actual painted navy
    full-bleed background, because "does the master have a `<p:bg>`" was
    checked before "does anything have a full-bleed shape" at all — nearly
    half that deck's slides use a layout that's actually dark navy, not
    the white this originally reported.
  - `text`: the most common color on a title placeholder, per slide — its
    own run color if set, else the layout's default style for that
    placeholder (`_layout_title_default_color`), since a title's *run*
    frequently carries no override at all and the actual styling lives on
    the layout (again, routine on a Google Slides export). Rejected if it
    doesn't contrast against the winning `background` (`_contrasts`, a
    lightness-gap check) — background and text are tallied independently
    across all slides, so without this guard they can land on the *same*
    color when a deck's most-common title color and most-common background
    happen to come from different subsets of slides (seen firsthand: an
    all-white "text" pick on an all-white "background" pick, an invisible
    title, from a deck that both used white titles on a few dark slides
    and was mostly white overall).
  - `accent`: the most common *non-neutral* color (`_is_neutral_color` — a
    strict 0.4 saturation floor, tuned after a too-loose 0.15 let a
    desaturated navy text shade outrank a deck's actual bright highlight
    color) among all text runs and shape fills, excluding whichever color
    already won `background`/`text`.
  Every color goes through `_resolve_color_format`, which handles both a
  literal RGB value and a theme (`schemeClr`) reference — including the
  *semantic* `tx1`/`bg1`/`tx2`/`bg2` slots real content overwhelmingly
  uses, resolved via that slide's own master's `_color_map` (its
  `<p:clrMap>`), never assumed to be a fixed `tx1`→`dk1`/`bg1`→`lt1`
  mapping, since a dark-background master can invert it. This is all
  still a heuristic, like `_find_blank_layout()` — "most common" is a
  proxy for "what a viewer actually associates with this deck," not a
  guarantee, and a deck that genuinely uses two accent colors about
  equally will pick whichever the `Counter` happens to see first on a tie.
  **Scope, explicitly**: this is 3 flat colors, nothing else — no logos or
  other pictures, no footer/page-number text, no decorative shapes, no
  fonts. "Match this slide" means "use its background/title/accent
  colors," not "reproduce its design" — confirmed against a real template
  slide literally named "Slide Option - Dark with Logo," where the fixed
  version correctly picks up its navy background and white title but
  obviously doesn't carry over its logo, its rounded-rectangle graphic
  pattern, or its footer. Reproducing any of those would be new,
  separately-scoped work, not a bug in what's built today.

  Gradients, on the other hand, *are* in scope, at reduced fidelity: any
  fill (background `<p:bg>`, a full-bleed shape, an accent shape)
  resolves through `_resolve_fill_color`, which handles a `SOLID` fill
  directly and reads a `GRADIENT` fill's first color stop (position 0) as
  a flat stand-in — a genuine gradient can't be reproduced, but a color
  read off one end is far closer than treating the shape as if it had no
  fill at all. That "as if no fill" behavior was the actual bug on a real
  slide (a section-header layout with a bright-blue-to-navy gradient
  background): the sampler only ever checked for `MSO_FILL_TYPE.SOLID`,
  so a gradient-filled full-bleed shape was silently skipped, falling
  through all the way to an unrelated, never-actually-visible master
  background. Reported as "I tried slide 18 ... I didn't see a
  difference" (the fallback happened to look close to the deck-wide
  default). Every caller that reads a fill's color goes through
  `_resolve_fill_color` now, not a direct `SOLID` check.
  Takes an optional `slide_index` (1-based) to sample a single specific
  slide instead of majority-voting across the deck — useful for a deck
  that genuinely has more than one distinct look (confirmed on a real
  customer deck: most slides teal-accented, a few amber, one navy — the
  deck-wide vote picks whichever is most common, but `slide_index` lets
  you pin an exact one instead). Raises `ValueError` for an out-of-range
  index, including against a deck with zero slides (e.g. one already
  stripped by `derive_template` — there's nothing left to sample against,
  so the slide choice has to be made *before* stripping, either by passing
  `slide_index` to `derive_template` itself or, in the web app, when
  `template_base64` still carries the deck's original slides).
- `derive_template(source_path: str, out_path: str, slide_index: int | None = None) -> str` strips every
  slide out of an existing `.pptx` via the standard python-pptx recipe
  (remove each `<p:sldId>` from `prs.slides._sldIdLst` and `drop_rel` its
  relationship — there's no public "delete slide" API), leaving only the
  masters/layouts/theme. Meant as a one-time distillation step: point it at
  someone's real 40-slide branded deck once, get back a small style-only
  file, use *that* as `template_path` from then on instead of carrying the
  original deck's slides along on every render. Verified against a deck
  with images and speaker notes, not just a bare theme file — those get
  dropped along with their slides. Before stripping, it also runs
  `_sample_deck_colors()` on the *original* slides and bakes the result
  into the saved file's own `<a:clrScheme>` (`_patch_theme_colors`,
  overwriting `lt1`/`dk1`/`accent1`) — since those slides are about to be
  gone, this is the only chance to capture them, and it's what makes a
  later plain `_theme_colors()` read of the *distilled* file (all
  `render_rack` can do once there's nothing left to sample) still reflect
  the deck's actual visual style. `_patch_theme_colors` reassigns the
  theme part's `.blob` directly rather than mutating an `.element` tree —
  the theme part loads as a generic (non-XML-aware) python-pptx `Part`,
  which has no live element to edit in place; `.blob` reassignment is the
  actual persistence mechanism for that part type.

### `qrack.py` (CLI)

`python qrack.py cluster.pdf [-o out.pptx] [--json] [--new CODE:N] [--label STR]
[--hide-stat KEY] [--list-stats] [--template FILE.pptx] [--template-slide N]
[--rack-sizes N,N,...]`
— `--json` dumps parsed config; `--new AH-96T:2` overrides the new-node guess;
`--template` appends the rack slide to an existing deck and picks up colors
sampled from its real content (see `render_rack`'s `template_path` above);
`--template-slide N` samples from just that 1-based slide instead of the
whole deck (requires `--template`; a plain `SystemExit` if given without
it); `--rack-sizes 15,35` overrides the auto-split node count per rack
(see `render_rack`'s `rack_sizes` / "Multi-rack layout" above; must sum to
the report's total node count) — a bad/missing template path, a malformed
`--rack-sizes` list, or a `ValueError` from `render_rack` itself (an
out-of-range `--template-slide`, mismatched `--rack-sizes`, or a report
needing more than 2 racks) is caught and re-raised as a clean `SystemExit`
using that error's own message, not a raw traceback.

### `derive_template.py` (CLI)

`python derive_template.py corp-deck.pptx [-o corp-template.pptx] [--slide N]`
— thin wrapper around `renderer.derive_template`; `--slide N` samples from
just that 1-based slide instead of the whole deck; same clean-`SystemExit`
handling for a bad input path or an out-of-range `--slide`.

## Domain facts the code encodes (don't rederive these wrong)

- **Rack units:** `ceil(node_height_in / 1.75)`, min 1. Qumulo nodes here are
  1.7 in = 1U, but the field is per-model so mixed-height clusters work. Sanity
  check: `sum(m.ru * m.count) == report.rack_u` (the report's "Rack Space
  Required"). Treat a mismatch as a parse bug. This `.ru` value alone (no
  separate lookup or config) is also what picks a node's chassis art in
  `renderer._draw_rack`: `ru == 1` or `ru == 2` gets a bundled product
  photo (`NODE_PHOTO_PATH`/`NODE_PHOTO_2U_PATH`, `qumulo_rack/assets/`),
  anything else falls back to the drawn vector server icon
  (`_draw_server`). Each bundled photo is a specific real model's chassis
  used generically for *any* node of that height, not matched per model
  code — by explicit product decision (not worried about being exact here,
  a photo of the right size reads better than a vector icon of the wrong
  one). Adding a new height's photo means: drop the asset in
  `qumulo_rack/assets/`, compute its crop fractions from its alpha
  bounding box the same way the existing ones were (see the comments
  above `NODE_PHOTO_CROP`/`NODE_PHOTO_2U_CROP` in `renderer.py`), and add
  one more branch next to the `ru == 2` one.
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

Five endpoints, because the new-node ambiguity makes a confirm step worth
having, previewing shouldn't require a download first, template
distillation shouldn't require a round-trip through the CLI, and the
rack-split editor shouldn't have to duplicate the auto-split algorithm in
JS:

- `POST /api/parse` (multipart PDF) → `200` with `{report: report.as_dict(),
  available_stats: [...]}` (see `renderer.available_stats` — every stat this
  report could show, for a selection checklist), or `422` with a clear
  message when it isn't a parseable Qumulo report.
- `POST /api/rack-split` takes `{report}` and returns
  `{rack_sizes: [n, ...]}` — `renderer.auto_rack_split`'s suggested
  node-count-per-rack split (always `[total node count]`, one rack, for a
  cluster that fits in one), for the confirm step's editable rack-split
  fields to start from. Called once right after `/api/parse` resolves,
  since the total node count it depends on never changes from anything
  else editable in the confirm step (only `new_count`, not `count`, is
  user-editable there).
- `POST /api/render` and `POST /api/preview` take the same body — `{report`
  (possibly edited by the user), `rack_label`, `visible_stats`,
  `template_base64`, `template_slide`, `rack_sizes`, `preview_slide}` (the
  last four are optional — a `.pptx`, base64-encoded, to append the rack
  slide to; `template_slide` a 1-based slide number to sample from instead
  of the whole deck, only meaningful when `template_base64` still has its
  original slides, since sampling has nothing to work with once they're
  stripped; `rack_sizes` an explicit node count per rack, see
  `render_rack`'s `rack_sizes`; `preview_slide` a 1-based index into just
  the slides *this render itself* added (not the whole deck, and not
  meaningful to `/api/render`, which always returns everything) — omit or
  `null` any of the four for the default behavior) — and differ only in
  what they stream back:
  `/api/render` streams the `.pptx`
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
- `POST /api/derive-template` takes `{template_base64, template_slide}`
  (the second optional, same meaning as above) and returns
  `{template_base64: <stripped>}` — the web equivalent of
  `derive_template.py` (same underlying function), for the UI's "style
  only" checkbox (see below): strip an uploaded deck down to just its
  theme/layouts before it's used or persisted, so the browser never has to
  hold onto (or `localStorage`-persist) the original full deck. An
  out-of-range `template_slide` comes back as a `422` with
  `_sample_deck_colors`'s own `ValueError` message.
  - **`/api/preview`'s rasterization is a two-step pipeline, not a single
    `soffice --convert-to png`:** `soffice`'s PNG export filter only ever
    rasterizes slide/page 1 of a multi-page conversion — there's no filter
    option to target another page. So preview converts to PDF first (which
    renders every page), then uses `pdftoppm -png -r 180 -f N -l N
    -singlefile` to pull out page `N` specifically. 180 DPI on the fixed
    13.333"×7.5" slide is what yields exactly `PREVIEW_WIDTH_PX` ×
    `PREVIEW_HEIGHT_PX` (2400×1350).
  - `N` is `template_slide_count + preview_slide` — `preview_slide`
    (1-based, default `1`) indexing into just our own added slides, not
    the deck's last — where `template_slide_count` is read from the
    template file itself (`len(Presentation(template_path).slides)`,
    before `render_rack` appends anything) and is `0` with no template.
    This used to unconditionally target `len(Presentation(pptx_path).slides)`
    (the whole deck's last slide), which was equivalent back when
    `render_rack` only ever appended exactly one slide — but a multi-rack
    cluster can now add several (see "Multi-rack layout" above), and its
    *last* one is the dedicated stats slide. Always targeting the last
    slide would have silently previewed only the stats panel for exactly
    the reports where a visual sanity check of the rack diagram matters
    most, and defaulting `preview_slide` to `1` keeps that fixed: the
    first rack slide, not the last slide overall, is what a caller sees
    with no `preview_slide` given. Reachable at all past slide 1 is a real
    need, not just tidiness — reported firsthand against a genuine 3-rack
    cluster ("I modified the rack count... and I'm only seeing 2 racks on
    the preview"): the render itself was correct (all 3 racks were in the
    downloaded `.pptx`), but before `preview_slide` existed, `/api/preview`
    had no way to show anything past the first rack slide at all, so the
    third rack and the aggregated stats were both invisible until download.
    `total_our_slides` (`len(Presentation(pptx_path).slides) -
    template_slide_count`, computed once the render is done since it
    depends on how many racks/slides *this particular* report and
    rack-split actually needed) is what bounds `preview_slide` — an
    out-of-range value is a clean `422` naming the valid range, not a
    silent clamp, so a stale page number left over from a bigger split
    doesn't quietly show the wrong slide. Both values also come back as
    `X-Slide-Index`/`X-Slide-Count` response headers (added to
    `expose_headers` on the CORS middleware, since this API allows
    cross-origin callers) so the caller can build pager UI without
    duplicating the slide-count math — see the web UI's
    `#preview-pager` below. Runs unconditionally (not just when a template
    is given) for simplicity — it happens to be a no-op difference for a
    single-slide, no-template render, since page 1 already is our slide.
  - The web UI's `#preview-pager` (Prev / "Slide N of M" / Next, next to
    the preview image) reads exactly those two headers on every preview
    response and hides itself whenever `X-Slide-Count` is `1` — the
    overwhelmingly common case, so nothing changes there. `currentPreviewSlide`
    (the page currently shown, sent as `preview_slide`) resets to `1`
    whenever the next render could plausibly produce a different slide
    count than the last one shown — a fresh report, an edited rack-split
    field, or "Reset to auto-split" — so a stale page index from a larger
    previous split can't end up requesting a page that no longer exists.
    It deliberately does *not* reset on every `scheduleAutoPreview()` call
    generally (e.g. a rack-label edit or a stat-visibility toggle, neither
    of which can change the slide count) — those keep whatever page the
    user was already looking at, since re-centering them back to slide 1
    on every keystroke would make paging forward and then tweaking the
    label pointlessly throw away where they were.

The UI shows the parsed config, lets the user fix the highlighted-as-new
selection (and optionally the rack label, stat selection, rack split, and a
template), then calls preview and/or render.

**Rack split**: right after `renderConfirm` sets `currentReport`, it fires
`/api/rack-split` (not awaited — a self-contained async call that fills in
its own UI once it resolves, nothing else in `renderConfirm` depends on
it) and hides the `#rack-split-section` block by default. If the response
has more than one rack, that section un-hides with one number input per
rack (`data-rack-idx="0"`, `"1"`, ...), pre-filled with the suggested
split, plus a "Reset to auto-split" button that restores those same
values. A cluster that fits in one rack never shows this section at all —
`buildPayload()` sends `rack_sizes: null` whenever it's hidden, so nothing
about the request changes for the common case. `validateRackSplit()`
(non-negative integers, summing to the report's total node count) runs
client-side on every edit and before every Preview/Generate click —
gating `scheduleAutoPreview()` itself (not each call site separately)
means every path that could trigger a render (rack label edits, stat
toggles, the rack-split fields themselves) automatically respects it —
so a typo is caught immediately with an inline message instead of costing
a round-trip to the server's own `rack_sizes` validation in
`_split_into_racks`.

The template picker has a "style only"
checkbox (checked by default) that, when a file is chosen, first round-trips
it through `/api/derive-template` before storing/using it — so by default
the browser only ever persists the small distilled file, not the original
branded deck — plus an optional "Match slide #" number field for a deck
with more than one distinct look. The checkbox is read at file-selection
time only (changing it afterward needs re-choosing the file, since the raw
upload isn't kept around once processed). What happens with the slide
number depends on the checkbox:
- **Style only checked**: same as the checkbox — read once at
  file-selection time and sent to that one `/api/derive-template` call.
  `currentTemplateSlide` stays set to whatever was used, though — it
  reflects "what's baked into the *current* distilled template," which is
  still true after distilling, not just before. (An earlier version reset
  it to `null` here on the theory that it's "spent" and irrelevant to
  later requests, which is correct for what gets *sent*, but
  `updateTemplateUi()` also used that same variable to decide what the
  field displays — so the field reset to blank right after a successful
  distill that *did* use the slide number, indistinguishable from the
  number having been silently ignored. Reported exactly that way: "it
  sets the slide # back to auto and I don't see any effect." Fixed by
  decoupling the two concerns — `currentTemplateSlide` always reflects
  the truth for display, `currentTemplateIsRaw` alone decides whether
  `buildPayload()` resends it.) Editing the field after a distill doesn't
  change anything about the current template — still inert until a new
  file is chosen — but the file-choice handler itself now also sets an
  explicit `Distilled using slide N.` status, since a field that already
  shows the right number isn't enough confirmation on its own that a
  fresh distill actually used it (rather than, say, silently keeping
  a stale value from before).
- **Style only unchecked** (raw mode, original deck kept): the slide
  number has to be resent on *every* `/api/render`/`/api/preview` call
  alongside the raw `template_base64`, since sampling happens fresh each
  time against the full deck (`currentTemplateIsRaw` tracks this so
  `buildPayload()` knows whether to include it) — and because that full
  deck is still sitting in the browser, an `input` listener on the field
  updates `currentTemplateSlide` and fires the debounced auto-preview live,
  no re-upload needed, once a preview is already showing.

  An edit to this field is otherwise invisible — no page motion, and the
  preview card it eventually updates can be well below the fold — so a
  dedicated `#template-slide-status` span next to the field gives an
  immediate, mode-aware confirmation on every `input` event: the
  "applies next time" explanation in style-only mode, "will apply once
  you click Preview" if no preview has been shown yet, an immediate
  "Updating preview…" otherwise, and a final "Preview updated for slide N"
  once it lands (or the failure message, if it didn't). That confirmation
  comes from `scheduleAutoPreview()`'s optional `onSettled(ok, err)`
  callback, added for exactly this — **never pass `scheduleAutoPreview`
  itself as a bare event listener** (`el.addEventListener('input',
  scheduleAutoPreview)`) now that it takes a parameter: the DOM would call
  it with the `Event` object as `onSettled`, and it crashes the instant
  that branch tries to invoke a non-function. Wrap it
  (`() => scheduleAutoPreview()`) at every call site, as the existing ones
  already do — caught by the jsdom regression suite (`test3.js`) the one
  time this slipped through here.

**`localStorage` schema versioning** (`TEMPLATE_STORAGE_VERSION`,
alongside `TEMPLATE_STORAGE_KEY`): the bug above had a second-order
consequence worth guarding against on its own. A browser that had saved a
template under the *old* buggy logic (correct bytes, but a `slide` field
wrongly reset to `null`) would silently reload that old, now-incorrect
combination as if it matched today's logic — because nothing about the
stored shape changed, only what a given value was supposed to *mean*. That's
strictly worse than losing the saved template: it looks like it's working
(a template is loaded, a name is shown) while actually producing the wrong
colors, which is exactly what got reported as "I don't see any effect."
`saveStoredTemplate()` now stamps every write with `TEMPLATE_STORAGE_VERSION`;
`loadStoredTemplate()` discards (and removes) anything that doesn't match,
rather than trying to interpret it. **Bump this version** whenever a change
alters what an existing stored field means (not just when a field is
added) — the cost of bumping unnecessarily is one harmless re-select; the
cost of not bumping when it mattered is a silent, hard-to-diagnose mismatch
between what's displayed and what's actually being sent.

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

- Back-end switch pair when `backend_ports > 0`.
- Auth, if the web app ever needs to leave a trusted network (currently none
  by design — see the "Docker shape" decision in project history).
