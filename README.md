# qrack — Qumulo sizing PDF → editable rack-diagram slide

Turns a `sizing.qumulo.com` cluster report (PDF) into a single editable
PowerPoint slide: a rack elevation (2 ToR switches + nodes stacked by RU,
new nodes highlighted, colour-coded cabling) alongside a stats panel
(capacity, performance, rack & power).

## Design: split extraction from rendering

The value is in keeping these two concerns separate:

- **`qumulo_rack/parser.py`** — *deterministic*, format-targeted. The sizing
  tool's output is stable, so it's parsed with `pdfplumber` + regex, no
  LLM/fuzzy reading required. Returns a `ClusterReport` dataclass.
- **`qumulo_rack/renderer.py`** — *deterministic geometry*. Same config → same
  slide, every time. That's what makes the output "standardized." Emits native
  `python-pptx` shapes (rectangles, connectors, text boxes), so the slide is
  fully editable in PowerPoint afterwards.

Because the core is a plain module, the same two functions back a CLI today and
can back a Claude skill or a web endpoint later — build the renderer once,
expose it however you like.

## Usage

```bash
python qrack.py cluster.pdf                 # -> cluster.rack.pptx
python qrack.py cluster.pdf -o out.pptx
python qrack.py cluster.pdf --json          # dump parsed config, no slide
python qrack.py cluster.pdf --new AH-96T:2  # override the new-node guess
python qrack.py cluster.pdf --label "Row 3 / Rack 12"
python qrack.py cluster.pdf --list-stats    # list this report's stat keys, no slide
python qrack.py cluster.pdf --hide-stat iops --hide-stat encoding
python qrack.py cluster.pdf --template corp-deck.pptx  # append to an existing deck, matching its real colors
python derive_template.py corp-deck.pptx               # -> corp-deck.template.pptx, no slides, same colors
python derive_template.py corp-deck.pptx --slide 5      # match slide 5 specifically, not the whole deck
python qrack.py cluster.pdf --rack-sizes 15,35          # node counts per rack, instead of auto-split
```

If you just want the rack slide styled like your company deck -- not literally
inserted into it -- run `derive_template.py` once to strip that deck down to
its theme/layouts (no slides, so it's small), and pass *that* file as
`--template` instead. Every `qrack.py` render then produces a single-slide
`.pptx` with the matching colors, rather than your whole original deck plus
one slide.

Some decks look different from slide to slide (a dark title slide, lighter
content slides, a differently-accented section divider, ...). By default,
colors are picked by majority vote across the whole deck; `--slide N` (on
either command) pins it to one specific 1-based slide instead. On
`derive_template.py` that slide choice has to be made up front, since the
distilled file has no slides left afterward for `qrack.py --template` to
sample from later.

A cluster too big for one 42U rack automatically splits across racks
(scaled down to fit). Up to 2 share one slide with the stats panel alongside
them; 3 or more move the stats to their own dedicated final slide instead,
which frees the whole slide width for racks -- up to 3 then share one
rack-only slide (wider per rack than the 2-up case, since nothing else
needs a share of that width), spilling onto additional rack-only slides
only for a cluster needing more than that. `--rack-sizes` overrides the auto-split with
your own node count per rack, e.g. `--rack-sizes 15,15,20` -- the counts must
add up to the report's total node count.

Dependencies: `pip install pdfplumber python-pptx`.

## What it derives

| Field | Source | Notes |
|---|---|---|
| Node models, counts | "All Nodes" section | handles heterogeneous clusters |
| RU per node | `ceil(Height / 1.75)` | cross-checked vs "Rack Space Required" |
| Cabling | Front-end ports (2/node) | each node dual-homed to both ToR switches |
| New vs existing | report's internal name (`…-96TB-…`) | **a guess** — override with `--new` |
| Capacity / perf / power | page-1 summary tables | |

## The one ambiguity worth knowing about

For an **expansion**, the PDF says how many nodes were added and the model
breakdown, but not definitively *which* model is new. The parser infers it from
the report's internal document name (which encodes the added node's size) and
**surfaces the guess** so you can confirm or override with `--new CODE:N`. This
is exactly the kind of decision a human should confirm — hence the override.

## Web app (Docker)

Upload a PDF, confirm/edit the new-vs-existing node split, download the
`.pptx`. FastAPI backend + a thin static HTML/JS front-end, both wrapping the
same `parse_report` / `render_rack` used by the CLI.

### Setup

The container serves plain HTTP on `8000` always, and additionally serves
HTTPS on `8443` if a cert/key pair is mounted at `/certs` — that's what
avoids browsers blocking the `.pptx` download as an "insecure download"
(Chrome/Brave-family browsers do this for anything fetched over plain HTTP).

**Generate a local certificate once** (needs [mkcert](https://github.com/FiloSottile/mkcert)):

```bash
mkdir -p certs
CAROOT=certs/mkcert-ca mkcert -cert-file certs/qrack-cert.pem -key-file certs/qrack-key.pem \
  <this-host's-LAN-IP> <this-host's-hostname> localhost 127.0.0.1
```

That single command creates the CA under `certs/mkcert-ca/` (a `rootCA.pem` /
`rootCA-key.pem` pair) the first time it runs, without touching *this* host's
own trust store — pointing `CAROOT` at a project-local directory keeps it
from installing anywhere system-wide, which we don't need since nobody
browses from this host.

`certs/` is gitignored and dockerignored — the private key never gets
committed or baked into the image; it's mounted in at `docker run` time.

```bash
docker build -t qrack .
docker run -d --name qrack \
  -p 8080:8000 -p 443:8443 \
  -v "$(pwd)/certs/qrack-cert.pem:/certs/qrack-cert.pem:ro" \
  -v "$(pwd)/certs/qrack-key.pem:/certs/qrack-key.pem:ro" \
  --restart unless-stopped qrack
```

- `-d --name qrack` runs it detached and gives it a name so you can
  `docker stop qrack` / `docker logs qrack` later.
- `--restart unless-stopped` brings it back after a host reboot or Docker
  restart.
- **HTTPS is mapped to the standard port 443, not a custom port.** A custom
  HTTPS port (e.g. `8443:8443`) can get silently dropped by corporate
  firewalls/VPNs that only allow the well-known web ports outbound — 443
  virtually always gets through. HTTP still uses a custom port (`8080` here)
  since that one's just a fallback, not the primary path.
- Change the left-hand number in `-p 8080:8000` to avoid conflicts with
  whatever else is running on this host — the right-hand container port
  never needs to change.

### Using it from another machine — and trusting the certificate

`-p` with no host IP given publishes on all of this host's network
interfaces, so it's already reachable from elsewhere on your network. Use:

```
https://<this-host's-LAN-IP-or-hostname>
```

(no port needed — 443 is HTTPS's default). The certificate is signed by a
**locally-generated CA that browsers don't trust by default**, so the first
visit from a new browser needs a one-time step:

1. Copy `certs/mkcert-ca/rootCA.pem` to the client machine.
2. **macOS**: double-click it to open Keychain Access, find the `mkcert ...`
   entry, open it, expand **Trust**, set "When using this certificate" to
   **Always Trust**, and enter your password to confirm.
3. **Fully quit and reopen the browser** (not just reload) — Chromium-family
   browsers (Chrome, Brave, Edge) cache certificate trust decisions per
   session, so a newly-trusted CA often doesn't take effect until restart.
4. **Firefox** keeps its own certificate store separate from the OS
   keychain — import the CA directly via Settings → Privacy & Security →
   Certificates → View Certificates → Authorities → Import, instead of step 2.

If a browser (this bit us with Brave) still shows "Not Secure" after all of
that even though clicking it says "Certificate is valid": it cached the
insecure verdict from a visit *before* the CA was trusted. Clear that site's
data (`brave://settings/content/all` or `chrome://settings/content/all` →
search the hostname/IP → delete) and fully restart the browser again.

If the page doesn't connect at all: a firewall/VPN in the path is more
likely to be blocking a *custom* port than port 443 — that's the whole
reason HTTPS is mapped to 443 above rather than something like `8443`.

There's currently **no authentication** — anyone who can reach the port can
use it. Fine on a trusted LAN/VPN; don't expose it directly to the internet
without adding auth or putting it behind something that provides it (see the
"Not yet built" list in `CLAUDE.md`).

### API

`POST /api/parse` (multipart PDF) returns `{report, available_stats}` — the
parsed config for the confirm step, plus every stat key/label/section this
report could show (for a selection checklist). `POST /api/render` and
`POST /api/preview` take the same body — that report (possibly edited),
`rack_label`, an optional `visible_stats: [key, ...]` (omit or leave `null`
to show everything), an optional `template_base64` (an existing `.pptx`,
base64-encoded, to append the rack slide to and match colors sampled from
its real content — omit or leave `null` for the default styling), and an
optional `template_slide` (a 1-based slide number to sample from instead of
the whole deck — only meaningful when `template_base64` still has its
original slides), and an optional `rack_sizes: [n, n, ...]` (an explicit
node count per rack, must sum to the report's total node count — omit or
leave `null` to auto-split by RU whenever the cluster needs more than one
rack) — `/api/render` streams back the `.pptx`, `/api/preview`
streams back a PNG rendered from that exact `.pptx` via LibreOffice, so
what you preview can't drift from what you'd download. No persistence —
each request renders into its own temp file/directory that's deleted after
streaming.

`POST /api/derive-template` takes `{template_base64, template_slide}` (the
second optional, same meaning as above) and returns
`{template_base64: <stripped>}` — the same slide-stripping as
`derive_template.py`, exposed for the web UI's "style only" checkbox below.

The web UI has an optional "PowerPoint template" file picker alongside the
sizing PDF upload, with a "style only" checkbox (checked by default) and a
"Match slide #" number field for a deck that looks different from slide to
slide: when "style only" is checked, a chosen template is distilled via
`/api/derive-template` before it's used, so what actually gets stored and
sent on every render is the small theme-only file, not the original
branded deck, and any slide number applies to that one-time distillation.
Uncheck it to use the template as-is and have the rack slide inserted
after its existing slides instead — a slide number here gets resent on
every render/preview, since the full deck (and its other slides to sample
from) is still around, so changing it once a preview is showing updates
that preview live, no re-upload needed. Either way, the resulting template
is remembered in the browser (`localStorage`) so you don't need to
re-upload it for every
report.

### Updating / stopping

```bash
docker stop qrack && docker rm qrack   # stop

# rebuild + redeploy (certs already exist under certs/ from setup above)
docker build -t qrack . && docker stop qrack && docker rm qrack && \
  docker run -d --name qrack \
    -p 8080:8000 -p 443:8443 \
    -v "$(pwd)/certs/qrack-cert.pem:/certs/qrack-cert.pem:ro" \
    -v "$(pwd)/certs/qrack-key.pem:/certs/qrack-key.pem:ro" \
    --restart unless-stopped qrack
```

## Delivery options (same core module behind each)

- **Web app** (above): the primary way to use this today.
- **CLI**: fastest for scripting / batch, or when you just want the JSON.
- **Claude skill** (lowest infra): drop the PDF in, Claude reads the fields,
  confirms the new-vs-existing split conversationally, runs the bundled
  renderer, returns the `.pptx`. Also gives you a fallback extractor if the
  report format ever drifts. Nothing to host.

## Extending the renderer

Common tweaks live in `renderer.py`:
- palette constants at the top (node/switch/cable colours)
- `_stat_rows()` — the full catalog of possible stats, grouped by section,
  as `(key, label, value)` rows; add a new stat here and it's automatically
  both drawable and selectable (see `available_stats()` and
  `render_rack(..., visible_stats=...)` — the CLI's `--hide-stat`/
  `--list-stats` and the web confirm step's stat checklist both run through
  this same catalog, so there's nowhere else to add a stat)
- `_draw_stats()` — layout only (card sizing, row height, side-by-side vs.
  stacked); adapts to however many sections/rows survive filtering
- `_node_sequence()` — node ordering (new nodes are floated to the top of
  their model group)
- `_split_into_racks()` / `_rack_geometry()` — multi-rack splitting: auto-fill
  by RU vs. a manual `rack_sizes` override, and the 1-up/2-up layout presets
  (see `CLAUDE.md`'s "Multi-rack layout" section for how 3+ racks reuse the
  2-up preset across additional slides instead of a third preset)
