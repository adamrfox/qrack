"""ClusterReport -> .pptx.

Deterministic geometry: every position is computed from the parsed config,
never from randomness or wall-clock time. Shapes are native python-pptx
rectangles / connectors / text boxes -- never flatten a slide to an image --
with one deliberate, narrow exception: 1U and 2U node chassis use a bundled
product photo (qumulo_rack/assets/) instead of the drawn server icon, by
explicit product decision. Each photo is a specific real model's chassis,
used generically for any node of that height regardless of its actual model
code -- exact-per-model photos aren't worth the asset-maintenance burden.
The NEW-node highlight and the code label stay separate vector overlays on
top of it, so those remain editable even though the chassis art itself
isn't. Any other height (3U+) falls back to the drawn vector server icon,
since there's no bundled photo for it.
"""

from __future__ import annotations

import colorsys
import re
from collections import Counter
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL_TYPE, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .parser import ClusterReport, NodeModel

ASSETS_DIR = Path(__file__).parent / "assets"

# --- palette --------------------------------------------------------------
# Centralize theming here -- nothing else in this module should hardcode a
# colour.

SLIDE_BG = RGBColor(0xFF, 0xFF, 0xFF)
PANEL_BG = RGBColor(0xF4, 0xF6, 0xF8)

RACK_FRAME = RGBColor(0x9A, 0xA5, 0xB1)
RACK_LABEL_TEXT = RGBColor(0x33, 0x3A, 0x42)

SWITCH_FILL = RGBColor(0x1B, 0x1F, 0x24)
SWITCH_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
SWITCH_EAR = RGBColor(0x33, 0x3A, 0x42)
SWITCH_PORT_FILL = RGBColor(0x3A, 0x41, 0x49)
SWITCH_PORT_LINE = RGBColor(0x53, 0x5C, 0x66)

NODE_FILL = RGBColor(0x3B, 0x4A, 0x5A)
NODE_NEW_FILL = RGBColor(0x2E, 0x8B, 0x57)
NODE_BORDER = RGBColor(0xFF, 0xFF, 0xFF)
NODE_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
NODE_EAR = RGBColor(0x20, 0x26, 0x2C)
NODE_BAY_FILL = RGBColor(0x2A, 0x34, 0x40)
NODE_BAY_LINE = RGBColor(0x1B, 0x22, 0x2A)
LED_GREEN = RGBColor(0x3D, 0xD6, 0x63)
LED_AMBER = RGBColor(0xF0, 0xA8, 0x3C)

NODE_LABEL_TEXT = RGBColor(0x14, 0x1A, 0x1F)  # dark text for the label strip outside the frame, on white

CABLE_SWITCH_A = RGBColor(0x2F, 0x80, 0xED)
CABLE_SWITCH_B = RGBColor(0xE0, 0x7A, 0x1B)

ACCENT = RGBColor(0x0A, 0x64, 0xA0)
TITLE_TEXT = RGBColor(0x14, 0x1A, 0x1F)
SUBTITLE_TEXT = RGBColor(0x5A, 0x64, 0x72)

STAT_HEADER_FILL = RGBColor(0x0A, 0x64, 0xA0)
STAT_HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
STAT_LABEL_TEXT = RGBColor(0x5A, 0x64, 0x72)
STAT_VALUE_TEXT = RGBColor(0x14, 0x1A, 0x1F)
STAT_CARD_BG = RGBColor(0xFF, 0xFF, 0xFF)
STAT_CARD_BORDER = RGBColor(0xDD, 0xE2, 0xE7)

NEW_BADGE_FILL = NODE_NEW_FILL
NEW_BADGE_TEXT = RGBColor(0xFF, 0xFF, 0xFF)

# --- slide geometry ---------------------------------------------------------

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

RACK_X = Inches(0.55)
RACK_Y = Inches(1.4)
RACK_W = Inches(3.7)
RACK_BOTTOM_MAX = Inches(6.75)
SWITCH_H = Inches(0.34)
RU_MIN = Inches(0.10)
RU_MAX = Inches(0.34)

# A physical rack is a fixed-size enclosure: the frame always spans this
# many rack units below the switches, regardless of how many nodes the
# report actually has -- a half-empty rack shows real empty space instead
# of every cluster stretching to fill the same slide real estate. Also the
# threshold for the "too many nodes for one rack" fallback (see
# _draw_rack), and for deciding where an *auto*-split between racks falls
# (see _split_into_racks) once a cluster needs more than one.
RACK_TOTAL_U = 42

# The report's own rack_u/node heights never include the ToR switches --
# they're networking gear, not part of the storage node count -- but they
# still take real physical space in an actual 42U rack. Assume 1U per
# switch (a common ToR form factor) when deciding how many nodes an
# auto-split rack can hold; a manually-specified split (see
# _split_into_racks's rack_sizes) isn't bound by this at all.
RACK_SWITCH_RESERVED_U = 2
RACK_NODE_CAPACITY_U = RACK_TOTAL_U - RACK_SWITCH_RESERVED_U

# Below this, the 1U node photo (stretched from its own fixed aspect ratio)
# doesn't have enough vertical resolution left to read as a server -- it
# degrades into visual noise well before real PowerPoint or LibreOffice hit
# any actual rendering limit. Node rows get at least this tall even when
# that means a small cluster renders "bigger" than strict 1/42-of-the-rack
# scale would dictate; only a cluster close to physically filling the rack
# compresses below it (see _draw_rack).
NODE_MIN_LEGIBLE_H = Inches(0.20)

# Cabling runs to the left of the rack and the model-code label lives to its
# right, outside the grey frame, instead of overlaid on the node art --
# STATS_X is pushed right (narrowing the stats panel) to give that label
# strip enough width to hold e.g. "AHG3-240T • NEW" legibly.
NODE_LABEL_GAP = Inches(0.15)
NODE_LABEL_W = Inches(1.3)

STATS_X = Inches(4.75) + NODE_LABEL_GAP + NODE_LABEL_W + Inches(0.15)
STATS_W = Inches(8.05) - (NODE_LABEL_GAP + NODE_LABEL_W + Inches(0.15))

# Second rack column, when a cluster needs two (see _rack_geometry) --
# narrower than the single-rack numbers above on both the frame and its
# label, to leave the stats panel enough width for a single (rather than
# side-by-side) column of cards.
RACK_W_2UP = Inches(2.5)
NODE_LABEL_W_2UP = Inches(1.0)
RACK_GAP_2UP = Inches(0.35)  # between rack 1's label strip and rack 2's frame
STATS_GAP_2UP = Inches(0.25)  # between rack 2's label strip and the stats panel
STATS_RIGHT_MARGIN = Inches(0.53)  # matches the single-rack layout's implied right margin


# --- shape helpers ----------------------------------------------------------


def _rect(slide, x, y, w, h, fill=None, line=None, line_w=Pt(0.75)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.shadow.inherit = False
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = line_w
    return shape


def _text(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri"):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    return box


def _label_in_shape(shape, text, size, color, bold=False, align=PP_ALIGN.CENTER, font="Calibri"):
    tf = shape.text_frame
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    return shape


def _connector(slide, x1, y1, x2, y2, color, weight=Pt(1.25)):
    # python-pptx's add_connector computes off/ext via min()/abs() on these
    # four values with no int-casting anywhere in that path (unlike
    # add_shape/add_textbox, which sanitize float input back to int
    # internally). Any float here -- e.g. from a stray `/2` on an Emu value
    # upstream -- lands in the XML as "1435608.0" instead of "1435608".
    # PowerPoint's parser rejects that outright ("needs repair") even though
    # it's schema-valid and LibreOffice/python-pptx tolerate it fine. Coerce
    # to int here, at the one chokepoint every connector call goes through,
    # rather than chasing every division site upstream.
    x1, y1, x2, y2 = (Emu(round(v)) for v in (x1, y1, x2, y2))
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    conn.shadow.inherit = False
    conn.line.color.rgb = color
    conn.line.width = weight
    return conn


def _oval(slide, x, y, w, h, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shape.shadow.inherit = False
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def _picture_locks(picture):
    return picture._element.find(qn("p:nvPicPr")).find(qn("p:cNvPicPr")).find(qn("a:picLocks"))


# --- rack unit iconography ------------------------------------------------
# Real chassis details (mounting ears, drive bays, status LEDs, switch
# ports), not plain boxes -- but native vector shapes throughout, never an
# image, so the deck stays hand-editable. Below a height threshold the fine
# detail (bays/LEDs) is dropped and only ears + a label remain, since a
# handful of tiny slivers reads as noise rather than a drive bay once a rack
# has enough nodes to compress each row.

EAR_W = Inches(0.045)
NODE_DETAIL_MIN_H = Inches(0.16)

# 1U node chassis photo. The source PNG is a 582x360 canvas but the actual
# artwork is a tight ~576x55px strip (aspect ~10.5:1, close to a real 1U's
# 19in/1.75in front-panel proportions) floating in transparent padding --
# crop that padding out so the strip fills the row without distorting it.
# Fractions computed once from the bundled asset's alpha bounding box
# (samples/1u-qumulo-node.png bbox == (4, 151, 580, 206)) with a small
# safety margin; re-derive them if the asset is ever replaced.
NODE_PHOTO_PATH = ASSETS_DIR / "1u-qumulo-node.png"
NODE_PHOTO_CROP = {"crop_left": 0.0034, "crop_right": 0.0, "crop_top": 0.4139, "crop_bottom": 0.4222}

# 2U node chassis photo -- a specific real model's chassis, used generically
# for any 2U-height node (not exact-match-per-model, same as the 1U photo
# above). Source PNG is 745x169; alpha bbox (20, 19, 725, 150), cropped with
# the same ~2px safety margin approach as the 1U asset; re-derive if the
# asset is ever replaced.
NODE_PHOTO_2U_PATH = ASSETS_DIR / "2u-qumulo-node.png"
NODE_PHOTO_2U_CROP = {"crop_left": 0.0242, "crop_right": 0.0242, "crop_top": 0.1006, "crop_bottom": 0.1006}


def _draw_server(slide, x, y, w, h, fill, is_new):
    shape = _rect(slide, x, y, w, h, fill=fill,
                  line=NODE_NEW_FILL if is_new else NODE_BORDER,
                  line_w=Pt(2) if is_new else Pt(0.75))
    _rect(slide, x, y, EAR_W, h, fill=NODE_EAR)
    _rect(slide, x + w - EAR_W, y, EAR_W, h, fill=NODE_EAR)

    if h >= NODE_DETAIL_MIN_H:
        inset = Inches(0.035)
        led_d = Inches(0.045)
        led_gap = Inches(0.03)
        led_x = x + w - EAR_W - inset - led_d
        led_cy = y + h / 2
        _oval(slide, led_x, led_cy - led_d - led_gap / 2, led_d, led_d, LED_GREEN)
        _oval(slide, led_x, led_cy + led_gap / 2, led_d, led_d, LED_AMBER)

        bay_x = x + EAR_W + inset
        bay_w = led_x - inset - bay_x
        bay_h = h - Inches(0.06)
        bay_y = y + (h - bay_h) / 2
        _rect(slide, bay_x, bay_y, bay_w, bay_h, fill=NODE_BAY_FILL, line=NODE_BAY_LINE, line_w=Pt(0.5))
        bay_slots = 6
        for i in range(1, bay_slots):
            bx = bay_x + bay_w * i / bay_slots
            _connector(slide, bx, bay_y, bx, bay_y + bay_h, NODE_BAY_LINE, Pt(0.5))
    return shape


def _draw_server_photo(slide, x, y, w, h, is_new, photo_path=NODE_PHOTO_PATH, photo_crop=NODE_PHOTO_CROP):
    pic = slide.shapes.add_picture(str(photo_path), x, y, w, h)
    pic.shadow.inherit = False
    for attr, value in photo_crop.items():
        setattr(pic, attr, value)
    # python-pptx always adds noChangeAspect="1" (a resize-handle hint, not a
    # rendering rule per the OOXML spec) but real PowerPoint appears to also
    # use it at render time here, where the target box is more elongated
    # than the cropped photo's native aspect -- rather than stretch
    # non-uniformly like every other renderer, it tiles the image to
    # preserve that native aspect, which read as "multiple nodes per row".
    # Dropping the lock forces the plain `<a:stretch>` fill.
    _picture_locks(pic).attrib.pop("noChangeAspect", None)

    # drawn last so the highlight frame isn't hidden under the photo
    border = _rect(slide, x, y, w, h, fill=None,
                    line=NODE_NEW_FILL if is_new else NODE_BORDER,
                    line_w=Pt(2) if is_new else Pt(0.75))
    return border


def _draw_switch(slide, x, y, w, h, label, font_size=Pt(11)):
    shape = _rect(slide, x, y, w, h, fill=SWITCH_FILL, line=NODE_BORDER, line_w=Pt(0.75))
    _rect(slide, x, y, EAR_W, h, fill=SWITCH_EAR)
    _rect(slide, x + w - EAR_W, y, EAR_W, h, fill=SWITCH_EAR)

    inset = Inches(0.05)
    label_w = w * 0.34
    label_x = x + EAR_W + inset
    _text(slide, label_x, y, label_w, h, label, font_size, SWITCH_TEXT, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    port_x = label_x + label_w + inset
    port_zone_w = x + w - EAR_W - inset - port_x
    rows, cols = 2, 12
    gap = Inches(0.02)
    port_w = (port_zone_w - gap * (cols - 1)) / cols
    port_h = (h - Inches(0.08) - gap) / rows
    top = y + Inches(0.04)
    for r in range(rows):
        for c in range(cols):
            px = port_x + c * (port_w + gap)
            py = top + r * (port_h + gap)
            _rect(slide, px, py, port_w, port_h, fill=SWITCH_PORT_FILL, line=SWITCH_PORT_LINE, line_w=Pt(0.25))
    return shape


# --- node ordering ------------------------------------------------------


def _node_sequence(models: list) -> list:
    """One entry per physical node, in rack-elevation order. Within each
    model's group, new nodes are floated to the top -- change ordering here,
    in one place, if that policy ever changes."""
    seq = []
    for m in models:
        new_n = min(m.new_count or 0, m.count)
        existing_n = max(m.count - new_n, 0)
        for _ in range(new_n):
            seq.append({"code": m.code, "is_new": True, "ru": m.ru})
        for _ in range(existing_n):
            seq.append({"code": m.code, "is_new": False, "ru": m.ru})
    return seq


def _split_into_racks(seq: list, rack_sizes: list[int] | None = None) -> list[list]:
    """Splits a flat node sequence (see `_node_sequence`) into one list per
    physical rack, in the same order.

    `rack_sizes`, when given, is an explicit node count per rack -- meant
    to be the user's edit of a previous auto-split result, so it must
    fully account for every node: raises `ValueError` if any count is
    negative or the counts don't sum to `len(seq)` exactly, rather than
    silently dropping or absorbing a mismatch into the last rack.

    When omitted, racks are filled automatically by cumulative RU up to
    `RACK_NODE_CAPACITY_U` per rack. The overwhelmingly common case --
    everything fits in one rack -- yields exactly `[seq]`, so this is a
    no-op for every report that predates multi-rack support.
    """
    if rack_sizes is not None:
        if any(n < 0 for n in rack_sizes):
            raise ValueError("rack sizes must be non-negative")
        if sum(rack_sizes) != len(seq):
            raise ValueError(f"rack sizes sum to {sum(rack_sizes)}, but there are {len(seq)} node(s) to place")
        racks = []
        i = 0
        for n in rack_sizes:
            racks.append(seq[i:i + n])
            i += n
        return racks

    racks = []
    current: list = []
    current_ru = 0
    for node in seq:
        if current and current_ru + node["ru"] > RACK_NODE_CAPACITY_U:
            racks.append(current)
            current = []
            current_ru = 0
        current.append(node)
        current_ru += node["ru"]
    racks.append(current)
    return racks


def _rack_geometry(num_racks: int) -> dict:
    """Positions for `num_racks` rack columns sharing one slide, plus
    where that leaves the stats panel. `num_racks == 1` reproduces the
    module's original single-rack constants exactly (`RACK_X`/`RACK_W`/
    `NODE_LABEL_W`/`STATS_X`/`STATS_W`) -- untouched, unscaled. 2 racks
    share the narrower `*_2UP` constants on both columns, leaving the
    stats panel one (rather than side-by-side) column of cards' worth of
    width. Anything beyond 2 doesn't fit on one slide at any legible
    scale -- that's a separate, not-yet-built multi-slide feature (see
    CLAUDE.md)."""
    if num_racks == 1:
        return {"rack_x": [RACK_X], "rack_w": RACK_W, "label_w": NODE_LABEL_W,
                "stats_x": STATS_X, "stats_w": STATS_W}
    if num_racks == 2:
        rack_x0 = RACK_X
        rack_x1 = rack_x0 + RACK_W_2UP + NODE_LABEL_GAP + NODE_LABEL_W_2UP + RACK_GAP_2UP
        stats_x = rack_x1 + RACK_W_2UP + NODE_LABEL_GAP + NODE_LABEL_W_2UP + STATS_GAP_2UP
        stats_w = SLIDE_W - stats_x - STATS_RIGHT_MARGIN
        return {"rack_x": [rack_x0, rack_x1], "rack_w": RACK_W_2UP, "label_w": NODE_LABEL_W_2UP,
                "stats_x": stats_x, "stats_w": stats_w}
    raise ValueError(f"this report needs {num_racks} racks; more than 2 on one slide isn't supported yet")


def _rack_labels(base_label: str | None, num_racks: int) -> list[str]:
    """`base_label` as-is for a single rack (unchanged default behavior);
    for multiple, a "Rack N" suffix on each so they're distinguishable --
    imperfect if `base_label` was itself already rack-specific (e.g. "Row
    3 / Rack 12"), but a free-text single label can't unambiguously name
    N racks without the user splitting it themselves, and this is at
    least never blank or ambiguous."""
    if num_racks == 1:
        return [base_label or "Rack 1"]
    prefix = base_label or "Rack"
    return [f"{prefix} — Rack {i + 1}" for i in range(num_racks)]


def auto_rack_split(report: ClusterReport) -> list[int]:
    """The node count per rack `render_rack` would use by default (no
    `rack_sizes` override) -- for a caller (e.g. the web confirm step) to
    show the suggested split and let the user edit it before rendering,
    without duplicating `_split_into_racks`'s auto-fill logic. Always
    `[report's total node count]` (one rack) for a cluster that fits in
    one, same as `render_rack` itself."""
    return [len(rack) for rack in _split_into_racks(_node_sequence(report.models))]


# --- formatting helpers ---------------------------------------------------


def _fmt(value, unit="", decimals=None, thousands=True):
    if value is None:
        return "—"
    if isinstance(value, float) and decimals is None:
        decimals = 0 if value == int(value) else 2
    if isinstance(value, (int, float)):
        if decimals is not None:
            s = f"{value:,.{decimals}f}" if thousands else f"{value:.{decimals}f}"
        else:
            s = f"{value:,}" if thousands else str(value)
        return f"{s}{unit}"
    return f"{value}{unit}"


def _stat_rows(report: ClusterReport, visible: set | None = None):
    """Every stat this report could show, grouped by section, as
    (key, label, value) rows. `visible`, when given, is a set of keys --
    rows outside it are dropped and sections left empty are dropped too.
    `visible=None` means "show everything" (today's default behavior)."""
    perf = report.perf
    capacity = [
        ("usable_capacity", "Usable Capacity", _fmt(report.usable_tb, " TB", 2)),
        ("raw_capacity", "Raw Capacity", _fmt(report.raw_tb, " TB", 0)),
        ("efficiency", "Efficiency", _fmt(report.efficiency, "%", 0)),
        ("encoding", "Encoding", report.encoding or "—"),
        ("can_scale_to", "Can Scale To", _fmt(report.scale_to_tb, " TB", 2)),
        ("license", "License", report.license or "—"),
    ]
    if report.is_expansion:
        capacity.insert(1, ("added_capacity", "Added Capacity", _fmt(report.added_capacity_tb, " TB", 2)))

    performance = [
        ("cached_read", "Cached Read", _fmt(perf.get("cached_read"), " MB/s")),
        ("uncached_read", "Uncached Read", _fmt(perf.get("uncached_read"), " MB/s")),
        ("sustained_write", "Sustained Write", _fmt(perf.get("sustained_write"), " MB/s")),
        ("burst_write", "Burst Write", _fmt(perf.get("burst_write"), " MB/s")),
        ("single_stream_write", "Single Stream Write", _fmt(perf.get("single_stream_write"), " MB/s")),
        ("ss_cached_read", "SS Cached Read", _fmt(perf.get("ss_cached_read"), " MB/s")),
        ("ss_uncached_read", "SS Uncached Read", _fmt(perf.get("ss_uncached_read"), " MB/s")),
        ("iops", "IOPS", _fmt(perf.get("iops"))),
    ]

    endurance = [
        ("write_volume_max", "Write Volume (Max)", report.write_volume_max or "—"),
        ("cluster_overwrite_cadence_max", "Cluster Overwrite Cadence (Max)", report.cluster_overwrite_cadence_max or "—"),
    ]

    rack_power = [
        ("rack_space", "Rack Space", _fmt(report.rack_u, "U")),
        ("weight", "Weight", _fmt(report.weight_lbs, " lbs", 0)),
        ("power", "Power", _fmt(report.watts, " W", 0)),
        ("amps_240", "Amps @240V", _fmt(report.amps_240, " A", 2)),
        ("amps_110", "Amps @110/115V", _fmt(report.amps_110, " A", 2)),
        ("thermal", "Thermal", _fmt(report.thermal_btu, " BTU/hr", 0)),
        ("drive_outage_tolerance", "Drive Outage Tolerance", _fmt(report.drive_outage_tolerance)),
        ("node_outage_tolerance", "Node Outage Tolerance", _fmt(report.node_outage_tolerance)),
        ("frontend_ports", "Front-end Ports", _fmt(report.frontend_ports_total)),
        ("backend_ports", "Back-end Ports", _fmt(report.backend_ports_total)),
        ("frontend_networking", "Front-end Networking", report.frontend_networking or "—"),
        ("backend_networking", "Back-end Networking", report.backend_networking or "—"),
    ]
    if report.is_expansion and report.additional_rack_u:
        rack_power.insert(1, ("additional_rack_space", "Additional Rack Space", _fmt(report.additional_rack_u, "U")))
    if report.is_expansion and report.added_watts:
        rack_power.append(("added_watts", "Additional Power", _fmt(report.added_watts, " W", 0)))
    if report.is_expansion and report.added_weight_lbs:
        rack_power.append(("added_weight", "Additional Weight", _fmt(report.added_weight_lbs, " lbs", 0)))

    # Best-effort catch-all (see parser._extract_catchall_stats): whatever
    # showed up in the summary tables that isn't one of the named fields
    # above, surfaced as-is rather than silently dropped. Keys are derived
    # from the PDF's own label text, so they're stable for a given report
    # but not guaranteed stable across a report-format change upstream.
    other = [
        (f"extra_{re.sub(r'[^a-z0-9]+', '_', label.lower()).strip('_')}", label, value)
        for label, value in report.extra_stats.items()
    ]

    sections = [
        ("Capacity", capacity),
        ("Performance", performance),
        ("Endurance", endurance),
        ("Rack & Power", rack_power),
        ("Other", other),
    ]
    if visible is not None:
        sections = [(title, [row for row in rows if row[0] in visible]) for title, rows in sections]
    return sections


def available_stats(report: ClusterReport) -> list:
    """Every stat this report could show, as flat {section, key, label}
    dicts with no values -- for a caller (e.g. the web confirm step) to
    build a selection checklist from. Built from the same catalog
    `_stat_rows` draws from, so the checklist can never drift from what
    can actually appear on the slide."""
    return [
        {"section": title, "key": key, "label": label}
        for title, rows in _stat_rows(report)
        for key, label, _value in rows
    ]


# --- rack elevation ---------------------------------------------------------


def _draw_rack(slide, seq: list, rack_label: str, rack_x, rack_w, label_w, label_text=None, muted_text=None):
    """Draws one physical rack: frame, switches, cabling, and `seq`'s nodes
    (one rack's slice of `_node_sequence`'s output, via `_split_into_racks`
    -- not necessarily the whole cluster). `rack_x`/`rack_w`/`label_w` come
    from `_rack_geometry`, so this same function draws every column
    regardless of how many racks share the slide.

    `label_text`/`muted_text`, when given, override the rack label /
    node-code-label color respectively -- used when rendering against a
    template, so these stay legible against a sampled background that may
    be dark (the module's own defaults assume a white background).

    Returns the frame's bottom Y. That's the same for every rack on a
    slide (the frame always spans the fixed 42U envelope regardless of
    content), so callers draw one shared legend below all of them via
    `_draw_legend` rather than getting one back per rack.
    """
    label_text = label_text or RACK_LABEL_TEXT
    total_ru = sum(n["ru"] for n in seq) or 1

    available = RACK_BOTTOM_MAX - RACK_Y - 2 * SWITCH_H
    frame_h = 2 * SWITCH_H + available  # the rack frame always spans the full 42U enclosure

    if total_ru > RACK_TOTAL_U:
        # More nodes than physically fit in one 42U rack -- shouldn't
        # normally happen once a report's `_split_into_racks` result is
        # respected, but compress to fit rather than overflowing the frame
        # if it ever does (e.g. a manual rack_sizes override that packs
        # one rack past capacity on purpose).
        ru_height = available / total_ru
    else:
        # Give nodes a legible floor height, even though that means small
        # clusters render "bigger" than strict 1/42-scale would dictate --
        # below this the 1U node photo stretches into unrecognizable noise.
        # Only compress toward true scale once a floor-sized stack of every
        # node this report has genuinely wouldn't fit in the available
        # space (i.e. the rack is close to physically full).
        ru_height = NODE_MIN_LEGIBLE_H
        if ru_height * total_ru > available:
            ru_height = available / total_ru
    ru_height = max(RU_MIN, min(RU_MAX, ru_height))

    frame_bottom_y = RACK_Y + frame_h
    nodes_top_y = frame_bottom_y - ru_height * total_ru  # bottom-align the node stack

    _text(
        slide, rack_x, RACK_Y - Inches(0.32), rack_w, Inches(0.28),
        rack_label, Pt(15), label_text, bold=True,
    )

    _rect(slide, rack_x - Inches(0.06), RACK_Y - Inches(0.06), rack_w + Inches(0.12),
          frame_h + Inches(0.12), fill=None, line=RACK_FRAME, line_w=Pt(1.5))

    # "ToR Switch A/B" at the normal Pt(11) wraps to two lines once the
    # switch (and its label sub-box, a fixed fraction of rack_w) narrows
    # for a multi-rack slide -- shrink to fit rather than let it wrap.
    switch_font_size = Pt(11) if rack_w >= RACK_W else Pt(8)
    y = RACK_Y
    for label in ("ToR Switch A", "ToR Switch B"):
        _draw_switch(slide, rack_x, y, rack_w, SWITCH_H, label, font_size=switch_font_size)
        y += SWITCH_H

    # Cabling runs to the left of the rack (not the right) so the strip to
    # the right of the frame is free for model-code labels instead.
    bus_a_x = rack_x - Inches(0.18)
    bus_b_x = rack_x - Inches(0.34)
    switch_a_mid_y = RACK_Y + SWITCH_H * 0.5
    switch_b_mid_y = RACK_Y + SWITCH_H * 1.5
    bus_bottom_y = frame_bottom_y
    if seq:
        # Each bus starts exactly where its switch's feed line ends, so the
        # two segments read as one continuous connected cable rather than a
        # floating gap between the switch and the top of the bus.
        _connector(slide, bus_a_x, switch_a_mid_y, bus_a_x, bus_bottom_y, CABLE_SWITCH_A, Pt(1.5))
        _connector(slide, bus_b_x, switch_b_mid_y, bus_b_x, bus_bottom_y, CABLE_SWITCH_B, Pt(1.5))
        _connector(slide, rack_x, switch_a_mid_y, bus_a_x, switch_a_mid_y, CABLE_SWITCH_A, Pt(1.5))
        _connector(slide, rack_x, switch_b_mid_y, bus_b_x, switch_b_mid_y, CABLE_SWITCH_B, Pt(1.5))

    label_x = rack_x + rack_w + NODE_LABEL_GAP
    y = nodes_top_y  # nodes sit at the bottom of the rack, not right below the switches
    for node in seq:
        h = ru_height * node["ru"]
        if node["ru"] == 1:
            _draw_server_photo(slide, rack_x, y, rack_w, h, node["is_new"])
        elif node["ru"] == 2:
            _draw_server_photo(slide, rack_x, y, rack_w, h, node["is_new"],
                                photo_path=NODE_PHOTO_2U_PATH, photo_crop=NODE_PHOTO_2U_CROP)
        else:
            fill = NODE_NEW_FILL if node["is_new"] else NODE_FILL
            _draw_server(slide, rack_x, y, rack_w, h, fill, node["is_new"])

        label = node["code"] + (" • NEW" if node["is_new"] else "")
        font_size = Pt(10) if h >= NODE_DETAIL_MIN_H else Pt(7)
        _text(slide, label_x, y, label_w, h, label, font_size,
              NODE_NEW_FILL if node["is_new"] else label_text,
              bold=node["is_new"], anchor=MSO_ANCHOR.MIDDLE)

        mid_y = y + h / 2
        _connector(slide, rack_x, mid_y, bus_a_x, mid_y, CABLE_SWITCH_A, Pt(0.75))
        _connector(slide, rack_x, mid_y, bus_b_x, mid_y, CABLE_SWITCH_B, Pt(0.75))
        y += h

    return frame_bottom_y


def _draw_legend(slide, x, legend_y, muted_text=None):
    """Drawn once per slide (not once per rack) below whichever rack's
    frame bottom the caller passes -- every rack on a slide shares the
    same frame height, so it doesn't matter which one."""
    muted_text = muted_text or SUBTITLE_TEXT
    sw = Inches(0.14)
    _rect(slide, x, legend_y, sw, sw, fill=NODE_NEW_FILL)
    _text(slide, x + sw + Inches(0.08), legend_y - Inches(0.02), Inches(1.4), Inches(0.2),
          "New node", Pt(9), muted_text)
    _connector(slide, x + Inches(1.55), legend_y + sw / 2, x + Inches(1.85), legend_y + sw / 2, CABLE_SWITCH_A, Pt(1.5))
    _text(slide, x + Inches(1.9), legend_y - Inches(0.02), Inches(1.0), Inches(0.2),
          "Switch A", Pt(9), muted_text)
    _connector(slide, x + Inches(2.75), legend_y + sw / 2, x + Inches(3.05), legend_y + sw / 2, CABLE_SWITCH_B, Pt(1.5))
    _text(slide, x + Inches(3.1), legend_y - Inches(0.02), Inches(1.0), Inches(0.2),
          "Switch B", Pt(9), muted_text)

    return legend_y + sw + Inches(0.1)


# --- stats panel -------------------------------------------------------


def _draw_stat_card(slide, cx, cy, w, title, rows, header_h, row_h, two_col, header_fill=STAT_HEADER_FILL,
                     card_pad=Inches(0.1), font_size=Pt(10.5)):
    """Draw one stat card and return its height. `rows` is already the
    final (label, value) list to show -- filtering happens upstream."""
    n_lines = (len(rows) + 1) // 2 if two_col else len(rows)
    card_h = header_h + row_h * n_lines + card_pad
    _rect(slide, cx, cy, w, card_h, fill=STAT_CARD_BG, line=STAT_CARD_BORDER, line_w=Pt(0.75))
    hshape = _rect(slide, cx, cy, w, header_h, fill=header_fill)
    _label_in_shape(hshape, title, Pt(12), STAT_HEADER_TEXT, bold=True, align=PP_ALIGN.LEFT)
    hshape.text_frame.margin_left = Inches(0.12)
    ry = cy + header_h + Inches(0.05)

    cols = [rows]
    col_w = w
    if two_col:
        half = (len(rows) + 1) // 2
        cols = [rows[:half], rows[half:]]
        col_w = w / 2
    for col_idx, col in enumerate(cols):
        ccx = cx + col_idx * col_w
        ry2 = ry
        for label, value in col:
            _text(slide, ccx + Inches(0.12), ry2, col_w * 0.58, row_h, label, font_size, STAT_LABEL_TEXT)
            _text(slide, ccx + col_w * 0.58, ry2, col_w * 0.4, row_h, value, font_size, STAT_VALUE_TEXT,
                  bold=True, align=PP_ALIGN.RIGHT)
            ry2 += row_h
    return card_h


def _draw_stats(slide, report: ClusterReport, visible_stats=None, header_fill=STAT_HEADER_FILL,
                 stats_x=STATS_X, stats_w=STATS_W, allow_pair=True, compact=False):
    """`stats_x`/`stats_w` come from `_rack_geometry`, so this panel fits
    whatever width sharing the slide with one or two racks leaves. `allow_pair`
    disables the first-two-sections-side-by-side layout below -- forced off
    by `render_rack` for the narrower 2-rack width, where two half-width
    cards would be too cramped to read a "label ... value" row in.

    `compact`, set alongside `allow_pair=False`, tightens header height,
    card spacing, and the row-height floor: forcing every card to a single
    column of rows (see `allow_pair` above) roughly doubles how many lines
    a report's full stat set needs, and without reclaiming this overhead
    a report with a realistic number of stats (not even an extreme
    everything-selected case) started overflowing past the slide's bottom
    edge even at the normal floor."""
    # Sections a user deselected entirely are dropped, not shown empty --
    # the layout below adapts to however many (0-4+) are left, and to
    # however many rows each has: a user can select anywhere from a
    # handful of stats up to every one this report has.
    sections = [(title, [(label, value) for _key, label, value in rows])
                for title, rows in _stat_rows(report, visible_stats) if rows]
    if not sections:
        return

    card_gap = Inches(0.1) if compact else Inches(0.18)
    header_h = Inches(0.24) if compact else Inches(0.32)
    row_h_floor = Inches(0.13) if compact else Inches(0.16)
    card_pad = Inches(0.06) if compact else Inches(0.1)
    stat_font_size = Pt(9) if compact else Pt(10.5)
    x = stats_x
    col_w = (stats_w - card_gap) / 2

    # Build the layout plan first (without knowing row_h yet): a list of
    # "visual rows", each either one full-width card or two side-by-side --
    # only the first two sections ever pair up, matching the original
    # Capacity/Performance-side-by-side design.
    # allow_pair also gates each full-width card's own internal two-column
    # row layout, not just whether two cards sit side by side -- both are
    # the same "is there enough width for 2 side-by-side label/value
    # groups" question, and a narrow stats panel (2 racks sharing the
    # slide) answers no to both, not just the first.
    top_row, rest = (sections[:2], sections[2:]) if allow_pair else ([], sections)
    if len(top_row) == 2:
        plan = [[(x, col_w, top_row[0][0], top_row[0][1], False),
                 (x + col_w + card_gap, col_w, top_row[1][0], top_row[1][1], False)]]
        plan += [[(x, stats_w, title, rows, allow_pair)] for title, rows in rest]
    else:
        plan = [[(x, stats_w, title, rows, allow_pair)] for title, rows in top_row + rest]

    def lines_needed(rows, two_col):
        return (len(rows) + 1) // 2 if two_col else len(rows)

    line_counts = [max(lines_needed(rows, two_col) for _, _, _, rows, two_col in vrow) for vrow in plan]

    # Scale row height to whatever total content there turns out to be,
    # so the panel always fits between the rack's top and the slide's
    # bottom regardless of how many stats got selected.
    available_h = SLIDE_H - Inches(0.3) - RACK_Y
    overhead = len(plan) * (header_h + card_pad) + max(0, len(plan) - 1) * card_gap
    row_h = (available_h - overhead) / (sum(line_counts) or 1)
    row_h = max(row_h_floor, min(Inches(0.28), row_h))

    y = RACK_Y
    for vrow in plan:
        h = max(_draw_stat_card(slide, cx, y, w, title, rows, header_h, row_h, two_col, header_fill, card_pad, stat_font_size)
                for cx, w, title, rows, two_col in vrow)
        y += h + card_gap


# --- templates -------------------------------------------------------
# Applying a user's template means two independent things: (1) if it has
# existing slides, our rack slide gets appended after them rather than
# replacing them, and (2) a few "chrome" colors (not the functional ones --
# new-node green and the two cable colors carry real meaning and stay fixed
# regardless of template, so an arbitrary brand palette can't make them
# illegible) switch to the template's own theme colors. Both need picking a
# sensible insertion layout and reading the template's theme -- neither is
# guaranteed to look right for every possible template, so this is a
# best-effort, same spirit as the new-node-guess: do the sensible thing,
# don't silently produce garbage if a template is unusual.

_NON_CONTENT_PLACEHOLDER_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}


def _find_blank_layout(prs: Presentation):
    """The layout with the fewest real content placeholders (date/footer/
    slide-number don't count -- the stock "Blank" layout carries those
    three and nothing else), preferring one literally named "blank" as a
    tie-breaker. There's no schema-level "this is the blank one" flag in
    OOXML -- layout order and naming are just convention -- so this is a
    heuristic, not a guarantee, for an arbitrary uploaded template."""
    layouts = list(prs.slide_layouts)
    if not layouts:
        return None

    def content_placeholder_count(layout):
        return sum(1 for ph in layout.placeholders if ph.placeholder_format.type not in _NON_CONTENT_PLACEHOLDER_TYPES)

    named_blank = [l for l in layouts if "blank" in (l.name or "").lower()]
    pool = named_blank or layouts
    return min(pool, key=content_placeholder_count)


def _theme_colors(slide_master) -> dict:
    """`{'accent1': RGBColor(...), 'dk1': ..., 'lt1': ..., ...}` from the
    given master's theme part. A color can be a literal `srgbClr` or a
    `sysClr` (a named system color, e.g. "windowText") carrying its actual
    RGB as a `lastClr` fallback attribute -- handle both. Returns {} rather
    than raising if the theme part is missing or unrecognizable; callers
    fall back to the module's own default colors either way."""
    try:
        theme_part = slide_master.part.part_related_by(RT.THEME)
        root = etree.fromstring(theme_part.blob)
    except Exception:
        return {}

    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:clrScheme", ns)
    if scheme is None:
        return {}

    colors = {}
    for child in scheme:
        name = etree.QName(child).localname  # e.g. "accent1", "dk1"
        srgb = child.find("a:srgbClr", ns)
        sys_clr = child.find("a:sysClr", ns)
        val = srgb.get("val") if srgb is not None else (sys_clr.get("lastClr") if sys_clr is not None else None)
        if val:
            colors[name] = RGBColor.from_string(val)
    return colors


_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

_LITERAL_SCHEME_KEYS = {
    MSO_THEME_COLOR.DARK_1: "dk1", MSO_THEME_COLOR.LIGHT_1: "lt1",
    MSO_THEME_COLOR.DARK_2: "dk2", MSO_THEME_COLOR.LIGHT_2: "lt2",
    MSO_THEME_COLOR.ACCENT_1: "accent1", MSO_THEME_COLOR.ACCENT_2: "accent2",
    MSO_THEME_COLOR.ACCENT_3: "accent3", MSO_THEME_COLOR.ACCENT_4: "accent4",
    MSO_THEME_COLOR.ACCENT_5: "accent5", MSO_THEME_COLOR.ACCENT_6: "accent6",
    MSO_THEME_COLOR.HYPERLINK: "hlink", MSO_THEME_COLOR.FOLLOWED_HYPERLINK: "folHlink",
}
_SEMANTIC_SCHEME_ATTRS = {
    MSO_THEME_COLOR.TEXT_1: "tx1", MSO_THEME_COLOR.BACKGROUND_1: "bg1",
    MSO_THEME_COLOR.TEXT_2: "tx2", MSO_THEME_COLOR.BACKGROUND_2: "bg2",
}


def _color_map(slide_master) -> dict:
    """`{'bg1': 'lt1', 'tx1': 'dk1', ...}` from the master's `<p:clrMap>` --
    real content overwhelmingly references the *semantic* bg1/tx1/bg2/tx2
    slots, not the literal dk1/lt1/dk2/lt2 scheme slots `_theme_colors()`
    is keyed by, and this mapping is what actually resolves one to the
    other. It's per-master and can be (and on a dark-background master,
    sometimes is) inverted from the usual bg1->lt1/tx1->dk1 assumption, so
    it's never safe to hardcode."""
    el = slide_master._element.find(qn("p:clrMap"))
    return dict(el.attrib) if el is not None else {}


def _resolve_color_format(color_format, theme_colors: dict, color_map: dict):
    """Best-effort `RGBColor` for a python-pptx `ColorFormat`, or `None` if
    it isn't explicitly set on this object (inherited from further up the
    style cascade -- not worth chasing further) or can't be resolved.
    Handles a literal RGB value and a theme reference alike, including the
    semantic tx1/bg1/tx2/bg2 slots (via that master's own `_color_map`,
    never a hardcoded guess) as well as the literal dk1/lt1/dk2/lt2/
    accentN/hlink/folHlink slots."""
    try:
        kind = color_format.type
    except AttributeError:
        return None
    if kind == MSO_COLOR_TYPE.RGB:
        try:
            return color_format.rgb
        except Exception:
            return None
    if kind != MSO_COLOR_TYPE.SCHEME:
        return None  # unset, or an HSL/PRESET/SCRGB/SYSTEM color -- not worth chasing
    theme_color = color_format.theme_color
    key = _LITERAL_SCHEME_KEYS.get(theme_color)
    if key is None:
        attr = _SEMANTIC_SCHEME_ATTRS.get(theme_color)
        key = color_map.get(attr) if attr else None
    return theme_colors.get(key) if key else None


def _is_neutral_color(rgb) -> bool:
    """True for near-white/near-black/muted colors -- the backgrounds,
    borders, and secondary text shades (greys, and desaturated navys/
    charcoals used as "the other text color" on light backgrounds) that
    dominate any deck's color usage but say nothing about its brand
    accent. The 0.4 saturation bar is deliberately strict: real accent
    colors (a bright amber, cyan, teal, ...) read close to fully
    saturated, while a merely-dark-but-not-vibrant color like a desaturated
    navy text shade lands well below it -- seen firsthand sampling a real
    deck, where a 0.15 bar let exactly that kind of navy through as the
    "accent" ahead of the deck's actual bright highlight color."""
    r, g, b = rgb[0] / 255, rgb[1] / 255, rgb[2] / 255
    _hue, lightness, saturation = colorsys.rgb_to_hls(r, g, b)
    return saturation < 0.4 or lightness > 0.93 or lightness < 0.07


def _contrasts(c1, c2, min_diff: float = 0.35) -> bool:
    """True if `c1` is a plausible text color against background `c2` --
    a simple lightness-gap proxy for "readable," not a full WCAG contrast
    ratio, but enough to catch a real failure mode: `_sample_deck_colors`
    tallies the deck's most common title color and most common background
    independently across all slides, so they can come from different
    subsets of slides (e.g. white titles on a few dark section-header
    slides, but a mostly-white deck overall) and land on the *same* color
    once paired up -- an invisible white-on-white title, seen firsthand."""
    l1 = colorsys.rgb_to_hls(c1[0] / 255, c1[1] / 255, c1[2] / 255)[1]
    l2 = colorsys.rgb_to_hls(c2[0] / 255, c2[1] / 255, c2[2] / 255)[1]
    return abs(l1 - l2) >= min_diff


def _blend(c1: RGBColor, c2: RGBColor, t: float) -> RGBColor:
    """Linear-interpolate from `c1` toward `c2` by fraction `t` (0=`c1`,
    1=`c2`). Used to derive a muted secondary-text shade from the sampled
    text/background colors that's always a step toward the background --
    correctly a light gray on a dark deck and a dark gray on a light one --
    rather than a single hardcoded gray tuned for a light background,
    which would be nearly invisible against a dark-themed template."""
    return RGBColor(*(round(c1[i] + (c2[i] - c1[i]) * t) for i in range(3)))


def _resolve_fill_color(fill, theme_colors: dict, color_map: dict):
    """Best-effort `RGBColor` for a python-pptx `FillFormat`, or `None` if
    it's unset or a type we don't read (picture/pattern/textured fills --
    rare for a background or accent shape, not worth chasing). Handles
    `SOLID` directly and, for `GRADIENT`, reads the color of the *first*
    stop (position 0) -- a gradient is drawn as a genuine color transition,
    which this renderer has no way to reproduce, but a flat color read off
    one end is far closer to the source deck than treating a gradient-
    filled shape as if it had no fill at all (the earlier behavior, which
    silently missed a real deck's gradient-filled background entirely and
    fell through to an unrelated, invisible-in-practice master color)."""
    try:
        kind = fill.type
    except Exception:
        return None
    if kind == MSO_FILL_TYPE.SOLID:
        return _resolve_color_format(fill.fore_color, theme_colors, color_map)
    if kind == MSO_FILL_TYPE.GRADIENT:
        try:
            stops = sorted(fill.gradient_stops, key=lambda s: s.position)
            return _resolve_color_format(stops[0].color, theme_colors, color_map) if stops else None
        except Exception:
            return None
    return None


def _effective_background_color(background, theme_colors: dict, color_map: dict):
    try:
        return _resolve_fill_color(background.fill, theme_colors, color_map)
    except Exception:
        return None


def _full_bleed_fill_color(shapes, slide_w, slide_h, theme_colors: dict, color_map: dict):
    """The fill of a shape sized and positioned to cover the entire slide
    -- the common Google Slides export pattern of faking a background with
    a plain rectangle instead of the real OOXML background-fill mechanism,
    which `_effective_background_color` alone would miss entirely. Only
    the first such shape counts; a real background is normally the
    bottom-most (first) shape in the tree anyway."""
    tolerance = int(min(slide_w, slide_h) * 0.02)
    for shape in shapes:
        try:
            if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
                continue
            if (abs(shape.left) > tolerance or abs(shape.top) > tolerance
                    or abs(shape.width - slide_w) > tolerance
                    or abs(shape.height - slide_h) > tolerance):
                continue
            color = _resolve_fill_color(shape.fill, theme_colors, color_map)
        except Exception:
            continue
        if color:
            return color
    return None


def _layout_title_default_color(layout, theme_colors: dict, color_map: dict):
    """A title placeholder's own text runs are frequently left with no
    explicit color at all -- the actual styling lives on the LAYOUT's copy
    of that placeholder (common in Google Slides exports, which bake
    per-layout text styles into `<a:lstStyle>` rather than per-run
    formatting). Reads that layout-level default straight from the XML,
    since python-pptx has no higher-level accessor for it."""
    for ph in layout.placeholders:
        if ph.placeholder_format.type not in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            continue
        lst_style = ph.text_frame._txBody.find(f"{_A_NS}lstStyle")
        lvl1 = lst_style.find(f"{_A_NS}lvl1pPr") if lst_style is not None else None
        def_rpr = lvl1.find(f"{_A_NS}defRPr") if lvl1 is not None else None
        fill = def_rpr.find(f"{_A_NS}solidFill") if def_rpr is not None else None
        if fill is None:
            continue
        srgb = fill.find(f"{_A_NS}srgbClr")
        if srgb is not None:
            return RGBColor.from_string(srgb.get("val"))
        scheme = fill.find(f"{_A_NS}schemeClr")
        if scheme is not None:
            val = scheme.get("val")
            key = val if val in theme_colors else color_map.get(val)
            if key:
                return theme_colors.get(key)
    return None


def _title_placeholder(slide):
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            return shape
    return None


def _sample_deck_colors(prs: Presentation, slide_index: int | None = None) -> dict:
    """Samples colors actually used across the deck's real slides, rather
    than reading the theme's abstract color scheme (`_theme_colors`) -- a
    real branded deck, especially a Google Slides export, doesn't reliably
    paint its content with its own declared theme colors (seen firsthand:
    one deck's theme scheme said "navy text on white", but most of its
    slides are actually solid-black backgrounds with white text -- the
    scheme was simply never applied to most of the deck). Returns up to
    `{'background':, 'text':, 'accent':}` as `RGBColor`, omitting whichever
    it couldn't determine -- callers fall back to `_theme_colors()` and
    then this module's own defaults for the rest.

    - `background`: the most common effective background across all
      slides (explicit slide/layout/master `<p:bg>`, or the Google Slides
      full-bleed-rectangle workaround via `_full_bleed_fill_color`).
    - `text`: the most common color used on a title placeholder, per
      slide (its own run color, falling back to its layout's default via
      `_layout_title_default_color` when the run itself has no override).
    - `accent`: the most common non-neutral color (see `_is_neutral_color`)
      among all text runs and shape fills that isn't already the winning
      background/text color -- a best-effort proxy for "the one color
      this deck uses to draw the eye," not a guarantee.

    Every color is resolved through that *slide's own* master (schemeClr
    references go through that master's `_color_map`, so this stays
    correct even across a deck with more than one master/theme).

    `slide_index`, when given, is a 1-based slide number that restricts
    sampling to that single slide instead of majority-voting across the
    whole deck -- for a deck that genuinely has more than one distinct
    look (e.g. alternating section-header and content styles), so the
    "most common" heuristic isn't the only option. Raises `ValueError` if
    out of range for the deck (including a deck with no slides at all,
    e.g. one already run through `derive_template`).
    """
    slides = list(prs.slides)
    if slide_index is not None:
        if not slides or not (1 <= slide_index <= len(slides)):
            raise ValueError(f"slide {slide_index} is out of range -- this deck has {len(slides)} slide(s)")
        slides = [slides[slide_index - 1]]

    bg_votes: Counter = Counter()
    title_votes: Counter = Counter()
    accent_votes: Counter = Counter()

    for slide in slides:
        layout = slide.slide_layout
        master = layout.slide_master
        theme_colors = _theme_colors(master)
        color_map = _color_map(master)

        # A full-bleed shape at a given level (slide/layout) visually
        # covers whatever <p:bg> that same level declares (or inherits),
        # so it has to be checked *before* falling through to the next
        # level's <p:bg> -- not after every level's <p:bg> has already
        # been tried. Got this backwards originally: a real deck had an
        # unused, never-visible white master-level <p:bg> that still won
        # over the layout's actual painted (navy) full-bleed background,
        # simply because "master background" was checked before any
        # full-bleed shape at all.
        bg = (_effective_background_color(slide.background, theme_colors, color_map)
              or _full_bleed_fill_color(slide.shapes, prs.slide_width, prs.slide_height, theme_colors, color_map)
              or _effective_background_color(layout.background, theme_colors, color_map)
              or _full_bleed_fill_color(layout.shapes, prs.slide_width, prs.slide_height, theme_colors, color_map)
              or _effective_background_color(master.background, theme_colors, color_map))
        if bg:
            bg_votes[bg] += 1

        title_shape = _title_placeholder(slide)
        title_color = None
        if title_shape is not None and title_shape.has_text_frame:
            for para in title_shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        title_color = _resolve_color_format(run.font.color, theme_colors, color_map)
                        if title_color:
                            break
                if title_color:
                    break
            if title_color is None:
                title_color = _layout_title_default_color(layout, theme_colors, color_map)
        if title_color:
            title_votes[title_color] += 1

        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if not run.text.strip():
                            continue
                        color = _resolve_color_format(run.font.color, theme_colors, color_map)
                        if color and not _is_neutral_color(color):
                            accent_votes[color] += 1
            try:
                color = _resolve_fill_color(shape.fill, theme_colors, color_map)
                if color and not _is_neutral_color(color):
                    accent_votes[color] += 1
            except Exception:
                pass

    result = {}
    if bg_votes:
        result["background"] = bg_votes.most_common(1)[0][0]
    bg_pick = result.get("background")
    for color, _n in title_votes.most_common():
        # The winning title color must actually read against the winning
        # background -- see `_contrasts`. When there's no background to
        # check against, take the plain top vote.
        if bg_pick is None or _contrasts(color, bg_pick):
            result["text"] = color
            break
    exclude = {result.get("background"), result.get("text")}
    for color, _n in accent_votes.most_common():
        if color not in exclude:
            result["accent"] = color
            break
    return result


_SAMPLED_TO_SCHEME_SLOT = {"background": "lt1", "text": "dk1", "accent": "accent1"}


def _patch_theme_colors(slide_master, sampled: dict) -> None:
    """Overwrites the master's theme `<a:clrScheme>` entries for
    lt1/dk1/accent1 with the colors `_sample_deck_colors` actually found
    in the deck's real content, so a later plain `_theme_colors()` read
    (e.g. once this file's own slides have been stripped by
    `derive_template` and there's nothing left to sample) picks up the
    sampled palette instead of the scheme's original, possibly-unused
    values. Normalizes to a literal `srgbClr` regardless of what was
    there before (a literal color or a `sysClr`). A no-op for any slot
    `sampled` didn't determine.

    Mutates `slide_master.part.part_related_by(RT.THEME).blob` directly --
    the theme part is a generic (non-XML-aware) `Part` in python-pptx, so
    unlike an `XmlPart` there's no live `.element` tree to edit in place;
    reassigning `.blob` is the actual persistence mechanism here.
    """
    try:
        theme_part = slide_master.part.part_related_by(RT.THEME)
        root = etree.fromstring(theme_part.blob)
    except Exception:
        return
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:clrScheme", ns)
    if scheme is None:
        return
    changed = False
    for sampled_key, slot in _SAMPLED_TO_SCHEME_SLOT.items():
        color = sampled.get(sampled_key)
        if color is None:
            continue
        el = scheme.find(f"a:{slot}", ns)
        if el is None:
            continue
        for child in list(el):
            el.remove(child)
        etree.SubElement(el, f"{_A_NS}srgbClr").set("val", str(color))
        changed = True
    if changed:
        theme_part.blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def derive_template(source_path: str, out_path: str, slide_index: int | None = None) -> str:
    """Strips every slide out of an existing .pptx, leaving only its slide
    masters/layouts/theme -- so a large branded deck can be distilled once
    into a small, content-free file for `render_rack`'s `template_path`
    (or the web app's template upload), instead of carrying every one of
    the source deck's original slides along on every render. What survives
    (masters, layouts, theme colors/fonts) is exactly what `render_rack`
    and `_theme_colors` actually read from a template -- nothing here
    depends on any slide *content*.

    python-pptx has no public "delete slide" API; this uses the documented
    community recipe of removing each slide's entry from the presentation's
    `<p:sldIdLst>` and dropping its relationship, which is what every slide
    deletion in python-pptx (there's no built-in method) is built on.

    Before stripping, samples the real colors actually used across those
    slides (`_sample_deck_colors`) and bakes them into the saved file's own
    theme scheme (`_patch_theme_colors`) -- since those slides are about to
    be gone, this is the only chance to capture them, and it means a later
    plain `_theme_colors()` read of the *distilled* file (which is what
    `render_rack` does once there are no slides left to sample) still
    reflects the deck's actual visual style rather than its possibly-
    unused declared theme.

    `slide_index`, when given, is a 1-based slide number in the *source*
    deck to sample from instead of majority-voting across all of it -- see
    `_sample_deck_colors`. Only meaningful here, against the original
    deck's real slides; there's nothing left to sample once this function
    is done stripping them, which is exactly why this is where that choice
    has to be made.
    """
    prs = Presentation(source_path)
    sampled = _sample_deck_colors(prs, slide_index=slide_index)

    slide_id_list = prs.slides._sldIdLst
    for slide_id in list(slide_id_list):
        r_id = slide_id.get(qn("r:id"))
        prs.part.drop_rel(r_id)
        slide_id_list.remove(slide_id)

    if sampled:
        _patch_theme_colors(prs.slide_masters[0], sampled)

    prs.save(out_path)
    return out_path


# --- entry point -------------------------------------------------------


def render_rack(report: ClusterReport, out_path: str, rack_label: str | None = None,
                 visible_stats=None, template_path: str | None = None,
                 template_slide: int | None = None, rack_sizes: list[int] | None = None) -> str:
    """`visible_stats`, when given, is an iterable of stat keys (see
    `available_stats`) -- only those appear in the stats panel, and a
    section left with none of its stats selected is omitted entirely.
    `None` (the default) shows every stat, matching the CLI's behavior.

    `rack_sizes`, when given, is an explicit node count per rack (see
    `_split_into_racks`); omit it to auto-split by RU whenever the cluster
    needs more than one rack (unchanged, exactly one rack, for the common
    case that doesn't). Up to 2 racks share one slide, scaled down (see
    `_rack_geometry`) to leave room for the stats panel, which always
    covers the *whole* cluster regardless of how many racks it's split
    across -- matching the source report, which never gives a per-rack
    breakdown either. More than 2 racks isn't supported yet (raises
    `ValueError`) -- that needs spilling onto additional slides, a
    separate not-yet-built piece (see CLAUDE.md).

    `template_path`, when given, is an existing .pptx: our rack slide is
    appended after any slides it already has, using a heuristically-chosen
    blank-ish layout from it, and a few chrome colors (title text, stat
    header bars, background) switch to colors sampled from that template's
    real content (`_sample_deck_colors`) -- falling back to its declared
    theme scheme (`_theme_colors`) for anything sampling couldn't
    determine (typically because the template has no slides left to
    sample from, e.g. one already run through `derive_template`, which
    bakes its own sampled colors into the theme for exactly this case).
    `template_slide`, when given, is a 1-based slide number that restricts
    that sampling to just that one slide instead of majority-voting across
    the whole template -- useful for a deck with more than one distinct
    look (raises `ValueError` if out of range, including a template with
    no slides left, e.g. one already run through `derive_template` --
    against a distilled template, pick the slide at `derive_template` time
    instead). Slide dimensions are always forced to this module's fixed
    16:9 design regardless of the template's own size -- if that differs
    from the template's native size, its *existing* slides (their shapes
    keep their original absolute positions) may look cropped or
    off-center against the new canvas size. Proportionally rescaling this
    renderer's geometry to match an arbitrary template size is a real but
    not-yet-built follow-up; for now this only reliably looks right for a
    same-aspect-ratio (16:9) template, or one with no existing slides to
    clash with.
    """
    visible_stats = set(visible_stats) if visible_stats is not None else None
    prs = Presentation(template_path) if template_path else Presentation()
    sampled = _sample_deck_colors(prs, slide_index=template_slide) if template_path else {}
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    layout = _find_blank_layout(prs)
    slide = prs.slides.add_slide(layout)

    # add_slide() clones the layout's own placeholders (title/subtitle/body/
    # etc.) onto the new slide -- normal python-pptx behavior, meant for
    # someone who's about to type into them. We never do that; every field
    # on this slide is our own explicit shape. On a template whose only
    # "blank-ish" layout still carries a few content placeholders (common
    # on real branded decks -- unlike python-pptx's own default template,
    # which has a true zero-placeholder "Blank" layout), those inherited
    # placeholders sit empty on top of/behind our own shapes at whatever
    # position and styling the layout gave them, which can visibly clash.
    # Strip them all immediately; nothing here ever reads from them.
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)

    theme = _theme_colors(layout.slide_master) if template_path else {}
    bg_color = sampled.get("background") or theme.get("lt1", SLIDE_BG)
    title_color = sampled.get("text") or theme.get("dk1", TITLE_TEXT)
    header_fill = sampled.get("accent") or theme.get("accent1", STAT_HEADER_FILL)
    # The module's own default secondary-text colors (SUBTITLE_TEXT,
    # RACK_LABEL_TEXT, NODE_LABEL_TEXT) assume a white background --
    # against a sampled/themed background that may be dark, blending
    # toward it from title_color keeps them a readable step away instead
    # of a near-invisible dark-gray-on-black.
    muted_text = _blend(title_color, bg_color, 0.35) if template_path else SUBTITLE_TEXT
    label_text = title_color if template_path else None

    bg = _rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=bg_color)
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)

    node_summary = f"{report.node_count} nodes" if report.node_count is not None else ""
    if report.is_expansion and report.added_nodes:
        node_summary = f"{(report.node_count or 0) - report.added_nodes} existing + {report.added_nodes} new nodes"

    title = report.title or "Qumulo Cluster"
    _text(slide, Inches(0.55), Inches(0.28), Inches(9), Inches(0.4), title, Pt(24), title_color, bold=True)
    subtitle_bits = [b for b in [node_summary, _fmt(report.usable_tb, " TB Usable", 2)] if b]
    _text(slide, Inches(0.55), Inches(0.68), Inches(9), Inches(0.3), " • ".join(subtitle_bits), Pt(13), muted_text)

    seq_all = _node_sequence(report.models)
    racks = _split_into_racks(seq_all, rack_sizes=rack_sizes)
    geometry = _rack_geometry(len(racks))
    labels = _rack_labels(rack_label, len(racks))

    frame_bottom = None
    for rack_x, rack_seq, rack_lbl in zip(geometry["rack_x"], racks, labels):
        frame_bottom = _draw_rack(slide, rack_seq, rack_lbl, rack_x, geometry["rack_w"], geometry["label_w"],
                                   label_text=label_text, muted_text=muted_text)

    legend_y = frame_bottom + Inches(0.18)
    rack_bottom = _draw_legend(slide, geometry["rack_x"][0], legend_y, muted_text)

    _draw_stats(slide, report, visible_stats, header_fill=header_fill,
                stats_x=geometry["stats_x"], stats_w=geometry["stats_w"],
                allow_pair=(len(racks) == 1), compact=(len(racks) > 1))

    if report.source_file:
        source_y = min(rack_bottom, SLIDE_H - Inches(0.32))
        _text(slide, Inches(0.55), source_y, Inches(6), Inches(0.25),
              f"Source: {report.source_file}", Pt(8), muted_text)

    prs.save(out_path)
    return out_path
