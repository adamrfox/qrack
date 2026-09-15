"""ClusterReport -> .pptx.

Deterministic geometry: every position is computed from the parsed config,
never from randomness or wall-clock time. Shapes are native python-pptx
rectangles / connectors / text boxes -- never flatten a slide to an image --
with one deliberate, narrow exception: 1U node chassis use a bundled product
photo (qumulo_rack/assets/) instead of the drawn server icon, by explicit
product decision. The NEW-node highlight and the code label stay separate
vector overlays on top of it, so those remain editable even though the
chassis art itself isn't.
"""

from __future__ import annotations

import re
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
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
# _draw_rack) -- proper multi-rack splitting is a separate, not-yet-built
# feature (see CLAUDE.md).
RACK_TOTAL_U = 42

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


def _draw_server_photo(slide, x, y, w, h, is_new):
    pic = slide.shapes.add_picture(str(NODE_PHOTO_PATH), x, y, w, h)
    pic.shadow.inherit = False
    for attr, value in NODE_PHOTO_CROP.items():
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


def _draw_switch(slide, x, y, w, h, label):
    shape = _rect(slide, x, y, w, h, fill=SWITCH_FILL, line=NODE_BORDER, line_w=Pt(0.75))
    _rect(slide, x, y, EAR_W, h, fill=SWITCH_EAR)
    _rect(slide, x + w - EAR_W, y, EAR_W, h, fill=SWITCH_EAR)

    inset = Inches(0.05)
    label_w = w * 0.34
    label_x = x + EAR_W + inset
    _text(slide, label_x, y, label_w, h, label, Pt(11), SWITCH_TEXT, bold=True, anchor=MSO_ANCHOR.MIDDLE)

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


def _draw_rack(slide, report: ClusterReport, rack_label: str):
    seq = _node_sequence(report.models)
    total_ru = sum(n["ru"] for n in seq) or 1

    available = RACK_BOTTOM_MAX - RACK_Y - 2 * SWITCH_H
    frame_h = 2 * SWITCH_H + available  # the rack frame always spans the full 42U enclosure

    if total_ru > RACK_TOTAL_U:
        # More nodes than physically fit in one 42U rack -- not yet handled
        # (see "multi-rack splitting" in CLAUDE.md); compress to fit rather
        # than overflowing the frame.
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
        slide, RACK_X, RACK_Y - Inches(0.32), RACK_W, Inches(0.28),
        rack_label, Pt(15), RACK_LABEL_TEXT, bold=True,
    )

    _rect(slide, RACK_X - Inches(0.06), RACK_Y - Inches(0.06), RACK_W + Inches(0.12),
          frame_h + Inches(0.12), fill=None, line=RACK_FRAME, line_w=Pt(1.5))

    y = RACK_Y
    for label in ("ToR Switch A", "ToR Switch B"):
        _draw_switch(slide, RACK_X, y, RACK_W, SWITCH_H, label)
        y += SWITCH_H

    # Cabling runs to the left of the rack (not the right) so the strip to
    # the right of the frame is free for model-code labels instead.
    bus_a_x = RACK_X - Inches(0.18)
    bus_b_x = RACK_X - Inches(0.34)
    switch_a_mid_y = RACK_Y + SWITCH_H * 0.5
    switch_b_mid_y = RACK_Y + SWITCH_H * 1.5
    bus_bottom_y = frame_bottom_y
    if seq:
        # Each bus starts exactly where its switch's feed line ends, so the
        # two segments read as one continuous connected cable rather than a
        # floating gap between the switch and the top of the bus.
        _connector(slide, bus_a_x, switch_a_mid_y, bus_a_x, bus_bottom_y, CABLE_SWITCH_A, Pt(1.5))
        _connector(slide, bus_b_x, switch_b_mid_y, bus_b_x, bus_bottom_y, CABLE_SWITCH_B, Pt(1.5))
        _connector(slide, RACK_X, switch_a_mid_y, bus_a_x, switch_a_mid_y, CABLE_SWITCH_A, Pt(1.5))
        _connector(slide, RACK_X, switch_b_mid_y, bus_b_x, switch_b_mid_y, CABLE_SWITCH_B, Pt(1.5))

    label_x = RACK_X + RACK_W + NODE_LABEL_GAP
    y = nodes_top_y  # nodes sit at the bottom of the rack, not right below the switches
    for node in seq:
        h = ru_height * node["ru"]
        if node["ru"] == 1:
            _draw_server_photo(slide, RACK_X, y, RACK_W, h, node["is_new"])
        else:
            fill = NODE_NEW_FILL if node["is_new"] else NODE_FILL
            _draw_server(slide, RACK_X, y, RACK_W, h, fill, node["is_new"])

        label = node["code"] + (" • NEW" if node["is_new"] else "")
        font_size = Pt(10) if h >= NODE_DETAIL_MIN_H else Pt(7)
        _text(slide, label_x, y, NODE_LABEL_W, h, label, font_size,
              NODE_NEW_FILL if node["is_new"] else NODE_LABEL_TEXT,
              bold=node["is_new"], anchor=MSO_ANCHOR.MIDDLE)

        mid_y = y + h / 2
        _connector(slide, RACK_X, mid_y, bus_a_x, mid_y, CABLE_SWITCH_A, Pt(0.75))
        _connector(slide, RACK_X, mid_y, bus_b_x, mid_y, CABLE_SWITCH_B, Pt(0.75))
        y += h

    legend_y = frame_bottom_y + Inches(0.18)
    sw = Inches(0.14)
    _rect(slide, RACK_X, legend_y, sw, sw, fill=NODE_NEW_FILL)
    _text(slide, RACK_X + sw + Inches(0.08), legend_y - Inches(0.02), Inches(1.4), Inches(0.2),
          "New node", Pt(9), SUBTITLE_TEXT)
    _connector(slide, RACK_X + Inches(1.55), legend_y + sw / 2, RACK_X + Inches(1.85), legend_y + sw / 2, CABLE_SWITCH_A, Pt(1.5))
    _text(slide, RACK_X + Inches(1.9), legend_y - Inches(0.02), Inches(1.0), Inches(0.2),
          "Switch A", Pt(9), SUBTITLE_TEXT)
    _connector(slide, RACK_X + Inches(2.75), legend_y + sw / 2, RACK_X + Inches(3.05), legend_y + sw / 2, CABLE_SWITCH_B, Pt(1.5))
    _text(slide, RACK_X + Inches(3.1), legend_y - Inches(0.02), Inches(1.0), Inches(0.2),
          "Switch B", Pt(9), SUBTITLE_TEXT)

    return legend_y + sw + Inches(0.1)


# --- stats panel -------------------------------------------------------


def _draw_stat_card(slide, cx, cy, w, title, rows, header_h, row_h, two_col, header_fill=STAT_HEADER_FILL):
    """Draw one stat card and return its height. `rows` is already the
    final (label, value) list to show -- filtering happens upstream."""
    n_lines = (len(rows) + 1) // 2 if two_col else len(rows)
    card_h = header_h + row_h * n_lines + Inches(0.1)
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
            _text(slide, ccx + Inches(0.12), ry2, col_w * 0.58, row_h, label, Pt(10.5), STAT_LABEL_TEXT)
            _text(slide, ccx + col_w * 0.58, ry2, col_w * 0.4, row_h, value, Pt(10.5), STAT_VALUE_TEXT,
                  bold=True, align=PP_ALIGN.RIGHT)
            ry2 += row_h
    return card_h


def _draw_stats(slide, report: ClusterReport, visible_stats=None, header_fill=STAT_HEADER_FILL):
    # Sections a user deselected entirely are dropped, not shown empty --
    # the layout below adapts to however many (0-4+) are left, and to
    # however many rows each has: a user can select anywhere from a
    # handful of stats up to every one this report has.
    sections = [(title, [(label, value) for _key, label, value in rows])
                for title, rows in _stat_rows(report, visible_stats) if rows]
    if not sections:
        return

    card_gap = Inches(0.18)
    header_h = Inches(0.32)
    x = STATS_X
    col_w = (STATS_W - card_gap) / 2

    # Build the layout plan first (without knowing row_h yet): a list of
    # "visual rows", each either one full-width card or two side-by-side --
    # only the first two sections ever pair up, matching the original
    # Capacity/Performance-side-by-side design.
    top_row, rest = sections[:2], sections[2:]
    if len(top_row) == 2:
        plan = [[(x, col_w, top_row[0][0], top_row[0][1], False),
                 (x + col_w + card_gap, col_w, top_row[1][0], top_row[1][1], False)]]
        plan += [[(x, STATS_W, title, rows, True)] for title, rows in rest]
    else:
        plan = [[(x, STATS_W, title, rows, True)] for title, rows in top_row + rest]

    def lines_needed(rows, two_col):
        return (len(rows) + 1) // 2 if two_col else len(rows)

    line_counts = [max(lines_needed(rows, two_col) for _, _, _, rows, two_col in vrow) for vrow in plan]

    # Scale row height to whatever total content there turns out to be,
    # so the panel always fits between the rack's top and the slide's
    # bottom regardless of how many stats got selected.
    available_h = SLIDE_H - Inches(0.3) - RACK_Y
    overhead = len(plan) * (header_h + Inches(0.1)) + max(0, len(plan) - 1) * card_gap
    row_h = (available_h - overhead) / (sum(line_counts) or 1)
    row_h = max(Inches(0.16), min(Inches(0.28), row_h))

    y = RACK_Y
    for vrow in plan:
        h = max(_draw_stat_card(slide, cx, y, w, title, rows, header_h, row_h, two_col, header_fill)
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


# --- entry point -------------------------------------------------------


def render_rack(report: ClusterReport, out_path: str, rack_label: str | None = None,
                 visible_stats=None, template_path: str | None = None) -> str:
    """`visible_stats`, when given, is an iterable of stat keys (see
    `available_stats`) -- only those appear in the stats panel, and a
    section left with none of its stats selected is omitted entirely.
    `None` (the default) shows every stat, matching the CLI's behavior.

    `template_path`, when given, is an existing .pptx: our rack slide is
    appended after any slides it already has, using a heuristically-chosen
    blank-ish layout from it, and a few chrome colors (title text, stat
    header bars, background) switch to that template's theme colors. Slide
    dimensions are always forced to this module's fixed 16:9 design
    regardless of the template's own size -- if that differs from the
    template's native size, its *existing* slides (their shapes keep their
    original absolute positions) may look cropped or off-center against
    the new canvas size. Proportionally rescaling this renderer's geometry
    to match an arbitrary template size is a real but not-yet-built
    follow-up; for now this only reliably looks right for a same-aspect-
    ratio (16:9) template, or one with no existing slides to clash with.
    """
    visible_stats = set(visible_stats) if visible_stats is not None else None
    prs = Presentation(template_path) if template_path else Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    layout = _find_blank_layout(prs)
    slide = prs.slides.add_slide(layout)

    theme = _theme_colors(layout.slide_master) if template_path else {}
    bg_color = theme.get("lt1", SLIDE_BG)
    title_color = theme.get("dk1", TITLE_TEXT)
    header_fill = theme.get("accent1", STAT_HEADER_FILL)

    bg = _rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=bg_color)
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)

    node_summary = f"{report.node_count} nodes" if report.node_count is not None else ""
    if report.is_expansion and report.added_nodes:
        node_summary = f"{(report.node_count or 0) - report.added_nodes} existing + {report.added_nodes} new nodes"

    title = report.title or "Qumulo Cluster"
    _text(slide, Inches(0.55), Inches(0.28), Inches(9), Inches(0.4), title, Pt(24), title_color, bold=True)
    subtitle_bits = [b for b in [node_summary, _fmt(report.usable_tb, " TB Usable", 2)] if b]
    _text(slide, Inches(0.55), Inches(0.68), Inches(9), Inches(0.3), " • ".join(subtitle_bits), Pt(13), SUBTITLE_TEXT)

    label = rack_label or "Rack 1"
    rack_bottom = _draw_rack(slide, report, label)
    _draw_stats(slide, report, visible_stats, header_fill=header_fill)

    if report.source_file:
        source_y = min(rack_bottom, SLIDE_H - Inches(0.32))
        _text(slide, Inches(0.55), source_y, Inches(6), Inches(0.25),
              f"Source: {report.source_file}", Pt(8), SUBTITLE_TEXT)

    prs.save(out_path)
    return out_path
