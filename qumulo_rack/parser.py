"""PDF -> ClusterReport.

Deterministic, format-targeted extraction from a sizing.qumulo.com cluster
report. No LLM / vision / fuzzy reading here -- see CLAUDE.md for why. If the
report format drifts and fields come back None, extend the regexes; do not
paper over misses with guesswork beyond the one documented ambiguity
(new-vs-existing node inference).
"""

from __future__ import annotations

import logging
import math
import os
import re
from dataclasses import dataclass, field, fields
from typing import BinaryIO, Optional, Union

import pdfplumber

# pdfminer logs a "Could not get FontBBox..." warning per glyph on some
# reports' embedded fonts. Harmless (we never touch font metrics), just
# noisy -- quiet it rather than let every parse spam stderr.
logging.getLogger("pdfminer").setLevel(logging.ERROR)

_NUM = r"[\d,]+\.?\d*"


def _num(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    return float(s.replace(",", ""))


def _int(s: Optional[str]) -> Optional[int]:
    n = _num(s)
    return None if n is None else int(n)


def _search(pattern: str, text: str, flags=0) -> Optional[str]:
    m = re.search(pattern, text, flags)
    return m.group(1) if m else None


@dataclass
class NodeModel:
    code: str
    description: Optional[str] = None
    raw_tb_per_node: Optional[float] = None
    count: int = 0
    height_in: Optional[float] = None
    frontend_ports: Optional[int] = None
    backend_ports: Optional[int] = None
    is_new: bool = False
    new_count: int = 0

    @property
    def ru(self) -> int:
        if not self.height_in:
            return 1
        return max(1, math.ceil(self.height_in / 1.75))

    def as_dict(self) -> dict:
        d = {f.name: getattr(self, f.name) for f in fields(self)}
        d["ru"] = self.ru
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "NodeModel":
        known = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in d.items() if k in known})


@dataclass
class ClusterReport:
    title: Optional[str] = None
    is_expansion: bool = False
    node_count: Optional[int] = None
    added_nodes: Optional[int] = None
    models: list = field(default_factory=list)  # list[NodeModel]
    usable_tb: Optional[float] = None
    raw_tb: Optional[float] = None
    added_capacity_tb: Optional[float] = None
    encoding: Optional[str] = None
    efficiency: Optional[float] = None
    scale_to_tb: Optional[float] = None
    perf: dict = field(default_factory=dict)
    drive_outage_tolerance: Optional[int] = None
    node_outage_tolerance: Optional[int] = None
    frontend_ports_total: Optional[int] = None
    backend_ports_total: Optional[int] = None
    frontend_networking: Optional[str] = None
    backend_networking: Optional[str] = None
    rack_u: Optional[int] = None
    additional_rack_u: Optional[int] = None
    weight_lbs: Optional[float] = None
    added_weight_lbs: Optional[float] = None
    watts: Optional[float] = None
    added_watts: Optional[float] = None
    amps_240: Optional[float] = None
    amps_110: Optional[float] = None
    thermal_btu: Optional[float] = None
    license: Optional[str] = None
    write_volume_max: Optional[str] = None
    cluster_overwrite_cadence_max: Optional[str] = None
    source_file: Optional[str] = None
    # Best-effort: label/value pairs found in the summary tables that don't
    # match any field above. Catches a new metric Qumulo adds to those
    # tables without a code change; won't catch a restructured report
    # format or a renamed section. See _extract_catchall_stats.
    extra_stats: dict = field(default_factory=dict)

    @property
    def total_ru_from_models(self) -> int:
        return sum(m.ru * m.count for m in self.models)

    @property
    def new_node_summary(self) -> str:
        parts = [f"{m.new_count}x {m.code}" for m in self.models if m.new_count]
        return ", ".join(parts) if parts else ""

    def as_dict(self) -> dict:
        d = {f.name: getattr(self, f.name) for f in fields(self) if f.name != "models"}
        d["models"] = [m.as_dict() for m in self.models]
        d["total_ru_from_models"] = self.total_ru_from_models
        d["new_node_summary"] = self.new_node_summary
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "ClusterReport":
        known = {f.name for f in fields(cls)}
        kwargs = {k: v for k, v in d.items() if k in known and k != "models"}
        kwargs["models"] = [NodeModel.from_dict(m) for m in d.get("models", [])]
        return cls(**kwargs)


# --- section-scoped field patterns (label -> regex over the whole doc) -----

def _clean_ws(s: Optional[str]) -> Optional[str]:
    """Collapse whitespace/line-wraps and drop stray "(Max)" labels that
    leak into a captured value when a field's number/unit wrap onto
    separate lines -- see cluster_overwrite_cadence_max below."""
    if s is None:
        return None
    return re.sub(r"\s*\(Max\)\s*|\s+", " ", s).strip()


_HEADER_FIELDS = {
    "usable_tb": (r"Capacity \(Usable\)\s+(%s)TB" % _NUM, _num),
    "license": (r"License\s+(\w+)", str),
    "scale_to_tb": (r"Can Scale to\s+(%s)TB" % _NUM, _num),
    "added_capacity_tb": (r"Added Capacity\s+(%s)TB" % _NUM, _num),
    "encoding": (r"Encoding\s+(\d+,\d+)", str),
    "raw_tb": (r"Capacity \(Raw\)\s+(%s)TB" % _NUM, _num),
    "efficiency": (r"Efficiency\s+(\d+)%", _num),
    "node_count": (r"Nodes in Cluster\s+(\d+)nodes", _int),
    "added_nodes": (r"Added Nodes\s+(\d+)nodes", _int),
    "drive_outage_tolerance": (r"Drive Outage Tolerance\s+(\d+)", _int),
    "node_outage_tolerance": (r"Node Outage Tolerance\s+(\d+)", _int),
    "write_volume_max": (r"Write Volume \(Max\)\s+([\d.,\s\-]+TB/day)", str),
    # "(Max)" sometimes precedes the value ("(Max) 2days"), sometimes
    # follows a wrapped number ("9 -\n(Max) 38days"), and the unit is
    # occasionally missing from the extracted text entirely -- capture up
    # to the next section header and clean up whatever's actually there
    # rather than assume one exact layout.
    "cluster_overwrite_cadence_max": (r"Cluster Overwrite Cadence\s+([\s\S]+?)\s*Networking Environmentals", _clean_ws),
    "frontend_ports_total": (r"Front-end Ports\s+(\d+)ports", _int),
    "backend_ports_total": (r"Back-end Ports\s+(\d+)ports", _int),
    "frontend_networking": (r"Front-end Networking\s+(\S.*?GbE)(?=\s+[A-Z])", str),
    "backend_networking": (r"Back-end Networking\s+(\S.*?GbE)(?=\s+[A-Z])", str),
    "rack_u": (r"Rack Space Required\s+(\d+)U", _int),
    "additional_rack_u": (r"Additional Rack Space required\s+(\d+)U", _int),
    "weight_lbs": (r"Weight \(Total\)\s+(%s)lbs" % _NUM, _num),
    "added_weight_lbs": (r"Additional Weight\s+(%s)lbs" % _NUM, _num),
    "watts": (r"Typical Watts \(Total\)\s+(%s)W" % _NUM, _num),
    "added_watts": (r"Additional Power required\s+(%s)W" % _NUM, _num),
    "amps_240": (r"Typical Amps @240V \(Total\)\s+(%s)A" % _NUM, _num),
    "amps_110": (r"Typical Amps @110/115V \(Total\)\s+(%s)A" % _NUM, _num),
    "thermal_btu": (r"Typical Thermal BTU/hr\s+(%s)BTU/hr" % _NUM, _num),
}

_PERF_FIELDS = {
    "cached_read": r"Cached Read\s+(%s)MB/s" % _NUM,
    "uncached_read": r"Uncached Read\s+(%s)MB/s" % _NUM,
    "sustained_write": r"Sustained Write\s+(%s)MB/s" % _NUM,
    "burst_write": r"Burst Write\s+(%s)MB/s" % _NUM,
    "single_stream_write": r"Single Stream Write\s+(%s)MB/s" % _NUM,
    "ss_cached_read": r"SS Cached Read\s+(%s)MB/s" % _NUM,
    "ss_uncached_read": r"SS Uncached Read\s+(%s)MB/s" % _NUM,
    "iops": r"IOPS\s+(%s)" % _NUM,
}

# Literal label text for every field already captured above (by
# _HEADER_FIELDS, _PERF_FIELDS, or hardcoded elsewhere in parse_report) --
# used to keep the catch-all pass below from re-surfacing a field we
# already have a proper name for. Keep this in sync when adding a field;
# drifting out of sync just risks a harmless duplicate in extra_stats; it's
# a best-effort catch-all, not something else's correctness depends on it.
_KNOWN_STAT_LABELS = [
    "Capacity (Usable)", "License", "Burst Write", "Can Scale to", "Cached Read", "IOPS",
    "Added Capacity", "Sustained Write", "Single Stream Write", "Encoding", "Uncached Read",
    "SS Cached Read", "Capacity (Raw)", "SS Uncached Read", "Efficiency", "Nodes in Cluster",
    "Drive Outage Tolerance", "Write Volume (Max)", "Added Nodes", "Node Outage Tolerance",
    "Cluster Overwrite Cadence", "Front-end Ports", "Rack Space Required", "Typical Watts (Total)",
    "Added Ports", "Additional Rack Space required", "Additional Power required", "Back-end Ports",
    "Height (Node)", "Typical Amps @110/115V (Total)", "Front-end Networking", "Width (Node)",
    "Typical Amps @240V (Total)", "Back-end Networking", "Depth (Node)", "Typical Thermal BTU/hr",
    "Weight (Total)", "Additional Weight",
]

# Generic "Label Words Value+Unit" shape, for the catch-all pass -- label is
# one or more Title-Case words (optionally hyphenated, optionally with a
# trailing parenthetical like "(Max)"), value is a number or a "N - M"
# range with an optional known unit. Deliberately conservative (numeric
# values only -- a new ACTIVE/Active-style status field won't be caught,
# same as label text containing a lowercase connector word like "to"/"of")
# rather than risk noisy garbage from a looser pattern.
_CATCHALL_RE = re.compile(
    r"(?<![\w(])"
    r"([A-Z][A-Za-z]*(?:-[A-Za-z]+)*(?:\s+[A-Z][A-Za-z]*(?:-[A-Za-z]+)*)*(?:\s*\([A-Za-z]+\))?)"
    r"\s+"
    r"(%s(?:\s*-\s*%s)?\s*(?:TB/day|BTU/hr|GbE|MB/s|lbs|days?|ports?|nodes?|%%|TB|GB|U\b|in\b|W\b|A\b)?)"
    % (_NUM, _NUM)
)


def _extract_catchall_stats(summary_text: str) -> dict:
    known_lower = [label.lower() for label in _KNOWN_STAT_LABELS]
    found = {}
    for m in _CATCHALL_RE.finditer(summary_text):
        label, value = m.group(1).strip(), m.group(2).strip()
        if len(label) < 3 or not value:
            continue
        if any(label.lower() == k or label.lower() in k or k in label.lower() for k in known_lower):
            continue
        if label not in found:
            found[label] = value
    return found


# The "All Nodes" table row (code / raw capacity / node count) plus its
# "MODEL RAW CAPACITY NODES" column-header row. On most reports these five
# tokens sit on two adjacent lines in that exact order, but a model with an
# extra footnote (e.g. a NIC-option row) can push the header words onto their
# own lines and out of order -- pdfplumber's line grouping follows the PDF's
# vertical text position, not a fixed table grammar. So: match the code/raw/
# count triple strictly (that part is stable), then just confirm the header
# words appear somewhere in the next couple hundred characters, in any order.
_MODEL_CORE_RE = re.compile(r"\n([A-Z][A-Z0-9\-]{1,19})\s*\n?\s*(%s)\s*TB\s+(\d+)\b" % _NUM)
_MODEL_HEADER_CONFIRM_WINDOW = 250


def _is_confirmed_model_header(full_text: str, after: int) -> bool:
    window = full_text[after:after + _MODEL_HEADER_CONFIRM_WINDOW]
    return "MODEL" in window and "RAW CAPACITY" in window and "NODES" in window

_BOILERPLATE_LINE_RE = re.compile(
    r"^\s*(\d{1,2}/\d{1,2}/\d{2,4},|All Nodes\s*$|https?://|Maximum Operating altitude|"
    r"Non-operating Humidity|MODEL\s+RAW CAPACITY\s+NODES)"
)

# Lines belonging to a per-model spec table (Performance / Networking /
# Environmentals). If the nearest non-boilerplate line above a model header
# is one of these, there's no real description text to find (it's on the
# far side of a page break) -- stop and report None rather than grabbing an
# unrelated spec line.
_FIELD_LABEL_LINE_RE = re.compile(
    r"^(CPU|Memory|HDDs|SSDs|Networking Speed|Network Connector|Management Network|"
    r"Weight|Height|Width|Depth|Power|Operating|Non-operating|Maximum|Typical|"
    r"Front-end|Back-end|Performance|Networking|Environmentals|Front view|Rear view|Supply)\b"
)

# The report's internal document name (footer/header text, not the upload
# filename) encodes the size of the model sizing.qumulo.com treated as the
# "driving" one, e.g. "...-QV-1UHG2-L2-96TB-860.68TB-2025-10-09". This is the
# one signal we have for which model was added in an expansion -- a guess,
# surfaced to the caller, never hidden. See CLAUDE.md.
_INTERNAL_NAME_RE = re.compile(r"-(%s)TB-%sTB-\d{4}-\d{2}-\d{2}\b" % (_NUM, _NUM))


def _extract_description(full_text: str, start: int) -> Optional[str]:
    preceding = full_text[:start]
    lines = [l.strip() for l in preceding.splitlines() if l.strip()]
    for line in reversed(lines):
        if _BOILERPLATE_LINE_RE.match(line):
            continue
        if _FIELD_LABEL_LINE_RE.match(line):
            return None
        return line
    return None


def _extract_models(full_text: str) -> list:
    matches = [m for m in _MODEL_CORE_RE.finditer(full_text) if _is_confirmed_model_header(full_text, m.end())]
    models = []
    for i, m in enumerate(matches):
        code, raw_tb, count = m.group(1), _num(m.group(2)), _int(m.group(3))
        block_end = matches[i + 1].start() if i + 1 < len(matches) else min(len(full_text), m.end() + 4000)
        block = full_text[m.end():block_end]
        frontend_ports = _int(_search(r"Front-end Ports\s+(\d+)ports", block))
        backend_ports = _int(_search(r"Back-end Ports\s+(\d+)ports", block))
        height_in = _num(_search(r"Height\s+(%s)in" % _NUM, block))
        description = _extract_description(full_text, m.start())
        models.append(
            NodeModel(
                code=code,
                description=description,
                raw_tb_per_node=raw_tb,
                count=count or 0,
                height_in=height_in,
                frontend_ports=frontend_ports,
                backend_ports=backend_ports,
            )
        )
    return models


def _flag_new_nodes(report: "ClusterReport", full_text: str) -> None:
    """Infer which model is new on an expansion, from the report's internal
    document name. A guess -- see module docstring and CLAUDE.md. Callers
    (CLI --new, web confirm step) may override the result."""
    if not report.is_expansion or not report.models:
        return
    inferred = _search(_INTERNAL_NAME_RE.pattern, full_text)
    inferred_tb = _num(inferred)
    if inferred_tb is None:
        return
    candidates = [m for m in report.models if m.raw_tb_per_node is not None and abs(m.raw_tb_per_node - inferred_tb) < 0.5]
    if len(candidates) != 1:
        return
    model = candidates[0]
    model.is_new = True
    model.new_count = min(model.count, report.added_nodes or model.count)


def parse_report(source: Union[str, BinaryIO], *, name: Optional[str] = None) -> ClusterReport:
    """Parse a sizing.qumulo.com PDF report.

    `source` is a file path or a file-like object (already opened, binary).
    Pass `name` to set `source_file` explicitly -- e.g. a web upload saved to
    a temp path still wants its original filename, not the temp path's. Do
    not trust an uploaded name for anything but display -- see CLAUDE.md
    web-upload notes.
    """
    if name is not None:
        source_file = name
    elif isinstance(source, str):
        source_file = os.path.basename(source)
    else:
        source_file = "upload.pdf"

    with pdfplumber.open(source) as pdf:
        page_texts = [p.extract_text() or "" for p in pdf.pages]
    full_text = "\n".join(page_texts)

    report = ClusterReport(source_file=source_file)

    if "Qumulo Cluster Expansion" in full_text:
        report.is_expansion = True
        report.title = "Qumulo Cluster Expansion"
    elif "New Qumulo Cluster" in full_text:
        report.is_expansion = False
        report.title = "New Qumulo Cluster"
    else:
        report.title = _search(r"^(.*Qumulo.*Cluster.*)$", full_text, re.MULTILINE)

    for attr, (pattern, cast) in _HEADER_FIELDS.items():
        raw = _search(pattern, full_text)
        if raw is not None:
            setattr(report, attr, cast(raw))

    for key, pattern in _PERF_FIELDS.items():
        raw = _search(pattern, full_text)
        if raw is not None:
            report.perf[key] = _num(raw)

    # Page-1 summary tables: stable section markers, regardless of which
    # specific metrics the sizing tool lists between them.
    summary_match = re.search(r"Capacity Performance([\s\S]*?)Capacity Planning", full_text)
    if summary_match:
        report.extra_stats = _extract_catchall_stats(summary_match.group(1))

    report.models = _extract_models(full_text)
    _flag_new_nodes(report, full_text)

    return report
