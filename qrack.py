#!/usr/bin/env python3
"""CLI: sizing PDF -> rack elevation .pptx.

    python qrack.py cluster.pdf [-o out.pptx] [--json] [--new CODE:N] [--label STR]
                                 [--hide-stat KEY] [--list-stats] [--template FILE.pptx]
                                 [--template-slide N] [--rack-sizes N,N,...]
"""

import argparse
import json
import sys

from qumulo_rack.parser import parse_report
from qumulo_rack.renderer import available_stats, render_rack


def _apply_new_override(report, overrides: list[str]):
    """--new CODE:N ... replaces the parser's new-node guess entirely."""
    by_code = {m.code: m for m in report.models}
    for spec in overrides:
        if ":" not in spec:
            raise SystemExit(f"--new expects CODE:N, got {spec!r}")
        code, n = spec.rsplit(":", 1)
        if code not in by_code:
            raise SystemExit(f"--new: unknown model code {code!r}; models are {sorted(by_code)}")
        try:
            n = int(n)
        except ValueError:
            raise SystemExit(f"--new: {spec!r} count is not an integer")
        m = by_code[code]
        if n > m.count:
            raise SystemExit(f"--new: {code} only has {m.count} nodes, cannot mark {n} as new")
        m.is_new = n > 0
        m.new_count = n
    overridden = {spec.rsplit(":", 1)[0] for spec in overrides}
    for m in report.models:
        if m.code not in overridden:
            m.is_new = False
            m.new_count = 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Qumulo sizing PDF -> editable rack-elevation .pptx")
    ap.add_argument("pdf", help="path to a sizing.qumulo.com cluster report PDF")
    ap.add_argument("-o", "--out", help="output .pptx path (default: <pdf>.rack.pptx)")
    ap.add_argument("--json", action="store_true", help="dump the parsed config as JSON and exit; no slide is produced")
    ap.add_argument("--new", action="append", default=[], metavar="CODE:N",
                     help="override the new-node guess, e.g. --new AH-96T:2 (repeatable)")
    ap.add_argument("--label", help="rack label shown on the slide, e.g. 'Row 3 / Rack 12'")
    ap.add_argument("--hide-stat", action="append", default=[], metavar="KEY",
                     help="hide a stat from the stats panel by key (repeatable); see --list-stats for keys")
    ap.add_argument("--list-stats", action="store_true",
                     help="list this report's available stat keys (for --hide-stat) and exit; no slide is produced")
    ap.add_argument("--template", metavar="FILE.pptx",
                     help="an existing .pptx to append the rack slide to, matching colors sampled from its "
                          "real slides (functional colors -- new-node green, cable colors -- stay fixed regardless)")
    ap.add_argument("--template-slide", type=int, metavar="N",
                     help="sample template colors from this 1-based slide number instead of the whole deck "
                          "(useful when the template has more than one distinct look); requires --template")
    ap.add_argument("--rack-sizes", metavar="N,N,...",
                     help="explicit node count per rack, e.g. --rack-sizes 15,35 (must sum to the total node "
                          "count); omit to auto-split by RU whenever the cluster needs more than one rack")
    args = ap.parse_args(argv)

    report = parse_report(args.pdf)

    if report.models and report.usable_tb is None:
        print("warning: usable_tb did not parse; this may not be a sizing.qumulo.com report", file=sys.stderr)
    if not report.models:
        print("error: no node models found in 'All Nodes' section; this doesn't look like a parseable sizing report", file=sys.stderr)
        return 1

    if args.list_stats:
        for stat in available_stats(report):
            print(f"{stat['key']:<24} {stat['section']} — {stat['label']}")
        return 0

    if args.new:
        _apply_new_override(report, args.new)
    elif report.is_expansion and not report.new_node_summary:
        print("warning: could not infer which model is new on this expansion; override with --new CODE:N", file=sys.stderr)

    if report.is_expansion:
        guess = report.new_node_summary or "(none inferred)"
        print(f"new-node guess: {guess} -- override with --new CODE:N if wrong", file=sys.stderr)

    if args.json:
        print(json.dumps(report.as_dict(), indent=2))
        return 0

    visible_stats = None
    if args.hide_stat:
        all_keys = {stat["key"] for stat in available_stats(report)}
        unknown = set(args.hide_stat) - all_keys
        if unknown:
            raise SystemExit(f"--hide-stat: unknown key(s) {sorted(unknown)}; see --list-stats for valid keys")
        visible_stats = all_keys - set(args.hide_stat)

    if args.template_slide is not None and not args.template:
        raise SystemExit("--template-slide requires --template")

    rack_sizes = None
    if args.rack_sizes:
        try:
            rack_sizes = [int(n) for n in args.rack_sizes.split(",")]
        except ValueError:
            raise SystemExit(f"--rack-sizes: expected comma-separated integers, got {args.rack_sizes!r}")

    out_path = args.out or (args.pdf.rsplit(".", 1)[0] + ".rack.pptx")
    try:
        render_rack(report, out_path, rack_label=args.label, visible_stats=visible_stats,
                    template_path=args.template, template_slide=args.template_slide,
                    rack_sizes=rack_sizes)
    except ValueError as exc:
        raise SystemExit(str(exc))
    except Exception as exc:
        if args.template:
            raise SystemExit(f"--template: couldn't open {args.template!r} as a .pptx: {exc}")
        raise
    print(f"wrote {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
