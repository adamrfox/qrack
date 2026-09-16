#!/usr/bin/env python3
"""CLI: strip all slides from an existing .pptx, leaving just its slide
masters/layouts/theme -- a one-time distillation of a large branded deck
into a small, reusable style-only file for qrack.py's --template flag (or
the web app's template upload), instead of carrying every original slide
along on each render.

    python derive_template.py corp-deck.pptx [-o corp-template.pptx] [--slide N]
"""

import argparse

from qumulo_rack.renderer import derive_template


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx", help="an existing .pptx to distill down to just its theme/layouts")
    ap.add_argument("-o", "--out", help="output .pptx path (default: <pptx>.template.pptx)")
    ap.add_argument("--slide", type=int, metavar="N",
                     help="sample colors from this 1-based slide number instead of the whole deck "
                          "(useful when the deck has more than one distinct look)")
    args = ap.parse_args(argv)

    out_path = args.out or (args.pptx.rsplit(".", 1)[0] + ".template.pptx")
    try:
        derive_template(args.pptx, out_path, slide_index=args.slide)
    except ValueError as exc:
        raise SystemExit(f"--slide: {exc}")
    except Exception as exc:
        raise SystemExit(f"couldn't open {args.pptx!r} as a .pptx: {exc}")
    print(f"wrote {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
