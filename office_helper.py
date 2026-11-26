#!/usr/bin/env python3
"""Office document helper for DSH Connectors.

Needs an interpreter with: python-docx, openpyxl, python-pptx.
`_env.ps1` probes for one automatically; override with "python" in
connectors.config.json.

Subcommands
-----------
  docx-read   <file> [--max-chars N]
  docx-write  <file> --title T [--line TEXT ...]
  xlsx-read   <file> [--rows N]
  pptx-read   <file>
"""

import argparse
import sys

SEP = "-" * 60


def _fail(msg, code=2):
    print("ERROR: " + str(msg), file=sys.stderr)
    return code


def cmd_docx_read(args):
    from docx import Document

    doc = Document(args.path)
    paras = [p.text for p in doc.paragraphs]
    body = "\n".join(paras)
    print("FILE: " + args.path)
    print("PARAGRAPHS: %d" % len(paras))
    print("TABLES: %d" % len(doc.tables))
    for i, t in enumerate(doc.tables):
        print("TABLE %d: %d rows x %d cols" % (i + 1, len(t.rows), len(t.columns)))
    print(SEP)
    if args.max_chars and len(body) > args.max_chars:
        print(body[: args.max_chars])
        print("... [truncated, %d chars total]" % len(body))
    else:
        print(body)
    return 0


def cmd_docx_write(args):
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    if args.title:
        h = doc.add_heading(args.title, level=1)
        for r in h.runs:
            r.font.size = Pt(18)
    for line in args.line or []:
        doc.add_paragraph(line)
    doc.save(args.path)
    print("WROTE: " + args.path)
    return 0


def cmd_xlsx_read(args):
    from openpyxl import load_workbook

    wb = load_workbook(args.path, data_only=True, read_only=True)
    print("FILE: " + args.path)
    print("SHEETS: " + ", ".join(wb.sheetnames))
    for name in wb.sheetnames:
        ws = wb[name]
        print(SEP)
        print("SHEET: %s  (max_row=%s, max_col=%s)" % (name, ws.max_row, ws.max_column))
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= args.rows:
                print("  ...")
                break
            cells = ["" if v is None else str(v) for v in row]
            print("  " + " | ".join(cells))
    wb.close()
    return 0


def cmd_pptx_read(args):
    from pptx import Presentation

    prs = Presentation(args.path)
    print("FILE: " + args.path)
    print("SLIDES: %d" % len(prs.slides))
    for i, slide in enumerate(prs.slides, 1):
        print(SEP)
        print("SLIDE %d" % i)
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    print("  " + txt.replace("\n", "\n  "))
    return 0


def build_parser():
    p = argparse.ArgumentParser(prog="office_helper.py")
    sub = p.add_subparsers(dest="cmd", required=True)

    r = sub.add_parser("docx-read")
    r.add_argument("path")
    r.add_argument("--max-chars", type=int, default=0)
    r.set_defaults(func=cmd_docx_read)

    w = sub.add_parser("docx-write")
    w.add_argument("path")
    w.add_argument("--title", default="")
    w.add_argument("--line", action="append")
    w.set_defaults(func=cmd_docx_write)

    x = sub.add_parser("xlsx-read")
    x.add_argument("path")
    x.add_argument("--rows", type=int, default=10)
    x.set_defaults(func=cmd_xlsx_read)

    s = sub.add_parser("pptx-read")
    s.add_argument("path")
    s.set_defaults(func=cmd_pptx_read)

    return p


def main():
    args = build_parser().parse_args()
    try:
        return args.func(args)
    except FileNotFoundError:
        return _fail("file not found: " + getattr(args, "path", "?"))
    except ImportError as e:
        return _fail(
            "missing library (%s) - run: python -m pip install "
            "python-docx openpyxl python-pptx" % e
        )
    except Exception as e:  # noqa: BLE001 - surface the real reason to the caller
        return _fail("%s: %s" % (type(e).__name__, e))


if __name__ == "__main__":
    sys.exit(main())
