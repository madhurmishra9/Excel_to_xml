"""Command-line interface: ``python -m excel_xml <command>``."""

from __future__ import annotations

import argparse
import sys

from . import convert
from .compare import check_xml_data
from .search import file_search


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel_xml",
        description="Convert between Excel and XML, and ingest remote feeds/pages.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("excel2xml", help="Excel workbook -> one XML file per sheet")
    p.add_argument("input", help="path to .xls/.xlsx/.xlsm")
    p.add_argument("--out-dir", default="", help="output folder (default: <name>_xml)")

    p = sub.add_parser("xml2excel", help="folder of XML files -> single Excel workbook")
    p.add_argument("xml_dir", help="folder containing .xml files")
    p.add_argument("--out", default="", help="output .xlsx path")

    p = sub.add_parser("compare", help="compare an Excel file against a folder of XML")
    p.add_argument("input", help="path to Excel file")
    p.add_argument("xml_dir", help="folder containing .xml files")

    p = sub.add_parser("search", help="find a file by name across all drives (Windows)")
    p.add_argument("filename", help="exact file name to search for")

    for name, help_text in (
        ("fetch2xml", "fetch a URL/cloud source -> XML files"),
        ("fetch2excel", "fetch a URL/cloud source -> Excel workbook"),
    ):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("source", help="http(s) URL or cloud URI (s3://, gs://, az://, gdrive://)")
        p.add_argument(
            "--kind",
            choices=["auto", "rss", "xml", "html", "data"],
            default="auto",
            help="force the source type instead of auto-detecting",
        )
        if name == "fetch2xml":
            p.add_argument("--out-dir", default="", help="output folder (default: feed_xml)")
        else:
            p.add_argument("--out", default="", help="output .xlsx path (default: feed.xlsx)")

    return parser


def main(argv=None) -> int:
    args = _build_parser().parse_args(argv)

    try:
        if args.command == "excel2xml":
            written = convert.excel_to_xml(args.input, args.out_dir)
            print("Wrote %d XML file(s):" % len(written))
            for path in written:
                print("  " + path)
        elif args.command == "xml2excel":
            out = convert.xml_to_excel(args.xml_dir, args.out)
            print("Wrote " + out)
        elif args.command == "compare":
            result = check_xml_data(args.input, args.xml_dir)
            print("100% Match" if result == 0 else "Differences found.")
            return result
        elif args.command == "search":
            file_search(args.filename)
        elif args.command == "fetch2xml":
            written = convert.remote_to_xml(args.source, args.out_dir, args.kind)
            print("Wrote %d XML file(s):" % len(written))
            for path in written:
                print("  " + path)
        elif args.command == "fetch2excel":
            out = convert.remote_to_excel(args.source, args.out, args.kind)
            print("Wrote " + out)
    except Exception as exc:  # surface a clean message instead of a traceback
        print("Error: %s" % exc, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
