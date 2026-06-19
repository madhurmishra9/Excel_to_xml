"""High-level conversion entry points.

Every function here is thin orchestration over the io / ingest modules; they all
flow through the same canonical model so local and remote paths share writers.
"""

from __future__ import annotations

import os
from typing import List

from . import ingest
from .excel_io import read_excel, write_excel
from .fetch import fetch
from .xml_io import read_xml_dir, write_xml


def excel_to_xml(excel_path: str, out_dir: str) -> List[str]:
    """Convert an Excel workbook to one XML file per sheet.

    Returns the list of written file paths.
    """
    model = read_excel(excel_path)
    if not out_dir:
        out_dir = os.path.splitext(os.path.basename(excel_path))[0] + "_xml"
    written = []
    for sheet_name, rows in model.items():
        written.append(write_xml(sheet_name, rows, out_dir))
    return written


def xml_to_excel(xml_dir: str, out_path: str = "") -> str:
    """Combine all XML files in a folder into a single Excel workbook."""
    model = read_xml_dir(xml_dir)
    if not out_path:
        base = os.path.basename(os.path.normpath(xml_dir)) or "workbook"
        out_path = base + ".xlsx"
    return write_excel(model, out_path)


def _remote_to_model(source: str, kind: str = "auto"):
    content, content_type = fetch(source)
    return ingest.to_model(content, kind=kind, content_type=content_type)


def remote_to_xml(source: str, out_dir: str, kind: str = "auto") -> List[str]:
    """Fetch a remote feed/URL and write it as XML (one file per sheet)."""
    model = _remote_to_model(source, kind)
    if not out_dir:
        out_dir = "feed_xml"
    written = []
    for sheet_name, rows in model.items():
        written.append(write_xml(sheet_name, rows, out_dir))
    return written


def remote_to_excel(source: str, out_path: str = "", kind: str = "auto") -> str:
    """Fetch a remote feed/URL and write it as a single Excel workbook."""
    model = _remote_to_model(source, kind)
    if not out_path:
        out_path = "feed.xlsx"
    return write_excel(model, out_path)
