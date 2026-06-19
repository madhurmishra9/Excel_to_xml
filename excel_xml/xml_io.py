"""Read and write the tabular XML format used by this tool.

Format (backward compatible with the original tool, plus a ``name`` attribute
that fixes sheet names that are not valid XML tags)::

    <data>
      <Sheet1 name="Sheet 1">
        <row1>
          <column1>value</column1>
          <column2>value</column2>
        </row1>
      </Sheet1>
    </data>

One file is written per sheet, named ``<sheet>.xml``.
"""

from __future__ import annotations

import glob
import os
import re
import xml.etree.ElementTree as ET
from typing import List, Tuple

from .workbook import new_model, safe_tag

# Characters not allowed in Windows file names.
_INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def _safe_filename(name: str) -> str:
    cleaned = _INVALID_FILENAME_CHARS.sub("_", name).strip()
    return cleaned or "sheet"


def write_xml(sheet_name: str, rows: List[List[str]], out_dir: str) -> str:
    """Write a single sheet to ``<out_dir>/<sheet_name>.xml``.

    Returns the path written.
    """
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    root = ET.Element("data")
    sheet_el = ET.SubElement(root, safe_tag(str(sheet_name)))
    # Always record the real sheet name so it round-trips even after the tag
    # was sanitised.
    sheet_el.set("name", str(sheet_name))

    for r_index, row in enumerate(rows, start=1):
        row_el = ET.SubElement(sheet_el, "row" + str(r_index))
        for c_index, cell in enumerate(row, start=1):
            col_el = ET.SubElement(row_el, "column" + str(c_index))
            col_el.text = "" if cell is None else str(cell)

    tree = ET.ElementTree(root)
    # Pretty-print when available (Python 3.9+); harmless otherwise.
    if hasattr(ET, "indent"):
        ET.indent(tree, space="  ")

    out_path = os.path.join(out_dir, _safe_filename(str(sheet_name)) + ".xml")
    tree.write(out_path, encoding="utf-8", xml_declaration=True)
    return out_path


def _row_sort_key(tag: str) -> int:
    m = re.search(r"(\d+)$", tag)
    return int(m.group(1)) if m else 0


def read_xml_file(path: str) -> Tuple[str, List[List[str]]]:
    """Parse one tabular XML file into ``(sheet_name, rows)``."""
    tree = ET.parse(path)
    root = tree.getroot()

    # The single child of <data> is the sheet element.  If the file has no
    # sheet element, fall back to the file name.
    sheet_el = next(iter(root), None)
    if sheet_el is None:
        return os.path.splitext(os.path.basename(path))[0], []

    sheet_name = sheet_el.get("name") or sheet_el.tag

    rows: List[List[str]] = []
    row_els = sorted(list(sheet_el), key=lambda el: _row_sort_key(el.tag))
    for row_el in row_els:
        col_els = sorted(list(row_el), key=lambda el: _row_sort_key(el.tag))
        rows.append([(col.text or "") for col in col_els])

    return sheet_name, rows


def model_from_data_root(root) -> "object":
    """Build a model from an already-parsed ``<data>`` root element.

    Handles a ``<data>`` document that contains one *or more* sheet children
    (e.g. a single fetched XML file holding several sheets).
    """
    model = new_model()
    for sheet_el in list(root):
        sheet_name = sheet_el.get("name") or sheet_el.tag
        rows: List[List[str]] = []
        for row_el in sorted(list(sheet_el), key=lambda el: _row_sort_key(el.tag)):
            cols = sorted(list(row_el), key=lambda el: _row_sort_key(el.tag))
            rows.append([(col.text or "") for col in cols])
        unique = sheet_name
        n = 1
        while unique in model:
            n += 1
            unique = "%s_%d" % (sheet_name, n)
        model[unique] = rows
    return model


def read_xml_dir(xml_dir: str):
    """Combine every ``*.xml`` file in a folder into one workbook model."""
    if not os.path.isdir(xml_dir):
        raise NotADirectoryError("XML directory not found: %s" % xml_dir)

    model = new_model()
    paths = sorted(glob.glob(os.path.join(xml_dir, "*.xml")))
    if not paths:
        raise FileNotFoundError("No .xml files found in %s" % xml_dir)

    for path in paths:
        sheet_name, rows = read_xml_file(path)
        # Avoid clobbering if two files resolve to the same sheet name.
        unique = sheet_name
        n = 1
        while unique in model:
            n += 1
            unique = "%s_%d" % (sheet_name, n)
        model[unique] = rows
    return model
