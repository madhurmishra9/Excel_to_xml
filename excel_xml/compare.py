"""Compare an Excel workbook against a folder of XML files.

Both sides are read into the canonical model and compared cell by cell, so this
works for ``.xlsx`` as well as ``.xls`` (the original tool silently broke on
``.xlsx``).  Differences are printed; the function returns ``0`` for a full
match and ``1`` when anything differs.
"""

from __future__ import annotations

from .excel_io import read_excel
from .xml_io import read_xml_dir


def _cell(rows, r, c):
    if r < len(rows) and c < len(rows[r]):
        return rows[r][c]
    return None


def check_xml_data(excel_path: str, xml_dir: str) -> int:
    """Return ``0`` if the Excel file and the XML folder match, else ``1``."""
    excel_model = read_excel(excel_path)
    xml_model = read_xml_dir(xml_dir)

    differences = 0
    checked_sheets = set()

    for sheet_name, excel_rows in excel_model.items():
        checked_sheets.add(sheet_name)
        if sheet_name not in xml_model:
            print("sheet %r is present in Excel but has no matching XML" % sheet_name)
            differences = 1
            continue
        xml_rows = xml_model[sheet_name]

        n_rows = max(len(excel_rows), len(xml_rows))
        for r in range(n_rows):
            erow = excel_rows[r] if r < len(excel_rows) else []
            xrow = xml_rows[r] if r < len(xml_rows) else []
            n_cols = max(len(erow), len(xrow))
            for c in range(n_cols):
                xls_val = _cell(excel_rows, r, c)
                xml_val = _cell(xml_rows, r, c)
                if (xls_val or "") != (xml_val or ""):
                    print(
                        "data differs in %s at %d row and %d column - "
                        "current value in xls is %s and in xml is %s"
                        % (sheet_name, r + 1, c + 1, xls_val, xml_val)
                    )
                    differences = 1

    for sheet_name in xml_model:
        if sheet_name not in checked_sheets:
            print("sheet %r is present in XML but not in the Excel file" % sheet_name)
            differences = 1

    return differences
