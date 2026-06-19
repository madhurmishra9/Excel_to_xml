"""XML writer/reader behaviour, including the sheet-name fix."""

from collections import OrderedDict

from excel_xml.workbook import safe_tag
from excel_xml.xml_io import read_xml_dir, read_xml_file, write_xml


def test_safe_tag():
    assert safe_tag("Sheet 1") == "Sheet_1"
    assert safe_tag("2023") == "_2023"          # cannot start with a digit
    assert safe_tag("xmlData")[0] == "_"        # cannot start with "xml"
    assert safe_tag("ok_name") == "ok_name"


def test_sheet_name_with_space_roundtrips(tmp_path):
    rows = [["a", "b"], ["1", "2"]]
    path = write_xml("My Sheet", rows, str(tmp_path))
    sheet_name, read_rows = read_xml_file(path)
    assert sheet_name == "My Sheet"   # recovered from the name attribute
    assert read_rows == rows


def test_read_xml_dir_combines_files(tmp_path):
    write_xml("One", [["x"]], str(tmp_path))
    write_xml("Two", [["y"]], str(tmp_path))
    model = read_xml_dir(str(tmp_path))
    assert set(model.keys()) == {"One", "Two"}
    assert model["One"] == [["x"]]
    assert model["Two"] == [["y"]]


def test_legacy_xml_without_name_attr(tmp_path):
    # Simulate an old-tool file: tag == sheet name, no name attribute.
    legacy = tmp_path / "Legacy.xml"
    legacy.write_text(
        "<data><Legacy><row1><column1>v</column1></row1></Legacy></data>",
        encoding="utf-8",
    )
    sheet_name, rows = read_xml_file(str(legacy))
    assert sheet_name == "Legacy"
    assert rows == [["v"]]
