"""Excel -> XML -> Excel round-trip (values compared as text)."""

from excel_xml.convert import excel_to_xml, xml_to_excel
from excel_xml.excel_io import read_excel


def test_excel_xml_excel_roundtrip(sample_xlsx, tmp_path):
    out_xml = tmp_path / "xml"
    written = excel_to_xml(sample_xlsx, str(out_xml))
    assert len(written) == 2  # one file per sheet

    rebuilt = tmp_path / "rebuilt.xlsx"
    xml_to_excel(str(out_xml), str(rebuilt))

    original = read_excel(sample_xlsx)
    result = read_excel(str(rebuilt))

    # Same sheet names (including the one with a space) and same text values.
    assert list(original.keys()) == list(result.keys())
    for sheet in original:
        assert original[sheet] == result[sheet]


def test_values_are_text(sample_xlsx, tmp_path):
    model = read_excel(sample_xlsx)
    # Numbers became strings on read.
    assert model["Data"][1] == ["apple", "3", "1.5"]
    # Empty cell is an empty string, not None.
    assert model["My Sheet"][1][1] == ""
