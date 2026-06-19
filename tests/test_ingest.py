"""Remote ingestion: RSS, arbitrary XML and HTML. Network is mocked."""

from excel_xml import convert, ingest
from tests.conftest import SAMPLE_HTML, SAMPLE_RSS, SAMPLE_XML


def test_detect_kind():
    assert ingest.detect_kind(SAMPLE_RSS) == "rss"
    assert ingest.detect_kind(SAMPLE_XML) == "xml"
    assert ingest.detect_kind(SAMPLE_HTML, "text/html") == "html"
    assert ingest.detect_kind(b"<data><S/></data>") == "data"


def test_rss_to_model():
    model = ingest.to_model(SAMPLE_RSS, kind="rss")
    rows = model["Example Feed"]
    assert rows[0][:2] == ["title", "link"]      # header
    assert rows[1][0] == "First"
    assert rows[2][0] == "Second"


def test_arbitrary_xml_flatten():
    model = ingest.to_model(SAMPLE_XML, kind="xml")
    rows = model["book"]
    # Columns are the record attribute + child tags.
    assert rows[0] == ["@id", "author", "title"]
    assert rows[1] == ["1", "A", "T1"]
    assert rows[2] == ["2", "B", "T2"]


def test_html_tables():
    model = ingest.to_model(SAMPLE_HTML, kind="html")
    rows = model["table"]
    assert rows[0] == ["City", "Pop"]
    assert rows[1] == ["Paris", "2M"]


def test_remote_to_excel_mocked(monkeypatch, tmp_path):
    monkeypatch.setattr(
        convert, "fetch", lambda source: (SAMPLE_RSS, "application/rss+xml")
    )
    out = tmp_path / "feed.xlsx"
    result = convert.remote_to_excel("http://example/feed", str(out))

    from excel_xml.excel_io import read_excel

    model = read_excel(result)
    assert "Example Feed" in model
    assert model["Example Feed"][1][0] == "First"
