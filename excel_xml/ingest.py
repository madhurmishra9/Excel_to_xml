"""Turn arbitrary remote content into the canonical workbook model.

Three source shapes are supported:

* **RSS / Atom feeds** - parsed with ``feedparser``; each entry becomes a row.
* **Arbitrary XML** - the repeating record element is auto-detected and each
  record becomes a row (child tags + ``@attributes`` become columns).  If the
  XML already matches this tool's own ``<data>`` schema it is read directly.
* **HTML pages** - every ``<table>`` becomes a sheet (``bs4``).
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections import OrderedDict
from typing import List, Optional

from .workbook import new_model, stringify_value
from .xml_io import model_from_data_root

# Feed fields surfaced as columns, in display order.  Only those present in at
# least one entry are emitted.
_FEED_FIELDS = ["title", "link", "published", "updated", "author", "summary", "id"]

_FIRST_TAG_RE = re.compile(rb"<\s*([A-Za-z_][\w.\-]*)")


def _localname(tag: str) -> str:
    """Strip an XML namespace, returning just the local element name."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _root_tag(content: bytes) -> Optional[str]:
    try:
        return ET.fromstring(content).tag
    except ET.ParseError:
        m = _FIRST_TAG_RE.search(content)
        return m.group(1).decode("ascii", "ignore") if m else None


def detect_kind(content: bytes, content_type: str = "") -> str:
    """Best-effort classification: ``'rss'``, ``'xml'``, ``'data'`` or ``'html'``."""
    ct = (content_type or "").lower()
    head = content[:512].lstrip()
    head_str = head.decode("utf-8", "ignore").lower()

    if "html" in ct or head_str.startswith("<!doctype html") or head_str.startswith("<html"):
        return "html"
    if "rss" in ct or "atom" in ct:
        return "rss"

    root_tag = _root_tag(content)
    if root_tag:
        local = _localname(root_tag).lower()
        if local in ("rss", "feed"):
            return "rss"
        if local == "data":
            return "data"
        return "xml"

    if "xml" in ct:
        return "xml"
    return "html"


# --------------------------------------------------------------------------- #
# RSS / Atom
# --------------------------------------------------------------------------- #
def _rss_to_model(content: bytes):
    import feedparser

    parsed = feedparser.parse(content)
    entries = parsed.entries or []

    present = [f for f in _FEED_FIELDS if any(e.get(f) for e in entries)]
    if not present:
        present = list(_FEED_FIELDS)

    rows: List[List[str]] = [list(present)]  # header row
    for entry in entries:
        rows.append([stringify_value(entry.get(field, "")) for field in present])

    title = ""
    if getattr(parsed, "feed", None):
        title = parsed.feed.get("title", "")
    sheet_name = title.strip() or "feed"

    model = new_model()
    model[sheet_name] = rows
    return model


# --------------------------------------------------------------------------- #
# Arbitrary XML -> flattened records
# --------------------------------------------------------------------------- #
def _find_record_group(root):
    """Return ``(parent, tag)`` of the largest set of same-tag siblings."""
    best = None  # (count, parent, tag)
    for parent in root.iter():
        counts: "OrderedDict[str, int]" = OrderedDict()
        for child in list(parent):
            counts[child.tag] = counts.get(child.tag, 0) + 1
        for tag, count in counts.items():
            if count > 1 and (best is None or count > best[0]):
                best = (count, parent, tag)
    if best is None:
        return None, None
    return best[1], best[2]


def _record_to_dict(record):
    fields: "OrderedDict[str, str]" = OrderedDict()
    for key, val in record.attrib.items():
        fields["@" + _localname(key)] = val
    children = list(record)
    if children:
        for child in children:
            key = _localname(child.tag)
            text = (child.text or "").strip()
            if key in fields and fields[key]:
                fields[key] = fields[key] + " | " + text
            else:
                fields[key] = text
    else:
        text = (record.text or "").strip()
        if text:
            fields["value"] = text
    return fields


def _xml_to_model(content: bytes):
    root = ET.fromstring(content)
    if _localname(root.tag).lower() == "data":
        return model_from_data_root(root)

    parent, tag = _find_record_group(root)
    if parent is None:
        # No repetition: treat each direct child as a record, else the root.
        records = list(root) or [root]
        record_tag = _localname(root.tag)
    else:
        records = [c for c in list(parent) if c.tag == tag]
        record_tag = _localname(tag)

    dicts = [_record_to_dict(rec) for rec in records]

    columns: "OrderedDict[str, int]" = OrderedDict()
    for d in dicts:
        for key in d:
            columns.setdefault(key, 1)
    headers = list(columns)

    rows: List[List[str]] = [headers]
    for d in dicts:
        rows.append([d.get(col, "") for col in headers])

    model = new_model()
    model[record_tag or "records"] = rows
    return model


# --------------------------------------------------------------------------- #
# HTML tables
# --------------------------------------------------------------------------- #
def _html_to_model(content: bytes):
    from bs4 import BeautifulSoup

    try:
        soup = BeautifulSoup(content, "lxml")
    except Exception:  # pragma: no cover - lxml is a dependency
        soup = BeautifulSoup(content, "html.parser")

    tables = soup.find_all("table")
    if not tables:
        raise ValueError("No <table> elements found in the HTML content")

    model = new_model()
    for i, table in enumerate(tables, start=1):
        rows: List[List[str]] = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["th", "td"])
            if not cells:
                continue
            rows.append([c.get_text(strip=True) for c in cells])
        if rows:
            model["table" if i == 1 else "table_%d" % i] = rows
    if not model:
        raise ValueError("Found <table> elements but no rows to extract")
    return model


# --------------------------------------------------------------------------- #
# Dispatch
# --------------------------------------------------------------------------- #
def to_model(content: bytes, kind: str = "auto", content_type: str = ""):
    """Convert remote ``content`` into a workbook model.

    ``kind`` may be ``'auto'`` (default), ``'rss'``, ``'xml'``, ``'data'`` or
    ``'html'``.
    """
    if kind == "auto":
        kind = detect_kind(content, content_type)

    if kind == "rss":
        return _rss_to_model(content)
    if kind in ("xml", "data"):
        return _xml_to_model(content)
    if kind == "html":
        return _html_to_model(content)
    raise ValueError("Unknown ingest kind %r" % kind)
