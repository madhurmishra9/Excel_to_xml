"""Canonical in-memory workbook model and shared value helpers.

A *workbook model* is an ``OrderedDict`` mapping a sheet name to its rows::

    OrderedDict[str, list[list[str]]]

i.e. ``model[sheet_name]`` is a list of rows and each row is a list of cell
*strings*.  Every reader in this package produces this structure and every
writer consumes it, so both conversion directions share the same code.
"""

from __future__ import annotations

import datetime
import re
from collections import OrderedDict
from typing import Dict, List

# Type alias for readability.
Model = "OrderedDict[str, List[List[str]]]"

DATE_FORMAT = "%m/%d/%Y"

# Excel limits sheet names to 31 characters and forbids a handful of chars.
_MAX_SHEET_NAME = 31
_INVALID_SHEET_CHARS = re.compile(r"[\[\]\*\?:/\\]")

# A valid XML element name (simplified): starts with a letter/underscore and is
# followed by letters, digits, hyphens, underscores or dots.
_XML_NAME_START = re.compile(r"[A-Za-z_]")
_XML_NAME_CHAR = re.compile(r"[A-Za-z0-9_.\-]")


def new_model() -> "OrderedDict[str, List[List[str]]]":
    """Return an empty workbook model."""
    return OrderedDict()


def stringify_value(value) -> str:
    """Convert any cell value coming out of Excel into a stable text form.

    ``None`` becomes an empty string, ``datetime``/``date`` are formatted as
    ``mm/dd/YYYY`` (matching the original tool's behaviour) and everything else
    is passed through ``str``.
    """
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime(DATE_FORMAT)
    return str(value)


def safe_tag(name: str) -> str:
    """Sanitise an arbitrary string into a valid XML element tag.

    Invalid characters are replaced with ``_`` and a leading underscore is
    added when the name would otherwise start with an invalid character (e.g.
    a digit).  The original, unsanitised name is preserved separately by the
    XML writer via a ``name`` attribute.
    """
    if not name:
        return "_"
    chars = []
    first = name[0]
    if _XML_NAME_START.match(first):
        chars.append(first)
    elif _XML_NAME_CHAR.match(first):
        # Valid inside a name but not as the first char (e.g. a digit): keep it
        # and prefix an underscore.
        chars.append("_" + first)
    else:
        chars.append("_")
    for ch in name[1:]:
        chars.append(ch if _XML_NAME_CHAR.match(ch) else "_")
    tag = "".join(chars)
    # XML names must not start with the reserved sequence "xml" (any case).
    if tag[:3].lower() == "xml":
        tag = "_" + tag
    return tag


def safe_sheet_name(name: str, existing: Dict[str, int]) -> str:
    """Return an Excel-safe, unique worksheet name.

    Strips characters Excel forbids, truncates to 31 chars and de-duplicates by
    appending ``_2``, ``_3`` ... ``existing`` is a mutable dict used to track
    names already handed out across calls.
    """
    cleaned = _INVALID_SHEET_CHARS.sub("_", name).strip() or "Sheet"
    cleaned = cleaned[:_MAX_SHEET_NAME]
    candidate = cleaned
    n = 1
    while candidate.lower() in existing:
        n += 1
        suffix = "_" + str(n)
        candidate = cleaned[: _MAX_SHEET_NAME - len(suffix)] + suffix
    existing[candidate.lower()] = 1
    return candidate
