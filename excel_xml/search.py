"""Find a file by name across all local drives (Windows-only feature).

``win32api`` (pywin32) is imported lazily so the rest of the tool works on
non-Windows platforms or when pywin32 is not installed.
"""

from __future__ import annotations

import os
from typing import List


def _get_drives() -> List[str]:
    try:
        import win32api
    except ImportError:
        raise ImportError(
            "File search requires pywin32 on Windows. Install it with: pip install pywin32"
        )
    drives = win32api.GetLogicalDriveStrings()
    return [d for d in drives.split("\0") if d]


def file_search(filename: str) -> List[str]:
    """Search every logical drive for ``filename`` and print the matches."""
    try:
        drives = _get_drives()
    except ImportError as exc:
        print(exc)
        return []

    results: List[str] = []
    for drive in drives:
        for root, _dirs, files in os.walk(drive):
            if filename in files:
                results.append(os.path.join(root, filename))

    if results:
        for i, path in enumerate(results, start=1):
            print("%d.%s" % (i, path))
    else:
        print("No file found with name %s" % filename)
    return results
