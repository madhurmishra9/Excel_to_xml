"""Fetch raw bytes from a remote source.

Supports plain HTTP(S) URLs via ``requests`` and authenticated cloud storage
via optional, lazily imported SDKs:

* ``s3://bucket/key``                      -> boto3
* ``gs://bucket/key``                      -> google-cloud-storage
* ``gdrive://<file-id>`` or a Drive URL    -> gdown (public files)
* ``az://container/blob`` or a
  ``*.blob.core.windows.net`` URL          -> azure-storage-blob

Cloud SDKs are imported only when the matching scheme is used, so the core tool
runs without them installed.  Credentials are taken from each provider's
standard environment / config (documented in the README), never hard-coded.
"""

from __future__ import annotations

import re
from typing import Tuple
from urllib.parse import urlparse

DEFAULT_TIMEOUT = 30
_USER_AGENT = "excel_xml/1.0 (+https://github.com/)"

_DRIVE_ID_RE = re.compile(r"/d/([A-Za-z0-9_-]+)")


def _missing(pkg: str, scheme: str) -> ImportError:
    return ImportError(
        "Fetching %s sources requires the %r package. "
        "Install the cloud extras: pip install -r requirements-cloud.txt" % (scheme, pkg)
    )


def _fetch_http(url: str) -> Tuple[bytes, str]:
    try:
        import requests
    except ImportError:  # pragma: no cover - requests is a core dependency
        raise _missing("requests", "http(s)")
    resp = requests.get(
        url, timeout=DEFAULT_TIMEOUT, headers={"User-Agent": _USER_AGENT}
    )
    resp.raise_for_status()
    content_type = resp.headers.get("Content-Type", "")
    return resp.content, content_type


def _fetch_s3(parsed) -> Tuple[bytes, str]:
    try:
        import boto3
    except ImportError:
        raise _missing("boto3", "s3://")
    bucket = parsed.netloc
    key = parsed.path.lstrip("/")
    obj = boto3.client("s3").get_object(Bucket=bucket, Key=key)
    return obj["Body"].read(), obj.get("ContentType", "")


def _fetch_gs(parsed) -> Tuple[bytes, str]:
    try:
        from google.cloud import storage
    except ImportError:
        raise _missing("google-cloud-storage", "gs://")
    bucket = parsed.netloc
    key = parsed.path.lstrip("/")
    blob = storage.Client().bucket(bucket).blob(key)
    return blob.download_as_bytes(), blob.content_type or ""


def _fetch_gdrive(source: str, parsed) -> Tuple[bytes, str]:
    try:
        import gdown
    except ImportError:
        raise _missing("gdown", "gdrive://")
    # Accept gdrive://<id>, a /d/<id>/ URL, or an ?id=<id> URL.
    file_id = parsed.netloc or parsed.path.lstrip("/")
    m = _DRIVE_ID_RE.search(source)
    if m:
        file_id = m.group(1)
    import io

    buf = io.BytesIO()
    gdown.download(id=file_id, output=buf, quiet=True)
    return buf.getvalue(), ""


def _fetch_azure(source: str, parsed) -> Tuple[bytes, str]:
    try:
        from azure.storage.blob import BlobClient
    except ImportError:
        raise _missing("azure-storage-blob", "az://")
    import os

    conn = os.environ.get("AZURE_STORAGE_CONNECTION_STRING")
    if parsed.scheme == "az":
        if not conn:
            raise RuntimeError(
                "AZURE_STORAGE_CONNECTION_STRING must be set for az:// sources"
            )
        container = parsed.netloc
        blob_name = parsed.path.lstrip("/")
        client = BlobClient.from_connection_string(conn, container, blob_name)
    else:  # full https blob URL
        client = BlobClient.from_blob_url(source)
    data = client.download_blob().readall()
    return data, ""


def fetch(source: str) -> Tuple[bytes, str]:
    """Return ``(content_bytes, content_type)`` for a URL or cloud URI."""
    parsed = urlparse(source)
    scheme = parsed.scheme.lower()

    if scheme in ("http", "https"):
        if "blob.core.windows.net" in parsed.netloc:
            return _fetch_azure(source, parsed)
        if "drive.google.com" in parsed.netloc:
            return _fetch_gdrive(source, parsed)
        return _fetch_http(source)
    if scheme == "s3":
        return _fetch_s3(parsed)
    if scheme == "gs":
        return _fetch_gs(parsed)
    if scheme == "gdrive":
        return _fetch_gdrive(source, parsed)
    if scheme == "az":
        return _fetch_azure(source, parsed)

    raise ValueError("Unsupported source scheme %r in %s" % (scheme, source))
