"""excel_xml - bidirectional Excel <-> XML conversion plus remote feed ingestion.

Public API::

    from excel_xml import (
        excel_to_xml, xml_to_excel,
        remote_to_xml, remote_to_excel,
        check_xml_data, file_search,
    )
"""

from .compare import check_xml_data
from .convert import (
    excel_to_xml,
    remote_to_excel,
    remote_to_xml,
    xml_to_excel,
)
from .search import file_search

__all__ = [
    "excel_to_xml",
    "xml_to_excel",
    "remote_to_xml",
    "remote_to_excel",
    "check_xml_data",
    "file_search",
]

__version__ = "1.0.0"
