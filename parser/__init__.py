"""
parser -- Input parsing package for ProtocolFormatter.

Public API
----------
    read_document(path)   → ParsedDocument
        Dispatch helper: routes .docx to docx_reader, .doc to doc_reader.

    ParsedDocument        — output dataclass
    ParsedParagraph       — paragraph-level dataclass
    ParsedTable           — table-level dataclass

    check_pandoc_available()   → (bool, str)
    check_node_available()     — re-exported from renderer for convenience
"""

from __future__ import annotations

from pathlib import Path

from loguru import logger

from .docx_reader import ParsedDocument, ParsedParagraph, ParsedTable, read_docx
from .doc_reader import check_pandoc_available, read_doc


def read_document(source_path: Path) -> ParsedDocument:
    """
    Parse a protocol source file and return a ``ParsedDocument``.

    Dispatches to the appropriate reader based on file extension:
        .docx  → docx_reader.read_docx()
        .doc   → doc_reader.read_doc()  (pandoc conversion then docx_reader)

    Parameters
    ----------
    source_path : Path
        Path to the source .docx or .doc file.

    Returns
    -------
    ParsedDocument

    Raises
    ------
    FileNotFoundError
        If the file does not exist.
    ValueError
        If the file extension is not supported.
    RuntimeError
        If pandoc is unavailable for a .doc file.
    """
    source_path = Path(source_path)
    suffix = source_path.suffix.lower()

    if suffix == ".docx":
        return read_docx(source_path)
    elif suffix == ".doc":
        return read_doc(source_path)
    else:
        raise ValueError(
            f"Unsupported file type: '{suffix}'. "
            "ProtocolFormatter accepts .docx and .doc files only."
        )


__all__ = [
    "read_document",
    "read_docx",
    "read_doc",
    "ParsedDocument",
    "ParsedParagraph",
    "ParsedTable",
    "check_pandoc_available",
]
