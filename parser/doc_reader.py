"""
parser/doc_reader.py -- Legacy .doc fallback parser for ProtocolFormatter.

Converts a .doc (legacy Word 97-2003) file to .docx using pandoc, then
delegates to ``docx_reader.read_docx()`` for structured extraction.

Pandoc is used rather than LibreOffice because:
  - It is lighter-weight and less likely to have GUI/sandbox conflicts.
  - It produces clean .docx output that python-docx can reliably parse.
  - It preserves bold/italic runs, tables, lists, and headings adequately.

If pandoc is not available on the system PATH (or at the configured path in
style_guide.yaml), a clear error is raised with installation instructions.

Typical usage
-------------
    from parser.doc_reader import read_doc
    parsed = read_doc(Path("legacy_protocol.doc"))
    # parsed is a ParsedDocument identical in structure to docx_reader output
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from loguru import logger

from .docx_reader import ParsedDocument, read_docx

from config import get_config


# ---------------------------------------------------------------------------
# Pandoc availability check
# ---------------------------------------------------------------------------

def _get_pandoc_executable() -> str:
    """
    Return the pandoc executable path.

    Checks style_guide.yaml for a ``paths.pandoc_executable`` key first;
    falls back to "pandoc" (assumes it is on the system PATH).
    """
    try:
        cfg = get_config()
        return cfg.get("paths", {}).get("pandoc_executable", "pandoc")
    except Exception:
        return "pandoc"


def check_pandoc_available(pandoc_exe: Optional[str] = None) -> tuple[bool, str]:
    """
    Check whether pandoc is available and return its version string.

    Parameters
    ----------
    pandoc_exe : str | None
        Path to the pandoc executable. If None, resolved from config.

    Returns
    -------
    tuple[bool, str]
        (True, version_string) or (False, error_message)
    """
    if pandoc_exe is None:
        pandoc_exe = _get_pandoc_executable()

    try:
        result = subprocess.run(
            [pandoc_exe, "--version"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode == 0:
            version_line = result.stdout.splitlines()[0] if result.stdout else "pandoc"
            return True, version_line
        return False, result.stderr.strip()
    except FileNotFoundError:
        return False, (
            f"pandoc executable not found: '{pandoc_exe}'. "
            "Install pandoc from https://pandoc.org/installing.html and "
            "ensure it is on your PATH, or set paths.pandoc_executable in "
            "configs/style_guide.yaml."
        )
    except subprocess.TimeoutExpired:
        return False, "pandoc version check timed out."


# ---------------------------------------------------------------------------
# Conversion
# ---------------------------------------------------------------------------

def _convert_doc_to_docx(
    source_path: Path,
    pandoc_exe: str,
    work_dir: Path,
) -> Path:
    """
    Convert a .doc file to .docx using pandoc.

    Parameters
    ----------
    source_path : Path
        Absolute path to the .doc file.
    pandoc_exe : str
        Path to the pandoc executable.
    work_dir : Path
        Temporary working directory for the output .docx.

    Returns
    -------
    Path
        Path to the converted .docx file.

    Raises
    ------
    RuntimeError
        If pandoc exits non-zero.
    """
    output_path = work_dir / (source_path.stem + "_converted.docx")

    cmd = [
        pandoc_exe,
        str(source_path),
        "--output", str(output_path),
        "--to", "docx",
        # Preserve basic formatting; don't wrap lines
        "--wrap=none",
    ]

    logger.debug("Converting .doc → .docx: {}", " ".join(cmd))

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            cwd=str(work_dir),
        )
    except FileNotFoundError:
        raise RuntimeError(
            f"pandoc executable not found at '{pandoc_exe}'. "
            "Install pandoc or update paths.pandoc_executable in style_guide.yaml."
        ) from None
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"pandoc timed out converting '{source_path.name}'. "
            "The file may be corrupt or unusually large."
        ) from None

    if result.returncode != 0:
        stderr = result.stderr.strip()
        raise RuntimeError(
            f"pandoc failed to convert '{source_path.name}' "
            f"(exit code {result.returncode}).\n"
            f"stderr: {stderr}"
        )

    if not output_path.exists():
        raise RuntimeError(
            f"pandoc reported success but output file not found: {output_path}"
        )

    logger.debug(
        "Converted '{}' → '{}' ({} bytes)",
        source_path.name,
        output_path.name,
        output_path.stat().st_size,
    )

    return output_path


# ---------------------------------------------------------------------------
# Main reader
# ---------------------------------------------------------------------------

def read_doc(source_path: Path) -> ParsedDocument:
    """
    Parse a legacy .doc file and return a ``ParsedDocument``.

    Converts the .doc to .docx in a temporary directory using pandoc, then
    delegates to ``docx_reader.read_docx()``. The temporary directory is
    cleaned up automatically.

    Parameters
    ----------
    source_path : Path
        Absolute or relative path to the .doc file.

    Returns
    -------
    ParsedDocument
        Identical in structure to the output of ``docx_reader.read_docx()``.
        The ``source_path`` field on the returned object reflects the
        original .doc path, not the temporary .docx.

    Raises
    ------
    FileNotFoundError
        If ``source_path`` does not exist.
    RuntimeError
        If pandoc fails to convert the file.
    ValueError
        If the converted .docx cannot be parsed by python-docx.
    """
    source_path = Path(source_path).resolve()

    if not source_path.exists():
        raise FileNotFoundError(f"Source file not found: {source_path}")

    suffix = source_path.suffix.lower()
    if suffix not in (".doc", ".docx"):
        logger.warning(
            "doc_reader received unexpected file extension '{}' for '{}'. "
            "Attempting conversion anyway.",
            suffix,
            source_path.name,
        )

    # If it's already a .docx, just pass through (avoids unnecessary conversion)
    if suffix == ".docx":
        logger.debug(
            "'{}' is already .docx — delegating directly to docx_reader.",
            source_path.name,
        )
        return read_docx(source_path)

    pandoc_exe = _get_pandoc_executable()

    # Verify pandoc is available before creating a temp dir
    available, msg = check_pandoc_available(pandoc_exe)
    if not available:
        raise RuntimeError(msg)

    logger.info("Converting legacy .doc: {}", source_path.name)

    with tempfile.TemporaryDirectory(prefix="protocol_formatter_") as tmp_dir:
        work_dir = Path(tmp_dir)

        # Copy source into work_dir to avoid issues with paths containing spaces
        local_copy = work_dir / source_path.name
        shutil.copy2(source_path, local_copy)

        converted_path = _convert_doc_to_docx(local_copy, pandoc_exe, work_dir)
        parsed = read_docx(converted_path)

    # Restore the original source path so downstream modules reference the
    # correct file for duplicate detection and error messages
    parsed.source_path = source_path

    logger.info(
        "Parsed legacy .doc '{}': {} paragraphs, {} tables, {} footnotes",
        source_path.name,
        len(parsed.paragraphs),
        len(parsed.tables),
        len(parsed.footnotes),
    )

    return parsed
