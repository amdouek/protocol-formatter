"""
node_renderer.py -- Python shim for ProtocolFormatter's Node.js renderer.

Serialises a validated Protocol schema instance to JSON, resolves the Node.js
executable and render.js paths from style_guide.yaml, invokes render.js as a
subprocess, and returns the absolute path of the generated .docx file.

The Node.js process receives the JSON payload via stdin and writes the output
.docx file to the path specified in the ``output_path`` field injected by this
module. On success render.js exits 0 and writes {"output": "<path>"} to stdout.
On failure it exits 1 and writes {"error": "...", "detail": "..."} to stderr.

Typical usage
-------------
    from renderer.node_renderer import render_protocol
    from schema import Protocol

    protocol = Protocol(...)
    output_path = render_protocol(protocol, output_dir=Path("output"))
    print(f"Written to: {output_path}")
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional

from loguru import logger

from config import get_config, PACKAGE_ROOT

_PACKAGE_ROOT = PACKAGE_ROOT


def _load_style_guide() -> dict:
    """Thin wrapper preserving the local function name for call sites."""
    return get_config()


def _resolve_node_executable(cfg: dict) -> str:
    """
    Return the Node.js executable path from config.

    Falls back to the system ``node`` if the configured value is the default
    sentinel "node". Raises RendererError if Node.js cannot be found.
    """
    node_exe = cfg.get("paths", {}).get("node_executable", "node")
    return node_exe


def _resolve_render_script(cfg: dict) -> Path:
    """
    Return the absolute path to render.js.

    The path in style_guide.yaml is relative to the package root.
    """
    rel = cfg.get("paths", {}).get("render_script", "renderer/templates/render.js")
    script = _PACKAGE_ROOT / rel
    if not script.exists():
        raise FileNotFoundError(
            f"render.js not found at {script}. "
            "Run Phase 2 of the build to create the renderer templates."
        )
    return script


class RendererError(RuntimeError):
    """Raised when the Node.js renderer subprocess fails."""


def render_protocol(
    protocol,
    output_dir: Optional[Path] = None,
    output_filename: Optional[str] = None,
) -> Path:
    """
    Render a validated Protocol instance to a .docx file via Node.js.

    Parameters
    ----------
    protocol:
        A validated ``schema.Protocol`` instance.
    output_dir:
        Directory in which to write the output file. Defaults to the
        ``output_dir`` value in style_guide.yaml, resolved relative to the
        current working directory.
    output_filename:
        Explicit output filename (without directory). If omitted, the filename
        is derived from the protocol title and version:
        ``<slug>_v<version>.docx``

    Returns
    -------
    Path
        Absolute path of the generated .docx file.

    Raises
    ------
    RendererError
        If Node.js exits non-zero or returns an unexpected response.
    FileNotFoundError
        If render.js or the Node.js executable cannot be located.
    """
    cfg = _load_style_guide()
    node_exe = _resolve_node_executable(cfg)
    render_script = _resolve_render_script(cfg)

    # Resolve output directory
    if output_dir is None:
        default_dir = cfg.get("paths", {}).get("output_dir", "output")
        output_dir = Path(default_dir)
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # Derive output filename
    if output_filename is None:
        slug = _title_to_slug(protocol.title)
        version = (protocol.version or "1.0").replace(".", "_")
        output_filename = f"{slug}_v{version}.docx"

    output_path = output_dir / output_filename

    # Serialise the protocol to JSON, injecting the output_path field
    payload_dict = json.loads(protocol.model_dump_json())
    payload_dict["output_path"] = str(output_path)
    payload_json = json.dumps(payload_dict, ensure_ascii=False, indent=None)

    logger.debug(
        "Invoking Node.js renderer: {} {}",
        node_exe,
        render_script,
    )
    logger.debug(
        "Payload: section_type={}, title={!r}, output={}",
        protocol.section_type,
        protocol.title,
        output_path,
    )

    # Invoke render.js with JSON payload via stdin
    try:
        result = subprocess.run(
            [node_exe, str(render_script)],
            input=payload_json,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=cfg.get("ollama", {}).get("request_timeout_seconds", 120),
        )
    except FileNotFoundError:
        raise RendererError(
            f"Node.js executable not found: '{node_exe}'. "
            "Install Node.js and ensure it is on your PATH, or set "
            "paths.node_executable in configs/style_guide.yaml."
        ) from None
    except subprocess.TimeoutExpired:
        raise RendererError(
            "Node.js renderer timed out. The protocol may be unusually large. "
            "Consider increasing ollama.request_timeout_seconds in style_guide.yaml."
        ) from None

    # Parse stdout (expected: {"output": "<path>"})
    stdout = result.stdout.strip()
    stderr = result.stderr.strip()

    if result.returncode != 0:
        # Attempt to parse structured error from stderr
        error_msg = "Node.js renderer exited with non-zero status."
        detail = stderr
        try:
            err_obj = json.loads(stderr)
            error_msg = err_obj.get("error", error_msg)
            detail = err_obj.get("detail", stderr)
        except (json.JSONDecodeError, AttributeError):
            pass

        logger.error("Renderer error: {}", error_msg)
        if detail:
            logger.error("Detail: {}", detail)

        raise RendererError(
            f"{error_msg}\n\nDetail:\n{detail}"
        )

    # Verify output path from stdout
    try:
        response = json.loads(stdout)
        rendered_path = Path(response["output"])
    except (json.JSONDecodeError, KeyError, TypeError) as exc:
        raise RendererError(
            f"Unexpected renderer stdout: {stdout!r}. "
            f"Expected JSON with 'output' key."
        ) from exc

    if not rendered_path.exists() and str(rendered_path) != "(stdout)":
        raise RendererError(
            f"Renderer reported success but output file does not exist: {rendered_path}"
        )

    logger.success("Protocol rendered: {}", rendered_path)
    return rendered_path


def check_node_available(node_exe: Optional[str] = None) -> tuple[bool, str]:
    """
    Check whether Node.js is available and return its version string.

    Parameters
    ----------
    node_exe:
        Path to the Node.js executable. If None, reads from style_guide.yaml.

    Returns
    -------
    tuple[bool, str]
        (True, version_string) if Node.js is available,
        (False, error_message) otherwise.
    """
    if node_exe is None:
        try:
            cfg = _load_style_guide()
            node_exe = _resolve_node_executable(cfg)
        except Exception as exc:
            return False, str(exc)

    try:
        result = subprocess.run(
            [node_exe, "--version"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode == 0:
            return True, result.stdout.strip()
        return False, result.stderr.strip()
    except FileNotFoundError:
        return False, f"Node.js executable not found: '{node_exe}'"
    except subprocess.TimeoutExpired:
        return False, "Node.js version check timed out."


def check_docx_package_available(node_exe: Optional[str] = None) -> tuple[bool, str]:
    """
    Check whether the ``docx`` npm package is available to the renderer.

    Runs a minimal Node.js require check from the renderer/templates/
    directory first (where render.js will run), then falls back to the
    package root (where npm install is typically run). This handles both
    local node_modules in templates/ and project-root node_modules that
    Node resolves via parent-directory walking.

    Parameters
    ----------
    node_exe:
        Path to the Node.js executable. If None, reads from style_guide.yaml.

    Returns
    -------
    tuple[bool, str]
    """
    if node_exe is None:
        try:
            cfg = _load_style_guide()
            node_exe = _resolve_node_executable(cfg)
        except Exception as exc:
            return False, str(exc)

    # Use require.resolve to find the package wherever Node's module
    # resolution locates it, then read its version from package.json
    # via the resolved path rather than a hardcoded relative path.
    check_script = (
        "try { "
        "const p = require.resolve('docx'); "
        "const fs = require('fs'); "
        "const path = require('path'); "
        "let dir = path.dirname(p); "
        "while (dir !== path.dirname(dir)) { "
        "  const pkg = path.join(dir, 'package.json'); "
        "  if (fs.existsSync(pkg)) { "
        "    const v = JSON.parse(fs.readFileSync(pkg, 'utf8')).version; "
        "    if (v) { process.stdout.write(v); process.exit(0); } "
        "  } "
        "  dir = path.dirname(dir); "
        "} "
        "process.stdout.write('installed'); "
        "process.exit(0); "
        "} catch(e) { "
        "process.stderr.write(e.message); "
        "process.exit(1); "
        "}"
    )

    # Try from the templates directory (where render.js runs), matching
    # the actual resolution context at render time.
    templates_dir = _PACKAGE_ROOT / "renderer" / "templates"
    search_dirs = [templates_dir, _PACKAGE_ROOT]

    for cwd in search_dirs:
        if not cwd.exists():
            continue
        try:
            result = subprocess.run(
                [node_exe, "-e", check_script],
                capture_output=True,
                text=True,
                timeout=10,
                cwd=str(cwd),
            )
            if result.returncode == 0:
                return True, result.stdout.strip()
        except FileNotFoundError:
            return False, f"Node.js executable not found: '{node_exe}'"
        except subprocess.TimeoutExpired:
            return False, "docx package check timed out."

    # Both search directories failed
    last_stderr = result.stderr.strip() if result else ""
    return False, (
        "docx npm package not found. "
        "Run 'npm install docx' in the project root or in renderer/templates/.\n"
        + last_stderr
    )


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def _title_to_slug(title: str) -> str:          # DEVNOTE001: This is a placeholder - need to revise it to be more aesthetic. Make title case, sep words with hyphens, and prepend with a section-protocol prefix (e.g. 3.1-protocol). Need to ensure edge cases are handled (e.g. non-alphanumeric chars, multiple spaces/hyphens, very long titles). Also need to ensure it doesn't produce an empty slug for weird titles (e.g. "!!!" → "untitled").
    """
    Convert a protocol title to a filesystem-safe slug.

    Rules:
        - Lowercase
        - Spaces and hyphens → underscore
        - Remove non-alphanumeric characters (except underscores)
        - Truncate to 60 characters

    Examples
    --------
    >>> _title_to_slug("RNA Extraction from Zebrafish Tissue")
    'rna_extraction_from_zebrafish_tissue'
    >>> _title_to_slug("In Situ Hybridisation (ISH) — wholemount")
    'in_situ_hybridisation_ish_wholemount'
    """
    import re
    slug = title.lower()
    slug = re.sub(r"[\s\-\u2014\u2013]+", "_", slug)   # spaces/dashes → _
    slug = re.sub(r"[^\w]", "", slug)                   # remove non-word chars
    slug = re.sub(r"_+", "_", slug)                     # collapse multiple _
    slug = slug.strip("_")
    return slug[:60]
