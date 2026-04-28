"""
config.py -- Centralised configuration loader for ProtocolFormatter.

Single source of truth for loading and caching style_guide.yaml. All modules
that need configuration should import from here rather than loading the YAML
independently.

Usage
-----
    from config import get_config, PACKAGE_ROOT, STYLE_GUIDE_PATH

    cfg = get_config()                  # cached; only reads disk once
    cfg = get_config(force_reload=True) # re-reads from disk
"""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Optional

import yaml

PACKAGE_ROOT: Path = Path(__file__).resolve().parent
STYLE_GUIDE_PATH: Path = PACKAGE_ROOT / "configs" / "style_guide.yaml"


def _load_from_disk() -> dict:
    """Read and parse style_guide.yaml. Raises FileNotFoundError if missing."""
    if not STYLE_GUIDE_PATH.exists():
        raise FileNotFoundError(
            f"style_guide.yaml not found at {STYLE_GUIDE_PATH}. "
            "Ensure the configs/ directory is present in the package root."
        )
    with STYLE_GUIDE_PATH.open("r", encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


# Module-level cache. Using a mutable container rather than lru_cache so
# that force_reload is straightforward without cache_clear() gymnastics.
_cached_config: Optional[dict] = None


def get_config(force_reload: bool = False) -> dict:
    """
    Return the parsed style_guide.yaml configuration dict.

    The result is cached after the first call. Subsequent calls return the
    same dict without re-reading disk. Pass ``force_reload=True`` to force
    a fresh read (useful in tests).

    Parameters
    ----------
    force_reload : bool
        If True, bypass the cache and re-read from disk.

    Returns
    -------
    dict
        Parsed YAML configuration.

    Raises
    ------
    FileNotFoundError
        If style_guide.yaml does not exist.
    """
    global _cached_config
    if _cached_config is None or force_reload:
        _cached_config = _load_from_disk()
        _verify_assets()
    return _cached_config


def _verify_assets() -> None:
    """
    Verify that critical non-Python assets are resolvable from PACKAGE_ROOT.

    Logs a warning if any expected file is missing. This catches packaging
    issues (e.g. non-editable installs where force-include paths did not
    resolve) early rather than at render time.
    """
    from loguru import logger

    expected = [
        STYLE_GUIDE_PATH,
        PACKAGE_ROOT / "renderer" / "templates" / "render.js",
        PACKAGE_ROOT / "renderer" / "templates" / "lib.js",
        PACKAGE_ROOT / "renderer" / "templates" / "lib_comp.js",
    ]
    for path in expected:
        if not path.exists():
            logger.warning(
                "Expected asset not found: {}. "
                "If running from a non-editable install, ensure the package "
                "was built with hatch and all force-include paths resolved.",
                path,
            )