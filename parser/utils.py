"""
parser/utils.py -- Shared normalisation helpers for ProtocolFormatter's input parser.

All functions in this module are pure (no I/O, no side-effects) and operate on
plain Python strings or simple data structures. They are consumed by both
docx_reader.py and doc_reader.py.

Responsibilities
----------------
- Whitespace normalisation (collapsing runs, stripping non-breaking spaces, etc.)
- Inline formatting marker conversion (python-docx run properties → **bold**/_italic_)
- Callout keyword detection (maps source text phrases to CalloutType values)
- Stopping point detection
- Footnote extraction and renumbering
- Cross-reference normalisation ("Section X.Y" preservation)
- Heading level heuristics for plain-text fallback paths
"""

from __future__ import annotations

import re
import unicodedata
from typing import Optional

from loguru import logger


# ---------------------------------------------------------------------------
# Whitespace normalisation
# ---------------------------------------------------------------------------

# Characters that should be treated as ordinary spaces
_SPACE_CHARS = (
    "\u00A0",   # NO-BREAK SPACE
    "\u200B",   # ZERO WIDTH SPACE
    "\u200C",   # ZERO WIDTH NON-JOINER
    "\u200D",   # ZERO WIDTH JOINER
    "\u2009",   # THIN SPACE
    "\u202F",   # NARROW NO-BREAK SPACE
    "\uFEFF",   # BYTE ORDER MARK / ZERO WIDTH NO-BREAK SPACE
    "\t",       # TAB — normalise to single space
)

_SPACE_TRANS = str.maketrans({ch: " " for ch in _SPACE_CHARS})

# Fancy punctuation → ASCII equivalents (best-effort; preserves meaning)
_FANCY_PUNCT = str.maketrans(
    {
        "\u2018": "'",    # LEFT SINGLE QUOTATION MARK
        "\u2019": "'",    # RIGHT SINGLE QUOTATION MARK
        "\u201C": '"',    # LEFT DOUBLE QUOTATION MARK
        "\u201D": '"',    # RIGHT DOUBLE QUOTATION MARK
        "\u2013": "-",    # EN DASH (in plain text contexts)
        "\u2014": "\u2014",  # EM DASH — preserve as-is (meaningful in headings)
        "\u2026": "...",  # HORIZONTAL ELLIPSIS
    }
)


def normalise_whitespace(text: str) -> str:
    """
    Normalise whitespace in a string extracted from a source document.

    Steps:
        1. Translate non-breaking and zero-width space variants to regular space.
        2. Collapse runs of spaces to a single space.
        3. Strip leading and trailing whitespace.

    Does NOT strip internal newlines; callers that want single-line strings
    should call this first, then handle newlines separately.

    Parameters
    ----------
    text : str
        Raw text from a parsed paragraph or run.

    Returns
    -------
    str
        Normalised text.

    Examples
    --------
    >>> normalise_whitespace("  hello\\u00A0 world  ")
    'hello world'
    >>> normalise_whitespace("\\tfoo\\u200Bbar")
    'foo bar'
    """
    if not text:
        return ""
    text = text.translate(_SPACE_TRANS)
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


def normalise_text(text: str, *, keep_punctuation: bool = True) -> str:
    """
    Full normalisation: whitespace + optional fancy-punctuation conversion.

    Parameters
    ----------
    text : str
    keep_punctuation : bool
        If True (default), preserves em dashes and other meaningful punctuation.
        If False, converts to ASCII equivalents.

    Returns
    -------
    str
    """
    text = normalise_whitespace(text)
    if not keep_punctuation:
        text = text.translate(_FANCY_PUNCT)
    return text


# ---------------------------------------------------------------------------
# Inline formatting marker conversion
# ---------------------------------------------------------------------------

def runs_to_marked_text(runs: list[dict]) -> str:
    """
    Convert a list of run dicts (as produced by docx_reader) into a single
    string with **bold** and _italic_ inline markers.

    Each run dict has the shape::

        {
            "text": str,
            "bold": bool,
            "italic": bool,
        }

    Markers are only applied when the run has non-empty text after
    whitespace normalisation. Adjacent runs with the same formatting are
    merged before marker application.

    Parameters
    ----------
    runs : list[dict]
        List of run dicts.

    Returns
    -------
    str
        Plain string with inline formatting markers.

    Examples
    --------
    >>> runs_to_marked_text([
    ...     {"text": "Add ", "bold": False, "italic": False},
    ...     {"text": "5 µL", "bold": True,  "italic": False},
    ...     {"text": " Buffer A.", "bold": False, "italic": False},
    ... ])
    'Add **5 µL** Buffer A.'
    """
    if not runs:
        return ""

    # Strategy: keep each run's text as-is (do NOT strip internal spaces via
    # normalise_whitespace — that strips trailing spaces off "Add " etc.).
    # Instead normalise only truly invisible characters (zero-width, nbsp).
    _invisible = str.maketrans({
        "\u00A0": " ", "\u200B": "", "\u200C": "", "\u200D": "",
        "\u2009": " ", "\u202F": " ", "\uFEFF": "",
    })

    # Build list of (text, bold, italic) segments, merging adjacent same-format runs
    segments: list[dict] = []
    for run in runs:
        raw = run.get("text", "").translate(_invisible)
        if not raw:
            continue
        bold = bool(run.get("bold"))
        italic = bool(run.get("italic"))
        if segments and segments[-1]["bold"] == bold and segments[-1]["italic"] == italic:
            segments[-1]["text"] += raw
        else:
            segments.append({"text": raw, "bold": bold, "italic": italic})

    # Apply markers. Spaces adjacent to formatted spans belong outside the markers.
    parts: list[str] = []
    for seg in segments:
        t = seg["text"]
        if not (seg["bold"] or seg["italic"]):
            parts.append(t)
            continue

        # Move surrounding spaces outside the marker delimiters
        inner = t.strip(" ")
        if not inner:
            parts.append(t)   # all-space run, keep as-is
            continue
        leading  = t[: len(t) - len(t.lstrip(" "))]
        trailing = t[len(t.rstrip(" ")):]

        if leading:
            parts.append(leading)

        if seg["bold"] and seg["italic"]:
            parts.append(f"**{inner}**")
        elif seg["bold"]:
            parts.append(f"**{inner}**")
        else:
            parts.append(f"_{inner}_")

        if trailing:
            parts.append(trailing)

    return "".join(parts)


def strip_markers(text: str) -> str:
    """
    Remove all **bold** and _italic_ inline markers from a string, leaving
    the plain text content.

    Parameters
    ----------
    text : str

    Returns
    -------
    str

    Examples
    --------
    >>> strip_markers("Add **5 µL** Buffer _A_.")
    'Add 5 µL Buffer A.'
    """
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"_(.+?)_", r"\1", text)
    return text


# ---------------------------------------------------------------------------
# Callout keyword detection
# ---------------------------------------------------------------------------

# Keywords per callout type, ordered by priority (highest first).
# A match on a higher-priority type short-circuits lower types.
_CALLOUT_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    (
        "critical",
        re.compile(
            r"\b(CRITICAL|critical step|do not|must not|protocol will fail|crucial)\b", # See notes in style_guide.yaml for potential variations to make
            re.IGNORECASE,
        ),
    ),
    (
        "caution",
        re.compile(
            r"\b(IMPORTANT|WARNING|CAUTION|hazard|toxic|corrosive|flammable|"
            r"carcinogen|biohazard|radiation|RNase|DNase|keep on ice|avoid freeze|handle gently|avoid freeze/thaw|sensitive to)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "tip",
        re.compile(
            r"\b(tip|optional|shortcut|alternatively|can also|to improve|"
            r"for best results|we recommend|works well)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "note",
        re.compile(
            r"\b(note|see section|refer to|for more information|background|"
            r"version|alternative|species-specific|see also|cross-reference)\b",
            re.IGNORECASE,
        ),
    ),
]

# Explicit prefix patterns that strongly indicate a callout label
_CALLOUT_PREFIX = re.compile(
    r"^(?P<label>IMPORTANT|WARNING|CAUTION|NOTE|TIP|CRITICAL)\s*[:\-–—]\s*",
    re.IGNORECASE,
)


def detect_callout_type(text: str) -> Optional[str]:
    """
    Heuristically determine whether a paragraph should be classified as a
    callout box, and if so, which type.

    Returns the callout_type string ("critical", "caution", "tip", "note")
    or None if the text does not match any callout pattern.

    Priority order: critical > caution > tip > note.

    Parameters
    ----------
    text : str
        Normalised paragraph text (inline markers stripped for matching).

    Returns
    -------
    str | None

    Examples
    --------
    >>> detect_callout_type("IMPORTANT: Keep samples on ice at all times.")
    'caution'
    >>> detect_callout_type("Note: See Section 3.2 for alternative method.")
    'note'
    >>> detect_callout_type("Add 5 µL Buffer A.")
    None
    """
    plain = strip_markers(text).strip()
    if not plain:
        return None

    # Check explicit prefix first (high confidence)
    m = _CALLOUT_PREFIX.match(plain)
    if m:
        label = m.group("label").lower()
        if label in ("important", "warning", "caution"):
            return "caution"
        if label == "critical":
            return "critical"
        if label == "tip":
            return "tip"
        if label == "note":
            return "note"

    # Keyword scan in priority order
    for callout_type, pattern in _CALLOUT_PATTERNS:
        if pattern.search(plain):
            return callout_type

    return None


def strip_callout_prefix(text: str) -> str:
    """
    Remove a leading callout label (e.g. "IMPORTANT: ", "Note: ") from text
    before storing it in the callout.text field.

    The renderer re-adds the label automatically.

    Parameters
    ----------
    text : str

    Returns
    -------
    str

    Examples
    --------
    >>> strip_callout_prefix("IMPORTANT: Keep on ice.")
    'Keep on ice.'
    >>> strip_callout_prefix("Note — See Section 3.2.")
    'See Section 3.2.'
    """
    return _CALLOUT_PREFIX.sub("", text).strip()


# ---------------------------------------------------------------------------
# Stopping point detection
# ---------------------------------------------------------------------------

_STOPPING_POINT_PATTERN = re.compile(
    r"\b(pause point|safe stopping point|can be stored|may be stored|"
    r"samples can be stored|can be left|store at this stage|"
    r"can stop here|overnight at)\b",
    re.IGNORECASE,
)


def is_stopping_point(text: str) -> bool:
    """
    Return True if the text describes a safe pause point in the procedure.

    Parameters
    ----------
    text : str
        Normalised step or paragraph text.

    Returns
    -------
    bool

    Examples
    --------
    >>> is_stopping_point("Samples can be stored at −20 °C overnight.")
    True
    >>> is_stopping_point("Add 5 µL Buffer A.")
    False
    """
    return bool(_STOPPING_POINT_PATTERN.search(strip_markers(text)))


# ---------------------------------------------------------------------------
# Footnote extraction and renumbering
# ---------------------------------------------------------------------------

# Inline footnote reference patterns (superscript numbers or bracketed refs)
_FOOTNOTE_REF_PATTERNS = [
    re.compile(r"\[(\d+)\]"),          # [1], [2], …
    re.compile(r"\((\d+)\)(?=\s|$)"),  # (1), (2) at end or followed by space
]


def extract_footnote_refs(text: str) -> tuple[str, list[int]]:
    """
    Find inline footnote reference markers in text and return the cleaned
    text plus the list of referenced footnote numbers.

    Only bracketed numeric references ([1], [2], …) are extracted; the
    python-docx footnote API provides footnotes separately and is the
    preferred source.

    Parameters
    ----------
    text : str

    Returns
    -------
    tuple[str, list[int]]
        (cleaned_text, [footnote_numbers_referenced])

    Examples
    --------
    >>> extract_footnote_refs("See method described previously [1] and [3].")
    ('See method described previously  and .', [1, 3])
    """
    refs: list[int] = []
    cleaned = text
    for pattern in _FOOTNOTE_REF_PATTERNS:
        for m in pattern.finditer(cleaned):
            try:
                refs.append(int(m.group(1)))
            except ValueError:
                pass
        cleaned = pattern.sub("", cleaned)
    cleaned = normalise_whitespace(cleaned)
    return cleaned, sorted(set(refs))


def renumber_footnotes(footnotes: dict[int, str]) -> list[str]:
    """
    Convert a dict of footnote_number → text into a sequential list of note
    strings for the Protocol.notes field, renumbered from 1.

    Parameters
    ----------
    footnotes : dict[int, str]
        Maps original footnote numbers (as found in the source document) to
        their text content.

    Returns
    -------
    list[str]
        Ordered list of note strings, renumbered 1, 2, 3, …

    Examples
    --------
    >>> renumber_footnotes({3: "Third footnote.", 1: "First footnote."})
    ['First footnote.', 'Third footnote.']
    """
    return [text for _, text in sorted(footnotes.items())]


# ---------------------------------------------------------------------------
# Cross-reference normalisation
# ---------------------------------------------------------------------------

# Matches "Section 3.2", "section 4", "Sec. 2.1", etc.
_XREF_PATTERN = re.compile(
    r"\b[Ss]ec(?:tion|\.)\s*(\d+(?:\.\d+)*)\b"
)


def normalise_cross_references(text: str) -> str:
    """
    Normalise cross-reference phrases to the canonical "Section X.Y" format.

    Handles: "section 3", "Sec. 4.1", "Section 2.3.1" → "Section 3",
    "Section 4.1", "Section 2.3.1".

    Parameters
    ----------
    text : str

    Returns
    -------
    str

    Examples
    --------
    >>> normalise_cross_references("See sec. 3.2 for details.")
    'See Section 3.2 for details.'
    """
    return _XREF_PATTERN.sub(lambda m: f"Section {m.group(1)}", text)

def normalise_centrifuge_units(text: str) -> str:
    """
    Normalise centrifuge speed notation to x g format.

    Converts italic/bold g markers that follow a number:
        12,000*g*  →  12,000 x g
        12,000_g_  →  12,000 x g
        12,000 x g →  12,000 x g
        12,000 xg  →  12,000 x g
    """
    # Remove italic/bold markers around g following a number
    text = re.sub(r'(\d)\s*\*g\*', r'\1 x g', text)
    text = re.sub(r'(\d)\s*_g_', r'\1 x g', text)
    # Normalise plain "x g" and "xg" variants
    text = re.sub(r'(\d)\s*[xX]\s*g\b', r'\1 x g', text)
    return text

def normalise_temperature_units(text: str) -> str:
    """
    Normalise temperature notation to use the proper degree symbol.

    Converts:
        4o C  →  4 °C
        37o C →  37 °C
        -20oC →  -20 °C
        55°C →  55 °C   (normalise spacing)
    """
    # Letter o used as degree symbol before C or F
    text = re.sub(r'(\d)\s*[oO]\s*([CF])\b', r'\1°\2', text)
    # Existing degree symbol but inconsistent spacing, also handles decimal temperatures like 37.5 °C
    text = re.sub(r'(\d+(?:\.\d+)?)\s*°\s*([CF])\b', r'\1 °\2', text)
    return text

# ---------------------------------------------------------------------------
# Heading level heuristics (for plain-text / pandoc fallback paths)
# ---------------------------------------------------------------------------

# Known H1-level section titles used across the compendium templates
_H1_TITLES = frozenset(
    {
        "overview",
        "materials",
        "procedure",
        "notes & variants",
        "notes and variants",
        "references",
        "reagent mixes & recipes",
        "reagent mixes and recipes",
        "prerequisites",
    }
)


def infer_heading_level(text: str, style_name: str = "") -> int:
    """
    Infer the semantic heading level (1–3) of a paragraph.

    Uses a combination of:
        1. The paragraph style name from python-docx (most reliable).
        2. Known H1 section titles from the protocol template.
        3. Text characteristics as a last resort.

    Parameters
    ----------
    text : str
        Normalised paragraph text.
    style_name : str
        python-docx paragraph style name (e.g. "Heading 1", "Heading 2").
        Empty string if not available (e.g. from pandoc plain-text output).

    Returns
    -------
    int
        Inferred heading level: 1, 2, or 3.
        Returns 0 if the paragraph is not a heading.

    Examples
    --------
    >>> infer_heading_level("Overview", "Heading 1")
    1
    >>> infer_heading_level("Sample preparation", "Heading 2")
    2
    >>> infer_heading_level("Reagents", "")
    3
    """
    # Style-name based (highest confidence)
    if style_name:
        s = style_name.lower()
        if "heading 1" in s or s == "h1":
            return 1
        if "heading 2" in s or s == "h2":
            return 2
        if "heading 3" in s or s == "h3":
            return 3
        if "title" in s:
            return 1

    # Known H1 titles
    plain = strip_markers(text).strip().lower()
    if plain in _H1_TITLES:
        return 1

    # Heuristic: short ALL-CAPS line is likely a heading
    if text.isupper() and len(text.split()) <= 8:
        return 2

    return 0


# ---------------------------------------------------------------------------
# Author attribution
# ---------------------------------------------------------------------------

_KNOWN_AUTHOR = re.compile(
    r"\b(Alon\s+Douek|A\.\s*Douek)\b",
    re.IGNORECASE,
)


def extract_author(text: str) -> str:
    """
    Attempt to extract an author name from a text string (e.g. a byline or
    metadata field from the source document).

    Returns "Alon Douek" if the canonical author is detected, otherwise
    returns the raw text stripped of excess whitespace. If text is empty,
    returns the default attribution "ARMI".

    Parameters
    ----------
    text : str

    Returns
    -------
    str

    Examples
    --------
    >>> extract_author("Protocol by Alon Douek, version 2.0")
    'Alon Douek'
    >>> extract_author("Contributed by the ARMI imaging team")
    'ARMI'
    """
    if not text or not text.strip():
        return "ARMI"
    if _KNOWN_AUTHOR.search(text):
        return "Alon Douek"
    # Reject software-generated Office default author names
    _DEFAULT_OFFICE_AUTHORS = re.compile(
        r"^(microsoft office user|windows user|user|owner|author|admin)$",
        re.IGNORECASE,
    )
    if _DEFAULT_OFFICE_AUTHORS.match(text.strip()):
        return "ARMI"
    # Return cleaned raw text; LLM extractor will refine further
    return normalise_whitespace(text)


# ---------------------------------------------------------------------------
# Section type detection
# ---------------------------------------------------------------------------

_COMP_KEYWORDS = re.compile(
    r"\b(bash|python|git|conda|pip|docker|node|npm|script|command|terminal|"
    r"shell|CLI|API|server|database|bioinformatics|pipeline|workflow|"
    r"jupyter|notebook|environment|package|install|deploy)\b",
    re.IGNORECASE,
)

_CODE_BLOCK_PATTERN = re.compile(r"```|^\s{4,}\S", re.MULTILINE)


def infer_section_type(full_text: str) -> str:
    """
    Heuristically infer whether a protocol is "wet_lab" or "computational"
    from its full extracted text.

    Computational indicators: presence of code blocks, bash/Python/git
    commands, or bioinformatics terminology.

    Parameters
    ----------
    full_text : str
        The entire extracted protocol text.

    Returns
    -------
    str
        "computational" or "wet_lab"
    """
    comp_keyword_hits = len(_COMP_KEYWORDS.findall(full_text))
    has_code_blocks = bool(_CODE_BLOCK_PATTERN.search(full_text))

    if has_code_blocks or comp_keyword_hits >= 3:
        logger.debug(
            "infer_section_type: computational "
            "(code_blocks={}, comp_keywords={})",
            has_code_blocks,
            comp_keyword_hits,
        )
        return "computational"

    return "wet_lab"


# ---------------------------------------------------------------------------
# Date normalisation
# ---------------------------------------------------------------------------

_DATE_PATTERNS = [
    # DD/MM/YYYY or DD-MM-YYYY
    re.compile(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})"),
    # Month name: "January 2025", "Jan 2025", "15 Jan 2025"
    re.compile(
        r"(\d{1,2}\s+)?"
        r"(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
        r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
        r"\s+(\d{4})",
        re.IGNORECASE,
    ),
]

_MONTH_MAP = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}


def normalise_date(text: str) -> Optional[str]:
    """
    Extract and normalise a date string to DD/MM/YYYY format.

    Returns None if no recognisable date is found.

    Parameters
    ----------
    text : str

    Returns
    -------
    str | None

    Examples
    --------
    >>> normalise_date("Last updated: 05-03-2024")
    '05/03/2024'
    >>> normalise_date("March 2024")
    '01/03/2024'
    """
    if not text:
        return None

    # DD/MM/YYYY pattern
    m = _DATE_PATTERNS[0].search(text)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        return f"{int(d):02d}/{int(mo):02d}/{y}"

    # Month name pattern
    m = _DATE_PATTERNS[1].search(text)
    if m:
        day_part = (m.group(1) or "").strip()
        month_str = m.group(2)[:3].lower()
        year = m.group(3)
        day = int(day_part) if day_part else 1
        month = _MONTH_MAP.get(month_str, "01")
        return f"{day:02d}/{month}/{year}"

    return None


# ---------------------------------------------------------------------------
# Miscellaneous
# ---------------------------------------------------------------------------

def clean_title(text: str) -> str:
    """
    Clean a raw title string: strip inline markers, normalise whitespace,
    remove trailing punctuation that is unlikely to be intentional.

    Parameters
    ----------
    text : str

    Returns
    -------
    str

    Examples
    --------
    >>> clean_title("**RNA Extraction Protocol:**")
    'RNA Extraction Protocol'
    """
    text = strip_markers(text)
    text = normalise_whitespace(text)
    text = text.rstrip(":;.,")
    return text


def split_into_sentences(text: str) -> list[str]:
    """
    Split a paragraph into individual sentences using a simple heuristic
    suitable for lab protocol text (avoids splitting on abbreviations like
    µL, mM, °C, etc.).

    Parameters
    ----------
    text : str

    Returns
    -------
    list[str]
    """
    # Don't split on common abbreviations
    # Insert sentinel before terminal punctuation only when followed by uppercase
    protected = re.sub(
        r"(?<!\b(?:µL|mL|mM|µM|nM|°C|pH|e\.g|i\.e|vs|et al|Fig|Tab|Ref|Sec|approx|min|max|vol|conc))"
        r"([.!?])\s+(?=[A-Z])",
        r"\1\n",
        text,
    )
    return [s.strip() for s in protected.split("\n") if s.strip()]
