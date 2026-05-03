"""
extractor/rule_extractor.py -- Heuristic fallback extractor for ProtocolFormatter.

Produces a schema-valid Protocol from a ParsedDocument using deterministic
rules, with no LLM call. Always succeeds, though fields may be incomplete
or approximate. Used when all LLM extraction attempts fail validation.

Design philosophy
-----------------
This extractor is intentionally conservative: it is better to produce a
sparse but valid schema that a human can fill in during the --review step
than to guess incorrectly and produce plausible-looking but wrong content.

Rules applied (in order)
------------------------
1. Title       — first H1 heading, or document core property, or filename stem
2. Author      — from ParsedDocument.author (already extracted by docx_reader)
3. Date        — from ParsedDocument.date
4. Section type — from ParsedDocument.section_type (parser inference)
5. Overview    — first non-heading, non-list paragraph after the title; or
                 first paragraph of the document
6. Materials   — tables directly following a "Reagents" or "Equipment" H3 heading
7. Procedure   — all H2 headings become ProcedureSections; paragraphs and list
                 items beneath each heading become ActionSteps, with callout
                 detection applied to each paragraph
8. Notes       — numbered list items under "Notes & variants" heading, plus
                 all document footnotes
9. References  — paragraphs under "References" heading
10. Mix tables — tables following a "mix" / "recipe" / "master mix" heading
11. Prerequisites — paragraphs under "Prerequisites" heading (comp protocols)
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from loguru import logger

from config import get_config

# Known acronyms that should remain fully uppercase in titles
_ACRONYMS = frozenset({
    "RNA", "DNA", "PCR", "RT", "FISH", "ISH", "FACS", "GFP", "RFP",
    "YFP", "CFP", "mRNA", "gRNA", "sgRNA", "crRNA", "siRNA", "shRNA",
    "ChIP", "ATAC", "HCR", "smFISH", "CRISPR", "LNP", "IHC", "IF",
    "PBS", "BSA", "EDTA", "DMSO", "EtOH", "dH2O", "ddH2O",
})

# Small words that should not be capitalised unless first or last
_LOWERCASE_WORDS = frozenset({
    "a", "an", "the", "and", "but", "or", "for", "nor", "on", "at",
    "to", "by", "in", "of", "up", "as", "is",
})

def _title_case(text: str) -> str:
    """
    Apply title case to a protocol title, preserving known acronyms and
    applying standard English title-case rules (small words lowercase
    unless first or last word).
    """
    words = text.split()
    result = []
    for i, word in enumerate(words):
        upper = word.upper()
        if upper in _ACRONYMS:
            result.append(upper)
        elif i == 0 or i == len(words) - 1:
            result.append(word.capitalize())
        elif word.lower() in _LOWERCASE_WORDS:
            result.append(word.lower())
        else:
            result.append(word.capitalize())
    return " ".join(result)

def _load_cfg() -> dict:
    """Thin wrapper for backward-compatible internal calls."""
    try:
        return get_config()
    except FileNotFoundError:
        return {}


# ---------------------------------------------------------------------------
# Heading classification helpers
# ---------------------------------------------------------------------------

_H1_CANONICAL = {
    "overview": "overview",
    "materials": "materials",
    "procedure": "procedure",
    "notes & variants": "notes",
    "notes and variants": "notes",
    "notes & variants:": "notes",
    "references": "references",
    "reagent mixes & recipes": "mixes",
    "reagent mixes and recipes": "mixes",
    "prerequisites": "prerequisites",
}


def _h1_role(text: str) -> Optional[str]:
    """Return the canonical role of an H1 heading, or None."""
    return _H1_CANONICAL.get(text.strip().lower())


# ---------------------------------------------------------------------------
# Step and callout detection
# ---------------------------------------------------------------------------

def _make_action_step(text: str) -> dict:
    return {"step_type": "action", "text": text, "children": []}


def _make_stopping_point(text: str) -> dict:
    return {"step_type": "stopping_point", "text": text, "children": []}


def _make_callout(callout_type: str, text: str) -> dict:
    from parser.utils import strip_callout_prefix
    return {
        "callout_type": callout_type,
        "text": strip_callout_prefix(text),
    }


def _classify_paragraph(para) -> str:
    """
    Classify a ParsedParagraph as one of:
      "callout_critical", "callout_caution", "callout_tip", "callout_note",
      "stopping_point", "step", or "prose"

    Parameters
    ----------
    para : ParsedParagraph

    Returns
    -------
    str
    """
    from parser.utils import detect_callout_type, is_stopping_point
    ct = detect_callout_type(para.raw_text)
    if ct:
        return f"callout_{ct}"
    if is_stopping_point(para.raw_text):
        return "stopping_point"
    if para.is_list_item:
        return "step"
    return "prose"


# ---------------------------------------------------------------------------
# Main heuristic extractor
# ---------------------------------------------------------------------------

def extract_protocol_heuristic(
    parsed_document,
    source_filename: Optional[str] = None,
    section_number_hint: Optional[str] = None,
    cfg: Optional[dict] = None,
):
    """
    Extract a Protocol from a ParsedDocument using deterministic heuristics.

    Always returns a schema-valid Protocol. Fields that cannot be determined
    are left at their schema defaults (empty lists, "ARMI" author, etc.).

    Parameters
    ----------
    parsed_document : ParsedDocument
    source_filename : str | None
    section_number_hint : str | None
    cfg : dict | None

    Returns
    -------
    Protocol
    """
    from schema import (
        Protocol, SectionType, ProcedureSection, ActionStep, StoppingPoint,
        Callout, CalloutType, MaterialsTable, MixTable, TableRow, Prerequisites, Step,
    )
    from parser.utils import normalise_whitespace, renumber_footnotes

    if cfg is None:
        cfg = _load_cfg()

    doc = parsed_document
    paras = doc.paragraphs
    tables = doc.tables
    fname = source_filename or doc.source_path.name

    logger.debug("Heuristic extraction from '{}'", fname)

    # ------------------------------------------------------------------
    # 1. Title
    # ------------------------------------------------------------------
    title = doc.title or ""
    if not title:
        for p in paras:
            if p.heading_level == 1 and not _h1_role(p.raw_text):
                title = p.raw_text
                break
    if not title:
        raw = Path(fname).stem.replace("_", " ").replace("-", " ")
        title = _title_case(raw)

    # ------------------------------------------------------------------
    # 2. Metadata from parsed document
    # ------------------------------------------------------------------
    author = doc.author or "ARMI"
    date = doc.date
    section_type_str = doc.section_type  # "wet_lab" or "computational"
    section_type = (
        SectionType.COMPUTATIONAL
        if section_type_str == "computational"
        else SectionType.WET_LAB
    )

    # ------------------------------------------------------------------
    # 3. Section number and name
    # ------------------------------------------------------------------
    sections_map = cfg.get("sections", {})
    section_number = section_number_hint or ""
    section_name = ""
    if section_number and section_number in sections_map:
        section_name = sections_map[section_number]
    elif section_type == SectionType.COMPUTATIONAL:
        section_number = section_number or "9"
        section_name = sections_map.get("9", "COMPUTATIONAL METHODS & BIOINFORMATICS")
    else:
        section_number = section_number or ""
        section_name = ""

    # ------------------------------------------------------------------
    # 4. Segment paragraphs by H1 section role
    # ------------------------------------------------------------------
    # Build a segment map: role → list of ParsedParagraph
    segments: dict[str, list] = {
        "pre_overview": [],
        "overview": [],
        "prerequisites": [],
        "materials": [],
        "procedure": [],
        "notes": [],
        "references": [],
        "mixes": [],
        "other": [],
    }
    current_role = "pre_overview"

    for para in paras:
        if para.heading_level == 1:
            role = _h1_role(para.raw_text)
            if role:
                current_role = role
                continue  # don't add the heading itself to the segment
            else:
                # Could be the document title — skip it
                continue
        segments[current_role].append(para)

    # ------------------------------------------------------------------
    # 5. Overview
    # ------------------------------------------------------------------
    overview_paras = segments["overview"] or segments["pre_overview"]
    from parser.utils import is_stopping_point
    overview_text = "\n\n".join(
        p.text for p in overview_paras
        if p.heading_level == 0
        and not p.is_list_item
        and not is_stopping_point(p.raw_text)
        and len(p.raw_text) > 20
    ).strip()
    if not overview_text:
        # Absolute fallback: first non-heading, non-list, non-stopping-point
        # paragraph in the document that reads like a description.
        # DEVNOTE-038: restrict the search to paragraphs that appear BEFORE
        # the first list item / numbered step. Without this bound, terminal
        # instructions at the end of the document (e.g. "store at -80 °C")
        # can be captured as the overview when no explicit Overview section
        # is present.
        first_step_idx = next(
            (i for i, p in enumerate(paras) if p.is_list_item),
            len(paras),
        )
        for p in paras[:first_step_idx]:
            if (p.heading_level == 0
                    and not p.is_list_item
                    and not is_stopping_point(p.raw_text)
                    and len(p.raw_text) > 40):
                overview_text = p.text
                break
    # If still nothing suitable found, leave blank rather than populate
    # with nonsense — the --review step will flag the empty field
    if not overview_text:
        overview_text = ""

    # ------------------------------------------------------------------
    # 6. Prerequisites (computational only)
    # ------------------------------------------------------------------
    prerequisites = None
    if section_type == SectionType.COMPUTATIONAL and segments["prerequisites"]:
        sw = acc = dep = None
        for p in segments["prerequisites"]:
            lower = p.raw_text.lower()
            if lower.startswith("software"):
                sw = re.sub(r"^software[:\s]+", "", p.raw_text, flags=re.IGNORECASE)
            elif lower.startswith("access"):
                acc = re.sub(r"^access[:\s]+", "", p.raw_text, flags=re.IGNORECASE)
            elif lower.startswith("depend"):
                dep = re.sub(r"^depend\w*[:\s]+", "", p.raw_text, flags=re.IGNORECASE)
        prerequisites = Prerequisites(software=sw, access=acc, dependencies=dep)

    # ------------------------------------------------------------------
    # 7. Materials tables (wet_lab only)
    # ------------------------------------------------------------------
    materials: list[MaterialsTable] = []
    if section_type == SectionType.WET_LAB:
        materials = _extract_materials(segments["materials"], tables)

    # ------------------------------------------------------------------
    # 8. Procedure
    # ------------------------------------------------------------------
    procedure = _extract_procedure(segments["procedure"])
    if not procedure:
        # Absolute fallback: one section containing all list items
        all_steps: list[Step] = [
                ActionStep(text=p.text, children=[])
                for p in paras
                if p.is_list_item and p.heading_level == 0
            ]
        if all_steps:
            procedure = [
                ProcedureSection(
                    heading="Procedure",
                    preamble=None,
                    steps=all_steps,
                    callouts=[],
                )
            ]
        else:
            # Last resort: one placeholder section
            procedure = [
                ProcedureSection(
                    heading="Procedure",
                    preamble=None,
                    steps=[ActionStep(text="[Procedure steps — review required]", children=[])],
                    callouts=[],
                )
            ]

    # ------------------------------------------------------------------
    # 9. Notes (including footnotes)
    # ------------------------------------------------------------------
    notes: list[str] = []
    for p in segments["notes"]:
        if p.is_list_item or (p.heading_level == 0 and p.raw_text.strip()):
            notes.append(p.text)
    # Append document footnotes
    if doc.footnotes:
        footnote_texts = renumber_footnotes(doc.footnotes)
        notes.extend(footnote_texts)

    # ------------------------------------------------------------------
    # 10. References
    # ------------------------------------------------------------------
    references: list[str] = [
        p.text for p in segments["references"]
        if p.heading_level == 0 and p.raw_text.strip()
    ]

    # ------------------------------------------------------------------
    # 11. Mix tables (wet_lab only)
    # ------------------------------------------------------------------
    mix_tables: list[MixTable] = []
    if section_type == SectionType.WET_LAB:
        mix_tables = _extract_mix_tables(segments["mixes"], tables)

    # ------------------------------------------------------------------
    # Assemble and validate
    # ------------------------------------------------------------------
    protocol = Protocol(
        title=title,
        subtitle=None,
        author=author,
        section_type=section_type,
        section_number=section_number,
        section_name=section_name,
        version="1.0",
        date=date,
        overview=overview_text,
        materials=materials,
        prerequisites=prerequisites,
        procedure=procedure,
        notes=notes,
        references=references,
        mix_tables=mix_tables,
        source_filename=fname,
    )

    logger.debug(
        "Heuristic extraction complete: {} sections, {} materials tables, "
        "{} notes, {} references",
        len(procedure),
        len(materials),
        len(notes),
        len(references),
    )

    return protocol


# ---------------------------------------------------------------------------
# Procedure extraction helpers
# ---------------------------------------------------------------------------

def _extract_procedure(procedure_paras: list) -> list:
    """
    Extract ProcedureSection objects from the list of paragraphs under the
    Procedure H1 heading.

    H2 headings → new ProcedureSection.
    List items / prose paragraphs → ActionStep or StoppingPoint or Callout.
    """
    from schema import ProcedureSection, ActionStep, StoppingPoint, Callout, CalloutType, Step
    from parser.utils import detect_callout_type, strip_callout_prefix

    if not procedure_paras:
        return []

    sections = []
    current_heading = "Procedure"
    current_preamble: Optional[str] = None
    current_steps = []
    current_callouts = []

    def _flush():
        if current_steps or current_callouts:
            sections.append(
                ProcedureSection(
                    heading=current_heading,
                    preamble=current_preamble,
                    steps=current_steps[:],
                    callouts=current_callouts[:],
                )
            )

    for para in procedure_paras:
        if para.heading_level == 2:
            _flush()
            current_heading = para.raw_text
            current_preamble = None
            current_steps = []
            current_callouts = []
            continue

        if para.heading_level == 3:
            # H3 inside procedure — treat as sub-section preamble
            continue

        classification = _classify_paragraph(para)

        if classification.startswith("callout_"):
            ct = classification.replace("callout_", "")
            current_callouts.append(
                Callout(callout_type=CalloutType(ct), text=_strip_callout_text(para.text))
            )

        elif classification == "stopping_point":
            current_steps.append(StoppingPoint(text=para.text, children=[]))

        elif classification == "step":
            # Check if a list item contains an inline CRITICAL/CRUCIAL statement
            # and if so, also emit a callout for it
            inline_ct = detect_callout_type(para.raw_text)
            if inline_ct in ("critical", "caution"):
                # Only promote to callout if not already covered in current_callouts
                already_present = any(
                    c.callout_type.value == inline_ct for c in current_callouts
                )
                if not already_present:
                    current_callouts.append(
                        Callout(
                            callout_type=CalloutType(inline_ct),
                            text=strip_callout_prefix(para.raw_text),
                        )
                    )
            step = ActionStep(text=para.text, children=[])
            if para.list_level <= 1 or not current_steps:
                current_steps.append(step)
            else:
                _attach_child(current_steps, step, para.list_level)

        elif classification == "prose":
            # Prose before first step in a section → preamble
            if not current_steps and current_preamble is None:
                current_preamble = para.text
            else:
                # Prose after steps: treat as a note-like action step
                current_steps.append(ActionStep(text=para.text, children=[]))

    _flush()
    return sections


def _attach_child(steps: list, child, list_level: int) -> None:
    """
    Attach a child step to the appropriate parent based on list nesting level.
    Traverses to the deepest reachable child of the last step.
    """
    if not steps:
        steps.append(child)
        return

    target = steps[-1]
    depth = 1
    while depth < list_level - 1 and target.children:
        target = target.children[-1]
        depth += 1

    target.children.append(child)


def _strip_callout_text(text: str) -> str:
    """Remove callout label prefix from text before storing."""
    from parser.utils import strip_callout_prefix
    return strip_callout_prefix(text)


# ---------------------------------------------------------------------------
# Materials extraction helpers
# ---------------------------------------------------------------------------

def _pad_cells(cells: list[str], width: int) -> list[str]:
    """Pad a cell list with em-dashes or trim to exactly ``width`` elements."""
    if len(cells) < width:
        return cells + ["\u2014"] * (width - len(cells))
    return cells[:width]


def _extract_materials(material_paras: list, all_tables) -> list:
    """
    Extract MaterialsTable objects from paragraphs and tables in the Materials
    section.

    H3 headings ("Reagents", "Equipment") become table headings.
    Tables at positions following each H3 are assigned to that heading.
    """
    from schema import MaterialsTable, TableRow

    if not material_paras and not all_tables:
        return []

    result = []

    # Collect H3 headings in document order for table assignment
    headings_seen = [
        p.raw_text for p in material_paras if p.heading_level == 3
    ]

    # Assign tables to headings by document order
    headings_seen = []
    for para in material_paras:
        if para.heading_level == 3:
            headings_seen.append(para.raw_text)

    # Use all tables that have ≥3 columns as materials tables
    mat_tables = [t for t in all_tables if t.rows and len(t.rows[0]) >= 2]

    for i, table in enumerate(mat_tables):
        heading = headings_seen[i] if i < len(headings_seen) else "Reagents"
        rows = [
            TableRow(
                cells=_pad_cells([c.strip() or "\u2014" for c in row[:3]], 3),
                bold=False,
            )
            for row in table.rows
            if not _is_header_row(row)
        ]
        if rows:
            result.append(MaterialsTable(heading=heading, rows=rows))

    return result


def _is_header_row(cells: list[str]) -> bool:
    """
    Heuristic: a row is a header if all cells are short and contain common
    header words.
    """
    header_words = {"reagent", "supplier", "cat", "equipment", "notes", "catalogue"}
    return all(
        any(w in cell.lower() for w in header_words)
        for cell in cells
        if cell.strip()
    )


# ---------------------------------------------------------------------------
# Mix table extraction helpers
# ---------------------------------------------------------------------------

def _extract_mix_tables(mix_paras: list, all_tables) -> list:
    """
    Extract MixTable objects from the post-References section.

    H3 headings → table headings.
    2-column tables → MixTable rows.
    """
    from schema import MixTable, TableRow

    if not mix_paras:
        return []

    result = []
    headings_seen = [
        p.raw_text for p in mix_paras if p.heading_level == 3
    ]

    two_col_tables = [
        t for t in all_tables
        if t.rows and len(t.rows[0]) == 2
    ]

    for i, table in enumerate(two_col_tables):
        heading = headings_seen[i] if i < len(headings_seen) else "Reagent Mix"
        rows = []
        for row in table.rows:
            if _is_header_row(row):
                continue
            cells = _pad_cells([c.strip() or "\u2014" for c in row[:2]], 2)
            first = cells[0] if cells else ""
            bold = bool(re.match(r"^(total|incubate)", first, re.IGNORECASE))
            rows.append(TableRow(cells=cells, bold=bold))
        if rows:
            result.append(MixTable(heading=heading, rows=rows))

    return result
