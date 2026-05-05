"""
extractor/prompts.py -- Prompt templates and style guide context for ProtocolFormatter.

This module is the single source of truth for everything that is injected into
the Ollama request. All prompt construction lives here so that prompt iteration
does not require changes to llm_extractor.py.

Architecture
------------
The extraction prompt is assembled from three layers:

  1. SYSTEM_PROMPT  — Fixed instructions that define the LLM's role, output
                      format contract, and JSON schema. Sent as the system
                      message in the Ollama messages array.

  2. Style guide context block — A concise, token-efficient summary of the
                      editorial conventions extracted from style_guide.yaml.
                      Injected at the top of the user message.

  3. Document content block — The ParsedDocument.full_text, clearly
                      delimited, injected after the style guide context.

The user message ends with an explicit extraction instruction that reminds the
model to output only valid JSON matching the specified schema.

JSON schema contract
--------------------
The LLM is instructed to produce a JSON object matching the Protocol Pydantic
model. The schema is embedded verbatim in the system prompt as a compact
TypeScript-style type definition (more token-efficient than JSON Schema).
Pydantic validates the output in llm_extractor.py; the prompt does not need to
enumerate every edge case — validation handles that.

DEVNOTE001 note
------------
Single-pass extraction. The full document text is passed in one request.
If the document is very long, the prompt + completion may approach
Qwen2.5:14b's 32k token limit. The token budget is estimated before sending
and a warning is logged if it exceeds 24k tokens (leaving ~8k for completion).
"""

from __future__ import annotations

from typing import Optional

from loguru import logger

from config import get_config


# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """\
You are a precise scientific document parser. Your task is to extract the \
structured content of a laboratory protocol from raw document text and return \
it as a single valid JSON object. You must output ONLY the JSON object — no \
preamble, no explanation, no markdown code fences.

## Output schema

Output a JSON object with exactly these fields:

{
  "title": string,                    // Full protocol title
  "subtitle": string | null,          // Brief subtitle, or null
  "author": string,                   // Author name; default "ARMI" if unknown
  "section_type": "wet_lab" | "computational",
  "section_number": string,           // e.g. "3.2" or "9.1"
  "section_name": string,             // e.g. "MOLECULAR BIOLOGY"
  "version": string,                  // e.g. "1.0"
  "date": string | null,              // DD/MM/YYYY or null
  "overview": string,                 // 1-3 sentences; supports **bold** _italic_
  "materials": MaterialsTable[],      // wet_lab only; [] for computational
  "prerequisites": Prerequisites | null, // computational only; null for wet_lab
  "procedure": ProcedureSection[],    // min 1 element
  "notes": string[],                  // numbered notes; footnotes go here
  "references": string[],
  "mix_tables": MixTable[],           // wet_lab only; [] for computational
  "source_filename": string | null
}

Where:

MaterialsTable = {
  "heading": string,                  // e.g. "Reagents"
  "rows": TableRow[]
}

TableRow = {
  "cells": string[],
  "bold": boolean                     // true for Total/Incubate rows
}

Prerequisites = {
  "software": string | null,
  "access": string | null,
  "dependencies": string | null
}

ProcedureSection = {
  "heading": string,                  // H2-level step group heading
  "preamble": string | null,          // intro paragraph before steps (comp only)
  "steps": Step[],
  "callouts": Callout[]
}

Step = ActionStep | StoppingPoint

ActionStep = {
  "step_type": "action",
  "text": string,                     // supports **bold** _italic_
  "code": string | null,              // code block content (computational only); null for wet_lab
  "code_language": string | null,     // e.g. "bash", "python", "R"; null if unknown
  "children": Step[]                  // nested sub-steps; arbitrary depth
}

StoppingPoint = {
  "step_type": "stopping_point",
  "text": string,
  "children": []
}

Callout = {
  "callout_type": "critical" | "caution" | "tip" | "note",
  "text": string,                     // do NOT include the label (e.g. "Caution:")
  "after_step": number | null         // 0-based index of triggering step,
                                      // or null for section-top placement
}

MixTable = {
  "heading": string,
  "rows": TableRow[]                  // columns: Component | Amount
}

## Critical rules

- Output ONLY valid JSON. No markdown, no prose, no code fences.
- "section_type" MUST be "wet_lab" or "computational". Never both.
- "materials" and "mix_tables" MUST be [] for computational protocols.
- "prerequisites" MUST be null for wet_lab protocols.
- Nested steps use "children": []. There is no depth limit.
- Stopping points: use step_type "stopping_point" ONLY for points where the \
  operator can walk away from the bench and safely resume the protocol later \
  (minutes to days). The discriminating signal is whether the protocol can be \
  paused indefinitely at that point, NOT whether the step mentions a time, \
  temperature, or incubation. Examples:
    * "Store overnight at 4 °C." → stopping_point
    * "Samples can be stored at -80 °C for up to 6 months." → stopping_point
    * "Incubate at room temperature for 5 min." → action (NOT stopping_point)
    * "Centrifuge at 12,000 × g for 15 min at 4 °C." → action
    * "Heat to 95 °C for 3 min." → action
    * "Pause point — samples may be left on ice for up to 1 hour." → stopping_point
  When in doubt, default to step_type "action". Stopping points are rare; \
  most protocols contain zero or one.
- Footnotes from the source document belong in "notes", not inline in step text.
- Cross-references to other protocols use the format "Section X.Y".
- "author" defaults to "ARMI" if the source does not name a specific author.
- Apply **bold** and _italic_ markers to key parameters in step text \
(temperatures, volumes, concentrations, durations).
- For computational protocols: when a step includes a terminal command, script, \
or code snippet, put the prose instruction in "text" and the executable content \
in "code". Do not embed code in the "text" field with backtick fencing. \
Set "code" and "code_language" to null for wet_lab protocols.
- Do not include callout labels ("Caution:", "Note:") in the callout text field.
- Callout deduplication: if the same safety or quality concern applies to multiple \
  procedure sections, place the callout ONCE in the first section where it is \
  relevant. Do not repeat identical or near-identical callouts across sections. \
  If a concern is a general reagent hazard (e.g. toxicity), place it in the \
  first procedure section or as a preamble callout, not in every section.
- Callout self-containment: callout text must be fully intelligible in isolation, \
  without requiring the reader to refer to surrounding steps. Replace all \
  context-dependent pronouns and references ("this phase", "the sample", \
  "here") with specific named entities ("the upper aqueous phase", "the RNA \
  pellet", "the TRIzol homogenate"). A reader seeing only the callout box \
  must understand exactly what it refers to.
- Callout positioning: each callout has an "after_step" field. Set after_step \
  to the 0-based index of the step that triggers the callout when the warning \
  applies to a specific action (e.g. a CRITICAL warning about avoiding phase \
  contamination during a transfer step → after_step is the index of the \
  transfer step within section.steps). Set after_step to null for general \
  section-level warnings that apply throughout the section (e.g. "Work in \
  RNase-free conditions throughout"). When in doubt, prefer inline placement \
  (specific step index) over null — inline callouts render adjacent to the \
  triggering action and preserve context.
- For unknown supplier or catalogue number, use "—".
"""


# ---------------------------------------------------------------------------
# Style guide context builder
# ---------------------------------------------------------------------------

def _load_style_guide() -> dict:
    """Thin wrapper preserving the local function name for call sites."""
    try:
        return get_config()
    except FileNotFoundError:
        logger.warning("style_guide.yaml not found. Using empty config.")
        return {}


def build_style_guide_context(cfg: Optional[dict] = None) -> str:
    """
    Build a concise, token-efficient style guide context block for injection
    into the user message.

    Extracts the most extraction-relevant conventions from style_guide.yaml:
    callout classification rules, stopping point keywords, and section registry.
    Verbose descriptions are omitted to save tokens.

    Parameters
    ----------
    cfg : dict | None
        Pre-loaded style guide config. If None, loaded from disk.

    Returns
    -------
    str
        Multi-line string ready for injection into the prompt.
    """
    if cfg is None:
        cfg = _load_style_guide()

    lines = ["## Editorial conventions\n"]

    # Callout classification
    callouts = cfg.get("callouts", {})
    if callouts:
        lines.append("### Callout classification (priority: critical > caution > tip > note)")
        for ct in ("critical", "caution", "tip", "note"):
            meta = callouts.get(ct, {})
            kws = meta.get("keywords", [])
            label = meta.get("label", ct)
            kw_str = ", ".join(f'"{k}"' for k in kws[:6])  # cap at 6 to save tokens
            lines.append(f'- **{label}**: triggered by {kw_str}')
        lines.append("")

    # Stopping points
    sp = cfg.get("stopping_points", {})
    if sp:
        sp_kws = sp.get("keywords", [])
        kw_str = ", ".join(f'"{k}"' for k in sp_kws[:6])
        lines.append(
            '### Stopping points\n'
            'A stopping point is a step where the operator can walk away from\n'
            'the bench and safely resume the protocol later (minutes to days of\n'
            'unattended pause). Storage steps, overnight incubations, and explicit\n'
            '"pause point" markers qualify. Short incubations, centrifugations,\n'
            'and timed reactions DO NOT qualify, even if they mention a duration\n'
            f'or temperature. Common stopping-point phrases: {kw_str}.\n'
            'Phrases alone are not sufficient — the operator must be able to\n'
            'walk away for an extended period without needing to intervene further.\n'
            'When in doubt, classify as step_type "action".\n'
        )

    # Section registry
    sections = cfg.get("sections", {})
    if sections:
        lines.append("### Compendium section registry")
        lines.append(
            "Use these section_name values when the section number is known:\n"
            + "\n".join(f'  "{k}": "{v}"' for k, v in sections.items())
        )
        lines.append("")

    # Author rule
    lines.append(
        '### Author attribution\n'
        'Use "Alon Douek" if the source names this author. '
        'Use "ARMI" for all other protocols unless a specific author is clearly stated.\n'
    )

    # Inline formatting
    lines.append(
        "### Inline formatting\n"
        "Apply **bold** to: temperatures, volumes, concentrations, durations, "
        "centrifuge speeds, critical reagent names.\n"
        "Centrifuge speeds must be formatted as **number × g** (e.g. 12,000 × g),"
        "never use the form '*g*', 'xg' or '_g_'.\n"
        "Apply _italic_ to: gene names, species names, Latin terms.\n"
    )

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# User message assembler
# ---------------------------------------------------------------------------

def build_user_message(
    document_text: str,
    source_filename: Optional[str] = None,
    section_number_hint: Optional[str] = None,
    section_type_hint: Optional[str] = None,
    cfg: Optional[dict] = None,
) -> str:
    """
    Assemble the complete user message for the extraction request.

    Parameters
    ----------
    document_text : str
        The full_text from ParsedDocument — the document content to extract from.
    source_filename : str | None
        Original filename, injected as a hint for title inference.
    section_number_hint : str | None
        If the section number is already known (e.g. from a batch manifest or
        filename convention), inject it to anchor the LLM's output.
    section_type_hint : str | None
        "wet_lab" or "computational" if already inferred by the parser.
        Reduces ambiguity for borderline documents.
    cfg : dict | None
        Pre-loaded style guide config. If None, loaded from disk.

    Returns
    -------
    str
        The complete user message string.
    """
    if cfg is None:
        cfg = _load_style_guide()

    style_context = build_style_guide_context(cfg)

    hint_lines = []
    if source_filename:
        hint_lines.append(f'- Source filename: "{source_filename}"')
    if section_number_hint:
        sections = cfg.get("sections", {})
        section_name = sections.get(section_number_hint, "")
        hint_lines.append(f'- Section number: {section_number_hint}')
        if section_name:
            hint_lines.append(f'- Section name: "{section_name}"')
    if section_type_hint:
        hint_lines.append(f'- Protocol type (inferred): {section_type_hint}')

    hint_block = ""
    if hint_lines:
        hint_block = (
            "## Contextual hints\n"
            + "\n".join(hint_lines)
            + "\n\n"
        )

    user_message = (
        f"{style_context}\n\n"
        f"{hint_block}"
        f"## Source document\n\n"
        f"{document_text.strip()}\n\n"
        f"## Task\n\n"
        f"Extract the protocol above into a single JSON object matching the schema "
        f"in your system prompt. Apply all editorial conventions listed above. "
        f"Output ONLY the JSON object. Do not wrap it in code fences or add any prose."
    )

    return user_message


# ---------------------------------------------------------------------------
# Retry prompt builder
# ---------------------------------------------------------------------------

def build_retry_message(
    previous_output: str,
    validation_errors: str,
    attempt: int,
) -> str:
    """
    Build a correction prompt for use on validation failure.

    Passes the model's previous (invalid) output back alongside the specific
    Pydantic validation errors, asking for a targeted fix.

    Parameters
    ----------
    previous_output : str
        The raw string the model returned on the previous attempt.
    validation_errors : str
        Human-readable validation error string from Pydantic.
    attempt : int
        The current attempt number (1-based), for logging context.

    Returns
    -------
    str
        A new user message that continues the conversation with the model.
    """
    return (
        f"Your previous response (attempt {attempt}) contained validation errors:\n\n"
        f"```\n{validation_errors}\n```\n\n"
        f"Your previous output was:\n\n"
        f"```json\n{previous_output[:2000]}"
        f"{'... [truncated]' if len(previous_output) > 2000 else ''}\n```\n\n"
        f"Fix ONLY the fields that caused validation errors. "
        f"Output the corrected complete JSON object and nothing else."
    )


# ---------------------------------------------------------------------------
# Token budget estimation
# ---------------------------------------------------------------------------

# Rough character-per-token estimate for Qwen2.5 on English/scientific text.
# Actual tokenisation varies; this is conservative (overestimates token count).
_CHARS_PER_TOKEN = 3.5

# Qwen2.5:14b context window
_CONTEXT_WINDOW = 32768

# Warn if estimated prompt tokens exceed this fraction of the context window
_WARN_THRESHOLD = 0.75


def estimate_token_count(text: str) -> int:
    """
    Estimate the token count of a text string using a character-ratio heuristic.

    Parameters
    ----------
    text : str

    Returns
    -------
    int
        Estimated token count.
    """
    return max(1, int(len(text) / _CHARS_PER_TOKEN))


def check_token_budget(
    system_prompt: str,
    user_message: str,
    max_tokens: int,
) -> tuple[int, bool]:
    """
    Estimate total token usage and warn if it approaches the context limit.

    Parameters
    ----------
    system_prompt : str
    user_message : str
    max_tokens : int
        The completion max_tokens value from config.

    Returns
    -------
    tuple[int, bool]
        (estimated_total_tokens, over_budget)
        over_budget is True if the estimate exceeds the warn threshold.
    """
    prompt_tokens = estimate_token_count(system_prompt + user_message)
    total = prompt_tokens + max_tokens
    over_budget = total > (_CONTEXT_WINDOW * _WARN_THRESHOLD)

    if over_budget:
        logger.warning(
            "Token budget warning: estimated ~{} tokens total "
            "(prompt ~{} + max_completion {}). "
            "Context window is {}. Consider DEV-001 two-pass strategy if extraction fails.",
            total,
            prompt_tokens,
            max_tokens,
            _CONTEXT_WINDOW,
        )
    else:
        logger.debug(
            "Token budget: estimated ~{} tokens total (prompt ~{} + max_completion {}).",
            total,
            prompt_tokens,
            max_tokens,
        )

    return total, over_budget
