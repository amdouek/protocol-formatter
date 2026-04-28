"""
extractor/llm_extractor.py -- Ollama LLM extraction with Pydantic validation.

Sends the parsed document text to the locally running Ollama server
(Qwen2.5:14b), receives a JSON completion, validates it against the Protocol
Pydantic schema, and returns a validated Protocol instance.

Retry strategy
--------------
On Pydantic validation failure the extractor makes up to ``max_retries``
additional attempts, each time injecting the previous (invalid) output and the
specific validation errors into a correction prompt. This gives the model
targeted feedback rather than asking it to start from scratch.

If all retries are exhausted the extractor falls back to rule_extractor.py
(heuristic extraction), which always produces a schema-valid Protocol even if
some fields are incomplete.

DEVNOTE001: Single-pass extraction (see prompts.py).

Conversation management
-----------------------
The Ollama /api/chat endpoint is used (messages array) rather than /api/generate
so that retry correction prompts can be passed as assistant/user turns within the
same logical conversation, giving the model the context of its own prior mistakes.
"""

from __future__ import annotations

import json
import re
import time
from typing import Optional

import urllib.request
import urllib.error
from loguru import logger

from pydantic import ValidationError

from config import get_config


def _load_cfg() -> dict:
    """Thin wrapper for backward-compatible internal calls."""
    return get_config()


# ---------------------------------------------------------------------------
# Custom exceptions
# ---------------------------------------------------------------------------

class OllamaUnavailableError(RuntimeError):
    """Raised when the Ollama server cannot be reached."""


class ExtractionError(RuntimeError):
    """Raised when all extraction attempts (LLM + heuristic) fail."""


# ---------------------------------------------------------------------------
# Ollama HTTP client
# ---------------------------------------------------------------------------

def _post_json(url: str, payload: dict, timeout: int) -> dict:
    """
    POST a JSON payload to a URL and return the parsed JSON response.

    Uses only the stdlib urllib to avoid an httpx/requests dependency.

    Parameters
    ----------
    url : str
    payload : dict
    timeout : int

    Returns
    -------
    dict

    Raises
    ------
    OllamaUnavailableError
        If the connection is refused or the server is unreachable.
    RuntimeError
        If the server returns a non-200 status or invalid JSON.
    """
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.URLError as exc:
        reason = str(exc.reason) if hasattr(exc, "reason") else str(exc)
        raise OllamaUnavailableError(
            f"Cannot reach Ollama server at {url}.\n"
            f"Ensure Ollama is running: `ollama serve`\n"
            f"Reason: {reason}"
        ) from exc

    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Ollama returned non-JSON response: {raw[:300]}"
        ) from exc


def check_ollama_available(base_url: str, timeout: int = 5) -> tuple[bool, str]:
    """
    Check whether the Ollama server is reachable.

    Parameters
    ----------
    base_url : str
    timeout : int

    Returns
    -------
    tuple[bool, str]
        (True, model_list_summary) or (False, error_message)
    """
    try:
        req = urllib.request.Request(f"{base_url}/api/tags")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        models = [m.get("name", "?") for m in data.get("models", [])]
        return True, f"Available models: {', '.join(models) or '(none pulled)'}"
    except Exception as exc:
        return False, str(exc)


# ---------------------------------------------------------------------------
# JSON extraction from model output
# ---------------------------------------------------------------------------

def _extract_json_from_response(text: str) -> str:
    """
    Extract the JSON object from the model's response text.

    Handles the common cases where the model wraps its output in:
        ```json ... ```
        ``` ... ```
        or just outputs the JSON directly.

    Returns the raw JSON string, or raises ValueError if no JSON object
    can be located.

    Parameters
    ----------
    text : str

    Returns
    -------
    str

    Raises
    ------
    ValueError
        If no JSON object is found in the response.
    """
    # Strip code fences if present
    fence_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if fence_match:
        return fence_match.group(1)

    # Find the outermost JSON object by balanced brace scanning
    start = text.find("{")
    if start == -1:
        raise ValueError("No JSON object found in model response.")

    depth = 0
    in_string = False
    escape_next = False

    for i, ch in enumerate(text[start:], start):
        if escape_next:
            escape_next = False
            continue
        if ch == "\\" and in_string:
            escape_next = True
            continue
        if ch == '"' and not escape_next:
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start : i + 1]

    raise ValueError(
        "Incomplete JSON object in model response "
        f"(opened at position {start}, never closed)."
    )


# ---------------------------------------------------------------------------
# Schema field post-processing
# ---------------------------------------------------------------------------

def _post_process_payload(payload: dict) -> dict:
    """
    Apply light post-processing to the raw parsed JSON before Pydantic
    validation.

    Rules applied:
    - Ensure "author" defaults to "ARMI" if missing or empty.
    - Ensure "version" defaults to "1.0" if missing.
    - Ensure "materials", "mix_tables", "notes", "references" default to [].
    - Ensure "procedure" is a list.
    - Strip callout labels from callout.text fields.
    - Normalise "section_type" to lowercase.

    Parameters
    ----------
    payload : dict

    Returns
    -------
    dict
    """
    from parser.utils import strip_callout_prefix, normalise_whitespace

    # Scalar defaults
    if not payload.get("author"):
        payload["author"] = "ARMI"
    if not payload.get("version"):
        payload["version"] = "1.0"

    # Normalise section_type
    st = payload.get("section_type", "")
    payload["section_type"] = st.lower().strip() if st else "wet_lab"
    
    # Normalise date to DD/MM/YYYY if the LLM returned a different format
    raw_date = payload.get("date")
    if raw_date and not re.match(r"^\d{2}/\d{2}/\d{4}$", str(raw_date)):
        from parser.utils import normalise_date
        normalised = normalise_date(str(raw_date))
        payload["date"] = normalised  # None if unrecognisable — validator accepts None

    # List defaults
    for key in ("materials", "mix_tables", "notes", "references"):
        if not isinstance(payload.get(key), list):
            payload[key] = []

    if not isinstance(payload.get("procedure"), list):
        payload["procedure"] = []

    # Strip callout prefixes recursively through procedure sections
    for section in payload.get("procedure", []):
        for callout in section.get("callouts", []):
            if "text" in callout:
                callout["text"] = strip_callout_prefix(
                    normalise_whitespace(callout["text"])
                )

    # Ensure mix_table rows that start with Total/Incubate have bold=True
    for mt in payload.get("mix_tables", []):
        for row in mt.get("rows", []):
            cells = row.get("cells", [])
            if cells and re.match(r"^(total|incubate)", cells[0], re.IGNORECASE):
                row["bold"] = True

    # Normalise table row widths: pad short rows with "—", trim long rows.
    # Mitigates burning LLM retries on minor cell-count mismatches which really don't matter.
    for mt in payload.get("materials", []):
        for row in mt.get("rows", []):
            cells = row.get("cells", [])
            if len(cells) < 3:
                cells.extend(["\u2014"] * (3 - len(cells)))
            elif len(cells) > 3:
                cells[:] = cells[:3]
            row["cells"] = cells

    for mt in payload.get("mix_tables", []):
        for row in mt.get("rows", []):
            cells = row.get("cells", [])
            if len(cells) < 2:
                cells.extend(["\u2014"] * (2 - len(cells)))
            elif len(cells) > 2:
                cells[:] = cells[:2]
            row["cells"] = cells

    return payload


# ---------------------------------------------------------------------------
# Main extraction function
# ---------------------------------------------------------------------------

def extract_protocol(
    parsed_document,
    source_filename: Optional[str] = None,
    section_number_hint: Optional[str] = None,
    cfg: Optional[dict] = None,
):
    """
    Extract a validated Protocol from a ParsedDocument using Ollama.

    Parameters
    ----------
    parsed_document : ParsedDocument
        Output of parser.read_document(). Provides full_text, inferred
        section_type, author, date, and source_path.
    source_filename : str | None
        Original filename for hint injection and source_filename field.
    section_number_hint : str | None
        If known (e.g. from batch manifest), passed as a prompt hint.
    cfg : dict | None
        Pre-loaded style_guide config. If None, loaded from disk.

    Returns
    -------
    Protocol
        Validated Protocol instance.

    Raises
    ------
    OllamaUnavailableError
        If the Ollama server cannot be reached.
    ExtractionError
        If all LLM attempts and the heuristic fallback all fail.
    """
    # Import here to avoid circular imports at module level
    from schema import Protocol
    from extractor.prompts import (
        SYSTEM_PROMPT,
        build_user_message,
        build_retry_message,
        check_token_budget,
    )
    from extractor.rule_extractor import extract_protocol_heuristic

    if cfg is None:
        cfg = _load_cfg()

    ollama_cfg = cfg.get("ollama", {})
    base_url: str = ollama_cfg.get("base_url", "http://localhost:11434")
    model: str = ollama_cfg.get("model", "qwen2.5:14b")
    max_tokens: int = int(ollama_cfg.get("max_tokens", 4096))
    max_retries: int = int(ollama_cfg.get("max_retries", 3))
    timeout: int = int(ollama_cfg.get("request_timeout_seconds", 120))

    # Derive hints from the parsed document
    fname = source_filename or parsed_document.source_path.name
    section_type_hint = parsed_document.section_type

    # Build the initial user message
    user_message = build_user_message(
        document_text=parsed_document.full_text,
        source_filename=fname,
        section_number_hint=section_number_hint,
        section_type_hint=section_type_hint,
        cfg=cfg,
    )

    # Check token budget (DEVNOTE001 warning)
    check_token_budget(SYSTEM_PROMPT, user_message, max_tokens)

    # Conversation history for retry continuation
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_message},
    ]

    last_raw: str = ""
    last_errors: str = ""

    for attempt in range(1, max_retries + 2):  # +1 for the first attempt
        is_retry = attempt > 1

        if is_retry:
            logger.info(
                "LLM extraction attempt {}/{} (retry after validation failure).",
                attempt,
                max_retries + 1,
            )
            retry_msg = build_retry_message(last_raw, last_errors, attempt - 1)
            messages.append({"role": "assistant", "content": last_raw})
            messages.append({"role": "user", "content": retry_msg})
        else:
            logger.info(
                "LLM extraction attempt 1/{} — model={}, tokens_max={}",
                max_retries + 1,
                model,
                max_tokens,
            )

        t_start = time.monotonic()

        payload = {
            "model": model,
            "messages": messages,
            "stream": False,
            "format": "json",
            "options": {
                "num_predict": max_tokens,
                "temperature": 0.1,   # low temperature for structured output
            },
        }

        try:
            response = _post_json(
                f"{base_url}/api/chat",
                payload,
                timeout,
            )
        except OllamaUnavailableError:
            raise
        except RuntimeError as exc:
            logger.error("Ollama request error on attempt {}: {}", attempt, exc)
            last_errors = str(exc)
            if attempt > max_retries:
                break
            continue

        elapsed = time.monotonic() - t_start
        logger.debug("Ollama responded in {:.1f}s", elapsed)

        # Extract content from response
        try:
            raw_content: str = response["message"]["content"]
        except (KeyError, TypeError) as exc:
            last_errors = f"Unexpected Ollama response structure: {exc}"
            last_raw = json.dumps(response)[:500]
            logger.warning("Unexpected response structure: {}", last_errors)
            if attempt > max_retries:
                break
            continue

        last_raw = raw_content

        # Extract JSON from the response text
        try:
            json_str = _extract_json_from_response(raw_content)
        except ValueError as exc:
            last_errors = f"No JSON object found in model response: {exc}"
            logger.warning("JSON extraction failed on attempt {}: {}", attempt, exc)
            if attempt > max_retries:
                break
            continue

        # Parse JSON
        try:
            raw_payload = json.loads(json_str)
        except json.JSONDecodeError as exc:
            last_errors = f"JSON parse error: {exc}"
            logger.warning("JSON parse failed on attempt {}: {}", attempt, exc)
            if attempt > max_retries:
                break
            continue

        # Post-process before Pydantic validation
        raw_payload = _post_process_payload(raw_payload)

        # Inject source_filename from the actual file, not the model's guess
        raw_payload["source_filename"] = fname

        # Pydantic validation
        try:
            protocol = Protocol.model_validate(raw_payload)
            logger.success(
                "Protocol extracted and validated on attempt {} in {:.1f}s: {!r}",
                attempt,
                elapsed,
                protocol.title,
            )
            return protocol

        except ValidationError as exc:
            last_errors = str(exc)
            logger.warning(
                "Pydantic validation failed on attempt {} ({} errors): {}",
                attempt,
                exc.error_count(),
                last_errors[:400],
            )
            if attempt > max_retries:
                break
            continue

    # All LLM attempts exhausted — fall back to heuristic extractor
    logger.warning(
        "All {} LLM extraction attempts failed. Falling back to heuristic extractor.",
        max_retries + 1,
    )
    try:
        protocol = extract_protocol_heuristic(parsed_document, cfg=cfg)
        logger.info(
            "Heuristic extraction succeeded: {!r}",
            protocol.title,
        )
        return protocol
    except Exception as fallback_exc:
        raise ExtractionError(
            f"All extraction methods failed.\n"
            f"Last LLM error: {last_errors}\n"
            f"Heuristic fallback error: {fallback_exc}"
        ) from fallback_exc
