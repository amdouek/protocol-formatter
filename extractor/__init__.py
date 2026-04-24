"""
extractor -- LLM and heuristic extraction package for ProtocolFormatter.

Public API
----------
    extract_protocol(parsed_document, ...)    → Protocol
        Primary entry point. Calls Ollama with retry, falls back to heuristic.

    extract_protocol_heuristic(parsed_document, ...)  → Protocol
        Heuristic-only extraction. Always succeeds; some fields may be sparse.

    check_ollama_available(base_url)          → (bool, str)
        Preflight check for the Ollama server.

    OllamaUnavailableError
    ExtractionError
"""
from extractor.llm_extractor import (
    extract_protocol,
    check_ollama_available,
    OllamaUnavailableError,
    ExtractionError,
)
from extractor.rule_extractor import extract_protocol_heuristic

__all__ = [
    "extract_protocol",
    "extract_protocol_heuristic",
    "check_ollama_available",
    "OllamaUnavailableError",
    "ExtractionError",
]
