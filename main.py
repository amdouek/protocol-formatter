"""
main.py — CLI entry point for ProtocolFormatter.

Usage
-----
    # Format a single protocol
    protocol-formatter format path/to/protocol.docx

    # Format with known section number
    protocol-formatter format path/to/protocol.docx --section 4

    # Format with human review step before rendering
    protocol-formatter format path/to/protocol.docx --review

    # Format a batch of protocols
    protocol-formatter batch path/to/protocols/ --section 4

    # Use heuristic extraction only (no Ollama required)
    protocol-formatter format path/to/protocol.docx --heuristic

    # Preflight: check all dependencies are available
    protocol-formatter check

Pipeline (per document)
-----------------------
    1. Duplicate detection (filename-based, DEV-003)
    2. Input parsing (.docx → ParsedDocument)
    3. LLM extraction (Ollama → Protocol) with heuristic fallback
    4. --review step (optional human-in-the-loop)
    5. Node.js rendering (Protocol → .docx)
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Annotated, Optional

import typer
import yaml
from loguru import logger
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Confirm, Prompt
from rich.syntax import Syntax
from rich.table import Table
from rich import print as rprint

from extractor import rule_extractor

app = typer.Typer(
    name="protocol-formatter",
    help="Convert raw laboratory protocol documents into consistently formatted Word files.",
    add_completion=False,
    no_args_is_help=True,
)

console = Console()

_PACKAGE_ROOT = Path(__file__).resolve().parent
_STYLE_GUIDE_PATH = _PACKAGE_ROOT / "configs" / "style_guide.yaml"


# ---------------------------------------------------------------------------
# Configuration helpers
# ---------------------------------------------------------------------------

def _load_cfg() -> dict:
    if not _STYLE_GUIDE_PATH.exists():
        console.print(
            f"[red]Error:[/red] style_guide.yaml not found at {_STYLE_GUIDE_PATH}.\n"
            "Ensure you are running from the protocol_formatter package root."
        )
        raise typer.Exit(code=1)
    with _STYLE_GUIDE_PATH.open("r", encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def _configure_logging(cfg: dict) -> None:
    """Configure loguru based on style_guide.yaml logging settings."""
    level = cfg.get("logging", {}).get("level", "INFO")
    logger.remove()
    logger.add(
        sys.stderr,
        level=level,
        format="<green>{time:HH:mm:ss}</green> | <level>{level: <8}</level> | {message}",
        colorize=True,
    )

def _validate_output_dir(explicit: Optional[Path], cfg: dict) -> Path:
    """
    Resolve and validate the output directory path without creating it.

    Catches obviously invalid paths (e.g. pointing at an existing file)
    before the pipeline runs. Actual directory creation is deferred to
    the renderer, which only creates it when output is ready to write.
    """
    output_dir = explicit or Path(cfg.get("paths", {}).get("output_dir", "output"))
    output_dir = Path(output_dir).resolve()

    if output_dir.exists() and not output_dir.is_dir():
        console.print(
            f"[red]Error:[/red] Output path exists but is not a directory: {output_dir}"
        )
        raise typer.Exit(code=1)

    return output_dir

# ---------------------------------------------------------------------------
# Pipeline result type
# ---------------------------------------------------------------------------

from enum import Enum
from dataclasses import dataclass


class ResultStatus(Enum):
    SUCCESS = "success"
    FAILED = "failed"
    SKIPPED = "skipped"


@dataclass
class PipelineResult:
    """Outcome of a single _process_one() invocation."""
    status: ResultStatus
    output_path: Optional[Path] = None
    reason: str = ""

# ---------------------------------------------------------------------------
# Duplicate detection (DEV-003)
# ---------------------------------------------------------------------------

class DuplicateDetector:
    """
    Filename-based duplicate guard for batch runs.

    Tracks filenames seen within a single CLI invocation. Case-insensitive,
    path-stripped comparison.
    """

    def __init__(self, case_sensitive: bool = False):
        self._seen: set[str] = set()
        self._case_sensitive = case_sensitive

    def _normalise(self, path: Path) -> str:
        name = path.name
        return name if self._case_sensitive else name.lower()

    def check(self, path: Path) -> bool:
        """
        Return True if this file has already been seen (duplicate).
        Register the file if it has not.
        """
        key = self._normalise(path)
        if key in self._seen:
            return True
        self._seen.add(key)
        return False

    def reset(self) -> None:
        self._seen.clear()


# ---------------------------------------------------------------------------
# Review step (human-in-the-loop)
# ---------------------------------------------------------------------------

def _run_review_step(protocol) -> object:
    """
    Present the extracted Protocol schema to the user for review and optional
    editing before rendering.

    Displays the protocol as pretty-printed JSON with syntax highlighting.
    Allows the user to:
      - Accept as-is → returns the original Protocol
      - Edit a field → prompts for field path and new value, re-validates,
        and loops until accepted or aborted
      - Abort → raises typer.Exit(1)

    Parameters
    ----------
    protocol : Protocol
        The validated Protocol instance to review.

    Returns
    -------
    Protocol
        The (possibly edited) Protocol instance.
    """
    from schema import Protocol

    console.print()
    console.print(Panel(
        "[bold cornflower_blue]Protocol Review[/bold cornflower_blue]\n"
        "Review the extracted schema before rendering.\n"
        "Press [bold]A[/bold] to accept, [bold]E[/bold] to edit a field, "
        "or [bold]Q[/bold] to abort.",
        expand=False,
    ))

    while True:
        # Pretty-print the schema as JSON with syntax highlighting
        json_str = protocol.model_dump_json(indent=2)
        syntax = Syntax(json_str, "json", theme="monokai", line_numbers=False,
                        word_wrap=True)
        console.print(syntax)
        console.print()

        # Summary table
        table = Table(show_header=False, box=None, padding=(0, 2))
        table.add_column("Field", style="bold")
        table.add_column("Value")
        table.add_row("Title", protocol.title)
        table.add_row("Type", str(protocol.section_type.value))
        table.add_row("Section", f"{protocol.section_number} — {protocol.section_name}")
        table.add_row("Author", protocol.author)
        table.add_row("Version", protocol.version)
        table.add_row("Procedure sections", str(len(protocol.procedure)))
        table.add_row("Notes", str(len(protocol.notes)))
        console.print(table)
        console.print()

        choice = Prompt.ask(
            "[bold]Accept / Edit / Quit[/bold]",
            choices=["a", "e", "q", "A", "E", "Q"],
            default="a",
        ).lower()

        if choice == "a":
            console.print("[green]✓[/green] Schema accepted.")
            return protocol

        elif choice == "q":
            console.print("[yellow]Aborted by user.[/yellow]")
            raise typer.Exit(code=1)

        elif choice == "e":
            field_path = Prompt.ask(
                "Field path to edit (e.g. [bold]title[/bold], "
                "[bold]procedure[0].heading[/bold], [bold]author[/bold])"
            )
            new_value_str = Prompt.ask("New value (enter as JSON — strings need quotes)")

            try:
                new_value = json.loads(new_value_str)
            except json.JSONDecodeError:
                # Treat unquoted input as a plain string
                new_value = new_value_str

            try:
                # Apply the edit to a dict representation, then re-validate
                data = json.loads(protocol.model_dump_json())
                _set_nested(data, field_path, new_value)
                protocol = Protocol.model_validate(data)
                console.print(f"[green]✓[/green] Field [bold]{field_path}[/bold] updated.")
            except Exception as exc:
                console.print(
                    f"[red]Edit failed:[/red] {exc}\n"
                    "The original value has been retained."
                )


def _set_nested(data: dict, path: str, value) -> None:
    """
    Set a value in a nested dict/list using a dotted path with optional
    integer indices (e.g. "procedure[0].heading").

    Parameters
    ----------
    data : dict
        The data structure to modify in-place.
    path : str
        Dotted path, e.g. "procedure[0].heading" or "title".
    value :
        The new value to set.

    Raises
    ------
    KeyError / IndexError / ValueError
        If the path is invalid or the index is out of range.
    """
    import re
    # Split on dots and bracket indices: "procedure[0].heading" → ["procedure", 0, "heading"]
    parts = []
    for token in re.split(r"\.(?![^\[]*\])", path):
        m = re.match(r"^(\w+)\[(\d+)\]$", token)
        if m:
            parts.append(m.group(1))
            parts.append(int(m.group(2)))
        else:
            parts.append(token)

    obj = data
    for part in parts[:-1]:
        if isinstance(part, int):
            obj = obj[part]
        else:
            obj = obj[part]

    last = parts[-1]
    if isinstance(last, int):
        obj[last] = value
    else:
        obj[last] = value


# ---------------------------------------------------------------------------
# Core pipeline
# ---------------------------------------------------------------------------

def _process_one(
    source_path: Path,
    output_dir: Path,
    section_hint: Optional[str],
    heuristic_only: bool,
    do_review: bool,
    cfg: dict,
    detector: DuplicateDetector,
) -> PipelineResult:
    """
    Run the full pipeline for a single source document.

    Returns a PipelineResult indicating the outcome.
    """
    from parser import read_document
    from extractor.llm_extractor import (
        extract_protocol,
        check_ollama_available,
        OllamaUnavailableError,
        ExtractionError,
    )
    from extractor.rule_extractor import extract_protocol_heuristic
    from renderer.node_renderer import render_protocol, RendererError

    # ── Duplicate detection ──────────────────────────────────────────────────
    if detector.check(source_path):
        console.print(
            f"[yellow]⚠  Duplicate:[/yellow] '{source_path.name}' has already been "
            "processed in this batch. Skipping."
        )
        return PipelineResult(ResultStatus.SKIPPED, reason="duplicate")

    console.print(f"\n[bold]Processing:[/bold] {source_path.name}")

    # ── Parse ────────────────────────────────────────────────────────────────
    try:
        with console.status("Parsing source document…"):
            parsed = read_document(source_path)
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        console.print(f"[red]Parse error:[/red] {exc}")
        return PipelineResult(ResultStatus.FAILED, reason=str(exc))

    console.print(
        f"  Parsed: {len(parsed.paragraphs)} paragraphs, "
        f"{len(parsed.tables)} tables, "
        f"type=[bold]{parsed.section_type}[/bold]"
    )

    # ── Extract ──────────────────────────────────────────────────────────────
    if heuristic_only:
        console.print("  [dim]--heuristic: skipping LLM, using rule extractor[/dim]")
        try:
            with console.status("Running heuristic extraction…"):
                protocol = extract_protocol_heuristic(
                    parsed,
                    source_filename=source_path.name,
                    section_number_hint=section_hint,
                    cfg=cfg,
                )
        except Exception as exc:
            console.print(f"[red]Heuristic extraction error:[/red] {exc}")
            return PipelineResult(ResultStatus.FAILED, reason=str(exc))
    else:
        # Preflight: check Ollama is running before attempting extraction
        base_url = cfg.get("ollama", {}).get("base_url", "http://localhost:11434")
        ok, msg = check_ollama_available(base_url, timeout=5)
        if not ok:
            console.print(
                f"[red]Ollama unavailable:[/red] {msg}\n"
                "Start Ollama with [bold]ollama serve[/bold], or use "
                "[bold]--heuristic[/bold] to skip LLM extraction."
            )
            return PipelineResult(ResultStatus.FAILED, reason=msg)

        try:
            with console.status(
                f"Extracting with Ollama ({cfg.get('ollama', {}).get('model', 'qwen2.5:14b')})…"
            ):
                protocol = extract_protocol(
                    parsed,
                    source_filename=source_path.name,
                    section_number_hint=section_hint,
                    cfg=cfg,
                )
        except OllamaUnavailableError as exc:
            console.print(f"[red]Ollama error:[/red] {exc}")
            return PipelineResult(ResultStatus.FAILED, reason=str(exc))
        except ExtractionError as exc:
            console.print(f"[red]Extraction failed:[/red] {exc}")
            return PipelineResult(ResultStatus.FAILED, reason=str(exc))

    console.print(
        f"  Extracted: [bold]{protocol.title!r}[/bold] "
        f"(author={protocol.author}, version={protocol.version})"
    )

    # ── Review ───────────────────────────────────────────────────────────────
    if do_review:
        try:
            protocol = _run_review_step(protocol)
        except typer.Exit:
            return PipelineResult(ResultStatus.FAILED, reason="aborted by user")

    # ── Render ───────────────────────────────────────────────────────────────
    try:
        with console.status("Rendering .docx…"):
            output_path = render_protocol(protocol, output_dir=output_dir)
    except RendererError as exc:
        console.print(f"[red]Render error:[/red] {exc}")
        return PipelineResult(ResultStatus.FAILED, reason=str(exc))
    except FileNotFoundError as exc:
        console.print(f"[red]Renderer not found:[/red] {exc}")
        return PipelineResult(ResultStatus.FAILED, reason=str(exc))

    console.print(f"  [green]✓[/green] Written to: [bold]{output_path}[/bold]")
    return PipelineResult(ResultStatus.SUCCESS, output_path=output_path)


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

@app.command()
def format(
    source: Annotated[
        Path,
        typer.Argument(help="Path to the source .docx or .doc protocol file."),
    ],
    section: Annotated[
        Optional[str],
        typer.Option(
            "--section", "-s",
            help="Compendium section number (e.g. '4'). Injected as a hint for the LLM.",
        ),
    ] = None,
    output: Annotated[
        Optional[Path],
        typer.Option(
            "--output", "-o",
            help="Output directory. Defaults to style_guide.yaml output_dir.",
        ),
    ] = None,
    review: Annotated[
        bool,
        typer.Option(
            "--review/--no-review",
            help="Pause for human review of the extracted schema before rendering.",
        ),
    ] = False,
    heuristic: Annotated[
        bool,
        typer.Option(
            "--heuristic",
            help="Use heuristic rule extractor only. Skips Ollama (no LLM required).",
        ),
    ] = False,
) -> None:
    """
    Format a single protocol source document into a standardised .docx file.
    """
    cfg = _load_cfg()
    _configure_logging(cfg)

    if not source.exists():
        console.print(f"[red]Error:[/red] File not found: {source}")
        raise typer.Exit(code=1)

    if source.suffix.lower() not in (".docx", ".doc"):
        console.print(
            f"[red]Error:[/red] Unsupported file type '{source.suffix}'. "
            "Expected .docx or .doc."
        )
        raise typer.Exit(code=1)

    output_dir = _validate_output_dir(output, cfg)

    detector = DuplicateDetector(
        case_sensitive=cfg.get("duplicate_detection", {}).get("case_sensitive", False)
    )

    result = _process_one(
        source_path=source.resolve(),
        output_dir=output_dir,
        section_hint=section,
        heuristic_only=heuristic,
        do_review=review,
        cfg=cfg,
        detector=detector,
    )

    if result.status != ResultStatus.SUCCESS:
        raise typer.Exit(code=1)


@app.command()
def batch(
    source_dir: Annotated[
        Path,
        typer.Argument(help="Directory containing .docx/.doc protocol source files."),
    ],
    section: Annotated[
        Optional[str],
        typer.Option(
            "--section", "-s",
            help="Section number applied to all files in this batch.",
        ),
    ] = None,
    output: Annotated[
        Optional[Path],
        typer.Option(
            "--output", "-o",
            help="Output directory. Defaults to style_guide.yaml output_dir.",
        ),
    ] = None,
    review: Annotated[
        bool,
        typer.Option(
            "--review/--no-review",
            help="Pause for review before rendering each document in the batch.",
        ),
    ] = False,
    heuristic: Annotated[
        bool,
        typer.Option(
            "--heuristic",
            help="Use heuristic rule extractor only. Skips Ollama.",
        ),
    ] = False,
    glob: Annotated[
        str,
        typer.Option(
            "--glob", "-g",
            help="Glob pattern for source files within the directory.",
        ),
    ] = "**/*.docx",
) -> None:
    """
    Format all protocol source documents in a directory.

    Files are processed in alphabetical order. Duplicate filenames within
    the batch are detected and skipped.
    """
    cfg = _load_cfg()
    _configure_logging(cfg)

    if not source_dir.exists() or not source_dir.is_dir():
        console.print(f"[red]Error:[/red] Directory not found: {source_dir}")
        raise typer.Exit(code=1)

    sources = sorted(source_dir.glob(glob))
    # Also pick up .doc files if the glob only covers .docx
    if glob.endswith("*.docx"):
        doc_glob = glob[:-len("*.docx")] + "*.doc"
        doc_sources = list(source_dir.glob(doc_glob))
        # Merge and deduplicate (by resolved path), maintaining sort order
        seen_paths = {p.resolve() for p in sources}
        for p in doc_sources:
            if p.resolve() not in seen_paths:
                sources.append(p)
        sources = sorted(sources)

    if not sources:
        console.print(
            f"[yellow]No files found[/yellow] matching '{glob}' in {source_dir}."
        )
        raise typer.Exit(code=0)

    console.print(
        Panel(
            f"[bold]Batch:[/bold] {len(sources)} file(s) in [italic]{source_dir}[/italic]\n"
            f"Section hint: {section or '(none)'} | "
            f"Review: {'yes' if review else 'no'} | "
            f"Heuristic: {'yes' if heuristic else 'no'}",
            expand=False,
        )
    )

    output_dir = _validate_output_dir(output, cfg)
    detector = DuplicateDetector(
        case_sensitive=cfg.get("duplicate_detection", {}).get("case_sensitive", False)
    )

    succeeded, failed, skipped = 0, 0, 0

    for i, source_path in enumerate(sources, 1):
        console.rule(f"[dim]{i}/{len(sources)}[/dim]")
        result = _process_one(
            source_path=source_path.resolve(),
            output_dir=output_dir,
            section_hint=section,
            heuristic_only=heuristic,
            do_review=review,
            cfg=cfg,
            detector=detector,
        )
        if result.status == ResultStatus.SUCCESS:
            succeeded += 1
        elif result.status == ResultStatus.SKIPPED:
            skipped += 1
        else:
            failed += 1

    console.rule()
    console.print(
        f"\n[bold]Batch complete:[/bold] "
        f"[green]{succeeded} succeeded[/green] | "
        f"[red]{failed} failed[/red] | "
        f"[yellow]{skipped} skipped (duplicates)[/yellow]"
    )

    if failed > 0:
        raise typer.Exit(code=1)


@app.command()
def check() -> None:
    """
    Run preflight checks on all dependencies.

    Verifies: Node.js, docx npm package, Ollama server, Ollama model,
    pandoc (for .doc support), and render.js existence.
    """
    cfg = _load_cfg()
    _configure_logging(cfg)

    console.print(Panel("[bold]ProtocolFormatter — Dependency Check[/bold]", expand=False))

    all_ok = True

    def _row(label: str, ok: bool, detail: str, optional: bool = False) -> None:
        nonlocal all_ok
        if ok:
            icon = "[green]✓[/green]"
        elif optional:
            icon = "[yellow]○[/yellow]"
        else:
            icon = "[red]✗[/red]"
        console.print(f"  {icon}  {label}: {detail}")
        if not ok and not optional:
            all_ok = False

    # Node.js
    from renderer.node_renderer import check_node_available, check_docx_package_available
    node_exe = cfg.get("paths", {}).get("node_executable", "node")
    ok, msg = check_node_available(node_exe)
    _row("Node.js", ok, msg)

    # docx npm package
    if ok:
        ok2, msg2 = check_docx_package_available(node_exe)
        _row("docx npm package", ok2, msg2 if ok2 else msg2)

    # render.js
    render_rel = cfg.get("paths", {}).get("render_script", "renderer/templates/render.js")
    render_path = _PACKAGE_ROOT / render_rel
    _row("render.js", render_path.exists(), str(render_path))

    # Ollama server
    from extractor.llm_extractor import check_ollama_available
    base_url = cfg.get("ollama", {}).get("base_url", "http://localhost:11434")
    ok3, msg3 = check_ollama_available(base_url, timeout=5)
    _row("Ollama server", ok3, msg3 if ok3 else f"{msg3} — run `ollama serve`")

    # Ollama model
    if ok3:
        model = cfg.get("ollama", {}).get("model", "qwen2.5:14b")
        model_present = model in msg3
        _row(
            f"Ollama model ({model})",
            model_present,
            "available" if model_present else f"not pulled — run `ollama pull {model}`",
        )

    # pandoc (optional — only needed for .doc files)
    from parser.doc_reader import check_pandoc_available
    pandoc_exe = cfg.get("paths", {}).get("pandoc_executable", "pandoc")
    ok4, msg4 = check_pandoc_available(pandoc_exe)
    _row(
        "pandoc (.doc support)",
        ok4,
        msg4 if ok4 else f"{msg4} [dim](optional — needed only for .doc files)[/dim]",
        optional=True,
    )

    console.print()
    if all_ok:
        console.print("[bold green]All checks passed.[/bold green]")
    else:
        console.print(
            "[bold red]Some checks failed.[/bold red] "
            "Fix the issues marked [red]✗[/red] above before running protocol-formatter."
        )
        raise typer.Exit(code=1)


@app.command()
def schema() -> None:
    """
    Print the Protocol JSON schema to stdout.

    Useful for inspecting the intermediate representation or for building
    test payloads.
    """
    from schema import Protocol
    import json as _json
    console.print_json(_json.dumps(Protocol.model_json_schema(), indent=2))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app()
