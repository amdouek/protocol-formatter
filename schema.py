"""
schema.py -- Pydantic v2 intermediate representation for ProtocolFormatter.

This module defines the template-agnostic data model that sits between the
input parser and the Node.js renderer. All fields describe content semantics,
not presentation. The renderer is responsible for all visual decisions.

Inline formatting convention
----------------------------
Text fields that flow into rendered paragraphs support two inline markers:
    **text**   → bold
    _text_     → italic

These markers are interpreted by the renderer and must not be escaped or
modified by the parser or extractor.

Section types
-------------
    wet_lab        → lib.js template (Georgia body, materials tables, etc.)
    computational  → lib_comp.js template (Calibri throughout, code blocks)
"""

from __future__ import annotations

from enum import Enum

import re
from typing import Annotated, Literal, Optional, Union
from pydantic import BaseModel, Discriminator, Field, Tag, field_validator, model_validator


# ---------------------------------------------------------------------------
# Enumerations
# ---------------------------------------------------------------------------

class SectionType(str, Enum):
    """Determines which Node.js template the renderer selects."""
    WET_LAB = "wet_lab"
    COMPUTATIONAL = "computational"


class CalloutType(str, Enum):
    """
    Visual callout variants.

    Critical  → red   — protocol will fail if ignored
    Caution   → amber — safety or quality risk
    Tip       → green — optional optimisation or shortcut
    Note      → blue  — contextual information, cross-references, caveats
    """
    CRITICAL = "critical"
    CAUTION = "caution"
    TIP = "tip"
    NOTE = "note"


class StepType(str, Enum):
    """Discriminator for the Step union."""
    ACTION = "action"
    STOPPING_POINT = "stopping_point"


# ---------------------------------------------------------------------------
# Shared leaf models
# ---------------------------------------------------------------------------

class Callout(BaseModel):
    """
    A callout box rendered within a procedure section.

    Callouts may render either at the top of their containing ProcedureSection
    (when ``after_step`` is None) or inline after a specific step (when
    ``after_step`` is the 0-based index of the triggering step).
    """
    callout_type: CalloutType = Field(
        ...,
        description="Visual variant of the callout box.",
    )
    text: str = Field(
        ...,
        min_length=1,
        description=(
            "Body text of the callout. Supports **bold** and _italic_ markers. "
            "Do not include the callout label (e.g. 'Caution:') — the renderer "
            "adds this automatically."
        ),
    )
    after_step: Optional[int] = Field(
        default=None,
        ge=0,
        description=(
            "0-based index of the step in section.steps after which this "
            "callout should render. None (default) places the callout at the "
            "top of the section, after any preamble but before the first step. "
            "Used for general section-level warnings. Set to a step index to "
            "render the callout inline immediately after that step. "
        ),
    )


class TableRow(BaseModel):
    """A single row in a materials, mix, or recipe table."""
    cells: list[str] = Field(
        ...,
        min_length=1,
        description="Ordered cell values for this row.",
    )
    bold: bool = Field(
        default=False,
        description=(
            "If True, all cells in this row are rendered bold. "
            "Used for 'Total' and 'Incubate' rows in mix/recipe tables."
        ),
    )


class MaterialsTable(BaseModel):
    """
    A three-column reagents or equipment table.

    Columns are always: Reagent / Equipment | Supplier | Cat. No. / Notes.
    Use '—' or 'Lab supply' for unknown supplier or catalogue values.
    """
    heading: str = Field(
        ...,
        description="Sub-section heading, e.g. 'Reagents' or 'Equipment & consumables'.",
    )
    rows: list[TableRow] = Field(
        ...,
        min_length=1,
        description="Data rows. Do not include the header row — the renderer adds it.",
    )

    @model_validator(mode="after")
    def validate_row_widths(self) -> "MaterialsTable":
        """Ensure every row has exactly 3 cells (Reagent / Supplier / Cat. No.)."""
        for i, row in enumerate(self.rows):
            if len(row.cells) != 3:
                raise ValueError(
                    f"MaterialsTable row {i} has {len(row.cells)} cell(s); "
                    f"expected 3 (Reagent / Supplier / Cat. No.)."
                )
        return self


class MixTable(BaseModel):
    """
    A two-column component/amount table for mixes, master mixes, or recipes.

    Rendered after the References section, before the end-of-protocol line.
    Rows whose first cell begins with 'Total' or 'Incubate' are automatically
    bolded by the renderer (the bold flag on TableRow may also be set explicitly).

    DEV-002: Placement is post-References, pre end-of-protocol footer.
    """
    heading: str = Field(
        ...,
        description="Table heading, e.g. 'Lysis buffer master mix (per sample)'.",
    )
    rows: list[TableRow] = Field(
        ...,
        min_length=1,
        description=(
            "Data rows. Columns are Component | Amount. "
            "Do not include the header row — the renderer adds it."
        ),
    )

    @model_validator(mode="after")
    def validate_row_widths(self) -> "MixTable":
        """Ensure every row has exactly 2 cells (Component / Amount)."""
        for i, row in enumerate(self.rows):
            if len(row.cells) != 2:
                raise ValueError(
                    f"MixTable row {i} has {len(row.cells)} cell(s); "
                    f"expected 2 (Component / Amount)."
                )
        return self


# ---------------------------------------------------------------------------
# Procedure step models
# ---------------------------------------------------------------------------

class ActionStep(BaseModel):
    """
    A discrete procedural action.

    Steps support arbitrary nesting via the children field, mirroring the
    indentation levels present in source documents (main step → sub-step →
    tertiary detail). There is no fixed depth limit.

    For computational protocols, a step may include an associated code block
    rendered after the step text in Courier New with a cornflower left border.
    The ``code`` field is separate from ``text`` so that prose instructions
    and executable commands have distinct semantic identities.
    """
    step_type: Literal[StepType.ACTION] = StepType.ACTION

    text: str = Field(
        ...,
        min_length=1,
        description=(
            "Instruction text for this step. Supports **bold** and _italic_ markers. "
            "Key parameters (temperatures, volumes, times) should be bolded."
        ),
    )
    code: Optional[str] = Field(
        default=None,
        description=(
            "Code block content associated with this step. Rendered after the "
            "step text in monospace font with a cornflower left border. Used in "
            "computational protocols for terminal commands, scripts, or config "
            "snippets. None for wet-lab protocols or steps with no code."
        ),
    )
    code_language: Optional[str] = Field(
        default=None,
        description=(
            "Optional language label for the code block (e.g. 'bash', 'python', "
            "'R'). Rendered as a muted comment line above the code. None if the "
            "language is not specified or not applicable."
        ),
    )
    children: list[Step] = Field(
        default_factory=list,
        description=(
            "Nested sub-steps. Each child is itself a Step (ActionStep or "
            "StoppingPoint), enabling arbitrary depth."
        ),
    )


class StoppingPoint(BaseModel):
    """
    A safe pause point in the procedure.

    Rendered inline within the step list with a ⏸ symbol. Triggered by source
    phrases such as 'pause point', 'can be stored at this stage', or
    'samples can be stored'.
    """
    step_type: Literal[StepType.STOPPING_POINT] = StepType.STOPPING_POINT

    text: str = Field(
        ...,
        min_length=1,
        description=(
            "Description of storage conditions or pause rationale, "
            "e.g. 'Samples can be stored at −20 °C for up to one week.' "
            "Supports **bold** and _italic_ markers."
        ),
    )
    children: list[Step] = Field(
        default_factory=list,
        description="Rarely populated. Included for structural consistency with ActionStep.",
    )


# Step is the discriminated union used throughout the procedure tree.
# The explicit Discriminator on step_type gives Pydantic a direct dispatch
# path, producing targeted validation errors (e.g. "step_type 'action'
# received but 'text' field missing") rather than the generic "none of the
# union variants matched" message that untagged unions produce.
Step = Annotated[
    Union[
        Annotated[ActionStep, Tag(StepType.ACTION)],
        Annotated[StoppingPoint, Tag(StepType.STOPPING_POINT)],
    ],
    Discriminator("step_type"),
]

# Rebuild models that reference Step forward reference.
ActionStep.model_rebuild()
StoppingPoint.model_rebuild()


# ---------------------------------------------------------------------------
# Procedure section
# ---------------------------------------------------------------------------

class ProcedureSection(BaseModel):
    """
    A named group of steps within the Procedure, corresponding to an H2 heading.

    Example headings: 'Sample preparation', 'Reaction assembly', 'Step 1 — Clone repository'.
    """
    heading: str = Field(
        ...,
        min_length=1,
        description="H2-level heading for this step group.",
    )
    preamble: Optional[str] = Field(
        default=None,
        description=(
            "Optional introductory paragraph rendered before the step list. "
            "Used in computational protocols to describe a step in plain language "
            "before showing commands. Supports **bold** and _italic_ markers."
        ),
    )
    steps: list[Step] = Field(
        default_factory=list,
        description="Ordered list of steps and stopping points in this section.",
    )
    callouts: list[Callout] = Field(
        default_factory=list,
        description=(
            "Callout boxes associated with this section. Each callout's "
            "after_step field controls placement: None for top-of-section "
            "(after any preamble, before the first step), or a 0-based step "
            "index for inline placement after that step."
        ),
    )

    @model_validator(mode="after")
    def validate_callout_positions(self) -> "ProcedureSection":
        """Ensure every callout's after_step is a valid index into self.steps."""
        n_steps = len(self.steps)
        for i, callout in enumerate(self.callouts):
            if callout.after_step is None:
                continue
            if callout.after_step >= n_steps:
                raise ValueError(
                    f"Callout {i} in section '{self.heading}' has "
                    f"after_step={callout.after_step}, but section has only "
                    f"{n_steps} step(s). Valid range: 0 to {n_steps - 1} "
                    f"inclusive, or None for top-of-section placement."
                )
        return self


# ---------------------------------------------------------------------------
# Computational-only models
# ---------------------------------------------------------------------------

class Prerequisites(BaseModel):
    """
    Software, access, and dependency requirements for a computational protocol.

    Present only when section_type == 'computational'. The renderer raises an
    error if this field is populated on a wet_lab protocol.
    """
    software: Optional[str] = Field(
        default=None,
        description="Required software, versions, and installation notes.",
    )
    access: Optional[str] = Field(
        default=None,
        description="Server access, API keys, or account requirements.",
    )
    dependencies: Optional[str] = Field(
        default=None,
        description="Required packages, libraries, or environment requirements.",
    )


# ---------------------------------------------------------------------------
# Top-level protocol model
# ---------------------------------------------------------------------------

class Protocol(BaseModel):
    """
    Template-agnostic intermediate representation of a single formatted protocol.

    This model is the contract between the Python pipeline and the Node.js
    renderer. The renderer consumes a JSON-serialised instance of this model
    and produces a .docx file.

    Field population rules
    ----------------------
    - materials and mix_tables are wet_lab only.
    - prerequisites is computational only.
    - These constraints are enforced by the model_validator below.
    - All text fields support **bold** and _italic_ inline markers.
    - Footnotes from source documents must be converted to numbered notes entries.
    - Author defaults to 'ARMI' when not recoverable from the source.
    """

    # --- Identity ---
    title: str = Field(
        ...,
        min_length=1,
        description="Full protocol title as it will appear in the document heading.",
    )
    subtitle: Optional[str] = Field(
        default=None,
        description="Brief subtitle describing scope, organism, or tool. Rendered in italics.",
    )
    author: str = Field(
        default="ARMI",
        description=(
            "Protocol author. Defaults to 'ARMI' when not recoverable from the source. "
            "Use the author's full name when identifiable."
        ),
    )
    section_type: SectionType = Field(
        ...,
        description="Determines the rendering template: 'wet_lab' or 'computational'.",
    )
    section_number: str = Field(
        ...,
        description=(
            "Compendium section identifier, e.g. '3.2' or '9.1'. "
            "Used in the document header and for cross-reference formatting."
        ),
    )
    section_name: str = Field(
        ...,
        description=(
            "Human-readable section name, e.g. 'MOLECULAR BIOLOGY' or "
            "'COMPUTATIONAL METHODS & BIOINFORMATICS'. Rendered in the document header."
        ),
    )
    version: str = Field(
        default="1.0",
        description="Protocol version string, e.g. '1.0' or '2.3'.",
    )
    date: Optional[str] = Field(
        default=None,
        description="Last-updated date in DD/MM/YYYY format.",
    )

    @field_validator("date")
    @classmethod
    def validate_date_format(cls, v: Optional[str]) -> Optional[str]:
        """Ensure date matches DD/MM/YYYY format when provided."""
        if v is None:
            return v
        if not re.match(r"^\d{2}/\d{2}/\d{4}$", v):
            raise ValueError(
                f"Date must be in DD/MM/YYYY format, got '{v}'. "
                "Use parser.utils.normalise_date() to convert other formats."
            )
        # Basic range check: day 01-31, month 01-12
        day, month, year = int(v[:2]), int(v[3:5]), int(v[6:])
        if not (1 <= month <= 12):
            raise ValueError(f"Invalid month {month:02d} in date '{v}'.")
        if not (1 <= day <= 31):
            raise ValueError(f"Invalid day {day:02d} in date '{v}'.")
        return v

    # --- Front matter ---
    overview: str = Field(
        ...,
        min_length=0,
        description=(
            "1–3 sentence overview of the protocol's purpose, expected duration, "
            "and biological or computational context. Mention critical prerequisites "
            "or linked protocols. Supports **bold** and _italic_ markers. "
            "Use an empty string if no overview is extractable — the --review step "
            "will flag it for manual completion."
        ),
    )

    # --- Wet-lab only ---
    materials: list[MaterialsTable] = Field(
        default_factory=list,
        description=(
            "Materials tables (Reagent / Supplier / Cat. No.). "
            "Wet-lab protocols only. Leave empty for computational protocols."
        ),
    )

    # --- Computational only ---
    prerequisites: Optional[Prerequisites] = Field(
        default=None,
        description=(
            "Software, access, and dependency prerequisites. "
            "Computational protocols only. Must be None for wet_lab protocols."
        ),
    )

    # --- Procedure ---
    procedure: list[ProcedureSection] = Field(
        ...,
        min_length=1,
        description="Ordered procedure sections. At least one section is required.",
    )

    # --- Back matter ---
    notes: list[str] = Field(
        default_factory=list,
        description=(
            "Numbered notes and variants rendered in the 'Notes & Variants' section. "
            "Footnotes from source documents must be converted to entries here. "
            "Each entry is one note; the renderer adds the number automatically."
        ),
    )
    references: list[str] = Field(
        default_factory=list,
        description=(
            "Reference list entries rendered in the 'References' section. "
            "Each entry is one reference in the author's preferred citation style."
        ),
    )
    mix_tables: list[MixTable] = Field(
        default_factory=list,
        description=(
            "Mix/recipe tables rendered after References, before the end-of-protocol "
            "line. Wet-lab protocols only. Leave empty for computational protocols. "
            "DEVNOTE002: Post-references placement."
        ),
    )

    # --- Duplicate detection ---
    source_filename: Optional[str] = Field(
        default=None,
        description=(
            "Original source filename (basename only, no path). Stored for "
            "duplicate detection. DEVNOTE003: Filename-based guard for batch runs."
        ),
    )

    # --- Cross-section validation ---
    @model_validator(mode="after")
    def validate_section_type_fields(self) -> "Protocol":
        """
        Enforce that wet_lab and computational fields are not mixed.

        Rules:
            - wet_lab protocols must not have prerequisites populated.
            - computational protocols must not have materials or mix_tables populated.
        """
        if self.section_type == SectionType.WET_LAB:
            if self.prerequisites is not None:
                raise ValueError(
                    "Field 'prerequisites' must be None for wet_lab protocols. "
                    "Prerequisites are a computational-only section."
                )
        elif self.section_type == SectionType.COMPUTATIONAL:
            if self.materials:
                raise ValueError(
                    "Field 'materials' must be empty for computational protocols. "
                    "Materials tables are a wet_lab-only section."
                )
            if self.mix_tables:
                raise ValueError(
                    "Field 'mix_tables' must be empty for computational protocols. "
                    "Mix/recipe tables are a wet_lab-only section."
                )
        return self


# ---------------------------------------------------------------------------
# Convenience re-exports
# ---------------------------------------------------------------------------

__all__ = [
    "SectionType",
    "CalloutType",
    "StepType",
    "Callout",
    "TableRow",
    "MaterialsTable",
    "MixTable",
    "ActionStep",
    "StoppingPoint",
    "Step",
    "ProcedureSection",
    "Prerequisites",
    "Protocol",
]
