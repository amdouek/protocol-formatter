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
from typing import Literal, Optional, Union
from pydantic import BaseModel, Field, model_validator


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
    An inline callout box rendered within a procedure section.

    Callouts may appear before or after any step within a ProcedureSection.
    The renderer places them at the position indicated by their containing list.
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


# ---------------------------------------------------------------------------
# Procedure step models
# ---------------------------------------------------------------------------

class ActionStep(BaseModel):
    """
    A discrete procedural action.

    Steps support arbitrary nesting via the children field, mirroring the
    indentation levels present in source documents (main step → sub-step →
    tertiary detail). There is no fixed depth limit.
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
Step = Union[ActionStep, StoppingPoint]

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
            "Callout boxes associated with this section. The renderer places each "
            "callout at the position it appears in this list, interleaved with steps "
            "if the source document indicates a specific position. If position is "
            "ambiguous, callouts are placed at the top of the section."
        ),
    )


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

    # --- Front matter ---
    overview: str = Field(
        ...,
        min_length=1,
        description=(
            "1–3 sentence overview of the protocol's purpose, expected duration, "
            "and biological or computational context. Mention critical prerequisites "
            "or linked protocols. Supports **bold** and _italic_ markers."
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
