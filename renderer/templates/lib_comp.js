/**
 * lib_comp.js -- Computational protocol template helpers for ProtocolFormatter.
 *
 * Extends lib.js for computational protocols. Re-uses all shared primitives
 * (header, footer, callouts, step rendering, notes, references) and overrides
 * or augments:
 *
 *   - Body and heading fonts unified to Calibri (no Georgia)
 *   - Prerequisites section (Software / Access / Dependencies)
 *   - Code block rendering: Courier New, cornflower left border, light grey bg
 *   - Procedure rendering with optional preamble paragraphs before steps
 *
 * Template identity
 * -----------------
 *   Body font   : Calibri 10.5pt
 *   Heading font: Calibri
 *   Code font   : Courier New 9.5pt
 *   Code block  : cornflower left border, light grey background
 *   All other visual conventions: identical to lib.js
 */

"use strict";

const {
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  BorderStyle,
  WidthType,
  ShadingType,
  AlignmentType,
} = require("docx");

const lib = require("./lib.js");

const {
  CONTENT_WIDTH_DXA,
  SIZE,
  FONT,
  COLOR,
  parseInline,
  spacer,
  buildHeader,
  buildFooter,
  buildNumberingConfig,
  heading1,
  heading2,
  heading3,
  calloutBox,
  renderSteps,
  renderOverview,
  renderNotes,
  renderReferences,
  endOfProtocol,
  renderTitleBlock,
} = lib;

// ---------------------------------------------------------------------------
// Computational font overrides
// ---------------------------------------------------------------------------

/**
 * Font options that replace Georgia body text with Calibri for computational
 * protocols. Pass this object to any shared helper that accepts fontOpts.
 */
const COMP_FONT_OPTS = {
  font: FONT.HEADING,   // Calibri - FONT.HEADING is Calibri in lib.js
  size: SIZE.BODY,      // 10.5pt
};

// ---------------------------------------------------------------------------
// Code block rendering
// ---------------------------------------------------------------------------

/**
 * Render a code block as a single-cell table with:
 *   - Courier New 9.5pt monospace text
 *   - Cornflower left border (accent bar matching H2 style)
 *   - Light grey background
 *   - Optional language comment line (e.g. "# Terminal / bash")
 *
 * Code lines are split on newline characters. Each line becomes a separate
 * Paragraph within the cell to avoid \n in TextRun (docx npm requirement).
 *
 * @param {string} code   - Raw code or command text (may contain \n).
 * @param {string} [lang] - Optional language label shown as a comment line.
 * @returns {Table}
 */
function codeBlock(code, lang = null) {
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const leftBorder = {
    style: BorderStyle.SINGLE,
    size: 20,
    color: COLOR.CORNFLOWER,
  };

  const cellBorders = {
    top: noBorder,
    bottom: noBorder,
    right: noBorder,
    left: leftBorder,
  };

  const codeRunOpts = {
    font: FONT.CODE,
    size: SIZE.CODE,
    color: COLOR.BLACK,
  };

  const lines = [];

  // Optional language comment line
  if (lang) {
    lines.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `# ${lang}`,
            ...codeRunOpts,
            color: "767676",   // muted grey for comment
            italics: true,
          }),
        ],
        spacing: { after: 40 },
      })
    );
  }

  // Split on literal \n or escaped \\n in the JSON payload
  const codeLines = code.split(/\r?\n/);
  for (const line of codeLines) {
    lines.push(
      new Paragraph({
        children: [new TextRun({ text: line || " ", ...codeRunOpts })],
        spacing: { after: 0 },
      })
    );
  }

  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH_DXA],
    margins: { top: 80, bottom: 80 },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: cellBorders,
            width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
            shading: { fill: COLOR.LIGHT_GREY, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 180, right: 120 },
            children: lines,
          }),
        ],
      }),
    ],
  });
}

// ---------------------------------------------------------------------------
// Prerequisites section
// ---------------------------------------------------------------------------

/**
 * Render the Prerequisites section (computational protocols only).
 *
 * Each non-null field (software, access, dependencies) is rendered as a
 * labelled paragraph:
 *   "Software: " [bold] + value text
 *
 * @param {object} prerequisites - Prerequisites schema object.
 * @returns {Array<Paragraph>}
 */
function renderPrerequisites(prerequisites) {
  if (!prerequisites) return [];

  const elements = [...heading1("Prerequisites")];
  const fields = [
    { key: "software", label: "Software" },
    { key: "access", label: "Access" },
    { key: "dependencies", label: "Dependencies" },
  ];

  for (const { key, label } of fields) {
    const value = prerequisites[key];
    if (!value) continue;

    elements.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `${label}:\u00A0`,
            font: FONT.HEADING,
            size: SIZE.BODY,
            bold: true,
            color: COLOR.NAVY,
          }),
          ...parseInline(value, COMP_FONT_OPTS),
        ],
        spacing: { after: 80 },
      })
    );
  }

  elements.push(spacer(80));
  return elements;
}

// ---------------------------------------------------------------------------
// Computational procedure rendering
// ---------------------------------------------------------------------------

/**
 * Render the Procedure section for a computational protocol.
 *
 * Differences from wet-lab renderProcedure:
 *   - Body font is Calibri (via COMP_FONT_OPTS) not Georgia
 *   - Section preamble paragraphs are rendered before steps
 *   - Steps containing code blocks are detected by the presence of a
 *     code_block field on an ActionStep; if present, a codeBlock() table is
 *     appended after the step text paragraph.
 *
 * Note: The current schema stores code blocks as ActionStep.text containing
 * backtick-fenced content (e.g. "`command --flag`"). The renderer detects
 * multi-line fenced blocks (```...```) and splits them into a prose paragraph
 * + codeBlock table. Single-line inline code (single backticks) is rendered
 * as Courier New inline within the step text.
 *
 * @param {object[]} procedureSections - Array of ProcedureSection objects.
 * @returns {Array<Paragraph|Table>}
 */
function renderCompProcedure(procedureSections) {
  const elements = [...heading1("Procedure")];

  for (const section of procedureSections) {
    elements.push(heading2(section.heading));

    // Preamble paragraph(s) before steps
    if (section.preamble) {
      elements.push(
        new Paragraph({
          children: parseInline(section.preamble, COMP_FONT_OPTS),
          spacing: { after: 80 },
        })
      );
    }

    // Callouts at top of section
    for (const callout of section.callouts || []) {
      elements.push(calloutBox(callout, COMP_FONT_OPTS));
      elements.push(spacer(60));
    }

    // Steps (rendered with Calibri font override)
    elements.push(...renderCompSteps(section.steps || [], 0));
    elements.push(spacer(80));
  }

  return elements;
}

/**
 * Recursively render computational protocol steps.
 * Detects fenced code blocks (``` ... ```) within ActionStep.text and
 * splits them into a prose paragraph + codeBlock table.
 * Inline backtick code (`...`) is rendered in Courier New inline.
 *
 * @param {object[]} steps - Array of Step objects.
 * @param {number}   depth - Current nesting depth.
 * @returns {Array<Paragraph|Table>}
 */
function renderCompSteps(steps, depth = 0) {
  const elements = [];

  for (const step of steps) {
    if (step.step_type === "stopping_point") {
      elements.push(lib.stoppingPoint(step.text, COMP_FONT_OPTS));
      continue;
    }

    // Detect fenced code block: text starts with ``` or contains ```
    const fencedMatch = step.text.match(/^```(?:(\w+)\n)?([\s\S]*?)```\s*$/s);
    const hasFencedBlock = Boolean(fencedMatch);

    if (hasFencedBlock) {
      const lang = fencedMatch[1] || null;
      const code = fencedMatch[2].trimEnd();
      elements.push(codeBlock(code, lang));
    } else {
      // Render inline backticks as Courier New runs
      const runs = parseCompInline(step.text);
      const numRef = getNumRef(depth);

      const para = numRef
        ? new Paragraph({
            numbering: { reference: numRef, level: 0 },
            children: runs,
            spacing: { after: 60 },
          })
        : new Paragraph({
            children: runs,
            spacing: { after: 60 },
            indent: { left: 720 + depth * 360 },
          });

      elements.push(para);
    }

    // Recurse into children
    if (step.children && step.children.length > 0) {
      elements.push(...renderCompSteps(step.children, depth + 1));
    }
  }

  return elements;
}

/**
 * Determine the numbering reference for a given depth level.
 * Mirrors the logic in lib.js NUMBERING_REFS.
 *
 * @param {number} depth
 * @returns {string|null}
 */
function getNumRef(depth) {
  const refs = [
    lib.NUMBERING_REFS.STEP_L0,
    lib.NUMBERING_REFS.STEP_L1,
    lib.NUMBERING_REFS.STEP_L2,
  ];
  return depth < refs.length ? refs[depth] : null;
}

/**
 * Parse text with bold/italic markers AND inline code backticks.
 * Returns an array of TextRun objects suitable for a Paragraph's children.
 *
 * Inline `code` spans are rendered in Courier New at CODE size.
 * All non-code spans are rendered in Calibri at BODY size.
 *
 * @param {string} text - Input text.
 * @returns {TextRun[]}
 */
function parseCompInline(text) {
  if (!text) return [new TextRun({ text: "", ...COMP_FONT_OPTS })];

  const runs = [];
  // Pattern: `code`, **bold**, _italic_, plain chunk, or stray punctuation
  const pattern = /`([^`]+)`|_\*\*(.+?)\*\*_|\*\*(.+?)\*\*|_(.+?)_|([^`*_]+|[`*_])/g;
  let match;

  while ((match = pattern.exec(text)) !== null) {
    if (match[1] !== undefined) {
      // Inline code
      runs.push(
        new TextRun({ text: match[1], font: FONT.CODE, size: SIZE.CODE })
      );
    } else if (match[2] !== undefined) {
      // Bold + italic
      runs.push(new TextRun({ text: match[2], ...COMP_FONT_OPTS, bold: true, italics: true }));
    } else if (match[3] !== undefined) {
      // Bold
      runs.push(new TextRun({ text: match[3], ...COMP_FONT_OPTS, bold: true }));
    } else if (match[4] !== undefined) {
      // Italic
      runs.push(new TextRun({ text: match[4], ...COMP_FONT_OPTS, italics: true }));
    } else if (match[5] !== undefined) {
      // Plain
      runs.push(new TextRun({ text: match[5], ...COMP_FONT_OPTS }));
    }
  }

  return runs.length > 0 ? runs : [new TextRun({ text, ...COMP_FONT_OPTS })];
}

// ---------------------------------------------------------------------------
// Main computational document builder
// ---------------------------------------------------------------------------

/**
 * Build a complete computational protocol Document from a validated Protocol
 * object.
 *
 * @param {object} protocol - Deserialised Protocol schema object.
 * @returns {Document}      - docx Document ready for Packer.toBuffer().
 */
function buildComputationalDocument(protocol) {
  const {
    Document,
  } = require("docx");

  // Abbreviated title for footer (truncate at 40 chars)
  const abbrevTitle =
    protocol.title.length > 40
      ? protocol.title.slice(0, 37) + "\u2026"
      : protocol.title;

  // Assemble body content (Calibri font throughout)
  const body = [
    ...renderTitleBlock(protocol, COMP_FONT_OPTS),
    ...renderCompOverview(protocol.overview),
    ...renderPrerequisites(protocol.prerequisites),
    ...renderCompProcedure(protocol.procedure),
    ...renderCompNotes(protocol.notes),
    ...renderReferences(protocol.references),
    spacer(40),
    endOfProtocol(),
  ];

  return new Document({
    numbering: {
      config: buildNumberingConfig(),
    },
    styles: {
      default: {
        document: {
          run: {
            font: FONT.HEADING,   // Calibri throughout
            size: SIZE.BODY,
            color: COLOR.BLACK,
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: lib.PAGE_WIDTH_DXA,
              height: lib.PAGE_HEIGHT_DXA,
            },
            margin: {
              top: lib.MARGIN_DXA,
              bottom: lib.MARGIN_DXA,
              left: lib.MARGIN_DXA,
              right: lib.MARGIN_DXA,
              header: 600,
              footer: 600,
            },
          },
        },
        headers: {
          default: buildHeader(protocol.section_number, protocol.section_name),
        },
        footers: {
          default: buildFooter(abbrevTitle, protocol.version || "1.0"),
        },
        children: body,
      },
    ],
  });
}

// ---------------------------------------------------------------------------
// Calibri overrides for shared section renderers
// ---------------------------------------------------------------------------

/**
 * Overview section rendered in Calibri (computational default).
 * DEV-004: Splits on \n\n to produce separate Paragraph objects.
 *
 * @param {string} text
 * @returns {Array<Paragraph>}
 */
function renderCompOverview(text) {
  const chunks = (text || "").split(/\n\n+/).filter(c => c.trim());
  if (chunks.length === 0) {
    chunks.push(text || "");
  }
  return [
    ...heading1("Overview"),
    ...chunks.map(chunk =>
      new Paragraph({
        children: parseInline(chunk.trim(), COMP_FONT_OPTS),
        spacing: { after: 120 },
      })
    ),
  ];
}

/**
 * Notes section rendered in Calibri.
 * @param {string[]} notes
 * @returns {Array<Paragraph>}
 */
function renderCompNotes(notes) {
  if (!notes || notes.length === 0) return [];

  const elements = [...heading1("Notes & Variants")];
  notes.forEach((note, i) => {
    elements.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `${i + 1}.\u00A0\u00A0`,
            ...COMP_FONT_OPTS,
            bold: true,
          }),
          ...parseInline(note, COMP_FONT_OPTS),
        ],
        spacing: { after: 80 },
        indent: { left: 360, hanging: 360 },
      })
    );
  });
  return elements;
}

// ---------------------------------------------------------------------------
// Exports
// ---------------------------------------------------------------------------

module.exports = {
  // Computational-specific
  COMP_FONT_OPTS,
  codeBlock,
  renderPrerequisites,
  renderCompProcedure,
  renderCompSteps,
  renderCompOverview,
  renderCompNotes,
  parseCompInline,
  buildComputationalDocument,
};
