/**
 * lib.js -- Wet-lab protocol template helpers for ProtocolFormatter.
 *
 * Exports building-block functions used by render.js to compose wet-lab
 * protocol documents. Also exports shared utilities (inline text parsing,
 * callout boxes, header/footer construction) that lib_comp.js re-uses.
 *
 * Template identity
 * -----------------
 *   Body font   : Georgia 10.5pt
 *   Heading font: Calibri
 *   H1          : 14pt, bold, navy (#1F3864), followed by cornflower rule
 *   H2          : 12pt, bold, cornflower left border paragraph
 *   H3          : 10.5pt, bold, Calibri
 *   Page        : A4, 2 cm margins on all sides
 *   Header      : "SECTION X — NAME  ·  Protocol Compendium" (tab-separated)
 *   Footer      : "Abbreviated title    Version X.X    Page N"
 *
 * Colour palette (hex, no leading # - see style_guide.yaml for full colour details)
 * ------------------------------------
 *   Navy        : 1A5FA8
 *   Cornflower  : 3B8BD4
 *   Light grey  : F5F5F5
 *   Callout backgrounds  : FFC5C5 / FFF8E1 / E8F5E9 / EBF3FB
 *   Callout borders      : FF0000 / E8A800 / 2E7D32 / 3B8BD4
 */

"use strict";

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  WidthType,
  ShadingType,
  VerticalAlign,
  PageNumber,
  TabStopType,
  TabStopPosition,
  LevelFormat,
} = require("docx");

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** A4 page dimensions in DXA (1440 DXA = 1 inch, 1 inch ≈ 2.54 cm). */
const PAGE_WIDTH_DXA = 11906;
const PAGE_HEIGHT_DXA = 16838;

/** 2 cm margins in DXA. 1 cm ≈ 567 DXA → 2 cm ≈ 1134 DXA. */
const MARGIN_DXA = 1134;

/** Usable content width = page width − left margin − right margin. */
const CONTENT_WIDTH_DXA = PAGE_WIDTH_DXA - MARGIN_DXA * 2; // 9638

/** Half-point font sizes (docx npm convention: 1pt = 2 units). */
const SIZE = {
  BODY: 21,       // 10.5pt
  H1: 28,         // 14pt
  H2: 24,         // 12pt
  H3: 21,         // 10.5pt
  FOOTER: 18,     // 9pt
  HEADER: 18,     // 9pt
  CODE: 19,       // 9.5pt  (used by lib_comp.js via re-export)
};

/** Font names. */
const FONT = {
  BODY: "Georgia",
  HEADING: "Calibri",
  CODE: "Courier New",
};

/** Colour palette (no leading #). */
const COLOR = {
  NAVY: "1A5FA8",
  CORNFLOWER: "3B8BD4",
  BLACK: "000000",
  WHITE: "FFFFFF",
  LIGHT_GREY: "F5F5F5",
  MID_GREY: "BFBFBF",

  // Callout backgrounds
  BG_CRITICAL: "FFC5C5",
  BG_CAUTION: "FFF8E1",
  BG_TIP: "E8F5E9",
  BG_NOTE: "EBF3FB",

  // Callout left-border colours
  BORDER_CRITICAL: "FF0000",
  BORDER_CAUTION: "E8A800",
  BORDER_TIP: "2E7D32",
  BORDER_NOTE: "3B8BD4",
};

/** Callout metadata keyed by callout_type string from the schema. */
const CALLOUT_META = {
  critical: {
    label: "CRITICAL",
    symbol: "⚠",
    bg: COLOR.BG_CRITICAL,
    border: COLOR.BORDER_CRITICAL,
  },
  caution: {
    label: "Caution",
    symbol: "⚠",
    bg: COLOR.BG_CAUTION,
    border: COLOR.BORDER_CAUTION,
  },
  tip: {
    label: "Tip",
    symbol: "💡",
    bg: COLOR.BG_TIP,
    border: COLOR.BORDER_TIP,
  },
  note: {
    label: "Note",
    symbol: "ℹ",
    bg: COLOR.BG_NOTE,
    border: COLOR.BORDER_NOTE,
  },
};

/** Stopping point symbol. */
const STOP_SYMBOL = "⏸";

// ---------------------------------------------------------------------------
// Inline text parser
// ---------------------------------------------------------------------------

/**
 * Parse a text string containing **bold** and _italic_ markers into an array
 * of docx TextRun objects.
 *
 * Rules:
 *   **text**  → bold TextRun
 *   _text_    → italic TextRun
 *   Plain text → normal TextRun
 *
 * Markers may not be nested. Unclosed markers are treated as literal text.
 *
 * @param {string} text     - Raw text with optional inline markers.
 * @param {object} baseOpts - Additional TextRun options applied to every run
 *                            (e.g. { font: "Calibri", size: 21 }).
 * @returns {TextRun[]}
 */
function parseInline(text, baseOpts = {}) {
  if (!text) return [new TextRun({ text: "", ...baseOpts })];

  const runs = [];
  // Regex alternation: **...** then _..._ then plain text chunk
  const pattern = /_\*\*(.+?)\*\*_|\*\*(.+?)\*\*|_(.+?)_|([^*_]+|[*_])/g;
  let match;

  while ((match = pattern.exec(text)) !== null) {
    if (match[1] !== undefined) {
      // Bold + italic span
      runs.push(new TextRun({ text: match[1], bold: true, italics: true, ...baseOpts }));
    } else if (match[2] !== undefined) {
      // Bold span
      runs.push(new TextRun({ text: match[2], bold: true, ...baseOpts }));
    } else if (match[3] !== undefined) {
      // Italic span
      runs.push(new TextRun({ text: match[3], italics: true, ...baseOpts }));
    } else if (match[4] !== undefined) {
      // Plain text
      runs.push(new TextRun({ text: match[4], ...baseOpts }));
    }
  }

  return runs.length > 0 ? runs : [new TextRun({ text, ...baseOpts })];
}

// ---------------------------------------------------------------------------
// Shared empty paragraph
// ---------------------------------------------------------------------------

/**
 * A blank spacer paragraph. Used between major sections.
 * @param {number} [spacingAfter=0] - spacing after in twips.
 */
function spacer(spacingAfter = 0) {
  return new Paragraph({
    children: [new TextRun("")],
    spacing: { after: spacingAfter },
  });
}

// ---------------------------------------------------------------------------
// Header and footer
// ---------------------------------------------------------------------------

/**
 * Build the document header paragraph.
 * Left: "SECTION X — SECTION NAME  ·  Protocol Compendium"
 *
 * @param {string} sectionNumber - e.g. "3"
 * @param {string} sectionName   - e.g. "CLONING & PLASMID ENGINEERING"
 * @returns {Header}
 */
function buildHeader(sectionNumber, sectionName) {
  const leftText = `SECTION ${sectionNumber} \u2014 ${sectionName}`;
  const rightText = "Protocol Compendium";

  return new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: leftText,
            font: FONT.HEADING,
            size: SIZE.HEADER,
            color: COLOR.NAVY,
            bold: true,
          }),
          new TextRun({
            text: `\u00A0\u00A0\u00B7\u00A0\u00A0`, // " · "
            font: FONT.HEADING,
            size: SIZE.HEADER,
            color: COLOR.CORNFLOWER,
          }),
          new TextRun({
            text: rightText,
            font: FONT.HEADING,
            size: SIZE.HEADER,
            color: COLOR.CORNFLOWER,
          }),
        ],
        border: {
          bottom: {
            style: BorderStyle.SINGLE,
            size: 6,
            color: COLOR.CORNFLOWER,
            space: 4,
          },
        },
        spacing: { after: 80 },
      }),
    ],
  });
}

/**
 * Build the document footer paragraph.
 * Layout (tab-separated): "Abbreviated title    Version X.X    Page N"
 *
 * @param {string} abbreviatedTitle - Short protocol title for the footer.
 * @param {string} version          - Version string, e.g. "1.0".
 * @returns {Footer}
 */
function buildFooter(abbreviatedTitle, version) {
  return new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: abbreviatedTitle,
            font: FONT.HEADING,
            size: SIZE.FOOTER,
            color: COLOR.MID_GREY,
          }),
          new TextRun({
            text: `\tVersion ${version}\t`,
            font: FONT.HEADING,
            size: SIZE.FOOTER,
            color: COLOR.MID_GREY,
          }),
          new TextRun({
            text: "Page\u00A0",
            font: FONT.HEADING,
            size: SIZE.FOOTER,
            color: COLOR.MID_GREY,
          }),
          new TextRun({
            children: [PageNumber.CURRENT],
            font: FONT.HEADING,
            size: SIZE.FOOTER,
            color: COLOR.MID_GREY,
          }),
        ],
        tabStops: [
          {
            type: TabStopType.CENTER,
            position: Math.round(CONTENT_WIDTH_DXA / 2),
          },
          {
            type: TabStopType.RIGHT,
            position: CONTENT_WIDTH_DXA,
          },
        ],
        border: {
          top: {
            style: BorderStyle.SINGLE,
            size: 4,
            color: COLOR.MID_GREY,
            space: 4,
          },
        },
        spacing: { before: 80 },
      }),
    ],
  });
}

// ---------------------------------------------------------------------------
// H1 heading with cornflower underrule
// ---------------------------------------------------------------------------

/**
 * Render a top-level section heading (H1).
 * Returns two paragraphs: the heading text in navy Calibri, followed by a
 * cornflower horizontal rule implemented as a paragraph bottom border.
 *
 * @param {string} text - Heading text (no inline markers as H1s are plain).
 * @returns {Paragraph[]}
 */
function heading1(text) {
  return [
    new Paragraph({
      children: [
        new TextRun({
          text,
          font: FONT.HEADING,
          size: SIZE.H1,
          bold: true,
          color: COLOR.NAVY,
        }),
      ],
      spacing: { before: 240, after: 0 },
      border: {
        bottom: {
          style: BorderStyle.SINGLE,
          size: 8,
          color: COLOR.CORNFLOWER,
          space: 4,
        },
      },
    }),
    spacer(120),
  ];
}

// ---------------------------------------------------------------------------
// H2 heading with cornflower left bar
// ---------------------------------------------------------------------------

/**
 * Render a procedure section heading (H2).
 * Styled with a thick cornflower left border to create the visual left-bar
 * effect seen in the template.
 *
 * @param {string} text - Heading text (no inline markers).
 * @returns {Paragraph}
 */
function heading2(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        font: FONT.HEADING,
        size: SIZE.H2,
        bold: true,
        color: COLOR.NAVY,
      }),
    ],
    border: {
      left: {
        style: BorderStyle.SINGLE,
        size: 24,           // thick left bar
        color: COLOR.CORNFLOWER,
        space: 8,
      },
    },
    indent: { left: 160 },
    spacing: { before: 200, after: 80 },
  });
}

// ---------------------------------------------------------------------------
// H3 sub-heading
// ---------------------------------------------------------------------------

/**
 * Render a sub-section heading (H3).
 * Used for Materials sub-headings ("Reagents", "Equipment & consumables").
 *
 * @param {string} text - Heading text.
 * @returns {Paragraph}
 */
function heading3(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        font: FONT.HEADING,
        size: SIZE.H3,
        bold: true,
        color: COLOR.BLACK,
      }),
    ],
    spacing: { before: 160, after: 60 },
  });
}

// ---------------------------------------------------------------------------
// Callout boxes
// ---------------------------------------------------------------------------

/**
 * Render a callout box as a single-cell table with a coloured left border
 * and tinted background.
 *
 * The label line ("⚠  Caution") is rendered bold; the body text follows on
 * a new paragraph within the same cell, supporting inline **bold** and
 * _italic_ markers.
 *
 * @param {object} callout          - Callout object from the schema.
 * @param {string} callout.callout_type - "critical" | "caution" | "tip" | "note"
 * @param {string} callout.text         - Body text with optional inline markers.
 * @param {object} [bodyFontOpts]       - Font overrides for body text runs.
 * @returns {Table}
 */
function calloutBox(callout, bodyFontOpts = {}) {
  const meta = CALLOUT_META[callout.callout_type];
  if (!meta) {
    throw new Error(`Unknown callout_type: "${callout.callout_type}"`);
  }

  const defaultFont = { font: FONT.BODY, size: SIZE.BODY, ...bodyFontOpts };
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

  // Left accent border
  const leftBorder = {
    style: BorderStyle.SINGLE,
    size: 24,
    color: meta.border,
  };

  const cellBorders = {
    top: noBorder,
    bottom: noBorder,
    right: noBorder,
    left: leftBorder,
  };

  // Label paragraph: "⚠  Caution" in bold
  const labelPara = new Paragraph({
    children: [
      new TextRun({
        text: `${meta.symbol}\u00A0\u00A0`,
        ...defaultFont,
        bold: true,
      }),
      new TextRun({
        text: meta.label,
        ...defaultFont,
        bold: true,
      }),
    ],
    spacing: { after: 60 },
  });

  // Body paragraph with inline marker parsing
  const bodyRuns = parseInline(callout.text, defaultFont);
  const bodyPara = new Paragraph({
    children: bodyRuns,
    spacing: { after: 0 },
  });

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
            shading: { fill: meta.bg, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 180, right: 120 },
            children: [labelPara, bodyPara],
          }),
        ],
      }),
    ],
  });
}

// ---------------------------------------------------------------------------
// Stopping point
// ---------------------------------------------------------------------------

/**
 * Render an inline stopping point paragraph with ⏸ symbol.
 *
 * @param {string} text - Stopping point description with optional inline markers.
 * @param {object} [fontOpts] - Font overrides.
 * @returns {Paragraph}
 */
function stoppingPoint(text, fontOpts = {}) {
  const opts = { font: FONT.BODY, size: SIZE.BODY, ...fontOpts };
  return new Paragraph({
    children: [
      new TextRun({ text: `${STOP_SYMBOL}\u00A0\u00A0`, ...opts, bold: true }),
      new TextRun({ text: "Safe stopping point\u00A0\u2014\u00A0", ...opts, bold: true }),
      ...parseInline(text, opts),
    ],
    spacing: { before: 60, after: 60 },
    indent: { left: 360 },
  });
}

// ---------------------------------------------------------------------------
// Procedure step rendering
// ---------------------------------------------------------------------------

/**
 * Numbering configuration references used for procedure steps.
 * Each nesting level uses a distinct reference so that step numbering resets
 * correctly at each indentation level.
 *
 * These references are declared in the Document numbering config (see
 * buildNumberingConfig) and referenced here by name.
 */
const NUMBERING_REFS = {
  STEP_L0: "step-level-0",   // main steps
  STEP_L1: "step-level-1",   // sub-steps (a, b, c …)
  STEP_L2: "step-level-2",   // tertiary (roman or bullet)
};

/**
 * Build the numbering configuration object for wet-lab procedure steps.
 * Returns the value for Document({ numbering: { config: [...] } }).
 *
 * Level 0 : arabic numerals   1. 2. 3.
 * Level 1 : lower-alpha       a. b. c.
 * Level 2 : lower-roman       i. ii. iii.
 *
 * @returns {object[]} Array of numbering config entries.
 */
function buildNumberingConfig() {
  function levelDef(level, format, text, indent) {
    return {
      level,
      format,
      text,
      alignment: AlignmentType.LEFT,
      style: {
        paragraph: {
          indent: { left: indent, hanging: 360 },
          spacing: { after: 60 },
        },
        run: { font: FONT.BODY, size: SIZE.BODY },
      },
    };
  }

  return [
    {
      reference: NUMBERING_REFS.STEP_L0,
      levels: [levelDef(0, LevelFormat.DECIMAL, "%1.", 720)],
    },
    {
      reference: NUMBERING_REFS.STEP_L1,
      levels: [levelDef(0, LevelFormat.LOWER_LETTER, "%1.", 1080)],
    },
    {
      reference: NUMBERING_REFS.STEP_L2,
      levels: [levelDef(0, LevelFormat.LOWER_ROMAN, "%1.", 1440)],
    },
  ];
}

/**
 * Recursively render a step tree into an array of Paragraphs and Tables.
 *
 * @param {object[]} steps    - Array of Step objects (ActionStep | StoppingPoint).
 * @param {number}   depth    - Current nesting depth (0 = top-level).
 * @param {object}   fontOpts - Font options for body text.
 * @returns {Array<Paragraph|Table>}
 */
function renderSteps(steps, depth = 0, fontOpts = {}) {
  const opts = { font: FONT.BODY, size: SIZE.BODY, ...fontOpts };
  const elements = [];

  // Choose the numbering reference for this depth level.
  const refMap = [
    NUMBERING_REFS.STEP_L0,
    NUMBERING_REFS.STEP_L1,
    NUMBERING_REFS.STEP_L2,
  ];
  // Beyond depth 2, fall back to bullet points with increasing indent.
  const numRef = depth < refMap.length ? refMap[depth] : null;

  for (const step of steps) {
    if (step.step_type === "stopping_point") {
      elements.push(stoppingPoint(step.text, fontOpts));
    } else {
      // ActionStep
      const para = numRef
        ? new Paragraph({
            numbering: { reference: numRef, level: 0 },
            children: parseInline(step.text, opts),
            spacing: { after: 60 },
          })
        : new Paragraph({
            bullet: { level: depth - refMap.length },
            children: parseInline(step.text, opts),
            spacing: { after: 60 },
            indent: { left: 720 + depth * 360 },
          });

      elements.push(para);

      // Recurse into children
      if (step.children && step.children.length > 0) {
        elements.push(...renderSteps(step.children, depth + 1, fontOpts));
      }
    }
  }

  return elements;
}

// ---------------------------------------------------------------------------
// Materials table (3-column: Reagent / Supplier / Cat. No.)
// ---------------------------------------------------------------------------

/**
 * Column widths for a 3-column materials table summing to CONTENT_WIDTH_DXA.
 * Reagent: ~50%, Supplier: ~28%, Cat. No.: ~22%
 */
const MAT_COL_WIDTHS = [
  Math.round(CONTENT_WIDTH_DXA * 0.50),  // 4819
  Math.round(CONTENT_WIDTH_DXA * 0.28),  // 2699
  CONTENT_WIDTH_DXA - Math.round(CONTENT_WIDTH_DXA * 0.50) - Math.round(CONTENT_WIDTH_DXA * 0.28), // remainder
];

/**
 * Build a header row cell for a materials or mix table.
 * @param {string} text       - Cell text.
 * @param {number} widthDxa   - Cell width in DXA.
 * @returns {TableCell}
 */
function headerCell(text, widthDxa) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: COLOR.CORNFLOWER };
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { fill: COLOR.LIGHT_GREY, type: ShadingType.CLEAR },
    borders: { top: border, bottom: border, left: border, right: border },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text,
            font: FONT.HEADING,
            size: SIZE.BODY,
            bold: true,
            color: COLOR.NAVY,
          }),
        ],
      }),
    ],
  });
}

/**
 * Build a data row cell for a materials or mix table.
 * @param {string}  text     - Cell text (plain, no inline markers in table cells).
 * @param {number}  widthDxa - Cell width in DXA.
 * @param {boolean} bold     - If true, render text bold.
 * @returns {TableCell}
 */
function dataCell(text, widthDxa, bold = false) {
  const border = { style: BorderStyle.SINGLE, size: 2, color: COLOR.MID_GREY };
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    borders: { top: border, bottom: border, left: border, right: border },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: text || "\u2014", // em-dash for empty cells
            font: FONT.BODY,
            size: SIZE.BODY,
            bold,
          }),
        ],
      }),
    ],
  });
}

/**
 * Render a MaterialsTable object into a docx Table.
 *
 * @param {object} materialsTable - MaterialsTable from the schema.
 * @returns {Table}
 */
function renderMaterialsTable(materialsTable) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell("Reagent / Equipment", MAT_COL_WIDTHS[0]),
      headerCell("Supplier", MAT_COL_WIDTHS[1]),
      headerCell("Cat. No. / Notes", MAT_COL_WIDTHS[2]),
    ],
  });

  const dataRows = materialsTable.rows.map(
    (row) =>
      new TableRow({
        children: [
          dataCell(row.cells[0], MAT_COL_WIDTHS[0], row.bold),
          dataCell(row.cells[1] || "\u2014", MAT_COL_WIDTHS[1], row.bold),
          dataCell(row.cells[2] || "\u2014", MAT_COL_WIDTHS[2], row.bold),
        ],
      })
  );

  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: MAT_COL_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ---------------------------------------------------------------------------
// Mix / recipe table (2-column: Component / Amount)
// ---------------------------------------------------------------------------

/**
 * Column widths for a 2-column mix table.
 * Component: ~65%, Amount: ~35%
 */
const MIX_COL_WIDTHS = [
  Math.round(CONTENT_WIDTH_DXA * 0.65),
  CONTENT_WIDTH_DXA - Math.round(CONTENT_WIDTH_DXA * 0.65),
];

/**
 * Render a MixTable object into a docx Table.
 * "Total" and "Incubate" rows are automatically bold, in addition to any
 * rows with the explicit bold flag set on the schema object.
 *
 * @param {object} mixTable - MixTable from the schema.
 * @returns {Table}
 */
function renderMixTable(mixTable) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell("Component", MIX_COL_WIDTHS[0]),
      headerCell("Amount", MIX_COL_WIDTHS[1]),
    ],
  });

  const dataRows = mixTable.rows.map((row) => {
    const firstCell = (row.cells[0] || "").trim();
    const autoBold =
      row.bold ||
      /^total/i.test(firstCell) ||
      /^incubate/i.test(firstCell);

    return new TableRow({
      children: [
        dataCell(row.cells[0], MIX_COL_WIDTHS[0], autoBold),
        dataCell(row.cells[1] || "\u2014", MIX_COL_WIDTHS[1], autoBold),
      ],
    });
  });

  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: MIX_COL_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ---------------------------------------------------------------------------
// Section rendering helpers
// ---------------------------------------------------------------------------

/**
 * Render the Overview section.
 * DEV-004: Splits on \n\n to produce separate Paragraph objects, since
 * the docx npm package ignores \n inside TextRun.
 *
 * @param {string} text - Overview text with optional inline markers.
 * @returns {Array<Paragraph>}
 */
function renderOverview(text) {
  const chunks = (text || "").split(/\n\n+/).filter(c => c.trim());
  if (chunks.length === 0) {
    chunks.push(text || "");
  }
  return [
    ...heading1("Overview"),
    ...chunks.map(chunk =>
      new Paragraph({
        children: parseInline(chunk.trim(), { font: FONT.BODY, size: SIZE.BODY }),
        spacing: { after: 120 },
      })
    ),
  ];
}

/**
 * Render the Materials section (wet-lab only).
 * @param {object[]} materialsTables - Array of MaterialsTable schema objects.
 * @returns {Array<Paragraph|Table>}
 */
function renderMaterials(materialsTables) {
  if (!materialsTables || materialsTables.length === 0) return [];

  const elements = [...heading1("Materials")];
  for (const mt of materialsTables) {
    elements.push(heading3(mt.heading));
    elements.push(renderMaterialsTable(mt));
    elements.push(spacer(100));
  }
  return elements;
}

/**
 * Render the full Procedure section.
 * Callouts associated with a ProcedureSection are placed at the top of that
 * section. If the source document implies a specific position for a callout,
 * the extractor should interleave callouts within the steps array as
 * separate callout-bearing wrapper objects; otherwise top-of-section
 * placement is the default.
 *
 * @param {object[]} procedureSections - Array of ProcedureSection objects.
 * @param {object}   [fontOpts]        - Font overrides for step text.
 * @returns {Array<Paragraph|Table>}
 */
function renderProcedure(procedureSections, fontOpts = {}) {
  const elements = [...heading1("Procedure")];

  for (const section of procedureSections) {
    elements.push(heading2(section.heading));

    // Preamble (used mainly in computational protocols but accepted here too)
    if (section.preamble) {
      const opts = { font: FONT.BODY, size: SIZE.BODY, ...fontOpts };
      elements.push(
        new Paragraph({
          children: parseInline(section.preamble, opts),
          spacing: { after: 80 },
        })
      );
    }

    // Callouts at top of section
    for (const callout of section.callouts || []) {
      elements.push(calloutBox(callout, fontOpts));
      elements.push(spacer(60));
    }

    // Steps
    elements.push(...renderSteps(section.steps || [], 0, fontOpts));
    elements.push(spacer(80));
  }

  return elements;
}

/**
 * Render the Notes & Variants section.
 * @param {string[]} notes - Array of note strings.
 * @returns {Array<Paragraph>}
 */
function renderNotes(notes) {
  if (!notes || notes.length === 0) return [];

  const elements = [...heading1("Notes & Variants")];
  notes.forEach((note, i) => {
    elements.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `${i + 1}.\u00A0\u00A0`,
            font: FONT.BODY,
            size: SIZE.BODY,
            bold: true,
          }),
          ...parseInline(note, { font: FONT.BODY, size: SIZE.BODY }),
        ],
        spacing: { after: 80 },
        indent: { left: 360, hanging: 360 },
      })
    );
  });
  return elements;
}

/**
 * Render the References section.
 * @param {string[]} references - Array of reference strings.
 * @returns {Array<Paragraph>}
 */
function renderReferences(references) {
  if (!references || references.length === 0) return [];

  const elements = [...heading1("References")];
  for (const ref of references) {
    elements.push(
      new Paragraph({
        children: parseInline(ref, { font: FONT.BODY, size: SIZE.BODY }),
        spacing: { after: 80 },
        indent: { left: 360, hanging: 360 },
      })
    );
  }
  return elements;
}

/**
 * Render the mix tables section (post-References, pre end-of-protocol).
 * DEVNOTE002: Placement is after References.
 *
 * @param {object[]} mixTables - Array of MixTable schema objects.
 * @returns {Array<Paragraph|Table>}
 */
function renderMixTables(mixTables) {
  if (!mixTables || mixTables.length === 0) return [];

  const elements = [...heading1("Reagent Mixes & Recipes")];
  for (const mt of mixTables) {
    elements.push(heading3(mt.heading));
    elements.push(renderMixTable(mt));
    elements.push(spacer(100));
  }
  return elements;
}

/**
 * Render the end-of-protocol line.
 * @returns {Paragraph}
 */
function endOfProtocol() {
  return new Paragraph({
    children: [
      new TextRun({
        text: "\u2500\u2500\u2500  End of Protocol  \u2500\u2500\u2500",
        font: FONT.BODY,
        size: SIZE.BODY,
        italics: true,
        color: COLOR.MID_GREY,
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200 },
  });
}

// ---------------------------------------------------------------------------
// Title block
// ---------------------------------------------------------------------------

/**
 * Render the protocol title block:
 *   - Title (large, bold, navy Calibri)
 *   - Subtitle (italic Georgia, if present)
 *   - Byline: "Author · Affiliation · Version X.X · Last updated DD/MM/YYYY"
 *
 * @param {object} protocol - Top-level Protocol schema object.
 * @returns {Paragraph[]}
 */
function renderTitleBlock(protocol) {
  const elements = [];

  // Title
  elements.push(
    new Paragraph({
      children: [
        new TextRun({
          text: protocol.title,
          font: FONT.HEADING,
          size: 40,          // 20pt for title
          bold: true,
          color: COLOR.NAVY,
        }),
      ],
      spacing: { before: 120, after: 60 },
    })
  );

  // Subtitle
  if (protocol.subtitle) {
    elements.push(
      new Paragraph({
        children: [
          new TextRun({
            text: protocol.subtitle,
            font: FONT.BODY,
            size: SIZE.BODY,
            italics: true,
            color: COLOR.BLACK,
          }),
        ],
        spacing: { after: 80 },
      })
    );
  }

  // Byline
  const bylineParts = [protocol.author || "ARMI"];
  bylineParts.push(`Version\u00A0${protocol.version || "1.0"}`);
  if (protocol.date) bylineParts.push(`Last updated\u00A0${protocol.date}`);
  const bylineText = bylineParts.join("\u00A0\u00A0\u00B7\u00A0\u00A0");

  elements.push(
    new Paragraph({
      children: [
        new TextRun({
          text: bylineText,
          font: FONT.BODY,
          size: SIZE.BODY,
          color: COLOR.MID_GREY,
        }),
      ],
      spacing: { after: 160 },
      border: {
        bottom: {
          style: BorderStyle.SINGLE,
          size: 4,
          color: COLOR.LIGHT_GREY,
          space: 4,
        },
      },
    })
  );

  elements.push(spacer(80));
  return elements;
}

// ---------------------------------------------------------------------------
// Main wet-lab document builder
// ---------------------------------------------------------------------------

/**
 * Build a complete wet-lab protocol Document from a validated Protocol object.
 *
 * @param {object} protocol - Deserialised Protocol schema object.
 * @returns {Document}      - docx Document ready for Packer.toBuffer().
 */
function buildWetLabDocument(protocol) {
  // Abbreviated title for footer (truncate at 40 chars)
  const abbrevTitle =
    protocol.title.length > 40
      ? protocol.title.slice(0, 37) + "\u2026"
      : protocol.title;

  // Assemble body content
  const body = [
    ...renderTitleBlock(protocol),
    ...renderOverview(protocol.overview),
    ...renderMaterials(protocol.materials),
    ...renderProcedure(protocol.procedure),
    ...renderNotes(protocol.notes),
    ...renderReferences(protocol.references),
    ...renderMixTables(protocol.mix_tables),
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
          run: { font: FONT.BODY, size: SIZE.BODY, color: COLOR.BLACK },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE_WIDTH_DXA, height: PAGE_HEIGHT_DXA },
            margin: {
              top: MARGIN_DXA,
              bottom: MARGIN_DXA,
              left: MARGIN_DXA,
              right: MARGIN_DXA,
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
// Exports
// ---------------------------------------------------------------------------

module.exports = {
  // Constants (re-exported for lib_comp.js)
  PAGE_WIDTH_DXA,
  PAGE_HEIGHT_DXA,
  MARGIN_DXA,
  CONTENT_WIDTH_DXA,
  SIZE,
  FONT,
  COLOR,
  CALLOUT_META,
  STOP_SYMBOL,
  NUMBERING_REFS,

  // Shared utility functions (re-exported for lib_comp.js)
  parseInline,
  spacer,
  buildHeader,
  buildFooter,
  buildNumberingConfig,
  heading1,
  heading2,
  heading3,
  calloutBox,
  stoppingPoint,
  renderSteps,
  renderOverview,
  renderNotes,
  renderReferences,
  endOfProtocol,
  renderTitleBlock,
  headerCell,
  dataCell,

  // Wet-lab specific
  renderMaterials,
  renderProcedure,
  renderMixTables,
  buildWetLabDocument,
};
