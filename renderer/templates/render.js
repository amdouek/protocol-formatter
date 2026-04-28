/**
 * render.js -- Node.js renderer entry point for ProtocolFormatter.
 *
 * Reads a JSON-serialised Protocol object from stdin, validates the required
 * fields, selects the appropriate template library (lib.js for wet_lab,
 * lib_comp.js for computational), builds the Document, and writes the output
 * .docx file to the path specified in the payload.
 *
 * Invocation (from Python node_renderer.py)
 * ------------------------------------------
 *   echo '<json>' | node render.js
 *   node render.js --input payload.json   (alternative file-based mode)
 *
 * Exit codes
 * ----------
 *   0  — success; output path written to stdout as JSON: {"output": "..."}
 *   1  — validation or rendering error; details written to stderr as JSON:
 *         {"error": "...", "detail": "..."}
 *
 * Stdin / file input
 * ------------------
 * The payload is a UTF-8 encoded JSON string matching the Protocol Pydantic
 * schema. The top-level field "output_path" specifies where the .docx file
 * should be written. If "output_path" is absent, the renderer writes to
 * stdout as raw binary (useful for piped workflows).
 *
 * Required top-level fields
 * -------------------------
 *   title           string
 *   section_type    "wet_lab" | "computational"
 *   section_number  string
 *   section_name    string
 *   overview        string
 *   procedure       array (min 1 element)
 *
 * Optional top-level fields
 * -------------------------
 *   subtitle, author, version, date,
 *   materials, prerequisites, notes, references,
 *   mix_tables, source_filename, output_path
 */

"use strict";

const fs   = require("fs");
const path = require("path");

const { Packer } = require("docx");
const { buildWetLabDocument }        = require("./lib.js");
const { buildComputationalDocument } = require("./lib_comp.js");

// ---------------------------------------------------------------------------
// Error helpers
// ---------------------------------------------------------------------------

/**
 * Write a structured error to stderr and exit with code 1.
 * @param {string} message  - Short error summary.
 * @param {string} [detail] - Optional longer explanation or stack trace.
 */
function fatal(message, detail = "") {
  process.stderr.write(
    JSON.stringify({ error: message, detail: String(detail) }) + "\n"
  );
  process.exit(1);
}

// ---------------------------------------------------------------------------
// Payload validation
// ---------------------------------------------------------------------------

const REQUIRED_FIELDS = [
  "title",
  "section_type",
  "section_number",
  "section_name",
  "overview",
  "procedure",
];

const VALID_SECTION_TYPES = new Set(["wet_lab", "computational"]);

/**
 * Validate the minimum required fields on the parsed payload.
 * Throws a TypeError with a descriptive message on failure.
 *
 * @param {object} payload - Parsed JSON object.
 */
function validatePayload(payload) {
  for (const field of REQUIRED_FIELDS) {
    if (payload[field] === undefined || payload[field] === null) {
      throw new TypeError(`Missing required field: "${field}"`);
    }
  }

  if (!VALID_SECTION_TYPES.has(payload.section_type)) {
    throw new TypeError(
      `Invalid section_type "${payload.section_type}". ` +
      `Expected "wet_lab" or "computational".`
    );
  }

  if (!Array.isArray(payload.procedure) || payload.procedure.length === 0) {
    throw new TypeError(
      `"procedure" must be a non-empty array. ` +
      `Got: ${JSON.stringify(payload.procedure)}`
    );
  }

  // Cross-field consistency checks (mirrors Pydantic validator)
  if (payload.section_type === "wet_lab" && payload.prerequisites != null) {
    throw new TypeError(
      `"prerequisites" must be null/absent for wet_lab protocols.`
    );
  }
  if (
    payload.section_type === "computational" &&
    Array.isArray(payload.materials) &&
    payload.materials.length > 0
  ) {
    throw new TypeError(
      `"materials" must be empty for computational protocols.`
    );
  }
  if (
    payload.section_type === "computational" &&
    Array.isArray(payload.mix_tables) &&
    payload.mix_tables.length > 0
  ) {
    throw new TypeError(
      `"mix_tables" must be empty for computational protocols.`
    );
  }
}

// ---------------------------------------------------------------------------
// Input reading
// ---------------------------------------------------------------------------

/**
 * Read the full JSON payload. Supports two modes:
 *   1. --input <filepath>  reads from a file
 *   2. Default             reads from stdin
 *
 * Returns a Promise that resolves to the raw JSON string.
 *
 * @returns {Promise<string>}
 */
function readPayload() {
  return new Promise((resolve, reject) => {
    // Check for --input flag
    const inputFlagIdx = process.argv.indexOf("--input");
    if (inputFlagIdx !== -1) {
      const filePath = process.argv[inputFlagIdx + 1];
      if (!filePath) {
        return reject(new Error("--input flag requires a file path argument."));
      }
      try {
        const content = fs.readFileSync(filePath, "utf8");
        resolve(content);
      } catch (err) {
        reject(new Error(`Cannot read input file "${filePath}": ${err.message}`));
      }
      return;
    }

    // Read from stdin
    let data = "";
    process.stdin.setEncoding("utf8");
    process.stdin.on("data", (chunk) => { data += chunk; });
    process.stdin.on("end", () => resolve(data));
    process.stdin.on("error", (err) =>
      reject(new Error(`stdin read error: ${err.message}`))
    );
  });
}

// ---------------------------------------------------------------------------
// Output writing
// ---------------------------------------------------------------------------

/**
 * Write the Document buffer to a file or stdout.
 *
 * @param {Document} doc        - Built docx Document object.
 * @param {string|null} outPath - Absolute or relative output file path,
 *                                or null to write raw bytes to stdout.
 * @returns {Promise<string>}   - Resolves to the output path used.
 */
async function writeOutput(doc, outPath) {
  const buffer = await Packer.toBuffer(doc);

  if (!outPath) {
    // Binary stdout mode — useful for piped workflows
    process.stdout.write(buffer);
    return "(stdout)";
  }

  // Ensure the output directory exists
  const dir = path.dirname(path.resolve(outPath));
  fs.mkdirSync(dir, { recursive: true });

  fs.writeFileSync(outPath, buffer);
  return path.resolve(outPath);
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  let rawJson;
  try {
    rawJson = await readPayload();
  } catch (err) {
    fatal("Failed to read input payload", err.message);
  }

  if (!rawJson || !rawJson.trim()) {
    fatal("Empty payload received. Provide JSON via stdin or --input <file>.");
  }

  // Parse JSON
  let payload;
  try {
    payload = JSON.parse(rawJson);
  } catch (err) {
    fatal("JSON parse error", err.message);
  }

  // Validate
  try {
    validatePayload(payload);
  } catch (err) {
    fatal("Payload validation error", err.message);
  }

  // Build document
  let doc;
  try {
    if (payload.section_type === "wet_lab") {
      doc = buildWetLabDocument(payload);
    } else {
      doc = buildComputationalDocument(payload);
    }
  } catch (err) {
    fatal("Document build error", err.stack || err.message);
  }

  // Write output
  const outPath = payload.output_path || null;
  let resolvedPath;
  try {
    resolvedPath = await writeOutput(doc, outPath);
  } catch (err) {
    fatal("Output write error", err.message);
  }

  // Report success.
  // In binary stdout mode (no output_path), the docx buffer has already been
  // written to stdout. Route the success metadata to stderr to avoid
  // corrupting the binary stream.
  const successJson = JSON.stringify({ output: resolvedPath }) + "\n";
  if (resolvedPath === "(stdout)") {
    process.stderr.write(successJson);
  } else {
    process.stdout.write(successJson);
  }
  process.exit(0);
}

main().catch((err) => {
  fatal("Unexpected renderer error", err.stack || err.message);
});
