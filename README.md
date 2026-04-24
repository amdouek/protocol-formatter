# ProtocolFormatter (protocol-formatter)
A local CLI tool that converts raw laboratory protocol documents (.docx, .doc) into consistently formatted Word files conforming to my Protocol Compendium template.

All processing runs entirely on-device; LLM inference is (currently) handled by [Ollama](https://ollama.com/) running `qwen2.5:14b` locally. No data is sent to external APIs.

## Background
As a predominantly wet-lab biologist, I've accumulated *many* bench protocols over the years from a variety of sources. Some are more detailed than others; some are neatly formatted, while others are a remarkable embodiment of chaos. After I finished my PhD, I wanted to do a bit of spring cleaning and assemble these protocols into a uniform, aesthetically pleasing format placed in an ordered protocol compendium.

However, graphic design is *not* my passion. After manually formatting a single protocol according to a consistent schema, I got bored and wanted a way to automate (or semi-automate) the process of converting the source protocol into my defined format. This tool is the consequence of that boredom.

Two document templates are supported:
- **Wet-lab:** Georgia body text, Calibri headings, navy/cornflower colour scheme, materials tables, mix/recipe tables, four callout types (Critical [Red], Caution [Orange], Tip [Green], Note [Blue]), and in-line stopping points.
- **Computational:** Calibri throughout, Courier New code blocks, prerequisites section, same callout and colour convention as above.

## Architecture
```
Raw .docx / .doc
      │
      ▼
 Input parser          python-docx (.docx) or pandoc subprocess (.doc)
      │                Extracts paragraphs, tables, footnotes, bold/italic runs
      ▼
 LLM extraction        Ollama (qwen2.5:14b) via local REST API
      │                Single-pass; retries up to 3× with targeted correction prompts
      │                Falls back to deterministic heuristic extractor on failure
      ▼
 Human review          Optional --review flag: inspect and edit the extracted
      │                schema as JSON before rendering (STRONGLY recommended)
      ▼
 Node.js renderer      Python shells out to render.js with a JSON payload
      │                render.js selects lib.js (wet-lab) or lib_comp.js (computational)
      ▼
 Output .docx
 ```

 The intermediate representation is a [Pydantic V2](https://pypi.org/project/pydantic/2.0/) `Protocol` model. This is a template-agnostic schema that describes content semantics instead of presentation. All visual decisions are instead handled by the renderer.

 ### Project structure
 ```
 protocol_formatter/
├── main.py                        # CLI entry point (Typer)
├── schema.py                      # Pydantic v2 intermediate models
├── pyproject.toml
├── configs/
│   └── style_guide.yaml           # Colours, fonts, callout keywords, paths
├── parser/
│   ├── docx_reader.py             # python-docx extraction
│   ├── doc_reader.py              # pandoc fallback for legacy .doc files
│   └── utils.py                   # Normalisation helpers
├── extractor/
│   ├── llm_extractor.py           # Ollama REST call, validation, retry logic
│   ├── prompts.py                 # Prompt templates and style guide context
│   └── rule_extractor.py          # Heuristic fallback extractor
├── renderer/
│   ├── node_renderer.py           # Python → Node.js shim
│   └── templates/
│       ├── lib.js                 # Wet-lab document builder
│       ├── lib_comp.js            # Computational document builder
│       └── render.js              # Entry point; accepts JSON payload
└── tests/
    └── test_roundtrip.py
```

## Requirements
| Dependency | Purpose |
|---|---|
| Python ≥ 3.10 | Runtime |
| Node.js (any recent LTS) | Document rendering via `docx` npm package |
| [Ollama](https://ollama.com) | Local LLM inference |
| `qwen2.5:14b` | Extraction model (pulled via Ollama) |
| pandoc *(optional)* | Legacy `.doc` file support |

## Installation
**1. Clone the repo**
```bash
git clone https://github.com/amdouek/protocol_formatter.git
cd protocol_formatter
```

**2. Install the Node.js rendering dependency**
... and then go back to project root.
```bash
cd renderer/templates
npm install docx
cd ../..
```

**3. Create and activate a Python .venv and install the ProtocolFormatter package**
```bash
python -m venv .venv

# If using Windows
.venv\Scripts\activate

# If using macOS or Linux
source .venv/bin/activate

pip install -e ".[dev]"
```

**4. Install Ollama and pull the LLM**
Download Ollama from [ollama.com](ollama.com), then:
```bash
ollama pull qwen2.5:14b     # Needs 9 GB space on device
```

**5. (Optional) Install pandoc for `.doc` support**
- Windows: `winget install --source winget --exact --id JohnMacFarlane.Pandoc`
- macOS: `brew install pandoc`
- Linux: See here: [Pandoc Linux Installation](https://pandoc.org/installing.html#linux)

**6. Run dependency check**
```bash
protocol-formatter check
```
All items (except the last item, if you haven't installed pandoc) should show the following:

```
╭──────────────────────────────────────╮
│ ProtocolFormatter — Dependency Check │
╰──────────────────────────────────────╯
  ✓  Node.js: v24.15.0
  ✓  docx npm package: 9.6.1
  ✓  render.js: C:\..\protocol_formatter\renderer\templates\render.js
  ✓  Ollama server: Available models: qwen2.5:14b
  ✓  Ollama model (qwen2.5:14b): available
  ✗  pandoc (.doc support): pandoc executable not found: 'pandoc'. Install pandoc from https://pandoc.org/installing.html and ensure it is on your PATH, 
or set paths.pandoc_executable in configs/style_guide.yaml. (optional)

Some checks failed. Fix the issues above before running protocol-formatter.
```

## Usage
**Format a single protocol**
```bash
protocol-formatter format path/to/protocol.docx --section 4 --review
```
`--section 4` gives the LLM a hint as to which section of the protocol compendium the protocol should belong to. Section labels are defined in `configs/style_guide.yaml`; here, `4` refers to *"TRANSGENESIS AND GENOME EDITING"*.

**Format a single protocol without the LLM using heuristic extraction**

This is mainly for a quick preview or for if Ollama is down. The output will be structurally correct but will need more manual checking.
```bash
protocol-formatter format path/to/protocol.docx --heuristic --review
```

**Batch processing**

To process a bunch of protocols in a single directory at once (you should have already grouped them by their destination section)
```bash
protocol-formatter batch path/to/protocols --section 4 --review
```

**Print the intermediate JSON schema**
```bash
protocol-formatter schema
```

### CLI flags
 
| Flag | Use with | Effect |
|---|---|---|
| `--section` | `format`, `batch` | Numerical; injects section number and name as a hint to the LLM |
| `--review` | `format`, `batch` | Pauses for human review of the extracted schema before rendering |
| `--heuristic` | `format`, `batch` | Skips Ollama; uses deterministic rule-based extraction only |
| `--output path/` | `format`, `batch` | Writes output to a custom directory (default: `output/`) |
| `--glob "*.docx"` | `batch` | Override the file pattern for batch discovery |

### Configuration
All tuneable params live in `configs/style_guide.yaml`. The following are the key settings:

```yaml
paths:
  node_executable: "node"          # Full path if Node is not on PATH
  pandoc_executable: "pandoc"      # Full path if pandoc is not on PATH
  output_dir: "output"             # Default output directory
 
ollama:
  base_url: "http://localhost:11434"
  model: "qwen2.5:14b"
  max_tokens: 4096
  max_retries: 3
  request_timeout_seconds: 300    # Increase for long protocols and/or if you get a timeout failure
```

Colour values (defined in hex), font sizes, callout keywords, and the compendium section registry are also defined here, and are propagated to both the prompt and the renderer at runtime.

## Compendium section registry
| Number | Section name |
|---|---|
| 1 | NUCLEIC ACID EXTRACTION & PURIFICATION |
| 2 | cDNA SYNTHESIS & GENE EXPRESSION |
| 3 | CLONING & PLASMID ENGINEERING |
| 4 | TRANSGENESIS AND GENOME EDITING |
| 5 | RNA PROBES & IN SITU HYBRIDISATION |
| 6 | HISTOLOGY, IMMUNOFLUORESCENCE & IMAGING |
| 7 | CELL & TISSUE PREPARATION |
| 8 | ANIMAL TECHNIQUES & REPRODUCTIVE BIOLOGY |
| 9 | COMPUTATIONAL METHODS & BIOINFORMATICS |

## Running the tests
```bash
pytest tests/test_roundtrip.py -v
```

Note: Tests that require Ollama or Node.js are automatically skipped if those dependencies aren't available. The full suite of tests requires Ollama running with `qwen2.5:14b` pulled.

## Notes
- `renderer/templates/lib.js` defines the byline as `const bylineParts = [protocol.author || "ARMI"];` as I made this purely for personal use while working at ARMI. This parameter will change as my institutional affiliation does, but I'm flagging it here for others who wish to fork this tool to change to match their affiliation. 
    - This also needs to be changed in `prompts.py`, in the `# Author rule` section of `build_style_guide_context`, in `utils.py` (`extract_author`), as well as in relevant heuristics in `rule_extractor.py`.
- Output files are always written to a new path (by default, a subdirectory called 'output' in the package root), and source docs are never modified.
- If duplicate filenames are detected in a batch run, these are skipped automatically to prevent processing the same doc more than once.
- If LLM extraction fails validation after all retries (default 4), the tool falls back to the heuristic extractor and flags the result for review.
- Use of the `--review` flag is **STRONGLY** recommended, especially for initial runs so you can judge how well the tool is handling your protocols. The exception to this is if you really don't care about the fidelity of the content, and you just want to leverage the stylistic conversion aspect of this tool.
- If you're trying to process longer protocol documents, you will likely need to increase `request_timeout_seconds` in `configs/style_guide.yaml` from the default 300s depending on your machine's GPU VRAM availability (this tool was built and validated on a laptop with 8 GB dedicated GPU memory and ~18 GB shared GPU memory).
- This tool was built specifically to use Qwen2.5's 14B model. Others may work, but backend parameters and other architectural elements will need tweaking. This is not a priority for me, but I may come back to it at some point.

## Acknowledgements
- Rendering uses the [`docx`](https://www.npmjs.com/package/docx) npm package (GitHub: https://github.com/dolanmiu/docx) under the MIT license
- Inference uses [Ollama](https://ollama.com) (GitHub: https://github.com/ollama/ollama) under the MIT license