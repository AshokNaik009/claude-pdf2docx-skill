# Docling setup & verification (`--docling` engine)

The `--docling` engine parses PDFs with [Docling](https://github.com/docling-project/docling)
(IBM's Python document parser) instead of the vision-LLM path. It needs a Python
virtual environment with the `docling` package installed. This is a **one-time**
setup; the other engines (default / `--pandoc` / `--markdown`) do **not** need it.

> The project already expects the venv at `.venv/` — see `config.js`:
> `pythonBin: ".venv/bin/python"`, `doclingScript: "scripts/docling_extract.py"`.

---

## Requirements

- **Python 3.12** (recommended — the project venv was built with it).
  ```bash
  python3.12 --version        # or: /opt/homebrew/bin/python3.12 --version
  ```
  If you don't have it: `brew install python@3.12`.
- **~1.9 GB of disk**: ~1.3 GB for the venv (pulls `torch`, `transformers`,
  `docling-ibm-models`, `rapidocr`) + ~0.5 GB of model weights downloaded on
  first run into `~/.cache/huggingface/`.
- Network access for the first install and the first parse (weight download).

---

## Install (pip)

From the project root (`pdf-to-docx/`):

```bash
# 1. Create the virtual environment (must be named .venv to match config.js)
python3.12 -m venv .venv          # or: /opt/homebrew/bin/python3.12 -m venv .venv

# 2. Upgrade pip, then install docling into the venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install docling
```

Notes
- The install is large (torch etc.) and can take a few minutes — this is normal.
- Do **not** `source .venv/bin/activate`; the project always calls the venv
  Python by its explicit path (`.venv/bin/python`), so activation is optional.
- Pin a version if you want reproducibility: `.venv/bin/pip install "docling==2.113.0"`.

### Known-good versions (as installed in this project)
| Package | Version |
|---|---|
| docling | 2.113.0 |
| docling-core | 2.87.1 |
| docling-ibm-models | 3.13.3 |
| torch | 2.13.0 |
| transformers | 5.8.1 |
| rapidocr | 3.9.1 |

---

## Verify

### 1. The exact check the app runs
`index.js` gates `--docling` on this (`doclingAvailable()` in
`src/doclingengine.js`) — it must exit 0:

```bash
.venv/bin/python -c "import docling" && echo "import docling OK"
```

### 2. Confirm the version
```bash
.venv/bin/pip show docling | grep -E "^(Name|Version)"
```

### 3. End-to-end parse on a sample PDF
The **first** run also downloads the layout/table model weights (one-time,
~0.5 GB → `~/.cache/huggingface/`), so it's slower; later runs are ~seconds.

```bash
.venv/bin/python scripts/docling_extract.py newtest-document.pdf work/verify-docling
```

Expected: it prints a summary and writes into `work/verify-docling/`:

| File | Contents |
|---|---|
| `doc.json` | Full DoclingDocument — layout, bboxes, element types, table grids (**this is what the code generator consumes**). |
| `doc.md` | `export_to_markdown()` — reading-order text + tables. |
| `doc.html` | `export_to_html()` — richer structure. |
| `images/` | Extracted picture assets (`pic_N.png`). |
| `SUMMARY.txt` | Element-type counts + per-page table/figure tallies. |

A healthy summary looks like:
```
pdf: newtest-document.pdf
convert_seconds: 4.x
pages: 7
tables: 1
pictures: 1  (saved 1)
texts: 67
element_type_counts:
  text: 50
  section_header: 12
  table: 1
  caption: 1
  picture: 1
```

> OCR is intentionally **disabled** (`do_ocr=False` in `scripts/docling_extract.py`)
> because these PDFs have a real text layer — it's ~9× faster with identical
> quality. If you feed a **scanned** PDF, re-enable OCR in that script.

### 4. Full pipeline through the flag
The real integration — parse **and** build the `.docx`:

```bash
node index.js all newtest-document.pdf --docling --fresh
# → outputs/newtest-document.docx
```

If Docling isn't installed, this fails fast with a clear message:
```
--docling requires the 'docling' Python package in the project venv.
Set it up:  python3 -m venv .venv && .venv/bin/pip install docling
```

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `--docling requires the 'docling' Python package…` | The venv is missing or `import docling` fails. Re-run the install; confirm the path is exactly `.venv/bin/python`. |
| First parse hangs / very slow | It's downloading model weights (~0.5 GB) on the first `convert()`. Let it finish once; subsequent runs are fast. Needs network. |
| `EPERM` / read fails on a PDF under `~/Downloads` | macOS TCC blocks CLI reads there (`existsSync` passes but the read fails). Move the PDF into the project directory. |
| Wrong Python picked up | Always invoke `.venv/bin/python …` explicitly (never rely on `activate` or a global `python`). |
| Want a clean reinstall | `rm -rf .venv && python3.12 -m venv .venv && .venv/bin/pip install docling`. To also clear weights: `rm -rf ~/.cache/huggingface/hub/models--docling-project--*`. |

---

## Where this plugs into the code

```
PDF ─▶ .venv/bin/python scripts/docling_extract.py ─▶ doc.json (+ md/html/images)
    ─▶ src/doclingcode.js  generateDoclingCode()     ─▶ work/<name>/pages/*.mjs (docx.js modules)
    ─▶ node index.js merge                            ─▶ outputs/<name>.docx
```

- `config.js` → `pythonBin`, `doclingScript`, `docling` flag defaults.
- `src/doclingengine.js` → spawns the venv Python, then calls the generator.
- `src/doclingcode.js` → turns `doc.json` into editable `docx.js` page modules.

See `README.md` for the full engine comparison and the Docling deep-dive.
