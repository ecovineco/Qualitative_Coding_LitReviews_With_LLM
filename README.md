# Literature Review Pipeline

---

## Part 1 — Introduction

This project is an automated pipeline that uses a Large Language Model (LLM) to code scientific and grey literature against a structured coding framework. It is built for researchers doing systematic or scoping literature reviews who want to extract verbatim, labelled evidence from a corpus of PDFs without hand-coding each one.

The pipeline supports multiple LLM providers — currently **Anthropic Claude** and **Azure OpenAI** — selected via a single config setting. Additional providers can be added with no changes outside `api.py`.

Inputs:

- A folder of PDF documents (Word and PowerPoint files are auto-converted on Windows).
- A coding framework: an Excel file listing labels, each with a description, inclusion criteria, and exclusion criteria.

Output:

- `coded_findings.xlsx` — the clean deliverable. One row per coded snippet, tagged with a label category, label code, page number, the LLM's reasoning, a self-assessed confidence level, a unique snippet identifier, a timestamp, and a stable hash for deduplication. Snippets that could not be located in their source PDF (status `not_found`) are excluded from this file, and it carries no verification columns.
- `coded_findings_verified.xlsx` — the full audit file. Every finding (including the `not_found` ones excluded above), each annotated with its verification verdict: whether the snippet was located verbatim in its source PDF, how, the match score, and the page it was actually found on. Verification runs automatically at the end of every pipeline run and costs no API credits.

The distinctive design choice is that the coding framework is **organised by categories**, and the pipeline runs **one batch per category**. For *N* documents and *C* categories, the pipeline submits *C* batches of *N* requests each — every request pairing one document with a prompt that contains only that category's labels. All results are merged into a single output file at the end. The rationale for this is given in Part 4.

---

## Part 2 — How to Use the Tool

### Step 1 — Install dependencies

```bash
pip install -r requirements.txt
```

Required packages: `requests`, `pandas`, `openpyxl`, `openai`, `PyMuPDF`, `rapidfuzz`.

`PyMuPDF` and `rapidfuzz` power the snippet-verification stage (Part 4.10): `PyMuPDF` extracts each PDF's own text and `rapidfuzz` matches snippets against it.

For Word or PowerPoint input files, also install `docx2pdf` and `pywin32`. This is **Windows + Microsoft Office only**, as the pipeline uses COM automation to convert these files to PDF.

### Step 2 — Set your API key

The key is read from an environment variable and must never be hard-coded in `config.py`.

For **Anthropic**:

```bash
# Linux / macOS
export ANTHROPIC_API_KEY="sk-ant-..."

# Windows PowerShell
$env:ANTHROPIC_API_KEY = "sk-ant-..."
```

For **Azure OpenAI**:

```bash
# Linux / macOS
export AZURE_OPENAI_API_KEY="..."

# Windows PowerShell
$env:AZURE_OPENAI_API_KEY = "..."
```

### Step 3 — Prepare your coding framework

Open `labels/coding_framework.xlsx`. It must have these five columns:

| Column | Required | What it is |
|---|---|---|
| `category` | Yes | Broader grouping (e.g. `CPVR_EFFECTIVENESS`). Multiple labels share a category. |
| `code` | Yes | Identifier within the category (e.g. `OBJECTIVES`). The pair `(category, code)` must be unique. |
| `description` | Yes | What the label covers, written as an instruction to a human coder. |
| `inclusion_criteria` | Yes | When to apply this label ("Code a passage if…"). |
| `exclusion_criteria` | Recommended | When NOT to apply this label ("Do NOT code if…"). |

Notes:

- The same code can appear under two different categories — for instance, both `CPVR_EFFECTIVENESS` and `CPVR_LEGAL` could each have a `BARRIERS` code. Uniqueness is enforced per category, not globally.
- Descriptions must be **self-contained**: the LLM sees only what is in the Excel.
- Inclusion and exclusion criteria should be **specific**. Telling the LLM what not to code is as valuable as telling it what to code.
- **Start small**. Set `MAX_PDFS = 3` in `config.py` and run on a handful of documents before processing the full corpus.

### Step 4 — Place your PDFs

Put them in `pdfs/` (or change `TARGET_FOLDER` in `config.py`). Use informative filenames such as `Author (Year) Title.pdf` — the filename is preserved in the output. PDFs must be text-based; scanned image PDFs cannot be read and will appear in `errors.xlsx`.

### Step 5 — Review settings in `config.py`

The key settings to review:

```python
TARGET_FOLDER = "./pdfs"                         # Where your PDFs are
LABELS_EXCEL  = "./labels/coding_framework.xlsx" # Your coding framework
OUTPUT_DIR    = "./outputs"                      # Where results are saved

LLM_PROVIDER  = "anthropic"   # "anthropic" or "azure"

# Anthropic-specific
ANTHROPIC_MODEL = "claude-opus-4-6"  # or "claude-sonnet-4-20250514" for speed / cost

# Azure-specific
AZURE_OPENAI_ENDPOINT    = "https://...cognitiveservices.azure.com/"
AZURE_OPENAI_API_VERSION = "2025-01-01-preview"
AZURE_OPENAI_MODEL       = "gpt-5.4"   # Azure deployment name

MAX_PDFS = None    # Set to an integer for testing
```

Only the settings for the selected `LLM_PROVIDER` matter; the others are ignored.

### Step 6 — Run the pipeline

**Full run:**

```bash
python main.py
```

The pipeline will: discover PDFs → upload each one once → loop over categories → submit a batch per category → poll → parse → merge results → verify snippets against their source PDFs → write `coded_findings.xlsx` (clean) and `coded_findings_verified.xlsx` (full audit).

**Dry run** (verify prompts and request structure without spending credits):

Set `DRY_RUN = True` in `config.py`, then `python main.py`. The pipeline still uploads files (needed to build requests) but stops before any batch is submitted.

**Resume** (after a crash or interruption):

```bash
python main.py --resume
```

Resume mode reads `outputs/batch_metadata.json` and `outputs/uploaded_file_ids.xlsx`, skips re-uploading, and skips submitting any category whose batch was already sent. Any categories that were not yet submitted are submitted fresh. All category batches are then polled, fetched, and parsed as usual.

### Step 7 — Read the output

After a successful run, `outputs/` will contain `coded_findings.xlsx` (the clean deliverable), `coded_findings_verified.xlsx` (the full audit file with verification verdicts), `errors.xlsx`, one `prompt_used__{CATEGORY}.txt` per category, `batch_metadata.json`, `uploaded_file_ids.xlsx`, and `pipeline.log`. The contents and schema of each file are described in Part 3.

---

## Part 3 — Project Structure

### Project layout

```
lit_review/
├── config.py          # Central settings (paths, model, flags, provider)
├── labels.py          # Load + validate the coding framework
├── prompt.py          # Build the LLM prompt from a set of labels
├── api.py             # LLM provider interface (abstract + concrete clients)
├── parser.py          # Parse + validate LLM JSON responses → Finding rows
├── verify.py          # Verify each snippet exists verbatim in its source PDF
├── export.py          # Write Excel + JSON outputs
├── main.py            # Pipeline orchestrator + CLI entry point
├── test_verify.py     # Fixture tests for the verification stage
├── requirements.txt
├── labels/
│   └── coding_framework.xlsx
├── pdfs/
└── outputs/
```

Each module has a single concern.

### 3.1 — The Python files

**`config.py`** — Central configuration. Defines paths (`TARGET_FOLDER`, `LABELS_EXCEL`, `OUTPUT_DIR`), the active LLM provider (`LLM_PROVIDER`), per-provider settings (`ANTHROPIC_MODEL`, `ANTHROPIC_VERSION`, `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_API_VERSION`, `AZURE_OPENAI_MODEL`, …), batch behaviour (`POLL_INTERVAL_SECONDS`, `MAX_PDFS`, `REQUEST_TIMEOUT`), and pipeline flags (`DRY_RUN`, `INCLUDE_CONFIDENCE`, `UPLOAD_RETRIES`, `UPLOAD_RETRY_DELAY`). API keys are read from the environment.

**`labels.py`** — Loads and validates the coding framework. Exposes:
- A `Label` dataclass with `category`, `code`, `description`, `inclusion_criteria`, `exclusion_criteria`, and a `to_prompt_block()` method that formats a single label for the prompt.
- `load_labels()` reads the Excel, normalises column names, checks required columns are present and non-empty, verifies that `(category, code)` pairs are unique, and returns a list of `Label` objects.
- `group_labels_by_category()` groups labels into an `OrderedDict` that preserves first-appearance order.
- `get_valid_codes()` returns the set of valid codes for a list of labels — used by the parser to detect hallucinated codes.

**`prompt.py`** — Builds the LLM prompt from a set of labels. The main function is `build_analysis_prompt(labels, include_confidence)`, which assembles a structured prompt instructing the LLM to act as a systematic reviewer, listing each label's description/inclusion/exclusion criteria, specifying the exact JSON output format, and enumerating the coding rules (verbatim quotes only, label codes must come from the framework, etc.). A stub `build_summary_prompt()` is also defined for an optional cross-document synthesis pass — see Part 5.

**`api.py`** — LLM provider interface. Defines:
- The dataclasses `BatchStatus` (batch state) and `BatchResultItem` (one result from a completed batch).
- `_make_custom_id(filename, idx)` — the single source of truth for batch custom-IDs (see Part 4 for rationale).
- `BaseLLMClient`, an abstract base class with four abstract methods (`upload_file`, `submit_batch`, `check_batch`, `get_results`) plus a concrete `poll_until_complete()` helper.
- `AnthropicClient`, the concrete implementation for Claude using Anthropic's Files API (beta) and Message Batches API.
- `AzureOpenAIClient`, the concrete implementation for Azure OpenAI. Since Azure OpenAI does not expose an equivalent asynchronous batch endpoint, this client runs requests synchronously inside `submit_batch()` and persists results to disk as JSONL at `outputs/azure_batches/{batch_id}.jsonl`. `check_batch()` and `get_results()` read from that file, so `--resume` works the same way as with Anthropic. PDFs are passed inline as base64 file content blocks; `upload_file()` simply verifies the file exists and returns its absolute path as the `file_id`.
- The `_PROVIDERS` registry mapping provider name to client class, and the `get_llm_client()` factory function used by the rest of the codebase.

**`parser.py`** — Parses and validates LLM responses. Defines:
- The `Finding` dataclass with all output fields, including `snippet_id`, `timestamp`, and `finding_hash` computed in `__post_init__`.
- The `ParseError` dataclass for errors, also auto-timestamped.
- `_clean_json_text()` strips common LLM artefacts (markdown fences, preamble text).
- `_make_snippet_id(filename, category, seq)` builds a human-readable unique identifier of the form `{stem}__{CATEGORY}__{NNNN}`.
- `parse_result_item()` processes one batch result: handles API failures, missing text, JSON-parse errors, invalid structures, missing required fields, and hallucinated codes (codes returned by the LLM that are not in the category's framework). Valid findings are wrapped in `Finding` objects.
- `parse_all_results()` aggregates over all result items in one category's batch.

**`export.py`** — Writes outputs to disk. Defines the canonical column orders `FINDINGS_COLUMNS`, `ERRORS_COLUMNS`, and `VERIFIED_FINDINGS_COLUMNS` (the finding columns plus the appended verification fields), and the functions `save_findings()` (used for the clean `coded_findings.xlsx`), `save_errors()`, `save_verified_findings()` (the merged `coded_findings_verified.xlsx`), `save_file_mapping()` / `load_file_mapping()` for the `(filename, file_id)` mapping, and `save_batch_metadata()` / `load_batch_metadata()` for the batch metadata used by `--resume`.

**`verify.py`** — Verifies that every coded snippet really appears in its source PDF — the pipeline's safeguard against fabricated quotes. Defines:
- `normalize()`, the normalisation pipeline applied identically to snippet and document (Unicode NFKC, curly-quote/dash/ellipsis folding, de-hyphenation of line breaks, whitespace collapse, optional case-folding) so that differences introduced purely by PDF extraction don't cause a true quote to be reported as missing.
- `extract_doc_text()`, which uses PyMuPDF to read and normalise a PDF page by page, caching the result so each document is read only once regardless of how many snippets came from it. It strips recurring page headers and footers first (margin text blocks that repeat across pages, with page numbers normalised away) — without this, a running header spliced into a sentence that spans a page break would break the match for a genuine verbatim quote. A document yielding almost no text is flagged as a scan (`no_text_layer`) rather than producing spurious failures.
- `verify_one()`, the **match ladder**: exact normalised match (`verified`) → fragmented match for elided `...` quotes (`verified_fragmented`) → `rapidfuzz` fuzzy match (`verified_fuzzy` above the threshold, `near_match` in the grey band) → `not_found`. Short snippets cannot be auto-verified by fuzzy match alone (they can hit a high score by chance), so they are demoted to `near_match`.
- `verify_findings()`, the orchestrator that runs the ladder over every finding and returns one `VerificationResult` each, plus `summarize()` for a status tally.
- A standalone CLI (`python verify.py`) that re-checks an existing findings file without re-invoking the LLM — it appends the verification columns to whatever rows it reads and writes `coded_findings_verified.xlsx` (it never modifies the input), which makes it cheap to re-run while tuning thresholds. The same `verify_findings()` function is called inline by `main.py` so verification always runs at the end of a full pipeline run.

**`main.py`** — Pipeline orchestrator and CLI entry point. Provides:
- `_setup_logging()` configures dual-stream logging (INFO to console, DEBUG to `outputs/pipeline.log`).
- Document-discovery utilities: `_convert_word_to_pdf()` and `_convert_pptx_to_pdf()` (Windows + Office only), `_convert_to_pdf()` dispatching by extension, and `discover_documents()` which finds all supported files in `TARGET_FOLDER` and converts non-PDFs into a `_converted_pdfs/` subfolder.
- `upload_pdfs()` calls `client.upload_file()` on every PDF.
- `build_custom_id_mapping()` reconstructs the `custom_id → filename` mapping using the same `_make_custom_id` helper as the submit side.
- `estimate_cost()` prints a rough per-run cost estimate, parameterised by provider (batch pricing for Anthropic, synchronous pricing for Azure).
- `_run_category()` runs the full per-category cycle: build prompt → save prompt copy → submit batch (or reuse on resume) → poll → fetch → parse.
- `run_pipeline()` is the main entry point. After parsing, it always runs the snippet-verification stage (`verify.verify_findings()`), then writes two files: `coded_findings.xlsx` via `save_findings()` containing only the findings that verified (every status except `not_found`) and no verification columns, and `coded_findings_verified.xlsx` via `save_verified_findings()` containing all findings annotated with their verdicts. It folds a verified/total tally and a kept/dropped count into the final summary. `main()` parses CLI arguments (`--resume`) and dispatches.

### 3.2 — How the files connect

The dependency graph is acyclic and shallow:

- `config.py` is imported by everything that needs settings.
- `labels.py` exposes `Label`, which `prompt.py` consumes to build the prompt text.
- `api.py` defines the LLM interface and is consumed by `main.py`. It only depends on `config.py`.
- `parser.py` consumes `BatchResultItem` from `api.py` and produces `Finding` and `ParseError` objects.
- `verify.py` consumes `Finding` objects (or rows of an existing findings file) and produces `VerificationResult` objects. It depends only on `config.py` (and `export.py` for its standalone CLI write); `export.py` never imports it, so the graph stays acyclic.
- `export.py` consumes `Finding` and `ParseError` from `parser.py` and `VerificationResult`-shaped rows, and writes them to disk.
- `main.py` ties everything together. It loads labels (`labels.py`), builds prompts per category (`prompt.py`), gets an LLM client (`api.py`), submits and polls batches, parses results (`parser.py`), writes outputs (`export.py`), and verifies snippets (`verify.py`).

A new provider only requires changes inside `api.py` (a new subclass plus an entry in `_PROVIDERS`); a new output format only requires changes inside `export.py`; a new prompt strategy only requires changes inside `prompt.py`.

### 3.3 — Contents of the `outputs/` folder

After a run, `outputs/` contains the following files.

**`coded_findings.xlsx`** — The clean deliverable. One row per coded snippet whose quote was located in its source PDF; snippets with verification status `not_found` are excluded (they remain in `coded_findings_verified.xlsx`). No verification columns are added here. Columns:

| Column | What it is |
|---|---|
| `snippet_id` | Unique identifier encoding source filename and category (e.g. `Adelaiye__2024__CPVR_EFFECTIVENESS__0001`). |
| `filename` | Source PDF filename. |
| `label_category` | Category this snippet was coded under. |
| `label_code` | Specific code within the category. |
| `snippet` | Verbatim quote from the document. |
| `page_number` | Page where the snippet was found. |
| `reasoning` | The LLM's explanation for why the snippet was coded under this label. |
| `confidence` | `high`, `medium`, or `low`. |
| `timestamp` | ISO-8601 timestamp of when the row was produced. |
| `finding_hash` | Stable hash for deduplication across runs. |

**`errors.xlsx`** — Any issues encountered during processing (failed uploads, malformed JSON, hallucinated codes, missing fields, API errors). Each row contains `filename`, `label_category`, `error_type`, `error_message`, `raw_text` (truncated to 2 000 characters), and `timestamp`.

**`coded_findings_verified.xlsx`** — The full audit file. Every finding (including the `not_found` rows excluded from `coded_findings.xlsx`), with all the `coded_findings.xlsx` columns above plus these appended verification columns:

| Column | What it is |
|---|---|
| `verification_status` | `verified` (exact), `verified_fragmented` (elided `...` quote, all parts found), `verified_fuzzy` (minor edits), `near_match` (likely present, flagged for review), `not_found` (could not be located — strongest fabrication signal), `no_text_layer` (scanned PDF, cannot verify), or `pdf_missing` (source file not on disk). |
| `match_score` | Best similarity score, 0–100. |
| `match_method` | `exact`, `fragmented`, `fuzzy`, or `none`. |
| `matched_page` | The *physical* PDF page the snippet was found on (may differ from `page_number` — see below). |
| `page_ok` | Whether `matched_page` is within `VERIFY_PAGE_TOLERANCE` of the claimed page. |
| `matched_text` | For non-exact matches, the document text that matched, so a human can adjudicate. |

**`prompt_used__{CATEGORY}.txt`** — One plain-text file per category, containing the exact prompt that was sent to the LLM for that category. Saved for reproducibility and audit purposes.

**`batch_metadata.json`** — JSON file with two arrays: `file_rows` (the `(filename, file_id)` pairs uploaded) and `batches` (the `(category, batch_id)` pairs submitted so far, in submission order). This is the file read by `--resume`.

**`uploaded_file_ids.xlsx`** — Two-column Excel (`filename`, `file_id`) saved immediately after each upload. Provides a separate, human-readable copy of the upload mapping.

**`pipeline.log`** — Full DEBUG log with timestamps. Console logs are at INFO level; this file captures DEBUG-level detail for post-mortem inspection.

**`azure_batches/{batch_id}.jsonl`** (Azure provider only) — One JSONL file per submitted batch, containing one line per request with the `custom_id` and the result (either `succeeded` with text, or `errored` with an error message). This is the file that the Azure client's `check_batch()` and `get_results()` read from, and that makes `--resume` work for Azure runs.

### 3.4 — Contents of the `labels/` folder

The `labels/` folder contains the coding framework — a single Excel file at `labels/coding_framework.xlsx`. This file is the only input that defines what the LLM will look for; it is loaded by `labels.load_labels()` and validated against the schema described in Part 2, Step 3.

The file must contain these columns: `category`, `code`, `description`, `inclusion_criteria`, and (recommended) `exclusion_criteria`. Each row defines one label. Rows that share a `category` value form a category, and the pipeline runs one LLM batch per category. The `(category, code)` pair must be unique across the file.

---

## Part 4 — Design Choices

### 4.1 — Category-based batching

**One batch per category, not one batch for everything.** A single prompt containing every label in a large framework forces the LLM to juggle many definitions simultaneously, pushing it toward shallower judgements and more hallucinated codes. Running separately per category keeps each prompt focused on a coherent set of labels. The trade-off is more batches, but on Anthropic batches are 50% of standard pricing and run in parallel server-side, so wall-clock time is dominated by the slowest batch rather than the sum.

**Categories drive the entire pipeline structure.** The `OrderedDict` returned by `group_labels_by_category()` preserves first-appearance order, so runs are deterministic and reproducible: the same Excel file produces the same sequence of category batches every time.

### 4.2 — Single-upload, multi-batch reuse (Anthropic)

PDFs are uploaded once and referenced by `file_id` across all category batches. Since one document is read against every category, uploading separately for each category would multiply the upload cost by the number of categories. Anthropic's Files API returns an ID that can be referenced from any number of batches, so a single upload covers them all. The `(filename, file_id)` mapping is saved to `uploaded_file_ids.xlsx` immediately after each upload — this makes `--resume` able to skip the upload step entirely.

### 4.3 — Provider abstraction

`api.BaseLLMClient` is an abstract base class with four methods (`upload_file`, `submit_batch`, `check_batch`, `get_results`) plus a concrete `poll_until_complete()` helper. The rest of the codebase only references the abstract interface and the `BatchStatus` / `BatchResultItem` dataclasses. Adding a new provider means writing one new class and registering it in the `_PROVIDERS` dict at the bottom of `api.py`. No other file needs to change.

This is why the same pipeline works for both Anthropic (with a real server-side batch API) and Azure OpenAI (where the client emulates batches using synchronous calls and on-disk JSONL).

### 4.4 — Single source of truth for `custom_id`

Batch `custom_id` values are round-tripped: the client sends them with each request and the API echoes them back on results. The pipeline uses that echo to attribute each result to a filename. If the submit-side and lookup-side built the ID with different formats, every attribution would silently fail and the `filename` column of the output would be meaningless. `api._make_custom_id(filename, idx)` is the single source of truth; both `AnthropicClient.submit_batch` / `AzureOpenAIClient.submit_batch` and `main.build_custom_id_mapping` call it.

### 4.5 — Three identifiers per finding

Each finding carries three different identifiers that serve distinct needs:

- `snippet_id` — Human-readable, encodes provenance (`{source_stem}__{CATEGORY}__{NNNN}`). Researchers can read it and immediately know where a row came from. Unique within a run.
- `timestamp` — ISO-8601 timestamp recording when the row was produced. Useful for audit trails and for knowing which rows came from which pipeline execution after multiple re-runs.
- `finding_hash` — SHA-256 of `(filename, category, code, snippet)`, truncated to 16 hex characters. Stable across runs, so re-running the pipeline (e.g. after adding new documents) lets you concatenate outputs and deduplicate on this column.

### 4.6 — Hallucination detection over silent dropping

LLMs sometimes invent labels. The parser does **not** silently drop them; it surfaces them as `hallucinated_code` errors with the raw text retained. This lets the researcher see what the LLM tried to code and either refine the label definition or add the code to the framework. The valid-codes check is **per category**: the set of valid codes passed into the parser for a given batch is the set of codes in that batch's category only, so a code that is valid in category A but not in B will be flagged if the LLM returns it during the category-B batch.

### 4.7 — Idempotent resume

`batch_metadata.json` always reflects what has been submitted, because it is rewritten incrementally after each category's submission. On `--resume`, the pipeline loads that file, builds a `{category: batch_id}` lookup, and for each category in the framework: if it is in the lookup, the existing `batch_id` is reused (no submission); otherwise a fresh batch is submitted and the metadata updated. Either way, all category batches end up being polled, fetched, and parsed. So a resume after a full-success run just re-downloads and re-parses (cheap and idempotent), and a resume after a partial crash submits exactly the missing categories and finishes the rest.

For Azure, the same logic works because the JSONL file at `outputs/azure_batches/{batch_id}.jsonl` plays the role that Anthropic's server-side storage plays for Claude batches: it is the durable record of a completed batch.

### 4.8 — Centralised configuration

`config.py` is the only place with hard-coded values. The rest of the codebase imports what it needs from there, so swapping providers, changing models, or relocating folders never requires editing more than one file.

### 4.9 — Cost considerations

Anthropic batch pricing is 50% of standard pricing. Very rough per-request estimates (one request = one document × one category):

Azure OpenAI is invoked synchronously by this pipeline (no batch discount); actual figures depend on the specific Azure deployment and SKU.

### 4.10 — Snippet verification as a separate, always-on stage

Every snippet is supposed to be a verbatim quote, but the model is trusted on that at coding time. The verification stage (`verify.py`) closes that gap by extracting each PDF's own text and confirming the snippet is really there. Several choices shape it:

**A graded match ladder, not a boolean test.** A naive `snippet in document_text` check reports huge numbers of *false* "not found" results, because PDF extraction and harmless model tidying introduce differences that are invisible to a human: line-break hyphenation, ligatures, smart quotes, en/em dashes, and exotic spaces. So both the snippet and the document are run through the *same* normalisation pipeline, and matching then climbs a ladder from cheap-and-certain to forgiving-and-scored — exact, fragmented (for elided `...` quotes), then fuzzy — recording which rung succeeded. A passage that genuinely isn't there falls all the way through to `not_found`, which is the strongest available signal of a fabricated quote.

**Surface, don't assert.** Mirroring the parser's treatment of hallucinated codes, verification never silently passes or fails a borderline case. The grey band between the fuzzy and near thresholds becomes `near_match`, and the matched document text is stored so a researcher can adjudicate. Short snippets cannot be auto-verified by fuzzy score alone (they can reach a high score by coincidence), so they are deliberately demoted to `near_match`.

**Extract once per document.** A corpus has many snippets per file, so each PDF's text is extracted and cached once and reused across all of that document's findings.

**Strip running headers and footers.** PDF extraction routinely splices a page's running header (author, journal, page number) into the middle of a sentence that continues onto the next page. Left in, that intrusion breaks the match for a quote that is genuinely verbatim. The extractor therefore detects margin text blocks whose content recurs across pages (ignoring the varying page number) and removes them before matching. On the bundled Adelaiye paper this is the difference between four spurious failures and all snippets verifying.

**Page numbers are a soft signal.** Models report the *printed* page number, which is offset from the physical PDF page by cover pages and front matter, so a page mismatch is recorded as a warning (`page_ok = False`) rather than a failure.

**Two files: a clean deliverable and a full audit.** The pipeline writes `coded_findings.xlsx` with only the findings that verified (every status except `not_found`) and no verification columns — the file a researcher actually works from — and `coded_findings_verified.xlsx` with every finding plus its verdict. Nothing is silently lost: a `not_found` is dropped from the clean file but preserved, with its score and matched text, in the audit file, so a reviewer can confirm whether it was a genuine fabrication or (as happened on the bundled paper before the header fix) an extraction artifact. Only `not_found` is filtered out; `near_match`, `no_text_layer`, and `pdf_missing` rows are kept in the clean file, since they are not confirmed-absent quotes.

**It always runs, and it's free.** Verification needs no API credits, so it runs automatically at the end of every full pipeline run. It is also exposed as a standalone CLI (`python verify.py`) that re-checks an existing findings file and writes `coded_findings_verified.xlsx` without modifying the input, which makes it cheap to tune the thresholds in `config.py` against your own corpus without re-coding anything. The behaviour of the normalisation pipeline and the ladder is locked down by `test_verify.py`, which exercises each known failure mode (hyphenation, ligatures, smart quotes, page-spanning quotes, elision, fabrication, and scanned PDFs).


## Part 5 — Data Privacy

This pipeline sends document content to either Azure OpenAI or Anthropic for processing. Both providers state that API-submitted data is not used to train their foundation models by default. Azure OpenAI additionally provides enterprise-grade security, compliance, and regional data residency controls through Microsoft's Azure platform; prompts and outputs are not shared with OpenAI and are not used to improve Microsoft or OpenAI models without explicit permission. For organisations operating within the European Union, Azure OpenAI can be deployed in EU Azure regions and is covered by Microsoft's compliance commitments, including support for GDPR compliance, EU Data Boundary controls for eligible services, and enterprise data governance requirements. Anthropic similarly states that inputs and outputs submitted through the Anthropic API are not used for model training unless the customer explicitly opts in. Users should ensure that any documents processed through the pipeline comply with their organisation's data governance, confidentiality, and regulatory requirements. For details, see Microsoft's Azure OpenAI data privacy documentation (https://learn.microsoft.com/en-us/legal/cognitive-services/openai/data-privacy), Microsoft's EU Data Boundary documentation (https://learn.microsoft.com/en-us/privacy/eudb/eu-data-boundary-overview), Microsoft's GDPR compliance documentation (https://learn.microsoft.com/en-us/azure/compliance/offerings/offering-gdpr), and Anthropic's API data usage policy (https://privacy.anthropic.com/en/articles/7996868-is-my-data-used-for-model-training).



## Part 6 — Next Steps

### Change 1 — Done

Verify, for each extracted snippet, that it really exists in the source text. **Implemented** in `verify.py` and run automatically at the end of every pipeline run (see Part 4.10). Snippets that cannot be located are dropped from `coded_findings.xlsx`, and the full verdict for every finding is written to `outputs/coded_findings_verified.xlsx`.

### Change 2

Identification of the page number should not be done by the LLM, but should be done in the verification phase (when verifying if the snippet actually exists): the page number within the actual pdf (and not the page number written in the document) should be retained then.

### Change 3

Metrics to identify false positives and false negatives. This still requires a hand-coded gold standard to measure against; the snippet-existence check in Change 1 already provides one half of the false-positive story (fabricated quotes), but precision and recall against expert coding remain future work.
