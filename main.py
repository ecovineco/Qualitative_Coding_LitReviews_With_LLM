"""Main orchestrator for the AI-powered literature review pipeline.

This script ties all modules together.  The pipeline flow is:

    1. Load configuration and the coding framework (labels).
    2. Group labels by category.
    3. Discover documents (PDF, Word, PowerPoint) in the target folder.
    4. Upload PDFs to the LLM provider (or resume from saved mapping).
    5. For each category (sequentially):
         a. Build a prompt containing ONLY that category's labels.
         b. Submit a batch of (prompt + document) requests — one per PDF.
         c. Save metadata (so resume can pick up).
         d. Poll until the batch completes.
         e. Download, parse, and validate results.
         f. Accumulate findings and errors.
    6. Export all findings (merged across categories) and the error log.

Usage:
    Full run from scratch::

        python main.py

    Resume from a previous run (skip upload; skip already-submitted
    category batches; re-fetch and parse all category results)::

        python main.py --resume

    Dry run (build requests but don't submit)::

        Set DRY_RUN = True in config.py, then:
        python main.py

Command-line arguments:
    --resume    Skip the upload step and skip submission for any category
                whose batch ID is already in the saved metadata.  Any
                categories that were not yet submitted are submitted now.
                All category batches are then polled, fetched, and parsed.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import config
from api import get_llm_client, BatchResultItem, _make_custom_id
from export import (
    load_batch_metadata,
    load_file_mapping,
    save_batch_metadata,
    save_errors,
    save_file_mapping,
    save_findings,
    save_verification,
)
from labels import (
    Label,
    get_valid_codes,
    group_labels_by_category,
    load_labels,
)
from parser import Finding, ParseError, parse_all_results
from prompt import build_analysis_prompt
from verify import VERIFIED_STATUSES, summarize, verify_findings


# ══════════════════════════════════════════════════════════════════════
# Logging setup
# ══════════════════════════════════════════════════════════════════════


def _setup_logging() -> None:
    """Configure logging for the pipeline.

    Logs are written to both the console (INFO level) and a file in
    the output directory (DEBUG level).
    """
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    log_file = os.path.join(config.OUTPUT_DIR, "pipeline.log")

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, mode="a", encoding="utf-8"),
        ],
    )
    # Reduce noise from the requests library.
    logging.getLogger("urllib3").setLevel(logging.WARNING)


logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════
# Document discovery & conversion
# ══════════════════════════════════════════════════════════════════════


def _convert_word_to_pdf(source_path: str, output_dir: str) -> str:
    """Convert a Word file to PDF using Microsoft Word via COM automation.

    Requires: ``pip install docx2pdf`` and Microsoft Word installed.
    """
    try:
        from docx2pdf import convert
    except ImportError:
        raise RuntimeError(
            "docx2pdf is required to convert Word files. "
            "Install it with: pip install docx2pdf"
        )

    os.makedirs(output_dir, exist_ok=True)
    source = Path(source_path)
    output_path = os.path.join(output_dir, source.stem + ".pdf")
    convert(str(source), output_path)

    if not os.path.exists(output_path):
        raise RuntimeError(f"Conversion failed, output not found: {output_path}")

    logger.info("Converted %s -> %s", source.name, output_path)
    return output_path


def _convert_pptx_to_pdf(source_path: str, output_dir: str) -> str:
    """Convert a PowerPoint file to PDF via Microsoft PowerPoint / COM.

    Requires: ``pip install pywin32`` and Microsoft PowerPoint installed.
    """
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32 is required to convert PowerPoint files. "
            "Install it with: pip install pywin32"
        )

    os.makedirs(output_dir, exist_ok=True)
    source = Path(source_path)
    output_path = str(Path(output_dir) / (source.stem + ".pdf"))

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    try:
        presentation = powerpoint.Presentations.Open(
            str(source.resolve()), WithWindow=False
        )
        presentation.SaveAs(str(Path(output_path).resolve()), 32)  # 32 = ppSaveAsPDF
        presentation.Close()
    finally:
        powerpoint.Quit()

    if not os.path.exists(output_path):
        raise RuntimeError(f"Conversion failed, output not found: {output_path}")

    logger.info("Converted %s -> %s", source.name, output_path)
    return output_path


def _convert_to_pdf(source_path: str, output_dir: str) -> str:
    """Dispatch conversion to the appropriate handler based on extension."""
    ext = Path(source_path).suffix.lower()
    if ext in {".docx", ".doc"}:
        return _convert_word_to_pdf(source_path, output_dir)
    elif ext in {".pptx", ".ppt"}:
        return _convert_pptx_to_pdf(source_path, output_dir)
    else:
        raise ValueError(f"Unsupported format for conversion: {ext}")


# Supported document extensions (case-insensitive matching).
_SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".doc", ".pptx", ".ppt"}


def discover_documents(
    target_folder: str, max_docs: int | None = None
) -> List[str]:
    """Find all supported documents and convert non-PDFs to PDF.

    Args:
        target_folder: Path to the directory containing documents.
        max_docs: If set, limit the number of documents returned (for testing).

    Returns:
        Sorted list of PDF file paths (originals + converted).

    Raises:
        FileNotFoundError: If the target folder does not exist.
        RuntimeError: If no supported documents are found.
    """
    folder = Path(target_folder)
    if not folder.exists():
        raise FileNotFoundError(f"Target folder not found: {folder}")

    all_files: List[Path] = []
    for f in sorted(folder.iterdir()):
        if f.is_file() and f.suffix.lower() in _SUPPORTED_EXTENSIONS:
            all_files.append(f)

    if not all_files:
        raise RuntimeError(
            f"No supported documents found in: {folder}\n"
            f"Supported formats: {', '.join(sorted(_SUPPORTED_EXTENSIONS))}"
        )

    pdf_paths: List[str] = []
    to_convert: List[Path] = []
    for f in all_files:
        if f.suffix.lower() == ".pdf":
            pdf_paths.append(str(f))
        else:
            to_convert.append(f)

    if to_convert:
        converted_dir = os.path.join(folder, "_converted_pdfs")
        logger.info(
            "Converting %d non-PDF document(s) to PDF...", len(to_convert)
        )
        for doc in to_convert:
            converted_pdf = _convert_to_pdf(str(doc), converted_dir)
            pdf_paths.append(converted_pdf)

    pdf_paths.sort()

    if max_docs is not None:
        pdf_paths = pdf_paths[:max_docs]

    logger.info(
        "Found %d document(s) in %s (%d native PDF, %d converted)",
        len(pdf_paths),
        folder,
        len(pdf_paths) - len(to_convert),
        len(to_convert),
    )
    return pdf_paths


# ══════════════════════════════════════════════════════════════════════
# Upload & custom-id helpers
# ══════════════════════════════════════════════════════════════════════


def upload_pdfs(pdf_paths: List[str]) -> List[Tuple[str, str]]:
    """Upload all PDFs and return ``[(filename, file_id), ...]``."""
    client = get_llm_client()
    file_rows: List[Tuple[str, str]] = []

    for i, pdf_path in enumerate(pdf_paths, start=1):
        filename = os.path.basename(pdf_path)
        logger.info("Uploading [%d/%d]: %s", i, len(pdf_paths), filename)
        file_id = client.upload_file(pdf_path)
        file_rows.append((filename, file_id))
        logger.info("  -> file_id: %s", file_id)

    return file_rows


def build_custom_id_mapping(
    file_rows: List[Tuple[str, str]],
) -> Dict[str, str]:
    """Build a mapping from batch ``custom_id`` to filename.

    Uses the same ``_make_custom_id`` helper as ``submit_batch`` so the
    two sides are guaranteed to agree.

    Args:
        file_rows: List of ``(filename, file_id)`` tuples.

    Returns:
        Dictionary mapping ``custom_id`` -> ``filename``.
    """
    mapping: Dict[str, str] = {}
    for idx, (filename, _) in enumerate(file_rows, start=1):
        custom_id = _make_custom_id(filename, idx)
        mapping[custom_id] = filename
    return mapping


# ══════════════════════════════════════════════════════════════════════
# Cost estimate
# ══════════════════════════════════════════════════════════════════════


def estimate_cost(num_files: int, num_categories: int, model: str) -> None:
    """Print a rough cost estimate for the entire run.

    The pipeline makes ``num_files * num_categories`` total requests
    (one per document per category).  Estimate scales linearly.

    The per-request figure depends on which provider is active:

    - Anthropic models use batch pricing (50% of standard).
    - Azure OpenAI calls in this pipeline are synchronous (no batch
      discount); figures here are placeholders since pricing depends
      on the specific deployment.
    """
    provider = config.LLM_PROVIDER.lower().strip()
    model_lc = model.lower()

    if provider == "anthropic":
        if "opus" in model_lc:
            cost_per_request = 0.06
        elif "sonnet" in model_lc:
            cost_per_request = 0.01
        else:
            cost_per_request = 0.005
        pricing_note = "batch pricing"
    elif provider == "azure":
        # Placeholder: actual cost depends on the Azure deployment SKU.
        cost_per_request = 0.02
        pricing_note = "synchronous pricing (no batch discount)"
    else:
        cost_per_request = 0.005
        pricing_note = "unknown provider"

    total_requests = num_files * num_categories
    estimated = total_requests * cost_per_request
    logger.info(
        "Rough cost estimate: ~$%.2f total "
        "(%d documents x %d categories = %d requests with %s [%s], %s; "
        "actual cost depends on document length)",
        estimated,
        num_files,
        num_categories,
        total_requests,
        model,
        provider,
        pricing_note,
    )


# ══════════════════════════════════════════════════════════════════════
# Per-category run
# ══════════════════════════════════════════════════════════════════════


def _sanitize_for_filename(text: str) -> str:
    """Sanitise a string so it's safe to use in a filename."""
    return re.sub(r"[^A-Za-z0-9_-]", "_", text)[:60].strip("_")


def _run_category(
    category: str,
    cat_labels: List[Label],
    file_rows: List[Tuple[str, str]],
    existing_batch_id: str | None,
    prompt_dir: str,
) -> Tuple[str, List[Finding], List[ParseError]]:
    """Run (or resume) a single category: submit → poll → fetch → parse.

    Args:
        category: Category name.
        cat_labels: Labels belonging to this category.
        file_rows: ``(filename, file_id)`` tuples for all documents.
        existing_batch_id: If resuming and this category was already
            submitted, its batch ID; otherwise ``None``.
        prompt_dir: Directory to save the per-category prompt in.

    Returns:
        ``(batch_id, findings, errors)`` for this category.
    """
    client = get_llm_client()

    # --- Build category-specific prompt & save a copy -------------------
    prompt = build_analysis_prompt(
        cat_labels,
        include_confidence=config.INCLUDE_CONFIDENCE,
    )
    safe_cat = _sanitize_for_filename(category)
    prompt_path = os.path.join(prompt_dir, f"prompt_used__{safe_cat}.txt")
    with open(prompt_path, "w", encoding="utf-8") as f:
        f.write(prompt)
    logger.info("[%s] Prompt saved to %s", category, prompt_path)

    valid_codes = get_valid_codes(cat_labels)

    # --- Submit batch (or reuse an already-submitted one on resume) -----
    if existing_batch_id:
        logger.info(
            "[%s] Reusing already-submitted batch '%s'",
            category,
            existing_batch_id,
        )
        batch_id = existing_batch_id
    else:
        if config.DRY_RUN:
            logger.info(
                "[%s] DRY RUN: would submit %d requests.",
                category,
                len(file_rows),
            )
            return ("DRY_RUN", [], [])
        logger.info(
            "[%s] Submitting batch with %d requests...",
            category,
            len(file_rows),
        )
        batch_id = client.submit_batch(file_rows, prompt)
        logger.info("[%s] Batch submitted: %s", category, batch_id)

    # --- Poll until complete -------------------------------------------
    logger.info("[%s] Polling batch '%s'...", category, batch_id)
    batch_status = client.poll_until_complete(batch_id)

    # --- Download & parse ----------------------------------------------
    logger.info("[%s] Downloading results...", category)
    results: List[BatchResultItem] = client.get_results(batch_status)

    custom_id_map = build_custom_id_mapping(file_rows)
    findings, errors = parse_all_results(
        results, custom_id_map, valid_codes, category
    )

    return (batch_id, findings, errors)


# ══════════════════════════════════════════════════════════════════════
# Main pipeline
# ══════════════════════════════════════════════════════════════════════


def run_pipeline(resume: bool = False) -> None:
    """Execute the full literature review pipeline.

    Args:
        resume: If ``True``, skip upload and skip submission for any
            category whose batch was already submitted in a previous run.
    """
    _setup_logging()
    logger.info("=" * 60)
    logger.info("AI-POWERED LITERATURE REVIEW PIPELINE")
    logger.info("=" * 60)

    # --- Paths for saved state ------------------------------------------
    mapping_path = os.path.join(config.OUTPUT_DIR, "uploaded_file_ids.xlsx")
    metadata_path = os.path.join(config.OUTPUT_DIR, "batch_metadata.json")
    findings_path = os.path.join(config.OUTPUT_DIR, "coded_findings.xlsx")
    errors_path = os.path.join(config.OUTPUT_DIR, "errors.xlsx")
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)

    # ── Step 1: Load labels & group by category ────────────────────────
    logger.info("Step 1: Loading coding framework...")
    labels = load_labels(config.LABELS_EXCEL)
    labels_by_category = group_labels_by_category(labels)
    logger.info(
        "Loaded %d labels across %d categories: %s",
        len(labels),
        len(labels_by_category),
        list(labels_by_category.keys()),
    )

    # ── Step 2: Get the uploaded file rows ─────────────────────────────
    if resume:
        logger.info("RESUME MODE: loading saved state...")
        file_rows, existing_batches = load_batch_metadata(metadata_path)
        logger.info(
            "Resuming with %d files and %d existing category batches.",
            len(file_rows),
            len(existing_batches),
        )
    else:
        logger.info("Step 2: Discovering documents...")
        pdf_paths = discover_documents(config.TARGET_FOLDER, config.MAX_PDFS)

        logger.info("Step 3: Uploading PDFs...")
        file_rows = upload_pdfs(pdf_paths)
        save_file_mapping(file_rows, mapping_path)

        existing_batches = []

    # ── Step 4: Cost estimate ──────────────────────────────────────────
    active_model = (
        config.AZURE_OPENAI_MODEL
        if config.LLM_PROVIDER.lower().strip() == "azure"
        else config.ANTHROPIC_MODEL
    )
    estimate_cost(
        num_files=len(file_rows),
        num_categories=len(labels_by_category),
        model=active_model,
    )

    # Turn existing_batches into a lookup for convenience.
    existing_by_cat: Dict[str, str] = {
        cat: bid for cat, bid in existing_batches
    }

    # ── Step 5: Loop over categories ───────────────────────────────────
    all_findings: List[Finding] = []
    all_errors: List[ParseError] = []
    # Build up the up-to-date (category, batch_id) list as we go, so the
    # metadata file always reflects what has actually been submitted.
    category_batches: List[Tuple[str, str]] = []

    for cat_idx, (category, cat_labels) in enumerate(
        labels_by_category.items(), start=1
    ):
        logger.info(
            "── Category %d/%d: %s (%d labels) ──",
            cat_idx,
            len(labels_by_category),
            category,
            len(cat_labels),
        )

        batch_id, findings, errors = _run_category(
            category=category,
            cat_labels=cat_labels,
            file_rows=file_rows,
            existing_batch_id=existing_by_cat.get(category),
            prompt_dir=config.OUTPUT_DIR,
        )

        if batch_id == "DRY_RUN":
            # Skip metadata + accumulation in dry-run mode.
            continue

        category_batches.append((category, batch_id))
        # Persist metadata incrementally so a crash after this point is
        # fully recoverable.
        save_batch_metadata(file_rows, category_batches, metadata_path)

        all_findings.extend(findings)
        all_errors.extend(errors)

    if config.DRY_RUN:
        logger.info("DRY RUN: pipeline stopped before submission.")
        return

    # ── Step 6: Export merged results ──────────────────────────────────
    logger.info("Step 6: Exporting merged results...")
    save_findings(all_findings, findings_path)
    save_errors(all_errors, errors_path)

    # ── Step 7: Verify snippets against their source PDFs ──────────────
    # Always run: confirms every coded snippet is a real verbatim quote
    # from its source document.  Needs no API credits.
    logger.info("Step 7: Verifying snippets against source PDFs...")
    verification = verify_findings(all_findings, pdf_dir=config.PDF_SOURCE_DIR)
    save_verification(verification, config.VERIFY_OUTPUT)
    verify_counts = summarize(verification)
    verified_total = sum(verify_counts.get(s, 0) for s in VERIFIED_STATUSES)

    # ── Summary ────────────────────────────────────────────────────────
    logger.info("=" * 60)
    logger.info("PIPELINE COMPLETE")
    logger.info("  Documents processed: %d", len(file_rows))
    logger.info("  Categories run:      %d", len(category_batches))
    logger.info("  Findings extracted:  %d", len(all_findings))
    logger.info("  Errors logged:       %d", len(all_errors))
    logger.info(
        "  Snippets verified:   %d/%d (%s)",
        verified_total,
        len(verification),
        {k: verify_counts[k] for k in sorted(verify_counts)},
    )
    logger.info("  Findings file:       %s", findings_path)
    logger.info("  Errors file:         %s", errors_path)
    logger.info("  Verification file:   %s", config.VERIFY_OUTPUT)
    logger.info("=" * 60)


# ══════════════════════════════════════════════════════════════════════
# CLI entry point
# ══════════════════════════════════════════════════════════════════════


def main() -> None:
    """Parse command-line arguments and run the pipeline."""
    parser = argparse.ArgumentParser(
        description="AI-powered systematic literature review pipeline.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Examples:
  python main.py               # Full run from scratch
  python main.py --resume      # Resume from saved batch state
        """,
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Skip upload; skip already-submitted category batches.",
    )
    args = parser.parse_args()
    run_pipeline(resume=args.resume)


if __name__ == "__main__":
    main()
