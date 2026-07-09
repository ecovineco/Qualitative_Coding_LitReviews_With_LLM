"""Export module for writing pipeline outputs to disk.

This module is responsible for saving all outputs:
    - File ID mapping (Excel) — links filenames to uploaded file IDs.
    - Coded findings (Excel) — the main deliverable.  Columns include
      ``snippet_id`` (unique, indicates source + category), ``timestamp``
      (when the finding was parsed), and ``label_category``.
    - Error log (Excel) — any issues encountered during processing.
    - Batch metadata (JSON) — list of (category, batch_id) pairs plus
      the uploaded file rows, for resume capability across multiple
      per-category batches.

Adding a new export format (e.g. CSV, SQLite, Google Sheets) only requires
adding a new function here — no changes to the rest of the codebase.

Typical usage::

    from export import save_findings, save_errors, save_file_mapping

    save_findings(findings, "outputs/coded_findings.xlsx")
    save_errors(errors, "outputs/errors.xlsx")
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd

from parser import Finding, ParseError

logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════
# Output column orders (single source of truth)
# ══════════════════════════════════════════════════════════════════════

FINDINGS_COLUMNS = [
    "snippet_id",
    "filename",
    "label_category",
    "label_code",
    "snippet",
    "page_number",
    "reasoning",
    "confidence",
    "timestamp",
    "finding_hash",
]

ERRORS_COLUMNS = [
    "filename",
    "label_category",
    "error_type",
    "error_message",
    "raw_text",
    "timestamp",
]

# Verification-specific fields appended after the standard finding fields
# to form the full audit file (``coded_findings_verified.xlsx``).  Note
# that ``page_number`` above is *not* repeated here: it is a base finding
# field, and verification populates it directly rather than exposing a
# separate "matched page" column.
_VERIFICATION_EXTRA_COLUMNS = [
    "verification_status",
    "match_score",
    "match_method",
    "matched_text",
]

VERIFIED_FINDINGS_COLUMNS = FINDINGS_COLUMNS + _VERIFICATION_EXTRA_COLUMNS


# ══════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════


def _ensure_dir(file_path: str | Path) -> None:
    """Create parent directories for a file path if they don't exist.

    Args:
        file_path: Path to a file (not a directory).
    """
    Path(file_path).parent.mkdir(parents=True, exist_ok=True)


# ══════════════════════════════════════════════════════════════════════
# File ID mapping
# ══════════════════════════════════════════════════════════════════════


def save_file_mapping(
    file_rows: List[Tuple[str, str]],
    output_path: str | Path,
) -> None:
    """Save the filename-to-file_id mapping as an Excel file.

    This mapping is saved immediately after upload so that the pipeline
    can resume without re-uploading if it crashes later.

    Args:
        file_rows: List of ``(filename, file_id)`` tuples.
        output_path: Where to save the ``.xlsx`` file.
    """
    _ensure_dir(output_path)
    df = pd.DataFrame(file_rows, columns=["filename", "file_id"])
    df.to_excel(str(output_path), index=False)
    logger.info("Saved file mapping (%d rows) to %s", len(df), output_path)


def load_file_mapping(mapping_path: str | Path) -> List[Tuple[str, str]]:
    """Load a previously saved file mapping from Excel.

    Useful for resuming the pipeline after a crash — skip the upload
    step and go straight to batch submission.

    Args:
        mapping_path: Path to the ``.xlsx`` file saved by
            ``save_file_mapping()``.

    Returns:
        List of ``(filename, file_id)`` tuples.

    Raises:
        FileNotFoundError: If the mapping file does not exist.
    """
    path = Path(mapping_path)
    if not path.exists():
        raise FileNotFoundError(f"File mapping not found: {path}")

    df = pd.read_excel(str(path), dtype=str)
    rows = list(zip(df["filename"].tolist(), df["file_id"].tolist()))
    logger.info("Loaded file mapping (%d rows) from %s", len(rows), path)
    return rows


# ══════════════════════════════════════════════════════════════════════
# Coded findings
# ══════════════════════════════════════════════════════════════════════


def save_findings(
    findings: List[Finding],
    output_path: str | Path,
) -> None:
    """Save coded findings to an Excel file.

    Each row is one finding (one snippet coded under one label from one
    document).  Columns are ordered per ``FINDINGS_COLUMNS``.

    Args:
        findings: List of ``Finding`` objects from the parser.
        output_path: Where to save the ``.xlsx`` file.
    """
    _ensure_dir(output_path)
    rows = [f.to_dict() for f in findings]
    df = pd.DataFrame(rows, columns=FINDINGS_COLUMNS)

    # Ensure deterministic column order even when some fields are missing.
    df = df[FINDINGS_COLUMNS]

    df.to_excel(str(output_path), index=False)
    logger.info("Saved %d findings to %s", len(findings), output_path)


# ══════════════════════════════════════════════════════════════════════
# Error log
# ══════════════════════════════════════════════════════════════════════


def save_errors(
    errors: List[ParseError],
    output_path: str | Path,
) -> None:
    """Save parsing/validation errors to an Excel file.

    Each row is one error encountered during result processing.  Columns
    are ordered per ``ERRORS_COLUMNS``.

    Args:
        errors: List of ``ParseError`` objects from the parser.
        output_path: Where to save the ``.xlsx`` file.
    """
    _ensure_dir(output_path)
    rows = [e.to_dict() for e in errors]
    df = pd.DataFrame(rows, columns=ERRORS_COLUMNS)
    df = df[ERRORS_COLUMNS]

    if not errors:
        logger.info("No errors to save — all results parsed successfully.")

    df.to_excel(str(output_path), index=False)
    logger.info("Saved %d errors to %s", len(errors), output_path)


# ══════════════════════════════════════════════════════════════════════
# Snippet verification
# ══════════════════════════════════════════════════════════════════════


def save_verified_findings(
    findings: List[Any],
    results: List[Any],
    output_path: str | Path,
) -> None:
    """Save the full audit file: every finding plus its verification info.

    Each row is one finding (with all the standard ``FINDINGS_COLUMNS``
    fields — including ``page_number``, which the caller is expected to
    have already set from the matching ``VerificationResult`` before
    calling this function) extended with the verification verdict for
    its snippet (status, match score/method, and the matched text for
    non-exact matches).  No findings are dropped — this is the complete
    record, including any ``not_found`` rows that are excluded from the
    deliverable ``coded_findings.xlsx``.

    Args:
        findings: List of ``parser.Finding`` objects, in order.
        results: List of ``verify.VerificationResult`` objects, in the
            same order as ``findings`` (one per finding).  Typed loosely
            so that ``export`` need not import ``verify`` (keeping the
            dependency one-directional and the import graph acyclic).
        output_path: Where to save the ``.xlsx`` file.
    """
    _ensure_dir(output_path)
    rows: List[Dict[str, Any]] = []
    for finding, result in zip(findings, results):
        row = finding.to_dict()
        verdict = result.to_dict()
        for col in _VERIFICATION_EXTRA_COLUMNS:
            row[col] = verdict.get(col)
        rows.append(row)
    df = pd.DataFrame(rows, columns=VERIFIED_FINDINGS_COLUMNS)
    df = df[VERIFIED_FINDINGS_COLUMNS]
    df.to_excel(str(output_path), index=False)
    logger.info("Saved %d verified findings to %s", len(rows), output_path)


# ══════════════════════════════════════════════════════════════════════
# Batch metadata (for resume capability across multiple category batches)
# ══════════════════════════════════════════════════════════════════════


def save_batch_metadata(
    file_rows: List[Tuple[str, str]],
    category_batches: List[Tuple[str, str]],
    output_path: str | Path,
) -> None:
    """Save batch metadata to a JSON file for resume capability.

    The pipeline submits one batch per category, sequentially.  This
    single file records all of them plus the shared upload mapping.  If
    the script crashes, ``--resume`` can figure out which categories
    were already submitted and pick up from there.

    Args:
        file_rows: The ``(filename, file_id)`` tuples submitted.
        category_batches: List of ``(category, batch_id)`` tuples, in
            the order they were submitted.
        output_path: Where to save the ``.json`` file.
    """
    _ensure_dir(output_path)
    metadata = {
        "file_rows": [
            {"filename": fn, "file_id": fid} for fn, fid in file_rows
        ],
        "batches": [
            {"category": cat, "batch_id": bid}
            for cat, bid in category_batches
        ],
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)
    logger.info(
        "Saved batch metadata (%d categories) to %s",
        len(category_batches),
        output_path,
    )


def load_batch_metadata(
    metadata_path: str | Path,
) -> Tuple[List[Tuple[str, str]], List[Tuple[str, str]]]:
    """Load batch metadata from a previously saved JSON file.

    Args:
        metadata_path: Path to the ``.json`` file saved by
            ``save_batch_metadata()``.

    Returns:
        A tuple of ``(file_rows, category_batches)``:
            - ``file_rows``: list of ``(filename, file_id)`` tuples.
            - ``category_batches``: list of ``(category, batch_id)``
              tuples, in submission order.

    Raises:
        FileNotFoundError: If the metadata file does not exist.
    """
    path = Path(metadata_path)
    if not path.exists():
        raise FileNotFoundError(f"Batch metadata not found: {path}")

    with open(path, "r", encoding="utf-8") as f:
        metadata = json.load(f)

    file_rows = [
        (row["filename"], row["file_id"]) for row in metadata.get("file_rows", [])
    ]
    category_batches = [
        (entry["category"], entry["batch_id"])
        for entry in metadata.get("batches", [])
    ]
    logger.info(
        "Loaded batch metadata: %d files, %d category batches",
        len(file_rows),
        len(category_batches),
    )
    return file_rows, category_batches
