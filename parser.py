"""Results parsing and validation module.

This module takes raw LLM text responses (expected to be JSON), parses them,
validates the structure against the coding framework, and produces clean
``Finding`` rows ready for export.

Each ``Finding`` is tagged with:
    - ``label_category``  — Which category batch produced this finding.
    - ``label_code``      — The specific code within that category.
    - ``snippet_id``      — A human-readable, globally unique identifier
                            that encodes the source document and category
                            (format: ``{stem}__{CATEGORY}__{NNNN}``).
    - ``timestamp``       — ISO-8601 timestamp set when the finding was
                            parsed (useful for audit trails and re-runs).

Validation catches:
    - Malformed JSON (LLM sometimes adds markdown fences or commentary).
    - Hallucinated label codes that don't exist in the framework.
    - Missing required fields in individual findings.

Typical usage::

    from parser import parse_all_results

    findings, errors = parse_all_results(
        results,
        custom_id_to_filename,
        valid_codes={"OBJECTIVES", "INNOVATION_MARKET", ...},
        category="CPVR_EFFECTIVENESS",
    )
"""

from __future__ import annotations

import hashlib
import json
import logging
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from api import BatchResultItem

logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════
# Data structures
# ══════════════════════════════════════════════════════════════════════


@dataclass
class Finding:
    """A single coded finding extracted from a document.

    Attributes:
        filename: Name of the source PDF.
        label_category: Category this finding was coded under.
        label_code: Code within the category.
        snippet: Verbatim quote from the document.
        page_number: Page where the snippet was found.
        reasoning: Why this snippet was coded under this label.
        confidence: LLM's self-assessed confidence (high/medium/low).
        snippet_id: Human-readable unique id that encodes source and
            category — format ``{stem}__{CATEGORY}__{NNNN}``.  Set at
            construction time by the parser.
        timestamp: ISO-8601 timestamp of when the finding was parsed.
        finding_hash: SHA-256 hash of (filename + category + code +
            snippet) — a stable identifier for deduplication across runs.
    """

    filename: str
    label_category: str
    label_code: str
    snippet: str
    page_number: Optional[int] = None
    reasoning: str = ""
    confidence: str = ""
    snippet_id: str = ""
    timestamp: str = ""
    finding_hash: str = ""

    def __post_init__(self) -> None:
        """Compute the deduplication hash and stamp the creation time."""
        raw = (
            f"{self.filename}|{self.label_category}|"
            f"{self.label_code}|{self.snippet}"
        )
        self.finding_hash = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]
        if not self.timestamp:
            self.timestamp = datetime.now().isoformat(timespec="seconds")

    def to_dict(self) -> Dict[str, Any]:
        """Convert the finding to a flat dictionary for export.

        Returns:
            Dictionary with all fields as string/int values.
        """
        return {
            "snippet_id": self.snippet_id,
            "filename": self.filename,
            "label_category": self.label_category,
            "label_code": self.label_code,
            "snippet": self.snippet,
            "page_number": self.page_number,
            "reasoning": self.reasoning,
            "confidence": self.confidence,
            "timestamp": self.timestamp,
            "finding_hash": self.finding_hash,
        }


@dataclass
class ParseError:
    """Records a parsing or validation error for a single result.

    Attributes:
        filename: Name of the source PDF.
        label_category: Category the batch was running (helps pinpoint
            which per-category pass produced the error).
        error_type: Category of the error (e.g. "json_parse_failed",
            "api_error", "no_text", "hallucinated_code").
        error_message: Human-readable error description.
        raw_text: The raw LLM response text (for debugging).
        timestamp: ISO-8601 timestamp of when the error was recorded.
    """

    filename: str
    label_category: str
    error_type: str
    error_message: str
    raw_text: str = ""
    timestamp: str = ""

    def __post_init__(self) -> None:
        if not self.timestamp:
            self.timestamp = datetime.now().isoformat(timespec="seconds")

    def to_dict(self) -> Dict[str, Any]:
        """Convert the error to a flat dictionary for export.

        Returns:
            Dictionary with all fields as string values.
        """
        return {
            "filename": self.filename,
            "label_category": self.label_category,
            "error_type": self.error_type,
            "error_message": self.error_message,
            "raw_text": self.raw_text[:2000],  # Excel cell-length safety.
            "timestamp": self.timestamp,
        }


# ══════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════


def _clean_json_text(text: str) -> str:
    """Remove common LLM artefacts from a JSON response.

    LLMs sometimes wrap JSON in markdown code fences or add preamble
    text.  This function strips those artefacts.

    Args:
        text: Raw text from the LLM.

    Returns:
        Cleaned text that is more likely to parse as valid JSON.
    """
    # Remove markdown code fences (```json ... ``` or ``` ... ```).
    text = re.sub(r"^```(?:json)?\s*\n?", "", text.strip())
    text = re.sub(r"\n?```\s*$", "", text.strip())

    # If the text doesn't start with '{', try to find the first '{'.
    if not text.startswith("{"):
        brace_pos = text.find("{")
        if brace_pos != -1:
            text = text[brace_pos:]

    return text.strip()


def _make_snippet_id(filename: str, category: str, seq: int) -> str:
    """Build a human-readable, unique snippet identifier.

    The id encodes both the source document and the category the finding
    was coded under, plus a sequence number that disambiguates multiple
    findings from the same document in the same category.

    Format: ``{safe_stem[:40]}__{safe_category[:30]}__{seq:04d}``

    Args:
        filename: Source PDF filename (with or without extension).
        category: Label category this finding belongs to.
        seq: 1-based sequence number within this (document, category).

    Returns:
        A string such as ``"Adelaiye__V_O___2024__Plant__CPVR_EFFECTIVENESS__0001"``.
    """
    stem = Path(filename).stem  # strip ".pdf"
    safe_stem = re.sub(r"[^A-Za-z0-9_-]", "_", stem)[:40].strip("_")
    safe_category = re.sub(r"[^A-Za-z0-9_-]", "_", category)[:30].strip("_")
    return f"{safe_stem}__{safe_category}__{seq:04d}"


# ══════════════════════════════════════════════════════════════════════
# Core parsing
# ══════════════════════════════════════════════════════════════════════


def parse_result_item(
    item: BatchResultItem,
    filename: str,
    category: str,
    valid_codes: Set[str],
) -> Tuple[List[Finding], List[ParseError]]:
    """Parse a single batch result item into findings and errors.

    This is the main entry point for processing one LLM response.  Every
    finding produced is tagged with the given ``category`` and receives a
    freshly-generated ``snippet_id`` and ``timestamp``.

    Args:
        item: A ``BatchResultItem`` from the API client.
        filename: The original PDF filename (for traceability).
        category: Category name that this batch was run under.  Stamped
            onto every finding and error produced.
        valid_codes: Set of valid label codes for this category.

    Returns:
        A tuple of ``(findings, errors)`` where:
            - ``findings`` is a list of valid ``Finding`` objects.
            - ``errors`` is a list of ``ParseError`` objects for any issues.
    """
    findings: List[Finding] = []
    errors: List[ParseError] = []

    # --- Handle API-level failures --------------------------------------
    if not item.success:
        errors.append(
            ParseError(
                filename=filename,
                label_category=category,
                error_type="api_error",
                error_message=item.error_message or "Unknown API error",
            )
        )
        return findings, errors

    # --- Handle missing text --------------------------------------------
    if not item.text:
        errors.append(
            ParseError(
                filename=filename,
                label_category=category,
                error_type="no_text",
                error_message="LLM returned no text content blocks.",
            )
        )
        return findings, errors

    # --- Parse JSON -----------------------------------------------------
    cleaned = _clean_json_text(item.text)
    try:
        parsed = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        errors.append(
            ParseError(
                filename=filename,
                label_category=category,
                error_type="json_parse_failed",
                error_message=str(exc),
                raw_text=item.text,
            )
        )
        return findings, errors

    # --- Extract findings list ------------------------------------------
    raw_findings = parsed.get("findings", [])
    if not isinstance(raw_findings, list):
        errors.append(
            ParseError(
                filename=filename,
                label_category=category,
                error_type="invalid_structure",
                error_message=(
                    "'findings' field is not a list: "
                    f"{type(raw_findings).__name__}"
                ),
                raw_text=item.text,
            )
        )
        return findings, errors

    # --- Validate and build each finding --------------------------------
    seq = 0  # 1-based sequence within (doc, category); incremented on each kept finding
    for idx, raw in enumerate(raw_findings):
        if not isinstance(raw, dict):
            errors.append(
                ParseError(
                    filename=filename,
                    label_category=category,
                    error_type="invalid_finding_item",
                    error_message=f"Finding #{idx} is not a dict: {type(raw).__name__}",
                )
            )
            continue

        label_code = str(raw.get("label_code", "")).strip()
        snippet = str(raw.get("snippet", "")).strip()

        # Required field checks.
        if not label_code or not snippet:
            errors.append(
                ParseError(
                    filename=filename,
                    label_category=category,
                    error_type="missing_fields",
                    error_message=(
                        f"Finding #{idx} missing label_code or snippet."
                    ),
                    raw_text=json.dumps(raw, ensure_ascii=False)[:500],
                )
            )
            continue

        # Validate label code against the (category-specific) framework.
        if label_code not in valid_codes:
            errors.append(
                ParseError(
                    filename=filename,
                    label_category=category,
                    error_type="hallucinated_code",
                    error_message=(
                        f"Finding #{idx}: label_code '{label_code}' is not "
                        f"in category '{category}'. "
                        f"Valid: {sorted(valid_codes)}"
                    ),
                    raw_text=json.dumps(raw, ensure_ascii=False)[:500],
                )
            )
            continue

        # Coerce page number if present.
        page = raw.get("page_number")
        if page is not None:
            try:
                page = int(page)
            except (ValueError, TypeError):
                page = None

        # Assign sequence and build the finding.
        seq += 1
        snippet_id = _make_snippet_id(filename, category, seq)

        finding = Finding(
            filename=filename,
            label_category=category,
            label_code=label_code,
            snippet=snippet,
            page_number=page,
            reasoning=str(raw.get("reasoning", "")).strip(),
            confidence=str(raw.get("confidence", "")).strip().lower(),
            snippet_id=snippet_id,
        )
        findings.append(finding)

    logger.info(
        "'%s' [%s]: parsed %d findings, %d errors",
        filename,
        category,
        len(findings),
        len(errors),
    )
    return findings, errors


def parse_all_results(
    results: List[BatchResultItem],
    custom_id_to_filename: Dict[str, str],
    valid_codes: Set[str],
    category: str,
) -> Tuple[List[Finding], List[ParseError]]:
    """Parse all batch result items for a single category batch.

    Convenience wrapper that calls ``parse_result_item()`` for each
    result and aggregates findings and errors.  All findings produced
    by this call are tagged with the given ``category``.

    Args:
        results: List of ``BatchResultItem`` from the API client.
        custom_id_to_filename: Mapping from ``custom_id`` to filename.
        valid_codes: Set of valid label codes for this category.
        category: Category name for this batch.

    Returns:
        A tuple of ``(all_findings, all_errors)``.
    """
    all_findings: List[Finding] = []
    all_errors: List[ParseError] = []

    for item in results:
        filename = custom_id_to_filename.get(item.custom_id, "unknown")
        item_findings, item_errors = parse_result_item(
            item, filename, category, valid_codes
        )
        all_findings.extend(item_findings)
        all_errors.extend(item_errors)

    logger.info(
        "Category '%s': total %d findings, %d errors from %d results",
        category,
        len(all_findings),
        len(all_errors),
        len(results),
    )
    return all_findings, all_errors
