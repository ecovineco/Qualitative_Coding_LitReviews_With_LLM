"""Configuration module for the AI-powered literature review pipeline.

This module centralises every user-configurable setting so that the rest of
the codebase never contains hard-coded paths, keys, or magic numbers.

Typical usage:
    Edit the values below **once** before running the pipeline.  Every other
    module imports what it needs from here::

        from config import LLM_PROVIDER, TARGET_FOLDER

Environment variables:
    ANTHROPIC_API_KEY        — Required when ``LLM_PROVIDER = "anthropic"``.
    AZURE_OPENAI_API_KEY     — Required when ``LLM_PROVIDER = "azure"``.

API keys are read from the environment so secrets never get committed to
version control.
"""

from __future__ import annotations

import os

# ──────────────────────────────────────────────────────────────────────
# Paths
# ──────────────────────────────────────────────────────────────────────

TARGET_FOLDER: str = "./pdfs"
"""Directory containing documents to analyse (PDF, Word, PowerPoint)."""

LABELS_EXCEL: str = "./labels/coding_framework.xlsx"
"""Path to the Excel file that defines the coding framework (labels)."""

OUTPUT_DIR: str = "./outputs"
"""Directory where all generated outputs are saved."""

# ──────────────────────────────────────────────────────────────────────
# LLM Provider Selection
# ──────────────────────────────────────────────────────────────────────

LLM_PROVIDER: str = "azure"
"""Which LLM provider to use.  Supported values:

    - "anthropic" — Claude via the Anthropic Message Batches API.
    - "azure"     — Azure OpenAI (GPT family) via synchronous chat
                    completions.

To add a new provider, subclass ``BaseLLMClient`` in ``api.py`` and
register it in ``api._PROVIDERS``.
"""

# ──────────────────────────────────────────────────────────────────────
# Anthropic provider settings
# ──────────────────────────────────────────────────────────────────────

ANTHROPIC_API_KEY: str = os.environ.get("ANTHROPIC_API_KEY", "")
"""Anthropic API key.  Set via environment variable for security."""

ANTHROPIC_MODEL: str = "claude-opus-4-6"
"""Anthropic model identifier.  Examples:
    - "claude-sonnet-4-20250514"  (fast, cost-effective)
    - "claude-opus-4-6"           (most capable, slower)
"""

ANTHROPIC_MAX_TOKENS: int = 8192
"""Maximum tokens for each Claude response."""

ANTHROPIC_VERSION: str = "2023-06-01"
"""Anthropic API version header value."""

ANTHROPIC_FILES_BETA: str = "files-api-2025-04-14"
"""Beta feature header required for the Files API."""

ANTHROPIC_BASE_URL: str = "https://api.anthropic.com"
"""Base URL for the Anthropic API."""

# ──────────────────────────────────────────────────────────────────────
# Azure OpenAI provider settings
# ──────────────────────────────────────────────────────────────────────

AZURE_OPENAI_API_KEY: str = os.environ.get("AZURE_OPENAI_API_KEY", "")
"""Azure OpenAI API key.  Set via environment variable for security."""

AZURE_OPENAI_ENDPOINT: str = (
    "https://foundry-be-1001215-114.cognitiveservices.azure.com/"
)
"""Azure OpenAI resource endpoint, including trailing slash."""

AZURE_OPENAI_API_VERSION: str = "2025-03-01-preview"
"""Azure OpenAI API version string."""

AZURE_OPENAI_MODEL: str = "gpt-5.4"
"""Azure OpenAI deployment name (not the underlying model family)."""

# ──────────────────────────────────────────────────────────────────────
# Backwards-compatible alias
# ──────────────────────────────────────────────────────────────────────

# Some older code paths may still import ``API_KEY``.  Resolve it to the
# active provider's key so nothing breaks if the symbol is referenced.
API_KEY: str = (
    AZURE_OPENAI_API_KEY if LLM_PROVIDER == "azure" else ANTHROPIC_API_KEY
)

# ──────────────────────────────────────────────────────────────────────
# Batch Processing
# ──────────────────────────────────────────────────────────────────────

POLL_INTERVAL_SECONDS: int = 30
"""How often (in seconds) to check whether a batch has finished.

Only relevant for providers with asynchronous batches (e.g. Anthropic).
"""

MAX_PDFS: int | None = None
"""Limit the number of PDFs processed (useful for testing).

Set to ``None`` to process all PDFs in ``TARGET_FOLDER``.
"""

REQUEST_TIMEOUT: int = 300
"""HTTP request timeout in seconds for API calls."""

# ──────────────────────────────────────────────────────────────────────
# Pipeline Behaviour
# ──────────────────────────────────────────────────────────────────────

DRY_RUN: bool = False
"""If ``True``, build batch requests and print them without submitting.

Useful for verifying the prompt and request structure before spending
API credits.
"""

INCLUDE_CONFIDENCE: bool = True
"""If ``True``, ask the LLM to return a confidence level for each finding."""

UPLOAD_RETRIES: int = 3
"""Number of retry attempts for failed PDF uploads."""

UPLOAD_RETRY_DELAY: int = 5
"""Seconds to wait between upload retry attempts."""

# ──────────────────────────────────────────────────────────────────────
# Snippet verification
# ──────────────────────────────────────────────────────────────────────

# After coding, every snippet is checked against its source PDF to
# confirm it is a real verbatim quote (see ``verify.py``).  This stage
# always runs at the end of the pipeline; it needs no API credits.

PDF_SOURCE_DIR: str = TARGET_FOLDER
"""Directory holding the original PDFs, used to verify snippets.

Defaults to ``TARGET_FOLDER`` — the same folder the documents were read
from.  (For the Anthropic provider the stored ``file_id`` is not a local
path, so verification always re-resolves PDFs from this folder by
filename.)
"""

VERIFY_OUTPUT: str = "./outputs/verification.xlsx"
"""Where the snippet-verification results are written.

The file joins back to ``coded_findings.xlsx`` on ``snippet_id`` (and on
``finding_hash`` for cross-run stability).
"""

VERIFY_FUZZY_THRESHOLD: int = 95
"""Fuzzy score (0-100) at or above which a snippet counts as verified.

Matches in ``[VERIFY_NEAR_THRESHOLD, VERIFY_FUZZY_THRESHOLD)`` are
flagged as ``near_match`` for human review rather than auto-verified.
"""

VERIFY_NEAR_THRESHOLD: int = 85
"""Fuzzy score below which a snippet is treated as ``not_found``."""

VERIFY_MIN_LEN_FOR_FUZZY: int = 40
"""Minimum normalised snippet length (chars) to allow a *fuzzy* verify.

Short snippets can reach a high fuzzy score by chance, so a fuzzy match
on a shorter snippet is demoted to ``near_match`` instead of
``verified_fuzzy``.
"""

VERIFY_PAGE_TOLERANCE: int = 1
"""Allowed difference between the claimed and matched page for ``page_ok``.

A soft signal only: models usually report the *printed* page number,
which is offset from the physical PDF page by front matter.
"""

VERIFY_NO_TEXT_LAYER_CHARS: int = 50
"""Below this many extracted characters, a PDF is treated as a scan.

Such documents have no usable text layer and yield ``no_text_layer``
(rather than falsely flagging every snippet as fabricated).
"""

VERIFY_STRIP_BOILERPLATE: bool = True
"""Strip recurring page headers / footers before matching.

PDF text extraction often injects a running header (author, journal,
page number) into the middle of a sentence that spans a page break,
which would otherwise break the match for a genuine verbatim quote.
When enabled, text blocks that sit in the page margins *and* repeat
across pages are removed before snippets are matched.
"""

VERIFY_MARGIN_FRACTION: float = 0.12
"""Fraction of page height treated as the top / bottom margin.

Only blocks within these margins are candidates for header / footer
removal, so repeated body text is never stripped.
"""
