"""Snippet verification stage for the literature review pipeline.

After the LLM has coded a corpus, every ``Finding`` carries a ``snippet``
that is supposed to be a *verbatim* quote from the source PDF.  This
module checks that claim: for each finding it extracts the source
document's own text and confirms that the snippet really appears in it.

This is the pipeline's only safeguard against fabricated quotes.  A
snippet that cannot be located in the source is the strongest available
signal that the model invented or distorted it.

Why a simple ``snippet in document_text`` test is not enough
------------------------------------------------------------
PDF text extraction is lossy and models silently tidy quotes, so a naive
substring test reports huge numbers of *false* "not found" results on
quotes that are genuinely present.  The usual culprits are:

    * Line-break hyphenation: ``inno-\\nvation`` in the PDF vs
      ``innovation`` in the snippet.
    * Ligatures: ``ﬁ`` / ``ﬂ`` extracted as single code points.
    * Smart punctuation: curly quotes, en/em dashes, ellipsis glyphs.
    * Invisible characters: non-breaking and zero-width spaces.
    * Whitespace: newlines and runs of spaces from column / justified
      layout where the snippet uses single spaces.
    * Model edits: dropped footnote markers, and elision where a snippet
      stitches two non-adjacent spans together with ``...``.

To handle these, both the snippet and the document text are run through
the *same* aggressive normalisation pipeline (see :func:`normalize`) and
then compared with a graded **match ladder** (see :func:`verify_one`):

    1. Exact match on the normalised text.            -> ``verified``
    2. Fragmented match for elided ``...`` quotes.     -> ``verified_fragmented``
    3. Fuzzy match (``rapidfuzz``) for minor edits.    -> ``verified_fuzzy`` / ``near_match``
    4. Nothing matched.                                -> ``not_found``

Two further statuses describe situations where verification could not be
attempted: ``no_text_layer`` (the PDF is a scan with no extractable
text) and ``pdf_missing`` (the source file could not be found on disk).

Typical usage
-------------
As a library, on the in-memory findings produced by the main pipeline::

    from verify import verify_findings
    results = verify_findings(findings, pdf_dir="./pdfs")

As a standalone CLI, re-running verification against an existing
``coded_findings.xlsx`` without spending any API credits::

    python verify.py
    python verify.py --findings outputs/coded_findings.xlsx --pdfs pdfs

Dependencies:
    PyMuPDF (``pip install PyMuPDF``) for PDF text extraction and
    ``rapidfuzz`` (``pip install rapidfuzz``) for fuzzy matching.  Both
    are listed in ``requirements.txt``.
"""

from __future__ import annotations

import argparse
import bisect
import logging
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

import config

logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════
# Status constants
# ══════════════════════════════════════════════════════════════════════

STATUS_VERIFIED = "verified"
"""Snippet found by exact match on the normalised document text."""

STATUS_VERIFIED_FRAGMENTED = "verified_fragmented"
"""All fragments of an elided (``...``) snippet were found."""

STATUS_VERIFIED_FUZZY = "verified_fuzzy"
"""Snippet matched above the fuzzy threshold (minor edits / extraction noise)."""

STATUS_NEAR_MATCH = "near_match"
"""Matched in the grey band — likely present but flagged for human review."""

STATUS_NOT_FOUND = "not_found"
"""Snippet could not be located — strongest signal of a fabricated quote."""

STATUS_NO_TEXT_LAYER = "no_text_layer"
"""The PDF has no extractable text (scanned image); cannot verify."""

STATUS_PDF_MISSING = "pdf_missing"
"""The source PDF could not be found on disk; cannot verify."""

#: Statuses that count as a successful verification of the snippet.
VERIFIED_STATUSES = frozenset(
    {STATUS_VERIFIED, STATUS_VERIFIED_FRAGMENTED, STATUS_VERIFIED_FUZZY}
)


# ══════════════════════════════════════════════════════════════════════
# Data structures
# ══════════════════════════════════════════════════════════════════════


@dataclass
class FindingRecord:
    """The minimal subset of a finding needed to verify its snippet.

    Using a lightweight record (rather than the full ``Finding`` from
    ``parser.py``) lets verification run both inline (fed ``Finding``
    objects by the pipeline) and standalone (built from the rows of an
    existing ``coded_findings.xlsx``) through the same code path.

    Attributes:
        snippet_id: Human-readable identifier of the finding; used to
            join verification results back to the findings table.
        finding_hash: Stable dedup hash of the finding (secondary join
            key, stable across runs).
        filename: Source PDF filename (used to locate the PDF on disk).
        page_number: Page the model claimed the snippet came from, if
            any.  May be ``None``.
        snippet: The verbatim quote to verify.
    """

    snippet_id: str
    finding_hash: str
    filename: str
    page_number: Optional[int]
    snippet: str


@dataclass
class VerificationResult:
    """The outcome of verifying a single snippet against its source.

    Attributes:
        snippet_id: Join key back to the findings table.
        finding_hash: Secondary, run-stable join key.
        filename: Source PDF filename.
        page_number: Page the model claimed (echoed from the finding).
        verification_status: One of the ``STATUS_*`` constants.
        match_score: Best similarity score on a 0-100 scale.  ``100.0``
            for exact / fragmented matches, the fuzzy score otherwise,
            and the best-effort partial score for ``not_found`` rows (so
            a near miss is distinguishable from total absence).
        match_method: Which rung of the ladder produced the result —
            ``"exact"``, ``"fragmented"``, ``"fuzzy"`` or ``"none"``.
        matched_page: The 1-based *physical* PDF page where the snippet
            was located, or ``None``.  May differ from ``page_number``
            because models usually report the *printed* page number,
            which is offset by cover pages and front matter.
        page_ok: ``True`` if ``matched_page`` is within
            ``config.VERIFY_PAGE_TOLERANCE`` of the claimed page,
            ``False`` if not, ``None`` if either page is unknown.
        matched_text: For non-exact matches, the slice of normalised
            document text that matched — kept so a human can adjudicate
            ``verified_fuzzy`` / ``near_match`` rows.  Empty for exact
            and fragmented matches.
    """

    snippet_id: str
    finding_hash: str
    filename: str
    page_number: Optional[int]
    verification_status: str
    match_score: float = 0.0
    match_method: str = "none"
    matched_page: Optional[int] = None
    page_ok: Optional[bool] = None
    matched_text: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """Flatten to a dict for export.

        Returns:
            A dictionary with one entry per output column.  ``matched_text``
            is truncated to keep Excel cells well within their length
            limit.
        """
        return {
            "snippet_id": self.snippet_id,
            "finding_hash": self.finding_hash,
            "filename": self.filename,
            "page_number": self.page_number,
            "verification_status": self.verification_status,
            "match_score": round(self.match_score, 1),
            "match_method": self.match_method,
            "matched_page": self.matched_page,
            "page_ok": self.page_ok,
            "matched_text": self.matched_text[:2000],
        }


@dataclass
class DocText:
    """Cached, normalised text of one source PDF.

    Built once per document and reused across all of that document's
    findings (a corpus typically has many snippets per file).

    Attributes:
        filename: Source PDF filename.
        has_text_layer: ``False`` if the PDF yielded almost no text
            (i.e. it is a scanned image and cannot be verified).
        full_cf: Whole document as a single normalised, case-folded
            string (pages joined by a space) — the haystack searched by
            the match ladder.  Joining pages means a snippet spanning a
            page break still matches.
        full_display: Same text, but case-preserved, so ``matched_text``
            slices read naturally.
        page_offsets: Start index of each page within ``full_cf``; used
            to map a match offset back to a page number.
    """

    filename: str
    has_text_layer: bool
    full_cf: str = ""
    full_display: str = ""
    page_offsets: List[int] = field(default_factory=list)

    def page_at(self, offset: int) -> Optional[int]:
        """Return the 1-based physical page containing a character offset.

        Args:
            offset: Index into ``full_cf`` (or ``full_display``).

        Returns:
            The 1-based page number, or ``None`` if there are no pages.
        """
        if not self.page_offsets:
            return None
        # bisect_right - 1 gives the index of the page whose start offset
        # is the greatest one that is <= ``offset``.
        return bisect.bisect_right(self.page_offsets, offset)


# ══════════════════════════════════════════════════════════════════════
# Text normalisation
# ══════════════════════════════════════════════════════════════════════

# Map of characters that look identical to a human but break exact
# matching.  Applied *after* NFKC (which already folds ligatures and many
# compatibility forms but leaves curly quotes and dashes untouched) and
# *before* whitespace is collapsed (de-hyphenation still needs newlines).
_CHAR_MAP = {
    # Single quotes / apostrophes -> straight apostrophe.
    "\u2018": "'", "\u2019": "'", "\u201a": "'", "\u201b": "'", "\u2032": "'",
    # Double quotes -> straight double quote.
    "\u201c": '"', "\u201d": '"', "\u201e": '"', "\u201f": '"', "\u2033": '"',
    # Dash family -> hyphen-minus.
    "\u2010": "-", "\u2011": "-", "\u2012": "-", "\u2013": "-",
    "\u2014": "-", "\u2015": "-", "\u2212": "-",
    # Ellipsis -> three dots (so it matches a model that typed "...").
    "\u2026": "...",
    # Spaces that are not a plain space -> plain space.
    "\u00a0": " ", "\u2007": " ", "\u202f": " ", "\u2009": " ", "\u200a": " ",
    # Zero-width and invisible marks -> removed.
    "\u200b": "", "\u200c": "", "\u200d": "", "\ufeff": "",
    # Soft hyphen (invisible hyphenation hint) -> removed.
    "\u00ad": "",
}
_TRANSLATION = str.maketrans(_CHAR_MAP)

# Join a word split across a line break: "inno-\nvation" -> "innovation".
# Only fires on hyphen + (optional spaces) + newline + (optional spaces),
# so genuine hyphenated compounds like "well-known" are left intact.
_DEHYPHEN_RE = re.compile(r"(\w)-[ \t]*\n[ \t]*(\w)")

_WHITESPACE_RE = re.compile(r"\s+")

# Markers a model uses to indicate elision inside a quote.
_ELLIPSIS_SPLIT_RE = re.compile(r"\s*(?:\[\s*\.\.\.\s*\]|\.\s*\.\s*\.)\s*")


def normalize(text: str, casefold: bool = False) -> str:
    """Normalise text so that visually-equal strings compare as equal.

    The same function is applied to both the snippet and the document so
    that differences introduced purely by PDF extraction or by harmless
    model tidying do not cause a true quote to be reported as missing.

    The steps, in order (order matters — de-hyphenation needs the
    newlines that the final whitespace collapse removes):

        1. Unicode NFKC normalisation (folds ligatures, compatibility
           forms, and many width variants).
        2. Character substitution (curly quotes, dashes, ellipsis,
           exotic and zero-width spaces) via :data:`_CHAR_MAP`.
        3. De-hyphenation of words broken across a line break.
        4. Collapse every run of whitespace to a single space and strip.
        5. Optionally case-fold for case-insensitive comparison.

    Args:
        text: The raw text to normalise.
        casefold: If ``True``, apply ``str.casefold()`` as the last step.

    Returns:
        The normalised string (empty if ``text`` was empty / whitespace).
    """
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.translate(_TRANSLATION)
    text = _DEHYPHEN_RE.sub(r"\1\2", text)
    text = _WHITESPACE_RE.sub(" ", text).strip()
    if casefold:
        text = text.casefold()
    return text


# ══════════════════════════════════════════════════════════════════════
# PDF text extraction (cached per document)
# ══════════════════════════════════════════════════════════════════════


def _import_fitz():
    """Import PyMuPDF, raising a clear, actionable error if it is absent.

    Returns:
        The imported ``fitz`` module.

    Raises:
        RuntimeError: If PyMuPDF is not installed.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:  # pragma: no cover - exercised via install state
        raise RuntimeError(
            "PyMuPDF is required for snippet verification. "
            "Install it with: pip install PyMuPDF"
        ) from exc
    return fitz


# Collapse digit runs so a running header's varying page number does not
# stop it from being recognised as the same line across pages.
_DIGIT_RUN_RE = re.compile(r"\d+")


def _boilerplate_signature(text: str) -> str:
    """Build a page-number-insensitive signature for a margin block.

    Two header lines that differ only by their page number (``... 458`` vs
    ``... 459``) collapse to the same signature, so the line is recognised
    as recurring boilerplate.

    Args:
        text: Raw text of a margin block.

    Returns:
        A normalised, case-folded, digit-collapsed signature (may be empty).
    """
    return _DIGIT_RUN_RE.sub("#", normalize(text, casefold=True)).strip()


def _detect_boilerplate(
    page_blocks: List[List[tuple]], heights: List[float]
) -> set:
    """Find header / footer signatures that recur across pages.

    A signature is treated as boilerplate when it appears, inside a page
    margin, on at least ``max(2, round(0.25 * n_pages))`` distinct pages.
    Restricting candidates to the margins means repeated *body* text is
    never removed.

    Args:
        page_blocks: Per page, a list of ``(text, y0, y1)`` block tuples.
        heights: Per page, the page height in points.

    Returns:
        The set of boilerplate signatures to strip.
    """
    from collections import defaultdict

    n = len(page_blocks)
    if not config.VERIFY_STRIP_BOILERPLATE or n < 2:
        return set()

    sig_pages: Dict[str, set] = defaultdict(set)
    for pno, blocks in enumerate(page_blocks):
        top = config.VERIFY_MARGIN_FRACTION * heights[pno]
        bottom = (1.0 - config.VERIFY_MARGIN_FRACTION) * heights[pno]
        for text, y0, y1 in blocks:
            if y1 <= top or y0 >= bottom:  # block lies in a margin
                sig = _boilerplate_signature(text)
                if sig:
                    sig_pages[sig].add(pno)

    threshold = max(2, round(0.25 * n))
    return {sig for sig, pages in sig_pages.items() if len(pages) >= threshold}


def extract_doc_text(pdf_path: str | Path, filename: str) -> DocText:
    """Extract and normalise the full text of a PDF, page by page.

    Recurring page headers and footers are detected and removed before
    normalisation (see :func:`_detect_boilerplate`).  This matters
    because PDF extraction routinely splices a running header into the
    middle of a sentence that spans a page break; without removal, a
    genuine verbatim quote crossing that break would fail to match.

    The result is meant to be cached and reused for every finding that
    came from this document.

    Args:
        pdf_path: Path to the PDF file on disk.
        filename: The finding's filename, stored on the result so it can
            be matched back even if ``pdf_path`` differs in casing/dir.

    Returns:
        A :class:`DocText`.  ``has_text_layer`` is ``False`` when the
        document yields fewer than ``config.VERIFY_NO_TEXT_LAYER_CHARS``
        characters of normalised text (a scanned image PDF).
    """
    fitz = _import_fitz()
    doc = fitz.open(str(pdf_path))
    try:
        # Per page, collect text blocks with their vertical position.  Each
        # block from PyMuPDF is (x0, y0, x1, y1, text, block_no, block_type);
        # block_type 1 is an image, which we skip.
        page_blocks: List[List[tuple]] = []
        heights: List[float] = []
        for page in doc:
            heights.append(page.rect.height)
            blocks = []
            for b in page.get_text("blocks"):
                if len(b) >= 7 and b[6] != 0:
                    continue  # image block
                text = b[4]
                if text and text.strip():
                    blocks.append((text, b[1], b[3]))
            page_blocks.append(blocks)
    finally:
        doc.close()

    boilerplate = _detect_boilerplate(page_blocks, heights)

    # Build each page's text from its non-boilerplate blocks.
    page_norms_cf: List[str] = []
    page_norms_display: List[str] = []
    for pno, blocks in enumerate(page_blocks):
        top = config.VERIFY_MARGIN_FRACTION * heights[pno]
        bottom = (1.0 - config.VERIFY_MARGIN_FRACTION) * heights[pno]
        kept: List[str] = []
        for text, y0, y1 in blocks:
            in_margin = y1 <= top or y0 >= bottom
            if in_margin and _boilerplate_signature(text) in boilerplate:
                continue
            kept.append(text)
        page_text = " ".join(kept)
        page_norms_display.append(normalize(page_text, casefold=False))
        page_norms_cf.append(normalize(page_text, casefold=True))

    # Build the full-document haystack and record each page's start offset
    # so a match index can be mapped back to a page.  Pages are joined by
    # a single space, matching the inter-token spacing produced by
    # ``normalize`` so a snippet crossing a page boundary still matches.
    offsets: List[int] = []
    cursor = 0
    for pcf in page_norms_cf:
        offsets.append(cursor)
        cursor += len(pcf) + 1  # +1 for the joining space
    full_cf = " ".join(page_norms_cf)
    full_display = " ".join(page_norms_display)

    has_text = len(full_cf.strip()) >= config.VERIFY_NO_TEXT_LAYER_CHARS
    if not has_text:
        logger.warning(
            "No usable text layer in '%s' (%d chars) — likely a scan.",
            filename,
            len(full_cf.strip()),
        )

    return DocText(
        filename=filename,
        has_text_layer=has_text,
        full_cf=full_cf,
        full_display=full_display,
        page_offsets=offsets,
    )


# ══════════════════════════════════════════════════════════════════════
# The match ladder
# ══════════════════════════════════════════════════════════════════════


def _fuzzy_partial(needle_cf: str, haystack_cf: str):
    """Best fuzzy alignment of ``needle_cf`` inside ``haystack_cf``.

    Wraps ``rapidfuzz.fuzz.partial_ratio_alignment``, which finds the
    best-matching window of the long string and returns both the score
    and the window's bounds in a single optimised call.

    Args:
        needle_cf: Normalised, case-folded snippet.
        haystack_cf: Normalised, case-folded document text.

    Returns:
        A ``ScoreAlignment`` (with ``.score``, ``.dest_start``,
        ``.dest_end``).

    Raises:
        RuntimeError: If ``rapidfuzz`` is not installed.
    """
    try:
        from rapidfuzz import fuzz
    except ImportError as exc:  # pragma: no cover - exercised via install state
        raise RuntimeError(
            "rapidfuzz is required for snippet verification. "
            "Install it with: pip install rapidfuzz"
        ) from exc
    return fuzz.partial_ratio_alignment(needle_cf, haystack_cf)


def verify_one(record: FindingRecord, doc: DocText) -> VerificationResult:
    """Run the match ladder for one snippet against its source document.

    The ladder tries the cheapest, most certain test first and falls
    through to progressively more forgiving (and lower-confidence) ones,
    recording which rung succeeded:

        1. **Exact** normalised substring  -> ``verified``.
        2. **Fragmented**: if the snippet contains ``...`` markers, every
           fragment must be present (searched in order)
           -> ``verified_fragmented``.
        3. **Fuzzy** partial match: score ``>= VERIFY_FUZZY_THRESHOLD``
           (and snippet long enough) -> ``verified_fuzzy``; score in
           ``[VERIFY_NEAR_THRESHOLD, VERIFY_FUZZY_THRESHOLD)`` ->
           ``near_match`` (flagged for review).
        4. Otherwise -> ``not_found``.

    A short snippet can hit a high fuzzy score by coincidence, so a fuzzy
    pass is only promoted to ``verified_fuzzy`` when the normalised
    snippet is at least ``config.VERIFY_MIN_LEN_FOR_FUZZY`` characters
    long; shorter fuzzy hits are demoted to ``near_match`` for a human to
    confirm.

    Args:
        record: The finding to verify.
        doc: The cached, normalised text of its source document.

    Returns:
        A populated :class:`VerificationResult`.
    """
    base = dict(
        snippet_id=record.snippet_id,
        finding_hash=record.finding_hash,
        filename=record.filename,
        page_number=record.page_number,
    )

    if not doc.has_text_layer:
        return VerificationResult(**base, verification_status=STATUS_NO_TEXT_LAYER)

    snippet_cf = normalize(record.snippet, casefold=True)
    if not snippet_cf:
        return VerificationResult(
            **base, verification_status=STATUS_NOT_FOUND, match_method="none"
        )

    # --- Rung 1: exact normalised substring -----------------------------
    idx = doc.full_cf.find(snippet_cf)
    if idx != -1:
        page = doc.page_at(idx)
        return VerificationResult(
            **base,
            verification_status=STATUS_VERIFIED,
            match_score=100.0,
            match_method="exact",
            matched_page=page,
            page_ok=_page_ok(record.page_number, page),
        )

    # --- Rung 2: fragmented (elided) snippet ----------------------------
    fragments = [
        f for f in _ELLIPSIS_SPLIT_RE.split(record.snippet) if f.strip()
    ]
    if len(fragments) > 1:
        positions: List[int] = []
        cursor = 0
        all_found = True
        for frag in fragments:
            frag_cf = normalize(frag, casefold=True)
            if len(frag_cf) < 3:  # ignore trivially short fragments
                continue
            pos = doc.full_cf.find(frag_cf, cursor)
            if pos == -1:  # not after the previous fragment — try anywhere
                pos = doc.full_cf.find(frag_cf)
            if pos == -1:
                all_found = False
                break
            positions.append(pos)
            cursor = pos + len(frag_cf)
        if all_found and positions:
            page = doc.page_at(min(positions))
            return VerificationResult(
                **base,
                verification_status=STATUS_VERIFIED_FRAGMENTED,
                match_score=100.0,
                match_method="fragmented",
                matched_page=page,
                page_ok=_page_ok(record.page_number, page),
            )

    # --- Rung 3: fuzzy partial match ------------------------------------
    alignment = _fuzzy_partial(snippet_cf, doc.full_cf)
    score = float(alignment.score)
    if score >= config.VERIFY_NEAR_THRESHOLD:
        page = doc.page_at(alignment.dest_start)
        window = doc.full_display[alignment.dest_start : alignment.dest_end]
        long_enough = len(snippet_cf) >= config.VERIFY_MIN_LEN_FOR_FUZZY
        status = (
            STATUS_VERIFIED_FUZZY
            if (score >= config.VERIFY_FUZZY_THRESHOLD and long_enough)
            else STATUS_NEAR_MATCH
        )
        return VerificationResult(
            **base,
            verification_status=status,
            match_score=score,
            match_method="fuzzy",
            matched_page=page,
            page_ok=_page_ok(record.page_number, page),
            matched_text=window,
        )

    # --- Rung 4: not found (report best score for diagnostics) ----------
    return VerificationResult(
        **base,
        verification_status=STATUS_NOT_FOUND,
        match_score=score,
        match_method="none",
    )


def _page_ok(claimed: Optional[int], matched: Optional[int]) -> Optional[bool]:
    """Whether the matched page is close enough to the claimed page.

    Treated as a *soft* signal: models usually report the printed page
    number, which is offset from the physical PDF page index by cover
    pages and front matter, so a mismatch is a warning, not a failure.

    Args:
        claimed: Page the model reported (may be ``None``).
        matched: Physical page the snippet was found on (may be ``None``).

    Returns:
        ``True`` / ``False`` within tolerance, or ``None`` if either page
        is unknown.
    """
    if claimed is None or matched is None:
        return None
    return abs(claimed - matched) <= config.VERIFY_PAGE_TOLERANCE


# ══════════════════════════════════════════════════════════════════════
# Orchestration
# ══════════════════════════════════════════════════════════════════════


def verify_findings(
    findings: Sequence[Any],
    pdf_dir: str | Path = "",
) -> List[VerificationResult]:
    """Verify the snippet of every finding against its source PDF.

    Documents are extracted once and cached, so the cost is one PDF read
    per *document*, not per finding.

    Args:
        findings: A sequence of finding-like objects.  Each must expose
            ``snippet_id``, ``finding_hash``, ``filename``, ``snippet``
            and ``page_number`` — both the ``Finding`` dataclass from
            ``parser.py`` and :class:`FindingRecord` qualify.
        pdf_dir: Directory holding the original PDFs.  Defaults to
            ``config.PDF_SOURCE_DIR``.

    Returns:
        A list of :class:`VerificationResult`, one per finding, in the
        same order as the input.
    """
    pdf_dir = Path(pdf_dir or config.PDF_SOURCE_DIR)
    records = [_as_record(f) for f in findings]

    doc_cache: Dict[str, DocText] = {}
    results: List[VerificationResult] = []

    for rec in records:
        doc = doc_cache.get(rec.filename)
        if doc is None:
            pdf_path = pdf_dir / rec.filename
            if not pdf_path.exists():
                logger.warning(
                    "Source PDF not found for verification: %s", pdf_path
                )
                doc = DocText(filename=rec.filename, has_text_layer=False)
                # Mark with a sentinel so we can distinguish "missing file"
                # from "scanned doc" below.
                doc.full_cf = "__PDF_MISSING__"
            else:
                logger.info("Extracting text for verification: %s", rec.filename)
                doc = extract_doc_text(pdf_path, rec.filename)
            doc_cache[rec.filename] = doc

        if doc.full_cf == "__PDF_MISSING__":
            results.append(
                VerificationResult(
                    snippet_id=rec.snippet_id,
                    finding_hash=rec.finding_hash,
                    filename=rec.filename,
                    page_number=rec.page_number,
                    verification_status=STATUS_PDF_MISSING,
                )
            )
            continue

        results.append(verify_one(rec, doc))

    _log_summary(results)
    return results


def _as_record(obj: Any) -> FindingRecord:
    """Coerce a finding-like object into a :class:`FindingRecord`.

    Args:
        obj: A ``Finding`` (or any object exposing the same attributes).

    Returns:
        A :class:`FindingRecord` with the fields verification needs.
    """
    if isinstance(obj, FindingRecord):
        return obj
    return FindingRecord(
        snippet_id=getattr(obj, "snippet_id", ""),
        finding_hash=getattr(obj, "finding_hash", ""),
        filename=getattr(obj, "filename", ""),
        page_number=getattr(obj, "page_number", None),
        snippet=getattr(obj, "snippet", ""),
    )


def summarize(results: Sequence[VerificationResult]) -> Dict[str, int]:
    """Count results by status.

    Args:
        results: The verification results to tally.

    Returns:
        A dict mapping each ``STATUS_*`` value to its count.
    """
    counts: Dict[str, int] = {}
    for r in results:
        counts[r.verification_status] = counts.get(r.verification_status, 0) + 1
    return counts


def _log_summary(results: Sequence[VerificationResult]) -> None:
    """Log a one-line summary of verification outcomes."""
    counts = summarize(results)
    verified = sum(counts.get(s, 0) for s in VERIFIED_STATUSES)
    logger.info(
        "Verification: %d/%d snippets verified. Breakdown: %s",
        verified,
        len(results),
        {k: counts[k] for k in sorted(counts)},
    )


# ══════════════════════════════════════════════════════════════════════
# Standalone entry point
# ══════════════════════════════════════════════════════════════════════


def _records_from_excel(findings_path: str | Path) -> List[FindingRecord]:
    """Build finding records from an existing ``coded_findings.xlsx``.

    Lets the verifier run standalone — re-checking a previous run's
    output without re-invoking the LLM.

    Args:
        findings_path: Path to a ``coded_findings.xlsx`` produced by the
            pipeline.

    Returns:
        A list of :class:`FindingRecord`.

    Raises:
        FileNotFoundError: If the findings file does not exist.
    """
    import pandas as pd

    path = Path(findings_path)
    if not path.exists():
        raise FileNotFoundError(f"Findings file not found: {path}")

    df = pd.read_excel(str(path), dtype=str).fillna("")
    records: List[FindingRecord] = []
    for _, row in df.iterrows():
        page_raw = str(row.get("page_number", "")).strip()
        try:
            page = int(float(page_raw)) if page_raw else None
        except (ValueError, TypeError):
            page = None
        records.append(
            FindingRecord(
                snippet_id=row.get("snippet_id", ""),
                finding_hash=row.get("finding_hash", ""),
                filename=row.get("filename", ""),
                page_number=page,
                snippet=row.get("snippet", ""),
            )
        )
    logger.info("Loaded %d findings from %s", len(records), path)
    return records


def run_verification_cli(
    findings_path: str | Path,
    pdf_dir: str | Path,
    output_path: str | Path,
) -> List[VerificationResult]:
    """Verify a findings Excel file and write the full audit file.

    This is the standalone path; the inline pipeline calls
    :func:`verify_findings` directly on its in-memory findings instead.
    It reads every row of ``findings_path``, verifies each snippet, and
    writes ``output_path`` (``coded_findings_verified.xlsx`` by default)
    containing the original rows extended with the verification columns.
    The input file is never modified, so this is safe to re-run while
    tuning thresholds.

    Args:
        findings_path: Path to a findings Excel file (e.g.
            ``coded_findings.xlsx``).
        pdf_dir: Directory holding the source PDFs.
        output_path: Where to write the verified audit Excel file.

    Returns:
        The list of :class:`VerificationResult`.
    """
    import pandas as pd

    # The verification-specific columns to append (single source of truth
    # is ``export``; imported lazily to keep the dependency one-directional
    # — ``export`` never imports ``verify``).
    from export import _VERIFICATION_EXTRA_COLUMNS

    df = pd.read_excel(str(findings_path), dtype=str).fillna("")
    records = _records_from_excel(findings_path)
    results = verify_findings(records, pdf_dir=pdf_dir)

    verdicts = {r.snippet_id: r.to_dict() for r in results}
    sids = df["snippet_id"] if "snippet_id" in df.columns else [""] * len(df)
    for col in _VERIFICATION_EXTRA_COLUMNS:
        df[col] = [verdicts.get(sid, {}).get(col) for sid in sids]

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(str(output_path), index=False)
    logger.info("Saved %d verified findings to %s", len(df), output_path)
    return results


def main() -> None:
    """Parse CLI arguments and run standalone verification."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    parser = argparse.ArgumentParser(
        description=(
            "Verify that each coded snippet appears verbatim in its "
            "source PDF."
        )
    )
    parser.add_argument(
        "--findings",
        default=str(Path(config.OUTPUT_DIR) / "coded_findings.xlsx"),
        help="Path to coded_findings.xlsx (default: outputs/coded_findings.xlsx).",
    )
    parser.add_argument(
        "--pdfs",
        default=config.PDF_SOURCE_DIR,
        help="Directory containing the source PDFs (default: config.PDF_SOURCE_DIR).",
    )
    parser.add_argument(
        "--out",
        default=config.VERIFIED_OUTPUT,
        help=(
            "Where to write the verified audit file "
            "(default: config.VERIFIED_OUTPUT, coded_findings_verified.xlsx)."
        ),
    )
    args = parser.parse_args()
    run_verification_cli(args.findings, args.pdfs, args.out)


if __name__ == "__main__":
    main()
