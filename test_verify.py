"""Fixture tests for the snippet verification stage (``verify.py``).

These tests do not touch any PDF or API.  They build a synthetic
:class:`~verify.DocText` from a known passage and assert that each
known extraction / model-editing failure mode lands in the correct
bucket of the match ladder.  This is what turns "I think it is robust"
into "every known failure mode is covered".

Run with::

    python test_verify.py        # plain assert-based, no pytest needed
    pytest test_verify.py         # also works under pytest
"""

from __future__ import annotations

import verify
from verify import (
    DocText,
    FindingRecord,
    STATUS_NEAR_MATCH,
    STATUS_NOT_FOUND,
    STATUS_VERIFIED,
    STATUS_VERIFIED_FRAGMENTED,
    STATUS_VERIFIED_FUZZY,
    normalize,
    verify_one,
)

# A short, two-"page" source passage with realistic PDF noise: a curly
# apostrophe, an em dash, a ligature, and a hyphenated line break.
PAGE_1 = (
    "The breeder\u2019s rights regime promotes innovation \u2014 "
    "and protects new plant varieties \ufb01rmly across the inno-\n"
    "vation pipeline."
)
PAGE_2 = "A second page discusses compulsory licensing and farmers rights."


def _make_doc() -> DocText:
    """Build a two-page :class:`DocText` the way ``extract_doc_text`` would."""
    pages_cf = [normalize(PAGE_1, casefold=True), normalize(PAGE_2, casefold=True)]
    pages_disp = [normalize(PAGE_1), normalize(PAGE_2)]
    offsets, cursor = [], 0
    for p in pages_cf:
        offsets.append(cursor)
        cursor += len(p) + 1
    return DocText(
        filename="sample.pdf",
        has_text_layer=True,
        full_cf=" ".join(pages_cf),
        full_display=" ".join(pages_disp),
        page_offsets=offsets,
    )


def _verify(snippet: str, page: int | None = 1) -> verify.VerificationResult:
    rec = FindingRecord(
        snippet_id="t",
        finding_hash="h",
        filename="sample.pdf",
        page_number=page,
        snippet=snippet,
    )
    return verify_one(rec, _make_doc())


def run() -> None:
    """Execute all fixture checks; raises ``AssertionError`` on failure."""

    # 1. normalize() folds the noise so visually-equal strings match.
    assert normalize("inno-\nvation") == "innovation", "de-hyphenation"
    assert normalize("\ufb01rmly") == "firmly", "ligature fold (NFKC)"
    assert normalize("breeder\u2019s") == "breeder's", "curly quote"
    assert normalize("a \u2014 b") == "a - b", "em dash"
    assert normalize("a\u00a0b") == "a b", "non-breaking space"

    # 2. Exact match across the smart quote -> verified.
    r = _verify("The breeder's rights regime promotes innovation")
    assert r.verification_status == STATUS_VERIFIED, r.verification_status
    assert r.match_score == 100.0
    assert r.matched_page == 1

    # 3. A quote crossing the hyphenated line break -> still verified.
    r = _verify("firmly across the innovation pipeline")
    assert r.verification_status == STATUS_VERIFIED, r.verification_status

    # 4. A quote crossing the PAGE BREAK -> still verified (pages joined).
    r = _verify("innovation pipeline. A second page discusses compulsory")
    assert r.verification_status == STATUS_VERIFIED, r.verification_status

    # 5. Elided quote with "..." -> verified_fragmented.
    r = _verify("The breeder's rights regime ... protects new plant varieties")
    assert r.verification_status == STATUS_VERIFIED_FRAGMENTED, r.verification_status

    # 6. Lightly edited long quote (one word changed) -> fuzzy verified.
    r = _verify(
        "The breeder's rights system promotes innovation \u2014 and protects "
        "new plant varieties firmly across the innovation pipeline"
    )
    assert r.verification_status == STATUS_VERIFIED_FUZZY, r.verification_status
    assert r.match_score >= 95

    # 7. Short, slightly-off snippet -> demoted to near_match, not auto-verified.
    r = _verify("breeders rihgts regime")  # typo, below min length for fuzzy verify
    assert r.verification_status in (STATUS_NEAR_MATCH, STATUS_NOT_FOUND), (
        r.verification_status
    )

    # 8. A fabricated quote -> not_found.
    r = _verify(
        "This study found a ninety percent increase in seed exports to "
        "the European Union following ratification."
    )
    assert r.verification_status == STATUS_NOT_FOUND, r.verification_status

    # 9. Page attribution: claimed page 5, content actually on page 2.
    r = _verify("compulsory licensing and farmers rights", page=5)
    assert r.matched_page == 2 and r.page_ok is False, (r.matched_page, r.page_ok)

    # 10. A scanned PDF (no text layer) -> no_text_layer, never not_found.
    empty = DocText(filename="scan.pdf", has_text_layer=False)
    rec = FindingRecord("t", "h", "scan.pdf", 1, "anything at all")
    assert verify_one(rec, empty).verification_status == verify.STATUS_NO_TEXT_LAYER

    test_detect_boilerplate()
    test_header_injection_across_page_break()

    print("All verification fixture tests passed.")


def test_detect_boilerplate() -> None:
    """Running margin headers are detected; repeated body text is not."""
    from verify import _detect_boilerplate

    height = 792.0
    pages = []
    for pno in range(4):
        pages.append(
            [
                # Top-margin running header; only the page number varies.
                (f"Smith, J. Journal of Testing Vol. 5 {100 + pno}", 18.0, 32.0),
                # Body text well clear of the margins.
                ("Distinct body content for this page", 300.0, 340.0),
            ]
        )
    boiler = _detect_boilerplate(pages, [height] * 4)
    assert any("journal of testing" in s for s in boiler), boiler
    assert not any("body content" in s for s in boiler), boiler


def test_header_injection_across_page_break() -> None:
    """A quote spanning a page break verifies once the header is stripped.

    Regression test for the real failure found in the Adelaiye paper: a
    running header was extracted into the middle of a sentence that
    continued onto the next page, breaking an otherwise verbatim quote.
    """
    import os
    import tempfile

    import config

    fitz = verify._import_fitz()
    tmp = tempfile.mkdtemp()
    pdf_path = os.path.join(tmp, "injection.pdf")

    doc = fitz.open()
    p1 = doc.new_page()
    p1.insert_text((72, 28), "Smith, J. Journal of Testing Vol. 5 100")  # header
    p1.insert_text((72, 300), "The Act established an Office in the country to grant")
    p2 = doc.new_page()
    p2.insert_text((72, 28), "Smith, J. Journal of Testing Vol. 5 101")  # header
    p2.insert_text((72, 300), "breeders rights and perform other necessary functions.")
    doc.save(pdf_path)
    doc.close()

    rec = FindingRecord(
        snippet_id="t",
        finding_hash="h",
        filename="injection.pdf",
        page_number=1,
        snippet="to grant breeders rights and perform other necessary functions",
    )

    # With stripping on (the default), the spanning quote verifies exactly.
    extracted = verify.extract_doc_text(pdf_path, "injection.pdf")
    assert verify_one(rec, extracted).verification_status == STATUS_VERIFIED

    # With stripping off, the injected header breaks the exact match —
    # proving the stripping is what fixes it.
    config.VERIFY_STRIP_BOILERPLATE = False
    try:
        raw = verify.extract_doc_text(pdf_path, "injection.pdf")
        assert verify_one(rec, raw).verification_status != STATUS_VERIFIED
    finally:
        config.VERIFY_STRIP_BOILERPLATE = True


if __name__ == "__main__":
    run()
