"""Prompt construction module for the literature review pipeline.

This module builds the analysis prompt that is sent to the LLM alongside
each PDF document.  The prompt is dynamically generated from the coding
framework (labels) so that the LLM receives the full inclusion/exclusion
criteria for each label.

Separating prompt logic into its own module makes it easy to:
    - Iterate on prompt engineering without touching other code.
    - A/B test different prompt versions.
    - Add new prompt strategies (e.g. chain-of-thought, multi-pass).

Typical usage::

    from labels import load_labels
    from prompt import build_analysis_prompt

    labels = load_labels("labels/coding_framework.xlsx")
    prompt_text = build_analysis_prompt(labels, include_confidence=True)
"""

from __future__ import annotations

from typing import List

from labels import Label


def build_analysis_prompt(
    labels: List[Label],
    include_confidence: bool = True,
) -> str:
    """Build the full analysis prompt from the coding framework.

    The prompt instructs the LLM to act as a systematic reviewer, analyse
    the attached PDF, and return structured JSON findings.

    Args:
        labels: List of ``Label`` objects defining the coding framework.
        include_confidence: If ``True``, the prompt asks the LLM to include
            a confidence score ("high", "medium", "low") for each finding.

    Returns:
        A string containing the complete prompt text.
    """
    # --- Build the labels section ---------------------------------------
    labels_section = "\n".join(label.to_prompt_block() for label in labels)

    # --- Build the JSON schema description ------------------------------
    finding_fields = (
        '        {\n'
        '            "label_code": "the label code from the framework above",\n'
        '            "snippet": "verbatim quote from the PDF",\n'
        '            "page_number": 1,\n'
        '            "reasoning": "brief explanation of why this snippet '
        'matches the label"'
    )

    if include_confidence:
        finding_fields += ',\n            "confidence": "high | medium | low"'

    finding_fields += "\n        }"

    # --- Assemble the full prompt ---------------------------------------
    prompt = f"""\
You are an expert systematic reviewer supporting qualitative literature analysis.

## Task

Analyse the attached PDF document and extract verbatim passages relevant to
ANY of the labels defined in the coding framework below.

## Coding Framework

{labels_section}

## Output Format

Return ONLY valid JSON (no markdown fences, no commentary, no preamble).
Use EXACTLY this structure:

{{
    "findings": [
{finding_fields}
    ]
}}

## Rules

1. Use verbatim quotes for the "snippet" field.  Do NOT paraphrase.
2. Provide best-effort page numbers based on the document.
3. The "label_code" MUST be one of the codes defined above.  Do NOT invent
   new codes.
4. If a label is not present in the document, simply omit it.  Do NOT
   fabricate content.
5. Keep each snippet concise but sufficient — typically 1 to 5 sentences.
6. A single passage may be relevant to multiple labels.  In that case,
   create one finding entry per label with the same snippet.
7. Pay close attention to the inclusion AND exclusion criteria for each
   label.  Only code a passage if it meets the inclusion criteria and does
   NOT meet the exclusion criteria.
"""

    if include_confidence:
        prompt += """\
8. For the "confidence" field, assess how clearly the snippet matches the
   label:
   - "high"   — clearly and directly relevant.
   - "medium" — relevant but somewhat ambiguous.
   - "low"    — tangentially relevant; included for completeness.
"""

    return prompt


def build_summary_prompt(labels: List[Label]) -> str:
    """Build a prompt for synthesising findings across documents.

    This is an optional second-pass prompt that can be used after initial
    coding to generate narrative summaries per label.

    Args:
        labels: List of ``Label`` objects defining the coding framework.

    Returns:
        A string containing the synthesis prompt text.

    Note:
        This function is provided as a starting point for the multi-pass
        analysis feature.  It is NOT called by the main pipeline by default.
    """
    codes = ", ".join(label.code for label in labels)
    return f"""\
You are an expert systematic reviewer.

You will be given a JSON object containing coded findings from multiple
academic documents.  Each finding has a label_code, a verbatim snippet,
and a reasoning field.

Your task: for each label code ({codes}), write a concise narrative
synthesis (3–5 sentences) that summarises the key themes across all
snippets coded under that label.

Return ONLY valid JSON:

{{
    "syntheses": [
        {{
            "label_code": "...",
            "summary": "narrative synthesis here"
        }}
    ]
}}
"""
