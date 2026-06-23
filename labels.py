"""Labels module for loading and validating the coding framework.

The coding framework is stored in an Excel file with one row per label.
Each label belongs to a **category** (a broader grouping), has a **code**
(identifier within the category), a description, and inclusion / exclusion
criteria that are fed directly into the LLM prompt.

The pipeline runs once per category: for each category, only the labels
belonging to that category are sent to the LLM.  This keeps each prompt
focused and lets different categories of coding live in a single Excel
file without forcing the LLM to juggle too many labels at once.

Expected Excel columns:
    category           — Grouping identifier (e.g. "CPVR_EFFECTIVENESS").
                         Multiple rows can share the same category.
    code               — Identifier within the category (e.g. "OBJECTIVES").
                         The pair (category, code) must be unique.
    description        — What this label covers.
    inclusion_criteria — When to apply this label.
    exclusion_criteria — When *not* to apply this label.  (Optional.)

Typical usage::

    from labels import load_labels, group_labels_by_category

    labels = load_labels("labels/coding_framework.xlsx")
    by_category = group_labels_by_category(labels)
    for cat, cat_labels in by_category.items():
        ...
"""

from __future__ import annotations

import logging
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Set

import pandas as pd

logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────────────────
# Data structure
# ──────────────────────────────────────────────────────────────────────

REQUIRED_COLUMNS = ["category", "code", "description", "inclusion_criteria"]
"""Columns that must be present and non-empty in the Excel file."""

OPTIONAL_COLUMNS = ["exclusion_criteria"]
"""Columns that are recommended but not strictly required."""


@dataclass(frozen=True)
class Label:
    """A single label in the coding framework.

    Attributes:
        category: Broader grouping (e.g. "CPVR_EFFECTIVENESS").  Multiple
            labels can share a category; the pipeline runs one batch per
            category.
        code: Identifier within the category (e.g. "OBJECTIVES").  The
            pair (category, code) must be unique across the file.
        description: What this label covers.
        inclusion_criteria: Rules for when a passage should receive this
            label.
        exclusion_criteria: Rules for when a passage should NOT receive
            this label.  May be empty.
    """

    category: str
    code: str
    description: str
    inclusion_criteria: str
    exclusion_criteria: str = ""

    def to_prompt_block(self) -> str:
        """Format this label as a structured text block for the LLM prompt.

        Returns:
            A multi-line string describing the label, ready to be inserted
            into a prompt template.  The category is NOT included in the
            block because each prompt is category-specific already.
        """
        block = (
            f"### Label: {self.code}\n"
            f"Description: {self.description}\n"
            f"Include if: {self.inclusion_criteria}\n"
        )
        if self.exclusion_criteria:
            block += f"Exclude if: {self.exclusion_criteria}\n"
        return block


# ──────────────────────────────────────────────────────────────────────
# Loading & validation
# ──────────────────────────────────────────────────────────────────────


def load_labels(excel_path: str | Path) -> List[Label]:
    """Load the coding framework from an Excel file.

    Reads the first sheet, validates required columns, checks that the
    pair ``(category, code)`` is unique across all rows, and returns a
    list of ``Label`` objects in the order they appear in the file.

    Args:
        excel_path: Path to the ``.xlsx`` file containing the labels.

    Returns:
        A list of validated ``Label`` objects, one per row.

    Raises:
        FileNotFoundError: If the Excel file does not exist.
        ValueError: If required columns are missing, if any required cell
            is empty, or if ``(category, code)`` pairs are not unique.
    """
    path = Path(excel_path)
    if not path.exists():
        raise FileNotFoundError(f"Labels file not found: {path}")

    logger.info("Loading labels from %s", path)
    df = pd.read_excel(path, dtype=str).fillna("")

    # --- Normalise column names (lowercase, strip whitespace) -----------
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

    # --- Check required columns -----------------------------------------
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Labels Excel is missing required columns: {missing}. "
            f"Found columns: {list(df.columns)}"
        )

    # --- Warn about missing optional columns ----------------------------
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            logger.warning(
                "Optional column '%s' not found in labels Excel. "
                "Adding empty column.",
                col,
            )
            df[col] = ""

    # --- Validate non-empty required fields -----------------------------
    for col in REQUIRED_COLUMNS:
        empty_mask = df[col].str.strip() == ""
        if empty_mask.any():
            bad_rows = list(df.index[empty_mask] + 2)  # +2 for header + 0-index
            raise ValueError(
                f"Column '{col}' has empty values in Excel rows: {bad_rows}"
            )

    # --- Check uniqueness of (category, code) pairs ---------------------
    df["_pair"] = df["category"].str.strip() + "||" + df["code"].str.strip()
    duplicates = df["_pair"][df["_pair"].duplicated()].unique().tolist()
    if duplicates:
        pretty = [p.replace("||", " / ") for p in duplicates]
        raise ValueError(
            f"Duplicate (category, code) pairs found: {pretty}. "
            f"The same code is allowed in different categories, but not twice "
            f"in the same category."
        )

    # --- Build Label objects --------------------------------------------
    labels: List[Label] = []
    for _, row in df.iterrows():
        label = Label(
            category=row["category"].strip(),
            code=row["code"].strip(),
            description=row["description"].strip(),
            inclusion_criteria=row["inclusion_criteria"].strip(),
            exclusion_criteria=row.get("exclusion_criteria", "").strip(),
        )
        labels.append(label)

    logger.info(
        "Loaded %d labels across %d categories",
        len(labels),
        len({l.category for l in labels}),
    )
    return labels


def group_labels_by_category(labels: List[Label]) -> Dict[str, List[Label]]:
    """Group labels by their category, preserving first-appearance order.

    Args:
        labels: Full list of ``Label`` objects from ``load_labels()``.

    Returns:
        An ``OrderedDict`` mapping category name -> list of labels in
        that category.  Categories are ordered by first appearance in the
        input list so behaviour is deterministic and reproducible.
    """
    grouped: "OrderedDict[str, List[Label]]" = OrderedDict()
    for label in labels:
        grouped.setdefault(label.category, []).append(label)
    return grouped


def get_valid_codes(labels: List[Label]) -> Set[str]:
    """Return the set of valid label codes in the given list.

    Used to detect LLM hallucinations — any returned ``label_code`` that
    is not in this set did not appear in the prompt.

    Note:
        Codes are only required to be unique within a category, so if you
        pass in labels from multiple categories you may get back a set
        where two labels share a code.  In practice this function is
        called per-category, so that is not an issue.

    Args:
        labels: List of ``Label`` objects.

    Returns:
        A set of label code strings.
    """
    return {label.code for label in labels}
