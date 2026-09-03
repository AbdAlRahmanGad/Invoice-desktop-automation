"""Geometry primitives for reconstructing a page's layout from OCR output.

Knows nothing about OCR engines or about sales orders -- it only groups
positioned text into rows and decides which label a value belongs to. Pure
functions over `TextBox`, so this layer is testable without running OCR.
"""

import re
from dataclasses import dataclass


#: Row/column tolerances as a fraction of page width, so they survive a change
#: of render scale (a 2x and a 4x render of the same page behave the same).
Y_TOL_RATIO = 8 / 1080
MAX_DIST_RATIO = 80 / 1080

_NON_ALNUM_RE = re.compile(r"[^A-Z0-9]")


@dataclass(frozen=True)
class TextBox:
    """One piece of recognised text, positioned by its centre point.

    `left` and `right` carry how wide the recognised region was, which matters
    when reading a ruled table: a box whose extent crosses a column line has
    swallowed two cells and has to be read again.
    """

    text: str
    x: float
    y: float
    score: float = 1.0
    left: float = 0.0
    right: float = 0.0


def normalize(text: str) -> str:
    """Normalize a label for matching; OCR drops inter-word spaces unpredictably."""
    return _NON_ALNUM_RE.sub("", text.upper())


def cluster_rows(boxes: list[TextBox], y_tol: float) -> list[list[TextBox]]:
    """Group boxes into rows by vertical proximity, each row sorted left to right.

    Args:
        boxes (list[TextBox]): Boxes in any order.
        y_tol (float): Max vertical gap between consecutive boxes in a row.

    Returns:
        list[list[TextBox]]: Rows, top to bottom.
    """
    ordered = sorted(boxes, key=lambda b: b.y)
    rows: list[list[TextBox]] = []
    for b in ordered:
        if rows and b.y - rows[-1][-1].y <= y_tol:
            rows[-1].append(b)
        else:
            rows.append([b])
    return [sorted(row, key=lambda b: b.x) for row in rows]


def nearest_label(candidates: dict[str, tuple[float, float]], box: TextBox, max_dist: float) -> str | None:
    """Pick the label a value box belongs to, by column then by vertical order.

    x alone is ambiguous when labels stack in one column (EMAIL directly above
    PHONE share an x-centre), so among labels in the same column we take the
    closest one *above* the value.

    Args:
        candidates (dict[str, tuple[float, float]]): field -> (x, y) of its label.
        box (TextBox): The value box being assigned.
        max_dist (float): Max horizontal distance to count as the same column.

    Returns:
        str | None: The matching field name, or None if no label is in range.
    """
    in_column = {f: (lx, ly) for f, (lx, ly) in candidates.items() if abs(lx - box.x) <= max_dist}
    if not in_column:
        return None
    above = {f: ly for f, (_, ly) in in_column.items() if ly < box.y}
    if above:
        return max(above, key=lambda f: above[f])
    return min(in_column, key=lambda f: abs(in_column[f][0] - box.x))
