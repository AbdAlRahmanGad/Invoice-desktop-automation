"""Reading a drawn table off the screen.

Fakturama's lists are NatTable canvases: they paint their cells rather than
building controls, so the accessibility tree holds no rows, no cells and no
headers at all. The only way to know what a list shows is to look at it, which
makes this the one place where the app-driving side borrows the OCR the rest
of the project uses on documents.

Cells are located by the table's own grid lines rather than by guessing from
text positions: numeric columns are right-aligned and text columns are
left-aligned, so a value can sit closer to a neighbouring header than to its
own, while the ruled line between them is unambiguous.
"""

from dataclasses import dataclass, field
from typing import Any

import ocr
from layout import TextBox, cluster_rows, normalize


#: Captures are upscaled before OCR: list rows are ~11px of text, which the
#: recogniser reads unreliably at native size.
SCALE = 2

#: A column separator is a pixel column darker than this for at least this
#: fraction of the header band. Only the header can be relied on to be ruled:
#: the address selector draws vertical lines across its header and nothing
#: below it, while a Data list rules the full height.
LINE_GREY = 210
LINE_COVERAGE = 0.9

#: Fraction of the capture to treat as the header when it has no ruled line
#: under its header row.
HEADER_BAND_FALLBACK = 0.05

#: More separators than a list plausibly has means the header band was read
#: as one dark smear -- a filled header does that -- so the boundaries are
#: taken from the ruled body instead.
MAX_COLUMNS = 12

#: Inset, in scaled pixels, when cropping a cell out of the grid, so the
#: crop holds the cell's text and none of its borders.
CELL_INSET = 2

#: Separator pixels this close together belong to the same line, and a line
#: no thicker than this. Thickness is what distinguishes a rule from a filled
#: band: a selected row is painted solid blue and a header is filled grey,
#: both dark right across the table, and neither is a column or row boundary.
LINE_GAP = 3
MAX_LINE_THICKNESS = 5

#: Row clustering tolerance, as a fraction of the capture's height.
ROW_TOL_RATIO = 0.02

#: How far above and below a row's text to read it when the table draws no
#: rules to bound it, as a fraction of the distance between rows.
UNRULED_BAND_RATIO = 0.35


def column_edges(image: Any, band: tuple[int, int] | None = None) -> list[float]:
    """Find the x of each vertical grid line in `image`.

    Args:
        image (Any): The upscaled capture, as a PIL image.
        band (tuple[int, int] | None): Rows of the image to look in, as
            (top, bottom). Defaults to the whole image.

    Returns:
        list[float]: Separator positions, left to right.
    """
    import numpy as np

    top, bottom = band or (0, image.height)
    grey = np.array(image.convert("L"), dtype=np.int16)[top:bottom]
    dark = (grey < LINE_GREY).sum(axis=0)
    columns = [x for x in range(image.width) if dark[x] >= (bottom - top) * LINE_COVERAGE]

    return _merge(columns)


def _merge(positions: list[int]) -> list[float]:
    """Collapse runs of adjacent pixel positions into one line each.

    Args:
        positions (list[int]): Dark rows or columns, in order.

    Returns:
        list[float]: The centre of each run.
    """
    runs: list[list[int]] = []
    for position in positions:
        if runs and position - runs[-1][-1] <= LINE_GAP:
            runs[-1].append(position)
        else:
            runs.append([position])
    return [sum(run) / len(run) for run in runs if len(run) <= MAX_LINE_THICKNESS]


def _header_band(image: Any) -> tuple[int, int]:
    """The slice of the capture holding the header row.

    Column separators are looked for here rather than over the whole picture,
    because a list may rule its header and leave its body unruled -- the
    address selector does exactly that. The header ends at the first ruled
    line across the table, which is found the same way the vertical ones are.

    Args:
        image (Any): The upscaled capture.

    Returns:
        tuple[int, int]: Top and bottom of the band.
    """
    import numpy as np

    lines = row_lines(image)
    if lines:
        return 0, int(lines[0])
    return 0, int(image.height * HEADER_BAND_FALLBACK)


def row_lines(image: Any) -> list[float]:
    """Find the y of each horizontal grid line in `image`.

    Args:
        image (Any): The upscaled capture.

    Returns:
        list[float]: Rule positions, top to bottom.
    """
    import numpy as np

    grey = np.array(image.convert("L"), dtype=np.int16)
    dark = (grey < LINE_GREY).sum(axis=1)
    return _merge([y for y in range(image.height) if dark[y] >= image.width * LINE_COVERAGE])


def _read_cells(
    image: Any, edges: list[float], band: tuple[float, float], columns: set[int] | None = None
) -> dict[int, str]:
    """Read one row by reading each of its cells on its own.

    Reading the row in one go is faster but wrong at the edges: the
    recogniser joins a cell's text to its neighbour's when they nearly touch,
    and a value truncated to "Northstar Office ..." abutting the rule comes
    back with the ZIP stuck to it. Cropping to the cell makes the boundary a
    fact rather than a guess -- and gives the recogniser a clean, isolated
    word, which also spares it the O-for-zero mistakes it makes on customer
    numbers in a crowded row.

    Args:
        image (Any): The upscaled capture.
        edges (list[float]): Column separators.
        band (tuple[float, float]): Top and bottom of the row.
        columns (set[int] | None): Which columns to read; all of them by default.

    Returns:
        dict[int, str]: Column index -> the text in that cell, empties left out.
    """
    top, bottom = int(band[0]) + CELL_INSET, int(band[1]) - CELL_INSET
    bounds = [0.0, *edges, float(image.width)]
    cells: dict[int, str] = {}
    for index in range(len(bounds) - 1):
        if columns is not None and index not in columns:
            continue
        crop = image.crop((int(bounds[index]) + CELL_INSET, top, int(bounds[index + 1]) - CELL_INSET, bottom))
        text = " ".join(box.text for box in ocr.read_image(crop) if box.text)
        if text:
            cells[index] = text
    return cells


def _band_for(y: float, lines: list[float], height: int, spacing: float) -> tuple[float, float]:
    """The ruled band containing `y`, or one guessed from the row spacing.

    Args:
        y (float): A row's vertical centre.
        lines (list[float]): Horizontal rules.
        height (int): The capture's height.
        spacing (float): Distance between rows, for unruled lists.

    Returns:
        tuple[float, float]: Top and bottom of the row.
    """
    # Never reach further than most of the way to the neighbouring row, ruled
    # or not: when a rule between two rows goes undetected, the nearest one
    # above can be the table's top border, and a band stretching back to it
    # pulls the header's own words into the cell ("Qty. 1.00").
    reach = abs(spacing) * UNRULED_BAND_RATIO
    top, bottom = max(0.0, y - reach), min(float(height), y + reach)

    above = [line for line in lines if line < y]
    below = [line for line in lines if line > y]
    if above:
        top = max(top, max(above))
    if below:
        bottom = min(bottom, min(below))
    return top, bottom


def _column_of(x: float, edges: list[float]) -> int:
    """Return which column `x` falls in, given the separators between them."""
    return sum(1 for edge in edges if x >= edge)


@dataclass(frozen=True)
class Row:
    """One row of a drawn list: what it says, and where it is."""

    cells: dict[str, str] = field(default_factory=dict)
    y: float = 0.0

    def get(self, column: str, default: str = "") -> str:
        """Return the text in `column`, or `default` when the cell is empty.

        Column names come from reading the header, so they carry the same
        noise as any other cell -- "First Name" comes back as "First Name."
        often enough that an exact key lookup is not safe. Names are compared
        with punctuation and case ignored, the way the labels are matched
        elsewhere.

        Args:
            column (str): The column's name as it reads on screen.
            default (str): What to return when there is no such cell.

        Returns:
            str: The cell's text.
        """
        wanted = normalize(column)
        for name, text in self.cells.items():
            if normalize(name) == wanted:
                return text
        return default


def header_positions(image: Any) -> dict[str, float]:
    """Where each column's header sits, in the capture's own pixels.

    Clicking a cell needs a position, and the ruled grid is not always
    readable -- the order's item table paints a filled header over light
    rules. The header text itself is unambiguous, so its words give the
    columns their x, which is all a click needs: cells are far wider than the
    error in a word's centre.

    Args:
        image (Any): A capture of the list, including its header row.

    Returns:
        dict[str, str]: Column name -> the x of its header, with the trailing
            dot Fakturama puts on abbreviations ("Qty.", "Pos.") removed.
    """
    from PIL import Image

    scaled = image.resize((image.width * SCALE, image.height * SCALE), Image.LANCZOS)
    rows = cluster_rows(ocr.read_image(scaled), scaled.height * ROW_TOL_RATIO)
    if not rows:
        return {}
    return {box.text.strip().rstrip("."): box.x / SCALE for box in rows[0] if box.text.strip()}


def read(image: Any) -> list[Row]:
    """Read a captured list into rows.

    The first text row is taken as the header, so each row's cells come back
    keyed by what the column is called on screen -- the same names the step
    definitions use ("Company", "First Name", "Name", "ZIP", "City").

    Each row also carries its vertical position in the capture, in the
    original image's pixels, because selecting a row means clicking it: the
    list draws its rows, so there is nothing else to aim at.

    Args:
        image (Any): A capture of the list, as a PIL image, including its
            header row.

    Returns:
        list[Row]: One per row below the header. Columns with no text in that
            row are absent rather than empty, and a column with no header text
            is keyed by its index.
    """
    from PIL import Image

    scaled = image.resize((image.width * SCALE, image.height * SCALE), Image.LANCZOS)
    lines = row_lines(scaled)
    edges = column_edges(scaled, _header_band(scaled))
    if not edges or len(edges) > MAX_COLUMNS:
        # A grey-filled header (the Data lists paint one) is dark all the way
        # across, so every pixel column in it looks like a rule. Those lists
        # do rule their bodies, which is the more reliable place to measure.
        edges = column_edges(scaled)

    # One pass over the whole list first: it names the columns, and says which
    # rows have anything in them. Only those rows are then read cell by cell,
    # so an empty list costs one reading rather than one per ruled row.
    rows = cluster_rows(ocr.read_image(scaled), scaled.height * ROW_TOL_RATIO)
    if not rows:
        return []

    names = _cells(rows[0], edges)
    centres = [sum(box.y for box in row) / len(row) for row in rows]
    spacing = (centres[1] - centres[0]) if len(centres) > 1 else scaled.height * ROW_TOL_RATIO * 2

    read_rows = []
    for row, centre in zip(rows[1:], centres[1:]):
        cells = _cells(row, edges)
        # A box that reaches across a column line has read two cells as one.
        # Those columns -- and only those -- are read again from their own
        # crops, where the boundary is a fact rather than a guess.
        straddled = {
            column
            for box in row
            for column in _columns_spanned(box, edges)
            if len(_columns_spanned(box, edges)) > 1
        }
        if straddled:
            band = _band_for(centre, lines, scaled.height, spacing)
            reread = _read_cells(scaled, edges, band, straddled)
            cells = {index: text for index, text in cells.items() if index not in straddled}
            cells.update(reread)
        read_rows.append(
            Row(
                cells={names.get(index, str(index)): text for index, text in sorted(cells.items())},
                y=centre / SCALE,
            )
        )
    return read_rows


def _columns_spanned(box: TextBox, edges: list[float]) -> list[int]:
    """Which columns a recognised box overlaps.

    Args:
        box (TextBox): A recognised piece of text.
        edges (list[float]): Column separators.

    Returns:
        list[int]: Column indices, in order; more than one means it straddles.
    """
    first, last = _column_of(box.left, edges), _column_of(box.right, edges)
    return list(range(first, last + 1))


def _cells(row: list[TextBox], edges: list[float]) -> dict[int, str]:
    """Group one row's boxes into cells by column index.

    Args:
        row (list[TextBox]): The row's text, left to right.
        edges (list[float]): Column separators.

    Returns:
        dict[int, str]: Column index -> the text in that cell.
    """
    cells: dict[int, list[str]] = {}
    for box in row:
        if box.text:
            cells.setdefault(_column_of(box.x, edges), []).append(box.text)
    return {index: " ".join(parts) for index, parts in cells.items()}
