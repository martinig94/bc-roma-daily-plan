"""Render the team sheet and each player block of the workbook as PNG images."""

import re
from datetime import date, datetime, time
from io import BytesIO
from typing import cast

from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX, Color
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image, ImageDraw, ImageFont, ImageOps

from daily_plan.fill import DATA_ROWS, HEADER_LABEL

TEAM_PREFIX = "bc_roma_daily"
BASE_FONT_SIZE = 14
DEFAULT_ZOOM = 2.0
MARGIN_PX = 12
PX_PER_CHAR = 7
PX_PER_POINT = 96 / 72
DEFAULT_COL_WIDTH = 8.43
DEFAULT_ROW_PX = 20
WHITE = "#FFFFFF"
BLACK = "#000000"
BORDER_PX = {"thin": 1, "medium": 2, "thick": 3, "double": 3}
NUMBER_PREFIX = re.compile(r"^#\d+\s*")

Bounds = tuple[int, int, int, int]


def render_images(
    workbook_bytes: bytes, plan_date: date, zoom: float = DEFAULT_ZOOM
) -> dict[str, bytes]:
    """Render the whole team sheet and a zoom on every player block.

    Args:
        workbook_bytes: Bytes of the filled xlsx file.
        plan_date: Date used in the image names.
        zoom: Scale factor applied to the rendered images.

    Returns:
        Mapping of image name (no extension) to PNG bytes: ``bc_roma_daily_<date>``
        for the team and ``<Surname>_<date>`` for each player.

    Raises:
        ValueError: If zoom is not positive.
    """
    if zoom <= 0:
        raise ValueError("zoom must be positive")
    workbook = load_workbook(BytesIO(workbook_bytes))
    sheet = cast(Worksheet, workbook.active)
    palette = [str(color) for color in workbook._colors]  # type: ignore[attr-defined]
    stamp = f"{plan_date:%Y-%m-%d}"

    used = (1, 1, sheet.max_column, sheet.max_row)
    images = {
        f"{TEAM_PREFIX}_{stamp}": _render_range(
            sheet, _trim_empty(sheet, used), zoom, palette
        )
    }
    for cell in (c for row in sheet.iter_rows() for c in row):
        if cell.value != HEADER_LABEL:
            continue
        title = sheet.cell(cell.row, cell.column + 1).value
        if not title:
            continue
        name = NUMBER_PREFIX.sub("", str(title)).title()
        bounds = (cell.column, cell.row, cell.column + 1, cell.row + DATA_ROWS)
        images[f"{name}_{stamp}"] = _render_range(sheet, bounds, zoom, palette)
    return images


def _trim_empty(sheet: Worksheet, bounds: Bounds) -> Bounds:
    """Drop trailing rows and columns that hold no value."""
    min_col, min_row, max_col, max_row = bounds
    filled = [
        (cell.row, cell.column)
        for row in sheet.iter_rows(min_row, max_row, min_col, max_col)
        for cell in row
        if cell.value not in (None, "")
    ]
    if not filled:
        return bounds
    return (
        min_col,
        min_row,
        max(column for _, column in filled),
        max(row for row, _ in filled),
    )


def _color(color: Color | None, palette: list[str]) -> str | None:
    """Resolve an rgb or indexed color (using the workbook palette) to ``#RRGGBB``."""
    if color is None:
        return None
    if color.type == "rgb" and isinstance(color.rgb, str):
        return f"#{color.rgb[-6:]}"
    if color.type == "indexed" and isinstance(color.indexed, int):
        colors = palette or COLOR_INDEX
        if color.indexed < len(colors):
            return f"#{colors[color.indexed][-6:]}"
    return None


def _text(value: object) -> str:
    """Format a cell value; times use the ``h.mm`` style of the template."""
    if isinstance(value, datetime | time):
        return f"{value.hour}.{value.minute:02d}"
    return "" if value is None else str(value)


def _render_range(
    sheet: Worksheet, bounds: Bounds, zoom: float, palette: list[str]
) -> bytes:
    """Draw a cell range as PNG with fills, fonts, borders and merged cells."""
    min_col, min_row, max_col, max_row = bounds
    font = ImageFont.load_default(size=round(BASE_FONT_SIZE * zoom))

    xs, ys = [0], [0]
    for col in range(min_col, max_col + 1):
        width = sheet.column_dimensions[get_column_letter(col)].width
        xs.append(xs[-1] + round((width or DEFAULT_COL_WIDTH) * PX_PER_CHAR * zoom))
    for row in range(min_row, max_row + 1):
        height = sheet.row_dimensions[row].height
        px = height * PX_PER_POINT if height else DEFAULT_ROW_PX
        ys.append(ys[-1] + round(px * zoom))

    spans: dict[tuple[int, int], tuple[int, int]] = {}
    covered: set[tuple[int, int]] = set()
    for merged in sheet.merged_cells.ranges:
        c0, r0 = max(merged.min_col, min_col), max(merged.min_row, min_row)
        c1, r1 = min(merged.max_col, max_col), min(merged.max_row, max_row)
        if c0 > c1 or r0 > r1:
            continue
        spans[(r0, c0)] = (r1, c1)
        covered.update((r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1))
        covered.discard((r0, c0))

    image = Image.new("RGB", (xs[-1] + 1, ys[-1] + 1), WHITE)
    draw = ImageDraw.Draw(image)
    boxes = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if (row, col) in covered:
                continue
            last_row, last_col = spans.get((row, col), (row, col))
            box = (
                xs[col - min_col],
                ys[row - min_row],
                xs[last_col - min_col + 1],
                ys[last_row - min_row + 1],
            )
            cell = sheet.cell(row, col)
            fill = None
            if cell.fill.fill_type == "solid":
                fill = _color(cell.fill.fgColor, palette)
                if fill:
                    draw.rectangle(box, fill=fill)
            boxes.append((cell, box, fill))

    for cell, (left, top, right, bottom), fill in boxes:
        text = _text(cell.value)
        if text:
            centered = cell.alignment.horizontal == "center"
            x = (left + right) / 2 if centered else left + 4 * zoom
            anchor = "mm" if centered else "lm"
            ink = _color(cell.font.color, palette) or BLACK
            if fill and ink.lower() == fill.lower():
                ink = BLACK
            y = (top + bottom) / 2
            draw.text((x, y), text, fill=ink, font=font, anchor=anchor)
            if cell.font.b:
                draw.text((x + 1, y), text, ink, font, anchor)
        for side, line in (
            (cell.border.left, (left, top, left, bottom)),
            (cell.border.right, (right, top, right, bottom)),
            (cell.border.top, (left, top, right, top)),
            (cell.border.bottom, (left, bottom, right, bottom)),
        ):
            if side.style:
                width = max(1, round(BORDER_PX.get(side.style, 1) * zoom))
                draw.line(line, fill=BLACK, width=width)

    image = ImageOps.expand(image, border=round(MARGIN_PX * zoom), fill=WHITE)
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()
