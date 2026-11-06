"""Render sections of the filled workbook as zoomed PNG images."""

from datetime import datetime, time
from io import BytesIO
from typing import cast

from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX, Color
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image, ImageDraw, ImageFont

HEADER_LABEL = "TIME"
BASE_FONT_SIZE = 14
PX_PER_CHAR = 7
PX_PER_POINT = 96 / 72
DEFAULT_COL_WIDTH = 8.43
DEFAULT_ROW_PX = 20
WHITE = "#FFFFFF"
DARK_LIMIT = 110
BORDER_PX = {"thin": 1, "medium": 2, "thick": 3, "double": 3}

Bounds = tuple[int, int, int, int]


def list_sections(sheet: Worksheet) -> dict[str, str]:
    """Split the sheet into one section per row of player blocks.

    Args:
        sheet: Filled worksheet.

    Returns:
        Mapping of section name to a range such as ``"A1:P13"``. ``full`` covers
        the whole used range; ``players_N`` covers each row of player blocks.
    """
    last_col = get_column_letter(sheet.max_column)
    starts = sorted(
        {c.row for row in sheet.iter_rows() for c in row if c.value == HEADER_LABEL}
    )
    sections = {"full": sheet.dimensions}
    for index, start in enumerate(starts):
        end = starts[index + 1] - 1 if index + 1 < len(starts) else sheet.max_row
        first = 1 if index == 0 else start
        sections[f"players_{index + 1}"] = f"A{first}:{last_col}{end}"
    return sections


def render_sections(workbook_bytes: bytes, zoom: float = 1.0) -> dict[str, bytes]:
    """Render every section of the active sheet to a PNG.

    Args:
        workbook_bytes: Bytes of the xlsx file.
        zoom: Scale factor (1.0 is default size).

    Returns:
        Mapping of section name to PNG bytes.

    Raises:
        ValueError: If zoom is not positive.
    """
    if zoom <= 0:
        raise ValueError("zoom must be positive")
    sheet = cast(Worksheet, load_workbook(BytesIO(workbook_bytes)).active)
    return {
        name: _render_range(sheet, cell_range, zoom)
        for name, cell_range in list_sections(sheet).items()
    }


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


def _color(color: Color | None) -> str | None:
    """Resolve an rgb or indexed openpyxl color to ``#RRGGBB`` (None otherwise)."""
    if color is None:
        return None
    if color.type == "rgb" and isinstance(color.rgb, str):
        return f"#{color.rgb[-6:]}"
    if color.type == "indexed" and isinstance(color.indexed, int):
        if color.indexed < len(COLOR_INDEX):
            return f"#{COLOR_INDEX[color.indexed][-6:]}"
    return None


def _contrast(background: str | None) -> str:
    """Return white ink on dark backgrounds, black otherwise."""
    if background is None:
        return "black"
    red, green, blue = (int(background[i : i + 2], 16) for i in (1, 3, 5))
    return (
        "white" if 0.299 * red + 0.587 * green + 0.114 * blue < DARK_LIMIT else "black"
    )


def _text(value: object) -> str:
    """Format a cell value; times use the ``h.mm`` style of the template."""
    if isinstance(value, datetime | time):
        return f"{value.hour}.{value.minute:02d}"
    return "" if value is None else str(value)


def _render_range(sheet: Worksheet, cell_range: str, zoom: float) -> bytes:
    """Draw one range as a PNG with fills, borders, merged and centered cells."""
    bounds = _trim_empty(sheet, cast(Bounds, range_boundaries(cell_range)))
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
    fill_colors: dict[int, str] = {}
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
            fill = _color(cell.fill.fgColor) if cell.fill.fill_type == "solid" else None
            if fill:
                draw.rectangle(box, fill=fill)
                fill_colors[id(cell)] = fill
            boxes.append((cell, box))

    for cell, (left, top, right, bottom) in boxes:
        text = _text(cell.value)
        if text:
            centered = cell.alignment.horizontal == "center"
            x = (left + right) / 2 if centered else left + 4 * zoom
            anchor = "mm" if centered else "lm"
            position = (x, (top + bottom) / 2)
            ink = _contrast(fill_colors.get(id(cell)))
            draw.text(position, text, fill=ink, font=font, anchor=anchor)
            if cell.font.b:
                draw.text((x + 1, position[1]), text, ink, font, anchor)
        for side, line in (
            (cell.border.left, (left, top, left, bottom)),
            (cell.border.right, (right, top, right, bottom)),
            (cell.border.top, (left, top, right, top)),
            (cell.border.bottom, (left, bottom, right, bottom)),
        ):
            if side.style:
                width = max(1, round(BORDER_PX.get(side.style, 1) * zoom))
                draw.line(line, fill="black", width=width)

    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()
