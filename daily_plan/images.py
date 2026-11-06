"""Render sections of a filled workbook as zoomed PNG images."""

from io import BytesIO
from typing import cast

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image, ImageDraw, ImageFont

SECTION_PREFIX = "section_"
BASE_FONT_SIZE = 14
CHAR_WIDTH_PX = 8
DEFAULT_COL_WIDTH = 13
DEFAULT_ROW_PX = 24
PADDING_PX = 6
GRID_COLOR = "#BBBBBB"
DEFAULT_FILL = "#FFFFFF"


def list_sections(workbook_bytes: bytes) -> dict[str, str]:
    """Find section ranges defined by ``section_*`` names in the workbook.

    Falls back to the whole used range when no section name exists.

    Args:
        workbook_bytes: Bytes of the xlsx file.

    Returns:
        Mapping of section name to a range such as ``"A1:F10"``.
    """
    workbook = load_workbook(BytesIO(workbook_bytes))
    sheet = workbook.active
    if sheet is None:
        return {}
    sections: dict[str, str] = {}
    for name, defined in workbook.defined_names.items():
        if not name.lower().startswith(SECTION_PREFIX):
            continue
        for _, reference in defined.destinations:
            sections[name[len(SECTION_PREFIX) :]] = reference.replace("$", "")
    return sections or {"full": sheet.dimensions}


def render_sections(workbook_bytes: bytes, zoom: float = 1.0) -> dict[str, bytes]:
    """Render every section of the workbook to a PNG.

    Args:
        workbook_bytes: Bytes of the xlsx file.
        zoom: Scale factor applied to the rendered image (1.0 is default size).

    Returns:
        Mapping of section name to PNG bytes.

    Raises:
        ValueError: If zoom is not positive.
    """
    if zoom <= 0:
        raise ValueError("zoom must be positive")
    workbook = load_workbook(BytesIO(workbook_bytes))
    sheet = workbook.active
    if sheet is None:
        return {}
    return {
        name: _render_range(sheet, cell_range, zoom)
        for name, cell_range in list_sections(workbook_bytes).items()
    }


def _trim_empty(
    sheet: Worksheet, bounds: tuple[int, int, int, int]
) -> tuple[int, int, int, int]:
    """Shrink bounds so trailing empty rows and columns are dropped."""
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


def _render_range(sheet: Worksheet, cell_range: str, zoom: float) -> bytes:
    """Draw one cell range as a PNG with fills, bold text and grid lines."""
    bounds = cast(tuple[int, int, int, int], range_boundaries(cell_range))
    min_col, min_row, max_col, max_row = _trim_empty(sheet, bounds)
    scale = zoom
    font = ImageFont.load_default(size=round(BASE_FONT_SIZE * scale))
    line_height = round((BASE_FONT_SIZE + 4) * scale)
    pad = round(PADDING_PX * scale)

    col_widths = [
        round(
            (sheet.column_dimensions[get_column_letter(col)].width or DEFAULT_COL_WIDTH)
            * CHAR_WIDTH_PX
            * scale
        )
        for col in range(min_col, max_col + 1)
    ]
    row_heights = []
    for row in range(min_row, max_row + 1):
        lines = max(
            (
                str(sheet.cell(row, col).value or "").count("\n") + 1
                for col in range(min_col, max_col + 1)
            ),
            default=1,
        )
        row_heights.append(
            max(round(DEFAULT_ROW_PX * scale), lines * line_height + pad)
        )

    image = Image.new("RGB", (sum(col_widths) + 1, sum(row_heights) + 1), DEFAULT_FILL)
    draw = ImageDraw.Draw(image)
    y = 0
    for row_index, height in enumerate(row_heights):
        x = 0
        for col_index, width in enumerate(col_widths):
            cell = sheet.cell(min_row + row_index, min_col + col_index)
            fill = DEFAULT_FILL
            if cell.fill.fill_type == "solid" and isinstance(
                cell.fill.fgColor.rgb, str
            ):
                fill = f"#{cell.fill.fgColor.rgb[-6:]}"
            draw.rectangle((x, y, x + width, y + height), fill=fill, outline=GRID_COLOR)
            text = "" if cell.value is None else str(cell.value)
            for line_number, line in enumerate(text.split("\n")):
                position = (x + pad, y + pad // 2 + line_number * line_height)
                draw.text(position, line, fill="black", font=font)
                if cell.font.bold:
                    draw.text((position[0] + 1, position[1]), line, "black", font)
            x += width
        y += height

    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()
