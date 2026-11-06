"""Fill an xlsx template with CSV data and the extra daily-plan inputs."""

import re
from dataclasses import dataclass
from datetime import time
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

PLACEHOLDER_PATTERN = re.compile(r"\{\{(\w+)\}\}")
TABLE_KEY = "csv_table"
TIME_FORMAT = "%H:%M"


@dataclass(frozen=True)
class Meeting:
    """One individual meeting.

    Attributes:
        who: Person met.
        at: Meeting start time.
    """

    who: str
    at: time


def read_csv(source: str | Path | BinaryIO) -> pd.DataFrame:
    """Read the uploaded CSV into a dataframe.

    Args:
        source: Path or binary file object.

    Returns:
        Parsed dataframe.

    Raises:
        ValueError: If the CSV is empty or cannot be parsed.
    """
    try:
        table = pd.read_csv(source)
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as error:
        raise ValueError(f"Invalid CSV file: {error}") from error
    return table


def format_meetings(meetings: list[Meeting]) -> str:
    """Format meetings as one line each, sorted by time.

    Args:
        meetings: Meetings to format.

    Returns:
        Multi-line string such as ``"09:30 - Anna"``.
    """
    ordered = sorted(meetings, key=lambda meeting: meeting.at)
    return "\n".join(f"{m.at.strftime(TIME_FORMAT)} - {m.who}" for m in ordered)


def fill_template(
    template: str | Path | BinaryIO,
    table: pd.DataFrame,
    practice_time: time,
    meetings: list[Meeting],
) -> bytes:
    """Fill the template and return the resulting workbook.

    Cells containing ``{{practice_time}}`` or ``{{individual_meetings}}`` are
    replaced. A cell holding only ``{{csv_table}}`` becomes the top-left corner of
    the CSV table (header row, then data rows).

    Args:
        template: Template workbook path or binary file object.
        table: CSV data to write.
        practice_time: Practice start time.
        meetings: Individual meetings.

    Returns:
        Bytes of the filled xlsx file.

    Raises:
        ValueError: If the template has no ``{{csv_table}}`` cell.
    """
    values = {
        "practice_time": practice_time.strftime(TIME_FORMAT),
        "individual_meetings": format_meetings(meetings),
    }
    workbook = load_workbook(template)
    sheet = workbook.active
    if sheet is None:
        raise ValueError("Template has no active sheet")

    anchor: tuple[int, int] | None = None
    for row in sheet.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            if cell.value.strip() == f"{{{{{TABLE_KEY}}}}}":
                anchor = (cell.row, cell.column)
                cell.value = None
            elif PLACEHOLDER_PATTERN.search(cell.value):
                cell.value = PLACEHOLDER_PATTERN.sub(
                    lambda match: values.get(match.group(1), match.group(0)),
                    cell.value,
                )
    if anchor is None:
        raise ValueError(f"Template has no {{{{{TABLE_KEY}}}}} cell")

    first_row, first_col = anchor
    clean = table.astype(object).where(table.notna(), None)
    for offset, name in enumerate(clean.columns):
        header = sheet.cell(first_row, first_col + offset, str(name))
        header.font = Font(bold=True)
    for row_offset, record in enumerate(clean.itertuples(index=False), start=1):
        for col_offset, value in enumerate(record):
            sheet.cell(first_row + row_offset, first_col + col_offset, value)

    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
