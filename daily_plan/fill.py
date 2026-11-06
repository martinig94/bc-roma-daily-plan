"""Build individual schedules from the general plan and fill the xlsx template."""

import re
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, time
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO, cast

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from daily_plan.roster import Player

HEADER_LABEL = "TIME"
TEAM_PRACTICE = "TEAM PRACTICE"
MEETING = "MEETING"
DATA_ROWS = 5
CLEAR_ROWS = 8
TIME_FORMAT = "h\\.mm"
TITLE_CELL = "A1"
INITIALS_SPLIT = re.compile(r"[\s,-]+")
CellStyle = tuple[Any, Any]
DEFAULT_STYLE: tuple[CellStyle, CellStyle] = ((PatternFill(), Font()),) * 2


@dataclass(frozen=True)
class Activity:
    """One entry of a player's schedule.

    Attributes:
        at: Start time.
        name: Activity name (e.g. ``"PHYSIO"``).
    """

    at: time
    name: str


@dataclass(frozen=True)
class Meeting:
    """An individual meeting.

    Attributes:
        initials: Initials of the player met.
        at: Meeting start time.
    """

    initials: str
    at: time


def read_plan(source: str | Path | BinaryIO, filename: str) -> pd.DataFrame:
    """Read the general plan (csv or xlsx) without interpreting headers.

    Args:
        source: Path or binary file object.
        filename: Original file name, used to pick the parser.

    Returns:
        Raw dataframe, first row being the header row.

    Raises:
        ValueError: If the file is empty or cannot be parsed.
    """
    try:
        if filename.lower().endswith(".csv"):
            return pd.read_csv(source, header=None)
        return pd.read_excel(source, header=None)
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as error:
        raise ValueError(f"Invalid plan file: {error}") from error


def _to_time(value: object) -> time:
    """Convert ``9.5``, ``"11.45"``, ``"11:45"`` or a time object to ``time``."""
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    try:
        hundredths = round(float(str(value).replace(":", ".").replace(",", ".")) * 100)
        return time(*divmod(hundredths, 100))
    except ValueError as error:
        raise ValueError(f"Invalid time in plan: {value!r}") from error


def build_schedules(
    plan: pd.DataFrame,
    roster: list[Player],
    practice_time: time | None = None,
    meetings: list[Meeting] | None = None,
) -> dict[str, list[Activity]]:
    """Expand the general plan into one schedule per player.

    Plan columns are ``TIME`` followed by (``ACTIVITY``, ``PLAYERS``) pairs. An
    activity without players (team practice, lifting) applies to everyone.

    Args:
        plan: Raw plan dataframe from ``read_plan``.
        roster: Roster used to resolve initials.
        practice_time: Overrides the time of TEAM PRACTICE when given.
        meetings: Individual meetings added to the players' schedules.

    Returns:
        Mapping of initials to time-sorted activities (every roster player).

    Raises:
        ValueError: On unknown initials or invalid times.
    """
    schedules: dict[str, list[Activity]] = {player.initials: [] for player in roster}
    unknown: set[str] = set()
    for row in plan.itertuples(index=False):
        if pd.isna(row[0]) or str(row[0]).strip().upper() == HEADER_LABEL:
            continue
        at = _to_time(row[0])
        for index in range(1, len(row), 2):
            name = row[index]
            if pd.isna(name):
                continue
            label = str(name).strip().upper()
            cell = row[index + 1] if index + 1 < len(row) else None
            if pd.isna(cell):
                targets = list(schedules)
            else:
                targets = [i.upper() for i in INITIALS_SPLIT.split(str(cell)) if i]
                unknown.update(set(targets) - schedules.keys())
            slot_time = (
                practice_time if label == TEAM_PRACTICE and practice_time else at
            )
            for initials in targets:
                if initials in schedules:
                    schedules[initials].append(Activity(slot_time, label))
    if unknown:
        raise ValueError(f"Unknown player initials: {', '.join(sorted(unknown))}")

    for meeting in meetings or []:
        schedules[meeting.initials].append(Activity(meeting.at, MEETING))
    return {
        initials: sorted(activities, key=lambda activity: activity.at)
        for initials, activities in schedules.items()
    }


def _find_slots(sheet: Worksheet) -> list[tuple[int, int]]:
    """Return (row, column) of every ``TIME`` header cell in reading order."""
    return [
        (cell.row, cell.column)
        for row in sheet.iter_rows()
        for cell in row
        if cell.value == HEADER_LABEL
    ]


def _clone_last_group(sheet: Worksheet) -> None:
    """Copy the last row of player blocks below itself (styles and headers)."""
    header_rows = sorted({row for row, _ in _find_slots(sheet)})
    start = header_rows[-1]
    stride = start - header_rows[-2]
    height = sheet.max_row - start + 1
    for offset in range(height):
        source_row, target_row = start + offset, start + stride + offset
        sheet.row_dimensions[target_row].height = sheet.row_dimensions[
            source_row
        ].height
        for column in range(1, sheet.max_column + 1):
            source = sheet.cell(source_row, column)
            target = sheet.cell(target_row, column)
            target._style = copy(source._style)  # type: ignore[union-attr]
            target.value = HEADER_LABEL if source.value == HEADER_LABEL else None


def _activity_styles(
    sheet: Worksheet, slots: list[tuple[int, int]]
) -> dict[str, tuple[CellStyle, CellStyle]]:
    """Learn each activity's colors (fill, font) from the template's example rows.

    Returns:
        Mapping of activity name to the (time cell, activity cell) styles.
    """
    styles: dict[str, tuple[CellStyle, CellStyle]] = {}
    for row, column in slots:
        for target_row in range(row + 1, row + DATA_ROWS + 1):
            time_cell = sheet.cell(target_row, column)
            name_cell = sheet.cell(target_row, column + 1)
            if isinstance(name_cell.value, str):
                styles.setdefault(
                    name_cell.value.strip().upper(),
                    (
                        (copy(time_cell.fill), copy(time_cell.font)),
                        (copy(name_cell.fill), copy(name_cell.font)),
                    ),
                )
    return styles


def fill_template(
    template: str | Path | BinaryIO,
    schedules: dict[str, list[Activity]],
    roster: list[Player],
    plan_date: date,
) -> bytes:
    """Write the schedules into the template's player blocks.

    Each ``TIME`` header cell starts a block: the cell to its right gets the player
    title and the rows below get time/activity pairs. Players are assigned to blocks
    in roster order; extra rows of blocks are cloned when there are too few.

    Args:
        template: Template workbook path or binary file object.
        schedules: Output of ``build_schedules``.
        roster: Players in block order.
        plan_date: Date shown in the sheet title.

    Returns:
        Bytes of the filled xlsx file.

    Raises:
        ValueError: If the template has no block or a player has too many entries.
    """
    workbook = load_workbook(template)
    sheet = cast(Worksheet, workbook.active)
    slots = _find_slots(sheet)
    if not slots:
        raise ValueError(f"Template has no '{HEADER_LABEL}' block header")
    while len(slots) < len(roster):
        _clone_last_group(sheet)
        slots = _find_slots(sheet)

    sheet[TITLE_CELL] = f"DAILY SCHEDULE {plan_date:%d/%m/%Y}"
    styles = _activity_styles(sheet, slots)
    header_rows = sorted({row for row, _ in slots})
    for row, column in slots:
        next_header = min(
            (r for r in header_rows if r > row), default=sheet.max_row + 1
        )
        for target_row in range(row + 1, min(row + CLEAR_ROWS + 1, next_header)):
            for target_column in (column, column + 1):
                cell = sheet.cell(target_row, target_column)
                cell.value = None
                cell.fill = PatternFill()
        sheet.cell(row, column + 1).value = None

    for (row, column), player in zip(slots, roster, strict=False):
        activities = schedules[player.initials]
        if len(activities) > DATA_ROWS:
            raise ValueError(
                f"{player.surname} has {len(activities)} entries, max is {DATA_ROWS}"
            )
        sheet.cell(row, column + 1).value = player.header
        for offset, activity in enumerate(activities, start=1):
            time_cell = sheet.cell(row + offset, column)
            name_cell = sheet.cell(row + offset, column + 1)
            time_cell.value = activity.at
            time_cell.number_format = TIME_FORMAT
            name_cell.value = activity.name
            for cell, (fill, font) in zip(
                (time_cell, name_cell),
                styles.get(activity.name, DEFAULT_STYLE),
                strict=True,
            ):
                cell.fill, cell.font = copy(fill), copy(font)

    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
