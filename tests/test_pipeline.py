"""Tests for schedule building, template filling, rendering and exports."""

import zipfile
from datetime import date, time
from io import BytesIO
from pathlib import Path
from typing import cast

import pytest
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image

from daily_plan.exports import zip_images
from daily_plan.fill import (
    Activity,
    Meeting,
    build_schedules,
    fill_template,
    read_plan,
)
from daily_plan.images import render_images
from daily_plan.roster import load_roster

PLAN = Path(__file__).parent / "data" / "611.xlsx"
TEMPLATE = Path(__file__).parent.parent / "templates" / "template_daily_bc_roma.xlsx"
PNG_MAGIC = b"\x89PNG"


@pytest.fixture
def roster():  # type: ignore[no-untyped-def]
    """Roster loaded from the repo file."""
    return load_roster()


@pytest.fixture
def plan():  # type: ignore[no-untyped-def]
    """Raw general plan of 6 November."""
    return read_plan(PLAN, PLAN.name)


def test_build_schedules_expands_players(plan, roster) -> None:  # type: ignore[no-untyped-def]
    """Watson gets his group activities plus the team rows."""
    schedules = build_schedules(plan, roster)
    assert [(a.at, a.name) for a in schedules["PW"]] == [
        (time(10, 20), "PHYSIO"),
        (time(11, 0), "PREHAB"),
        (time(11, 20), "PDC"),
        (time(11, 45), "TEAM PRACTICE"),
        (time(13, 15), "TEAM LIFTING"),
    ]
    assert len(schedules) == len(roster)


def test_practice_time_and_meeting(plan, roster) -> None:  # type: ignore[no-untyped-def]
    """Practice time overrides the plan and meetings are added."""
    schedules = build_schedules(plan, roster, time(12, 0), [Meeting("NM", time(9, 0))])
    assert schedules["NM"][0] == Activity(time(9, 0), "MEETING")
    assert Activity(time(12, 0), "TEAM PRACTICE") in schedules["NM"]


def test_unknown_initials_raise(plan, roster) -> None:  # type: ignore[no-untyped-def]
    """Initials missing from the roster are reported."""
    with pytest.raises(ValueError, match="Unknown player initials"):
        build_schedules(plan, roster[:-1])


def test_fill_template(plan, roster) -> None:  # type: ignore[no-untyped-def]
    """Blocks are filled, old data cleared and extra players get cloned blocks."""
    filled = fill_template(
        TEMPLATE, build_schedules(plan, roster), roster, date(2026, 11, 6)
    )
    sheet = cast(Worksheet, load_workbook(BytesIO(filled)).active)
    assert sheet["A1"].value == "DAILY SCHEDULE 06/11/2026"
    assert sheet["C5"].value == "#7 EMEJURU"
    assert sheet["F5"].value == "#10 STEVENSON"
    assert sheet["F6"].value == "PHYSIO"
    assert sheet["F9"].value == "TEAM LIFTING"
    assert sheet["F10"].value is None
    headers = [c.value for row in sheet.iter_rows() for c in row if c.value]
    assert "OLYNYK" in headers and "#9 FLOWERS" in headers


def test_render_images_names_and_zoom(plan, roster) -> None:  # type: ignore[no-untyped-def]
    """One team image plus one zoom per player, named by player and date."""
    day = date(2026, 11, 6)
    filled = fill_template(TEMPLATE, build_schedules(plan, roster), roster, day)
    small, large = render_images(filled, day, 1.0), render_images(filled, day, 2.0)
    assert len(small) == len(roster) + 1
    assert "bc_roma_daily_2026-11-06" in small
    assert {"Mannion_2026-11-06", "Olynyk_2026-11-06", "Emejuru_2026-11-06"} <= set(
        small
    )
    assert all(data.startswith(PNG_MAGIC) for data in small.values())
    key = "Watson_2026-11-06"
    assert Image.open(BytesIO(large[key])).width > Image.open(BytesIO(small[key])).width
    archive = zipfile.ZipFile(BytesIO(zip_images(small)))
    assert sorted(archive.namelist()) == sorted(f"{n}.png" for n in small)


def test_meeting_label_and_color(plan, roster) -> None:  # type: ignore[no-untyped-def]
    """A meeting keeps its label (upper-cased) and fill color."""
    meeting = Meeting("NM", time(9, 0), "video review", "FFF2CC")
    schedules = build_schedules(plan, roster, None, [meeting])
    assert schedules["NM"][0] == Activity(time(9, 0), "VIDEO REVIEW", "FFF2CC")


def test_many_entries_grow_blocks(plan, roster) -> None:  # type: ignore[no-untyped-def]
    """Players with more than 5 entries push the following blocks down."""
    day = date(2026, 11, 6)
    meetings = [Meeting("PW", time(8, minute), "TALK", "BDD7EE") for minute in range(4)]
    schedules = build_schedules(plan, roster, None, meetings)
    assert len(schedules["PW"]) == 9
    filled = fill_template(TEMPLATE, schedules, roster, day)
    sheet = cast(Worksheet, load_workbook(BytesIO(filled)).active)
    values = [c.value for row in sheet.iter_rows() for c in row if c.value]
    assert values.count("TALK") == 4
    assert "#3 WATSON" in values and "#13 ODZEBE" in values
    assert sheet["L24"].fill.fgColor.rgb.endswith("BDD7EE")
    images = render_images(filled, day, 1.0)
    assert (
        Image.open(BytesIO(images["Watson_2026-11-06"])).height
        > Image.open(BytesIO(images["Mannion_2026-11-06"])).height
    )
