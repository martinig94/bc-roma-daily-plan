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
from daily_plan.images import render_sections
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


def test_render_sections_zoom(plan, roster) -> None:  # type: ignore[no-untyped-def]
    """Each section yields a PNG, zoom enlarges it, zip holds them all."""
    filled = fill_template(
        TEMPLATE, build_schedules(plan, roster), roster, date(2026, 11, 6)
    )
    small, large = render_sections(filled, 1.0), render_sections(filled, 2.0)
    assert {"full", "players_1", "players_4"} <= set(small)
    assert all(data.startswith(PNG_MAGIC) for data in small.values())
    assert (
        Image.open(BytesIO(large["full"])).width
        > Image.open(BytesIO(small["full"])).width
    )
    archive = zipfile.ZipFile(BytesIO(zip_images(small)))
    assert sorted(archive.namelist()) == sorted(f"{n}.png" for n in small)
