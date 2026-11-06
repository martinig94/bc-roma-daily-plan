"""Tests for template filling, image rendering and exports."""

from datetime import time
from io import BytesIO

import pandas as pd
import pytest
from openpyxl import load_workbook
from PIL import Image

from daily_plan.exports import zip_images
from daily_plan.fill import Meeting, fill_template, format_meetings
from daily_plan.images import render_sections
from scripts.make_sample_template import build_template

PNG_MAGIC = b"\x89PNG"


@pytest.fixture
def template() -> BytesIO:
    """Sample template as an in-memory file."""
    buffer = BytesIO()
    build_template().save(buffer)
    buffer.seek(0)
    return buffer


@pytest.fixture
def filled(template: BytesIO) -> bytes:
    """Workbook filled with small sample data."""
    table = pd.DataFrame({"player": ["Anna", "Bo"], "load": [1.5, None]})
    meetings = [Meeting("Zed", time(10, 0)), Meeting("Anna", time(9, 30))]
    return fill_template(template, table, time(18, 0), meetings)


def test_format_meetings_sorted() -> None:
    """Meetings are listed in time order."""
    text = format_meetings([Meeting("B", time(11, 0)), Meeting("A", time(9, 5))])
    assert text == "09:05 - A\n11:00 - B"


def test_fill_template_values(filled: bytes) -> None:
    """Placeholders and table are written, NaN becomes empty."""
    sheet = load_workbook(BytesIO(filled)).active
    assert sheet["B3"].value == "18:00"
    assert sheet["B4"].value == "09:30 - Anna\n10:00 - Zed"
    assert [sheet.cell(6, c).value for c in (1, 2)] == ["player", "load"]
    assert sheet["A7"].value == "Anna"
    assert sheet["B8"].value is None


def test_missing_table_placeholder_raises() -> None:
    """A template without the table anchor is rejected."""
    from openpyxl import Workbook

    buffer = BytesIO()
    Workbook().save(buffer)
    buffer.seek(0)
    with pytest.raises(ValueError, match="csv_table"):
        fill_template(buffer, pd.DataFrame({"a": [1]}), time(9, 0), [])


def test_render_sections_zoom(filled: bytes) -> None:
    """Each section yields a PNG and zoom enlarges it."""
    small = render_sections(filled, 1.0)
    large = render_sections(filled, 2.0)
    assert set(small) == {"overview", "schedule"}
    assert all(data.startswith(PNG_MAGIC) for data in small.values())
    width_small = Image.open(BytesIO(small["schedule"])).width
    width_large = Image.open(BytesIO(large["schedule"])).width
    assert width_large > width_small


def test_zip_images_contains_files(filled: bytes) -> None:
    """The zip holds one png per section."""
    import zipfile

    archive = zipfile.ZipFile(BytesIO(zip_images(render_sections(filled))))
    assert sorted(archive.namelist()) == ["overview.png", "schedule.png"]
