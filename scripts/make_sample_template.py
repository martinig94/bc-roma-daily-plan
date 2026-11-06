"""Generate a sample xlsx template with placeholders and section names."""

from pathlib import Path
from typing import cast

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet

OUTPUT = Path(__file__).parent.parent / "templates" / "daily_plan_template.xlsx"


def build_template() -> Workbook:
    """Build the sample template workbook.

    Returns:
        Workbook using ``{{practice_time}}``, ``{{individual_meetings}}`` and
        ``{{csv_table}}`` placeholders, with ``section_*`` defined names.
    """
    workbook = Workbook()
    sheet = cast(Worksheet, workbook.active)
    sheet.title = "Daily plan"
    sheet["A1"] = "BC Roma - Daily plan"
    sheet["A1"].font = Font(bold=True, size=16)
    sheet["A1"].fill = PatternFill("solid", fgColor="DDEBF7")
    sheet["A3"], sheet["B3"] = "Practice time", "{{practice_time}}"
    sheet["A4"], sheet["B4"] = "Individual meetings", "{{individual_meetings}}"
    sheet["B4"].alignment = Alignment(wrap_text=True, vertical="top")
    sheet["A6"] = "{{csv_table}}"
    sheet.column_dimensions["A"].width = 24
    for letter in "BCDEF":
        sheet.column_dimensions[letter].width = 20
    for name, reference in {
        "section_overview": "'Daily plan'!$A$1:$F$4",
        "section_schedule": "'Daily plan'!$A$6:$F$40",
    }.items():
        workbook.defined_names[name] = DefinedName(name, attr_text=reference)
    return workbook


if __name__ == "__main__":
    OUTPUT.parent.mkdir(exist_ok=True)
    build_template().save(OUTPUT)
    print(f"Wrote {OUTPUT}")
