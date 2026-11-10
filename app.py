"""Streamlit interface: general plan in, individual schedules xlsx and PNGs out."""

import os
import smtplib
from datetime import date, time
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from daily_plan.exports import save_outputs, zip_images
from daily_plan.fill import (
    MEETING,
    MEETING_COLORS,
    Meeting,
    build_schedules,
    fill_template,
    read_plan,
)
from daily_plan.images import TEAM_PREFIX, render_images
from daily_plan.mailer import MailConfigError, send_images
from daily_plan.roster import load_roster

load_dotenv()

DEFAULT_TEMPLATE = Path(__file__).parent / "templates" / "template_daily_bc_roma.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DEFAULT_PRACTICE = time(11, 45)


def as_time(value: object) -> time:
    """Convert a time editor cell (``time`` or ``"HH:MM[:SS]"`` string) to ``time``.

    Args:
        value: Cell value returned by ``st.data_editor``.

    Returns:
        The value as a ``time`` object.

    Raises:
        ValueError: If the value is not a valid time string.
    """
    if isinstance(value, time):
        return value
    return time.fromisoformat(str(value))


def meetings_from_table(table: pd.DataFrame) -> list[Meeting]:
    """Build meetings from the editor table, skipping rows without who or time.

    Args:
        table: Editor output with ``Who``, ``Time``, ``Label`` and ``Color`` columns.

    Returns:
        One meeting per complete row; empty labels fall back to ``MEETING``.
    """
    return [
        Meeting(
            initials_by_surname[str(row.Who)],
            as_time(row.Time),
            row.Label if isinstance(row.Label, str) else MEETING,
            MEETING_COLORS.get(str(row.Color)),
        )
        for row in table.dropna(subset=["Who", "Time"]).itertuples()
    ]


roster = load_roster()
initials_by_surname = {player.surname: player.initials for player in roster}

st.set_page_config(
    page_title="BC Roma daily plan", page_icon=":material/sports_basketball:"
)
LOGO = Path(__file__).parent / "assets" / "logo.svg"
DESCRIPTION = (
    "Upload the general plan, set practice time and individual meetings, and get "
    "each player's daily schedule as xlsx and PNG images, ready to email."
)

_, logo_col, _ = st.columns([1, 2, 1])
logo_col.image(LOGO, width="stretch", alt="BC Roma logo")
st.title("BC Roma daily plan", text_alignment="center")
st.caption(DESCRIPTION, text_alignment="center")

with st.container(border=True):
    plan_file = st.file_uploader("General plan (csv or xlsx)", type=["csv", "xlsx"])
    template_file = st.file_uploader(
        "Template xlsx (optional, default template used if empty)", type="xlsx"
    )
    plan_date = st.date_input("Date", value=date.today(), format="DD/MM/YYYY")
    practice_time = st.time_input("Practice time", value=DEFAULT_PRACTICE)
    st.caption("Individual meetings")
    meetings_table = st.data_editor(
        pd.DataFrame(
            {
                "Who": pd.Series(dtype="str"),
                "Time": pd.Series(dtype="object"),
                "Label": pd.Series(dtype="str"),
                "Color": pd.Series(dtype="str"),
            }
        ),
        num_rows="dynamic",
        column_config={
            "Who": st.column_config.SelectboxColumn(
                "Who", options=list(initials_by_surname)
            ),
            "Time": st.column_config.TimeColumn("Time", format="HH:mm"),
            "Label": st.column_config.TextColumn(
                "Label", help=f"Text in the schedule, {MEETING} if empty"
            ),
            "Color": st.column_config.SelectboxColumn(
                "Color", options=list(MEETING_COLORS), help="Row color, optional"
            ),
        },
        width="stretch",
        alt="Individual meetings: who and at what time",
    )
    save_to_disk = st.checkbox("Save files automatically to the output folder", True)
    generate = st.button("Generate", type="primary", disabled=plan_file is None)

if generate and plan_file is not None:
    meetings = meetings_from_table(meetings_table)
    try:
        schedules = build_schedules(
            read_plan(plan_file, plan_file.name), roster, practice_time, meetings
        )
        workbook = fill_template(
            BytesIO(template_file.getvalue()) if template_file else DEFAULT_TEMPLATE,
            schedules,
            roster,
            plan_date,
        )
        images = render_images(workbook, plan_date)
    except ValueError as error:
        st.error(str(error), title="Could not generate the schedules")
    else:
        stem = f"{TEAM_PREFIX}_{plan_date:%Y-%m-%d}"
        st.session_state["result"] = {"stem": stem, "xlsx": workbook, "images": images}
        if save_to_disk:
            directory = Path(os.environ.get("OUTPUT_DIR", "output"))
            saved = save_outputs(directory, workbook, images, stem)
            st.success(f"Saved {len(saved)} files to {directory.resolve()}")

result = st.session_state.get("result")
if result:
    xlsx_col, zip_col = st.columns(2)
    xlsx_col.download_button(
        "Download xlsx",
        result["xlsx"],
        f"{result['stem']}.xlsx",
        XLSX_MIME,
        icon=":material/download:",
        width="stretch",
    )
    zip_col.download_button(
        "Download PNGs (zip)",
        zip_images(result["images"]),
        f"{result['stem']}_images.zip",
        "application/zip",
        icon=":material/image:",
        width="stretch",
    )
    with st.container(border=True):
        st.subheader("Send by email")
        recipient = st.text_input(
            "Recipient email", value=os.environ.get("MAIL_TO", ""), key="recipient"
        )
        send = st.button(
            "Send images by email",
            icon=":material/mail:",
            disabled=not recipient.strip(),
        )
        if send:
            try:
                send_images(
                    result["images"],
                    recipient.strip(),
                    f"BC Roma daily schedule {plan_date:%d/%m/%Y}",
                    "Daily schedule images attached.",
                )
            except (MailConfigError, smtplib.SMTPException, OSError) as error:
                st.error(str(error), title="Email not sent")
            else:
                st.success(f"Email sent to {recipient.strip()}")

    with st.expander(f"Preview images ({len(result['images'])})"):
        for name, data in result["images"].items():
            st.image(data, caption=name, alt=f"Schedule image {name}")
