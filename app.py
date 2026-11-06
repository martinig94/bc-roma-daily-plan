"""Streamlit interface: general plan in, individual schedules xlsx and PNGs out."""

import os
import smtplib
from datetime import date, time
from io import BytesIO
from pathlib import Path
from typing import cast

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from daily_plan.exports import save_outputs, zip_images
from daily_plan.fill import Meeting, build_schedules, fill_template, read_plan
from daily_plan.images import render_sections
from daily_plan.mailer import MailConfigError, send_images
from daily_plan.roster import load_roster

load_dotenv()

DEFAULT_TEMPLATE = Path(__file__).parent / "templates" / "template_daily_bc_roma.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DEFAULT_PRACTICE = time(11, 45)

roster = load_roster()
initials_by_surname = {player.surname: player.initials for player in roster}

st.set_page_config(
    page_title="BC Roma daily plan", page_icon=":material/sports_basketball:"
)
st.title("BC Roma daily plan")

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
            {"Who": pd.Series(dtype="str"), "Time": pd.Series(dtype="object")}
        ),
        num_rows="dynamic",
        column_config={
            "Who": st.column_config.SelectboxColumn(
                "Who", options=list(initials_by_surname)
            ),
            "Time": st.column_config.TimeColumn("Time", format="HH:mm"),
        },
        width="stretch",
        alt="Individual meetings: who and at what time",
    )
    zoom = st.slider("PNG zoom", 1.0, 3.0, 2.0, step=0.25)
    save_to_disk = st.checkbox("Save files automatically to the output folder", True)
    generate = st.button("Generate", type="primary", disabled=plan_file is None)

if generate and plan_file is not None:
    meetings = [
        Meeting(initials_by_surname[str(row.Who)], cast(time, row.Time))
        for row in meetings_table.dropna().itertuples()
    ]
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
        images = render_sections(workbook, zoom)
    except ValueError as error:
        st.error(str(error), title="Could not generate the schedules")
    else:
        stem = f"daily_schedule_{plan_date:%Y-%m-%d}"
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
    for name, data in result["images"].items():
        st.image(data, caption=name, alt=f"Section {name} of the daily schedule")

    with st.form("email"):
        recipient = st.text_input("Send images to", value=os.environ.get("MAIL_TO", ""))
        send = st.form_submit_button("Send by email", icon=":material/mail:")
    if send:
        try:
            send_images(
                result["images"],
                recipient,
                f"BC Roma daily schedule {plan_date:%d/%m/%Y}",
                "Daily schedule images attached.",
            )
        except (MailConfigError, smtplib.SMTPException, OSError) as error:
            st.error(str(error), title="Email not sent")
        else:
            st.success(f"Email sent to {recipient}")
