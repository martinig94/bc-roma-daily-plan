# BC Roma daily plan

Streamlit app: upload the general plan (csv or xlsx, e.g. `611.xlsx`), set the date,
practice time and individual meetings, and get the individual schedules sheet filled
from `templates/template_daily_bc_roma.xlsx`. Download the xlsx, export section PNGs
(auto-saved to `output/`) and email them.

## Run

```bash
uv sync
cp .env.example .env   # fill SMTP settings for email
uv run streamlit run app.py
```

## How it works

- Plan columns: `TIME`, then (`ACTIVITY`, `PLAYERS`) pairs. Times are `hh.mm`
  (`9.5` means 9.50). Initials are joined with `-`. A row with an activity and no
  players (TEAM PRACTICE, TEAM LIFTING) applies to everyone.
- `templates/roster.csv` maps initials to jersey number and surname, in block order.
  Edit it when the roster changes.
- Each `TIME` header in the template starts a player block (5 rows). Colors per
  activity are learned from the template's example rows. Extra blocks are cloned
  when the roster has more players than blocks.
- "Practice time" overrides the time of TEAM PRACTICE. Each individual meeting adds a
  MEETING row to that player.
- PNGs: `bc_roma_daily_<date>.png` for the whole team plus a zoom on each player
  block named `<Surname>_<date>.png` (date as `YYYY-MM-DD`). Colors come from the
  template's own palette.

## Email setup

Fill these in `.env` (never commit it): `SMTP_HOST`, `SMTP_PORT` (587 STARTTLS or
465 SSL), `SMTP_USER`, `SMTP_PASSWORD`, `MAIL_FROM`, and optionally `MAIL_TO` as the
default recipient. Gmail needs an app password. Then generate a plan in the app and
use "Send by email". The email carries all the PNGs as attachments.

## Checks

```bash
uv run ruff check . && uv run mypy daily_plan app.py tests && uv run pytest
```
