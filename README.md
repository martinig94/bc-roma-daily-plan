# BC Roma daily plan

Streamlit app: upload a CSV, set practice time and individual meetings, fill an xlsx
template, download it, export section PNGs and email them.

## Run

```bash
uv sync
cp .env.example .env   # fill SMTP settings for email
uv run streamlit run app.py
```

## Template

Default: `templates/daily_plan_template.xlsx` (regenerate with
`uv run python scripts/make_sample_template.py`). Or upload your own in the app.

- `{{practice_time}}`, `{{individual_meetings}}`: replaced in any cell.
- `{{csv_table}}`: top-left cell where the CSV (header + rows) is written.
- Defined names starting with `section_` (e.g. `section_overview`) mark the ranges
  exported as PNGs. Without them the whole sheet is exported.

## Checks

```bash
uv run ruff check . && uv run mypy daily_plan app.py scripts && uv run pytest
```
