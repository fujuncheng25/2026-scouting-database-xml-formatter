# 2026 Scouting Database XML/Excel Formatter

A read-only web tool that connects to the `2026-scouting-backend` PostgreSQL database and renders `team_match_record` data in a human-friendly grid.

This repository is fully independent from `2026-scouting-backend` source code. It only reads that backend's database through environment-configured connection settings.

## Features

- Configure DB connection through `.env`
- Filter by `scoutEventId`
- Web UI pages by `matchNumber` (one page per match)
- Headers like `Red Alliance 254` / `Blue Alliance 971` in the same cell
- Bool output normalized from `0/1`/`true/false` to `No/Yes`
- Download Excel (`.xlsx`) with one sheet per `matchNumber`
- Download XML (`.xml`) with one `<Page>` per `matchNumber`
- Read-only behavior (`SELECT` only)
- No runtime/code dependency on `2026-scouting-backend`

## Quick Start

1. Install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Create env file:

```bash
cp .env.example .env
```

3. Fill DB values in `.env`.

4. Run:

```bash
python app.py
```

5. Open:

`http://localhost:5000`

## Notes

- The app auto-detects camelCase or snake_case column naming for `team_match_record` embedded fields.
- If embedded fields are stored as JSON objects (`autonomous`, `teleop`, `endAndAfterGame`), the app will also read nested values automatically.
- If your schema differs significantly from `2026-scouting-backend`, update candidate column names in `app.py`.
