# 2026 Scouting Database Local-First Sheet Editor

This app now uses a local-first architecture:

- `team_match_record` is fetched from a formal remote PostgreSQL DB (read-only source).
- Remote data is synced into this project's own local PostgreSQL cache.
- Users edit sheet cells directly in-browser (Excel-like table editing).
- Autosave writes edits only to local DB (`local_edits`) and never to the remote DB.
- Merge rule is time-layer based: later changes win. If remote updates are newer after a sync, they override older local edits.

## What Was Added

- Dual database connections:
	- remote source DB (`REMOTE_POSTGRES_*`)
	- local editable DB (`LOCAL_POSTGRES_*`)
- Local schema management on startup:
	- `local_snapshots`
	- `local_edits`
	- `local_users`
- Sync API:
	- `POST /api/sync`
- Cell autosave API:
	- `POST /api/edit`
- Optional URL-based auth mode:
	- set `REQUIRE_AUTH=true`
	- generate random users with `POST /auth/bootstrap`
	- use links with `?u=<username>&p=<password>`
- Export (`/export/excel`, `/export/xml`) now reads from effective local-first merged view.
- Docker support with app + local PostgreSQL via `docker-compose.yml`.

## Merge Behavior

The effective sheet value is computed per cell/field:

1. Start from latest synced remote snapshot (`local_snapshots.payload`).
2. If local edit timestamp is newer than or equal to remote record timestamp, show local edit.
3. If remote record timestamp is newer, remote value wins.

This matches the rule: later layer has higher priority.

## Local Run (without Docker)

1. Install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Configure env:

```bash
cp .env.example .env
```

3. Start app:

```bash
python app.py
```

4. Open:

`http://localhost:5000`

## Docker Run

1. Copy env:

```bash
cp .env.example .env
```

2. Set your remote DB credentials in `.env` (`REMOTE_POSTGRES_*`).

If the remote PostgreSQL runs on your host machine instead of another network host, set `REMOTE_POSTGRES_HOST=host.docker.internal` when using Docker.

3. Start app + local PostgreSQL:

```bash
docker compose up --build
```

4. Open:

`http://localhost:5000`

## Key Env Variables

- `REMOTE_POSTGRES_HOST`, `REMOTE_POSTGRES_PORT`, `REMOTE_POSTGRES_USER`, `REMOTE_POSTGRES_PASSWORD`, `REMOTE_POSTGRES_DB`
- `LOCAL_POSTGRES_HOST`, `LOCAL_POSTGRES_PORT`, `LOCAL_POSTGRES_USER`, `LOCAL_POSTGRES_PASSWORD`, `LOCAL_POSTGRES_DB`
- `REQUIRE_AUTH` (`false` by default)
- `PORT` (default `5000`)

## Notes

- If `REQUIRE_AUTH=false`, auth links are optional.
- The UI includes a "Create Random User URL" button for quick collaborator links.
- This project still does not write to the formal remote database.
- If remote sync is unavailable, the UI falls back to local cached data instead of crashing.
