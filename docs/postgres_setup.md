# Local Postgres setup (Docker)

You're running Postgres locally instead of Snowflake. Single command brings the database up, schema initializes automatically, takes about 30 seconds.

## Prerequisites

- Docker Desktop running on your laptop
- Python venv already set up (you've done this)

## Step 1: Start Postgres

From the `sap-error-intelligence` folder:

```cmd
cd C:\sap-error-ai\sap-error-intelligence
docker compose up -d
```

What this does:

- Pulls the official `postgres:16` image (~150 MB, one-time)
- Creates a container named `artilegenz-postgres` listening on `localhost:5432`
- Auto-runs `db/postgres_ddl.sql` on first start, creating both schemas and all tables
- Persists data in a Docker volume named `artilegenz_postgres_data` — survives container restarts

Wait ~30 seconds for it to be healthy:

```cmd
docker compose ps
```

You want `STATUS` to show `Up (healthy)`. If it says `Up` without `(healthy)`, give it another 10 seconds.

## Step 2: Verify the schema is in place

```cmd
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "\dt raw.*"
```

Should show:

```
              List of relations
 Schema |    Name     | Type  |   Owner
--------+-------------+-------+------------
 raw    | raw_errors  | table | artilegenz
 raw    | watermarks  | table | artilegenz
```

And:

```cmd
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "\dt intel.*"
```

Should show four tables: `corpus`, `classifications`, `critiques`, `human_feedback`.

## Step 3: Configure `.env`

In `C:\sap-error-ai\sap-error-intelligence\.env`, set:

```
DB_ENGINE=postgres

PG_HOST=localhost
PG_PORT=5432
PG_USER=artilegenz
PG_PASSWORD=artilegenz_local_dev
PG_DATABASE=sap_errors

ARTILEGENZ_INGEST_TOKEN=<paste your generated token>
```

You can leave `ANTHROPIC_API_KEY`, `VOYAGE_API_KEY`, `PINECONE_API_KEY` blank for now — they're only needed when you run the classifier.

## Step 4: Test the FastAPI ingest endpoint

In your activated venv:

```cmd
.venv\Scripts\activate
uvicorn api.app_ingest:app --port 8000
```

You should see:

```
INFO:     Started server process [...]
INFO:     Uvicorn running on http://0.0.0.0:8000
```

In a second cmd window:

```cmd
curl http://localhost:8000/healthz
curl http://localhost:8000/ingest/errors/health
```

Both should return JSON. The second one should include `"db_engine":"postgres"` and `"auth_configured":"yes"`.

## Step 5: Smoke-test the ingest endpoint

Send a fake error record:

```cmd
curl -X POST http://localhost:8000/ingest/errors ^
  -H "Content-Type: application/json" ^
  -H "Authorization: Bearer <your-token>" ^
  -d "{\"system_id\":\"S4P\",\"client\":\"100\",\"host\":\"test-laptop\",\"batch_size\":1,\"records\":[{\"source\":\"ST22\",\"error_id\":\"TEST-001\",\"occurred_at\":\"20260425T120000Z\",\"severity\":\"ERROR\",\"short_text\":\"smoke test record\"}]}"
```

Expected response:

```json
{"status":"ok","received":1,"inserted":1,"duplicates":0}
```

Run the **exact same command again** — you should now see `"inserted":0,"duplicates":1`. That's the dedup working correctly.

Verify in Postgres:

```cmd
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT hash_key, source, error_id, short_text FROM raw.raw_errors;"
```

You should see one row with the test record.

## What's running where

```
laptop ─┬─ Docker container (artilegenz-postgres)  ← stores all data
        ├─ FastAPI on :8000                         ← receives ABAP POSTs
        └─ ngrok                                    ← exposes :8000 publicly
```

ABAP on the SAP server POSTs to the ngrok URL → FastAPI → backend abstraction → Postgres.

## Common operations

**Stop Postgres** (data persists):
```cmd
docker compose stop
```

**Start it again**:
```cmd
docker compose start
```

**Reset everything** (loses data):
```cmd
docker compose down -v
docker compose up -d
```

**Connect to psql interactively**:
```cmd
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors
```

Type `\q` to exit, `\dt schema.*` to list tables, `\d table_name` to describe one.

**Logs from Postgres**:
```cmd
docker compose logs -f postgres
```

## Switching to Snowflake later

When your Snowflake account is reactivated, change `.env`:

```
DB_ENGINE=snowflake
```

…and fill in the `SF_*` variables. Restart FastAPI. Same code, same SQL surface, same Python imports — the backend abstraction handles the rest.

(Note: the current Snowflake backend implementation has stubs for some methods; we'll restore the full Snowflake path when you're ready to switch.)
