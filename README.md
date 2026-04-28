# sap-error-intelligence

AI/vector layer for the SAP Error Intelligence Platform.

Scrapes a 10K+ corpus of SAP error-resolution pairs from public communities, embeds them with Voyage AI, indexes in Pinecone, and serves a Claude-powered classifier that proposes ranked solutions with calibrated confidence scores. Includes self-critique and auto-retry below 70% confidence.

---

## What this repo does

| Stage | Module |
|---|---|
| Scrape SAP Community, SO, Reddit, blogs | `scraper/*.py` |
| Build & filter corpus | `scraper/corpus_builder.py` |
| Embed via Voyage voyage-3 | `embeddings/voyage_client.py` |
| Index in Pinecone | `embeddings/pinecone_index.py` |
| Single-pass Claude classifier | `classifier/single_pass.py` |
| Self-critique pass | `classifier/self_critique.py` |
| Auto-retry below 70% | `classifier/retry_policy.py` |
| Composite confidence | `classifier/confidence.py` |
| Orchestrator (full pipeline) | `classifier/orchestrator.py` |
| REST API | `api/fastapi_app.py` |
| Dashboard | `frontend/` (React + Vite + Tailwind) |

---

## Architecture

```
        ┌─────────────────────────────────────────┐
        │  Public Sources (10K corpus, one-time)  │
        │  SAP Community + SO + Reddit + blogs    │
        └──────────────────┬──────────────────────┘
                           │
                  ┌────────▼────────┐
                  │ corpus_builder  │  quality filter, dedup, redact
                  └────────┬────────┘
                           │
                  ┌────────▼────────┐
                  │  Voyage voyage-3│  1024-dim
                  └────────┬────────┘
                           │
                  ┌────────▼────────┐
                  │   Pinecone      │  cosine, sap-errors namespace
                  └─────────────────┘

        ┌─────────────────────────────────────────┐
        │   New error in Snowflake.RAW_ERRORS     │  (from sap-error-ingestion)
        └──────────────────┬──────────────────────┘
                           │
                  ┌────────▼────────┐
                  │  Voyage embed   │
                  └────────┬────────┘
                           │
                  ┌────────▼────────┐
                  │  Pinecone top-K │  K=8 default, K=20 on retry
                  └────────┬────────┘
                           │
                  ┌────────▼────────┐
                  │  Claude pass 1  │  → ranked proposals + self-confidence
                  └────────┬────────┘
                           │
                  ┌────────▼────────┐
                  │  Claude pass 2  │  → critique, adjust, add/remove
                  └────────┬────────┘
                           │
              ┌────────────▼────────────┐
              │   composite confidence  │  0.45 vec + 0.45 model + 0.10 corp
              └────────────┬────────────┘
                           │
                ┌──────────▼──────────┐
                │  composite < 0.70?  │
                └──┬───────────────┬──┘
            yes   │               │   no
                  ▼               ▼
            retry K=20      land in CLASSIFICATIONS
                  │               │
                  └───────────────┘
                           │
                  ┌────────▼────────┐
                  │  FastAPI + React│  triage / drill-down / analytics
                  └─────────────────┘
```

---

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env  # fill in keys
```

### Build the corpus (one-time, ~3-6 hours)

```bash
# Scrape (writes to local SQLite + Snowflake INTEL.CORPUS)
python -m scraper.corpus_builder --target 10000

# Embed and index
python -m embeddings.indexer --reindex-all
```

### Run the API

```bash
uvicorn api.fastapi_app:app --reload --port 8000
```

### Run the dashboard

```bash
cd frontend && npm install && npm run dev
```

---

## Notes on scraping ethics

The user has explicitly accepted ToS risk for aggressive scraping. The implementation is still polite by default (rate-limited, robots.txt-aware, exponential backoff on 429s) — these are practical reliability measures, not ethical compromises. Before any commercial use, revisit ToS for each source.
