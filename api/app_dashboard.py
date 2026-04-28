"""Combined FastAPI app for ABAP ingest + dashboard.

Run::

    uvicorn api.app_dashboard:app --host 0.0.0.0 --port 8000

Endpoints:
    POST /ingest/errors          -- ABAP push (existing)
    GET  /ingest/errors/health   -- health (existing)
    GET  /healthz                -- top-level health
    GET  /api/stats              -- dashboard tiles
    GET  /api/errors             -- list view
    GET  /api/errors/{hash_key}  -- detail
    POST /api/feedback           -- thumbs up/down
    POST /api/classify-now/{h}   -- trigger classifier
"""
from __future__ import annotations

import logging

import structlog
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from config import settings
from api.ingest import router as ingest_router
from api.dashboard import router as dashboard_router

structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.processors.JSONRenderer(),
    ],
    wrapper_class=structlog.make_filtering_bound_logger(
        getattr(logging, settings.log_level.upper(), logging.INFO)
    ),
)
log = structlog.get_logger("dashboard_app")


app = FastAPI(title="Artilegenz Error Intelligence", version="1.0")

# CORS for localhost dev (frontend runs on :5173 by default with vite)
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
    ],
    allow_methods=["*"],
    allow_headers=["*"],
    allow_credentials=False,
)
app.include_router(ingest_router)
app.include_router(dashboard_router)


@app.get("/healthz")
def healthz() -> dict:
    return {"status": "ok", "db_engine": settings.db_engine}


@app.on_event("startup")
def _startup() -> None:
    log.info("dashboard_app_startup", db_engine=settings.db_engine)
