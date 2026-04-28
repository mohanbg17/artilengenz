"""Minimal FastAPI app for the ABAP ingest smoke test.

Mounts only the ingest router. Use this to validate the SAP→FastAPI→DB loop
without dragging in Snowflake or the classifier.

Run::

    uvicorn api.app_ingest:app --port 8000

Once the loop is proven, switch to api.fastapi_app for the full surface
(triage, drilldown, analytics, classifier sweep).
"""
from __future__ import annotations

import logging

import structlog
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from config import settings
from api.ingest import router as ingest_router

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
log = structlog.get_logger("ingest_app")


app = FastAPI(title="Artilegenz Ingest API", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
app.include_router(ingest_router)


@app.get("/healthz")
def healthz() -> dict:
    return {"status": "ok", "db_engine": settings.db_engine}


@app.on_event("startup")
def _startup() -> None:
    log.info("ingest_app_startup", db_engine=settings.db_engine)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=settings.api_host, port=settings.api_port)
