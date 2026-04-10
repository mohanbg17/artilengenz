"""FastAPI application entrypoint for the NLP-SAP Query Engine."""

from __future__ import annotations

import logging
import sys
from contextlib import asynccontextmanager
from typing import AsyncGenerator

import structlog
from fastapi import FastAPI, Request, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from api.routes import router
from nlp_sap.config import get_settings
from nlp_sap.orchestrator import QueryOrchestrator

# ── Logging setup ─────────────────────────────────────────────────────────────
settings = get_settings()

structlog.configure(
    wrapper_class=structlog.make_filtering_bound_logger(
        getattr(logging, settings.log_level.upper(), logging.INFO)
    )
)
logging.basicConfig(stream=sys.stdout, level=settings.log_level.upper())
logger = structlog.get_logger()


# ── Application lifespan ──────────────────────────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncGenerator[None, None]:
    logger.info("Starting NLP-SAP Query Engine", env=settings.app_env, mock=settings.mock_sap)
    orchestrator = QueryOrchestrator()
    app.state.orchestrator = orchestrator
    yield
    await orchestrator.close()
    logger.info("NLP-SAP Query Engine shut down")


# ── FastAPI app ───────────────────────────────────────────────────────────────
app = FastAPI(
    title="NLP-SAP Query Engine",
    description=(
        "Natural language interface to SAP S4/HANA and ECC for finance "
        "and logistics reports. Powered by Claude AI."
    ),
    version="0.1.0",
    lifespan=lifespan,
    docs_url="/docs",
    redoc_url="/redoc",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"] if settings.app_env == "development" else [],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

app.include_router(router, prefix="/api/v1")


# ── Global exception handler ──────────────────────────────────────────────────
@app.exception_handler(Exception)
async def unhandled_exception(request: Request, exc: Exception) -> JSONResponse:
    logger.error("Unhandled exception", path=request.url.path, error=str(exc))
    return JSONResponse(
        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        content={"detail": "Internal server error", "error": str(exc)},
    )


@app.get("/health", tags=["health"])
async def health() -> dict:
    """Health check endpoint."""
    return {"status": "ok", "mock_sap": settings.mock_sap, "version": "0.1.0"}
