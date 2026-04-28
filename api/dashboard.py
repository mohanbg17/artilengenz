"""Dashboard API endpoints.

Read-only endpoints that drive the frontend dashboard, plus a single POST
endpoint for thumbs up / thumbs down feedback.

All endpoints are unauthenticated (MVP scope: localhost only). Never expose
this router publicly without adding auth.

Mounted at /api/* by app_dashboard.py.
"""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

import psycopg2
import psycopg2.extras
import structlog
from fastapi import APIRouter, HTTPException, Query, status
from pydantic import BaseModel, Field

log = structlog.get_logger(__name__)

router = APIRouter(prefix="/api", tags=["dashboard"])


PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


# ===========================================================================
# Helpers
# ===========================================================================
def _conn():
    return psycopg2.connect(PG_DSN)


def _serialize_row(row: Dict[str, Any]) -> Dict[str, Any]:
    """Convert non-JSON types (datetime, Decimal) to strings/floats."""
    out: Dict[str, Any] = {}
    for k, v in row.items():
        if isinstance(v, datetime):
            out[k] = v.isoformat()
        elif hasattr(v, "__float__") and not isinstance(v, (int, float, bool)):
            out[k] = float(v)
        else:
            out[k] = v
    return out


# ===========================================================================
# /api/stats -- top-level dashboard tiles
# ===========================================================================
@router.get("/stats")
def get_stats() -> Dict[str, Any]:
    """Counts for the dashboard header."""
    with _conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("""
                SELECT
                    (SELECT COUNT(*) FROM raw.raw_errors)                                AS total_errors,
                    (SELECT COUNT(*) FROM intel.classifications WHERE status='classified') AS total_classified,
                    (SELECT COUNT(DISTINCT error_hash_key) FROM intel.classifications
                       WHERE status='classified')                                          AS unique_classified,
                    (SELECT COUNT(*) FROM intel.classifications WHERE status='failed')   AS total_failed,
                    (SELECT COUNT(*) FROM intel.corpus
                       WHERE embedding_version IS NOT NULL)                              AS corpus_embedded,
                    (SELECT MAX(occurred_at) FROM raw.raw_errors)                        AS latest_error_at;
            """)
            stats = _serialize_row(dict(cur.fetchone()))

            cur.execute("""
                SELECT severity, COUNT(*) AS n
                FROM raw.raw_errors
                WHERE severity IS NOT NULL
                GROUP BY severity ORDER BY 2 DESC;
            """)
            stats["by_severity"] = [_serialize_row(dict(r)) for r in cur.fetchall()]

            cur.execute("""
                SELECT badge, COUNT(*) AS n,
                       AVG(composite_confidence)::numeric(4,3) AS avg_conf
                FROM intel.classifications
                WHERE status='classified'
                GROUP BY badge ORDER BY 2 DESC;
            """)
            stats["by_badge"] = [_serialize_row(dict(r)) for r in cur.fetchall()]

            cur.execute("""
                SELECT source, COUNT(*) AS n
                FROM raw.raw_errors GROUP BY source ORDER BY 2 DESC;
            """)
            stats["by_source"] = [_serialize_row(dict(r)) for r in cur.fetchall()]

            cur.execute("""
                SELECT
                    SUM((sonnet_tokens->>'input_tokens')::int)                AS sonnet_input,
                    SUM((sonnet_tokens->>'output_tokens')::int)               AS sonnet_output,
                    SUM((sonnet_tokens->>'cache_read_input_tokens')::int)     AS sonnet_cache_reads,
                    SUM((opus_tokens->>'input_tokens')::int)                  AS opus_input,
                    SUM((opus_tokens->>'output_tokens')::int)                 AS opus_output
                FROM intel.classifications WHERE status='classified';
            """)
            t = cur.fetchone()
            if t:
                stats["tokens"] = _serialize_row(dict(t))

    return stats


# ===========================================================================
# /api/errors -- list view
# ===========================================================================
@router.get("/errors")
def list_errors(
    limit: int = Query(50, ge=1, le=500),
    offset: int = Query(0, ge=0),
    classified_only: bool = Query(False),
    severity: Optional[str] = Query(None),
    source: Optional[str] = Query(None),
    badge: Optional[str] = Query(None),
) -> Dict[str, Any]:
    """List errors with optional filters. Joins latest classification per error."""
    where: List[str] = ["1=1"]
    params: Dict[str, Any] = {"limit": limit, "offset": offset}

    if classified_only:
        where.append("c.classification_id IS NOT NULL")
    if severity:
        where.append("r.severity = %(severity)s")
        params["severity"] = severity
    if source:
        where.append("r.source = %(source)s")
        params["source"] = source
    if badge:
        where.append("c.badge = %(badge)s")
        params["badge"] = badge

    sql = f"""
        SELECT
            r.hash_key, r.source, r.system_id, r.occurred_at, r.severity,
            r.error_id, r.transaction, r.program, r.user_name,
            LEFT(r.short_text, 280) AS short_text_preview,
            c.classification_id, c.composite_confidence, c.badge,
            c.top_proposal_title, c.created_at AS classified_at
        FROM raw.raw_errors r
        LEFT JOIN intel.v_latest_classifications c
               ON c.error_hash_key = r.hash_key
        WHERE {' AND '.join(where)}
        ORDER BY r.occurred_at DESC
        LIMIT %(limit)s OFFSET %(offset)s;
    """

    with _conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(sql, params)
            rows = [_serialize_row(dict(r)) for r in cur.fetchall()]

            cur.execute(
                f"SELECT COUNT(*) AS total FROM raw.raw_errors r "
                f"LEFT JOIN intel.v_latest_classifications c ON c.error_hash_key=r.hash_key "
                f"WHERE {' AND '.join(where)};",
                {k: v for k, v in params.items() if k not in ("limit", "offset")},
            )
            total = cur.fetchone()["total"]

    return {"errors": rows, "total": total, "limit": limit, "offset": offset}


# ===========================================================================
# /api/errors/{hash_key} -- detail
# ===========================================================================
@router.get("/errors/{hash_key}")
def get_error_detail(hash_key: str) -> Dict[str, Any]:
    """Full raw error + latest classification + retrieved citations."""
    with _conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT * FROM raw.raw_errors WHERE hash_key = %s""",
                (hash_key,),
            )
            err = cur.fetchone()
            if not err:
                raise HTTPException(status_code=404, detail="error not found")
            err = _serialize_row(dict(err))

            cur.execute(
                """SELECT * FROM intel.classifications
                   WHERE error_hash_key = %s AND status='classified'
                   ORDER BY created_at DESC LIMIT 1""",
                (hash_key,),
            )
            cls_row = cur.fetchone()
            classification = _serialize_row(dict(cls_row)) if cls_row else None

            # Existing feedback for this classification (if any)
            feedback = None
            if classification:
                cur.execute(
                    """SELECT * FROM intel.human_feedback
                       WHERE classification_id = %s
                       ORDER BY reviewed_at DESC LIMIT 1""",
                    (classification["classification_id"],),
                )
                fb = cur.fetchone()
                if fb:
                    feedback = _serialize_row(dict(fb))

    return {
        "error": err,
        "classification": classification,
        "feedback": feedback,
    }


# ===========================================================================
# /api/feedback -- thumbs up / thumbs down
# ===========================================================================
class FeedbackPayload(BaseModel):
    classification_id: str
    accepted: bool                                      # True = thumbs up, False = thumbs down
    reviewer: str = Field(default="anonymous", max_length=128)
    comments: Optional[str] = Field(default=None, max_length=2000)
    corrected_solution: Optional[str] = Field(default=None, max_length=8000)


@router.post("/feedback", status_code=status.HTTP_201_CREATED)
def submit_feedback(payload: FeedbackPayload) -> Dict[str, Any]:
    """Insert a feedback record, returning the new feedback_id."""
    with _conn() as conn:
        with conn.cursor() as cur:
            # Verify classification exists
            cur.execute(
                "SELECT 1 FROM intel.classifications WHERE classification_id = %s",
                (payload.classification_id,),
            )
            if not cur.fetchone():
                raise HTTPException(
                    status_code=404,
                    detail=f"classification {payload.classification_id} not found",
                )

            # Generate feedback_id
            seed = f"{payload.classification_id}:{payload.reviewer}:{datetime.utcnow().isoformat()}"
            feedback_id = hashlib.sha256(seed.encode()).hexdigest()[:32]

            cur.execute(
                """INSERT INTO intel.human_feedback
                     (feedback_id, classification_id, reviewer, reviewed_at,
                      accepted, corrected_solution, comments)
                   VALUES (%s, %s, %s, CURRENT_TIMESTAMP, %s, %s, %s);""",
                (feedback_id, payload.classification_id, payload.reviewer,
                 payload.accepted, payload.corrected_solution, payload.comments),
            )
            conn.commit()

    log.info("feedback_recorded",
             classification_id=payload.classification_id,
             accepted=payload.accepted, reviewer=payload.reviewer)
    return {"feedback_id": feedback_id, "status": "recorded"}


# ===========================================================================
# /api/classify-now -- trigger classifier on demand (synchronous)
# ===========================================================================
@router.post("/classify-now/{hash_key}")
def classify_now(hash_key: str) -> Dict[str, Any]:
    """Run the classifier synchronously on an error. ~30-60 sec."""
    # Lazy import so this module can load without classifier deps when only used for reads
    try:
        import sys
        # Ensure classifier dir is on path
        proj = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        clf_path = os.path.join(proj, "classifier")
        if clf_path not in sys.path:
            sys.path.insert(0, clf_path)
        from classifier_core import classify  # noqa: WPS433
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(
            status_code=500,
            detail=f"classifier module unavailable: {exc}",
        )

    try:
        result = classify(hash_key)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except Exception as exc:  # noqa: BLE001
        log.exception("classify_now_failed", hash_key=hash_key[:16] + "...")
        raise HTTPException(status_code=500, detail=f"classification failed: {exc}")

    return {
        "classification_id": result.classification_id,
        "confidence": result.final_confidence,
        "badge": result.badge,
        "summary_md": result.final_summary_md,
        "root_cause": result.final_root_cause,
        "remediation_steps": result.final_remediation,
        "severity": result.final_severity,
    }
