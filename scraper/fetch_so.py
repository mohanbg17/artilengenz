"""Stack Overflow scraper for SAP knowledge corpus.

Uses the Stack Exchange API (https://api.stackexchange.com) to fetch
SAP-tagged questions with accepted answers.

Filters for:
  * Has an accepted answer
  * Question score >= 1 (filter out spam/garbage only)
  * Body length >= 100 chars (filter out trivial questions)
  * English language
  * Created in last 5 years (more relevant to S/4HANA era)

NO keyword filtering -- we want broad SAP knowledge, including
"how does X work" content that helps the classifier reason.

Writes results to a JSON file for downstream cleaning + loading.
No auth required (unauth: 300 req/day quota).

Usage:
    python fetch_so.py --target 500 --output ./scraped_so.json
    python fetch_so.py --target 50 --output ./smoke.json     # quick smoke test
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

import requests
import structlog

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(format="%(message)s", stream=sys.stdout, level=logging.INFO)
structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.dev.ConsoleRenderer(),
    ],
)
log = structlog.get_logger("fetch_so")

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
SO_API_BASE = "https://api.stackexchange.com/2.3"
SO_API_KEY = os.getenv("SO_API_KEY", "")
PAGE_SIZE = 100

# Tags to query (broadest first; we'll page through each)
SAP_TAGS = [
    "abap",         # largest pool, ABAP development
    "sap",          # general SAP umbrella
    "sapui5",       # frontend / Fiori
    "sap-fiori",
    "sap-erp",
    "sap-hana",
    "hana-sql-script",
    "sap-basis",
    "sap-bw",
    "sap-cloud-platform",
    "sap-cap",
    "sap-rap",
]

# Time window: last 5 years
NOW_UNIX = int(datetime.now(timezone.utc).timestamp())
FIVE_YEARS_UNIX = int((datetime.now(timezone.utc) - timedelta(days=5 * 365)).timestamp())

# Politeness
REQUEST_DELAY_SEC = 1.0


# ---------------------------------------------------------------------------
# API client
# ---------------------------------------------------------------------------
class SOClient:
    def __init__(self, api_key: Optional[str] = None) -> None:
        self.session = requests.Session()
        self.api_key = api_key or SO_API_KEY
        self.quota_remaining: Optional[int] = None

    def _request(self, path: str, params: Dict[str, Any]) -> Dict[str, Any]:
        url = f"{SO_API_BASE}/{path}"
        params["site"] = "stackoverflow"
        if self.api_key:
            params["key"] = self.api_key
        time.sleep(REQUEST_DELAY_SEC)
        resp = self.session.get(url, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        self.quota_remaining = data.get("quota_remaining")
        if data.get("backoff"):
            log.warning("so_backoff_requested", backoff_sec=data["backoff"])
            time.sleep(data["backoff"])
        return data

    def search_questions(
        self,
        tag: str,
        page: int = 1,
        from_unix: int = FIVE_YEARS_UNIX,
        to_unix: int = NOW_UNIX,
    ) -> Dict[str, Any]:
        """Search /questions for tagged + accepted-answer questions."""
        params = {
            "page": page,
            "pagesize": PAGE_SIZE,
            "tagged": tag,
            "fromdate": from_unix,
            "todate": to_unix,
            "order": "desc",
            "sort": "votes",
            "filter": "withbody",
        }
        return self._request("questions", params)

    def fetch_answers_bulk(self, answer_ids: List[int]) -> List[Dict[str, Any]]:
        """Fetch full content for a batch of answer IDs (max 100 per call)."""
        if not answer_ids:
            return []
        chunks = [answer_ids[i : i + 100] for i in range(0, len(answer_ids), 100)]
        out: List[Dict[str, Any]] = []
        for chunk in chunks:
            params = {
                "pagesize": PAGE_SIZE,
                "filter": "withbody",
            }
            ids_str = ";".join(str(a) for a in chunk)
            data = self._request(f"answers/{ids_str}", params)
            out.extend(data.get("items", []))
        return out


# ---------------------------------------------------------------------------
# Quality filters -- relaxed
# ---------------------------------------------------------------------------
def passes_quality_filter(question: Dict[str, Any]) -> bool:
    """Relaxed: must be answered (accepted), score >= 1, body >= 100 chars."""
    if not question.get("is_answered"):
        return False
    if not question.get("accepted_answer_id"):
        return False
    if (question.get("score") or 0) < 1:
        return False

    # Body length check (filter out trivial questions)
    q_body = question.get("body") or ""
    q_body_text = re.sub(r"<[^>]+>", " ", q_body)
    if len(q_body_text) < 100:
        return False

    # Exclude obviously off-topic content
    title = (question.get("title") or "").lower()
    if any(skip in title for skip in [
        "interview", "salary", "career", "best book", "where to learn",
        "should i learn", "how to start", "free course", "best course",
        "certification cost", "job opportunity", "future of",
    ]):
        return False

    return True


def passes_answer_quality(answer: Dict[str, Any]) -> bool:
    """Validate the accepted answer is substantive."""
    if not answer.get("is_accepted"):
        return False
    if (answer.get("score") or 0) < 1:
        return False
    body = answer.get("body") or ""
    body_text = re.sub(r"<[^>]+>", " ", body)
    if len(body_text) < 80:
        return False
    return True


# ---------------------------------------------------------------------------
# Module classifier (lightweight tag-based)
# ---------------------------------------------------------------------------
MODULE_TAG_MAP = {
    "abap":               "ABAP",
    "abap-oo":            "ABAP",
    "sap-basis":          "BASIS",
    "sap-fiori":          "FIORI",
    "sapui5":             "FIORI",
    "sap-hana":           "HANA",
    "hana":               "HANA",
    "hana-sql-script":    "HANA",
    "sap-bw":             "BW",
    "sap-bw-on-hana":     "BW",
    "sap-erp":            "ERP",
    "fico":               "FI",
    "sap-fi":             "FI",
    "sap-mm":             "MM",
    "sap-sd":             "SD",
    "sap-pi":             "PI",
    "sap-cpi":            "CPI",
    "sap-pp":             "PP",
    "sap-ps":             "PS",
    "sap-cloud-platform": "BTP",
    "sap-btp":            "BTP",
    "sap-rap":            "RAP",
    "sap-cap":            "CAP",
}


def classify_module(tags: List[str]) -> str:
    for t in tags:
        if t in MODULE_TAG_MAP:
            return MODULE_TAG_MAP[t]
    if "sap" in tags:
        return "GENERAL"
    return "UNKNOWN"


# ---------------------------------------------------------------------------
# Main scrape loop
# ---------------------------------------------------------------------------
def scrape(target: int, output_path: str) -> None:
    client = SOClient()
    collected: Dict[int, Dict[str, Any]] = {}

    log.info("scrape_starting", target=target, tags=SAP_TAGS, api_key=bool(client.api_key))

    for tag in SAP_TAGS:
        if len(collected) >= target:
            log.info("target_reached_breaking", count=len(collected))
            break

        log.info("tag_starting", tag=tag, collected_so_far=len(collected))
        page = 1
        max_pages = 25
        consecutive_empty_pages = 0

        while page <= max_pages and len(collected) < target:
            try:
                data = client.search_questions(tag=tag, page=page)
            except requests.HTTPError as e:
                log.warning("so_http_error", tag=tag, page=page, err=str(e))
                break
            except Exception as e:  # noqa: BLE001
                log.warning("so_request_failed", tag=tag, page=page, err=str(e))
                break

            items = data.get("items", [])
            has_more = data.get("has_more", False)
            quota = data.get("quota_remaining")

            if not items:
                consecutive_empty_pages += 1
                if consecutive_empty_pages >= 2:
                    break
            else:
                consecutive_empty_pages = 0

            kept_this_page = 0
            for q in items:
                qid = q["question_id"]
                if qid in collected:
                    continue
                if passes_quality_filter(q):
                    collected[qid] = q
                    kept_this_page += 1
                    if len(collected) >= target:
                        break

            log.info(
                "page_done",
                tag=tag, page=page, fetched=len(items),
                kept=kept_this_page, total=len(collected),
                quota_remaining=quota,
            )

            if not has_more:
                break
            page += 1

    log.info("questions_collected", count=len(collected))

    # Fetch the accepted answers in bulk
    answer_ids = [q["accepted_answer_id"] for q in collected.values() if q.get("accepted_answer_id")]
    log.info("fetching_answers", count=len(answer_ids))
    answers = client.fetch_answers_bulk(answer_ids)
    answers_by_id = {a["answer_id"]: a for a in answers}

    # Stitch question + accepted answer
    paired: List[Dict[str, Any]] = []
    for q in collected.values():
        aid = q.get("accepted_answer_id")
        a = answers_by_id.get(aid)
        if not a or not passes_answer_quality(a):
            continue

        record = {
            "source_platform": "stackoverflow",
            "source_id": str(q["question_id"]),
            "source_url": q.get("link") or f"https://stackoverflow.com/questions/{q['question_id']}",
            "title": q.get("title") or "",
            "tags": q.get("tags") or [],
            "module": classify_module(q.get("tags") or []),
            "question_body": q.get("body") or "",
            "answer_body": a.get("body") or "",
            "question_score": q.get("score", 0),
            "answer_score": a.get("score", 0),
            "view_count": q.get("view_count", 0),
            "answer_count": q.get("answer_count", 0),
            "is_accepted": True,
            "creation_date_unix": q.get("creation_date"),
            "fetched_at": datetime.now(timezone.utc).isoformat(),
        }
        paired.append(record)

    log.info("paired_records", count=len(paired))

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(paired, f, ensure_ascii=False, indent=2)

    log.info("scrape_complete", output_path=output_path, records_written=len(paired))


# ---------------------------------------------------------------------------
# Entry
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="Scrape SAP knowledge from Stack Overflow")
    parser.add_argument("--target", type=int, default=500, help="Number of records to collect")
    parser.add_argument("--output", default="scraped_so.json", help="Output JSON file path")
    args = parser.parse_args()
    scrape(target=args.target, output_path=args.output)


if __name__ == "__main__":
    main()
