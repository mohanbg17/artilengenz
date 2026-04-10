"""Intent Classifier — uses Claude to classify NL queries into SAP intents.

Architecture:
  1. Build a system prompt injecting the full intent catalogue + schema hints.
  2. Send user query to Claude (claude-sonnet-4-6) with structured JSON output.
  3. Parse and validate the JSON response into a ParsedIntent.
  4. Fall back to keyword matching if LLM fails or returns low-confidence result.

Two-pass strategy for reliability:
  - Pass 1: LLM (high accuracy, slower)
  - Pass 2: keyword fallback (fast, lower accuracy)
  Final confidence = weighted blend of LLM confidence + keyword match score.
"""

from __future__ import annotations

import json
import logging
import re
from datetime import date

import anthropic

from nlp_sap.config import Settings, get_settings
from nlp_sap.nlp.models import ParsedIntent, SAPModule
from nlp_sap.schema.registry import SchemaRegistry, get_registry

logger = logging.getLogger(__name__)

_SYSTEM_PROMPT = """You are an expert SAP S4/HANA and ECC query analyst.
Your task is to analyze a natural language query and extract:
1. The SAP business intent (from the catalogue below)
2. All relevant filter parameters (company code, fiscal year, etc.)
3. A confidence score (0.0 to 1.0) for your classification

## Available Intents
{intent_catalogue}

## Current Date
{today}

## Output Format
Respond ONLY with a valid JSON object — no markdown, no prose, no code fences.
Use null for fields that are not mentioned or cannot be inferred.

```json
{{
  "intent_name": "<intent from catalogue>",
  "module": "<FI|CO|MM|SD|WM>",
  "confidence": <0.0-1.0>,
  "description": "<one-line description of what was asked>",
  "company_code": "<4-digit code or null>",
  "fiscal_year": "<4-digit year or null>",
  "fiscal_period": "<2-digit period 01-16 or null>",
  "fiscal_periods": ["<list of periods if a range or quarter was mentioned>"],
  "gl_account": "<10-digit account or null>",
  "cost_center": "<cost center or null>",
  "profit_center": "<profit center or null>",
  "customer": "<customer number or null>",
  "vendor": "<vendor number or null>",
  "material": "<material number or null>",
  "plant": "<4-char plant code or null>",
  "storage_location": "<4-char storage location or null>",
  "document_number": "<document number or null>",
  "document_type": "<2-char document type or null>",
  "sales_org": "<4-char sales org or null>",
  "po_number": "<10-digit PO number or null>",
  "date_from": "<YYYY-MM-DD or null>",
  "date_to": "<YYYY-MM-DD or null>",
  "amount_min": <number or null>,
  "amount_max": <number or null>,
  "overdue_only": <true|false>,
  "max_rows": <integer, default 500>
}}
```

Important rules:
- If the query is ambiguous between two intents, pick the most likely one and lower confidence.
- If no intent matches, use "gl_balance_inquiry" with confidence 0.2.
- For relative dates like "this month" or "Q3 2024", resolve to exact dates based on today's date.
- Extract ALL filter values mentioned — company codes, years, amounts, etc.
- confidence > 0.85 = high certainty; 0.6-0.85 = moderate; < 0.6 = low.
"""


class IntentClassifier:
    """Classifies user queries into SAP intents using Claude + keyword fallback."""

    def __init__(
        self,
        settings: Settings | None = None,
        registry: SchemaRegistry | None = None,
    ) -> None:
        self._settings = settings or get_settings()
        self._registry = registry or get_registry()
        self._client = anthropic.Anthropic(
            api_key=self._settings.anthropic_api_key.get_secret_value()
        )

    def _build_system_prompt(self) -> str:
        return _SYSTEM_PROMPT.format(
            intent_catalogue=self._registry.intent_summary_for_prompt(),
            today=date.today().isoformat(),
        )

    def _keyword_score(self, query: str, intent_name: str, intent_def: dict) -> float:
        """Simple keyword overlap score [0, 1]."""
        q_lower = query.lower()
        keywords: list[str] = intent_def.get("keywords", [])
        hits = sum(1 for kw in keywords if kw.lower() in q_lower)
        return min(hits / max(len(keywords), 1), 1.0)

    def _keyword_fallback(self, query: str) -> tuple[str, float]:
        """Return best-matching intent name and score using keyword matching."""
        best_intent = "gl_balance_inquiry"
        best_score = 0.0
        for name, defn in self._registry.all_intents().items():
            score = self._keyword_score(query, name, defn)
            if score > best_score:
                best_score = score
                best_intent = name
        return best_intent, min(best_score * 0.7, 0.7)  # cap fallback at 0.7

    def _parse_llm_json(self, content: str) -> dict:
        """Extract and parse JSON from LLM response, handling markdown fences."""
        # Strip code fences if present
        content = re.sub(r"```(?:json)?", "", content).strip()
        # Find first { … } block
        match = re.search(r"\{.*\}", content, re.DOTALL)
        if not match:
            raise ValueError("No JSON object found in LLM response")
        return json.loads(match.group())

    async def classify(self, query: str) -> ParsedIntent:
        """Classify a natural language query into a ParsedIntent.

        Uses Claude for primary classification with keyword-matching as fallback.
        """
        # ── LLM pass ────────────────────────────────────────────────────────
        llm_data: dict = {}
        llm_intent: str = ""
        llm_confidence: float = 0.0

        try:
            import asyncio

            loop = asyncio.get_running_loop()

            def _call_llm() -> str:
                response = self._client.messages.create(
                    model=self._settings.llm_model,
                    max_tokens=self._settings.llm_max_tokens,
                    temperature=self._settings.llm_temperature,
                    system=self._build_system_prompt(),
                    messages=[{"role": "user", "content": query}],
                )
                return response.content[0].text

            raw_text = await loop.run_in_executor(None, _call_llm)
            llm_data = self._parse_llm_json(raw_text)
            llm_intent = llm_data.get("intent_name", "")
            llm_confidence = float(llm_data.get("confidence", 0.5))
            logger.debug("LLM classified '%s' → %s (%.2f)", query, llm_intent, llm_confidence)

        except Exception as exc:
            logger.warning("LLM classification failed: %s — using keyword fallback", exc)

        # ── Keyword fallback pass ────────────────────────────────────────────
        kw_intent, kw_score = self._keyword_fallback(query)

        # ── Blend: LLM wins unless it failed / low confidence ────────────────
        if llm_intent and llm_confidence >= 0.4:
            final_intent = llm_intent
            # Keyword score as a secondary signal — boosts confidence if they agree
            agreement_bonus = 0.05 if llm_intent == kw_intent else 0.0
            final_confidence = min(llm_confidence + agreement_bonus, 1.0)
        else:
            final_intent = kw_intent
            final_confidence = kw_score
            llm_data = {}  # Clear since LLM failed

        # Validate intent exists in registry
        intent_def = self._registry.get_intent(final_intent)
        if not intent_def:
            final_intent = kw_intent
            final_confidence = kw_score * 0.8

        module_str = llm_data.get("module") or (
            self._registry.get_intent(final_intent) or {}
        ).get("module", "FI")
        try:
            module = SAPModule(module_str)
        except ValueError:
            module = SAPModule.UNKNOWN

        # Resolve relative dates from LLM (it may emit symbolic tokens)
        date_from = llm_data.get("date_from") or self._resolve_date_token("date_from", query)
        date_to = llm_data.get("date_to") or self._resolve_date_token("date_to", query)

        return ParsedIntent(
            intent_name=final_intent,
            module=module,
            confidence=final_confidence,
            description=llm_data.get("description", ""),
            company_code=llm_data.get("company_code"),
            fiscal_year=llm_data.get("fiscal_year") or self._extract_year(query),
            fiscal_period=llm_data.get("fiscal_period"),
            fiscal_periods=llm_data.get("fiscal_periods") or [],
            gl_account=llm_data.get("gl_account"),
            cost_center=llm_data.get("cost_center"),
            profit_center=llm_data.get("profit_center"),
            customer=llm_data.get("customer"),
            vendor=llm_data.get("vendor"),
            material=llm_data.get("material"),
            plant=llm_data.get("plant"),
            storage_location=llm_data.get("storage_location"),
            document_number=llm_data.get("document_number"),
            document_type=llm_data.get("document_type"),
            sales_org=llm_data.get("sales_org"),
            po_number=llm_data.get("po_number"),
            date_from=date_from,
            date_to=date_to,
            amount_min=llm_data.get("amount_min"),
            amount_max=llm_data.get("amount_max"),
            overdue_only=bool(llm_data.get("overdue_only", False)),
            max_rows=int(llm_data.get("max_rows", 500)),
            raw_llm_json=llm_data,
        )

    # ── Helpers ─────────────────────────────────────────────────────────────

    @staticmethod
    def _extract_year(query: str) -> str | None:
        """Regex-extract a 4-digit year from the query."""
        match = re.search(r"\b(20\d{2})\b", query)
        return match.group(1) if match else None

    @staticmethod
    def _resolve_date_token(field: str, query: str) -> str | None:
        """Resolve 'today', 'this month' style expressions to ISO dates."""
        today = date.today()
        q = query.lower()
        if "today" in q:
            return today.isoformat()
        if "this month" in q:
            return date(today.year, today.month, 1).isoformat() if field == "date_from" else today.isoformat()
        return None
