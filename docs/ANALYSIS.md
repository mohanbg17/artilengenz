# NLP-SAP Query Engine — Analysis

## Advantages

### Business Value
| Advantage | Detail |
|-----------|--------|
| **Zero SAP training required** | Business users ask in plain English — no T-codes, no ABAP, no report configuration |
| **Democratizes data access** | Controllers, buyers, warehouse managers get self-service analytics without IT tickets |
| **Reduces time-to-insight** | Query → result in seconds vs. hours for a custom ABAP report |
| **Cross-module queries** | One interface spans FI, CO, MM, SD — unlike monolithic SAP fiori tiles |
| **Audit trail** | Every query is logged with intent, filters, confidence score, and execution plan |
| **Export-ready** | CSV / Excel download built in — no SAP GUI dependency |

### Technical Value
| Advantage | Detail |
|-----------|--------|
| **LLM-powered** | Claude extracts entities and resolves ambiguity that regex/rule-based NLU cannot handle |
| **Schema-grounded** | The LLM is constrained to known SAP tables/BAPIs — it cannot hallucinate arbitrary queries |
| **Dual-mode connectivity** | Single codebase supports both S4/HANA (OData) and ECC (RFC) — no fork required |
| **Confidence-aware** | Wilson CI tells users how certain the system is — prevents silent wrong results |
| **Mock mode** | Runs fully offline during development — no expensive SAP sandbox needed |
| **BAPI-first for writes** | Write operations route through BAPIs (not raw table updates) — respects SAP data integrity |
| **Async throughput** | FastAPI + async connectors handle concurrent queries without blocking |

---

## Disadvantages

### Technical Limitations
| Disadvantage | Detail | Mitigation |
|-------------|--------|------------|
| **LLM latency** | Claude API adds 1–3 s per query on top of SAP round-trip | Cache common intents; use fast models for classification |
| **LLM cost** | ~$0.003–$0.015 per query at current Claude pricing | Batch, cache, or use smaller model for intent routing |
| **RFC_READ_TABLE risks** | Generic table reads bypass authorization objects on some systems | Prefer BAPIs/OData; make RFC_READ_TABLE opt-in per table |
| **OData coverage gaps** | Not all SAP tables/fields are exposed via standard OData APIs | Fall back to RFC for uncovered data |
| **No write operations** | Current scope is read-only reporting | Extend with BAPI calls for creation/update flows |
| **Large result sets** | max_rows capped at 5000 — unsuitable for bulk exports | Use background jobs + streaming for large extracts |
| **SAP authorization** | The RFC/OData user must have adequate authorizations — misconfigured auth = no data or excess access | Principle of least privilege; separate read-only RFC user |

### Business / Organisational Limitations
| Disadvantage | Detail |
|-------------|--------|
| **Schema maintenance** | New SAP fields/tables require YAML updates and re-testing |
| **Multilingual queries** | Currently optimised for English — other languages require prompt engineering |
| **Ambiguous queries** | "Show me costs" has 5+ valid intents — the system picks one and signals low confidence |
| **Training data dependency** | Intent keyword lists need domain expert review per company's SAP config |
| **SAP customisation** | Customer-specific Z-tables / custom fields are not in the default schema registry |
| **Compliance / audit** | In regulated industries, AI-generated queries may require human approval before execution |
| **SSO complexity** | Integrating with SAP SSO (SAML, Kerberos) requires additional middleware |

---

## Technical Challenges

### 1. SAP Semantic Ambiguity
SAP field names (`DMBTR`, `WRBTR`, `HSL`) are cryptic.  
The NLP engine must map "invoice amount" to the correct currency field
(local currency vs. document currency vs. group currency) based on context.
**Mitigation**: Schema registry field descriptions + few-shot LLM examples.

### 2. Fiscal Period vs Calendar Period
SAP fiscal years can start in any month (e.g., April–March in UK).  
"Q1 2024" means January–March in calendar terms but could be periods 10–12
in SAP if the fiscal year starts in April.  
**Mitigation**: Config-driven fiscal year variant mapping per company code.

### 3. Client / Mandant Isolation
SAP is multi-client. The NLP system must ensure queries are scoped to the
correct client (MANDT) to avoid cross-client data leakage.  
**Mitigation**: MANDT always injected server-side; never derived from user input.

### 4. ABAP Table Buffering
Some SAP tables (e.g., configuration tables like T001) are client-buffered.
RFC_READ_TABLE reads the database directly, bypassing the buffer, which can
return stale data.  
**Mitigation**: Use BAPIs for configuration data; OData services handle buffering correctly.

### 5. Pagination and Large Results
SAP OData v2 does not support server-side cursors.
`$skip + $top` pagination causes re-execution of the full query each page.  
**Mitigation**: Default max 500 rows; large exports via async background task.

### 6. Authorization Object Complexity
Each SAP RFC/OData call requires specific authorization objects
(e.g., `F_BKPF_BUK` for company code access, `M_BEST_BSA` for purchasing).
A misconfigured service account can silently return empty results.  
**Mitigation**: Startup connectivity test; explicit BAPI RETURN message parsing.

### 7. ECC vs S4/HANA Table Differences
`BSEG` (ECC) does not exist as the primary line item store in S4/HANA —
`ACDOCA` (Universal Journal) replaces it. Some tools still write to BSEG
compatibility views.  
**Mitigation**: System-type flag in config drives routing; dual table mappings
in `QueryBuilder`.

### 8. Data Volume and Performance
Finance tables like `BKPF`/`BSEG` can have hundreds of millions of rows.
An unfiltered query would time out or cause system instability.  
**Mitigation**: Required filter validation (company code + year minimum);
max_rows hard cap; BAPI calls that handle server-side aggregation.

### 9. LLM Hallucination of SAP Entities
Without grounding, an LLM might generate a non-existent company code,
G/L account, or table name.  
**Mitigation**:
- Schema registry acts as a validation whitelist
- Field value validation in `ConfidenceScorer._schema_consistency()`
- Low confidence triggers clarification prompt

### 10. RFC Connection Stability
Long-running RFC connections (pyrfc) can drop under network interruptions.
The RFC protocol has no built-in reconnection.  
**Mitigation**: `pyrfc.Connection.alive` check before each call + auto-reconnect.

### 11. SAP Unicode and Special Characters
SAP systems can be Unicode or non-Unicode. Material numbers, text fields, and
descriptions may contain umlauts, kanji, etc.  
**Mitigation**: RFC_READ_TABLE DELIMITER avoids encoding issues; OData returns UTF-8.

---

## Confidence Interval — Interpretation Guide

### What the CI Means
The 95% Wilson CI `[lower, upper]` answers:
> "If we ran 100 similar queries of this type and complexity,
>  in 95 of them the model would classify correctly with a probability
>  between `lower` and `upper`."

This is **not** the probability the current classification is correct —
it is a frequentist interval on the classification success rate.

### Example Interpretations

| Query | Score | 95% CI | Label | Action |
|-------|-------|--------|-------|--------|
| "Show AR open items for customer 10001 in company code 1000" | 0.93 | [0.84, 0.97] | High | Execute immediately |
| "Show me costs" | 0.58 | [0.42, 0.73] | Low | Ask: "Which cost module? Cost centers, profit centers, or purchase costs?" |
| "List vendor invoices overdue by 90 days" | 0.79 | [0.67, 0.88] | Medium | Execute; display "Interpreted as: AP Open Items — is that right?" |

### Why Wilson Score (not Wald)?
The Wald interval `p̂ ± z√(p̂(1-p̂)/n)` produces invalid intervals when
confidence is near 0 or 1 (the most common case for well-tuned NLP).  
Wilson is coverage-accurate across the full [0,1] range.

### Effective Sample Size (n)
We set `n = evidence_count × 10` where `evidence_count` is the number of
distinct signals observed (LLM vote, keyword hit, entity completeness, schema
consistency). More signals → narrower CI → higher certainty in the estimate.

### Recommended UI Thresholds

| Label | Suggested UX |
|-------|-------------|
| **High** (≥0.85) | Show results; no warning |
| **Medium** (0.65–0.84) | Show results with "Did you mean: `{description}`?" banner |
| **Low** (<0.65) | Show disambiguation dialog; don't auto-execute |

---

## Clarifying Questions — Decisions Made

These are questions that could change the design significantly.
Current defaults are listed; adjust via `.env`.

| Question | Current Default | Impact if Changed |
|----------|----------------|-----------------|
| SAP connectivity | OData (S4/HANA) | RFC mode: needs SAP NW RFC SDK |
| Auth type | Basic Auth | OAuth2: needs token endpoint config |
| LLM provider | Anthropic Claude | Others: change `IntentClassifier` |
| Fiscal year variant | Calendar year (Jan–Dec) | Non-standard: add FY mapping |
| Multi-language support | English only | Others: multilingual prompt engineering |
| Write operations | Read-only | Enable: add BAPI POST flows |
| Custom Z-tables | Not included | Add: extend `sap_schemas.yaml` |
| Row security | Service account level | Row-level: integrate SAP auth objects |
| Max result rows | 500 (API), 5000 (export) | Tune per hardware/network |
| Deployment | Docker single-container | K8s: add horizontal scaling |
