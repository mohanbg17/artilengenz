# NLP-SAP Query Engine — Architecture

## Overview

A natural-language interface that lets business users query SAP S4/HANA and ECC
for finance (FI/CO) and logistics (MM/SD/WM) data without writing ABAP, SQL,
or navigating SAP transaction codes.

```
┌────────────────────────────────────────────────────────────────────┐
│                         User / Caller                              │
│  "What are the open AR items for customer 10001 in company 1000?"  │
└──────────────────────────────┬─────────────────────────────────────┘
                               │ HTTP POST /api/v1/query
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                        FastAPI Application                          │
│  • Request validation (Pydantic)                                    │
│  • Auth middleware (future: JWT)                                     │
│  • Export endpoints (CSV / Excel)                                   │
└──────────────────────────────┬──────────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                       QueryOrchestrator                             │
│  Coordinates the full pipeline in sequence                          │
└──┬────────────────┬────────────────┬───────────────────────────────-┘
   │                │                │
   ▼                ▼                ▼
┌──────────┐  ┌──────────┐  ┌──────────────┐
│  Intent  │  │Confidence│  │Query Builder │
│Classifier│  │  Scorer  │  │              │
│          │  │          │  │ • Table plan │
│ Claude   │  │ Wilson   │  │ • BAPI plan  │
│ LLM +    │  │ Score CI │  │ • OData plan │
│ keyword  │  │          │  │              │
│ fallback │  │          │  │ S4 vs ECC    │
└────┬─────┘  └────┬─────┘  └──────┬───────┘
     │              │               │
     └──────────────┴───────────────┘
                    │
                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      SAP Connector Layer                            │
│                                                                     │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────────────┐  │
│  │ ODataConnec- │  │ RFCConnector │  │    MockSAPConnector      │  │
│  │ tor (httpx)  │  │ (pyrfc)      │  │ (fixtures, no SAP needed)│  │
│  │              │  │              │  │                          │  │
│  │ S4/HANA APIs │  │ BAPIs +      │  │  Development / Testing   │  │
│  │ OAuth2/Basic │  │ RFC_READ_    │  │                          │  │
│  │              │  │ TABLE (ECC)  │  │                          │  │
│  └──────┬───────┘  └──────┬───────┘  └──────────────────────────┘  │
└─────────┼─────────────────┼───────────────────────────────────────-─┘
          │                 │
          ▼                 ▼
┌─────────────────────────────────────────────────────────────────────┐
│            SAP S4/HANA (OData v2/v4)  /  SAP ECC (RFC)             │
│                                                                     │
│  FI: ACDOCA, BKPF, BSEG, BSID, BSIK, FAGLFLEXT, SKA1              │
│  CO: COSP, COSS, CSKS                                               │
│  MM: EKKO, EKPO, MARA, MARC, MSEG, MCHB                            │
│  SD: VBAK, VBAP, VBRK, LIPS                                         │
│  WM: LQUA, LGPLA                                                    │
└─────────────────────────────────────────────────────────────────────┘
```

## Component Details

### 1. Intent Classifier (`src/nlp_sap/nlp/intent_classifier.py`)

**Primary strategy — LLM (Claude claude-sonnet-4-6)**
- Constructs a system prompt injecting the full intent catalogue + SAP schema hints
- Sends user query; expects structured JSON response with intent + all entity fields
- Temperature = 0.0 (deterministic — critical for consistent query generation)

**Fallback strategy — Keyword matching**
- Each intent has a keyword list in `config/intent_patterns.yaml`
- TF-IDF-style overlap score over the raw query
- Caps at 0.7 confidence — always lower than a good LLM classification

**Blending rule**
- If LLM confidence ≥ 0.4 → use LLM intent; bonus +0.05 if keyword agrees
- Otherwise → keyword fallback

### 2. Confidence Scorer (`src/nlp_sap/nlp/confidence.py`)

Produces a calibrated confidence score and **Wilson score 95% CI** from 4 signals:

| Signal | Weight | Description |
|--------|--------|-------------|
| LLM self-reported confidence | 45% | Raw probability from the LLM |
| Keyword overlap | 20% | Fraction of intent keywords present in query |
| Entity completeness | 20% | Required filter fields were extracted |
| Schema consistency | 15% | Extracted values are SAP-valid (length, format) |

**Wilson CI formula** (n = evidence signals × 10):
```
lower = (p̂ + z²/2n − z√(p̂(1−p̂)/n + z²/4n²)) / (1 + z²/n)
upper = (p̂ + z²/2n + z√(p̂(1−p̂)/n + z²/4n²)) / (1 + z²/n)
```
Where z = 1.96 (95% CI). Wilson is preferred over Wald at confidence extremes.

**Confidence Labels**
- ≥ 0.85 → **High** — execute and return
- 0.65–0.84 → **Medium** — return with a review suggestion
- < 0.65 → **Low** — prompt user for clarification

### 3. Query Builder (`src/nlp_sap/nlp/query_builder.py`)

Maps each intent to the optimal SAP data access path:

| Intent | S4/HANA | ECC |
|--------|---------|-----|
| GL Balance | OData `API_GLACCOUNTLINEITEM_SRV` / ACDOCA | Table: FAGLFLEXT |
| AR Open Items | BAPI `BAPI_AR_ACC_GETOPENITEMS` | BAPI (same) |
| AP Open Items | BAPI `BAPI_AP_ACC_GETOPENITEMS` | BAPI (same) |
| Document Search | Table: BKPF + JOIN BSEG | Same |
| Cost Center | Table: COSP + COSS | Same |
| Purchase Orders | OData `API_PURCHASEORDER_PROCESS_SRV` | Table: EKKO + JOIN EKPO |
| Sales Orders | OData `API_SALES_ORDER_SRV` | Table: VBAK + JOIN VBAP |
| Inventory | Table: MCHB | Same |

### 4. Connector Layer

| Connector | When Used | Protocol |
|-----------|-----------|----------|
| `ODataConnector` | S4/HANA (default) | HTTPS + JSON, Basic/OAuth2 |
| `RFCConnector` | ECC / BAPI calls | SAP NW RFC SDK (`pyrfc`) |
| `MockSAPConnector` | Dev/Test (`MOCK_SAP=true`) | In-memory fixtures |

### 5. Schema Registry (`src/nlp_sap/schema/registry.py`)

YAML-driven catalogue of:
- 20+ SAP tables across FI, CO, MM, SD, WM with full field definitions
- BAPI signatures (import/export parameters, table parameters)
- OData service and entity set names
- Intent keyword lists and example queries
- Filter aliases (e.g. `"invoice"` → `["RE","KR","DR"]`)

## Supported SAP Modules

| Module | Transactions / Areas |
|--------|---------------------|
| **FI** (Financial Accounting) | G/L, AR, AP, Asset Accounting, Bank |
| **CO** (Controlling) | Cost Centers, Profit Centers, Internal Orders |
| **MM** (Materials Management) | Purchasing, Goods Movements, Inventory |
| **SD** (Sales & Distribution) | Sales Orders, Billing, Delivery |
| **WM** (Warehouse Management) | Stock in Bins, Quants |

## Data Flow for a Typical Query

```
Input: "What is the total revenue for customer C10001 in fiscal year 2024?"

1. IntentClassifier
   → intent_name = "revenue_report"
   → module = SD
   → customer = "C10001"
   → fiscal_year = "2024"
   → confidence = 0.93

2. ConfidenceScorer
   → score = 0.88 (weighted blend)
   → CI = [0.79, 0.94] (Wilson 95%)
   → label = "High"

3. QueryBuilder (S4/HANA mode)
   → primary_table = VBRK
   → filters = {KUNAG: "C10001", FKDAT: {gte: "2024-01-01", lte: "2024-12-31"}}
   → fields = [VBELN, FKART, KUNAG, FKDAT, NETWR, WAERK]
   → aggregations = [sum_revenue, monthly_breakdown]

4. ODataConnector / MockSAPConnector
   → GET /sap/opu/odata/sap/...?$filter=SoldToParty eq 'C10001'...
   → Returns 12 billing document rows

5. NLPQueryResponse
   → data: [{VBELN: "90000001", NETWR: "145230.50", ...}, ...]
   → confidence: 0.88
   → confidence_interval: {lower: 0.79, upper: 0.94}
   → confidence_label: "High"
```
