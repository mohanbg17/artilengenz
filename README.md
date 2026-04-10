# NLP-SAP Query Engine

> Natural language interface to SAP S4/HANA and ECC for finance and logistics reports.
> Ask questions in plain English — get SAP data back.

## Quick Start

```bash
# 1. Install dependencies (Python 3.11+)
pip install -e ".[dev]"

# 2. Configure environment
cp .env.example .env
# Edit .env — set ANTHROPIC_API_KEY (keep MOCK_SAP=true for local dev)

# 3. Run the API server
uvicorn api.main:app --reload --host 0.0.0.0 --port 8080

# 4. Try a query
curl "http://localhost:8080/api/v1/query?q=Show+AR+open+items+for+customer+10001"
```

## Example Queries

```
Finance (FI/CO)
  "What is the GL balance for account 400000 in company code 1000 for 2024?"
  "Show all open AR items for customer C10001 overdue today"
  "List vendor invoices due this month for company code 2000"
  "Compare actual vs planned costs for cost center CC1001 in Q1 2024"
  "Show the trial balance for fiscal year 2024"

Logistics (MM/SD)
  "What is the status of purchase order 4500012345?"
  "List all open POs for vendor V20001 created this year"
  "What is the current stock of material M-001 in plant 1000?"
  "Show all goods receipts for PO 4500012345"
  "Which sales orders for customer C10001 are not yet delivered?"
  "Total revenue by month for sales org 1000 in 2024"
```

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/api/v1/query?q=...` | Run a query (browser/curl friendly) |
| `POST` | `/api/v1/query` | Run a query with full options |
| `POST` | `/api/v1/query/export` | Download result as CSV or Excel |
| `GET` | `/api/v1/intents` | List all supported intents |
| `GET` | `/api/v1/schema/{TABLE}` | Get SAP table field definitions |
| `GET` | `/api/v1/ping` | Check SAP connectivity |
| `GET` | `/health` | Health check |
| `GET` | `/docs` | Swagger UI |

## Response Format

```json
{
  "query": "Show AR open items for customer C10001",
  "intent": "ar_open_items",
  "module": "FI",
  "confidence": 0.88,
  "confidence_label": "High",
  "confidence_interval": {
    "lower": 0.79,
    "upper": 0.94
  },
  "data": [
    {"BUKRS": "1000", "KUNNR": "C10001", "BELNR": "1400000001",
     "DMBTR": "45230.50", "WAERS": "USD", "FAEDT": "2024-03-15"}
  ],
  "total_count": 12,
  "has_more": false,
  "execution_plan": {
    "primary_table": "BSID",
    "query_type": "bapi",
    "bapi": "BAPI_AR_ACC_GETOPENITEMS",
    "filters": {"BUKRS": "1000", "KUNNR": "C10001"},
    "elapsed_ms": 87.3
  },
  "warnings": [],
  "source": "mock"
}
```

## Supported SAP Modules & Tables

| Module | Key Tables / APIs |
|--------|-------------------|
| **FI — Financial Accounting** | ACDOCA, BKPF, BSEG, BSID, BSIK, BSAD, BSAK, SKA1, FAGLFLEXT |
| **CO — Controlling** | COSP, COSS, CSKS |
| **MM — Materials Management** | EKKO, EKPO, MARA, MARC, MSEG, MCHB |
| **SD — Sales & Distribution** | VBAK, VBAP, VBRK, LIPS |
| **WM — Warehouse Management** | LQUA, LGPLA |

## SAP Connectivity Options

### Option 1: S4/HANA OData (default)
```env
SAP_SYSTEM_TYPE=S4HANA
SAP_HOST=my-s4.example.com
SAP_AUTH_TYPE=basic
SAP_USERNAME=nlp_user
SAP_PASSWORD=secret
MOCK_SAP=false
```

### Option 2: ECC RFC
```env
SAP_SYSTEM_TYPE=ECC
SAP_RFC_ENABLED=true
SAP_RFC_HOST=10.0.0.1
SAP_RFC_SYSNR=00
MOCK_SAP=false
```

### Option 3: Mock (no SAP needed)
```env
MOCK_SAP=true
```

## Architecture

See [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for the full pipeline diagram.

## Advantages, Disadvantages & Challenges

See [`docs/ANALYSIS.md`](docs/ANALYSIS.md) for a detailed analysis including:
- Business and technical advantages
- Known limitations and workarounds  
- 11 technical challenges with mitigations
- Confidence interval interpretation guide

## Running Tests

```bash
# Unit tests (no SAP, no LLM needed)
pytest tests/unit/ -v

# Integration tests (mock SAP, LLM mocked)
pytest tests/integration/ -v

# All tests with coverage
pytest --cov=nlp_sap --cov-report=term-missing
```

## Docker

```bash
docker build -t nlp-sap-query .
docker run -p 8080:8080 --env-file .env nlp-sap-query
```

## Project Structure

```
artilengenz/
├── src/nlp_sap/
│   ├── config.py                # Pydantic settings
│   ├── orchestrator.py          # Main pipeline coordinator
│   ├── connectors/
│   │   ├── base.py              # Abstract connector interface
│   │   ├── odata.py             # SAP OData v2/v4 (S4/HANA)
│   │   ├── rfc.py               # SAP RFC / pyrfc (ECC)
│   │   ├── mock.py              # Mock connector (dev/test)
│   │   └── factory.py           # Connector selection
│   ├── nlp/
│   │   ├── models.py            # Pydantic models (intent, plan, response)
│   │   ├── intent_classifier.py # Claude LLM + keyword fallback
│   │   ├── query_builder.py     # Intent → SAP query plan
│   │   └── confidence.py        # Wilson score CI engine
│   └── schema/
│       └── registry.py          # YAML schema loader / lookup
├── api/
│   ├── main.py                  # FastAPI app + lifespan
│   └── routes.py                # API endpoints
├── config/
│   ├── sap_schemas.yaml         # SAP table/BAPI/OData definitions
│   └── intent_patterns.yaml     # Intent keywords + examples
├── tests/
│   ├── unit/                    # Fast, no-network tests
│   └── integration/             # Full pipeline with mocked SAP+LLM
├── docs/
│   ├── ARCHITECTURE.md          # System design
│   └── ANALYSIS.md              # Advantages, disadvantages, challenges, CI
├── .env.example                 # Environment variable template
└── pyproject.toml               # Dependencies
```
