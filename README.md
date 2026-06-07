# IDoc Error-Healing Platform

> AI-assisted detection, classification, and resolution of failed SAP IDocs — built on an all open-source stack.

![License](https://img.shields.io/badge/license-MIT-green)
![Python](https://img.shields.io/badge/python-3.11+-blue)
![FastAPI](https://img.shields.io/badge/api-FastAPI-009688)
![PostgreSQL](https://img.shields.io/badge/store-PostgreSQL%20%2B%20pgvector-336791)
![Docker](https://img.shields.io/badge/deploy-Docker%20Compose-2496ED)

Failed SAP IDocs (status 51 / 56) are a constant operational drain — each one is triaged and fixed by hand. This platform automates that loop: it ingests failed IDocs over a direct RFC/IDoc interface, grounds a two-pass Claude classifier in a vector store of previously resolved cases (RAG), proposes a concrete fix, routes it to a human operator for approval, and reposts the corrected IDoc to SAP. Every outcome is fed back into the store so the system improves over time.

Validated at **99.3% classification accuracy across 5,000 synthetic IDocs.**

---

## Architecture & Data Flow

![Architecture and data flow of the IDoc Error-Healing Platform](docs/architecture.png)

The flow, end to end:

1. **Capture** — A failed IDoc (status 51 / 56) is detected in SAP S/4HANA.
2. **Inbound** — It is pulled into the platform over a direct **SAP RFC / IDoc interface (PyRFC)** — no middleware.
3. **Persist + embed** — A **FastAPI** service ingests the IDoc, persists it, and generates an embedding (**Voyage AI**).
4. **Retrieve** — The embedding queries **PostgreSQL + pgvector** to retrieve similar, previously resolved cases (RAG).
5. **Classify & resolve** — A **two-pass Claude** pipeline runs: *Pass 1 (Sonnet)* triages the error and assigns a category; *Pass 2 (Opus)* determines the root cause and proposes a concrete fix, grounded in the retrieved context.
6. **Review** — The proposed resolution surfaces in a **React operator console** for human-in-the-loop review and approval.
7. **Heal** — On approval, the corrected IDoc is reposted to SAP via the RFC/IDoc interface; the outcome is written back to the store for continuous learning.

---

## Why it is built this way

- **Two-pass classification** separates *cheap, fast triage* (Sonnet) from *expensive, careful reasoning* (Opus), so most volume is handled at low cost and only ambiguous cases pay for the deeper pass.
- **RAG grounding** keeps proposed fixes anchored in what has actually worked before, rather than free-form model output.
- **Human-in-the-loop** keeps an operator in control of every change applied to SAP — the system recommends, the human approves.
- **One open-source datastore** — PostgreSQL with the `pgvector` extension serves as both the vector store and the system-of-record (state, resolutions, audit), removing a managed vector-DB dependency.

---

## Tech stack

| Layer | Technology |
|-------|------------|
| API & orchestration | Python · FastAPI (async, REST + event-driven) |
| LLM classification | Anthropic Claude — Sonnet (triage) + Opus (resolution) |
| Embeddings | Voyage AI |
| Vector store + state | PostgreSQL + pgvector |
| Operator UI | React |
| SAP connectivity | PyRFC (RFC / IDoc port) |
| Packaging | Docker Compose |

> Voyage AI and Claude are external API dependencies; the rest of the stack is open source and self-hosted.

---

## Quickstart

```bash
# 1. Clone (IDoc MVP lives on the claude/idoc-mvp branch)
git clone -b claude/idoc-mvp https://github.com/mohanbg17/artilegenz.git
cd artilegenz

# 2. Configure environment (see below)
cp .env.example .env
#   then edit .env with your keys / connection details

# 3. Launch the full stack (API, Postgres+pgvector, operator console)
docker compose up --build

# API:              http://localhost:8000
# Operator console: http://localhost:3000
# API docs (Swagger): http://localhost:8000/docs
```

### Configuration

Set these in `.env` (never commit real secrets):

```ini
ANTHROPIC_API_KEY=your_anthropic_key
VOYAGE_API_KEY=your_voyage_key
DATABASE_URL=postgresql://idoc:idoc@db:5432/idoc

# SAP RFC / IDoc connection (optional — synthetic mode runs without SAP)
SAP_ASHOST=your_sap_host
SAP_SYSNR=00
SAP_CLIENT=100
SAP_USER=your_sap_user
SAP_PASSWD=your_sap_password
```

Running without a live SAP system? Start in **synthetic mode** to replay the bundled synthetic IDoc dataset:

```bash
docker compose run api python -m idoc_healing.cli --synthetic
```

---

## Results

| Metric | Value |
|--------|-------|
| Classification accuracy | **99.3%** |
| Evaluation set | 5,000 synthetic IDocs |
| Human-in-the-loop | Required before any repost to SAP |

---

## Project structure

```
idoc-error-healing/
├── docs/
│   └── architecture.png         # architecture & data-flow diagram
├── api/                         # FastAPI service
│   ├── ingestion.py             # RFC/IDoc inbound
│   ├── embeddings.py            # Voyage AI embedding
│   ├── retrieval.py             # pgvector RAG
│   ├── classifier.py            # two-pass Claude (Sonnet → Opus)
│   └── resolution.py            # auto-heal + repost
├── console/                     # React operator console
├── db/                          # PostgreSQL + pgvector schema/migrations
├── data/                        # synthetic IDoc fixtures
├── docker-compose.yml
├── .env.example
└── README.md
```

---

## Roadmap

- [ ] Pluggable embedding backends (open-source embeddings option)
- [ ] Confidence-based auto-approval for high-certainty, low-risk fixes
- [ ] Expanded message-type coverage (ORDERS, INVOIC, DESADV, …)
- [ ] Evaluation harness with category-level precision/recall reporting

---

## Disclaimer

This project was built and validated entirely on **synthetic IDoc data**. It contains no proprietary, customer, or employer data, configuration, or code. It is an independent, open-source demonstration of an AI-assisted SAP error-resolution pattern.

## License

Released under the MIT License. See [`LICENSE`](LICENSE).
