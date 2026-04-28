-- Postgres DDL for SAP Error Intelligence Platform
-- Mirrors the Snowflake schema, with type / syntax adjustments.

CREATE SCHEMA IF NOT EXISTS raw;
CREATE SCHEMA IF NOT EXISTS intel;

-- ============================================================================
-- raw.raw_errors - landing table for live SAP errors
-- ============================================================================
CREATE TABLE IF NOT EXISTS raw.raw_errors (
  hash_key         VARCHAR(64) PRIMARY KEY,
  source           VARCHAR(20) NOT NULL,
  system_id        VARCHAR(200) NOT NULL,
  occurred_at      TIMESTAMP NOT NULL,
  extracted_at     TIMESTAMP NOT NULL,
  severity         VARCHAR(20),
  short_text       VARCHAR(2000),
  long_text        TEXT,
  error_id         VARCHAR(200),
  user_name        VARCHAR(64),
  transaction      VARCHAR(64),
  program          VARCHAR(128),
  abap_source      VARCHAR(256),
  line_no          INTEGER,
  exception_class  VARCHAR(128),
  message_class    VARCHAR(64),
  message_number   VARCHAR(10),
  message_v1       VARCHAR(256),
  message_v2       VARCHAR(256),
  message_v3       VARCHAR(256),
  message_v4       VARCHAR(256),
  instance         VARCHAR(64),
  work_process     VARCHAR(20),
  object           VARCHAR(64),
  sub_object       VARCHAR(64),
  job_name         VARCHAR(128),
  raw              JSONB
);
CREATE INDEX IF NOT EXISTS ix_raw_errors_system_source_extracted
  ON raw.raw_errors (system_id, source, extracted_at);
CREATE INDEX IF NOT EXISTS ix_raw_errors_occurred
  ON raw.raw_errors (occurred_at DESC);

-- ============================================================================
-- raw.watermarks - high-water mark per (source, system)
-- ============================================================================
CREATE TABLE IF NOT EXISTS raw.watermarks (
  source              VARCHAR(20),
  system_id           VARCHAR(200),
  last_occurred_at    TIMESTAMP,
  updated_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (source, system_id)
);

-- ============================================================================
-- intel.* - intelligence layer tables
-- ============================================================================
CREATE TABLE IF NOT EXISTS intel.corpus (
  corpus_id            VARCHAR(64) PRIMARY KEY,
  source_platform      VARCHAR(40),
  source_url           VARCHAR(2000),
  error_signature      VARCHAR(512),
  error_text           TEXT,
  proposed_solution    TEXT,
  system_module        VARCHAR(64),
  upvote_score         INTEGER,
  accepted_flag        BOOLEAN,
  scraped_at           TIMESTAMP,
  embedding_version    VARCHAR(40),
  pinecone_id          VARCHAR(128)
);
CREATE INDEX IF NOT EXISTS ix_corpus_unembedded
  ON intel.corpus (embedding_version) WHERE embedding_version IS NULL;

CREATE TABLE IF NOT EXISTS intel.classifications (
  classification_id    VARCHAR(64) PRIMARY KEY,
  error_hash_key       VARCHAR(64) NOT NULL,
  created_at           TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  top_proposal_title   VARCHAR(1024),
  composite_confidence NUMERIC(5, 4),
  badge                VARCHAR(20),
  retried              BOOLEAN,
  proposals            JSONB,
  citations            JSONB,
  critique_notes       TEXT,
  model_version        VARCHAR(40)
);
CREATE INDEX IF NOT EXISTS ix_classifications_error
  ON intel.classifications (error_hash_key);
CREATE INDEX IF NOT EXISTS ix_classifications_confidence
  ON intel.classifications (composite_confidence DESC);

CREATE TABLE IF NOT EXISTS intel.critiques (
  critique_id          VARCHAR(64) PRIMARY KEY,
  classification_id    VARCHAR(64) NOT NULL,
  created_at           TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  critique_text        TEXT,
  confidence_delta     NUMERIC(6, 4),
  added_proposals      JSONB,
  removed_proposals    JSONB
);

CREATE TABLE IF NOT EXISTS intel.human_feedback (
  feedback_id          VARCHAR(64) PRIMARY KEY,
  classification_id    VARCHAR(64) NOT NULL,
  reviewer             VARCHAR(128),
  reviewed_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  accepted             BOOLEAN,
  corrected_solution   TEXT,
  comments             VARCHAR(8000)
);
