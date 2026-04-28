-- ============================================================================
-- Migration: intel.raw_embeddings_log
-- Purpose : Track which raw_errors have been embedded into Pinecone, version
--           the embedding model, and record failure state for retry logic.
-- Run as  : artilegenz user against sap_errors database
-- ============================================================================

CREATE TABLE IF NOT EXISTS intel.raw_embeddings_log (
  hash_key            VARCHAR(64) PRIMARY KEY
                      REFERENCES raw.raw_errors(hash_key) ON DELETE CASCADE,
  embedding_version   VARCHAR(40) NOT NULL,
  pinecone_id         VARCHAR(128) NOT NULL,
  pinecone_namespace  VARCHAR(64) NOT NULL,
  embedded_at         TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  text_length         INTEGER,
  truncated           BOOLEAN DEFAULT FALSE,
  embed_error         TEXT,
  retry_count         INTEGER DEFAULT 0,
  last_attempt_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS ix_raw_embeddings_log_version
  ON intel.raw_embeddings_log (embedding_version);

CREATE INDEX IF NOT EXISTS ix_raw_embeddings_log_failed
  ON intel.raw_embeddings_log (last_attempt_at)
  WHERE embed_error IS NOT NULL;

-- ============================================================================
-- Helper view: raw_errors that still need embedding
--   - never embedded, OR
--   - embedded with an older model version (set via env), OR
--   - failed last attempt and retry_count < 3
-- ============================================================================
CREATE OR REPLACE VIEW intel.v_pending_embeddings AS
SELECT
  r.hash_key,
  r.source,
  r.system_id,
  r.occurred_at,
  r.severity,
  r.short_text,
  r.long_text,
  r.error_id,
  r.transaction,
  r.program,
  r.user_name,
  r.instance,
  COALESCE(l.retry_count, 0) AS retry_count,
  l.embedding_version       AS current_version,
  l.embed_error             AS last_error
FROM raw.raw_errors r
LEFT JOIN intel.raw_embeddings_log l ON l.hash_key = r.hash_key
WHERE
  -- Not embedded yet
  l.hash_key IS NULL
  -- OR embedded with older version (worker checks current target version)
  -- OR failed but eligible for retry (retry_count < 3)
  OR (l.embed_error IS NOT NULL AND COALESCE(l.retry_count, 0) < 3)
ORDER BY r.occurred_at ASC;

-- ============================================================================
-- Sanity reporting view
-- ============================================================================
CREATE OR REPLACE VIEW intel.v_embedding_status AS
SELECT
  COUNT(*) FILTER (WHERE l.hash_key IS NOT NULL AND l.embed_error IS NULL) AS embedded,
  COUNT(*) FILTER (WHERE l.hash_key IS NULL)                                AS pending,
  COUNT(*) FILTER (WHERE l.embed_error IS NOT NULL)                         AS failed,
  COUNT(*)                                                                   AS total
FROM raw.raw_errors r
LEFT JOIN intel.raw_embeddings_log l ON l.hash_key = r.hash_key;

-- Verify table created
\d intel.raw_embeddings_log;
SELECT 'Migration complete' AS status;
