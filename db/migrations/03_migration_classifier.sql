-- ============================================================================
-- Migration 03: Extend intel.classifications for Sonnet+Opus pipeline
-- ============================================================================
-- Adds columns the classifier writes:
--   * sonnet_response (JSONB): raw first-pass output
--   * opus_critique  (JSONB): critique-pass output
--   * adaptive_split (JSONB): how many records pulled from each namespace
--   * sonnet_tokens / opus_tokens (JSONB): cost tracking
--   * status (VARCHAR): 'pending', 'classified', 'failed'
--   * error_message (TEXT): failure detail
-- ============================================================================

DO $$
BEGIN
  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='sonnet_response') THEN
    ALTER TABLE intel.classifications ADD COLUMN sonnet_response JSONB;
  END IF;

  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='opus_critique') THEN
    ALTER TABLE intel.classifications ADD COLUMN opus_critique JSONB;
  END IF;

  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='adaptive_split') THEN
    ALTER TABLE intel.classifications ADD COLUMN adaptive_split JSONB;
  END IF;

  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='sonnet_tokens') THEN
    ALTER TABLE intel.classifications ADD COLUMN sonnet_tokens JSONB;
  END IF;

  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='opus_tokens') THEN
    ALTER TABLE intel.classifications ADD COLUMN opus_tokens JSONB;
  END IF;

  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='status') THEN
    ALTER TABLE intel.classifications ADD COLUMN status VARCHAR(20) DEFAULT 'classified';
  END IF;

  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='error_message') THEN
    ALTER TABLE intel.classifications ADD COLUMN error_message TEXT;
  END IF;

  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='retry_count') THEN
    ALTER TABLE intel.classifications ADD COLUMN retry_count INTEGER DEFAULT 0;
  END IF;

  IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                 WHERE table_schema='intel' AND table_name='classifications'
                   AND column_name='summary_md') THEN
    ALTER TABLE intel.classifications ADD COLUMN summary_md TEXT;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS ix_classifications_status
  ON intel.classifications (status);
CREATE INDEX IF NOT EXISTS ix_classifications_created
  ON intel.classifications (created_at DESC);

-- Latest classification per error (for the "current diagnosis" view)
CREATE OR REPLACE VIEW intel.v_latest_classifications AS
SELECT DISTINCT ON (error_hash_key)
  classification_id, error_hash_key, created_at,
  top_proposal_title, composite_confidence, badge,
  status, summary_md, model_version
FROM intel.classifications
WHERE status = 'classified'
ORDER BY error_hash_key, created_at DESC;

-- "Pending classification" view: raw_errors that don't have a successful classification
CREATE OR REPLACE VIEW intel.v_pending_classifications AS
SELECT
  r.hash_key,
  r.source,
  r.system_id,
  r.occurred_at,
  r.severity,
  r.short_text,
  r.error_id,
  COALESCE(c2.retry_count, 0) AS retry_count
FROM raw.raw_errors r
LEFT JOIN intel.v_latest_classifications c ON c.error_hash_key = r.hash_key
LEFT JOIN (
  SELECT error_hash_key, MAX(retry_count) AS retry_count
  FROM intel.classifications
  WHERE status = 'failed'
  GROUP BY error_hash_key
) c2 ON c2.error_hash_key = r.hash_key
WHERE
  c.classification_id IS NULL
  AND COALESCE(c2.retry_count, 0) < 3;

CREATE OR REPLACE VIEW intel.v_classification_status AS
SELECT
  COUNT(*) FILTER (WHERE c.classification_id IS NOT NULL AND c.status = 'classified') AS classified,
  COUNT(*) FILTER (WHERE c.classification_id IS NULL)                                  AS pending,
  COUNT(*) FILTER (WHERE c.status = 'failed')                                          AS failed,
  COUNT(*)                                                                              AS total
FROM raw.raw_errors r
LEFT JOIN intel.v_latest_classifications c ON c.error_hash_key = r.hash_key;

SELECT 'Migration 03 complete' AS status;
