-- ============================================================================
-- Migration 02: Extend intel.corpus for scraping pipeline
-- Purpose : Add columns needed by the SO scraper (raw_html, fetched_at,
--           dedup tracking, embed status). Backwards-compatible -- existing
--           columns untouched.
-- ============================================================================

-- Add scraper-tracking columns if they don't already exist
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_schema = 'intel' AND table_name = 'corpus' AND column_name = 'fetched_at'
  ) THEN
    ALTER TABLE intel.corpus ADD COLUMN fetched_at TIMESTAMP;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_schema = 'intel' AND table_name = 'corpus' AND column_name = 'source_id'
  ) THEN
    ALTER TABLE intel.corpus ADD COLUMN source_id VARCHAR(64);
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_schema = 'intel' AND table_name = 'corpus' AND column_name = 'tags'
  ) THEN
    ALTER TABLE intel.corpus ADD COLUMN tags TEXT[];
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_schema = 'intel' AND table_name = 'corpus' AND column_name = 'view_count'
  ) THEN
    ALTER TABLE intel.corpus ADD COLUMN view_count INTEGER;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_schema = 'intel' AND table_name = 'corpus' AND column_name = 'answer_count'
  ) THEN
    ALTER TABLE intel.corpus ADD COLUMN answer_count INTEGER;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_schema = 'intel' AND table_name = 'corpus' AND column_name = 'question_score'
  ) THEN
    ALTER TABLE intel.corpus ADD COLUMN question_score INTEGER;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_schema = 'intel' AND table_name = 'corpus' AND column_name = 'language'
  ) THEN
    ALTER TABLE intel.corpus ADD COLUMN language VARCHAR(10) DEFAULT 'en';
  END IF;
END $$;

-- Unique constraint on (source_platform, source_id) prevents duplicates across re-runs
CREATE UNIQUE INDEX IF NOT EXISTS ux_corpus_source
  ON intel.corpus (source_platform, source_id)
  WHERE source_id IS NOT NULL;

-- Helper view: corpus rows that need embedding (matches embedding-worker pattern)
CREATE OR REPLACE VIEW intel.v_corpus_pending_embeddings AS
SELECT
  corpus_id,
  source_platform,
  source_url,
  error_signature,
  error_text,
  proposed_solution,
  system_module,
  upvote_score,
  accepted_flag,
  tags
FROM intel.corpus
WHERE embedding_version IS NULL
   OR embedding_version != 'voyage-3-large';

CREATE OR REPLACE VIEW intel.v_corpus_status AS
SELECT
  COUNT(*) FILTER (WHERE embedding_version IS NOT NULL) AS embedded,
  COUNT(*) FILTER (WHERE embedding_version IS NULL)     AS pending,
  COUNT(*)                                              AS total,
  COUNT(*) FILTER (WHERE source_platform = 'stackoverflow') AS from_stackoverflow,
  COUNT(*) FILTER (WHERE source_platform = 'sap_community') AS from_sap_community
FROM intel.corpus;

SELECT 'Migration 02 complete' AS status;
