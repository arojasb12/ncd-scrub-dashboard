-- NCD Scrub Dashboard — initial schema
-- Run once: psql $DATABASE_URL -f migrations/001_create_tables.sql

BEGIN;

CREATE TABLE IF NOT EXISTS scrub_entries (
    id              SERIAL PRIMARY KEY,
    section         VARCHAR(50)  NOT NULL,     -- billing | admin | integrity | eligibility
    category        VARCHAR(100) NOT NULL,     -- display name on dashboard
    scrub_date      DATE         NOT NULL,     -- date extracted from filename
    value           INTEGER,                   -- count metric
    amount          NUMERIC(12, 2),            -- dollar metric (nullable)
    source_file     TEXT,                      -- SharePoint URL for traceability
    created_at      TIMESTAMPTZ  DEFAULT NOW(),

    CONSTRAINT uq_scrub_entry UNIQUE (section, category, scrub_date)
);

CREATE INDEX IF NOT EXISTS idx_scrub_entries_section_date
    ON scrub_entries (section, scrub_date);

CREATE INDEX IF NOT EXISTS idx_scrub_entries_category
    ON scrub_entries (category);

COMMENT ON TABLE  scrub_entries IS 'Parsed scrub metrics — one row per section/category/date';
COMMENT ON COLUMN scrub_entries.section     IS 'Dashboard section: billing, admin, integrity, eligibility';
COMMENT ON COLUMN scrub_entries.category    IS 'Metric display name, e.g. Billing Alignment, FBD in the Past';
COMMENT ON COLUMN scrub_entries.scrub_date  IS 'Date the scrub was run (from filename MM.DD.YYYY)';
COMMENT ON COLUMN scrub_entries.value       IS 'Count metric (nullable — some entries are amount-only)';
COMMENT ON COLUMN scrub_entries.amount      IS 'Dollar amount metric (nullable — most entries are count-only)';
COMMENT ON COLUMN scrub_entries.source_file IS 'SharePoint web URL of the source .xlsx file';

COMMIT;
