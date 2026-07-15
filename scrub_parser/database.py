"""
Postgres storage layer for the NCD Scrub Dashboard.

Handles connection management, dedup checks, inserts, and queries.

Expects DATABASE_URL env var (postgres://user:pass@host:port/dbname).
"""

from __future__ import annotations

import datetime
import logging
import os
from contextlib import contextmanager
from decimal import Decimal
from typing import Optional

try:
    import psycopg2
    import psycopg2.extras
except ImportError:
    psycopg2 = None  # type: ignore[assignment]

from scrub_parser.models import ScrubResult

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS scrub_entries (
    id              SERIAL PRIMARY KEY,
    section         VARCHAR(50) NOT NULL,
    category        VARCHAR(100) NOT NULL,
    scrub_date      DATE NOT NULL,
    value           INTEGER,
    amount          NUMERIC(12, 2),
    source_file     TEXT,
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(section, category, scrub_date)
);

CREATE INDEX IF NOT EXISTS idx_scrub_entries_section_date
    ON scrub_entries(section, scrub_date);

CREATE INDEX IF NOT EXISTS idx_scrub_entries_category
    ON scrub_entries(category);
"""


class Database:
    """Postgres connection manager for scrub entries."""

    def __init__(self, database_url: Optional[str] = None):
        self.database_url = database_url or os.environ.get("DATABASE_URL", "")
        self._conn = None

    @contextmanager
    def connection(self):
        """Yield a psycopg2 connection, creating one if needed."""
        if psycopg2 is None:
            raise ImportError(
                "psycopg2 is required for database operations. "
                "Install it: pip install psycopg2-binary"
            )
        if self._conn is None or self._conn.closed:
            self._conn = psycopg2.connect(self.database_url)
        try:
            yield self._conn
        except Exception:
            self._conn.rollback()
            raise
        else:
            self._conn.commit()

    def close(self):
        if self._conn and not self._conn.closed:
            self._conn.close()

    # ── Schema ─────────────────────────────────────────────────────────────

    def create_tables(self):
        """Run the schema migration (idempotent)."""
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(SCHEMA_SQL)
        logger.info("Database schema ensured")

    # ── Dedup ──────────────────────────────────────────────────────────────

    def entry_exists(self, section: str, category: str,
                     scrub_date: datetime.date) -> bool:
        """Check if an entry already exists for this section + category + date."""
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT 1 FROM scrub_entries "
                    "WHERE section = %s AND category = %s AND scrub_date = %s",
                    (section, category, scrub_date)
                )
                return cur.fetchone() is not None

    # ── Insert ─────────────────────────────────────────────────────────────

    def insert_entry(self, result: ScrubResult) -> bool:
        """
        Insert a scrub entry. Returns True if inserted, False if duplicate.
        Uses ON CONFLICT to skip duplicates gracefully.
        """
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO scrub_entries (section, category, scrub_date,
                                              value, amount, source_file)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (section, category, scrub_date) DO NOTHING
                    RETURNING id
                    """,
                    (result.section.value, result.category, result.scrub_date,
                     result.value, result.amount, result.source_file)
                )
                row = cur.fetchone()
                if row:
                    logger.info("Inserted: %s / %s / %s → v=%s, a=%s",
                                result.section.value, result.category,
                                result.scrub_date, result.value, result.amount)
                    return True
                else:
                    logger.debug("Skipped (duplicate): %s / %s / %s",
                                 result.section.value, result.category,
                                 result.scrub_date)
                    return False

    def insert_batch(self, results: list[ScrubResult]) -> tuple[int, int]:
        """
        Insert multiple entries. Returns (inserted_count, skipped_count).
        """
        inserted = 0
        skipped = 0
        for r in results:
            if self.insert_entry(r):
                inserted += 1
            else:
                skipped += 1
        return inserted, skipped

    # ── Upsert ─────────────────────────────────────────────────────────────

    def upsert_entry(self, result: ScrubResult) -> bool:
        """Insert or update an existing entry. Returns True always."""
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO scrub_entries (section, category, scrub_date,
                                              value, amount, source_file)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (section, category, scrub_date)
                    DO UPDATE SET
                        value       = EXCLUDED.value,
                        amount      = EXCLUDED.amount,
                        source_file = EXCLUDED.source_file,
                        created_at  = NOW()
                    """,
                    (result.section.value, result.category, result.scrub_date,
                     result.value, result.amount, result.source_file)
                )
        return True

    # ── Query ──────────────────────────────────────────────────────────────

    def get_entries(self,
                    section: Optional[str] = None,
                    category: Optional[str] = None,
                    start_date: Optional[datetime.date] = None,
                    end_date: Optional[datetime.date] = None,
                    ) -> list[dict]:
        """
        Query scrub entries with optional filters.
        Returns list of dicts.
        """
        clauses = []
        params = []

        if section:
            clauses.append("section = %s")
            params.append(section)
        if category:
            clauses.append("category = %s")
            params.append(category)
        if start_date:
            clauses.append("scrub_date >= %s")
            params.append(start_date)
        if end_date:
            clauses.append("scrub_date <= %s")
            params.append(end_date)

        where = " AND ".join(clauses) if clauses else "1=1"
        sql = f"""
            SELECT id, section, category, scrub_date, value, amount,
                   source_file, created_at
            FROM scrub_entries
            WHERE {where}
            ORDER BY scrub_date DESC, section, category
        """

        with self.connection() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql, params)
                return [dict(row) for row in cur.fetchall()]

    def get_latest_per_category(self, section: Optional[str] = None) -> list[dict]:
        """Get the most recent entry for each category."""
        where = f"WHERE section = '{section}'" if section else ""
        sql = f"""
            SELECT DISTINCT ON (section, category)
                   id, section, category, scrub_date, value, amount,
                   source_file, created_at
            FROM scrub_entries
            {where}
            ORDER BY section, category, scrub_date DESC
        """
        with self.connection() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql)
                return [dict(row) for row in cur.fetchall()]
