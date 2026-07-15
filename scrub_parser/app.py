"""
NCD Scrub Dashboard — Web API

Endpoints:
  POST /api/ingest          Power Automate sends files here
  GET  /api/entries          Dashboard reads data (filterable)
  GET  /api/latest           Latest value per category
  POST /api/entries          Manual add for unresolved scrub types
  GET  /api/scrub-types      List configured scrub types
  GET  /api/health           Health check

Auth: X-API-Key header required on all /api/ routes.
      Set SCRUB_API_KEY env var.
"""

from __future__ import annotations

import datetime
import io
import logging
import os
from decimal import Decimal, InvalidOperation
from typing import Optional

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from scrub_parser.config import SCRUB_CONFIGS
from scrub_parser.database import Database
from scrub_parser.ingest import detect_or_raise, detect_scrub_type
from scrub_parser.models import Section, ScrubResult
from scrub_parser.runner import Runner

logger = logging.getLogger(__name__)

# ── App setup ──────────────────────────────────────────────────────────────

app = FastAPI(
    title="NCD Scrub Dashboard API",
    version="1.0.0",
    docs_url="/api/docs",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=os.environ.get("CORS_ORIGINS", "*").split(","),
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Shared dependencies ───────────────────────────────────────────────────

_db: Optional[Database] = None


def get_db() -> Database:
    global _db
    if _db is None:
        _db = Database()
        _db.create_tables()
    return _db


def get_runner() -> Runner:
    return Runner(graph=None, db=get_db())


API_KEY = os.environ.get("SCRUB_API_KEY", "")


def verify_api_key(x_api_key: str = Header(default="")):
    """Require a valid API key on all /api/ routes."""
    if not API_KEY:
        return  # no key configured → open (dev mode)
    if x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid or missing API key")


# ── Pydantic models ───────────────────────────────────────────────────────

class IngestResponse(BaseModel):
    status: str
    scrub_type: str
    filename: str
    scrub_date: Optional[str] = None
    entries_inserted: int = 0
    entries_skipped: int = 0
    errors: list[str] = []


class ManualEntry(BaseModel):
    section: str
    category: str
    scrub_date: datetime.date
    value: Optional[int] = None
    amount: Optional[float] = None
    source: str = "manual"


class EntryOut(BaseModel):
    id: int
    section: str
    category: str
    scrub_date: datetime.date
    value: Optional[int] = None
    amount: Optional[float] = None
    source_file: Optional[str] = None
    created_at: Optional[str] = None


# ── Ingest endpoint (Power Automate sends files here) ─────────────────────

@app.post("/api/ingest", response_model=IngestResponse,
          dependencies=[Depends(verify_api_key)])
async def ingest_file(
    file: UploadFile = File(...),
    folder_path: str = Form(default=""),
    source_url: str = Form(default=""),
    scrub_type_override: str = Form(default=""),
):
    """
    Receive a scrub .xlsx file, auto-detect the scrub type, parse it,
    and store the results.

    Power Automate sends:
      - file: the .xlsx file content
      - folder_path: (optional) SharePoint folder the file came from
      - source_url: (optional) SharePoint web URL for traceability
      - scrub_type_override: (optional) skip auto-detection, use this key
    """
    filename = file.filename or "unknown.xlsx"

    # detect scrub type
    try:
        if scrub_type_override:
            scrub_key = scrub_type_override
        else:
            scrub_key = detect_or_raise(filename, folder_path)
    except ValueError as e:
        return IngestResponse(
            status="error",
            scrub_type="unknown",
            filename=filename,
            errors=[str(e)],
        )

    # read file into memory
    content = await file.read()
    buf = io.BytesIO(content)

    # run the parser
    runner = get_runner()
    stats = runner.run_from_buffer(
        scrub_key=scrub_key,
        file_buf=buf,
        filename=filename,
        source_url=source_url or filename,
    )

    scrub_date_str = None
    from scrub_parser.utils import extract_date_from_filename
    d = extract_date_from_filename(filename)
    if d:
        scrub_date_str = d.isoformat()

    return IngestResponse(
        status="ok" if not stats.errors else "partial",
        scrub_type=scrub_key,
        filename=filename,
        scrub_date=scrub_date_str,
        entries_inserted=stats.entries_inserted,
        entries_skipped=stats.entries_skipped,
        errors=stats.errors,
    )


# ── Query endpoints (dashboard reads from here) ───────────────────────────

@app.get("/api/entries", dependencies=[Depends(verify_api_key)])
def get_entries(
    section: Optional[str] = Query(default=None),
    category: Optional[str] = Query(default=None),
    start_date: Optional[datetime.date] = Query(default=None),
    end_date: Optional[datetime.date] = Query(default=None),
):
    """
    Get scrub entries with optional filters.
    Used by the dashboard to populate charts and tables.
    """
    db = get_db()
    rows = db.get_entries(
        section=section,
        category=category,
        start_date=start_date,
        end_date=end_date,
    )
    # serialize dates and decimals
    for row in rows:
        for k, v in row.items():
            if isinstance(v, datetime.date):
                row[k] = v.isoformat()
            elif isinstance(v, datetime.datetime):
                row[k] = v.isoformat()
            elif isinstance(v, Decimal):
                row[k] = float(v)
    return rows


@app.get("/api/latest", dependencies=[Depends(verify_api_key)])
def get_latest(section: Optional[str] = Query(default=None)):
    """Get the most recent entry for each category."""
    db = get_db()
    rows = db.get_latest_per_category(section=section)
    for row in rows:
        for k, v in row.items():
            if isinstance(v, datetime.date):
                row[k] = v.isoformat()
            elif isinstance(v, datetime.datetime):
                row[k] = v.isoformat()
            elif isinstance(v, Decimal):
                row[k] = float(v)
    return rows


# ── Manual entry (for unresolved scrub types) ─────────────────────────────

@app.post("/api/entries", dependencies=[Depends(verify_api_key)])
def add_manual_entry(entry: ManualEntry):
    """
    Manually add a scrub entry (for types without automated file sources,
    like PTD vs Inactive Date).
    """
    try:
        section = Section(entry.section)
    except ValueError:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid section '{entry.section}'. "
                   f"Must be one of: {[s.value for s in Section]}"
        )

    result = ScrubResult(
        section=section,
        category=entry.category,
        scrub_date=entry.scrub_date,
        value=entry.value,
        amount=Decimal(str(entry.amount)) if entry.amount is not None else None,
        source_file=entry.source,
    )

    db = get_db()
    inserted = db.upsert_entry(result)
    return {"status": "ok", "upserted": True}


# ── Scrub types reference ─────────────────────────────────────────────────

@app.get("/api/scrub-types", dependencies=[Depends(verify_api_key)])
def list_scrub_types():
    """List all configured scrub types and their output categories."""
    return [
        {
            "key": cfg.key,
            "display_name": cfg.display_name,
            "outputs": [
                {
                    "section": out.section.value,
                    "category": out.category,
                    "is_amount": out.is_amount,
                }
                for out in cfg.outputs
            ],
        }
        for cfg in SCRUB_CONFIGS
    ]


# ── Health check ───────────────────────────────────────────────────────────

@app.get("/api/health")
def health():
    """Health check — no auth required."""
    return {"status": "ok", "version": "1.0.0"}


# ── Startup ────────────────────────────────────────────────────────────────

@app.on_event("startup")
async def startup():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
    )
    logger.info("NCD Scrub Dashboard API starting up")
    get_db()  # ensure schema exists on boot
