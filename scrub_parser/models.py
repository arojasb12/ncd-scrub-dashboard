"""
Data models for the NCD Scrub Dashboard parser.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import Callable, Optional


class Section(str, Enum):
    BILLING = "billing"
    ADMIN = "admin"
    INTEGRITY = "integrity"
    ELIGIBILITY = "eligibility"


@dataclass
class ScrubResult:
    """One parsed metric from a scrub file."""
    section: Section
    category: str
    scrub_date: datetime.date
    value: Optional[int] = None       # count
    amount: Optional[Decimal] = None  # dollar amount
    source_file: str = ""


@dataclass
class ScrubTypeConfig:
    """Declarative config for one scrub type."""
    key: str                          # unique identifier, e.g. "billing_date_alignment"
    display_name: str                 # human-readable, e.g. "Billing Date Alignment"
    folder_path: str                  # SharePoint relative folder
    search_query: str                 # Graph API search text
    parser_func: str                  # name of the parser function in parsers.py
    outputs: list[OutputMapping]      # what dashboard rows this scrub produces
    file_pattern: str = ""            # regex for matching filenames (optional override)
    notes: str = ""


@dataclass
class OutputMapping:
    """Maps a parser output key to a dashboard section + category."""
    output_key: str       # key returned by the parser, e.g. "fbd_past" or "count"
    section: Section
    category: str         # display name on dashboard, e.g. "FBD in the Past"
    is_amount: bool = False  # True if this output is a dollar value, not a count


@dataclass
class ParsedFile:
    """Intermediate record: one file that was found and parsed."""
    scrub_key: str
    filename: str
    scrub_date: datetime.date
    source_url: str
    results: dict[str, int | Decimal]  # output_key → value
    errors: list[str] = field(default_factory=list)
