"""
Scrub type auto-detection.

Given a filename and/or folder path from Power Automate, determine
which scrub type config to use for parsing.
"""

from __future__ import annotations

import logging
import re
from typing import Optional

from scrub_parser.config import SCRUB_CONFIGS, ScrubTypeConfig

logger = logging.getLogger(__name__)

# ── Filename keyword patterns per scrub type ───────────────────────────────
# Order matters — more specific patterns first to avoid false positives.

_FILENAME_RULES: list[tuple[str, list[str]]] = [
    ("ach_rebill",               ["ach rebill", "ach hold rebill"]),
    ("billing_date_alignment",   ["billing date alignment"]),
    ("fbd",                      ["first billing date", "fbd"]),
    ("nbd",                      ["next billing date", "nbd in past"]),
    ("mass_terms",               ["mass term"]),
    ("combined_eligibility",     ["combined eligibility"]),
    ("daily_dupes",              ["daily dupe"]),
    ("realm_health",             ["realm health"]),
    ("account_updater",          ["account updater"]),
    ("overaged_dependents",      ["overaged dependent", "age out"]),
    ("all_hold_reasons",         ["hold reason", "all hold"]),
    ("zero_product_fee",         ["$0 product fee", "0 product fee"]),
    ("srs_aged",                 ["sr", "service request", "aged 30"]),
    ("incomplete_accounts",      ["incomplete account"]),
    ("no_dependent_products",    ["no product", "dependents scrub"]),
    ("missing_mop",              ["missing mop"]),
    ("test_member_accounts",     ["test member"]),
]

# ── Folder path patterns ──────────────────────────────────────────────────

_FOLDER_RULES: list[tuple[str, list[str]]] = [
    ("ach_rebill",               ["ach rebill"]),
    ("billing_date_alignment",   ["billing date alignment"]),
    ("fbd",                      ["first billing date"]),
    ("nbd",                      ["next billing date", "nbd in past"]),
    ("mass_terms",               ["mass term"]),
    ("combined_eligibility",     ["eligibility error reports/2026",
                                  "eligibility error reports\\2026"]),
    ("daily_dupes",              ["daily dupe"]),
    ("realm_health",             ["realm health"]),
    ("account_updater",          ["account updater"]),
    ("overaged_dependents",      ["overaged dependent"]),
    ("all_hold_reasons",         ["all hold reasons"]),
    ("zero_product_fee",         ["$0 product fee", "0 product fee"]),
    ("srs_aged",                 ["srs aged", "service requests"]),
    ("incomplete_accounts",      ["incomplete account"]),
    ("no_dependent_products",    ["no products on dependents"]),
    ("missing_mop",              ["missing mop"]),
    ("test_member_accounts",     ["test member account"]),
]


def detect_scrub_type(filename: str,
                      folder_path: str = ""
                      ) -> Optional[str]:
    """
    Auto-detect scrub type key from a filename and/or folder path.

    Returns the scrub config key (e.g. 'billing_date_alignment') or None
    if no match is found.

    Checks filename first (more specific), then folder path as fallback.
    """
    fn_lower = filename.lower()
    fp_lower = folder_path.lower()

    # 1. try file_pattern regex from configs (most precise)
    for cfg in SCRUB_CONFIGS:
        if cfg.file_pattern:
            if re.search(cfg.file_pattern, filename, re.IGNORECASE):
                logger.info("Detected '%s' via file_pattern match on '%s'",
                            cfg.key, filename)
                return cfg.key

    # 2. try filename keyword rules
    for key, keywords in _FILENAME_RULES:
        if any(kw in fn_lower for kw in keywords):
            logger.info("Detected '%s' via filename keyword in '%s'", key, filename)
            return key

    # 3. try folder path rules
    if fp_lower:
        for key, keywords in _FOLDER_RULES:
            if any(kw in fp_lower for kw in keywords):
                logger.info("Detected '%s' via folder path '%s'", key, folder_path)
                return key

    logger.warning("Could not detect scrub type for file='%s', folder='%s'",
                   filename, folder_path)
    return None


def detect_or_raise(filename: str, folder_path: str = "") -> str:
    """Like detect_scrub_type but raises ValueError if no match."""
    key = detect_scrub_type(filename, folder_path)
    if key is None:
        raise ValueError(
            f"Could not determine scrub type for '{filename}' "
            f"(folder: '{folder_path}'). "
            "Check that the filename or folder matches a known scrub pattern."
        )
    return key
