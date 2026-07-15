"""
Scrub type configuration registry.

Each ScrubTypeConfig declares:
  - Where to find files on SharePoint
  - Which parser function to call
  - How to map parser output keys → dashboard section + category
"""

from scrub_parser.models import OutputMapping, Section, ScrubTypeConfig

SCRUB_CONFIGS: list[ScrubTypeConfig] = [

    # ── 1. Billing Date Alignment ──────────────────────────────────────────
    ScrubTypeConfig(
        key="billing_date_alignment",
        display_name="Billing Date Alignment",
        folder_path="Scrubs/Billing Date Alignment Scrub/2026/",
        search_query="Billing Date Alignment Scrub",
        parser_func="parse_billing_date_alignment",
        outputs=[
            OutputMapping("count", Section.BILLING, "Billing Alignment"),
        ],
    ),

    # ── 2. FBD in the Past + Blank FBD ─────────────────────────────────────
    ScrubTypeConfig(
        key="fbd",
        display_name="FBD in the Past / Blank FBD",
        folder_path="Scrubs/First Billing Date in the Past Scrub/2026/",
        search_query="First Billing Date in the Past Scrub",
        parser_func="parse_fbd",
        outputs=[
            OutputMapping("fbd_past", Section.BILLING, "FBD in the Past"),
            OutputMapping("fbd_blank", Section.BILLING, "Blank FBD"),
        ],
    ),

    # ── 3. NBD in the Past + Blank NBD ─────────────────────────────────────
    ScrubTypeConfig(
        key="nbd",
        display_name="NBD in the Past / Blank NBD",
        folder_path="Scrubs/Next Billing Date in the Past Scrub/NBD in Past/2026/",
        search_query="Next Billing Date in the Past Scrub",
        parser_func="parse_nbd",
        outputs=[
            OutputMapping("nbd_past", Section.BILLING, "NBD in the Past"),
            OutputMapping("nbd_blank", Section.BILLING, "Blank NBD"),
        ],
    ),

    # ── 4. Mass Terms ──────────────────────────────────────────────────────
    ScrubTypeConfig(
        key="mass_terms",
        display_name="Mass Terms",
        folder_path="Mass Terms/2026/",
        search_query="Mass Terms",
        parser_func="parse_mass_terms",
        outputs=[
            OutputMapping("count", Section.ADMIN, "Mass Terms (Count)"),
            OutputMapping("amount", Section.ADMIN, "Mass Terms (Count)", is_amount=True),
        ],
        notes="Subfolder structure: {month}/MM.DD.YYYY/",
    ),

    # ── 5. ACH Rebill ─────────────────────────────────────────────────────
    ScrubTypeConfig(
        key="ach_rebill",
        display_name="ACH Rebill",
        folder_path="Accounting + Billing/ACH Rebill/2026/",
        search_query="ACH Rebill Cancel 2 Attempts",
        parser_func="parse_ach_rebill",
        file_pattern=r"ACH Rebill Cancel 2 Attempts.*\.xlsx",
        outputs=[
            OutputMapping("total_records", Section.BILLING, "FPN Update (ACH Rebill 2+ Attempts)"),
            OutputMapping("recapture_amount", Section.BILLING, "ACH Rebill (Recapture)", is_amount=True),
        ],
        notes="Subfolder: {month}/MM.DD.YYYY ACH Hold Rebill/",
    ),

    # ── 6. Incomplete Accounts ─────────────────────────────────────────────
    ScrubTypeConfig(
        key="incomplete_accounts",
        display_name="Incomplete Accounts",
        folder_path="Scrubs/Incomplete Account Scrub/2026/",
        search_query="Incomplete Account Scrub",
        parser_func="parse_incomplete_accounts",
        outputs=[
            OutputMapping("count", Section.INTEGRITY, "Incomplete Accounts"),
        ],
    ),

    # ── 7. No Products on Dependents ───────────────────────────────────────
    ScrubTypeConfig(
        key="no_dependent_products",
        display_name="No Products on Dependents",
        folder_path="Scrubs/No Products on Dependents Scrub/",
        search_query="No Products on Dependents Scrub",
        parser_func="parse_no_dependent_products",
        outputs=[
            OutputMapping("count", Section.INTEGRITY, "No Dependent Products"),
        ],
        notes="Subfolder: {month} 2026/",
    ),

    # ── 8. Missing MOP ────────────────────────────────────────────────────
    ScrubTypeConfig(
        key="missing_mop",
        display_name="Missing MOP",
        folder_path="Scrubs/Missing MOP/2026/",
        search_query="Missing MOP",
        parser_func="parse_missing_mop",
        outputs=[
            OutputMapping("count", Section.INTEGRITY, "Missing MOP"),
        ],
    ),

    # ── 9. SRs Aged 30+ Days ──────────────────────────────────────────────
    ScrubTypeConfig(
        key="srs_aged",
        display_name="SRs Aged 30+ Days",
        folder_path="Scrubs/Service Requests/SRs Aged 30+ Days Scrubs/2026/",
        search_query="SRs Aged 30+ Days",
        parser_func="parse_srs_aged",
        outputs=[
            OutputMapping("count", Section.ADMIN, "SRs Aged 30 Days"),
        ],
    ),

    # ── 10. Test Member Accounts ───────────────────────────────────────────
    ScrubTypeConfig(
        key="test_member_accounts",
        display_name="Test Member Accounts",
        folder_path="Scrubs/Test Member Account Scrub/2026/",
        search_query="Test Member Account Scrub",
        parser_func="parse_test_member_accounts",
        outputs=[
            OutputMapping("count", Section.INTEGRITY, "Test Member Accounts"),
        ],
    ),

    # ── 11. Overaged Dependents (Age Outs) ─────────────────────────────────
    ScrubTypeConfig(
        key="overaged_dependents",
        display_name="Overaged Dependents",
        folder_path="Scrubs/Overaged Dependent Scrub/2026/",
        search_query="Overaged Dependent Scrub",
        parser_func="parse_overaged_dependents",
        outputs=[
            OutputMapping("final_notice", Section.ADMIN, "Age Outs – Final Notice"),
            OutputMapping("thirty_day", Section.ADMIN, "Age Outs – 30-Day Notice"),
            OutputMapping("count", Section.ADMIN, "Age Outs – Final Notice"),
            # 'count' is fallback when file doesn't split into two sections
        ],
        notes="If file doesn't split Final/30-Day, 'count' maps to Final Notice",
    ),

    # ── 12. All Hold Reasons ───────────────────────────────────────────────
    ScrubTypeConfig(
        key="all_hold_reasons",
        display_name="All Hold Reasons",
        folder_path="Scrubs/All Hold Reasons Tracking/2026/",
        search_query="All Hold Reasons Tracking",
        parser_func="parse_all_hold_reasons",
        outputs=[
            OutputMapping("count", Section.ADMIN, "Policies on Hold 2+ Mo."),
        ],
        notes="Filtered count — requires date math (hold_date > 2 months before file date)",
    ),

    # ── 13. $0 Product Fee ─────────────────────────────────────────────────
    ScrubTypeConfig(
        key="zero_product_fee",
        display_name="$0 Product Fee",
        folder_path="Scrubs/$0 Product Fee Scrub/2026/",
        search_query="$0 Product Fee Scrub",
        parser_func="parse_zero_product_fee",
        outputs=[
            OutputMapping("count", Section.INTEGRITY, "$0 Product Fee"),
        ],
        notes="Files may be split into parts (- 1.xlsx, - 2.xlsx) — sum across same date",
    ),

    # ── 14. Daily Dupes ────────────────────────────────────────────────────
    ScrubTypeConfig(
        key="daily_dupes",
        display_name="Daily Dupes",
        folder_path="Scrubs/Eligibility Error Reports/Daily Dupe Reports/",
        search_query="Daily Dupe Reports",
        parser_func="parse_daily_dupes",
        outputs=[
            OutputMapping("count", Section.ELIGIBILITY, "Daily Dupes"),
        ],
        notes="Subfolder: {month} 2026/; files are daily",
    ),

    # ── 15. Combined Eligibility Errors ────────────────────────────────────
    ScrubTypeConfig(
        key="combined_eligibility",
        display_name="Combined Eligibility Errors",
        folder_path="Scrubs/Eligibility Error Reports/2026/",
        search_query="Combined Eligibility Errors",
        file_pattern=r".*Combined Eligibility Errors\.xlsx",
        parser_func="parse_combined_eligibility",
        outputs=[
            OutputMapping("metlife", Section.ELIGIBILITY, "MetLife (EDI)"),
            OutputMapping("vsp", Section.ELIGIBILITY, "VSP (EDI)"),
            OutputMapping("nwfa", Section.ELIGIBILITY, "NWFA (EDI)"),
            OutputMapping("oed", Section.ELIGIBILITY, "OED Errors (EDI)"),
        ],
        notes="Weekly files (usually Sunday dates)",
    ),

    # ── 16. Realm Health Errors ────────────────────────────────────────────
    ScrubTypeConfig(
        key="realm_health",
        display_name="Realm Health Errors",
        folder_path="Scrubs/Eligibility Error Reports/Realm Health/",
        search_query="Realm Health",
        parser_func="parse_realm_health",
        outputs=[
            OutputMapping("metlife_carrier", Section.ELIGIBILITY, "MetLife (Carrier)"),
            OutputMapping("vsp_carrier", Section.ELIGIBILITY, "VSP Carrier"),
        ],
        notes="Carrier determined by header row content",
    ),

    # ── 17. Account Updater ────────────────────────────────────────────────
    ScrubTypeConfig(
        key="account_updater",
        display_name="Account Updater",
        folder_path="Scrubs/Account Updater/2026/",
        search_query="Account Updater",
        parser_func="parse_account_updater",
        outputs=[
            OutputMapping("count", Section.BILLING, "Account Updater"),
            OutputMapping("amount", Section.BILLING, "Account Updater", is_amount=True),
        ],
    ),
]


# Indexed by key for O(1) lookup
SCRUB_CONFIG_MAP: dict[str, ScrubTypeConfig] = {c.key: c for c in SCRUB_CONFIGS}


def get_config(key: str) -> ScrubTypeConfig:
    """Retrieve a scrub config by its key."""
    cfg = SCRUB_CONFIG_MAP.get(key)
    if cfg is None:
        raise KeyError(f"No scrub config for key '{key}'. "
                       f"Available: {sorted(SCRUB_CONFIG_MAP.keys())}")
    return cfg
