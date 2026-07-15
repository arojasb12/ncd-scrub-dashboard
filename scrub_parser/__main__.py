"""
CLI entry point for the NCD Scrub Dashboard parser.

Usage:
    python -m scrub_parser run-all [--start-date 2026-01-01] [--end-date 2026-07-14] [--dry-run]
    python -m scrub_parser run-single billing_date_alignment [--dry-run]
    python -m scrub_parser parse-local billing_date_alignment path/to/file.xlsx [--date 2026-07-14]
    python -m scrub_parser init-db
    python -m scrub_parser list-types
"""

from __future__ import annotations

import argparse
import datetime
import logging
import os
import sys
from pathlib import Path

# Allow running as `python -m scrub_parser` from project root
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scrub_parser.config import SCRUB_CONFIGS


def _get_db():
    from scrub_parser.database import Database
    return Database()


def _get_graph():
    from scrub_parser.sharepoint import GraphClient
    return GraphClient.from_env()


def _get_runner(dry_run=False):
    from scrub_parser.runner import Runner
    return Runner(_get_graph(), _get_db(), dry_run=dry_run)


def setup_logging(verbose: bool = False):
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def parse_date(s: str) -> datetime.date:
    return datetime.datetime.strptime(s, "%Y-%m-%d").date()


def cmd_run_all(args):
    db = _get_db()
    db.create_tables()
    runner = _get_runner(dry_run=args.dry_run)
    stats = runner.run_all(
        start_date=args.start_date,
        end_date=args.end_date,
    )
    print("\n" + "=" * 60)
    print("RUN COMPLETE")
    print("=" * 60)
    print(stats.summary())


def cmd_run_single(args):
    db = _get_db()
    db.create_tables()
    runner = _get_runner(dry_run=args.dry_run)
    stats = runner.run_single(
        args.scrub_key,
        start_date=args.start_date,
        end_date=args.end_date,
    )
    print("\n" + "=" * 60)
    print(f"RUN COMPLETE — {args.scrub_key}")
    print("=" * 60)
    print(stats.summary())


def cmd_parse_local(args):
    from scrub_parser.runner import Runner
    db = _get_db()
    db.create_tables()
    runner = Runner(graph=None, db=db, dry_run=args.dry_run)
    stats = runner.run_from_local_file(
        args.scrub_key,
        args.file_path,
        scrub_date=args.date,
    )
    print("\n" + "=" * 60)
    print(f"LOCAL PARSE COMPLETE — {args.scrub_key}")
    print("=" * 60)
    print(stats.summary())


def cmd_init_db(args):
    db = _get_db()
    db.create_tables()
    print("Database schema initialized.")


def cmd_list_types(args):
    print(f"\n{'KEY':<30} {'DISPLAY NAME':<35} {'SECTION':<12} CATEGORIES")
    print("-" * 110)
    for cfg in SCRUB_CONFIGS:
        cats = ", ".join(o.category for o in cfg.outputs if not o.is_amount)
        sections = ", ".join(sorted({o.section.value for o in cfg.outputs}))
        print(f"{cfg.key:<30} {cfg.display_name:<35} {sections:<12} {cats}")
    print(f"\n{len(SCRUB_CONFIGS)} scrub types configured.")


def main():
    parser = argparse.ArgumentParser(
        prog="scrub_parser",
        description="NCD Scrub Dashboard — automated file parser",
    )
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Enable debug logging")

    sub = parser.add_subparsers(dest="command", required=True)

    # run-all
    p_all = sub.add_parser("run-all", help="Process all scrub types")
    p_all.add_argument("--start-date", type=parse_date, default=None)
    p_all.add_argument("--end-date", type=parse_date, default=None)
    p_all.add_argument("--dry-run", action="store_true",
                       help="Log what would be inserted without writing to DB")
    p_all.set_defaults(func=cmd_run_all)

    # run-single
    p_one = sub.add_parser("run-single", help="Process one scrub type")
    p_one.add_argument("scrub_key", help="Scrub type key (see list-types)")
    p_one.add_argument("--start-date", type=parse_date, default=None)
    p_one.add_argument("--end-date", type=parse_date, default=None)
    p_one.add_argument("--dry-run", action="store_true")
    p_one.set_defaults(func=cmd_run_single)

    # parse-local
    p_local = sub.add_parser("parse-local",
                             help="Parse a local .xlsx file (for testing/backfill)")
    p_local.add_argument("scrub_key")
    p_local.add_argument("file_path")
    p_local.add_argument("--date", type=parse_date, default=None,
                         help="Override scrub date (default: extract from filename)")
    p_local.add_argument("--dry-run", action="store_true")
    p_local.set_defaults(func=cmd_parse_local)

    # init-db
    p_db = sub.add_parser("init-db", help="Create/update database schema")
    p_db.set_defaults(func=cmd_init_db)

    # list-types
    p_list = sub.add_parser("list-types", help="Show all configured scrub types")
    p_list.set_defaults(func=cmd_list_types)

    args = parser.parse_args()
    setup_logging(args.verbose)
    args.func(args)


if __name__ == "__main__":
    main()
