#!/usr/bin/env python3
"""
Daily runner — call this script each day to regenerate the game plan report.

Usage:
    python scripts/run_daily.py
    python scripts/run_daily.py --matchups "BOS@MIL 7:30PM -5.5 O/U:228,LAL@GSW 10PM +3 O/U:225"

Add to cron (runs at 11 AM daily):
    0 11 * * * cd /path/to/DFS-Perfect-Lineup && python scripts/run_daily.py
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from generate_report import generate_report, parse_matchups
import argparse
from datetime import datetime


def main():
    parser = argparse.ArgumentParser(description="Daily DFS Game Plan Generator")
    parser.add_argument("--matchups", default="",
                        help="Matchups string: 'AWAY@HOME TIME SPREAD O/U:XXX,...'")
    parser.add_argument("--date", default=datetime.now().strftime("%B %d, %Y"))
    args = parser.parse_args()

    matchups = parse_matchups(args.matchups) if args.matchups else []

    out_path = generate_report(matchups=matchups, date_str=args.date)
    print(f"Report generated: {out_path}")


if __name__ == "__main__":
    main()
