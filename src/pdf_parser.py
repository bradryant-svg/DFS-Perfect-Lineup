"""
PDF Parser — extracts perfect lineup data from LineStar/DraftKings PDF exports.

Each PDF contains one perfect lineup with:
  date, contest_type, position, player, salary, actual_pts
"""

import re
import os
import glob
import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent


def extract_lineup_from_pdf(filepath: str) -> dict:
    """
    Parse a single perfect lineup PDF.
    Returns dict with keys: date, contest, players (list of dicts).
    """
    try:
        import PyPDF2
    except ImportError:
        raise ImportError("Run: pip install PyPDF2")

    data = {"date": None, "contest": "", "players": [], "file": os.path.basename(filepath)}

    try:
        with open(filepath, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            text = ""
            for i in range(min(3, len(reader.pages))):
                text += reader.pages[i].extract_text() + "\n"
    except Exception as exc:
        print(f"[pdf_parser] Could not read {filepath}: {exc}")
        return data

    # ── Date ──────────────────────────────────────────────────────────────
    date_match = re.search(
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d+,\s+\d+", text
    )
    data["date"] = date_match.group(0) if date_match else "Unknown"

    # ── Contest type ──────────────────────────────────────────────────────
    ctype_match = re.search(r"(Late Swap|Early Only|Main|Afternoon),\s*(Main|All Day)?", text)
    data["contest"] = ctype_match.group(0).strip() if ctype_match else "Main"

    # ── Games on the slate ─────────────────────────────────────────────────
    games_match = re.search(r"(\d+)\s+Games", text)
    data["num_games"] = int(games_match.group(1)) if games_match else 0

    # ── Players ───────────────────────────────────────────────────────────
    # Pattern: position + whitespace + player name + $salary + score
    # PyPDF2 extracts the dollar sign as ord 36 ($). Use \$ via a non-raw string
    # to ensure the literal dollar is matched correctly across Python versions.
    player_pattern = re.compile(
        r'(PG|SG|SF|PF|C|G|F|UTIL)\s+([A-Z][\w ]+?)\s+[$](\d+)\s+([\d.]+)'
    )
    for pos, player, salary, pts in player_pattern.findall(text):
        data["players"].append(
            {
                "position": pos.strip(),
                "player": player.strip(),
                "salary": int(salary),
                "actual_pts": float(pts),
            }
        )

    return data


def load_all_perfect_lineups(pdf_dir: str = None) -> pd.DataFrame:
    """
    Parse every PDF in *pdf_dir* (defaults to repo root) and return
    a flat DataFrame of all player appearances in perfect lineups.
    """
    if pdf_dir is None:
        pdf_dir = str(BASE_DIR)

    pdf_files = glob.glob(os.path.join(pdf_dir, "*.pdf"))
    # Exclude the Starting Lineups PDF (not a perfect lineup file)
    pdf_files = [f for f in pdf_files if "Starting Lineup" not in f]

    records = []
    for fp in pdf_files:
        lineup = extract_lineup_from_pdf(fp)
        if not lineup["players"]:
            continue
        for player in lineup["players"]:
            records.append(
                {
                    "date": lineup["date"],
                    "contest": lineup["contest"],
                    "num_games": lineup["num_games"],
                    "file": lineup["file"],
                    **player,
                }
            )

    df = pd.DataFrame(records)
    if not df.empty:
        # Normalise date
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.sort_values("date")
        print(f"[pdf_parser] Loaded {len(df)} player-lineup rows from {len(pdf_files)} PDFs")
    return df


if __name__ == "__main__":
    df = load_all_perfect_lineups()
    print(df.to_string(index=False))
