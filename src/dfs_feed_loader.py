"""
DFS Feed Loader
---------------
Loads and normalises the BigDataBall NBA DFS Excel feed
(e.g. 03-17-2026-nba-season-dfs-feed.xlsx).

Column mapping (row 2 of the sheet):
  BIGDATABALL DATASET | GAME ID | DATE | PLAYER ID | PLAYER |
  OWN TEAM | OPPONENT TEAM | STARTER | VENUE | MINUTES |
  USAGE RATE | DAYS REST | DK_POS | FD_POS | YH_POS |
  DK_SALARY | FD_SALARY | YH_SALARY | DK_FPTS | FD_FPTS | YH_FPTS
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta

BASE_DIR = Path(__file__).parent.parent

CLEAN_COLS = {
    "BIGDATABALL\nDATASET":          "DATASET",
    "GAME ID":                        "GAME_ID",
    "DATE":                           "DATE",
    "PLAYER ID":                      "PLAYER_ID",
    "PLAYER":                         "PLAYER",
    "OWN\nTEAM":                      "TEAM",
    "OPPONENT\nTEAM":                 "OPP",
    "STARTER (Y/N)":                  "STARTER",
    "VENUE (R/H/N)":                  "VENUE",
    "MINUTES":                        "MIN",
    "USAGE RATE":                     "USG",
    'DAYS\nREST\n\n"3+"=season debut\n"0"= back-to-back': "REST",
    "DRAFTKINGS":                     "DK_POS",
    "FANDUEL":                        "FD_POS",
    "YAHOO":                          "YH_POS",
    "for DRAFTKINGS\n\"Classic\" Contests": "DK_SAL",
    "for FANDUEL\n\"Full Roster\" Contests": "FD_SAL",
    "for\nYAHOO\n\"Full Slate\" Contests": "YH_SAL",
    "DK_FPTS":                        "DK_FPTS",
    "FD_FPTS":                        "FD_FPTS",
    "YH_FPTS":                        "YH_FPTS",
}

# Alternate flat header name (pandas adds .1 to duplicate column names)
RENAME_AFTER_LOAD = {
    # Positional columns (appear first in the sheet)
    "DRAFTKINGS":   "DK_POS",
    "FANDUEL":      "FD_POS",
    "YAHOO":        "YH_POS",
    # Salary columns use the long "for DRAFTKINGS..." header — handled in CLEAN_COLS
    # Fantasy-point columns are the .1 duplicates
    "DRAFTKINGS.1": "DK_FPTS",
    "FANDUEL.1":    "FD_FPTS",
    "YAHOO.1":      "YH_FPTS",
}


def find_feed_file() -> Path:
    """Find the most recent nba-season-dfs-feed xlsx in the repo root."""
    candidates = sorted(BASE_DIR.glob("*nba*dfs*feed*.xlsx"), reverse=True)
    if candidates:
        return candidates[0]
    candidates = sorted(BASE_DIR.glob("*.xlsx"), reverse=True)
    for c in candidates:
        if "Sample" not in c.name:
            return c
    return None


def load_dfs_feed(filepath: str = None) -> pd.DataFrame:
    """
    Load the BigDataBall DFS feed and return a clean DataFrame.
    Row 1 = multi-header; row 2 = actual column names; data starts row 4.
    """
    if filepath is None:
        fp = find_feed_file()
        if fp is None:
            print("[dfs_feed] No feed file found.")
            return pd.DataFrame()
        filepath = str(fp)

    print(f"[dfs_feed] Loading {filepath} ...")

    # Read with header=1 (0-indexed) so row 2 becomes column names
    df = pd.read_excel(
        filepath,
        sheet_name="NBA-2025-26-DFS",
        header=1,        # row index 1 = row 2 in spreadsheet
        skiprows=[2],    # skip the note row (row 3)
        engine="openpyxl",
    )

    # Drop completely empty rows
    df = df.dropna(how="all")

    # Rename duplicate column names that pandas suffixes with .1 .2
    df.rename(columns=RENAME_AFTER_LOAD, inplace=True)

    # Apply clean names
    df.rename(columns=CLEAN_COLS, inplace=True)

    # Drop rows without a player name
    if "PLAYER" in df.columns:
        df = df[df["PLAYER"].notna() & (df["PLAYER"] != "PLAYER")]

    # Parse date
    if "DATE" in df.columns:
        df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")

    # Numeric coercion
    for col in ["MIN", "USG", "DK_SAL", "DK_FPTS", "FD_FPTS", "YH_FPTS"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Normalise team names to uppercase abbreviations (feed uses full city names)
    team_map = {
        "Atlanta": "ATL", "Boston": "BOS", "Brooklyn": "BKN",
        "Charlotte": "CHA", "Chicago": "CHI", "Cleveland": "CLE",
        "Dallas": "DAL", "Denver": "DEN", "Detroit": "DET",
        "Golden State": "GSW", "Houston": "HOU", "Indiana": "IND",
        "LA Clippers": "LAC", "LA Lakers": "LAL", "Los Angeles Clippers": "LAC",
        "Los Angeles Lakers": "LAL", "Memphis": "MEM", "Miami": "MIA",
        "Milwaukee": "MIL", "Minnesota": "MIN", "New Orleans": "NOP",
        "New York": "NYK", "Oklahoma City": "OKC", "Orlando": "ORL",
        "Philadelphia": "PHI", "Phoenix": "PHX", "Portland": "POR",
        "Sacramento": "SAC", "San Antonio": "SAS", "Toronto": "TOR",
        "Utah": "UTA", "Washington": "WAS",
    }
    for col in ["TEAM", "OPP"]:
        if col in df.columns:
            df[col] = df[col].map(team_map).fillna(df[col])

    print(f"[dfs_feed] Loaded {len(df)} rows, {df['DATE'].nunique()} game-dates")
    return df


# ---------------------------------------------------------------------------
# Computed metrics
# ---------------------------------------------------------------------------

def position_weakness_l20(df: pd.DataFrame, n_games: int = 20) -> pd.DataFrame:
    """
    For each team return avg DK FPTS allowed to each opposing position
    over the last *n_games* games they played.

    Returns: DataFrame [TEAM, PG, SG, SF, PF, C, TOTAL, Weakest, Rating]
    """
    if df.empty:
        return pd.DataFrame()

    # Only use rows where we have DK_FPTS and position
    needed = ["DATE", "OPP", "DK_POS", "DK_FPTS"]
    if not all(c in df.columns for c in needed):
        return pd.DataFrame()

    sub = df[needed].dropna(subset=["DK_FPTS", "OPP", "DK_POS"]).copy()

    # Get last N game-dates per team (as opponent)
    # Keep OPP as a regular column while filtering; don't use it as groupby key
    sub = sub.copy()
    sub["_rank"] = sub.groupby("OPP")["DATE"].rank(method="dense", ascending=True)
    max_rank_per_opp = sub.groupby("OPP")["_rank"].transform("max")
    sub = sub[sub["_rank"] > (max_rank_per_opp - n_games)]
    sub = sub.drop(columns=["_rank"]).reset_index(drop=True)

    # Flatten DK position (e.g. "PF/C" → "PF" then "C")
    rows = []
    for _, row in sub.iterrows():
        positions = [p.strip() for p in str(row["DK_POS"]).split("/")]
        for pos in positions:
            if pos in ("PG", "SG", "SF", "PF", "C"):
                rows.append({"TEAM": row["OPP"], "POS": pos, "DK_FPTS": row["DK_FPTS"]})
    flat = pd.DataFrame(rows)

    if flat.empty:
        return pd.DataFrame()

    pivot = (
        flat.groupby(["TEAM", "POS"])["DK_FPTS"]
        .mean()
        .unstack(fill_value=0)
        .round(1)
    )
    for pos in ["PG", "SG", "SF", "PF", "C"]:
        if pos not in pivot.columns:
            pivot[pos] = 0.0

    pivot = pivot[["PG", "SG", "SF", "PF", "C"]]
    pivot["TOTAL"] = pivot.sum(axis=1).round(1)

    # Weakest position(s)
    def weakest(row):
        top = row[["PG", "SG", "SF", "PF", "C"]].nlargest(2)
        return "+".join(f"{pos}({val})" for pos, val in top.items())

    pivot["Weakest"] = pivot.apply(weakest, axis=1)

    # Star rating
    max_total = pivot["TOTAL"].max()
    def star_rate(total):
        pct = total / max_total if max_total > 0 else 0
        filled = round(pct * 5)
        return "★" * filled + "☆" * (5 - filled)

    pivot["Rating"] = pivot["TOTAL"].apply(star_rate)
    pivot = pivot.reset_index().rename(columns={"TEAM": "Team"})
    pivot = pivot.sort_values("TOTAL", ascending=False)
    return pivot


def frequency_return(df: pd.DataFrame, n_games: int = 20) -> pd.DataFrame:
    """
    For each player compute:
      - L20 avg DK FPTS
      - 6x hit rate  (actual >= salary/1000 * 6)
      - 7x hit rate
      - 8x hit rate
      - Floor (10th percentile)
      - Ceiling (90th percentile)
    Using all games where they have a DK salary entry.
    """
    if df.empty:
        return pd.DataFrame()

    needed = ["DATE", "PLAYER", "TEAM", "OPP", "DK_POS", "DK_SAL", "DK_FPTS"]
    if not all(c in df.columns for c in needed):
        return pd.DataFrame()

    sub = df[needed].dropna(subset=["DK_FPTS", "DK_SAL"]).copy()
    sub = sub[sub["DK_SAL"] > 0]

    records = []
    for player, grp in sub.groupby("PLAYER"):
        grp = grp.sort_values("DATE").tail(n_games)
        if len(grp) < 3:
            continue

        l20_avg = grp["DK_FPTS"].mean()
        # Use last salary for targets
        last_sal = grp["DK_SAL"].iloc[-1]
        last_team = grp["TEAM"].iloc[-1]
        last_opp = grp["OPP"].iloc[-1]
        last_pos = grp["DK_POS"].iloc[-1]

        tgt_6x = last_sal / 1000 * 6
        tgt_7x = last_sal / 1000 * 7
        tgt_8x = last_sal / 1000 * 8

        hits_6x = (grp["DK_FPTS"] >= tgt_6x).sum()
        hits_7x = (grp["DK_FPTS"] >= tgt_7x).sum()
        hits_8x = (grp["DK_FPTS"] >= tgt_8x).sum()

        n = len(grp)
        records.append({
            "Player":    player,
            "Tm":        last_team,
            "vs":        last_opp,
            "Pos":       last_pos,
            "Salary":    f"${int(last_sal):,}",
            "L20 Avg":   round(l20_avg, 1),
            "6x Tgt":    round(tgt_6x, 1),
            "6x Hits":   f"{hits_6x}/{n}",
            "ADJ 6x%":   f"{round(hits_6x/n*100)}%",
            "7x Tgt":    round(tgt_7x, 1),
            "7x Hits":   f"{hits_7x}/{n}",
            "ADJ 7x%":   f"{round(hits_7x/n*100)}%",
            "8x Hits":   f"{hits_8x}/{n}",
            "ADJ 8x%":   f"{round(hits_8x/n*100)}%",
            "Floor":     round(float(grp["DK_FPTS"].quantile(0.10)), 1),
            "Ceiling":   round(float(grp["DK_FPTS"].quantile(0.90)), 1),
            # raw for sorting
            "_6x_pct":   hits_6x / n,
            "_sal_raw":  last_sal,
        })

    freq_df = pd.DataFrame(records).sort_values("_6x_pct", ascending=False)
    return freq_df


def player_season_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Season averages per player."""
    if df.empty:
        return pd.DataFrame()
    needed = ["PLAYER", "TEAM", "OPP", "DK_POS", "DK_SAL", "DK_FPTS", "MIN"]
    sub = df[[c for c in needed if c in df.columns]].dropna(subset=["DK_FPTS"])
    agg = (
        sub.groupby("PLAYER")
        .agg(
            TEAM=("TEAM", "last"),
            POS=("DK_POS", "last"),
            FPPG=("DK_FPTS", "mean"),
            SALARY=("DK_SAL", "last"),
            GP=("DK_FPTS", "count"),
            MIN=("MIN", "mean"),
        )
        .round({"FPPG": 1, "MIN": 1})
        .reset_index()
        .sort_values("FPPG", ascending=False)
    )
    return agg


if __name__ == "__main__":
    df = load_dfs_feed()
    print(df.head(5).to_string())
    pw = position_weakness_l20(df)
    print("\nPosition Weakness L20:")
    print(pw.head(10).to_string(index=False))
