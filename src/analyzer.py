"""
DFS Perfect Lineup Analyzer
----------------------------
Ingests scraped lineup + NBA stats data and surfaces:
  1. Most-frequent players in winning lineups
  2. Opponent (defense) trends – which teams give up the most DK pts
  3. Position-level vulnerability by team
  4. Correlation between salary, ownership, and actual points
  5. Value picks (pts / $1k salary)
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta

DATA_DIR = Path(__file__).parent.parent / "data"
RAW_DIR = DATA_DIR / "raw"
PROC_DIR = DATA_DIR / "processed"
PROC_DIR.mkdir(parents=True, exist_ok=True)

# DraftKings scoring weights for NBA
DK_SCORING = {
    "PTS": 1.0,
    "REB": 1.25,
    "AST": 1.5,
    "STL": 2.0,
    "BLK": 2.0,
    "TOV": -0.5,
    "THREE_PM": 0.5,  # bonus 3PM
    # Double-double and triple-double bonuses applied elsewhere
}

DOUBLE_DOUBLE_BONUS = 1.5
TRIPLE_DOUBLE_BONUS = 3.0


# ---------------------------------------------------------------------------
# DK fantasy point calculation
# ---------------------------------------------------------------------------

def calc_dk_fpts(row: pd.Series) -> float:
    """Calculate DraftKings fantasy points from a box score row."""
    try:
        pts = float(row.get("PTS", 0) or 0)
        reb = float(row.get("REB", 0) or 0)
        ast = float(row.get("AST", 0) or 0)
        stl = float(row.get("STL", 0) or 0)
        blk = float(row.get("BLK", 0) or 0)
        tov = float(row.get("TOV", 0) or 0)
        fg3m = float(row.get("FG3M", 0) or 0)

        total = (
            pts * 1.0
            + reb * 1.25
            + ast * 1.5
            + stl * 2.0
            + blk * 2.0
            + tov * -0.5
            + fg3m * 0.5
        )

        # Double / triple-double bonuses
        dd_cats = sum(
            1
            for v in [pts, reb, ast, stl, blk]
            if v >= 10
        )
        if dd_cats >= 3:
            total += TRIPLE_DOUBLE_BONUS
        elif dd_cats >= 2:
            total += DOUBLE_DOUBLE_BONUS

        return round(total, 2)
    except Exception:
        return 0.0


# ---------------------------------------------------------------------------
# Load & normalize data
# ---------------------------------------------------------------------------

def load_nba_gamelog(filepath: str = None) -> pd.DataFrame:
    """
    Load NBA game log CSV.  Accepts the official stats.nba.com column names
    OR a user-supplied file with similar columns.  Falls back to the most
    recent file in data/raw/.
    """
    if filepath:
        df = pd.read_csv(filepath)
    else:
        # Try most-recent nba_gamelog_*.csv
        files = sorted(RAW_DIR.glob("nba_gamelog_*.csv"), reverse=True)
        if not files:
            # Look for any CSV that looks like box scores
            files = sorted(RAW_DIR.glob("*.csv"), reverse=True)
        if not files:
            print("[analyzer] No NBA game log found. Run scraper first.")
            return pd.DataFrame()
        df = pd.read_csv(files[0])
        print(f"[analyzer] Loaded game log from {files[0]}")

    # Normalize common column aliases
    col_map = {
        "PLAYER_NAME": "PLAYER",
        "TEAM_ABBREVIATION": "TEAM",
        "MATCHUP": "MATCHUP",
        "WL": "WL",
        "GAME_DATE": "DATE",
        "MIN": "MIN",
        "FGM": "FGM", "FGA": "FGA",
        "FG_PCT": "FG_PCT",
        "FG3M": "FG3M", "FG3A": "FG3A",
        "FTM": "FTM", "FTA": "FTA",
        "OREB": "OREB", "DREB": "DREB", "REB": "REB",
        "AST": "AST", "STL": "STL", "BLK": "BLK", "TOV": "TOV",
        "PF": "PF", "PTS": "PTS",
    }
    df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)

    # Extract opponent from MATCHUP string (e.g. "LAL vs. GSW" or "LAL @ GSW")
    if "MATCHUP" in df.columns and "OPP" not in df.columns:
        df["OPP"] = df["MATCHUP"].str.extract(r"(?:vs\.|@)\s+(\w+)")

    # Compute DK fantasy points if not present
    if "DK_FPTS" not in df.columns:
        df["DK_FPTS"] = df.apply(calc_dk_fpts, axis=1)

    # Ensure DATE is parsed
    if "DATE" in df.columns:
        df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")

    return df


def load_perfect_lineups(date_str: str = None) -> pd.DataFrame:
    """Load one or all RotoGrinders perfect lineup CSVs."""
    if date_str:
        f = RAW_DIR / f"rotogrinders_{date_str}.csv"
        if f.exists():
            return pd.read_csv(f)
        return pd.DataFrame()

    files = list(RAW_DIR.glob("rotogrinders_*.csv"))
    if not files:
        return pd.DataFrame()
    dfs = [pd.read_csv(f) for f in files]
    return pd.concat(dfs, ignore_index=True)


# ---------------------------------------------------------------------------
# Trend analysis
# ---------------------------------------------------------------------------

def analyze_opponent_defense(df: pd.DataFrame, n_games: int = 15) -> pd.DataFrame:
    """
    Rank teams by DK points allowed per game (by position).
    Higher = softer defense = better streaming target.
    """
    if df.empty or "OPP" not in df.columns:
        return pd.DataFrame()

    recent = df.copy()
    if "DATE" in recent.columns:
        cutoff = recent["DATE"].max() - timedelta(days=n_games * 1.5)
        recent = recent[recent["DATE"] >= cutoff]

    # If POSITION available, break down by position
    group_cols = ["OPP"]
    if "POSITION" in recent.columns:
        group_cols = ["OPP", "POSITION"]

    agg = (
        recent.groupby(group_cols)["DK_FPTS"]
        .agg(["mean", "count"])
        .rename(columns={"mean": "avg_dk_fpts_allowed", "count": "games"})
        .reset_index()
        .sort_values("avg_dk_fpts_allowed", ascending=False)
    )
    return agg


def analyze_perfect_lineup_trends(pl_df: pd.DataFrame) -> dict:
    """
    From historical perfect lineups, find:
      - Most frequent players
      - Most frequent opponents faced by players in perfect lineups
      - Average salary used
      - Average score
    """
    if pl_df.empty:
        return {}

    results = {}

    # Most-used players
    results["top_players"] = (
        pl_df.groupby("player")
        .agg(
            appearances=("player", "count"),
            avg_pts=("actual_pts", "mean"),
            avg_salary=("salary", "mean"),
            avg_ownership=("ownership", "mean"),
        )
        .reset_index()
        .sort_values("appearances", ascending=False)
        .head(25)
    )

    # Opponents that appear most as the team being scored on
    if "opponent" in pl_df.columns:
        results["top_opponents"] = (
            pl_df.groupby("opponent")
            .agg(
                times_targeted=("opponent", "count"),
                avg_pts_allowed=("actual_pts", "mean"),
            )
            .reset_index()
            .sort_values("times_targeted", ascending=False)
            .head(20)
        )

    # Position-level breakdown
    if "position" in pl_df.columns:
        results["position_breakdown"] = (
            pl_df.groupby("position")
            .agg(
                appearances=("position", "count"),
                avg_pts=("actual_pts", "mean"),
                avg_salary=("salary", "mean"),
            )
            .reset_index()
            .sort_values("avg_pts", ascending=False)
        )

    # Value metric
    pl_df = pl_df.copy()
    pl_df["value"] = pl_df["actual_pts"] / (pl_df["salary"] / 1000)
    results["top_value_players"] = (
        pl_df.groupby("player")["value"]
        .mean()
        .reset_index()
        .sort_values("value", ascending=False)
        .head(20)
    )

    return results


def player_vs_opponent_splits(df: pd.DataFrame, player: str, opponent: str) -> dict:
    """Return a player's stats specifically against one opponent."""
    sub = df[(df["PLAYER"] == player) & (df["OPP"] == opponent)]
    if sub.empty:
        return {}
    return {
        "games": len(sub),
        "avg_dk_fpts": round(sub["DK_FPTS"].mean(), 2),
        "avg_pts": round(sub["PTS"].mean(), 2) if "PTS" in sub.columns else "N/A",
        "avg_reb": round(sub["REB"].mean(), 2) if "REB" in sub.columns else "N/A",
        "avg_ast": round(sub["AST"].mean(), 2) if "AST" in sub.columns else "N/A",
    }


# ---------------------------------------------------------------------------
# Daily game plan builder
# ---------------------------------------------------------------------------

def build_game_plan(
    gamelog_df: pd.DataFrame,
    todays_matchups: list[dict],  # [{"home": "BOS", "away": "NYK"}, ...]
    n_recommendations: int = 10,
) -> pd.DataFrame:
    """
    Given today's matchups, rank players by expected DK value.

    Strategy:
      1. For each player in todays_matchups, compute:
         - Season avg DK FPTS
         - Last 5 game avg DK FPTS (recent form)
         - DK FPTS vs today's opponent (all-time this season)
         - Opponent defensive rank (avg DK FPTS allowed)
      2. Composite score = weighted blend
      3. Return top players per matchup

    Returns DataFrame: player, team, opponent, proj_dk_fpts, last5, vs_opp_avg, def_rank
    """
    if gamelog_df.empty or not todays_matchups:
        return pd.DataFrame()

    opp_def = analyze_opponent_defense(gamelog_df)
    def_lookup = {}
    if not opp_def.empty and "OPP" in opp_def.columns:
        def_lookup = opp_def.set_index("OPP")["avg_dk_fpts_allowed"].to_dict()

    all_teams = set()
    for m in todays_matchups:
        all_teams.add(m.get("home", "").upper())
        all_teams.add(m.get("away", "").upper())

    playing_df = gamelog_df[gamelog_df["TEAM"].isin(all_teams)]
    if playing_df.empty:
        return pd.DataFrame()

    # Build opponent lookup for each team playing today
    team_to_opp = {}
    for m in todays_matchups:
        home = m.get("home", "").upper()
        away = m.get("away", "").upper()
        team_to_opp[home] = away
        team_to_opp[away] = home

    records = []
    for player, grp in playing_df.groupby("PLAYER"):
        team = grp["TEAM"].iloc[-1]
        opp = team_to_opp.get(team, "")

        season_avg = grp["DK_FPTS"].mean()

        # Last 5 games
        last5 = grp.sort_values("DATE").tail(5)["DK_FPTS"].mean() if "DATE" in grp.columns else season_avg

        # vs today's opponent
        vs_opp = grp[grp["OPP"] == opp]["DK_FPTS"].mean() if opp else np.nan

        # Opponent defensive rating (higher = softer defense)
        def_rank_val = def_lookup.get(opp, season_avg)

        # Composite score (weighted)
        composite = (
            0.35 * last5
            + 0.30 * season_avg
            + 0.20 * (vs_opp if not np.isnan(vs_opp) else season_avg)
            + 0.15 * def_rank_val
        )

        records.append(
            {
                "PLAYER": player,
                "TEAM": team,
                "OPP": opp,
                "SEASON_AVG": round(season_avg, 2),
                "LAST5_AVG": round(last5, 2),
                "VS_OPP_AVG": round(vs_opp, 2) if not np.isnan(vs_opp) else "N/A",
                "OPP_DEF_RANK": round(def_rank_val, 2),
                "PROJ_DK_FPTS": round(composite, 2),
            }
        )

    result = pd.DataFrame(records).sort_values("PROJ_DK_FPTS", ascending=False)
    path = PROC_DIR / f"game_plan_{datetime.now().strftime('%Y_%m_%d')}.csv"
    result.to_csv(path, index=False)
    return result


if __name__ == "__main__":
    df = load_nba_gamelog()
    if not df.empty:
        print(df.head())
        opp_def = analyze_opponent_defense(df)
        print(opp_def.head(10))
