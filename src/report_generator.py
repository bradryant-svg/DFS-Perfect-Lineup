"""
DFS Daily Game Plan Report Generator
--------------------------------------
Produces a rich text / HTML report with:
  1. Today's matchups and projected slate
  2. Top player recommendations per game
  3. Defensive vulnerability rankings
  4. Perfect lineup trends (opponent patterns)
  5. Recommended DraftKings lineup (6-man: PG/SG/SF/PF/C/FLEX + CPT equivalent)
"""

import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
from tabulate import tabulate

from analyzer import (
    load_nba_gamelog,
    load_perfect_lineups,
    analyze_opponent_defense,
    analyze_perfect_lineup_trends,
    build_game_plan,
)

REPORTS_DIR = Path(__file__).parent.parent / "data" / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

BANNER = """
╔══════════════════════════════════════════════════════════════════╗
║          🏀  DFS PERFECT LINEUP  –  DAILY GAME PLAN  🏀          ║
╚══════════════════════════════════════════════════════════════════╝
"""

DK_POSITIONS = ["PG", "SG", "SF", "PF", "C", "UTIL"]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _section(title: str, char: str = "─") -> str:
    line = char * 66
    return f"\n{line}\n  {title}\n{line}\n"


def _fmt_table(df: pd.DataFrame, cols: list = None, max_rows: int = 20) -> str:
    if df is None or df.empty:
        return "  (no data)\n"
    sub = df[cols].head(max_rows) if cols else df.head(max_rows)
    return tabulate(sub, headers="keys", tablefmt="simple", showindex=False) + "\n"


def star_rating(score: float, max_score: float) -> str:
    """Convert numeric score to ★ string."""
    if max_score == 0:
        return "☆☆☆☆☆"
    ratio = score / max_score
    filled = round(ratio * 5)
    return "★" * filled + "☆" * (5 - filled)


# ---------------------------------------------------------------------------
# Report sections
# ---------------------------------------------------------------------------

def section_header(date_str: str) -> str:
    return (
        BANNER
        + f"  Generated : {datetime.now().strftime('%A, %B %d %Y  %I:%M %p')}\n"
        + f"  Slate Date: {date_str}\n"
        + f"  Site      : DraftKings  |  Sport: NBA\n"
    )


def section_matchup_overview(matchups: list[dict]) -> str:
    out = _section("TODAY'S MATCHUPS")
    if not matchups:
        return out + "  No matchups loaded.\n"
    for m in matchups:
        home = m.get("home", "???")
        away = m.get("away", "???")
        tipoff = m.get("time", "TBD")
        out += f"  {away:>5}  @  {home:<5}   {tipoff}\n"
    return out


def section_defensive_rankings(df: pd.DataFrame) -> str:
    out = _section("DEFENSIVE VULNERABILITY  (most DK pts allowed to opponents)")
    opp_def = analyze_opponent_defense(df, n_games=15)
    if opp_def.empty:
        return out + "  (no data)\n"

    opp_def = opp_def.copy()
    max_pts = opp_def["avg_dk_fpts_allowed"].max()
    opp_def["rating"] = opp_def["avg_dk_fpts_allowed"].apply(
        lambda x: star_rating(x, max_pts)
    )
    cols = ["OPP", "avg_dk_fpts_allowed", "games", "rating"]
    available_cols = [c for c in cols if c in opp_def.columns]
    out += _fmt_table(opp_def, cols=available_cols, max_rows=30)
    out += "\n  ★★★★★ = softest defense (target)   ☆☆☆☆☆ = toughest defense (avoid)\n"
    return out


def section_perfect_lineup_trends(pl_df: pd.DataFrame) -> str:
    out = _section("PERFECT LINEUP TRENDS  (historical DraftKings contest winners)")
    trends = analyze_perfect_lineup_trends(pl_df)

    if not trends:
        return out + "  (no perfect lineup data – run scraper to populate)\n"

    out += "\n  ── Top Players by Appearances in Winning Lineups ──\n"
    if "top_players" in trends:
        out += _fmt_table(trends["top_players"], max_rows=15)

    out += "\n  ── Opponents Most Targeted in Winning Lineups ──\n"
    if "top_opponents" in trends:
        out += (
            "  These teams appear most as the OPPONENT when players\n"
            "  land in perfect lineups — meaning they allow the most fantasy production.\n\n"
        )
        out += _fmt_table(trends["top_opponents"], max_rows=15)

    out += "\n  ── Best Value Picks (avg DK pts per $1k salary) ──\n"
    if "top_value_players" in trends:
        out += _fmt_table(trends["top_value_players"], max_rows=10)

    return out


def section_player_recommendations(game_plan: pd.DataFrame, matchups: list[dict]) -> str:
    out = _section("PLAYER RECOMMENDATIONS BY MATCHUP")

    if game_plan.empty:
        return out + "  (no game plan data)\n"

    for m in matchups:
        home = m.get("home", "").upper()
        away = m.get("away", "").upper()
        game_label = f"{away} @ {home}"
        out += f"\n  ┌─ {game_label} ─{'─' * max(0, 50 - len(game_label))}┐\n"

        sub = game_plan[game_plan["TEAM"].isin([home, away])].head(10)
        if sub.empty:
            out += "  │  No players found.\n"
        else:
            for _, row in sub.iterrows():
                marker = "🔥" if row["PROJ_DK_FPTS"] >= 45 else ("✅" if row["PROJ_DK_FPTS"] >= 35 else "  ")
                vs_opp = row["VS_OPP_AVG"] if row["VS_OPP_AVG"] != "N/A" else "---"
                out += (
                    f"  │ {marker} {row['PLAYER']:<22} {row['TEAM']:<4} vs {row['OPP']:<4}"
                    f"  Proj: {row['PROJ_DK_FPTS']:>5}  L5: {row['LAST5_AVG']:>5}"
                    f"  vOpp: {str(vs_opp):>5}\n"
                )
        out += f"  └{'─' * 56}┘\n"
    return out


def section_top_targets(game_plan: pd.DataFrame, top_n: int = 20) -> str:
    out = _section(f"TOP {top_n} TARGETS ACROSS FULL SLATE")
    if game_plan.empty:
        return out + "  (no data)\n"

    top = game_plan.head(top_n).copy()
    top["RANK"] = range(1, len(top) + 1)
    cols = ["RANK", "PLAYER", "TEAM", "OPP", "PROJ_DK_FPTS", "LAST5_AVG", "VS_OPP_AVG", "OPP_DEF_RANK"]
    available = [c for c in cols if c in top.columns]
    out += _fmt_table(top, cols=available, max_rows=top_n)
    return out


def section_recommended_lineup(game_plan: pd.DataFrame, salary_cap: int = 50000) -> str:
    """
    Build a recommended 8-player DraftKings Classic lineup:
    PG, SG, SF, PF, C, G, F, UTIL (within salary cap)
    Uses a greedy approach: best projected player per position.
    """
    out = _section("RECOMMENDED DRAFTKINGS LINEUP  (salary-aware, $50,000 cap)")

    if game_plan.empty or "PROJ_DK_FPTS" not in game_plan.columns:
        return out + "  (insufficient data to build lineup)\n"

    # Simulate salary if not present
    gp = game_plan.copy()
    if "SALARY" not in gp.columns:
        # Estimate salary from projected score (proxy)
        proj_max = gp["PROJ_DK_FPTS"].max()
        gp["SALARY"] = (gp["PROJ_DK_FPTS"] / proj_max * 10000).round(-2).clip(3000, 10000).astype(int)

    # DraftKings Classic NBA slots
    slots = [
        ("PG", ["PG"]),
        ("SG", ["SG"]),
        ("SF", ["SF"]),
        ("PF", ["PF"]),
        ("C",  ["C"]),
        ("G",  ["PG", "SG"]),
        ("F",  ["SF", "PF"]),
        ("UTIL", ["PG", "SG", "SF", "PF", "C"]),
    ]

    used_players = set()
    lineup = []
    remaining_salary = salary_cap

    for slot_name, eligible_positions in slots:
        if "POSITION" in gp.columns:
            candidates = gp[
                gp["POSITION"].isin(eligible_positions)
                & ~gp["PLAYER"].isin(used_players)
                & (gp["SALARY"] <= remaining_salary - (len(slots) - len(lineup) - 1) * 3000)
            ]
        else:
            candidates = gp[~gp["PLAYER"].isin(used_players)]

        if candidates.empty:
            candidates = gp[~gp["PLAYER"].isin(used_players)]

        if candidates.empty:
            break

        pick = candidates.sort_values("PROJ_DK_FPTS", ascending=False).iloc[0]
        lineup.append(
            {
                "SLOT": slot_name,
                "PLAYER": pick["PLAYER"],
                "TEAM": pick["TEAM"],
                "OPP": pick.get("OPP", ""),
                "PROJ_DK_FPTS": pick["PROJ_DK_FPTS"],
                "SALARY": pick["SALARY"],
            }
        )
        used_players.add(pick["PLAYER"])
        remaining_salary -= pick["SALARY"]

    if not lineup:
        return out + "  (could not build lineup)\n"

    lu_df = pd.DataFrame(lineup)
    total_proj = lu_df["PROJ_DK_FPTS"].sum()
    total_sal = lu_df["SALARY"].sum()

    out += _fmt_table(lu_df, max_rows=10)
    out += f"\n  Total Projected DK FPTS : {total_proj:.2f}\n"
    out += f"  Total Salary Used       : ${total_sal:,}  (cap: ${salary_cap:,})\n"
    out += f"  Remaining Salary        : ${salary_cap - total_sal:,}\n"
    return out


def section_key_insights(df: pd.DataFrame, pl_df: pd.DataFrame, game_plan: pd.DataFrame) -> str:
    out = _section("KEY INSIGHTS & TRENDS")

    insights = []

    # Softest defenses today
    opp_def = analyze_opponent_defense(df)
    if not opp_def.empty and "OPP" in opp_def.columns and not game_plan.empty:
        todays_opps = game_plan["OPP"].dropna().unique().tolist()
        soft = opp_def[opp_def["OPP"].isin(todays_opps)].head(3)
        for _, row in soft.iterrows():
            insights.append(
                f"🎯 {row['OPP']} has allowed {row['avg_dk_fpts_allowed']:.1f} avg DK pts "
                f"to opponents over last 15 games – prioritize players facing them."
            )

    # Hot players (last 5 > season avg by 20%)
    if not game_plan.empty:
        hot = game_plan[
            (game_plan["LAST5_AVG"] > game_plan["SEASON_AVG"] * 1.20)
            & (game_plan["LAST5_AVG"] >= 35)
        ].head(5)
        for _, row in hot.iterrows():
            insights.append(
                f"🔥 {row['PLAYER']} is ON FIRE: L5 avg {row['LAST5_AVG']:.1f} vs "
                f"season avg {row['SEASON_AVG']:.1f} (+{((row['LAST5_AVG']/row['SEASON_AVG'])-1)*100:.0f}%)"
            )

    # Perfect lineup repeat offenders
    trends = analyze_perfect_lineup_trends(pl_df)
    if "top_opponents" in trends and not trends["top_opponents"].empty:
        top_opp = trends["top_opponents"].iloc[0]
        insights.append(
            f"📊 Historically, players facing {top_opp['opponent']} appear most often "
            f"in perfect lineups ({top_opp['times_targeted']} times, "
            f"{top_opp['avg_pts_allowed']:.1f} avg DK pts)."
        )

    if not insights:
        insights.append("Run the scraper to populate trend data for richer insights.")

    for insight in insights:
        out += f"\n  {insight}\n"

    return out


# ---------------------------------------------------------------------------
# Master report builder
# ---------------------------------------------------------------------------

def generate_daily_report(
    todays_matchups: list[dict],
    gamelog_path: str = None,
    date_str: str = None,
) -> str:
    """
    Full pipeline:
      1. Load data
      2. Build game plan
      3. Compose report
      4. Write to file + return as string
    """
    if date_str is None:
        date_str = datetime.now().strftime("%Y-%m-%d")

    print(f"[report] Loading NBA game log...")
    df = load_nba_gamelog(gamelog_path)

    print(f"[report] Loading perfect lineups...")
    pl_df = load_perfect_lineups()

    print(f"[report] Building game plan for {len(todays_matchups)} matchups...")
    game_plan = build_game_plan(df, todays_matchups)

    # Compose report
    report = ""
    report += section_header(date_str)
    report += section_matchup_overview(todays_matchups)
    report += section_defensive_rankings(df)
    report += section_perfect_lineup_trends(pl_df)
    report += section_player_recommendations(game_plan, todays_matchups)
    report += section_top_targets(game_plan)
    report += section_recommended_lineup(game_plan)
    report += section_key_insights(df, pl_df, game_plan)

    footer = (
        "\n"
        + "─" * 66
        + "\n"
        + "  DISCLAIMER: This report is for educational/entertainment use only.\n"
        + "  DFS involves financial risk. Always conduct your own research.\n"
        + "─" * 66
        + "\n"
    )
    report += footer

    # Save report
    out_path = REPORTS_DIR / f"dfs_game_plan_{date_str}.txt"
    out_path.write_text(report)
    print(f"[report] Report saved to {out_path}")

    return report


if __name__ == "__main__":
    # Example usage with hardcoded matchups
    matchups = [
        {"home": "BOS", "away": "NYK", "time": "7:30 PM ET"},
        {"home": "GSW", "away": "LAL", "time": "10:00 PM ET"},
        {"home": "DEN", "away": "PHX", "time": "9:00 PM ET"},
    ]
    report = generate_daily_report(matchups)
    print(report)
