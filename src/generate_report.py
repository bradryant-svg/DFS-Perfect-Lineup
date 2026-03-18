"""
DFS Daily Game Plan Report Generator
=====================================
Produces an Excel workbook (.xlsx) that mirrors the Sample.xlsx format:

  Sheet 1 – DFS Rules          (master rules derived from historical trends)
  Sheet 2 – Frequency Return   (6x/7x/8x hit rates for today's players)
  Sheet 3 – Slate + Injuries   (matchup analysis, spreads, O/U, key info)
  Sheet 4 – Position Weakness  (L20 DK pts allowed per position by team)
  Sheet 5 – Revenge & Former   (traded/revenge narrative players)
  Sheet 6 – Lineup Builds      (3 recommended lineup constructions)
  Sheet 7 – Stack Strategy     (priority stacks by game)
  Sheet 8 – Full Player Pool   (all eligible players with freq stats)
  Sheet 9 – Salary Guide       (positional salary targets)

Usage:
    python generate_report.py
    python generate_report.py --date 2026-03-18 --matchups "BOS@MIL,LAL@GSW"
"""

import sys
import os
import argparse
import pandas as pd
import numpy as np
from datetime import datetime, date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# Make sure src/ is importable
sys.path.insert(0, str(Path(__file__).parent))

from pdf_parser import load_all_perfect_lineups
from dfs_feed_loader import (
    load_dfs_feed,
    position_weakness_l20,
    frequency_return,
    player_season_stats,
)

BASE_DIR = Path(__file__).parent.parent
REPORTS_DIR = BASE_DIR / "data" / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

# ── Colour palette ────────────────────────────────────────────────────────
C_HEADER_BG    = "1F3864"   # dark navy
C_HEADER_FG    = "FFFFFF"
C_SUBHEAD_BG   = "2F5597"   # medium navy
C_SUBHEAD_FG   = "FFFFFF"
C_GREEN_BG     = "C6EFCE"   # light green (≥80% 6x)
C_GREEN_FG     = "276221"
C_LGREEN_BG    = "FFEB9C"   # yellow (≥55%)
C_LGREEN_FG    = "9C5700"
C_RED_BG       = "FFC7CE"   # red (expensive/low return)
C_RED_FG       = "9C0006"
C_ALT_ROW      = "EBF0FA"   # alternating row tint
C_GOLD         = "FFD700"
C_ORANGE_BG    = "FCE4D6"
C_TITLE_BG     = "16365C"


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _font(bold=False, color="000000", size=11) -> Font:
    return Font(bold=bold, color=color, size=size)


def _border() -> Border:
    thin = Side(style="thin", color="AAAAAA")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _align(horizontal="left", wrap=False) -> Alignment:
    return Alignment(horizontal=horizontal, vertical="center", wrap_text=wrap)


def _set_col_widths(ws, widths: dict):
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def _header_row(ws, row: int, values: list, bg=C_HEADER_BG, fg=C_HEADER_FG, size=11):
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = _font(bold=True, color=fg, size=size)
        cell.fill = _fill(bg)
        cell.alignment = _align("center")
        cell.border = _border()


def _data_row(ws, row: int, values: list, alt=False, bold=False, bg=None):
    bg_color = bg or (C_ALT_ROW if alt else "FFFFFF")
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = _font(bold=bold)
        cell.fill = _fill(bg_color)
        cell.alignment = _align()
        cell.border = _border()


# ============================================================================
# Sheet 1 — DFS Rules
# ============================================================================

MASTER_RULES = [
    ("#",  "Rule",                                      "Evidence"),
    ("1",  "SALARY $49,500-$50,000",                    "Use every dollar."),
    ("2",  "CHEAPEST = HIGH VALUE. $3-4k punt.",         "9/10 perfects."),
    ("3",  "MID-PRICE $5-7.5k WINS.",                   "~50% of perfects."),
    ("4R", "HIGH O/U + CLOSE SPREAD = #1 (REVISED)",    "Close spread > high O/U alone."),
    ("5",  "SWEET SPOT SPREAD: 7-12 pts",               "Allows both teams to score."),
    ("6",  "GAME STACK: players from both sides",        "Multi-team stacks dominate."),
    ("7",  "AVOID BLOWOUT LOSERS",                      "Garbage time minutes unreliable."),
    ("8",  "TARGET WORST DEFENSES (L20 DK pts allowed)","Exploit weakest Ds."),
    ("9",  "BACK-TO-BACK = RISK",                       "Avoid 0-day rest stars."),
    ("10", "OWNERSHIP <20% = GPP EDGE",                 "Contrarian plays for tournaments."),
    ("11", "3+ REST DAYS = FRESH LEGS",                 "Better performance after rest."),
    ("12", "HOME TEAM BIAS in close spreads",           "Home teams cover more often."),
    ("13", "POSITION COMBO: C+PG stack wins most",      "Historical perfect lineup data."),
    ("14", "BLOWOUT FAVORITE: stack winners only",      "Losers' pts redistributed late."),
    ("15", "PDF PLAYER (3x+ in perfects) = LOCK",       "High-frequency historical picks."),
    ("16", "TANKING TEAM OPPONENTS = VALUE",             "Blow-out + weak D = feast."),
    ("17", "USAGE RATE >25% = ELITE TARGET",            "High-usage = more DK pts."),
    ("18", "SALARY GUIDE: 1 superstar + 2 mid + punt",  "Proven stack construction."),
]


def build_rules_sheet(ws, pl_df: pd.DataFrame, today_str: str):
    ws.title = "DFS Rules"
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = f"MASTER DFS RULES — Based on {len(pl_df['file'].unique()) if not pl_df.empty else 0} Perfect Lineups Analysed"
    c.font = _font(bold=True, color=C_HEADER_FG, size=13)
    c.fill = _fill(C_TITLE_BG)
    c.alignment = _align("center")

    ws.row_dimensions[1].height = 28

    # Today note
    ws.merge_cells("A2:C2")
    ws["A2"].value = f"Report Generated: {today_str}  |  Recommendations powered by historical PDF data + BigDataBall DFS Feed"
    ws["A2"].font = _font(size=10, color="444444")
    ws["A2"].alignment = _align("center")

    # Headers
    _header_row(ws, 4, ["#", "Rule", "Evidence"], size=11)
    ws.row_dimensions[4].height = 20

    # Rules
    for i, (num, rule, evidence) in enumerate(MASTER_RULES[1:], start=5):
        alt = (i % 2 == 0)
        _data_row(ws, i, [num, rule, evidence], alt=alt)
        ws.row_dimensions[i].height = 18

    _set_col_widths(ws, {"A": 6, "B": 50, "C": 55})


# ============================================================================
# Sheet 2 — Frequency Return (6x/7x/8x)
# ============================================================================

def build_frequency_sheet(ws, freq_df: pd.DataFrame, today_str: str):
    ws.title = "Frequency Return (6x-7x-8x)"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:P1")
    ws["A1"].value = f"FREQUENCY RETURN — 6x/7x/8x Hit Rates (L20 × Matchup Boost)  |  {today_str}"
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=12)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    ws.merge_cells("A2:P2")
    ws["A2"].value = "Green ≥80% 6x | Light green ≥55% | Yellow ≥40% | Red = expensive + low return. Sorted by adj 6x%."
    ws["A2"].font = _font(size=9, color="333333")
    ws["A2"].alignment = _align("center")

    headers = [
        "Player", "Tm", "vs", "Pos", "Salary",
        "L20 Avg", "6x Tgt", "6x Hits", "ADJ 6x%",
        "7x Tgt", "7x Hits", "ADJ 7x%",
        "8x Hits", "ADJ 8x%", "Floor", "Ceiling"
    ]
    _header_row(ws, 4, headers)

    if freq_df.empty:
        ws.cell(row=5, column=1, value="Run the scraper/loader to populate this sheet.")
        return

    display_cols = [c for c in headers if c in freq_df.columns]
    for row_idx, (_, row) in enumerate(freq_df[display_cols].iterrows(), start=5):
        alt = row_idx % 2 == 0
        values = [row.get(c, "") for c in display_cols]

        # Colour code based on ADJ 6x%
        raw_6x = float(str(row.get("ADJ 6x%", "0%")).replace("%", "") or 0)
        if raw_6x >= 80:
            bg = C_GREEN_BG
        elif raw_6x >= 55:
            bg = C_LGREEN_BG
        elif raw_6x >= 40:
            "F2F2F2"
            bg = None
        else:
            bg = None  # low rate stays white

        _data_row(ws, row_idx, values, alt=alt, bg=bg)
        ws.row_dimensions[row_idx].height = 16

    _set_col_widths(ws, {
        "A": 24, "B": 6, "C": 6, "D": 8, "E": 9,
        "F": 9, "G": 8, "H": 8, "I": 8,
        "J": 8, "K": 8, "L": 8,
        "M": 8, "N": 8, "O": 8, "P": 10,
    })


# ============================================================================
# Sheet 3 — Slate + Injuries
# ============================================================================

def build_slate_sheet(ws, matchups: list[dict], today_str: str,
                      pos_weak: pd.DataFrame, freq_df: pd.DataFrame):
    ws.title = "Slate + Injuries"
    ws.sheet_view.showGridLines = False

    title = f"NBA DFS GAME PLAN — {datetime.now().strftime('%A %B %d, %Y')}"
    ws.merge_cells("A1:I1")
    ws["A1"].value = title
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=14)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    # Slate summary
    games_str = " • ".join(f"{m['away']}@{m['home']}" for m in matchups) if matchups else "No matchups loaded"
    ws.merge_cells("A2:I2")
    ws["A2"].value = f"{len(matchups)}-Game Slate | {games_str}"
    ws["A2"].font = _font(size=10)
    ws["A2"].alignment = _align("center")

    headers = ["Game", "Time", "Spread", "O/U", "Vegas Rule",
               "Key Info", "Environment", "Top Targets", "Priority"]
    _header_row(ws, 4, headers)

    if not matchups:
        ws.cell(row=5, column=1, value="No matchups provided. Use --matchups flag.")
        return

    # For each matchup, build an insight row
    pw_lookup = {}
    if not pos_weak.empty and "Team" in pos_weak.columns:
        pw_lookup = pos_weak.set_index("Team")["TOTAL"].to_dict()

    # Top targets from frequency sheet per game
    def top_targets_for_game(away, home):
        if freq_df.empty:
            return "Check frequency sheet"
        teams = [away, home]
        sub = freq_df[freq_df["Tm"].isin(teams)].head(4)
        if sub.empty:
            return "—"
        return ", ".join(
            f"{row['Player']} {row['Salary']} ({row['ADJ 6x%']} 6x)"
            for _, row in sub.iterrows()
        )

    def build_rule_note(away, home, spread_pts, ou):
        notes = []
        away_def = pw_lookup.get(away, 150)
        home_def = pw_lookup.get(home, 150)
        worst_def = away if away_def > home_def else home
        worst_val = max(away_def, home_def)

        if worst_val >= 165:
            notes.append(f"Rule 8: {worst_def} allows {worst_val:.0f} L20!")
        if abs(spread_pts) >= 15:
            notes.append(f"Rule 14: BLOWOUT ({abs(spread_pts):.0f} pts)")
        if abs(spread_pts) <= 12 and ou >= 228:
            notes.append("Rule 4R: HIGH O/U + CLOSE SPREAD")
        elif 7 <= abs(spread_pts) <= 12:
            notes.append("Rule 5: Sweet spot spread")
        if ou >= 230:
            notes.append(f"Rule 4R: O/U {ou} ELITE")
        return " + ".join(notes) if notes else "Standard game"

    def priority_stars(spread, ou):
        score = 0
        if ou >= 230: score += 2
        elif ou >= 225: score += 1
        if abs(spread) <= 7: score += 2
        elif abs(spread) <= 12: score += 1
        worst_d = max(pw_lookup.values()) if pw_lookup else 150
        if worst_d >= 165: score += 2
        filled = min(score, 5)
        return "★" * filled + "☆" * (5 - filled)

    row_num = 5
    for m in matchups:
        away = m.get("away", "???").upper()
        home = m.get("home", "???").upper()
        game = f"{away}@{home}"
        time = m.get("time", "TBD")
        spread_str = m.get("spread", "—")
        ou = float(str(m.get("ou", "225")).replace("+", "") or 225)
        spread_pts = float(str(m.get("spread_pts", "5")).replace("+", "").replace("-", "") or 5)

        rule_note = build_rule_note(away, home, spread_pts, ou)
        targets = top_targets_for_game(away, home)

        away_def = pw_lookup.get(away, 0)
        home_def = pw_lookup.get(home, 0)
        env = (
            f"{away} allows {away_def:.0f} L20 | "
            f"{home} allows {home_def:.0f} L20 | "
            f"O/U {ou}"
        )

        key_info = m.get("key_info", f"{away} vs {home}. Check injury report.")
        priority = priority_stars(spread_pts, ou)

        values = [game, time, spread_str, ou, rule_note, key_info, env, targets, priority]
        _data_row(ws, row_num, values, alt=(row_num % 2 == 0), bold=False)
        ws.row_dimensions[row_num].height = 60
        # Wrap text in key columns
        for col_idx in [5, 6, 7, 8]:
            ws.cell(row=row_num, column=col_idx).alignment = _align("left", wrap=True)
        row_num += 1

    _set_col_widths(ws, {
        "A": 12, "B": 10, "C": 12, "D": 7, "E": 30,
        "F": 35, "G": 35, "H": 40, "I": 10,
    })


# ============================================================================
# Sheet 4 — Position Weakness (L20)
# ============================================================================

def build_pos_weakness_sheet(ws, pos_weak: pd.DataFrame, matchups: list[dict], today_str: str):
    ws.title = "Position Weakness (L20)"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:J1")
    ws["A1"].value = f"DEFENSIVE POSITION WEAKNESS — L20 Games (Avg DK Pts ALLOWED to Opposing Starters)  |  {today_str}"
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=12)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    ws.merge_cells("A2:J2")
    ws["A2"].value = "Higher = weaker defense (exploit). Red cells ≥37. Yellow ≥33."
    ws["A2"].font = _font(size=9)
    ws["A2"].alignment = _align("center")

    headers = ["Team", "PG", "SG", "SF", "PF", "C", "TOTAL", "Weakest", "Tonight's Exploit", "Rating"]
    _header_row(ws, 4, headers)

    today_teams = set()
    for m in matchups:
        today_teams.add(m.get("away", "").upper())
        today_teams.add(m.get("home", "").upper())

    team_opp_map = {}
    for m in matchups:
        away = m.get("away", "").upper()
        home = m.get("home", "").upper()
        team_opp_map[away] = home
        team_opp_map[home] = away

    if pos_weak.empty:
        ws.cell(row=5, column=1, value="Run the DFS feed loader to compute this sheet.")
        return

    row_num = 5
    for _, r in pos_weak.iterrows():
        team = str(r.get("Team", ""))
        pg_val = float(r.get("PG", 0) or 0)
        sg_val = float(r.get("SG", 0) or 0)
        sf_val = float(r.get("SF", 0) or 0)
        pf_val = float(r.get("PF", 0) or 0)
        c_val  = float(r.get("C", 0) or 0)
        total  = float(r.get("TOTAL", 0) or 0)
        weakest = str(r.get("Weakest", ""))
        rating = str(r.get("Rating", ""))

        # Build exploit note
        opp = team_opp_map.get(team, "")
        if opp:
            exploit = f"{opp} faces {team} tonight. Exploit: {weakest}"
        elif team in today_teams:
            exploit = f"Playing tonight. Exploit: {weakest}"
        else:
            exploit = ""

        values = [team, pg_val, sg_val, sf_val, pf_val, c_val, total, weakest, exploit, rating]
        alt = row_num % 2 == 0
        _data_row(ws, row_num, values, alt=alt)

        # Colour code weakness cells (cols 2-6)
        for col_idx, val in enumerate([pg_val, sg_val, sf_val, pf_val, c_val], start=2):
            cell = ws.cell(row=row_num, column=col_idx)
            if val >= 37:
                cell.fill = _fill(C_RED_BG)
                cell.font = _font(bold=True, color=C_RED_FG)
            elif val >= 33:
                cell.fill = _fill(C_LGREEN_BG)
                cell.font = _font(color=C_LGREEN_FG)

        ws.row_dimensions[row_num].height = 18
        row_num += 1

    _set_col_widths(ws, {
        "A": 8, "B": 8, "C": 8, "D": 8, "E": 8, "F": 8,
        "G": 9, "H": 25, "I": 45, "J": 10,
    })


# ============================================================================
# Sheet 5 — Revenge & Former Teams
# ============================================================================

def build_revenge_sheet(ws, pl_df: pd.DataFrame, df: pd.DataFrame,
                        matchups: list[dict], today_str: str):
    ws.title = "Revenge & Former Teams"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    ws["A1"].value = f"REVENGE GAMES & FORMER TEAM CONNECTIONS — {today_str}"
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=12)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    ws.merge_cells("A2:H2")
    ws["A2"].value = "Players facing former teams, traded mid-season, or with motivation boosts. Higher DFS value."
    ws["A2"].font = _font(size=9)
    ws["A2"].alignment = _align("center")

    headers = ["Player", "Now", "vs", "Salary", "Connection", "DFS Impact", "Adj 6x%", "Rating"]
    _header_row(ws, 4, headers)

    today_teams = set()
    for m in matchups:
        today_teams.add(m.get("away", "").upper())
        today_teams.add(m.get("home", "").upper())

    # Derive potential revenge plays from frequency sheet
    if df.empty or not today_teams:
        ws.cell(row=5, column=1, value="Add matchup data to see revenge plays.")
        return

    # Find players who were recently traded (played for multiple teams this season)
    needed = ["PLAYER", "TEAM", "DATE"]
    if not all(c in df.columns for c in needed):
        ws.cell(row=5, column=1, value="Insufficient data columns.")
        return

    # Players on today's teams
    playing_players = df[df["TEAM"].isin(today_teams)][["PLAYER", "TEAM"]].drop_duplicates()

    # Find players who have records on other teams (traded)
    all_team_records = df[["PLAYER", "TEAM"]].drop_duplicates()
    multi_team = (
        all_team_records.groupby("PLAYER")["TEAM"]
        .nunique()
        .reset_index()
        .rename(columns={"TEAM": "num_teams"})
    )
    multi_team = multi_team[multi_team["num_teams"] > 1]

    revenge_players = playing_players[playing_players["PLAYER"].isin(multi_team["PLAYER"])]

    # Get salary and 6x from freq
    freq_lookup = {}
    if not pl_df.empty and "player" in pl_df.columns:
        freq_lookup = (
            pl_df.groupby("player")["actual_pts"]
            .count()
            .to_dict()
        )

    row_num = 5
    for _, rp in revenge_players.iterrows():
        player = rp["PLAYER"]
        current_team = rp["TEAM"]

        # What teams did they previously play for?
        all_teams = all_team_records[all_team_records["PLAYER"] == player]["TEAM"].tolist()
        prev_teams = [t for t in all_teams if t != current_team]

        if not prev_teams:
            continue

        # Today's opponent
        opp = None
        for m in matchups:
            if current_team in [m.get("away", "").upper(), m.get("home", "").upper()]:
                opp = m.get("home", "").upper() if current_team == m.get("away", "").upper() else m.get("away", "").upper()
                break

        connection = f"Former team(s): {', '.join(set(prev_teams))}"
        is_revenge = opp and opp in prev_teams
        rating = "★★★★★" if is_revenge else "★★★"
        impact = f"Facing ex-team {opp}!" if is_revenge else f"Traded player on {current_team}"

        # Lookup salary
        sal_row = df[df["PLAYER"] == player].sort_values("DATE").tail(1)
        salary = f"${int(sal_row['DK_SAL'].iloc[0]):,}" if not sal_row.empty and "DK_SAL" in sal_row.columns and not pd.isna(sal_row["DK_SAL"].iloc[0]) else "—"

        appearances = freq_lookup.get(player, 0)
        adj_6x = f"{min(appearances*5, 50)}%" if appearances > 0 else "—"

        _data_row(ws, row_num, [player, current_team, opp or "—", salary, connection, impact, adj_6x, rating],
                  alt=(row_num % 2 == 0))
        ws.row_dimensions[row_num].height = 18
        row_num += 1

    if row_num == 5:
        ws.cell(row=5, column=1, value="No traded/revenge players identified for today's slate.")

    _set_col_widths(ws, {
        "A": 24, "B": 6, "C": 6, "D": 9,
        "E": 40, "F": 35, "G": 10, "H": 10,
    })


# ============================================================================
# Sheet 6 — Lineup Builds
# ============================================================================

def build_lineup_builds_sheet(ws, freq_df: pd.DataFrame, pos_weak: pd.DataFrame,
                               matchups: list[dict], today_str: str):
    ws.title = "Lineup Builds"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")
    ws["A1"].value = f"LINEUP BUILDS — {today_str} ({len(matchups)}-Game Slate)"
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=13)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    today_teams = set()
    for m in matchups:
        today_teams.add(m.get("away", "").upper())
        today_teams.add(m.get("home", "").upper())

    if freq_df.empty or not today_teams:
        ws.cell(row=3, column=1, value="No frequency data available. Run the loader first.")
        return

    # Filter to today's teams
    today_freq = freq_df[freq_df["Tm"].isin(today_teams)].copy()

    # Build 3 lineups
    DK_SLOTS = ["PG", "SG", "SF", "PF", "C", "G", "F", "UTIL"]
    SLOT_ELIGIBLE = {
        "PG": ["PG"],
        "SG": ["SG"],
        "SF": ["SF"],
        "PF": ["PF"],
        "C":  ["C"],
        "G":  ["PG", "SG"],
        "F":  ["SF", "PF"],
        "UTIL": ["PG", "SG", "SF", "PF", "C"],
    }
    SALARY_CAP = 50000

    def build_one_lineup(df_pool: pd.DataFrame, label: str, exclusions: set = None) -> list:
        used = set(exclusions or [])
        lineup = []
        remaining = SALARY_CAP

        def parse_salary(s):
            return int(str(s).replace("$", "").replace(",", "") or 0)

        for slot in DK_SLOTS:
            eligible_pos = SLOT_ELIGIBLE[slot]
            candidates = df_pool[
                df_pool["Pos"].apply(lambda p: any(ep in str(p) for ep in eligible_pos))
                & ~df_pool["Player"].isin(used)
            ].copy()
            candidates["_sal"] = candidates["Salary"].apply(parse_salary)
            budget_reserve = (8 - len(lineup) - 1) * 3000
            candidates = candidates[candidates["_sal"] <= remaining - budget_reserve]

            if candidates.empty:
                candidates = df_pool[~df_pool["Player"].isin(used)].copy()
                candidates["_sal"] = candidates["Salary"].apply(parse_salary)
                candidates = candidates[candidates["_sal"] <= remaining - budget_reserve]

            if candidates.empty:
                lineup.append({"slot": slot, "player": "—", "team": "—",
                                "salary": 0, "proj": 0, "edge": "No eligible player"})
                continue

            pick = candidates.sort_values("_6x_pct", ascending=False).iloc[0]
            lineup.append({
                "slot": slot,
                "player": pick["Player"],
                "team": pick["Tm"],
                "salary": pick["_sal"],
                "proj": pick["L20 Avg"],
                "edge": f"vs {pick['vs']} | {pick['ADJ 6x%']} 6x | L20:{pick['L20 Avg']}",
            })
            used.add(pick["Player"])
            remaining -= pick["_sal"]

        return lineup

    row_num = 3
    for lu_num in range(1, 4):
        lu_label = f"LU{lu_num}: RECOMMENDED LINEUP {lu_num}"
        ws.merge_cells(f"A{row_num}:E{row_num}")
        ws[f"A{row_num}"].value = lu_label
        ws[f"A{row_num}"].font = _font(bold=True, color="FFFFFF", size=11)
        ws[f"A{row_num}"].fill = _fill(C_SUBHEAD_BG)
        ws[f"A{row_num}"].alignment = _align("center")

        # Stack different teams for diversity
        lineup = build_one_lineup(today_freq, lu_label)

        headers = ["Slot", "Player", "Team", "Salary", "Proj Pts", "Edge"]
        _header_row(ws, row_num + 1, headers, bg=C_HEADER_BG)

        total_sal = 0
        total_proj = 0.0
        for j, slot_info in enumerate(lineup, start=row_num + 2):
            _data_row(ws, j, [
                slot_info["slot"],
                slot_info["player"],
                slot_info["team"],
                f"${slot_info['salary']:,}" if slot_info["salary"] else "—",
                slot_info["proj"],
                slot_info["edge"],
            ], alt=(j % 2 == 0))
            ws.row_dimensions[j].height = 18
            total_sal += slot_info["salary"]
            total_proj += slot_info["proj"]

        summary_row = row_num + 2 + len(lineup)
        _data_row(ws, summary_row, [
            "TOTAL", "", "", f"${total_sal:,}", round(total_proj, 1), f"Salary remaining: ${SALARY_CAP - total_sal:,}"
        ], bold=True, bg=C_ALT_ROW)

        row_num = summary_row + 2  # gap between lineups

    _set_col_widths(ws, {"A": 8, "B": 26, "C": 7, "D": 10, "E": 10, "F": 45})


# ============================================================================
# Sheet 7 — Stack Strategy
# ============================================================================

def build_stack_strategy_sheet(ws, matchups: list[dict], pos_weak: pd.DataFrame,
                                freq_df: pd.DataFrame, today_str: str):
    ws.title = "Stack Strategy"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:G1")
    ws["A1"].value = f"STACK STRATEGY — {today_str}"
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=13)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    headers = ["Priority", "Game", "Stars", "Vegas", "Core Stack", "Bring-Back", "Why"]
    _header_row(ws, 3, headers)

    pw_lookup = {}
    if not pos_weak.empty and "Team" in pos_weak.columns:
        pw_lookup = pos_weak.set_index("Team")["TOTAL"].to_dict()

    def game_stars(away, home, spread_pts, ou):
        score = 0
        worst_d = max(pw_lookup.get(away, 0), pw_lookup.get(home, 0))
        if worst_d >= 165: score += 2
        elif worst_d >= 155: score += 1
        if abs(spread_pts) <= 7: score += 2
        elif abs(spread_pts) <= 12: score += 1
        if ou >= 230: score += 2
        elif ou >= 225: score += 1
        f = min(score, 5)
        return "★" * f + "☆" * (5 - f)

    def top_players_for_team(team, n=3):
        if freq_df.empty:
            return team + " players"
        sub = freq_df[freq_df["Tm"] == team].head(n)
        if sub.empty:
            return f"{team} (no freq data)"
        return " + ".join(
            f"{row['Player']} {row['Salary']} ({row['ADJ 6x%']} 6x)"
            for _, row in sub.iterrows()
        )

    # Sort matchups by priority
    def matchup_score(m):
        ou = float(str(m.get("ou", "225")) or 225)
        sp = abs(float(str(m.get("spread_pts", "5")).replace("+", "").replace("-", "") or 5))
        worst_d = max(
            pw_lookup.get(m.get("away", "").upper(), 0),
            pw_lookup.get(m.get("home", "").upper(), 0)
        )
        return ou + (15 - sp) + worst_d / 10

    sorted_matchups = sorted(matchups, key=matchup_score, reverse=True)

    row_num = 4
    for priority, m in enumerate(sorted_matchups, 1):
        away = m.get("away", "???").upper()
        home = m.get("home", "???").upper()
        game = f"{away}@{home}"
        ou = m.get("ou", "225")
        spread_str = m.get("spread", "—")
        spread_pts = float(str(m.get("spread_pts", "5")).replace("+", "").replace("-", "") or 5)

        stars = game_stars(away, home, spread_pts, float(str(ou).replace("+", "") or 225))

        vegas = f"O/U {ou}\n{spread_str}"
        core = top_players_for_team(away, 3) + "\n" + top_players_for_team(home, 2)
        bring_back = top_players_for_team(away if spread_pts > 0 else home, 2)

        away_def = pw_lookup.get(away, 0)
        home_def = pw_lookup.get(home, 0)
        why_parts = []
        if max(away_def, home_def) >= 165:
            worst_t = away if away_def > home_def else home
            why_parts.append(f"Rule 8: {worst_t} allows {max(away_def, home_def):.0f} L20!")
        if abs(spread_pts) <= 7:
            why_parts.append(f"Rule 4R: Close spread ({spread_pts:.0f} pts)")
        elif 7 < abs(spread_pts) <= 12:
            why_parts.append(f"Rule 5: Sweet spot spread")
        elif abs(spread_pts) >= 15:
            why_parts.append(f"Rule 14: Blowout ({spread_pts:.0f} pts)")
        if float(str(ou).replace("+","") or 0) >= 230:
            why_parts.append(f"Rule 4R: Elite O/U {ou}")
        why = ". ".join(why_parts) if why_parts else "Standard slate game."

        _data_row(ws, row_num, [f"#{priority}", game, stars, vegas, core, bring_back, why],
                  alt=(row_num % 2 == 0))
        for col_idx in [4, 5, 6, 7]:
            ws.cell(row=row_num, column=col_idx).alignment = _align("left", wrap=True)
        ws.row_dimensions[row_num].height = 60
        row_num += 1

    _set_col_widths(ws, {
        "A": 8, "B": 12, "C": 10, "D": 14,
        "E": 45, "F": 35, "G": 45,
    })


# ============================================================================
# Sheet 8 — Full Player Pool
# ============================================================================

def build_player_pool_sheet(ws, freq_df: pd.DataFrame, today_str: str,
                             matchups: list[dict]):
    ws.title = "Full Player Pool"
    ws.sheet_view.showGridLines = False

    today_teams = set()
    for m in matchups:
        today_teams.add(m.get("away", "").upper())
        today_teams.add(m.get("home", "").upper())

    ws.merge_cells("A1:I1")
    ws["A1"].value = f"FULL PLAYER POOL — {today_str}"
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=12)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    headers = ["Game", "Player", "Team", "Pos", "Salary", "FPPG", "Adj 6x%", "Adj 7x%", "Notes"]
    _header_row(ws, 3, headers)

    if freq_df.empty:
        ws.cell(row=4, column=1, value="Run the DFS feed loader to populate this sheet.")
        return

    today_pool = freq_df[freq_df["Tm"].isin(today_teams)].copy() if today_teams else freq_df.copy()

    # Attach game string
    team_game_map = {}
    for m in matchups:
        away = m.get("away", "").upper()
        home = m.get("home", "").upper()
        game_str = f"{away}@{home} {m.get('date', today_str)}"
        team_game_map[away] = game_str
        team_game_map[home] = game_str

    row_num = 4
    for _, r in today_pool.iterrows():
        game = team_game_map.get(r.get("Tm", ""), f"{r.get('Tm', '')} game")
        notes = ""
        raw_6x = float(str(r.get("ADJ 6x%", "0%")).replace("%", "") or 0)
        if raw_6x >= 80:
            notes = "🔥 Elite value"
        elif raw_6x >= 60:
            notes = "✅ High floor"

        values = [
            game, r.get("Player"), r.get("Tm"), r.get("Pos"), r.get("Salary"),
            r.get("L20 Avg"), r.get("ADJ 6x%"), r.get("ADJ 7x%"), notes
        ]
        _data_row(ws, row_num, values, alt=(row_num % 2 == 0))
        ws.row_dimensions[row_num].height = 16
        row_num += 1

    _set_col_widths(ws, {
        "A": 28, "B": 24, "C": 6, "D": 10,
        "E": 10, "F": 9, "G": 9, "H": 9, "I": 18,
    })


# ============================================================================
# Sheet 9 — Salary Guide
# ============================================================================

def build_salary_guide_sheet(ws, freq_df: pd.DataFrame, pos_weak: pd.DataFrame,
                              matchups: list[dict], today_str: str):
    ws.title = "Salary Guide"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")
    ws["A1"].value = f"SALARY & POSITION GUIDE — Frequency + L20 Weakness  |  {today_str}"
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=12)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    today_teams = set()
    for m in matchups:
        today_teams.add(m.get("away", "").upper())
        today_teams.add(m.get("home", "").upper())

    # Build top narrative line
    pw_lookup = {}
    if not pos_weak.empty and "Team" in pos_weak.columns:
        pw_lookup = pos_weak.set_index("Team")["TOTAL"].to_dict()
    today_worst = {t: pw_lookup.get(t, 0) for t in today_teams}
    worst_team = max(today_worst, key=today_worst.get) if today_worst else "—"
    worst_val = today_worst.get(worst_team, 0)

    ws.merge_cells("A2:F2")
    ws["A2"].value = (
        f"Today's worst defense: {worst_team} ({worst_val:.0f} L20 pts allowed). "
        "Target players facing them. Use McBride-style punts for salary relief."
    )
    ws["A2"].font = _font(size=9)
    ws["A2"].alignment = _align("center")

    headers = ["Position", "Targets", "Strategy", "Lock", "Avoid", "Freq Notes"]
    _header_row(ws, 4, headers)

    positions = ["PG", "SG", "SF", "PF", "C", "G (PG 70%)"]
    pos_keys   = ["PG", "SG", "SF", "PF", "C", "PG"]

    if freq_df.empty or not today_teams:
        ws.cell(row=5, column=1, value="No data available. Populate with DFS feed loader.")
        return

    today_freq = freq_df[freq_df["Tm"].isin(today_teams)]

    pw_by_pos = {}
    if not pos_weak.empty:
        for _, r in pos_weak.iterrows():
            team = r.get("Team", "")
            if team in today_teams:
                for pos in ["PG", "SG", "SF", "PF", "C"]:
                    pw_by_pos.setdefault(pos, []).append(float(r.get(pos, 0) or 0))

    row_num = 5
    for display_pos, filter_pos in zip(positions, pos_keys):
        sub = today_freq[today_freq["Pos"].str.contains(filter_pos, na=False)]
        targets = ", ".join(
            f"{r['Player']} {r['Salary']}" for _, r in sub.head(3).iterrows()
        ) or "—"
        lock = sub.head(1)["Player"].values[0] if not sub.empty else "—"
        avoid = "Low O/U games"
        avg_allowed = np.mean(pw_by_pos.get(filter_pos, [0]))
        freq_note = (
            f"{filter_pos} weakness: {avg_allowed:.1f} avg allowed tonight"
            if avg_allowed > 0 else "No L20 data"
        )
        strategy = f"Target {filter_pos} vs weak Ds. Stack with game-stack approach."

        _data_row(ws, row_num, [display_pos, targets, strategy, lock, avoid, freq_note],
                  alt=(row_num % 2 == 0))
        ws.row_dimensions[row_num].height = 20
        row_num += 1

    _set_col_widths(ws, {"A": 12, "B": 40, "C": 40, "D": 24, "E": 20, "F": 35})


# ============================================================================
# Sheet 10 — Perfect Lineup Trends
# ============================================================================

def build_trends_sheet(ws, pl_df: pd.DataFrame, today_str: str):
    ws.title = "Perfect Lineup Trends"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")
    ws["A1"].value = f"PERFECT LINEUP HISTORICAL TRENDS — {len(pl_df['file'].unique()) if not pl_df.empty else 0} Slates Analysed"
    ws["A1"].font = _font(bold=True, color=C_HEADER_FG, size=12)
    ws["A1"].fill = _fill(C_TITLE_BG)
    ws["A1"].alignment = _align("center")

    if pl_df.empty:
        ws.cell(row=3, column=1, value="No PDF data loaded.")
        return

    row_num = 3

    # Most frequent players
    ws.cell(row=row_num, column=1).value = "Most Frequent Players in Perfect Lineups"
    ws.cell(row=row_num, column=1).font = _font(bold=True, size=11)
    ws.cell(row=row_num, column=1).fill = _fill(C_SUBHEAD_BG)
    ws.cell(row=row_num, column=1).font = _font(bold=True, color="FFFFFF")
    row_num += 1

    top_players = (
        pl_df.groupby("player")
        .agg(appearances=("player", "count"), avg_pts=("actual_pts", "mean"),
             avg_salary=("salary", "mean"))
        .reset_index()
        .sort_values("appearances", ascending=False)
        .head(20)
    )
    _header_row(ws, row_num, ["Player", "Appearances", "Avg DK Pts", "Avg Salary"])
    row_num += 1
    for _, r in top_players.iterrows():
        _data_row(ws, row_num,
                  [r["player"], r["appearances"], round(r["avg_pts"], 1), f"${int(r['avg_salary']):,}"],
                  alt=(row_num % 2 == 0))
        ws.row_dimensions[row_num].height = 16
        row_num += 1

    row_num += 2

    # Position combo analysis
    ws.cell(row=row_num, column=1).value = "Position Combos in Perfect Lineups"
    ws.cell(row=row_num, column=1).font = _font(bold=True, color="FFFFFF")
    ws.cell(row=row_num, column=1).fill = _fill(C_SUBHEAD_BG)
    row_num += 1

    pos_combo = (
        pl_df.groupby(["file", "date"])["position"]
        .apply(lambda x: " + ".join(sorted(set(x))))
        .reset_index()
        .groupby("position")["file"]
        .count()
        .reset_index()
        .rename(columns={"position": "Combo", "file": "Count"})
        .sort_values("Count", ascending=False)
        .head(15)
    )
    total = pos_combo["Count"].sum()
    pos_combo["Pct"] = (pos_combo["Count"] / total * 100).round(1).astype(str) + "%"

    _header_row(ws, row_num, ["Position Combo", "Count", "Percent"])
    row_num += 1
    for _, r in pos_combo.iterrows():
        _data_row(ws, row_num, [r["Combo"], r["Count"], r["Pct"]], alt=(row_num % 2 == 0))
        row_num += 1

    row_num += 2

    # Salary range analysis
    ws.cell(row=row_num, column=1).value = "Salary Range Distribution in Perfect Lineups"
    ws.cell(row=row_num, column=1).font = _font(bold=True, color="FFFFFF")
    ws.cell(row=row_num, column=1).fill = _fill(C_SUBHEAD_BG)
    row_num += 1

    pl_df_copy = pl_df.copy()
    pl_df_copy["sal_tier"] = pd.cut(
        pl_df_copy["salary"],
        bins=[0, 4000, 5500, 7500, 9000, 15000],
        labels=["$3k-4k (punt)", "$4k-5.5k (value)", "$5.5k-7.5k (mid)", "$7.5k-9k (star)", "$9k+ (superstar)"]
    )
    sal_dist = (
        pl_df_copy["sal_tier"]
        .value_counts()
        .reset_index()
        .rename(columns={"index": "Tier", "sal_tier": "Count"})
    )
    _header_row(ws, row_num, ["Salary Tier", "Count in Perfects"])
    row_num += 1
    for _, r in sal_dist.iterrows():
        _data_row(ws, row_num, [str(r.iloc[0]), r.iloc[1]], alt=(row_num % 2 == 0))
        row_num += 1

    _set_col_widths(ws, {"A": 30, "B": 15, "C": 15, "D": 15})


# ============================================================================
# Master build function
# ============================================================================

def generate_report(matchups: list[dict] = None, date_str: str = None) -> Path:
    if date_str is None:
        date_str = datetime.now().strftime("%B %d, %Y")
    if matchups is None:
        matchups = []

    print("\n" + "="*60)
    print("  DFS PERFECT LINEUP — DAILY GAME PLAN GENERATOR")
    print("="*60 + "\n")

    # 1. Load data
    print("[1/4] Loading NBA DFS feed...")
    df = load_dfs_feed()

    print("[2/4] Parsing perfect lineup PDFs...")
    pl_df = load_all_perfect_lineups()

    print("[3/4] Computing metrics...")
    pos_weak = position_weakness_l20(df) if not df.empty else pd.DataFrame()
    freq_df  = frequency_return(df) if not df.empty else pd.DataFrame()

    print(f"[4/4] Building Excel report for {date_str}...")

    # 2. Create workbook
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    build_rules_sheet(wb.create_sheet(),       pl_df, date_str)
    build_frequency_sheet(wb.create_sheet(),   freq_df, date_str)
    build_slate_sheet(wb.create_sheet(),       matchups, date_str, pos_weak, freq_df)
    build_pos_weakness_sheet(wb.create_sheet(), pos_weak, matchups, date_str)
    build_revenge_sheet(wb.create_sheet(),     pl_df, df, matchups, date_str)
    build_lineup_builds_sheet(wb.create_sheet(), freq_df, pos_weak, matchups, date_str)
    build_stack_strategy_sheet(wb.create_sheet(), matchups, pos_weak, freq_df, date_str)
    build_player_pool_sheet(wb.create_sheet(), freq_df, date_str, matchups)
    build_salary_guide_sheet(wb.create_sheet(), freq_df, pos_weak, matchups, date_str)
    build_trends_sheet(wb.create_sheet(),      pl_df, date_str)

    # 3. Save
    safe_date = datetime.now().strftime("%Y-%m-%d")
    out_path = REPORTS_DIR / f"DFS_Game_Plan_{safe_date}.xlsx"
    wb.save(str(out_path))
    print(f"\n✅ Report saved: {out_path}\n")
    return out_path


# ============================================================================
# CLI entry point
# ============================================================================

def parse_matchups(matchup_str: str) -> list[dict]:
    """
    Parse matchup string like "BOS@MIL 7:30PM -5.5 O/U:228,LAL@GSW 10PM +3 O/U:225"
    into list of dicts.
    """
    matchups = []
    for part in matchup_str.split(","):
        part = part.strip()
        if "@" not in part:
            continue
        tokens = part.split()
        game = tokens[0]
        away, home = game.split("@")
        m = {
            "away": away.upper().strip(),
            "home": home.upper().strip(),
            "time": tokens[1] if len(tokens) > 1 else "TBD",
            "spread": tokens[2] if len(tokens) > 2 else "—",
            "spread_pts": tokens[2].replace("+", "").replace("pk", "0") if len(tokens) > 2 else "5",
            "ou": tokens[3].replace("O/U:", "").replace("O/U", "").strip() if len(tokens) > 3 else "225",
        }
        matchups.append(m)
    return matchups


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate DFS daily game plan report")
    parser.add_argument("--date", default=datetime.now().strftime("%B %d, %Y"),
                        help="Slate date string (default: today)")
    parser.add_argument("--matchups", default="",
                        help="Comma-separated matchups: 'BOS@MIL 7:30PM -5.5 O/U:228'")
    args = parser.parse_args()

    matchups = parse_matchups(args.matchups) if args.matchups else []

    # If no matchups supplied, use today's matchups from the March 17 sample
    if not matchups:
        matchups = [
            {"away": "DET", "home": "WAS", "time": "7:00 PM", "spread": "DET -15",
             "spread_pts": "15", "ou": "232"},
            {"away": "OKC", "home": "ORL", "time": "7:00 PM", "spread": "OKC -9.5",
             "spread_pts": "9.5", "ou": "222.5"},
            {"away": "MIA", "home": "CHA", "time": "7:00 PM", "spread": "CHA -3.5",
             "spread_pts": "3.5", "ou": "233.5"},
            {"away": "IND", "home": "NYK", "time": "7:30 PM", "spread": "NYK -16.5",
             "spread_pts": "16.5", "ou": "225.5"},
            {"away": "CLE", "home": "MIL", "time": "8:00 PM", "spread": "CLE -9.5",
             "spread_pts": "9.5", "ou": "227.5"},
            {"away": "PHX", "home": "MIN", "time": "8:00 PM", "spread": "MIN -5",
             "spread_pts": "5",   "ou": "224"},
        ]

    generate_report(matchups=matchups, date_str=args.date)
