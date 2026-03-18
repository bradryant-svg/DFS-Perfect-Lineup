"""
DFS Perfect Lineup Scraper
Scrapes DraftKings NBA contest results and optimal lineups from:
- RotoGrinders
- DraftKings lobby results
- NumberFire
- FantasyPros
"""

import requests
import pandas as pd
import json
import time
import os
import re
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from pathlib import Path

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/html",
}

DATA_DIR = Path(__file__).parent.parent / "data" / "raw"
DATA_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# DraftKings API helpers
# ---------------------------------------------------------------------------

def fetch_dk_contests(sport="NBA", date_str: str = None) -> list[dict]:
    """
    Pull today's (or a given date's) DraftKings NBA contest list from
    the public DraftKings lobby API.
    """
    if date_str is None:
        date_str = datetime.now().strftime("%Y-%m-%d")

    url = "https://www.draftkings.com/lobby/getcontests"
    params = {"sport": sport}

    try:
        resp = requests.get(url, params=params, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        contests = data.get("Contests", [])
        print(f"[scraper] Found {len(contests)} DK contests for {sport}")
        return contests
    except Exception as exc:
        print(f"[scraper] DK contest fetch failed: {exc}")
        return []


def fetch_dk_draftables(draft_group_id: int) -> list[dict]:
    """Return draftable players for a DK draft group."""
    url = f"https://api.draftkings.com/lineups/v1/draftgroups/{draft_group_id}/draftables"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        return resp.json().get("draftables", [])
    except Exception as exc:
        print(f"[scraper] DK draftables fetch failed: {exc}")
        return []


# ---------------------------------------------------------------------------
# RotoGrinders optimal lineup scraper
# ---------------------------------------------------------------------------

def scrape_rotogrinders_optimal(date_str: str = None) -> pd.DataFrame:
    """
    Scrape RotoGrinders' posted optimal / winning lineups for a given date.
    Returns DataFrame with columns:
        date, contest_name, position, player, team, opponent,
        salary, actual_pts, ownership
    """
    if date_str is None:
        date_str = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

    url = f"https://rotogrinders.com/resultsdb/nba?date={date_str}&site=draftkings"
    records = []

    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")

        # RotoGrinders renders data in a JSON blob inside a <script> tag
        for script in soup.find_all("script"):
            if script.string and "resultsData" in (script.string or ""):
                raw = re.search(r"resultsData\s*=\s*(\[.*?\]);", script.string, re.DOTALL)
                if raw:
                    results = json.loads(raw.group(1))
                    for entry in results:
                        lineup = entry.get("lineup", [])
                        for slot in lineup:
                            records.append(
                                {
                                    "date": date_str,
                                    "contest_name": entry.get("contest_name", ""),
                                    "contest_type": entry.get("contest_type", ""),
                                    "position": slot.get("position", ""),
                                    "player": slot.get("name", ""),
                                    "team": slot.get("team", ""),
                                    "opponent": slot.get("opp", ""),
                                    "salary": slot.get("salary", 0),
                                    "actual_pts": slot.get("fpts", 0),
                                    "ownership": slot.get("ownership", 0),
                                }
                            )
    except Exception as exc:
        print(f"[scraper] RotoGrinders scrape failed: {exc}")

    df = pd.DataFrame(records)
    if not df.empty:
        path = DATA_DIR / f"rotogrinders_{date_str}.csv"
        df.to_csv(path, index=False)
        print(f"[scraper] Saved {len(df)} rows to {path}")
    return df


# ---------------------------------------------------------------------------
# Basketball-Reference game log scraper
# ---------------------------------------------------------------------------

def fetch_bbref_game_log(season: int = 2025) -> pd.DataFrame:
    """
    Scrape Basketball-Reference schedule & results for box score links,
    then pull player game logs.
    Returns a combined player game log DataFrame.
    """
    url = f"https://www.basketball-reference.com/leagues/NBA_{season}_games.html"
    records = []

    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
        table = soup.find("table", {"id": "schedule"})
        if table is None:
            print("[scraper] BBRef schedule table not found")
            return pd.DataFrame()

        rows = table.find("tbody").find_all("tr")
        for row in rows:
            if row.get("class") and "thead" in row.get("class"):
                continue
            cells = row.find_all("td")
            if len(cells) < 5:
                continue
            records.append(
                {
                    "date": row.find("th").get_text(strip=True) if row.find("th") else "",
                    "away_team": cells[1].get_text(strip=True),
                    "away_pts": cells[2].get_text(strip=True),
                    "home_team": cells[3].get_text(strip=True),
                    "home_pts": cells[4].get_text(strip=True),
                }
            )
    except Exception as exc:
        print(f"[scraper] BBRef game log fetch failed: {exc}")

    df = pd.DataFrame(records)
    if not df.empty:
        path = DATA_DIR / f"bbref_schedule_{season}.csv"
        df.to_csv(path, index=False)
        print(f"[scraper] Saved {len(df)} BBRef game rows to {path}")
    return df


# ---------------------------------------------------------------------------
# NBA Stats API – player game logs (official stats.nba.com)
# ---------------------------------------------------------------------------

def fetch_nba_player_gamelog(season: str = "2024-25", per_mode: str = "Totals") -> pd.DataFrame:
    """
    Pull full player game log for the current season from stats.nba.com.
    """
    url = "https://stats.nba.com/stats/leaguegamelog"
    params = {
        "Counter": "1000",
        "DateFrom": "",
        "DateTo": "",
        "Direction": "DESC",
        "LeagueID": "00",
        "PlayerOrTeam": "P",
        "Season": season,
        "SeasonType": "Regular Season",
        "Sorter": "DATE",
        "PerMode": per_mode,
    }
    nba_headers = {
        **HEADERS,
        "Referer": "https://www.nba.com/",
        "Accept": "application/json",
        "x-nba-stats-origin": "stats",
        "x-nba-stats-token": "true",
    }

    try:
        resp = requests.get(url, params=params, headers=nba_headers, timeout=30)
        resp.raise_for_status()
        raw = resp.json()
        result_set = raw["resultSets"][0]
        columns = result_set["headers"]
        rows = result_set["rowSet"]
        df = pd.DataFrame(rows, columns=columns)
        path = DATA_DIR / f"nba_gamelog_{season.replace('-','_')}.csv"
        df.to_csv(path, index=False)
        print(f"[scraper] Saved {len(df)} NBA game log rows to {path}")
        return df
    except Exception as exc:
        print(f"[scraper] NBA stats API failed: {exc}")
        return pd.DataFrame()


# ---------------------------------------------------------------------------
# Convenience runner
# ---------------------------------------------------------------------------

def run_daily_scrape(date_str: str = None) -> dict:
    """Run all scrapers and return dict of DataFrames."""
    print(f"\n{'='*60}")
    print(f"  DFS SCRAPER  |  {date_str or datetime.now().strftime('%Y-%m-%d')}")
    print(f"{'='*60}\n")

    results = {}
    results["rotogrinders"] = scrape_rotogrinders_optimal(date_str)
    time.sleep(2)
    results["nba_gamelog"] = fetch_nba_player_gamelog()
    return results


if __name__ == "__main__":
    run_daily_scrape()
