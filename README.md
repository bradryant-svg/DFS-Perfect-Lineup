# DFS Perfect Lineup — NBA Daily Game Plan Bot

Automated DraftKings NBA DFS analysis bot that reads your historical perfect lineup PDFs and NBA season stats feed, then generates a comprehensive daily game plan Excel report — styled after your `Sample.xlsx`.

---

## What It Does

1. **Parses all perfect lineup PDFs** (LineStar/DraftKings exports) to find which players, teams, and opponent matchups keep appearing in winning lineups.
2. **Loads the BigDataBall NBA DFS season feed** (`03-17-2026-nba-season-dfs-feed.xlsx`) to compute real DK fantasy points, salary data, usage rates, and positional stats.
3. **Computes defensive vulnerability rankings** — which teams allow the most DK pts to each position over the last 20 games (L20).
4. **Calculates frequency return rates** — 6x/7x/8x salary hit rates for every player.
5. **Generates a full Excel report** with 10 sheets identical in style to `Sample.xlsx`.

---

## Generated Report Sheets

| Sheet | Contents |
|-------|----------|
| **DFS Rules** | 18 master rules derived from historical perfect lineup trends |
| **Frequency Return (6x-7x-8x)** | Hit rates for every player at today's salary |
| **Slate + Injuries** | Game-by-game breakdown: spread, O/U, Vegas rule, top targets, priority |
| **Position Weakness (L20)** | Avg DK pts allowed per position (PG/SG/SF/PF/C) by each team |
| **Revenge & Former Teams** | Traded players and motivation-boost narrative angles |
| **Lineup Builds** | 3 recommended DraftKings Classic lineups (salary-capped) |
| **Stack Strategy** | Priority stack order for all games on the slate |
| **Full Player Pool** | Every eligible player with frequency stats |
| **Salary Guide** | Positional salary targets and strategy notes |
| **Perfect Lineup Trends** | Historical analysis from your PDF perfect lineup library |

---

## Quick Start

### Install dependencies
```bash
pip install -r requirements.txt
```

### Run the daily report (uses today's date + example matchups)
```bash
python src/generate_report.py
```

### Run with your actual matchups
```bash
python scripts/run_daily.py --matchups "MIA@CHA 7:00PM CHA-3.5 O/U:233.5,DET@WAS 7:00PM DET-15 O/U:232,CLE@MIL 8:00PM CLE-9.5 O/U:227.5"
```

### Matchup format
```
AWAY@HOME TIME SPREAD O/U:NUMBER,...
```
Example: `BOS@MIL 7:30PM -5.5 O/U:228,LAL@GSW 10PM +3 O/U:225`

---

## Files

```
DFS-Perfect-Lineup/
├── src/
│   ├── generate_report.py     # Main report generator (run this daily)
│   ├── pdf_parser.py          # Extracts lineup data from perfect lineup PDFs
│   ├── dfs_feed_loader.py     # Loads & processes the BigDataBall NBA DFS feed
│   ├── analyzer.py            # Additional trend analysis helpers
│   └── scraper.py             # Live web scrapers (DraftKings API, BBRef, etc.)
├── scripts/
│   └── run_daily.py           # Daily CLI runner
├── data/
│   └── reports/               # Generated Excel reports (auto-created)
├── Sample.xlsx                # Reference report format
├── 03-17-2026-nba-season-dfs-feed.xlsx  # BigDataBall NBA DFS data
├── *.pdf                      # Perfect lineup PDF exports (LineStar)
└── requirements.txt
```

---

## Adding New Data

### New perfect lineup PDFs
Drop any new LineStar perfect lineup PDF exports into the repo root. The parser auto-detects all `*.pdf` files except "Starting Lineups" PDFs.

### Updated NBA DFS feed
Replace `03-17-2026-nba-season-dfs-feed.xlsx` with the latest BigDataBall export. The loader auto-finds the most recent `*nba*dfs*feed*.xlsx` file.

---

## How the Report Is Built

### Defensive Vulnerability (Position Weakness L20)
For each team, the system looks at the last 20 game-dates and computes the average DK fantasy points allowed to opposing players at each position (PG/SG/SF/PF/C). Higher = weaker defense = better matchup for offensive players.

### Frequency Return (6x/7x/8x)
For each player's last 20 games with a DK salary:
- **6x target** = salary / 1000 × 6
- **6x hit rate** = % of games where actual DK points ≥ 6x target
- Same logic for 7x and 8x

### Lineup Construction
Uses a greedy salary-cap optimizer:
1. Sort all today's players by adjusted 6x hit rate
2. Fill each DK Classic slot (PG/SG/SF/PF/C/G/F/UTIL) with the best eligible player
3. Respect the $50,000 salary cap with budget reserve per remaining slot

### Stack Strategy
Games are prioritized by:
- Defensive weakness of the opponent (higher L20 pts allowed = better)
- Over/Under (higher = more scoring expected)
- Spread tightness (close spread = both teams score = stack both sides)

---

## Automating Daily Reports

Add to cron (runs at 11 AM every day):
```bash
0 11 * * * cd /path/to/DFS-Perfect-Lineup && python scripts/run_daily.py >> logs/daily.log 2>&1
```

---

## Key Insights From Your Perfect Lineup History

Based on the 16 perfect lineup PDFs provided:

- **Cade Cunningham** is the #1 most frequent perfect lineup player (5 appearances, 64.8 avg DK pts)
- **Reed Sheppard** appears 4 times — excellent value at $5-6k range
- **Alperen Sengun** averages 66.1 DK pts when he lands in perfects
- **Low-salary punts ($3-4k)** appear in 9 out of 10 perfect lineups — essential for salary flexibility
- **Game stacking** (players from both teams) dominates — 100% of perfects use multi-team stacks
