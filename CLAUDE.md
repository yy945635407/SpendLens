# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A family bill analyzer web app that ingests iCost-format Excel files, generates analysis (charts, PPT, PDF), and provides budget/rule/simulation tools. Single-directory, zero-database design.

## Commands

```bash
# Start dev server (port 5050, debug auto-reload)
/Users/ylyt_bot/miniconda3/bin/python3 app.py

# Kill any process on port 5050
lsof -ti:5050 | xargs kill -9

# Run analyzer directly on a local Excel file
/Users/ylyt_bot/miniconda3/bin/python3 -c "
import json
from analyzer import analyze
r = analyze('/path/to/iCost_202605.xlsx')
print(json.dumps(r, ensure_ascii=False, indent=2))
"
```

No tests, no linter, no build step. Python 3.13 via miniconda at `/Users/ylyt_bot/miniconda3/bin/python3`.

## Architecture

Data flows: **Excel → analyzer.py → {charts, budget, ppt/pdf builder} → HTTP response or file**

### Core Modules

| File | Role |
|------|------|
| `app.py` | Flask server — 10 routes, all file I/O via temp files, no concurrency concerns |
| `analyzer.py` | Parses iCost Excel (openpyxl, sheet `收支账单`). Column mapping: `row[0]=date, row[1]=type(收入/支出), row[2]=amount, row[3]=cat1, row[4]=cat2, row[5]=account, row[7]=note, row[9]=tag`. Expense amounts are negative in source, made absolute in output. Returns ~20-field dict |
| `charts.py` | matplotlib xkcd-mode charts → PNG BytesIO. Font auto-detection for Chinese (scans system font dirs). Output never touches disk |
| `ppt_builder.py` | python-pptx, 10 slides with pink cartoon theme. Color palette in `C` dict |
| `pdf_builder.py` | fpdf2, uses Arial Unicode font at `/Library/Fonts/Arial Unicode.ttf` for Chinese. `_rgb_tuple()` handles both hex strings and tuples |
| `budget.py` | 6 concerns in one file: health scoring, budget CRUD, rule engine, what-if simulation, share snapshots, monthly data persistence. All storage via JSON files in project dir |

### Key Data Flow

`analyze()` returns a dict consumed by all downstream consumers. Fields include: `month`, `total_income`, `total_expense`, `balance`, `savings_rate`, `cat1_list`, `daily_list`, `weekly_list`, `food_list`, `account_list`, `income_list`, `health`, `budget_comparison`.

After `analyze()`, `app.py` attaches `health_score()` results and `compare_budget()` results to the dict before passing to chart/PPT/PDF builders.

### Storage

All runtime data stored as JSON files in project root or subdirs — **no database**:
- `budget_config.json` — budget categories and limits
- `auto_rules.json` — tag rules (pattern → tag)
- `shares/<uuid>.json` — share snapshots
- `data/<YYYY-MM>.json` — persisted monthly data for comparison

### Color Palette (Pink Cartoon Theme)

Consistent across all outputs (PPT, PDF, HTML, charts):
- Primary: `#FF6B8A` (hot pink)
- Dark bg: `#5C2D3E`
- Positive: `#7BC8A4` (mint green)
- Gold accent: `#FFD4B8`
- Muted text: `#C4909E`
- Background: `#FFF0F5`

### Font Handling

- **matplotlib**: Auto-scans system font dirs for Chinese-capable fonts (Kaiti, PingFang, STHeiti). `_find_cn_font()` in charts.py.
- **fpdf2 PDF**: Hardcoded to `/Library/Fonts/Arial Unicode.ttf`, registered as font family `'CN'`. TTC fonts don't work with fpdf2.
- **PPT/HTML**: Use system fonts via CSS/font-family, no special handling needed.

### PPT Slide Structure (10 slides)

1. Cover (dark bg, decorative circles)
2. Overview (6 stat cards + health score banner)
3. Spending pie chart + category breakdown
4. Food bar chart
5. Weekly trend line chart
6. Account bar chart
7. Income doughnut chart
8. Calendar heatmap
9. Budget vs actual comparison
10. Summary & suggestions (dark bg)

`CARD_COLORS` list controls stat card accent colors; `_hex_to_rgb()` converts health score hex to `RGBColor` for pptx.

### Share Feature

`POST /share` accepts `{"data": <analyze result dict>}`, saves as JSON snapshot with a 10-char UUID, returns `share_url`. `GET /share/<id>` renders `share.html` (Jinja2 template, read-only view). No auth — URL is the access token.
