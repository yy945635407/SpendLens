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

# Generate full PPT + PDF for visual testing
/Users/ylyt_bot/miniconda3/bin/python3 -c "
from analyzer import analyze
from charts import all_charts
from ppt_builder import build_ppt
from pdf_builder import build_pdf
data = analyze('/Users/ylyt_bot/Desktop/家庭账单/iCost_202605.xlsx')
charts = all_charts(data)
with open('/tmp/test_pink.pptx','wb') as f: f.write(build_ppt(data, charts, theme='pink').getvalue())
with open('/tmp/test_blue.pptx','wb') as f: f.write(build_ppt(data, charts, theme='blue').getvalue())
with open('/tmp/test_pink.pdf','wb') as f: f.write(build_pdf(data, charts, theme='pink').getvalue())
print('Done')
"
```

No tests, no linter, no build step. Python 3.13 via miniconda at `/Users/ylyt_bot/miniconda3/bin/python3`.

## Architecture

Data flows: **Excel → analyzer.py → {charts, budget, ppt/pdf builder} → HTTP response or file**

### Core Modules

| File | Role |
|------|------|
| `app.py` | Flask server — routes for upload/analyze, chart images, PPT/PDF generation, budget CRUD, rules, simulation, share. Caches analysis in `_analysis_cache` (max 20 entries) |
| `analyzer.py` | Parses iCost Excel (openpyxl, sheet `收支账单`). Column mapping: `row[0]=date, row[1]=type(收入/支出), row[2]=amount, row[3]=cat1, row[4]=cat2, row[5]=account, row[7]=note, row[9]=tag`. Expense amounts are negative in source, made absolute in output. Returns ~20-field dict |
| `charts.py` | matplotlib xkcd-mode charts → PNG BytesIO. Font auto-detection for Chinese. All chart colors unified to theme palette (no more earthy tones) |
| `ppt_builder.py` | python-pptx, up to 10 slides. Theme-aware via `set_ppt_theme()`. Key helpers: `_add_image()` (auto aspect-ratio with max_w/max_h), `_ensure_text_on_top()` (Z-order fix), `_add_multiline()` (auto-font-fit) |
| `pdf_builder.py` | fpdf2 with bundled WenQuanYi font. Theme-aware via `set_pdf_theme()`. `_chart_rgb()` flattens RGBA→RGB to prevent mobile PDF rendering artifacts |
| `budget.py` | 6 concerns: health scoring, budget CRUD, rule engine, what-if simulation, share snapshots, monthly data persistence. All storage via JSON files |

### Key Data Flow

`analyze()` returns a dict consumed by all downstream consumers. Fields include: `month`, `total_income`, `total_expense`, `balance`, `savings_rate`, `cat1_list`, `daily_list`, `weekly_list`, `food_list`, `account_list`, `income_list`, `health`, `budget_comparison`.

After `analyze()`, `app.py` attaches `health_score()` results and `compare_budget()` results to the dict before passing to chart/PPT/PDF builders.

### Storage

All runtime data stored as JSON files in project root or subdirs — **no database**:
- `budget_config.json` — budget categories and limits
- `auto_rules.json` — tag rules (pattern → tag)
- `shares/<uuid>.json` — share snapshots
- `data/<YYYY-MM>.json` — persisted monthly data for comparison

## Dual Theme System

The app supports two themes: **Sakura Pink** (default) and **Ocean Blue**. All 5 surfaces (web, mini program, PPT, PDF, charts) share the same theme tokens.

### Theme Switching

```python
# Charts
from charts import set_theme
set_theme('blue')        # switches global palette for all charts

# PPT
from ppt_builder import set_ppt_theme, build_ppt
set_ppt_theme('blue')    # or: build_ppt(data, charts, theme='blue')

# PDF
from pdf_builder import set_pdf_theme, build_pdf
set_pdf_theme('blue')    # or: build_pdf(data, charts, theme='blue')
```

### Web / Mini Program

- Web: `<html data-theme="blue">` or click the theme dots in the header. Saved to `localStorage('spendlens-theme')`.
- Mini Program: add `class="theme-blue"` to the page element. See `miniprogram/app.wxss` for CSS variable overrides.

### Color Palettes

| Token | Pink | Blue |
|-------|------|------|
| Primary | `#FF6B8A` | `#5B8DEF` |
| Dark BG | `#5C2D3E` | `#1E2A4A` |
| Text | `#3D1E2A` | `#2A3550` |
| Muted | `#C4909E` | `#90A4C4` |
| Positive | `#7BC8A4` | `#5BC8A4` |

Full design tokens in `DESIGN_SYSTEM.md`.

## PPT Builder Details

### Helper Functions

- `_add_image(slide, buf, x, y, w=None, h=None, max_w=None, max_h=None)` — adds image preserving aspect ratio. Specify either `w` or `h`, optional `max_w`/`max_h` clamp. Uses PIL to read native dimensions.
- `_add_multiline(slide, lines, x, y, w, h, size, auto_fit=True)` — multi-line text with auto font-sizing when content exceeds box height.
- `_ensure_text_on_top(slide)` — reorders XML spTree to push text shapes above charts/shapes. Called on all slides at end of `build_ppt()`.
- `_add_status_dot(slide, x, y, size, color)` — solid colored circle (replaces low-contrast emoji indicators).
- `_hex_to_rgb(hex_str)` — `'#7BC8A4'` → `RGBColor`
- `_PT = 12700` — EMU-to-point conversion constant, used in dynamic font sizing.

### PIL Import

PIL is imported at module level with fallback:
```python
try:
    from PIL import Image as PILImage
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False
```

### Color Access

`C` is a module-level dict built from hex strings. `set_ppt_theme('blue')` replaces it. All functions reference `C` directly (not imported) so they always see the current theme.

## PDF Builder Details

- `_chart_rgb(buf)` — converts chart PNGs from RGBA to RGB (white background) to fix mobile PDF viewer rendering artifacts (opaque stickers blocking content). Returns the original buffer unchanged if already RGB.
- `_rgb_tuple(val)` — converts hex string or tuple to `(R, G, B)` int tuple for fpdf2.
- `set_pdf_theme('blue')` replaces the global `C` dict with blue theme colors.

## Chart Color Rules

All chart functions use `_colors(n)` which returns the current theme's palette (pink `CC` or blue `CC_BLUE`). The old earthy palette (`#E8815F`, `#5B9A8B`, `#D4A853`, `#B0A090`) has been removed. Chart colors are never hardcoded — they always flow from the active theme.

Heatmap max color is `#5C2D3E` (pink dark) / `#1E2A4A` (blue dark), not `#FF3D6A`.

## Font Handling

- **matplotlib/charts**: Auto-scans system font dirs for Chinese-capable fonts. Uses WenQuanYi Micro Hei from `fonts/` directory.
- **fpdf2 PDF**: Uses WenQuanYi Micro Hei bundled in `fonts/`, registered as family `'CN'`. TTC fonts don't work with fpdf2.
- **PPT**: Arial throughout (system default, wide Chinese support).
- **Web**: Inter + JetBrains Mono from Google Fonts, system fallback.
- **Mini Program**: System font stack only (no custom fonts).

## Share Feature

`POST /share` accepts `{"data": <analyze result dict>}`, saves as JSON snapshot with a 10-char UUID, returns `share_url`. `GET /share/<id>` renders `share.html` (Jinja2 template, read-only view). No auth — URL is the access token.
