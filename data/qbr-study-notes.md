# CoWork QBR Reference Study Notes

Study completed for bead pb-ba2. These notes serve as a distilled reference for downstream implementation tasks (pb-0hq, pb-cwv, pb-m46).

## Architecture Summary

Three-layer pipeline: **Template Analyzer** -> **Data Processor** -> **Presentation Generator**

## Data Ingestion (for pb-0hq)

### Primary Data Source
- **Internal Performance Tracker** (Excel .xlsx)
- Location pattern: `No7/FY26/Int. Performance Tracker/[Month]/[Month] No7 Internal Tracker.xlsx`
- Key sheets: `RAW DATA`, `MTD Reporting`, `Daily`, `Channel Deep Dive`
- Channels: EMAIL, AFFILIATE, PPC, ORGANIC, DIRECT, SOCIAL, DISPLAY
- Metrics: Revenue, Orders, Sessions, New Customers, Cost (per channel, per day)

### Supplementary Data Sources
| File Type | Format | Contents |
|-----------|--------|----------|
| CRM data | .xlsm | Email/CRM campaign performance |
| Affiliate data | .xlsm | Partner-level breakdown |
| Offer Performance | .csv | Promotion redemption data |
| Product Sales | .csv | SKU-level product sales |

### Calculated Fields
| Metric | Formula |
|--------|---------|
| Variance vs Target | `(Actual - Target) / Target * 100` |
| YoY Variance | `(Actual - LY) / LY * 100` |
| QoQ Variance | `(Q_Current - Q_Previous) / Q_Previous * 100` |
| ROAS | `Revenue / Spend` |
| COS | `Spend / Revenue` |

### Missing Data Policy
- Missing target: show actual only, note "Target N/A"
- Missing LY: show QoQ instead
- Missing channel data: exclude that channel's slide
- Missing month: note gap, never interpolate

## Template Schema (for pb-cwv)

### Report Structure (14 slides)
| Slide | Type | Data Source |
|-------|------|-------------|
| 1 | Cover + KPIs | Tracker: MTD totals |
| 2 | Table of Contents | Static |
| 3 | Section Divider | Static |
| 4 | Executive Summary | Tracker: Channel totals + YoY |
| 5 | Daily Performance | Tracker: Daily sheet |
| 6 | Promotion Performance | Offer Performance CSV |
| 7 | Product Performance | Product Sales CSV |
| 8 | Section Divider | Static |
| 9 | CRM Performance | CRM file or Tracker: EMAIL |
| 10 | Affiliate Performance | Affiliate file or Tracker: AFFILIATE |
| 11 | SEO Performance | Tracker: ORGANIC channel |
| 12 | Section Divider | Static |
| 13 | Upcoming Promotions | Manual input |
| 14 | Next Steps | Manual input |

### Slide Types and Layout Patterns
- **KPI donuts**: Revenue, AOV, CVR, COS, NC Sign Ups — with variance indicators
- **Channel deep-dives**: Two-column layout (Q[X] Recap | Q[X+1] Strategy)
- **Executive Summary**: Three theme boxes + bullet points + Looking Ahead sidebar
- **Section dividers**: Full blue background, white centered title
- **Performance slides**: Left = chart/graph, Right = KPI donuts

### Design System
- Colors: Primary Blue #0066CC, Dark Blue #003366, Positive #00AA00, Negative #CC0000
- Typography: Titles 36-44pt bold, Headers 24-28pt bold, Body 14-16pt, KPI numbers 48-60pt bold

### Formatting Rules
- Currency: <$1k = `$XXX`, $1k-$999k = `$XXXk`, $1m+ = `$X.Xm`
- Percentages: Rates `X.X%`, Variances `+X.X%`/`-X.X%`, Point changes `+X.X ppts`
- Numbers: <1k = `XXX`, 1k-999k = `X,XXX`/`XXXk`, 1m+ = `X.Xm`

## Mapping Config (for pb-m46)

### Data-to-Slide Mapping
Each slide type needs specific data fields mapped to template placeholders:

**Cover slide**: Total Revenue, Total Orders, AOV, CVR (from MTD totals)
**Executive Summary**: Top 3 themes (derived from channel performance), YoY headline, Looking Ahead items
**Performance Overview**: Revenue vs Target by month (chart data), KPI actuals + targets + variances
**Channel Deep-Dives**: Per-channel GMV, CVR, ROAS/COS, AOV, Spend — each with target & YoY variance
**Daily Performance**: Day-by-day revenue/orders (from Daily sheet)
**Promotion Performance**: Offer-level redemption data (from CSV)
**Product Performance**: SKU-level sales (from CSV)

### Core Principles
1. Data-first: extract and analyse BEFORE writing narrative
2. No recycling: every output reflects its own unique data
3. Honest gaps: missing data = "N/A", never fabricated
4. Template-driven: formatting from template, not hardcoded
5. Similarity target: <80% vs previous month's output

## Tech Stack
- python-pptx >= 0.6.23 (PPTX manipulation)
- pandas >= 2.1.0 (data processing)
- openpyxl >= 3.1.0 (Excel reading)
- markitdown >= 0.1.0 (PPTX-to-text QA)
- Pillow >= 10.0.0 (thumbnail QA)
- pyyaml >= 6.0 (schema/config)

## Key Lessons from ATTEMPT1/ATTEMPT2
- Find-and-replace from previous month results in 93%+ similarity (FAIL)
- Must verify data exists in tracker before writing about it
- Each month has a different story (Dec 2025 = holiday peak +13% YoY; Jan 2026 = post-holiday reset -20% YoY)
- Tables must reflect actual data or be marked N/A
