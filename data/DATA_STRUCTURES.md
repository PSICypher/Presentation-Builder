# Example Data Files - Structure Analysis

Analysis of all data sources found in the prior art project (`/Users/merwes/CLAUDE/CoWork/QBRs/`). This document defines the schemas the `processor` module must handle.

---

## Data Source Categories

| Category | Format | Encoding | Primary Use | Slide Mapping |
|----------|--------|----------|-------------|---------------|
| Internal Performance Tracker | `.xlsx` | Standard | Revenue, orders, sessions, costs by channel | Slides 1, 4, 5, 9-11 |
| Product Sales Report | `.csv` | UTF-16 LE, tab-delimited | SKU-level product performance | Slide 7 |
| Offer Performance | `.csv` | UTF-16 LE, tab-delimited | Promotion/offer redemption tracking | Slide 6 |
| CRM/Email Performance | `.xlsm` | Standard | Email campaign metrics | Slide 9 |
| Affiliate Publisher | `.xlsm` | Standard | Affiliate partner breakdown | Slide 10 |
| Target Phasing | `.csv` | UTF-8, comma-delimited | Daily channel targets | Target variance calculations |
| Trading/Pricing | `.csv` | UTF-8, comma-delimited | Promotional pricing by SKU | Supplementary context |

---

## 1. Internal Performance Tracker (PRIMARY)

**File pattern:** `{Month} {Year} No7 Internal Tracker.xlsx`
**Location:** `No7/FY26/Int. Performance Tracker/{NN}.{Month} {Year}/`
**Size:** ~2.5MB per month

### Sheet: RAW DATA (25,775 rows in January 2026)

The core data source. Daily channel-level performance from Sept 2023 onward.

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `COS Year` | int | Year | `2026` |
| `COS Month` | int | Month number (1-12) | `1` |
| `COS Day` | int | Day of month | `15` |
| `COS Channel` | str | Marketing channel | `AFFILIATE` |
| `COS Locale` | str | Locale code | `en_US` |
| `COS Orders` | int | Order count | `42` |
| `COS New Customers` | int | New customer count | `21` |
| `COS COS%` | float | Cost of sale as decimal | `0.09` |
| `COS CAC` | float | Customer acquisition cost ($) | `10.2` |
| `COS CPA` | float | Cost per acquisition ($) | `5.1` |
| `COS Cost` | float | Total cost ($) | `214` |
| `COS Revenue` | float | Revenue ($) | `2383` |
| `COS Sessions` | int | Session count | `451` |
| `COS AOV` | float | Average order value ($) | `56.7` |
| `COS Conversion` | float | Conversion rate as decimal | `0.093` |

**Channels:** `AFFILIATE`, `DIRECT`, `DISPLAY`, `EMAIL`, `INFLUENCER`, `ORGANIC`, `OTHER`, `PPC`, `SOCIAL`

**Ingestion notes:**
- All numeric values are stored as numbers (no comma formatting)
- COS%, Conversion are decimals not percentages (0.09 = 9%)
- Data spans multiple years (2023-2026) for YoY comparison
- Filter to target month + comparison period for processing

### Sheet: MTD Reporting (54 rows)

Month-to-date summary with targets. Pre-calculated aggregations.

- Contains actual vs target comparisons
- Run-rate projections
- Layout is formatted for human viewing (non-tabular header region)

### Sheet: Daily (70 rows)

Day-by-day breakdown with targets.

| Column Group | Columns |
|-------------|---------|
| Affiliate spend | Date, AFFILIATE, Spend Target, % |
| Revenue | Date, Revenue Actuals, Revenue Target, % |
| New Customers | Date, New Customer Actuals, New customer Target, % |

**Ingestion notes:**
- Dates are datetime objects (`2026-01-01 00:00:00`)
- Percentage columns are decimal variance (`-0.129` = -12.9%)
- Multi-column-group layout (columns repeat for different metrics)

### Sheet: Channel Deep Dive (15 rows)

Aggregated channel comparison (YoY variance).

| Column | Type | Description |
|--------|------|-------------|
| `Split by 3` | str | Channel name or "Total" |
| `Orders` | float | YoY variance as decimal |
| `New_Customers` | float | YoY variance as decimal |
| `COS%` | float | YoY variance as decimal |
| `Cost - Currency` | float | YoY variance as decimal |
| `Revenue - Currency` | float | YoY variance as decimal |
| `Sessions - Filter` | float | YoY variance as decimal |
| `AOV` | float | YoY variance as decimal |
| `Conversion%` | float | YoY variance as decimal |

**Ingestion note:** Values are already YoY variances (e.g., `0.070` = +7.0% YoY), not absolute values.

### Sheet: Data Dump (117 rows)

Raw MTD data paste area with channel breakdowns.

| Column | Description |
|--------|-------------|
| `Split by 1` | Primary dimension (e.g., year) |
| `Split by 2` | Secondary dimension (e.g., month) |
| `Split by 3` | Tertiary dimension (e.g., channel) |
| `Orders` | Order count |
| `New_Customers` | New customer count |
| `COS%` | Cost of sale |
| `CAC` | Customer acquisition cost |
| `CPA` | Cost per acquisition |
| `Cost - Currency` | Total cost |
| `Revenue - Currency` | Total revenue |
| `Sessions - Filter` | Session count |
| `Impressions` | Ad impressions |
| `Clicks` | Ad clicks |
| `AOV` | Average order value |
| `Conversion%` | Conversion rate |

### Sheet: Manual update- Daily Target (1,128 rows)

Daily revenue targets spanning Aug 2024 onward.

| Column | Type | Description |
|--------|------|-------------|
| `Target Date` | datetime | Target date |
| `Target Locale` | str | Locale (often blank) |
| `Targets Week of the Year` | int | ISO week number |
| `Targets Year` | int | Year |
| `Targets Month` | int | Month number |
| `Targets Day` | int | Day of month |
| `Targets` | float | Revenue target ($) |
| `SPEND TARGET` | float | Spend target ($) |
| `ORDER TARGET` | float | Order target |
| `TRAFFIC` | float | Traffic/session target |

### Sheet: Manual update- Channel Targets (45 rows)

Monthly channel-level targets for revenue, sessions, orders, AOV, CVR.

**Layout:** Non-standard. Row-based channel labels with column groups for metrics. Requires positional parsing.

### Other Sheets

| Sheet | Rows | Purpose | Processing Priority |
|-------|------|---------|-------------------|
| `YTD` | 87 | Year-to-date monthly summary with YoY | Medium |
| `Last Week` | 13 | Weekly comparison | Low |
| `This week` | 8 | Current week performance | Low |
| `Warehouse Forecast` | 335 | Shipment/fulfillment data | Low |
| `N` | 38 | Channel revenue by month matrix | Medium |
| `Sheet1` | 11 | H1 revenue comparison FY23 vs FY24 | Low |
| `Sheet2` | 31 | Unlabelled numeric data | Skip |
| `IGNORE` | 19 | Flagged to ignore | Skip |
| `YTD Targets` | 10 | YTD target aggregation | Medium |

---

## 2. Product Sales Report CSV

**File pattern:** `Product Sales Report {Brand} {Region} {Quarter}.csv`
**Encoding:** UTF-16 LE with BOM, tab-delimited, CRLF line endings
**Size:** ~4.2MB, 8,236 data rows, 69 columns

### Schema

**Dimension columns (product/time hierarchy):**

| Column | Description | Example Values |
|--------|-------------|----------------|
| `Dimension 1` | Product name or "Grand Total" | `No7 Advanced Retinol 30ml` |
| `Dimension 2` | Month number or "Total" | `10`, `11`, `12`, `Total` |
| `Dimension 3` | Week/sub-period or "Total" | `1` through `5`, `Total` |

**Metric columns (22 metrics x 3 variants = 66 columns):**

Each metric has three columns following the pattern:
- `{Metric} (Analysis)` — Current period value
- `{Metric} (Comparison)` — Prior period value
- `{Metric} (vs. Comp)` — Variance as percentage string

| Metric | Unit | Example Analysis | Example vs. Comp |
|--------|------|-----------------|------------------|
| Units | integer (comma-formatted string) | `63,571` | `170.6%` |
| Orders | integer (comma-formatted string) | `49,156.000000000` | `160.8%` |
| AOV | float | `23.53` | `20.7%` |
| Product Revenue | integer (comma-formatted string) | `1,138,771` | `215.2%` |
| Shipping Revenue | integer (comma-formatted string) | `17,753` | `198.4%` |
| Total Revenue | integer (comma-formatted string) | `1,156,524` | `214.9%` |
| Avg. Items Per Order | float | `1.29` | `3.8%` |
| Avg. Selling Price | float | `17.91` | `16.5%` |
| Unique Products Sold | integer | `206` | `13.8%` |
| New Customers | integer (comma-formatted string) | `17,296` | `134.7%` |
| Returning Customers | integer (comma-formatted string) | `31,763` | `180.5%` |
| Total Customers | integer (comma-formatted string) | `49,059.000000000` | `162.4%` |
| Markdown | integer (comma-formatted string) | `298,341` | `131.2%` |
| Markdown % | percentage string | `17.9%` | `-3.52 ppts` |
| Promo Discount | integer (comma-formatted string) | `225,902` | `103.2%` |
| Promo Discount % | percentage string | `16.6%` | `-6.98 ppts` |
| Total Discount | integer (comma-formatted string) | `524,243` | `118.2%` |
| Total Discount % | percentage string | `31.5%` | `-8.41 ppts` |
| Discounted Units | integer (comma-formatted string) | `56,082` | `155.0%` |
| Discounted Units % | percentage string | `88.2%` | `-5.37 ppts` |
| Discounted Product Revenue | integer (comma-formatted string) | `926,018` | `191.6%` |
| Discounted Product Revenue % | percentage string | `80.1%` | `-6.40 ppts` |

**Ingestion notes:**
- Column names have trailing spaces: `'Dimension 1 '`
- Most numeric values arrive as comma-formatted strings (`"63,571"`) — must strip commas before parsing
- Some integers have float artifacts: `"49,156.000000000"` — truncate to int
- Percentage columns contain `%` suffix or `ppts` suffix — parse accordingly
- Variance columns mix `%` (relative change) and `ppts` (absolute point change)
- Many cells are `NaN` for product-level rows missing comparison data
- 153 unique products, 7 time periods (months 10-12 + Total), up to 61 sub-periods

---

## 3. Offer Performance CSV

**File pattern:** `Offer Performance {Brand} {Region} {Quarter}.csv`
**Encoding:** UTF-16 LE with BOM, tab-delimited, CRLF line endings
**Size:** ~1.1MB, 7,122 data rows, 10 columns

### Schema

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `Dimension 1` | str | Offer/promotion name or "Grand Total" | `[Affiliate - Cartera] Black Friday...` |
| `Dimension 2` | str | Channel | `AFFILIATE`, `EMAIL`, `PPC`, `DIRECT`, `ORGANIC`, `SOCIAL`, `INTERNAL REFERRAL`, `OTHER`, `REFERRAL`, `Total` |
| `Dimension 3` | str | Month number or "Total" | `1`, `10`, `11`, `12`, `Total` |
| `Dimension 4` | str | Week/sub-period or "Total" | `1` through `31`, `Total` |
| `Redemptions` | str | Redemption count (comma-formatted) | `26,338` |
| `% Change Redemptions` | str | YoY change with arrow | `-21.6% ↓` |
| `Revenue` | str | Revenue (comma-formatted) | `2,173,984` |
| `% Change Revenue` | str | YoY change | `-21.9%` |
| `Discount Amount` | str | Discount value (comma-formatted) | `248,421` |
| `% Change Discount Amount` | str | YoY change with arrow | `-24.8% ↓` |

**Ingestion notes:**
- Column names have trailing spaces: `'Redemptions '`, `'Revenue '`, `'Discount Amount  '`
- Some percentage change values include Unicode arrows (↓ ↑) — strip before parsing
- 247 unique offer names — some contain special characters, pipes (`||`), dates
- Channels in Dimension 2 match the tracker channel names
- 13 unique channel values including `INTERNAL REFERRAL`, `REFERRAL`, `OTHER`

---

## 4. CRM/Email Performance (Excel)

**File pattern:** `CRM {Brand} {Region} {Quarter}.xlsm`
**Sheet:** `Custom Dates Dimension Table`
**Size:** 12KB, 8 data rows, 31 columns

### Schema

**Dimension columns (first 3 columns):**
- Column A: Grand Total / Total
- Column B: Total
- Column C: Campaign type (`Manual`, `Automated`, `Total`)

**Metric columns (14 metrics, each with an `vs Comp` companion):**

| Metric | vs Comp Type | Description |
|--------|-------------|-------------|
| `Emails Sent` | % change | Total emails sent |
| `Emails Delivered` | % change | Successfully delivered |
| `Emails Opened` | % change | Unique opens |
| `Open Rate` | ppts change | Open rate (decimal) |
| `Unsubscribes` | % change | Unsubscribe count |
| `Unsibscribe Rate` | ppts change | Unsubscribe rate (decimal) [sic: typo in source] |
| `Link Clicks` | % change | Click count |
| `Click-Through Rate` | ppts change | CTR (decimal) |
| `Sessions` | % change | Sessions driven |
| `Landing Page Bounce Rate` | ppts change | Bounce rate (decimal) |
| `Orders` | % change | Orders driven |
| `CVR` | ppts change | Conversion rate (decimal) |
| `Revenue` | % change | Revenue ($) |
| `AOV` | % change | Average order value ($) |

**Ingestion notes:**
- Column headers have inconsistent spacing (leading/trailing spaces)
- Typo in source: `Unsibscribe Rate` — handle both spellings
- Very small dataset (8 rows) — likely aggregated by campaign type
- `vs Comp` columns contain decimals (e.g., `-0.106` = -10.6%)
- Rate columns (Open Rate, CVR, etc.) are decimals not percentages

---

## 5. Affiliate Publisher Performance (Excel)

**File pattern:** `Affiliate Publisher {Quarter} {Brand} {Region}.xlsm`
**Sheet:** `Table - Custom Dates`
**Size:** 68KB, 194 data rows, 88 columns

### Schema

**Dimension columns:**

| Column | Description | Example |
|--------|-------------|---------|
| `Dimension 1` | Publisher ID or "Grand Total" | `1004849`, `101248` |
| `Dimension 2` | Publisher name | `REWARDOO`, `TAKEADS GMBH` |
| `Dimension 3` | Sub-dimension or "Total" | `Total`, month numbers |
| `Influencer Filter` | Publisher type | `Affiliate`, `Influencer`, `Total` |

**Metric groups (all with Analysis/Comparison/vs Comp triplets):**

| Group | Metrics |
|-------|---------|
| Revenue | Revenue, Revenue Delta, Revenue Mix |
| Engagement | RPS (Revenue Per Session) |
| Orders | Orders, Items per Order |
| Value | AOV |
| Discounting | Discount %, Discount $ |
| Profitability | Gross Profit %, Gross Profit $ |
| Cost | Cost, Total Commission, Total Tenancy |
| Efficiency | CoS, CAC, CPA, CPS |
| Traffic | Sessions, Bounces, Bounce Rate, CVR |
| Customers | NC Mix, New Customers, Returning Customers |
| Publisher Activity | Session-Active Publishers (CountD), Sale-Active Publishers (CountD) |

**Ingestion notes:**
- 88 columns — largest schema of all files
- Column headers have inconsistent spacing
- Typo in source: `Sale-Actvie Publishers` — handle both spellings
- Many empty cells for inactive publishers
- `vs Comp` columns alternate between `(%)` and `(ppts)` notation
- Revenue Mix columns have very long names containing dimension descriptions

---

## 6. Target Phasing CSV

**File pattern:** `{Month}-Targets.csv` or `Sept.csv`
**Encoding:** UTF-8, comma-delimited
**Size:** ~120KB, 420-434 data rows

### Schema

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `Target_Type_Id` | str | Granularity | `Daily` |
| `Date` | str | ISO date | `2025-09-01` |
| `Site_Id` | str | Brand site | `No 7` |
| `Locale_Id` | str | Locale | `en_US` |
| `Channel_Id` | str | Channel | `AFFILIATE`, `DIRECT`, `DISPLAY`, `EMAIL`, `INFLUENCER`, `ORGANIC`, `OTHER`, `PPC`, `SOCIAL` |
| `Notes` | float | Notes (usually NaN) | `NaN` |
| `Gross_Revenue_Target` | float | Gross revenue target ($) | `519.596052` |
| `Net_Revenue_Target` | float | Net revenue target ($) | `552.761532` |
| `Marketing_Spend_Target` | float | Marketing spend target ($) | `215.0358` |
| `Session_Target` | int | Session target | `121` |
| `Order_Target` | int | Order target | `9` |
| `New_Customer_Target` | int | New customer target | `13` |

**Ingestion notes:**
- Clean, well-structured data — easiest to ingest
- Daily granularity per channel per date
- Revenue targets are precise floats (not rounded)
- Channel names match RAW DATA sheet channels exactly
- One row per channel per day (~14 channels x 30 days = ~420 rows)

---

## 7. Trading/Pricing CSV

**File pattern:** `{promotion description}.csv`
**Encoding:** UTF-8, comma-delimited
**Size:** 5-20KB, 37-138 data rows

### Schema

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `Master SKU` | float | Parent SKU ID | `12569556.0` |
| `Master Title` | str | Parent product name | `No7 Laboratories Line Correcting Booster Serum (Various Sizes)` |
| `SKU` | int | Variant SKU ID | `11994411` |
| `Title` | str | Variant product name with attributes | `...Serum (Various Sizes) - Size:15ml` |
| `Subsite` | str | Subsite/locale | `en_US` |
| `Currency` | str | Currency code | `USD` |
| `Country Group` | str | Geographic scope | `Worldwide` |
| `RRP` | float | Recommended retail price | `42.99` |
| `Price` | float | Promotional price | `34.392` |
| `Expiry Date` | str | Promotion end date (often empty) | `NaN` |
| `Campaign Details` | str | Campaign notes (often empty) | `NaN` |

**Ingestion notes:**
- Master SKU is NaN for variant rows under the same parent
- Price precision varies (some are exact percentage discounts: `34.392 = 42.99 * 0.8`)
- Useful for promotion performance cross-referencing but not primary report data

---

## 8. Historical Comparison Data (CSV)

**File pattern:** `{Month}{Year}-Data.csv`
**Encoding:** UTF-16 LE, tab-delimited
**Size:** ~10KB, 33 data rows

### Schema

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| `Dimension 1` | str | Year or "Grand Total" | `2024` |
| `Dimension 2` | str | Month number | `10` |
| `Dimension 3` | str | Day or "Total" | `1`, `Total` |
| `Orders` | str | Order count (comma-formatted) | `4,659` |
| `New Customers` | str | New customer count (comma-formatted) | `1,725` |
| `CoS` | str | Cost of sale (percentage string) | `96.90%` |
| `CAC` | float | Customer acquisition cost | `142.8` |
| `CPA` | float | Cost per acquisition | `52.9` |
| `Cost` | str | Total cost (comma-formatted) | `246,410` |
| `Phased` | str | Phased percentage | `3%` |
| `Revenue` | str | Revenue (comma-formatted) | `254,388` |
| `Phased.1` | str | Revenue phased percentage | `3%` |

**Ingestion notes:**
- Prior-year daily data for YoY calculation
- `Phased` columns show % of month's total allocated to that day
- Same dimension pattern as other THG exports

---

## Common Ingestion Challenges

### 1. Encoding Detection
- THG platform exports use UTF-16 LE with BOM + tab delimiter
- Internal/target files use standard UTF-8 + comma delimiter
- Processor must auto-detect or accept encoding parameter

### 2. Numeric String Parsing
Comma-formatted strings need cleaning before numeric conversion:
```
"63,571" -> 63571
"49,156.000000000" -> 49156
"1,138,771" -> 1138771
```

### 3. Percentage/Variance Parsing
Multiple formats in the same dataset:
```
"170.6%"     -> 1.706 (relative change)
"-3.52 ppts" -> -0.0352 (absolute point change)
"-21.6% ↓"   -> -0.216 (with Unicode arrow)
"96.90%"     -> 0.969 (absolute rate)
```
Context determines interpretation: `vs. Comp` columns are relative changes; `COS%`, `CVR` columns are absolute rates.

### 4. Column Name Inconsistency
- Trailing spaces: `'Dimension 1 '` vs `'Dimension 1'`
- Leading spaces: `'   vs Comp  '`
- Typos: `'Unsibscribe Rate'`, `'Sale-Actvie Publishers'`
- Recommendation: `.strip()` all column names on ingestion

### 5. Mixed Types in Columns
Pandas reads many numeric columns as `object`/`str` due to:
- Comma formatting in numbers
- `%` and `ppts` suffixes
- `NaN` mixed with formatted strings
- Floating point artifacts in integers

### 6. Dimension Hierarchy
THG exports use generic `Dimension 1/2/3/4` naming. Actual meaning varies by file:
- Product Sales: Dimension 1 = Product, Dimension 2 = Month, Dimension 3 = Week
- Offer Performance: Dimension 1 = Offer, Dimension 2 = Channel, Dimension 3 = Month, Dimension 4 = Week
- Affiliate: Dimension 1 = Publisher ID, Dimension 2 = Publisher Name, Dimension 3 = Sub-period
- Historical: Dimension 1 = Year, Dimension 2 = Month, Dimension 3 = Day

---

## Channel Name Mapping

Consistent across all data sources:

| Channel ID | Full Name | Appears In |
|-----------|-----------|------------|
| `AFFILIATE` | Affiliate Marketing | Tracker, Targets, Offer Perf |
| `DIRECT` | Direct Traffic | Tracker, Targets, Offer Perf |
| `DISPLAY` | Display Advertising | Tracker, Targets |
| `EMAIL` | Email/CRM | Tracker, Targets, Offer Perf |
| `INFLUENCER` | Influencer Marketing | Tracker, Targets |
| `ORGANIC` | Organic/SEO | Tracker, Targets, Offer Perf |
| `OTHER` | Other/Unattributed | Tracker, Targets, Offer Perf |
| `PPC` | Pay-Per-Click/Paid Search | Tracker, Targets, Offer Perf |
| `SOCIAL` | Paid Social | Tracker, Targets, Offer Perf |
| `INTERNAL REFERRAL` | Internal Referral | Offer Perf only |
| `REFERRAL` | External Referral | Offer Perf only |

---

## Slide-to-Data Mapping

| Slide | Data Source | Key Fields to Extract |
|-------|-----------|----------------------|
| 1 - Cover + KPIs | Tracker: MTD Reporting | Total Revenue, Orders, AOV, CVR, NC, vs Target |
| 4 - Executive Summary | Tracker: RAW DATA (aggregated) | Channel totals, YoY variance, narrative drivers |
| 5 - Daily Performance | Tracker: Daily sheet | Daily revenue, target, cumulative |
| 6 - Promotion Performance | Offer Performance CSV | Top offers by redemption/revenue, channel split |
| 7 - Product Performance | Product Sales Report CSV | Top products by revenue, units, discount impact |
| 9 - CRM Performance | CRM Excel or Tracker EMAIL | Emails sent/opened/clicked, CVR, revenue, vs comp |
| 10 - Affiliate Performance | Affiliate Excel or Tracker AFFILIATE | Top publishers, revenue, commission, CoS |
| 11 - SEO Performance | Tracker: ORGANIC channel | Sessions, revenue, CVR, YoY |

---

## Derived Metrics (Processor Must Calculate)

| Metric | Formula | Source Fields |
|--------|---------|--------------|
| YoY Variance | `(current - prior) / prior` | RAW DATA filtered by year |
| vs Target Variance | `(actual - target) / target` | RAW DATA + Targets |
| ROAS | `revenue / spend` | RAW DATA: Revenue, Cost |
| COS | `cost / revenue` | RAW DATA: Cost, Revenue |
| CVR | `orders / sessions` | RAW DATA: Orders, Sessions |
| CAC | `cost / new_customers` | RAW DATA: Cost, New Customers |
| CPA | `cost / orders` | RAW DATA: Cost, Orders |
| AOV | `revenue / orders` | RAW DATA: Revenue, Orders |
| NC Mix | `new_customers / total_customers` | RAW DATA |
| Run Rate | `(mtd_actual / days_elapsed) * days_in_month` | Daily sheet |

---

## Recommended Ingestion Priority

1. **RAW DATA sheet** — Foundation for all channel metrics and YoY
2. **Target Phasing CSV** — Required for vs-target variance on every slide
3. **Product Sales Report CSV** — Slide 7 (product performance)
4. **Offer Performance CSV** — Slide 6 (promotion performance)
5. **CRM Excel** — Slide 9 (email deep-dive)
6. **Affiliate Excel** — Slide 10 (affiliate deep-dive)
7. **Daily sheet** — Slide 5 (daily performance chart)
8. **Trading CSVs** — Supplementary context only
