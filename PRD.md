# Product Requirements Document — BeGifted Sales Dashboard

- **Owner:** aoengnatchasmith@gmail.com
- **Platform:** Google Apps Script (V8 runtime) + HTML Web App (served via `doGet`)
- **Timezone:** `Asia/Bangkok`
- **Deployment:** Web app, `executeAs: USER_DEPLOYING`, `access: ANYONE`
- **Apps Script Project ID:** `1PYgKjKIyR8XzESqBHxvID1fXQFKohLdtmu2wCce0ybytbm1BB0LEOpFd`
- **Source files:** `Code.gs` + `Dashboard.html` + `appsscript.json`

---

## 0. Purpose

Consolidate BeGifted's monthly sales data — spread across multiple Google Sheets (one file per month) — into a single normalized master dataset, classify each transaction by customer lifecycle, and expose it as a filterable sales dashboard. The goal is to give the sales team and management real-time visibility without anyone having to build monthly reports by hand.

The system consists of 4 main stages (Extract → Build → Analyze → Dashboard) plus automatic refresh 3× daily via time-based triggers.

---

## Step 1 — Extract

### 1.1 Source Files (FILES array — 14 entries, Apr 2025 – May 2026)

| # | Label | Google Sheet ID | Package Sheet |
|---|---|---|---|
| 1 | 2025-04 Apr | `161E5AKy7mNp7xloqF77hHOOy7OBs1YRGQFAA3UTn98s` | `SalesRecord` |
| 2 | 2025-05 May | `1WUY91SetwtXWrq3dLG6BJDliAgKO7ZMQFRQzh5twyTo` | `SalesRecord` |
| 3 | 2025-06 Jun | `1fP5JxN2of6Q_NyluXWXWZkbeh7lIMj8dLSleEzTSGA0` | `(1)PackageSales` |
| 4 | 2025-07 Jul | `1yd6QZQFamlRxnhYFGIDrfIZHuEzGtkdLpLmNDDtq1DE` | `(1)PackageSales` |
| 5 | 2025-08 Aug | `1wfPrvBF73L1AumyiSJ5TL919CeTt7rOwLe6LJ9HpMM4` | `(1)PackageSales` |
| 6 | 2025-09 Sep | `1y-IEH4E2sO_XGs5NlPcrlMEfkoTTeS1WVB_Oskn-01w` | `(1)PackageSales` |
| 7 | 2025-10 Oct | `1Ont-CPISkfunIS01oWo9dVVLfPL9RZspVJdeZ6YsMRo` | `(1)PackageSales` |
| 8 | 2025-11 Nov | `1mmujgodFgUi3lyVqwX7fJCmQNRFUncgQKj0l9Ys71Fw` | `(1)PackageSales` |
| 9 | 2025-12 Dec | `1W3V-bNBOJLtT0Lml5_OMF2tXf4PVcUj4X0GS0PhGC5c` | `(1)PackageSales` |
| 10 | 2026-01 Jan | `1z9LAQbZ-V2GYLm_NA5lkkhR8fdXqyiUzW9EuiHJyeJM` | `(1)PackageSales` |
| 11 | 2026-02 Feb | `1dRZjgRP3f0isr-ssZxobwhlsw1v8WWzR0v4zMR82o3k` | `(1)PackageSales` |
| 12 | 2026-03 Mar | `1G3wgBV9KnSyqNiSwHKULmbtgEbJnnLTCR-zDBqalS4w` | `(1)PackageSales` |
| 13 | 2026-04 Apr | `1HHtZ6YYCqK8wI6nYvVXpwgHSrqoFzcPOD7mMz8hQVJk` | `(1)PackageSales` |
| 14 | 2026-05 May | `1wrIEfBKFp325nFYfeKW7r7znfpT_cN6s4VMwxjEEfXA` | `(1)PackageSales` |

The `pkgSheet` field overrides the default `SRC_PACKAGE = '(1)PackageSales'` when set. Apr–May 2025 use `SalesRecord` because the tab was renamed in Jun 2025.

### 1.2 Source Sheet Format Detection

`extractNormalSales()` auto-detects format by checking column headers:

| | Old Format (Apr 2025 – Nov 2025, Jan–Mar 2026) | New Format (Dec 2025+) |
|---|---|---|
| Payment Date | `วันที่ชำระเงิน` | `Payment Date` |
| Sales Rep | `ผู้ขาย` | `Sales Person` |
| Amount | `ยอดชำระสุทธิ` | `Total Price` |
| Filter | Skip if payment date empty | Only rows where `Already Paid? = TRUE` |
| Pre-filled Enrollment Type | Not available | `Enrollment Type` column (Apr 2026+) |

### 1.3 Enrollment Type — Pre-filled vs. Inferred

Starting **Apr 2026**, source sheets contain a pre-filled `Enrollment Type` column. The extractor reads it directly and maps values:

| Source value | Mapped to |
|---|---|
| `Trial` | `Trial` |
| `New` | `New Student` |
| `Renew` / `Renewal` | `Renewal` |

For **all months before Apr 2026**, enrollment type is left blank in staging and classified in Step 3 from package name.

### 1.4 Staging Sheets

- Normal Sales → `NormalSales_<MM><YYYY>` (e.g. `NormalSales_042025`)
- Additional Sales → `AdditionalSales_<MM><YYYY>`
- Both have `Header row = 3` (`HEADER_ROW = 3`)
- A `Source Month` column is added (e.g. `"2025-09 Sep"`) for traceability

### 1.5 Row Filtering

A row is **skipped** if:
1. `Student's Nickname` is empty
2. Payment date is empty or invalid (old format)
3. `Already Paid?` is not TRUE (new format)

### 1.6 Entry point
- **Manual:** `runStep1_Extract()`

---

## Step 2 — Build

Combines all staging sheets into two master sheets (`runStep2_Build`).

### 2.1 `MasterNormalized_NormalSales`

| Staging | Master | Notes |
|---|---|---|
| `Student's Nickname` | `Student's Nickname` | unchanged |
| `Program` | `Program` | unchanged |
| `Package` | `Package Hours` | renamed |
| `No. of Student` | `No. of Student` | unchanged |
| amount column | `Payment Amount` | renamed |
| rep column | `Sales Representative` | renamed |
| date column | `Payment Date` | renamed |
| `Source Month` | `Source Month` | unchanged |
| `Enrollment Type` | `Enrollment Type` | pre-filled value passed through (blank for pre-Apr 2026) |
| *(new)* | `Program (Wise Name)` | populated in Step 3 |
| *(new)* | `Package Hours (Clean)` | populated in Step 3 |
| `Valid Until` | `Valid Until` | unchanged |

**Formatting:** header `#003087` white bold; alternating stripes `#F0F4FF`/`#FFFFFF`; dates `dd/mm/yyyy`; amounts `#,##0.00`; row 1 frozen.

### 2.2 `MasterNormalized_AdditionalSales`

Columns: `Student's Nickname`, `Sales Type`, `Package`, `Payment Amount`, `Payment Date`, `Source Month`.

### 2.3 Entry point
- **Manual:** `runStep2_Build()` — requires Step 1 first

---

## Step 3 — Analyze & Transform

### 3.1 Enrollment Type Classification

**For rows that already have a pre-filled Enrollment Type (Apr 2026+):** value is kept as-is, no re-classification.

**For rows without pre-filled type (Apr 2025 – Mar 2026):** classify from package name:

1. Group by `Student's Nickname` (lowercased + trimmed)
2. Sort by `Payment Date` oldest → newest
3. Apply rules:

| Condition | Enrollment Type |
|---|---|
| Package contains "Trial" (case-insensitive) | **Trial** |
| First non-Trial row AND student has a prior Trial | **New Student** |
| First non-Trial row AND no prior Trial | **Renewal** |
| Any subsequent non-Trial row | **Renewal** |

**Visual:** Trial `#DBEAFE`, New Student `#D1FAE5`, Renewal `#FEF3C7`.

### 3.2 Program (Wise Name)

Adds `Program (Wise Name)` via `PROGRAM_MAP` lookup. Full mapping table (unchanged from original PRD).

**⚠️ Known gap:** `16+ Master` not yet in the map.

### 3.3 Package Hours (Clean)

Strips parentheticals via regex `/\s*\(.*?\)/g` — e.g. `"30-hr (free extra 1 hr)"` → `"30-hr"`.

### 3.4 Churn Status

Populated **only on latest row per student** (all others show `—`).

| Status | Condition |
|---|---|
| **Active** | `Valid Until + 14 days` ≥ today |
| **Retained** | Past grace AND has a newer payment |
| **Churned** | Past grace AND no newer payment |
| **N/A** | Trial-only student OR no `Valid Until` |

Grace period: `GRACE_DAYS = 14`. Uses `Valid Until` from most recent non-Trial row.

**Visual:** Active `#D1FAE5`, Retained `#FEF3C7`, Churned `#FEE2E2`, N/A `#F3F4F6`.

### 3.5 Dashboard Cache (`buildDashboardCache`)

Runs immediately after Step 3. Stores data in hidden sheet `Dashboard_Cache`:

**Cache layout:**
- **Column A, rows 1..N:** `normalDays` JSON split into ≤40KB chunks. Count stored as `normalChunks` in B1.
- **B1:** aggregates JSON — `normalChunks`, `addDays`, `pkgCount`, `progCount`, `addPkgCount`, `repArr`, `dayCount`, `completionRate`, `weekBandPct`, `completionMonths`, totals, churn stats, `lastUpdated`
- **C1:** `churnList[]` — per-student array with `validUntil` + `status`

**`completionRate[1..31]`** — computed from all complete months: cumulative % of monthly revenue typically received by day D. Used for projection.

**`weekBandPct[0..5]`** — avg % of monthly revenue per 5-day band (1-5, 6-10, 11-15, 16-20, 21-25, 26-31). Used for Payment Concentration chart.

**`mergeCacheParts_()`** reads B1 for chunk count, loops A1..AN, merges with aggregates and churnList.

### 3.6 Entry point
- **Manual:** `runStep3_Analyze()`
- **Alert:** rows processed + Enrollment Type + Churn Status breakdown

### 3.7 Helper: Unmapped Programs

`debugUnmappedPrograms()` — highlights unmapped rows orange, creates `UnmappedPrograms_Log` sheet. **Run after every Step 3.**

---

## Step 4 — Dashboard (Web App)

### 4.1 Architecture

- `doGet()` — serves `Dashboard.html`
- `getDashboardData()` — reads cache via `mergeCacheParts_()`, returns full JSON
- `getSSId()` — reads Spreadsheet ID from `PropertiesService` (set on first editor run); avoids `getActiveSpreadsheet()` failing in web app context
- All filtering and chart rendering happen **client-side** — no backend call on period change
- `lastUpdated` shown in Bangkok timezone in top bar

### 4.2 Layout (Single Page)

```
Topbar:  BeGifted logo · Last updated: DD Mon, HH:MM (data updates 3 times a day)
Tabbar:  Sales Overview | Period buttons | Custom date range

Row 1:   [Normal Sales] [Additional Sales] [Total Revenue] [Trial→Paid] [Renewal Rate] [Churned·Replacement]
Row GT:  [🎯 Goal Tracker — ฿4M/mo  (collapsible)]
Row 2:   [Monthly Revenue + toggles]  [Pipeline Health]  [Sales Rep Ranking]
Row 3:   [Package Hours]  [Program (Wise Name) + toggle]  [Payment Concentration + toggle]
```

### 4.3 Period Selector

| Button | Period |
|---|---|
| **All** (default) | Apr 2025 → present |
| **2025** | Apr 2025 – Dec 2025 |
| **2026** | Jan 2026 → present |
| **Q1 2026** | Jan – Mar 2026 |
| **This Month** | Current month (dynamic — always current calendar month) |
| **Custom** | User-defined via date pickers |

Month label format: all months show year suffix e.g. `Apr '25`, `Jan '26`.

### 4.4 KPI Cards (6 — top row)

| # | Card | Formula | Color |
|---|---|---|---|
| 1 | **Normal Sales** | Sum of `rev` from `normalDays` in period | Blue |
| 2 | **Additional Sales** | Sum of `rev` from `addDays` in period | Orange |
| 3 | **Total Revenue** | Normal + Additional | Navy |
| 4 | **Trial → Paid Conversion** | `uniqueNewStudents / uniqueTrials × 100%` | Green |
| 5 | **Renewal Rate** | `Renewals / (New + Renewals) × 100%` — excludes Trial | Yellow |
| 6 | **Churned · Replacement** | Churned count; Replacement = `New / Churned` — filtered by `validUntil` within period | Red |

### 4.5 Goal Tracker (collapsible row between KPIs and charts)

A full-width collapsible bar between the KPI row and the chart rows.

**Collapsed:** shows summary — `🎯 Goal Tracker — ฿4M/mo · FY26 avg ฿X · Gap ฿Y` + pill `▼ ฿Y to go`.

**Expanded:** overlays chart area as an absolute-positioned panel containing:
- **Progress bar:** `฿X / ฿4M (N%)`
- **"To close the gap" section:** 3 cards (Renewal / New Student / Trial) showing FY26 avg baseline + target needed, calculated by scaling current mix proportionally to close the gap
- **"Or improve existing metrics" section:** 2 cards (↑ Renewal Rate / ↑ Avg Package Value)
- **Footer note:** method explanation

**Responsive logic:**
- **This Month filter** → uses actual current month revenue as baseline; shows FY26 avg for reference
- **All other periods** → uses FY26 monthly average as baseline

**Formula (gap distribution):**
```
blendedAvgRev = FY26 monthly avg revenue / avg monthly transactions
extraTxnNeeded = gap / blendedAvgRev
needRenewal    = extraTxnNeeded × (renewalTxns / totalTxns)
needNew        = extraTxnNeeded × (newTxns / totalTxns)
needTrial      = extraTxnNeeded × (trialTxns / totalTxns)
```

### 4.6 Monthly Revenue Chart — 5 Modes

Toggle buttons: **Normal Sales** | **Additional** | **Combined** | **YoY** | **MoM**

#### Normal Sales mode
- Stacked bar: Trial (`#60A5FA`) / New Student (`#34D399`) / Renewal (`#FBBF24`)
- Revenue per type stored as `revT / revN / revR`
- **FY2025 avg line** (blue dashed) — spans 2025 months only
- **FY2026 avg line** (green dashed) — spans 2026 months only
- **Goal ฿4M line** (red dashed) — full width
- FY avg lines show value on hover; Y-axis max set to `max(tallest bar × 1.3, ฿4M × 1.1)`
- Includes projection (see 4.6.1)

#### Additional mode
- Single orange bar — no projection, no avg lines

#### Combined mode
- Normal stacked (Trial/New/Renewal) + Additional on top — 4 layers; no projection

#### YoY mode
- Dropdown: select month (only months with a valid same-month prior-year entry shown)
- Side-by-side bars: selected month vs. same month prior year
- Breakdown: Trial / New Student / Renewal / Additional / Total
- Pill: `▲ +X% YoY · Apr '26: ฿X vs Apr '25: ฿Y`

#### MoM mode
- Dropdown: select month (only months with a valid prior month shown)
- Side-by-side bars: selected month vs. previous month
- Same breakdown as YoY
- Pill: `▼ -X% MoM · Apr '26: ฿X vs Mar '26: ฿Y`

### 4.6.1 Projection Logic (Normal mode — `calcProjection`)

Only shown when filter period extends to today or later (not shown for 2025, Q1 2026, etc.).

Uses `window._nbm_all` (full unfiltered data) so projection is unaffected by period filter.

**Option B — Hybrid Blend:**

```
THRESHOLD = 30%
actualPct = actual_this_month / fy26_hist_avg

if actualPct < 30%:
  wActual = 0%  → trust history 100%
else:
  wActual = min(80%, (actualPct - 30%) / 70%)
wHistory = 1 - wActual

actualProj   = actual / completionRate[todayDom]   (from cache)
historyProj  = weighted avg(last 3 complete months, weights 1:2:3)
projected    = actualProj × wActual + historyProj × wHistory
```

**Rendering:**
- Current month: actual bar + semi-transparent gap segment (`rgba(150,150,220,0.15)`, dashed border) showing estimated remaining
- Next month: projected bar at 30% opacity
- Tooltip footer: `Est: ฿X · X% actual · Y% hist`

### 4.7 Pipeline Health (`#cFunnel`)
Line chart — Trials / New Students / Renewals per month. Colors: blue / green / yellow.

### 4.8 Sales Rep Ranking (`#cRep`)
Table: rank, name, txn count, revenue breakdown pills (Trial/New/Renewal), stacked txn-ratio bar, total revenue. Sorted by Normal Sales revenue descending. Top 8 shown by default with expand button.

### 4.9 Package Hours (`#cPkg`)
Horizontal bar — transaction count per `Package Hours (Clean)`. Normal Sales only. Top 8 default, expandable. Bar thickness auto-scales.

### 4.10 Program (Wise Name) (`#cProg`) — toggle: Normal / Additional
- **Normal:** `Program (Wise Name)` from `progCount`, orange shades
- **Additional:** `Package` from `addPkgCount`, orange shades
- Top 8 default, expandable.

### 4.11 Payment Concentration (`#cDays`) — toggle: By Day of Week / By Week of Month

#### By Day of Week
- Bar chart — total transaction count per day (Mon–Sun)
- Hover shows: txn count + **% of revenue** for that day
- Peak day highlighted `#003087`; others `rgba(0,48,135,0.2)`

#### By Week of Month
- 6 bands: 1-5 / 6-10 / 11-15 / 16-20 / 21-25 / 26-31
- **Two bars per band:** selected period actual % (orange) vs. N-month historical avg (navy, semi-transparent)
- Responsive to period filter — actual % recomputes from filtered `normalDays`
- Hover shows: `Day X-Y · Selected period: X% · N-month avg: Y% · ▲/▼ diff vs avg`
- Historical avg (`weekBandPct`) pre-computed in Step 3 from all complete months

---

## 5. Automation (Daily Refresh — 3× per day)

### 5.1 `setupTriggers()` — one-time setup
Deletes existing triggers, then creates **3 time-based triggers** for `dailyRefresh()`:

| Time | Bangkok |
|---|---|
| 01:00 AM | overnight refresh |
| 12:00 PM | midday refresh |
| 06:00 PM | end-of-day refresh |

### 5.2 `dailyRefresh()` — automated pipeline
Runs Step 1 → Step 2 → `runStep3_Analyze_silent()` → rebuild cache. No UI alerts.

**⚠️ Limitation:** silent path does **not** recompute Churn Status. Values from last manual Step 3 remain.  
**Recommendation:** run `runStep3_Analyze()` manually at least once a week.

### 5.3 `removeTriggers()` — wipe all project triggers

---

## 6. Entry Points Summary

| Function | Purpose | Invoked from |
|---|---|---|
| `runStep1_Extract()` | Extract all staging sheets (14 months) | Manual |
| `runStep2_Build()` | Build 2 master sheets | Manual |
| `runStep3_Analyze()` | Classify + build cache + alerts | Manual (after Step 2) |
| `setupTriggers()` | Install 3× daily triggers | Manual (one-time) |
| `dailyRefresh()` | Automated pipeline | Time-based trigger |
| `removeTriggers()` | Clear all triggers | Manual |
| `debugUnmappedPrograms()` | Find programs missing from map | Manual (maintenance) |
| `doGet()` | Serve the web app | HTTP request |
| `getDashboardData()` | Return cached payload | Called from Dashboard.html |
| `getSSId()` | Get Spreadsheet ID from Properties | Internal |

---

## 7. Known Issues & Future Work

### 7.1 Data Quality

| Issue | Impact | Mitigation |
|---|---|---|
| Source `Program` not in `PROGRAM_MAP` | Fragments groupings | Run `debugUnmappedPrograms()` + extend map |
| `16+ Master` not yet mapped | Own bucket in charts | Add to `PROGRAM_MAP` |
| Row missing `Valid Until` | Churn Status = N/A | Enforce in source sheet |
| Duplicate nicknames (different students) | Wrong Enrollment/Churn | Use `Student ID` if feasible |
| Trial with Payment Amount > ฿1000 | Still classified as Trial | Add sanity-check in future |

### 7.2 System

| Issue | Impact | Mitigation |
|---|---|---|
| Daily trigger doesn't recompute churn | Churn Status stale | Run Step 3 manually weekly |
| Source file IDs hard-coded in `FILES` | New month requires code edit | Move `FILES` to config sheet or Script Properties |
| Cache split across column A chunks + B1 + C1 | Growth past ~200KB may need more splits | Monitor cache-size logs |
| Projection uses hybrid blend (not seasonality-aware) | Bimodal payment pattern makes partial-month estimates imprecise | Future: full percentile completion curve |
| Goal ฿4M/mo is hardcoded | Cannot change without code edit | Move to Script Properties or config cell |

---

## 8. Helpers & Utilities

| Function | Role |
|---|---|
| `colMap(headers)` | Build `{header → index}` map for column lookup by name |
| `str(row, idx)` | Read cell as trimmed string |
| `num(row, idx)` | Read cell as number; NaN → empty string |
| `val(row, idx)` | Read cell raw (used for Dates) |
| `formatDate(d)` | Convert Date to `YYYY-MM-DD` |
| `writeStaging(ss, name, headers, rows)` | Write staging sheet with formatting |
| `writeMaster(ss, name, headers, rows)` | Write master sheet with formatting + frozen row |
| `showAlert(msg)` | Log + try to show UI alert (graceful fallback under triggers) |
| `mergeCacheParts_(cacheSh)` | Read B1 for chunk count → loop A1..AN → merge full payload |
| `getSSId()` | Read SS ID from PropertiesService; saves on first editor run |

---

## 9. Run Book (for the operator)

**First-time setup:**
1. Open Apps Script editor → run `setupTriggers()` (grants OAuth scopes + sets SS_ID in Properties)
2. Run `runStep1_Extract()` → verify log shows rows for all 14 months
3. Run `runStep2_Build()` → inspect master sheets
4. Run `runStep3_Analyze()` → verify Enrollment Type and Churn Status
5. Deploy as web app → open URL

**Monthly (when a new month's source file is added):**
1. Edit `FILES` in `Code.gs` → add `{ id, mm, yyyy, label }` (add `pkgSheet` if tab name differs)
2. Run Step 1 → Step 2 → Step 3
3. Run `debugUnmappedPrograms()` → extend `PROGRAM_MAP` if needed
4. Update `Dashboard.html` period buttons if new year/quarter needed

**Weekly:**
- Run `runStep3_Analyze()` manually to recompute Churn Status against today's date

**Troubleshooting:**
- Dashboard stale → hard-refresh browser (`Cmd/Ctrl + Shift + R`)
- All charts blank / RAW undefined → check console for JS errors; common cause is runtime error in `renderGoalTracker` or `renderRevChart` stopping execution
- Dashboard shows cache error → re-run Step 3
- `dailyRefresh` column mismatch error → ensure `normalHdrs` count matches `normalRows` push count (currently **12 columns** each)
- Trigger not firing → run `removeTriggers()` then `setupTriggers()`
- New month has 0 rows after Step 1 → check tab name in source sheet; verify `Already Paid?` column exists for new-format sheets (Dec 2025+)
- Enrollment Type wrong for Apr 2026+ → verify source sheet has `Enrollment Type` column in header row 3; check values are `Trial`/`New`/`Renew`
