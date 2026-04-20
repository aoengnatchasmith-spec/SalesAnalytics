# Product Requirements Document — BeGifted Sales Dashboard

- **Owner:** aoengnatchasmith@gmail.com
- **Platform:** Google Apps Script (V8 runtime) + HTML Web App (served via `doGet`)
- **Timezone:** `Asia/Bangkok`
- **Deployment:** Web app, `executeAs: USER_DEPLOYING`, `access: ANYONE`
- **Apps Script Project ID:** `1PYgKjKIyR8XzESqBHxvID1fXQFKohLdtmu2wCce0ybytbm1BB0LEOpFd`
- **Source files:** `code.js` + `Dashboard.html` + `appsscript.json`

---

## 0. Purpose

Consolidate BeGifted's monthly sales data — spread across multiple Google Sheets (one file per month) — into a single normalized master dataset, classify each transaction by customer lifecycle, and expose it as a filterable sales dashboard (by period, revenue type, and Sales Rep). The goal is to give the sales team and management real-time visibility without anyone having to build monthly reports by hand.

The system consists of 4 main stages (Extract → Build → Analyze → Dashboard) plus an automatic daily refresh at 01:00 Bangkok time.

---

## Step 1 — Extract: Pull raw data before any transformation

### 1.1 Source files

Code reference: `code.js:16–21` (constant `FILES`)

| # | Label | Google Sheet ID |
|---|---|---|
| 1 | (5) 2026 01 BeGifted Sales Record | `1z9LAQbZ-V2GYLm_NA5lkkhR8fdXqyiUzW9EuiHJyeJM` |
| 2 | (5) 2026 02 BeGifted Sales Record | `1dRZjgRP3f0isr-ssZxobwhlsw1v8WWzR0v4zMR82o3k` |
| 3 | (5) 2026 03 BeGifted Sales Record | `1G3wgBV9KnSyqNiSwHKULmbtgEbJnnLTCR-zDBqalS4w` |
| 4 | (5) 2026 04 BeGifted Sales Record | `1HHtZ6YYCqK8wI6nYvVXpwgHSrqoFzcPOD7mMz8hQVJk` |

Data is pulled live via `SpreadsheetApp.openById()` on every run (no caching between runs). If a file can't be opened (missing permissions / deleted), the function logs a `FAIL` entry and skips that file — it does **not** abort the whole pipeline.

### 1.2 Sheet `(1)PackageSales` → Staging `NormalSales_<MM><YYYY>`

- Header row is at **row 3** (constant `HEADER_ROW = 3`)
- Columns are read by header **name**, not by index — resilient to column reordering in the source sheet

**Columns extracted** (code reference: `code.js:81–98` — `extractNormalSales`):

| Source column | Description |
|---|---|
| `วันที่ชำระเงิน` | Actual payment date — used as the primary Transaction/Payment Date |
| `ผู้ขาย` | Sales representative who owns the deal |
| `Student's Nickname` | Student nickname — primary grouping key |
| `Program` | Course/program name |
| `Package` | Hours or package type (e.g. Trial, 10-hr, 20-hr) |
| `No. of Student` | Number of students per session (1/2/3 STU) |
| `ยอดชำระสุทธิ` | Net payment amount |
| `Valid Until` | Package expiry date — used by churn logic in Step 3 |

**Staging sheet name:** `NormalSales_<MM><YYYY>` — e.g. `NormalSales_012026`, `NormalSales_022026`, etc.

**Provenance:** a `Source Month` column is added to the staging sheet (values like `"2026-01 Jan"`) so that every row can be traced back to its source file in Step 2.

### 1.3 Sheet `(2)AdditionalSales` → Staging `AdditionalSales_<MM><YYYY>`

- Header also on row 3
- Columns extracted (code reference: `code.js:100–116` — `extractAdditionalSales`):

| Source column | Description |
|---|---|
| `วันที่ชำระเงิน` | Payment date |
| `Student's Nickname` | Student nickname |
| `Sales Type` | Type of add-on revenue (e.g. Stationery, Textbook) |
| `Package` | Package detail |
| `ยอดชำระสุทธิ` | Net payment amount |

**Staging sheet name:** `AdditionalSales_<MM><YYYY>` — e.g. `AdditionalSales_012026`

### 1.4 Row filtering rules (both sheets)

A row is **skipped** (not written to staging) if:
1. `Student's Nickname` is empty
2. `วันที่ชำระเงิน` is empty or an invalid date → treated as **payment not yet received**, not a real transaction

**Rationale:** when a customer commits to a package but hasn't transferred the money yet, the upstream sheet often has the row open but with no payment date. Filtering these out prevents phantom revenue in downstream reports.

### 1.5 Entry point

- **Manual:** `runStep1_Extract()` — invoked from the Apps Script editor
- **Output alert:** `"Step 1 complete"` with per-sheet row counts

---

## Step 2 — Build: Create the Master Normalized Data Model

Combine all staging sheets from Step 1 into two master sheets (code reference: `code.js:120–168` — `runStep2_Build`).

### 2.1 `MasterNormalized_NormalSales`

Concatenates `NormalSales_012026`, `NormalSales_022026`, `NormalSales_032026`, `NormalSales_042026` and renames columns as follows:

| Staging (Step 1) | Master (Step 2) | Notes |
|---|---|---|
| `Student's Nickname` | `Student's Nickname` | unchanged |
| `Program` | `Program` | unchanged |
| `Package` | `Package Hours` | renamed |
| `No. of Student` | `No. of Student` | unchanged |
| `ยอดชำระสุทธิ` | `Payment Amount` | renamed to English |
| `ผู้ขาย` | `Sales Representative` | renamed to English |
| `วันที่ชำระเงิน` | `Payment Date` | renamed to English |
| `Source Month` | `Source Month` | unchanged — source file traceability |
| *(new)* | `Enrollment Type` | empty — populated in Step 3 |
| *(new)* | `Program (Wise Name)` | empty — populated in Step 3 |
| *(new)* | `Package Hours (Clean)` | empty — populated in Step 3 |
| `Valid Until` | `Valid Until` | unchanged — used by churn logic in Step 3 |

**Formatting applied automatically:**
- Header row: background `#003087`, white text, bold
- Alternating row stripes: `#F0F4FF` / `#FFFFFF`
- `Payment Date` + `Valid Until`: formatted as `dd/mm/yyyy`
- `Payment Amount`: formatted as `#,##0.00`
- Row 1 frozen; columns auto-resized

### 2.2 `MasterNormalized_AdditionalSales`

Concatenates `AdditionalSales_012026` … `AdditionalSales_042026`:

| Staging (Step 1) | Master (Step 2) | Notes |
|---|---|---|
| `Student's Nickname` | `Student's Nickname` | unchanged |
| `Sales Type` | `Sales Type` | unchanged |
| `Package` | `Package` | unchanged |
| `ยอดชำระสุทธิ` | `Payment Amount` | renamed |
| `วันที่ชำระเงิน` | `Payment Date` | renamed |
| `Source Month` | `Source Month` | unchanged |

### 2.3 Entry point

- **Manual:** `runStep2_Build()` — requires Step 1 to have run first
- **Alert:** reports the number of rows written to each master sheet

---

## Step 3 — Analyze & Transform

Code reference: `code.js:172–338` — `runStep3_Analyze` (manual) and `code.js:638–704` — `runStep3_Analyze_silent` (used by the daily trigger).

### 3.1 Enrollment Type Classification

Adds an `Enrollment Type` column to `MasterNormalized_NormalSales`, with three possible values: **Trial / New Student / Renewal**.

**Procedure:**

1. Group rows by `Student's Nickname` (lowercased + trimmed, so `"Nong"` = `"nong"` = `" Nong "`)
2. Within each group, sort by `Payment Date` **oldest → newest**
3. For each row, apply these rules:

| Package Hours | Enrollment Type |
|---|---|
| `Trial` (case-insensitive) | **Trial** |
| Other (first non-trial row for the student) AND preceding row is a Trial | **New Student** |
| Other (first non-trial row) AND no Trial precedes it | **Renewal** |
| Other (any subsequent non-trial row) | **Renewal** |

**Implications:**
- Students who "skip Trial" (go straight to a paid package) are counted as **Renewal** on their first row, even though they're technically brand new — this is a data limitation: the source sheet has no explicit "new customer" flag.
- Students with **2 rows (Trial + 10-hr)** → Trial + New Student
- Students with **3+ rows (Trial + 10-hr + 20-hr …)** → Trial + New Student + Renewal + Renewal …
- Students with **multiple rows but no Trial** (e.g. 10-hr + 20-hr) → Renewal across the board

**Note:** the current code does **not** validate `Payment Amount < 1000` for Trial rows — if a row has `Package Hours = "Trial"` but a payment over 1000 THB, it's still classified as Trial. This could be added as a sanity-check in future.

**Visual:** background color applied per type — Trial `#DBEAFE` (light blue), New Student `#D1FAE5` (light green), Renewal `#FEF3C7` (light yellow).

### 3.2 Program (Wise Name) — Standardization

Adds a `Program (Wise Name)` column by looking up `Program` in the `PROGRAM_MAP` constant (`code.js:37–53`). Values not found in the map pass through unchanged.

| Source `Program` | → `Program (Wise Name)` |
|---|---|
| School Curriculum | Y2-8 / G1-7 (Int.) |
| School Curriculum (2 STU) | (2-STU) Y2-8 / G1-7 (Int.) |
| School Curriculum (3 STU) | (3-STU) Y2-8 / G1-7 (Int.) |
| School Curriculum Master | Y2-8 / G1-7 (Int.) Master |
| School Curriculum Master (2 STU) | (2-STU) Y2-8 / G1-7 (Int.) Master |
| School Curriculum Master (3 STU) | (3-STU) Y2-8 / G1-7 (Int.) Master |
| Admission Exam Prep 11+/13+ | 11+/13+ |
| Admission Exam Prep 11+/13+ (2 STU) | (2-STU) 11+/13+ |
| Admission Exam Prep 11+/13+ (3 STU) | (3-STU) 11+/13+ |
| Admission Exam Prep 11+/13+ Master | 11+/13+ Master |
| Admission Exam Prep 16+ | 16+ |
| Admission Exam Prep 16+ (2 STU) | (2-STU) 16+ |
| Admission Exam Prep 16+ (3 STU) | (3-STU) 16+ |
| IGCSE | Y9-11 / G8-10 (Int.) |
| IGCSE (2 STU) | (2-STU) Y9-11 / G8-10 (Int.) |
| IGCSE (3 STU) | (3-STU) Y9-11 / G8-10 (Int.) |
| IGCSE Master | Y9-11 / G8-10 (Int.) Master |
| IGCSE Master (2 STU) | (2-STU) Y9-11 / G8-10 (Int.) Master |
| A-level OR IB Diploma | Y12-13 / G11-12 (Int.) |
| A-level | Y12-13 / G11-12 (Int.) |
| IB Diploma | Y12-13 / G11-12 (Int.) |
| A-level (2 STU) OR IB Diploma (2 STU) | (2-STU) Y12-13 / G11-12 (Int.) |
| A-level (2 STU) | (2-STU) Y12-13 / G11-12 (Int.) |
| IB Diploma (2 STU) | (2-STU) Y12-13 / G11-12 (Int.) |
| A-level (3 STU) OR IB Diploma (3 STU) | (3-STU) Y12-13 / G11-12 (Int.) |
| A-level (3 STU) | (3-STU) Y12-13 / G11-12 (Int.) |
| IB Diploma (3 STU) | (3-STU) Y12-13 / G11-12 (Int.) |
| A-Level Master OR IB Diploma Master | Y12-13 / G11-12 (Int.) Master |
| A-Level Master | Y12-13 / G11-12 (Int.) Master |
| IB Diploma Master | Y12-13 / G11-12 (Int.) Master |
| A-Level Master (2 STU) OR IB Diploma Master (2 STU) | (2-STU) Y12-13 / G11-12 (Int.) Master |
| A-Level Master (2 STU) | (2-STU) Y12-13 / G11-12 (Int.) Master |
| IB Diploma Master (2 STU) | (2-STU) Y12-13 / G11-12 (Int.) Master |
| A-Level Master (3 STU) OR IB Diploma Master (3 STU) | (3-STU) Y12-13 / G11-12 (Int.) Master |
| A-Level Master (3 STU) | (3-STU) Y12-13 / G11-12 (Int.) Master |
| IB Diploma Master (3 STU) | (3-STU) Y12-13 / G11-12 (Int.) Master |
| GED | GED |
| GED (2 STU) | (2-STU) GED |
| GED (3 STU) | (3-STU) GED |
| SAT | SAT |
| SAT (2 STU) | (2-STU) SAT |
| SAT (3 STU) | (3-STU) SAT |
| SAT Master | SAT Master |
| SAT Master (2 STU) | (2-STU) SAT Master |
| SAT Master (3 STU) | (3-STU) SAT Master |
| IELTS/TOEFL | IELTS |
| IELTS/TOEFL (2 STU) | (2-STU) IELTS |
| IELTS/TOEFL (3 STU) | (3-STU) IELTS |
| IELTS/TOEFL Master | IELTS Master |
| IELTS/TOEFL Master (2 STU) | (2-STU) IELTS Master |
| University | University |
| University (2 STU) | (2-STU) University |
| University (3 STU) | (3-STU) University |
| University Master | University Master |
| University Master (2 STU) | (2-STU) University Master |
| University Master (3 STU) | (3-STU) University Master |
| English Master Class | English Masterclass |
| Interview Prep | Interview Prep |

**⚠️ Courses not yet in the map (known gap):**
- **16+ Master** — a new course that did not exist in legacy systems. It currently passes through with its original name. Needs to be added to `PROGRAM_MAP` once its real source-column value is seen in the sheets.

### 3.3 Package Hours (Clean)

Adds a `Package Hours (Clean)` column that strips the parenthetical from `Package Hours`, e.g.:
- `"30-hr (free extra 1 hr)"` → `"30-hr"`
- `"60-hr (free extra 3 hrs)"` → `"60-hr"`
- `"Trial"` → `"Trial"` (no parenthetical, unchanged)

Implementation: regex `/\s*\(.*?\)/g` replaced with empty string.

**Rationale:** keeps dashboard groupings (e.g. the Package Hours chart) from splintering across near-identical values like `30-hr` vs `30-hr (free extra 1 hr)`.

### 3.4 Churn Status Classification

Adds a `Churn Status` column, populated **only on the latest row for each student** (all other rows show a dash `—`). Four possible values:

| Status | Condition |
|---|---|
| **Active** | `Valid Until + 14-day grace period` ≥ today (still within course or grace window) |
| **Retained** | Past the grace period, but the student made a newer payment |
| **Churned** | Past the grace period, and no newer payment exists |
| **N/A** | Trial-only student, OR no `Valid Until` available |

**Grace period:** 14 days (constant `GRACE_DAYS = 14`) — meaning a student whose package expired 10 days ago is **not** yet counted as churned, because they may still renew.

**Which `Valid Until` is used:** the one on the **most recent non-Trial row** for the student (Trial `Valid Until` is ignored because it doesn't represent a real paid commitment).

**Visual:** Active `#D1FAE5` (green), Retained `#FEF3C7` (yellow), Churned `#FEE2E2` (red), N/A `#F3F4F6` (gray).

**⚠️ Limitation:** Churn Status is computed **only** inside `runStep3_Analyze()` (manual). The daily trigger's silent path does **not** recompute churn (see Section 5 — Known Issues).

### 3.5 Post-analysis: Build Dashboard Cache

Immediately after Step 3 finishes, `buildDashboardCache(ss)` is called (`code.js:342–492`). Details in Step 4.

### 3.6 Entry point

- **Manual:** `runStep3_Analyze()`
- **Alert:** summary of rows processed + Enrollment Type breakdown + Churn Status breakdown

### 3.7 Helper: Unmapped Programs

`debugUnmappedPrograms()` (`code.js:708–731`) — finds `Program` values that don't exist in `PROGRAM_MAP` and:
- Highlights those rows in orange `#FED7AA` in the master sheet
- Creates/updates the `UnmappedPrograms_Log` sheet with the unmatched program names and their row counts
- **Should be run after every Step 3** when you suspect a new course may have been added upstream

---

## Step 4 — Dashboard (Web App)

### 4.1 Architecture

- `doGet()` (`code.js:496–500`) — serves `Dashboard.html` as a web app
- `getDashboardData()` (`code.js:502–511`) — reads the cache sheet `Dashboard_Cache` (3 cells) and parses the JSON
- The cache is split across **3 cells** because a single Google Sheets cell is capped at 50,000 characters:
  - Cell (1,1): `normalDays` — per-day breakdown of Normal Sales
  - Cell (1,2): `aggregates` — YTD totals, `pkgCount`, `progCount`, `repArr`, `dayCount`, etc.
  - Cell (1,3): `churnList` — per-student churn entries with `validUntil` + status
- The cache sheet is **hidden** from end users

### 4.2 Period Selector (top of page)

| Button | Period |
|---|---|
| **YTD** (default) | 2026-01-01 → today |
| **Q1 2026** | Jan–Mar 2026 |
| **Apr 2026** | current month |
| **Custom** | custom date range, via `setPeriod()` in the client JS |

When the period changes, **all KPIs and charts re-render client-side** — no backend call is made (the full payload is already in the cached JS objects).

### 4.3 KPI Cards (6 cards — top row)

| # | Card | Description | Source |
|---|---|---|---|
| 1 | **Normal Sales** | Sum of `Payment Amount` from `MasterNormalized_NormalSales` over the selected period | `normalDays[].rev` |
| 2 | **Additional Sales** | Sum of `Payment Amount` from `MasterNormalized_AdditionalSales` | `addDays[].rev` |
| 3 | **Total Revenue** | Normal + Additional | computed |
| 4 | **Trial → Paid Conversion** | `unique New Students / unique Trials` (by unique nickname) | `uniqueNewStudents / uniqueTrials` |
| 5 | **Renewal Rate** | % of unique students who have at least one Renewal row | `uniqueRenewals` |
| 6 | **Churned · Replacement** | Count of Churned students vs. new students acquired in the period | filtered `churnList` |

### 4.4 Charts (Row 2 — 3 charts)

#### (a) Revenue over Time (`#cRev`)
- Toggle between 3 modes: **Normal / Additional / Combined**
- Bar chart — x-axis is month; bars show stacked breakdowns of `Payment Amount`
- **Normal mode:** stacked bar split into Trial / New Student / Renewal **plus projection bars** (see 4.4.1)
- **Additional mode:** single-color bar (orange `#E8712B`) — no projection
- **Combined mode:** stacked bar (Trial / New / Renewal / Additional) — no projection

#### 4.4.1 Projection Logic (Normal mode only) — `calcProjection()`

Lives in `Dashboard.html:410–479` — runs **client-side** on every chart re-render (projections are not stored in the backend).

**Two projection modes are produced:**

**(A) "Est. remaining" — filling out the current partial month**

Applied only if the last month in the data equals the month the user is currently viewing (i.e. the month is in progress and not yet complete).

Formula:
```
daysInMonth    = total days in the month (e.g. April = 30)
daysPassed     = today's day-of-month (e.g. 17)
factor         = daysInMonth / daysPassed        (e.g. 30/17 = 1.76)
full_month_est = actual_current * factor         (annualize every tracked metric)
gap            = full_month_est − actual_current
```

**Rendering:** the `gap` is drawn as a semi-transparent segment (`rgba(150,150,220,0.15)`) with a dashed border, stacked on top of the current month's bar. This visually communicates "here's how much more we'd expect if the current pace holds."

Tooltip on this segment shows: `Full month est: ฿X (฿Y remaining)`.

**(B) "Projected" — next month (or next two months)**

Uses a **weighted moving average** over the last 3 months, weighting the most recent month the heaviest:

```
weights     = [1, 2, 3]   (oldest : middle : newest)
wavg[field] = Σ(value × weight) / Σ(weight)
```

**Worked example:** if Jan = 100k, Feb = 120k, Mar = 150k (all actuals):
```
wavg = (100×1 + 120×2 + 150×3) / (1+2+3)
     = (100 + 240 + 450) / 6
     = 131.67k
```
→ April projection = 131.67k

**When the current month is partial:** the `full_month_est` (annualized) value is substituted for the current month's actual inside the weighted-average calculation. This prevents the incomplete month from dragging the average downward.

**How many projection bars are rendered:**
| Situation | Projection |
|---|---|
| Last month in data = current month (partial) | Next 1 month |
| Last month in data = a past month (already complete) | Next 2 months |

**Visual treatment of projection bars:**
- Same colors as actual bars, but at **30% opacity** — `rgba(...,0.3)` instead of solid
- Legend adds a "Projected" chip
- Tooltip title appends `(Projected)` or `(Partial month)` to the month label

**Fields projected:** `rev` (total), `revT` (Trial $), `revN` (New Student $), `revR` (Renewal $), `trial`, `newS`, `renew` (transaction counts).

**Caveats / Limitations:**
- This is a trend-based projection — it **does not account for seasonality** (school terms, holidays, exam cycles)
- If only 1–2 months of data exist, the weighted average uses whatever is available (lower accuracy)
- An outlier month (e.g. a big dip or spike) will skew the projection
- The projection **does not incorporate** churn rate, marketing campaigns, or one-off events
- Treat projections as a **baseline comparison**, not a real forecast

#### (b) Pipeline Health Funnel (`#cFunnel`)
- Line chart showing the funnel: Trial → New Student → Renewal → Retained
- Uses unique student counts within the period

#### (c) Sales Rep Ranking (`#cRep`)
- Bar chart, sorted descending
- Shows `revenue` + `transaction count` per sales rep (from `repArr`)
- Source: `Sales Representative` column in `MasterNormalized_NormalSales`

### 4.5 Charts (Row 3 — 3 charts)

#### (a) Package Hours (`#cPkg`)
- Bar chart — transaction count per package (uses `Package Hours (Clean)`, so `(free extra 1 hr)` variants are merged)
- Source: `pkgCount` (Normal Sales only)

#### (b) Program (Wise Name) (`#cProg`)
- Bar chart — transaction count per program
- Toggle: **Normal / Additional**
- Uses the standardized `Program (Wise Name)` (Normal) or `Package` (Additional)

#### (c) Payment Day of Week (`#cDays`)
- Bar chart with 7 bars (Sun–Sat)
- The peak bar is highlighted in solid `#003087`; others use `rgba(0,48,135,0.2)`
- Source: `dayCount` — derived from `Payment Date.getDay()` per transaction

### 4.6 What the dashboard answers (mapped to the original questions)

| Question | Chart / KPI |
|---|---|
| 1. Total Normal Sales revenue per month | KPI #1 + Revenue Chart (Normal mode) |
| 2. Enrollment Type breakdown (Trial / New / Renewal) | Pipeline Funnel + KPI #4, #5 |
| 3. Most popular package (Normal) | Package Hours Chart |
| 4. Most popular program (Normal) | Program Chart (Normal mode) |
| 5. Top-performing sales rep | Sales Rep Ranking Chart |
| 6. Which day of the week customers pay most | Payment Day of Week Chart |
| 7. Total Additional Sales revenue per month | KPI #2 + Revenue Chart (Additional mode) |
| 8. Most popular package (Additional) | Program Chart (Additional mode) |

### 4.7 Churn — answering the open question

> *"I wasn't sure whether we should use Valid Until to calculate churn, or which column(s) could drive it."*

**✅ The implementation uses `Valid Until` + a 14-day grace period:**

Formula:
```
if Valid Until + 14 days >= today   →  Active (not yet churned)
elif a newer payment exists after grace   →  Retained
else                                        →  Churned
```

**Columns used:**
- `Valid Until` (primary) — package expiry date
- `Payment Date` (secondary) — used to check whether the student renewed after grace expired
- `Package Hours` (auxiliary) — so Trial-only students aren't counted in churn metrics

**Why a 14-day grace period:**
- Customers often renew 1–2 weeks late; cutting churn the moment the package expires would give an overly pessimistic number
- 14 days is a reasonably conservative threshold (adjustable via `GRACE_DAYS` in `code.js:244`)

**Other signals that were considered but not used:**
- `Payment Date` alone (without Valid Until) — rejected, since we wouldn't know how long each package lasts
- Converting `Package Hours` into duration — rejected, since we don't know the student's hours-per-week consumption rate

---

## 5. Automation (Daily Refresh)

Code reference: `code.js:532–634`

### 5.1 `setupTriggers()` — one-time setup
- Deletes any existing project triggers first (prevents duplicates)
- Creates a time-based trigger that invokes `dailyRefresh()` every day at **01:00 Asia/Bangkok**

### 5.2 `dailyRefresh()` — automated pipeline
Runs Step 1 → Step 2 → `runStep3_Analyze_silent()` → rebuild cache, with no UI alerts.

**⚠️ Limitation:** `runStep3_Analyze_silent` does **not** recompute Churn Status (unlike manual `runStep3_Analyze`), because:
- The `Churn Status` values written during the last manual Step 3 run remain in place after each daily refresh
- If manual Step 3 is never re-run, `Churn Status` will become stale relative to today's date
- **Recommendation:** run `runStep3_Analyze()` manually at least once a week, or patch `runStep3_Analyze_silent` to include churn logic

### 5.3 `removeTriggers()` — wipe all project triggers

---

## 6. Entry Points Summary

| Function | Purpose | Invoked from |
|---|---|---|
| `runStep1_Extract()` | Extract 8 staging sheets | Manual (Apps Script editor) |
| `runStep2_Build()` | Build 2 master sheets | Manual |
| `runStep3_Analyze()` | Classify + build cache + alerts | Manual (after Step 2) |
| `setupTriggers()` | Install daily 01:00 trigger | Manual (one-time) |
| `dailyRefresh()` | Automated pipeline | Time-based trigger |
| `removeTriggers()` | Clear all triggers | Manual |
| `debugUnmappedPrograms()` | Find programs missing from MAP | Manual (maintenance) |
| `doGet()` | Serve the web app | HTTP request (user opens URL) |
| `getDashboardData()` | Return cached payload | AJAX from Dashboard.html |

---

## 7. Known Issues & Future Work

### 7.1 Data quality

| Issue | Impact | Mitigation |
|---|---|---|
| Source `Program` value not in `PROGRAM_MAP` | Passes through as-is — dashboard splits it into its own bucket | Run `debugUnmappedPrograms()` + add to `PROGRAM_MAP` |
| `16+ Master` not yet mapped | New course — may fragment dashboard groupings | Add to `PROGRAM_MAP` once the real source value appears |
| Row missing `Valid Until` | Churn Status = N/A | Enforce `Valid Until` in the source sheet |
| Duplicate nicknames (different students) | Incorrect merge — wrong Enrollment/Churn classifications | Use a `Student ID` field instead of nickname, if feasible |

### 7.2 System

| Issue | Impact | Mitigation |
|---|---|---|
| Daily trigger doesn't recompute churn | Churn Status frozen at the last manual Step 3 | Run Step 3 manually at least weekly, or patch the silent function |
| Source file IDs hard-coded | Adding a new month requires a code edit | Move `FILES` to a config sheet or Script Properties |
| Cache 3-cell split | Data growth past ~150 KB will require further splitting | Monitor cache-size logs; split further when needed |
| Timezone hard-coded to `Asia/Bangkok` | Not usable outside Thailand | Acceptable for the current use case |

---

## 8. Helpers & Utilities (for code maintainers)

Code reference: `code.js:735–789`

| Function | Role |
|---|---|
| `colMap(headers)` | Build a `{header → index}` map for column lookup by name |
| `str(row, idx)` | Read cell as a trimmed string |
| `num(row, idx)` | Read cell as a number; NaN → empty string |
| `val(row, idx)` | Read cell raw (used for Dates) |
| `formatDate(d)` | Convert a Date to `YYYY-MM-DD` (guards against invalid dates) |
| `writeStaging(ss, name, headers, rows)` | Write a staging sheet with formatting |
| `writeMaster(ss, name, headers, rows)` | Write a master sheet with formatting + frozen row |
| `showAlert(msg)` | Log + try to show a UI alert (graceful fallback under triggers) |
| `mergeCacheParts_(cacheSh)` | Read 3 cells + parse JSON + merge into a single payload object |

---

## 9. Run Book (for the operator)

**First-time setup:**
1. Open the Apps Script editor → run `setupTriggers()` (grants OAuth scopes)
2. Run `runStep1_Extract()` → check the log shows rows for every month
3. Run `runStep2_Build()` → inspect the master sheets
4. Run `runStep3_Analyze()` → verify Enrollment Type and Churn Status classification
5. Deploy as a web app → use that URL to open the dashboard

**Monthly (when a new month's source file is added):**
1. Edit `FILES` in `code.js` → add a new entry `{ id, mm, yyyy, label }`
2. Run `runStep1_Extract()` + `runStep2_Build()` + `runStep3_Analyze()`
3. Run `debugUnmappedPrograms()` → if unmapped values appear, extend `PROGRAM_MAP`

**Weekly:**
- Run `runStep3_Analyze()` manually so Churn Status is recomputed against today's date (the daily refresh doesn't do this)

**Troubleshooting:**
- Dashboard shows "Cache error" → re-run Step 3
- Trigger isn't firing → run `removeTriggers()` then `setupTriggers()` again
- Numbers don't reconcile → inspect the `Source Month` column in the master sheet and compare against the source file
