# COGS Savings Dashboard

Read-only, MCP-driven savings analysis of **direct RM/PM issued to the shop floor** across
10 SBUs, comparing the current issue rate against **Last Year (LY)** and the **approved
procurement budget**.

**Deliverable:** `COGS_Savings_Dashboard.html` — a single self-contained, interactive
dashboard (no internet, no server needed; just open it in a browser).

---

## Quick start

```powershell
# refresh actuals + budget from the MCP server / Google Sheet, then rebuild the dashboard
powershell -NoProfile -ExecutionPolicy Bypass -File .opencode\cogs-extract.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .opencode\cogs-build.ps1
```

Then open `COGS_Savings_Dashboard.html`.

Or run the agent:

```
opencode run --agent cogs-savings-dashboard --auto "refresh the COGS savings dashboard and summarise"
```

---

## Components

| File | Purpose |
|---|---|
| `.opencode\cogs-extract.ps1` | Extracts actuals from the iBOS MCP server (read-only, paginated) + budget rates from the Google Sheet → `.opencode\logs\cogs-data.json` |
| `.opencode\cogs-dashboard-template.html` | Dashboard template (HTML/CSS/JS, data placeholder) |
| `.opencode\cogs-build.ps1` | Injects the dataset into the template → `COGS_Savings_Dashboard.html` |
| `.opencode\agent\cogs-savings-dashboard.md` | Agent definition (rules, formulas, scope, drill queries) |

MCP is the **primary source of actuals**; nothing is hard-coded and no data is written back.

## SBU scope

`ACCL` (BU 4) · `AIL` (BU 224, Akij Ispat) · `ARMCL` (175) · `ABSL` (220) · `APFIL` (8) ·
`HRML` (188) · `AEL` (144) · `FAL` (189) · `AAFL` (232) · `ALEL` (237).
Add an SBU by extending the `$Sbus` array in `cogs-extract.ps1` only.

## Fiscal year

July → June. Current FY **2026-27**; LY **2025-26**. **YTD** = Jul → selected month;
**MTD** = selected month. Periods are dynamic and shift with the data.

## Actuals definition

`wms.tblInventoryTransaction*` where `TransactionGroupName = 'Issue Inventory'` **and**
`strTransactionTypeName IN ('Issue For ShopFloor','Issue For Shop Floor')` (both spellings exist),
active rows, from 2025-07-01. Issues are signed negative → `qty = ABS(qty)`, `rate = ABS(value)/ABS(qty)`.

**Direct RM/PM only:** RM = category `Direct Raw Materials`; PM = item type `Packaging Materials`,
plus ACCL cement bags (typed `Raw Materials` but packaging). Excluded: `Non RM`,
`In-direct/Indirect Raw Materials`, Finished/Semi-finished/Trading goods, MRO.

## Formulas

```
LY Savings     = (LY Rate     − Current Rate) × Current Qty
Budget Savings = (Budget Rate − Current Rate) × Current Qty
```
Computed at **item level first**, then aggregated (never SBU-average-rate × SBU-qty).
Positive = saving, negative = adverse. Zero-current-qty items are excluded from rate savings.

## Budget source

Google Sheet `17WltMox1pa-GxnbxNGfBVIFVeymy8n-4q7B7M03u9e4`, tab **`Consolidated Budget Rate all SBU`**
(monthly BDT rates, Jul-2026 … Jun-2027). *Note:* the gid in the original brief (`852000096`) points to
*06_Procurement Trading Cluster* — a quantity plan with no rates/codes — so the rate tab is used instead.
Match order: SBU+code → SBU+normalised description → fuzzy (≥0.6). HRML/FAL/AAFL/ALEL have no budget rows
and correctly show **"No Budget"** (not zero).

## Dashboard pages

1. **Executive Summary** — KPI cards (RM/PM qty, issue value, avg rate, LY savings, budget savings,
   savings %, adverse impact) · SBU-wise savings · monthly trend (Jul→Jun) · top-10 savings · top-10 adverse.
2. **SBU Analysis** — SBU selector, current vs LY vs budget, monthly trend, item drill-down
   (click an item row for its month-by-month breakdown).
3. **Item Analysis** — full item table with current / LY / budget rate and savings, conditional
   formatting, **Export CSV**.
4. **Data Quality** — items without LY rate, items without budget, unmapped/duplicate budget records,
   fuzzy matches, coverage & classification exceptions.

## Data-quality behaviour

Missing LY rate → **N/A**; missing budget → **No Budget** (never silently zero). Fuzzy budget matches and
unmapped/duplicate budget rows are listed for review. **ALEL** has no direct RM/PM shop-floor issues (all
`Non RM`) — flagged as a master-data matter, not zero savings.

## Calibration (FY2026-27 YTD, Jul–Sep 2026)

918 items extracted (600 with current-period qty); total issue value ≈ **1,156 cr**.
ACCL clinker current rate ≈ 7,774 vs budget 8,343 (favorable). Total **LY savings ≈ −48.4 cr**
(adverse) and **budget savings ≈ +38.7 cr**.

## Sharing / publishing

The dashboard is self-contained, but the recommended way to share it is **GitHub Pages**:

- `docs/index.html` is the page; `docs/data.json` is the dataset it loads at runtime.
- GitHub Pages source = **`main` branch → `/docs`** → live at `https://<user>.github.io/<repo>/`.
- `docs/index.html` has a **↻ Refresh data** button that reloads `docs/data.json` (no rebuild needed).
- **Automatic refresh:** `.github/workflows/refresh.yml` re-extracts from the MCP server + Google Sheet
  daily (and on demand) and commits `docs/`. Needs repo secrets `COGS_MCP_URL`, `COGS_MCP_KEY`,
  `COGS_GCLIENT_ID`, `COGS_GCLIENT_SECRET`, `COGS_GREFRESH`.
- **Manual refresh:** `cogs-extract.ps1` → `cogs-build.ps1` → `git add docs && git commit && git push`.

> ⚠️ This data is confidential (live procurement rates/savings). Use a **private** repo
> (private GitHub Pages requires GitHub Pro/Team) — do not publish on a public URL.

Other options:
- Serve it on your own network as a real web page (access-controlled):
  ```powershell
  powershell -NoProfile -ExecutionPolicy Bypass -File .opencode\cogs-serve.ps1 -User affan -Pass <password> -Port 8080
  powershell -NoProfile -ExecutionPolicy Bypass -File .opencode\cogs-serve.ps1 -Stop
  ```
  It prints Local + Network URLs; stop with `-Stop`.
- Email `COGS_Savings_Dashboard.html` as an attachment, or drop it in a shared network/OneDrive folder.

## Notes / limitations

- The MCP response is capped at 200 rows, so extraction pages with `OFFSET/FETCH`; the pipeline runs
  in a script (not chat) to stay within that limit.
- Dashboard drill is SBU → Item → Month; **transaction-level** rows are produced on request by the agent.
- Read-only throughout: only `SELECT` statements are issued; no writes to any database.
