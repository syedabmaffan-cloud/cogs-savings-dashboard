# COGS Savings Dashboard

Interactive, self-contained savings dashboard for **direct RM/PM issued to the shop floor**
across 10 Akij SBUs, comparing the current issue rate against **Last Year** and the
**approved procurement budget**.

**Live site (GitHub Pages):** https://syedabmaffan-cloud.github.io/cogs-savings-dashboard/

---

## How it works

```
iBOS MCP server (actuals) ─┐
                           ├─► .opencode/cogs-extract.ps1 ─► .opencode/logs/cogs-data.json
ABP Google Sheet (budget) ─┘                                        │
                                                                    ▼
                              .opencode/cogs-build.ps1 ─► docs/data.json + docs/index.html
                                                                    │
                                                          GitHub Pages serves docs/
```

`docs/index.html` loads `docs/data.json` at runtime, so refreshing the data does **not** require
rebuilding the page. There is also a **↻ Refresh data** button in the page.

## Refresh

**Automatic** — `.github/workflows/refresh.yml` runs daily at 01:00 UTC (and on demand via *Actions → Run workflow*),
re-extracts from the MCP server + Google Sheet and commits `docs/`. Pages redeploys automatically.
Requires repo secrets: `COGS_MCP_URL` + `COGS_MCP_TOKEN` (enterprise-api-gateway bearer token),
`COGS_GCLIENT_ID`, `COGS_GCLIENT_SECRET`, `COGS_GREFRESH`. (Use `COGS_MCP_KEY` instead of `COGS_MCP_TOKEN` for the legacy `X-API-Key` endpoint.)

**Manual** —
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .opencode\cogs-extract.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .opencode\cogs-build.ps1
git add docs ; git commit -m "refresh" ; git push
```

## Shared Material Planning ETA

Procurement updates ETA dates and quantities directly on the **Material Planning** tab. Values live in
**Supabase** (PostgREST) and are visible to everyone who opens the dashboard.

One-time setup:

1. Create a Supabase project (free tier is enough).
2. SQL Editor → paste `.opencode/supabase-eta.sql` → **Run**.
3. Project Settings → API → copy **Project URL** and the **anon public** key.
4. Set them in `ETA_CFG` in `.opencode/cogs-dashboard-template.html`, then rebuild and deploy:
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .opencode\cogs-build.ps1
   ```

Leave `ETA_CFG` empty to run local-only (values stay in the browser's cache).

Notes:

- **Anyone with the link can read and edit** — there are no logins. The anon key is public by design.
- Every write is appended to `mp_eta_history`; there is no delete policy, so a bad edit can be recovered.
- Users enter their name in the tab; each row's tooltip shows who last touched it and when.
- The page refreshes ETAs every 90 seconds (when idle on the tab) and has a **Refresh ETA** button.
- If the shared store is unreachable, edits are kept locally and retried automatically.

## Files

| Path | Purpose |
|---|---|
| `docs/index.html` | The dashboard (GitHub Pages entry point) |
| `docs/data.json` | Refreshable dataset the page loads |
| `COGS_Savings_Dashboard.html` | Same page as one offline-capable file (embedded snapshot) |
| `.opencode/cogs-extract.ps1` | Reads actuals from the iBOS MCP server (read-only, paginated) + budget from the Google Sheet |
| `.opencode/cogs-build.ps1` | Writes `docs/data.json` + `docs/index.html` |
| `.opencode/cogs-dashboard-template.html` | Dashboard template |
| `.opencode/supabase-eta.sql` | Schema for the shared Material Planning ETA register (run once in Supabase) |
| `.opencode/cogs-serve.ps1` / `cogs-server.js` | Optional local/LAN web server (with optional login) |
| `.opencode/agent/cogs-savings-dashboard.md` | Agent definition (rules, formulas, scope, drill queries) |
| `.github/workflows/refresh.yml` | Scheduled refresh |

## Scope & method

- **SBUs:** ACCL · AIL · ARMCL · ABSL · APFIL · HRML · AEL · FAL · AAFL · ALEL.
- **Fiscal year:** July → June. Current FY 2026-27; LY 2025-26. **YTD** = Jul → selected month; **MTD** = selected month.
- **Actuals:** `wms.tblInventoryTransaction*`, `TransactionGroupName='Issue Inventory'` and
  `strTransactionTypeName IN ('Issue For ShopFloor','Issue For Shop Floor')`, active rows.
- **RM/PM:** RM = `Direct Raw Materials`; PM = `Packaging Materials` (plus ACCL cement bags).
- **Savings:** `(LY|Budget rate − Current rate) × Current-period qty`, computed at item level then summed.
- **Budget:** Google Sheet tab **`Consolidated Budget Rate all SBU`** (Jul-2026 … Jun-2027). Missing values show
  **N/A** / **No Budget** (never zero). See the **Data Quality** page for fuzzy/unmapped/duplicate mappings.

> ⚠️ The dataset contains live procurement rates and savings. Keep the repository **private**
> (GitHub Pages on a private repo requires GitHub Pro/Team) or restrict access appropriately.

## Notes

- Read-only: only `SELECT` statements are ever issued; no database writes.
- The MCP server caps each response at 200 rows, so extraction pages with `OFFSET/FETCH`.
- Drill-down is SBU → Item → Month in the page; transaction-level rows are produced on demand by the agent.
