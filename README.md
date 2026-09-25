# RJJ Creative Brief Dashboard (Pod B3)

Asana-backed dashboard for Pod B3 ("RJJ" — Rocky Iannaci, Julian DiVito, Jay
Jeong), comparing weekly creative brief volume by account and creative type.

Runs as a Google Apps Script web app: a scheduled sync pulls and filters
Asana data into a Google Sheet, and the dashboard page only ever reads that
sheet, so it loads fast regardless of how large the Asana project is.

## Layout

- `src/config.gs` — pod config (people, accounts, pod label). Add a pod by
  duplicating one entry in `POD_CONFIGS` and pointing it at that pod's Asana
  project/filter; nothing else needs to change.
- `src/asana.gs` — Asana REST client (pagination, auth). No filtering here.
- `src/sync.gs` — filters tasks to this pod's briefs, aggregates, writes to
  the `Data` and `SyncLog` sheet tabs. `syncAsanaData()` is the entrypoint.
- `src/code.gs` — `doGet()` and the read-only aggregation the dashboard
  calls (`getDashboardData()`, `refreshData()`). Never calls Asana directly.
- `src/bigquery.gs` — service-account BigQuery REST client (JWT auth +
  query runner). Generic; not specific to ad performance.
- `src/adperformance.gs` — pulls creative/ad spend + funnel data from
  BigQuery, writes it to the `AdPerformance` sheet tab, and aggregates it
  for the Ad Performance dashboard tab. Optional: does nothing until
  `BIGQUERY_SERVICE_ACCOUNT_KEY` is set (see below).
- `src/index.html` — the dashboard UI (Chart.js via CDN).

## One-time setup

1. Install clasp and log in (opens a browser for Google OAuth):
   ```
   npm install -g @google/clasp
   clasp login
   ```
2. Create the Apps Script project (run from the repo root) and copy the
   resulting script ID into `.clasp.json`:
   ```
   clasp create --type webapp --title "RJJ Creative Dashboard" --rootDir ./src
   ```
   If `.clasp.json` already exists with a placeholder ID, `clasp create`
   will complain — either delete the placeholder file first, or manually
   paste the new script ID into it.
3. Push the code:
   ```
   clasp push
   ```
4. Open the project in the Apps Script editor (`clasp open`) and set the
   Asana token as a Script Property (Project Settings → Script Properties):
   - Key: `ASANA_TOKEN`
   - Value: your Asana Personal Access Token

   (Alternatively, temporarily paste the token into `setAsanaToken_()` in
   `asana.gs`, run that function once from the editor, then remove the
   literal value from the file before committing.)
5. Before relying on the filters, confirm the exact "Client Name" custom
   field values in the live Asana project match the strings in
   `POD_CONFIGS.B3.accounts` in `config.gs` (Asana may abbreviate them in
   the UI). Update the config if they differ.
6. Run `syncAsanaData` once manually from the editor (select it in the
   function dropdown, click Run) to do the first sync and create the data
   spreadsheet. Check the `SyncLog` tab in the generated sheet — it logs
   how many tasks were excluded by the account rule vs. the person rule,
   which makes a mismatched client-name string easy to spot.
7. Run `createHourlyTrigger_` once from the editor to install the automatic
   scheduled sync. Re-running it is safe — it clears any existing
   `syncAsanaData` trigger first so it never double-schedules.
8. Deploy as a web app:
   ```
   clasp deploy
   ```
   Then in the Apps Script editor: Deploy → Manage deployments → copy the
   `.../exec` URL. Bookmark that URL — it's the dashboard.

## Updating after code changes

Always redeploy to the **same deployment ID** your bookmark uses — plain
`clasp deploy` with no ID creates a brand new deployment with a brand new
URL, silently orphaning your bookmark on the old code.

```
clasp push
clasp deploy --deploymentId <your-bookmarked-deployment-id>
```

Find your deployment IDs with `clasp deployments` if you don't have it
handy — it's the id segment in your bookmarked `.../s/<id>/exec` URL.

## Ad Performance tab (optional, BigQuery-backed)

Shows creative/ad spend and funnel data (spend, leads, CPL, ROAS, etc.)
next to brief volume, sourced from the same BigQuery tables Looker
Studio's LSR report reads from — not via LSR itself. LSR's own connector
(`lsr-mcp`) authenticates by reusing a signed-in Chrome session on a
local machine; that only works for a local Claude Desktop session, not a
cloud Apps Script web app, so this tab talks to BigQuery directly with a
service account instead.

**Setup:**

1. Someone with admin rights on the `cdm-ads` GCP project creates a
   service account, grants it **BigQuery Data Viewer** on the dataset
   containing the ad performance table and **BigQuery Job User** at the
   project level, and generates a JSON key for it.
2. Open the Apps Script editor and add that key as a Script Property:
   - Key: `BIGQUERY_SERVICE_ACCOUNT_KEY`
   - Value: the full JSON key file contents (paste it in directly — never
     commit it to this repo or paste it into a chat).
3. **Verify the config in `config.gs` before trusting any numbers.**
   `AD_PERFORMANCE_CONFIG` (dataset/table/column names) and each pod's
   `adPerformanceClientNameMap` / `adPerformanceMediaBuyerMap` are
   best-guess placeholders based on the LSR connector's captured schema —
   they were never run against a live BigQuery connection. Run
   `syncAdPerformanceData` once from the editor, then check the
   `AdPerformanceSyncLog` tab: `unmapped_client_names_seen` and
   `unmapped_media_buyers_seen` list any real BigQuery value that didn't
   match the config, the same way `SyncLog` flags Asana name mismatches.
   Fix the config and re-run until both are empty.
4. Run `createAdPerformanceHourlyTrigger_` once to keep it refreshed
   automatically. This is a separate trigger from the Asana sync's, by
   design — a BigQuery problem should never be able to take down the
   Asana-backed parts of the dashboard.

## Adding a second pod (B1 or B2) later

Add a new entry to `POD_CONFIGS` in `config.gs` with that pod's `label`,
`people`, `accounts`, and `asanaProjectGid`, then either change
`ACTIVE_POD_KEY` for a dedicated deployment, or extend `doGet`/`code.gs` to
pick the pod from a URL parameter. `sync.gs` and `code.gs` don't need any
other changes since they only ever read from `getActivePodConfig()`.
