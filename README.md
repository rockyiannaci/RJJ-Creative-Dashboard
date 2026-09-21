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

```
clasp push
clasp deploy
```

## Adding a second pod (B1 or B2) later

Add a new entry to `POD_CONFIGS` in `config.gs` with that pod's `label`,
`people`, `accounts`, and `asanaProjectGid`, then either change
`ACTIVE_POD_KEY` for a dedicated deployment, or extend `doGet`/`code.gs` to
pick the pod from a URL parameter. `sync.gs` and `code.gs` don't need any
other changes since they only ever read from `getActivePodConfig()`.
