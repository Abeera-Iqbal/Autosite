# Review Report Runner

A small website that lets someone pick **Burger King** or **Denny's**, pick a
**month** (and, for Denny's, a day range), and run your existing Selenium /
Playwright scraping pipeline as a GitHub Action. When the run finishes, the
site shows download links for the resulting **XLSX** and **JSON** reports.

```
project/
├── automation/                 # your existing scripts, now period-aware
│   ├── config.py                # reads TARGET_YEAR/MONTH/DAY + RUN_ID from env
│   ├── json_export.py           # NEW – mirrors a finished xlsx report to .json
│   ├── download.py               (BK)   – uses TARGET_MONTH_LABEL instead of "today"
│   ├── dennys.py                (Denny's) – uses TARGET_MONTH/DATE_FROM_DAY/DATE_TO_DAY
│   ├── combinefiles.py          (BK)   – now uploads Final_BK_Report.xlsx + .json, tagged
│   ├── dennys-record.py         (Denny's) – now uploads Dennys_record.xlsx + .json, tagged
│   ├── dennys-summary.py        (Denny's) – now uploads Dennys_Summary.xlsx + .json, tagged
│   └── ... (unchanged intermediate steps)
├── .github/workflows/
│   ├── bk-report.yml            # workflow_dispatch: month, year, run_id
│   └── dennys-report.yml        # workflow_dispatch: month, year, date_from_day, date_to_day, run_id
├── api/                         # Vercel serverless functions (hold the GitHub PAT)
│   ├── trigger.js               # POST -> dispatches the right workflow, resolves its run id
│   └── status.js                # GET  -> polls run status/conclusion
├── frontend/                    # static site (no build step)
│   ├── index.html
│   ├── style.css
│   └── app.js
└── vercel.json
```

## How it fits together

1. The frontend posts `{ reportType, year, month, dateFromDay?, dateToDay? }`
   to `/api/trigger`.
2. `trigger.js` calls GitHub's `workflow_dispatch` API for `bk-report.yml` or
   `dennys-report.yml`, passing along a generated `run_id` (e.g.
   `bk-202607-1720000000`) as a workflow input.
3. Because GitHub's dispatch endpoint doesn't hand back a run id, `trigger.js`
   immediately polls the workflow's runs list for the newest run and returns
   its numeric id to the frontend.
4. The frontend polls `/api/status?workflowRunId=...` until GitHub reports
   `completed`.
5. Each workflow's last step (`combinefiles.py`, or `dennys-record.py` /
   `dennys-summary.py`) uploads the final report **twice** — as `.xlsx` and
   `.json` — to your existing Test Records API, with `RUN_ID` baked into the
   filename (e.g. `Final_BK_Report_2026-07_bk-202607-1720000000.xlsx`).
6. The frontend polls that same Test Records API directly
   (`GET https://distapi.cybussolutions.com/api/v1/test-records/files`),
   filters for filenames containing the `run_id`, and shows download links.

## One-time setup

### 1. Repo secrets (Settings → Secrets and variables → Actions)

| Secret | Used by |
|---|---|
| `BK_USERNAME`, `BK_PASSWORD` | `bk-report.yml` |
| `DS_USERNAME`, `DS_PASSWORD` | `dennys-report.yml` |

### 2. Push this repo to GitHub

Commit everything as-is — the workflow files are already under
`.github/workflows/`, so GitHub picks them up automatically once pushed.

### 3. Deploy the frontend + API functions to Vercel

```bash
npm i -g vercel   # if you don't have it
vercel
```

In the Vercel project settings, add these **Environment Variables**:

| Variable | Value |
|---|---|
| `GH_OWNER` | your GitHub username or org |
| `GH_REPO` | the repo name (this repo) |
| `GH_PAT` | a GitHub PAT with `repo` + `workflow` scopes |
| `GH_BRANCH` | branch to dispatch from (default `main` if unset) |

**Never** put `GH_PAT` in the frontend — it only ever lives in the
serverless functions (`api/trigger.js`, `api/status.js`), which is why those
exist instead of calling the GitHub API straight from the browser.

### 4. Try it

Open the deployed site → pick BK or Denny's → pick a month → **Run report**.
The status card tracks the Action run, and finished files show up as
download links once the pipeline uploads them.

## Notes / things to double check

- `download.py`'s DCR URL still has `methodology=2026` hardcoded — update
  that if BK's insights portal changes its methodology year.
- The Test Records API (`distapi.cybussolutions.com`) must allow
  cross-origin `GET /files` requests from your frontend's domain for step 6
  above to work from the browser. If it doesn't, add a third serverless
  function (`api/files.js`) that proxies that call server-side — same
  pattern as `status.js`.
- `RUN_ID` also has a safe local fallback (`YYYYMM-local`) so any script can
  still be run manually without the workflow/frontend involved.
