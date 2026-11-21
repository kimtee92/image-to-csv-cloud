# Image → CSV (Gemini-powered, cloud-ready)

Self-contained FastAPI app that turns scanned Daily Time Records into their
client's timekeeping workbook, using **Google Gemini** via an OpenAI-compatible
endpoint, with a bundled static frontend. No local GPU dependency — everything
needed to deploy is in this one directory.

**Client-first.** You pick the client; that determines both how their DTR is read
(their form's fixed column set) and which workbook the batch is compiled into.
There is no general-purpose "table or text" mode — this is a DTR tool.

**Batch processing:** upload multiple files at once (browse, drag-drop, or
paste). Files are extracted in parallel (3 at a time) and listed in a
paginated overview — page size adjustable (3/5/10/20, default 5). Each file
gets its own CSV: download individually, use "Download All", or compile the whole
batch into the client's workbook with "Excel (batch)". Click any processed file to
open it in the full editor (editable table, header/footer fields, AI amendments).

```
image-to-csv-cloud/
├── app/
│   ├── main.py          # FastAPI backend: OCR, per-client schema locking, validation
│   ├── clients.py       # client registry: input layout + output template per client
│   ├── xlsx_fill.py     # writes rows into a client's own .xlsx, in place
│   ├── templates/       # one timekeeping workbook per client
│   ├── requirements.txt
│   └── static/
│       └── index.html   # frontend (single file, no build step)
├── benchmark/           # hand-read ground truth + scorer
├── Dockerfile
├── docker-compose.yml    # local dev only
├── .env.example
└── .env                  # local secrets (gitignored) — NOT for prod
```

## Configuration

Environment variables (see `.env.example`):

| Var | Default | Notes |
|---|---|---|
| `GEMINI_API_KEY` | — | **required** |
| `GEMINI_MODEL` | `google/gemini-3.5-flash-lite` | **the `google/` prefix is required on Vertex** — a bare name 400s every call |
| `GEMINI_BASE_URL` | Vertex OpenAI-compat endpoint | production points at Vertex AI, not the public Gemini endpoint |
| `PORT` | `8080` | listen port (Cloud Run injects this) |
| `LOSSLESS_MIN_SIDE` | `2000` | upscale small scans so thin strokes span more pixels |
| `LOSSLESS_MAX_PIXELS` | `12000000` | cap after upscaling, to bound payload size |
| `REASONING_EFFORT` | `minimal` | not portable between models — `gemini-3.8-flash` rejects `minimal` outright |
| `DTR_ESCALATE` | `1` | retry unsettled rows against narrower, more magnified crops |
| `ESCALATION_MODEL` | *(empty)* | send that retry to another model; empty = same model |
| `TIME_VERIFY` | `1` | verify in/out/total arithmetic (printed forms) and row chronology (DTRs), repairing misread digits |
| `OCR_TIMEOUT` | `170` | seconds for the extraction call |
| `VERIFY_TIMEOUT` | `90` | seconds for the verification re-read; `OCR_TIMEOUT + VERIFY_TIMEOUT` must stay under Cloud Run's request timeout (300s) |

## Batch size

**There is no limit on how many files you can upload**, and batch position does not
affect accuracy. The "Rows per page" selector (default 5) controls only how many
result cards are *displayed* at once — every uploaded file is processed regardless.
This was verified by running the same 20 forms in forward and reverse order: the
errors followed the individual forms, not their position in the batch (files 1–5
scored 100% then 98.1%, files 6+ scored 96.5% then 99.7% — the ranking inverts
with the ordering, because the harder forms simply sort later alphabetically).

What *does* scale with batch size is the absolute number of errors: at ~97% per-group
accuracy, 5 sheets shows a bug or two while 98 sheets shows dozens. Same rate, more
occurrences — which is why the review flags below matter more as batches grow.

Two silent failure modes are therefore caught explicitly:

- **Empty extractions** (~1–2%) are treated as failures, not successes, so a blank
  CSV can't hide mid-batch. They're retried once automatically, as are transient
  5xx errors, then marked failed with a Retry button.
- **Gemini rate limits (429).** A large batch can outrun the project's quota. The
  upstream 429 is passed through as a **429** rather than flattened into a 502,
  because the two need opposite handling — a 502 is worth retrying in a second, a
  quota exhaustion needs tens of seconds. Rate-limited files get 5 attempts with
  exponential back-off (3s, 6s, 12s, 24s — honouring `Retry-After` when the API
  sends one) instead of the single 1.2s retry other errors get, so a batch no
  longer loses files purely for being large. Observed repeatedly while testing 20
  files back-to-back: before this, one or two files per run failed with
  `RESOURCE_EXHAUSTED`; after, 20/20 complete.
- **Row-count outliers** (~3.3% of extractions returned a different row count than
  the same form usually yields — a dropped or duplicated day). Since a batch is
  normally many copies of one form, any file whose row count differs from the batch's
  dominant count is badged in red. Only applied when the batch has ≥4 files and ≥60%
  share one count, so mixed form types don't flag everything.

## Accuracy on dense forms (timesheets/DTRs)

Two things drive digit-level accuracy, both measured on 20 real timesheet scans
(3 runs each, scored against the totals the forms print themselves):

**1. Don't process the image.** The old enhancement prepass was the single
largest source of errors. Its 3×3 median filter closes the gap in small digits —
`18:45` reads as `16:45` — and greyscaling collapses black-on-red highlighted
cells to almost no contrast. Sending an unfiltered PNG instead moved accuracy
**83.3% → 93.1%** (paired sign test p=0.004). That prepass has since been
**deleted** rather than left behind an env switch someone could flip back on.

**2. Let the form check itself.** These forms print `TOTAL` alongside `Time in`
and `Time out`, so `TOTAL = out − in` is a free correctness check. When a group
fails it, the backend re-reads just those columns from the image — *without*
showing the model its previous answer, since priming it with a misread anchors it
to that misread — and accepts the new values only if they both resolve the
arithmetic and look like a digit-level fix (same length, edit distance ≤ 2).
Anything it can't resolve is returned as `flagged_cells` and highlighted amber in
the UI with a per-file "⚠ N to review" badge, so a bad row in a 98-file batch is
visible instead of silently wrong. Editing a flagged cell clears its flag.

A 2-minute tolerance is applied: these are spreadsheets that compute from
second-precision stamps but print to the minute, so in, out and total can each
round by up to a minute. Real digit errors measured were ≥3 minutes.

Model choice was tested too: `gemini-3.5-flash-lite` was **not** better here
(93.1% vs 91.2%, paired p=0.52 — no measurable difference) and ran ~18% slower
(8.2s vs 6.7s median), so this stays on `gemini-3.1-flash-lite`.

## Handwritten DTRs (the `dtr` layout)

The handwritten Daily Time Records in `reference2/` are a *different* form from the
printed timesheets above: `MORNING / AFTERNOON / EVENING · IN / OUT`, with **no
per-section TOTAL column**. That matters, because it means both existing safety
nets are inert on them — `_find_time_groups` finds nothing without a TOTAL, and
`_find_missing_rows` can't read their numeric `6-26-26` dates. Before this mode,
a handwritten DTR got a single unvalidated pass.

Measured on the 10 `reference2/` scans against hand-built ground truth (2 reps,
scored per time-cell, ambiguous cells excluded, dropped rows counted as errors):

| | generic extractor | DTR mode |
|---|---|---|
| Time-cell accuracy | 97.8% | **99.0–99.2%** |
| Times resolved to 24-hour | 71.2% | **100%** |
| Distinct column schemas across the batch | **9** | **1** |
| Chronology-breaking cells surfaced | 0 | all of them |

Four things do the work:

**1. A fixed 11-column schema.** The same form was coming back with 9, 10 or 11
columns and a different header spelling per page — the empty-column stripper drops
whichever of `NO. OF DAYS/HRS` / `O.T. HRS` is blank on that page, and on degraded
scans the model renames the evening pair `OVERTIME IN/OUT` after the payroll card
printed beside the grid. Pages of one batch then refuse to line up. A recognised
DTR is now remapped onto `DTR_HEADERS` and skips the stripper.

**2. 24-hour normalisation.** Staff write afternoons and evenings in 12-hour form,
so a column would mix `02:46` and `14:48`. The row's own `DAILY TIME SCHEDULE`
cell resolves the ambiguity (`12:15-9:15` ⇒ the shift ends 21:15).

**3. Chronology as the free check.** With no TOTAL to verify against, the
invariant is that a row's six stamps run forward through one day. A row that goes
backwards is provably misread. The check tolerates un-normalised 12-hour input —
each cell may read as itself or itself + 12h — so a correctly-transcribed
12-hour page doesn't light up every row.

**4. Repair by zooming in, not by asking again.** A blind re-read of the full page
returns the *same* misread every time, and re-sending the page at 3000 or 4000px
changes nothing — the model tokenises an image to a fixed budget, so extra
resolution is discarded. Cropping is the only lever that works: a full-width
horizontal band at 2.5× recovered the cell 6/6 times where the full page failed
6/6. Corrections are accepted only when they're a digit-level change (same length,
edit distance ≤ 2) *and* they make the row chronological; anything else becomes a
`flagged_cells` entry for a human. Two overlapping bands are tried, full page
width, so no column is ever cut off on a scan without a punch card beside it.

`Any document` mode still canonicalises and time-checks anything that turns out to
be a DTR; it just uses the generic extraction prompt. Detection is deliberately
strict — the printed timesheets above also have three in/out pairs, so a DTR must
additionally name its morning and afternoon sessions and have no per-section TOTAL.

Two knobs were tried and rejected: `reasoning_effort: low` was **worse** (95.6% vs
98.8%, from dropped rows), and missing-row detection is deliberately *not* enabled
for DTRs — rest days and day-offs are normal gaps here, so it would false-alarm on
nearly every form.

## Break-detail timesheets (`Document Type: Timesheet (with breaks)`)

The other client's forms are a third shape: a date column plus `REGULAR HOURS`,
`LUNCH BREAK`, `COFFEE BREAK` and `OVERTIME`, each with its own in / out / TOTAL.
They drift for the same reason DTRs did, only worse — the *printed* forms come in
12-, 13-, 14-, 15- and 17-column variants (some omit `DUTY LOCATION`, some print
only a lunch total and no lunch in/out, some put `TARDINESS` last). So the
extractor legitimately returned a different schema per page.

`TIMESHEET_HEADERS` is the union of every variant seen across 20 real scans, and a
recognised form is remapped onto all 17 columns with `""` for whatever that page
doesn't print. Verified: the 12-, 14- and 15-column variants all collapse to one
schema with every value in its correct slot.

**Detection is mutually exclusive with the DTR path.** `_is_timesheet_table`
returns False whenever `_is_dtr_table` matches, so the two modes can never both
claim a form. Checked against all 20 printed scans (20/20 timesheet, 0 DTR) and
the DTR header variants (all DTR, 0 timesheet).

Three things were needed before a real 20-file batch actually landed on one schema
— each found by driving the deployed app in Playwright, not by unit tests:

- **The frontend had its own empty-column stripper** (`normalizeTablePayload`)
  which undid the server's locked schema, turning 17 columns back into 12. It now
  skips stripping when `is_dtr` or `is_timesheet` is set. The backend was correct
  the whole time — only the displayed and exported table was wrong, which is
  exactly the sort of bug that never shows up in an API-level test.
- **A date column is not required to lock.** On a bad read the model sometimes
  drops the DATE column entirely; refusing to canonicalise in that case is
  backwards, since the page whose extraction went wrong is the one that most needs
  to keep the batch's shape. It now locks with an empty `DATE`, which is visible.
- **`_timesheet_positional_fallback`** handles runs where the model drops the
  section prefixes and returns `Date, Time In, Time out, TOTAL, BREAK OUT, ...`.
  Section order on these forms is fixed, so the 1st in/out/TOTAL triple is regular
  hours, the 2nd lunch, the 3rd coffee, the 4th overtime. It engages only when the
  named pass found no `REGULAR HOURS` pair **and** at least three triples are
  present — a lone in/out/TOTAL is far too weak a signal to claim a document on.

Verified end to end against the live service: 20/20 files, 17 columns each,
**1 distinct schema**, review flags down from 72–84 per page to 0–6 on all but one.

### Why `auto` is the default

It didn't used to be, and that was a real bug. Defaulting the picker to `dtr` meant
the DTR prompt — *"use EXACTLY these 11 headers: DAILY TIME SCHEDULE, DATE,
MORNING IN …"* — was forced onto break-detail timesheets that have no morning or
afternoon columns at all. The model half-complied: pages came back with `MORNING IN`
headings over regular-hours data, 5–11 columns instead of 12–17, and 70–84 bogus
review flags per page (the chronology check firing on values mapped into slots they
don't belong in).

Pinning a header list is only safe once you know which form you're looking at. So
`auto` runs the generic prompt and then canonicalises whichever shape the result
turns out to be; the two specific modes exist for batches you know are all one form,
and the UI says so.

## Batch → timekeeping workbook (`Excel (batch)`)

`POST /api/dtr-workbook` compiles a whole batch into one `.xlsx` with two tabs,
matching the client's master workbook ("DTR - Timekeeping - Sports.xlsx"):

- **Sheet1** — one block per employee on a 20-row stride (rows 3, 23, 43 …):
  name row, header row, 16 data rows, 2 blanks. Columns `A` Schedule (as written),
  `B` Day, `C` Date, `D–I` the six punches as **Excel time serials**, `J` O.T. HRS,
  `K` Reason. `N–V` and `Y–AB` are the client's own formulas; `AD` is the
  normalised schedule they drive off, `AE` their REMARKS input.
- **DTR Summary** — the client's VLOOKUPs, one row per employee.

Design notes worth keeping:

- **The formulas are not re-authored.** They were extracted from the master
  mechanically and verified to re-render its own cells character-for-character,
  then row-shifted per block. Fed the same punches, the generated workbook
  reproduces every one of the master's computed values exactly (checked cell by
  cell after a LibreOffice recalculation — 0 differences).
- **`AD` must never be blank.** The Late/UT formula parses it unconditionally, so
  an empty schedule yields `#VALUE!` rather than 0. Rows with no readable schedule
  inherit that employee's most common shift and are marked in the `REVIEW` column.
- **Schedules are normalised** (`10:00 to 7:00` → `10:00-7:00`): the computed
  columns locate the separator with `FIND("-", AD..)`, so the "to" form staff
  write would break all of them.
- **Two-punch DTRs** put the single out time in `EVENING OUT`, not `MORNING OUT` —
  that is the column the TOTAL/OT formulas read as the end of the shift.
- **Pages are merged by employee name** (normalised), rows de-duplicated and sorted
  by date; a merge across differing spellings raises a warning rather than
  happening silently. The key is letters-only so `"DELA CRUZ, JUAN"` and
  `"Dela Cruz , Juan"` are one person — but a scan with no readable NAME field
  falls back to the *file* name, and those are often all digits (`0017.jpg`).
  Those reduced to an empty key and silently merged **an entire 20-file batch into
  one employee**, so anything with fewer than three letters now keys on its digits.
- **Warning text is sanitised before it goes in the `X-Export-Warnings` header.**
  Starlette encodes header values as latin-1, so one character outside it raises
  `UnicodeEncodeError` and turns a finished workbook into a 500. The warnings quote
  employee names, so `Vergaño` alone was enough — as was the em dash the messages
  themselves used.
- **`AG` = REVIEW** carries "check this row" notes. It sits outside the client's
  model on purpose — `AE` is their WDO/SH/LH input and the computed columns read
  it, so writing notes there would silently change OT and holiday pay.

Two bugs in the master were **repaired** here rather than reproduced. Both were
verified by recalculating the generated file and diffing against the master:

- **Column `T` (WDO ND)** pointed its four operands at a dead external reference
  (`[1]Sheet1!#REF!`), so the column errored the moment REMARKS contained WDO or
  SH. Each use pins down the operand independently — the guard pairs it with time
  out, one term measures night hours before 06:00 from it, one compares time out
  against it for a midnight crossing, one takes `MAX(it, 22:00)` — and all four are
  the time-in column `D`. The guard is also character-identical to the ones in
  `Q`/`S`/`V`, which read `D`. Substituted accordingly: a rest-day shift punched
  08:00→23:30 now returns 1:30 of night differential instead of `#REF!`, and rows
  without a WDO/SH remark are still 0, so no existing number moves.
- **The Summary VLOOKUPs were each one position short of the column their own
  heading names** — "DAYS" returned N (total hours) rather than O (days worked),
  "OT" returned days worked, and so on down the row — while LATE/UT reached index
  22, which is blank column X, so it read 0 for everybody. Each lookup now points
  at the column its label names, and `TOTAL HRS` is kept as its own column so the
  figure that used to sit under "DAYS" is still on the sheet. For ALFON that turns
  `DAYS 52.27 · OT 5.00 · NDOT 6.00 · WDO 1.50 · LATE/UT 0` into
  `TOTAL HRS 52.27 · DAYS 5.00 · OT 6.00 · NDOT 1.50 · LATE/UT 7` — the same
  underlying values, under the right headings, plus the 7 minutes of lateness the
  old lookup could never see. Sheet1 itself is untouched.

Writing goes through `app/xlsx_writer.py`, a ~200-line stdlib xlsx writer, so the
Cloud Run image doesn't gain openpyxl for what is a few thousand cells.

**Accuracy note:** the schedule column feeds `AD`, which drives the OT/NDOT split —
so a misread *schedule* moves payroll numbers even when all six punches are
correct. Rows whose times failed the chronology check, and rows with an unreadable
schedule, are both marked in `REVIEW`.

### `DTR_CROSSCHECK` — a second opinion (off by default)

The errors that survive everything above are all *chronologically valid* misreads,
so no invariant can see them. Setting `DTR_CROSSCHECK=1` reads every row a second
time from a magnified crop and flags cells where the two readings differ.

Measured on the 10 reference scans, 252 time cells:

| | |
|---|---|
| the two readings agree | 249 / 252 (98.8%) |
| disagree, **crop** right | 1 |
| disagree, **full page** right | 1 |
| disagree, both wrong | 1 |

So the crop is **not** a reliable arbiter — auto-applying it would fix one cell and
break another. Nothing is ever overwritten on its say-so; disagreements are only
flagged. What it buys is detection: on page-0002 the flag lands exactly on the bad
cell, which nothing else in the pipeline can find. It is genuinely imperfect —
on page-0010 it flagged a *different* cell and missed the real error — and it
roughly doubles cost and time per file. Hence off by default.

## Run locally

```bash
docker compose up -d --build
# open http://localhost:8080
```

Already verified working end-to-end (health check + a real OCR call) against
the Gemini key/model you provided — it's running right now at
**http://localhost:8080** for you to try in a browser.

Useful while testing:
```bash
docker compose logs -f     # tail backend logs
docker compose down        # stop it
```

## Deploy — recommended: Google Cloud Run

**Why Cloud Run over Fly.io / Render / Railway for this app:**
- This service is stateless and only does light CPU work (image
  resize/preprocess with Pillow); Gemini does all the heavy lifting off-box.
  That means it's a near-perfect fit for a scale-to-zero request-billed
  platform — you pay ~nothing between requests.
- Free tier: 2M requests/month, 360k GB-seconds, 180k vCPU-seconds free,
  every month, forever (not a trial). A low/moderate-traffic OCR tool like
  this will likely stay entirely inside the free tier.
- Same network as Gemini (both Google infra) → lower latency per request
  than calling Gemini from another cloud.
- One-command deploy straight from source, no registry/config wrangling
  needed (`gcloud run deploy --source .`), and it builds this exact
  Dockerfile via Cloud Build.
- Fly.io/Render free tiers either sleep with slow cold starts on their free
  plans or need a paid plan to avoid it; Railway has no meaningful free
  tier anymore. Cloud Run's scale-to-zero cold start is short (this image
  is small, no ML weights baked in) and the free quota is genuinely free.

### 0. One-time account/project setup (skip if you already have a GCP project)

```bash
# Install: https://cloud.google.com/sdk/docs/install
gcloud auth login

# Create a fresh project (id must be globally unique) or reuse an existing one
gcloud projects create image-to-csv-app --name="Image to CSV"
gcloud config set project image-to-csv-app

# A billing account must be linked for Cloud Run to deploy, even though
# you'll likely stay inside the free tier. List accounts, then link one:
gcloud billing accounts list
gcloud billing projects link image-to-csv-app --billing-account=BILLING_ACCOUNT_ID

gcloud services enable run.googleapis.com secretmanager.googleapis.com cloudbuild.googleapis.com
```

### 1. Store the Gemini API key in Secret Manager

Don't ship the key as a plain `--set-env-vars` value in prod — Secret Manager
keeps it out of the Cloud Run service config/logs:
```bash
printf '%s' 'YOUR_GEMINI_API_KEY' | gcloud secrets create gemini-api-key --data-file=-
```

### 2. Deploy from this directory

```bash
gcloud run deploy image-to-csv \
  --source . \
  --region us-central1 \
  --allow-unauthenticated \
  --memory 512Mi --cpu 1 \
  --min-instances 0 --max-instances 3 \
  --set-env-vars GEMINI_MODEL=gemini-3.1-flash-lite \
  --set-secrets GEMINI_API_KEY=gemini-api-key:latest
```

`gcloud` builds the `Dockerfile` via Cloud Build and prints a `*.run.app`
HTTPS URL when done — that's your live app.

### 3. Verify the deploy

```bash
SERVICE_URL=$(gcloud run services describe image-to-csv --region us-central1 --format='value(status.url)')
curl -s "$SERVICE_URL/api/health"
```
Expect `{"status":"ok","ocr_ready":true,...}`. If `ocr_ready` is `false`,
check the secret is attached correctly (step 1) and that Gemini's API is
reachable/quota isn't exhausted.

### Redeploy after code changes (automatic: push to GitHub)

Two branches, two Cloud Run services, one pipeline:

| Branch | Trigger | Service | URL |
|---|---|---|---|
| `dev` | `deploy-staging-on-push-dev` | `image-to-csv-staging` | https://image-to-csv-staging-358885129386.us-central1.run.app |
| `main` | `deploy-on-push-main` | `image-to-csv` (live) | https://image-to-csv-358885129386.us-central1.run.app |

Workflow, no local machine needed:

1. Make changes on `dev`: GitHub's web editor, a Claude Code session on the
   repo, or a local clone. Each push deploys to **staging** in a few minutes.
2. Test on the staging URL. It uses the same model and settings as live
   (max 2 instances, scales to zero when idle).
3. When it's right, open a pull request `dev` → `main` on GitHub and merge it
   (works from the GitHub mobile app). The merge deploys to **live**.

If a step fails (preflight, recalc, build, health check), nothing is deployed and
the service keeps serving its previous revision. Build results also show as checks
on each commit and pull request on GitHub.

**Spreadsheet check (`recalc` step).** Every build fills each client template with
fixed test shifts, recalculates it in LibreOffice and compares every formula on the
filled rows with what Excel 365 computed for the same export (`ci/golden/*.json`).
It fails on `#NAME?`, on error cells, or on any value that differs. Debian's
LibreOffice lacks `LET`/`XLOOKUP` like Excel 2016/2019, so a 365-only function that
reaches an export is caught as `#NAME?`.

When a change is *meant* to alter computed values (a new or updated client
template, new rules), re-record the expected values on a PC with Excel and Python
3.11, review the diff, and commit it:
```powershell
powershell -ExecutionPolicy Bypass -File ci\record_golden.ps1
```

Both triggers run [cloudbuild.yaml](cloudbuild.yaml); the staging one only overrides `_SERVICE`:

1. **preflight**: `PREFLIGHT_ONLY=1 ./deploy.sh` on `python:3.11-slim`, the same
   syntax, import and client-workbook OOXML checks as a manual deploy. A failure
   stops the build before anything ships.
2. **build / push**: Docker build of the repo, tagged with the commit SHA, into
   Artifact Registry `cloud-run-source-deploy/image-to-csv`.
3. **deploy**: `gcloud run deploy image-to-csv --image …:<sha>`. Only the image
   changes; env vars, secrets, memory and scaling stay as set on the service.
4. **health**: `GET /api/health` must answer, or the build is marked failed.

The build runs as the `cloudbuild-deployer` service account (Cloud Build builder,
Run developer, Artifact Registry writer, log writer, and actAs on the runtime
compute SA). Watch builds in the console under Cloud Build → History, or:
```bash
gcloud builds list --region us-central1 --limit 5
```
Roll back by redirecting traffic to a previous revision:
```bash
gcloud run services update-traffic image-to-csv --region us-central1 --to-revisions REVISION=100
```

### Redeploy manually

`./deploy.sh` still works for a one-off deploy from a machine with gcloud. It
stages, preflights, then deploys:
```bash
./deploy.sh
```

It uploads a temporary staging tree rather than the working directory, because
`gcloud run deploy --source .` sends the *whole* tree to a Cloud Storage staging
bucket and `.dockerignore` does not prevent that — it only controls what lands in
the image. Staging keeps the client scans in `reference/` and `reference2/` out of
the upload entirely.

`app/*.py` is globbed rather than listed file by file (the `Dockerfile` does the
same). `main.py` imports its siblings as top-level modules, so a module left out
of a hand-maintained list builds a perfectly good image that then dies on the
first request with `ImportError`.

Preflight aborts before spending a Cloud Build round trip if: `main.py` or
`index.html` is missing, anything fails to compile, the staged modules differ from
`app/*.py`, or **any module imports something that is neither stdlib, nor pinned in
`requirements.txt`, nor a file in `app/`** — which catches a deleted or misspelled
sibling anywhere in the tree, not just in the entrypoint.

Everything is overridable by environment variable, and only the ones you actually
export are sent:

```bash
DTR_CROSSCHECK=1 ./deploy.sh                 # turn the second-opinion pass on
MEMORY=2Gi CONCURRENCY=2 ./deploy.sh          # more headroom per request
GEMINI_SECRET=gemini-api-key ./deploy.sh      # (re)attach the key from Secret Manager
DRY_RUN=1 ./deploy.sh                         # print the gcloud command, change nothing
```

Recognised: `SERVICE REGION MEMORY CPU CONCURRENCY MIN_INSTANCES MAX_INSTANCES
TIMEOUT GEMINI_MODEL DTR_CROSSCHECK IMAGE_PREPROCESS TIME_VERIFY OCR_TIMEOUT
VERIFY_TIMEOUT LOSSLESS_MIN_SIDE LOSSLESS_MAX_PIXELS GEMINI_SECRET DRY_RUN`.

Two defaults are deliberate:

- **`--memory 1Gi`, `--concurrency 4`.** A DTR re-read decodes the source scan and
  builds a magnified band crop (capped at 8M pixels) on top of the upscaled model
  copy, so peak memory per request is well above what a single-pass extraction
  needed. Raise memory before raising concurrency.
- **`--timeout 300`.** `OCR_TIMEOUT` (170) + `VERIFY_TIMEOUT` (90) = 260s of model
  calls have to finish inside the request timeout, or a slow pair is dropped with
  no result at all.

It uses `--update-env-vars` / `--update-secrets`, not `--set-*`: the `set` forms
**replace** the service's entire environment, silently dropping anything configured
out of band. After deploying it prints the service URL and its `/api/health`.

Plain `gcloud run deploy image-to-csv --source . --region us-central1` still works
if you want to bypass the script, and Cloud Run keeps all previously-set flags.

### Rotate the API key later
```bash
printf '%s' 'NEW_KEY' | gcloud secrets versions add gemini-api-key --data-file=-
gcloud run services update image-to-csv --region us-central1 --set-secrets GEMINI_API_KEY=gemini-api-key:latest
```

### Logs / debugging
```bash
gcloud run services logs read image-to-csv --region us-central1 --limit 50
```

### Custom domain (optional)
```bash
gcloud run domain-mappings create --service image-to-csv --domain your-domain.com --region us-central1
```
Then add the DNS records `gcloud` prints out at your registrar.

## Alternatives (if you'd rather not use GCP)

| Platform | Cost model | Cold start | Fit |
|---|---|---|---|
| **Cloud Run** (recommended) | Pay-per-use, generous perpetual free tier | Fast (small image) | Best — see above |
| Fly.io | Pay-per-use, small free allowance | Fast if kept warm | Good, similar model, smaller free tier |
| Render | Free web service sleeps after inactivity | Slow (~30s+) on free tier | OK for a demo, not for real usage |
| Railway | No real free tier ($5 min/mo usage-based) | Fast | Fine, but costs from day one |

All of these build the same `Dockerfile` unchanged and inject `PORT` the
same way Cloud Run does — no code changes needed to switch.

**Fly.io:**
```bash
fly launch --dockerfile Dockerfile --no-deploy   # answer prompts, skip deploy
fly secrets set GEMINI_API_KEY=YOUR_KEY GEMINI_MODEL=gemini-3.1-flash-lite
fly deploy
```

**Render:** connect this directory as a repo in the Render dashboard, choose
"Docker" as the environment, and add `GEMINI_API_KEY` / `GEMINI_MODEL` under
Environment → Environment Variables. No CLI needed.

**Railway:**
```bash
railway init
railway variables set GEMINI_API_KEY=YOUR_KEY GEMINI_MODEL=gemini-3.1-flash-lite
railway up
```
