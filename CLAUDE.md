# CLAUDE.md

Guidance for AI assistants working in this repository.

## What this project is

**Prospection Intake Bot** — a field-sales prospecting capture pipeline for RAS Intérim.
Commercial reps ("commerciaux") send business cards, storefront photos, GPS and voice notes
to a **Telegram bot**. The bot OCRs and enriches each visit into a structured company record
(SIRET, NAF, address, phone, effectif, etc.), stores it, and each day a scheduled job exports
those records to formatted Excel files and emails them to reps, agency managers, and the admin.

Two independent components make up the system:

1. **`worker.js`** — a **Cloudflare Worker** (the Telegram bot + enrichment engine + storage).
   This is the runtime that reps interact with. It runs 24/7 on Cloudflare, stores data in KV.
2. **`scripts/` + `.github/workflows/`** — **Python exporters** run on a cron by GitHub Actions.
   They pull the day's records from the worker's `/dump` endpoint and email Excel/photo bundles.

The two talk over HTTP only: Actions → `GET {WORKER_BASE_URL}/dump` (auth via `EXPORT_TOKEN`).
User-facing text, comments, commit messages, and agency labels are all in **French** — match that.

## Repository layout

```
worker.js                      Cloudflare Worker — the entire Telegram bot (~5900 lines, single file)
wrangler.toml                  Worker deploy config (KV binding, secrets list — deploy with `npx wrangler deploy`)

scripts/
  export_and_mail.py           MAIN exporter entrypoint (called by both GitHub workflows)
  excel_export.py              Excel/CRM workbook builder (imported by export_and_mail.py)
  scripts/export_and_mail.py   STALE nested duplicate — NOT used by any workflow; ignore / do not edit

.github/workflows/
  prospection-daily.yml        Cron: agency-manager recap @17:47 FR, admin recap @17:49 FR (Mon–Fri)
  prospection-close.yml        Manual (workflow_dispatch): one rep's individual export

requirements.txt               Python deps for the exporters (openpyxl, pillow, pytesseract, easyocr, ...)
config.yml                     Legacy per-agency email routing (superseded by MAIL_ROUTING_JSON secret)
data/cumul_2026-*.json         Monthly rollup stats snapshots (per-agency prospects/clients/commandes)

consolidate.py                 Standalone/alternate: rebuild an Excel from GitHub *issues* labelled `intake`.
                               NOT wired into the live pipeline; kept as a fallback path.
moulinette.yml                 Orphan workflow file (references scripts/notify_and_export.py, which does
                               NOT exist). Not under .github/workflows, so it never runs. Treat as dead.

Table.2.xlsx, users.json,      Legacy artifacts / empty placeholders. Not part of the live flow.
config, out, reports
```

`out/` and `reports/` are git-ignored generated-output dirs (the export writes to `/tmp/prospection_out`
in CI). `__pycache__`, `.venv`, `node_modules`, `dist` are git-ignored.

## The Telegram bot (`worker.js`)

One large self-contained ES-module Worker. No build step, no bundler — it is deployed as-is.
It is organized top-to-bottom into banner-delimited sections (`// ===...`). Key regions:

- **Entry / routing** — `export default { fetch }` → `handle(req, event)` dispatches by `url.pathname`.
- **HTTP endpoints:**
  - `GET  /health`, `GET /status` — liveness / config introspection
  - `GET  /dump?date=YYYY-MM-DD&kind=prospects|closes|photos|cards` — NDJSON export (auth: `X-Export-Token`)
  - `POST /telegram` — the Telegram webhook (auth: secret header)
  - `POST /remind`, `POST /update` — admin/OCR-callback endpoints (Bearer / `X-Export-Token`)
  - `GET  /test-brave`, `/test-scrape`, `/test-gouv` — manual API probes
- **Security:** `timingSafeEqual` for secret comparison, `checkExportToken` / `checkTelegramSecretAsync`,
  `isSafeUrl` SSRF guard on all outbound scraping.
- **Storage (Cloudflare KV, binding `PROSPECTION_KV` via `kv()`):** sessions (7d TTL), prefs (30d),
  KV indexes `idx:day:{date}:{agency}` and `idx:dup:{date}:{agency}` for O(1) dedup, prospects (90d TTL).
- **Session / visit model:** `getSession`/`setSession`, multi-visit index per chat, `emptyDraft`,
  `sanitizeProspectDraft`, auto-save after full enrichment.
- **OCR (hybrid):** Gemini Vision (`geminiVisionJsonFromImageBytes`) + Google Vision, fused; voice notes
  via `geminiTranscribeAudio`.
- **Enrichment:** `enrichCompany` / `enrichDraftFromCurrentData` — API Gouv (recherche-entreprises),
  Google Places, Brave Search, website scraping, and BAN (`banNormalizeAddress`) reverse-geocode.
- **Telegram UX:** inline keyboards, `force_reply`, spinner messages, pinned active form, quick commands
  `/n` (new visit), `/s` (save), `/l` (list), `/g` (GPS).
- **Handlers:** `handleMessage` (text/photo/voice/location) and `handleCallback` (button taps).
- **Background work:** `scheduleBackgroundTask` (via `event.waitUntil`) dispatches GitHub Actions
  (`dispatchGitHubWithRetry`) for heavy OCR without blocking the webhook response.

### Worker secrets (set via `npx wrangler secret put <NAME>`, never commit)

`TELEGRAM_TOKEN`, `TELEGRAM_WEBHOOK_SECRET`, `EXPORT_TOKEN`, `GEMINI_API_KEY`,
`GOOGLE_VISION_API_KEY`, `GOOGLE_PLACES_API_KEY`, `BRAVE_SEARCH_KEY`, `GITHUB_TOKEN`,
`GITHUB_OWNER`, `GITHUB_REPO`, `GITHUB_WORKFLOW_FILE`, `BREVO_API_KEY`, `BREVO_SENDER_EMAIL`,
`MANAGER_CHAT_ID`. The KV namespace `id` in `wrangler.toml` must be filled in before deploy.

## The exporters (`scripts/`)

`export_and_mail.py` is the entrypoint invoked by both workflows. Flow:

1. `fetch_prospects(date)` → `GET {WORKER_BASE_URL}/dump` (NDJSON), authed with `EXPORT_TOKEN`.
2. `normalize_prospect` / `filter_rows` — flatten the `draft`, dedup by `visit_id` then `SIRET|name`.
3. `excel_export.py` builds two workbooks per recipient set:
   - `.xlsx` — styled consultation copy (agency colors, embedded card-photo thumbnails, score fills)
   - `.xls` — flat CRM-import copy (`CRM_HEADERS` / `FIELD_MAP` define the column contract)
     Both from `export_commercial` / `export_agency_manager` / `export_admin`.
4. `build_photos_zip` — downloads card + façade images fresh from Telegram by `file_id` into a ZIP.
5. `send_brevo` — emails attachments via the Brevo API.

**Three modes**, chosen by the `SEND_MODE` env var:
- `individual` — one rep's export → commercial + manager + admin (xlsx + xls + zip each). Needs `AGENCY` + `INITIALS`.
- `agency_manager` — per-agency daily recap → each manager (xls + zip). Loops `GR, VR, GRS, SLS`.
- `admin` — all-agencies recap → admin `SL` (xlsx + zip).

Agencies are `GR, VR, GRS, SLS`. Email routing comes from the `MAIL_ROUTING_JSON` secret, with a
hardcoded `DEFAULT_ROUTING` fallback baked into `export_and_mail.py`.

### Exporter env vars (provided by the workflows from GitHub secrets)

`WORKER_BASE_URL`, `EXPORT_TOKEN`, `TELEGRAM_TOKEN`, `BREVO_API_KEY`, `BREVO_SENDER_EMAIL`,
`BREVO_SENDER_NAME`, `MAIL_ROUTING_JSON`, `GEMINI_API_KEY`, `GOOGLE_PLACES_API_KEY`,
`SEND_MODE`, `RUN_DATE` (Europe/Paris date; empty = today), and `AGENCY`/`INITIALS` for individual mode.

## Development workflows

There is no test suite, linter config, or package.json in this repo. Verify changes manually.

**Worker (`worker.js`):**
```bash
npx wrangler deploy            # deploy to Cloudflare (name: prospection-intake-worker)
npx wrangler tail              # live logs
```
The single-file structure is intentional; keep the banner-section organization and the existing
helper conventions (`safeFetchJson`/`safeFetchText` for all outbound HTTP, `withCache` for cached
API calls, `kv()` for storage, `log()` for structured logging).

**Exporters (`scripts/`):**
```bash
pip install -r requirements.txt
# Run locally by exporting the env vars above, then:
python scripts/export_and_mail.py
```
In CI the workflows also `apt-get install tesseract-ocr tesseract-ocr-fra tesseract-ocr-eng` and
`pip install -r requirements.txt` before running. To trigger manually: run **Prospection Daily Recaps**
(pick `agency_manager`/`admin` + optional date) or **Prospection Close Individual** (agency + initials
+ date) from the GitHub Actions tab.

## Conventions & gotchas

- **Language:** all user-facing strings, comments, and commit messages are French. Preserve this.
- **Versioning:** the codebase self-labels **v8.4** in banners/strings (`Prospection Bot v8.4`).
  When you change behavior, keep version references consistent.
- **Column contract:** `CRM_HEADERS`, `FIELD_MAP`, and the worker's draft field names must stay in
  sync — the `.xls` is imported directly into a CRM, so renaming/reordering columns is breaking.
- **Secrets never live in the repo.** They come from Cloudflare secrets (worker) and GitHub Actions
  secrets (exporters). `wrangler.toml` and workflow files only *list* their names.
- **Ignore the stale duplicates:** `scripts/scripts/export_and_mail.py`, `moulinette.yml`, and
  `consolidate.py` are not part of the live pipeline. Edit `scripts/export_and_mail.py` for exporter
  changes and `.github/workflows/*.yml` for scheduling changes.
- **Dedup is layered:** the worker dedups on write (KV `idx:dup:*`), the exporter dedups again on read
  (`visit_id`, then `SIRET|name`). Keep both when touching identity logic.

## Git workflow

- Work on the designated feature branch; commit with clear, descriptive (French-friendly) messages.
- Push with `git push -u origin <branch>`. Do not open a PR unless explicitly asked.
- Do not commit generated output (`out/`, `reports/`) or any secret value.
