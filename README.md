# Pharma RFP / Grant Monitor (Windows)

Checks the company URLs listed in the Excel workbook and records page-level funding signals. It does not discover linked pages or PDFs and does not verify individual opportunities.

## What it does

- Reads the hyperlink target in each Excel `URL` cell.
- Fetches each listed page with `requests` and BeautifulSoup, with an optional Playwright fallback for JavaScript-heavy pages.
- Applies keyword rules to the page text to identify grant/RFP-related signals and changes since the previous run.
- Sends an optional email when the existing change-detection rules identify a positive change. Treat alerts as leads to review on the source page.
- Writes reports with a stable set of columns. Fields requiring per-opportunity extraction (such as validated dates, eligibility, and application URL) remain blank in this simple mode.

## Windows setup

1. Put the Excel workbook in the same folder as `main.py` and set its name in `.env` (`INPUT_XLSX`).
2. Install Python 3.11 or 3.12, then open PowerShell in this folder and run:

   ```powershell
   py -3.11 -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   python -m playwright install chromium
   Copy-Item .env.example .env
   ```

3. Edit `.env` with your SMTP settings and comma- or semicolon-separated `ALERT_TO` recipients. Use an app password if your mail provider requires one. Do not commit or share `.env`.
4. Run a small check, then a normal run:

   ```powershell
   python main.py --max-urls 2
   python main.py
   ```

5. Use `python main.py --baseline` to store the current state without sending initial change alerts.

## Report columns

Both CSV files use these columns in order:

`company`, `source_url`, `final_url`, `discovered_from`, `http_status`, `extraction_method`, `page_title`, `opportunity_type`, `opportunity_title`, `opportunity_status`, `status_reason`, `opportunity_id`, `published_date`, `published_date_text`, `deadline_text`, `deadline_date`, `eligibility`, `application_url`, `evidence`, `funding_terms`, `relevance_score`, `change_type`, `changed_fields`, `flag`, `checked_at`, `error`.

The tool checks only URLs from the workbook. It does not crawl for additional RFP pages or PDFs, extract or validate actual deadlines, reliably establish whether a specific opportunity is open, or distinguish a new RFP from an older RFP mentioned on a page. The corresponding columns are left blank or contain simple page-level heuristic values.

## Output files

- `output/last_run_report.csv` — results for the current run.
- `output/rfp_grant_matches.csv` — current page-level matches.
- `output/monitor_state.json` — saved fingerprints used by change detection.
- `logs/monitor.log` — run log.

If Google Sheets sync is configured, the cumulative `rfp_grant_matches.csv` contents (including its header row) replace the configured report worksheet on each run. The `Email Alerts` worksheet is replaced on each run with only alerts successfully sent for that scrape. Email alert tabs and messages include company, RFP URL, evidence, RFP/CGA/IME keywords, RFP title, and detected specialty (from the monitor's configured topic terms; `Not identified` if none match). `Email Alerts History` retains sent alerts for a rolling 30 days and adds an `Alert Sent At (UTC)` timestamp; older rows are removed on each run. Older timestamped rows from the previous combined tab are migrated into the history tab, and undated legacy rows are archived separately. Set `GOOGLE_SHEETS_SPREADSHEET_ID`, `GOOGLE_SHEETS_WORKSHEET`, `GOOGLE_SHEETS_EMAIL_WORKSHEET`, `GOOGLE_SHEETS_EMAIL_HISTORY_WORKSHEET`, and `GOOGLE_SERVICE_ACCOUNT_JSON` in `.env`, share the spreadsheet with the service account email from the JSON credentials, and install dependencies with `pip install -r requirements.txt`. A sync error is logged and does not stop local report or email processing.

The relevance score and status signals are heuristic indicators, not confirmation that an RFP or grant is open. Check the cited source page before acting.

## Run automatically on weekdays at 8 PM IST

After a baseline run and a successful email check, create a Weekly task in Windows Task Scheduler for Monday through Friday at 8:00 PM. Set the program to this project's `.venv\Scripts\python.exe`, arguments to `main.py`, and **Start in** to the project folder. The computer must be powered on and connected to the internet. Check `logs\monitor.log` and the output CSVs after the first scheduled run.

## Limitations

- Only URLs already present in the workbook are checked; unlinked opportunity pages and PDFs are not found.
- Keyword matching can miss unfamiliar phrasing or flag generic funding language. Page-level matches may not represent a newly announced opportunity.
- Sites requiring login, blocking automated access, or relying on unsupported JavaScript layouts may not be read successfully. Scanned PDFs are not processed in this simple mode.
