from __future__ import annotations
import argparse
import csv
import hashlib
import html
import json
import logging
import os
import re
import smtplib
import time
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urlparse
import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openpyxl import load_workbook
from io import BytesIO
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sync_playwright = None

ROOT = Path(__file__).resolve().parent
load_dotenv(ROOT / ".env")
UTC = timezone.utc
INPUT_XLSX = Path(os.getenv("INPUT_XLSX", str(ROOT / "RFP Scrape Pharma URLs.xlsx")))
OUTPUT_DIR = ROOT / "output"
LOG_DIR = ROOT / "logs"
STATE_FILE = OUTPUT_DIR / "monitor_state.json"
MATCHES_FILE = OUTPUT_DIR / "rfp_grant_matches.csv"
RUN_REPORT = OUTPUT_DIR / "last_run_report.csv"
REPORT_COLUMNS = [
    "company", "source_url", "final_url", "discovered_from", "http_status",
    "extraction_method", "page_title", "opportunity_type", "opportunity_title",
    "opportunity_status", "status_reason", "opportunity_id", "published_date",
    "published_date_text", "deadline_text", "deadline_date", "eligibility",
    "application_url", "evidence", "funding_terms", "relevance_score", "change_type",
    "changed_fields", "flag", "checked_at", "error",
]

# Tune these lists to reduce or increase alert sensitivity.
FUNDING_TERMS = [
    "request for proposal", "requests for proposals", "rfp", "call for proposals",
    "independent medical education", 
    "call for grant applications", "CGA",
    "funding opportunity", "apply for funding", "submit a proposal", "IME", "external funding",
    #"#investigator-sponsored", "#investigator initiated", 
    #"grant application", "medical education grant", "educational grant",
    #"research grant", "#sponsorship application",
    #"call for grants", "calls for grants", "grant opportunity", "grant opportunities",
]
TOPIC_TERMS = [
    "oncology", "cancer", "tumor", "tumour", "hematology", "haematology",
    "neuroscience", "neurology", "neurologic", "neurological", "parkinson",
    "migraine", "epilepsy", "multiple sclerosis", "alzheimer", "dementia",
    "brain", "neurodegenerative", "immuno-oncology", "precision oncology",
    "breast cancer", "lung cancer", "colorectal cancer", "prostate cancer",
    "pancreatic cancer", "ovarian cancer", "lymphoma", "leukemia", "leukaemia",
    "multiple myeloma", "solid tumor", "solid tumour", "tumor microenvironment",
    "car t", "cell therapy", "radiotherapy", "stroke", "cerebrovascular",
    "movement disorder", "alzheimer's disease", "seizure", "headache",
    "neuromuscular", "neuroimmunology", "motor neuron disease", "amyotrophic lateral sclerosis",
    "spinal muscular atrophy", "huntington's disease", "neuropathy", "cognitive impairment",
]
OPEN_TERMS = [
    "now accepting", "accepting applications", "currently accepting",
    "application deadline", "due date", "submit by", "deadline",
    "open for applications", "applications are open", "proposal deadline",
]
CLOSED_TERMS = [
    "closed", "expired", "deadline has passed", "no longer accepting",
    "applications are no longer being accepted",
]
RFP_TERMS = ["request for proposal", "requests for proposals", "rfp", "call for proposals", "calls for proposals"]
EMAIL_CATEGORY_PATTERNS = {
    "RFP": re.compile(r"\b(?:RFP|requests? for proposals?|calls? for proposals?)\b", re.I),
    "CGA": re.compile(r"\bCGA\b|\bcall for grant applications?\b", re.I),
    "IME": re.compile(r"\bIME\b|\bindependent medical education\b", re.I),
}
MAX_TEXT_CHARS = int(os.getenv("MAX_TEXT_CHARS", "1500000"))
REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "25"))
DELAY_SECONDS = float(os.getenv("DELAY_SECONDS", "1.0"))
USE_PLAYWRIGHT_FALLBACK = os.getenv("USE_PLAYWRIGHT_FALLBACK", "true").lower() in {"1", "true", "yes"}
USER_AGENT = os.getenv("USER_AGENT", "PharmaRFPMonitor/1.0 (polite public-page monitoring)")

def setup_logging():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(LOG_DIR / "monitor.log", encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )

def normalize_url(url):
    if not isinstance(url, str):
        return ""
    url = url.strip()
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    if not re.match(r"^https?://", url, re.I):
        return ""
    return url

def load_input(path: Path):
    if not path.exists():
        raise FileNotFoundError(
            f"Input workbook not found: {path}\n"
            "Place the workbook beside main.py or set INPUT_XLSX in .env."
        )
    # Read actual cell hyperlinks, because the uploaded workbook's URL cells display
    # friendly labels while the real URLs are stored as hyperlink targets.
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    header_map = {h.lower(): i for i, h in enumerate(headers)}
    company_idx = header_map.get("company")
    url_idx = header_map.get("url")
    desc_idx = header_map.get("description")
    if company_idx is None or url_idx is None:
        raise ValueError("Workbook must contain Company and URL columns.")

    records = []
    for row_idx in range(2, ws.max_row + 1):
        company_cell = ws.cell(row_idx, company_idx + 1)
        url_cell = ws.cell(row_idx, url_idx + 1)
        desc_cell = ws.cell(row_idx, desc_idx + 1) if desc_idx is not None else None
        company = str(company_cell.value or "").strip()
        # Reload formula workbook cell for hyperlink target; data_only workbook
        # can preserve hyperlink in most files, but read the target directly.
        url = normalize_url(url_cell.hyperlink.target if url_cell.hyperlink else url_cell.value)
        description = str(desc_cell.value or "").strip() if desc_cell else ""
        if company and url:
            records.append({"company": company, "url": url, "description": description})
        elif company or url:
            logging.warning("Skipping row %s: missing company or usable URL", row_idx)
    # Deduplicate by company+URL
    unique, seen = [], set()
    for r in records:
        key = (r["company"].casefold(), r["url"].rstrip("/").casefold())
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return unique

def make_session():
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept": "text/html,application/xhtml+xml"})
    retry = Retry(total=2, connect=2, read=2, backoff_factor=0.7,
                  status_forcelist=[429, 500, 502, 503, 504],
                  allowed_methods=["GET"], raise_on_status=False)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

def looks_like_pdf(url, content_type, content):
    return (
        "application/pdf" in (content_type or "").lower()
        or urlparse(url).path.lower().endswith(".pdf")
        or (content or b"").lstrip().startswith(b"%PDF-")
    )

def extract_pdf_text(pdf_bytes):
    """
    Extract text from both ordinary PDFs and many scanned PDFs.
    PyMuPDF is the primary extractor. If the PDF has no usable text,
    optional OCR can be attempted with pytesseract + a local Tesseract install.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:
        raise RuntimeError("PDF support requires PyMuPDF. Install with: pip install pymupdf") from exc

    pages_text = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page_number, page in enumerate(doc, start=1):
            try:
                page_text = page.get_text("text", sort=True) or ""
            except Exception:
                page_text = ""
            if page_text.strip():
                pages_text.append(f"\n[PDF page {page_number}]\n{page_text.strip()}")

    text = "\n".join(pages_text).strip()
    if len(text) >= 80:
        return text[:MAX_TEXT_CHARS], "pymupdf"

    # OCR is optional because it requires a separate Tesseract installation.
    try:
        import pytesseract
        from PIL import Image
        import fitz
        ocr_pages = []
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page_number, page in enumerate(doc, start=1):
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image = Image.open(BytesIO(pix.tobytes("png")))
                page_text = pytesseract.image_to_string(image) or ""
                if page_text.strip():
                    ocr_pages.append(f"\n[PDF page {page_number} - OCR]\n{page_text.strip()}")
        ocr_text = "\n".join(ocr_pages).strip()
        if len(ocr_text) > len(text):
            return ocr_text[:MAX_TEXT_CHARS], "ocr"
    except Exception:
        pass

    if text:
        return text[:MAX_TEXT_CHARS], "pymupdf_low_text"
    return "", "no_extractable_text; scanned PDF may require OCR"

def extract_with_requests(session, url):
    response = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
    response.raise_for_status()
    content_type = response.headers.get("Content-Type", "").lower()
    if looks_like_pdf(response.url, content_type, response.content):
        pdf_text, pdf_method = extract_pdf_text(response.content)
        pdf_note = "" if pdf_text.strip() else "No text extracted from PDF; scanned PDF may need OCR or manual review"
        return response.url, Path(urlparse(response.url).path).name or "PDF document", pdf_text, response.status_code, pdf_note, pdf_method
    soup = BeautifulSoup(response.text, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "nav", "footer", "header"]):
        tag.decompose()
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))[:MAX_TEXT_CHARS]
    return response.url, title, text, response.status_code, "", "requests"

def extract_with_playwright(url):
    if sync_playwright is None:
        raise RuntimeError("Playwright is not installed.")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=USER_AGENT)
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(1800)
            final_url = page.url
            title = page.title()
            text = page.locator("body").inner_text(timeout=10000)
            return final_url, title, re.sub(r"\s+", " ", text)[:MAX_TEXT_CHARS], 200, "", "playwright"
        finally:
            browser.close()

def find_terms(text, terms):
    low = text.casefold()
    return [term for term in terms if term.casefold() in low]

def classify(company, url, description, text):
    combined = " ".join([company, url, description, text])
    funding = find_terms(combined, FUNDING_TERMS)
    topics = find_terms(combined, TOPIC_TERMS)
    open_signals = find_terms(combined, OPEN_TERMS)
    closed_signals = find_terms(combined, CLOSED_TERMS)
    # Only flag pages where funding-related wording is present. Topic matches
    # are separately recorded so reviewers can prioritize relevant opportunities.
    return {
        "funding_terms": funding,
        "topic_terms": topics,
        "open_signals": open_signals,
        "closed_signals": closed_signals,
        "flag": bool(funding),
        "priority": "HIGH" if funding and topics and open_signals and not closed_signals
                    else "MEDIUM" if funding and topics
                    else "REVIEW" if funding else "NONE",
    }

def opportunity_evidence(text):
    """Return concise funding-related sentences used for positive-change detection."""
    sentences = re.split(r"(?<=[.!?])\s+|\s*[•▪\n]+\s*", text)
    evidence = []
    seen = set()
    for sentence in sentences:
        clean = re.sub(r"\s+", " ", sentence).strip(" -:;•")
        low = clean.casefold()
        has_funding = any(term.casefold() in low for term in FUNDING_TERMS)
        has_positive = any(term.casefold() in low for term in OPEN_TERMS)
        has_rfp = any(term.casefold() in low for term in RFP_TERMS)
        has_opportunity = any(term in low for term in ("opportunity", "accepting", "applications are open", "apply now", "submit by", "deadline"))
        if clean and has_funding and (has_positive or (has_rfp and has_opportunity)):
            key = low
            if key not in seen:
                seen.add(key)
                evidence.append(clean[:700])
    return evidence[:100]

def email_categories(text):
    """Return only the supported RFP, CGA, and IME categories for email content."""
    return [category for category, pattern in EMAIL_CATEGORY_PATTERNS.items() if pattern.search(text or "")]

def email_evidence_snippets(evidence_items):
    """Split page evidence into shorter excerpts and retain only RFP/CGA/IME text."""
    snippets = []
    seen = set()
    for item in evidence_items:
        # The source page is flattened before extraction, so split on common
        # sentence, list, and table separators to avoid emailing whole sections.
        parts = re.split(r"(?<=[.!?])\s+|\s*[;|\u2022\u25aa]\s*|\s+(?=\d{4}\b)", item or "")
        for part in parts:
            clean = re.sub(r"\s+", " ", part).strip(" -:;\u2022\u25aa")
            key = clean.casefold()
            if clean and email_categories(clean) and key not in seen:
                seen.add(key)
                snippets.append(clean[:500])
    return snippets

EMAIL_ALERT_COLUMNS = [
    "Pharma Company Name", "RFP URL", "Evidence Text", "Keywords", "Title of the RFP", "Specialty",
]

def build_email_alert_rows(alerts):
    """Build the alert records shown in emails and the alert spreadsheet tabs."""
    rows = []
    ordered_alerts = sorted(
        alerts,
        key=lambda alert: (0 if alert["opportunity_type"] == "RFP" else 1, alert["company"].casefold()),
    )
    for alert in ordered_alerts:
        evidence = email_evidence_snippets(alert.get("new_evidence", []))
        if not evidence:
            continue
        keywords = list(dict.fromkeys(category for item in evidence for category in email_categories(item)))
        rows.append([
            alert["company"],
            alert["final_url"],
            "; ".join(evidence),
            "; ".join(keywords),
            alert["opportunity_title"],
            "; ".join(dict.fromkeys(alert.get("topic_terms", []))) or "Not identified",
        ])
    return rows

def get_state():
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            logging.warning("State file is invalid; starting with empty state.")
    return {}

def save_state(state):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")

def sync_google_sheet(csv_path):
    """Replace the configured worksheet contents with the cumulative matches CSV."""
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
    worksheet_name = os.getenv("GOOGLE_SHEETS_WORKSHEET", "").strip()
    credentials_setting = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not (spreadsheet_id and worksheet_name and credentials_setting):
        logging.info("Google Sheets sync is not configured; skipping spreadsheet update.")
        return False
    # Accept either the spreadsheet ID or its full Google Sheets URL in .env.
    spreadsheet_match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", spreadsheet_id)
    spreadsheet_key = spreadsheet_match.group(1) if spreadsheet_match else spreadsheet_id

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise RuntimeError("Google Sheets sync requires gspread and google-auth; install requirements.txt.") from exc

    credentials_path = Path(credentials_setting)
    if not credentials_path.is_absolute():
        credentials_path = ROOT / credentials_path
    credentials = Credentials.from_service_account_file(
        str(credentials_path),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as csv_file:
        values = list(csv.reader(csv_file))
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_key)
    logging.info(
        "Google Sheets target: spreadsheet='%s' | URL=%s | configured worksheet='%s' | CSV=%s",
        spreadsheet.title, spreadsheet.url, worksheet_name, csv_path.resolve(),
    )
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        row_count = max(100, len(values))
        column_count = max(1, max((len(row) for row in values), default=1))
        worksheet = spreadsheet.add_worksheet(
            title=worksheet_name,
            rows=str(row_count),
            cols=str(column_count),
        )
        logging.info("Created missing Google Sheets worksheet '%s'.", worksheet_name)

    # Replace the tab contents with the CSV, then read the tab back to verify.
    worksheet.clear()
    if values:
        worksheet.update(range_name="A1", values=values)
    expected_rows = len(values)
    expected_cols = max((len(row) for row in values), default=0)
    written_values = worksheet.get_all_values()
    written_rows = len(written_values)
    logging.info(
        "Google Sheets write verification: tab_url=%s | expected_rows_including_header=%d | readback_rows=%d | expected_columns=%d",
        worksheet.url, expected_rows, written_rows, expected_cols,
    )
    # Sheets omits trailing blank cells when reading rows back, so normalize row
    # widths before comparing the data to avoid false mismatches.
    def pad_rows(rows):
        return [list(row) + [""] * max(0, expected_cols - len(row)) for row in rows]

    if pad_rows(written_values) != pad_rows(values):
        logging.warning(
            "Google Sheets read-back differs from CSV. CSV header=%r; sheet header=%r",
            values[0] if values else [], written_values[0] if written_values else [],
        )
        raise RuntimeError("Google Sheets write completed but read-back did not match the CSV.")
    logging.info(
        "Google Sheet updated and verified: spreadsheet='%s' | worksheet='%s' | %d data row(s).",
        spreadsheet.title, worksheet_name, max(0, len(values) - 1),
    )
    return True

def sync_email_alerts_to_google_sheet(alerts):
    """Replace the current-alert tab and maintain a separate 30-day history tab."""
    current_rows = build_email_alert_rows(alerts)
    sent_at = datetime.now(UTC).isoformat(timespec="seconds")
    history_headers = EMAIL_ALERT_COLUMNS + ["Alert Sent At (UTC)"]
    old_email_headers = EMAIL_ALERT_COLUMNS[:-1]
    old_url_email_headers = ["Pharma Company Name", "URL", "Evidence Text", "Keywords", "Title of the RFP"]
    old_history_headers = [old_email_headers + ["Alert Sent At (IST)"], old_url_email_headers + ["Alert Sent At (IST)"],
                           old_email_headers + ["Alert Sent At (UTC)"], old_url_email_headers + ["Alert Sent At (UTC)"]]
    new_history_rows = [row + [sent_at] for row in current_rows]

    spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
    credentials_setting = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    current_name = os.getenv("GOOGLE_SHEETS_EMAIL_WORKSHEET", "Email Alerts").strip() or "Email Alerts"
    history_name = os.getenv("GOOGLE_SHEETS_EMAIL_HISTORY_WORKSHEET", "Email Alerts History").strip() or "Email Alerts History"
    if not (spreadsheet_id and credentials_setting):
        logging.info("Google Sheets email-alert sync is not configured; skipping.")
        return False
    if current_name == history_name:
        raise ValueError("Current email alerts and email history must use different worksheet names.")
    spreadsheet_match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", spreadsheet_id)
    spreadsheet_key = spreadsheet_match.group(1) if spreadsheet_match else spreadsheet_id

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise RuntimeError("Google Sheets sync requires gspread and google-auth; install requirements.txt.") from exc

    credentials_path = Path(credentials_setting)
    if not credentials_path.is_absolute():
        credentials_path = ROOT / credentials_path
    credentials = Credentials.from_service_account_file(
        str(credentials_path), scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    spreadsheet = gspread.authorize(credentials).open_by_key(spreadsheet_key)

    def get_or_create_tab(name, columns):
        try:
            return spreadsheet.worksheet(name)
        except gspread.WorksheetNotFound:
            tab = spreadsheet.add_worksheet(title=name, rows="1000", cols=str(columns))
            logging.info("Created Google Sheets worksheet '%s'.", name)
            return tab

    current_tab = get_or_create_tab(current_name, len(EMAIL_ALERT_COLUMNS))
    history_tab = get_or_create_tab(history_name, len(history_headers))
    current_values = current_tab.get_all_values()
    history_values = history_tab.get_all_values()

    def records_for_tab(values, name):
        values = [row for row in values if any(str(cell).strip() for cell in row)]
        if not values:
            return [], []
        if values[0] == history_headers:
            rows = [row[:len(history_headers)] + [""] * max(0, len(history_headers) - len(row)) for row in values[1:]]
            return [normalize_history_timestamp(row) for row in rows], []
        if values[0] in old_history_headers:
            rows = []
            for old_row in values[1:]:
                old_row = old_row + [""] * max(0, 6 - len(old_row))
                row = old_row[:5] + ["Not identified"] + [old_row[5]]
                rows.append(normalize_history_timestamp(row))
            return rows, []
        if values[0] == EMAIL_ALERT_COLUMNS:
            return [], [row[:len(EMAIL_ALERT_COLUMNS)] + [""] * max(0, len(EMAIL_ALERT_COLUMNS) - len(row)) for row in values[1:]]
        if values[0] in (old_email_headers, old_url_email_headers):
            return [], [row[:5] + ["Not identified"] for row in values[1:]]
        raise RuntimeError(f"Worksheet '{name}' has unexpected headers; leaving it unchanged.")

    def normalize_history_timestamp(row):
        """Convert stored timestamp instants to a consistent UTC ISO string."""
        try:
            timestamp = datetime.fromisoformat(row[-1])
            if timestamp.tzinfo is None:
                timestamp = timestamp.replace(tzinfo=UTC)
            row[-1] = timestamp.astimezone(UTC).isoformat(timespec="seconds")
        except ValueError:
            pass  # The retention pass below will discard malformed timestamps.
        return row

    history_rows, legacy_history_rows = records_for_tab(history_values, history_name)
    current_history_rows, legacy_current_rows = records_for_tab(current_values, current_name)
    known_timestamped_payloads = {
        tuple(row[:len(EMAIL_ALERT_COLUMNS)]) for row in history_rows + current_history_rows
    }
    legacy_rows = legacy_history_rows + [
        row for row in legacy_current_rows
        if tuple(row) not in known_timestamped_payloads
    ]

    # Preserve pre-timestamp data in a clearly labelled archive before migrating tabs.
    if legacy_rows:
        archive_headers = EMAIL_ALERT_COLUMNS + ["Timestamp Status"]
        archive = None
        archive_values = []
        for suffix in range(1, 100):
            archive_name = "Email Alerts Legacy Archive" if suffix == 1 else f"Email Alerts Legacy Archive {suffix}"
            try:
                candidate = spreadsheet.worksheet(archive_name)
                candidate_values = candidate.get_all_values()
            except gspread.WorksheetNotFound:
                candidate = spreadsheet.add_worksheet(title=archive_name, rows="1000", cols=str(len(archive_headers)))
                candidate_values = []
            if not candidate_values or candidate_values[0] == archive_headers:
                archive, archive_values = candidate, candidate_values
                break
            logging.warning("Leaving archive worksheet '%s' unchanged because its headers differ.", archive_name)
        if archive is None:
            raise RuntimeError("Could not find or create a compatible legacy email-alert archive worksheet.")
        if not archive_values:
            archive.update(range_name="A1", values=[archive_headers], value_input_option="RAW")
            archive.format("A1:G1", {"textFormat": {"bold": True}})
            archive_values = [archive_headers]
        known_legacy = {tuple(row[:len(EMAIL_ALERT_COLUMNS)]) for row in archive_values[1:]}
        archive_rows = [
            row + ["Timestamp unavailable (pre-change)"]
            for row in legacy_rows
            if tuple(row) not in known_legacy
        ]
        if archive_rows:
            archive.append_rows(archive_rows, value_input_option="RAW")
        logging.info("Archived %d pre-timestamp alert row(s).", len(archive_rows))

    # Existing current-tab history is migrated to the history tab. Identical
    # rows are deduplicated because it may already have been copied on a prior run.
    candidates = history_rows + current_history_rows + new_history_rows
    cutoff = datetime.now(UTC) - timedelta(days=30)
    retained_rows, expired_rows, seen = [], 0, set()
    for row in candidates:
        key = tuple(row)
        if key in seen:
            continue
        seen.add(key)
        try:
            timestamp = datetime.fromisoformat(row[-1])
            if timestamp.tzinfo is None:
                timestamp = timestamp.replace(tzinfo=UTC)
        except ValueError:
            expired_rows += 1
            continue
        if timestamp.astimezone(UTC) >= cutoff:
            retained_rows.append(row)
        else:
            expired_rows += 1

    updated_history = [history_headers] + retained_rows
    history_tab.clear()
    history_tab.update(range_name="A1", values=updated_history, value_input_option="RAW")
    history_tab.format("A1:G1", {"textFormat": {"bold": True}})
    history_tab.freeze(rows=1)
    if history_tab.get_all_values() != updated_history:
        raise RuntimeError(f"Google Sheets read-back did not match '{history_name}'.")

    # This tab represents only emails successfully sent in the current scrape.
    updated_current = [EMAIL_ALERT_COLUMNS] + (current_rows if alerts else [])
    current_tab.clear()
    current_tab.update(range_name="A1", values=updated_current, value_input_option="RAW")
    current_tab.format("A1:F1", {"textFormat": {"bold": True}})
    current_tab.freeze(rows=1)
    if current_tab.get_all_values() != updated_current:
        raise RuntimeError(f"Google Sheets read-back did not match current alerts in '{current_name}'.")

    logging.info(
        "Google Sheets email tabs updated: current='%s' (%d alert(s)); history='%s' (%d retained, %d expired); window=30 days.",
        current_name, len(current_rows) if alerts else 0, history_name, len(retained_rows), expired_rows,
    )
    return True

def send_email_alert(alerts):
    host = os.getenv("SMTP_HOST", "").strip()
    recipients = [r.strip() for r in re.split(r"[,;]", os.getenv("ALERT_TO", "")) if r.strip()]
    sender = os.getenv("ALERT_FROM", "").strip()
    if not (host and recipients and sender):
        logging.info("Email not configured; alerts will be written to CSV/log only.")
        return False
    port = int(os.getenv("SMTP_PORT", "587"))
    username = os.getenv("SMTP_USERNAME", "")
    password = os.getenv("SMTP_PASSWORD", "")
    use_tls = os.getenv("SMTP_USE_TLS", "true").lower() in {"1", "true", "yes"}

    msg = EmailMessage()
    msg["Subject"] = "Pharma RFP Monitor Alert"
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    text_blocks, html_blocks = [], []
    for row in build_email_alert_rows(alerts):
        fields = list(zip(EMAIL_ALERT_COLUMNS, row))
        text_blocks.append("\n".join(f"{label}: {value}" for label, value in fields))
        html_blocks.append("<section>" + "<br>".join(
            f"<strong>{html.escape(label)}:</strong> {html.escape(value)}" for label, value in fields
        ) + "</section>")
    if not html_blocks:
        logging.info("No RFP, CGA, or IME evidence remains after email filtering.")
        return False
    spreadsheet_setting = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
    spreadsheet_match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", spreadsheet_setting)
    spreadsheet_id = spreadsheet_match.group(1) if spreadsheet_match else spreadsheet_setting
    spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit" if spreadsheet_id else ""
    text_body = "\n\n".join(text_blocks)
    html_body = "<hr>".join(html_blocks)
    if spreadsheet_url:
        text_body += f"\n\nGoogle Spreadsheet: {spreadsheet_url}"
        html_body += (
            "<hr><p><strong>Google Spreadsheet:</strong> "
            f'<a href="{html.escape(spreadsheet_url, quote=True)}">Open spreadsheet</a></p>'
        )
    msg.set_content(text_body)
    msg.add_alternative("<html><body>" + html_body + "</body></html>", subtype="html")
    with smtplib.SMTP(host, port, timeout=25) as server:
        if use_tls:
            server.starttls()
        if username:
            server.login(username, password)
        server.send_message(msg, to_addrs=recipients)
    logging.info("Email alert sent to %s recipient(s)", len(recipients))
    return True

def main():
    parser = argparse.ArgumentParser(description="Monitor pharma company URLs for grant/RFP-related content.")
    parser.add_argument("--input", default=str(INPUT_XLSX), help="Excel workbook path")
    parser.add_argument("--max-urls", type=int, default=0, help="Test with only the first N URLs")
    parser.add_argument("--force-alert", action="store_true", help="Email every current match, even if unchanged")
    parser.add_argument("--baseline", action="store_true", help="Save current state without sending alerts (recommended before enabling notifications)")
    args = parser.parse_args()

    setup_logging()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    records = load_input(Path(args.input))
    if args.max_urls:
        records = records[:args.max_urls]
    if not records:
        raise SystemExit("No usable URL hyperlinks found in the workbook.")
    logging.info("Loaded %d company URL(s)", len(records))

    session = make_session()
    state = get_state()
    now = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    results, alerts = [], []
    for i, rec in enumerate(records, start=1):
        logging.info("[%d/%d] Checking %s — %s", i, len(records), rec["company"], rec["url"])
        status, error, final_url, title, text = 0, "", rec["url"], "", ""
        method = "requests"
        try:
            final_url, title, text, status, note, method = extract_with_requests(session, rec["url"])
            if note:
                error = note
            if USE_PLAYWRIGHT_FALLBACK and not method.startswith(("pymupdf", "ocr", "no_extractable_text")) and (len(text) < 250 or status in (401, 403, 429)):
                try:
                    final_url, title, text, status, note, method = extract_with_playwright(rec["url"])
                    error = note or error
                except Exception as browser_exc:
                    error = (error + "; " if error else "") + f"Playwright fallback: {browser_exc}"
        except Exception as exc:
            error = str(exc)
            if USE_PLAYWRIGHT_FALLBACK:
                try:
                    final_url, title, text, status, note, method = extract_with_playwright(rec["url"])
                    error = note
                except Exception as browser_exc:
                    error += f"; Playwright fallback: {browser_exc}"

        cls = classify(rec["company"], rec["url"], rec["description"], text)
        fingerprint = hashlib.sha256((final_url + "\n" + text).encode("utf-8", errors="ignore")).hexdigest()
        key = rec["company"].casefold() + "|" + rec["url"].casefold()
        previous = state.get(key, {})
        is_new_or_changed = previous.get("fingerprint") != fingerprint
        evidence = opportunity_evidence(text)
        prior_evidence = set(previous.get("opportunity_evidence", []))
        new_evidence = [item for item in evidence if item not in prior_evidence]
        low_evidence = " ".join(evidence).casefold()
        is_rfp = any(term in low_evidence for term in RFP_TERMS)
        positive_change = bool(new_evidence) and cls["flag"]
        opportunity_type = "RFP" if is_rfp else "Grant"
        if cls["closed_signals"]:
            opportunity_status = "CLOSED"
            status_reason = "Closed wording detected: " + "; ".join(cls["closed_signals"])
        elif cls["open_signals"]:
            opportunity_status = "OPEN"
            status_reason = "Open wording detected: " + "; ".join(cls["open_signals"])
        else:
            opportunity_status = "UNKNOWN"
            status_reason = "No open/closed status terms detected"
        relevance_score = (2 if cls["flag"] else 0) + (1 if cls["topic_terms"] else 0) + (1 if cls["open_signals"] and not cls["closed_signals"] else 0)
        result = {
            "company": rec["company"], "source_url": rec["url"], "final_url": final_url,
            "discovered_from": "", "http_status": status, "extraction_method": method,
            "page_title": title, "opportunity_type": opportunity_type,
            "opportunity_title": title, "opportunity_status": opportunity_status,
            "status_reason": status_reason, "opportunity_id": "", "published_date": "",
            "published_date_text": "", "deadline_text": "", "deadline_date": "",
            "eligibility": "", "application_url": "",
            "evidence": "; ".join(new_evidence or evidence),
            "funding_terms": "; ".join(cls["funding_terms"]), "relevance_score": relevance_score,
            "change_type": "NEW_OR_CHANGED" if is_new_or_changed else "UNCHANGED",
            "changed_fields": "page_content" if is_new_or_changed else "",
            "flag": cls["flag"], "checked_at": now, "error": error,
        }
        results.append(result)
        email_evidence = email_evidence_snippets(new_evidence)
        if cls["flag"] and email_evidence and (positive_change or args.force_alert) and not args.baseline:
            alert = dict(result)
            alert.update({
                "priority": cls["priority"],
                "change_summary": "New RFP related evidence" if is_rfp else "New grant opportunity evidence",
                "new_evidence": email_evidence,
                "topic_terms": cls["topic_terms"],
            })
            alerts.append(alert)
        state[key] = {"fingerprint": fingerprint, "checked_at": now, "flag": cls["flag"], "final_url": final_url,
                      "opportunity_evidence": evidence}
        if error:
            logging.warning("Issue for %s: %s", rec["company"], error)
        elif cls["flag"]:
            logging.info("FLAG [%s] %s | funding=%s | topics=%s",
                         cls["priority"], rec["company"], cls["funding_terms"], cls["topic_terms"])
        time.sleep(DELAY_SECONDS)

    # Write full run report and cumulative current flagged results.
    pd.DataFrame(results, columns=REPORT_COLUMNS).to_csv(RUN_REPORT, index=False, encoding="utf-8-sig")
    old = pd.read_csv(MATCHES_FILE).to_dict("records") if MATCHES_FILE.exists() else []
    by_key = {(str(r.get("company", "")).casefold(), str(r.get("source_url", "")).casefold()): r for r in old}
    for r in results:
        k = (r["company"].casefold(), r["source_url"].casefold())
        if r["flag"]:
            by_key[k] = r
        elif k in by_key:
            del by_key[k]
    pd.DataFrame(list(by_key.values()), columns=REPORT_COLUMNS).to_csv(MATCHES_FILE, index=False, encoding="utf-8-sig")

    # Sync the cumulative matches after writing the local CSV. A Sheets failure
    # must not prevent local reports or email alerts from being produced.
    try:
        sync_google_sheet(MATCHES_FILE)
    except Exception:
        logging.exception("Google Sheets update failed; local CSV and email-alert processing will continue.")

    state_saved = not alerts or args.baseline
    email_sent = False
    if alerts and not args.baseline:
        try:
            state_saved = send_email_alert(alerts)
            email_sent = state_saved
        except Exception:
            logging.exception("Email notification failed; see CSV output for matches.")
            state_saved = False
    try:
        sync_email_alerts_to_google_sheet(alerts if email_sent else [])
    except Exception:
        logging.exception("Google Sheets email-alert history update failed.")
    if state_saved:
        save_state(state)
    else:
        logging.warning("State was not advanced so these alerts can be retried on the next run.")
    logging.info("Finished. Flagged this run: %d; positive-change alerts: %d", sum(r["flag"] for r in results), len(alerts))
    logging.info("Run report: %s", RUN_REPORT)
    logging.info("Current flagged matches: %s", MATCHES_FILE)


if __name__ == "__main__":
    main()
