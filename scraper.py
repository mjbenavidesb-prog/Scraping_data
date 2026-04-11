"""
UNMSM Admision Scraper
======================
Extracts all applicant results from https://admision.unmsm.edu.pe/Website20262/A/A.html
career by career, handling DataTables JS pagination, and exports to Excel.

Strategy for DataTables pagination:
  The page loads ALL records into memory (client-side DataTables) but only
  displays 50 at a time.  We bypass pagination entirely by using the
  DataTables JavaScript API to pull every row's text in a single JS call,
  completely avoiding DOM pagination.  If that fails we fall back to
  changing the page-length selector to "All" (-1).
"""

import logging
import os
import time
from datetime import datetime

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ── Configuration ──────────────────────────────────────────────────────────────
BASE_URL = "https://admision.unmsm.edu.pe/Website20262/A/A.html"
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"resultados_admision_{TIMESTAMP}.xlsx")

PAGE_LOAD_WAIT = 15   # seconds to wait for page elements
DRAW_WAIT = 3         # seconds after triggering DataTable redraw

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ── Driver ─────────────────────────────────────────────────────────────────────
def build_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    # Suppress browser logging noise
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # Try webdriver-manager for automatic ChromeDriver resolution; fall back to
    # Selenium 4's built-in driver manager (works when Chrome is installed).
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)
    except Exception:
        return webdriver.Chrome(options=options)


# ── Helpers ────────────────────────────────────────────────────────────────────
def wait_for_table(driver: webdriver.Chrome, timeout: int = PAGE_LOAD_WAIT) -> bool:
    """Return True if a <table> with at least one <tbody tr> is found."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )
        return True
    except Exception:
        return False


def get_datatable_id(driver: webdriver.Chrome) -> str | None:
    """Return the id of the first DataTable on the page, or None."""
    return driver.execute_script("""
        try {
            var tables = $.fn.dataTable.tables();
            return tables.length > 0 ? tables[0].id : null;
        } catch(e) { return null; }
    """)


# ── Primary extraction: JS DataTables API ──────────────────────────────────────
#
# Cell extraction rules (discovered by inspecting the UNMSM page HTML):
#   • Puntaje  → <td data-score="461.375"></td>        ← value in data-score, text empty
#   • Mérito   → <td data-merit="41"></td>             ← value in data-merit, text empty
#   • Nombres/Escuela → <span class="obfuscated" data-auth="BASE64">text</span>
#                       ← base64 decoded for correct UTF-8 (Ñ, Á, É, …)
#   • Other cells → plain innerText
#
_JS_EXTRACT_ALL = """
try {
    var dt = $('#%s').DataTable();
    var result = [];
    dt.rows().every(function() {
        var node = this.node();
        if (!node) return;
        var cells = node.querySelectorAll('td');
        var row = [];
        cells.forEach(function(c) {
            var text;
            if (c.hasAttribute('data-score')) {
                // Puntaje: numeric value stored as attribute, cell body is empty
                text = c.getAttribute('data-score') || '';
            } else if (c.hasAttribute('data-merit')) {
                // Mérito E.P: same pattern
                text = c.getAttribute('data-merit') || '';
            } else {
                var span = c.querySelector('span.obfuscated[data-auth]');
                if (span) {
                    // Names / Escuela are base64-encoded UTF-8 to prevent simple scraping
                    try {
                        text = decodeURIComponent(escape(atob(span.getAttribute('data-auth'))));
                    } catch(e) {
                        text = span.innerText.trim();
                    }
                } else {
                    text = c.innerText.trim();
                }
            }
            row.push(text);
        });
        result.push(row);
    });
    return result;
} catch(e) { return null; }
"""


def extract_via_js_api(driver: webdriver.Chrome, table_id: str) -> list[list[str]] | None:
    """
    Pull ALL rows from DataTables memory via the JS API.
    Returns a list of rows (each row is a list of cell strings), or None on failure.
    This works regardless of the current page view — no pagination needed.
    """
    data = driver.execute_script(_JS_EXTRACT_ALL % table_id)
    if data and len(data) > 0:
        return data
    return None


# ── Fallback extraction: change page-length to "All" ──────────────────────────
_JS_SHOW_ALL = """
try {
    var dt = $('#%s').DataTable();
    dt.page.len(-1).draw(false);
    return dt.rows({ search: 'applied' }).count();
} catch(e) { return -1; }
"""


_JS_FALLBACK_ROWS = """
(function() {
    var rows = document.querySelectorAll('#%s tbody tr');
    var result = [];
    rows.forEach(function(tr) {
        var cells = tr.querySelectorAll('td');
        var row = [];
        cells.forEach(function(c) {
            var text;
            if (c.hasAttribute('data-score')) {
                text = c.getAttribute('data-score') || '';
            } else if (c.hasAttribute('data-merit')) {
                text = c.getAttribute('data-merit') || '';
            } else {
                var span = c.querySelector('span.obfuscated[data-auth]');
                if (span) {
                    try { text = decodeURIComponent(escape(atob(span.getAttribute('data-auth')))); }
                    catch(e) { text = span.innerText.trim(); }
                } else {
                    text = c.innerText.trim();
                }
            }
            row.push(text);
        });
        result.push(row);
    });
    return result;
})();
"""


def extract_via_show_all(driver: webdriver.Chrome, table_id: str) -> list[list[str]]:
    """
    Fallback: set DataTable page-length to -1 (All) and scrape visible DOM rows.
    Uses the same attribute-aware extraction as the primary JS path.
    """
    total = driver.execute_script(_JS_SHOW_ALL % table_id)
    log.info(f"    Fallback: showing all {total} rows, waiting for redraw...")
    time.sleep(DRAW_WAIT)
    return driver.execute_script(_JS_FALLBACK_ROWS % table_id) or []


# ── Career list ────────────────────────────────────────────────────────────────
def collect_career_links(driver: webdriver.Chrome) -> list[tuple[str, str]]:
    """
    Load the main page and return a list of (career_name, absolute_url) tuples.
    Links are found inside <table> elements.
    """
    log.info(f"Loading main page: {BASE_URL}")
    driver.get(BASE_URL)

    try:
        WebDriverWait(driver, PAGE_LOAD_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table a"))
        )
    except Exception:
        log.error("Timed out waiting for career links on the main page.")
        return []

    elements = driver.find_elements(By.CSS_SELECTOR, "table a")
    seen = set()
    careers = []
    for el in elements:
        name = el.text.strip()
        href = el.get_attribute("href") or ""
        if name and href and href not in seen:
            seen.add(href)
            careers.append((name, href))

    log.info(f"Found {len(careers)} career links.")
    return careers


# ── Per-career scraping ────────────────────────────────────────────────────────
COLUMNS = ["Carrera", "Código", "Apellidos y Nombres", "Escuela",
           "Puntaje", "Mérito E.P", "Observación"]


_NETWORK_ERRORS = ("ERR_INTERNET_DISCONNECTED", "ERR_NAME_NOT_RESOLVED",
                   "ERR_CONNECTION_REFUSED", "ERR_CONNECTION_TIMED_OUT",
                   "net::ERR_")
MAX_RETRIES = 4
RETRY_BACKOFF = [5, 15, 30, 60]   # seconds between retries


def scrape_career(driver: webdriver.Chrome, name: str, url: str) -> list[dict]:
    """
    Navigate to a career page and return all applicant records.
    Retries on transient network errors with exponential back-off.
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            driver.get(url)
        except Exception as exc:
            msg = str(exc)
            is_network = any(e in msg for e in _NETWORK_ERRORS)
            if is_network and attempt < MAX_RETRIES:
                wait = RETRY_BACKOFF[attempt - 1]
                log.warning(f"  Network error on attempt {attempt}/{MAX_RETRIES}, "
                            f"retrying in {wait}s... ({msg.splitlines()[0]})")
                time.sleep(wait)
                continue
            log.warning(f"  Could not navigate to {url}: {msg.splitlines()[0]}")
            return []

        if not wait_for_table(driver):
            # Could be a transient load failure — retry if network-related
            if attempt < MAX_RETRIES:
                wait = RETRY_BACKOFF[attempt - 1]
                log.warning(f"  No table rows on attempt {attempt}/{MAX_RETRIES}, "
                            f"retrying in {wait}s...")
                time.sleep(wait)
                continue
            log.warning(f"  No table rows found for '{name}', skipping.")
            return []

        # Page loaded — proceed
        break
    else:
        return []

    # Small pause for DataTables to fully initialise
    time.sleep(1)

    table_id = get_datatable_id(driver)
    if not table_id:
        log.warning(f"  No DataTable detected for '{name}', skipping.")
        return []

    # ── Strategy 1: JS API (fastest — no DOM pagination) ──────────────────────
    raw_rows = extract_via_js_api(driver, table_id)

    # ── Strategy 2: Show-all fallback ─────────────────────────────────────────
    if raw_rows is None:
        log.info(f"  JS API returned nothing, trying show-all fallback...")
        raw_rows = extract_via_show_all(driver, table_id)

    records = []
    for row in raw_rows:
        if len(row) >= 6:
            records.append({
                "Carrera":             name,
                "Código":              row[0],
                "Apellidos y Nombres": row[1],
                "Escuela":             row[2],
                "Puntaje":             row[3],
                "Mérito E.P":          row[4],
                "Observación":         row[5],
            })

    return records


# ── Excel export ───────────────────────────────────────────────────────────────
def export_excel(df: pd.DataFrame, path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)

    # Numeric coercion
    df["Puntaje"] = pd.to_numeric(df["Puntaje"], errors="coerce")
    df["Mérito E.P"] = pd.to_numeric(df["Mérito E.P"], errors="coerce")

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")
        ws = writer.sheets["Resultados"]

        # Column widths: A=Carrera B=Código C=Nombres D=Escuela E=Puntaje F=Mérito G=Obs
        for col_letter, width in zip("ABCDEFG", [28, 12, 40, 40, 10, 10, 30]):
            ws.column_dimensions[col_letter].width = width

        # Styles
        header_fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        passed_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
        thin = Side(style="thin")
        thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Header row
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        # Data rows
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            observacion = str(row[6].value or "")   # column G = Observación
            is_vacante = "VACANTE" in observacion.upper()
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                if is_vacante:
                    cell.fill = passed_fill

        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "A2"

    log.info(f"Excel saved: {path}")


# ── Main ───────────────────────────────────────────────────────────────────────
def main() -> None:
    driver = build_driver()
    all_records: list[dict] = []

    try:
        careers = collect_career_links(driver)
        if not careers:
            log.error("No careers found. Exiting.")
            return

        total = len(careers)
        for idx, (name, url) in enumerate(careers, start=1):
            log.info(f"[{idx}/{total}] {name}")
            records = scrape_career(driver, name, url)
            all_records.extend(records)
            log.info(f"  → {len(records):,} applicants extracted "
                     f"(running total: {len(all_records):,})")
    finally:
        driver.quit()

    log.info(f"\nTotal records across all careers: {len(all_records):,}")

    if not all_records:
        log.error("No data was extracted. Nothing to save.")
        return

    df = pd.DataFrame(all_records, columns=COLUMNS)
    export_excel(df, OUTPUT_FILE)
    print(f"\nDone!  Output file: {OUTPUT_FILE}")
    print(f"Total rows: {len(df):,}  |  Careers: {df['Carrera'].nunique()}")


if __name__ == "__main__":
    main()
