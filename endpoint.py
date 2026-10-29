#!/usr/bin/env python3
"""
Dynatrace MDA — Slow Endpoint Report with Response Time Hotspot Reasons

For each endpoint whose average response time > 5 s:
1. Reads the MDA table, including virtual scrolling.
2. Finds the endpoint row and opens its Response time hotspots view.
3. Extracts hotspot findings using legacy selectors and visible-page fallbacks.
4. Classifies findings into External and Internal reasons.
5. Writes Excel and Word reports.

Time window is selected automatically based on when the script is run (IST):
    Run 11:00–14:59  →  report window 08:00–11:00
    Run 15:00–17:59  →  report window 11:00–15:00
    Run 18:00–23:59  →  report window 15:00–18:00

Install:
    pip install playwright python-docx openpyxl
    playwright install chromium
"""

import os
import re
import shutil
import time
from datetime import datetime, timezone, timedelta

from playwright.sync_api import sync_playwright


# ========================= CONFIG =========================

IST = timezone(timedelta(hours=5, minutes=30))


def get_mda_url():
    now = datetime.now(IST)
    total_minutes = now.hour * 60 + now.minute

    if 11 * 60 <= total_minutes < 15 * 60:
        start_h, end_h = 8, 11
    elif 15 * 60 <= total_minutes < 18 * 60:
        start_h, end_h = 11, 15
    elif 18 * 60 <= total_minutes < 24 * 60:
        start_h, end_h = 15, 18
    else:
        raise RuntimeError(
            f"Script run at {now.strftime('%H:%M')} IST is outside all "
            "expected windows.\n"
            "Expected run times:\n"
            "  08:00–11:00 report  →  run between 11:00–14:59\n"
            "  11:00–15:00 report  →  run between 15:00–17:59\n"
            "  15:00–18:00 report  →  run between 18:00–23:59"
        )

    d     = now.strftime("%Y-%m-%d")
    start = f"{d}T{start_h:02d}%3A00%3A00%2B05%3A30"
    end   = f"{d}T{end_h:02d}%3A00%3A00%2B05%3A30"

    print(
        f"Time window selected: {start_h:02d}:00 – {end_h:02d}:00 IST "
        f"on {d}  (script run at {now.strftime('%H:%M')} IST)",
        flush=True,
    )

    return (
        "https://abcdapmsaas.apps.dynatrace.com/ui/apps/dynatrace.classic.mda"
        "/ui/diagnostictools/mda"
        f"?gtf={start}+to+{end}"
        "&gf=all&mdaId=topweb&metric=RESPONSE_TIME"
        "&dimension=%7BRequest%3AName%7D&mergeServices=false"
        "&aggregation=AVERAGE&percentile=80&chart=COLUMN"
        "&servicefilter=0%1E26%112%1052%11abc-%100%115000000%144611686018427387"
    )


MDA_URL = get_mda_url()

TOP_N        = 20
THRESHOLD_MS = 5_000
IGNORE       = ("schedule", "terminatesessionandupdatelead")

OUT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "endpoint_response_time_report"
)

COL_ENDPOINT = 0
COL_SERVICE  = 1
COL_REQUESTS = 3
COL_RESPONSE = 4
MIN_COLS     = 5

REFINE_XPATH = (
    "/html/body/app-root/ng-component/app-shell/main/ng-component"
    "/page-content/div/mda-chart/div/div[1]/div[2]/refine-button"
)


# ========================= KEYWORD LISTS =========================

EXT_KW = [
    "calls to http",
    "calls to https",
    "calls to get ",
    "calls to post ",
    "calls to put ",
    "calls to delete ",
    "network io",
    "external call",
    "remote call",
]

INT_KW = [
    "active wait",
    "suspension",
    "sql quer",
    "queries or procedures",
    "sql modification",
    "select ",
    "insert ",
    "update ",
    "delete ",
    "merge ",
    "lock wait",
    "database",
    "jdbc",
    "service execution",
    "database usage",
]

_SKIP_LABELS = {
    # Dynatrace Distribution category headers (left panel)
    "interaction with services and queues",
    "interaction with services",
    "interaction with third-party services",
    "interaction with databases",
    "database usage",
    "service execution",
    # Page-level labels (circle, breadcrumb, section titles)
    "response time",
    "response time analysis",
    "top findings",
    "distribution",
    "need help",
    # Single-word noise fragments from split
    "network",
    "memory",
    "cpu",
    "queues",
    "services",
    "databases",
    # Port-only fragments
    "port 443",
    "port 80",
}

_BARE_NUM  = re.compile(r"^\d+(?:\.\d+)?$")
_BARE_UNIT = re.compile(
    r"^(?:ns|us|µs|μs|ms|s|sec|secs|min|mins|h|hr)$", re.I
)
_DURATION  = re.compile(
    r"^\d+(?:\.\d+)?\s*(?:ns|us|µs|μs|ms|s|sec|secs|min|mins|h|hr)$", re.I
)


# ========================= HELPERS =========================

def clean(x):
    return " ".join(
        str(x or "").replace("\n", " ").replace("\r", " ").split()
    )


def extract_service(cell_text):
    text = clean(cell_text)
    return text.split(" tomcat")[0].strip() if " tomcat" in text else text


def parse_duration_ms(value):
    original = clean(value)
    if not original or "rarely occurred" in original.lower():
        return None
    value = original.lower().replace(",", "")
    units = {
        "ns": 1e-6, "us": 1e-3, "µs": 1e-3, "μs": 1e-3,
        "ms": 1,
        "s": 1_000, "sec": 1_000, "secs": 1_000,
        "min": 60_000, "mins": 60_000,
        "h": 3_600_000, "hr": 3_600_000,
    }
    unit_pattern = "|".join(sorted(units.keys(), key=len, reverse=True))
    parts = re.findall(rf"(\d+(?:\.\d+)?)\s*({unit_pattern})\b", value)
    if not parts:
        return None
    return sum(float(n) * units[u] for n, u in parts)


def ignored(endpoint):
    return any(kw in clean(endpoint).lower() for kw in IGNORE)


def text_has_reason(text):
    text_lower = clean(text).lower()
    return any(
        kw in text_lower
        for kw in (
            EXT_KW + INT_KW
            + ["response time", "wait time", "sql queries",
               "queries", "procedures", "hotspot"]
        )
    )


# ========================= REASON CLASSIFICATION =========================

def _is_noise_clause(clause):
    """Return True for bare numbers, bare units, standalone durations,
    and Dynatrace card category headers that carry no finding information."""
    c = clause.strip()
    if not c:
        return True
    if _BARE_NUM.match(c):
        return True
    if _BARE_UNIT.match(c):
        return True
    if _DURATION.match(c):
        return True
    if c.lower() in _SKIP_LABELS:
        return True
    return False


def classify_reason(text):
    """
    Classify hotspot findings into (external, internal).

    Rules (matching the manual report convention):
    1. Split on ' and ' / '|', strip noise clauses (bare numbers, units,
       category headers, short meaningless fragments).
    2. Findings are already limited to the top 2 by read_hotspot_findings,
       so classify whatever arrives.
    3. If ANY external clause is present (Network IO, Calls to URL, etc.):
         - Reason(External) = only the external clauses (Network IO + URL calls).
         - Reason(Internal) = NA.
         - Suspension / Active wait time are DROPPED when external is present.
    4. If findings are ONLY Active wait time and/or Suspension (no external URL,
       no DB query):
         - Reason(External) = NA.
         - Reason(Internal) = those clauses joined by ' and '.
    5. If internal DB/SQL findings exist with no external clause:
         - Reason(External) = NA.
         - Reason(Internal) = those clauses.
    """
    text = clean(text)
    if not text:
        return "NA", "NA"

    # Split on conjunctions used by both the legacy and card layouts.
    clauses = [
        clean(c)
        for c in re.split(r"\s+(?:and|\|)\s+", text, flags=re.IGNORECASE)
        if clean(c)
    ]

    # Remove noise before any classification.
    clauses = [c for c in clauses if not _is_noise_clause(c)]

    if not clauses:
        return "NA", "NA"

    def is_external_clause(clause):
        lower = clause.lower()
        return (
            any(kw in lower for kw in EXT_KW)
            or bool(re.search(r"https?://|url\.|\.com\b|\.in\b", lower))
        )

    def is_wait_or_suspension(clause):
        lower = clause.lower()
        return "active wait" in lower or "suspension" in lower

    # ── Helpers ───────────────────────────────────────────────────────────
    def is_url_call(clause):
        """True only when the clause references an actual outbound URL call."""
        lower = clause.lower()
        return bool(re.search(r"https?://", lower)) or any(
            kw in lower
            for kw in ["calls to http", "calls to https",
                       "calls to get ", "calls to post ",
                       "calls to put ", "calls to delete "]
        )

    def is_network_io(clause):
        return "network io" in clause.lower()

    def is_active_wait(clause):
        return "active wait" in clause.lower()

    def is_suspension(clause):
        return "suspension" in clause.lower()

    url_call_present = any(is_url_call(c) for c in clauses)

    if url_call_present:
        # External URL found → ALL top-2 findings go to External, Internal = NA.
        external = " and ".join(dict.fromkeys(clauses)) if clauses else "NA"
        return external, "NA"

    # No URL call → ALL top-2 findings go to Internal, External = NA.
    internal = " and ".join(dict.fromkeys(clauses)) if clauses else "NA"
    return "NA", internal


# ========================= BROWSER HELPERS =========================

def find_frame(page, timeout=120):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for frame in page.frames:
            try:
                if (
                    frame.locator("dt-table").count()
                    or frame.locator("refine-button").count()
                ):
                    return frame
            except Exception:
                pass
        time.sleep(0.5)
    return None


def signature(frame):
    try:
        rows = frame.locator("dt-table dt-row")
        return "|".join(
            clean(rows.nth(i).inner_text())
            for i in range(min(rows.count(), 10))
        )
    except Exception:
        return ""


# ========================= REFINE =========================

def refine(frame):
    button = frame.locator(f"xpath={REFINE_XPATH}")
    if not button.count():
        button = frame.locator("refine-button")
    if not button.count():
        raise RuntimeError("MDA Refine button not found.")

    button.first.click(force=True, timeout=10_000)
    print("Refine button clicked; checking for filter panel...", flush=True)

    panel_deadline = time.time() + 15
    apply_clicked  = False
    panel_seen     = False
    apply_pattern  = re.compile(r"^(apply|update|run|submit|refresh)(\b|$)", re.I)

    while time.time() < panel_deadline:
        try:
            panel_candidates = [
                frame.locator("[role='dialog']"),
                frame.locator("[role='menu']"),
                frame.locator(".cdk-overlay-pane"),
                frame.locator("refine-panel"),
            ]
            panel_seen = any(
                c.count() and c.last.is_visible()
                for c in panel_candidates
            )
            buttons = frame.locator("button")
            for i in range(buttons.count()):
                candidate = buttons.nth(i)
                if not candidate.is_visible():
                    continue
                label = clean(
                    candidate.get_attribute("aria-label")
                    or candidate.inner_text()
                )
                if apply_pattern.search(label):
                    candidate.click(force=True, timeout=5_000)
                    print(f"Clicked Refine panel action: {label!r}", flush=True)
                    apply_clicked = True
                    break
            if apply_clicked:
                break
        except Exception:
            pass
        time.sleep(0.4)

    if panel_seen and not apply_clicked:
        print("Refine panel detected; no Apply/Update button found. Waiting...", flush=True)
    elif not panel_seen:
        print("No separate Refine panel; waiting for table refresh...", flush=True)

    deadline     = time.time() + 180
    last_sig     = None
    stable_count = 0
    last_status  = 0

    while time.time() < deadline:
        try:
            rows      = frame.locator("dt-table dt-row")
            current   = signature(frame)
            row_count = rows.count()

            page_text = ""
            try:
                page_text = clean(
                    frame.locator("body").inner_text(timeout=2_000)
                ).lower()
            except Exception:
                pass

            loading = any(
                m in page_text
                for m in ("loading data", "loading...", "please wait",
                          "fetching data", "updating data")
            )
            pending = any(
                "rarely occurred" in clean(rows.nth(i).inner_text()).lower()
                for i in range(min(row_count, 20))
            )

            if current and current == last_sig and not loading and not pending:
                stable_count += 1
            else:
                stable_count = 0
            last_sig = current

            if (
                row_count > 0
                and current
                and stable_count >= 3
                and not loading
                and not pending
            ):
                print("MDA Refine completed; table is stable.", flush=True)
                return

            elapsed = int(180 - (deadline - time.time()))
            if elapsed - last_status >= 15:
                print(
                    f"Still waiting for MDA table to settle ({elapsed}s; "
                    f"rows={row_count}, loading={loading}, pending={pending})...",
                    flush=True,
                )
                last_status = elapsed
        except Exception:
            pass
        time.sleep(0.7)

    raise RuntimeError(
        "MDA table did not become stable within 3 minutes after Refine."
    )


# ========================= READ VIRTUAL TABLE =========================

def collect_rows(frame):
    table = frame.locator("dt-table").first
    if not table.count():
        raise RuntimeError("MDA table not found.")

    try:
        table.evaluate("""el => {
            let n = el;
            while (n) {
                if (n.scrollHeight > n.clientHeight + 10) n.scrollTop = 0;
                n = n.parentElement;
            }
            window.scrollTo(0, 0);
        }""")
    except Exception:
        pass

    found, seen = [], set()
    deadline = time.time() + 180
    rounds   = 0

    while time.time() < deadline and rounds < 1_000:
        rounds += 1
        try:
            rows = frame.locator("dt-table dt-row")
            for i in range(rows.count()):
                try:
                    r = rows.nth(i)
                    if not r.is_visible(timeout=100):
                        continue
                    cells = [clean(v) for v in r.locator("dt-cell").all_inner_texts()]
                    if len(cells) < MIN_COLS:
                        continue
                    key = tuple(cells)
                    if key not in seen:
                        seen.add(key)
                        found.append(cells)
                except Exception:
                    pass

            moved = table.evaluate("""el => {
                const step = Math.max(300, Math.floor((innerHeight || 700) * 0.7));
                let n = el;
                while (n) {
                    if (n.scrollHeight > n.clientHeight + 10) {
                        const b = n.scrollTop;
                        n.scrollTop = Math.min(b + step, n.scrollHeight - n.clientHeight);
                        return { moved: n.scrollTop > b + 1,
                                 bottom: n.scrollTop + n.clientHeight >= n.scrollHeight - 5 };
                    }
                    n = n.parentElement;
                }
                const d = document.scrollingElement || document.documentElement;
                const b = d.scrollTop;
                d.scrollTop = Math.min(b + step, d.scrollHeight - d.clientHeight);
                return { moved: d.scrollTop > b + 1,
                         bottom: d.scrollTop + d.clientHeight >= d.scrollHeight - 5 };
            }""")

            print(f"Scan {rounds}: {len(found)} unique rows", flush=True)

            if moved.get("bottom"):
                time.sleep(0.8)
                rows2  = frame.locator("dt-table dt-row")
                before = len(found)
                for i in range(rows2.count()):
                    try:
                        cells = [
                            clean(v)
                            for v in rows2.nth(i).locator("dt-cell").all_inner_texts()
                        ]
                        if len(cells) >= MIN_COLS and tuple(cells) not in seen:
                            seen.add(tuple(cells))
                            found.append(cells)
                    except Exception:
                        pass
                if len(found) == before:
                    break
        except Exception:
            pass
        time.sleep(0.3)

    print(f"Completed MDA scan: {len(found)} rows total.", flush=True)
    return found


# ========================= FILTER + TOP N =========================

def get_top_n(frame):
    result, seen = [], set()

    for idx, cells in enumerate(collect_rows(frame)):
        if idx < 10:
            print(f"RAW MDA ROW {idx + 1}: {cells}", flush=True)
        if len(cells) < MIN_COLS:
            continue
        endpoint = clean(cells[COL_ENDPOINT])
        if not endpoint or ignored(endpoint):
            continue
        service      = extract_service(cells[COL_SERVICE])
        requests     = clean(cells[COL_REQUESTS])
        response_raw = clean(cells[COL_RESPONSE])
        ms           = parse_duration_ms(response_raw)
        if ms is None or ms <= THRESHOLD_MS:
            continue
        print(
            f"QUALIFIED: {endpoint} | {service} | {requests} req | {response_raw}",
            flush=True,
        )
        key = (endpoint.lower(), service.lower())
        if key in seen:
            continue
        seen.add(key)
        result.append({
            "Endpoint":         endpoint,
            "Service name":     service,
            "Requests":         requests,
            "Response time":    response_raw,
            "Reason(External)": "NA",
            "Reason(Internal)": "NA",
            "_ms":              ms,
        })

    result.sort(key=lambda x: x["_ms"], reverse=True)
    result = result[:TOP_N]
    for i, row in enumerate(result, 1):
        row["#"] = i

    print(f"\nEndpoints > 5 s: {len(result)}", flush=True)
    for row in result:
        print(
            f"  {row['#']:2}. {row['Endpoint']:<45} "
            f"{row['Service name']:<38} {row['Requests']:>5} req  "
            f"{row['Response time']}",
            flush=True,
        )
    return result


# ========================= HOTSPOT REASON FETCHING =========================

HOTSPOT_MENU_TEXT = "Response time hotspots"
FINDINGS_SEL      = '[uitestid*="rta-overview-finding"][uitestid*="-name"]'
FINDINGS_CSS      = "table.dnk-db tbody tr td.dnk-g div.dnk-a.dnk-r"

REASON_SIGNAL_KW = [
    "calls to http", "calls to https", "calls to get ", "calls to post ",
    "calls to put ", "calls to delete ", "active wait", "network io", "sql",
    "suspension", "queries or procedures", "sql modification", "database usage",
    "sql queries or procedures", "service execution", "response time", "wait time",
]


def scroll_table_to_top(frame):
    try:
        frame.evaluate("""() => {
            window.scrollTo(0, 0);
            const t = document.querySelector('dt-table');
            if (t) {
                let n = t;
                while (n) {
                    if (n.scrollHeight > n.clientHeight + 10) n.scrollTop = 0;
                    n = n.parentElement;
                }
            }
        }""")
    except Exception:
        pass


def find_row_element(frame, endpoint_name, scroll_attempts=15):
    scroll_table_to_top(frame)
    time.sleep(0.3)

    for attempt in range(scroll_attempts):
        rows = frame.locator("dt-table dt-row")
        for i in range(rows.count()):
            try:
                r     = rows.nth(i)
                cells = r.locator("dt-cell")
                if cells.count() == 0:
                    continue
                if (
                    clean(cells.first.inner_text(timeout=500)).lower()
                    == endpoint_name.lower()
                ):
                    return r
            except Exception:
                pass

        if attempt == scroll_attempts - 1:
            break
        try:
            frame.evaluate("""() => {
                const t = document.querySelector('dt-table');
                if (t) {
                    let n = t;
                    while (n) {
                        if (n.scrollHeight > n.clientHeight + 10) {
                            n.scrollTop = Math.min(n.scrollTop + 250, n.scrollHeight);
                            return;
                        }
                        n = n.parentElement;
                    }
                }
                window.scrollBy(0, 250);
            }""")
        except Exception:
            pass
        time.sleep(0.3)

    return None


def click_three_dots(row_el):
    try:
        row_el.scroll_into_view_if_needed(timeout=5_000)
    except Exception:
        pass
    try:
        row_el.hover(timeout=5_000)
    except Exception:
        pass
    time.sleep(0.8)

    try:
        drill = row_el.locator("service-drilldown")
        if drill.count():
            drill.last.hover(timeout=3_000)
            time.sleep(0.8)
    except Exception:
        pass

    button = None
    for selector in (
        "dt-context-dialog button",
        "service-drilldown button",
        "button[aria-label='Open context dialog']",
    ):
        try:
            candidate = row_el.locator(selector)
            if candidate.count():
                button = candidate.last
                break
        except Exception:
            pass

    if button is None:
        print("  [DEBUG] Context/3-dots button does not exist in row", flush=True)
        return False

    try:
        if button.is_visible(timeout=2_000):
            button.click(force=True, timeout=5_000)
            print("  3-dots clicked", flush=True)
            return True
    except Exception:
        pass

    try:
        row_el.hover(timeout=3_000)
        drill = row_el.locator("service-drilldown")
        if drill.count():
            drill.last.hover(timeout=3_000)
        time.sleep(1.0)
        if button.is_visible(timeout=2_000):
            button.click(force=True, timeout=5_000)
            print("  3-dots clicked after re-hover", flush=True)
            return True
    except Exception:
        pass

    try:
        button.evaluate("""el => {
            el.dispatchEvent(new MouseEvent('mousedown',
                {bubbles:true,cancelable:true,view:window}));
            el.dispatchEvent(new MouseEvent('mouseup',
                {bubbles:true,cancelable:true,view:window}));
            el.click();
        }""")
        print("  3-dots clicked via DOM", flush=True)
        return True
    except Exception as e:
        print(f"  [DEBUG] DOM 3-dots click failed: {e}", flush=True)
        return False


def open_three_dots_and_click_hotspot(frame, page, row_el, endpoint_name):
    if not click_three_dots(row_el):
        print("  WARNING: could not open 3-dots menu", flush=True)
        return False

    time.sleep(1.2)

    for scope in (page, frame):
        for selector in (
            "a.link:has-text('Response time hotspots')",
            "a:has-text('Response time hotspots')",
        ):
            try:
                opt = scope.locator(selector)
                if opt.count():
                    opt.first.click(timeout=4_000)
                    print("  Clicked: 'Response time hotspots'", flush=True)
                    return True
            except Exception:
                pass

        try:
            opt = scope.get_by_text(HOTSPOT_MENU_TEXT, exact=False)
            if opt.count():
                opt.first.click(timeout=4_000)
                print("  Clicked: 'Response time hotspots'", flush=True)
                return True
        except Exception:
            pass

    print(
        "  [DEBUG] 'Response time hotspots' not found. Menu contents:",
        flush=True,
    )
    for sel in (
        ".cdk-overlay-pane", ".cdk-overlay-container",
        "dt-context-dialog-panel", "[class*='context-dialog']",
        "service-drilldown-link-group",
    ):
        for scope in (page, frame):
            try:
                el = scope.locator(sel)
                if el.count():
                    t = clean(el.first.inner_text(timeout=700))
                    if t:
                        print(f"    [{sel}] -> {t[:500]}", flush=True)
            except Exception:
                pass

    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    return False


def read_hotspot_findings(page, endpoint_name, wait_sec=25):
    deadline = time.time() + wait_sec

    while time.time() < deadline:
        scopes = [page] + list(page.frames)

        # 1. Legacy test-id selector — top 2 findings only
        for scope in scopes:
            try:
                findings = scope.locator(FINDINGS_SEL)
                texts = []
                for i in range(min(findings.count(), 30)):
                    t = clean(findings.nth(i).inner_text(timeout=700))
                    if t and not _is_noise_clause(t) and len(t.split()) >= 2:
                        texts.append(t)
                    if len(texts) == 2:
                        break
                if texts:
                    result = " | ".join(dict.fromkeys(texts))
                    print(f"  Findings by test-id (top {len(texts)}): {result[:300]}", flush=True)
                    return result
            except Exception:
                pass

        # 2. Legacy table CSS — top 2 findings only
        for scope in scopes:
            try:
                cells = scope.locator(FINDINGS_CSS)
                texts = []
                for i in range(min(cells.count(), 30)):
                    t = clean(cells.nth(i).inner_text(timeout=700))
                    if t and not _is_noise_clause(t) and len(t.split()) >= 2:
                        texts.append(t)
                    if len(texts) == 2:
                        break
                if texts:
                    result = " | ".join(dict.fromkeys(texts))
                    print(f"  Findings by table CSS (top {len(texts)}): {result[:300]}", flush=True)
                    return result
            except Exception:
                pass

        # 3. Card-based Top findings section
        for scope in scopes:
            try:
                body_text = scope.locator("body").inner_text(timeout=1_000)
                if not body_text:
                    continue

                lines = [clean(l) for l in body_text.splitlines() if clean(l)]
                heading_idx = next(
                    (i for i, l in enumerate(lines) if "top findings" in l.lower()),
                    None,
                )

                if heading_idx is not None:
                    # Collect only the top 2 meaningful finding lines.

                    section_lines = []
                    for line in lines[heading_idx + 1:]:
                        lower = line.lower()
                        if "need help" in lower or "distribution" in lower:
                            break
                        if (
                            not _is_noise_clause(line)
                            and len(line.split()) >= 2
                        ):
                            section_lines.append(line)
                        if len(section_lines) == 2:
                            break

                    candidate = clean(" | ".join(section_lines))
                    if candidate:
                        print(
                            f"  Findings from Top findings section (top {len(section_lines)}): "
                            f"{candidate[:500]}",
                            flush=True,
                        )
                        return candidate

                # 4. Full-page fallback when on the analysis URL
                lower_body    = clean(body_text).lower()
                loading_markers = (
                    "loading data", "loading...", "please wait", "fetching data"
                )
                concrete_kw = tuple(
                    kw for kw in (EXT_KW + INT_KW)
                    if kw not in ("response time", "wait time")
                )
                has_concrete = any(kw in lower_body for kw in concrete_kw)

                if any(m in lower_body for m in loading_markers):
                    time.sleep(0.5)
                    continue

                if "responsetimeanalysis" in page.url.lower() and has_concrete:
                    candidate = clean(body_text)
                    candidate = re.sub(
                        r"(?i)\bservices?\s+[^|]+?\s+details\b", "", candidate
                    )
                    candidate = clean(candidate)
                    if candidate and has_concrete:
                        print(
                            f"  Findings from visible page text (fallback): "
                            f"{candidate[:500]}",
                            flush=True,
                        )
                        return candidate

            except Exception:
                pass

        time.sleep(0.5)

    print(
        f"  [DEBUG] No findings after {wait_sec}s. "
        f"Endpoint: {endpoint_name}. URL: {page.url}",
        flush=True,
    )
    for scope in [page] + list(page.frames):
        try:
            body_text = clean(scope.locator("body").inner_text(timeout=1_500))
            if body_text:
                print(f"  [DEBUG] Visible text preview: {body_text[:700]}", flush=True)
                print(
                    f"  [DEBUG] reason-signal-present={text_has_reason(body_text)}",
                    flush=True,
                )
                break
        except Exception:
            pass

    return ""


def fetch_reasons(frame, page, rows):
    print("\n======= Fetching Response Time Hotspot Reasons =======", flush=True)

    for i, row in enumerate(rows):
        ep = row["Endpoint"]
        print(f"\n[{i + 1}/{len(rows)}] {ep}", flush=True)

        if page.url != MDA_URL:
            page.goto(MDA_URL, wait_until="domcontentloaded", timeout=60_000)
            frame = find_frame(page, timeout=30)
            if not frame:
                print("  ERROR: lost MDA frame", flush=True)
                break
            time.sleep(1)

        scroll_table_to_top(frame)
        time.sleep(0.3)

        row_el = find_row_element(frame, ep)
        if row_el is None:
            print("  WARNING: row not found — skipping", flush=True)
            continue

        initiated = open_three_dots_and_click_hotspot(frame, page, row_el, ep)
        if not initiated:
            print("  WARNING: could not open hotspot — skipping", flush=True)
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            continue

        time.sleep(2)

        raw = read_hotspot_findings(page, ep, wait_sec=25)

        if raw:
            print(
                f"  Raw: {raw[:300]}{'...' if len(raw) > 300 else ''}",
                flush=True,
            )
            external, internal = classify_reason(raw)
        else:
            external = internal = "NA"

        row["Reason(External)"] = external
        row["Reason(Internal)"] = internal
        print(f"  External : {external[:150]}", flush=True)
        print(f"  Internal : {internal[:150]}", flush=True)

        if page.url != MDA_URL:
            try:
                page.go_back(wait_until="domcontentloaded", timeout=30_000)
            except Exception as e:
                print(f"  [DEBUG] Browser back failed: {e}", flush=True)

            frame = find_frame(page, timeout=30)
            if not frame:
                page.goto(MDA_URL, wait_until="domcontentloaded", timeout=60_000)
                frame = find_frame(page, timeout=30)
                if not frame:
                    print("  ERROR: Could not restore MDA frame.", flush=True)
                    break
        else:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass

        time.sleep(0.5)

    return rows


# ========================= EXCEL =========================

def make_excel(rows):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "Top_20_Endpoint_Response_Time_Greater_Than_5s.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Response Time > 5s"

    headers = [
        "#", "Endpoint", "Service name", "Requests",
        "Response time", "Reason(External)", "Reason(Internal)",
    ]
    ws.append(headers)

    for row in rows:
        ws.append([
            row["#"], row["Endpoint"], row["Service name"], row["Requests"],
            row["Response time"], row["Reason(External)"], row["Reason(Internal)"],
        ])

    red_fill   = PatternFill("solid", fgColor="C00000")
    white_font = Font(color="FFFFFF", bold=True, name="Arial", size=10)
    side       = Side(style="thin", color="000000")
    border     = Border(left=side, right=side, top=side, bottom=side)

    for cell in ws[1]:
        cell.fill      = red_fill
        cell.font      = white_font
        cell.border    = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_cells in ws.iter_rows(min_row=2):
        for cell in row_cells:
            cell.border = border
            cell.font   = Font(name="Arial", size=10)
            if cell.column >= 6:
                cell.alignment = Alignment(
                    horizontal="left", vertical="top", wrap_text=True
                )
            else:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

    for col, width in zip("ABCDEFG", [5, 42, 34, 9, 14, 60, 55]):
        ws.column_dimensions[col].width = width

    for row_cells in ws.iter_rows(min_row=2):
        ws.row_dimensions[row_cells[0].row].height = 45

    ws.freeze_panes         = "A2"
    ws.sheet_view.showGridLines = False
    wb.save(path)
    return path


# ========================= WORD =========================

def make_word(rows):
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor

    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "Top_20_Endpoint_Response_Time_Greater_Than_5s.docx")

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.4)
    sec.left_margin = sec.right_margin = Inches(0.4)

    p   = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Top 20 Endpoint Response Time Analysis")
    run.bold      = True
    run.font.size = Pt(15)

    doc.add_paragraph(
        f"Average Response Time > 5 s  |  "
        f"Excluded: schedule, terminateSessionAndUpdateLead  |  "
        f"Generated: {datetime.now(IST).strftime('%d-%b-%Y %H:%M:%S')} IST"
    ).paragraph_format.space_after = Pt(6)

    headers = [
        "#", "Endpoint", "Service name", "Requests",
        "Response time", "Reason(External)", "Reason(Internal)",
    ]

    table = doc.add_table(rows=1, cols=len(headers))
    table.style     = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit   = False

    def shade_red(cell):
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "C00000")
        cell._tc.get_or_add_tcPr().append(shd)

    def write_cell(cell, value, bold=False, white=False, align="center"):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = (
            WD_ALIGN_PARAGRAPH.LEFT if align == "left"
            else WD_ALIGN_PARAGRAPH.CENTER
        )
        r = p.add_run(str(value))
        r.bold      = bold
        r.font.name = "Arial"
        r.font.size = Pt(7)
        if white:
            r.font.color.rgb = RGBColor(255, 255, 255)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for i, h in enumerate(headers):
        shade_red(table.rows[0].cells[i])
        write_cell(table.rows[0].cells[i], h, bold=True, white=True)

    for row in rows:
        cells = table.add_row().cells
        vals  = [
            row["#"], row["Endpoint"], row["Service name"], row["Requests"],
            row["Response time"], row["Reason(External)"], row["Reason(Internal)"],
        ]
        for i, v in enumerate(vals):
            write_cell(cells[i], v, align="left" if i >= 5 else "center")

    for i, w in enumerate([0.3, 2.8, 2.0, 0.65, 0.95, 2.5, 2.5]):
        for table_row in table.rows:
            table_row.cells[i].width = Inches(w)

    doc.save(path)
    return path


# ========================= MAIN =========================

def main():
    if os.path.exists(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page    = context.new_page()

        try:
            print("Opening Dynatrace MDA...", flush=True)
            page.goto(MDA_URL, wait_until="domcontentloaded", timeout=90_000)

            frame = find_frame(page)
            if not frame:
                raise RuntimeError("Could not find the Dynatrace MDA table frame.")
            print("MDA table found.", flush=True)

            refine(frame)
            rows = get_top_n(frame)

            if rows:
                rows = fetch_reasons(frame, page, rows)
            else:
                print("No rows — skipping reason fetch.", flush=True)

            excel_path = make_excel(rows)
            word_path  = make_word(rows)

            print("\n========== COMPLETED ==========", flush=True)
            print(f"Rows in report : {len(rows)}", flush=True)
            print(f"Excel          : {excel_path}", flush=True)
            print(f"Word           : {word_path}", flush=True)

        finally:
            browser.close()


if __name__ == "__main__":
    main()