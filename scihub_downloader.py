"""
Download articles from sci-hub.reatk.com using DOIs listed in an Excel file.

Usage:
    pip install -r requirements.txt
    playwright install chromium
    python scihub_downloader.py Articles_to_Download.xlsx [--out downloads] [--headless] [--limit N]
                                [--min-delay 6] [--max-delay 15] [--skip-doi-check]

For each row:
  1. cleans the DOI and checks that it really exists (doi.org handle registry)
     -> non-existent / malformed DOIs are skipped and never sent to Sci-Hub
  2. opens https://sci-hub.reatk.com/{DOI} like a person would (random pauses, mouse
     movement, scrolling)
  3. if the article loaded -> clicks save/download and stores it under the Excel filename
  4. otherwise -> skipped
Every row prints DOWNLOADED / SKIPPED / ALREADY EXISTS. Results go to results.csv and
finished files are skipped on re-run.
"""
import argparse
import csv
import random
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote

import openpyxl
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

BASE = "https://sci-hub.reatk.com/"
DOI_API = "https://doi.org/api/handles/"   # responseCode 1 = DOI exists, 100 = not found
LOAD_TIMEOUT = 4_000       # ms max to wait for the article to appear (override with --load-timeout)
DOWNLOAD_TIMEOUT = 60_000  # ms to wait for the download to start

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Elements that indicate an article PDF was loaded
PDF_SELECTOR = "embed[type*='pdf'], iframe#pdf, embed#pdf, iframe[src*='.pdf'], embed[src*='.pdf'], object[type*='pdf']"
# Sci-Hub's own "↓ save" button on the left panel
SAVE_SELECTOR = "button:has-text('save'), a:has-text('save'), [onclick*='location.href'][onclick*='download']"

DOI_RE = re.compile(r"^10\.\d{4,9}/\S+$")

# Text Sci-Hub shows when it doesn't have the article -> skip immediately instead of waiting
NOT_FOUND_RE = re.compile(
    r"article not found|not found|doesn.?t have|does not have|no article|unable to find|"
    r"couldn.?t find|could not find|not available|try to search again|404",
    re.I,
)


# ----------------------------------------------------------------------------- helpers
def pause(lo, hi):
    """Sleep a random, human-like amount of time."""
    time.sleep(random.uniform(lo, hi))


def clean_doi(raw):
    """Strip URL prefixes / whitespace so we only keep the bare DOI."""
    doi = (raw or "").strip()
    doi = re.sub(r"^(https?://)?(dx\.)?doi\.org/", "", doi, flags=re.I)
    doi = re.sub(r"^doi:\s*", "", doi, flags=re.I)
    return doi.strip()


def doi_exists(context, doi):
    """
    True  -> DOI is registered
    False -> DOI is malformed or not registered (doi.org says 'not found')
    None  -> could not check (network error); caller decides what to do
    """
    if not DOI_RE.match(doi):
        return False
    for attempt in range(2):
        try:
            resp = context.request.get(DOI_API + quote(doi, safe="/"), timeout=15_000)
            if resp.status == 404:
                return False
            if resp.ok:
                return resp.json().get("responseCode") == 1
        except Exception:
            pass
        pause(1, 2)
    return None


def human_look_around(page):
    """Move the mouse and scroll a little, the way someone skimming a page would."""
    try:
        vp = page.viewport_size or {"width": 1280, "height": 800}
        for _ in range(random.randint(1, 2)):
            page.mouse.move(
                random.randint(80, vp["width"] - 80),
                random.randint(80, vp["height"] - 80),
                steps=random.randint(5, 12),
            )
            pause(0.1, 0.3)
    except Exception:
        pass


def human_click(page, locator):
    """Hover over the element, wait a beat, then click."""
    try:
        box = locator.bounding_box()
        if box:
            page.mouse.move(
                box["x"] + box["width"] * random.uniform(0.25, 0.75),
                box["y"] + box["height"] * random.uniform(0.3, 0.7),
                steps=random.randint(6, 14),
            )
            pause(0.1, 0.3)
    except Exception:
        pass
    locator.click(delay=random.randint(30, 90))


# ----------------------------------------------------------------------------- excel
def read_rows(xlsx_path):
    ws = openpyxl.load_workbook(xlsx_path, read_only=True).active
    header = [str(c).strip() if c else "" for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    doi_i = header.index("DOI")
    name_i = next((i for i, h in enumerate(header) if h.lower().startswith("filename")), None)
    id_i = header.index("MasterID") if "MasterID" in header else None
    for n, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        doi = clean_doi(str(row[doi_i])) if row[doi_i] else ""
        fname = row[name_i] if name_i is not None and row[name_i] else None
        mid = row[id_i] if id_i is not None else n
        yield n, mid, doi, fname


def safe_name(name, fallback):
    name = name or fallback
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name if name.lower().endswith(".pdf") else name + ".pdf"


# ----------------------------------------------------------------------------- sci-hub
def article_loaded(page, timeout=LOAD_TIMEOUT):
    """
    True  -> the page shows an article PDF / save button
    False -> "not found" message detected (instant) or nothing appeared within `timeout` ms
    Polls every 100 ms so a missing article is skipped as soon as the page says so.
    """
    found = page.locator(f"{PDF_SELECTOR}, {SAVE_SELECTOR}")
    deadline = time.time() + timeout / 1000
    while True:
        try:
            if found.count():
                return True
            body = page.inner_text("body", timeout=1000)
            if NOT_FOUND_RE.search(body):
                return False
        except Exception:
            pass
        if time.time() >= deadline:
            return False
        time.sleep(0.1)


def download_article(page, context, dest: Path):
    """Returns a Playwright Download (still transferring), True (already saved) or False."""
    # 1) Preferred: click the download/save button and capture the download
    save = page.locator(SAVE_SELECTOR).first
    if save.count():
        try:
            with page.expect_download(timeout=DOWNLOAD_TIMEOUT) as dl:
                human_click(page, save)
            return dl.value  # transfer keeps running in the background; saved by finalize()
        except PWTimeout:
            pass

    # 2) Fallback: Chrome's PDF viewer toolbar lives outside the DOM, so fetch the embedded PDF URL directly
    src = page.evaluate(
        """() => {
            const el = document.querySelector('embed, iframe#pdf, iframe[src*=".pdf"], object');
            return el ? (el.src || el.data || el.getAttribute('src')) : null;
        }"""
    )
    if src:
        src = page.evaluate("(u) => new URL(u, location.href).href", src)
        resp = context.request.get(src, timeout=DOWNLOAD_TIMEOUT)
        body = resp.body()
        if resp.ok and body[:4] == b"%PDF":
            dest.write_bytes(body)
            return True
    return False


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--out", default="downloads")
    ap.add_argument("--headless", action="store_true", help="run without a visible browser (default: visible)")
    ap.add_argument("--limit", type=int, default=0, help="only process the first N rows (for testing)")
    ap.add_argument("--min-delay", type=float, default=3, help="min seconds between articles (default 3)")
    ap.add_argument("--max-delay", type=float, default=7, help="max seconds between articles (default 7)")
    ap.add_argument("--post-click-delay", type=float, nargs=2, default=[1, 2], metavar=("MIN", "MAX"),
                    help="seconds to wait after clicking download before moving on (default 1 2); "
                         "the file keeps downloading in the background")
    ap.add_argument("--load-timeout", type=float, default=LOAD_TIMEOUT / 1000,
                    help="max seconds to wait for an article to appear before skipping (default 4)")
    ap.add_argument("--skip-delay", type=float, nargs=2, default=[0.3, 0.8], metavar=("MIN", "MAX"),
                    help="seconds to wait after a not-found article (default 0.3 0.8)")
    ap.add_argument("--fast", action="store_true",
                    help="minimal waiting: 0.5-1.5 s between articles, no mouse movement, no long breaks")
    ap.add_argument("--skip-doi-check", action="store_true", help="don't verify DOIs against doi.org first")
    args = ap.parse_args()

    if args.fast:
        args.min_delay, args.max_delay = 0.5, 1.5
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    rows = list(read_rows(args.xlsx))
    if args.limit:
        rows = rows[: args.limit]

    counts = {"downloaded": 0, "skipped": 0, "exists": 0, "invalid": 0}
    seen = {}  # DOI -> first saved path, so duplicate DOIs aren't fetched twice
    log = open(out / "results.csv", "a", newline="", encoding="utf-8")
    writer = csv.writer(log)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context(
            accept_downloads=True,
            user_agent=USER_AGENT,
            viewport={"width": 1366, "height": 800},
            locale="en-US",
        )
        page = context.new_page()
        pending = {}  # in-flight download: {"dl", "dest", "tag", "mid", "doi"}

        def finalize():
            """Wait for the in-flight download (usually already done) and save it."""
            if not pending:
                return
            d = dict(pending)
            pending.clear()
            try:
                d["dl"].save_as(d["dest"])  # blocks only until the transfer completes
                counts["downloaded"] += 1
                seen[d["doi"]] = d["dest"]
                print(f"{d['tag']} -> DOWNLOADED ({d['dest'].name})")
                writer.writerow([d["mid"], d["doi"], "downloaded", d["dest"].name])
            except Exception as e:
                counts["skipped"] += 1
                msg = f"download failed: {str(e).splitlines()[0]}"
                print(f"{d['tag']} -> SKIPPED ({msg})")
                writer.writerow([d["mid"], d["doi"], "skipped", msg])
            log.flush()

        fetched_since_break = 0
        next_break = random.randint(10, 20)  # take a longer "coffee break" every so often

        for i, (row_no, mid, doi, fname) in enumerate(rows, 1):
            if row_no < 388:
                continue
            tag = f"[{i}/{len(rows)}] {mid} ({doi or 'no DOI'})"

            def record(status, detail):
                print(f"{tag} -> {status} ({detail})")
                writer.writerow([mid, doi, status.lower(), detail])
                log.flush()

            if not doi:
                counts["skipped"] += 1
                record("SKIPPED", "no DOI")
                continue

            dest = out / safe_name(fname, f"{mid}.pdf")
            if dest.exists():
                counts["exists"] += 1
                print(f"{tag} -> ALREADY EXISTS, skipping ({dest.name})")
                continue

            # same DOI already downloaded earlier in this run -> copy instead of re-fetching
            if pending and pending["doi"] == doi:
                finalize()
            if doi in seen and seen[doi].exists():
                dest.write_bytes(seen[doi].read_bytes())
                counts["downloaded"] += 1
                record("DOWNLOADED", f"duplicate DOI, copied from {seen[doi].name}")
                continue

            # ---- 1) does the DOI exist? ----
            if not args.skip_doi_check:
                exists = doi_exists(context, doi)
                if exists is False:
                    counts["invalid"] += 1
                    counts["skipped"] += 1
                    record("SKIPPED", "DOI does not exist")
                    continue
                if exists is None:
                    print(f"{tag}    (could not verify DOI, trying Sci-Hub anyway)")

            # ---- 2) open it on Sci-Hub and download ----
            got_file = False
            not_found = False
            try:
                page.goto(BASE + doi, wait_until="domcontentloaded", timeout=10_000)
                finalize()                           # previous file has been downloading meanwhile
                if article_loaded(page, args.load_timeout * 1000):
                    if not args.fast:
                        human_look_around(page)
                        pause(0.2, 0.6)
                    result = download_article(page, context, dest)
                    if result is True:
                        counts["downloaded"] += 1
                        seen[doi] = dest
                        record("DOWNLOADED", dest.name)
                        got_file = True
                    elif result:
                        pending.update(dl=result, dest=dest, tag=tag, mid=mid, doi=doi)
                        print(f"{tag} -> download started ({dest.name})")
                        got_file = True
                    else:
                        counts["skipped"] += 1
                        record("SKIPPED", "download failed")
                else:
                    counts["skipped"] += 1
                    record("SKIPPED", "article not available")
                    not_found = True
            except Exception as e:  # network error, timeout, etc.
                counts["skipped"] += 1
                record("SKIPPED", f"error: {str(e).splitlines()[0]}")

            # ---- 3) human-like wait before the next article ----
            fetched_since_break += 1
            if i < len(rows):
                if fetched_since_break >= next_break and not args.fast:
                    wait = random.uniform(15, 40)
                    print(f"    ... taking a longer break ({wait:.0f}s)")
                    fetched_since_break = 0
                    next_break = random.randint(10, 20)
                elif got_file:
                    wait = random.uniform(*args.post_click_delay) if not args.fast else random.uniform(0.3, 0.8)
                elif not_found:
                    wait = random.uniform(*args.skip_delay)
                else:
                    wait = random.uniform(args.min_delay, args.max_delay)
                time.sleep(wait)

        finalize()  # last download
        browser.close()

    log.close()
    print(
        f"\nDone. Downloaded: {counts['downloaded']} | Skipped: {counts['skipped']} "
        f"(of which non-existent DOIs: {counts['invalid']}) | Already existed: {counts['exists']}"
    )


if __name__ == "__main__":
    sys.exit(main())