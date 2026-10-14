#!/usr/bin/env python3
"""
Booking.com Dubai hotel scraper
================================

What this script does
---------------------
1. Fetches the FIRST page of search results through the Crawio API
   (optional - skipped with a clear message if no API key is set).
2. Opens the same search URL in a separate Chromium browser (Playwright),
   directly against Booking.com, and clicks "Load more results" up to
   MAX_LOAD_MORE_CLICKS times, saving an HTML snapshot after every step.
   NOTE: Crawio only fetches the initial page. The Playwright browser is a
   completely separate session and does NOT go through Crawio.
3. Parses every HTML (Crawio + each snapshot) with BeautifulSoup.
4. Combines the results and removes duplicates.
5. Saves everything to booking_dubai_hotels.xlsx.

Setup
-----
    pip install requests beautifulsoup4 pandas openpyxl playwright
    playwright install chromium

Run
---
    # Optional: set your Crawio key (the environment variable wins over the
    # CRAWIO_API_KEY placeholder further down in this file).
    #   Windows (PowerShell):  $env:CRAWIO_API_KEY = "your-key"
    #   macOS / Linux:         export CRAWIO_API_KEY="your-key"
    python booking_dubai_scraper.py

This script does not try to bypass CAPTCHAs or any other access restriction.
If Booking.com does not show hotel cards, the script says so and stops.
"""

import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

TARGET_URL = (
    "https://www.booking.com/searchresults.html"
    "?ss=Dubai%2C+United+Arab+Emirates&dest_id=-782831&dest_type=city"
    "&group_adults=2&no_rooms=1&group_children=0&lang=en-us"
)

# Crawio API. The CRAWIO_API_KEY environment variable takes priority.
# Never paste a real key into a file you share or upload.
CRAWIO_API_KEY = "CRAWIO_API_KEY"
CRAWIO_ENDPOINT = "https://api.crawio.com/scrape"
CRAWIO_TIMEOUT_SECONDS = 200

# Playwright settings
HEADLESS = False                 # False = you can watch the browser (good for tutorials)
MAX_LOAD_MORE_CLICKS = 10        # hard limit, so the loop can never run forever
PAGE_LOAD_TIMEOUT_MS = 60_000
CARDS_WAIT_TIMEOUT_MS = 30_000
NEW_CARDS_WAIT_TIMEOUT_MS = 15_000
BUTTON_SEARCH_SECONDS = 25       # how long to scroll around looking for the button

# Output files
EXCEL_FILE = "booking_dubai_hotels.xlsx"
SNAPSHOT_DIR = Path("booking_snapshots")
COLUMNS = [
    "Hotel Name",
    "Review Label",
    "Review Score",
    "Review Count",
    "Description",
    "Image URL",
    "Hotel URL",
]

# Selectors
CARD_SELECTOR = 'div[data-testid="property-card"]'
LOAD_MORE_SELECTOR = "button:has-text('Load more results')"

BOOKING_BASE_URL = "https://www.booking.com"


# ---------------------------------------------------------------------------
# PART 1: CRAWIO API (initial page only)
# ---------------------------------------------------------------------------

def get_crawio_api_key():
    """Return the Crawio API key, or "" if none is configured.

    The environment variable wins. The placeholder text counts as "not set".
    """
    key = os.environ.get("CRAWIO_API_KEY", "").strip() or CRAWIO_API_KEY.strip()
    if not key or key == "PASTE_YOUR_API_KEY_HERE":
        return ""
    return key


def looks_like_html(text):
    """Quick check: does this string look like an HTML document?"""
    if not isinstance(text, str) or len(text) < 200:
        return False
    lowered = text.lower()
    return "<html" in lowered or "<div" in lowered or "<body" in lowered


def find_html_in_json(data):
    """Search parsed JSON for an HTML string.

    Common keys are checked first (html, content, body, data, result ...),
    then the search goes deeper into nested dicts and lists.
    """
    preferred_keys = ["html", "content", "body", "data", "result",
                      "response", "page", "text", "source"]

    if isinstance(data, str):
        return data if looks_like_html(data) else None

    if isinstance(data, dict):
        for key in preferred_keys:
            if key in data:
                found = find_html_in_json(data[key])
                if found:
                    return found
        for value in data.values():
            found = find_html_in_json(value)
            if found:
                return found

    if isinstance(data, list):
        for item in data:
            found = find_html_in_json(item)
            if found:
                return found

    return None


def fetch_html_with_crawio(url, api_key):
    """Fetch one page through the Crawio API.

    Returns the HTML string, or None if anything goes wrong (a message is
    printed so the user knows why).
    """
    if not api_key:
        print("[Crawio] No API key set - skipping Crawio and using Playwright only.")
        print("         (Set the CRAWIO_API_KEY environment variable to enable it.)")
        return None

    print("[Crawio] Sending request for the initial page (can take up to 200 seconds)...")
    headers = {"x-api-key": api_key, "Content-Type": "application/json"}
    payload = {"url": url, "method": "GET"}

    try:
        response = requests.post(
            CRAWIO_ENDPOINT,
            headers=headers,
            json=payload,
            timeout=CRAWIO_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
    except requests.exceptions.Timeout:
        print(f"[Crawio] The request timed out after {CRAWIO_TIMEOUT_SECONDS} seconds.")
        return None
    except requests.exceptions.HTTPError as error:
        status = error.response.status_code if error.response is not None else "?"
        print(f"[Crawio] HTTP error {status}: {error}")
        return None
    except requests.exceptions.RequestException as error:
        print(f"[Crawio] Network error: {error}")
        return None

    # Try JSON first.
    try:
        data = response.json()
    except ValueError:
        # Not JSON - maybe the API returned raw HTML.
        if looks_like_html(response.text):
            print("[Crawio] Response was raw HTML (not JSON) - using it directly.")
            return response.text
        print("[Crawio] The response was not valid JSON and does not look like HTML.")
        return None

    html = find_html_in_json(data)
    if not html:
        print("[Crawio] The JSON response did not contain any HTML.")
        return None

    print(f"[Crawio] Success - received {len(html):,} characters of HTML.")
    return html


# ---------------------------------------------------------------------------
# PART 2: HOTEL EXTRACTION (BeautifulSoup)
# ---------------------------------------------------------------------------

def clean_text(text):
    """Collapse whitespace and trim a string."""
    return re.sub(r"\s+", " ", text or "").strip()


def normalize_hotel_url(href):
    """Turn a hotel link into an absolute https URL without query/fragment."""
    if not href:
        return ""
    absolute = urljoin(BOOKING_BASE_URL + "/", href.strip())
    parts = urlsplit(absolute)
    if not parts.netloc:
        return ""
    return f"https://{parts.netloc}{parts.path}"


def to_float(text):
    """Pull the first decimal number out of a string, e.g. 'Scored 8.4' -> 8.4."""
    match = re.search(r"\d+(?:[.,]\d+)?", text or "")
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", "."))
    except ValueError:
        return None


def to_int(text):
    """Pull the first whole number out of a string, e.g. '3,045 reviews' -> 3045."""
    match = re.search(r"\d[\d,\.]*", text or "")
    if not match:
        return None
    digits = re.sub(r"[^\d]", "", match.group(0))
    return int(digits) if digits else None


def parse_review_aria_label(aria_label):
    """Parse e.g. 'Scored 8.4, Very Good, 3,045 reviews, Opens ...'.

    Returns (score, label, count); any value may be None.
    """
    match = re.search(
        r"Scored\s+([\d.,]+)\s*,\s*(.+?)\s*,\s*([\d,\.]+)\s+reviews?",
        aria_label or "",
        flags=re.IGNORECASE,
    )
    if not match:
        return None, "", None
    return to_float(match.group(1)), clean_text(match.group(2)), to_int(match.group(3))


def extract_description(card, hotel_name, review_count_text):
    """Pick the best description from the card.

    div.fff1944c52 is a shared class (it is also used for review count, price,
    location ...), so we filter out anything that does not look like a
    description. If nothing reliable is found we return "".
    """
    unwanted_words = [
        "review", "scored", "aed", "per night", "night", "km from", "from centre",
        "from center", "show on map", "metro", "price", "taxes", "fees",
        "free cancellation", "no prepayment", "sign in", "genius",
    ]
    for element in card.select("div.fff1944c52"):
        text = clean_text(element.get_text(" "))
        lowered = text.lower()
        if len(text) < 30:                      # too short to be a description
            continue
        if text == review_count_text or text == hotel_name:
            continue
        if any(word in lowered for word in unwanted_words):
            continue
        if element.find("div", class_="fff1944c52"):   # wrapper around other blocks
            continue
        return text
    return ""


def parse_card(card):
    """Extract one hotel record (a dict) from a property card."""
    # Name
    name_el = card.select_one('[data-testid="title"]')
    name = clean_text(name_el.get_text(" ")) if name_el else ""

    # Review data from the main selectors
    label_el = card.select_one(".f63b14ab7a.f546354b44.becbee2f63")
    score_el = card.select_one(".f63b14ab7a.dff2e52086")
    count_el = card.select_one(".fff1944c52.fb14de7f14.eaa8455879")

    review_label = clean_text(label_el.get_text(" ")) if label_el else ""
    review_score = to_float(score_el.get_text(" ")) if score_el else None
    count_text = clean_text(count_el.get_text(" ")) if count_el else ""
    review_count = to_int(count_text) if count_el else None

    # Fallback: aria-label on the review link
    if review_score is None or review_count is None or not review_label:
        review_link = card.select_one('[data-testid="review-score-link"]')
        if review_link and review_link.get("aria-label"):
            aria_score, aria_label, aria_count = parse_review_aria_label(
                review_link["aria-label"]
            )
            if review_score is None:
                review_score = aria_score
            if not review_label:
                review_label = aria_label
            if review_count is None:
                review_count = aria_count

    description = extract_description(card, name, count_text)

    image_el = card.select_one('img[data-testid="image"]')
    image_url = (image_el.get("src") or "").strip() if image_el else ""

    link_el = card.select_one('a[data-testid="title-link"]')
    hotel_url = normalize_hotel_url(link_el.get("href")) if link_el else ""

    return {
        "Hotel Name": name,
        "Review Label": review_label,
        "Review Score": review_score,
        "Review Count": review_count,
        "Description": description,
        "Image URL": image_url,
        "Hotel URL": hotel_url,
    }


def parse_hotels(html):
    """Parse a full HTML page and return a list of hotel dicts."""
    soup = BeautifulSoup(html, "html.parser")
    hotels = []
    for card in soup.select(CARD_SELECTOR):
        try:
            record = parse_card(card)
        except Exception as error:  # one broken card should not stop everything
            print(f"   [Parser] Skipped one card because of an error: {error}")
            continue
        if record["Hotel Name"] or record["Hotel URL"]:
            hotels.append(record)
    return hotels


# ---------------------------------SSS------------------------------------------
# PART 3: "LOAD MORE" PAGINATION (Playwright)
# ---------------------------------------------------------------------------

def count_cards(page):
    """Number of hotel cards currently on the page."""
    try:
        return page.locator(CARD_SELECTOR).count()
    except PlaywrightError:
        return 0


def save_snapshot(page, name):
    """Save the current page HTML to a file and return the HTML string."""
    SNAPSHOT_DIR.mkdir(exist_ok=True)
    html = page.content()
    (SNAPSHOT_DIR / name).write_text(html, encoding="utf-8")
    return html


def dismiss_popups(page):
    """Close the cookie banner (reject non-essential) and the sign-in popup.

    Every step is optional: if a popup is not there we simply move on.
    """
    # OneTrust cookie banner -> reject non-essential cookies
    try:
        reject = page.locator("#onetrust-reject-all-handler")
        if reject.count() > 0 and reject.first.is_visible():
            reject.first.click(timeout=3000)
            print("   Closed cookie banner (rejected non-essential cookies).")
            page.wait_for_timeout(500)
    except PlaywrightError:
        pass

    # Sign-in / Genius popup
    try:
        dismiss = page.locator(
            'button[aria-label*="Dismiss" i], button[aria-label="Close" i]'
        )
        if dismiss.count() > 0 and dismiss.first.is_visible():
            dismiss.first.click(timeout=3000)
            print("   Closed a popup.")
            page.wait_for_timeout(500)
    except PlaywrightError:
        pass


def find_load_more_button(page):
    """Scroll up and down looking for the 'Load more results' button.

    Search pattern: bottom -> up a little -> bottom again, repeated for up to
    BUTTON_SEARCH_SECONDS. Returns a Playwright locator or None.
    """
    deadline = time.time() + BUTTON_SEARCH_SECONDS
    while time.time() < deadline:
        # Main text-based selector, then a broader role/text fallback.
        candidates = [
            page.locator(LOAD_MORE_SELECTOR),
            page.get_by_role("button", name=re.compile(r"load more|show more results", re.I)),
            page.locator("button, a").filter(has_text=re.compile(r"load more", re.I)),
        ]
        for candidate in candidates:
            try:
                if candidate.count() > 0 and candidate.first.is_visible():
                    return candidate.first
            except PlaywrightError:
                continue

        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(800)
            page.evaluate("window.scrollBy(0, -700)")
            page.wait_for_timeout(600)
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(800)
        except PlaywrightError:
            break
    return None


def click_with_fallbacks(button):
    """Try a normal click, then a forced click, then a JavaScript click."""
    try:
        button.scroll_into_view_if_needed(timeout=5000)
    except PlaywrightError:
        pass

    try:
        button.click(timeout=5000)
        return True
    except PlaywrightError:
        print("   Normal click failed, trying a forced click...")

    try:
        button.click(timeout=5000, force=True)
        return True
    except PlaywrightError:
        print("   Forced click failed, trying a JavaScript click...")

    try:
        button.evaluate("element => element.click()")
        return True
    except PlaywrightError as error:
        print(f"   JavaScript click failed too: {error}")
        return False


def wait_for_more_cards(page, previous_count):
    """Wait until the card count goes above previous_count.

    If it does not grow, scroll a bit and try once more.
    Returns the new card count (may equal previous_count).
    """
    script = (
        "(n) => document.querySelectorAll('div[data-testid=\"property-card\"]')"
        ".length > n"
    )
    for attempt in (1, 2):
        try:
            page.wait_for_function(script, arg=previous_count,
                                   timeout=NEW_CARDS_WAIT_TIMEOUT_MS)
            break
        except PlaywrightTimeoutError:
            if attempt == 1:
                print("   No new cards yet - scrolling and retrying once...")
                try:
                    page.evaluate("window.scrollBy(0, -500)")
                    page.wait_for_timeout(700)
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                except PlaywrightError:
                    pass
    page.wait_for_timeout(1000)  # let the page settle
    return count_cards(page)


def print_diagnostics(page):
    """Print useful information about the page state when we stop early."""
    print("   --- Diagnostics ---")
    try:
        position = page.evaluate(
            "() => ({y: Math.round(window.scrollY), "
            "height: document.body.scrollHeight, view: window.innerHeight})"
        )
        print(f"   Scroll position: {position['y']} / page height {position['height']} "
              f"(viewport {position['view']})")
    except PlaywrightError:
        print("   Scroll position: unavailable")
    print(f"   Cards on page: {count_cards(page)}")

    try:
        related = page.locator("button, a").filter(
            has_text=re.compile(r"load|more|show|result", re.I)
        )
        total = min(related.count(), 15)
        print(f"   Related buttons/links found: {related.count()} (showing {total})")
        for i in range(total):
            text = clean_text(related.nth(i).inner_text(timeout=2000))
            if text:
                print(f"     - {text[:80]}")
    except PlaywrightError:
        print("   Could not read related buttons/links.")


def save_debug_files(page, reason):
    """Save a screenshot and the page HTML for troubleshooting."""
    SNAPSHOT_DIR.mkdir(exist_ok=True)
    try:
        page.screenshot(path=str(SNAPSHOT_DIR / "debug_screenshot.png"), full_page=False)
        (SNAPSHOT_DIR / "debug_page.html").write_text(page.content(), encoding="utf-8")
        print(f"   Debug files saved in '{SNAPSHOT_DIR}/' "
              f"(debug_screenshot.png, debug_page.html). Reason: {reason}")
    except PlaywrightError as error:
        print(f"   Could not save debug files: {error}")


def scrape_with_playwright(url):
    """Open Booking.com in Chromium and click 'Load more results' repeatedly.

    This is a separate browser session that talks directly to Booking.com
    (it does NOT use Crawio).

    Returns a list of (label, html) snapshots - one for the initial page and
    one after every successful click.
    """
    snapshots = []

    try:
        with sync_playwright() as playwright:
            print("[Playwright] Starting Chromium...")
            try:
                browser = playwright.chromium.launch(headless=HEADLESS)
            except PlaywrightError as error:
                print(f"[Playwright] Could not start Chromium: {error}")
                print("             Did you run:  playwright install chromium ?")
                return snapshots

            context = browser.new_context(
                viewport={"width": 1366, "height": 900},
                locale="en-US",
            )
            page = context.new_page()

            try:
                print("[Playwright] Opening the search page...")
                try:
                    page.goto(url, wait_until="domcontentloaded",
                              timeout=PAGE_LOAD_TIMEOUT_MS)
                except PlaywrightTimeoutError:
                    print("[Playwright] The page took too long to load.")
                    return snapshots
                except PlaywrightError as error:
                    print(f"[Playwright] Navigation error: {error}")
                    return snapshots

                dismiss_popups(page)

                try:
                    page.wait_for_selector(CARD_SELECTOR, timeout=CARDS_WAIT_TIMEOUT_MS)
                except PlaywrightTimeoutError:
                    print("[Playwright] No hotel cards appeared. The page may be showing "
                          "a verification/CAPTCHA screen or a different layout.")
                    print("             This script does not try to bypass access "
                          "restrictions, so it is stopping here.")
                    save_debug_files(page, "no property cards found")
                    return snapshots

                dismiss_popups(page)
                initial_count = count_cards(page)
                print(f"[Playwright] Initial card count: {initial_count}")
                html = save_snapshot(page, "snapshot_00_initial.html")
                snapshots.append(("Playwright initial page", html))

                clicks_done = 0
                stop_reason = ""

                for click_number in range(1, MAX_LOAD_MORE_CLICKS + 1):
                    print(f"\n[Load more] Click {click_number} of up to {MAX_LOAD_MORE_CLICKS}")
                    dismiss_popups(page)

                    cards_before = count_cards(page)
                    button = find_load_more_button(page)
                    if button is None:
                        stop_reason = "the 'Load more results' button is gone"
                        print("   Button not found - no more results to load.")
                        print_diagnostics(page)
                        save_debug_files(page, stop_reason)
                        break

                    if not click_with_fallbacks(button):
                        stop_reason = "the button could not be clicked"
                        print_diagnostics(page)
                        save_debug_files(page, stop_reason)
                        break

                    cards_after = wait_for_more_cards(page, cards_before)
                    print(f"   Cards before: {cards_before} -> after: {cards_after}")

                    if cards_after <= cards_before:
                        stop_reason = "no new cards appeared after the click"
                        print("   No new cards loaded - stopping.")
                        print_diagnostics(page)
                        save_debug_files(page, stop_reason)
                        break

                    clicks_done += 1
                    html = save_snapshot(page, f"snapshot_{click_number:02d}.html")
                    snapshots.append((f"Playwright after click {click_number}", html))
                else:
                    stop_reason = f"reached the click limit ({MAX_LOAD_MORE_CLICKS})"

                print(f"\n[Playwright] Finished after {clicks_done} successful click(s): "
                      f"{stop_reason}.")

            except PlaywrightTimeoutError as error:
                print(f"[Playwright] Timed out: {error}")
            except PlaywrightError as error:
                print(f"[Playwright] Browser error: {error}")
            finally:
                try:
                    browser.close()
                except PlaywrightError:
                    pass

    except Exception as error:  # last-resort safety net
        print(f"[Playwright] Unexpected error: {error}")

    return snapshots


# ---------------------------------------------------------------------------
# PART 4: COMBINE AND CLEAN
# ---------------------------------------------------------------------------

def hotel_key(hotel):
    """Unique key for a hotel: normalized URL, or lowercase name as fallback."""
    return hotel["Hotel URL"] or hotel["Hotel Name"].strip().lower()


def is_empty(value):
    """True for None, empty strings, and NaN."""
    return value is None or value == "" or (isinstance(value, float) and pd.isna(value))


def merge_hotels(hotel_lists):
    """Combine several lists of hotels into one record per hotel.

    Fields that are empty in an earlier record are filled from later ones.
    The original order (first time each hotel was seen) is preserved.
    """
    merged = {}
    for hotels in hotel_lists:
        for hotel in hotels:
            key = hotel_key(hotel)
            if not key:
                continue
            if key not in merged:
                merged[key] = dict(hotel)
            else:
                existing = merged[key]
                for column in COLUMNS:
                    if is_empty(existing.get(column)) and not is_empty(hotel.get(column)):
                        existing[column] = hotel[column]
    return list(merged.values())


# ---------------------------------------------------------------------------
# PART 5: EXCEL EXPORT
# ---------------------------------------------------------------------------

def export_to_excel(hotels, filename):
    """Write the hotels to an Excel file. Returns True on success."""
    dataframe = pd.DataFrame(hotels, columns=COLUMNS)
    widths = {
        "Hotel Name": 40, "Review Label": 16, "Review Score": 14,
        "Review Count": 14, "Description": 70, "Image URL": 50, "Hotel URL": 60,
    }
    try:
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            dataframe.to_excel(writer, index=False, sheet_name="Dubai Hotels")
            sheet = writer.sheets["Dubai Hotels"]
            for position, column in enumerate(COLUMNS, start=1):
                letter = sheet.cell(row=1, column=position).column_letter
                sheet.column_dimensions[letter].width = widths[column]
            sheet.freeze_panes = "A2"
        return True
    except PermissionError:
        print(f"[Excel] Cannot write '{filename}'. If it is open in Excel, close it "
              f"and run the script again.")
        return False
    except OSError as error:
        print(f"[Excel] Could not save the file: {error}")
        return False


# ---------------------------------------------------------------------------
# PART 6: TERMINAL OUTPUT HELPERS
# ---------------------------------------------------------------------------

def print_preview(hotels, limit=10):
    """Print the first few records in a readable format."""
    print(f"\nPreview of the first {min(limit, len(hotels))} hotels:")
    print("-" * 70)
    for number, hotel in enumerate(hotels[:limit], start=1):
        description = hotel["Description"]
        if len(description) > 90:
            description = description[:87] + "..."
        print(f"{number:>2}. {hotel['Hotel Name'] or '(no name)'}")
        print(f"    Score: {hotel['Review Score']}  |  Label: {hotel['Review Label'] or '-'}"
              f"  |  Reviews: {hotel['Review Count']}")
        print(f"    URL: {hotel['Hotel URL'] or '-'}")
        if description:
            print(f"    {description}")
    print("-" * 70)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    """Run the whole workflow from start to finish."""
    print("=" * 70)
    print("Booking.com Dubai hotel scraper")
    print("=" * 70)

    hotel_lists = []

    # Step 1: Crawio (initial page only)
    print("\nSTEP 1: Initial page via Crawio")
    crawio_html = fetch_html_with_crawio(TARGET_URL, get_crawio_api_key())
    if crawio_html:
        crawio_hotels = parse_hotels(crawio_html)
        print(f"[Crawio] Hotels extracted: {len(crawio_hotels)}")
        hotel_lists.append(crawio_hotels)
    else:
        print("[Crawio] Continuing with Playwright only.")

    # Step 2: Playwright (separate browser session, direct to Booking.com)
    print("\nSTEP 2: Load more results via Playwright (direct browser session)")
    snapshots = scrape_with_playwright(TARGET_URL)

    # Step 3: Parse every snapshot
    if snapshots:
        print("\nSTEP 3: Parsing snapshots")
        for label, html in snapshots:
            snapshot_hotels = parse_hotels(html)
            print(f"   {label}: {len(snapshot_hotels)} hotels")
            hotel_lists.append(snapshot_hotels)

    # Step 4: Combine
    unique_hotels = merge_hotels(hotel_lists)
    print(f"\nTotal unique hotels: {len(unique_hotels)}")

    if not unique_hotels:
        print("No hotels were collected, so no Excel file was created.")
        print("Check the debug files in the snapshot folder and your Crawio key.")
        return 1

    # Step 5: Excel
    print("\nSTEP 4: Saving to Excel")
    if export_to_excel(unique_hotels, EXCEL_FILE):
        print(f"[Excel] Saved {len(unique_hotels)} hotels to: {EXCEL_FILE}")
    else:
        return 1

    print_preview(unique_hotels)
    print("\nDone.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nStopped by user.")
        sys.exit(130)