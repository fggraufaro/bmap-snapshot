"""Shared mechanics for pulling products off FFIEC's public bulk-download
portal (https://cdr.ffiec.gov/public/PWS/DownloadBulkData.aspx) -- no
account or API key, but no static download URL either. It's an ASP.NET
WebForms page where the product list and the valid-years dropdown are
linked via an autopostback: DatesDropDownList is present but EMPTY on a
plain GET until the product listbox's onchange postback
(__doPostBack('ctl00$MainContentHolder$ListBox1','')) runs server-side
and repopulates it for that specific product (confirmed 2026-10-02 by
reading the page's own onchange handler, not guessed).

So every pull is a 3-step dance:
  1. GET the page for an initial __VIEWSTATE.
  2. POST simulating the ListBox1 postback -- returns a new __VIEWSTATE
     plus the now-populated year dropdown for that product.
  3. POST again with a chosen year + format + the Download button's
     name/value -- this response IS the zip file's bytes.

Also needs a browser-like User-Agent; FFIEC's WAF 403s the default
python-requests UA (confirmed 2026-10-02).
"""

import io
import time
import zipfile
from datetime import datetime

import requests
from lxml import html as lxml_html

# FFIEC's bulk files can be tens of MB; transient drops mid-download
# (IncompleteRead, connection resets) happen and aren't worth failing the
# whole pull over -- confirmed 2026-10-02 pulling the ~36MB Rank product.
_RETRYABLE = (requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError,
              requests.exceptions.Timeout)

PORTAL_URL = "https://cdr.ffiec.gov/public/PWS/DownloadBulkData.aspx"

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}


def _viewstate_fields(doc):
    def field(name):
        els = doc.xpath(f'//input[@name="{name}"]')
        return els[0].get("value", "") if els else ""
    return field("__VIEWSTATE"), field("__VIEWSTATEGENERATOR")


def _scrape_form_state(session, product_value):
    r = session.get(PORTAL_URL, headers=HEADERS, timeout=30)
    r.raise_for_status()
    doc = lxml_html.fromstring(r.text)
    viewstate, viewstategen = _viewstate_fields(doc)
    if not viewstate:
        raise RuntimeError("Could not find __VIEWSTATE on the FFIEC download page -- page structure may have changed.")

    data = {
        "__EVENTTARGET": "ctl00$MainContentHolder$ListBox1",
        "__EVENTARGUMENT": "",
        "__LASTFOCUS": "",
        "__VIEWSTATE": viewstate,
        "__VIEWSTATEGENERATOR": viewstategen,
        "ctl00$MainContentHolder$ListBox1": product_value,
    }
    r2 = session.post(PORTAL_URL, data=data, headers=HEADERS, timeout=30)
    r2.raise_for_status()
    doc2 = lxml_html.fromstring(r2.text)
    viewstate2, viewstategen2 = _viewstate_fields(doc2)

    years = doc2.xpath('//select[@name="ctl00$MainContentHolder$DatesDropDownList"]/option')
    year_options = [(o.get("value"), o.text_content().strip()) for o in years]
    if not year_options:
        raise RuntimeError("DatesDropDownList still empty after selecting the product -- FFIEC page structure may have changed.")
    return viewstate2, viewstategen2, year_options


def _pick_year_value(year_options, requested_year):
    if requested_year is None:
        return year_options[0]  # newest listed first, verified 2026-10-02
    for value, label in year_options:
        if label.strip() == str(requested_year):
            return value, label
    available = ", ".join(label for _, label in year_options)
    raise RuntimeError(f"Year {requested_year} not available. Options: {available}")


def fetch_zip(product_value, requested_year=None):
    session = requests.Session()
    viewstate, viewstategen, year_options = _scrape_form_state(session, product_value)
    value, label = _pick_year_value(year_options, requested_year)
    print(f"Pulling {product_value} for reporting year {label} (dropdown value {value})...")

    data = {
        "__LASTFOCUS": "",
        "__EVENTTARGET": "",
        "__EVENTARGUMENT": "",
        "__VIEWSTATE": viewstate,
        "__VIEWSTATEGENERATOR": viewstategen,
        "ctl00$MainContentHolder$ListBox1": product_value,
        "ctl00$MainContentHolder$DatesDropDownList": value,
        "ctl00$MainContentHolder$FormatType": "TSVRadioButton",
        "ctl00$MainContentHolder$TabStrip1$Download_0": "Download",
    }
    max_attempts = 3
    for attempt in range(1, max_attempts + 1):
        try:
            r = session.post(PORTAL_URL, data=data, headers=HEADERS, timeout=180)
            r.raise_for_status()
            content = r.content  # force full read here so a mid-stream drop raises now, not later
            break
        except _RETRYABLE as e:
            if attempt == max_attempts:
                raise
            wait = 5 * attempt
            print(f"  download error ({e!r}), retrying in {wait}s (attempt {attempt}/{max_attempts})...")
            time.sleep(wait)
    if r.headers.get("Content-Type", "").startswith("text/html"):
        raise RuntimeError("FFIEC returned an HTML page instead of a zip -- form fields likely stale/rejected.")
    print(f"  {len(content) / 1e6:.1f} MB downloaded")
    return zipfile.ZipFile(io.BytesIO(content)), label


def parse_period(text):
    # "3/31/2026 11:59:59 PM" -> date(2026, 3, 31). A real DATE, not text --
    # a text CYCLE_DATE column elsewhere in this pipeline caused a
    # lexicographic-max bug (see CLAUDE.md), not repeating that here.
    return datetime.strptime(text.strip(), "%m/%d/%Y %I:%M:%S %p").date()
