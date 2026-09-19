"""a58 Phase 3 - competitor news monitoring via GDELT, using the gdeltdoc
library (https://pypi.org/project/gdeltdoc/) instead of hand-rolled
requests/subprocess calls.

Ported from Session 2's build_gdelt_news.py (smoke-tested, unmodified logic
carried over below) onto gdeltdoc for cleaner HTTP/JSON handling - notably
gdeltdoc.errors.RateLimitError is a specific 429 exception rather than
string-matching response text, and its JSON parser has a max-depth
illegal-character strip that's more robust than a bare json.loads.

IMPORTANT - query construction: gdeltdoc's Filters class has no documented
way to express "company name AND (phrase1 OR phrase2 OR phrase3)" - its
`keyword` parameter only supports a single phrase OR a flat OR-list, not a
required term plus a separate OR-group. Filters.near/.repeat, however, are
typed as plain `Optional[str]` and get appended to the query verbatim with
no validation of their content (confirmed by reading gdeltdoc/filters.py
directly - `near` and `repeat` are meant to carry the output of the
near()/repeat() helpers, but nothing enforces that). We use `near=` as a
raw pass-through slot for the pre-built category OR-group, which
reproduces byte-for-byte the same query structure
(`"CompanyName" (phrase1 OR phrase2 OR phrase3)`) Session 2's script
already validated as correct - not a new, unverified query shape.

gdeltdoc has no built-in rate limiting or retry logic (confirmed against
its source - it's a thin request/parse wrapper), so the 7s throttle and
retry-on-429 discipline from the original script is reimplemented here
around gd.article_search(), not assumed away by switching libraries.

Scope: same 58-institution CFPB/a20-vetted priority list as the original
pilot (ingestion/gdelt_institutions.json). Scaling to the full
branch_target_competitors universe is an explicit non-goal here - at this
rate limit it would take hours (see module docstring precedent in the
original script).
"""

import json
import re
import sys
import time
from pathlib import Path

from gdeltdoc import Filters, GdeltDoc
from gdeltdoc.errors import RateLimitError

from ingestion.supabase_client import upsert

INSTITUTIONS_FILE = Path(__file__).parent / "gdelt_institutions.json"

CATEGORIES = {
    "branch_closing": 'branch closing OR branch closure OR closes branch OR closing branches',
    "branch_opening": 'new branch OR opens branch OR opening branch OR branch opening',
    "m_and_a": 'acquisition OR acquires OR merger OR "to acquire" OR "to merge"',
}

MIN_INTERVAL = 7.0  # GDELT enforces 5s; pad further so a retry's own wait
                     # can't land the next real call inside the window
MAX_ATTEMPTS = 6
TIMESPAN = "14d"
MAX_RECORDS = 10

_last_request_at = [0.0]
_gd = GdeltDoc()


def short_name(cfpb_name):
    """Strip legal-entity suffixes that don't appear in ordinary news prose,
    so the GDELT query matches how a reporter would actually write the
    bank's name. Identical logic to the original script."""
    n = re.sub(
        r",?\s*(NATIONAL ASSOCIATION|CORPORATION|CORP\.?|INC\.?|BANCORP|BANCSHARES,?\s*INC\.?|BANCSHARES|FSB)\.?$",
        "", cfpb_name, flags=re.I,
    ).strip()
    return n or cfpb_name


def _throttle():
    wait = MIN_INTERVAL - (time.monotonic() - _last_request_at[0])
    if wait > 0:
        time.sleep(wait)


def gdelt_search(company_name, category_term):
    """Returns a list of article dicts (title/url/seendate/domain), or None
    if every retry was rate-limited - callers must not treat None as
    "zero results"."""
    category_group = f"({category_term}) "
    filters = Filters(
        keyword=company_name,   # -> '"CompanyName" ' via gdeltdoc's own quoting
        near=category_group,    # raw pass-through - see module docstring
        timespan=TIMESPAN,
        num_records=MAX_RECORDS,
    )

    for _attempt in range(MAX_ATTEMPTS):
        _throttle()
        _last_request_at[0] = time.monotonic()
        try:
            df = _gd.article_search(filters)
        except RateLimitError:
            continue
        except Exception as e:
            print(f"    non-rate-limit error, retrying: {e}", flush=True)
            continue
        return df.to_dict("records") if not df.empty else []
    return None


def main():
    with open(INSTITUTIONS_FILE, encoding="utf-8") as f:
        institutions = json.load(f)

    rows = []
    gave_up = []
    n = len(institutions)
    for i, inst in enumerate(institutions):
        inst_key = inst["inst_key"]
        name = short_name(inst["cfpb_company_name"])
        for category, term in CATEGORIES.items():
            articles = gdelt_search(name, term)
            if articles is None:
                gave_up.append((inst_key, category))
                print(f"  [{i + 1}/{n}] {name} / {category}: GAVE UP (rate-limited after retries)", flush=True)
                continue

            # Noise filter: GDELT's DOC API matches keywords anywhere in the
            # full article text, so a boilerplate roundup article that lists
            # the bank in a ticker table matches even though the bank is
            # never the subject. Requiring the bank's name in the TITLE
            # itself is a cheap, effective filter - unchanged from the
            # original script, which confirmed this was necessary via a
            # live false-positive (BANK OZK matched to an unrelated 13F
            # filing roundup article).
            for a in articles:
                title = a.get("title") or ""
                if name.lower() not in title.lower():
                    continue
                rows.append({
                    "inst_key": inst_key,
                    "institution_name": name,
                    "category": category,
                    "article_url": a.get("url"),
                    "title": title,
                    "seendate": a.get("seendate"),
                    "source_domain": a.get("domain"),
                })

        if (i + 1) % 10 == 0 or i == n - 1:
            print(f"  [{i + 1}/{n}] {len(rows)} matches so far", flush=True)

    print(f"Done. {len(rows)} article matches across {n} institutions. "
          f"{len(gave_up)} (institution, category) pairs gave up after retries.")
    if gave_up:
        print(f"  Gave up on: {gave_up}")

    if not rows:
        print("No rows to write - aborting without touching raw.raw_gdelt_news.")
        return

    sent = upsert("raw_gdelt_news", rows, on_conflict="inst_key,category,article_url", schema="raw", batch_size=500)
    print(f"Upserted {sent} rows into raw.raw_gdelt_news.")


if __name__ == "__main__":
    sys.exit(main())
