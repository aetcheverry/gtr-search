"""
Search UKRI Gateway to Research (GtR) for projects and save the results as a CSV.
Uses the GtR API v3 (https://gtr.ukri.org/api/v3, launched November 2026).

Usage (Python 3.8+, standard library only):
    python gtr_search.py
    python gtr_search.py "single-molecule conductance" "host-guest sensing"
    python gtr_search.py "heavy metal sensor water" --funder "EPSRC,Innovate UK" --since 2018
    python gtr_search.py "water sensor" --funder "Innovate UK" --region "North West" --status active
    python gtr_search.py --compare gtr_2026-10-20_a.csv gtr_2026-10-27_b.csv

Which projects are fetched (filters are applied by GtR, so excluded projects cost nothing):
    --funder LIST     funders, comma-separated; "all" for every funder (default EPSRC)
    --status S        active, closed or both (default both)
    --since YEAR      projects starting in or after YEAR
    --until YEAR      projects starting in or before YEAR
    --region LIST     lead organisation regions, e.g. "North West,Yorkshire and the Humber"
    --category LIST   grant categories, e.g. "Research Grant,Fellowship"
    --amount LIST     value bands: lt100k, 100k-1m, 1m-10m, gte10m
    --fetch F         which matches to fetch first: relevance (default), recent, ending, value
    --pages N         result pages per term (default 2; page size PAGE_SIZE)

What happens to them:
    --sort S          CSV order: score, recent, ending, value or gtr (GtR's order)
    --score / --no-score   switch keyword scoring on or off (default USE_SCORING)
    --keywords LIST   keywords for the relevance score (whole words only)
    --exclude LIST    phrases to ignore when scoring, e.g. "gap junction"
    --no-coi          leave out co-investigators
    --quick           skip the per-project detail lookups (no abstract, co-investigators,
                      score or postcode; much faster)
    --compare FILES   add new=yes/no against earlier result files
    --out FILE        output file (default gtr_<date>_<first-term>.csv; never overwrites)
    --refresh         ignore the disk cache
    --list-values     show the funders, categories, regions etc. GtR has for these terms
                      and filters, with counts, then stop (no CSV)

Run python gtr_search.py -h for the full list. Lookups are cached in gtr_cache_v3.json.
No registration or API key is needed. GtR allows 10 requests per second; the script
stays well below that and slows down by itself if it gets HTTP 429 responses.
"""
import argparse
import csv
import datetime
import json
import os
import re
import socket
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

# ---------------------------------------------------------------- settings

DEFAULT_TERMS = [
    "single-molecule conductance",
    "molecular electronics sensor",
    "single molecule junction",
    "supramolecular sensor",
    "host-guest sensing",
    "molecular recognition electrical detection",
]
DEFAULT_FUNDERS = "EPSRC"
DEFAULT_STATUS = "both"       # active, closed or both
DEFAULT_FETCH = "relevance"   # relevance, recent, ending or value
DEFAULT_SORT = None           # None = "score" with scoring on, "gtr" with scoring off
DEFAULT_PAGES = 2
PAGE_SIZE = 100               # results per page requested from GtR (falls back to 20 if refused)
USE_SCORING = True            # set to False to skip keyword scoring by default
DEFAULT_KEYWORDS = [
    "junction", "break junction", "STM", "conductance", "host-guest",
    "supramolecular", "macrocycle", "recognition", "sensor",
]
# Phrases ignored before scoring, so their words don't count as keyword hits
# ("gap junction" no longer scores for "junction"). Plurals are covered automatically.
EXCLUDE_PHRASES = [
    "gap junction", "tight junction", "adherens junction", "neuromuscular junction",
    "cell junction", "skin conductance", "membrane conductance",
    "p-n junction", "pn junction", "Josephson junction", "magnetic tunnel junction",
]
# Person links on an award whose rel or relName matches this are treated as
# co-investigators. API v3 uses PROJECT_CO_LEAD_UK ("Project co-lead (UK)");
# the pattern also catches international co-leads and older CoI labels.
COI_PATTERN = r"co[\s_-]?lead|co[\s_-]?i"
CACHE_FILE = "gtr_cache_v3.json"
RETRIES = 4                   # extra attempts after a timeout, HTTP 429 or 5xx

# Pauses (seconds) between requests. GtR allows 10 per second; these keep well below.
# After each HTTP 429 the script also slows itself down for the rest of the run.
PAUSE_PAGE = 0.5              # between search result pages
PAUSE_DETAIL = 0.2            # after each award or organisation lookup
SLOWDOWN_STEP = 1.5           # pauses are multiplied by this after each 429 ...
SLOWDOWN_MAX = 4.0            # ... up to this many times their starting length

API = "https://gtr.ukri.org/api/v3"
SITE = "https://gtr.ukri.org"
HEADERS = {"Accept": "application/json", "User-Agent": "gtr_search.py (research landscape script)"}

STATUSES = {"active": ["Active"], "closed": ["Closed"], "both": []}
FETCH_FIELDS = {"relevance": "score", "recent": "startDate", "ending": "endDate",
                "value": "totalGrantValue"}
SORTS = ["score", "recent", "ending", "value", "gtr"]
AMOUNTS = ["lt100k", "100k-1m", "1m-10m", "gte10m"]
MAX_RESULTS = 10000           # GtR returns at most this many results per query

# Original columns first (unchanged order), newer columns at the end
COLUMNS = ["grant_ref", "title", "pi", "co_investigators", "lead_organisation", "department",
           "status", "start", "end", "value_gbp", "category", "funder", "search_term",
           "other_refs", "abstract", "gtr_link",
           "score", "matched_keywords", "new", "lead_org_town", "lead_org_region",
           "lead_org_postcode", "lead_org_country", "project_partners"]


# ---------------------------------------------------------------- network

_slow = 1.0                   # current pause multiplier (raised by 429 responses)


def pause(seconds):
    time.sleep(seconds * _slow)


def _slow_down():
    global _slow
    new = min(_slow * SLOWDOWN_STEP, SLOWDOWN_MAX)
    if new > _slow:
        _slow = new
        print(f"    Rate-limited: pauses now {_slow:.1f}x longer for the rest of this run.")


def _retry_after(err):
    try:
        return min(int(err.headers.get("Retry-After", "")), 60)
    except (TypeError, ValueError, AttributeError):
        return None


def error_detail(err):
    """GtR's error message from an HTTPError body, if there is one."""
    try:
        body = json.loads(err.read().decode("utf-8", "replace"))
        return body.get("detailedDescription") or body.get("errorMessage") or ""
    except Exception:
        return ""


def get_json(url):
    """GET a GtR URL, retrying with backoff on timeouts, 429 and 5xx."""
    delay = 2
    for attempt in range(RETRIES + 1):
        req = urllib.request.Request(url, headers=HEADERS)
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if (e.code != 429 and e.code < 500) or attempt == RETRIES:
                raise
            if e.code == 429:
                _slow_down()
            wait, msg = _retry_after(e) or delay, f"HTTP {e.code}"
        except (urllib.error.URLError, socket.timeout, TimeoutError, ConnectionError) as e:
            if attempt == RETRIES:
                raise
            wait, msg = delay, str(e)
        print(f"    {msg}; retrying in {wait}s (attempt {attempt + 2}/{RETRIES + 1})")
        time.sleep(wait)
        delay *= 2


def awards_url(term, page, size, filters, sortfield):
    params = [("query", term), ("page", page), ("size", size),
              ("sortfield", sortfield), ("sortorder", "D")]
    for name, values in filters.items():
        params += [(name, v) for v in values]      # repeated parameter = OR
    return f"{API}/awards?{urllib.parse.urlencode(params)}"


# ---------------------------------------------------------------- disk cache

class Cache:
    """Small JSON cache for award and organisation lookups. Failed lookups are not cached."""

    def __init__(self, path, refresh=False):
        self.path, self.data, self.dirty = path, {}, False
        if not refresh and os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    self.data = json.load(f)
                print(f"Loaded {len(self.data)} cached lookups from {path}")
            except (OSError, ValueError) as e:
                print(f"  Could not read cache {path} ({e}); starting afresh.")

    def get(self, key, lookup):
        if key in self.data:
            return self.data[key]
        value = lookup()
        if value is not None:
            self.data[key] = value
            self.dirty = True
        return value

    def save(self):
        if not self.dirty:
            return
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.data, f)
        os.replace(tmp, self.path)
        self.dirty = False


# ---------------------------------------------------------------- record parsing

def iso_date(v):
    """'2020-10-01T00:00:00Z' -> '2020-10-01'."""
    return str(v or "")[:10]


def person_label(name):
    name = name or {}
    return name.get("preferredName") or " ".join(filter(None, [name.get("firstName"), name.get("surname")]))


def award_row(item, term):
    """One CSV row from an award in the search results."""
    a = item.get("awardListItem") or item          # list entries wrap the award
    value = a.get("totalGrantValue", "")
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    ui = a.get("uiPath") or ""
    return {
        "_id": a.get("id", ""),
        "_lead_org_id": a.get("leadOrgId", ""),
        "_lead_person_id": a.get("leadPersonId", ""),
        "search_term": term,
        "grant_ref": a.get("grantRefNumber", ""),
        "title": a.get("projectTitle", ""),
        "pi": person_label(a.get("leadPersonName")),
        "lead_organisation": a.get("leadOrgName", ""),
        "lead_org_region": a.get("leadOrgRegion", ""),
        "department": "",                          # filled from the award detail
        "status": a.get("projectStatus", ""),
        "start": iso_date(a.get("startDate")),
        "end": iso_date(a.get("endDate")),
        "value_gbp": value,
        "category": a.get("grantCategory", ""),
        "funder": a.get("fundingOrgName", ""),
        "gtr_link": SITE + ui if ui else "",
        "abstract": "", "co_investigators": "", "lead_org_town": "",
        "lead_org_postcode": "", "lead_org_country": "", "project_partners": "",
    }


def award_detail(cache, award_id):
    """Abstract plus person and organisation links from /awards/{id} (best effort)."""
    def lookup():
        try:
            d = get_json(f"{API}/awards/{urllib.parse.quote(award_id)}")
        except Exception:
            return None
        finally:
            pause(PAUSE_DETAIL)
        extra = d.get("projectAdditionalData") or {}
        people, orgs = [], []
        for link in d.get("links") or []:
            el = link.get("entityLink") or {}
            p, o = link.get("personListItem"), link.get("organisationListItem")
            if p:
                people.append({"id": p.get("id", ""), "name": person_label(p.get("personName")),
                               "rel": el.get("rel", ""), "relName": el.get("relName", "")})
            elif o:
                orgs.append({"name": o.get("name", ""), "region": o.get("region", ""),
                             "rel": el.get("rel", "")})
        return {"abstract": " ".join((extra.get("abstractText") or "").split()),
                "people": people, "orgs": orgs}
    return cache.get(f"award2:{award_id}", lookup) or {}


def org_detail(cache, org_id):
    """Postcode and country from /organisations/{id} (best effort)."""
    def lookup():
        try:
            o = get_json(f"{API}/organisations/{urllib.parse.quote(org_id)}")
        except Exception:
            return None
        finally:
            pause(PAUSE_DETAIL)
        return {"postcode": o.get("postcode", ""), "country": o.get("country", "")}
    return cache.get(f"org:{org_id}", lookup) or {}


# ---------------------------------------------------------------- helpers

def split_list(s):
    return [x.strip() for x in (s or "").split(",") if x.strip()]


def norm_title(t):
    return " ".join((t or "").lower().split())


def slugify(s, n=40):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")[:n].strip("-") or "search"


def default_outfile(terms, using_defaults):
    date = datetime.date.today().isoformat()
    slug = "default-terms" if using_defaults else slugify(terms[0])
    if not using_defaults and len(terms) > 1:
        slug += f"-plus{len(terms) - 1}"
    name, n = f"gtr_{date}_{slug}.csv", 2
    while os.path.exists(name):
        name, n = f"gtr_{date}_{slug}_{n}.csv", n + 1
    return name


def _phrase_pattern(phrase):
    """Whole-word pattern for a keyword or phrase: plural -s/-es allowed, spaces and
    hyphens interchangeable ('break junction' = 'break-junctions'), nothing else
    ('sensor' does not match 'sensory')."""
    words = [re.escape(w) for w in re.split(r"[\s-]+", phrase.lower().strip()) if w]
    return re.compile(r"\b" + r"[\s-]+".join(words) + r"(?:s|es)?\b")


def make_scorer(keywords, exclude=()):
    """Each keyword scores 2 if in the title, 1 if in the abstract. Whole words only,
    and any excluded phrase is blanked out of the text before matching."""
    pats = [(k, _phrase_pattern(k)) for k in keywords]
    excl = [_phrase_pattern(e) for e in exclude]

    def clean(text):
        text = text.lower()
        for pat in excl:
            text = pat.sub(" ", text)
        return text

    def score(title, abstract):
        t, a = clean(title), clean(abstract)
        total, hits = 0, []
        for k, pat in pats:
            s = (2 if pat.search(t) else 0) + (1 if pat.search(a) else 0)
            if s:
                total += s
                hits.append(k)
        return total, "; ".join(hits)
    return score


def load_known(paths):
    """Grant refs and normalised titles from earlier result files."""
    refs, titles = set(), set()
    for path in paths:
        if not os.path.exists(path):
            sys.exit(f"--compare: file not found: {path}")
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            for row in csv.DictReader(f):
                for field in ("grant_ref", "other_refs"):
                    refs.update(r.strip() for r in (row.get(field) or "").split(";") if r.strip())
                if row.get("title"):
                    titles.add(norm_title(row["title"]))
    print(f"Comparing against {len(refs)} grant refs from {len(paths)} earlier file(s).")
    return refs, titles


def merge_duplicates(rows):
    """Joint grants appear once per institution with the same title. Keep one row per
    title (the first, usually the lead), and record the other grant refs and search terms."""
    merged, by_title = [], {}
    for r in rows:
        key = norm_title(r["title"])
        if key in by_title:
            keep = by_title[key]
            if r["grant_ref"] and r["grant_ref"] != keep["grant_ref"]:
                keep["other_refs"] = "; ".join(filter(None, [keep.get("other_refs"), r["grant_ref"]]))
            if r["search_term"] not in keep["search_term"]:
                keep["search_term"] += "; " + r["search_term"]
            continue
        r["other_refs"] = ""
        by_title[key] = r
        merged.append(r)
    print(f"Removed {len(rows) - len(merged)} duplicate entries (same project at partner institutions); "
          f"{len(merged)} unique projects left.")
    return merged


def sort_rows(rows, how):
    if how == "score":
        rows.sort(key=lambda r: (r["score"], r["start"]), reverse=True)
    elif how == "recent":
        rows.sort(key=lambda r: (r["start"], r["end"]), reverse=True)
    elif how == "ending":
        rows.sort(key=lambda r: (r["end"], r["start"]), reverse=True)
    elif how == "value":
        rows.sort(key=lambda r: r["value_gbp"] if isinstance(r["value_gbp"], (int, float)) else -1,
                  reverse=True)
    # "gtr": keep the order GtR returned


def parse_args():
    ap = argparse.ArgumentParser(description="Search UKRI Gateway to Research (API v3) and save a CSV.")
    ap.add_argument("terms", nargs="*", help="search terms (default: the built-in list)")
    g = ap.add_argument_group("which projects are fetched (filtered by GtR)")
    g.add_argument("--funder", default=DEFAULT_FUNDERS,
                   help='funders, comma-separated, e.g. "EPSRC,Innovate UK"; "all" for every funder')
    g.add_argument("--status", choices=list(STATUSES), default=DEFAULT_STATUS,
                   help=f"project status (default {DEFAULT_STATUS})")
    g.add_argument("--since", type=int, metavar="YEAR", help="projects starting in or after YEAR")
    g.add_argument("--until", type=int, metavar="YEAR", help="projects starting in or before YEAR")
    g.add_argument("--region", help='lead organisation regions, comma-separated, e.g. "North West"')
    g.add_argument("--category", help='grant categories, comma-separated, e.g. "Research Grant,Fellowship"')
    g.add_argument("--amount", help="value bands, comma-separated: " + ", ".join(AMOUNTS))
    g.add_argument("--fetch", choices=list(FETCH_FIELDS), default=DEFAULT_FETCH,
                   help="which matches GtR returns first (default relevance)")
    g.add_argument("--pages", type=int, default=DEFAULT_PAGES, help="pages per term")
    o = ap.add_argument_group("output")
    o.add_argument("--sort", choices=SORTS, default=DEFAULT_SORT,
                   help="CSV order (default: score with scoring on, gtr with scoring off)")
    o.add_argument("--score", dest="score", action="store_true", default=USE_SCORING,
                   help="score by keywords" + (" (default)" if USE_SCORING else ""))
    o.add_argument("--no-score", dest="score", action="store_false",
                   help="skip keyword scoring" + ("" if USE_SCORING else " (default)"))
    o.add_argument("--keywords", help="comma-separated relevance keywords")
    o.add_argument("--exclude", help="comma-separated phrases to ignore when scoring "
                                     "(replaces the EXCLUDE_PHRASES list)")
    o.add_argument("--no-coi", action="store_true", help="leave out co-investigators")
    o.add_argument("--quick", action="store_true",
                   help="skip per-project lookups: no abstract, co-investigators, score or postcode")
    o.add_argument("--compare", nargs="+", metavar="CSV", help="earlier result files to flag new projects against")
    o.add_argument("--out", help="output CSV (default: dated name that never overwrites)")
    o.add_argument("--refresh", action="store_true", help="ignore the disk cache")
    o.add_argument("--list-values", action="store_true",
                   help="list the funders, categories, regions etc. GtR has for these terms and "
                        "filters, with counts, then stop (no CSV)")
    args = ap.parse_args()

    bad = [a for a in split_list(args.amount) if a not in AMOUNTS]
    if bad:
        ap.error(f"--amount: unknown band(s) {', '.join(bad)}; use {', '.join(AMOUNTS)}")
    if args.since and args.until and args.since > args.until:
        ap.error("--since is later than --until")
    if args.quick:
        args.score, args.no_coi = False, True
    if args.sort is None:
        args.sort = "score" if args.score else "gtr"
    elif args.sort == "score" and not args.score:
        print("Note: --sort score needs keyword scoring, which is off; sorting by most recent instead.")
        args.sort = "recent"
    return args


def build_filters(args):
    filters = {}
    if args.funder.strip().lower() not in ("", "all"):
        filters["funder"] = split_list(args.funder)
    if STATUSES[args.status]:
        filters["status"] = STATUSES[args.status]
    if args.since or args.until:
        this_year = datetime.date.today().year
        first = args.since or 2006               # earliest year GtR accepts
        last = args.until or max(this_year + 5, first)
        filters["year"] = [str(y) for y in range(first, last + 1)]
    for name in ("region", "category", "amount"):
        if getattr(args, name):
            filters[name] = split_list(getattr(args, name))
    return filters


# ---------------------------------------------------------------- main

def search(terms, filters, pages, sortfield):
    global PAGE_SIZE
    seen, rows, checked = set(), [], False
    for term in terms:
        page = 1
        while page <= pages:
            try:
                data = get_json(awards_url(term, page, PAGE_SIZE, filters, sortfield))
            except urllib.error.HTTPError as e:
                detail = error_detail(e)
                if e.code == 400 and PAGE_SIZE > 20 and not checked:
                    print(f"  GtR refused page size {PAGE_SIZE} ({detail or 'HTTP 400'}); using 20 instead.")
                    PAGE_SIZE, checked = 20, True
                    continue
                if e.code != 404:
                    print(f"  '{term}' page {page}: request failed ({e}{': ' + detail if detail else ''})")
                break
            except Exception as e:
                print(f"  '{term}' page {page}: request failed ({e})")
                break
            items = data.get("awardList") or []
            paging = data.get("paging") or {}
            if not checked:
                checked = True
                got = paging.get("size")
                if got and got != PAGE_SIZE:
                    print(f"  Note: GtR returned pages of {got}, not {PAGE_SIZE}.")
                    PAGE_SIZE = got
                if items and not (items[0].get("awardListItem") or {}).get("projectTitle"):
                    print("  Unexpected response format. Keys returned:", sorted(items[0].keys()))
            if not items:
                break
            kept = 0
            for item in items:
                r = award_row(item, term)
                if not r["_id"] or r["_id"] in seen:
                    continue
                seen.add(r["_id"])
                rows.append(r)
                kept += 1
            total = paging.get("totalSize", "?")
            print(f"  '{term}' page {page}: {len(items)} projects, {kept} new kept ({total} matches in GtR)")
            # Past the last page GtR repeats the last page, so stop using its page count
            if page >= (paging.get("totalPages") or page) or page * PAGE_SIZE >= MAX_RESULTS:
                break
            page += 1
            pause(PAUSE_PAGE)
        pause(PAUSE_PAGE)
    return rows


def list_values(terms, filters):
    """Print the facet values (funder, category, region, ...) GtR reports for each search,
    summed across the search terms. One small request per term; no CSV is written."""
    totals, matches, failed = {}, 0, []
    for term in terms:
        data = None
        for size in (1, 20):                       # tiny page; 20 if GtR refuses size=1
            try:
                data = get_json(awards_url(term, 1, size, filters, "score"))
                break
            except urllib.error.HTTPError as e:
                if e.code != 400:
                    print(f"  '{term}': request failed ({e})")
                    break
            except Exception as e:
                print(f"  '{term}': request failed ({e})")
                break
        if data is None:
            failed.append(term)
            continue
        n = (data.get("paging") or {}).get("totalSize", 0) or 0
        matches += n
        print(f"  '{term}': {n} matches")
        for facet in data.get("facets") or []:
            name = facet.get("facetName") or "?"
            bucket = totals.setdefault(name, {})
            for v in facet.get("values") or []:
                key = str(v.get("value", ""))
                bucket[key] = bucket.get(key, 0) + (v.get("count") or 0)
        pause(PAUSE_PAGE)

    if not totals:
        print("GtR returned no value lists for these searches." if not failed else "No value lists retrieved.")
        return
    note = " (counts summed across terms, so a project matching two terms counts twice)" if len(terms) > 1 else ""
    print(f"\n{matches} matches in total{note}.")
    for name in sorted(totals):
        values = sorted(totals[name].items(), key=lambda kv: (-kv[1], kv[0]))
        print(f"\n{name}:")
        for value, count in values:
            print(f"  {count:>7}  {value}")
    print('\nUse these exactly as shown, e.g. --category "Research Grant" or --funder "Innovate UK".')


def main():
    args = parse_args()
    using_defaults = not args.terms
    terms = args.terms or DEFAULT_TERMS
    filters = build_filters(args)
    if args.list_values:
        shown = {k: (v if k != "year" else [f"{v[0]}-{v[-1]}"]) for k, v in filters.items()}
        print("Filters:", "; ".join(f"{k}={', '.join(v)}" for k, v in shown.items()) or "none")
        list_values(terms, filters)
        return
    keywords = split_list(args.keywords) if args.keywords else DEFAULT_KEYWORDS
    exclude = split_list(args.exclude) if args.exclude is not None else EXCLUDE_PHRASES
    scorer = make_scorer(keywords, exclude) if args.score else None
    outfile = args.out or default_outfile(terms, using_defaults)
    known = load_known(args.compare) if args.compare else None
    cache = Cache(CACHE_FILE, refresh=args.refresh)
    coi_re = re.compile(COI_PATTERN, re.I)

    shown = {k: (v if k != "year" else [f"{v[0]}-{v[-1]}"]) for k, v in filters.items()}
    print("Filters:", "; ".join(f"{k}={', '.join(v)}" for k, v in shown.items()) or "none",
          f"| fetching by {args.fetch} | sorting by {args.sort}")

    rows = search(terms, filters, args.pages, FETCH_FIELDS[args.fetch])
    if not rows:
        print("No results. If every request failed, check your internet connection.")
        return
    rows = merge_duplicates(rows)

    n = len(rows)
    print(f"Looking up abstracts, co-investigators and postcodes for {n} projects..."
          if not args.quick else f"Quick mode: no per-project lookups for {n} projects.")
    done, dropped, t0, person_rels = [], 0, time.time(), set()
    try:
        for i, r in enumerate(rows, 1):
            award_id, org_id, lead_id = r.pop("_id"), r.pop("_lead_org_id"), r.pop("_lead_person_id")
            # Backstop for the year filter (GtR filters by start year already)
            year = r["start"][:4]
            if year.isdigit() and ((args.since and int(year) < args.since) or
                                   (args.until and int(year) > args.until)):
                dropped += 1
                continue
            if not args.quick:
                d = award_detail(cache, award_id)
                r["abstract"] = d.get("abstract", "")
                people, orgs = d.get("people") or [], d.get("orgs") or []
                person_rels.update(f"{p['rel']} ({p['relName']})" for p in people)
                r["department"] = "; ".join(o["name"] for o in orgs if o["rel"] == "LEAD_RO_DEPARTMENT")
                r["project_partners"] = "; ".join(
                    o["name"] + (f" ({o['region']})" if o["region"] else "")
                    for o in orgs if o["rel"] == "PROJECT_PARTNER")
                if not args.no_coi:
                    r["co_investigators"] = "; ".join(
                        p["name"] for p in people
                        if p["id"] != lead_id and (coi_re.search(p["rel"]) or coi_re.search(p["relName"])))
                if org_id:
                    o = org_detail(cache, org_id)
                    r["lead_org_postcode"] = o.get("postcode", "")
                    r["lead_org_country"] = o.get("country", "")
            r["score"], r["matched_keywords"] = scorer(r["title"], r["abstract"]) if scorer else ("", "")
            if known:
                refs = [r["grant_ref"]] + [x.strip() for x in r["other_refs"].split(";") if x.strip()]
                is_old = any(x in known[0] for x in refs if x) or norm_title(r["title"]) in known[1]
                r["new"] = "no" if is_old else "yes"
            else:
                r["new"] = ""
            done.append(r)

            if not args.quick:
                elapsed = time.time() - t0
                remaining = elapsed / i * (n - i)
                title = r["title"][:45] + ("..." if len(r["title"]) > 45 else "")
                print(f"  [{i}/{n}] {i * 100 // n:3d}%  ~{int(remaining // 60)}m{int(remaining % 60):02d}s left  "
                      f"{r['grant_ref'] or '-':<14} {r['pi'] or '?':<22.22} "
                      f"{r['lead_organisation'] or '?':<22.22} {title}")
                if i % 25 == 0:
                    cache.save()
    except KeyboardInterrupt:
        print(f"\nStopped. Saving the {len(done)} projects looked up so far (re-runs reuse the cache).")
    finally:
        cache.save()

    sort_rows(done, args.sort)
    with open(outfile, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        w.writeheader()
        w.writerows(done)

    funders = filters.get("funder")
    msg = f"Saved {len(done)} {', '.join(funders) if funders else 'UKRI'} projects to {outfile}"
    if dropped:
        msg += f" ({dropped} dropped by the year check)"
    if known:
        msg += f"; {sum(r['new'] == 'yes' for r in done)} new since the compared files"
    print(msg + ".")
    if person_rels:
        print("  Person link types seen:", ", ".join(sorted(person_rels)))
    if not args.quick and not args.no_coi and done and not any(r["co_investigators"] for r in done):
        print("  Note: no co-investigators found. If a co-investigator type is listed above, "
              "adjust COI_PATTERN to match it.")


if __name__ == "__main__":
    main()
