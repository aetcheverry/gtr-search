# gtr-search

A command-line script that searches UKRI's [Gateway to Research](https://gtr.ukri.org/) (GtR) for funded projects and saves the results as a CSV file for Excel. I use it to map the UK funding landscape around a research topic: which projects already exist nearby, who leads them, which companies partner on them, and which ones are new since the last search.

It uses the GtR API v3 (launched November 2026). Plain Python 3.8+, standard library only, no API key.

## Quick start

```
python gtr_search.py "molecular junction" "break junction"
```

With no search terms, it runs the built-in list in `DEFAULT_TERMS` at the top of the script. Results go to a dated file such as `gtr_2026-11-05_molecular-junction-plus1.csv`; existing files are never overwritten.

## What it does

1. Runs each search term against GtR, with the filters applied by GtR itself (funder, status, start year, region, grant category, value band).
2. Removes duplicates and merges joint grants (the same title at several institutions) into one row, listing the other grant references in `other_refs`.
3. For each project, looks up the abstract, co-investigators, lead department, project partners and lead organisation postcode.
4. Scores each project against a keyword list (2 points for a title match, 1 for an abstract match, whole words only), ignoring phrases such as "gap junction" that cause false hits.
5. Optionally flags projects not seen in earlier result files.
6. Writes the CSV, sorted by score, date or value.

## Options

Which projects are fetched (GtR applies these filters, so excluded projects cost nothing):

| Option | What it does |
|---|---|
| `--funder LIST` | Funders, comma-separated, e.g. `"EPSRC,Innovate UK"`; `all` for every funder (default EPSRC) |
| `--status S` | `active`, `closed` or `both` (default both) |
| `--since YEAR` / `--until YEAR` | Projects starting in or after / in or before a year |
| `--region LIST` | Lead organisation regions, e.g. `"North West,Yorkshire and the Humber"` |
| `--category LIST` | Grant categories, e.g. `"Research Grant,Fellowship"` |
| `--amount LIST` | Value bands: `lt100k`, `100k-1m`, `1m-10m`, `gte10m` |
| `--fetch F` | Which matches GtR returns first: `relevance` (default), `recent`, `ending`, `value` |
| `--pages N` | Result pages per term, 100 projects each (default 2) |

What happens to them:

| Option | What it does |
|---|---|
| `--sort S` | CSV order: `score`, `recent`, `ending`, `value` or `gtr` (GtR's order) |
| `--score` / `--no-score` | Switch keyword scoring on or off |
| `--keywords LIST` | Keywords for the score (replaces `DEFAULT_KEYWORDS`) |
| `--exclude LIST` | Phrases to ignore when scoring (replaces `EXCLUDE_PHRASES`) |
| `--no-coi` | Leave out co-investigators |
| `--quick` | Skip the per-project lookups (no abstract, co-investigators, score or postcode); takes seconds |
| `--compare FILES` | Add a `new` column (yes/no) against earlier result files, matched by grant reference or title |
| `--out FILE` | Output file name |
| `--refresh` | Ignore the cache and look everything up again |

Defaults for most options can be changed at the top of the script. Run `python gtr_search.py -h` for the full list.

## Examples

Active Innovate UK projects in the North West, newest first:

```
python gtr_search.py "water quality" "heavy metal" "electrochemical sensor" --funder "Innovate UK" --region "North West" --status active --sort recent
```

Every funder, projects since 2018, flagging what is new since earlier runs:

```
python gtr_search.py "single-molecule conductance" "host-guest" --funder all --since 2018 --compare old_results.csv
```

Fellowships in a field, with custom scoring keywords:

```
python gtr_search.py "spin crossover" "molecular magnetism" --funder all --category "Fellowship" --keywords "junction,spin crossover,single-molecule magnet,anisotropy"
```

## Output columns

`grant_ref, title, pi, co_investigators, lead_organisation, department, status, start, end, value_gbp, category, funder, search_term, other_refs, abstract, gtr_link, score, matched_keywords, new, lead_org_town, lead_org_region, lead_org_postcode, lead_org_country, project_partners`

The file is UTF-8 with a byte-order mark, so Excel opens accented names correctly. `lead_org_town` is always empty because API v3 does not provide it; the column is kept so the layout matches earlier result files.

## Good to know

- **Search is strict.** The v3 search matches terms much more narrowly than the old GtR search. Several short terms find more than one long one.
- **Cache.** Lookups are saved in `gtr_cache_v3.json` next to the script, so re-runs are fast. Stopping a run with Ctrl+C is safe: the cache is saved and a CSV is written with the projects done so far.
- **Rate limit.** GtR allows 10 requests per second. The script stays well below that, retries on timeouts and server errors, and slows down by itself after an HTTP 429 response.
- **Data gaps.** Some records list no co-investigators even where the grant had them, and studentships (numeric references) have no PI.
- **Co-investigators.** API v3 labels them `PROJECT_CO_LEAD_UK`. If GtR changes the label, the script prints the person link types it sees at the end of each run; adjust `COI_PATTERN` to match.

## Data licence

GtR data is published by UKRI under the [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/). API documentation: [Getting started with the GtR API](https://gtr.ukri.org/find-research-and-innovation/getting-started-with-the-gtr-api).
