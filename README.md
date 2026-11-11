# gtr-search

A command-line script that searches UKRI's [Gateway to Research](https://gtr.ukri.org/) (GtR) for funded projects and saves the results as a CSV file for Excel. I use it to map the UK funding landscape around a research topic: which projects already exist nearby, who leads them, which companies partner on them, and which ones are new since the last search.

It uses the GtR API v3 (launched November 2026). Plain Python 3.8+, standard library only, no API key.

It comes in two forms:

- **GtR Search app (Windows)**: a desktop window for people who don't use Python. Download `GtR-Search.exe` from the [Releases](../../releases) page and double-click it; nothing needs installing.
- **`gtr_search.py`**: the command-line script, described from [Quick start](#quick-start) onwards.

## The app

Work through the numbered tabs, then click **Run search**:

1. **Search terms**: one per line.
2. **Filters**: funders, project status, start years, region, award type and value.
3. **Options**: which matches to fetch first, pages per term, sort order, co-investigators, quick mode, and earlier result files to compare against.
4. **Scoring**: the keywords used to rank results and the phrases to ignore.

**Show available values** lists the funders, award types and regions GtR has for your terms and filters, with counts, without saving anything. Each input has a short note under it, and the **Help** tab explains every option.

Results are saved as CSV files in `Documents\GtR Search` (changeable at the bottom of the window); **Open results** opens the latest one in Excel. The app also keeps its settings and lookup cache in that folder, so it remembers your last search and later runs are faster. **Stop** is safe: projects looked up so far are saved.

The first time it runs, Windows may show "Windows protected your PC", because the app is not signed by a registered publisher. Click **More info** → **Run anyway**. If your organisation blocks unsigned programs, ask IT to allow it, or run `python gtr_gui.py` instead (needs Python, with `gtr_search.py` in the same folder).

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
| `--out-dir FOLDER` | Folder for the output file (default: current folder) |
| `--refresh` | Ignore the cache and look everything up again |
| `--list-values` | Show the funders, categories, regions etc. GtR has for these terms and filters, with counts, then stop (no CSV) |

Defaults for most options can be changed at the top of the script. Run `python gtr_search.py -h` for the full list.

## Allowed values

To see the values GtR currently has for your own search, with counts, use `--list-values`. It makes one small request per term and writes no CSV:

```
python gtr_search.py "water quality" "heavy metal" --funder all --list-values
```

Add `--funder all` to see every funder; otherwise only the default funder (EPSRC) is counted. Other filters narrow the lists too, e.g. adding `--funder "Innovate UK"` shows only the categories and regions within Innovate UK.

Values are case-sensitive and must be typed as shown. Put lists in double quotes and separate items with commas, e.g. `--region "North West,Wales"`. Several values in one option mean "any of these"; different options combine as "all of these".

**`--funder`**: GtR treats this as an open list that changes over time. Current values:

| Value | Funder |
|---|---|
| `AHRC` | Arts and Humanities Research Council |
| `BBSRC` | Biotechnology and Biological Sciences Research Council |
| `EPSRC` | Engineering and Physical Sciences Research Council |
| `ESRC` | Economic and Social Research Council |
| `Innovate UK` | Innovate UK (business-led projects) |
| `MRC` | Medical Research Council |
| `NERC` | Natural Environment Research Council |
| `STFC` | Science and Technology Facilities Council |
| `Research England` | Research England |
| `Horizon Europe Guarantee` | UK guarantee funding for Horizon Europe awards |
| `all` | No funder filter (script option, not a GtR value) |

Cross-council schemes such as Future Leaders Fellowships may be listed under a funder name not shown here. To find them, search with `--funder all` and check the `funder` column.

**`--status`**: `active` (project still running), `closed` (project finished) or `both`.

**`--region`**: region of the lead organisation (UK ITL1 regions, a fixed list):

`North East`, `North West`, `Yorkshire and the Humber`, `East Midlands`, `West Midlands`, `East of England`, `London`, `South East`, `South West`, `Wales`, `Scotland`, `Northern Ireland`, `Channel Islands`, `Non-UK`, `Unknown`

**`--category`**: type of award. GtR treats this as an open list that it revises occasionally. Common values:

| Value | Typical use |
|---|---|
| `Research Grant` | Standard research council grants |
| `Fellowship` | Personal fellowships |
| `Training Grant` | Doctoral training and similar training awards |
| `Studentship` | Individual PhD studentships (numeric grant references, usually no PI) |
| `EU-Funded` | Awards funded through EU programmes |
| `Collaborative R&D` | Innovate UK collaborative projects (seen in results) |

The `category` column of any results file shows the values in use for your search.

**`--amount`**: total value of the award:

| Value | Meaning |
|---|---|
| `lt100k` | Less than £100k |
| `100k-1m` | £100k to £1m |
| `1m-10m` | £1m to £10m |
| `gte10m` | £10m or more |

**`--since` / `--until`**: project start year, from 2006 to 2036.

**`--fetch`**: which matches GtR returns first, before any local sorting:

| Value | Meaning |
|---|---|
| `relevance` | Best match to the search terms (default) |
| `recent` | Latest start date first |
| `ending` | Latest end date first |
| `value` | Highest award value first |

**`--sort`**: order of rows in the CSV:

| Value | Meaning |
|---|---|
| `score` | Highest keyword score first, then newest start date (default with scoring on) |
| `recent` | Newest start date first |
| `ending` | Latest end date first |
| `value` | Highest award value first |
| `gtr` | The order GtR returned (default with scoring off) |

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

## Building and releasing the app

The app is `gtr_gui.py`, a Tkinter window that calls `gtr_search.py` (`run()` with options from `make_args()`), so the two always behave the same. It is packaged into a single `GtR-Search.exe` with [PyInstaller](https://pyinstaller.org/).

**Automatic build (recommended):** the workflow in `.github/workflows/build-windows.yml` builds the .exe on a Windows machine at GitHub. To publish a new version:

1. Update `__version__` at the top of `gtr_search.py` (e.g. `1.1.0`) and commit.
2. Tag and push:
   ```
   git tag v1.1.0
   git push origin v1.1.0
   ```
3. After a few minutes the release appears on the Releases page with `GtR-Search.exe` attached. You can also run the workflow by hand from the **Actions** tab; the .exe is then under that run's **Artifacts**.

While the repository is private, only people with access to it can download releases.

**Local build:** on a Windows PC with Python, double-click `build_exe.bat` (or run it from a terminal). The app appears in `dist\GtR-Search.exe`.

## Files

| File | What it is |
|---|---|
| `gtr_search.py` | The search script (command line, and the engine behind the app) |
| `gtr_gui.py` | The desktop window |
| `build_exe.bat` | Builds `GtR-Search.exe` locally |
| `.github/workflows/build-windows.yml` | Builds the .exe at GitHub and attaches it to releases |

## Data licence

GtR data is published by UKRI under the [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/). API documentation: [Getting started with the GtR API](https://gtr.ukri.org/find-research-and-innovation/getting-started-with-the-gtr-api).
