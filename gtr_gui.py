"""
GtR Search – desktop window for gtr_search.py.

Searches UKRI Gateway to Research and saves the results as a CSV file for Excel,
without using the command line. Run with:  python gtr_gui.py
(or use the packaged "GtR Search.exe", which needs no Python).

Settings, the lookup cache and (by default) the results go in Documents\\GtR Search.
Needs gtr_search.py in the same folder. Standard library only (Tkinter).
"""
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import gtr_search as gs

APP_NAME = "GtR Search"
FUNDERS = [
    ("EPSRC", "Engineering and Physical Sciences"),
    ("BBSRC", "Biotechnology and Biological Sciences"),
    ("MRC", "Medical Research Council"),
    ("NERC", "Natural Environment"),
    ("STFC", "Science and Technology Facilities"),
    ("AHRC", "Arts and Humanities"),
    ("ESRC", "Economic and Social"),
    ("Innovate UK", "business-led projects"),
    ("Research England", ""),
    ("Horizon Europe Guarantee", "UK guarantee for Horizon Europe"),
]
REGIONS = ["North East", "North West", "Yorkshire and the Humber", "East Midlands",
           "West Midlands", "East of England", "London", "South East", "South West",
           "Wales", "Scotland", "Northern Ireland", "Channel Islands", "Non-UK", "Unknown"]
CATEGORIES = ["Research Grant", "Fellowship", "Training Grant", "Studentship",
              "EU-Funded", "Collaborative R&D"]
AMOUNTS = [("lt100k", "Less than £100k"), ("100k-1m", "£100k to £1m"),
           ("1m-10m", "£1m to £10m"), ("gte10m", "£10m or more")]
STATUS_LABELS = [("both", "Active and closed"), ("active", "Active only"), ("closed", "Closed only")]
FETCH_LABELS = {"Most relevant first": "relevance", "Newest start date first": "recent",
                "Latest end date first": "ending", "Highest value first": "value"}
SORT_LABELS = {"Keyword score, highest first": "score", "Newest start date first": "recent",
               "Latest end date first": "ending", "Highest value first": "value",
               "GtR's own order": "gtr"}
YEARS = [""] + [str(y) for y in range(2006, 2037)]
HINT_COLOUR = "#5f6368"
HELP_HEADINGS = ["How to use GtR Search", "1  Search terms", "2  Filters", "3  Options", "4  Scoring",
                 "Available values", "Results", "Good to know"]

HELP_TEXT = """How to use GtR Search

GtR Search looks up funded projects on UKRI's Gateway to Research (gtr.ukri.org) and saves them as a spreadsheet (CSV) that opens in Excel. Work through the numbered tabs, then click Run search at the bottom.

1  Search terms
Type one search term per line. Each line is a separate search; the results are combined and duplicates removed.
•  Use short phrases of two or three words, e.g. "molecular junction", "break junction", "water quality". GtR matches the words in each term quite strictly, so several short terms find more than one long one.
•  No quotation marks, AND/OR or wildcards are needed.
•  "Use default terms" fills in the built-in list.

2  Filters
GtR applies these before sending results, so a narrow search is also a fast one.
•  Funders: tick the councils you want, or "All funders".
      EPSRC – engineering and physical sciences
      BBSRC – biotechnology and biological sciences
      MRC – medical research
      NERC – natural environment
      STFC – science and technology facilities (physics, astronomy, large facilities)
      AHRC – arts and humanities
      ESRC – economic and social sciences
      Innovate UK – business-led innovation projects; useful when looking for industry partners
      Research England – funding to English universities
      Horizon Europe Guarantee – UK funding for awards won under the EU's Horizon Europe programme
•  Status: active projects are still running; closed ones have finished.
•  Start year: leave either box blank for no limit.
•  Region: the region of the lead organisation. Tick none to include every region.
•  Award type: tick none to include every type. Use "Other" for a type not listed, exactly as GtR writes it (see "Show available values"); separate several with commas.
•  Award value: tick none to include every value.

3  Options
•  Fetch first: which matches GtR sends first. "Most relevant first" suits most searches; "Newest start date first" brings the latest projects, even weaker matches.
•  Pages per term: each page is up to 100 projects. 2 pages is plenty for most terms.
•  Sort results by: the order of rows in the spreadsheet.
•  Co-investigators, quick mode, refresh: see the notes next to each.
•  Compare with earlier results: pick spreadsheets from earlier searches; a "new" column then marks projects not in them (yes/no).

4  Scoring
Each project gets a score for how well it matches your keywords: 2 points if a keyword is in the title, 1 if it is in the abstract. Rows are then sorted with the best matches first.
•  Keywords: one per line. Whole words only; plurals are matched automatically ("sensor" also finds "sensors", but not "sensory"). Hyphens and spaces count as the same ("host-guest" = "host guest").
•  Ignore these phrases: removed from the text before scoring, to stop false matches, e.g. "gap junction" so that "junction" does not score for cell biology.
•  Untick "Score by keywords" to keep GtR's own ranking.

Available values
"Show available values" asks GtR which funders, award types and regions occur for your search terms and filters, with the number of projects for each. It takes a few seconds and writes no file. Use it to check spelling for the "Other" award type box, or to see where the projects for a topic sit.

Results
Results are saved in the output folder with a dated name such as gtr_2026-11-11_water-quality-plus2.csv; existing files are never overwritten. "Open results" opens the last file in Excel; "Open folder" shows the folder.
Columns include grant reference, title, lead person (pi), co-investigators, lead organisation and department, status, start and end dates, value in £, funder, abstract, a link to the GtR page, the keyword score, project partners and the lead organisation's region and postcode.

Good to know
•  The first run of a search is the slowest: each project is looked up once and remembered, so later runs are much faster.
•  Stop is safe: the projects looked up so far are saved to the spreadsheet.
•  Some GtR records list no co-investigators even when the grant had them, and PhD studentships have no lead person.
•  GtR data is published by UKRI under the Open Government Licence v3.0.
"""


# ---------------------------------------------------------------- folders and settings

def documents_folder():
    """The user's Documents folder (follows OneDrive redirection on Windows)."""
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes
            from uuid import UUID

            class GUID(ctypes.Structure):
                _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                            ("Data3", wintypes.WORD), ("Data4", wintypes.BYTE * 8)]

            uid = UUID("{FDD39AD0-238F-46AF-ADB4-6C85480369C7}")   # FOLDERID_Documents
            guid = GUID(uid.time_low, uid.time_mid, uid.time_hi_version,
                        (wintypes.BYTE * 8).from_buffer_copy(uid.bytes[8:]))
            path = ctypes.c_wchar_p()
            if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, None,
                                                         ctypes.byref(path)) == 0:
                folder = path.value
                ctypes.windll.ole32.CoTaskMemFree(path)
                return folder
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


APP_DIR = os.path.join(documents_folder(), APP_NAME)
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")
CACHE_FILE = os.path.join(APP_DIR, gs.CACHE_FILE)


def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_settings(data):
    try:
        os.makedirs(APP_DIR, exist_ok=True)
        tmp = SETTINGS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=1)
        os.replace(tmp, SETTINGS_FILE)
    except OSError:
        pass


def open_path(path):
    if sys.platform == "win32":
        os.startfile(path)                        # opens CSVs in Excel, folders in Explorer
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ---------------------------------------------------------------- small widgets

def hint(parent, text, wrap=380, **grid):
    lbl = ttk.Label(parent, text=text, foreground=HINT_COLOUR, wraplength=wrap, justify="left",
                    style="Hint.TLabel")
    if grid:
        lbl.grid(**grid)
    return lbl


class LinesBox(ttk.Frame):
    """Multi-line text box returning a list of non-empty lines."""

    def __init__(self, parent, height=8, width=40):
        super().__init__(parent)
        self.text = tk.Text(self, height=height, width=width, wrap="none", undo=True,
                            font=("Segoe UI", 10) if sys.platform == "win32" else None)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    def get_lines(self):
        return [ln.strip() for ln in self.text.get("1.0", "end").splitlines() if ln.strip()]

    def set_lines(self, lines):
        self.text.delete("1.0", "end")
        self.text.insert("1.0", "\n".join(lines))


# ---------------------------------------------------------------- the window

class App:
    def __init__(self, root):
        self.root = root
        self.queue = queue.Queue()
        self.worker = None
        self.stop_event = None
        self.last_file = None
        self.closing = False
        root.title(f"{APP_NAME} {gs.__version__} – UKRI Gateway to Research")
        root.minsize(940, 700)
        style = ttk.Style()
        style.configure("Hint.TLabel", font=("Segoe UI", 8) if sys.platform == "win32" else ("TkDefaultFont", 8))
        style.configure("Big.TButton", padding=(12, 4))

        gs.set_logger(lambda text: self.queue.put(("log", text)))

        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 4))
        self._tab_terms()
        self._tab_filters()
        self._tab_options()
        self._tab_scoring()
        self._tab_values()
        self._tab_help()
        self._bottom()
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        self.apply_settings(load_settings())
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(100, self.poll)

    # ------------------------------------------------------------ tabs

    def _tab_terms(self):
        f = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(f, text="1  Search terms")
        ttk.Label(f, text="Search terms – one per line").grid(row=0, column=0, sticky="w")
        hint(f, "Short phrases of two or three words work best, e.g. molecular junction, break junction, "
                "water quality. Each line is searched separately and the results are combined. "
                "No quotation marks or AND/OR needed.", wrap=800, row=1, column=0, columnspan=2,
             sticky="w", pady=(0, 6))
        self.terms = LinesBox(f, height=14, width=60)
        self.terms.grid(row=2, column=0, sticky="nsew")
        side = ttk.Frame(f)
        side.grid(row=2, column=1, sticky="n", padx=(10, 0))
        ttk.Button(side, text="Use default terms",
                   command=lambda: self.terms.set_lines(gs.DEFAULT_TERMS)).grid(row=0, column=0, sticky="ew")
        ttk.Button(side, text="Clear", command=lambda: self.terms.set_lines([])).grid(
            row=1, column=0, sticky="ew", pady=4)
        f.columnconfigure(0, weight=1)
        f.rowconfigure(2, weight=1)

    def _tab_filters(self):
        f = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(f, text="2  Filters")
        left, right = ttk.Frame(f), ttk.Frame(f)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right.grid(row=0, column=1, sticky="nsew")
        f.columnconfigure(0, weight=1)
        f.columnconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        # Funders
        lf = ttk.LabelFrame(left, text="Funders", padding=8)
        lf.grid(row=0, column=0, sticky="ew")
        self.all_funders = tk.BooleanVar(value=False)
        ttk.Checkbutton(lf, text="All funders", variable=self.all_funders,
                        command=self._funders_toggled).grid(row=0, column=0, columnspan=2, sticky="w")
        self.funder_vars, self.funder_checks = {}, []
        for i, (code, name) in enumerate(FUNDERS):
            v = tk.BooleanVar(value=(code == "EPSRC"))
            cb = ttk.Checkbutton(lf, text=code, variable=v)
            cb.grid(row=1 + i // 2, column=i % 2, sticky="w", padx=(0, 16))
            self.funder_vars[code] = v
            self.funder_checks.append(cb)
        hint(lf, "Tick one or more, or All funders. Innovate UK projects are company-led: useful for "
                 "finding industry partners. See Help for what each funder covers.", wrap=330,
             row=7, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # Status
        lf = ttk.LabelFrame(left, text="Project status", padding=8)
        lf.grid(row=1, column=0, sticky="ew", pady=8)
        self.status = tk.StringVar(value="both")
        for i, (value, label) in enumerate(STATUS_LABELS):
            ttk.Radiobutton(lf, text=label, value=value, variable=self.status).grid(row=0, column=i, sticky="w",
                                                                                 padx=(0, 12))

        # Years
        lf = ttk.LabelFrame(left, text="Start year", padding=8)
        lf.grid(row=2, column=0, sticky="ew")
        self.since, self.until = tk.StringVar(), tk.StringVar()
        ttk.Label(lf, text="From").grid(row=0, column=0)
        ttk.Combobox(lf, textvariable=self.since, values=YEARS, width=7).grid(row=0, column=1, padx=(4, 12))
        ttk.Label(lf, text="To").grid(row=0, column=2)
        ttk.Combobox(lf, textvariable=self.until, values=YEARS, width=7).grid(row=0, column=3, padx=4)
        hint(lf, "Year the project started. Leave a box blank for no limit.", wrap=330, row=1, column=0,
             columnspan=4, sticky="w", pady=(4, 0))

        # Value
        lf = ttk.LabelFrame(left, text="Award value", padding=8)
        lf.grid(row=3, column=0, sticky="ew", pady=8)
        self.amount_vars = {}
        for i, (code, label) in enumerate(AMOUNTS):
            v = tk.BooleanVar()
            ttk.Checkbutton(lf, text=label, variable=v).grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 12))
            self.amount_vars[code] = v
        hint(lf, "Total value of the award. Tick none to include every value.", wrap=330, row=2, column=0,
             columnspan=2, sticky="w", pady=(4, 0))

        # Regions
        lf = ttk.LabelFrame(right, text="Region of the lead organisation", padding=8)
        lf.grid(row=0, column=0, sticky="ew")
        self.region_vars = {}
        for i, name in enumerate(REGIONS):
            v = tk.BooleanVar()
            ttk.Checkbutton(lf, text=name, variable=v).grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 10))
            self.region_vars[name] = v
        hint(lf, "Tick none to include every region.", row=8, column=0, columnspan=2, sticky="w",
             pady=(4, 0))

        # Categories
        lf = ttk.LabelFrame(right, text="Award type", padding=8)
        lf.grid(row=1, column=0, sticky="ew", pady=8)
        self.category_vars = {}
        for i, name in enumerate(CATEGORIES):
            v = tk.BooleanVar()
            ttk.Checkbutton(lf, text=name, variable=v).grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 10))
            self.category_vars[name] = v
        ttk.Label(lf, text="Other").grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.category_other = tk.StringVar()
        ttk.Entry(lf, textvariable=self.category_other, width=24).grid(row=3, column=1, sticky="ew", pady=(6, 0))
        hint(lf, "Tick none to include every type. 'Other': any other type exactly as GtR writes it "
                 "(see Show available values); separate several with commas.", wrap=360, row=4, column=0,
             columnspan=2, sticky="w", pady=(4, 0))

    def _tab_options(self):
        f = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(f, text="3  Options")
        row = 0
        ttk.Label(f, text="Fetch first").grid(row=row, column=0, sticky="w", padx=(0, 10))
        self.fetch = tk.StringVar(value="Most relevant first")
        ttk.Combobox(f, textvariable=self.fetch, values=list(FETCH_LABELS), state="readonly",
                     width=28).grid(row=row, column=1, sticky="w")
        hint(f, "Which matches GtR sends first. 'Newest start date first' brings the latest projects, "
                "including weaker matches.", wrap=420, row=row, column=2, sticky="w", padx=10)
        row += 1
        ttk.Label(f, text="Pages per term").grid(row=row, column=0, sticky="w", pady=6, padx=(0, 10))
        self.pages = tk.StringVar(value=str(gs.DEFAULT_PAGES))
        ttk.Spinbox(f, from_=1, to=100, textvariable=self.pages, width=6).grid(row=row, column=1, sticky="w")
        hint(f, "Up to 100 projects per page. 2 is plenty for most terms.", wrap=420, row=row, column=2,
             sticky="w", padx=10)
        row += 1
        ttk.Label(f, text="Sort results by").grid(row=row, column=0, sticky="w", padx=(0, 10))
        self.sort = tk.StringVar(value="Keyword score, highest first")
        ttk.Combobox(f, textvariable=self.sort, values=list(SORT_LABELS), state="readonly",
                     width=28).grid(row=row, column=1, sticky="w")
        hint(f, "Order of rows in the spreadsheet. Keyword score needs scoring switched on (tab 4).",
             wrap=420, row=row, column=2, sticky="w", padx=10)
        row += 1
        ttk.Separator(f).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        self.coi = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="Include co-investigators", variable=self.coi).grid(row=row, column=0,
                                                                                  columnspan=2, sticky="w")
        row += 1
        self.quick = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Quick mode", variable=self.quick).grid(row=row, column=0, columnspan=2,
                                                                       sticky="w", pady=4)
        hint(f, "Much faster, but leaves out abstracts, co-investigators, postcodes and keyword scores.",
             wrap=420, row=row, column=2, sticky="w", padx=10)
        row += 1
        self.refresh = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Refresh saved lookups", variable=self.refresh).grid(row=row, column=0,
                                                                                    columnspan=2, sticky="w")
        hint(f, "Looks every project up again instead of using the copies saved by earlier runs. "
                "Only needed if GtR has updated its records.", wrap=420, row=row, column=2, sticky="w", padx=10)
        row += 1
        ttk.Separator(f).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        ttk.Label(f, text="Compare with earlier results").grid(row=row, column=0, columnspan=2, sticky="w")
        hint(f, "Pick spreadsheets from earlier searches. A 'new' column then marks projects that are "
                "not in them.", wrap=420, row=row, column=2, sticky="w", padx=10)
        row += 1
        box = ttk.Frame(f)
        box.grid(row=row, column=0, columnspan=3, sticky="nsew")
        self.compare_list = tk.Listbox(box, height=5, selectmode="extended")
        self.compare_list.grid(row=0, column=0, sticky="nsew")
        btns = ttk.Frame(box)
        btns.grid(row=0, column=1, sticky="n", padx=(8, 0))
        ttk.Button(btns, text="Add files…", command=self._add_compare).grid(row=0, column=0, sticky="ew")
        ttk.Button(btns, text="Remove selected", command=self._remove_compare).grid(row=1, column=0,
                                                                                   sticky="ew", pady=4)
        ttk.Button(btns, text="Clear", command=lambda: self.compare_list.delete(0, "end")).grid(
            row=2, column=0, sticky="ew")
        box.columnconfigure(0, weight=1)
        f.columnconfigure(2, weight=1)
        f.rowconfigure(row, weight=1)

    def _tab_scoring(self):
        f = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(f, text="4  Scoring")
        self.score = tk.BooleanVar(value=gs.USE_SCORING)
        ttk.Checkbutton(f, text="Score by keywords", variable=self.score).grid(row=0, column=0, sticky="w")
        hint(f, "Each keyword found adds 2 points if it is in the title and 1 if it is in the abstract. "
                "Untick to keep GtR's own ranking.", wrap=800, row=1, column=0, columnspan=2, sticky="w",
             pady=(0, 8))
        ttk.Label(f, text="Keywords – one per line").grid(row=2, column=0, sticky="w")
        ttk.Label(f, text="Ignore these phrases – one per line").grid(row=2, column=1, sticky="w", padx=(12, 0))
        hint(f, "Whole words only; plurals are matched automatically ('sensor' finds 'sensors', not "
                "'sensory'). Hyphens and spaces count as the same.", wrap=400, row=3, column=0, sticky="w")
        hint(f, "Removed before scoring to stop false matches, e.g. 'gap junction' so that 'junction' "
                "does not score for cell biology.", wrap=400, row=3, column=1, sticky="w", padx=(12, 0))
        self.keywords = LinesBox(f, height=12, width=36)
        self.keywords.grid(row=4, column=0, sticky="nsew", pady=4)
        self.exclude = LinesBox(f, height=12, width=36)
        self.exclude.grid(row=4, column=1, sticky="nsew", pady=4, padx=(12, 0))
        ttk.Button(f, text="Reset to defaults",
                   command=lambda: self.keywords.set_lines(gs.DEFAULT_KEYWORDS)).grid(row=5, column=0, sticky="w")
        ttk.Button(f, text="Reset to defaults",
                   command=lambda: self.exclude.set_lines(gs.EXCLUDE_PHRASES)).grid(row=5, column=1, sticky="w",
                                                                                 padx=(12, 0))
        f.columnconfigure(0, weight=1)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(4, weight=1)

    def _tab_values(self):
        f = ttk.Frame(self.notebook, padding=12)
        self.values_tab = f
        self.notebook.add(f, text="Available values")
        self.values_summary = tk.StringVar(
            value="Click 'Show available values' below to see which funders, award types and regions occur "
                  "for your search terms and filters, with the number of projects for each.")
        ttk.Label(f, textvariable=self.values_summary, wraplength=850, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self.values_tree = ttk.Treeview(f, columns=("count",), height=16)
        self.values_tree.heading("#0", text="Value")
        self.values_tree.heading("count", text="Projects")
        self.values_tree.column("#0", width=520)
        self.values_tree.column("count", width=100, anchor="e")
        sb = ttk.Scrollbar(f, orient="vertical", command=self.values_tree.yview)
        self.values_tree.configure(yscrollcommand=sb.set)
        self.values_tree.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")
        hint(f, "Counts are added up across your search terms, so a project matching two terms counts twice.",
             wrap=850, row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        f.columnconfigure(0, weight=1)
        f.rowconfigure(1, weight=1)

    def _tab_help(self):
        f = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(f, text="Help")
        txt = ScrolledText(f, wrap="word", padx=10, pady=8,
                           font=("Segoe UI", 10) if sys.platform == "win32" else None)
        base = ("Segoe UI", 10) if sys.platform == "win32" else ("TkDefaultFont", 10)
        txt.tag_configure("h1", font=(base[0], 14, "bold"), spacing1=4, spacing3=6)
        txt.tag_configure("h2", font=(base[0], 11, "bold"), spacing1=10, spacing3=3)
        txt.tag_configure("body", lmargin2=18, spacing1=2)
        for line in HELP_TEXT.splitlines():
            if line == HELP_HEADINGS[0]:
                tag = "h1"
            elif line in HELP_HEADINGS:
                tag = "h2"
            else:
                tag = "body"
            if line.strip() or tag != "body":
                txt.insert("end", line + "\n", tag)
        txt.configure(state="disabled")
        txt.grid(row=0, column=0, sticky="nsew")
        f.columnconfigure(0, weight=1)
        f.rowconfigure(0, weight=1)

    def _bottom(self):
        f = ttk.Frame(self.root, padding=(8, 0, 8, 8))
        f.grid(row=1, column=0, sticky="nsew")
        out = ttk.Frame(f)
        out.grid(row=0, column=0, sticky="ew")
        ttk.Label(out, text="Save results in").grid(row=0, column=0, sticky="w")
        self.out_dir = tk.StringVar(value=APP_DIR)
        ttk.Entry(out, textvariable=self.out_dir).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(out, text="Browse…", command=self._browse_out).grid(row=0, column=2)
        out.columnconfigure(1, weight=1)

        bar = ttk.Frame(f)
        bar.grid(row=1, column=0, sticky="ew", pady=6)
        self.btn_values = ttk.Button(bar, text="Show available values", command=self.show_values)
        self.btn_run = ttk.Button(bar, text="Run search", style="Big.TButton", command=self.run_search)
        self.btn_stop = ttk.Button(bar, text="Stop", command=self.stop, state="disabled")
        self.btn_open = ttk.Button(bar, text="Open results", command=self.open_results, state="disabled")
        self.btn_folder = ttk.Button(bar, text="Open folder", command=self.open_folder)
        for i, b in enumerate((self.btn_values, self.btn_run, self.btn_stop, self.btn_open, self.btn_folder)):
            b.grid(row=0, column=i, padx=(0, 6))

        prog = ttk.Frame(f)
        prog.grid(row=2, column=0, sticky="ew")
        self.progress = ttk.Progressbar(prog, mode="determinate", maximum=100, length=320)
        self.progress.grid(row=0, column=0, sticky="w")
        self.status_text = tk.StringVar(value="Ready.")
        ttk.Label(prog, textvariable=self.status_text).grid(row=0, column=1, sticky="w", padx=10)
        prog.columnconfigure(1, weight=1)

        self.log = ScrolledText(f, height=7, wrap="word", state="disabled",
                                font=("Consolas", 9) if sys.platform == "win32" else None)
        self.log.grid(row=3, column=0, sticky="nsew", pady=(6, 0))
        f.columnconfigure(0, weight=1)

    # ------------------------------------------------------------ helpers

    def _funders_toggled(self):
        state = "disabled" if self.all_funders.get() else "normal"
        for cb in self.funder_checks:
            cb.configure(state=state)

    def _add_compare(self):
        paths = filedialog.askopenfilenames(title="Earlier result files", initialdir=self.out_dir.get() or APP_DIR,
                                            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        existing = set(self.compare_list.get(0, "end"))
        for p in paths:
            if p not in existing:
                self.compare_list.insert("end", p)

    def _remove_compare(self):
        for i in reversed(self.compare_list.curselection()):
            self.compare_list.delete(i)

    def _browse_out(self):
        folder = filedialog.askdirectory(title="Save results in", initialdir=self.out_dir.get() or APP_DIR)
        if folder:
            self.out_dir.set(folder)

    def write_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip("\n").lstrip("\n") + "\n" if text.strip() else "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def set_running(self, running):
        for b in (self.btn_values, self.btn_run):
            b.configure(state="disabled" if running else "normal")
        self.btn_stop.configure(state="normal" if running else "disabled")
        if not running:
            self.progress.stop()
            self.progress.configure(mode="determinate")

    # ------------------------------------------------------------ settings

    def collect_settings(self):
        return {
            "terms": self.terms.get_lines(),
            "all_funders": self.all_funders.get(),
            "funders": [c for c, v in self.funder_vars.items() if v.get()],
            "status": self.status.get(),
            "since": self.since.get(), "until": self.until.get(),
            "regions": [r for r, v in self.region_vars.items() if v.get()],
            "categories": [c for c, v in self.category_vars.items() if v.get()],
            "category_other": self.category_other.get(),
            "amounts": [a for a, v in self.amount_vars.items() if v.get()],
            "fetch": self.fetch.get(), "pages": self.pages.get(), "sort": self.sort.get(),
            "coi": self.coi.get(), "quick": self.quick.get(),
            "compare": list(self.compare_list.get(0, "end")),
            "score": self.score.get(),
            "keywords": self.keywords.get_lines(), "exclude": self.exclude.get_lines(),
            "out_dir": self.out_dir.get(),
            "geometry": self.root.geometry(),
        }

    def apply_settings(self, s):
        self.terms.set_lines(s.get("terms", gs.DEFAULT_TERMS))
        self.all_funders.set(s.get("all_funders", False))
        chosen = set(s.get("funders", ["EPSRC"]))
        for c, v in self.funder_vars.items():
            v.set(c in chosen)
        self._funders_toggled()
        self.status.set(s.get("status", "both"))
        self.since.set(s.get("since", ""))
        self.until.set(s.get("until", ""))
        for name, store in (("regions", self.region_vars), ("categories", self.category_vars),
                            ("amounts", self.amount_vars)):
            chosen = set(s.get(name, []))
            for k, v in store.items():
                v.set(k in chosen)
        self.category_other.set(s.get("category_other", ""))
        if s.get("fetch") in FETCH_LABELS:
            self.fetch.set(s["fetch"])
        self.pages.set(s.get("pages", str(gs.DEFAULT_PAGES)))
        if s.get("sort") in SORT_LABELS:
            self.sort.set(s["sort"])
        self.coi.set(s.get("coi", True))
        self.quick.set(s.get("quick", False))
        self.compare_list.delete(0, "end")
        for p in s.get("compare", []):
            self.compare_list.insert("end", p)
        self.score.set(s.get("score", gs.USE_SCORING))
        self.keywords.set_lines(s.get("keywords", gs.DEFAULT_KEYWORDS))
        self.exclude.set_lines(s.get("exclude", gs.EXCLUDE_PHRASES))
        self.out_dir.set(s.get("out_dir") or APP_DIR)
        if s.get("geometry"):
            try:
                self.root.geometry(s["geometry"])
            except tk.TclError:
                pass

    # ------------------------------------------------------------ building options

    def build_args(self, list_values=False):
        """Options for gtr_search.run(), or None (after showing a message) if something is wrong."""
        terms = self.terms.get_lines()
        if not terms:
            self.notebook.select(0)
            messagebox.showwarning(APP_NAME, "Add at least one search term (tab 1), or click "
                                             "'Use default terms'.")
            return None
        if self.all_funders.get():
            funder = "all"
        else:
            chosen = [c for c, v in self.funder_vars.items() if v.get()]
            if not chosen:
                self.notebook.select(1)
                messagebox.showwarning(APP_NAME, "Tick at least one funder, or 'All funders' (tab 2).")
                return None
            funder = ",".join(chosen)
        years = {}
        for name, var in (("since", self.since), ("until", self.until)):
            text = var.get().strip()
            if text and not (text.isdigit() and len(text) == 4):
                self.notebook.select(1)
                messagebox.showwarning(APP_NAME, f"Start year '{text}' is not a year. Use four digits, "
                                                 "or leave the box blank.")
                return None
            years[name] = int(text) if text else None
        try:
            pages = int(self.pages.get())
        except ValueError:
            pages = 0
        if not 1 <= pages <= 100:
            self.notebook.select(2)
            messagebox.showwarning(APP_NAME, "Pages per term must be a whole number from 1 to 100.")
            return None
        categories = [c for c, v in self.category_vars.items() if v.get()]
        categories += gs.split_list(self.category_other.get())
        out_dir = self.out_dir.get().strip() or APP_DIR
        options = dict(
            terms=terms, funder=funder, status=self.status.get(),
            since=years["since"], until=years["until"],
            region=",".join(r for r, v in self.region_vars.items() if v.get()) or None,
            category=",".join(categories) or None,
            amount=",".join(a for a, v in self.amount_vars.items() if v.get()) or None,
            fetch=FETCH_LABELS.get(self.fetch.get(), "relevance"), pages=pages,
            sort=SORT_LABELS.get(self.sort.get()),
            score=self.score.get(), no_coi=not self.coi.get(), quick=self.quick.get(),
            keywords=",".join(self.keywords.get_lines()) or None,
            exclude=",".join(self.exclude.get_lines()),
            compare=list(self.compare_list.get(0, "end")) or None,
            refresh=self.refresh.get(), out_dir=out_dir, cache_file=CACHE_FILE,
            list_values=list_values,
        )
        try:
            os.makedirs(APP_DIR, exist_ok=True)
            return gs.make_args(**options)
        except (ValueError, OSError) as e:
            messagebox.showwarning(APP_NAME, str(e))
            return None

    # ------------------------------------------------------------ actions

    def run_search(self):
        args = self.build_args()
        if args:
            save_settings(self.collect_settings())
            self.refresh.set(False)            # one-off: don't keep refreshing on later runs
            self.start(args, "Searching GtR…")

    def show_values(self):
        args = self.build_args(list_values=True)
        if args:
            save_settings(self.collect_settings())
            self.start(args, "Asking GtR which values occur…")

    def start(self, args, status):
        self.write_log("")
        self.stop_event = threading.Event()
        self.set_running(True)
        self.status_text.set(status)
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)

        def work():
            try:
                result = gs.run(args, progress=self.on_progress, stop=self.stop_event)
                self.queue.put(("done", result))
            except FileNotFoundError as e:
                self.queue.put(("error", str(e)))
            except Exception as e:                       # show anything unexpected, don't crash
                self.queue.put(("error", f"Something went wrong: {e}"))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def on_progress(self, i, n, remaining, row):          # called from the worker thread
        self.queue.put(("progress", i, n, remaining, row.get("title", "")))

    def stop(self):
        if self.stop_event:
            self.stop_event.set()
            self.status_text.set("Stopping…")
            self.btn_stop.configure(state="disabled")

    def open_results(self):
        if self.last_file and os.path.exists(self.last_file):
            open_path(self.last_file)

    def open_folder(self):
        folder = self.out_dir.get().strip() or APP_DIR
        os.makedirs(folder, exist_ok=True)
        open_path(folder)

    def show_value_lists(self, values):
        tree = self.values_tree
        tree.delete(*tree.get_children())
        if not values or not values["facets"]:
            self.values_summary.set("GtR returned no value lists for these searches.")
            return
        per_term = ", ".join(f"{t} ({n})" for t, n in values["per_term"].items())
        self.values_summary.set(f"{values['matches']} matches in total: {per_term}. "
                                "Use the values exactly as shown.")
        for name in sorted(values["facets"]):
            items = values["facets"][name]
            parent = tree.insert("", "end", text=f"{name}  ({len(items)} values)", open=True)
            for value, count in items:
                tree.insert(parent, "end", text=value or "(blank)", values=(f"{count:,}",))
        self.notebook.select(self.values_tab)

    # ------------------------------------------------------------ queue polling

    def poll(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self.write_log(msg[1])
                elif kind == "progress":
                    _, i, n, remaining, title = msg
                    if str(self.progress.cget("mode")) != "determinate":
                        self.progress.stop()
                        self.progress.configure(mode="determinate")
                    self.progress["value"] = i * 100 / n
                    mins, secs = int(remaining // 60), int(remaining % 60)
                    short = title[:40] + ("…" if len(title) > 40 else "")
                    self.status_text.set(f"{i} of {n} projects  ·  about {mins}m{secs:02d}s left  ·  {short}")
                elif kind == "done":
                    self.finished(msg[1])
                elif kind == "error":
                    self.set_running(False)
                    self.status_text.set("Stopped with an error.")
                    self.write_log(msg[1])
                    messagebox.showerror(APP_NAME, msg[1])
        except queue.Empty:
            pass
        if self.closing and not (self.worker and self.worker.is_alive()):
            self.root.destroy()
            return
        self.root.after(100, self.poll)

    def finished(self, result):
        self.set_running(False)
        if result.get("values") is not None:
            self.show_value_lists(result["values"])
            self.status_text.set("Value lists ready (see the Available values tab).")
            return
        if result.get("outfile"):
            self.last_file = result["outfile"]
            self.btn_open.configure(state="normal")
            self.progress["value"] = 100
            text = f"Saved {result['saved']} projects to {os.path.basename(result['outfile'])}"
            if result.get("new") is not None:
                text += f" ({result['new']} new)"
            self.status_text.set(text + (" – stopped early." if result.get("stopped") else "."))
        else:
            self.progress["value"] = 0
            self.status_text.set("Stopped – no file written." if result.get("stopped")
                                 else "No results – see the messages below.")

    def on_close(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(APP_NAME, "A search is still running. Stop it and close?\n\n"
                                                 "Projects looked up so far will be saved."):
                return
            self.stop()
            self.closing = True
            save_settings(self.collect_settings())
            return                                          # poll() closes once the search has stopped
        save_settings(self.collect_settings())
        self.root.destroy()


def check_engine():
    """Make sure gtr_search.py is the version this window was written for."""
    missing = [name for name in ("run", "make_args", "set_logger", "split_list") if not hasattr(gs, name)]
    if missing or not hasattr(gs, "__version__"):
        where = getattr(gs, "__file__", "gtr_search.py")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_NAME, "The gtr_search.py next to this program is an older version:\n\n"
                                       f"{where}\n\n"
                                       "Replace it with the gtr_search.py that came with gtr_gui.py "
                                       "(same download or release), then start the program again.")
        root.destroy()
        sys.exit(1)


def main():
    check_engine()
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)   # sharp text on high-DPI screens
        except Exception:
            pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
