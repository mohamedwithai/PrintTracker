#!/usr/bin/env python3
"""
Print Tracker - Kyocera TASKalfa 2553ci
---------------------------------------
* Reads the printer's own page counters (paper size x colour, duplex) over HTTP.
  Every counter change becomes a tracked print event  -> nothing is missed.
* Watches the Windows print queue (pywin32) to learn document name, tray
  (cassette / bypass), paper size and colour of each job.
* Serves a live dashboard + pricing back end on http://localhost:8765

Only the Python standard library is required. pywin32 (bundled with Anaconda)
adds document names and tray detection.
"""
import json
import os
import re
import sqlite3
import sys
import threading
import time
import urllib.request
import webbrowser
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE = os.path.dirname(os.path.abspath(__file__))
CFG_PATH = os.path.join(BASE, "config.json")
DEFAULT_CFG = {
    "printer_ip": "192.168.50.10",
    "port": 8765,
    "printer_name_match": "kyocera|taskalfa",
    "open_browser": True,
}


def load_cfg():
    cfg = dict(DEFAULT_CFG)
    try:
        with open(CFG_PATH, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
    except FileNotFoundError:
        try:
            with open(CFG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except OSError:
            pass
    except Exception as e:  # bad json -> keep defaults
        print("config.json could not be read, using defaults:", e)
    if os.environ.get("PT_PRINTER_URL"):
        cfg["printer_url"] = os.environ["PT_PRINTER_URL"]
    if os.environ.get("PT_PORT"):
        cfg["port"] = int(os.environ["PT_PORT"])
    if os.environ.get("PT_NO_BROWSER"):
        cfg["open_browser"] = False
    cfg.setdefault("printer_url", "http://%s" % cfg["printer_ip"])
    return cfg


CFG = load_cfg()
DB_PATH = os.environ.get("PT_DB") or os.path.join(BASE, "print_tracker.db")

# ---------------------------------------------------------------- categories
CATS = [
    # key, group, label, colour
    ("a4_bw_single", "A4", "B&W - Single-sided", "bw"),
    ("a4_bw_duplex", "A4", "B&W - Duplex", "bw"),
    ("a4_color_single", "A4", "Colour - Single-sided", "color"),
    ("a4_color_duplex", "A4", "Colour - Duplex", "color"),
    ("a3_tray_bw", "A3 - Standard tray", "B&W", "bw"),
    ("a3_tray_color", "A3 - Standard tray", "Colour", "color"),
    ("a3_bypass_bw", "A3 - Bypass / manual tray", "B&W", "bw"),
    ("a3_bypass_color", "A3 - Bypass / manual tray", "Colour", "color"),
    ("master_legal", "Offset masters", "Legal master (B&W)", "bw"),
    ("master_12x18", "Offset masters", "12 x 18 in master (B&W)", "bw"),
    ("other_bw", "Other sizes (uncategorised)", "B&W", "bw"),
    ("other_color", "Other sizes (uncategorised)", "Colour", "color"),
]
CAT_KEYS = [c[0] for c in CATS]

# paper stock types: key, label. Each print category uses one of them.
PAPERS = [
    ("a4", "A4 paper"),
    ("a3", "A3 paper"),
    ("legal_master", "Legal master sheet"),
    ("master_12x18", "12 x 18 in master sheet"),
    ("other", "Other paper"),
]
PAPER_KEYS = [p[0] for p in PAPERS]


def paper_of(cat):
    if cat.startswith("a4_"):
        return "a4"
    if cat.startswith("a3_"):
        return "a3"
    if cat == "master_legal":
        return "legal_master"
    if cat == "master_12x18":
        return "master_12x18"
    return "other"


def sheets_of(cat, pages):
    """Sheets of paper used: duplex puts 2 printed sides on one sheet."""
    return (pages + 1) // 2 if cat.endswith("duplex") else pages

# index of each paper size inside the printer's counter arrays
SIZES = ["A4", "B5", "A5", "Folio", "Legal", "Letter", "Statement", "Banner 1",
         "Banner 2", "Other (Single)", "Other (Double)", "A3", "B4", "Ledger"]
I_A4, I_LEGAL, I_OTH1, I_OTH2, I_A3 = 0, 4, 9, 10, 11

# ---------------------------------------------------------------- database
_db_lock = threading.RLock()
_db = sqlite3.connect(DB_PATH, check_same_thread=False)
_db.row_factory = sqlite3.Row


def db_init():
    with _db_lock:
        _db.executescript("""
        CREATE TABLE IF NOT EXISTS events(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts REAL NOT NULL, ts_end REAL NOT NULL,
            category TEXT NOT NULL, pages INTEGER NOT NULL,
            doc TEXT, app TEXT, via TEXT, size TEXT, tray_src TEXT, edited INTEGER DEFAULT 0);
        CREATE INDEX IF NOT EXISTS ev_ts ON events(ts);
        CREATE TABLE IF NOT EXISTS kv(k TEXT PRIMARY KEY, v TEXT);
        CREATE TABLE IF NOT EXISTS rates(category TEXT PRIMARY KEY, rate REAL NOT NULL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS paper_rates(paper TEXT PRIMARY KEY, rate REAL NOT NULL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS purchases(
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts REAL NOT NULL, paper TEXT NOT NULL,
            supplier TEXT, packets REAL NOT NULL, sheets_per_packet INTEGER NOT NULL,
            total_cost REAL NOT NULL, note TEXT);
        CREATE TABLE IF NOT EXISTS jobs(
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts REAL NOT NULL, customer TEXT, title TEXT,
            qty TEXT, customer_amount REAL NOT NULL DEFAULT 0, customer_paid REAL NOT NULL DEFAULT 0,
            supplier TEXT, supplier_amount REAL NOT NULL DEFAULT 0, supplier_paid REAL NOT NULL DEFAULT 0,
            other_cost REAL NOT NULL DEFAULT 0, note TEXT);
        CREATE TABLE IF NOT EXISTS expenses(
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts REAL NOT NULL, title TEXT, amount REAL NOT NULL);
        CREATE TABLE IF NOT EXISTS recurring(
            id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, amount REAL NOT NULL, last_ym INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS scan_events(
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts REAL NOT NULL, n INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS job_items(
            id INTEGER PRIMARY KEY AUTOINCREMENT, job_id INTEGER NOT NULL, title TEXT,
            qty REAL NOT NULL DEFAULT 1, rate REAL NOT NULL DEFAULT 0, amount REAL NOT NULL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS capital(
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts REAL NOT NULL, kind TEXT NOT NULL,
            amount REAL NOT NULL, note TEXT);
        """)
        # a master print can belong to an offset job; "sheets" overrides the sheet
        # count when front & back were printed on one master sheet
        have = {r[1] for r in _db.execute("PRAGMA table_info(events)")}
        if "job_id" not in have:
            _db.execute("ALTER TABLE events ADD COLUMN job_id INTEGER")
        if "sheets" not in have:
            _db.execute("ALTER TABLE events ADD COLUMN sheets INTEGER")
        if "personal" not in have:      # printed for the owner himself: not a sale, paper taken out at cost
            _db.execute("ALTER TABLE events ADD COLUMN personal INTEGER DEFAULT 0")
        pcols = {r[1] for r in _db.execute("PRAGMA table_info(purchases)")}
        if "loose" not in pcols:
            _db.execute("ALTER TABLE purchases ADD COLUMN loose INTEGER DEFAULT 0")
        ecols = {r[1] for r in _db.execute("PRAGMA table_info(expenses)")}
        if "due_ts" not in ecols:       # bills entered when they arrive, paid later
            _db.execute("ALTER TABLE expenses ADD COLUMN due_ts REAL")
            _db.execute("ALTER TABLE expenses ADD COLUMN paid INTEGER DEFAULT 1")
            _db.execute("ALTER TABLE expenses ADD COLUMN paid_ts REAL")
            _db.execute("ALTER TABLE expenses ADD COLUMN grp INTEGER")   # rows of one split bill
        if "paid_by" not in {r[1] for r in _db.execute("PRAGMA table_info(expenses)")}:
            # 'owner' = paid from the owner's own pocket -> counts as money he put into the business
            _db.execute("ALTER TABLE expenses ADD COLUMN paid_by TEXT DEFAULT 'business'")
            _db.execute("ALTER TABLE recurring ADD COLUMN paid_by TEXT DEFAULT 'business'")
            # the owner said rent, the printer service and the toner so far came from his own money
            _db.execute("UPDATE expenses SET paid_by='owner' WHERE title LIKE '%rent%' OR title LIKE '%toner%' OR title LIKE '%service%'")
            _db.execute("UPDATE recurring SET paid_by='owner' WHERE title LIKE '%rent%'")
        if "bill_no" not in pcols:      # several lines of one supplier bill share a bill number
            _db.execute("ALTER TABLE purchases ADD COLUMN bill_no TEXT")
        if "paid_by" not in {r[1] for r in _db.execute("PRAGMA table_info(purchases)")}:
            _db.execute("ALTER TABLE purchases ADD COLUMN paid_by TEXT DEFAULT 'business'")
            # the owner said bill 01 was bought with his own money
            _db.execute("UPDATE purchases SET paid_by='owner' WHERE bill_no='01'")
        if not _db.execute("SELECT 1 FROM kv WHERE k='bills_numbered'").fetchone():
            # give every older purchase (no bill no.) a number: one per date + supplier
            nxt = _next_bill()
            groups = {}
            for r in _db.execute("SELECT id, ts, supplier FROM purchases WHERE IFNULL(bill_no,'')='' ORDER BY ts, id"):
                key = (datetime.fromtimestamp(r["ts"]).date(), r["supplier"] or "")
                if key not in groups:
                    groups[key] = nxt
                    nxt = "%0*d" % (len(nxt), int(nxt) + 1)
                _db.execute("UPDATE purchases SET bill_no=? WHERE id=?", (groups[key], r["id"]))
            _db.execute("INSERT INTO kv(k,v) VALUES('bills_numbered','true')")
        # the printer's "Other (Double)" counter (12x18 masters) ticks twice per sheet;
        # halve rows recorded before this was known (once)
        if not _db.execute("SELECT 1 FROM kv WHERE k='oth2_halved'").fetchone():
            _db.execute("UPDATE events SET pages=(pages+1)/2 WHERE edited=0 AND "
                        "((category='master_12x18' AND size='12x18') OR size LIKE 'Other (Double)%')")
            _db.execute("UPDATE events SET sheets=MIN(sheets, pages) WHERE sheets IS NOT NULL")
            _db.execute("INSERT INTO kv(k,v) VALUES('oth2_halved','true')")
        for k in CAT_KEYS:
            _db.execute("INSERT OR IGNORE INTO rates(category, rate) VALUES(?, 0)", (k,))
        for k in PAPER_KEYS:
            _db.execute("INSERT OR IGNORE INTO paper_rates(paper, rate) VALUES(?, 0)", (k,))
        _db.commit()


def _next_bill():
    """Next purchase bill number: highest number used so far + 1, same width (01, 02 ...)."""
    best, width = 0, 2
    with _db_lock:
        rows = _db.execute("SELECT bill_no FROM purchases WHERE IFNULL(bill_no,'')<>''").fetchall()
    for (b,) in rows:
        if str(b).isdigit() and int(b) >= best:
            best, width = int(b), max(width, len(str(b)))
    return "%0*d" % (width, best + 1)


def kv_get(k, default=None):
    with _db_lock:
        r = _db.execute("SELECT v FROM kv WHERE k=?", (k,)).fetchone()
    return json.loads(r["v"]) if r else default


def kv_set(k, v):
    with _db_lock:
        _db.execute("INSERT INTO kv(k,v) VALUES(?,?) ON CONFLICT(k) DO UPDATE SET v=excluded.v",
                    (k, json.dumps(v)))
        _db.commit()


# ---------------------------------------------------------------- printer
MODEL = "/js/jssrc/model/dvcinfo/dvccounter/DvcInfo_Counter_%s.model.htm"


def _get(path):
    url = CFG["printer_url"].rstrip("/") + path
    req = urllib.request.Request(url, headers={
        # Command Center RX answers 500 without these two headers
        "Referer": CFG["printer_url"].rstrip("/") + "/startwlm/Start_Wlm.htm",
        "Cookie": "rtl=0",
        "User-Agent": "Mozilla/5.0 PrintTracker",
    })
    with urllib.request.urlopen(req, timeout=5) as r:
        return r.read().decode("utf-8", "replace")


def _arr(text, name):
    m = re.search(name + r"\s*=\s*\[(.*?)\]", text, re.S)
    if not m:
        raise ValueError("counter %s not found" % name)
    vals = re.findall(r"'([^']*)'", m.group(1))
    out = [int(v) if v.isdigit() else 0 for v in vals]
    if len(out) < len(SIZES):
        raise ValueError("counter %s too short" % name)
    return out[:len(SIZES)]


def _num(text, name):
    m = re.search(r"_pp\." + name + r"\s*=\s*(?:parseInt)?\('(\d+)'", text)
    return int(m.group(1)) if m else 0


def read_printer():
    """One snapshot of the printer's lifetime counters."""
    ps = _get(MODEL % "PaperSize")
    snap = {"bw": _arr(ps, "counterBlackWhite"), "col": _arr(ps, "counterFullColor")}
    # "Other (Double)" = large non-standard paper (12x18), counted 2 per sheet by the printer
    for a in (snap["bw"], snap["col"]):
        if len(a) > I_OTH2:
            a[I_OTH2] //= 2
    dp = _get(MODEL % "Duplex")
    snap["dup2"] = _num(dp, "Duplex2")
    try:
        pc = _get(MODEL % "PrnCounter")
        snap["copy"] = _num(pc, "copytotal")
    except Exception:
        snap["copy"] = None
    try:
        sc = _get(MODEL % "ScanCounter")
        # scanCopy = scans made while copying, scanOther = scan-to-PC/email/USB etc.
        snap["scan"] = _num(sc, "scanCopy") + _num(sc, "scanBlackWhite") + _num(sc, "scanOther")
    except Exception:
        snap["scan"] = None
    return snap


# ---------------------------------------------------------------- spooler
def app_from_doc(doc):
    d = doc or ""
    m = re.match(r"^(Microsoft (?:Word|Excel|PowerPoint)|Adobe [A-Za-z ]+?|Notepad|Paint)\s+-\s+", d, re.I)
    if m:
        return m.group(1), d[m.end():]
    ext = (re.search(r"\.([a-z0-9]{2,5})\s*$", d, re.I) or [None, ""])[1].lower()
    table = {"pdf": "PDF", "ai": "Adobe Illustrator", "psd": "Adobe Photoshop", "doc": "Microsoft Word",
             "docx": "Microsoft Word", "xls": "Microsoft Excel", "xlsx": "Microsoft Excel",
             "cdr": "CorelDRAW", "jpg": "Image", "jpeg": "Image", "png": "Image", "txt": "Notepad",
             "indd": "Adobe InDesign", "pptx": "Microsoft PowerPoint"}
    if ext in table:
        return table[ext], d
    if re.search(r"^https?:|www\.", d, re.I):
        return "Web browser", d
    return "", d


class Job(object):
    def __init__(self, key):
        self.key = key
        self.doc = ""
        self.kind = None      # 'A4' | 'A3' | 'Legal' | '12x18' | 'OTHER' | None (unknown)
        self.color = None     # True / False / None
        self.duplex = None
        self.bypass = None    # True / False / None (auto / unknown)
        self.paper = ""
        self.tray = ""
        self.pages = 0
        self.copies = 1
        self.first = time.time()
        self.alive = time.time()   # last time seen in queue or matched to counters

    def public(self):
        return {"doc": self.doc, "paper": self.paper or (self.kind or "?"), "tray": self.tray,
                "color": self.color, "duplex": self.duplex, "pages": self.pages, "copies": self.copies}


class Spooler(threading.Thread):
    """Polls the Windows print queue. Works without pywin32 in 'off' mode."""
    KEEP = 180  # seconds a job stays matchable after it was last seen / matched

    def __init__(self, on_activity):
        threading.Thread.__init__(self, daemon=True)
        self.jobs = {}
        self.lock = threading.Lock()
        self.mode = "off"
        self.note = ""
        self.printers = []
        self.on_activity = on_activity
        self._caps = {}
        try:
            import win32print  # noqa
            self.w = win32print
            self.mode = "on"
        except Exception as e:
            self.w = None
            self.note = "pywin32 not available (%s) - document names and tray detection are off" % e

    # -- helpers
    def _capnames(self, name, port):
        if name in self._caps:
            return self._caps[name]
        bins, papers = {}, {}
        try:
            ids = self.w.DeviceCapabilities(name, port, 6)     # DC_BINS
            nms = self.w.DeviceCapabilities(name, port, 12)    # DC_BINNAMES
            bins = dict(zip(ids or [], [str(n).strip("\x00 ") for n in (nms or [])]))
        except Exception:
            pass
        try:
            ids = self.w.DeviceCapabilities(name, port, 2)     # DC_PAPERS
            nms = self.w.DeviceCapabilities(name, port, 16)    # DC_PAPERNAMES
            papers = dict(zip(ids or [], [str(n).strip("\x00 ") for n in (nms or [])]))
        except Exception:
            pass
        self._caps[name] = (bins, papers)
        return self._caps[name]

    @staticmethod
    def kind_of(paper_id, name, w, l):
        """Paper kind from DEVMODE paper id / form name / size in 0.1 mm."""
        n = (name or "").lower()
        if paper_id == 9 or re.search(r"\ba4\b", n):
            return "A4"
        if paper_id == 8 or re.search(r"\ba3\b", n):
            return "A3"
        if paper_id == 5 or "legal" in n:
            return "Legal"
        if re.search(r"12\s*[x×]\s*18", n):
            return "12x18"
        if w and l:
            a, b = sorted((int(w), int(l)))
            if abs(a - 3048) < 40 and abs(b - 4572) < 40:
                return "12x18"
            if abs(a - 2970) < 30 and abs(b - 4200) < 30:
                return "A3"
            if abs(a - 2100) < 30 and abs(b - 2970) < 30:
                return "A4"
            if abs(a - 2159) < 30 and abs(b - 3556) < 30:
                return "Legal"
        if paper_id or n:
            return "OTHER"
        return None

    @staticmethod
    def bypass_of(bin_id, bin_name):
        n = (bin_name or "").lower()
        if re.search(r"mp|multi|bypass|manual", n) or bin_id in (4, 5):   # DMBIN_MANUAL / ENVMANUAL
            return True
        if re.search(r"auto", n) or bin_id in (7, 15) or (not n and not bin_id):
            return None
        if re.search(r"cassette|tray|drawer|upper|lower|middle", n) or bin_id in (1, 2, 3):
            return False
        return None

    def _targets(self):
        rx = re.compile(CFG.get("printer_name_match") or ".", re.I)
        flags = self.w.PRINTER_ENUM_LOCAL | self.w.PRINTER_ENUM_CONNECTIONS
        allp = self.w.EnumPrinters(flags, None, 2)
        hit = [p for p in allp if rx.search((p.get("pPrinterName") or "") + " " + (p.get("pDriverName") or ""))]
        if not hit:
            skip = re.compile(r"pdf|xps|onenote|fax|anydesk|send to", re.I)
            hit = [p for p in allp if not skip.search(p.get("pPrinterName") or "")]
        return [(p["pPrinterName"], p.get("pPortName") or "") for p in hit]

    def _scan(self, name, port, handle):
        try:
            jobs = self.w.EnumJobs(handle, 0, 100, 2)
        except Exception:
            jobs = self.w.EnumJobs(handle, 0, 100, 1)
        now = time.time()
        for j in jobs:
            key = (name, j.get("JobId"), j.get("Submitted") and str(j.get("Submitted")))
            with self.lock:
                job = self.jobs.get(key)
                if job is None:
                    job = self.jobs[key] = Job(key)
                    new = True
                else:
                    new = False
                job.alive = now
                job.doc = j.get("pDocument") or job.doc
                job.pages = max(job.pages, int(j.get("TotalPages") or 0))
                dm = j.get("pDevMode")
                if dm is not None:
                    try:
                        bins, papers = self._capnames(name, port)
                        pid = int(getattr(dm, "PaperSize", 0) or 0)
                        pname = papers.get(pid) or str(getattr(dm, "FormName", "") or "").strip("\x00 ")
                        job.paper = pname or job.paper
                        k = self.kind_of(pid, pname, getattr(dm, "PaperWidth", 0), getattr(dm, "PaperLength", 0))
                        job.kind = k or job.kind
                        c = int(getattr(dm, "Color", 0) or 0)
                        job.color = True if c == 2 else (False if c == 1 else job.color)
                        d = int(getattr(dm, "Duplex", 0) or 0)
                        job.duplex = (d in (2, 3)) if d else job.duplex
                        job.copies = int(getattr(dm, "Copies", 1) or 1)
                        b = int(getattr(dm, "DefaultSource", 0) or 0)
                        job.tray = bins.get(b, "") or job.tray
                        job.bypass = self.bypass_of(b, bins.get(b, ""))
                    except Exception:
                        pass
            if new:
                print("[queue] %s | %s | %s | tray: %s" % (job.doc, job.paper or job.kind or "?",
                      {True: "colour", False: "b&w", None: "?"}[job.color], job.tray or "auto/unknown"))
            self.on_activity()

    def run(self):
        if self.mode != "on":
            return
        handles = {}
        last_enum = 0
        while True:
            try:
                if time.time() - last_enum > 30:
                    targets = self._targets()
                    self.printers = [t[0] for t in targets]
                    for name, port in targets:
                        if name not in handles:
                            try:
                                handles[name] = (self.w.OpenPrinter(name), port)
                            except Exception:
                                pass
                    for name in list(handles):
                        if name not in self.printers:
                            try:
                                self.w.ClosePrinter(handles[name][0])
                            except Exception:
                                pass
                            del handles[name]
                    last_enum = time.time()
                for name, (h, port) in list(handles.items()):
                    try:
                        self._scan(name, port, h)
                    except Exception:
                        try:
                            self.w.ClosePrinter(h)
                        except Exception:
                            pass
                        del handles[name]
                        last_enum = 0
                self._expire()
            except Exception as e:
                self.note = "queue watcher error: %s" % e
                time.sleep(3)
            time.sleep(0.25)

    def _expire(self):
        cut = time.time() - 3600
        with self.lock:
            for k in [k for k, j in self.jobs.items() if j.alive < cut]:
                del self.jobs[k]

    # -- used by the tracker
    def match(self, kind, strict=False):
        """Most recent still-relevant job on the given paper kind."""
        now = time.time()
        with self.lock:
            act = [j for j in self.jobs.values() if now - j.alive <= self.KEEP]
            c = [j for j in act if j.kind == kind]
            if not c and not strict:
                c = [j for j in act if j.kind is None]
            if not c:
                return None
            j = max(c, key=lambda x: x.first)
            j.alive = now
            return j

    def active(self):
        now = time.time()
        with self.lock:
            return [j.public() for j in self.jobs.values() if now - j.alive <= 20]


# ---------------------------------------------------------------- tracker
class Tracker(threading.Thread):
    MERGE = 25        # seconds: counter ticks of the same job are merged into one row
    MERGE_DOC = 120   # ...longer when the document name is known (slow master prints)

    def __init__(self):
        threading.Thread.__init__(self, daemon=True)
        self.hot_until = time.time() + 60
        self.online = False
        self.last_ok = 0
        self.error = ""
        self.life = None
        self.spool = Spooler(self.poke)

    def poke(self):
        self.hot_until = time.time() + 150

    def run(self):
        while True:
            try:
                snap = read_printer()
                self.online, self.last_ok, self.error = True, time.time(), ""
                self.life = sum(snap["bw"]) + sum(snap["col"])
                self.process(snap)
            except Exception as e:
                self.online, self.error = False, str(e)
            time.sleep(1.5 if time.time() < self.hot_until else 6)

    # -- allocate one counter delta to categories
    def process(self, snap):
        prev = kv_get("counters")
        if prev is None:
            kv_set("counters", snap)
            kv_set("since", time.time())
            print("[start] baseline taken - tracking starts now")
            return
        dbw = [n - o for n, o in zip(snap["bw"], prev["bw"])]
        dco = [n - o for n, o in zip(snap["col"], prev["col"])]
        if min(dbw + dco) < 0:      # counters were reset on the machine
            kv_set("counters", snap)
            return
        total = sum(dbw) + sum(dco)

        dscan = 0
        if snap.get("scan") is not None and prev.get("scan") is not None:
            dscan = snap["scan"] - prev["scan"]
            if dscan < 0:           # scan counter was reset on the machine
                dscan = 0

        if total == 0:
            if dscan > 0:
                with _db_lock:
                    _db.execute("INSERT INTO scan_events(ts, n) VALUES(?, ?)", (time.time(), dscan))
                    _db.execute("INSERT INTO kv(k,v) VALUES('counters',?) ON CONFLICT(k) DO UPDATE SET v=excluded.v",
                                (json.dumps(snap),))
                    _db.commit()
                self.poke()
            elif snap != prev:
                kv_set("counters", snap)
            return
        self.poke()
        ddup = max(0, snap["dup2"] - prev.get("dup2", snap["dup2"]))
        via = "print"
        if snap.get("copy") is not None and prev.get("copy") is not None:
            if snap["copy"] - prev["copy"] >= total:
                via = "copy"
        sp = self.spool if via == "print" else None
        out = []  # (category, pages, size, job, tray_src)

        # A4 - split single / duplex
        a, c = dbw[I_A4], dco[I_A4]
        if a or c:
            dup = min(ddup, a + c)
            if c == 0:
                dup_c = 0
            elif a == 0:
                dup_c = dup
            else:
                dup_c = min(c, int(round(dup * c / float(a + c))))
            dup_b = min(a, dup - dup_c)
            job = sp.match("A4") if sp else None
            for cat, n in (("a4_bw_duplex", dup_b), ("a4_bw_single", a - dup_b),
                           ("a4_color_duplex", dup_c), ("a4_color_single", c - dup_c)):
                if n > 0:
                    out.append((cat, n, "A4", job, ""))

        # A3 - standard tray or bypass
        if dbw[I_A3] or dco[I_A3]:
            job = sp.match("A3") if sp else None
            if job is not None and job.bypass is not None:
                byp, src = job.bypass, "driver"
            else:
                byp, src = (kv_get("a3_default", "bypass") == "bypass"), "default"
            where = "bypass" if byp else "tray"
            if dbw[I_A3]:
                out.append(("a3_%s_bw" % where, dbw[I_A3], "A3", job, src))
            if dco[I_A3]:
                out.append(("a3_%s_color" % where, dco[I_A3], "A3", job, src))

        # Legal - B&W is an offset master
        if dbw[I_LEGAL]:
            out.append(("master_legal", dbw[I_LEGAL], "Legal", sp.match("Legal") if sp else None, ""))
        if dco[I_LEGAL]:
            out.append(("other_color", dco[I_LEGAL], "Legal", sp.match("Legal") if sp else None, ""))

        # 12x18 / Legal masters and other custom sizes - the printer's own size
        # sensor buckets any non-standard paper (most offset masters) into these
        # two generic "Other" counters instead of the named Legal/12x18 slots,
        # so which master (if any) each one defaults to is a per-shop setting
        # (Pricing & settings -> "Other paper counters"); OTH_DEFAULT below is
        # only the out-of-the-box guess, kept for anyone who hasn't set it.
        OTH_DEFAULT = {I_OTH1: "master_legal", I_OTH2: "master_12x18"}
        OTH_LABEL = {"master_12x18": "12x18", "master_legal": "Legal"}
        for i in (I_OTH1, I_OTH2):
            if dbw[i]:
                setting = kv_get("other%d_default" % (1 if i == I_OTH1 else 2), OTH_DEFAULT[i])
                job = None
                if sp:      # the document's own paper decides when it is clearly a master size
                    job = sp.match("12x18", strict=True) or sp.match("Legal", strict=True)
                if job is not None:
                    cat = "master_12x18" if job.kind == "12x18" else "master_legal"
                else:
                    # PDF readers often send a custom page size instead of "Legal", so an
                    # unrecognised paper name must not override the shop's setting
                    job = sp.match("OTHER") if sp else None
                    cat = setting if setting in ("master_12x18", "master_legal") else "other_bw"
                label = OTH_LABEL.get(cat) or (SIZES[i] + (" · " + job.paper if job and job.paper else ""))
                out.append((cat, dbw[i], label, job, ""))
            if dco[i]:
                job = None
                if sp:
                    job = sp.match("12x18", strict=True) or sp.match("OTHER")
                out.append(("other_color", dco[i], (job.paper if job and job.paper else SIZES[i]), job, ""))

        # every remaining size
        for i, name in enumerate(SIZES):
            if i in (I_A4, I_A3, I_LEGAL, I_OTH1, I_OTH2):
                continue
            if dbw[i]:
                out.append(("other_bw", dbw[i], name, sp.match("OTHER") if sp else None, ""))
            if dco[i]:
                out.append(("other_color", dco[i], name, sp.match("OTHER") if sp else None, ""))

        now = time.time()
        with _db_lock:
            for cat, n, size, job, src in out:
                doc = job.doc if job else ("Copy at machine panel" if via == "copy" else "")
                app, doc = app_from_doc(doc) if job else ("", doc)
                row = _db.execute(
                    "SELECT id FROM events WHERE category=? AND IFNULL(doc,'')=? AND via=? AND edited=0 "
                    "AND ts_end>=? ORDER BY id DESC LIMIT 1", (cat, doc, via, now - (self.MERGE_DOC if doc else self.MERGE))).fetchone()
                if row:
                    _db.execute("UPDATE events SET pages=pages+?, ts_end=? WHERE id=?", (n, now, row["id"]))
                else:
                    _db.execute("INSERT INTO events(ts,ts_end,category,pages,doc,app,via,size,tray_src) "
                                "VALUES(?,?,?,?,?,?,?,?,?)", (now, now, cat, n, doc, app, via, size, src))
                print("[count] +%d  %s  %s" % (n, cat, doc))
            if dscan > 0:
                _db.execute("INSERT INTO scan_events(ts, n) VALUES(?, ?)", (now, dscan))
                print("[count] +%d  scan" % dscan)
            _db.execute("INSERT INTO kv(k,v) VALUES('counters',?) ON CONFLICT(k) DO UPDATE SET v=excluded.v",
                        (json.dumps(snap),))
            _db.commit()


TRACKER = None


# ---------------------------------------------------------------- api
def range_bounds(rng):
    """(start, end) timestamps. Custom range: d20261115_20261119 (both days included)."""
    m = re.match(r"d(\d{8})_(\d{8})$", rng or "")
    if m:
        try:
            a = datetime.strptime(m.group(1), "%Y%m%d")
            b = datetime.strptime(m.group(2), "%Y%m%d")
            if b < a:
                a, b = b, a
            return a.timestamp(), (b + timedelta(days=1)).timestamp()
        except ValueError:
            pass
    return range_start(rng), float("inf")


def range_start(rng):
    now = datetime.now()
    day = now.replace(hour=0, minute=0, second=0, microsecond=0)
    if rng == "today":
        return day.timestamp()
    if rng == "7":
        return (day - timedelta(days=6)).timestamp()
    if rng == "month":
        return day.replace(day=1).timestamp()
    return 0.0


def _f(v, default=0.0):
    try:
        return max(0.0, float(v))
    except (TypeError, ValueError):
        return default


def _ts(v):
    """'YYYY-MM-DD' -> timestamp at noon local time; anything else -> now."""
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").replace(hour=12).timestamp()
    except Exception:
        return time.time()


def _txt(v, n=200):
    return str(v or "").strip()[:n]


def _ym(dt):
    return dt.year * 12 + dt.month - 1


def _ym_ts(ym, day=1):
    return datetime(ym // 12, ym % 12 + 1, day, 12).timestamp()


def run_recurring():
    """Add this month's entry (and any missed months) for every monthly expense."""
    cur = _ym(datetime.now())
    with _db_lock:
        for r in _db.execute("SELECT * FROM recurring").fetchall():
            for ym in range(r["last_ym"] + 1, cur + 1):
                _db.execute("INSERT INTO expenses(ts,title,amount,paid_by) VALUES(?,?,?,?)",
                            (_ym_ts(ym), r["title"], r["amount"], r["paid_by"] or "business"))
            if r["last_ym"] < cur:
                _db.execute("UPDATE recurring SET last_ym=? WHERE id=?", (cur, r["id"]))
        _db.commit()


def api_state(rng, inner=False):
    start, end = range_bounds(rng)
    endq = min(end, 1e12)
    run_recurring()
    with _db_lock:
        rates = {r["category"]: float(r["rate"]) for r in _db.execute("SELECT * FROM rates")}
        manual = {r["paper"]: float(r["rate"]) for r in _db.execute("SELECT * FROM paper_rates")}
        rows = [dict(r) for r in _db.execute(
            "SELECT id, category, pages, job_id, sheets, personal FROM events WHERE ts>=? AND ts<?", (start, endq))]
        personal_all = [dict(r) for r in _db.execute(
            "SELECT id, ts, category, pages, sheets, doc FROM events WHERE personal=1")]
        evs = [dict(r) for r in _db.execute(
            "SELECT * FROM events WHERE ts>=? AND ts<? ORDER BY ts_end DESC, id DESC LIMIT 300", (start, endq))]
        used_all = [dict(r) for r in _db.execute(
            "SELECT category, SUM(COALESCE(sheets, pages)) p, SUM((pages+1)/2) s FROM events GROUP BY category")]
        items = [dict(r) for r in _db.execute("SELECT * FROM job_items ORDER BY id")]
        masters = [dict(r) for r in _db.execute(
            "SELECT id, ts, category, pages, sheets, job_id, doc FROM events "
            "WHERE category IN ('master_legal','master_12x18') ORDER BY ts DESC LIMIT 300")]
        proofs = [dict(r) for r in _db.execute(
            "SELECT id, ts, category, pages, job_id, doc FROM events "
            "WHERE category NOT IN ('master_legal','master_12x18') AND (job_id IS NOT NULL OR ts>=?) "
            "ORDER BY ts DESC LIMIT 400", (time.time() - 45 * 86400,))]
        purchases = [dict(r) for r in _db.execute("SELECT * FROM purchases ORDER BY ts DESC, id DESC LIMIT 500")]
        jobs = [dict(r) for r in _db.execute("SELECT * FROM jobs ORDER BY ts DESC, id DESC LIMIT 500")]
        expenses = [dict(r) for r in _db.execute("SELECT * FROM expenses ORDER BY ts DESC, id DESC LIMIT 500")]
        recurring = [dict(r) for r in _db.execute("SELECT id, title, amount, paid_by FROM recurring ORDER BY id")]
        capital = [dict(r) for r in _db.execute("SELECT * FROM capital ORDER BY ts DESC, id DESC")]
        scans = _db.execute("SELECT SUM(n) n FROM scan_events WHERE ts>=? AND ts<?", (start, endq)).fetchone()["n"] or 0

    # cost of one sheet = average of everything bought; manual value only when nothing was bought
    bought = {k: [0.0, 0.0] for k in PAPER_KEYS}     # sheets, money
    spend = 0.0
    for p in purchases:
        p["sheets"] = int(round(p["packets"] * p["sheets_per_packet"]))
        p["per_sheet"] = round(p["total_cost"] / p["sheets"], 4) if p["sheets"] else 0.0
        if p["paper"] in bought:
            bought[p["paper"]][0] += p["sheets"]
            bought[p["paper"]][1] += p["total_cost"]
        if start <= p["ts"] < end:
            spend += p["total_cost"]
    pcost, psrc = {}, {}
    for k in PAPER_KEYS:
        if bought[k][0] > 0:
            pcost[k], psrc[k] = bought[k][1] / bought[k][0], "purchase"
        elif manual.get(k, 0) > 0:
            pcost[k], psrc[k] = manual[k], "manual"
        else:
            pcost[k], psrc[k] = 0.0, "none"

    def money(cat, pages, sheets=None, linked=False):
        sh = sheets if sheets is not None else sheets_of(cat, pages)
        sale = round(pages * rates.get(cat, 0.0), 2)        # selling price already includes paper
        # linked to an offset job: the paper is charged inside that job instead; and an
        # ordinary print linked to a job is a free proof for that customer, not a sale
        pc = 0.0 if linked else round(sh * pcost[paper_of(cat)], 2)
        if linked and not cat.startswith("master_"):
            sale = 0.0
        return sh, sale, pc

    def master_charge(e):
        """What a master print costs the offset job it belongs to: sheets x paper + sides x rate."""
        sh = e["sheets"] if e["sheets"] is not None else sheets_of(e["category"], e["pages"])
        paper = round(sh * pcost[paper_of(e["category"])], 2)
        work = round(e["pages"] * rates.get(e["category"], 0.0), 2)
        return sh, paper, work

    agg = {k: [0, 0, 0, 0.0, 0.0] for k in CAT_KEYS}   # pages, sheets, entries, sales, paper cost
    pers_pages = 0
    pers_cost = 0.0
    for r in rows:
        if r.get("personal"):
            # own use: no sale, and its paper is money the owner took out, not a business cost
            sh, _, pc = money(r["category"], r["pages"], r["sheets"])
            pers_pages += r["pages"]; pers_cost += pc
            continue
        a = agg.get(r["category"])
        if a is None:
            continue
        sh, sale, pc = money(r["category"], r["pages"], r["sheets"], bool(r["job_id"]))
        a[0] += r["pages"]; a[1] += sh; a[2] += 1; a[3] += sale; a[4] += pc

    cats, tp, tsale, tpaper = [], 0, 0.0, 0.0
    psheets = {k: 0 for k in PAPER_KEYS}
    for key, group, label, colour in CATS:
        p, sh, n, sale, pc = agg[key]
        tp += p; tsale += sale; tpaper += pc
        pk = paper_of(key)
        psheets[pk] += sh
        rate = rates.get(key, 0.0)
        per_page_cost = pcost[pk] / (2.0 if key.endswith("duplex") else 1.0)
        cats.append({"key": key, "group": group, "label": label, "colour": colour,
                     "pages": p, "sheets": sh, "jobs": n, "paper": pk, "rate": rate,
                     "paper_cost_sheet": round(pcost[pk], 4),
                     "margin_page": round(rate - per_page_cost, 4),
                     "sales": round(sale, 2), "paper_cost": round(pc, 2),
                     "profit": round(sale - pc, 2), "amount": round(sale, 2)})

    used_tot = {k: 0 for k in PAPER_KEYS}
    for r in used_all:
        if r["category"] in CAT_KEYS:
            used_tot[paper_of(r["category"])] += (r["s"] if r["category"].endswith("duplex") else r["p"]) or 0
    papers = []
    for k, lbl in PAPERS:
        papers.append({"key": k, "label": lbl, "cost": round(pcost[k], 4), "cost_src": psrc[k],
                       "manual": manual.get(k, 0.0), "sheets": psheets[k],
                       "amount": round(psheets[k] * pcost[k], 2),
                       "bought": int(bought[k][0]), "bought_cost": round(bought[k][1], 2),
                       "used_total": used_tot[k], "stock": int(bought[k][0]) - used_tot[k]})

    for e in evs:
        sh, sale, pc = money(e["category"], e["pages"], e.get("sheets"), bool(e.get("job_id")))
        if e.get("personal"):
            sale = 0.0
        e.update(sheets=sh, amount=sale, paper_cost=pc, profit=round(sale - pc, 2))

    job_ids = {j["id"] for j in jobs}
    by_items, by_masters, free_masters = {}, {}, []
    for it in items:
        by_items.setdefault(it["job_id"], []).append(it)
    for m in masters:
        sh, paper, work = master_charge(m)
        m.update(sheets=sh, paper_cost=paper, work=work, charge=round(paper + work, 2))
        if m["job_id"] in job_ids:
            by_masters.setdefault(m["job_id"], []).append(m)
        else:
            free_masters.append(m)
    by_proofs, free_prints = {}, []
    clabel = {k: (g + " · " + l) for k, g, l, c in CATS}
    for e in proofs:
        sh = sheets_of(e["category"], e["pages"])
        e.update(sheets=sh, label=clabel.get(e["category"], e["category"]),
                 paper_cost=round(sh * pcost[paper_of(e["category"])], 2),
                 would_sell=round(e["pages"] * rates.get(e["category"], 0.0), 2))
        if e["job_id"] in job_ids:
            by_proofs.setdefault(e["job_id"], []).append(e)
        elif not e["job_id"]:
            free_prints.append(e)
    jsale = jsup = jother = jitems = jmaster = jproof = recv = pay = 0.0
    for j in jobs:
        j["items"] = by_items.get(j["id"], [])
        j["masters"] = by_masters.get(j["id"], [])
        j["items_cost"] = round(sum(i["amount"] for i in j["items"]), 2)
        j["master_cost"] = round(sum(m["charge"] for m in j["masters"]), 2)
        j["proofs"] = by_proofs.get(j["id"], [])
        j["proof_cost"] = round(sum(e["paper_cost"] for e in j["proofs"]), 2)
        j["cost"] = round(j["supplier_amount"] + j["other_cost"] + j["items_cost"] + j["master_cost"] + j["proof_cost"], 2)
        j["profit"] = round(j["customer_amount"] - j["cost"], 2)
        j["receivable"] = round(j["customer_amount"] - j["customer_paid"], 2)
        j["payable"] = round(j["supplier_amount"] - j["supplier_paid"], 2)
        recv += j["receivable"]; pay += j["payable"]          # outstanding, all time
        if start <= j["ts"] < end:
            jsale += j["customer_amount"]; jsup += j["supplier_amount"]; jother += j["other_cost"]
            jitems += j["items_cost"]; jmaster += j["master_cost"]; jproof += j["proof_cost"]
    exp = sum(x["amount"] for x in expenses if start <= x["ts"] < end)
    sales = tsale + jsale
    direct = tpaper + jsup + jother + jitems + jmaster + jproof
    gross = sales - direct
    pnl = {"print_sales": round(tsale, 2), "job_sales": round(jsale, 2), "sales": round(sales, 2),
           "paper_cost": round(tpaper, 2), "supplier_cost": round(jsup, 2), "job_other": round(jother, 2),
           "job_items": round(jitems, 2), "job_masters": round(jmaster, 2), "job_proofs": round(jproof, 2),
           "direct": round(direct, 2), "gross": round(gross, 2), "expenses": round(exp, 2),
           "net": round(gross - exp, 2), "purchase_spend": round(spend, 2),
           "receivable": round(recv, 2), "payable": round(pay, 2),
           "bills_due": round(sum(x["amount"] for x in expenses if x.get("paid") == 0), 2),
           "personal_pages": pers_pages, "personal_cost": round(pers_cost, 2)}
    # owner's money: cash he put in (lump sums + expenses paid from his own pocket) - cash he took out.
    # This is shown since the start, whatever period is selected; it does not change profit.
    owner = None
    if not inner:
        net_all = pnl["net"] if rng == "all" else api_state("all", inner=True)["pnl"]["net"]
        own_exp = [x for x in expenses if x.get("paid_by") == "owner" and x.get("paid") != 0]
        own_pur = [x for x in purchases if x.get("paid_by") == "owner"]
        put_in = (sum(c["amount"] for c in capital if c["kind"] == "in") + sum(x["amount"] for x in own_exp)
                  + sum(x["total_cost"] for x in own_pur))
        for e in personal_all:
            e["paper_cost"] = money(e["category"], e["pages"], e["sheets"])[2]
        took = sum(c["amount"] for c in capital if c["kind"] == "out") + sum(e["paper_cost"] for e in personal_all)
        owner = {"put_in": round(put_in, 2), "took_out": round(took, 2), "net_all": round(net_all, 2),
                 "own_expenses": round(sum(x["amount"] for x in own_exp), 2),
                 "own_purchases": round(sum(x["total_cost"] for x in own_pur), 2),
                 "personal_prints": [{"ts": e["ts"], "doc": e["doc"], "pages": e["pages"], "cost": e["paper_cost"]} for e in personal_all],
                 "lump_in": round(sum(c["amount"] for c in capital if c["kind"] == "in"), 2),
                 "balance": round(put_in - took + net_all, 2)}
    t = TRACKER
    return {
        "now": time.time(), "range": rng, "start": start, "end": (None if end == float("inf") else end), "since": kv_get("since"),
        "printer": {"online": t.online, "last_ok": t.last_ok, "error": t.error,
                    "url": CFG["printer_url"], "lifetime": t.life},
        "queue": {"mode": t.spool.mode, "note": t.spool.note, "printers": t.spool.printers,
                  "active": t.spool.active()},
        "settings": {"a3_default": kv_get("a3_default", "bypass"),
                     "other1_default": kv_get("other1_default", "master_legal"),
                     "other2_default": kv_get("other2_default", "master_12x18")},
        "cats": cats, "papers": papers, "total_pages": tp, "pnl": pnl, "scans": int(scans),
        "total_amount": round(tsale, 2), "events": evs,
        "free_masters": free_masters[:60], "free_prints": free_prints[:25], "purchases": purchases, "jobs": jobs, "expenses": expenses, "recurring": recurring, "capital": capital, "owner": owner, "next_bill": _next_bill(),
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "PrintTracker/1.0"

    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if not isinstance(body, bytes):
            body = json.dumps(body).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path, _, q = self.path.partition("?")
        if path in ("/", "/index.html"):
            try:
                with open(os.path.join(BASE, "dashboard.html"), "rb") as f:
                    return self._send(200, f.read(), "text/html; charset=utf-8")
            except OSError:
                return self._send(500, b"dashboard.html is missing next to app.py", "text/plain")
        if path == "/api/state":
            m = re.search(r"range=(\w+)", q)
            return self._send(200, api_state(m.group(1) if m else "today"))
        self._send(404, {"error": "not found"})

    def do_POST(self):
        host = (self.headers.get("Host") or "").split(":")[0]
        if host not in ("localhost", "127.0.0.1"):
            return self._send(403, {"error": "local only"})
        try:
            n = int(self.headers.get("Content-Length") or 0)
            data = json.loads(self.rfile.read(n).decode("utf-8") or "{}")
        except Exception:
            return self._send(400, {"error": "bad json"})
        path = self.path.split("?")[0]
        if path == "/api/rates":
            with _db_lock:
                for k, v in (data.get("rates") or {}).items():
                    if k in CAT_KEYS:
                        try:
                            v = max(0.0, float(v))
                        except (TypeError, ValueError):
                            continue
                        _db.execute("UPDATE rates SET rate=? WHERE category=?", (v, k))
                for k, v in (data.get("paper") or {}).items():
                    if k in PAPER_KEYS:
                        try:
                            v = max(0.0, float(v))
                        except (TypeError, ValueError):
                            continue
                        _db.execute("UPDATE paper_rates SET rate=? WHERE paper=?", (v, k))
                _db.commit()
            return self._send(200, {"ok": True})
        if path == "/api/settings":
            if data.get("a3_default") in ("bypass", "tray"):
                kv_set("a3_default", data["a3_default"])
            if data.get("other1_default") in ("none", "master_12x18", "master_legal"):
                kv_set("other1_default", data["other1_default"])
            if data.get("other2_default") in ("none", "master_12x18", "master_legal"):
                kv_set("other2_default", data["other2_default"])
            return self._send(200, {"ok": True})
        if path == "/api/event":
            if "personal" in data:
                with _db_lock:
                    _db.execute("UPDATE events SET personal=? WHERE id=?",
                                (1 if data.get("personal") else 0, int(data.get("id") or 0)))
                    _db.commit()
                return self._send(200, {"ok": True})
            if data.get("category") in CAT_KEYS:
                with _db_lock:
                    _db.execute("UPDATE events SET category=?, edited=1, tray_src='manual' WHERE id=?",
                                (data["category"], int(data.get("id") or 0)))
                    _db.commit()
            return self._send(200, {"ok": True})
        if path == "/api/purchase":
            # one bill = several lines (A4 packets, Legal masters, 12x18 masters ...)
            lines = data.get("lines") if isinstance(data.get("lines"), list) else [data]
            rows = []
            for n, ln in enumerate(lines, 1):
                loose = ln.get("mode") == "loose"
                if loose:       # bought as single sheets, not by the packet
                    packets, spp = 1.0, int(_f(ln.get("loose_sheets")))
                else:
                    packets, spp = _f(ln.get("packets")), int(_f(ln.get("sheets_per_packet")))
                where = (" (line %d)" % n) if len(lines) > 1 else ""
                if ln.get("paper") not in PAPER_KEYS or packets <= 0 or spp <= 0:
                    return self._send(400, {"error": ("Choose the paper and enter the number of sheets" if loose else
                                                      "Choose the paper and enter packets and sheets per packet") + where + "."})
                if _f(ln.get("total_cost")) <= 0:
                    return self._send(400, {"error": "Enter the amount paid" + where + "."})
                rows.append((ln["paper"], packets, spp, _f(ln.get("total_cost")), 1 if loose else 0, _txt(ln.get("note")),
                             int(_f(ln.get("id")))))
            if not rows:
                return self._send(400, {"error": "Add at least one line."})
            ts, sup, bill, note = _ts(data.get("date")), _txt(data.get("supplier")), _txt(data.get("bill_no"), 60), _txt(data.get("note"))
            by = "owner" if data.get("paid_by") == "owner" else "business"
            with _db_lock:
                if not bill:
                    bill = _next_bill()
                # editing a saved bill: lines with an id are updated, removed lines deleted
                for rid in data.get("remove_ids") or []:
                    _db.execute("DELETE FROM purchases WHERE id=?", (int(_f(rid)),))
                for paper, packets, spp, cost, loose, lnote, rid in rows:
                    vals = (ts, paper, sup, packets, spp, cost, lnote or note, loose, bill, by)
                    if rid:
                        _db.execute("UPDATE purchases SET ts=?,paper=?,supplier=?,packets=?,sheets_per_packet=?,total_cost=?,"
                                    "note=?,loose=?,bill_no=?,paid_by=? WHERE id=?", vals + (rid,))
                    else:
                        _db.execute("INSERT INTO purchases(ts,paper,supplier,packets,sheets_per_packet,total_cost,note,loose,bill_no,paid_by) "
                                    "VALUES(?,?,?,?,?,?,?,?,?,?)", vals)
                _db.commit()
            return self._send(200, {"ok": True})
        if path == "/api/job":
            if not (_txt(data.get("customer")) or _txt(data.get("title"))):
                return self._send(400, {"error": "Enter the customer or the job name."})
            vals = (_ts(data.get("date")), _txt(data.get("customer")), _txt(data.get("title")), _txt(data.get("qty"), 60),
                    _f(data.get("customer_amount")), _f(data.get("customer_paid")), _txt(data.get("supplier")),
                    _f(data.get("supplier_amount")), _f(data.get("supplier_paid")), _f(data.get("other_cost")),
                    _txt(data.get("note")))
            with _db_lock:
                if data.get("id"):
                    _db.execute("UPDATE jobs SET ts=?,customer=?,title=?,qty=?,customer_amount=?,customer_paid=?,"
                                "supplier=?,supplier_amount=?,supplier_paid=?,other_cost=?,note=? WHERE id=?",
                                vals + (int(data["id"]),))
                    jid = int(data["id"])
                else:
                    jid = _db.execute("INSERT INTO jobs(ts,customer,title,qty,customer_amount,customer_paid,supplier,"
                                      "supplier_amount,supplier_paid,other_cost,note) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                      vals).lastrowid
                _db.commit()
            return self._send(200, {"ok": True, "id": jid})
        if path == "/api/job_item":
            qty, rate = _f(data.get("qty"), 1) or 1, _f(data.get("rate"))
            if not _txt(data.get("title")) or qty <= 0 or rate <= 0:
                return self._send(400, {"error": "Enter what it is, the quantity and the price of one."})
            with _db_lock:
                if not _db.execute("SELECT 1 FROM jobs WHERE id=?", (int(data.get("job_id") or 0),)).fetchone():
                    return self._send(400, {"error": "Job not found."})
                _db.execute("INSERT INTO job_items(job_id,title,qty,rate,amount) VALUES(?,?,?,?,?)",
                            (int(data["job_id"]), _txt(data.get("title")), qty, rate, round(qty * rate, 2)))
                _db.commit()
            return self._send(200, {"ok": True})
        if path == "/api/job_link":
            eid = int(data.get("event_id") or 0)
            with _db_lock:
                e = _db.execute("SELECT pages, category FROM events WHERE id=?", (eid,)).fetchone()
                if not e:
                    return self._send(400, {"error": "Print not found."})
                if "job_id" in data:
                    jid = int(data.get("job_id") or 0) or None
                    _db.execute("UPDATE events SET job_id=? WHERE id=?", (jid, eid))
                if "sheets" in data and e["category"].startswith("master_"):
                    # 0 is allowed: sides printed on the back of a sheet already counted in another entry
                    sh = None if data.get("sheets") is None else max(0, min(e["pages"], int(_f(data.get("sheets")))))
                    _db.execute("UPDATE events SET sheets=? WHERE id=?", (sh, eid))
                _db.commit()
            return self._send(200, {"ok": True})
        if path == "/api/expense":
            if _f(data.get("amount")) <= 0:
                return self._send(400, {"error": "Enter the amount."})
            ts, title, amount = _ts(data.get("date")), _txt(data.get("title")), _f(data.get("amount"))
            due = _ts(data.get("due")) if data.get("due") else None
            by = "owner" if data.get("paid_by") == "owner" else "business"
            paid = 1 if data.get("paid") else 0
            months = min(24, max(1, int(_f(data.get("months"), 1) or 1)))
            d = datetime.fromtimestamp(ts)
            with _db_lock:
                if data.get("id"):
                    # edit one saved expense; due date / paid apply to the whole split bill
                    eid = int(_f(data.get("id")))
                    r = _db.execute("SELECT grp FROM expenses WHERE id=?", (eid,)).fetchone()
                    if not r:
                        return self._send(400, {"error": "Expense not found."})
                    g = r["grp"] or eid
                    _db.execute("UPDATE expenses SET ts=?, title=?, amount=? WHERE id=?", (ts, title, amount, eid))
                    _db.execute("UPDATE expenses SET paid_by=? WHERE id=? OR grp=?", (by, g, g))
                    _db.execute("UPDATE expenses SET due_ts=?, paid=?, paid_ts=CASE WHEN ?=1 THEN COALESCE(paid_ts, ?) END "
                                "WHERE id=? OR grp=?", (due, paid, paid, time.time(), g, g))
                elif data.get("repeat"):
                    # monthly expense: first entry from the chosen month, then one every month
                    _db.execute("INSERT INTO recurring(title,amount,last_ym,paid_by) VALUES(?,?,?,?)",
                                (title, amount, _ym(d) - 1, by))
                elif months > 1:
                    # one bill that covers several months: split it, ending in the bill's month
                    part = round(amount / months, 2)
                    g = None
                    for i in range(months):
                        a = round(amount - part * (months - 1), 2) if i == 0 else part
                        when = ts if i == 0 else _ym_ts(_ym(d) - i, min(d.day, 28))
                        rid = _db.execute("INSERT INTO expenses(ts,title,amount,due_ts,paid,paid_ts,grp,paid_by) VALUES(?,?,?,?,?,?,?,?)",
                                          (when, "%s (%d of %d months)" % (title, months - i, months), a,
                                           due, paid, time.time() if paid else None, g, by)).lastrowid
                        if g is None:
                            g = rid
                            _db.execute("UPDATE expenses SET grp=? WHERE id=?", (g, g))
                else:
                    _db.execute("INSERT INTO expenses(ts,title,amount,due_ts,paid,paid_ts,paid_by) VALUES(?,?,?,?,?,?,?)",
                                (ts, title, amount, due, paid, time.time() if paid else None, by))
                _db.commit()
            run_recurring()
            return self._send(200, {"ok": True})
        if path == "/api/capital":
            amt = _f(data.get("amount"))
            if amt <= 0:
                return self._send(400, {"error": "Enter the amount."})
            kind = "out" if data.get("kind") == "out" else "in"
            with _db_lock:
                if data.get("id"):
                    _db.execute("UPDATE capital SET ts=?, kind=?, amount=?, note=? WHERE id=?",
                                (_ts(data.get("date")), kind, amt, _txt(data.get("note")), int(_f(data.get("id")))))
                else:
                    _db.execute("INSERT INTO capital(ts,kind,amount,note) VALUES(?,?,?,?)",
                                (_ts(data.get("date")), kind, amt, _txt(data.get("note"))))
                _db.commit()
            return self._send(200, {"ok": True})
        if path == "/api/expense_paid":
            eid = int(_f(data.get("id")))
            with _db_lock:
                r = _db.execute("SELECT grp FROM expenses WHERE id=?", (eid,)).fetchone()
                if r:
                    g = r["grp"] or eid
                    p = 1 if data.get("paid", True) else 0
                    _db.execute("UPDATE expenses SET paid=?, paid_ts=? WHERE id=? OR grp=?",
                                (p, time.time() if p else None, g, g))
                    _db.commit()
            return self._send(200, {"ok": True})
        if path == "/api/delete":
            if data.get("kind") == "purchase_bill" and _txt(data.get("bill_no")):
                with _db_lock:
                    _db.execute("DELETE FROM purchases WHERE bill_no=?", (_txt(data.get("bill_no")),))
                    _db.commit()
                return self._send(200, {"ok": True})
            table = {"purchase": "purchases", "job": "jobs", "expense": "expenses", "recurring": "recurring",
                     "job_item": "job_items", "event": "events", "capital": "capital"}.get(data.get("kind"))
            if table:
                with _db_lock:
                    _db.execute("DELETE FROM %s WHERE id=?" % table, (int(data.get("id") or 0),))
                    if table == "jobs":
                        _db.execute("DELETE FROM job_items WHERE job_id=?", (int(data.get("id") or 0),))
                        _db.execute("UPDATE events SET job_id=NULL WHERE job_id=?", (int(data.get("id") or 0),))
                    _db.commit()
            return self._send(200, {"ok": True})
        if path == "/api/reset":
            with _db_lock:
                _db.execute("DELETE FROM events")
                _db.commit()
            kv_set("since", time.time())
            return self._send(200, {"ok": True})
        self._send(404, {"error": "not found"})


def main():
    global TRACKER
    db_init()
    TRACKER = Tracker()
    TRACKER.spool.start()
    TRACKER.start()
    port = int(CFG["port"])
    try:
        srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    except OSError as e:
        print("Could not start on port %d (%s). Is Print Tracker already running?" % (port, e))
        if CFG.get("open_browser"):
            webbrowser.open("http://localhost:%d" % port)
        return 1
    url = "http://localhost:%d" % port
    print("=" * 60)
    print(" Print Tracker is running  ->  %s" % url)
    print(" Printer: %s" % CFG["printer_url"])
    print(" Queue watcher: %s %s" % (TRACKER.spool.mode, TRACKER.spool.note))
    print(" Keep this window open. Close it to stop tracking.")
    print("=" * 60)
    if CFG.get("open_browser"):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
