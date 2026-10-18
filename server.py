"""
Port Troubleshooter - one page to troubleshoot a switch port: documented cable path
(Patchmanager), what the IP list expects there, the snagging-list history and the
live CloudVision events, side by side.

- Standard library only (Python 3.8+), ThreadingHTTPServer, default port 8092.
- Inputs are imported in the page (or dropped into data/): Patchmanager cable export (CSV),
  Patchmanager all-equipment export (CSV), IP list "View" (xlsx), Connectivity
  snagging list (xlsx). The type is detected from the content. The last import of each type
  is kept in data/ (git-ignored) and loaded again at start.
- CloudVision (read-only, optional): token in env CV_TOKEN or cv_token.txt (git-ignored),
  URL in config.json. Inventory every 10 min, events every 2 min, in the background -
  viewers never cause CloudVision requests, except the explicit "Locate" button.
- PM <-> CloudVision switch names: from the snagging list's name sheet, overrides in
  switches.json (git-ignored, editable in the page).

    python server.py               # http://<this pc>:8092/
    python server.py --port 8110
"""
import argparse
import sys
import csv
import gzip
import hashlib
import io
import json
import logging
import os
import re
import threading
import time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from logging.handlers import RotatingFileHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import model
from cloudvision import CloudVision, CloudVisionError

BASE_DIR = Path(__file__).resolve().parent
PAGE_FILE = BASE_DIR / "index.html"
DATA_DIR = BASE_DIR / "data"
META_FILE = DATA_DIR / "meta.json"
CONFIG_FILE = BASE_DIR / "config.json"
CONFIG_EXAMPLE = BASE_DIR / "config.example.json"
SWITCHES_FILE = BASE_DIR / "switches.json"
FUTURE_FILE = BASE_DIR / "future.json"
SERIALS_FILE = BASE_DIR / "serials.json"
SERIALS_EXAMPLE = BASE_DIR / "serials.example.json"
NOTES_FILE = BASE_DIR / "notes.json"
TOKEN_FILE = BASE_DIR / "cv_token.txt"
LOG_FILE = BASE_DIR / "service.log"
MAX_UPLOAD = 60 * 1024 * 1024
CV_INVENTORY_SEC = 600
CV_EVENTS_SEC = 120
FILENAME_RE = re.compile(r"^[\w .,()\[\]+&'\-]{1,160}\.(csv|xlsx|xlsm)$", re.I)
SWITCH_RE = re.compile(r"^[\w .\-/()]{1,60}$")
LOCATE_RE = re.compile(r"^[\w.:\-]{2,64}$")

log = logging.getLogger("porttrouble")
log.setLevel(logging.INFO)
_h = RotatingFileHandler(LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
_h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
log.addHandler(_h)
log.addHandler(logging.StreamHandler())

LOCK = threading.RLock()
STORE = model.Store()
CV = {"enabled": False, "url": "", "error": "", "devices": [], "events": [], "devices_at": "",
      "events_at": "", "raw_sample": ""}
CV_WAKE = threading.Event()
KIND_NAMES = {"cables": "Patchmanager cables", "equipment": "Patchmanager equipment", "portplan": "IP list",
              "snags": "snagging list", "neuron": "Neuron signal list", "cvports": "CloudVision ports"}
FIXED_KINDS = ("cables", "equipment", "portplan", "snags", "neuron")


def kind_name(key):
    """'cvports:DCBESWLIVSPI001' -> 'CloudVision ports DCBESWLIVSPI001'."""
    k, _, host = key.partition(":")
    return (KIND_NAMES.get(k, k) + (" " + host if host else ""))
MISMATCH_CACHE = {"key": None, "rows": []}
PORTS_CACHE = {"key": None, "rows": []}

WORKLIST_COLS = [("source", "Source"), ("switch", "Switch (Patchmanager)"), ("cv_switch", "Switch (CloudVision)"),
                 ("intf", "Interface"), ("what", "Problem"), ("destination", "Destination"), ("status", "Status"),
                 ("since", "Since"), ("who", "Who"), ("jira", "Jira"), ("nr", "Snag Nr")]
SHUT_COLS = [("switch", "Switch (Patchmanager)"), ("cv_switch", "Switch (CloudVision)"), ("intf", "Interface"),
             ("device", "Future device"), ("device_port", "Device port"), ("via", "Found via"), ("source", "Marked by"),
             ("note", "Note")]
CHAIN_COLS = [("severity", "Severity"), ("fault", "Fault"), ("depth", "Step"), ("from", "From device"),
              ("from_port", "From port"), ("to", "To device"), ("to_port", "To port"), ("cables", "Cables")]
PORT_COLS = [("switch", "Switch (Patchmanager)"), ("cv_switch", "Switch (CloudVision)"), ("intf", "Interface"),
             ("device", "Device"), ("device_port", "Device port"), ("panels", "Patch panels"), ("cables", "Cables"),
             ("length", "Length"), ("iplist", "IP list says"), ("status", "Status"), ("snags", "Open snags"),
             ("future", "Future device")]
DEVICE_COLS = [("device", "Device"), ("kind", "Type"), ("template", "Template"), ("location", "Location"),
               ("position", "Position"), ("alias", "Alias"), ("links", "Connections"), ("future", "Future use")]
MISMATCH_COLS = [("severity", "Severity"), ("kind", "Mismatch"), ("switch", "Switch (Patchmanager)"),
                 ("cv_switch", "Switch (CloudVision)"), ("intf", "Interface"), ("source", "Source"), ("says", "Source says"),
                 ("reference", "Compared with"), ("ref_says", "That says"), ("detail", "Detail")]


def read_version():
    """Git commit of this folder (or 'zip <date>' without git) - shown in the page to see whether
    an update arrived."""
    try:
        head = (BASE_DIR / ".git" / "HEAD").read_text().strip()
        if head.startswith("ref:"):
            ref = BASE_DIR / ".git" / head[5:].strip()
            if ref.exists():
                return ref.read_text().strip()[:7]
            packed = (BASE_DIR / ".git" / "packed-refs").read_text()
            m = re.search(r"^([0-9a-f]{40}) " + re.escape(head[5:].strip()) + "$", packed, re.M)
            return m.group(1)[:7] if m else "?"
        return head[:7]
    except OSError:
        return "zip " + datetime.fromtimestamp((BASE_DIR / "server.py").stat().st_mtime).strftime("%Y-%m-%d %H:%M")


VERSION = read_version()
COMPUTE_LOCK = threading.Lock()     # the big lists: one at a time, without blocking other requests
STATE = {"loading": True, "lists_ready": False}


def mismatches():
    """Whole-site comparison - computed once per data state."""
    key = (json.dumps(STORE.meta, sort_keys=True), json.dumps(STORE.overrides, sort_keys=True))
    if MISMATCH_CACHE["key"] != key:
        with COMPUTE_LOCK:
            if MISMATCH_CACHE["key"] != key:
                MISMATCH_CACHE["rows"] = STORE.mismatches()
                MISMATCH_CACHE["key"] = key
    return MISMATCH_CACHE["rows"]


def ports_overview():
    """All ports (a few seconds) - computed once per data state."""
    key = (json.dumps(STORE.meta, sort_keys=True), json.dumps(STORE.overrides, sort_keys=True),
           json.dumps(STORE.future, sort_keys=True), CV["events_at"])
    if PORTS_CACHE["key"] != key:
        with COMPUTE_LOCK:
            if PORTS_CACHE["key"] != key:
                PORTS_CACHE["rows"] = STORE.ports_overview(CV["events"])
                PORTS_CACHE["key"] = key
    return PORTS_CACHE["rows"]


# ---------------------------------------------------------------- notes, Jira links, manual worklist
# One list in notes.json (git-ignored): comments and manual worklist items on a port or a
# device. {id, target: {type: port|device, switch, nums | device}, kind: comment|work,
#  text, jira, user, created, status (work: open|done), done_by, done_at}
NOTES_LOCK = threading.Lock()
JIRA_RE = re.compile(r"^[\w\s#:/.?=&%~\-,;]{0,300}$")


def load_notes():
    try:
        return json.loads(NOTES_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []


def save_notes(notes):
    tmp = NOTES_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(notes, indent=1, ensure_ascii=False), encoding="utf-8")
    tmp.replace(NOTES_FILE)


def note_target(t):
    """Validate a target from the browser -> canonical dict + key, or raise ValueError."""
    if not isinstance(t, dict):
        raise ValueError("bad target")
    if t.get("type") == "port":
        sw = str(t.get("switch", "")).strip()
        nums = t.get("nums")
        if not SWITCH_RE.match(sw) or not isinstance(nums, list) or not 1 <= len(nums) <= 3 or \
                not all(isinstance(n, int) and 0 <= n < 1000 for n in nums):
            raise ValueError("bad port")
        sw = STORE.pm_name(sw) or sw
        return {"type": "port", "switch": sw, "nums": nums}
    if t.get("type") == "device":
        d = str(t.get("device", "")).strip()
        if not d or len(d) > 120:
            raise ValueError("bad device")
        return {"type": "device", "device": d}
    raise ValueError("bad target")


def same_target(a, b):
    if a.get("type") != b.get("type"):
        return False
    if a["type"] == "port":
        return model.norm_name(a["switch"]) == model.norm_name(b["switch"]) and list(a["nums"]) == list(b["nums"])
    return model.norm_name(a["device"]) == model.norm_name(b["device"])


def known_users():
    """Names for the user pick list: everyone who wrote a note, plus the snagging list's 'who'."""
    names = {n.get("user", "").strip() for n in load_notes()} | {n.get("done_by", "").strip() for n in load_notes()}
    names |= {s.get("who", "").strip() for s in STORE.snags}
    return sorted((n for n in names if n and len(n) <= 60 and not n.startswith("#") and n != "0"), key=str.lower)


def manual_work_items():
    """Open manual worklist items in the worklist's row format."""
    out = []
    for n in load_notes():
        if n.get("kind") != "work" or n.get("status") == "done":
            continue
        t = n["target"]
        if t["type"] == "port":
            pm = t["switch"]
            out.append({"source": "manual", "switch": pm, "cv_switch": STORE.cv_name(pm) or "", "nums": t["nums"],
                        "intf": model.intf_name(t["nums"]), "device": "", "what": n.get("text", ""),
                        "destination": "", "status": "open (manual)", "since": n.get("created", "")[:16].replace("T", " "),
                        "who": n.get("user", ""), "jira": n.get("jira", ""), "nr": "", "note_id": n["id"], "future": ""})
        else:
            out.append({"source": "manual", "switch": "", "cv_switch": "", "nums": [], "intf": "", "device": t["device"],
                        "what": n.get("text", ""), "destination": t["device"], "status": "open (manual)",
                        "since": n.get("created", "")[:16].replace("T", " "), "who": n.get("user", ""),
                        "jira": n.get("jira", ""), "nr": "", "note_id": n["id"], "future": ""})
    return out


def worklist():
    return manual_work_items() + STORE.worklist(CV["events"])


# ---------------------------------------------------------------- help page (README.md)
REPO_URL = "https://github.com/t-schaefer/port-troubleshooter/blob/main/"


def md_to_html(md):
    """The little Markdown README.md uses: headings, paragraphs, lists, tables, `code`,
    **bold**, *italic*, [links]. Relative links point to the GitHub repo."""
    import html as H

    def inline(t):
        t = H.escape(t, quote=False)
        t = re.sub(r"`([^`]+)`", r"<code>\1</code>", t)
        t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
        t = re.sub(r"(?<![\w*])\*(?!\s)(.+?)\*(?!\w)", r"<i>\1</i>", t)

        def link(m):
            text, url = m.group(1), m.group(2)
            if not re.match(r"^(https?:|#|/)", url):
                url = REPO_URL + url.lstrip("./")
            return '<a class="lnk" href="%s" target="_blank" rel="noopener">%s</a>' % (H.escape(url), text)
        return re.sub(r"\[([^\]]+)\]\(([^)]+)\)", link, t)

    out, para, items, table = [], [], [], []

    def flush():
        if para:
            out.append("<p>%s</p>" % inline(" ".join(para)))
            para.clear()
        if items:
            out.append("<ul>%s</ul>" % "".join("<li>%s</li>" % inline(i) for i in items))
            items.clear()
        if table:
            rows = [[c.strip() for c in r.strip().strip("|").split("|")] for r in table
                    if not re.match(r"^\s*\|?\s*:?-{2,}", r)]
            head, body = rows[0], rows[1:]
            th = "".join("<th>%s</th>" % inline(c) for c in head) if any(head) else ""
            out.append("<table>%s<tbody>%s</tbody></table>" % ("<thead><tr>%s</tr></thead>" % th if th else "",
                       "".join("<tr>%s</tr>" % "".join("<td>%s</td>" % inline(c) for c in r) for r in body)))
            table.clear()

    for line in md.splitlines():
        if not line.strip():
            flush()
            continue
        m = re.match(r"^(#{1,3})\s+(.*)$", line)
        if m:
            flush()
            n = len(m.group(1)) + 1   # README '#' = page title -> h2 inside the panel
            out.append("<h%d>%s</h%d>" % (n, inline(m.group(2)), n))
        elif line.lstrip().startswith("|"):
            if para or items:
                flush()
            table.append(line)
        elif re.match(r"^\s*[-*]\s+", line):
            if para or table:
                flush()
            items.append(re.sub(r"^\s*[-*]\s+", "", line))
        elif items and line.startswith("  "):
            items[-1] += " " + line.strip()      # continued list item
        else:
            if items or table:
                flush()
            para.append(line.strip())
    flush()
    return "\n".join(out)


def warm_up():
    """Build the big lists (Ports, Mismatches) in the background after start / import / settings
    change, so a visitor never waits for them and other requests are not held up."""
    def work():
        try:
            t = time.time()
            ports_overview()
            mismatches()
            STATE["lists_ready"] = True
            log.info("lists ready (%.1fs)", time.time() - t)
        except Exception:
            log.exception("warm-up failed")
    threading.Thread(target=work, daemon=True).start()


def to_csv(rows, cols):
    """Excel-friendly CSV: UTF-8 with BOM, one header row."""
    f = io.StringIO()
    w = csv.writer(f)
    w.writerow([h for _, h in cols])
    for r in rows:
        w.writerow([" + ".join(map(str, r.get(k))) if isinstance(r.get(k), list) else r.get(k, "") for k, _ in cols])
    return ("\ufeff" + f.getvalue()).encode("utf-8")


# ---------------------------------------------------------------- config / files

def load_config():
    if not CONFIG_FILE.exists() and CONFIG_EXAMPLE.exists():
        CONFIG_FILE.write_text(CONFIG_EXAMPLE.read_text(encoding="utf-8"), encoding="utf-8")
        log.info("created config.json from config.example.json")
    # settings missing in config.json (added in a later version) come from config.example.json
    cfg = {}
    for f in (CONFIG_EXAMPLE, CONFIG_FILE):
        try:
            cfg.update(json.loads(f.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            pass
    return cfg


def load_serials():
    """CloudVision hostname -> device serial (for the per-port link into CloudVision). Defaults
    from serials.example.json, edits from the Switches tab in serials.json; inventory from the
    CloudVision API (when configured) fills the rest."""
    out = {}
    for f in (SERIALS_EXAMPLE, SERIALS_FILE):
        try:
            out.update({k.upper(): v for k, v in json.loads(f.read_text(encoding="utf-8")).items() if v})
        except (OSError, ValueError):
            pass
    for d in CV.get("devices") or []:
        if d.get("hostname") and d.get("id"):
            out.setdefault(d["hostname"].upper(), d["id"])
    return out


def cv_token():
    t = os.environ.get("CV_TOKEN", "").strip()
    if not t and TOKEN_FILE.exists():
        t = TOKEN_FILE.read_text(encoding="utf-8").strip()
    return t


def load_saved():
    """Load the last import of each kind from data/."""
    DATA_DIR.mkdir(exist_ok=True)
    try:
        meta = json.loads(META_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        meta = {}
    for key, m in meta.items():
        f = DATA_DIR / m.get("stored", "")
        if f.is_file():
            try:
                STORE.load(key.split(":")[0], m["file"], f.read_bytes(), when=m.get("loaded"))
                log.info("loaded %s: %s (%s)", kind_name(key), m["file"], STORE.meta[key]["count"])
            except Exception as e:  # a broken file must not stop the service
                log.warning("could not load %s from %s: %s", key, f, e)
    try:
        STORE.set_overrides(json.loads(SWITCHES_FILE.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    try:
        STORE.set_future(json.loads(FUTURE_FILE.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass


# ---------------------------------------------------------------- update from latest files
# Folders the service looks in for new exports (config.json "import_dirs", relative to this
# folder or absolute). Only files in these folders can be imported this way - the browser
# picks a file by an id from the last scan, never by path.
SCAN = {"files": {}, "kinds": {}}          # id -> path;  (path, mtime, size) -> kind


def import_dirs():
    dirs = load_config().get("import_dirs") or ["inbox"]
    out = []
    for d in dirs:
        p = Path(os.path.expanduser(str(d)))
        p = p if p.is_absolute() else BASE_DIR / p
        out.append(p)
    return out


def scan_sources(limit_per_dir=300):
    """Newest candidate files per input type in the import folders."""
    cands = []
    for d in import_dirs():
        if not d.is_dir():
            continue
        try:
            files = [f for f in d.iterdir() if f.is_file() and f.suffix.lower() in (".csv", ".xlsx", ".xlsm")
                     and not f.name.startswith(("~$", "."))]
        except OSError:
            continue
        files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        for f in files[:limit_per_dir]:
            st = f.stat()
            key = (str(f), st.st_mtime, st.st_size)
            if key not in SCAN["kinds"]:
                try:
                    SCAN["kinds"][key] = model.detect_kind(f.name, f.read_bytes() if st.st_size < MAX_UPLOAD else b"")
                except Exception:
                    SCAN["kinds"][key] = None
            kind = SCAN["kinds"][key]
            if kind == "cvports":
                try:
                    kind = model.meta_key(kind, f.name)
                except ValueError:
                    kind = None
            if kind:
                fid = hashlib.sha1(("%s|%s|%s" % key).encode()).hexdigest()[:16]
                SCAN["files"][fid] = str(f)
                cands.append({"id": fid, "kind": kind, "name": f.name, "folder": str(f.parent), "size": st.st_size,
                              "mtime": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"), "_key": key})
    out = {}
    try:
        meta = json.loads(META_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        meta = {}
    extra = sorted({c["kind"] for c in cands if c["kind"].startswith("cvports:")} | {k for k in meta if k.startswith("cvports:")})
    for kind in list(FIXED_KINDS) + extra:
        mine = sorted((c for c in cands if c["kind"] == kind), key=lambda c: c["mtime"], reverse=True)[:5]
        src = meta.get(kind, {}).get("source")
        for c in mine:
            c["loaded"] = bool(src) and src.get("path") == c["_key"][0] and src.get("mtime") == c["_key"][1] \
                and src.get("size") == c["_key"][2]
            c.pop("_key")
        newest = mine[0] if mine else None
        cur = STORE.meta.get(kind)
        if not newest or newest["loaded"]:
            newer = False
        elif src:                       # imported from a folder before: compare file dates
            newer = newest["mtime"] > src.get("mtime_iso", "")
        elif cur:                       # uploaded by hand: newer when the file is younger than that upload
            newer = newest["mtime"] > cur.get("loaded", "")
        else:
            newer = True
        out[kind] = {"current": cur, "candidates": mine, "newer": newer, "label": kind_name(kind)}
    return {"dirs": [str(d) + ("" if d.is_dir() else "  (missing)") for d in import_dirs()], "kinds": out}


def import_file(filename, data, source=None):
    kind = model.detect_kind(filename, data)
    if not kind:
        raise ValueError("Not recognised. Expected: Patchmanager cable export (CSV), Patchmanager equipment "
                         "export (CSV), IP list View (xlsx), Connectivity snagging list (xlsx), Neuron signal list "
                         "(xlsx) or a CloudVision Ethernet-interfaces export (CSV).")
    key = model.meta_key(kind, filename)          # 'cvports:<HOST>' - one per switch
    with LOCK:
        n = STORE.load(kind, filename, data)
        ext = filename.rsplit(".", 1)[1].lower()
        base = key.replace(":", "-")
        stored = "%s.%s" % (base, ext)
        for old in DATA_DIR.glob(base + ".*"):
            old.unlink()
        (DATA_DIR / stored).write_bytes(data)
        try:
            meta = json.loads(META_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            meta = {}
        meta[key] = {"file": filename, "stored": stored, "loaded": STORE.meta[key]["loaded"]}
        if source:
            meta[key]["source"] = source
        META_FILE.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    log.info("IMPORT %s: %s (%d rows)", kind_name(key), filename, n)
    return key, n


# ---------------------------------------------------------------- CloudVision poll

def cv_loop():
    cfg = load_config()
    last_inv = 0
    while True:
        token = cv_token()
        url = cfg.get("cv_url", "")
        CV["enabled"], CV["url"] = bool(token and url), url
        if CV["enabled"]:
            try:
                client = CloudVision(url, token, verify_tls=cfg.get("cv_verify_tls", True))
                if time.time() - last_inv > CV_INVENTORY_SEC or not CV["devices"]:
                    CV["devices"] = client.devices()
                    CV["devices_at"] = datetime.now().isoformat(timespec="seconds")
                    last_inv = time.time()
                id_to_host = {d["id"]: d["hostname"] for d in CV["devices"]}
                CV["events"] = client.events(since_hours=cfg.get("event_hours", 72), id_to_host=id_to_host)
                CV["events_at"] = datetime.now().isoformat(timespec="seconds")
                CV["raw_sample"] = json.dumps(CV["events"][:2], indent=1)[:2000]
                CV["error"] = ""
                warm_up()
            except CloudVisionError as e:
                CV["error"] = str(e)
                log.warning("CloudVision: %s", e)
            except Exception as e:
                CV["error"] = "unexpected: %s" % e
                log.exception("CloudVision poll failed")
        CV_WAKE.wait(CV_EVENTS_SEC)
        CV_WAKE.clear()
        cfg = load_config()


def cv_public():
    return {**{k: CV[k] for k in ("enabled", "url", "error", "devices_at", "events_at")},
            "devices": len(CV["devices"]), "events": len(CV["events"]), "token": bool(cv_token())}


# ---------------------------------------------------------------- HTTP

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass

    def _send(self, body, ctype="application/json", code=200, pre_gz=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        gz = (pre_gz is not None or len(body) > 2048) and "gzip" in self.headers.get("Accept-Encoding", "")
        if gz:
            body = pre_gz if pre_gz is not None else gzip.compress(body, 5)
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Server-Ms", "%.0f" % ((time.time() - getattr(self, "_t0", time.time())) * 1000))
        self.send_header("Access-Control-Expose-Headers", "X-Server-Ms")
        if gz:
            self.send_header("Content-Encoding", "gzip")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(json.dumps(obj, default=str), code=code)

    def _err(self, msg, code=400):
        self._json({"error": msg}, code)

    def _cached_json(self, cache, rows):
        """Big lists: JSON + gzip built once per data state, then sent as bytes."""
        if cache.get("body_key") is not cache.get("rows") or "body" not in cache:
            raw = json.dumps({"items": rows}, default=str).encode("utf-8")
            cache.update(body=raw, gz=gzip.compress(raw, 5), body_key=rows)
        self._send(cache["body"], pre_gz=cache["gz"])

    def do_GET(self):
        t = self._t0 = time.time()
        try:
            self._get()
        finally:
            dt = time.time() - t
            if dt > 0.5:
                log.warning("SLOW GET %s %.1fs", self.path[:120], dt)

    def do_POST(self):
        t = self._t0 = time.time()
        try:
            self._post()
        finally:
            dt = time.time() - t
            if dt > 2:
                log.warning("SLOW POST %s %.1fs", self.path[:120], dt)

    def _get(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path == "/":
                self._send(PAGE_FILE.read_bytes(), "text/html; charset=utf-8")
            elif u.path == "/api/help":
                try:
                    md = (BASE_DIR / "README.md").read_text(encoding="utf-8")
                except OSError:
                    md = "# Help\n\nREADME.md not found."
                self._json({"html": md_to_html(md)})
            elif u.path == "/api/status":
                with LOCK:
                    cfg = load_config()
                    self._json({"store": STORE.status(), "cv": cv_public(), "jira_url": cfg.get("jira_url", ""),
                                "loading": STATE["loading"], "lists_ready": STATE["lists_ready"],
                                "version": VERSION, "python": sys.version.split()[0], "platform": sys.platform,
                                "cv_url": cfg.get("cv_url", ""), "serials": load_serials()})
            elif u.path == "/api/search":
                with LOCK:
                    self._json({"results": STORE.search(q.get("q", "")[:200])})
            elif u.path == "/api/port":
                sw, nums = q.get("switch", ""), model.parse_intf(q.get("intf", ""))
                if not SWITCH_RE.match(sw) or not nums:
                    return self._err("bad switch / interface")
                with LOCK:
                    self._json(STORE.port_card(sw, nums, CV["events"], nav=True))
            elif u.path == "/api/device":
                name = q.get("name", "").strip()
                if not name or len(name) > 120:
                    return self._err("bad device name")
                with LOCK:
                    self._json(STORE.device_card(name, CV["events"]))
            elif u.path == "/api/cable":
                with LOCK:
                    if not STORE.cabling:
                        return self._err("no cable export imported")
                    idx = int(q.get("idx", "-1"))
                    if not 0 <= idx < len(STORE.cabling.cables):
                        return self._err("bad cable")
                    self._json(STORE.cable_card(idx, q.get("side")))
            elif u.path == "/api/mismatches":
                if STATE["loading"]:
                    return self._json({"items": [], "loading": True})
                rows = mismatches()
                self._cached_json(MISMATCH_CACHE, rows)
            elif u.path == "/api/ports":
                if STATE["loading"]:
                    return self._json({"items": [], "loading": True})
                rows = ports_overview()
                self._cached_json(PORTS_CACHE, rows)
            elif u.path == "/api/devices":
                with LOCK:
                    self._json({"items": STORE.devices_overview()})
            elif u.path == "/api/notes":
                try:
                    tgt = note_target(json.loads(q.get("target", "{}")))
                except ValueError as e:
                    return self._err(str(e))
                notes = [n for n in load_notes() if same_target(n["target"], tgt)]
                self._json({"notes": sorted(notes, key=lambda n: n.get("created", ""), reverse=True), "users": known_users()})
            elif u.path == "/api/sources":
                self._json(scan_sources())
            elif u.path == "/api/chain":
                name = q.get("device", "").strip()
                if not name or len(name) > 120:
                    return self._err("bad device name")
                with LOCK:
                    self._json(STORE.signal_chain(name, CV["events"]))
            elif u.path == "/api/future":
                with LOCK:
                    rows = STORE.ports_to_shut()
                    self._json({"devices": STORE.future_devices(), "ports": rows, "eos": model.eos_shutdown(rows)})
            elif u.path in ("/export/worklist.csv", "/export/mismatches.csv", "/export/shut.csv", "/export/shut-eos.txt",
                            "/export/chain.csv", "/export/ports.csv", "/export/devices.csv"):
                ctype, ext = "text/csv; charset=utf-8", "csv"
                with LOCK:
                    if u.path == "/export/ports.csv":
                        body, name = to_csv(ports_overview(), PORT_COLS), "ports"
                    elif u.path == "/export/devices.csv":
                        body, name = to_csv(STORE.devices_overview(), DEVICE_COLS), "devices"
                    elif u.path == "/export/chain.csv":
                        dev = q.get("device", "").strip()[:120]
                        ch = STORE.signal_chain(dev, CV["events"])
                        body = to_csv(ch["faults"], CHAIN_COLS)
                        name = "chain-" + re.sub(r"[^\w.\-]+", "_", ch["root"])
                    elif u.path == "/export/worklist.csv":
                        body, name = to_csv(worklist(), WORKLIST_COLS), "worklist"
                    elif u.path == "/export/mismatches.csv":
                        body, name = to_csv(mismatches(), MISMATCH_COLS), "mismatches"
                    elif u.path == "/export/shut.csv":
                        body, name = to_csv(STORE.ports_to_shut(), SHUT_COLS), "ports-to-shut"
                    else:
                        body, name = model.eos_shutdown(STORE.ports_to_shut()).encode("utf-8"), "shutdown-config"
                        ctype, ext = "text/plain; charset=utf-8", "txt"
                self.send_response(200)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Disposition", 'attachment; filename="port-troubleshooter-%s-%s.%s"' % (
                    name, datetime.now().strftime("%Y%m%d-%H%M"), ext))
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            elif u.path == "/api/worklist":
                with LOCK:
                    self._json({"items": worklist()})
            elif u.path == "/api/switches":
                with LOCK:
                    cvhosts = {d["hostname"].upper() for d in CV["devices"]}
                    sw = STORE.switches()
                    serials = load_serials()
                    for s in sw:
                        s["serial"] = serials.get((s["cv"] or "").upper(), "")
                        s["override"] = s["pm"] in STORE.overrides
                        s["in_cv"] = (s["cv"].upper() in cvhosts) if CV["devices"] and s["cv"] else None
                        s["dup"] = s["cv"].upper() in STORE.mapping_dups if s["cv"] else False
                    self._json({"switches": sw, "dups": STORE.mapping_dups})
            elif u.path == "/api/cv":
                self._json({**cv_public(), "sample": CV["raw_sample"], "recent": CV["events"][:200]})
            elif u.path == "/api/cv/locate":
                term = q.get("q", "").strip()
                if not LOCATE_RE.match(term):
                    return self._err("enter an IP, MAC or hostname")
                if not CV["enabled"]:
                    return self._err("CloudVision is not configured")
                cfg = load_config()
                try:
                    hits = CloudVision(cfg["cv_url"], cv_token(), verify_tls=cfg.get("cv_verify_tls", True)).locate(term)
                except CloudVisionError as e:
                    return self._err(str(e), 502)
                id_to_host = {d["id"]: d["hostname"] for d in CV["devices"]}
                with LOCK:
                    for h in hits:
                        h["device"] = id_to_host.get(h["device_id"], h["device_id"])
                        h["switch"] = STORE.pm_name(h["device"]) or h["device"]
                self._json({"hits": hits})
            else:
                self._err("not found", 404)
        except BrokenPipeError:
            pass
        except Exception as e:
            log.exception("GET %s failed", u.path)
            self._err("server error: %s" % e, 500)

    def _post(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            n = int(self.headers.get("Content-Length", "0"))
            if n > MAX_UPLOAD:
                return self._err("file too large (max %d MB)" % (MAX_UPLOAD // 1024 // 1024), 413)
            body = self.rfile.read(n)
            if u.path == "/api/import":
                name = os.path.basename(q.get("name", ""))
                if not FILENAME_RE.match(name):
                    return self._err("only .csv / .xlsx / .xlsm files")
                try:
                    kind, count = import_file(name, body)
                except ValueError as e:
                    return self._err(str(e))
                warm_up()
                self._json({"kind": kind, "count": count})
            elif u.path == "/api/import-latest":
                data = json.loads(body or b"{}")
                ids = data.get("ids") or []
                if not isinstance(ids, list) or not ids or len(ids) > 10:
                    return self._err("choose 1-10 files")
                done = []
                for fid in ids:
                    path = SCAN["files"].get(str(fid))
                    if not path:
                        return self._err("file list is out of date - reopen the menu")
                    f = Path(path)
                    if not any(f.parent == d for d in import_dirs()):
                        return self._err("file is not in an import folder")
                    st = f.stat()
                    try:
                        kind, count = import_file(f.name, f.read_bytes(), source={
                            "path": str(f), "mtime": st.st_mtime, "size": st.st_size,
                            "mtime_iso": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds")})
                    except ValueError as e:
                        return self._err("%s: %s" % (f.name, e))
                    done.append({"file": f.name, "kind": kind, "count": count})
                warm_up()
                self._json({"imported": done})
            elif u.path == "/api/switches":
                data = json.loads(body or b"{}")
                if not isinstance(data, dict) or len(data) > 500:
                    return self._err("expected {PM name: CloudVision hostname}")
                clean = {}
                for k, v in data.items():
                    k, v = str(k).strip(), str(v).strip()
                    if not (SWITCH_RE.match(k) and (not v or re.match(r"^[\w.\-]{1,64}$", v))):
                        return self._err("bad entry %r" % k)
                    if v:
                        clean[k] = v
                with LOCK:
                    SWITCHES_FILE.write_text(json.dumps(clean, indent=2, sort_keys=True), encoding="utf-8")
                    STORE.set_overrides(clean)
                log.info("switches.json saved (%d overrides)", len(clean))
                warm_up()
                self._json({"ok": True, "overrides": len(clean)})
            elif u.path == "/api/future":
                data = json.loads(body or b"{}")
                dev = str(data.get("device", "")).strip()
                note = str(data.get("note", "")).strip()[:200]
                if not dev or len(dev) > 120:
                    return self._err("bad device")
                with LOCK:
                    fut = dict(STORE.future)
                    if data.get("future"):
                        fut[dev] = {"note": note, "since": datetime.now().strftime("%Y-%m-%d")}
                    else:
                        fut.pop(dev, None)
                    FUTURE_FILE.write_text(json.dumps(fut, indent=2, sort_keys=True), encoding="utf-8")
                    STORE.set_future(fut)
                log.info("FUTURE %s %s %s", "set" if data.get("future") else "cleared", dev, note)
                warm_up()
                self._json({"ok": True})
            elif u.path == "/api/notes":
                data = json.loads(body or b"{}")
                action = data.get("action", "add")
                with NOTES_LOCK:
                    notes = load_notes()
                    if action == "add":
                        try:
                            tgt = note_target(data.get("target"))
                        except ValueError as e:
                            return self._err(str(e))
                        kind = "work" if data.get("kind") == "work" else "comment"
                        text = str(data.get("text", "")).strip()[:2000]
                        jira = str(data.get("jira", "")).strip()[:300]
                        user = str(data.get("user", "")).strip()[:60]
                        if not user:
                            return self._err("enter your name")
                        if not text and not jira:
                            return self._err("enter a comment or a Jira link")
                        if jira and not JIRA_RE.match(jira):
                            return self._err("Jira: a key (ABC-123) or a link")
                        n = {"id": hashlib.sha1(("%s%s%s" % (time.time(), user, text)).encode()).hexdigest()[:12],
                             "target": tgt, "kind": kind, "text": text, "jira": jira, "user": user,
                             "created": datetime.now().isoformat(timespec="seconds")}
                        if kind == "work":
                            n["status"] = "open"
                        notes.append(n)
                        log.info("NOTE %s by %s on %s: %s %s", kind, user, json.dumps(tgt), text[:80], jira)
                    elif action in ("done", "reopen", "delete"):
                        nid = str(data.get("id", ""))
                        user = str(data.get("user", "")).strip()[:60]
                        hit = next((n for n in notes if n["id"] == nid), None)
                        if not hit:
                            return self._err("note not found")
                        if action == "delete":
                            notes.remove(hit)
                        elif action == "done":
                            if not user:
                                return self._err("enter your name")
                            hit.update(status="done", done_by=user, done_at=datetime.now().isoformat(timespec="seconds"))
                        else:
                            hit.update(status="open", done_by="", done_at="")
                        log.info("NOTE %s %s by %s", action, nid, user or "-")
                    else:
                        return self._err("unknown action")
                    save_notes(notes)
                self._json({"ok": True})
            elif u.path == "/api/serials":
                data = json.loads(body or b"{}")
                if not isinstance(data, dict) or len(data) > 500:
                    return self._err("expected {CloudVision hostname: serial}")
                clean = {}
                for k, v in data.items():
                    k, v = str(k).strip().upper(), str(v).strip()
                    if not re.match(r"^[\w.\-]{1,64}$", k) or (v and not re.match(r"^[A-Za-z0-9\-]{4,40}$", v)):
                        return self._err("bad serial for %s" % k)
                    if v:
                        clean[k] = v
                SERIALS_FILE.write_text(json.dumps(clean, indent=2, sort_keys=True), encoding="utf-8")
                log.info("serials.json saved (%d)", len(clean))
                self._json({"ok": True})
            elif u.path == "/api/cv/refresh":
                CV_WAKE.set()
                self._json({"ok": True})
            else:
                self._err("not found", 404)
        except BrokenPipeError:
            pass
        except Exception as e:
            log.exception("POST %s failed", u.path)
            self._err("server error: %s" % e, 500)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8092)
    args = ap.parse_args()
    load_config()
    for d in import_dirs():
        if d == BASE_DIR / "inbox":
            d.mkdir(exist_ok=True)
    ThreadingHTTPServer.request_queue_size = 128
    ThreadingHTTPServer.daemon_threads = True
    srv = ThreadingHTTPServer(("0.0.0.0", args.port), Handler)

    def start():   # listen at once - the page shows "loading" until the data is in
        t = time.time()
        with LOCK:
            load_saved()
        STATE["loading"] = False
        log.info("data loaded (%.1fs)", time.time() - t)
        threading.Thread(target=cv_loop, daemon=True).start()
        warm_up()
    threading.Thread(target=start, daemon=True).start()
    log.info("Port Troubleshooter on http://0.0.0.0:%d/  (Ctrl+C stops it)", args.port)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        log.info("stopped")


if __name__ == "__main__":
    main()
