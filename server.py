"""
Port Troubleshooter - one page to troubleshoot a switch port: documented cable path
(Patchmanager), what the port plan expects there, the snagging-list history and the
live CloudVision events, side by side.

- Standard library only (Python 3.8+), ThreadingHTTPServer, default port 8102.
- Inputs are imported in the page (or dropped into data/): Patchmanager cable export (CSV),
  Patchmanager all-equipment export (CSV), IP list "View" (xlsx, port plan), Connectivity
  snagging list (xlsx). The type is detected from the content. The last import of each type
  is kept in data/ (git-ignored) and loaded again at start.
- CloudVision (read-only, optional): token in env CV_TOKEN or cv_token.txt (git-ignored),
  URL in config.json. Inventory every 10 min, events every 2 min, in the background -
  viewers never cause CloudVision requests, except the explicit "Locate" button.
- PM <-> CloudVision switch names: from the snagging list's name sheet, overrides in
  switches.json (git-ignored, editable in the page).

    python server.py               # http://<this pc>:8102/
    python server.py --port 8110
"""
import argparse
import csv
import gzip
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
MISMATCH_CACHE = {"key": None, "rows": []}

WORKLIST_COLS = [("source", "Source"), ("switch", "Switch (Patchmanager)"), ("cv_switch", "Switch (CloudVision)"),
                 ("intf", "Interface"), ("what", "Problem"), ("destination", "Destination"), ("status", "Status"),
                 ("since", "Since"), ("who", "Who"), ("jira", "Jira"), ("nr", "Snag Nr")]
SHUT_COLS = [("switch", "Switch (Patchmanager)"), ("cv_switch", "Switch (CloudVision)"), ("intf", "Interface"),
             ("device", "Future device"), ("device_port", "Device port"), ("via", "Found via"), ("source", "Marked by"),
             ("note", "Note")]
CHAIN_COLS = [("severity", "Severity"), ("fault", "Fault"), ("depth", "Step"), ("from", "From device"),
              ("from_port", "From port"), ("to", "To device"), ("to_port", "To port"), ("cables", "Cables")]
MISMATCH_COLS = [("severity", "Severity"), ("kind", "Mismatch"), ("switch", "Switch (Patchmanager)"),
                 ("cv_switch", "Switch (CloudVision)"), ("intf", "Interface"), ("source", "Source"), ("says", "Source says"),
                 ("reference", "Compared with"), ("ref_says", "That says"), ("detail", "Detail")]


def mismatches():
    """Whole-site comparison takes a few seconds - computed once per data state."""
    key = (json.dumps(STORE.meta, sort_keys=True), json.dumps(STORE.overrides, sort_keys=True))
    if MISMATCH_CACHE["key"] != key:
        MISMATCH_CACHE["rows"] = STORE.mismatches()
        MISMATCH_CACHE["key"] = key
    return MISMATCH_CACHE["rows"]


def to_csv(rows, cols):
    """Excel-friendly CSV: UTF-8 with BOM, one header row."""
    f = io.StringIO()
    w = csv.writer(f)
    w.writerow([h for _, h in cols])
    for r in rows:
        w.writerow([r.get(k, "") for k, _ in cols])
    return ("\ufeff" + f.getvalue()).encode("utf-8")


# ---------------------------------------------------------------- config / files

def load_config():
    if not CONFIG_FILE.exists() and CONFIG_EXAMPLE.exists():
        CONFIG_FILE.write_text(CONFIG_EXAMPLE.read_text(encoding="utf-8"), encoding="utf-8")
        log.info("created config.json from config.example.json")
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


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
    for kind, m in meta.items():
        f = DATA_DIR / m.get("stored", "")
        if f.is_file():
            try:
                STORE.load(kind, m["file"], f.read_bytes(), when=m.get("loaded"))
                log.info("loaded %s: %s (%s)", kind, m["file"], STORE.meta[kind]["count"])
            except Exception as e:  # a broken file must not stop the service
                log.warning("could not load %s from %s: %s", kind, f, e)
    try:
        STORE.set_overrides(json.loads(SWITCHES_FILE.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    try:
        STORE.set_future(json.loads(FUTURE_FILE.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass


def import_file(filename, data):
    kind = model.detect_kind(filename, data)
    if not kind:
        raise ValueError("Not recognised. Expected: Patchmanager cable export (CSV), Patchmanager equipment "
                         "export (CSV), IP list View (xlsx) or Connectivity snagging list (xlsx).")
    with LOCK:
        n = STORE.load(kind, filename, data)
        ext = filename.rsplit(".", 1)[1].lower()
        stored = "%s.%s" % (kind, ext)
        for old in DATA_DIR.glob(kind + ".*"):
            old.unlink()
        (DATA_DIR / stored).write_bytes(data)
        try:
            meta = json.loads(META_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            meta = {}
        meta[kind] = {"file": filename, "stored": stored, "loaded": STORE.meta[kind]["loaded"]}
        META_FILE.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    log.info("IMPORT %s: %s (%d rows)", kind, filename, n)
    return kind, n


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

    def _send(self, body, ctype="application/json", code=200):
        if isinstance(body, str):
            body = body.encode("utf-8")
        gz = len(body) > 2048 and "gzip" in self.headers.get("Accept-Encoding", "")
        if gz:
            body = gzip.compress(body, 5)
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if gz:
            self.send_header("Content-Encoding", "gzip")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(json.dumps(obj, default=str), code=code)

    def _err(self, msg, code=400):
        self._json({"error": msg}, code)

    def do_GET(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path == "/":
                self._send(PAGE_FILE.read_bytes(), "text/html; charset=utf-8")
            elif u.path == "/api/status":
                with LOCK:
                    self._json({"store": STORE.status(), "cv": cv_public()})
            elif u.path == "/api/search":
                with LOCK:
                    self._json({"results": STORE.search(q.get("q", "")[:200])})
            elif u.path == "/api/port":
                sw, nums = q.get("switch", ""), model.parse_intf(q.get("intf", ""))
                if not SWITCH_RE.match(sw) or not nums:
                    return self._err("bad switch / interface")
                with LOCK:
                    self._json(STORE.port_card(sw, nums, CV["events"]))
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
                with LOCK:
                    self._json({"items": mismatches()})
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
                            "/export/chain.csv"):
                ctype, ext = "text/csv; charset=utf-8", "csv"
                with LOCK:
                    if u.path == "/export/chain.csv":
                        dev = q.get("device", "").strip()[:120]
                        ch = STORE.signal_chain(dev, CV["events"])
                        body = to_csv(ch["faults"], CHAIN_COLS)
                        name = "chain-" + re.sub(r"[^\w.\-]+", "_", ch["root"])
                    elif u.path == "/export/worklist.csv":
                        body, name = to_csv(STORE.worklist(CV["events"]), WORKLIST_COLS), "worklist"
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
                    self._json({"items": STORE.worklist(CV["events"])})
            elif u.path == "/api/switches":
                with LOCK:
                    cvhosts = {d["hostname"].upper() for d in CV["devices"]}
                    sw = STORE.switches()
                    for s in sw:
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

    def do_POST(self):
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
                self._json({"kind": kind, "count": count})
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
    ap.add_argument("--port", type=int, default=8102)
    args = ap.parse_args()
    load_config()
    load_saved()
    threading.Thread(target=cv_loop, daemon=True).start()
    ThreadingHTTPServer.request_queue_size = 128
    ThreadingHTTPServer.daemon_threads = True
    srv = ThreadingHTTPServer(("0.0.0.0", args.port), Handler)
    log.info("Port Troubleshooter on http://0.0.0.0:%d/  (Ctrl+C stops it)", args.port)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        log.info("stopped")


if __name__ == "__main__":
    main()
