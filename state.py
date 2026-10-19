"""
Everything people enter in the page, and the history - one SQLite file (state.db, git-ignored).
Standard library only (sqlite3).

  settings  small dictionaries by name: switches (PM name -> CloudVision hostname overrides),
            serials (hostname -> CloudVision serial), future (device -> note), ignores
            (mismatch key -> reason / user / time)
  notes     comments, Jira links and manual worklist items
  imports   every import: when, which input, which file, how many rows
  builds    every new evaluation state: mismatch counts per bucket
  build_keys the mismatch keys of the last builds (to show "new" / "resolved")

The old JSON files (notes.json, ignores.json, future.json, serials.json, switches.json) are
taken over on the first start and renamed to *.json.migrated (kept, not deleted).
"""
import json
import sqlite3
import threading
from datetime import datetime
from pathlib import Path

KEEP_BUILDS_WITH_KEYS = 30


class StateDB:
    def __init__(self, path):
        self.path = Path(path)
        self.lock = threading.Lock()
        self.db = sqlite3.connect(str(self.path), check_same_thread=False)
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.executescript("""
            CREATE TABLE IF NOT EXISTS settings (name TEXT PRIMARY KEY, json TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS notes (id TEXT PRIMARY KEY, json TEXT NOT NULL, created TEXT);
            CREATE TABLE IF NOT EXISTS imports (id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT, kind TEXT,
                                                file TEXT, rows INTEGER, how TEXT);
            CREATE TABLE IF NOT EXISTS builds (id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT, reason TEXT,
                                               real INTEGER, incomplete INTEGER, naming INTEGER, total INTEGER);
            CREATE TABLE IF NOT EXISTS build_keys (build_id INTEGER, key TEXT, bucket TEXT, kind TEXT,
                                                   switch TEXT, intf TEXT, says TEXT);
            CREATE INDEX IF NOT EXISTS bk_build ON build_keys(build_id);
        """)
        self.db.commit()

    # -- small dictionaries ---------------------------------------------------------------
    def get(self, name, default=None):
        with self.lock:
            row = self.db.execute("SELECT json FROM settings WHERE name=?", (name,)).fetchone()
        return json.loads(row[0]) if row else ({} if default is None else default)

    def put(self, name, value):
        with self.lock:
            self.db.execute("INSERT OR REPLACE INTO settings(name, json) VALUES(?, ?)",
                            (name, json.dumps(value, ensure_ascii=False)))
            self.db.commit()

    # -- notes --------------------------------------------------------------------------------
    def notes(self):
        with self.lock:
            return [json.loads(r[0]) for r in self.db.execute("SELECT json FROM notes ORDER BY created")]

    def save_note(self, note):
        with self.lock:
            self.db.execute("INSERT OR REPLACE INTO notes(id, json, created) VALUES(?, ?, ?)",
                            (note["id"], json.dumps(note, ensure_ascii=False), note.get("created", "")))
            self.db.commit()

    def delete_note(self, nid):
        with self.lock:
            self.db.execute("DELETE FROM notes WHERE id=?", (nid,))
            self.db.commit()

    # -- history ------------------------------------------------------------------------------
    def log_import(self, kind, file, rows, how):
        with self.lock:
            self.db.execute("INSERT INTO imports(at, kind, file, rows, how) VALUES(?, ?, ?, ?, ?)",
                            (datetime.now().isoformat(timespec="seconds"), kind, file, rows, how))
            self.db.commit()

    def log_build(self, mismatches, ignored_keys, reason):
        """Store a new evaluation state - only when something changed against the last one.
        Returns True when a new build was written."""
        live = [m for m in mismatches if m.get("key") not in ignored_keys]
        keys = {m["key"] for m in live}
        counts = {b: sum(1 for m in live if m["bucket"] == b) for b in ("real", "incomplete", "naming")}
        with self.lock:
            last = self.db.execute("SELECT id, real, incomplete, naming FROM builds ORDER BY id DESC LIMIT 1").fetchone()
            if last:
                old = {r[0] for r in self.db.execute("SELECT key FROM build_keys WHERE build_id=?", (last[0],))}
                if old == keys and tuple(last[1:]) == (counts["real"], counts["incomplete"], counts["naming"]):
                    return False
            cur = self.db.execute("INSERT INTO builds(at, reason, real, incomplete, naming, total) VALUES(?, ?, ?, ?, ?, ?)",
                                  (datetime.now().isoformat(timespec="seconds"), reason, counts["real"],
                                   counts["incomplete"], counts["naming"], len(live)))
            bid = cur.lastrowid
            self.db.executemany("INSERT INTO build_keys(build_id, key, bucket, kind, switch, intf, says) VALUES(?, ?, ?, ?, ?, ?, ?)",
                                [(bid, m["key"], m["bucket"], m["kind"], m["switch"], str(m.get("intf", "")),
                                  str(m.get("says", ""))[:200]) for m in live])
            old_ids = [r[0] for r in self.db.execute("SELECT id FROM builds ORDER BY id DESC LIMIT -1 OFFSET ?",
                                                     (KEEP_BUILDS_WITH_KEYS,))]
            if old_ids:
                self.db.executemany("DELETE FROM build_keys WHERE build_id=?", [(i,) for i in old_ids])
            self.db.commit()
        return True

    def history(self, n_builds=90, n_imports=50):
        """Builds (counts over time), imports, and what changed between the last two builds."""
        with self.lock:
            builds = [dict(zip(("id", "at", "reason", "real", "incomplete", "naming", "total"), r)) for r in
                      self.db.execute("SELECT id, at, reason, real, incomplete, naming, total FROM builds ORDER BY id DESC LIMIT ?",
                                      (n_builds,))][::-1]
            imports = [dict(zip(("at", "kind", "file", "rows", "how"), r)) for r in
                       self.db.execute("SELECT at, kind, file, rows, how FROM imports ORDER BY id DESC LIMIT ?", (n_imports,))]
            diff = {"new": [], "resolved": [], "since": ""}
            if len(builds) >= 2:
                a, b = builds[-2]["id"], builds[-1]["id"]
                cols = "key, bucket, kind, switch, intf, says"
                old = {r[0]: r for r in self.db.execute("SELECT %s FROM build_keys WHERE build_id=?" % cols, (a,))}
                new = {r[0]: r for r in self.db.execute("SELECT %s FROM build_keys WHERE build_id=?" % cols, (b,))}
                as_dict = lambda r: dict(zip(("key", "bucket", "kind", "switch", "intf", "says"), r))
                diff = {"new": [as_dict(new[k]) for k in new.keys() - old.keys()],
                        "resolved": [as_dict(old[k]) for k in old.keys() - new.keys()],
                        "since": builds[-2]["at"]}
        return {"builds": builds, "imports": imports, "diff": diff}

    # -- take over the old JSON files -------------------------------------------------------------
    def migrate_json(self, base, log=None):
        base = Path(base)
        for name, fname in (("switches", "switches.json"), ("serials", "serials.json"),
                            ("future", "future.json"), ("ignores", "ignores.json")):
            f = base / fname
            if f.exists():
                try:
                    data = json.loads(f.read_text(encoding="utf-8"))
                    current = self.get(name)
                    current.update(data)
                    self.put(name, current)
                    f.rename(f.with_name(fname + ".migrated"))
                    if log:
                        log.info("state.db: took over %s (%d entries)", fname, len(data))
                except (OSError, ValueError) as e:
                    if log:
                        log.warning("state.db: could not take over %s: %s", fname, e)
        f = base / "notes.json"
        if f.exists():
            try:
                for n in json.loads(f.read_text(encoding="utf-8")):
                    self.save_note(n)
                f.rename(f.with_name("notes.json.migrated"))
                if log:
                    log.info("state.db: took over notes.json")
            except (OSError, ValueError) as e:
                if log:
                    log.warning("state.db: could not take over notes.json: %s", e)
