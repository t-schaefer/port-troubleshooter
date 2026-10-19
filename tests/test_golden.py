"""Regression test against the REAL data of this installation (data/, never in git).

Only numbers are checked in (tests/golden.json): how many mismatches per category, how
many ports per status, how many devices per type ... - no names, nothing project-specific.
A change in the rules that shifts these numbers fails the test and shows the difference,
so it is a decision ("yes, that is the intended effect") instead of a surprise.

    python -m unittest tests.test_golden          # check (skipped when data/ is empty)
    python tests/test_golden.py --update          # accept the current numbers as expected

golden.json is per data set: after importing new exports, run --update once.
"""
import collections
import hashlib
import json
import os
import sys
import unittest

BASE = os.path.join(os.path.dirname(__file__), "..")
sys.path.insert(0, BASE)
import model  # noqa: E402

DATA = os.path.join(BASE, "data")
GOLDEN = os.path.join(os.path.dirname(__file__), "golden.json")


def load_store():
    meta = json.load(open(os.path.join(DATA, "meta.json"), encoding="utf-8"))
    s = model.Store()
    for key, m in meta.items():
        s.load(key.split(":")[0], m["file"], open(os.path.join(DATA, m["stored"]), "rb").read())
    return s, meta


def fingerprint(s, meta):
    """The numbers that describe the result - and some invariants that must always hold."""
    mis = s.mismatches()
    ports = s.ports_overview()
    devs = s.devices_overview()
    ncheck = collections.Counter()
    for r in s.neuron:
        ch = s.neuron_check(r)
        ncheck["ok" if not ch else "issue"] += 1
    return {
        # which data set (hash only - file names can be project specific, the repo is public)
        "inputs": hashlib.sha1(json.dumps({k.split(":")[0]: m.get("file") for k, m in sorted(meta.items())},
                                          sort_keys=True).encode()).hexdigest()[:12],
        "counts": {
            "cables": len(s.cabling.cables) if s.cabling else 0,
            "ip_list_rows": len(s.plan),
            "snags": len(s.snags),
            "neuron_rows": len(s.neuron),
        },
        "mismatch_bucket": dict(sorted(collections.Counter(m["bucket"] for m in mis).items())),
        "mismatch_kind": dict(sorted(collections.Counter("%s | %s" % (m["bucket"], m["kind"]) for m in mis).items())),
        "port_status": dict(sorted(collections.Counter(p["status"] for p in ports).items())),
        "device_kind": dict(sorted(collections.Counter(d["kind"] for d in devs).items())),
        "neuron_check": dict(sorted(ncheck.items())),
    }


def invariants(s):
    """Things that must hold for ANY data set."""
    problems = []
    # two different tops must never be merged into one identity key
    seen = {}
    if s.cabling:
        for c in s.cabling.cables:
            for e in (c.a, c.b):
                if e.top:
                    k = model.norm_name(e.top)
                    other = seen.setdefault(k, e.top)
                    if model._compact(other) != model._compact(e.top) and other != e.top:
                        problems.append("identity collision: %r / %r" % (other, e.top))
    # every mismatch is explained
    for m in s.mismatches():
        if not m.get("meaning") or m.get("bucket") not in ("real", "incomplete", "naming"):
            problems.append("unexplained mismatch: %s %s %s" % (m["switch"], m["intf"], m["kind"]))
            break
    return problems[:10]


def has_data():
    return os.path.exists(os.path.join(DATA, "meta.json"))


@unittest.skipUnless(has_data(), "no data/ in this installation - golden test skipped")
class GoldenTest(unittest.TestCase):
    maxDiff = None

    @classmethod
    def setUpClass(cls):
        cls.store, cls.meta = load_store()

    def test_invariants(self):
        self.assertEqual(invariants(self.store), [])

    def test_numbers_unchanged(self):
        if not os.path.exists(GOLDEN):
            self.skipTest("no tests/golden.json yet - run: python tests/test_golden.py --update")
        golden = json.load(open(GOLDEN, encoding="utf-8"))
        now = fingerprint(self.store, self.meta)
        if golden.get("inputs") != now["inputs"]:
            self.skipTest("golden.json was made for other input files - run --update after checking the result")
        self.assertEqual(golden, now, "numbers changed - if intended: python tests/test_golden.py --update")


if __name__ == "__main__":
    if "--update" in sys.argv:
        s, meta = load_store()
        fp = fingerprint(s, meta)
        json.dump(fp, open(GOLDEN, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
        print("golden.json updated:", json.dumps(fp["mismatch_bucket"]), "| invariants:", invariants(s) or "ok")
    else:
        unittest.main()
