"""
Port troubleshooter - data model. Standard library only, no network.

Inputs (all optional, each replaces its predecessor when imported again):
  cables     Patchmanager cable export (CSV)          -> documented cabling, path tracing
  equipment  Patchmanager "all equipment" export (CSV) -> template, rack/U, defect/reserved ports
  portplan   IP list ("View" xlsx: Switch | Port | Device | VLAN | Cable No blocks)
  snags      Connectivity snagging list (xlsx: "Master" sheet + PM <-> CloudVision name sheet)

Names: Patchmanager calls a switch "DCB.ESW A01", CloudVision "DCBESWCPREDG001". The
mapping comes from the snagging list (sheet with columns "PM" / "arista cloud") and can be
overridden in switches.json.

Interfaces: CloudVision "Ethernet49" / "Ethernet5/2" / "Ethernet3/1/1" ->
  Port(slot, port, lane): fixed switch Ethernet49 -> (None, 49, None)
                          breakout     Ethernet5/2 -> (None, 5, 2)
                          modular      Ethernet3/1/1 -> (3, 1, 1)  (Patchmanager: line card LC03, port 1)
Whether a two-number name is breakout (port/lane) or modular (slot/port) depends on the
switch: a switch whose Patchmanager ports sit under LCxx line cards is modular.
"""
import csv
import functools
import hashlib
import io
import re
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

from xlsx import read_xlsx

# ---------------------------------------------------------------- helpers

def norm_port(p):
    """'01' -> '1', ' Eth ' -> 'ETH'; used to compare Patchmanager / plan port labels."""
    p = str(p or "").strip().upper()
    return str(int(p)) if p.isdigit() else p


_NORM_RE = re.compile(r"[\s.\-_]")
_ZERO_RE = re.compile(r"(?<!\d)0+(?=\d)")
_RUN_RE = re.compile(r"[A-Z]+|\d+|[^\sA-Z\d.\-_]+")


@functools.lru_cache(maxsize=200000)
def _compact(s):
    return _NORM_RE.sub("", s).upper()


@functools.lru_cache(maxsize=200000)
def _norm(s):
    # letter / number blocks, separators between them ignored
    return "~".join(_RUN_RE.findall(s.upper()))


def compact(s):
    """Everything glued together: 'DCB.VS 02' -> 'DCBVS02' (for partial search only)."""
    return _compact(s if isinstance(s, str) else str(s or ""))


def norm_name(s):
    """Identity of a name: case, spaces, dots, dashes do not matter, but the letter / number
    blocks do: 'DCB.ESW A01' == 'dcb esw a01' == 'DCB-ESW-A01', 'DCR.SGXR 103' == 'DCR.SGXR103',
    yet 'DCB.ESW B01' (access switch B01) != 'DCB.ESWB 01' (blue leaf 01). Cached."""
    return _norm(s if isinstance(s, str) else str(s or ""))


@functools.lru_cache(maxsize=200000)
def _ndig(s):
    return _ZERO_RE.sub("", _norm(s))


def norm_digits(s):
    """norm_name without leading zeros in numbers: 'DCB.VS 02' == 'DCB.VS 2' == 'dcbvs2'."""
    return _ndig(s if isinstance(s, str) else str(s or ""))


def same_device(a, b):
    return norm_digits(a) == norm_digits(b)


def same_switch_port(a, b):
    """{switch, nums} == {switch, nums}; plan '1/25' on a fixed switch == port 25."""
    if not a or not b or not same_device(a["switch"], b["switch"]):
        return False
    na, nb = list(a["nums"]), list(b["nums"])
    strip = lambda n: n[1:] if len(n) == 2 and n[0] == 1 else n
    return na == nb or strip(na) == strip(nb) or na[:1] == nb[:1] and (len(na) == 1 or len(nb) == 1)


PORT_SYNONYMS = [(r"MGNT|MANAGEMENT", "MGMT"), (r"ETHERNET|ENET|LAN|LOM|NIC|RJ45|EXTLAN", "ETH")]


def port_relation(a, b):
    """'same' | 'naming' (ETH = ETH.1, MGNT = MGMT, PCLAN1 = PC LAN 1 ...) | 'different'."""
    if norm_name(a) == norm_name(b) or not a or not b:
        return "same"
    return "naming" if same_port(a, b) else "different"


def likely_typo(a, b):
    """Same numbers in the same order, letters differ by one edit (insert / delete / replace /
    swap two neighbours) and are long enough: 'OC.PRIN 01' ~ 'OC.PRINT 01', 'SPDE 1' ~ 'SDPE 1'."""
    # a different room (DCB / DCR, CS1 / CS2 ...) is never a typo
    if "." in a and "." in b and norm_name(a.split(".", 1)[0]) != norm_name(b.split(".", 1)[0]):
        return False
    na, nb = norm_digits(a), norm_digits(b)
    if re.findall(r"\d+", na) != re.findall(r"\d+", nb):
        return False
    la, lb = re.sub(r"[\d~]", "", na), re.sub(r"[\d~]", "", nb)
    if min(len(la), len(lb)) < 4 or la == lb:
        return False
    if abs(len(la) - len(lb)) > 1:
        return False
    if len(la) == len(lb):
        diff = [i for i in range(len(la)) if la[i] != lb[i]]
        return len(diff) == 1 or (len(diff) == 2 and diff[1] == diff[0] + 1 and la[diff[0]] == lb[diff[1]] and la[diff[1]] == lb[diff[0]])
    short, long_ = (la, lb) if len(la) < len(lb) else (lb, la)
    return any(long_[:i] + long_[i + 1:] == short for i in range(len(long_)))


def same_port(a, b):
    """Same physical port despite naming: 'PCLAN1' == 'PC LAN 1', 'ETH.1' == 'Ethernet' ==
    'ENET 1', 'MGNT' == 'MGMT', 'ETH.3' == '03', '10G.1' == 'Expansion / 10G.1 / 10G.1'.
    Different numbers stay different: 'ETH.1' != 'ETH.2'."""
    def keys(p):
        parts = [x.strip() for x in str(p).split("/") if x.strip()] + [str(p)]
        out = set()
        for x in parts:
            k = _ZERO_RE.sub("", compact(x))
            for pat, rep in PORT_SYNONYMS:
                k = re.sub(pat, rep, k)
            out.add(k)
            if k.endswith("1") and not k[:-1].endswith(tuple("0123456789")):
                out.add(k[:-1])          # 'ETH1' also as 'ETH' (a single port)
            num = re.sub(r"\D", "", k)
            if num and re.fullmatch(r"(ETH|PORT|P)?\d+", k):
                out.add("#" + num)       # 'ETH3' / '03' / 'P3' -> '#3'
        out.add(_ZERO_RE.sub("", compact(str(p).replace("/", ""))))  # 'S5 / D1' == 'S5. D1'
        return out

    def last_num(p):
        m = re.search(r"(\d+)\D*$", str(p))
        return int(m.group(1)) if m else None
    if keys(a) & keys(b):
        return True
    # module paths written differently ('10G.2' / '10G / 2 / 2', '1.7' / '07 / 07'): the
    # trailing number decides, as long as both have a module/lane structure
    la, lb = last_num(a), last_num(b)
    structured = lambda p: bool(re.search(r"[./]", str(p)))
    return la is not None and la == lb and structured(a) and structured(b) and \
        not (re.sub(r"[\d\s./]", "", str(a)).upper() in ("ETH", "LAN") and re.sub(r"[\d\s./]", "", str(b)).upper() in ("ETH", "LAN"))


def port_sort(end):
    """Sort key for ports: numbers numerically, then names."""
    p = norm_port(end.port)
    return (end.parent or "", end.label if end.label != end.top else "", 0 if p.isdigit() else 1, int(p) if p.isdigit() else 0, p)


INTF_RE = re.compile(r"^\s*(?:et(?:hernet)?)?\s*(\d+)(?:/(\d+))?(?:/(\d+))?\s*$", re.I)


def parse_intf(s):
    """'Ethernet3/1/1' -> (3, 1, 1); 'Et49' -> (49,); '5/2' -> (5, 2). None when not an interface."""
    m = INTF_RE.match(str(s or ""))
    if not m:
        return None
    return tuple(int(x) for x in m.groups() if x is not None)


def intf_name(nums):
    return "Ethernet" + "/".join(str(n) for n in nums)


def excel_date(v):
    """Excel serial (46154) or text -> 'YYYY-MM-DD' / the text itself."""
    s = str(v or "").strip()
    if re.fullmatch(r"\d{5}(\.\d+)?", s):
        return (date(1899, 12, 30) + timedelta(days=int(float(s)))).isoformat()
    # CloudVision copy-paste: 'Aug 21, 2026 13:04:24 GMT+1' -> '2026-08-21 13:04' (sortable)
    m = re.match(r"^([A-Z][a-z]{2}) (\d{1,2}), (\d{4})(?: (\d\d):(\d\d))?", s)
    if m:
        try:
            d = datetime.strptime("%s %s %s" % m.group(1, 2, 3), "%b %d %Y").date().isoformat()
            return d + (" %s:%s" % m.group(4, 5) if m.group(4) else "")
        except ValueError:
            pass
    return s


def parse_attrs(s):
    """'Installed=false&Stamp=x 09<COLON>44' -> {'Installed': 'false', 'Stamp': 'x 09:44'}"""
    out = {}
    for part in str(s or "").split("&"):
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip()] = v.replace("<COLON>", ":").strip()
    return out


def read_csv_text(data):
    if isinstance(data, (bytes, bytearray)):
        for enc in ("utf-8-sig", "cp1252"):
            try:
                data = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
    # short rows give None values, long rows a None key - both become harmless
    return [{k: (v or "") for k, v in r.items() if k is not None} for r in csv.DictReader(io.StringIO(data))]


# Patch panels, outlets, cassettes: a path that ends on one of these is not finished.
PASSIVE_RE = re.compile(r"leviton|patch ?panel|\bpp-|gop|splice|rosenberger|d-type|canford|termination panel|"
                        r"hdx|cassette|outlet|floor ?box", re.I)
PASSIVE_LABEL_RE = re.compile(r"\.(FPP|EPP|CPP|BPP|NS|PP)\b", re.I)


def is_passive(end):
    return bool(PASSIVE_RE.search(end.get("template") or "") or PASSIVE_LABEL_RE.search(end.get("top") or ""))


# ---------------------------------------------------------------- detection

def detect_kind(filename, data):
    """Which input a file is: 'cables' | 'equipment' | 'portplan' | 'snags' | 'neuron' |
    'cvports' | None."""
    name = (filename or "").lower()
    if name.endswith(".csv"):
        head = data[:4000].decode("utf-8-sig", "replace") if isinstance(data, bytes) else data[:4000]
        first = head.splitlines()[0] if head else ""
        if "Cable Identifier" in first and "Connected Equipment (Left)" in first:
            return "cables"
        if "Equipment Identifier" in first and "Equipment Template" in first:
            return "equipment"
        if "Interface" in first and "Operational Status" in first and "Transceiver Type" in first:
            return "cvports"
        return None
    if name.endswith((".xlsx", ".xlsm")):
        sheets = read_xlsx(data, max_rows=40)
        for rows in sheets.values():
            for r in rows:
                cells = [c.strip().lower() for c in r]
                if "status manually" in cells or ("port cv" in cells and "switch cad" in cells):
                    return "snags"
                if "neuron" in cells and "card" in cells and any(c.startswith("bnc") for c in cells):
                    return "neuron"
        for rows in sheets.values():
            for r in rows:
                cells = [c.strip().lower() for c in r]
                if "switch" in cells and "port" in cells and "device" in cells:
                    return "portplan"
    return None


CVPORTS_FILE_RE = re.compile(r"cvp-ethernet-interfaces-(.+?)-(\d{4})-(\d\d)-(\d\d)(?:-(\d\d)-(\d\d)-(\d\d))?", re.I)


def cvports_time(filename):
    """Export time from CloudVision's file name: '...-2026-10-18-14-40-35.csv' -> '2026-10-18 14:40:35'."""
    m = CVPORTS_FILE_RE.search(filename or "")
    if not m:
        return ""
    y, mo, d, h, mi, se = m.group(2), m.group(3), m.group(4), m.group(5), m.group(6), m.group(7)
    return "%s-%s-%s" % (y, mo, d) + (" %s:%s:%s" % (h, mi, se) if h else "")


def meta_key(kind, filename):
    """Inputs that exist once ('cables') use the kind; CloudVision port exports exist once per
    switch: 'cvports:DCBESWLIVSPI001' (switch name from the CloudVision file name)."""
    if kind == "cvports":
        m = CVPORTS_FILE_RE.search(filename or "")
        if not m:
            raise ValueError("CloudVision interface export: keep CloudVision's file name "
                             "(cvp-ethernet-interfaces-<switch>-<date>.csv) - the switch name is only in it")
        return "cvports:" + m.group(1).upper()
    return kind


def parse_cvports(rows):
    """CloudVision 'Ethernet interfaces' export -> {nums: row}."""
    out = {}
    for r in rows:
        nums = parse_intf(r.get("Interface", ""))
        if not nums:
            continue
        out[nums] = {"intf": r.get("Interface", "").strip(), "description": r.get("Description", "").strip(),
                     "status": r.get("Operational Status", "").strip(), "speed": r.get("Speed", "").strip(),
                     "optic": r.get("Transceiver Type", "").strip(), "mtu": r.get("MTU", "").strip(),
                     "mode": r.get("Switch Port Type", "").strip(), "vlan": r.get("VLAN", "").strip(),
                     "changed": r.get("Last Link Status Change", "").strip()}
    return out


# ---------------------------------------------------------------- Neuron signal list

NEURON_COLS = {  # (group, column) in the two header rows -> field
    ("evs neuron", "room"): "room", ("evs neuron", "neuron"): "neuron", ("evs neuron", "card"): "card",
    ("evs neuron", "bnc conektor"): "bnc", ("evs neuron", "bnc connector"): "bnc", ("evs neuron", "chanel"): "channel",
    ("evs neuron", "channel"): "channel", ("evs neuron", "mode(use)"): "mode",
    ("patch panel", "room"): "pp_room", ("patch panel", "rack"): "pp_rack", ("patch panel", "pp number"): "pp",
    ("patch panel", "port"): "pp_port", ("patch panel", "rear cable"): "rear", ("patch panel", "front cable"): "front",
    ("endpoint", "room"): "ep_room", ("endpoint", "name"): "ep", ("endpoint", "port"): "ep_port",
    ("signal name", ""): "signal", ("revision", "comment"): "comment", ("revision", "revised by"): "revised_by",
    ("revision", "patch manager updated?"): "pm_updated", ("revision", "date"): "rev_date",
    ("cerebrum source", "mnemonic"): "src_mnemonic", ("cerebrum source", "button"): "src_button",
    ("cerebrum source", "umd"): "src_umd", ("cerebrum source", "qvest button"): "src_qbutton",
    ("cerebrum destination", "mnemonic"): "dst_mnemonic", ("cerebrum destination", "button"): "dst_button",
    ("cerebrum destination", "umd"): "dst_umd", ("cerebrum destination", "qvest button"): "dst_qbutton",
    ("contribution partner", "partner"): "partner",
}


def parse_neuron(sheets):
    """Neuron signal list: two header rows (group: 'EVS Neuron' / 'Patch Panel' / 'Endpoint' /
    'Cerebrum Source' ..., then column names). One row per BNC."""
    out = []
    for sheet, rows in sheets.items():
        for hi in range(1, min(len(rows), 15)):
            cells = [c.strip().lower() for c in rows[hi]]
            if "neuron" in cells and "card" in cells and any(c.startswith("bnc") for c in cells):
                break
        else:
            continue
        groups, head, cols, grp = rows[hi - 1], rows[hi], {}, ""
        for i in range(max(len(groups), len(head))):
            if i < len(groups) and groups[i].strip():
                grp = groups[i].strip().lower()
            col = head[i].strip().lower() if i < len(head) else ""
            field = NEURON_COLS.get((grp, col)) or (NEURON_COLS.get(("signal name", "")) if grp == "signal name" and not col else None)
            if field and field not in cols:
                cols[field] = i
        for ri in range(hi + 1, len(rows)):
            r = rows[ri]
            g = {f: (r[i].strip() if i < len(r) else "") for f, i in cols.items()}
            if not (g.get("neuron") and g.get("card") and g.get("bnc")):
                continue
            g["device"] = "%s.%s" % (g.get("room", ""), g["neuron"]) if g.get("room") else g["neuron"]
            g["pp_device"] = "%s.%s" % (g.get("pp_room") or g.get("room", ""), g["pp"]) if g.get("pp") else ""
            m = re.match(r"\s*(\d+)", g.get("channel", ""))
            g["ch"] = int(m.group(1)) if m else None
            g["dir"] = "out" if re.search(r"out", g.get("channel", ""), re.I) else "in" if re.search(r"\bin\b", g.get("channel", ""), re.I) else "bi"
            g["rev_date"] = excel_date(g.get("rev_date", ""))
            g["sheet"], g["row"] = sheet, ri + 1
            out.append(g)
    return out


CABLE_ID_RE = re.compile(r"^[A-Z]{1,4}\.?[A-Z]?\d{3,6}$")


# ---------------------------------------------------------------- Patchmanager cables

class End:
    """One end of a cable: equipment (top level / parent / label) and port."""
    __slots__ = ("top", "parent", "label", "port", "template", "location", "cabinet", "upos", "spec")

    def __init__(self, row, side):
        e = "Connected Equipment (%s) " % side
        c = "Connectors (%s) Connected Port " % side
        self.top = row.get(e + "Top Level Equipment Label", "").strip()
        self.parent = row.get(e + "Parent Equipment Label", "").strip()
        self.label = row.get(e + "Label", "").strip()
        self.port = row.get(c + "Label", "").strip()
        self.template = row.get(e + "Equipment Template Name", "").strip()
        self.location = row.get(e + "Position Location Name", "").strip()
        self.cabinet = row.get(e + "Position Cabinet Name", "").strip()
        self.upos = row.get(e + "Position Cabinet Position U Position", "").strip()
        self.spec = row.get(c + "Port Specification Template Name", "").strip()

    def key(self):
        """Identity of the connection point. Parent only counts when it adds information
        (line card LC03 under a chassis), not when it repeats the top level."""
        parent = self.parent if self.parent not in ("", self.top) else ""
        return (self.top, parent, self.label, norm_port(self.port))

    def module(self):
        """(top, parent, label) without port - a cassette / panel module."""
        return self.key()[:3]

    def empty(self):
        return not (self.top or self.label)

    def name(self):
        """Readable: 'DCB.FPP 104 / A : 02', 'DCB.VR 01 / LC01 / 28 : 28'."""
        parts = [self.top]
        if self.parent and self.parent != self.top:
            parts.append(self.parent)
        if self.label and self.label != self.top:
            parts.append(self.label)
        return " / ".join(p for p in parts if p) + (" : " + self.port if self.port else "")

    def to_dict(self):
        return {"name": self.name(), "top": self.top, "parent": self.parent, "label": self.label,
                "port": self.port, "template": self.template, "location": self.location,
                "cabinet": self.cabinet, "u": self.upos, "spec": self.spec}


class Cable:
    __slots__ = ("idx", "id", "label", "template", "length", "date", "attrs", "description", "a", "b")

    def __init__(self, idx, row):
        self.idx = idx
        self.id = row.get("Cable Identifier", "").strip()
        self.label = row.get("Cable Label", "").strip()
        self.template = row.get("Template Name", "").strip()
        self.length = row.get("Length (Calculated)", "").strip()
        self.date = row.get("History Date", "").strip()
        self.attrs = parse_attrs(row.get("User Attributes", ""))
        self.description = row.get("Description", "").strip()
        self.a = End(row, "Left")
        self.b = End(row, "Right")

    def other(self, side):
        return self.b if side == "a" else self.a

    def flags(self):
        """Patchmanager tick boxes that matter on site: not installed / not labelled / to be checked."""
        f = []
        if self.attrs.get("Installed") == "false":
            f.append("not installed")
        if self.attrs.get("Labeled") == "false":
            f.append("not labelled")
        if self.attrs.get("To Be Checked") == "true":
            f.append("to be checked")
        return f

    def to_dict(self):
        return {"id": self.id, "label": self.label, "template": self.template, "length": self.length,
                "date": self.date, "flags": self.flags(), "description": self.description,
                "stamp": self.attrs.get("Stamp", "")}


class Cabling:
    def __init__(self, rows):
        self.cables = [Cable(i, r) for i, r in enumerate(rows)]
        self.at = defaultdict(list)       # End.key() -> [(cable, side)]
        self.at_module = defaultdict(list)  # module key, port "" (trunks)  -> [(cable, side)]
        self.by_top = defaultdict(list)   # top level label -> [(cable, side)]
        self.by_id = defaultdict(list)
        for c in self.cables:
            self.by_id[c.id.upper()].append(c)
            for side, end in (("a", c.a), ("b", c.b)):
                if end.empty():
                    continue
                self.at[end.key()].append((c, side))
                self.by_top[norm_name(end.top)].append((c, side))
                if not end.port:
                    self.at_module[end.module()].append((c, side))

    # -- tracing ------------------------------------------------------------
    def trace(self, cable, side, max_hops=40):
        """Follow the cabling from `side` of `cable` (the start end) to the far end.

        At a patch panel the same point has a second cable (front / rear): continue there.
        A trunk (loose tube, MPO) is documented on the cassette without a port: jump to the
        trunk's far cassette and continue at the SAME port number (assumption, flagged).
        Returns {"hops": [...], "end": ..., "status": "complete" | "open" | "ambiguous" | "loop"}.
        """
        hops = [{"type": "end", **getattr(cable, side).to_dict()}]
        seen = {cable.idx}
        cur, cur_side = cable, side
        status = "open"
        for _ in range(max_hops):
            far = cur.other(cur_side)
            hops.append({"type": "cable", **cur.to_dict()})
            if far.empty():
                hops.append({"type": "end", "name": "(not connected in Patchmanager)", "missing": True})
                status = "open"
                break
            hops.append({"type": "end", **far.to_dict()})
            nxt = [(c, s) for c, s in self.at.get(far.key(), []) if c.idx not in seen]
            if len(nxt) == 1:
                cur, cur_side = nxt[0]
                seen.add(cur.idx)
                continue
            if len(nxt) > 1:
                status = "ambiguous"
                hops.append({"type": "note", "text": "%d further cables at %s: %s" % (
                    len(nxt), far.name(), ", ".join(c.id or "?" for c, _ in nxt[:8]))})
                break
            # trunk on the module (cassette) without port -> same port on the far module
            if far.port:
                trunks = [(c, s) for c, s in self.at_module.get(far.module(), []) if c.idx not in seen]
                dests = {}
                for c, s in trunks:
                    o = c.other(s)
                    if not o.empty():
                        dests.setdefault(o.module(), (c, s, o))
                if len(dests) == 1:
                    tc, ts, o = next(iter(dests.values()))
                    for c, _ in trunks:
                        seen.add(c.idx)
                    hops.append({"type": "cable", **tc.to_dict(), "trunk": True,
                                 "trunk_cables": [c.id for c, _ in trunks]})
                    tgt = (o.module()[0], o.module()[1], o.module()[2], norm_port(far.port))
                    cont = [(c, s) for c, s in self.at.get(tgt, []) if c.idx not in seen]
                    hops.append({"type": "end", "name": "%s : %s" % (
                        " / ".join(x for x in (o.top, o.parent if o.parent not in ("", o.top) else "",
                                               o.label if o.label != o.top else "") if x), far.port),
                        "top": o.top, "label": o.label, "port": far.port, "template": o.template,
                        "location": o.location, "cabinet": o.cabinet, "u": o.upos, "assumed": True})
                    if len(cont) == 1:
                        cur, cur_side = cont[0]
                        seen.add(cur.idx)
                        continue
                    status = "open" if not cont else "ambiguous"
                    break
                if len(dests) > 1:
                    status = "ambiguous"
                    hops.append({"type": "note", "text": "trunks from %s go to %d different panels: %s" % (
                        far.name(), len(dests), ", ".join(" / ".join(k[::2]) for k in dests))})
                    break
            status = "open" if is_passive(hops[-1]) else "complete"
            break
        else:
            status = "loop"
        return {"hops": hops, "status": status}

    # -- switch ports --------------------------------------------------------
    # Chassis switches: CloudVision counts the supervisor slots, Patchmanager (LC01..) and the
    # IP list (1/1/1) count line cards only -> CloudVision Ethernet3/1/1 = LC01 port 1 = IP list 1/1/1.
    slot_offset = 2

    def is_modular(self, switch):
        return any(re.fullmatch(r"LC\d+", e.parent or "") for c, s in self.by_top.get(norm_name(switch), [])
                   for e in [getattr(c, s)])

    def switch_port_ends(self, switch, nums):
        """Cable ends on the given Patchmanager switch for interface numbers `nums`
        (see module doc). Returns [(cable, side)]."""
        modular = self.is_modular(switch)
        if modular and len(nums) >= 2:
            slot, port = nums[0], nums[1]
        else:
            slot, port = None, nums[0]
        out = []
        for c, s in self.by_top.get(norm_name(switch), []):
            e = getattr(c, s)
            if norm_port(e.port) != str(port):
                continue
            if slot is not None:
                m = re.fullmatch(r"LC0*(\d+)", e.parent or "")
                if not m or int(m.group(1)) + self.slot_offset != slot:
                    continue
            out.append((c, s))
        return out

    def ports_of(self, switch):
        """All documented interfaces of a switch: sorted list of (nums, [(cable, side)])."""
        res = defaultdict(list)
        for c, s in self.by_top.get(norm_name(switch), []):
            e = getattr(c, s)
            p = norm_port(e.port)
            if not p.isdigit():
                continue
            m = re.fullmatch(r"LC0*(\d+)", e.parent or "")
            nums = (int(m.group(1)) + self.slot_offset, int(p)) if m else (int(p),)
            res[nums].append((c, s))
        return sorted(res.items())


# ---------------------------------------------------------------- Patchmanager equipment

class Equipment:
    def __init__(self, rows):
        self.by_label = defaultdict(list)
        # a part (line card, cassette, SFP) has its parent's identifier as Location;
        # follow that up to the top-level device
        by_id = {",".join(x.strip() for x in r.get("Equipment Identifier", "").split(",")): r for r in rows}

        def top_of(r):
            """Shortest prefix of the identifier path that is itself equipment = the device.
            ('...,DCB.A03,DCB.VR 01,LC01,LC01' -> 'DCB.VR 01'; slots like 'LC01' are not rows)"""
            parts = [x.strip() for x in r.get("Equipment Identifier", "").split(",")]
            for n in range(1, len(parts) + 1):
                hit = by_id.get(",".join(parts[:n]))
                if hit:
                    return hit.get("Equipment Label", "").strip()
            return r.get("Equipment Label", "").strip()

        for r in rows:
            label = r.get("Equipment Label", "").strip()
            if not label:
                continue
            ident = [p.strip() for p in r.get("Equipment Identifier", "").split(",")]
            attrs = parse_attrs(r.get("Equipment Attributes", ""))
            top = top_of(r)
            alias = attrs.get("Alias", "").replace("<EQUALS>", "=")
            alias = alias.split("=")[-1].strip()  # 'Edit1<EQUALS>CS01 AMU 01' -> 'CS01 AMU 01'
            self.by_label[norm_name(label)].append({
                "label": label, "top": top, "alias": alias, "asset": attrs.get("Asset Tag", ""),
                "template": r.get("Equipment Template", "").strip(),
                "location": " > ".join(p for p in r.get("Equipment Location", "").split(",")[2:]),
                "path": ident,
                "position": r.get("Equipment Position", "").strip(),
                "description": r.get("Equipment Description", "").strip(),
                "attrs": attrs,
                "defect": [norm_port(p) for p in (r.get("Defect Ports Front", "") + "," + r.get("Defect Ports Rear", "")).split(",") if p.strip()],
                "reserved": [norm_port(p) for p in (r.get("Reserved Ports Front", "") + "," + r.get("Reserved Ports Rear", "")).split(",") if p.strip()],
            })
        self.count = sum(len(v) for v in self.by_label.values())

    def get(self, label, under=None):
        """Equipment by label; `under` (a top-level label) picks the right one when a label
        like 'LC01' or '49' exists many times."""
        items = self.by_label.get(norm_name(label), [])
        if under and len(items) > 1:
            u = norm_name(under)
            items = [i for i in items if any(norm_name(p) == u for p in i["path"])] or items
        return items[0] if items else None


# ---------------------------------------------------------------- IP list (switch IP list, "View" xlsx)

PLAN_DEV_RE = re.compile(r"^\s*(.*?)\s*--\s*(.*)$")


def parse_portplan(sheets):
    """Every sheet: rows with 'Switch' / 'Port' / 'Device' headers start blocks of columns.
    Block layout either Switch|Port|Device|VLAN|Cable or Switch|<slot>|Port|Device|VLAN|Cable.
    Returns list of {switch, nums, device, equipment, eqport, where, vlan, cable, sheet, row}."""
    out = []
    for sheet, rows in sheets.items():
        blocks = None
        for ri, r in enumerate(rows):
            cells = [c.strip().lower() for c in r]
            if "switch" in cells and "port" in cells and "device" in cells:
                blocks = []
                for ci, c in enumerate(cells):
                    if c != "switch":
                        continue
                    # Port at +1 (no slot) or +2 (slot in between, its header empty)
                    if ci + 1 < len(cells) and cells[ci + 1] == "port":
                        blocks.append((ci, None, ci + 1))
                    elif ci + 2 < len(cells) and cells[ci + 2] == "port":
                        blocks.append((ci, ci + 1, ci + 2))
                continue
            if not blocks:
                continue
            for sc, slc, pc in blocks:
                g = lambda i: r[i].strip() if i is not None and i < len(r) else ""
                sw, slot, port, dev = g(sc), g(slc), g(pc), g(pc + 1)
                if not sw or not port:
                    continue
                if not (port.isdigit() and (not slot or slot.isdigit())):
                    continue
                nums = (int(slot), int(port)) if slot else (int(port),)
                eq, eqport, where = "", "", ""
                if "---" in dev:  # core / border sheets: 'DCR.BSWR 01.1.1 --- DCR.VR 01.1.1' (link, both ends)
                    sides = [x.strip() for x in dev.split("---")]
                    other = [x for x in sides if not norm_name(x).startswith(norm_name(sw))] or sides[-1:]
                    dev = other[0] + " -- "
                m = PLAN_DEV_RE.match(dev)
                if m:
                    left, where = m.group(1), m.group(2).strip()
                    # 'DCB.VS 02.PCLAN1' -> equipment 'DCB.VS 02', port 'PCLAN1'
                    mm = re.match(r"^(.*\S)\.([^.\s]+(?:\.\d+)?)$", left)
                    eq, eqport = (mm.group(1), mm.group(2)) if mm else (left, "")
                    # 'SFP.1', 'ETH.1': the port is the last two parts when the last is a number
                    mm2 = re.match(r"^(.*\S)\.([A-Za-z0-9]+\.\d+)$", left)
                    if mm2:
                        eq, eqport = mm2.group(1), mm2.group(2)
                out.append({"switch": sw, "nums": list(nums), "device": dev, "equipment": eq,
                            "eqport": eqport, "where": where, "vlan": g(pc + 2), "cable": g(pc + 3),
                            "sheet": sheet.strip(), "row": ri + 1,
                            "free": dev.lower() in ("", "free", "spare")})
    # the same block can appear on two sheets (e.g. 'Access IDF1' and a personal copy): keep the first
    seen, uniq = set(), []
    for p in out:
        k = (norm_name(p["switch"]), tuple(p["nums"]), norm_name(p["device"]))
        if k not in seen:
            seen.add(k)
            uniq.append(p)
    return uniq


# ---------------------------------------------------------------- snagging list

SNAG_COLS = {  # our field -> header names seen in the list (lower case)
    "nr": ["nr"],
    "source": ["source port + cv"],
    "cv_switch": ["switch name  cv", "switch name cv", "switch"],
    "pm_switch": ["switch cad"],
    "cv_port": ["port cv", "port"],
    "pm_port": ["port cad"],
    "error_type": ["error type"],
    "error": ["error description"],
    "comment": ["comment"],
    "destination": ["destination"],
    "identifier": ["identifier arista"],
    "status": ["status manually"],
    "cv_status": ["status in arista"],
    "first": ["date first occurance", "date first occurrence"],
    "fixed": ["date rectification"],
    "action": ["what has been done when"],
    "who": ["name"],
    "cabling": ["cabling"],
    "jira": ["jira"],
}
CLOSED_STATUS = {"fixed", "closed", "duplicate", "ignore", "ignore for now"}


def parse_snags(sheets):
    """Returns (snags, mapping). snags: rows of sheets that have a 'status manually' column
    (the Master list). mapping: {PM name: CV hostname} from a sheet with 'PM' / 'arista cloud'."""
    snags, mapping = [], {}
    has_master = any(n.strip().lower() == "master" for n in sheets)
    for sheet, rows in sheets.items():
        hdr_i, cols = None, {}
        for ri, r in enumerate(rows[:15]):
            cells = [c.strip().lower() for c in r]
            if "pm" in cells and any("arista" in c for c in cells):
                pi = cells.index("pm")
                ai = next(i for i, c in enumerate(cells) if "arista" in c)
                for r2 in rows[ri + 1:]:
                    pm = r2[pi].strip() if pi < len(r2) else ""
                    cv = r2[ai].strip() if ai < len(r2) else ""
                    if pm and cv and cv != "0":
                        mapping[pm] = cv
                break
            if "status manually" in cells:
                hdr_i = ri
                for field, names in SNAG_COLS.items():
                    for n in names:
                        if n in cells:
                            cols[field] = cells.index(n)
                            break
                break
        # other sheets with the same header are raw event dumps - only the Master list counts
        if hdr_i is None or (has_master and sheet.strip().lower() != "master"):
            continue
        for ri in range(hdr_i + 1, len(rows)):
            r = rows[ri]
            if not any(c.strip() for c in r):
                continue
            g = {f: (r[i].strip() if i < len(r) else "") for f, i in cols.items()}
            g = {f: ("" if v == "0" and f not in ("nr",) else v) for f, v in g.items()}  # '0' = empty formula
            if not (g.get("cv_switch") or g.get("pm_switch") or g.get("source")):
                continue
            for f in ("first", "fixed"):
                g[f] = excel_date(g.get(f, ""))
            if g.get("jira", "").startswith("#"):  # '#VALUE!' from a broken formula
                g["jira"] = ""
            st = g.get("status", "").strip().lower()
            g["open"] = not any(st.startswith(x) for x in CLOSED_STATUS)
            g["sheet"], g["row"] = sheet, ri + 1
            port = g.get("cv_port") or g.get("pm_port")
            g["nums"] = list(parse_intf(port) or ())
            snags.append(g)
    return snags, mapping


# ---------------------------------------------------------------- the whole picture

class Store:
    """Everything imported, plus the lookups across sources."""

    def __init__(self):
        self.cabling = None
        self.equipment = None
        self.plan = []
        self.snags = []
        self.mapping = {}       # PM switch name -> CV hostname (snagging list + overrides)
        self.overrides = {}
        self._pm_by_norm, self._cv_by_norm, self._plan_sw_by_norm = {}, {}, {}
        self.future = {}        # device -> {"note", "since"}: marked "future use" in the page
        self.neuron = []        # Neuron signal list rows (one per BNC)
        self.cvports = {}       # CloudVision hostname (upper) -> {nums: interface row}
        self.meta = {}          # kind -> {"file", "loaded", "count"}
        self._rebuild_names()

    # -- import --------------------------------------------------------------
    def load(self, kind, filename, data, when=None):
        when = when or datetime.now().isoformat(timespec="seconds")
        if kind == "cables":
            self.cabling = Cabling(read_csv_text(data))
            n = len(self.cabling.cables)
        elif kind == "equipment":
            self.equipment = Equipment(read_csv_text(data))
            n = self.equipment.count
        elif kind == "portplan":
            self.plan = parse_portplan(read_xlsx(data))
            n = len(self.plan)
        elif kind == "snags":
            self.snags, self.mapping_snags = parse_snags(read_xlsx(data))
            n = len(self.snags)
        elif kind == "neuron":
            self.neuron = parse_neuron(read_xlsx(data))
            n = len(self.neuron)
        elif kind == "cvports":
            host = meta_key(kind, filename).split(":", 1)[1]
            rows = parse_cvports(read_csv_text(data))
            exported = cvports_time(filename)
            for r in rows.values():
                r["exported"], r["host"] = exported, host
            self.cvports[host] = rows
            n = len(rows)
        else:
            raise ValueError("unknown kind %r" % kind)
        self.meta[meta_key(kind, filename)] = {"file": filename, "loaded": when, "count": n}
        self._rebuild_names()
        return n

    def set_overrides(self, overrides):
        self.overrides = {k: v for k, v in (overrides or {}).items() if k and v}
        self._rebuild_names()

    def _rebuild_names(self):
        self.mapping = dict(getattr(self, "mapping_snags", {}) or {})
        self.mapping.update(self.overrides)
        self.cv_to_pm = {}
        dup = defaultdict(list)
        for pm, cv in self.mapping.items():
            dup[cv.upper()].append(pm)
            self.cv_to_pm.setdefault(cv.upper(), pm)
        self.mapping_dups = {cv: pms for cv, pms in dup.items() if len(pms) > 1}
        self._pm_by_norm = {norm_name(pm): pm for pm in self.mapping}
        self._cv_by_norm = {norm_name(pm): cv for pm, cv in self.mapping.items()}
        self._plan_sw_by_norm = {}
        for p in self.plan:
            self._plan_sw_by_norm.setdefault(norm_name(p["switch"]), p["switch"])
        # IP list slot numbers of chassis switches -> CloudVision numbering (see Cabling.slot_offset);
        # nums_doc keeps what the list says
        modular = {}
        for p in self.plan:
            p.setdefault("nums_doc", list(p["nums"]))
            sw = norm_name(p["switch"])
            if sw not in modular:
                modular[sw] = bool(self.cabling) and self.cabling.is_modular(p["switch"])
            nd = p["nums_doc"]
            p["nums"] = [nd[0] + self.cabling.slot_offset, nd[1]] if modular[sw] and len(nd) == 2 else list(nd)
        self.plan_by = defaultdict(list)
        for p in self.plan:
            self.plan_by[(norm_name(p["switch"]), tuple(p["nums"]))].append(p)
        self.snags_by = defaultdict(list)
        for s in self.snags:
            sw = self.pm_name(s.get("pm_switch") or s.get("cv_switch"))
            if sw and s["nums"]:
                self.snags_by[(norm_name(sw), tuple(s["nums"]))].append(s)
        self._build_name_index()
        self._shut_cache = None
        self._noff = None
        self._swports = None

    # -- device names -------------------------------------------------------------
    def _build_name_index(self):
        """Every name a device is known by -> the device (its Patchmanager top-level label).
        Sources: Patchmanager label / alias / asset tag / description, cable ends, cable
        description, IP list device names, snagging-list destinations, CloudVision hostnames."""
        idx = defaultdict(dict)   # strict key -> {device: (name, via)}
        comp = defaultdict(dict)  # glued-together key ('DCBVS2') -> {device: (name, via)}

        def add(name, device, via):
            name, device = str(name or "").strip(" *"), str(device or "").strip(" *")
            if len(norm_name(name)) < 2 or not device:
                return
            for k in {norm_name(name), norm_digits(name)}:
                idx[k].setdefault(device, (name, via))
            for k in {compact(name), _ZERO_RE.sub("", compact(name))}:   # for typed-together searches
                comp[k].setdefault(device, (name, via))

        self.switch_set = {norm_name(k) for k in self.mapping} | {norm_name(p["switch"]) for p in self.plan}
        self.switch_names = {norm_name(k): k for k in self.mapping}
        for p in self.plan:
            self.switch_names.setdefault(norm_name(p["switch"]), p["switch"])
        if self.equipment:
            for items in self.equipment.by_label.values():
                for e in items:
                    top = e["top"]
                    if norm_name(e["label"]) == norm_name(top):
                        add(e["label"], top, "Patchmanager")
                        add(e["description"], top, "Patchmanager description")
                        if re.match(r"^Arista (CCS|DCS)-", e["template"]):
                            self.switch_set.add(norm_name(top))
                            self.switch_names.setdefault(norm_name(top), top)
                    elif re.search(r"[A-Za-z]", e["label"]) and len(e["label"]) > 4 and "." in e["label"]:
                        add(e["label"], top, "Patchmanager (part of %s)" % top)   # 'DCB.ICM 02' inside a frame
                    add(e["alias"], top, "Patchmanager alias")
                    add(e["asset"], top, "asset tag")
        if self.cabling:
            for c in self.cabling.cables:
                for end in (c.a, c.b):
                    add(end.top, end.top, "Patchmanager cabling")
                if c.description:
                    for end in (c.a, c.b):
                        if norm_name(c.description) in norm_name(end.top) or len(c.description) > 6:
                            add(c.description, end.top, "cable %s description" % c.id)
        self._known = {k: next(iter(v)) for k, v in idx.items()}
        device_for = lambda name: self.device_split(name)[0]

        for p in self.plan:
            if p["equipment"] and not p["free"]:
                left = PLAN_DEV_RE.match(p["device"])
                p["_dev"], p["_devport"] = self.device_split(left.group(1) if left else p["equipment"])
                if not p["_devport"]:
                    p["_devport"] = p["eqport"]
                add(p["equipment"], p["_dev"], "IP list (%s)" % p["sheet"])
        for s in self.snags:
            d = (s.get("destination") or "").strip(" *")
            if d:
                s["_dev"], s["_devport"] = self.device_split(d)
                if s["_dev"] == d:  # unknown device: strip a trailing port-looking part
                    m = re.match(r"^(.*?)\.((?:[A-Za-z]*\d+|ETH|LAN|MGMT|MGNT)(?:\.\w+)?)$", d)
                    if m:
                        s["_dev"], s["_devport"] = m.group(1), m.group(2)
                add(s["_dev"], s["_dev"], "snagging list destination")
        for r in getattr(self, "neuron", []):
            dev = device_for(r["device"])
            for f, via in (("src_mnemonic", "Cerebrum source"), ("dst_mnemonic", "Cerebrum destination"),
                           ("src_umd", "Cerebrum source UMD"), ("dst_umd", "Cerebrum destination UMD"),
                           ("src_qbutton", "Cerebrum button"), ("dst_qbutton", "Cerebrum button"),
                           ("signal", "Neuron signal name")):
                if r.get(f):
                    add(r[f], dev, "%s (%s %s BNC %s)" % (via, r["device"], r["card"], r["bnc"]))
        for pm, cv in self.mapping.items():
            add(cv, pm, "CloudVision hostname")
            add(pm, pm, "Patchmanager")
        # names that are only known from plan / snags but are cabled under another spelling:
        # point them at the cabled device when the normalised name is the same
        self.name_index = idx
        self.name_compact = comp

    def device_split(self, name):
        """'DCB.ICM 02.1.SFP' -> ('DCB.ICM 02', '1.SFP') when 'DCB.ICM 02' is a known device;
        'DCR.ESW B01.50' -> ('DCR.ESW B01', '50'). Unknown: (name, '')."""
        name = str(name or "").strip(" *")
        known = getattr(self, "_known", {})
        # CloudVision-style link description: 'P2P_dcreswcprcor001_Ethernet15/2'
        m = re.match(r"^P2P_([\w\-]+?)_((?:Ethernet|Et)\d+(?:/\d+)*)$", name, re.I)
        if m:
            return self.pm_name(m.group(1)) or m.group(1), m.group(2)
        # 'DCRSRVIRAXSV003 *DCR.VS 23' / 'A; B': several names - take the first known one
        pieces = [x.strip() for x in re.split(r"[*;]|\s{2,}", name) if x.strip()]
        if len(pieces) > 1:
            for x in pieces:
                d, rest = self.device_split(x)
                if norm_name(d) in known:
                    return d, rest
            name = pieces[0]
        # free text after the name: 'DCB.BKP 02 DCB B06 Veeam ...' -> 'DCB.BKP 02'
        words = name.split(" ")
        for n in range(len(words), 1, -1):
            d = known.get(norm_name(" ".join(words[:n])))
            if d and n < len(words):
                return d, ""
        parts = name.split(".")
        for n in range(len(parts), 0, -1):
            d = known.get(norm_name(".".join(parts[:n])))
            if d:
                return d, ".".join(parts[n:])
        # unknown device: names end in ' <number>' ('OC.IP/HL 01.100G A.1' -> 'OC.IP/HL 01');
        # hostnames have no space ('MC1STBDGTPLY001.ETH.1' -> 'MC1STBDGTPLY001')
        m = re.match(r"^(.*?\s\d+[A-Za-z]?)\.(.+)$", name) or re.match(r"^([^.\s]+\d)\.(.+)$", name)
        if m:
            return m.group(1), m.group(2)
        return name, ""

    def device_of(self, name):
        """Best device for a name (exact key), or None."""
        for k in (norm_name(name), norm_digits(name)):
            if k in self.name_index:
                return next(iter(self.name_index[k]))
        comp = getattr(self, "name_compact", {})
        for k in (compact(name), _ZERO_RE.sub("", compact(name))):
            if k in comp and len(comp[k]) == 1:      # only when unambiguous
                return next(iter(comp[k]))
        return None

    def find_devices(self, q, limit=40):
        """Name search: exact, then prefix, then substring (>= 3 characters)."""
        keys = {norm_name(q), norm_digits(q)}
        found = {}  # device -> (rank, name, via)
        for k in keys:
            for dev, (name, via) in self.name_index.get(k, {}).items():
                found.setdefault(dev, (0, name, via))
        comp = getattr(self, "name_compact", {})
        for k in {compact(q), _ZERO_RE.sub("", compact(q))}:
            for dev, (name, via) in comp.get(k, {}).items():
                found.setdefault(dev, (0, name, via))
        k0 = compact(q)
        if len(found) < limit and len(k0) >= 2:
            for key, devs in comp.items():
                for k in (k0,):
                    rank = 1 if key.startswith(k) else 2 if (len(k) >= 3 and k in key) else None
                    if rank is None:
                        continue
                    for dev, (name, via) in devs.items():
                        if dev not in found or found[dev][0] > rank:
                            found[dev] = (rank, name, via)
        # one result per device, prefer the Patchmanager spelling of the device itself
        out = sorted(found.items(), key=lambda kv: (kv[1][0], len(kv[0]), kv[0]))
        return [(dev, name, via) for dev, (rank, name, via) in out[:limit]]

    def is_switch(self, name):
        return norm_name(name) in getattr(self, "switch_set", set())

    def device_card(self, name, cv_events=None):
        """Everything about one device: Patchmanager data, every cabled port traced to its far
        end (with the switch port to open), IP list rows and snags that name it."""
        dev = self.device_of(name) or name
        nd = norm_name(dev)
        card = {"device": dev, "names": [], "equipment": None, "connections": [], "plan": [], "snags": [],
                "is_switch": self.is_switch(dev), "cv_switch": self.cv_name(dev) or ""}
        names = {}
        for key, devs in self.name_index.items():
            if dev in devs:
                n, via = devs[dev]
                names.setdefault(norm_name(n), (n, via))
        card["names"] = sorted(names.values(), key=lambda x: x[0].lower())
        if self.equipment:
            e = self.equipment.get(dev)
            if e:
                card["equipment"] = {k: e[k] for k in ("label", "template", "location", "position", "description", "alias", "asset")}
        if self.cabling:
            for c, s in sorted(self.cabling.by_top.get(nd, []), key=lambda cs: port_sort(getattr(cs[0], cs[1]))):
                me = getattr(c, s)
                path = self.cabling.trace(c, s)
                ends = [h for h in path["hops"] if h["type"] == "end" and not h.get("missing")]
                far = ends[-1] if len(ends) > 1 else None
                link = None
                if far and self.is_switch(far.get("top")) and not card["is_switch"]:
                    link = self._switch_ref(far)
                elif card["is_switch"]:
                    link = self._switch_ref(me.to_dict())
                card["connections"].append({"port": " / ".join(x for x in (
                    me.parent if me.parent not in ("", me.top) else "", me.label if me.label != me.top else "", me.port) if x),
                    "far": far, "status": path["status"], "path": path, "switch_link": link})
        for p in self.plan:
            if norm_name(p.get("_dev") or p["equipment"]) == nd or (nd and norm_name(p["switch"]) == nd and not p["free"]):
                card["plan"].append(p)
        for s in self.snags:
            if nd and norm_name(s.get("_dev")) == nd or (card["is_switch"] and norm_name(self.pm_name(
                    s.get("pm_switch") or s.get("cv_switch")) or "") == nd):
                card["snags"].append(s)
        card["compare"] = self._compare_device(card)
        card["future"] = self.future_of(dev)
        card["neuron"] = self.neuron_rows(dev) if self.neuron else []
        card["neuron_summary"] = [x for x in self.neuron_summary() if same_device(x["device"], dev)] if self.neuron else []
        return card

    def _compare_device(self, card):
        """Per device port: switch port in the IP list vs. where Patchmanager cabling goes.
        verdict: ok | switch (cabled to another switch port) | no cable (plan only) | not in plan."""
        if card["is_switch"]:
            return []
        rows, used = [], set()
        conns = [c for c in card["connections"] if c["switch_link"]]
        for p in card["plan"]:
            dport = p.get("_devport") or p["eqport"]
            plan_ref = {"switch": p["switch"], "nums": p["nums"]}
            hit = next((c for c in conns if dport and same_port(dport, c["port"])), None)
            if not hit:  # same switch port, other device-port spelling
                hit = next((c for c in conns if same_switch_port(c["switch_link"], plan_ref)), None)
            if hit:
                used.add(id(hit))
                verdict = "ok" if same_switch_port(hit["switch_link"], plan_ref) else "switch"
                rows.append({"port": dport or hit["port"], "plan": plan_ref, "cabled": hit["switch_link"],
                             "cabled_port": hit["port"], "verdict": verdict})
            else:
                rows.append({"port": dport, "plan": plan_ref, "cabled": None, "cabled_port": "", "verdict": "no cable"})
        for c in conns:
            if id(c) not in used:
                rows.append({"port": c["port"], "plan": None, "cabled": c["switch_link"], "cabled_port": c["port"],
                             "verdict": "not in plan"})
        return rows

    # -- the ports of one switch (for stepping through them) -----------------------------------
    def switch_ports(self, switch):
        """All ports known for a switch, sorted: CloudVision export interfaces (exact, with lanes)
        where there is an export, otherwise / plus the documented ports (Patchmanager, IP list,
        snags). A base port is dropped when its lanes are listed."""
        pm = self.pm_name(switch) or switch
        key = norm_name(pm)
        cache = getattr(self, "_swports", None)
        if cache is None:
            cache = self._swports = {}
        if key in cache:
            return cache[key]
        nums = set()
        host = (self.cv_name(pm) or "").upper()
        nums |= set(self.cvports.get(host, {}).keys())
        if self.cabling:
            nums |= {tuple(n) for n, _ in self.cabling.ports_of(pm)}
        modular = bool(self.cabling) and self.cabling.is_modular(pm)
        for p in self.plan:
            if norm_name(p["switch"]) == key and not p["free"]:
                n = tuple(p["nums"])
                nums.add(n[1:] if len(n) == 2 and n[0] == 1 and not modular else n)
        for (sw, n) in self.snags_by:
            if sw == key:
                nums.add(tuple(n))
        # a base port (Ethernet3/13) is covered by its lanes (Ethernet3/13/1..4)
        out = sorted(n for n in nums if not any(len(m) > len(n) and m[:len(n)] == n for m in nums))
        cache[key] = out
        return out

    def neighbours(self, switch, nums):
        """(index, total, previous, next) of a port in switch_ports - stops at the ends."""
        ports = self.switch_ports(switch)
        if not ports:
            return None
        nums = tuple(nums)
        if nums in ports:
            i = ports.index(nums)
        else:   # e.g. base port Ethernet3/1 while the export lists Ethernet3/1/1
            i = next((k for k, p in enumerate(ports) if p[:len(nums)] == nums or p > nums), len(ports) - 1)
        return {"index": i + 1, "total": len(ports), "prev": list(ports[i - 1]) if i > 0 else None,
                "next": list(ports[i + 1]) if i + 1 < len(ports) else None,
                "ports": [intf_name(p) for p in ports]}

    # -- CloudVision interface exports ------------------------------------------------------
    def cv_ports_for(self, pm, nums, modular=None):
        """CloudVision export rows (lanes) for a switch port, [] when no export of that switch."""
        host = (self.cv_name(pm) or "").upper()
        rows = self.cvports.get(host) if host else None
        if not rows:
            return []
        nums = tuple(nums)
        if nums in rows:
            return [rows[nums]]
        if modular is None:
            modular = bool(self.cabling) and self.cabling.is_modular(pm)
        base = nums[:2] if modular else nums[:1]
        return [rows[k] for k in sorted(rows) if k[:len(base)] == base]

    # -- Neuron signal list ---------------------------------------------------------------------
    def neuron_offsets(self):
        """Per IPG card and per patch panel: a numbering offset between the Neuron list and
        Patchmanager that holds for most rows (e.g. a whole card documented +16). Reported once,
        instead of as one error per BNC. {("ipg", device, card) | ("pp", device): (offset, rows)}"""
        if getattr(self, "_noff", None) is not None:
            return self._noff
        diffs = defaultdict(list)
        digits = lambda p: int(re.search(r"\d+", str(p)).group()) if re.search(r"\d+", str(p or "")) else None
        for r in self.neuron if self.cabling else []:
            for c in self.cabling.by_id.get((r.get("rear") or "").upper(), [])[:1]:
                ipg = next((e for e in (c.a, c.b) if same_device(e.top, r["device"])), None)
                pp = next((e for e in (c.a, c.b) if r["pp_device"] and same_device(e.top, r["pp_device"])), None)
                if ipg and digits(ipg.port) is not None and r.get("ch") is not None:
                    diffs[("ipg", norm_name(r["device"]), norm_name(r["card"]))].append(digits(ipg.port) - r["ch"])
                if pp and digits(pp.port) is not None and digits(r.get("pp_port")) is not None:
                    diffs[("pp", norm_name(r["pp_device"]))].append(digits(pp.port) - digits(r["pp_port"]))
        self._noff = {}
        for k, ds in diffs.items():
            off, n = Counter(ds).most_common(1)[0]
            if off != 0 and n >= 4 and n >= 0.6 * len(ds):
                self._noff[k] = (off, n)
        return self._noff

    def neuron_summary(self):
        """One 'check' line per IPG card / patch panel with a consistent numbering offset."""
        out = []
        for k, (off, n) in sorted(self.neuron_offsets().items()):
            r = next((x for x in self.neuron if (k[0] == "ipg" and norm_name(x["device"]) == k[1] and norm_name(x["card"]) == k[2])
                      or (k[0] == "pp" and norm_name(x["pp_device"]) == k[1])), None)
            if not r:
                continue
            what = "%s card %s: Patchmanager IPG port = list channel %+d" % (r["device"], r["card"], off) if k[0] == "ipg" \
                else "%s: Patchmanager panel port = list port %+d" % (r["pp_device"], off)
            out.append({"device": r["device"] if k[0] == "ipg" else r["pp_device"], "kind": k[0],
                        "text": "%s for %d rows - whole %s shifted, check which document is right" % (
                            what, n, "card" if k[0] == "ipg" else "panel")})
        return out

    def neuron_check(self, r):
        """[(severity, text)]: does Patchmanager have the list's rear / front cable where the list
        says? Patchmanager names the IPG port by channel + direction ('04>' / '>01'), the panel
        port R12 / F12."""
        out = []
        if not self.cabling:
            return out
        def digits(p):
            m = re.search(r"\d+", str(p or ""))
            return int(m.group()) if m else None
        rear = (r.get("rear") or "").strip()
        if CABLE_ID_RE.match(rear):
            cables = self.cabling.by_id.get(rear.upper(), [])
            if not cables:
                out.append(("warning", "Rear cable %s is not in Patchmanager" % rear))
            for c in cables[:1]:
                ipg = next((e for e in (c.a, c.b) if same_device(e.top, r["device"])), None)
                pp = next((e for e in (c.a, c.b) if r["pp_device"] and same_device(e.top, r["pp_device"])), None)
                if not ipg:
                    out.append(("error", "Rear cable %s does not end on %s in Patchmanager: %s <-> %s" % (rear, r["device"], c.a.name(), c.b.name())))
                else:
                    if norm_name(r["card"]) not in (norm_name(ipg.label), norm_name(ipg.parent)):
                        out.append(("error", "Rear cable %s: Patchmanager card %s, list card %s" % (rear, ipg.label or ipg.parent, r["card"])))
                    pnum = digits(ipg.port)
                    off = self.neuron_offsets().get(("ipg", norm_name(r["device"]), norm_name(r["card"])), (0,))[0]
                    if pnum is not None and r.get("ch") is not None and off and pnum - r["ch"] == off:
                        pass  # part of a card-wide shift, reported once (neuron_summary)
                    elif pnum is not None and pnum not in (r.get("ch"), digits(r.get("bnc"))):
                        out.append(("error", "Rear cable %s: Patchmanager IPG port %s, list channel %s (BNC %s)" % (rear, ipg.port, r.get("channel"), r.get("bnc"))))
                if r["pp_device"] and not pp and (c.a.empty() or c.b.empty()):
                    out.append(("warning", "Rear cable %s: Patchmanager has only one end connected (%s) - the panel end is missing" % (
                        rear, (c.b if c.a.empty() else c.a).name())))
                elif r["pp_device"] and not pp:
                    out.append(("error", "Rear cable %s does not end on %s in Patchmanager: %s <-> %s" % (rear, r["pp_device"], c.a.name(), c.b.name())))
                elif pp and digits(pp.port) != digits(r.get("pp_port")):
                    off = self.neuron_offsets().get(("pp", norm_name(r["pp_device"])), (0,))[0]
                    if not (off and digits(pp.port) is not None and digits(r.get("pp_port")) is not None
                            and digits(pp.port) - digits(r.get("pp_port")) == off):
                        out.append(("error", "Rear cable %s: Patchmanager panel port %s, list port %s" % (rear, pp.port, r.get("pp_port"))))
        front = (r.get("front") or "").strip()
        if CABLE_ID_RE.match(front):
            cables = self.cabling.by_id.get(front.upper(), [])
            if not cables:
                out.append(("warning", "Front cable %s is not in Patchmanager" % front))
            for c in cables[:1]:
                pp = next((e for e in (c.a, c.b) if r["pp_device"] and same_device(e.top, r["pp_device"])), None)
                other = c.b if pp is c.a else c.a
                if r["pp_device"] and not pp:
                    out.append(("error", "Front cable %s does not end on %s in Patchmanager: %s <-> %s" % (front, r["pp_device"], c.a.name(), c.b.name())))
                elif pp and digits(pp.port) != digits(r.get("pp_port")):
                    out.append(("error", "Front cable %s: Patchmanager panel port %s, list port %s" % (front, pp.port, r.get("pp_port"))))
                ep = (r.get("ep") or "").strip()
                # the list often writes the endpoint without the room: 'SWI 01.SDPE 1' = 'DCR.SWI 01 / SDPE 1'
                cands = {norm_name(other.top), norm_name(other.top.split(".", 1)[-1])}
                if pp and ep and not other.empty() and not any(c_ and c_ in norm_name(ep) for c_ in cands) \
                        and norm_name(ep) not in norm_name(other.name()):
                    out.append(("error", "Front cable %s goes to %s in Patchmanager, list endpoint %s %s" % (front, other.name(), ep, r.get("ep_port", ""))))
        return out

    def neuron_rows(self, dev):
        """Neuron list rows that involve a device (the IPG, its patch panel or the endpoint), with checks."""
        out = []
        for r in self.neuron:
            if same_device(r["device"], dev) or (r["pp_device"] and same_device(r["pp_device"], dev)) or \
                    (r.get("ep") and norm_name(dev) and norm_name(r["ep"]).startswith(norm_name(dev))):
                out.append(dict(r, checks=self.neuron_check(r)))
        return out

    def _switch_ref(self, end):
        """A cable end on a switch -> {switch, nums} for the port card, or None."""
        p = norm_port(end.get("port"))
        if not p.isdigit():
            return None
        m = re.fullmatch(r"LC0*(\d+)", end.get("parent") or "")
        off = self.cabling.slot_offset if self.cabling else 0
        return {"switch": end.get("top"), "nums": [int(m.group(1)) + off, int(p)] if m else [int(p)]}

    # -- names -------------------------------------------------------------------
    def pm_name(self, name):
        """Any switch name (PM or CV hostname, any case) -> Patchmanager name, or None."""
        if not name:
            return None
        n = name.strip()
        if n.upper() in self.cv_to_pm:
            return self.cv_to_pm[n.upper()]
        nn = norm_name(n)
        if nn in self._pm_by_norm:
            return self._pm_by_norm[nn]
        if self.cabling and nn in self.cabling.by_top:
            c, s = self.cabling.by_top[nn][0]
            return getattr(c, s).top
        return self._plan_sw_by_norm.get(nn)

    def cv_name(self, pm):
        if not pm:
            return None
        return self._cv_by_norm.get(norm_name(pm))

    def switches(self):
        """All switches we know of (PM names) with their CV hostname."""
        names = {}
        for pm in self.mapping:
            names[norm_name(pm)] = pm
        for p in self.plan:
            names.setdefault(norm_name(p["switch"]), p["switch"])
        return sorted(({"pm": pm, "cv": self.cv_name(pm) or ""} for pm in names.values()), key=lambda x: x["pm"])

    # -- the port card -----------------------------------------------------------
    def port_card(self, switch, nums, cv_events=None, with_shut=True, nav=False):
        """Everything about one switch port. `switch` PM or CV name, `nums` tuple."""
        pm = self.pm_name(switch) or switch
        cv = self.cv_name(pm) or (switch if switch and switch != pm else "")
        nums = tuple(nums)
        card = {"switch": pm, "cv_switch": cv, "nums": list(nums), "intf": intf_name(nums),
                "paths": [], "plan": [], "snags": [], "events": [], "checks": [], "equipment": None}
        modular = False
        if self.cabling:
            modular = self.cabling.is_modular(pm)
            # breakout lane (Ethernet5/2) and modular lane (Ethernet3/1/1) share the cable
            base = nums[:2] if modular else nums[:1]
            for c, s in self.cabling.switch_port_ends(pm, base):
                card["paths"].append(self.cabling.trace(c, s))
        card["modular"] = modular
        # IP list: exact numbers first, then without the lane
        keys = [nums, nums[:2], nums[:1]] if not modular else [nums, nums[:2]]
        if not modular:
            keys.append((1, nums[0]))   # core plans write fixed ports as module 1 / port N
        for k in keys:
            hits = self.plan_by.get((norm_name(pm), tuple(k)), [])
            if hits:
                card["plan"] = hits
                break
        # snag history: exact interface, plus the same base port (lanes of one QSFP)
        for (sw, n), items in self.snags_by.items():
            if sw == norm_name(pm) and (n == nums or n[:len(base_len(nums, modular))] == base_len(nums, modular)):
                card["snags"].extend(items)
        card["snags"].sort(key=lambda s: (s.get("first") or "", s.get("nr") or ""))
        if cv_events:
            card["events"] = [e for e in cv_events
                              if (e.get("device") or "").upper() == (cv or "").upper()
                              and parse_intf(e.get("interface")) == nums]
        if self.equipment:
            card["equipment"] = self.equipment.get(pm)
        card["cv_ports"] = self.cv_ports_for(pm, nums, modular)
        if nav:           # only for the card the user opens - not for the bulk lists
            card["nav"] = self.neighbours(pm, nums)
        card["compare"] = self._compare_port(card)
        card["shut"] = self.shut_for(pm, nums) if with_shut else None
        card["checks"] = self._checks(card)
        return card

    def _compare_port(self, card):
        """What each source says is connected to this switch port: [{source, device, port, detail, verdict}].
        verdict: ok | device (different device) | port (same device, other port) | '' (nothing to compare)."""
        rows = []
        for p in card["paths"]:
            ends = [h for h in p["hops"] if h["type"] == "end" and not h.get("missing")]
            if len(ends) > 1:
                far = ends[-1]
                port = " / ".join(x for x in (far.get("parent") if far.get("parent") not in ("", far.get("top")) else "",
                                              far.get("label") if far.get("label") != far.get("top") else "", far.get("port")) if x)
                rows.append({"source": "Patchmanager cabling", "device": far.get("top") or far.get("name"), "port": port,
                             "detail": "%s, %d cables" % (p["status"], sum(1 for h in p["hops"] if h["type"] == "cable")),
                             "open": p["status"] != "complete",
                             "assumed": any(h.get("assumed") for h in p["hops"])})
        for p in card["plan"]:
            if not p["free"]:
                rows.append({"source": "IP list", "device": p.get("_dev") or p["equipment"], "port": p.get("_devport") or p["eqport"],
                             "detail": "%s r%s, VLAN %s" % (p["sheet"], p["row"], p["vlan"])})
        for lane in card.get("cv_ports", [])[:1]:
            desc = lane["description"].strip(" *")
            if desc:
                dev, dport = self.device_split(desc)
                rows.append({"source": "Switch description (CloudVision)", "device": dev, "port": dport,
                             "detail": "%s %s" % (lane["intf"], lane["status"])})
        seen = set()
        for s in sorted(card["snags"], key=lambda x: x.get("first") or "", reverse=True):
            if s.get("_dev") and (s["_dev"], s.get("_devport")) not in seen and tuple(s["nums"]) == tuple(card["nums"]):
                seen.add((s["_dev"], s.get("_devport")))
                rows.append({"source": "Snagging list", "device": s["_dev"], "port": s.get("_devport", ""),
                             "detail": "#%s %s" % (s.get("nr", ""), s.get("status") or "open")})
        # verdict against the Patchmanager cabling (the reference); without cabling against the plan.
        # Two cables on one port: the one the other sources agree with is the reference.
        pm_rows = [r for r in rows if r["source"] == "Patchmanager cabling"]
        others = [r for r in rows if r["source"] != "Patchmanager cabling"]
        ref = next((p for p in pm_rows if any(self.device_relation(o["device"], p["device"]) != "different"
                                              for o in others)), None) or \
            (pm_rows[0] if pm_rows else None) or next((r for r in rows if r["source"] == "IP list"), None)
        for r in rows:
            rel = self.device_relation(r["device"], ref["device"]) if ref and r is not ref else "same"
            prel = port_relation(r["port"], ref["port"]) if ref and r is not ref else "same"
            r["naming"] = ""
            if r is ref or not ref:
                r["verdict"] = "ref" if r is ref else ""
            elif not norm_name(r["device"]) or str(r["device"]).startswith("."):
                r["verdict"] = "incomplete"  # plan row without device name ('.ETH.1 -- CS9/Ceiling')
            elif r["source"].startswith("Switch description") and re.fullmatch(r"[A-Za-z0-9\-]+\d", r["device"]) \
                    and not self.pm_name(r["device"]):
                r["verdict"] = "unmapped"  # a CloudVision hostname without Patchmanager name (Switches tab)
            elif ref.get("open") and rel == "different":
                r["verdict"] = "unknown"   # cabling stops at a panel: the device behind it is not documented
            elif ref.get("assumed") and rel == "different":
                r["verdict"] = "trunk"     # differs, but the cabling crossed a trunk with an assumed port
            elif r["source"] == "Patchmanager cabling" and rel == "different":
                r["verdict"] = "double"    # a second cable on the same switch port
            elif rel == "different":
                r["verdict"] = "device"
            elif r["port"] and ref["port"] and prel == "different":
                r["verdict"] = "port"
            elif rel == "naming" or prel == "naming":
                r["verdict"] = "naming"    # same device / port, written differently
                r["naming"] = "; ".join(x for x in (
                    "%s vs %s" % (r["device"], ref["device"]) if rel == "naming" else "",
                    "port %s vs %s" % (r["port"], ref["port"]) if prel == "naming" else "") if x)
            else:
                r["verdict"] = "ok"
        return rows

    def device_relation(self, a, b):
        """'same' | 'naming' (same device, spelt differently) | 'different'.
        naming: leading zeros / spaces / dots (DCB.WS 2 = DCB.WS 02), room prefix left out
        (WS 202 = DCB.WS 202), CloudVision hostname vs Patchmanager name, or a one-letter typo
        with all numbers equal (OC.PRIN 01 = OC.PRINT 01). A different number or room stays different."""
        if norm_name(a) == norm_name(b):
            return "same"
        if not norm_name(a) or not norm_name(b):
            return "different"
        if norm_digits(a) == norm_digits(b):
            return "naming"
        pa, pb = self.pm_name(a), self.pm_name(b)
        if (pa and same_device(pa, b)) or (pb and same_device(pb, a)) or (pa and pb and same_device(pa, pb)):
            return "naming"
        sa, sb = str(a), str(b)
        for x, y in ((sa, sb), (sb, sa)):     # room prefix missing on one side
            if "." in x and x.count(".") > y.count(".") and norm_digits(x.split(".", 1)[1]) == norm_digits(y):
                return "naming"
        if likely_typo(sa, sb):
            return "naming"
        return "different"

    def _checks(self, card):
        """Plain-language findings: documentation gaps and mismatches."""
        out = []
        if card.get("shut"):
            sh = card["shut"]
            out.append(("bad", "%s is marked FUTURE USE (%s) - this port should be shut down until it goes live." % (
                sh["device"], sh["source"])))
        if self.cabling is not None and not card["paths"]:
            out.append(("warn", "No cable documented in Patchmanager on %s %s." % (card["switch"], card["intf"])))
        far_ends = []
        for p in card["paths"]:
            ends = [h for h in p["hops"] if h["type"] == "end" and not h.get("missing")]
            far = ends[-1] if len(ends) > 1 else None
            if far:
                far_ends.append(far)
            if p["status"] == "open":
                out.append(("warn", "Path ends at %s (a panel / outlet) - nothing patched beyond it in Patchmanager." % (far["name"] if far else "?")))
            elif p["status"] == "ambiguous":
                out.append(("warn", "Path cannot be followed unambiguously (see note in the path)."))
            if any(h.get("assumed") for h in p["hops"]):
                out.append(("info", "A trunk was crossed: the port number on the far panel is assumed to be the same."))
            # one line, not one per cable: most cables are still "not installed" in Patchmanager
            flagged = [h for h in p["hops"] if h["type"] == "cable" and h.get("flags")]
            if any("to be checked" in h["flags"] for h in flagged):
                out.append(("warn", "Patchmanager: cable %s marked 'to be checked'." % ", ".join(
                    h["id"] for h in flagged if "to be checked" in h["flags"])))
            if flagged:
                out.append(("info", "%d of %d cables on this path are not ticked as installed / labelled in Patchmanager." % (
                    len(flagged), sum(1 for h in p["hops"] if h["type"] == "cable"))))
        live = [p for p in card["plan"] if not p["free"]]
        cmp_rows = card.get("compare", [])
        ref = next((r for r in cmp_rows if r.get("verdict") == "ref"), None)
        for r in cmp_rows:
            if r["verdict"] == "unmapped":
                out.append(("info", "The switch description names CloudVision host %s, which has no Patchmanager name yet%s - add it under Switches." % (
                    r["device"], " (probably %s)" % ref["device"] if ref and self.is_switch(ref["device"]) else "")))
        if any(r["verdict"] == "trunk" for r in cmp_rows):
            out.append(("warn", "IP list names another device, but the Patchmanager path crosses a trunk where the port "
                                "number is assumed - check the trunk on site before trusting either."))
        if any(r["verdict"] == "incomplete" for r in cmp_rows):
            out.append(("info", "The IP list row has no device name."))
        if any(r["verdict"] == "unknown" for r in cmp_rows):
            out.append(("warn", "Cannot compare with the IP list: the Patchmanager path stops at %s." % ref["device"]))
        for r in cmp_rows:
            if r["verdict"] == "device":
                out.append(("bad", "%s says %s, %s says %s." % (r["source"], r["device"] + (" " + r["port"] if r["port"] else ""),
                                                             ref["source"], ref["device"] + (" " + ref["port"] if ref["port"] else ""))))
            elif r["verdict"] == "port":
                out.append(("warn", "%s: same device %s, but port %s instead of %s." % (r["source"], r["device"], r["port"], ref["port"])))
        if ref and len(cmp_rows) > 1 and all(r["verdict"] in ("ok", "ref", "naming") for r in cmp_rows):
            out.append(("ok", "All sources agree: %s%s." % (ref["device"], " " + ref["port"] if ref["port"] else "")))
        for r in cmp_rows:
            if r["verdict"] == "naming":
                out.append(("info", "Naming only, technically the same: %s writes %s." % (r["source"], r["naming"])))
            elif r["verdict"] == "double":
                out.append(("bad", "Two cables documented on this port in Patchmanager: %s and %s." % (ref["device"], r["device"])))
        if card["plan"] and not live and far_ends:
            out.append(("warn", "IP list lists this port as free, but Patchmanager has a cable on it."))
        eq = card.get("equipment")
        if eq:
            p = str(card["nums"][1] if card.get("modular") and len(card["nums"]) > 1 else card["nums"][0])
            if p in eq["defect"]:
                out.append(("bad", "Patchmanager marks port %s of %s as DEFECT." % (p, eq["label"])))
            if p in eq["reserved"]:
                out.append(("info", "Port %s is reserved in Patchmanager." % p))
        if not card["cv_switch"]:
            out.append(("warn", "No CloudVision hostname known for %s (add it under Switches)." % card["switch"]))
        lanes = card.get("cv_ports") or []
        if lanes:
            up = [l for l in lanes if l["status"] == "Connected"]
            documented = bool(card["paths"]) or bool(live)
            fully = bool(far_ends and any(p["status"] == "complete" for p in card["paths"])) or bool(live)
            states = ", ".join("%s %s" % (l["intf"], l["status"]) for l in lanes)
            if fully and not up:
                out.append(("bad", "Documented as connected, but the link is DOWN in CloudVision (%s)." % states))
            elif up and not documented:
                desc = lanes[0]["description"]
                out.append(("warn", "Link is UP in CloudVision, but nothing is documented on this port (%s)%s." % (
                    states, " - only the switch description says %s" % desc if desc else "")))
            elif up and card.get("shut"):
                out.append(("bad", "Marked future use, but the link is UP in CloudVision."))
            elif up:
                l = up[0]
                out.append(("ok", "Link up in CloudVision: %s %s %s (last change %s)." % (l["intf"], l["speed"], l["optic"], l["changed"])))
        if any(s["open"] for s in card["snags"]):
            out.append(("warn", "Open entry in the snagging list."))
        return out

    # -- search ---------------------------------------------------------------------
    def search(self, q, limit=60):
        """Free text -> list of results {kind: port | device | cable, title, sub, ...}.
        Port:   'Ethernet12 on dcbeswcpredg001', 'DCB.ESW A01 12', 'DCB.ESW A01.12', 'swhost Et5/2'
        Cable:  'F03639'
        Device: any name it is known by (label, alias, asset tag, description, IP list,
                snag destination, CloudVision hostname), also partial ('AMU 01', 'vs 2')."""
        q = (q or "").strip()
        if not q:
            return []
        res = []
        m = re.match(r"^\s*(et\S*)\s+on\s+(\S+)", q, re.I)
        if m:
            sw, intf = m.group(2), m.group(1)
        else:
            m = re.match(r"^(.*?)[\s.:]+((?:et(?:hernet)?)?\s*\d+(?:/\d+){0,2})\s*$", q, re.I)
            sw, intf = (m.group(1), m.group(2)) if m else (None, None)
        nums = parse_intf(intf) if intf else None
        pm = self.pm_name(sw) if sw else None
        if pm and nums and self.is_switch(pm):
            res.append({"kind": "port", "title": "%s %s" % (pm, intf_name(nums)), "sub": self.cv_name(pm) or "",
                        "switch": pm, "nums": list(nums)})
        if self.cabling:
            for c in self.cabling.by_id.get(q.upper(), []):
                res.append({"kind": "cable", "title": "Cable %s" % c.id,
                            "sub": "%s  <->  %s" % (c.a.name(), c.b.name()), "cable": c.idx})
        for dev, name, via in self.find_devices(q, limit=limit):
            e = self.equipment.get(dev) if self.equipment else None
            sub = " · ".join(x for x in ((e or {}).get("template", ""), (e or {}).get("location", "")) if x)
            if norm_name(name) != norm_name(dev):
                sub = ("%s: %s" % (via, name)) + (" · " + sub if sub else "")
            res.append({"kind": "switch" if self.is_switch(dev) else "device", "title": dev, "sub": sub, "device": dev})
        return res[:limit]

    def cable_card(self, idx, side=None):
        c = self.cabling.cables[idx]
        out = {"cable": c.to_dict(), "a": c.a.to_dict(), "b": c.b.to_dict(), "paths": []}
        for s in ([side] if side in ("a", "b") else ["a", "b"]):
            out["paths"].append(self.cabling.trace(c, s))
        return out

    def worklist(self, cv_events=None):
        """Open snagging entries + CloudVision interface events not already in the list."""
        items = []
        for s in self.snags:
            if not s["open"]:
                continue
            pm = self.pm_name(s.get("pm_switch") or s.get("cv_switch")) or s.get("pm_switch") or s.get("cv_switch")
            items.append({"source": "snag", "switch": pm, "cv_switch": s.get("cv_switch") or self.cv_name(pm) or "",
                          "nums": s["nums"], "intf": intf_name(s["nums"]) if s["nums"] else s.get("cv_port", ""),
                          "what": s.get("error") or s.get("error_type") or s.get("comment") or "",
                          "status": s.get("status") or "open", "since": s.get("first", ""),
                          "who": s.get("who", ""), "jira": s.get("jira", ""), "nr": s.get("nr", ""),
                          "destination": s.get("destination", "")})
        for i in items:
            sh = self.shut_for(i["switch"] or "", i["nums"]) if i["nums"] else None
            i["future"] = sh["device"] if sh else ""
        known = {(norm_name(i["switch"] or ""), tuple(i["nums"])) for i in items}
        for e in cv_events or []:
            nums = parse_intf(e.get("interface"))
            pm = self.pm_name(e.get("device")) or e.get("device")
            if not nums or (norm_name(pm), nums) in known:
                continue
            items.append({"source": "cloudvision", "switch": pm, "cv_switch": e.get("device", ""),
                          "nums": list(nums), "intf": intf_name(nums), "what": e.get("title", ""),
                          "status": "CloudVision " + (e.get("severity") or "").lower(), "since": e.get("time", ""),
                          "who": "", "jira": "", "nr": "", "destination": "",
                          "future": (self.shut_for(pm, nums) or {}).get("device", "")})
        return items

    # -- future use ------------------------------------------------------------------
    def set_future(self, future):
        self.future = {k: v for k, v in (future or {}).items() if k}
        self._shut_cache = None

    def future_devices(self):
        """{device: {"source", "note", "since"}} - marked in this tool, or ticked
        'Future Expansion' on the device in Patchmanager."""
        out = {}
        if self.equipment:
            for items in self.equipment.by_label.values():
                for e in items:
                    if e["label"] == e["top"] and e["attrs"].get("Future Expansion") == "true":
                        out[e["label"]] = {"source": "Patchmanager 'Future Expansion'", "note": "", "since": ""}
        for dev, v in self.future.items():
            out[dev] = {"source": "marked here", "note": v.get("note", ""), "since": v.get("since", "")}
        return out

    def future_of(self, device):
        if not device:
            return None
        fut = self.future_devices()
        for d, v in fut.items():
            if same_device(d, device):
                return dict(v, device=d)
        return None

    def ports_to_shut(self):
        """Every switch port that leads to a future device (Patchmanager cabling or IP list):
        [{switch, cv_switch, nums, intf, device, device_port, via, source, note}]"""
        rows, seen = [], set()
        for dev, info in sorted(self.future_devices().items()):
            card = self.device_card(dev)
            cands = [(c["switch_link"], c["port"], "Patchmanager cabling") for c in card["connections"] if c["switch_link"]]
            cands += [({"switch": p["switch"], "nums": p["nums"]}, p.get("_devport") or p["eqport"], "IP list")
                      for p in card["plan"]]
            for link, dport, via in cands:
                pm = self.pm_name(link["switch"]) or link["switch"]
                nums = list(link["nums"])
                if len(nums) == 2 and nums[0] == 1 and not (self.cabling and self.cabling.is_modular(pm)):
                    nums = nums[1:]   # core plan '1/25' = port 25
                key = (norm_name(pm), tuple(nums))
                if key in seen:
                    continue
                seen.add(key)
                # another source naming a different device on this port = do NOT shut blindly
                pc = self.port_card(pm, nums, with_shut=False)
                others = sorted({"%s: %s" % (r["source"], r["device"]) for r in pc["compare"]
                                 if r["verdict"] != "incomplete" and norm_name(r["device"])
                                 and not str(r["device"]).startswith(".") and self.device_relation(r["device"], dev) == "different"
                                 and not self.future_of(r["device"])})
                rows.append({"switch": pm, "cv_switch": self.cv_name(pm) or "", "nums": nums, "intf": intf_name(nums),
                             "device": dev, "device_port": dport, "via": via, "source": info["source"], "note": info["note"],
                             "conflict": "; ".join(others)})
        rows.sort(key=lambda r: (norm_name(r["switch"]), r["nums"]))
        return rows

    def shut_keys(self):
        if getattr(self, "_shut_cache", None) is None:
            self._shut_cache = {(norm_name(r["switch"]), tuple(r["nums"])): r for r in self.ports_to_shut()}
        return self._shut_cache

    def shut_for(self, switch, nums):
        """The ports-to-shut row for a switch port (lanes count as the port), or None."""
        keys = self.shut_keys()
        n = tuple(nums)
        sw = norm_name(self.pm_name(switch) or switch)
        return keys.get((sw, n)) or keys.get((sw, n[:2])) or keys.get((sw, n[:1]))

    # -- signal chain --------------------------------------------------------------------
    def device_links(self, dev):
        """Every documented link of a device: its port -> through panels -> the next device.
        [{from_port, to, to_port, to_end, path, cables}] (to = '' when the path dead-ends)."""
        out = []
        if not self.cabling:
            return out
        for c, s in sorted(self.cabling.by_top.get(norm_name(dev), []), key=lambda cs: port_sort(getattr(cs[0], cs[1]))):
            me = getattr(c, s)
            path = self.cabling.trace(c, s)
            ends = [h for h in path["hops"] if h["type"] == "end" and not h.get("missing")]
            far = ends[-1] if len(ends) > 1 else None
            to = far.get("top", "") if far and not is_passive(far) else ""
            out.append({"from_port": end_port(me.to_dict()), "to": to, "to_port": end_port(far) if far else "",
                        "to_end": far, "path": path,
                        "cables": [h for h in path["hops"] if h["type"] == "cable"]})
        return out

    HUB_LINKS = 16   # a device with more cabled ports than this ends a branch (KVM, frames, routers)

    def signal_chain(self, start, cv_events=None, max_depth=6, max_links=150):
        """Follow every cable from a device through panels AND through devices, breadth first,
        until the chain ends. Switches end a branch (going through them would pull in the whole
        network) - their port is a leaf with its port card. Every link gets its faults.
        Returns {root, links: [... in display order ...], faults: [...], truncated}"""
        root = self.device_of(start) or start
        links, seen_dev, seen_cable = [], {norm_name(root)}, set()
        truncated = False

        def walk(dev, depth, parent_idx):
            nonlocal truncated
            if depth > max_depth:
                truncated = True
                return
            for l in self.device_links(dev):
                ids = tuple(sorted(h.get("id", "") + str(h.get("date", "")) for h in l["cables"]))
                if ids in seen_cable:
                    continue          # the link back to where we came from
                seen_cable.add(ids)
                if len(links) >= max_links:
                    truncated = True
                    return
                to = l["to"]
                sw = bool(to) and self.is_switch(to)
                again = bool(to) and norm_name(to) in seen_dev
                hub = bool(to) and not sw and not again and \
                    len(self.cabling.by_top.get(norm_name(to), [])) > self.HUB_LINKS
                link = {"depth": depth, "parent": parent_idx, "from": dev, "from_port": l["from_port"], "to": to,
                        "to_port": l["to_port"], "to_name": (l["to_end"] or {}).get("name", ""), "status": l["path"]["status"],
                        "cables": [{"id": h.get("id", ""), "template": h.get("template", ""), "length": h.get("length", ""),
                                    "flags": h.get("flags", []), "trunk": bool(h.get("trunk"))} for h in l["cables"]],
                        "hops": l["path"]["hops"], "is_switch": sw, "again": again, "hub": hub,
                        "hub_links": len(self.cabling.by_top.get(norm_name(to), [])) if hub else 0,
                        "switch_link": self._switch_ref(l["to_end"]) if sw else None,
                        "future": bool(to) and bool(self.future_of(to))}
                link["faults"] = self._link_faults(link, cv_events)
                links.append(link)
                idx = len(links) - 1
                if to and not sw and not again and not hub:
                    seen_dev.add(norm_name(to))
                    walk(to, depth + 1, idx)
                elif hub:
                    seen_dev.add(norm_name(to))

        if self.is_switch(root):
            # a switch as start: one level only (its ports)
            for l in self.device_links(root):
                link = {"depth": 1, "parent": None, "from": root, "from_port": l["from_port"], "to": l["to"],
                        "to_port": l["to_port"], "to_name": (l["to_end"] or {}).get("name", ""), "status": l["path"]["status"],
                        "cables": [{"id": h.get("id", ""), "template": h.get("template", ""), "length": h.get("length", ""),
                                    "flags": h.get("flags", []), "trunk": bool(h.get("trunk"))} for h in l["cables"]],
                        "hops": l["path"]["hops"], "is_switch": False, "again": False, "hub": False, "hub_links": 0,
                        "switch_link": self._switch_ref({"top": root, "parent": "", "port": l["from_port"]}),
                        "future": bool(l["to"]) and bool(self.future_of(l["to"]))}
                link["faults"] = self._link_faults(link, cv_events)
                links.append(link)
        else:
            walk(root, 1, None)
        faults = []
        for i, l in enumerate(links):
            for sev, text in l["faults"]:
                faults.append({"link": i, "severity": sev, "fault": text, "from": l["from"], "from_port": l["from_port"],
                               "to": l["to"] or l["to_name"], "to_port": l["to_port"], "depth": l["depth"],
                               "cables": " + ".join(c["id"] or "?" for c in l["cables"])})
        order = {"error": 0, "warning": 1, "info": 2}
        faults.sort(key=lambda f: (order.get(f["severity"], 3), f["depth"], f["link"]))
        return {"root": root, "links": links, "faults": faults, "truncated": truncated,
                "devices": len({norm_name(l["to"]) for l in links if l["to"]}) + 1}

    def _link_faults(self, link, cv_events=None):
        """[(severity, text)] for one link of the chain."""
        out = []
        for c in link["cables"]:
            if "to be checked" in c["flags"]:
                out.append(("warning", "Cable %s is marked 'to be checked' in Patchmanager" % (c["id"] or "?")))
        if link["status"] == "open":
            out.append(("warning", "Patchmanager path ends at %s - nothing patched beyond it" % (link["to_name"] or "?")))
        elif link["status"] in ("ambiguous", "loop"):
            out.append(("warning", "Patchmanager path is %s at %s" % (link["status"], link["to_name"] or "?")))
        if any(c["trunk"] for c in link["cables"]):
            out.append(("info", "Path crosses a trunk - far port number assumed"))
        # defect ports on either device (Patchmanager)
        if self.equipment:
            for dev, port in ((link["from"], link["from_port"]), (link["to"], link["to_port"])):
                e = self.equipment.get(dev) if dev else None
                last = norm_port(str(port).split("/")[-1]) if port else ""
                if e and last and last in e["defect"]:
                    out.append(("error", "Port %s of %s is marked DEFECT in Patchmanager" % (port, dev)))
        # the switch end: snags, CloudVision, documentation
        ref = link.get("switch_link")
        if ref:
            pc = self.port_card(ref["switch"], ref["nums"], cv_events, with_shut=False)
            for snag in pc["snags"]:
                if snag["open"]:
                    out.append(("error", "Open snag #%s on %s %s: %s (%s)" % (
                        snag.get("nr", ""), pc["switch"], pc["intf"],
                        snag.get("error") or snag.get("error_type") or snag.get("comment") or snag.get("destination") or "no description",
                        snag.get("status") or "open")))
            for e in pc["events"]:
                if e.get("open"):
                    out.append(("error", "CloudVision on %s %s: %s" % (pc["switch"], pc["intf"], e.get("title"))))
            for r in pc["compare"]:
                if r["verdict"] == "device":
                    out.append(("error", "Documentation mismatch on %s %s: %s says %s" % (pc["switch"], pc["intf"], r["source"], r["device"])))
                elif r["verdict"] == "port":
                    out.append(("warning", "Documentation mismatch on %s %s: %s says port %s" % (pc["switch"], pc["intf"], r["source"], r["port"])))
            if len([r for r in pc["compare"] if r["source"] == "Patchmanager cabling"]) > 1:
                out.append(("error", "Two cables documented on %s %s" % (pc["switch"], pc["intf"])))
        if link.get("future"):
            out.append(("info", "%s is marked future use" % link["to"]))
        return out

    # -- overviews ------------------------------------------------------------------------
    def ports_overview(self, cv_events=None):
        """One row per switch port (and per cable if a port has two), plus IP-list ports
        without a documented cable: switch, port, device, device port, patch panels on the
        path, cables, what the IP list says, status."""
        rows, covered = [], set()
        names = getattr(self, "switch_names", {})
        for nsw in sorted(getattr(self, "switch_set", set())):
            sw = names.get(nsw) or self.pm_name(nsw) or nsw
            ports = self.cabling.ports_of(sw) if self.cabling else []
            for nums, _ in ports:
                card = self.port_card(sw, nums, cv_events)
                covered.add((nsw, tuple(nums)))
                rows += self._overview_rows(card)
        for p in self.plan:
            if p["free"]:
                continue
            nsw = norm_name(p["switch"])
            nums = tuple(p["nums"])
            if len(nums) == 2 and nums[0] == 1 and not (self.cabling and self.cabling.is_modular(p["switch"])):
                nums = nums[1:]
            if (nsw, nums) in covered:
                continue
            covered.add((nsw, nums))
            rows += self._overview_rows(self.port_card(p["switch"], nums, cv_events))
        # ports only CloudVision knows (link up or a description, nothing documented)
        for host, ports in self.cvports.items():
            pm = self.pm_name(host) or host
            nsw = norm_name(pm)
            modular = bool(self.cabling) and self.cabling.is_modular(pm)
            done = set()
            for k, l in sorted(ports.items()):
                base_nums = k[:2] if modular else k[:1]
                if (nsw, base_nums) in covered or base_nums in done:
                    continue
                if l["status"] != "Connected" and not l["description"]:
                    continue
                done.add(base_nums)
                rows += self._overview_rows(self.port_card(pm, base_nums, cv_events))
        return rows

    def _overview_rows(self, card):
        live = [p for p in card["plan"] if not p["free"]]
        plan = ("%s %s" % (live[0].get("_dev") or live[0]["equipment"], live[0].get("_devport") or live[0]["eqport"])).strip() if live else ""
        verdicts = {r["verdict"] for r in card["compare"]}
        status = ("mismatch" if "device" in verdicts else "two cables" if "double" in verdicts else
                  "check" if "trunk" in verdicts else "port differs" if "port" in verdicts else
                  "naming only" if "naming" in verdicts else "")
        open_snags = sum(1 for s in card["snags"] if s["open"])
        lanes = card.get("cv_ports") or []
        up = sum(1 for l in lanes if l["status"] == "Connected")
        link = "" if not lanes else lanes[0]["status"] if len(lanes) == 1 else "%d/%d up" % (up, len(lanes))
        base = {"switch": card["switch"], "cv_switch": card["cv_switch"], "nums": card["nums"], "intf": card["intf"],
                "link": link, "link_up": up > 0, "cv_desc": lanes[0]["description"] if lanes else "",
                "speed": lanes[0]["speed"] if lanes else "", "cv_time": lanes[0]["exported"] if lanes else "",
                "cv_intf": lanes[0]["intf"] if lanes else "",
                "iplist": plan, "snags": open_snags, "events": len(card["events"]),
                "future": (card.get("shut") or {}).get("device", "")}
        if not card["paths"]:
            return [dict(base, device="", device_port="", panels=[], cables=[], length="",
                         status=status or ("no cable" if self.cabling else ""))]
        out = []
        for p in card["paths"]:
            ends = [h for h in p["hops"] if h["type"] == "end" and not h.get("missing")]
            far = ends[-1] if len(ends) > 1 else None
            panels = []
            for h in ends[1:]:
                if is_passive(h) and h.get("top") and h.get("top") not in panels:
                    panels.append(h["top"])
            cables = [h for h in p["hops"] if h["type"] == "cable"]
            total = 0.0
            for h in cables:
                m = re.match(r"([\d.]+)\s*m", h.get("length") or "")
                total += float(m.group(1)) if m else 0
            dev = far.get("top", "") if far and not is_passive(far) else ""
            st = status or ("ends at panel" if p["status"] == "open" else p["status"] if p["status"] != "complete" else "ok")
            out.append(dict(base, device=dev, device_port=end_port(far) if dev else "", panels=panels,
                            cables=[h.get("id") or "?" for h in cables], length=("%.1f m" % total) if total else "",
                            status=st, end=(far or {}).get("name", "") if not dev else ""))
        if len(out) > 1:
            for r in out:
                r["status"] = "two cables"
        return out

    def devices_overview(self):
        """Every device, switch and patch panel: name, type, template, location, connections."""
        seen = {}
        if self.equipment:
            for items in self.equipment.by_label.values():
                for e in items:
                    if e["label"] == e["top"]:
                        seen.setdefault(norm_name(e["label"]), {"device": e["label"], "template": e["template"],
                                                                "location": e["location"], "position": e["position"],
                                                                "alias": e["alias"]})
        if self.cabling:
            for key, ends in self.cabling.by_top.items():
                c, sd = ends[0]
                end = getattr(c, sd)
                d = seen.setdefault(key, {"device": end.top, "template": end.template, "location": " > ".join(
                    x for x in (end.location, end.cabinet) if x), "position": "", "alias": ""})
                d["links"] = len(ends)
        fut = {norm_name(k) for k in self.future_devices()}
        out = []
        for key, d in seen.items():
            if not d["device"]:
                continue
            kind = "switch" if self.is_switch(d["device"]) else \
                "patch panel" if is_passive({"top": d["device"], "template": d["template"]}) else "device"
            out.append(dict(d, kind=kind, links=d.get("links", 0), future=key in fut))
        out.sort(key=lambda d: norm_name(d["device"]))
        return out

    def mismatches(self):
        """Every switch port (IP list + documented cabling + snags) where the sources disagree.
        Rows: {switch, cv_switch, intf, nums, severity, kind, source, says, reference, ref_says, detail}"""
        ports = set()
        for p in self.plan:
            if not p["free"]:
                ports.add((p["switch"], tuple(p["nums"])))
        for (sw, nums) in self.snags_by:
            pm = self.pm_name(sw)
            if pm:
                ports.add((pm, nums))
        if self.cabling:
            for sw in list(getattr(self, "switch_set", set())):
                pm = self.pm_name(sw) or sw
                for nums, _ in self.cabling.ports_of(pm):
                    ports.add((pm, nums))
        for host, cvp in self.cvports.items():
            pm = self.pm_name(host) or host
            modular = bool(self.cabling) and self.cabling.is_modular(pm)
            for k, l in cvp.items():
                if l["status"] == "Connected" or l["description"]:
                    ports.add((pm, k[:2] if modular else k[:1]))
        out, seen = [], set()
        for sw, nums in sorted(ports, key=lambda x: (norm_name(x[0]), x[1])):
            card = self.port_card(sw, nums)
            key = (norm_name(card["switch"]), tuple(card["nums"][:2] if card.get("modular") else card["nums"][:1]))
            if key in seen and len(nums) > 1 and not card.get("modular"):
                continue
            seen.add(key)
            ref = next((r for r in card["compare"] if r["verdict"] == "ref"), None)
            pm_paths = [r for r in card["compare"] if r["source"] == "Patchmanager cabling"]
            for r in card["compare"]:
                if r["verdict"] in ("ok", "ref", ""):
                    continue
                sev = {"device": "error", "port": "warning", "trunk": "check", "unknown": "info", "incomplete": "info",
                       "unmapped": "info", "naming": "naming", "double": "error"}[r["verdict"]]
                kind = {"device": "different device", "port": "different port", "trunk": "different device (trunk assumed)",
                        "unknown": "cabling ends at panel", "incomplete": "IP list without device name",
                        "unmapped": "CloudVision hostname not mapped", "naming": "naming only (same device)",
                        "double": "two cables on one port"}[r["verdict"]]
                out.append({"switch": card["switch"], "cv_switch": card["cv_switch"], "intf": card["intf"],
                            "nums": card["nums"], "severity": sev, "kind": kind, "source": r["source"],
                            "says": (r["device"] + " " + (r["port"] or "")).strip(), "reference": ref["source"] if ref else "",
                            "ref_says": (ref["device"] + " " + (ref["port"] or "")).strip() if ref else "",
                            "detail": r.get("detail", "")})
            for lvl, text in card["checks"]:
                if "in CloudVision" in text and lvl in ("bad", "warn"):
                    out.append({"switch": card["switch"], "cv_switch": card["cv_switch"], "intf": card["intf"],
                                "nums": card["nums"], "severity": "error" if lvl == "bad" else "warning",
                                "kind": "documented, link down" if "DOWN" in text else "link up, not documented"
                                if "nothing is documented" in text else "future use, link up",
                                "source": "CloudVision export", "says": text, "reference": "", "ref_says": "", "detail": ""})
            if self.cabling is not None and not pm_paths and any(not p["free"] for p in card["plan"]):
                live = [p for p in card["plan"] if not p["free"]][0]
                out.append({"switch": card["switch"], "cv_switch": card["cv_switch"], "intf": card["intf"], "nums": card["nums"],
                            "severity": "warning", "kind": "in IP list, no cable in Patchmanager", "source": "IP list",
                            "says": live["device"], "reference": "Patchmanager cabling", "ref_says": "(nothing)",
                            "detail": "%s r%s" % (live["sheet"], live["row"])})
        for sm in self.neuron_summary():
            out.append({"switch": sm["device"], "cv_switch": "", "intf": "(whole %s)" % ("card" if sm["kind"] == "ipg" else "panel"),
                        "nums": [], "device": sm["device"], "severity": "check", "kind": "Neuron list vs Patchmanager - numbering shift",
                        "source": "Neuron signal list", "says": sm["text"], "reference": "Patchmanager", "ref_says": "", "detail": ""})
        for r in self.neuron:
            for sev, text in self.neuron_check(r):
                out.append({"switch": r["device"], "cv_switch": "", "intf": "card %s BNC %s" % (r["card"], r["bnc"]),
                            "nums": [], "device": r["device"], "severity": sev, "kind": "Neuron list vs Patchmanager",
                            "source": "Neuron signal list", "says": text, "reference": "Patchmanager", "ref_says": "",
                            "detail": "%s r%s" % (r["sheet"], r["row"])})
        return self._explain(out)

    # -- plain-language interpretation of a mismatch ------------------------------------------
    BUCKETS = {   # kind -> (bucket, order)
        "different device": ("real", 1), "two cables on one port": ("real", 2), "different port": ("real", 3),
        "documented, link down": ("real", 4), "future use, link up": ("real", 5),
        "different device (trunk assumed)": ("real", 6), "Neuron list vs Patchmanager": ("real", 7),
        "Neuron list vs Patchmanager - numbering shift": ("real", 8),
        "link up, not documented": ("incomplete", 1), "in IP list, no cable in Patchmanager": ("incomplete", 2),
        "cabling ends at panel": ("incomplete", 3), "IP list without device name": ("incomplete", 4),
        "CloudVision hostname not mapped": ("incomplete", 5), "naming only (same device)": ("naming", 1),
    }

    def _explain(self, rows):
        """Adds to every mismatch: bucket (real | incomplete | naming), 'sources' (what each
        source says about this port, side by side), 'meaning' and 'action' in plain words."""
        cards = {}
        for m in rows:
            kind = m["kind"]
            bucket, order = self.BUCKETS.get(kind, ("incomplete", 9))
            if kind == "Neuron list vs Patchmanager" and m["severity"] != "error":
                bucket = "incomplete"
            m["bucket"], m["order"] = bucket, order
            # stable identity: the same finding on the same port with the same values
            m["key"] = hashlib.sha1("|".join((kind, norm_name(m["switch"]), str(m.get("intf", "")), m.get("source", ""),
                                              norm_name(m.get("says", "")), norm_name(m.get("ref_says", "")))).encode()).hexdigest()[:16]
            # what each source says about this switch port
            sources = []
            if m.get("nums"):
                key = (norm_name(m["switch"]), tuple(m["nums"]))
                if key not in cards:
                    cards[key] = self.port_card(m["switch"], m["nums"], with_shut=False)
                card = cards[key]
                for r in card["compare"]:
                    sources.append({"source": r["source"], "says": ("%s %s" % (r["device"], r["port"] or "")).strip(),
                                    "verdict": r["verdict"]})
                lanes = card.get("cv_ports") or []
                if lanes:
                    l0 = lanes[0]
                    sources.append({"source": "CloudVision link", "says": "%s %s (export %s)" % (
                        l0["status"], l0["speed"], l0["exported"]), "verdict": "ok" if l0["status"] == "Connected" else "device"})
            m["sources"] = sources
            port = "%s %s" % (m["switch"], m["intf"])
            pm = next((x["says"] for x in sources if x["source"] == "Patchmanager cabling" and x["verdict"] == "ref"), m.get("ref_says", ""))
            says, src = m.get("says", ""), m.get("source", "")
            if kind == "different device":
                m["meaning"] = "%s expects %s on %s, but the Patchmanager cable goes to %s. Either the cable is patched to the wrong device, or one document is wrong." % (src, says, port, pm or m.get("ref_says"))
                m["action"] = "Check on site what is plugged into %s, then correct the wrong document." % port
            elif kind == "two cables on one port":
                m["meaning"] = "Patchmanager documents two different cables on the one switch port %s (%s and %s) - physically only one is possible." % (port, pm, says)
                m["action"] = "Check on site which cable is really on %s and delete the other one in Patchmanager." % port
            elif kind == "different port":
                m["meaning"] = "Same device, different device port: %s says %s, Patchmanager says %s." % (src, says, pm)
                m["action"] = "Check which port of the device is really cabled to %s and correct the document." % port
            elif kind == "different device (trunk assumed)":
                m["meaning"] = "%s expects %s, but the Patchmanager path to %s crosses a fibre trunk whose strand is not documented, so the far end is a guess." % (src, says, pm)
                m["action"] = "Check the trunk / cassette on site and document the strand in Patchmanager - then the tool can compare properly."
            elif kind == "documented, link down":
                m["meaning"] = "A device is documented on %s, but CloudVision shows no link there (at the export time)." % port
                m["action"] = "Check power, cable, optics and patching on site - or update the documentation if the device is gone."
            elif kind == "future use, link up":
                m["meaning"] = "The device on %s is marked as future use, but the link is up." % port
                m["action"] = "Shut the port (see To shut) - or clear the future-use mark if the device is live."
            elif kind == "link up, not documented":
                m["meaning"] = "CloudVision shows a live link on %s, but no document says what is connected." % port
                m["action"] = "Find out what is plugged in (CloudVision LLDP / MAC) and document it in Patchmanager and the IP list."
            elif kind == "in IP list, no cable in Patchmanager":
                m["meaning"] = "The IP list plans %s on %s, but Patchmanager has no cable on this port." % (says, port)
                m["action"] = "Add the cable in Patchmanager - or remove the line from the IP list if it is no longer planned."
            elif kind == "cabling ends at panel":
                m["meaning"] = "The Patchmanager path from %s stops at a patch panel (%s); what is behind it is not documented." % (port, pm)
                m["action"] = "Complete the patching in Patchmanager from that panel to the device."
            elif kind == "IP list without device name":
                m["meaning"] = "The IP list has a line for %s, but without a device name (%s)." % (port, says)
                m["action"] = "Add the device name in the IP list."
            elif kind == "CloudVision hostname not mapped":
                m["meaning"] = "The switch description names %s, which has no Patchmanager name yet." % says
                m["action"] = "Add the hostname under Switches (Override column)."
            elif kind == "naming only (same device)":
                m["meaning"] = "Technically the same: %s writes %s, Patchmanager %s." % (src, says, pm)
                m["action"] = "Optional: align the spelling in %s so searches find both." % src
            elif kind == "Neuron list vs Patchmanager - numbering shift":
                m["meaning"] = says
                m["action"] = "Check on site which numbering is right (Neuron list or Patchmanager) and correct the other one."
            elif kind == "Neuron list vs Patchmanager":
                m["meaning"] = "The Neuron signal list and Patchmanager disagree on %s: %s." % (m["intf"], says)
                m["action"] = "Check the BNC cable on site and correct the Neuron list or Patchmanager." if m["bucket"] == "real" \
                    else "Add / complete the cable in Patchmanager."
            else:
                m["meaning"], m["action"] = says, ""
        rows.sort(key=lambda m: ({"real": 0, "incomplete": 1, "naming": 2}[m["bucket"]], m["order"], norm_name(m["switch"]), m.get("nums") or []))
        return rows

    def status(self):
        return {"meta": self.meta, "switches": len(self.switches()), "mapping": len(self.mapping),
                "mapping_dups": self.mapping_dups}


def end_port(end):
    """Readable port of a cable end: 'LC03 / 1', 'UIC 1 / MGMT', 'ETH'."""
    if not end:
        return ""
    top = end.get("top", "")
    parts = [end.get("parent") if end.get("parent") not in ("", None, top) else "",
             end.get("label") if end.get("label") not in ("", None, top) else "", end.get("port", "")]
    out = []
    for p in parts:
        if p and (not out or norm_name(out[-1]) != norm_name(p)):
            out.append(p)
    return " / ".join(out)


def eos_shutdown(rows):
    """Arista EOS config text to shut the given ports, one block per switch. Only text - this
    tool never logs in to a switch."""
    out, by = [], defaultdict(list)
    for r in rows:
        by[(r["cv_switch"] or r["switch"], r["switch"])].append(r)
    for (host, pm), items in sorted(by.items()):
        out.append("! %s  (%s)" % (host, pm))
        out.append("configure")
        for r in items:
            desc = re.sub(r"[^\w .:/\-]", "", "FUTURE %s %s" % (r["device"], r["device_port"]))[:80].strip()
            if r.get("conflict"):  # documentation disagrees: only a comment, never a shutdown
                out += ["! CHECK FIRST - not shut: %s - documentation also says %s" % (r["intf"], r["conflict"]),
                        "! interface %s" % r["intf"], "!    shutdown"]
                continue
            out += ["interface %s" % r["intf"], "   description %s" % desc, "   shutdown"]
        out += ["end", ""]
    return "\n".join(out)


def base_len(nums, modular):
    """The part of the interface that identifies the physical cage (lanes share it)."""
    return tuple(nums[:2] if modular else nums[:1])
