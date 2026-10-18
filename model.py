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
import io
import re
from collections import defaultdict
from datetime import date, datetime, timedelta

from xlsx import read_xlsx

# ---------------------------------------------------------------- helpers

def norm_port(p):
    """'01' -> '1', ' Eth ' -> 'ETH'; used to compare Patchmanager / plan port labels."""
    p = str(p or "").strip().upper()
    return str(int(p)) if p.isdigit() else p


def norm_name(s):
    """Compare names without case, spaces, dots, dashes: 'DCB.ESW A01' == 'dcb esw a01'."""
    return re.sub(r"[\s.\-_]", "", str(s or "")).upper()


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
    """Which input a file is: 'cables' | 'equipment' | 'portplan' | 'snags' | None."""
    name = (filename or "").lower()
    if name.endswith(".csv"):
        head = data[:4000].decode("utf-8-sig", "replace") if isinstance(data, bytes) else data[:4000]
        first = head.splitlines()[0] if head else ""
        if "Cable Identifier" in first and "Connected Equipment (Left)" in first:
            return "cables"
        if "Equipment Identifier" in first and "Equipment Template" in first:
            return "equipment"
        return None
    if name.endswith((".xlsx", ".xlsm")):
        sheets = read_xlsx(data, max_rows=40)
        for rows in sheets.values():
            for r in rows:
                cells = [c.strip().lower() for c in r]
                if "status manually" in cells or ("port cv" in cells and "switch cad" in cells):
                    return "snags"
        for rows in sheets.values():
            for r in rows:
                cells = [c.strip().lower() for c in r]
                if "switch" in cells and "port" in cells and "device" in cells:
                    return "portplan"
    return None


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
                if not m or int(m.group(1)) != slot:
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
            nums = (int(m.group(1)), int(p)) if m else (int(p),)
            res[nums].append((c, s))
        return sorted(res.items())


# ---------------------------------------------------------------- Patchmanager equipment

class Equipment:
    def __init__(self, rows):
        self.by_label = defaultdict(list)
        for r in rows:
            label = r.get("Equipment Label", "").strip()
            if not label:
                continue
            ident = [p.strip() for p in r.get("Equipment Identifier", "").split(",")]
            self.by_label[norm_name(label)].append({
                "label": label,
                "template": r.get("Equipment Template", "").strip(),
                "location": " > ".join(p for p in r.get("Equipment Location", "").split(",")[2:]),
                "path": ident,
                "position": r.get("Equipment Position", "").strip(),
                "description": r.get("Equipment Description", "").strip(),
                "attrs": parse_attrs(r.get("Equipment Attributes", "")),
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


# ---------------------------------------------------------------- port plan (IP list view)

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
    return out


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
        else:
            raise ValueError("unknown kind %r" % kind)
        self.meta[kind] = {"file": filename, "loaded": when, "count": n}
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
        self.plan_by = defaultdict(list)
        for p in self.plan:
            self.plan_by[(norm_name(p["switch"]), tuple(p["nums"]))].append(p)
        self.snags_by = defaultdict(list)
        for s in self.snags:
            sw = self.pm_name(s.get("pm_switch") or s.get("cv_switch"))
            if sw and s["nums"]:
                self.snags_by[(norm_name(sw), tuple(s["nums"]))].append(s)

    # -- names -------------------------------------------------------------------
    def pm_name(self, name):
        """Any switch name (PM or CV hostname, any case) -> Patchmanager name, or None."""
        if not name:
            return None
        n = name.strip()
        if n.upper() in self.cv_to_pm:
            return self.cv_to_pm[n.upper()]
        nn = norm_name(n)
        for pm in self.mapping:
            if norm_name(pm) == nn:
                return pm
        if self.cabling and nn in self.cabling.by_top:
            c, s = self.cabling.by_top[nn][0]
            return getattr(c, s).top
        for p in self.plan:
            if norm_name(p["switch"]) == nn:
                return p["switch"]
        return None

    def cv_name(self, pm):
        if not pm:
            return None
        for k, v in self.mapping.items():
            if norm_name(k) == norm_name(pm):
                return v
        return None

    def switches(self):
        """All switches we know of (PM names) with their CV hostname."""
        names = {}
        for pm in self.mapping:
            names[norm_name(pm)] = pm
        for p in self.plan:
            names.setdefault(norm_name(p["switch"]), p["switch"])
        return sorted(({"pm": pm, "cv": self.cv_name(pm) or ""} for pm in names.values()), key=lambda x: x["pm"])

    # -- the port card -----------------------------------------------------------
    def port_card(self, switch, nums, cv_events=None):
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
        # port plan: exact numbers first, then without the lane
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
        card["checks"] = self._checks(card)
        return card

    def _checks(self, card):
        """Plain-language findings: documentation gaps and mismatches."""
        out = []
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
        if live and far_ends:
            want = norm_name(live[0]["equipment"])
            got = [norm_name(f.get("top")) for f in far_ends] + [norm_name(f.get("label")) for f in far_ends]
            if want and want not in got:
                out.append(("bad", "Port plan says %s, Patchmanager cabling ends at %s." % (
                    live[0]["equipment"] or live[0]["device"], ", ".join(f["name"] for f in far_ends))))
            elif want:
                out.append(("ok", "Port plan and Patchmanager agree: %s." % live[0]["equipment"]))
        elif card["plan"] and not live and far_ends:
            out.append(("warn", "Port plan lists this port as free, but Patchmanager has a cable on it."))
        eq = card.get("equipment")
        if eq:
            p = str(card["nums"][1] if card.get("modular") and len(card["nums"]) > 1 else card["nums"][0])
            if p in eq["defect"]:
                out.append(("bad", "Patchmanager marks port %s of %s as DEFECT." % (p, eq["label"])))
            if p in eq["reserved"]:
                out.append(("info", "Port %s is reserved in Patchmanager." % p))
        if not card["cv_switch"]:
            out.append(("warn", "No CloudVision hostname known for %s (add it under Switches)." % card["switch"]))
        if any(s["open"] for s in card["snags"]):
            out.append(("warn", "Open entry in the snagging list."))
        return out

    # -- search ---------------------------------------------------------------------
    def search(self, q, limit=60):
        """Free text -> list of results {kind, title, sub, switch, nums | cable}.
        Understands 'Ethernet12 on dcbeswcpredg001', 'DCB.ESW A01 12', 'DCB.ESW A01.12',
        'dcbeswcpredg001 Et5/2', a cable id ('F03639'), equipment ('DCB.VS 02')."""
        q = (q or "").strip()
        if not q:
            return []
        res = []
        m = re.match(r"^\s*(et\S*)\s+on\s+(\S+)", q, re.I)
        if m:
            sw, intf = m.group(2), m.group(1)
        else:
            m = re.match(r"^(.*?)[\s.:]+((?:et(?:hernet)?)?\s*\d+(?:/\d+){0,2})\s*$", q, re.I)
            sw, intf = (m.group(1), m.group(2)) if m else (q, None)
        pm = self.pm_name(sw)
        nums = parse_intf(intf) if intf else None
        if pm and nums:
            res.append({"kind": "port", "title": "%s %s" % (pm, intf_name(nums)), "sub": self.cv_name(pm) or "",
                        "switch": pm, "nums": list(nums)})
        elif pm and self.cabling:
            for n, _ in self.cabling.ports_of(pm)[:limit]:
                res.append({"kind": "port", "title": "%s %s" % (pm, intf_name(n)), "sub": "", "switch": pm, "nums": list(n)})
        if self.cabling:
            for c in self.cabling.by_id.get(q.upper(), []):
                res.append({"kind": "cable", "title": "Cable %s" % c.id, "sub": "%s  <->  %s" % (c.a.name(), c.b.name()),
                            "cable": c.idx})
        if not res:
            nq = norm_name(q)
            # equipment / devices: port plan device names and cable ends
            for p in self.plan:
                if nq and (nq in norm_name(p["device"])):
                    res.append({"kind": "port", "title": "%s %s" % (p["switch"], intf_name(p["nums"])),
                                "sub": p["device"], "switch": p["switch"], "nums": p["nums"]})
                    if len(res) >= limit:
                        break
            if self.cabling and len(res) < limit:
                seen = set()
                for c in self.cabling.cables:
                    for side in ("a", "b"):
                        e = getattr(c, side)
                        if nq and nq in norm_name(e.top) and (c.idx, side) not in seen:
                            seen.add((c.idx, side))
                            res.append({"kind": "cable", "title": "%s" % e.name(), "sub": "cable %s (%s)" % (c.id, c.template),
                                        "cable": c.idx, "side": side})
                    if len(res) >= limit:
                        break
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
        known = {(norm_name(i["switch"] or ""), tuple(i["nums"])) for i in items}
        for e in cv_events or []:
            nums = parse_intf(e.get("interface"))
            pm = self.pm_name(e.get("device")) or e.get("device")
            if not nums or (norm_name(pm), nums) in known:
                continue
            items.append({"source": "cloudvision", "switch": pm, "cv_switch": e.get("device", ""),
                          "nums": list(nums), "intf": intf_name(nums), "what": e.get("title", ""),
                          "status": "CloudVision " + (e.get("severity") or "").lower(), "since": e.get("time", ""),
                          "who": "", "jira": "", "nr": "", "destination": ""})
        return items

    def status(self):
        return {"meta": self.meta, "switches": len(self.switches()), "mapping": len(self.mapping),
                "mapping_dups": self.mapping_dups}


def base_len(nums, modular):
    """The part of the interface that identifies the physical cage (lanes share it)."""
    return tuple(nums[:2] if modular else nums[:1])
