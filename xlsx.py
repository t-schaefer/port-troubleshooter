"""
Minimal .xlsx / .xlsm reader - standard library only (zipfile + ElementTree).

read_xlsx(path_or_bytes) -> {sheet name: [[cell, ...], ...]}
Cells are strings ("" when empty); numbers come back as their stored text
("46154", "3.5"). Formulas give their cached value. Macros are never run
(they are not even read).
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET

_M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS = {"m": _M, "r": _R}


def _col(ref):
    n = 0
    for c in re.match(r"[A-Z]+", ref).group():
        n = n * 26 + ord(c) - 64
    return n - 1


def read_xlsx(src, max_rows=None):
    z = zipfile.ZipFile(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", _NS):
            shared.append("".join(t.text or "" for t in si.iter("{%s}t" % _M)))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    out = {}
    for s in wb.find("m:sheets", _NS):
        target = rels[s.get("{%s}id" % _R)].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        rows = []
        for _, el in ET.iterparse(z.open(target)):
            if el.tag != "{%s}row" % _M:
                continue
            cells = {}
            for c in el.findall("m:c", _NS):
                v, typ = c.find("m:v", _NS), c.get("t")
                if typ == "s" and v is not None:
                    val = shared[int(v.text)]
                elif typ == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter("{%s}t" % _M))
                else:
                    val = v.text if v is not None else ""
                if val:
                    cells[_col(c.get("r"))] = val
            # keep the row position (empty rows included) so row numbers match Excel
            rn = int(el.get("r", len(rows) + 1))
            while len(rows) < rn - 1:
                rows.append([])
            rows.append([cells.get(i, "") for i in range(max(cells) + 1)] if cells else [])
            el.clear()
            if max_rows and len(rows) >= max_rows:
                break
        out[s.get("name")] = rows
    return out
