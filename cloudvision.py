"""
Read-only CloudVision (CVaaS / CVP) client - standard library only.

Auth: a CloudVision service-account token (Settings > Access Control > Service Accounts,
role with read access), sent as "Authorization: Bearer <token>". The token is read from
the environment variable CV_TOKEN or from cv_token.txt next to server.py (git-ignored);
it is never sent to the browser, never logged.

Used Resource APIs (REST form, GET only - this module has no write code path):
  inventory  /api/resources/inventory/v1/Device/all      hostname <-> device id (serial)
  events     /api/resources/event/v1/Event/all           interface events (errors, link down ...)
  endpoint   /api/resources/endpointlocation/v1/EndpointLocation?key.searchTerm=<ip|mac|name>
                                                         where a device / IP is plugged in

UNVERIFIED: written from Arista's Resource API docs, not yet run against
cv-prod-uk-1.arista.io. Field names are read defensively; check the "raw" sample in the
page's CloudVision tab on the first real connection.
"""
import json
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request


class CloudVisionError(Exception):
    pass


def _stream(text):
    """Resource API GetAll answers a JSON array, or one JSON object per line /
    concatenated. Yields each object's result.value (or the object)."""
    text = text.strip()
    if not text:
        return
    try:
        doc = json.loads(text)
        items = doc if isinstance(doc, list) else [doc]
    except ValueError:
        items, dec, i = [], json.JSONDecoder(), 0
        while i < len(text):
            while i < len(text) and text[i] in " \r\n\t,":
                i += 1
            if i >= len(text):
                break
            obj, i = dec.raw_decode(text, i)
            items.append(obj)
    for it in items:
        if isinstance(it, dict) and "result" in it:
            it = it["result"]
        if isinstance(it, dict) and "value" in it:
            yield it["value"], it.get("time")
        elif isinstance(it, dict) and "error" not in it:
            yield it, None
        elif isinstance(it, dict):
            raise CloudVisionError(str(it.get("error"))[:300])


def _find(obj, *names):
    """First value of any of `names` anywhere in a nested dict/list (CloudVision wraps
    values differently per API version: {"value": x}, {"values": {...}}, ...)."""
    stack = [obj]
    while stack:
        o = stack.pop(0)
        if isinstance(o, dict):
            for n in names:
                if n in o and o[n] not in (None, "", {}, []):
                    v = o[n]
                    if isinstance(v, dict) and "value" in v:
                        v = v["value"]
                    return v
            stack.extend(o.values())
        elif isinstance(o, list):
            stack.extend(o)
    return None


class CloudVision:
    def __init__(self, base_url, token, timeout=20, verify_tls=True):
        self.base = base_url.rstrip("/")
        if not self.base.startswith("https://"):
            raise CloudVisionError("CloudVision URL must start with https://")
        self.token = token
        self.timeout = timeout
        self.ctx = ssl.create_default_context() if verify_tls else ssl._create_unverified_context()

    def _get(self, path, params=None):
        url = self.base + path + ("?" + urllib.parse.urlencode(params) if params else "")
        req = urllib.request.Request(url, headers={"Authorization": "Bearer " + self.token,
                                                   "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=self.timeout, context=self.ctx) as r:
                return r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")[:300]
            if e.code in (401, 403):
                raise CloudVisionError("CloudVision refused the token (HTTP %d) - check the service account / role." % e.code)
            raise CloudVisionError("HTTP %d from %s: %s" % (e.code, path, body))
        except (urllib.error.URLError, OSError) as e:
            raise CloudVisionError("cannot reach %s: %s" % (self.base, getattr(e, "reason", e)))

    # -- inventory -------------------------------------------------------------
    def devices(self):
        """[{id, hostname, model, version, streaming}]"""
        out = []
        for v, _ in _stream(self._get("/api/resources/inventory/v1/Device/all")):
            out.append({
                "id": _find(v.get("key", {}), "deviceId", "device_id") or "",
                "hostname": v.get("hostname") or _find(v, "hostname") or "",
                "model": v.get("modelName") or "",
                "version": v.get("softwareVersion") or "",
                "streaming": v.get("streamingStatus") or "",
            })
        return out

    # -- events -------------------------------------------------------------------
    def events(self, since_hours=72, id_to_host=None, limit=2000):
        """Interface-related events of the last `since_hours`, newest first:
        [{key, time, severity, title, description, device, device_id, interface, ack, open}]"""
        id_to_host = id_to_host or {}
        cutoff = time.time() - since_hours * 3600
        out = []
        for v, _ in _stream(self._get("/api/resources/event/v1/Event/all")):
            ts = _find(v.get("key", {}), "timestamp") or ""
            t = _iso_to_epoch(ts)
            if t and t < cutoff:
                continue
            dev_id = _find(v, "deviceId", "device_id") or ""
            intf = _find(v, "interfaceId", "interface_id", "intfId") or ""
            if not intf:
                continue
            end = _find(v, "endTime", "end_time")
            out.append({
                "key": _find(v.get("key", {}), "key") or "",
                "time": ts,
                "severity": str(v.get("severity") or "").replace("EVENT_SEVERITY_", ""),
                "title": v.get("title") or "",
                "description": v.get("description") or "",
                "type": v.get("eventType") or "",
                "device_id": dev_id,
                "device": id_to_host.get(dev_id, dev_id),
                "interface": intf,
                "ack": bool(_find(v.get("ack") or {}, "ack")),
                "open": not end,
            })
        out.sort(key=lambda e: e["time"], reverse=True)
        return out[:limit]

    # -- endpoint location ------------------------------------------------------------
    def locate(self, term):
        """Where is an IP / MAC / hostname plugged in? [{device, device_id, interface, mac, ip, raw}]"""
        out = []
        txt = self._get("/api/resources/endpointlocation/v1/EndpointLocation", {"key.searchTerm": term})
        for v, _ in _stream(txt):
            for loc in _all(v, "deviceMap") or [v]:
                out.append({
                    "device_id": _find(loc, "deviceId") or "",
                    "interface": _find(loc, "interface", "interfaceName") or "",
                    "mac": _find(loc, "macAddress") or "",
                    "ip": _find(loc, "ipAddress", "ipV4Address") or "",
                    "raw": json.dumps(loc)[:600],
                })
        return out


def _all(obj, name):
    v = _find(obj, name)
    if isinstance(v, dict):
        vals = v.get("values", v)
        return list(vals.values()) if isinstance(vals, dict) else vals
    return v


def _iso_to_epoch(ts):
    """'2026-08-22T15:19:45.123456789Z' -> epoch seconds (UTC); 0 when unreadable."""
    import calendar
    import re
    m = re.match(r"(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)", str(ts or ""))
    if not m:
        return 0
    return calendar.timegm(tuple(int(x) for x in m.groups()) + (0, 0, 0))
