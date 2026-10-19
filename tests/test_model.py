"""Unit tests on small synthetic data - no real files, no network.

    python -m unittest discover -s tests
"""
import csv
import io
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import model  # noqa: E402
from cloudvision import _stream, _iso_to_epoch  # noqa: E402

HDR = ["Cable Identifier", "Cable Label", "Connected Equipment (Left) Top Level Equipment Label",
       "Connected Equipment (Left) Parent Equipment Label", "Connected Equipment (Left) Label",
       "Connectors (Left) Connected Port Label", "Connected Equipment (Left) Equipment Template Name",
       "Connected Equipment (Right) Top Level Equipment Label", "Connected Equipment (Right) Parent Equipment Label",
       "Connected Equipment (Right) Label", "Connectors (Right) Connected Port Label",
       "Connected Equipment (Right) Equipment Template Name", "Template Name", "User Attributes", "Length (Calculated)"]


def cables_csv(rows):
    f = io.StringIO()
    w = csv.writer(f)
    w.writerow(HDR)
    for r in rows:
        w.writerow(list(r) + ["Installed=true&Labeled=true", "1m"])
    return f.getvalue().encode()


ROWS = [
    # switch port 12 -> panel A port 12 (front)
    ("N1", "", "SW 01", "", "SW 01", "12", "Arista CCS-720XP-48Y6", "X.EPP 01", "", "X.EPP 01", "12", "Leviton 24 EPP", "Copper Patch Cord"),
    # panel rear port 12 -> outlet 3
    ("T1", "", "X.EPP 01", "", "X.EPP 01", "12", "Leviton 24 EPP", "R.NS 01", "", "R.NS 01", "03", "Excel GOP Box", "Fixed Copper-1"),
    # outlet 3 -> device
    ("N2", "", "R.NS 01", "", "R.NS 01", "03", "Excel GOP Box", "CAM 01", "", "CAM 01", "ETH", "Mobotix", "Copper Patch Cord"),
    # modular chassis LC03 port 1 -> fibre cassette A port 05, trunk A<->A, far cassette port 05 -> server
    ("F1", "", "VR 01", "LC03", "1", "1", "QSFP", "D.FPP 01", "D.FPP 01", "A", "05", "Leviton HDX", "SM LC"),
    ("TR1", "", "D.FPP 01", "D.FPP 01", "A", "", "", "E.FPP 01", "E.FPP 01", "A", "", "", "8 Strand SM Loose Tube"),
    ("F2", "", "E.FPP 01", "E.FPP 01", "A", "05", "Leviton HDX", "SRV 01", "", "SRV 01", "NIC1", "Dell", "SM LC"),
    # port 14 ends at a panel with nothing behind it
    ("N3", "", "SW 01", "", "SW 01", "14", "Arista CCS-720XP-48Y6", "X.EPP 01", "", "X.EPP 01", "14", "Leviton 24 EPP", "Copper Patch Cord"),
]


class ModelTest(unittest.TestCase):
    def setUp(self):
        self.s = model.Store()
        self.s.load("cables", "c.csv", cables_csv(ROWS))
        self.s.mapping_snags = {"SW 01": "SWHOST01", "VR 01": "VRHOST01"}
        self.s.plan = [{"switch": "SW 01", "nums": [12], "device": "CAM 01.ETH -- R/Wall", "equipment": "CAM 01",
                        "eqport": "ETH", "where": "R/Wall", "vlan": "134", "cable": "", "sheet": "t", "row": 1, "free": False}]
        self.s._rebuild_names()

    def test_parse_intf(self):
        self.assertEqual(model.parse_intf("Ethernet3/1/1"), (3, 1, 1))
        self.assertEqual(model.parse_intf("Et49"), (49,))
        self.assertEqual(model.parse_intf(" 5/2 "), (5, 2))
        self.assertIsNone(model.parse_intf("Management1"))

    def test_copper_path_through_panel_and_outlet(self):
        card = self.s.port_card("SWHOST01", (12,))
        self.assertEqual(card["switch"], "SW 01")
        p = card["paths"][0]
        ends = [h["name"] for h in p["hops"] if h["type"] == "end"]
        self.assertEqual(ends, ["SW 01 : 12", "X.EPP 01 : 12", "R.NS 01 : 03", "CAM 01 : ETH"])
        self.assertEqual(p["status"], "complete")
        self.assertIn("ok", [c[0] for c in card["checks"]])  # plan and cabling agree

    def test_modular_port_and_trunk(self):
        # chassis: CloudVision slot = Patchmanager LCxx + 2 (supervisors count in CloudVision)
        card = self.s.port_card("VR 01", (5, 1, 1))
        self.assertTrue(card["modular"])
        p = card["paths"][0]
        ends = [h["name"] for h in p["hops"] if h["type"] == "end"]
        self.assertEqual(ends[-1], "SRV 01 : NIC1")
        self.assertTrue(any(h.get("assumed") for h in p["hops"]))

    def test_chassis_slot_offset(self):
        self.assertEqual(self.s.port_card("VR 01", (3, 1, 1))["paths"], [])   # LC01 has no cable here
        self.assertEqual([n for n, _ in self.s.cabling.ports_of("VR 01")], [(5, 1)])
        self.s.plan.append({"switch": "VR 01", "nums": [3, 1], "device": "SRV 01.NIC1 -- R", "equipment": "SRV 01",
                            "eqport": "NIC1", "where": "R", "vlan": "", "cable": "", "sheet": "t", "row": 2, "free": False})
        self.s._rebuild_names()
        self.s._rebuild_names()   # twice: the shift must not add up
        self.assertEqual(self.s.plan[-1]["nums"], [5, 1])
        self.assertEqual(self.s.port_card("VR 01", (5, 1, 1))["plan"][0]["equipment"], "SRV 01")

    def test_dead_end_at_panel_is_open(self):
        card = self.s.port_card("SW 01", (14,))
        self.assertEqual(card["paths"][0]["status"], "open")

    def test_plan_mismatch(self):
        self.s.plan[0].update(equipment="OTHER 01", device="OTHER 01.ETH -- R/Wall")
        self.s._rebuild_names()
        card = self.s.port_card("SW 01", (12,))
        self.assertIn("bad", [c[0] for c in card["checks"]])
        self.assertEqual([r["verdict"] for r in card["compare"]], ["ref", "device"])

    def test_plan_other_port_same_device(self):
        self.s.plan[0].update(device="CAM 01.IO -- R/Wall", eqport="IO")
        self.s._rebuild_names()
        card = self.s.port_card("SW 01", (12,))
        self.assertEqual([r["verdict"] for r in card["compare"]], ["ref", "port"])

    def test_future_device_ports_to_shut(self):
        self.s.set_future({"CAM 01": {"note": "phase 2"}})
        rows = self.s.ports_to_shut()
        self.assertEqual([(r["switch"], r["intf"], r["device"]) for r in rows], [("SW 01", "Ethernet12", "CAM 01")])
        card = self.s.port_card("SWHOST01", (12,))
        self.assertTrue(card["shut"])
        self.assertIn("FUTURE", " ".join(t for _, t in card["checks"]))
        eos = model.eos_shutdown(rows)
        self.assertIn("! SWHOST01", eos)
        self.assertIn("interface Ethernet12\n   description FUTURE CAM 01 ETH\n   shutdown", eos)
        # the plan names another device on the port: listed, but not shut in the config
        self.s.plan[0].update(device="OTHER 01.ETH -- R/Wall", equipment="OTHER 01")
        self.s._rebuild_names()
        rows = self.s.ports_to_shut()
        self.assertIn("OTHER 01", rows[0]["conflict"])
        self.assertIn("! CHECK FIRST", model.eos_shutdown(rows))
        self.assertNotIn("\n   shutdown", model.eos_shutdown(rows))
        self.s.set_future({})
        self.assertFalse(self.s.port_card("SW 01", (12,))["shut"])

    def test_p2p_destination(self):
        self.assertEqual(self.s.device_split("P2P_swhost01_Ethernet15/2"), ("SW 01", "Ethernet15/2"))

    def test_naming_vs_real_mismatch(self):
        rel = self.s.device_relation
        self.assertEqual(rel("DCB.WS 2", "DCB.WS 02"), "naming")
        self.assertEqual(rel("WS 202", "DCB.WS 202"), "naming")
        self.assertEqual(rel("OC.PRIN 01", "OC.PRINT 01"), "naming")
        self.assertEqual(rel("SWHOST01", "SW 01"), "naming")          # CloudVision hostname vs PM name
        self.assertEqual(rel("06-25.PIR 01", "6-26.PIR 01"), "different")
        self.assertEqual(rel("DCB.XCL 21", "DCR.XCL 21"), "different")
        self.assertEqual(rel("CS20.AIF 01", "CS29.AIF 01"), "different")
        self.assertEqual(model.port_relation("ETH", "ETH.1"), "naming")
        self.assertEqual(model.port_relation("ETH.2", "ETH.1"), "different")
        self.s.plan[0].update(device="CAM 1.ETH.1 -- R/Wall", equipment="CAM 1", eqport="ETH.1")
        self.s._rebuild_names()
        card = self.s.port_card("SW 01", (12,))
        self.assertEqual([r["verdict"] for r in card["compare"]], ["ref", "naming"])
        self.assertIn("ok", [c[0] for c in card["checks"]])

    def test_fusion_lldp(self):
        state = {"devices": [
            {"name": "CAM 01", "ip": "10.0.0.1", "config": {"lldpFetchedAt": "2026-10-19T10:00:00", "lldp": {"neighbor": [
                {"interface": 1, "chassis": "AA:AA", "port": "Ethernet12", "system_name": "swhost01.example.local"},
                {"interface": 2, "chassis": "", "port": ""}]}}},
            {"name": "CAM 02", "ip": "10.0.0.2", "config": {"lldp": {"neighbor": [
                {"interface": 1, "chassis": "aa:aa", "port": "Ethernet13"}]}}},      # name missing -> MAC
        ]}
        rows = model.parse_fusion_state(state)
        self.assertEqual([(r["device"], r["iface"], r["switch_name"], r["port"]) for r in rows],
                         [("CAM 01", "e1", "swhost01", "Ethernet12"), ("CAM 02", "e1", "", "Ethernet13")])
        self.s.set_lldp(rows)
        self.assertEqual([r["switch"] for r in self.s.lldp], ["SW 01", "SW 01"])     # second via the learnt MAC
        self.assertTrue(self.s.lldp[1].get("via_mac"))
        # LLDP agrees with Patchmanager on SW 01 / 12 ...
        card = self.s.port_card("SW 01", (12,))
        lrow = next(r for r in card["compare"] if r["source"] == "LLDP (Fusion)")
        self.assertEqual(lrow["verdict"], "ok")
        # ... and sees CAM 02 on port 13, where nothing is documented -> listed
        self.assertEqual(self.s.port_card("SW 01", (13,))["lldp"][0]["device"], "CAM 02")

    def test_red_blue(self):
        self.assertEqual(model.switch_color("DCB.VR 01"), "blue")
        self.assertEqual(model.switch_color("IDF1.VSWR 01"), "red")
        self.assertEqual(model.switch_color("DCB.ESW A01"), "")
        rows = [{"device": "FUS 01", "iface": "e1", "switch_name": "IDF2.VSWB 01", "chassis": "", "port": "Ethernet1", "fetched": ""},
                {"device": "FUS 01", "iface": "e2", "switch_name": "IDF1.VSWR 01", "chassis": "", "port": "Ethernet1", "fetched": ""}]
        self.s.set_lldp(rows)
        self.assertEqual([p["kind"] for p in self.s.lldp_problems()], ["Red / Blue swapped"])

    def test_same_port(self):
        self.assertTrue(model.same_port("PCLAN1", "PC LAN 1"))
        self.assertTrue(model.same_port("ETH.1", "Eth"))
        self.assertTrue(model.same_port("10G.1", "Expansion / 10G.1 / 10G.1"))
        self.assertFalse(model.same_port("ETH.2", "ETH.1"))

    def test_search_forms(self):
        for q in ("Ethernet12 on swhost01", "SW 01 12", "SW 01.12", "swhost01 Et12"):
            r = self.s.search(q)
            self.assertTrue(r and r[0]["kind"] == "port" and r[0]["nums"] == [12], q)
        self.assertEqual(self.s.search("N2")[0]["kind"], "cable")

    def test_portplan_blocks(self):
        sheets = {"Access": [["", "Switch", "Port", "Device", "VLAN", "Cable No", "", "Switch", "", "Port", "Device", "VLAN", "Cable No"],
                             ["", "SW 01", "3", "PC 01.ETH.1 -- R/Desk", "274", "12000", "", "VR 01", "2", "4", "free", "free", "free"]]}
        plan = model.parse_portplan(sheets)
        self.assertEqual([(p["switch"], p["nums"], p["equipment"], p["eqport"]) for p in plan],
                         [("SW 01", [3], "PC 01", "ETH.1"), ("VR 01", [2, 4], "", "")])
        self.assertTrue(plan[1]["free"])

    def test_snags(self):
        sheets = {"Master": [["Nr", "Switch name  CV", "Switch CAD", "Port CV", "Error description", "status manually", "date first occurance", "name"],
                             ["1", "SWHOST01", "SW 01", "Ethernet12", "Symbol errors", "fixed", "46154", "AN"],
                             ["2", "SWHOST01", "SW 01", "Ethernet12", "Link down", "", "", "0"]],
                  "Names": [["PM", "arista cloud"], ["SW 01", "SWHOST01"], ["FW 01", "0"]],
                  "events-dump": [["Nr", "Switch name  CV", "Port CV", "status manually"], ["", "x", "Ethernet1", ""]]}
        snags, mapping = model.parse_snags(sheets)
        self.assertEqual(mapping, {"SW 01": "SWHOST01"})
        self.assertEqual(len(snags), 2)  # dump sheet ignored when a Master exists
        self.assertEqual(snags[0]["first"], "2026-05-12")
        self.assertFalse(snags[0]["open"])
        self.assertTrue(snags[1]["open"])
        self.assertEqual(snags[1]["who"], "")

    def test_cv_stream_formats(self):
        a = '{"result": {"value": {"hostname": "a"}}}\n{"result": {"value": {"hostname": "b"}}}'
        b = '[{"result": {"value": {"hostname": "a"}}}]'
        self.assertEqual([v["hostname"] for v, _ in _stream(a)], ["a", "b"])
        self.assertEqual([v["hostname"] for v, _ in _stream(b)], ["a"])
        self.assertEqual(_iso_to_epoch("1970-01-01T00:01:00.5Z"), 60)


if __name__ == "__main__":
    unittest.main()
