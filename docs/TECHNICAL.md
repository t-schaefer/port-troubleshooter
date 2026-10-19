# Port Troubleshooter – technical notes

Short overview for users: [../README.md](../README.md). This file: inputs, matching rules, CloudVision, deployment.

```text

One page per switch port: the documented cable path from Patchmanager (through every
patch panel, outlet and trunk), what the IP list expects there, the snagging-list
history and the live CloudVision interface events - side by side, plus a worklist of
everything that is still open.

  Search accepts:   Ethernet49 on idf2eswcpredg03      (paste straight from CloudVision)
                    DCB.ESW A01 12   /   DCB.ESW A01.12   /   dcbeswcpredg001 Et12
                    F03639                              (Patchmanager cable ID)
                    any name a device is known by - also partial, case / dots / leading
                    zeros ignored ('dcb vs 2' finds DCB.VS 02):
                      Patchmanager label, alias (CS01 AMU 01), asset tag (QID...), description
                      (Dataminer), IP list name (MC1STBDGTPLY001), snagging-list
                      destination, CloudVision hostname
  Port card:        Sources compared (Patchmanager cabling / IP list / snagging list:
                    same / other device / other port), Copy job card, Copy link, Print,
                    ◀ port / port ▶ to step through neighbouring ports.
  Device card:      every cabled port traced to its far end with a link to the switch port,
                    IP list vs. cabling per device port, all names the device is known by.
  Mismatches tab:   every switch port where the sources disagree (errors / warnings /
                    check on site / info), click opens the port.
  CSV dump:         "Download CSV" on Worklist and Mismatches (Excel-ready, UTF-8):
                    /export/worklist.csv, /export/mismatches.csv

  Signal chain:     "Signal chain" on a device card: follows every cable through panels AND
                    through devices, shown as one tree. Switches and hubs (devices with more
                    than 16 cabled ports - KVM, frames, servers) end a branch; open them to go
                    on. Every link is checked for faults: open snag, open CloudVision event,
                    Patchmanager 'to be checked' / DEFECT port, documentation mismatch, two
                    cables on one port, path ending at a panel. Fault list + CSV download.
  Lists:            clicking a row in Worklist / Mismatches / To shut opens the details in a
                    side panel - the list keeps its filter, sort and position (Esc closes).
                    Filters and sort are remembered in the browser and shown above the table.

  Notes & Jira:     comments, Jira links and manual worklist items per port / device in
                    notes.json (git-ignored, written atomically). Users are free text with a
                    pick list of names already used (notes + snagging list "name"). Manual
                    items stay open until "Done" (who + when recorded). config.json
                    "jira_url" turns keys (ABC-123) into links to <jira_url>/browse/ABC-123.
  Ports tab:        every switch port (cabled or in the IP list) with device, patch panels,
                    cables, IP list and status; group by switch / device / patch panel; CSV.
  Devices tab:      every device / switch / patch panel; click shows all its connections.

  Chassis slots:    CloudVision counts the supervisor slots, Patchmanager (LC01..) and the
                    IP list (1/1/1) count line cards: CloudVision Ethernet3/1/1 = LC01 port 1
                    = IP list 1/1/1 (Cabling.slot_offset = 2). Port cards use CloudVision names.
  CV port exports:  CloudVision "Ethernet interfaces" CSV per switch; switch + export time
                    from the file name (cvp-ethernet-interfaces-<host>-<YYYY-MM-DD-hh-mm-ss>.csv).
                    Status / speed / optic / description per port; description is compared
                    like a source. Not live - the export time is shown.
  CV links:         <cv_url>/cv/interface/overview/<serial>/<interface>. Serials per hostname
                    in serials.example.json (defaults) + serials.json (Switches tab).
  Neuron list:      rear / front cable checked against Patchmanager (IPG port = channel with
                    direction, panel port R/F + number); card- or panel-wide numbering offsets
                    are reported once instead of per BNC.

  Fusion LLDP:      every 5 min the service reads the Fusion Dashboard's /api/state
                    (config fusion_url, default http://localhost:8090; read-only, from the
                    dashboard's cache - the Fusions are not asked again). The Fusion API sends
                    no switch name, only chassis MAC + port ({"interface":1,"chassis":
                    "c4:ca:2b:..","port":"Ethernet4/22/4"}). MAC -> switch: typed in under
                    Switches (wins), else learnt by vote from the documentation (CloudVision
                    descriptions, IP list): a device documented on <switch> <port> with the
                    matching e1/e2 (device port ...1 / ...2) votes for <switch>; >= 2 votes
                    and 60 %. LLDP is a source on the port card (compared by device
                    only), "LLDP sees another device" is a real mismatch, "LLDP shows the far
                    end, cabling incomplete" where the Patchmanager path stops at a panel.
  Red / Blue:       e1 must be on a Red, e2 on a Blue 2110 switch: ...SWR / ...SWB by type,
                    the spines by room (DCR = Red, DCB = Blue). Swapped / same switch / wrong
                    network are real mismatches.

  state.db:         SQLite (stdlib) with notes, ignores, future, serials, switch-name
                    overrides and the history (imports, evaluation states with mismatch
                    counts, mismatch keys of the last 30 states for "new / resolved"). Old
                    JSON files are taken over at the first start (*.json.migrated kept).
  Data states:      an import builds a complete new data state next to the running one
                    (files + settings + last LLDP + Ports / Mismatches lists) and then switches
                    over in one step - other requests keep being answered meanwhile; while a
                    list is rebuilt, visitors get the previous one.
  Golden test:      tests/test_golden.py checks the numbers of the local data/ against
                    tests/golden.json (counts only); --update after intended changes.

  Future use:       "Mark as future use" on a device card (stored in future.json, git-
                    ignored) - or "Future Expansion" ticked on the device in Patchmanager.
                    Its switch ports (via cabling AND IP list) are listed under "To shut",
                    flagged on the port card and in the worklist, with CSV and Arista EOS
                    config (text only - the tool never touches a switch). A port where the
                    documentation names another device is marked "conflict - check first"
                    and only written as a comment, never as "shutdown".

  Names / identity: case, spaces, dots and dashes do not matter, but the letter / number
                    blocks do: DCR.SGXR 103 = DCR.SGXR103, DCB.ESW B01 != DCB.ESWB 01.
  Naming only:      same device / port, written differently - leading zeros (DCB.WS 2 =
                    DCB.WS 02), room prefix left out (WS 202 = DCB.WS 202), CloudVision
                    hostname, one-letter typo with all numbers equal (OC.PRIN 01 = OC.PRINT 01;
                    never across rooms: DCB / DCR stay different), ETH = ETH.1, MGNT = MGMT.
                    Shown blue, not counted as an error. Two cables on one port: the one the
                    other sources agree with is the reference.
  Speed:            the service listens at once and loads its data in the background (page
                    shows "Loading data on the server..."); Ports / Mismatches are built in the
                    background; requests slower than 0.5 s are logged as "SLOW ..." in
                    service.log.

  Port names are compared tolerantly: MGNT = MGMT, ETH.1 = Ethernet = ENET 1 = LAN,
  ETH.3 = 03, S5. D1 = S5 / D1, 10G.2 = 10G / 2 / 2. Different numbers stay different.


INPUTS (tab "Data & CloudVision", drag & drop - the type is detected from the content)
------
  Patchmanager cable export (CSV)       Cable Identifier, Connected Equipment (Left/Right) ...
  Patchmanager equipment export (CSV)   Equipment Identifier, Template, Position, Defect Ports ...
  IP list "View" (xlsx)                 sheets with Switch | Port | Device | VLAN | Cable No blocks
  Connectivity snagging list (xlsx)     sheet "Master" (status manually, ...) and the
                                        PM / "arista cloud" name sheet (switch name mapping)

"⟳ Update data" (top right) scans the import folders - config.json "import_dirs",
default ["inbox"] (the inbox\ folder next to server.py, created at start; add e.g. a
share or "~/Downloads"). It shows the newest file of each type next to what is loaded and
ticks the ones that are newer. Only files inside these folders can be imported this way.

A new import replaces the previous file of the same type. The last file of each type is
kept in data\ and loaded again when the service starts. data\ is git-ignored - the files
contain project data and never go into the repository.

Not imported: the .xlsm IP list with macros - export / save the "View" as .xlsx.


HOW THE PATH IS FOLLOWED
------------------------
- From the switch port along the cable to the other end.
- At a patch panel / outlet the same port has a second cable (front / rear): continue there.
- Fibre trunks (loose tube, MPO) are documented in Patchmanager on the cassette WITHOUT a
  port. The tool jumps to the trunk's far cassette and continues at the SAME port number.
  This is an assumption - shown as "trunk - same port assumed" / "assumed".
- Status: complete = ends at a device;  ends at panel / open = last point is a panel or
  outlet with nothing patched beyond it (e.g. MPO breakout cassettes);  ambiguous = more
  than one cable continues from the same point.
- Interface names: Ethernet49 = port 49;  Ethernet5/2 = port 5 lane 2 (breakout);
  on chassis switches (ports under LCxx in Patchmanager) Ethernet3/1/1 = LC03 port 1 lane 1.

Checks shown on a port card: no cable documented, path open / ambiguous, IP list vs.
Patchmanager far end (mismatch = red), port marked DEFECT / reserved in Patchmanager,
cable "to be checked", no CloudVision hostname known, open snag.


SWITCH NAMES (tab "Switches")
-----------------------------
Patchmanager "DCB.ESW A01" <-> CloudVision "DCBESWCPREDG001" comes from the snagging list's
name sheet. Wrong / missing entries: type the hostname in the Override column and Save
(stored in switches.json, git-ignored). Duplicate hostnames are flagged - as of the
2026-10-18 list: IDF1/IDF2.ESW 11 and 12 share hostnames with ESW 09 and 10.


CLOUDVISION (optional, read-only)  -  UNVERIFIED
-----------
1. In CloudVision (cv-prod-uk-1.arista.io): Settings > Access Control > Service Accounts,
   create an account with a READ-ONLY role, generate a token.
2. Put the token into cv_token.txt next to server.py (one line), or set the environment
   variable CV_TOKEN. Never commit it (cv_token.txt is git-ignored).
3. URL / TLS / time window: config.json (created from config.example.json).
4. Restart the service. The Data tab shows device / event counts or the error.

What the service asks CloudVision for, in the background (viewers never trigger it):
  inventory  every 10 min   /api/resources/inventory/v1/Device/all
  events     every 2 min    /api/resources/event/v1/Event/all  (interface events, last 72 h)
  locate     on button only /api/resources/endpointlocation/v1/EndpointLocation (IP/MAC/host)
There is no write code path. The Resource API calls are written from Arista's docs and NOT
yet tested against the real CVaaS tenant - on the first connection check "Raw sample" in
the Data tab and fix the field mapping in cloudvision.py if needed.


RUN
---
  Mac / quick test:   python3 server.py            -> http://localhost:8092/
  Windows test:       Start-PortTroubleshooter.bat -> http://localhost:8098/
  Windows service:    Install-PortTroubleshooter.bat   (Scheduled Task "PortTroubleshooterService",
                      SYSTEM, at boot, auto-restart, firewall rule for 8092)
  Update:             Update-PortTroubleshooter.bat    (git pull --ff-only, restart)
  Stop:               Stop-PortTroubleshooter.bat      Remove: uninstall-task-windows.ps1

Python 3.8+, standard library only. Settings at the top of server.py.
Tests: python -m unittest discover -s tests
```
