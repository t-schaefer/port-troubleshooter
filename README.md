# Port Troubleshooter

Find a broken or wrongly documented connection in seconds. The tool puts all port
documentation side by side, shows where the sources disagree and traces every cable
end to end. It only reads; it never changes a switch or a document.

## How it works

It combines four exports that you load into the tool, plus live data if configured:

| Source | What it tells |
|---|---|
| **Patchmanager cables** (CSV) | which cable runs where, through every patch panel |
| **Patchmanager equipment** (CSV) | rack, U position, aliases, asset tags, defect ports |
| **IP list** (View, xlsx) | which device *should* be on which switch port |
| **Connectivity snagging list** (xlsx) | known problems, status, who fixed what |
| **CloudVision** (optional, live) | current interface errors and link events |

Loading files: click **⟳ Update data** at the top right and upload the exports from your
PC (several at once, the type is detected automatically). Alternatively copy them into the
`inbox` folder on the server and pick them in the same menu. Your data is never in git, so a
new installation starts empty.

## What you can do

- **Search by any name.** Device label, alias, asset tag, hostname, part of a name,
  a cable ID, or a CloudVision line such as `Ethernet49 on idf2eswcpredg03`.
- **Port card.** The full cable path from switch port to device, what the IP list says,
  snagging history and live events. Sources that disagree are marked in red.
- **Device card.** Every port of a device and where it goes, IP list vs. cabling.
- **Signal chain.** Follows a device's cables through panels *and* devices and lists every
  faulty link: open snag, CloudVision error, defect port, documentation mismatch.
- **Worklist.** Everything open (snagging list plus CloudVision) in one sortable list.
- **Mismatches.** Every port where Patchmanager, the IP list and the snagging list disagree.
- **To shut.** Mark a device as *future use*: its switch ports are listed with ready-made
  Arista `shutdown` config. Ports where the documentation disagrees are never included.
- **Notes & Jira.** Every port and device card has notes: who (type a name or pick one
  used before), a comment and a Jira link. Tick "also put it on the worklist" – or press
  **+ Add to worklist** – and it stays in the Worklist until someone marks it done.
- **Hand it on.** Copy a job card for a technician, copy a link, print, or download a CSV.

Tip: in a list, clicking a row opens the details on the right and the list stays where it
was. Press Esc to close.

## Run it

| | |
|---|---|
| Open | `http://<server>:8092/` |
| Install as a Windows service | `Install-PortTroubleshooter.bat` |
| Update to the latest version | `Update-PortTroubleshooter.bat` |
| Stop / quick test | `Stop-PortTroubleshooter.bat` / `Start-PortTroubleshooter.bat` |
| Mac / Linux | `python3 server.py` |

You need Python 3.8 or newer. Nothing else to install.

## Good to know

- Your data stays local. Imports are kept in `data/`, which is never uploaded to git.
- Jira keys like `PLS-123` link to qvest-media.atlassian.net (`jira_url` in the config).
- CloudVision is optional: put a read-only token in `cv_token.txt` and restart.
- Details on matching rules, settings and CloudVision: [docs/TECHNICAL.md](docs/TECHNICAL.md)
