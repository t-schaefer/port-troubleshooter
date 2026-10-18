# Handover: add the Port Troubleshooter to the tools landing page

For whoever maintains `tools-landing` (github.com/t-schaefer/tools-landing).

## The service

| | |
|---|---|
| Name | Port Troubleshooter |
| Repository | https://github.com/t-schaefer/port-troubleshooter |
| Folder on the server | next to the landing page: `..\port-troubleshooter\` (same parent as `tools-landing\`) |
| Port | **8092** (HTTP, `/`) |
| Windows Scheduled Task | `PortTroubleshooterService` (SYSTEM, at boot, auto-restart) |
| Install | `Install-PortTroubleshooter.bat` (self-elevates, opens firewall for 8092) |
| Update | `update-windows.ps1` (`git pull --ff-only`, or GitHub `main.zip` without git) |
| Stop | `stop-windows.ps1` |
| Health check | `GET http://<host>:8092/api/status` → 200 JSON |
| Logo | none yet – suggestion: add `logos/arista.svg` (+ `arista-white.svg`) and use `"logo": "arista"`; until then `""` |

## Install on the server (once)

```bat
cd <parent folder of tools-landing>
git clone https://github.com/t-schaefer/port-troubleshooter.git
port-troubleshooter\Install-PortTroubleshooter.bat
```

Then import the data in the page (**⟳ Update data**, or drag the files onto the
Data & CloudVision tab). Put new exports in `port-troubleshooter\inbox\` and they show up
in the update menu.

## Entry for `tools.example.json` (and the server's `tools.json`)

Paths are relative to the `tools-landing` folder, like the Neuron and Fusion entries.

```json
{
  "name": "Port Troubleshooter",
  "description": "Find broken or wrongly documented connections: cable path from Patchmanager, IP list, snagging list and CloudVision side by side; mismatches, signal chain, ports to shut.",
  "port": 8092,
  "path": "/",
  "host": "",
  "logo": "",
  "note": "New – feedback to Thorsten.",
  "task": "PortTroubleshooterService",
  "update": "../port-troubleshooter/update-windows.ps1",
  "stop": "../port-troubleshooter/stop-windows.ps1"
}
```

## Notes

- Port 8092 sits between Fusion (8090), MuoN config (8091) and Neuron (8101); the quick-test
  batch `Start-PortTroubleshooter.bat` uses 8098.
- The service keeps its own state next to it (`data\`, `config.json`, `switches.json`,
  `future.json`, `cv_token.txt`, `inbox\`, `service.log`). All are git-ignored, so the
  update never touches them.
- It is read-only towards the network: no switch is ever configured. The "To shut" EOS
  config is text for an engineer to apply.
- User overview: `README.md`; technical notes: `docs/TECHNICAL.md`.
