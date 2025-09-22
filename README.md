# Ping Tester

Keep tabs on any fleet of machines straight from Excel. Feed the tool an `.xlsx` inventory and it will fan out pings in parallel, split multi-interface cells automatically, and hand back both a clear console summary and a polished workbook you can share.

## Highlights
- Multi-IP aware: a single cell like `172.25.22.5, 172.25.22.6` is expanded and checked host-by-host.
- Styled Excel report: results land in `{input}_results.xlsx` with color-coded status, frozen header row, filters, and auto-fit columns.
- Smart header detection for common name/IP labels, with explicit overrides when needed.
- Cross-platform: uses the system `ping` command on Windows, macOS, and Linux.
- Fast by default thanks to configurable worker pools and timeouts.

## Quick Start
```powershell
# Optional: create a virtual environment
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install --upgrade pip openpyxl

# Run against the default ip.xlsx
py tester.py

# Or point to a different workbook and change settings
py tester.py inventory.xlsx --sheet "Prod" --timeout 1500 --workers 32
```

macOS/Linux users can swap `py` for `python3` and source the virtual environment with `source .venv/bin/activate`.

## Excel Expectations
- Default input file is `ip.xlsx` (override by passing a path or using `--sheet`).
- The first row should contain headers; auto-detection looks for typical name/IP keywords.
- Use `--name-column` / `--ip-column` if the tool cannot guess correctly.

Example rows:

| VM Name | IP Addresses |
| --- | --- |
| web-01 | 192.168.1.10 |
| firewall | 172.25.22.5, 172.25.22.6 |

## Command Reference
Flag | Description
---- | -----------
`excel` | Workbook to inspect (defaults to `ip.xlsx`).
`--sheet` | Worksheet name; uses the active sheet when omitted.
`--name-column` | Header text to treat as the machine name column.
`--ip-column` | Header text to treat as the IP column.
`--timeout` | Ping timeout per host in milliseconds (default: `1000`).
`--workers` | Number of concurrent ping workers (default: `16`).
`--output` | Destination for the generated report (defaults to `{input_file}_results.xlsx`).

## What You Get
- Console table summarising each interface and its `UP`/`DOWN` status.
- Excel report with:
  - A bold navy header row frozen in place.
  - Auto-filter enabled, so you can instantly slice by up/down.
  - Auto-sized columns for clean readability.
  - Green fill for `UP`, red for `DOWN`.

Exit status is `0` only when every interface responds. Any unreachable host returns `1` so you can plug the command straight into automation.

## Troubleshooting
- "openpyxl is required": install it (`pip install openpyxl`).
- "'ping' command not found": ensure your OS packages supply a `ping` binary.
- Firewalls can block ICMP; double-check policy if you receive unexpected `DOWN` results.

## Project Files
- `tester.py` – command-line entrypoint.
- `ip.xlsx` – sample/working inventory (you can replace it).
- `.gitignore` – ignores virtual env artefacts and generated Excel outputs.

You are encouraged to add a project license if you plan to publish the repository.
