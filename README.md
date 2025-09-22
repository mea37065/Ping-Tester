# Ping Tester

Ping Tester turns an Excel inventory into a fast, reliable health check. Point it at a workbook, and it will fan out pings in parallel, split up any cells that list more than one IP address, and publish the results both to the terminal and to a polished Excel report you can hand off to teammates.

## Why You Will Like It
- Understand hosts with multiple interfaces; entries like "172.25.22.5, 172.25.22.6" are expanded automatically.
- Receive a styled workbook (`{input}_results.xlsx`) with a frozen header, auto-filter, and colour-coded status cells.
- Let the tool guess sensible column headers, or override them when the sheet uses custom labels.
- Run it anywhere the `ping` command exists: Windows, macOS, or Linux.
- Tune worker count and timeouts to match your environment.

## Getting Started
1. Make sure Python 3.8 or newer is installed.
2. (Optional) Create a virtual environment and install dependencies:
   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install --upgrade pip openpyxl
   ```
   On macOS or Linux replace the activation command with `source .venv/bin/activate` and use `python3` instead of `py`.
3. Run the script:
   ```powershell
   py tester.py
   ```
   To target another workbook or adjust behaviour:
   ```powershell
   py tester.py inventory.xlsx --sheet "Production" --timeout 2000 --workers 32
   ```

## Excel Input Tips
- The default input file is `ip.xlsx` in the project root. Pass a different path as the first argument when needed.
- Keep headers on the first row. The loader recognises common variants for name and IP columns, for example `Host`, `VM`, `IP`, or `Address`.
- Use `--name-column` and `--ip-column` when the sheet uses unusual labels.

Example rows:

| VM Name | IP Addresses         |
| ------- | -------------------- |
| web-01  | 192.168.1.10         |
| fw-01   | 172.25.22.5, 172.25.22.6 |

## Command Reference

| Flag          | Description                                                            |
| ------------- | ---------------------------------------------------------------------- |
| `excel`       | Path to the Excel workbook (defaults to `ip.xlsx`).                    |
| `--sheet`     | Worksheet name (defaults to the active sheet).                         |
| `--name-column` | Header text to use for the machine name column.                     |
| `--ip-column` | Header text to use for the IP address column.                          |
| `--timeout`   | Ping timeout per host in milliseconds (default `1000`).                |
| `--workers`   | Number of concurrent workers (default `16`).                           |
| `--output`    | Destination for the generated report (defaults to `{input}_results.xlsx`). |

## Output Overview
- Console table showing each interface and whether it responded (`UP`) or not (`DOWN`).
- Excel workbook with:
  - Frozen header row for easier scrolling.
  - Auto-filter enabled across the dataset.
  - Auto-sized columns.
  - Green fill for `UP` and red fill for `DOWN` cells.

The script exits with status `0` only when every interface responds. Any failure yields exit code `1`, making the tool easy to drop into automation.

## Troubleshooting
- Missing `openpyxl`: install it with `pip install openpyxl`.
- `'ping' command not found`: ensure your operating system provides the `ping` utility and that it is on `PATH`.
- Unexpected `DOWN` entries can be caused by network firewalls or devices that block ICMP.

## Project Files
- `tester.py` - command-line entry point.
- `ip.xlsx` - sample inventory workbook.
- `.gitignore` - excludes virtual environments and generated result sheets.
- `LICENSE` - the MIT license for this project.

## License
Ping Tester is distributed under the MIT License. Read the [LICENSE](LICENSE) file for the full text.
