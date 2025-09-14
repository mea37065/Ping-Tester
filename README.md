# Ping Tester

A simple cross-platform CLI tool to ping many hosts listed in an Excel (.xlsx) file. It auto-detects common column headers, runs pings concurrently, and prints a compact summary table of results.

## Features
- Auto-detects header columns for IP and name (with overrides)
- Concurrent pinging for speed on large lists
- Works on Windows, macOS, and Linux (uses the system `ping`)
- Clear summary with per-host UP/DOWN status

## Requirements
- Python 3.8+
- Python package: `openpyxl`
- System `ping` command available on PATH

## Installation
- Optional virtualenv (recommended)
  - Windows PowerShell:
    ```powershell
    py -m venv .venv
    .\.venv\Scripts\Activate.ps1
    pip install --upgrade pip
    pip install openpyxl
    ```
  - macOS/Linux:
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate
    pip install --upgrade pip
    pip install openpyxl
    ```

## Excel Input
- Default file name: `ip.xlsx` in the project root.
- Sheet: active sheet by default; override with `--sheet`.
- Header auto-detection:
  - IP column candidates: `ip`, `ipaddress`, `ipv4`, `address`
  - Name column candidates: `vm`, `vmname`, `name`, `hostname`, `host`, `machine`, `computer`, `server`
- You can force specific headers with `--ip-column` and `--name-column`.

Example layout (first row is headers):

| VM | IP |
| --- | --- |
| web-01 | 192.168.1.10 |
| db-01 | 192.168.1.11 |

## Usage
Run the script directly with Python. The default input file is `ip.xlsx`.

- Basic (uses `ip.xlsx` active sheet):
  - Windows PowerShell:
    ```powershell
    py tester.py
    ```
  - macOS/Linux:
    ```bash
    python3 tester.py
    ```

- Specify a different file and sheet, and force headers:
  ```bash
  python3 tester.py my-hosts.xlsx --sheet "Inventory" --name-column "Hostname" --ip-column "IP Address"
  ```

- Tune concurrency and timeouts:
  ```bash
  python3 tester.py ip.xlsx --workers 32 --timeout 1500
  ```

Arguments:
- `excel`: Path to the Excel file (default: `ip.xlsx`)
- `--sheet`: Worksheet name (default: workbook active sheet)
- `--name-column`: Header text for the name column (override auto-detection)
- `--ip-column`: Header text for the IP column (override auto-detection)
- `--timeout`: Ping timeout per host in milliseconds (default: 1000)
- `--workers`: Number of concurrent workers (default: 16; set `1` for sequential)

## Output
The tool prints a simple summary table, for example:

```
VM Name   IP Address     Status
-------   ----------     ------
web-01    192.168.1.10   UP
db-01     192.168.1.11   DOWN
```

Exit code is `0` if all hosts are UP, otherwise `1`.

## Notes and Tips
- Windows vs Unix ping:
  - Windows: `ping -n 1 -w <ms>`
  - Unix: `ping -c 1 -W <s>` (timeout in whole seconds)
- If you see many DOWN results, ensure targets allow ICMP Echo (firewall/router may block).
- For large lists, increase `--workers` to improve throughput. Be mindful of network and host rate limits.
- If auto-detection fails, pass `--ip-column` and `--name-column` explicitly and ensure headers are on the first row.

## Troubleshooting
- `Error: openpyxl is required ...`: Install with `pip install openpyxl`.
- `Error: 'ping' command not found ...`: Ensure `ping` exists in your system PATH. On minimal containers, install the OS’s `iputils`/`inetutils`.
- File not found: Verify the Excel file path, or pass the correct filename.
- Wrong sheet: Use `--sheet` with the exact worksheet name.

## Project Files
- `tester.py`: Main CLI script.
- `.gitignore`: Ignores `output/` and `__pycache__/` (feel free to adjust).

## License
Specify a license for your project if you plan to share it publicly.

