# Process SR Report

A small Windows utility you start manually when you want it active. It watches a folder (Downloads by default) for TeamBinder reports, **only files that start with `ETSA-TSA-` and end with `.xls`**, and that are **no older than 12 hours**. When a file appears, it:

1. Waits until the download completes.
2. Converts `.xls` to `.xlsx` using Excel (Excel must be installed).
3. Transforms Column **B** (starting at row **19**) by appending the hyperlink once, as: `display text ### url`.
4. Copies `Processed_YYYYMMDD_HHMMSS.xlsx` into a **SharePoint library folder you select** on first run. That folder **must be synced locally via OneDrive**; OneDrive then uploads the file automatically.

> **Why it works:** SharePoint libraries can be synced to a local folder using the OneDrive client; any file copied there is uploaded to SharePoint. citeturn8search69

## First run
- The app explains the **SharePoint Sync** requirement and prompts you to select the synced library folder.
- It then offers to change the watch folder (default: your **Downloads**).
- Your choices are stored in `%APPDATA%\ProcessSRReport\settings.json`.
- Re-run the wizard any time via: `ProcessSRReport.exe --reset`

## Buildless usage for end users
Download the `.exe` from the **Actions** artifact and double‑click it. There is **no startup install**; run it only when you want it active.

## For maintainers: GitHub Actions build
This repository includes a workflow that builds a single-file Windows `.exe` with PyInstaller every push and on demand. The artifact is named **ProcessSRReport**.

## Local build (optional)
If you prefer to build locally:
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --noconfirm --onefile --name ProcessSRReport watch_and_process.py
```

## Tech notes
- Folder watching uses the Python **watchdog** library for low-latency file system events. citeturn8search57
- Excel automation for `.xls → .xlsx` uses **pywin32** (Excel COM). citeturn8search46
- Hyperlink extraction in `.xlsx` uses **openpyxl** (via `cell.hyperlink.target`) and a fallback for `=HYPERLINK()` formulas. citeturn8search40
