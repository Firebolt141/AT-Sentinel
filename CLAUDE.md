# AT Sentinel — Claude Code Setup Guide

This file tells Claude Code exactly how to set up and run this project.
If you are a human reading this, see **README.md** for the full user guide.

---

## What this project does

AT Sentinel automates daily SAP UAT test cycle management:
- Reads condition script Excel files and auto-populates execution dates and status
- Updates the Test Cycle List (サイクル一覧) on SharePoint via OneDrive sync
- Sends reminders to executors and a daily summary to the manager
- Generates reports: daywise progress, workstream breakdown, executor breakdown

---

## Setup steps (Claude, follow these in order)

### 1. Install dependencies
```bash
pip install -r requirements.txt
```
Note: `pywin32` is Windows-only. On Linux/macOS it will be skipped — email/Teams sending
won't work, but the app runs fully in dry-run mode for testing.

### 2. Verify Python version
```bash
python --version
```
Must be 3.10 or higher (uses X | Y union type hints).

### 3. Configure the app

Open `AT_Sentinel.py` and update the `CONFIG` dict near the top:

| Key | What to set |
|-----|-------------|
| `folder` | Full local path to the OneDrive-synced folder containing all Excel files |
| `manager_email` | Manager's email address for daily summary reports |
| `executor_emails` | Dict mapping executor display names → email addresses |
| `teams_channel_email` | Teams channel email (channel → ••• → Get email address) |
| `test_mode` | Keep `True` until executor emails are verified |
| `dry_run` | Set `True` to simulate without writing or sending |

These values can also be changed at runtime in the Streamlit sidebar — no restart needed.

### 4. Verify profile column indices

The default `AT-SAP` profile has placeholder column indices for the new review columns
(`review_end`, `review_plan_end`, `review_date`, `review_pic`, `review_result`).
These must be verified against the actual Excel files before going live.

Use the **sidebar → Edit / Add Profiles** section in the UI to adjust indices.
All indices are 0-based.

### 5. Start the app
```bash
streamlit run app.py
```
The app opens at http://localhost:8501 in your browser.

### 6. First run (dry-run recommended)
1. In the sidebar, set **Dry Run = ON**
2. Set **Test Mode = ON**
3. Click **▶ Run Now**
4. Review the summary — check that cycles are found and status changes look correct
5. Once verified, turn off Dry Run and Test Mode for the live run

---

## Scheduling (Windows Task Scheduler — 3 daily runs)

The spec calls for notifications at 09:00, 12:00, and 17:00.
Create three Task Scheduler entries, each running:
```
python "C:\path\to\AT_Sentinel.py" <slot>
```
where `<slot>` is `morning`, `midday`, or `evening`.

Or run the Streamlit app and click **▶ Run Now** with the appropriate time slot selected.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| "Cycle list file not found" | Wrong folder path or pattern | Check `folder` config and `cycle_list_pattern` in profile |
| "No condition scripts found" | Pattern mismatch | Verify `condition_script_pattern` matches filenames |
| Email fails / win32com not found | pywin32 not installed or not Windows | `pip install pywin32`; Outlook must be open and signed in |
| Status not updating | Column indices wrong | Open profile editor in sidebar and verify all col indices |
| Review dates not tracking | `review_date` col index wrong | Verify `review_date` index in profile's script_cols |
| Duplicate cycle ID warning | Two rows reconstruct to same ID | Check for duplicate area+seq_no entries in cycle list |

---

## Key files

| File | Purpose |
|------|---------|
| `AT_Sentinel.py` | Core engine: file discovery, analysis, update, notifications, reports |
| `app.py` | Streamlit web UI |
| `requirements.txt` | Python dependencies |
| `profiles.json` | User-saved profile overrides (auto-created on first profile save) |
| `runs_log.json` | Run history (auto-created on first run) |
