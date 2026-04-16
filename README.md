# AT Sentinel 🛡️
### SAP UAT Test Cycle Automation

AT Sentinel eliminates the daily manual overhead of managing SAP UAT test cycles.
It reads condition script Excel files, auto-populates execution dates and status in the
Test Cycle List, and sends targeted reminders — all without any API keys or IT approvals.

---

## Problems it solves

| # | Problem | Solution |
|---|---------|----------|
| 1 | Executors never enter actual dates in real time | Auto-reads condition scripts and backfills dates when steps are marked OK |
| 2 | "Complete" status set without script completion | Status derived solely from script content, not manual input |
| 3 | 2+ hours/day collecting status in standups and chats | Daily summary report distributed automatically |
| 4 | Prep reminders for upcoming tests are missed | Flags cycles starting today/tomorrow; alerts for overdue |
| 5 | No visibility into workstream or executor workload | Reports by area and executor with plan vs. actual |

---

## Prerequisites

- **Python 3.10+**
- **Windows** — required for Outlook email and Teams notifications (via `pywin32`)
  - On macOS/Linux the app runs but notifications are disabled; use dry-run mode
- **Microsoft Outlook** installed locally and signed in
- **OneDrive** sync folder containing the test Excel files (synced from SharePoint)

---

## Installation

```bash
# 1. Clone or extract the project
git clone <repo-url>
cd AT-Sentinel

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the app
streamlit run app.py
```

The app opens at **http://localhost:8501**.

---

## Quick start

1. Open the app in your browser
2. In the **sidebar**, set the **Test Files Folder** to your OneDrive-synced path
3. Set your **Manager Email**
4. Enable **Dry Run** (so nothing is written yet)
5. Click **▶ Run Now**
6. Review the results — verify cycles are found and status changes look correct
7. Disable Dry Run when ready for the live run

---

## Configuration

All settings are accessible in the sidebar. Changes take effect immediately — no restart needed.

### Test Files Folder
Full local path to the folder containing all Excel test files.
Example: `C:\Users\YourName\OneDrive - Company\SAP_UAT\AT1_Files`

### Manager Email
Receives the daily summary report (all time slots).

### Notification Time Slot
Controls which reminder types are sent when you click **▶ Run Now**:

| Slot | Time | Sends |
|------|------|-------|
| Morning | 9:00 | Overdue cycles + cycles starting today |
| Midday | 12:00 | Same as Morning |
| Evening | 17:00 | Above + cycles starting tomorrow + cycles due tomorrow |

### Notifications
- **Email** — sends via locally installed Outlook (no API needed)
- **Teams** — posts to a channel via its built-in email address  
  (Channel → ••• → Get email address)

### Options
| Toggle | Effect |
|--------|--------|
| Test Mode | Redirects all executor emails to manager email (safe for testing) |
| Dry Run | Computes changes but does NOT write Excel or send notifications |

### Executor Emails
Map executor display names to email addresses. Used for per-person reminder dispatch.
In Test Mode, all mail goes to the manager email regardless of this table.

---

## Execution Status Lifecycle

```
00.未開始 (Not Started)
    │  ≥1 step has an actual date entered
    ▼
10.実行中 (In Progress)
    │  Any step result = NG  ──►  SIR (regression found)
    │  All steps OK + all execution dates filled
    ▼
20.レビュー中 (In Review)
    │  All steps have a review completion date
    ▼
30.完了 (Complete)
```

Cancelled cycles (`99.キャンセル`) are excluded from all automation.
Deleted cycles (Deletion Flag = X) are also skipped entirely.

---

## Profile System

A **profile** bundles all settings that differ between test stages (AT-SAP, BLT, AT-IF, SIT):
- File name patterns (regex)
- Excel sheet names
- Column indices (0-based) for both the cycle list and condition scripts
- Cycle ID prefix and format

**AT-SAP** is the fully verified reference profile.
**BLT / AT-IF / SIT** ship as templates — verify column indices before use.

To edit a profile: **Sidebar → ✏️ Edit / Add Profiles**.
Changes are saved to `profiles.json` and persist across restarts.

---

## Reports (📊 Reports tab)

After running a sync (or clicking **📁 Load Data**), four reports are available:

| Report | Description |
|--------|-------------|
| 📋 Cycle & Step | Per-cycle table: status, step counts, plan/actual/review dates |
| 📅 Daywise | Cumulative line chart — Start Plan/Actual, Comp Plan/Actual(Exe), Comp Actual(Review) |
| 🏢 Streamwise | Workstream breakdown: total vs execution/review actual/ahead/delay |
| 👤 Executorwise | Same breakdown by executor name |

---

## Scheduling (3 daily runs — Windows Task Scheduler)

Create three Task Scheduler entries:

| Time | Slot argument | Command |
|------|-------------|---------|
| 09:00 | `morning` | `python "C:\path\AT_Sentinel.py" morning` |
| 12:00 | `midday` | `python "C:\path\AT_Sentinel.py" midday` |
| 17:00 | `evening` | `python "C:\path\AT_Sentinel.py" evening` |

Settings for each task:
- Run whether user is logged on or not
- Wake computer to run this task
- Run as soon as possible after a missed start

---

## Column Mapping Reference

### Test Cycle List (テストサイクル一覧) — 0-indexed

| Index | Column | Field |
|-------|--------|-------|
| 0 | A | Cycle ID (formula — reconstructed from H+C) |
| 2 | C | Sequence number |
| 7 | H | Area code |
| 8 | I | Regression flag (X = regression target) |
| 9 | J | Deletion flag (X = skip row) |
| 24 | Y | Plan start date (latest) |
| 25 | Z | Plan end date (latest) |
| 26 | AA | Actual start date |
| 27 | AB | Actual end date |
| 28 | AC | Review completion date *(verify per stage)* |
| 29 | AD | Review plan end date *(verify per stage, optional)* |
| 32 | AG | Executor |
| 33 | AH | Execution status |
| 34 | AI | Total test steps |
| 35 | AJ | Completed test steps |

### Condition Script (テスト仕様書兼結果記述書) — 0-indexed

| Index | Column | Field |
|-------|--------|-------|
| 8 | I | Excluded flag (X = skip step) |
| 25 | Z | Executor name |
| 26 | AA | Planned execution date |
| 27 | AB | Actual execution date |
| 28 | AC | Result (OK/NG/-) |
| 30 | AE | Retest executor |
| 31 | AF | Retest planned date |
| 32 | AG | Retest actual date |
| 33 | AH | Retest result |
| 35 | AJ | Reviewer name *(verify per stage)* |
| 36 | AK | Review actual date *(verify per stage)* |
| 37 | AL | Review result *(verify per stage)* |

All indices are placeholders for the new review columns — verify them in your actual
Excel files and update via the Profile Editor in the sidebar.

---

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| Cycle list not found | Wrong folder or pattern | Check folder path; verify `cycle_list_pattern` in profile |
| No condition scripts found | Pattern mismatch | Check `condition_script_pattern` matches filenames |
| Email fails | Outlook not open / pywin32 missing | Open Outlook; `pip install pywin32` |
| Status not changing | Column index wrong | Open profile editor; verify `exec_status` col index |
| Review dates not tracked | `review_date` col wrong | Verify `review_date` index in script_cols |
| Step counts are 0 | `excluded` col wrong or sentinel not found | Check `excluded` col index and that sentinel row (col A = "e") exists |
| Duplicate cycle ID warning | Two rows with same area+seq_no | Check for duplicate entries in cycle list |

---

## File Reference

| File | Purpose |
|------|---------|
| `AT_Sentinel.py` | Core automation engine |
| `app.py` | Streamlit web UI |
| `requirements.txt` | Python dependencies |
| `profiles.json` | Saved profile overrides (auto-created) |
| `runs_log.json` | Run history log (auto-created) |
| `CLAUDE.md` | Setup guide for Claude Code users |
