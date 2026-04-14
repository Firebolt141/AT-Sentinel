"""
================================================================================
SAP UAT Test Management Automation
================================================================================
Version  : 2.1
Author   : Adarsh A. Shetty
Platform : Windows (local OneDrive sync + Outlook / Teams)

--------------------------------------------------------------------------------
PURPOSE
--------------------------------------------------------------------------------
Automates the daily overhead of managing ~500 SAP UAT test cycles across
domestic (Japan) and international (India) sales-order scenarios.

Problems this script solves:
  #1  Actual execution dates are never entered by executors in real time.
      → Reads condition script files and auto-populates actual dates when all
        steps are marked OK. No manual entry required.
  #2  False "Complete" status without script completion (regression risk).
      → Execution status is derived solely from script content, not manual input.
  #3  2+ hours/day in standups and 20+ individual chats to collect status.
      → Daily summary report is generated and distributed automatically at 8am.
  #4  No visibility into upcoming tests; prep reminders are missed.
      → Flags cycles starting within N days; alerts for overdue and stalled.
  #5  Scheduling conflicts (multiple cycles on same day/executor) are invisible.
      → Report surfaces workload concentration per executor per day.

--------------------------------------------------------------------------------
END-TO-END FLOW
--------------------------------------------------------------------------------
  1. DISCOVER FILES
     Scans the configured local folder (OneDrive-synced path) for:
       • One cycle list file  — filename matches regex: サイクル一覧
       • N condition scripts  — filename matches regex: 仕様書兼結果記述書.*AT1TC
     Only .xlsx files are considered. Excel lock files (~$*) are skipped.

  2. ANALYZE CONDITION SCRIPTS  (analyze_script)
     For each condition script file:
       a. Opens the sheet "テスト仕様書兼結果記述書"
       b. Skips header rows (rows 0–3) and excluded steps (col 8 = 'X')
       c. Stops processing at sentinel row where col 0 = 'e'
       d. A row is a valid step if it has a step number (col 1) OR result data
          — handles merged cells where step number only appears in first merged row
       e. Counts total active steps vs completed (result = OK) steps
       f. Collects first and last actual execution date from OK rows
          — retest date (col 32) takes priority over first-run date (col 27)
       g. Collects earliest planned execution date from any step (col 26)
       h. Detects NG results (regression indicator)
       i. Extracts executor names for backfill into cycle list

  3. UPDATE CYCLE LIST  (update_cycle_list)
     Opens テストサイクル一覧 sheet (read-write, formulas preserved).
     For each data row (from row 4 onward):
       a. Reconstructs cycle ID from area (col H) + seq_no (col C)
          The formula in col A cannot be evaluated by openpyxl, so the ID is
          rebuilt using the same logic: f"AT1TC_{area}{seq_no:03d}"
       b. Skips rows where deletion flag (col J / index 9) = 'X'
          NOTE: col I / index 8 is the REGRESSION flag — do not confuse the two
       c. Looks up the matching script result by cycle ID
       d. Updates these columns (guarded as noted):
            plan_start_latest (col 24) ← script's earliest planned date  [only if blank]
            actual_start      (col 26) ← script's first OK actual date   [only if blank]
            actual_end        (col 27) ← script's last OK actual date    [only if blank, all_ok only]
            exec_status       (col 33) ← derived from completion state   [always overwritten]
            total_steps       (col 34) ← count of active steps           [always overwritten]
            complete_steps    (col 35) ← count of OK steps               [always overwritten]
            executor          (col 32) ← from script's exec_pic column   [only if blank]
       e. Saves the workbook in place — OneDrive syncs it back to SharePoint

  4. DETECT REMINDERS  (get_reminders)
     Re-reads the (now updated) cycle list and flags:
       • upcoming — plan_start_latest within CONFIG["reminder_days_ahead"] days,
                    cycle not yet started
       • overdue  — plan_end_latest has passed, cycle not complete
       • stalled  — in-progress, plan_end within last 5 days, no actual_end set

  5. SEND NOTIFICATIONS  (notify)
     Dispatches to one or both channels based on CONFIG["notify_channels"]:

       "email"
         Sends via the locally installed Outlook app using win32com.
         No API keys or auth tokens — uses your already-logged-in session.
         Executor reminders go to individual executor email addresses.
         In test_mode, all executor mail is redirected to manager_email.

       "teams"
         Posts to a Teams channel via its built-in email address.
         Every Teams channel has a unique email (find it via channel → ••• →
         "Get email address"). Messages appear as email cards in the channel.
         No webhook, no Azure AD app registration, no IT approval required.
         Store the channel address in CONFIG["teams_channel_email"].

       Both channels can be active simultaneously.
       CONFIG["notify_channels"] = ["email", "teams"]  → both
       CONFIG["notify_channels"] = ["email"]            → email only
       CONFIG["notify_channels"] = ["teams"]            → Teams only

       UI NOTE: A future UI layer will expose channel selection per-run.
       The config values here serve as the default / fallback.

  6. REPORTS
     • Manager daily summary — HTML table: completed / started / overdue /
                               stalled / upcoming cycles
     • Executor reminders    — per-executor HTML message per reminder type
                               (upcoming / overdue / stalled)

--------------------------------------------------------------------------------
FILE NAMING CONVENTIONS
--------------------------------------------------------------------------------
  Cycle list (one per folder):
    Must contain:  サイクル一覧
    Example:       AT1_SAP内結合テストサイクル一覧_v1.0.xlsx

  Condition scripts (one per cycle):
    Must contain:  仕様書兼結果記述書  AND  AT1TC
    Example:       ⑧_TE586_SAP内結合テスト仕様書兼結果記述書_AT1TC_SDFI001.xlsx

  Cycle ID mapping (filename → cycle list):
    Filename uses  "SDFI001"  but cycle list stores  "SD/FI001"
    The slash is inserted automatically:
      if len(area_code) > 2 → area_code[:2] + "/" + area_code[2:]
    Examples:  SDFI → SD/FI  |  MM → MM  |  FI → FI

--------------------------------------------------------------------------------
COLUMN MAPPING — テストサイクル一覧  (0-indexed)
--------------------------------------------------------------------------------
  Index  0  (col A) : テストサイクルID       — FORMULA; reconstructed from H + C
  Index  1  (col B) : サイクル名称
  Index  2  (col C) : 領域内サイクル連番     — used in cycle ID reconstruction
  Index  7  (col H) : 担当領域               — area code (SD/FI, MM, FI, etc.)
  Index  8  (col I) : リグレッション対象フラグ — 'X' = regression target (NOT deletion)
  Index  9  (col J) : 削除フラグ             — 'X' = deleted; row is skipped entirely
  Index 24  (col Y) : 実行開始予定日(最新)   — plan_start_latest; backfilled from script
  Index 25  (col Z) : 実行完了予定日(最新)   — plan_end_latest; not auto-populated
  Index 26  (col AA): 実行開始実績日         — actual_start; set from first OK date
  Index 27  (col AB): 実行完了実績日         — actual_end; set from last OK date (all_ok)
  Index 32  (col AG): 実行担当者             — executor; backfilled from script
  Index 33  (col AH): 実行ステータス         — exec_status; always synced by automation
  Index 34  (col AI): 総テストステップ数     — total_steps; always synced
  Index 35  (col AJ): 完了テストステップ数   — complete_steps; always synced

  Header rows: 3 (data starts at Excel row 4 / openpyxl min_row=4)

--------------------------------------------------------------------------------
COLUMN MAPPING — テスト仕様書兼結果記述書  (0-indexed)
--------------------------------------------------------------------------------
  Index  8  (col I) : テスト対象外           — 'X' = skip this step entirely
  Index 25  (col Z) : 初回実行 担当者        — executor name
  Index 26  (col AA): 初回実行 実行予定日    — planned_date; feeds earliest_planned
  Index 27  (col AB): 初回実行 実行日        — actual_date; feeds first/last_actual_date
  Index 28  (col AC): 初回実行 テスト結果    — OK / NG / -
  Index 30  (col AE): 再実行 担当者
  Index 31  (col AF): 再実行 実行予定日
  Index 32  (col AG): 再実行 実行日          — takes priority over first-run actual_date
  Index 33  (col AH): 再実行 テスト結果      — takes priority over first-run result

  Header rows: 4 (data starts at Excel row 5 / row_idx >= 4 in iter_rows)
  Sentinel:   Processing stops when col 0 = 'e'

--------------------------------------------------------------------------------
EXECUTION STATUS VALUES  (dropdown from the 'list' sheet)
--------------------------------------------------------------------------------
  00.未開始    — Not started   (no steps complete, no NG)
  10.実行中    — In progress   (≥1 step complete, OR any NG present)
  20.レビュー中 — Reviewing    (not set by automation — manual use only)
  30.完了      — Complete      (all active steps = OK, zero NG)
  99.キャンセル — Cancelled    (excluded from all automation processing)

--------------------------------------------------------------------------------
REQUIREMENTS
--------------------------------------------------------------------------------
  Python   : 3.10+  (uses X | Y union type hints)
  Packages : pip install openpyxl pywin32
  OS       : Windows only (win32com requires the Windows COM interface)
  Outlook  : Must be installed locally and signed in
  OneDrive : Test file folder must be synced to a local path

--------------------------------------------------------------------------------
SCHEDULING  (Windows Task Scheduler — recommended setup)
--------------------------------------------------------------------------------
  Action   : Start a program
  Program  : C:\path\to\python.exe
  Arguments: "C:\path\to\this_script.py"
  Start in : C:\path\to\script_folder
  Trigger  : Daily at 08:00
  Settings : Run whether user is logged on or not
             Wake computer to run this task
             Run task as soon as possible after a scheduled start is missed

--------------------------------------------------------------------------------
KNOWN LIMITATIONS
--------------------------------------------------------------------------------
  • openpyxl cannot evaluate Excel formulas, so the cycle ID in col A (which
    is a formula) is rebuilt from area (col H) and seq_no (col C) instead.
    If the formula logic changes in the Excel file, update reconstruct_cycle_id.

  • Merged cells in condition scripts: if a step spans multiple rows via merge,
    openpyxl returns None for the cell values in non-anchor rows. The script
    handles this by accepting rows with either a step number OR result data,
    but very unusual merge layouts may miscount steps.

  • Teams channel email posts appear as email cards in the channel, not as
    native chat messages. For richer card formatting, replace the Teams sender
    with a Power Automate HTTP-triggered flow.

  • test_mode = True redirects ALL executor emails to manager_email.
    Set to False only after fully populating executor_emails in CONFIG.

  • plan_end_latest (col 25) is intentionally not auto-populated by this script.
    It is left for manual entry or a future enhancement.

--------------------------------------------------------------------------------
CHANGE LOG
--------------------------------------------------------------------------------
  v2.1  Added plan_start_latest backfill from script's earliest planned date
        Added Teams channel email notification (no webhook / API required)
        Added notify_channels config for dual email + Teams delivery
        Added notify() dispatcher; run() uses it instead of direct send calls
        Added this detailed header documentation

  v2.0  Full rewrite — fully local, no Azure / Graph API / Power Automate
        Cycle ID reconstructed from area + seq_no (formula workaround)
        Deletion flag (col 9) correctly distinguished from regression flag (col 8)
        Executor backfill, step count sync, retest-takes-priority logic

  v1.0  Initial prototype
================================================================================
"""

import os
import io
import re
import logging
from datetime import datetime, date, timedelta
from pathlib import Path

import openpyxl

# ─────────────────────────────────────────────────────────────
# CONFIG  ← edit these before first run
# ─────────────────────────────────────────────────────────────
CONFIG = {
    # Full path to the folder containing ALL test files
    "folder": r"C:\Users\adarsh.a.shetty\Downloads",

    # Filename patterns to identify files (regex)
    # Cycle list: contains "サイクル一覧" in the name
    "cycle_list_pattern": r"サイクル一覧",

    # Condition scripts: contains "仕様書兼結果記述書" AND "AT1TC" (to exclude unit test scripts)
    "condition_script_pattern": r"仕様書兼結果記述書.*AT1TC|AT1TC.*仕様書兼結果記述書",

    # How many days ahead to warn executors about upcoming cycles
    "reminder_days_ahead": 2,

    # Your email — ALL emails (manager report + executor reminders) go here during testing
    "manager_email": "adarsh.a.shetty@accenture.com",

    # TEST MODE: all executor reminder emails are redirected to manager_email above.
    # Set to False only when you are ready to send to real executors.
    "test_mode": True,

    # Executor name → email mapping
    # Not used while test_mode=True (all mail goes to manager_email instead)
    # Fill these in before going live.
    "executor_emails": {
        # "Suzuki":  "saeko.suzuki@accenture.com",
        # "鈴木":    "saeko.suzuki@accenture.com",
        # "Tanaka":  "tanaka.x@accenture.com",
        # "田中":    "tanaka.x@accenture.com",
    },

    # Log file location
    "log_file": r"C:\Users\adarsh.a.shetty\Downloads\sap_test_automation.log",

    # ── Notification channels ──────────────────────────────────────────────
    # List any combination of "email" and "teams".
    # "email"  → sends via local Outlook app (win32com, no API needed)
    # "teams"  → posts to a Teams channel via its email address (see below)
    # A future UI layer will let the user override this per run.
    "notify_channels": ["email", "teams"],

    # Teams channel email address.
    # How to find it: open the Teams channel → click ••• → "Get email address"
    # Leave blank ("") to disable Teams notifications even if "teams" is listed above.
    "teams_channel_email": "",
}

# ─────────────────────────────────────────────────────────────
# COLUMN INDICES (0-based) — from your actual files
# ─────────────────────────────────────────────────────────────

# テストサイクル一覧 sheet
CYCLE_COLS = {
    # NOTE: col 0 (テストサイクルID) is a formula ="AT1TC_"&H5&TEXT(C5,"000")
    # We DO NOT read col 0 directly. Instead we reconstruct the ID from:
    #   area (col 7) + seq_no (col 2) to match against script filenames.
    "cycle_id":          0,   # テストサイクルID (formula — read via reconstruct_cycle_id)
    "seq_no":            2,   # 領域内サイクル連番 (e.g. 1, 2, 3)
    "cycle_name":        1,   # サイクル名称
    "area":              7,   # 担当領域 (e.g. "SD/FI")
    "regression_flag":   8,   # リグレッション対象フラグ (X = regression, NOT deletion)
    "deletion_flag":     9,   # 削除フラグ (X = deleted, skip)
    "plan_start_latest": 24,  # 実行開始予定日(最新)
    "plan_end_latest":   25,  # 実行完了予定日(最新)
    "actual_start":      26,  # 実行開始実績日
    "actual_end":        27,  # 実行完了実績日
    "executor":          32,  # 実行担当者
    "exec_status":       33,  # 実行ステータス
    "total_steps":       34,  # 総テストステップ数
    "complete_steps":    35,  # 完了テストステップ数
}

# テスト仕様書兼結果記述書 sheet
SCRIPT_COLS = {
    "excluded":       8,   # テスト対象外 ('X' = skip this step)
    "exec_pic":       25,  # 初回実行 担当者
    "planned_date":   26,  # 初回実行 実行予定日
    "actual_date":    27,  # 初回実行 実行日
    "result":         28,  # 初回実行 テスト結果 (OK / NG / -)
    "retest_pic":     30,  # 再実行 担当者
    "retest_planned": 31,  # 再実行 実行予定日
    "retest_date":    32,  # 再実行 実行日
    "retest_result":  33,  # 再実行 テスト結果
}

# Execution status dropdown values (from the 'list' sheet)
STATUS = {
    "not_started": "00.未開始",
    "in_progress": "10.実行中",
    "reviewing":   "20.レビュー中",
    "complete":    "30.完了",
    "cancelled":   "99.キャンセル",
}

# Header rows to skip in condition scripts (0-indexed rows 0-3 are headers, data starts row 4)
SCRIPT_HEADER_ROWS = 4
# Header rows to skip in cycle list (rows 0-2 are headers, data starts row 3)
CYCLE_HEADER_ROWS = 3

# ─────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────
def setup_logging():
    # Guard against duplicate handlers when called multiple times (e.g. Streamlit reruns)
    root = logging.getLogger()
    if root.handlers:
        return
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(CONFIG["log_file"], encoding="utf-8"),
        ]
    )

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────
# FILE DISCOVERY
# ─────────────────────────────────────────────────────────────
def find_files(folder: str) -> tuple[str | None, list[str]]:
    """
    Scan folder for:
    - One cycle list file (matches cycle_list_pattern)
    - All condition script files (match condition_script_pattern)

    Returns (cycle_list_path, [script_paths])
    """
    cycle_list = None
    scripts = []

    for f in Path(folder).iterdir():
        if not f.suffix.lower() == ".xlsx":
            continue
        if f.name.startswith("~$"):  # skip Excel temp/lock files
            continue

        name = f.name
        if re.search(CONFIG["cycle_list_pattern"], name):
            cycle_list = str(f)
            log.info(f"Cycle list: {name}")
        elif re.search(CONFIG["condition_script_pattern"], name):
            scripts.append(str(f))
            log.info(f"Condition script: {name}")

    log.info(f"Found {len(scripts)} condition scripts")
    return cycle_list, scripts


def reconstruct_cycle_id(area: str, seq_no) -> str | None:
    """
    Reconstruct the cycle ID from area + seq_no columns — same logic as the Excel formula:
      ="AT1TC_"&H5&TEXT(C5,"000")
    e.g. area="SD/FI", seq_no=1 → "AT1TC_SD/FI001"
    """
    if not area or seq_no is None:
        return None
    try:
        return f"AT1TC_{str(area).strip()}{int(seq_no):03d}"
    except (ValueError, TypeError):
        return None


def extract_cycle_id_from_filename(filename: str) -> str | None:
    """
    Extract cycle ID from the condition script filename.
    e.g. '⑧_TE586_SAP内結合テスト仕様書兼結果記述書_AT1TC_SDFI001.xlsx'
         → 'AT1TC_SD/FI001'

    The filename uses 'SDFI001' but the cycle list uses 'SD/FI001'.
    We insert '/' before the last letter+digits block if area contains no slash.
    """
    # Match pattern like AT1TC_SDFI001 or AT1TC_SD001 etc.
    match = re.search(r'(AT\dTC)_([A-Z]+)(\d+)', filename)
    if not match:
        return None

    prefix = match.group(1)   # e.g. AT1TC
    area   = match.group(2)   # e.g. SDFI or SD
    num    = match.group(3)   # e.g. 001

    # Insert slash: SDFI → SD/FI, MM → MM, FI → FI
    # Rule: split after first 2 uppercase chars if total > 2
    if len(area) > 2:
        area_with_slash = area[:2] + "/" + area[2:]
    else:
        area_with_slash = area

    return f"{prefix}_{area_with_slash}{num}"


# ─────────────────────────────────────────────────────────────
# CONDITION SCRIPT ANALYSIS
# ─────────────────────────────────────────────────────────────
def analyze_script(filepath: str) -> dict | None:
    """
    Read a condition script and return completion status.

    Returns dict with:
      cycle_id          - extracted from filename
      total_steps       - active (non-excluded) steps
      completed_steps   - steps with result = OK
      all_ok            - True if all active steps are OK and no NG
      has_ng            - True if any step has NG
      first_actual_date - earliest actual execution date
      last_actual_date  - latest actual execution date
      executor          - name(s) from 実行担当者 column
      earliest_planned  - earliest planned execution date (for reminders)
    """
    cycle_id = extract_cycle_id_from_filename(Path(filepath).name)
    if not cycle_id:
        log.warning(f"Could not extract cycle ID from: {Path(filepath).name}")
        return None

    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    except Exception as e:
        log.warning(f"Could not open {Path(filepath).name}: {e}")
        return None

    sheet_name = "テスト仕様書兼結果記述書"
    if sheet_name not in wb.sheetnames:
        log.warning(f"Sheet '{sheet_name}' not found in {Path(filepath).name}")
        return None

    ws = wb[sheet_name]
    c = SCRIPT_COLS

    total = 0
    completed = 0
    has_ng = False
    actual_dates = []
    planned_dates = []
    executors = set()

    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx < SCRIPT_HEADER_ROWS:
            continue

        # Stop at sentinel row
        if row and row[0] == "e":
            break

        # Skip empty rows
        if not row or all(v is None for v in row):
            continue

        # Skip excluded steps (col 8 = テスト対象外)
        if len(row) > c["excluded"] and row[c["excluded"]] == "X":
            continue

        # A valid test step has EITHER a step number (col 1) OR test result data
        # Handles merged cells where step number only appears in the first merged row
        has_step_no = row[1] is not None
        has_result_data = (
            len(row) > c["result"] and row[c["result"]] in ("OK", "NG", "-")
        ) or (
            len(row) > c["retest_result"] and row[c["retest_result"]] in ("OK", "NG", "-")
        ) or (
            len(row) > c["actual_date"] and row[c["actual_date"]] is not None
        ) or (
            len(row) > c["planned_date"] and row[c["planned_date"]] is not None
        )

        if not has_step_no and not has_result_data:
            continue

        total += 1

        # Collect executor name
        pic = row[c["exec_pic"]] if len(row) > c["exec_pic"] else None
        if pic and str(pic).strip() not in ("", "-"):
            executors.add(str(pic).strip())

        # Collect planned date
        planned = row[c["planned_date"]] if len(row) > c["planned_date"] else None
        if planned and isinstance(planned, (datetime, date)):
            d = planned.date() if isinstance(planned, datetime) else planned
            planned_dates.append(d)

        # Determine result — retest takes priority over first run
        retest_result = row[c["retest_result"]] if len(row) > c["retest_result"] else None
        first_result  = row[c["result"]]         if len(row) > c["result"]         else None
        result = retest_result if retest_result and retest_result != "-" else first_result

        retest_date = row[c["retest_date"]] if len(row) > c["retest_date"] else None
        first_date  = row[c["actual_date"]] if len(row) > c["actual_date"] else None
        actual = retest_date if retest_date else first_date

        if result == "OK":
            completed += 1
            if actual and isinstance(actual, (datetime, date)):
                d = actual.date() if isinstance(actual, datetime) else actual
                actual_dates.append(d)
        elif result == "NG":
            has_ng = True

    wb.close()

    return {
        "cycle_id":          cycle_id,
        "total_steps":       total,
        "completed_steps":   completed,
        "all_ok":            total > 0 and completed == total and not has_ng,
        "has_ng":            has_ng,
        "first_actual_date": min(actual_dates)  if actual_dates  else None,
        "last_actual_date":  max(actual_dates)  if actual_dates  else None,
        "executor":          ", ".join(sorted(executors)) if executors else None,
        "earliest_planned":  min(planned_dates) if planned_dates else None,
    }


# ─────────────────────────────────────────────────────────────
# CYCLE LIST UPDATE
# ─────────────────────────────────────────────────────────────
def update_cycle_list(cycle_list_path: str, script_results: dict[str, dict]) -> list[dict]:
    """
    Open the Test Cycle List, update status/dates/counts from script results,
    save in place. Returns list of changed rows for the report.

    IMPORTANT: We use load_workbook without data_only so formulas are preserved.
    We ONLY write to columns that contain plain values (not formulas).
    Formula columns (flags, report tabs) are left untouched.
    """
    wb = openpyxl.load_workbook(cycle_list_path)  # NOT read_only, NOT data_only
    ws = wb["テストサイクル一覧"]
    c = CYCLE_COLS
    today = date.today()
    changes = []

    for row in ws.iter_rows(min_row=CYCLE_HEADER_ROWS + 1):
        # Reconstruct cycle ID from area + seq_no (col 0 is a formula openpyxl can't evaluate)
        area_val   = row[c["area"]].value
        seq_val    = row[c["seq_no"]].value
        cycle_id   = reconstruct_cycle_id(area_val, seq_val)

        if not cycle_id:
            continue

        # Skip header remnants and sentinel rows
        raw_col0 = str(row[c["cycle_id"]].value or "").strip()
        if raw_col0 in ("e", "テストサイクルID", "Test cycle ID"):
            continue

        # Skip deleted cycles (col 9 = deletion flag, NOT col 8 which is regression flag)
        if row[c["deletion_flag"]].value == "X":
            continue
        result = script_results.get(cycle_id)
        if not result:
            continue  # no matching script found for this cycle

        current_status = row[c["exec_status"]].value or STATUS["not_started"]

        # Determine new status
        if result["all_ok"]:
            new_status = STATUS["complete"]
        elif result["completed_steps"] > 0 or result["has_ng"]:
            new_status = STATUS["in_progress"]
        else:
            new_status = STATUS["not_started"]

        row_changed = False

        # planned start date — backfill from script if blank in ichiran
        if result["earliest_planned"] and not row[c["plan_start_latest"]].value:
            row[c["plan_start_latest"]].value = result["earliest_planned"]
            row_changed = True

        # actual start date — only set if blank
        if result["first_actual_date"] and not row[c["actual_start"]].value:
            row[c["actual_start"]].value = result["first_actual_date"]
            row_changed = True

        # actual end date — only set when all OK and currently blank
        if result["all_ok"] and result["last_actual_date"] and not row[c["actual_end"]].value:
            row[c["actual_end"]].value = result["last_actual_date"]
            row_changed = True

        # execution status
        if new_status != current_status:
            row[c["exec_status"]].value = new_status
            row_changed = True

        # step counts — always sync
        row[c["total_steps"]].value    = result["total_steps"]
        row[c["complete_steps"]].value = result["completed_steps"]

        # executor — backfill if blank
        if result["executor"] and not row[c["executor"]].value:
            row[c["executor"]].value = result["executor"]

        if row_changed:
            plan_end = row[c["plan_end_latest"]].value
            if isinstance(plan_end, datetime):
                plan_end = plan_end.date()

            changes.append({
                "cycle_id":        cycle_id,
                "cycle_name":      row[c["cycle_name"]].value,
                "area":            row[c["area"]].value,
                "old_status":      current_status,
                "new_status":      new_status,
                "completed_steps": result["completed_steps"],
                "total_steps":     result["total_steps"],
                "executor":        result["executor"],
                "plan_end":        plan_end,
                "all_ok":          result["all_ok"],
                "has_ng":          result["has_ng"],
            })

    wb.save(cycle_list_path)
    wb.close()
    log.info(f"Cycle list saved — {len(changes)} rows updated")
    return changes


# ─────────────────────────────────────────────────────────────
# REMINDER DETECTION
# ─────────────────────────────────────────────────────────────
def get_reminders(cycle_list_path: str, today: date) -> list[dict]:
    """
    Scan cycle list for cycles needing attention:
    - upcoming:  plan start within N days, not yet started
    - overdue:   plan end has passed, not complete
    - stalled:   in progress, plan end within last 5 days, no actual end
    """
    wb = openpyxl.load_workbook(cycle_list_path, read_only=True, data_only=True)
    ws = wb["テストサイクル一覧"]
    c = CYCLE_COLS
    days_ahead = CONFIG["reminder_days_ahead"]
    reminders = []

    for row in ws.iter_rows(min_row=CYCLE_HEADER_ROWS + 1, values_only=True):
        if not row:
            continue

        # Reconstruct cycle ID from area + seq_no
        area_val = row[c["area"]]
        seq_val  = row[c["seq_no"]]
        cycle_id = reconstruct_cycle_id(area_val, seq_val)
        if not cycle_id:
            continue

        # Skip sentinel rows
        raw_col0 = str(row[c["cycle_id"]] or "").strip()
        if raw_col0 in ("e", "テストサイクルID", "Test cycle ID"):
            continue

        # Skip deleted (col 9), NOT regression (col 8)
        if row[c["deletion_flag"]] == "X":
            continue

        status     = row[c["exec_status"]]     or STATUS["not_started"]
        plan_start = row[c["plan_start_latest"]]
        plan_end   = row[c["plan_end_latest"]]
        actual_end = row[c["actual_end"]]
        executor   = row[c["executor"]]

        if isinstance(plan_start, datetime): plan_start = plan_start.date()
        if isinstance(plan_end,   datetime): plan_end   = plan_end.date()

        # Skip completed or cancelled
        if status in (STATUS["complete"], STATUS["cancelled"]):
            continue

        reminder_type = None

        if plan_start and 0 <= (plan_start - today).days <= days_ahead:
            reminder_type = "upcoming"
        elif plan_end and plan_end < today and not actual_end:
            reminder_type = "overdue"
        elif status == STATUS["in_progress"] and plan_end and 0 <= (today - plan_end).days <= 5:
            reminder_type = "stalled"

        if reminder_type:
            reminders.append({
                "cycle_id":   cycle_id,
                "cycle_name": row[c["cycle_name"]],
                "executor":   executor,
                "plan_start": plan_start,
                "plan_end":   plan_end,
                "status":     status,
                "type":       reminder_type,
            })

    wb.close()
    return reminders


# ─────────────────────────────────────────────────────────────
# EMAIL VIA LOCAL OUTLOOK (win32com — no API needed)
# ─────────────────────────────────────────────────────────────
def send_email_outlook(to: str, subject: str, body_html: str):
    """
    Send an email using the locally installed Outlook app via win32com.
    No API, no auth — uses your already-logged-in Outlook session.
    """
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To      = to
        mail.Subject = subject
        mail.HTMLBody = body_html
        mail.Send()
        log.info(f"Email sent → {to}")
    except ImportError:
        log.error("win32com not installed. Run: pip install pywin32")
    except Exception as e:
        log.error(f"Failed to send email to {to}: {e}")


def resolve_email(executor_name: str) -> str | None:
    """
    Look up executor email from the name→email mapping in CONFIG.
    In test_mode, always returns the manager email so nothing goes to real executors.
    """
    if not executor_name:
        return None

    # TEST MODE: redirect everything to manager
    if CONFIG.get("test_mode"):
        return CONFIG["manager_email"]

    mapping = CONFIG["executor_emails"]
    # Try exact match first
    if executor_name in mapping:
        return mapping[executor_name]
    # Try partial match (e.g. "Tanaka K." matches "Tanaka")
    for name, email in mapping.items():
        if name.lower() in executor_name.lower():
            return email
    return None


def send_teams_channel(subject: str, body_html: str):
    """
    Post a message to the configured Teams channel via its email address.

    Every Teams channel has a built-in email address (channel → ••• → Get email
    address). Sending an email to that address creates a card in the channel —
    no webhook, no Azure AD app registration, no IT approval needed.

    The same win32com Outlook send is reused, so no extra dependencies are added.
    Messages appear as email cards in the Teams channel feed.
    """
    ch_email = CONFIG.get("teams_channel_email", "").strip()
    if not ch_email:
        log.warning("Teams notification skipped — teams_channel_email is not configured.")
        return
    send_email_outlook(ch_email, subject, body_html)
    log.info(f"Teams channel notified → {ch_email}")


def notify(to_email: str | None, subject: str, body_html: str, *, post_to_teams: bool = False):
    """
    Dispatch a notification to all configured channels.

    Parameters
    ----------
    to_email     : recipient for the "email" channel; pass None to skip email
    subject      : email / Teams card subject line
    body_html    : HTML body (rendered in Outlook and as Teams email card)
    post_to_teams: if True AND "teams" is in notify_channels, also posts to
                   the Teams channel. Use True for manager-level broadcasts;
                   False for per-executor messages that shouldn't flood the channel.

    Channel selection is controlled by CONFIG["notify_channels"].
    A future UI layer will expose this choice per run without changing CONFIG.
    """
    channels = CONFIG.get("notify_channels", ["email"])

    if "email" in channels and to_email:
        send_email_outlook(to_email, subject, body_html)

    if "teams" in channels and post_to_teams:
        send_teams_channel(subject, body_html)


# ─────────────────────────────────────────────────────────────
# REPORT BUILDER
# ─────────────────────────────────────────────────────────────
def build_manager_report(changes: list[dict], reminders: list[dict], today: date) -> str:
    """Build an HTML daily summary email for the manager."""
    completed = [c for c in changes if c["new_status"] == STATUS["complete"]]
    started   = [c for c in changes if c["new_status"] == STATUS["in_progress"]
                                    and c["old_status"] == STATUS["not_started"]]
    overdue   = [r for r in reminders if r["type"] == "overdue"]
    upcoming  = [r for r in reminders if r["type"] == "upcoming"]
    stalled   = [r for r in reminders if r["type"] == "stalled"]

    def table(rows, cols):
        html = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px">'
        html += "<tr>" + "".join(f'<th style="background:#f0f0f0">{c}</th>' for c in cols) + "</tr>"
        for r in rows:
            html += "<tr>" + "".join(f"<td>{v}</td>" for v in r) + "</tr>"
        html += "</table>"
        return html

    html = f"""
    <div style="font-family:Arial,sans-serif;max-width:900px">
    <h2 style="color:#4a4a8a">📊 SAP Test Cycle Daily Report — {today.strftime('%Y/%m/%d')}</h2>
    <p style="font-size:15px">
      <b>✅ Newly completed:</b> {len(completed)} &nbsp;|&nbsp;
      <b>🚀 Newly started:</b> {len(started)} &nbsp;|&nbsp;
      <b>🔴 Overdue:</b> {len(overdue)} &nbsp;|&nbsp;
      <b>⏰ Stalled:</b> {len(stalled)} &nbsp;|&nbsp;
      <b>📅 Upcoming ({CONFIG['reminder_days_ahead']}d):</b> {len(upcoming)}
    </p>
    """

    if completed:
        html += "<h3 style='color:green'>✅ Completed Today</h3>"
        html += table(
            [(c["cycle_id"], c["cycle_name"], c["area"],
              f"{c['completed_steps']}/{c['total_steps']}", c["executor"] or "-")
             for c in completed],
            ["Cycle ID", "Name", "Area", "Steps", "Executor"]
        )

    if overdue:
        html += "<h3 style='color:red'>🔴 Overdue — Action Required</h3>"
        html += table(
            [(r["cycle_id"], r["cycle_name"], str(r["plan_end"]), r["executor"] or "-", r["status"])
             for r in overdue],
            ["Cycle ID", "Name", "Plan End", "Executor", "Status"]
        )

    if stalled:
        html += "<h3 style='color:orange'>⏰ Stalled — Please Follow Up</h3>"
        html += table(
            [(r["cycle_id"], r["cycle_name"], str(r["plan_end"]), r["executor"] or "-")
             for r in stalled],
            ["Cycle ID", "Name", "Plan End", "Executor"]
        )

    if upcoming:
        html += f"<h3 style='color:#4a4a8a'>📅 Starting within {CONFIG['reminder_days_ahead']} days</h3>"
        html += table(
            [(r["cycle_id"], r["cycle_name"], str(r["plan_start"]), r["executor"] or "-")
             for r in upcoming],
            ["Cycle ID", "Name", "Plan Start", "Executor"]
        )

    if started:
        html += "<h3>🚀 Newly Started</h3>"
        html += table(
            [(c["cycle_id"], c["cycle_name"], c["area"],
              f"{c['completed_steps']}/{c['total_steps']}", c["executor"] or "-")
             for c in started],
            ["Cycle ID", "Name", "Area", "Steps", "Executor"]
        )

    html += "<br><p style='color:gray;font-size:11px'>Generated automatically by SAP Test Automation</p></div>"
    return html


def build_executor_reminder(reminder: dict) -> str:
    """Build an HTML reminder email for an individual executor."""
    r = reminder
    if r["type"] == "upcoming":
        heading = f"📅 Upcoming Test Cycle — {r['cycle_id']}"
        msg = f"""Your test cycle <b>{r['cycle_id']} — {r['cycle_name']}</b> is scheduled
        to start on <b>{r['plan_start']}</b>.<br><br>
        Please make sure your test data and SAP environment are ready before the start date.
        Check the condition script file and confirm all pre-conditions are met."""
    elif r["type"] == "overdue":
        heading = f"⚠️ Overdue Test Cycle — {r['cycle_id']}"
        msg = f"""Your test cycle <b>{r['cycle_id']} — {r['cycle_name']}</b> was due by
        <b>{r['plan_end']}</b> and has not been marked complete.<br><br>
        Please update the condition script with your execution results (actual dates + OK/NG)
        as soon as possible. Current status: <b>{r['status']}</b>."""
    else:
        heading = f"🔄 Please Update Test Results — {r['cycle_id']}"
        msg = f"""Test cycle <b>{r['cycle_id']} — {r['cycle_name']}</b> is in progress
        but results have not been fully entered.<br><br>
        Please open the condition script and enter the actual execution date and
        OK/NG result for each completed step."""

    return f"""
    <div style="font-family:Arial,sans-serif;max-width:700px">
    <h2 style="color:#4a4a8a">{heading}</h2>
    <p style="font-size:14px">{msg}</p>
    <p style="font-size:13px;color:gray">
      This is an automated reminder from the SAP Test Management system.<br>
      Please do not reply to this email.
    </p>
    </div>
    """


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def run():
    setup_logging()
    today = date.today()
    log.info(f"{'='*60}")
    log.info(f"SAP Test Automation — {today}")
    log.info(f"{'='*60}")

    folder = CONFIG["folder"]
    if not os.path.exists(folder):
        log.error(f"Folder not found: {folder}")
        log.error("Check the 'folder' path in CONFIG and make sure OneDrive is synced.")
        return

    # 1. Find files
    cycle_list_path, script_paths = find_files(folder)

    if not cycle_list_path:
        log.error("Could not find the Test Cycle List file. Check 'cycle_list_pattern' in CONFIG.")
        return

    if not script_paths:
        log.error("No condition script files found. Check 'condition_script_pattern' in CONFIG.")
        return

    # 2. Analyze all condition scripts
    log.info(f"Analyzing {len(script_paths)} condition scripts...")
    script_results = {}
    for path in script_paths:
        result = analyze_script(path)
        if result:
            script_results[result["cycle_id"]] = result
            status_str = "✅ DONE" if result["all_ok"] else f"{result['completed_steps']}/{result['total_steps']} steps"
            log.info(f"  {result['cycle_id']}: {status_str} {'⚠️ HAS NG' if result['has_ng'] else ''}")

    log.info(f"Successfully analyzed: {len(script_results)} scripts")

    # 3. Update cycle list
    log.info("Updating Test Cycle List...")
    changes = update_cycle_list(cycle_list_path, script_results)
    # OneDrive will auto-sync the updated file to SharePoint

    # 4. Get reminders
    reminders = get_reminders(cycle_list_path, today)
    log.info(f"Reminders: {len(reminders)} cycles need attention")

    # 5. Send individual executor reminders
    sent_to = set()
    for reminder in reminders:
        executor_name = reminder.get("executor")
        if not executor_name:
            continue

        # Handle multiple executors in one field (e.g. "Tanaka, Suzuki")
        for name in [n.strip() for n in str(executor_name).split(",")]:
            email = resolve_email(name)
            if not email:
                log.warning(f"No email mapping for executor: '{name}' — add to CONFIG['executor_emails']")
                continue

            # Avoid duplicate emails to same person
            reminder_key = f"{email}_{reminder['cycle_id']}_{reminder['type']}"
            if reminder_key in sent_to:
                continue
            sent_to.add(reminder_key)

            subject = f"[SAP Test{'  TEST MODE' if CONFIG.get('test_mode') else ''}] {reminder['type'].upper()} — {reminder['cycle_id']} (executor: {name})"
            body    = build_executor_reminder(reminder)
            # post_to_teams=False: executor reminders go to email only to avoid
            # flooding the shared Teams channel with per-person messages.
            notify(email, subject, body, post_to_teams=False)

    # 6. Send daily summary to manager (email + Teams channel)
    report_html = build_manager_report(changes, reminders, today)
    notify(
        CONFIG["manager_email"],
        f"[SAP Test] Daily Report {today.strftime('%Y/%m/%d')} — "
        f"{len([c for c in changes if c['new_status'] == STATUS['complete']])} completed, "
        f"{len([r for r in reminders if r['type'] == 'overdue'])} overdue",
        report_html,
        post_to_teams=True,   # manager summary always goes to the Teams channel
    )

    log.info("=== Run complete ===")
    log.info(f"Changes: {len(changes)} | Reminders sent: {len(sent_to)} | Report: sent to {CONFIG['manager_email']}")


if __name__ == "__main__":
    run()
