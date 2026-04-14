"""
AT Sentinel — Streamlit UI
===========================
Run with:  streamlit run app.py
Requires:  pip install streamlit openpyxl pywin32
"""

import json
import streamlit as st
from datetime import date, datetime
from pathlib import Path

st.set_page_config(page_title="AT Sentinel", page_icon="🛡️", layout="wide")

import AT_Sentinel as sentinel
from AT_Sentinel import (
    find_files, analyze_script, update_cycle_list,
    get_reminders, notify, build_manager_report,
    build_executor_reminder, resolve_email, STATUS, CONFIG,
)

RUNS_LOG = Path(__file__).parent / "runs_log.json"


def _append_run_log(entry: dict) -> None:
    """Append a run-summary entry to the local JSON log file."""
    history: list = []
    if RUNS_LOG.exists():
        try:
            with open(RUNS_LOG, "r", encoding="utf-8") as f:
                history = json.load(f)
        except Exception:
            history = []
    history.append(entry)
    try:
        with open(RUNS_LOG, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2, default=str)
    except Exception as exc:
        st.warning(f"Could not write run history: {exc}")


# ── Sidebar — Configuration ────────────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ Configuration")

    folder        = st.text_input("Test Files Folder", value=CONFIG["folder"])
    manager_email = st.text_input("Manager Email",     value=CONFIG["manager_email"])

    st.divider()
    st.subheader("Notifications")
    col_a, col_b  = st.columns(2)
    email_on      = col_a.checkbox("✉️ Email", value="email" in CONFIG["notify_channels"])
    teams_on      = col_b.checkbox("💬 Teams", value="teams" in CONFIG["notify_channels"])
    teams_ch_email = st.text_input(
        "Teams Channel Email",
        value=CONFIG["teams_channel_email"],
        disabled=not teams_on,
        help="Channel → ••• → Get email address",
    )

    st.divider()
    st.subheader("Options")
    test_mode     = st.toggle(
        "🧪 Test Mode", value=CONFIG["test_mode"],
        help="Redirects all executor emails to manager email",
    )
    dry_run       = st.toggle(
        "🔵 Dry Run", value=CONFIG.get("dry_run", False),
        help="Compute all changes but do NOT write to Excel or send notifications",
    )
    reminder_days = st.slider("Remind N days ahead", 1, 7, value=CONFIG["reminder_days_ahead"])

    st.divider()
    st.subheader("Executor Emails")
    st.caption("Add or edit executor name → email mappings used for reminders.")

    # Seed session state from CONFIG on first load only
    if "exec_emails" not in st.session_state:
        st.session_state.exec_emails = [
            {"Name": k, "Email": v}
            for k, v in CONFIG["executor_emails"].items()
        ]

    exec_emails_edited = st.data_editor(
        st.session_state.exec_emails,
        num_rows="dynamic",
        column_config={
            "Name":  st.column_config.TextColumn("Executor Name", required=True),
            "Email": st.column_config.TextColumn("Email Address", required=True),
        },
        use_container_width=True,
        key="exec_emails_editor",
    )
    # Persist current edits for next rerun
    st.session_state.exec_emails = exec_emails_edited

# ── Apply sidebar values to CONFIG ────────────────────────────────────────────
new_exec_emails = {
    row["Name"]: row["Email"]
    for row in (exec_emails_edited or [])
    if str(row.get("Name") or "").strip() and str(row.get("Email") or "").strip()
}
sentinel.CONFIG.update({
    "folder":              folder,
    "manager_email":       manager_email,
    "test_mode":           test_mode,
    "dry_run":             dry_run,
    "reminder_days_ahead": reminder_days,
    "notify_channels":     [ch for ch, on in [("email", email_on), ("teams", teams_on)] if on],
    "teams_channel_email": teams_ch_email,
    "executor_emails":     new_exec_emails,
})

# ── Main area ──────────────────────────────────────────────────────────────────
st.title("🛡️ AT Sentinel")
st.caption("SAP UAT Test Cycle Automation — daily sync, reminders & reporting")

if test_mode:
    st.info("🧪 **Test mode ON** — all executor emails redirected to manager address", icon="ℹ️")
if dry_run:
    st.info(
        "🔵 **Dry Run ON** — changes will be computed but NOT written to Excel or sent as notifications",
        icon="ℹ️",
    )

run_clicked = st.button("▶ Run Now", type="primary", use_container_width=True)

if run_clicked:
    sentinel.setup_logging()
    today = date.today()

    # Pre-initialise so all vars remain in scope after the status block
    changes         = []
    missing_scripts = []
    reminders       = []
    overdue         = []
    stalled         = []
    upcoming        = []

    with st.status("Running AT Sentinel…", expanded=True) as status_widget:

        # 1 ── Discover files
        st.write("📁 Scanning folder for test files…")
        cycle_list_path, script_paths = find_files(folder)

        if not cycle_list_path:
            status_widget.update(label="Failed — cycle list not found", state="error")
            st.error("Cycle list file not found. Check **folder path** and **cycle_list_pattern** in CONFIG.")
            st.stop()

        if not script_paths:
            status_widget.update(label="Failed — no condition scripts found", state="error")
            st.error("No condition script files found. Check **condition_script_pattern** in CONFIG.")
            st.stop()

        st.write(f"✅ Found **{len(script_paths)}** condition script(s)")

        # 2 ── Analyze scripts
        st.write("🔍 Analyzing condition scripts…")
        script_results = {}
        for path in script_paths:
            result = analyze_script(path)
            if result:
                script_results[result["cycle_id"]] = result
        st.write(f"✅ Successfully analyzed **{len(script_results)}** script(s)")

        # 3 ── Update cycle list
        st.write("📝 Updating Test Cycle List…")
        changes, missing_scripts = update_cycle_list(cycle_list_path, script_results)
        save_note = " (dry run — not saved)" if dry_run else " — OneDrive will sync automatically"
        st.write(f"✅ Updated **{len(changes)}** row(s){save_note}")
        if missing_scripts:
            st.write(f"⚠️ **{len(missing_scripts)}** active cycle(s) have no matching script file")

        # 4 ── Reminders
        st.write("🔔 Detecting reminders…")
        reminders = get_reminders(cycle_list_path, today)
        overdue   = [r for r in reminders if r["type"] == "overdue"]
        stalled   = [r for r in reminders if r["type"] == "stalled"]
        upcoming  = [r for r in reminders if r["type"] == "upcoming"]
        st.write(f"✅ {len(overdue)} overdue · {len(stalled)} stalled · {len(upcoming)} upcoming")

        # 5 ── Notifications
        active_channels = sentinel.CONFIG["notify_channels"]
        if not active_channels:
            st.write("⚠️ No notification channels selected — skipping")
        elif dry_run:
            st.write("🔵 Dry run — notifications skipped")
        else:
            st.write(f"📨 Sending via: **{', '.join(active_channels)}**…")
            sent_keys: set[str] = set()
            for reminder in reminders:
                executor_name = reminder.get("executor")
                if not executor_name:
                    continue
                for name in [n.strip() for n in str(executor_name).split(",")]:
                    email_addr = resolve_email(name)
                    if not email_addr:
                        continue
                    key = f"{email_addr}_{reminder['cycle_id']}_{reminder['type']}"
                    if key in sent_keys:
                        continue
                    sent_keys.add(key)
                    subject = (
                        f"[SAP Test{'  TEST MODE' if test_mode else ''}] "
                        f"{reminder['type'].upper()} — {reminder['cycle_id']} (executor: {name})"
                    )
                    notify(email_addr, subject, build_executor_reminder(reminder), post_to_teams=False)

            report_html   = build_manager_report(changes, reminders, today, missing_scripts)
            completed_cnt = len([c for c in changes if c["new_status"] == STATUS["complete"]])
            notify(
                manager_email,
                f"[SAP Test] Daily Report {today.strftime('%Y/%m/%d')} — "
                f"{completed_cnt} completed, {len(overdue)} overdue",
                report_html,
                post_to_teams=True,
            )
            st.write("✅ Notifications dispatched")

        run_label = (
            f"🔵 Dry Run complete — {today.strftime('%Y/%m/%d %H:%M')} — no changes written"
            if dry_run else
            f"✅ Run complete — {today.strftime('%Y/%m/%d %H:%M')}"
        )
        status_widget.update(label=run_label, state="complete")

    # ── Compute summary row lists (after status block — changes/reminders are final) ──
    completed_rows = [c for c in changes if c["new_status"] == STATUS["complete"]]
    started_rows   = [c for c in changes if c["new_status"] == STATUS["in_progress"]
                                         and c["old_status"] == STATUS["not_started"]]

    # ── Persist run to history log ─────────────────────────────────────────────
    _append_run_log({
        "timestamp":       datetime.now().isoformat(timespec="seconds"),
        "completed":       len(completed_rows),
        "started":         len(started_rows),
        "overdue":         len(overdue),
        "stalled":         len(stalled),
        "upcoming":        len(upcoming),
        "missing_scripts": len(missing_scripts),
        "dry_run":         dry_run,
    })

    # ── Dry Run banner ─────────────────────────────────────────────────────────
    if dry_run:
        st.warning("🔵 **Dry Run — no changes written to Excel, no notifications sent.**")

    # ── Summary metrics ────────────────────────────────────────────────────────
    st.subheader("Summary")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("✅ Completed", len(completed_rows))
    m2.metric("🚀 Started",   len(started_rows))
    m3.metric("🔴 Overdue",   len(overdue))
    m4.metric("⏰ Stalled",   len(stalled))
    m5.metric("📅 Upcoming",  len(upcoming))

    # ── Detail tables ──────────────────────────────────────────────────────────
    def _show_table(rows: list, col_map: dict, label: str) -> None:
        if not rows:
            return
        with st.expander(label):
            st.dataframe(
                [{display: r.get(key, "-") for display, key in col_map.items()} for r in rows],
                use_container_width=True,
            )

    _show_table(completed_rows, {
        "Cycle ID": "cycle_id", "Name": "cycle_name", "Area": "area",
        "Executor": "executor",
    }, f"✅ Completed ({len(completed_rows)})")

    _show_table(overdue, {
        "Cycle ID": "cycle_id", "Name": "cycle_name",
        "Plan End": "plan_end", "Executor": "executor", "Status": "status",
    }, f"🔴 Overdue ({len(overdue)})")

    _show_table(stalled, {
        "Cycle ID": "cycle_id", "Name": "cycle_name",
        "Plan End": "plan_end", "Executor": "executor",
    }, f"⏰ Stalled ({len(stalled)})")

    _show_table(upcoming, {
        "Cycle ID": "cycle_id", "Name": "cycle_name",
        "Plan Start": "plan_start", "Executor": "executor",
    }, f"📅 Upcoming ({len(upcoming)})")

    _show_table(started_rows, {
        "Cycle ID": "cycle_id", "Name": "cycle_name",
        "Area": "area", "Executor": "executor",
    }, f"🚀 Newly Started ({len(started_rows)})")

    if missing_scripts:
        _show_table(missing_scripts, {
            "Cycle ID": "cycle_id", "Name": "cycle_name", "Area": "area",
            "Current Status": "exec_status", "Executor": "executor",
        }, f"⚠️ No Script Found ({len(missing_scripts)})")

# ── Run History (always visible at page bottom) ────────────────────────────────
st.divider()
with st.expander("📋 Run History"):
    if RUNS_LOG.exists():
        try:
            with open(RUNS_LOG, "r", encoding="utf-8") as f:
                history = json.load(f)
            if history:
                st.dataframe(list(reversed(history)), use_container_width=True)
            else:
                st.info("No runs recorded yet.")
        except Exception as exc:
            st.error(f"Could not load run history: {exc}")
    else:
        st.info("No runs recorded yet.")
