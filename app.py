"""
AT Sentinel — Streamlit UI
===========================
Run with:  streamlit run app.py
Requires:  pip install -r requirements.txt
"""

import json
import pandas as pd
import streamlit as st
from datetime import date, datetime
from pathlib import Path

st.set_page_config(page_title="AT Sentinel", page_icon="🛡️", layout="wide")

import AT_Sentinel as sentinel
from AT_Sentinel import (
    find_files, analyze_script, update_cycle_list,
    get_reminders, notify, build_manager_report,
    build_executor_reminder, resolve_email,
    load_profiles, save_profiles,
    read_cycle_data, build_daywise_report,
    build_streamwise_report, build_executorwise_report,
    build_cycle_step_report, build_merged_condition_report,
    get_executor_reminder_data, build_review_reminder, build_remark_reminder,
    STATUS, CONFIG, DEFAULT_PROFILES,
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


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ Configuration")

    # ── 1. Test stage profile ──────────────────────────────────────────────────
    st.subheader("Test Stage")
    profiles = load_profiles()
    profile_names = list(profiles.keys())
    default_idx = profile_names.index(CONFIG.get("active_profile", "AT-SAP")) \
                  if CONFIG.get("active_profile", "AT-SAP") in profile_names else 0

    active_profile_name = st.selectbox(
        "Active Profile",
        options=profile_names,
        index=default_idx,
        help="Select the test stage. Column indices and patterns are loaded from this profile.",
    )
    sentinel.CONFIG["active_profile"] = active_profile_name
    active_profile = profiles[active_profile_name]

    if not active_profile.get("configured", True):
        st.warning(
            f"⚠️ **{active_profile_name}** uses AT-SAP column defaults. "
            "Verify column indices below before running.",
            icon="⚠️",
        )

    # ── 1a. Profile editor ─────────────────────────────────────────────────────
    with st.expander("✏️ Edit / Add Profiles"):
        col_new, col_del = st.columns(2)
        new_name = col_new.text_input("New profile name", key="new_profile_name",
                                      placeholder="e.g. SIT-Phase2")
        if col_new.button("➕ Add", use_container_width=True):
            if new_name and new_name not in profiles:
                profiles[new_name] = {
                    **DEFAULT_PROFILES["AT-SAP"],
                    "display_name": new_name,
                    "configured": False,
                    "cycle_cols": dict(DEFAULT_PROFILES["AT-SAP"]["cycle_cols"]),
                    "script_cols": dict(DEFAULT_PROFILES["AT-SAP"]["script_cols"]),
                }
                save_profiles(profiles)
                sentinel.CONFIG["active_profile"] = new_name
                st.rerun()
        if col_del.button("🗑️ Delete", use_container_width=True,
                          disabled=active_profile_name in DEFAULT_PROFILES):
            del profiles[active_profile_name]
            save_profiles(profiles)
            sentinel.CONFIG["active_profile"] = "AT-SAP"
            st.rerun()
        if active_profile_name in DEFAULT_PROFILES:
            st.caption("Built-in profiles cannot be deleted (only edited).")

        st.divider()

        with st.form(f"profile_form_{active_profile_name}"):
            p = active_profile

            st.markdown("**Basic settings**")
            new_display   = st.text_input("Display Name",
                                          value=p.get("display_name", active_profile_name))
            new_cl_pat    = st.text_input("Cycle List Pattern (regex)",
                                          value=p.get("cycle_list_pattern", ""))
            new_scr_pat   = st.text_input("Script Pattern (regex)",
                                          value=p.get("condition_script_pattern", ""))
            c1, c2        = st.columns(2)
            new_cl_sheet  = c1.text_input("Cycle List Sheet",
                                          value=p.get("cycle_sheet_name", ""))
            new_scr_sheet = c2.text_input("Script Sheet",
                                          value=p.get("script_sheet_name", ""))
            c1, c2, c3, c4 = st.columns(4)
            new_cl_hdr    = c1.number_input("Cycle header rows",
                                            value=p.get("cycle_header_rows", 3),
                                            min_value=0, step=1)
            new_scr_hdr   = c2.number_input("Script header rows",
                                            value=p.get("script_header_rows", 4),
                                            min_value=0, step=1)
            new_prefix    = c3.text_input("Cycle ID prefix",
                                          value=p.get("cycle_id_prefix", "AT1TC"))
            new_slash     = c4.checkbox("Area slash (SDFI→SD/FI)",
                                        value=p.get("cycle_id_area_slash", True))

            # Cycle list columns
            st.markdown("**Cycle list columns** (0-indexed)")
            _CC_LABELS = {
                "cycle_id":          "Cycle ID (formula col)",
                "seq_no":            "Seq No",
                "cycle_name":        "Cycle Name",
                "area":              "Area",
                "regression_flag":   "Regression Flag",
                "deletion_flag":     "Deletion Flag",
                "plan_start_latest": "Plan Start (latest)",
                "plan_end_latest":   "Plan End (latest)",
                "actual_start":      "Actual Start",
                "actual_end":        "Actual End",
                "review_end":        "Review Completion Date",
                "review_plan_end":   "Review Plan End (latest)",
                "executor":          "Executor",
                "exec_status":       "Exec Status",
                "total_steps":       "Total Steps",
                "complete_steps":    "Complete Steps",
            }
            cc_rows = [
                {"Column": lbl, "_key": k, "Index": p["cycle_cols"].get(k, 0)}
                for k, lbl in _CC_LABELS.items()
            ]
            cc_edited = st.data_editor(
                cc_rows,
                column_config={
                    "Column": st.column_config.TextColumn("Column", disabled=True),
                    "_key":   st.column_config.TextColumn(disabled=True),
                    "Index":  st.column_config.NumberColumn("Col index", min_value=0, step=1),
                },
                hide_index=True,
                use_container_width=True,
                key=f"cc_editor_{active_profile_name}",
            )

            # Script columns
            st.markdown("**Script columns** (0-indexed)")
            _SC_LABELS = {
                "excluded":       "Excluded flag",
                "exec_pic":       "Executor name",
                "planned_date":   "Planned date",
                "actual_date":    "Actual date",
                "result":         "Result (OK/NG/-)",
                "retest_pic":     "Retest executor",
                "retest_planned": "Retest planned date",
                "retest_date":    "Retest actual date",
                "retest_result":  "Retest result",
                "review_pic":     "Review executor",
                "review_date":    "Review actual date",
                "review_result":  "Review result (OK/NG/-)",
            }
            sc_rows = [
                {"Column": lbl, "_key": k, "Index": p["script_cols"].get(k, 0)}
                for k, lbl in _SC_LABELS.items()
            ]
            sc_edited = st.data_editor(
                sc_rows,
                column_config={
                    "Column": st.column_config.TextColumn("Column", disabled=True),
                    "_key":   st.column_config.TextColumn(disabled=True),
                    "Index":  st.column_config.NumberColumn("Col index", min_value=0, step=1),
                },
                hide_index=True,
                use_container_width=True,
                key=f"sc_editor_{active_profile_name}",
            )

            if st.form_submit_button("💾 Save Profile", use_container_width=True):
                new_cc = {row["_key"]: int(row["Index"]) for row in cc_edited}
                new_sc = {row["_key"]: int(row["Index"]) for row in sc_edited}
                profiles[active_profile_name] = {
                    **p,
                    "display_name":             new_display,
                    "cycle_list_pattern":        new_cl_pat,
                    "condition_script_pattern":  new_scr_pat,
                    "cycle_sheet_name":          new_cl_sheet,
                    "script_sheet_name":         new_scr_sheet,
                    "cycle_header_rows":         int(new_cl_hdr),
                    "script_header_rows":        int(new_scr_hdr),
                    "cycle_id_prefix":           new_prefix,
                    "cycle_id_area_slash":        new_slash,
                    "configured":                True,
                    "cycle_cols":                new_cc,
                    "script_cols":               new_sc,
                }
                save_profiles(profiles)
                st.success(f"Profile '{active_profile_name}' saved.")
                st.rerun()

    st.divider()

    # ── 2. Folder & email ──────────────────────────────────────────────────────
    folder        = st.text_input("Test Files Folder", value=CONFIG["folder"])
    manager_email = st.text_input("Manager Email",     value=CONFIG["manager_email"])

    st.divider()

    # ── 3. Notifications ───────────────────────────────────────────────────────
    st.subheader("Notifications")
    col_a, col_b   = st.columns(2)
    email_on       = col_a.checkbox("✉️ Email", value="email" in CONFIG["notify_channels"])
    teams_on       = col_b.checkbox("💬 Teams", value="teams" in CONFIG["notify_channels"])
    teams_ch_email = st.text_input(
        "Teams Channel Email",
        value=CONFIG["teams_channel_email"],
        disabled=not teams_on,
        help="Channel → ••• → Get email address",
    )

    st.divider()

    # ── 4. Options ─────────────────────────────────────────────────────────────
    st.subheader("Options")
    test_mode     = st.toggle("🧪 Test Mode",  value=CONFIG["test_mode"],
                              help="Redirects all executor emails to manager email")
    dry_run       = st.toggle("🔵 Dry Run",    value=CONFIG.get("dry_run", False),
                              help="Compute all changes but do NOT write to Excel or send notifications")
    reminder_days = st.slider("Remind N days ahead", 1, 7, value=CONFIG["reminder_days_ahead"])

    st.divider()

    # ── 5. Notification time slot ──────────────────────────────────────────────
    st.subheader("Notification Time Slot")
    time_slot = st.radio(
        "When are you running?",
        options=["morning", "midday", "evening"],
        format_func=lambda x: {
            "morning": "Morning  9:00 — overdue + starting today",
            "midday":  "Midday  12:00 — overdue + starting today",
            "evening": "Evening 17:00 — above + starting/due tomorrow",
        }[x],
        index=0,
        horizontal=False,
    )

    st.divider()

    # ── 6. Executor emails ─────────────────────────────────────────────────────
    st.subheader("Executor Emails")
    st.caption("Executor name → email mappings used for reminder dispatch.")

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

# Status banners
if test_mode:
    st.info("🧪 **Test mode ON** — all executor emails redirected to manager address", icon="ℹ️")
if dry_run:
    st.info(
        "🔵 **Dry Run ON** — changes will be computed but NOT written to Excel or sent as notifications",
        icon="ℹ️",
    )
if not active_profile.get("configured", True):
    st.warning(
        f"⚠️ Profile **{active_profile_name}** uses AT-SAP column defaults — "
        "open the sidebar profile editor to verify column indices before running.",
        icon="⚠️",
    )

# Resolve profile once — shared by both tabs
run_profile = load_profiles().get(
    active_profile_name,
    DEFAULT_PROFILES.get(active_profile_name, DEFAULT_PROFILES["AT-SAP"]),
)

tab_run, tab_reminders, tab_reports = st.tabs(["▶ Run", "🔔 Reminders", "📊 Reports"])

# ── RUN TAB ───────────────────────────────────────────────────────────────────
with tab_run:
    run_clicked = st.button("▶ Run Now", type="primary", use_container_width=True)

    if run_clicked:
        sentinel.setup_logging()
        today = date.today()

        changes         = []
        missing_scripts = []
        reminders       = []

        with st.status("Running AT Sentinel…", expanded=True) as status_widget:

            # 1 — Discover files
            st.write(f"📁 Scanning folder (profile: **{active_profile_name}**)…")
            cycle_list_path, script_paths = find_files(folder, profile=run_profile)

            if not cycle_list_path:
                status_widget.update(label="Failed — cycle list not found", state="error")
                st.error("Cycle list file not found. Check **folder path** and the profile's **cycle_list_pattern**.")
                st.stop()

            if not script_paths:
                status_widget.update(label="Failed — no condition scripts found", state="error")
                st.error("No condition script files found. Check the profile's **condition_script_pattern**.")
                st.stop()

            # Store for Reports and Reminders tabs
            st.session_state["last_cycle_list_path"] = cycle_list_path
            st.session_state["last_script_paths"]    = script_paths
            st.write(f"✅ Found **{len(script_paths)}** condition script(s)")

            # 2 — Analyze scripts
            st.write("🔍 Analyzing condition scripts…")
            script_results = {}
            for path in script_paths:
                result = analyze_script(path, profile=run_profile)
                if result:
                    script_results[result["cycle_id"]] = result
            st.write(f"✅ Successfully analyzed **{len(script_results)}** script(s)")

            # 3 — Update cycle list
            st.write("📝 Updating Test Cycle List…")
            changes, missing_scripts = update_cycle_list(
                cycle_list_path, script_results, profile=run_profile
            )
            save_note = " (dry run — not saved)" if dry_run else " — OneDrive will sync automatically"
            st.write(f"✅ Updated **{len(changes)}** row(s){save_note}")
            if missing_scripts:
                st.write(f"⚠️ **{len(missing_scripts)}** active cycle(s) have no matching script file")

            # 4 — Reminders
            st.write(f"🔔 Detecting reminders (slot: **{time_slot}**)…")
            reminders = get_reminders(cycle_list_path, today, profile=run_profile, time_slot=time_slot)
            overdue        = [r for r in reminders if r["type"] == "overdue"]
            starting_today = [r for r in reminders if r["type"] == "starting_today"]
            st_tomorrow    = [r for r in reminders if r["type"] == "starting_tomorrow"]
            due_tomorrow   = [r for r in reminders if r["type"] == "due_tomorrow"]
            st.write(
                f"✅ {len(overdue)} overdue · {len(starting_today)} starting today"
                + (f" · {len(st_tomorrow)} starting tomorrow · {len(due_tomorrow)} due tomorrow"
                   if time_slot == "evening" else "")
            )

            # 5 — Notifications
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
                        notify(email_addr, subject, build_executor_reminder(reminder),
                               post_to_teams=False)

                report_html   = build_manager_report(changes, reminders, today, missing_scripts)
                completed_cnt = len([c for c in changes if c["new_status"] == STATUS["complete"]])
                notify(
                    manager_email,
                    f"[SAP Test/{active_profile_name}] Daily Report {today.strftime('%Y/%m/%d')} — "
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

        # ── Summary metrics ────────────────────────────────────────────────────
        completed_rows = [c for c in changes if c["new_status"] == STATUS["complete"]]
        in_review_rows = [c for c in changes if c["new_status"] == STATUS["reviewing"]]
        started_rows   = [c for c in changes if c["new_status"] == STATUS["in_progress"]
                                             and c["old_status"] == STATUS["not_started"]]

        _append_run_log({
            "timestamp":       datetime.now().isoformat(timespec="seconds"),
            "profile":         active_profile_name,
            "time_slot":       time_slot,
            "completed":       len(completed_rows),
            "in_review":       len(in_review_rows),
            "started":         len(started_rows),
            "overdue":         len(overdue),
            "missing_scripts": len(missing_scripts),
            "dry_run":         dry_run,
        })

        if dry_run:
            st.warning("🔵 **Dry Run — no changes written to Excel, no notifications sent.**")

        st.subheader("Summary")
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("✅ Completed",      len(completed_rows))
        m2.metric("🔍 In Review",      len(in_review_rows))
        m3.metric("🚀 Started",        len(started_rows))
        m4.metric("🔴 Overdue",        len(overdue))
        m5.metric("🟢 Starting Today", len(starting_today))
        m6.metric("⚠️ No Script",     len(missing_scripts))

        # ── Detail tables ──────────────────────────────────────────────────────
        def _show_table(rows: list, col_map: dict, label: str) -> None:
            if not rows:
                return
            with st.expander(label):
                st.dataframe(
                    [{display: r.get(key, "-") for display, key in col_map.items()} for r in rows],
                    use_container_width=True,
                )

        _show_table(completed_rows, {
            "Cycle ID": "cycle_id", "Name": "cycle_name", "Area": "area", "Executor": "executor",
        }, f"✅ Completed ({len(completed_rows)})")

        _show_table(in_review_rows, {
            "Cycle ID": "cycle_id", "Name": "cycle_name", "Area": "area", "Executor": "executor",
        }, f"🔍 Moved to In Review ({len(in_review_rows)})")

        _show_table(overdue, {
            "Cycle ID": "cycle_id", "Name": "cycle_name",
            "Plan End": "plan_end", "Executor": "executor", "Status": "status",
        }, f"🔴 Overdue ({len(overdue)})")

        _show_table(due_tomorrow, {
            "Cycle ID": "cycle_id", "Name": "cycle_name",
            "Plan End": "plan_end", "Executor": "executor", "Status": "status",
        }, f"⚡ Due Tomorrow ({len(due_tomorrow)})")

        _show_table(starting_today, {
            "Cycle ID": "cycle_id", "Name": "cycle_name",
            "Plan Start": "plan_start", "Executor": "executor",
        }, f"🟢 Starting Today ({len(starting_today)})")

        _show_table(st_tomorrow, {
            "Cycle ID": "cycle_id", "Name": "cycle_name",
            "Plan Start": "plan_start", "Executor": "executor",
        }, f"📅 Starting Tomorrow ({len(st_tomorrow)})")

        _show_table(started_rows, {
            "Cycle ID": "cycle_id", "Name": "cycle_name", "Area": "area", "Executor": "executor",
        }, f"🚀 Newly Started ({len(started_rows)})")

        if missing_scripts:
            _show_table(missing_scripts, {
                "Cycle ID": "cycle_id", "Name": "cycle_name", "Area": "area",
                "Current Status": "exec_status", "Executor": "executor",
            }, f"⚠️ No Script Found ({len(missing_scripts)})")

# ── REMINDERS TAB ─────────────────────────────────────────────────────────────
with tab_reminders:
    st.subheader("🔔 Send Reminders")
    st.caption("Send targeted reminders to individual executors for execution and/or review tasks.")

    rem_cycle_path = st.session_state.get("last_cycle_list_path")
    if not rem_cycle_path:
        st.info("Run a sync first, or click **Load Data** to scan the folder.")
        if st.button("📁 Load Data", key="load_reminders"):
            with st.spinner("Scanning folder…"):
                cl, sps = find_files(folder, profile=run_profile)
                if cl:
                    st.session_state["last_cycle_list_path"] = cl
                    st.session_state["last_script_paths"]    = sps
                    st.rerun()
                else:
                    st.error("Cycle list not found.")

    if rem_cycle_path:
        try:
            rem_data = get_executor_reminder_data(rem_cycle_path, date.today(),
                                                   profile=run_profile)
        except Exception as e:
            st.error(f"Could not load reminder data: {e}")
            rem_data = {}

        if not rem_data:
            st.success("No pending reminders — all cycles are on track.")
        else:
            # ── Summary table ──
            summary = [
                {"Executor":        name,
                 "Exec Pending":    len(v["exec_cycles"]),
                 "Review Pending":  len(v["review_cycles"])}
                for name, v in sorted(rem_data.items())
            ]
            st.dataframe(summary, use_container_width=True)

            st.divider()

            # ── Executor selector ──
            all_executors = list(rem_data.keys())
            selected = st.multiselect(
                "Select executors to send reminders to:",
                options=all_executors,
                default=all_executors,
                help="Choose specific executors or keep all selected.",
            )

            # ── Reminder type ──
            c1, c2, c3 = st.columns(3)
            send_exec   = c1.checkbox("Execution reminders",  value=True)
            send_review = c2.checkbox("Review reminders",     value=True)
            send_remark = c3.checkbox("Remark filling reminder", value=False)

            st.divider()

            # ── Preview ──
            if selected:
                with st.expander("📋 Preview — cycles that will be notified"):
                    for name in selected:
                        v = rem_data[name]
                        st.markdown(f"**{name}**")
                        if send_exec and v["exec_cycles"]:
                            st.markdown("*Execution:*")
                            for cyc in v["exec_cycles"]:
                                st.markdown(
                                    f"  - {cyc['cycle_id']} | {cyc['exec_status']} "
                                    f"| Plan End: {cyc['plan_end'] or '—'}")
                        if send_review and v["review_cycles"]:
                            st.markdown("*Review:*")
                            for cyc in v["review_cycles"]:
                                st.markdown(
                                    f"  - {cyc['cycle_id']} | Review Plan End: "
                                    f"{cyc.get('review_plan_end') or '—'}")

            # ── Send buttons ──
            col_sel, col_all = st.columns(2)
            send_selected = col_sel.button("📨 Send to Selected", type="primary",
                                            disabled=not selected or dry_run)
            send_all      = col_all.button("📨 Send to All",
                                            disabled=dry_run)

            if dry_run:
                st.warning("🔵 Dry Run ON — no emails will be sent.")

            def _do_send(targets: list[str]):
                sent = 0
                for name in targets:
                    email_addr = resolve_email(name)
                    if not email_addr:
                        st.warning(f"No email mapping for '{name}' — skipping.")
                        continue
                    v = rem_data.get(name, {})
                    if send_exec:
                        for cyc in v.get("exec_cycles", []):
                            from AT_Sentinel import build_executor_reminder as _ber
                            reminder_dict = {
                                "type": "overdue" if (cyc["plan_end"] and cyc["plan_end"] < date.today())
                                         else "starting_today",
                                "cycle_id": cyc["cycle_id"], "cycle_name": cyc["cycle_name"],
                                "plan_start": cyc["plan_start"], "plan_end": cyc["plan_end"],
                                "status": cyc["exec_status"],
                            }
                            notify(email_addr,
                                   f"[SAP Test] Execution Reminder — {cyc['cycle_id']} ({name})",
                                   _ber(reminder_dict), post_to_teams=False)
                            sent += 1
                    if send_review:
                        for cyc in v.get("review_cycles", []):
                            notify(email_addr,
                                   f"[SAP Test] Review Reminder — {cyc['cycle_id']} ({name})",
                                   build_review_reminder(cyc), post_to_teams=False)
                            sent += 1
                    if send_remark:
                        all_cycs = v.get("exec_cycles", []) + v.get("review_cycles", [])
                        for cyc in all_cycs:
                            notify(email_addr,
                                   f"[SAP Test] Remark Filling — {cyc['cycle_id']} ({name})",
                                   build_remark_reminder(cyc), post_to_teams=False)
                            sent += 1
                return sent

            if send_selected and selected:
                with st.spinner("Sending reminders…"):
                    n = _do_send(selected)
                st.success(f"✅ Sent {n} reminder(s) to {len(selected)} executor(s).")

            if send_all:
                with st.spinner("Sending reminders to all…"):
                    n = _do_send(all_executors)
                st.success(f"✅ Sent {n} reminder(s) to {len(all_executors)} executor(s).")


# ── REPORTS TAB ───────────────────────────────────────────────────────────────
with tab_reports:
    st.subheader("📊 Reports")

    report_cycle_path  = st.session_state.get("last_cycle_list_path")
    report_script_paths = st.session_state.get("last_script_paths", [])

    if not report_cycle_path:
        st.info("Run a sync first, or click **Load Data** to scan the configured folder.")
        if st.button("📁 Load Data", key="load_reports"):
            with st.spinner("Scanning folder for cycle list…"):
                cl, sps = find_files(folder, profile=run_profile)
                if cl:
                    st.session_state["last_cycle_list_path"] = cl
                    st.session_state["last_script_paths"]    = sps
                    st.rerun()
                else:
                    st.error("Cycle list not found. Check folder path and profile settings.")

    if report_cycle_path:
        st.caption(f"Data source: `{report_cycle_path}`")

        # ── Dropdown selector (spec: "dropdown to select which data to see") ──
        report_choice = st.selectbox(
            "Select report:",
            options=[
                "📋 Cycle & Step Basis",
                "📄 Merged Condition File",
                "📅 Execution Status — Daywise",
                "🏢 Execution Status — Streamwise",
                "👤 Execution Status — Executorwise",
            ],
        )

        try:
            if report_choice == "📋 Cycle & Step Basis":
                st.markdown("**Per-cycle execution and review progress**")
                cycle_rows = build_cycle_step_report(report_cycle_path, profile=run_profile)
                if cycle_rows:
                    st.dataframe(cycle_rows, use_container_width=True)
                    st.caption(f"Total active cycles: {len(cycle_rows)}")
                    csv = pd.DataFrame(cycle_rows).to_csv(index=False).encode("utf-8")
                    st.download_button("⬇️ Download CSV", csv, "cycle_step_report.csv", "text/csv")
                else:
                    st.info("No cycle data found.")

            elif report_choice == "📄 Merged Condition File":
                st.markdown("**All test steps merged from every condition script**")
                if report_script_paths:
                    with st.spinner("Merging condition scripts…"):
                        merged = build_merged_condition_report(report_script_paths, profile=run_profile)
                    if merged:
                        df_merged = pd.DataFrame(merged)
                        st.dataframe(df_merged, use_container_width=True)
                        st.caption(f"Total steps: {len(merged)} across {len(report_script_paths)} scripts")
                        csv = df_merged.to_csv(index=False).encode("utf-8")
                        st.download_button("⬇️ Download CSV", csv, "merged_condition_report.csv", "text/csv")
                    else:
                        st.info("No step data found.")
                else:
                    st.warning("No script paths available. Run a sync first to load scripts.")

            elif report_choice == "📅 Execution Status — Daywise":
                st.markdown("**Cumulative execution progress by date (planned vs. actual)**")
                daywise = build_daywise_report(report_cycle_path, profile=run_profile)
                if daywise["dates"]:
                    df_day = pd.DataFrame(daywise["series"], index=daywise["dates"])
                    st.line_chart(df_day, use_container_width=True)
                    st.caption(f"Total cycles: {daywise['total']}")
                    with st.expander("Show data table"):
                        st.dataframe(df_day, use_container_width=True)
                        csv = df_day.to_csv().encode("utf-8")
                        st.download_button("⬇️ Download CSV", csv, "daywise_report.csv", "text/csv")
                else:
                    st.info("No date data available.")

            elif report_choice == "🏢 Execution Status — Streamwise":
                st.markdown("**Planned till date, Actual till date, Ahead, Delayed — by workstream**")
                stream_rows = build_streamwise_report(report_cycle_path, profile=run_profile)
                if stream_rows:
                    display = [
                        {"Workstream":     r["area"],
                         "Total":          r["total"],
                         "Exec Plan":      r["exec_plan"],
                         "Exec Actual":    r["exec_actual"],
                         "Exec Ahead":     r["exec_ahead"],
                         "Exec Delay":     r["exec_delay"],
                         "Review Plan":    r["review_plan"],
                         "Review Actual":  r["review_actual"],
                         "Review Ahead":   r["review_ahead"],
                         "Review Delay":   r["review_delay"]}
                        for r in stream_rows
                    ]
                    st.dataframe(display, use_container_width=True)
                    chart_rows = [r for r in stream_rows if r.get("area") != "Total"]
                    if chart_rows:
                        df_s = pd.DataFrame([
                            {"Workstream": r["area"],
                             "Exec Plan": r["exec_plan"], "Exec Actual": r["exec_actual"],
                             "Review Plan": r["review_plan"], "Review Actual": r["review_actual"]}
                            for r in chart_rows
                        ]).set_index("Workstream")
                        st.bar_chart(df_s, use_container_width=True)
                    csv = pd.DataFrame(display).to_csv(index=False).encode("utf-8")
                    st.download_button("⬇️ Download CSV", csv, "streamwise_report.csv", "text/csv")
                else:
                    st.info("No workstream data found.")

            elif report_choice == "👤 Execution Status — Executorwise":
                st.markdown("**Planned till date, Actual till date, Ahead, Delayed — by executor**")
                exec_rows = build_executorwise_report(report_cycle_path, profile=run_profile)
                if exec_rows:
                    display = [
                        {"Executor":       r["executor"],
                         "Total":          r["total"],
                         "Exec Plan":      r["exec_plan"],
                         "Exec Actual":    r["exec_actual"],
                         "Exec Ahead":     r["exec_ahead"],
                         "Exec Delay":     r["exec_delay"],
                         "Review Plan":    r["review_plan"],
                         "Review Actual":  r["review_actual"],
                         "Review Ahead":   r["review_ahead"],
                         "Review Delay":   r["review_delay"]}
                        for r in exec_rows
                    ]
                    st.dataframe(display, use_container_width=True)
                    chart_rows = [r for r in exec_rows if r.get("executor") != "Total"]
                    if chart_rows:
                        df_e = pd.DataFrame([
                            {"Executor": r["executor"],
                             "Exec Plan": r["exec_plan"], "Exec Actual": r["exec_actual"],
                             "Review Plan": r["review_plan"], "Review Actual": r["review_actual"]}
                            for r in chart_rows
                        ]).set_index("Executor")
                        st.bar_chart(df_e, use_container_width=True)
                    csv = pd.DataFrame(display).to_csv(index=False).encode("utf-8")
                    st.download_button("⬇️ Download CSV", csv, "executorwise_report.csv", "text/csv")
                else:
                    st.info("No executor data found.")

        except Exception as e:
            st.error(f"Could not generate report: {e}")

# ── Run History ────────────────────────────────────────────────────────────────
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
