# !/usr/bin/env python3
"""main.py
#
#   OneDrive Versioned Backup Interactive - main script
# =======================================================
# This script helps you back up your OneDrive to a local disk. It is copy-only
# (does not delete or modify files in OneDrive). It prunes only old dated
# folders it created under your backup root.
#
# PURPOSE
#--------
#   - Mirror your OneDrive into a new, dated folder on a local disk each run.
#   - Keep only the last N days of backups (prune old dated folders).
#   - Optionally register a Windows Scheduled Task to run automatically.
#   - Provide interactive prompts with clear defaults so beginners can run it.
#
# SAFE TO READ NOTES
#-------------------
#   - This script only *copies* data out of OneDrive into a backup root.
#   - It never modifies your OneDrive files.
#   - It prunes only dated backup folders it created under your backup root.
#
# REQUIREMENTS
#--------------
#   - Windows 10/11 with 'robocopy' and 'schtasks' available (built into
# Windows).
#   - Run from a normal Command Prompt or PowerShell. Admin may be needed to
#     register a task at "highest" privileges.
#
# BACKUP LAYOUT EXAMPLE
#----------------------
#   D:\OneDriveBackup\
#     2025-10-31_09-00\
#     2025-11-01_09-00\
#     ...
#
# RETURN CODES
#-------------
#   - 0  success
#   - >0 error (including robocopy exit codes >= 8)
"""

import os
import re
import sys
import shutil
import subprocess
import datetime as dt
from pathlib import Path
from typing import Tuple, List

# -----------------------------
# DEFAULTS SHOWN TO THE USER
# -----------------------------
DEFAULT_RETENTION_DAYS = 30               # keep 30 days of dated folders
DEFAULT_BACKUP_ROOT = r"D:\OneDriveBackup"  # target drive and folder
DEFAULT_TASK_NAME = "OneDriveVersionedBackup"  # Windows Scheduled Task name
DEFAULT_SCHEDULE_TYPE = "DAILY"           # DAILY, HOURLY, MINUTE supported here
DEFAULT_START_TIME = "09:00"              # 24-hour HH:MM for DAILY schedule
DEFAULT_MODIFIER = 1                      # HOURLY: every 1 hour; MINUTE: every 1 minute

# Regex used to detect our dated snapshot folders (YYYY-MM-DD_HH-MM)
STAMP_RE = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}$")

# -----------------------------
# UTILITIES: INPUT + VALIDATION
# -----------------------------
def prompt_with_default(prompt_text: str, default_value: str) -> str:
    """
    Show a prompt with a [default]. The user can press Enter to accept the default.
    Always returns a string (possibly empty if default is empty).
    """
    response = input(f"{prompt_text} [{default_value}]: ").strip()
    return response if response else str(default_value)

def prompt_int_with_default(prompt_text: str, default_value: int, min_value: int = 1) -> int:
    """
    Ask for an integer. If empty, use default.
    Validate it is >= min_value. Re-prompt on invalid input.
    """
    while True:
        raw = input(f"{prompt_text} [{default_value}]: ").strip()
        if raw == "":
            return int(default_value)
        if raw.isdigit():
            val = int(raw)
            if val >= min_value:
                return val
        print(f"Enter an integer >= {min_value}.")

def prompt_yes_no_default(prompt_text: str, default_yes: bool = False) -> bool:
    """
    Ask a yes/no question with a default.
    Returns True for yes, False for no.
    """
    default_str = "Y/n" if default_yes else "y/N"
    while True:
        raw = input(f"{prompt_text} [{default_str}]: ").strip().lower()
        if raw == "" and not default_yes:
            return False
        if raw == "" and default_yes:
            return True
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print("Answer y or n.")

def validate_time_hhmm(hhmm: str) -> bool:
    """
    Validate time string is 24-hour HH:MM where HH in 00..23 and MM in 00..59.
    """
    if not re.fullmatch(r"\d{2}:\d{2}", hhmm):
        return False
    hh, mm = hhmm.split(":")
    return 0 <= int(hh) <= 23 and 0 <= int(mm) <= 59

def prompt_time_hhmm(prompt_text: str, default_value: str) -> str:
    """
    Prompt for a time in HH:MM with validation. Re-prompt on invalid input.
    """
    while True:
        t = prompt_with_default(prompt_text, default_value)
        if validate_time_hhmm(t):
            return t
        print("Use 24-hour HH:MM, e.g., 09:00 or 18:30.")

def prompt_schedule() -> Tuple[str, str, int]:
    """
    Ask the user how often to run the task.
    - DAILY: needs a Start Time HH:MM. Modifier not used by schtasks for daily.
    - HOURLY: ask for every N hours (modifier). Start time optional in schtasks,
              but we set one to keep behavior predictable.
    - MINUTE: ask for every N minutes (modifier). Start time not needed, but we
              provide one for consistency. schtasks will begin at next aligned interval.
    Returns (schedule_type, start_time, modifier)
    """
    print("\nSchedule type options supported here:")
    print("  DAILY  - run once each day at the time you choose")
    print("  HOURLY - run every N hours")
    print("  MINUTE - run every N minutes")

    while True:
        sched = prompt_with_default("Choose schedule type (DAILY|HOURLY|MINUTE)", DEFAULT_SCHEDULE_TYPE).upper()
        if sched in ("DAILY", "HOURLY", "MINUTE"):
            break
        print("Type DAILY, HOURLY, or MINUTE.")

    start_time = DEFAULT_START_TIME
    modifier = DEFAULT_MODIFIER

    if sched == "DAILY":
        # schtasks uses /SC DAILY with /ST HH:MM
        start_time = prompt_time_hhmm("Start time (HH:MM) for daily run", DEFAULT_START_TIME)
        modifier = 1  # ignored by schtasks for DAILY; kept for uniform return signature
    elif sched == "HOURLY":
        # schtasks uses /SC HOURLY with /MO <n>; optionally /ST sets the first start
        modifier = prompt_int_with_default("Every how many hours? (modifier /MO)", DEFAULT_MODIFIER, min_value=1)
        start_time = prompt_time_hhmm("Start time (HH:MM) for first run", DEFAULT_START_TIME)
    elif sched == "MINUTE":
        # schtasks uses /SC MINUTE with /MO <n>; /ST is optional
        modifier = prompt_int_with_default("Every how many minutes? (modifier /MO)", DEFAULT_MODIFIER, min_value=1)
        # Start time for minute schedules is optional; we still set one for clarity.
        start_time = prompt_time_hhmm("Start time (HH:MM) to align first run", DEFAULT_START_TIME)

    return sched, start_time, modifier

# -----------------------------
# PATHS AND ROBUSTNESS HELPERS
# -----------------------------
def onedrive_path() -> Path:
    """
    Resolve OneDrive root:
      - Prefer %OneDrive% (works for personal and business names like 'OneDrive - Contoso').
      - Fallback to %UserProfile%\OneDrive.
    """
    p = os.getenv("OneDrive")
    return Path(p) if p else Path(os.path.expandvars(r"%UserProfile%\OneDrive"))

def timestamp_stamp() -> str:
    """
    Produce a stamp like 2025-10-31_09-00 for folder names. This matches STAMP_RE.
    """
    return dt.datetime.now().strftime("%Y-%m-%d_%H-%M")

# -----------------------------
# CORE BACKUP AND PRUNING
# -----------------------------
def run_robocopy(src: Path, dst: Path) -> int:
    """
    Execute robocopy to mirror OneDrive into a new dated folder.
    /MIR mirrors. /FFT relaxes timestamp precision differences.
    /R:1 /W:1 keeps retries short so the task doesn't hang for long.
    Returns robocopy's exit code. Codes < 8 are considered success.
    """
    dst.mkdir(parents=True, exist_ok=True)
    cmd = [
        "robocopy",
        str(src),
        str(dst),
        "/MIR",
        "/FFT",
        "/R:1",
        "/W:1",
    ]
    # Print exactly what we run so beginners can see and learn.
    print("\nRunning:", " ".join(cmd))
    completed = subprocess.run(cmd, capture_output=True, text=True)
    print(completed.stdout)
    if completed.returncode >= 8:
        print(completed.stderr, file=sys.stderr)
    return completed.returncode

def prune_old_backups(root: Path, retention_days: int) -> None:
    """
    Delete backup folders older than retention_days.
    We only touch folders whose names match our date stamp pattern.
    This avoids touching any unrelated folders in the backup root.
    """
    cutoff = dt.datetime.now() - dt.timedelta(days=retention_days)
    for child in root.iterdir():
        if child.is_dir() and STAMP_RE.match(child.name):
            try:
                d = dt.datetime.strptime(child.name, "%Y-%m-%d_%H-%M")
            except ValueError:
                # Name looked like our pattern but did not parse; skip defensively.
                continue
            if d < cutoff:
                print(f"Pruning old backup: {child}")
                shutil.rmtree(child, ignore_errors=True)

def run_once(backup_root: Path, retention_days: int) -> int:
    """
    One full backup cycle:
      1) Resolve OneDrive path and verify it exists.
      2) Create a new dated destination folder.
      3) Run robocopy to mirror into that folder.
      4) If success (<8), prune old dated folders beyond retention.
    """
    src = onedrive_path()
    if not src.exists():
        print(f"OneDrive path not found: {src}", file=sys.stderr)
        return 1

    backup_root.mkdir(parents=True, exist_ok=True)
    dst = backup_root / timestamp_stamp()
    rc = run_robocopy(src, dst)
    if rc < 8:
        prune_old_backups(backup_root, retention_days)
        return 0
    return rc

# -----------------------------
# SCHEDULED TASK MANAGEMENT
# -----------------------------
def build_schtasks_command(task_name: str,
                           schedule_type: str,
                           start_time_hhmm: str,
                           modifier: int,
                           python_exe: str,
                           script_path: Path,
                           backup_root: Path,
                           retention_days: int) -> List[str]:
    """
    Create the 'schtasks /Create' command, using:
      - /SC <DAILY|HOURLY|MINUTE>
      - /ST HH:MM for alignment
      - /MO <n> for HOURLY or MINUTE
      - /TR "python this_script.py --run-now-mode" (we pass args so Task runs the backup)
    We run the script with --headless-run to skip prompts when the Task executes.
    """
    # Build the command line the task will run each time.
    # We call ourselves with a special flag that does a single run with given arguments.
    run_args = (
        f'"{python_exe}" "{script_path}" '
        f'--headless-run '
        f'--backup-root "{backup_root}" '
        f'--retention-days {retention_days}'
    )

    cmd = [
        "schtasks",
        "/Create",
        "/TN", task_name,
        "/SC", schedule_type,
        "/TR", run_args,
        "/RL", "HIGHEST",
        "/F"  # overwrite if exists
    ]

    # /ST is valid across schedule types. We include it for predictable alignment.
    if validate_time_hhmm(start_time_hhmm):
        cmd.extend(["/ST", start_time_hhmm])

    # /MO applies to HOURLY and MINUTE. It is ignored for DAILY.
    if schedule_type in ("HOURLY", "MINUTE") and modifier >= 1:
        cmd.extend(["/MO", str(modifier)])

    return cmd

def install_task(task_name: str,
                 schedule_type: str,
                 start_time_hhmm: str,
                 modifier: int,
                 backup_root: Path,
                 retention_days: int) -> int:
    """
    Register or overwrite a Windows Scheduled Task that runs this script
    in headless mode at the configured schedule.
    """
    python_exe = sys.executable
    script_path = Path(__file__).resolve()
    cmd = build_schtasks_command(
        task_name, schedule_type, start_time_hhmm, modifier,
        python_exe, script_path, backup_root, retention_days
    )

    print("\nRegistering Scheduled Task with command:")
    print(" ".join(cmd))
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        print(res.stderr, file=sys.stderr)
    else:
        print(res.stdout)
    return res.returncode

def stop_task(task_name: str) -> None:
    """
    Try to disable the task. If not found, try delete. Report outcome.
    """
    print(f"\nStopping task '{task_name}'...")
    disable = subprocess.run(["schtasks", "/Change", "/TN", task_name, "/Disable"],
                             capture_output=True, text=True)
    if disable.returncode == 0:
        print(f"Task '{task_name}' disabled.")
        return
    delete = subprocess.run(["schtasks", "/Delete", "/TN", task_name, "/F"],
                            capture_output=True, text=True)
    if delete.returncode == 0:
        print(f"Task '{task_name}' deleted.")
    else:
        print("No task found to stop or delete.", file=sys.stderr)

# -----------------------------
# MAIN INTERACTIVE FLOW
# -----------------------------
def interactive_main() -> int:
    """
    Interactive session:
      1) Ask user for all key values with defaults shown.
      2) Optionally run a one-time backup now.
      3) Optionally create or update the Scheduled Task.
      4) Optionally stop the Scheduled Task.
    All steps are optional. Nothing runs unless confirmed.
    """
    print("\n=== OneDrive Versioned Backup (Interactive) ===")

    # Ask for retention days with default
    retention_days = prompt_int_with_default(
        "Retention in days (how many days of dated backups to keep)",
        DEFAULT_RETENTION_DAYS,
        min_value=1
    )

    # Ask for backup root with default
    backup_root_str = prompt_with_default(
        "Backup root folder (should NOT be inside OneDrive)",
        DEFAULT_BACKUP_ROOT
    )
    backup_root = Path(backup_root_str).expanduser()

    # Ask for Windows Scheduled Task name with default
    task_name = prompt_with_default(
        "Scheduled Task name",
        DEFAULT_TASK_NAME
    )

    # Ask schedule details (type, start time, modifier)
    schedule_type, start_time, modifier = prompt_schedule()

    # Ask whether to run a backup now (useful to test paths and permissions)
    if prompt_yes_no_default("Run a one-time backup now?", default_yes=True):
        rc = run_once(backup_root, retention_days)
        if rc >= 8:
            print("Backup failed (robocopy exit code >= 8). Fix issues and try again.", file=sys.stderr)
            # We still allow user to continue to scheduling after a failed test, if desired.

    # Ask whether to install or update the scheduled task with the given settings
    if prompt_yes_no_default("Install or update the Scheduled Task with these settings?", default_yes=True):
        rc = install_task(task_name, schedule_type, start_time, modifier, backup_root, retention_days)
        if rc != 0:
            print("Task registration failed. You may need to run your shell as Administrator.", file=sys.stderr)

    # Ask whether to stop the task right now (disable or delete)
    if prompt_yes_no_default("Do you want to stop (disable/delete) the Scheduled Task now?", default_yes=False):
        stop_task(task_name)

    print("\nDone.")
    return 0

# -----------------------------
# HEADLESS ENTRY FOR SCHEDULED TASK
# -----------------------------
def headless_run(backup_root: Path, retention_days: int) -> int:
    """
    Special entry used by the Scheduled Task. It runs one backup cycle
    without prompts, then exits.
    """
    return run_once(backup_root, retention_days)

# -----------------------------
# ARG PARSE LITE
# -----------------------------
def parse_args(argv: List[str]) -> dict:
    """
    Minimal flag parsing to support:
      --headless-run                : run once without prompts (used by Task Scheduler)
      --backup-root <path>
      --retention-days <int>
    Anything else falls back to interactive mode.
    """
    args = {
        "mode": "interactive",
        "backup_root": DEFAULT_BACKUP_ROOT,
        "retention_days": DEFAULT_RETENTION_DAYS,
    }
    i = 0
    while i < len(argv):
        tok = argv[i]
        if tok == "--headless-run":
            args["mode"] = "headless"
            i += 1
        elif tok == "--backup-root" and i + 1 < len(argv):
            args["backup_root"] = argv[i + 1]
            i += 2
        elif tok == "--retention-days" and i + 1 < len(argv):
            try:
                args["retention_days"] = int(argv[i + 1])
            except ValueError:
                print("Invalid --retention-days. Using default.", file=sys.stderr)
            i += 2
        else:
            # Unknown token -> interactive mode will ignore it safely.
            i += 1
    return args

# -----------------------------
# PROGRAM ENTRY
# -----------------------------
if __name__ == "__main__":
    parsed = parse_args(sys.argv[1:])
    if parsed["mode"] == "headless":
        # Called by the Scheduled Task. No prompts.
        sys.exit(headless_run(Path(parsed["backup_root"]), int(parsed["retention_days"])))
    else:
        # Normal manual run. Ask everything with defaults.
        sys.exit(interactive_main())
