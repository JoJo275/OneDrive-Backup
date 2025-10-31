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
#
# USAGE EXAMPLES
#--------------
#   Interactive mode (default):
#     python main.py
#   
#   Headless mode (used by scheduled task):
#     python main.py --headless-run --backup-root "D:\\OneDriveBackup" --retention-days 30
#
# ROBOCOPY EXIT CODES
#-------------------
#   0 - No files were copied. No failure was encountered. No files were mismatched.
#   1 - One or more files were copied successfully.
#   2 - Extra files or directories were detected. Examine output log for details.
#   3 - Some files were copied. Additional files were present.
#   4 - Mismatched files or directories were detected. Examine output log.
#   5 - Some files were copied. Some files were mismatched.
#   6 - Additional files and mismatched files exist. No files were copied.
#   7 - Files were copied, a file mismatch was present, and additional files were present.
#   8+ - Serious error. Robocopy did not copy any files. Check error log.
#
# SCHEDULED TASK BEHAVIOR
#-----------------------
#   The script registers itself to run in headless mode with the Windows Task Scheduler.
#   The task will execute with the following characteristics:
#   - Runs at highest privilege level (if admin rights available)
#   - Can run whether user is logged on or not
#   - Automatically overwrites existing task with same name
#   - Supports DAILY, HOURLY, and MINUTE schedules
#
# AUTHOR: 
# VERSION: 1.0.0
# LAST MODIFIED: 2025-01-01
"""

# -----------------------------
# IMPORTS USED IN THIS SCRIPT
# -----------------------------
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
# These defaults are carefully chosen for typical use cases.
# Users can override them during interactive prompts.

DEFAULT_RETENTION_DAYS = 30                    # Keep 30 days of dated folders (approximately 1 month)
DEFAULT_BACKUP_ROOT = r"D:\OneDriveBackup"     # Target drive and folder (D: assumed as common backup drive)
DEFAULT_TASK_NAME = "OneDriveVersionedBackup"  # Windows Scheduled Task name (descriptive, no spaces)
DEFAULT_SCHEDULE_TYPE = "DAILY"                # DAILY, HOURLY, MINUTE supported (DAILY is most common)
DEFAULT_START_TIME = "09:00"                   # 24-hour HH:MM for DAILY schedule (9 AM typical work start)
DEFAULT_MODIFIER = 1                           # HOURLY: every 1 hour; MINUTE: every 1 minute (minimum interval)

# Regular expression pattern to identify our dated snapshot folders
# Format: YYYY-MM-DD_HH-MM (e.g., 2025-01-15_09-30)
# This pattern ensures we only touch folders we created, not user data
STAMP_RE = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}$")

# -----------------------------
# UTILITIES: INPUT + VALIDATION
# -----------------------------
def prompt_with_default(prompt_text: str, default_value: str) -> str:
    """
    Display an interactive prompt with a default value shown in brackets.
    
    This function provides a user-friendly way to collect input with sensible defaults.
    The user can press Enter without typing to accept the default value.
    
    Args:
        prompt_text (str): The question or instruction to display to the user.
                          Should not include the brackets for the default.
        default_value (str): The value to use if user presses Enter without input.
                           This value is displayed in square brackets.
    
    Returns:
        str: The user's input if provided, otherwise the default value.
             Always returns a string, even for numeric defaults.
    
    Example:
        >>> name = prompt_with_default("Enter your name", "John")
        Enter your name [John]: Jane
        >>> print(name)
        'Jane'
        
        >>> path = prompt_with_default("Backup location", "D:\\Backup")
        Backup location [D:\\Backup]: <user presses Enter>
        >>> print(path)
        'D:\\Backup'
    
    Note:
        - Leading and trailing whitespace is stripped from user input
        - Empty string defaults are handled correctly
        - The function always returns a string type
    """
    # Format the prompt with the default value in square brackets
    response = input(f"{prompt_text} [{default_value}]: ").strip()
    
    # Return user input if provided, otherwise return the default
    # Convert default to string to ensure consistent return type
    return response if response else str(default_value)

def prompt_int_with_default(prompt_text: str, default_value: int, min_value: int = 1) -> int:
    """
    Prompt for an integer value with validation and a default option.
    
    This function repeatedly prompts until valid input is received. It ensures
    the returned value meets the minimum threshold requirement.
    
    Args:
        prompt_text (str): The instruction text to display to the user.
                          Should describe what integer is being requested.
        default_value (int): The integer to use if user presses Enter without input.
                           Must be >= min_value to be valid.
        min_value (int, optional): The minimum acceptable value. Defaults to 1.
                                  Used to prevent invalid inputs like 0 or negative numbers.
    
    Returns:
        int: A validated integer that is >= min_value.
    
    Example:
        >>> days = prompt_int_with_default("Retention days", 30, min_value=1)
        Retention days [30]: 45
        >>> print(days)
        45
        
        >>> hours = prompt_int_with_default("Interval in hours", 24, min_value=1)
        Interval in hours [24]: 0
        Enter an integer >= 1.
        Interval in hours [24]: 12
        >>> print(hours)
        12
    
    Validation Process:
        1. Check if input is empty (use default)
        2. Check if input contains only digits
        3. Check if integer value meets minimum requirement
        4. Re-prompt if any validation fails
    
    Note:
        - Non-numeric input triggers re-prompt with error message
        - Values below min_value trigger re-prompt
        - Function loops indefinitely until valid input received
    """
    while True:
        # Get raw input from user
        raw = input(f"{prompt_text} [{default_value}]: ").strip()
        
        # Handle empty input - return default
        if raw == "":
            return int(default_value)
        
        # Validate numeric input
        if raw.isdigit():
            val = int(raw)
            # Check minimum value constraint
            if val >= min_value:
                return val
        
        # Invalid input - show error and loop
        print(f"Enter an integer >= {min_value}.")

def prompt_yes_no_default(prompt_text: str, default_yes: bool = False) -> bool:
    """
    Present a yes/no question with a default answer indicated by capitalization.
    
    This function follows Unix convention where the capitalized option is the default.
    For example: [Y/n] means Yes is default, [y/N] means No is default.
    
    Args:
        prompt_text (str): The yes/no question to ask the user.
                          Should be phrased as a question but without the question mark.
        default_yes (bool, optional): If True, default to Yes. If False, default to No.
                                     Defaults to False (No) for safety.
    
    Returns:
        bool: True if user chooses yes, False if user chooses no.
    
    Example:
        >>> proceed = prompt_yes_no_default("Continue with backup", default_yes=True)
        Continue with backup [Y/n]: <Enter>
        >>> print(proceed)
        True
        
        >>> delete = prompt_yes_no_default("Delete old files", default_yes=False)
        Delete old files [y/N]: n
        >>> print(delete)
        False
    
    Accepted Inputs:
        - For Yes: 'y', 'yes' (case-insensitive)
        - For No: 'n', 'no' (case-insensitive)
        - Empty input: uses the default based on default_yes parameter
    
    Note:
        - Invalid inputs trigger re-prompt with instruction
        - Input is case-insensitive
        - Empty input always returns the default value
    """
    # Format the prompt options based on default
    # Capital letter indicates the default option
    default_str = "Y/n" if default_yes else "y/N"
    
    while True:
        # Get user input and normalize to lowercase
        raw = input(f"{prompt_text} [{default_str}]: ").strip().lower()
        
        # Handle empty input - return default
        if raw == "" and not default_yes:
            return False
        if raw == "" and default_yes:
            return True
        
        # Check for affirmative responses
        if raw in ("y", "yes"):
            return True
        
        # Check for negative responses
        if raw in ("n", "no"):
            return False
        
        # Invalid input - show instruction and loop
        print("Answer y or n.")

def validate_time_hhmm(hhmm: str) -> bool:
    """
    Validate a time string in 24-hour HH:MM format.
    
    This function checks both the format and the logical validity of the time.
    It ensures hours are 00-23 and minutes are 00-59.
    
    Args:
        hhmm (str): Time string to validate, expected format "HH:MM".
    
    Returns:
        bool: True if the time is valid, False otherwise.
    
    Example:
        >>> validate_time_hhmm("09:30")
        True
        >>> validate_time_hhmm("24:00")  # Invalid hour
        False
        >>> validate_time_hhmm("12:60")  # Invalid minute
        False
        >>> validate_time_hhmm("9:30")   # Missing leading zero
        False
    
    Validation Rules:
        1. Must match pattern: exactly 2 digits, colon, 2 digits
        2. Hour component must be 00-23 (24-hour format)
        3. Minute component must be 00-59
    
    Note:
        - Leading zeros are required (09:05, not 9:5)
        - 24:00 is invalid (use 00:00 for midnight)
        - Does not validate semantic meaning (e.g., business hours)
    """
    # Check format: must be exactly HH:MM with all digits
    if not re.fullmatch(r"\d{2}:\d{2}", hhmm):
        return False
    
    # Split and validate components
    hh, mm = hhmm.split(":")
    
    # Validate hour is 00-23 and minute is 00-59
    return 0 <= int(hh) <= 23 and 0 <= int(mm) <= 59

def prompt_time_hhmm(prompt_text: str, default_value: str) -> str:
    """
    Prompt for a time in 24-hour HH:MM format with validation.
    
    This function ensures the user provides a valid time string, re-prompting
    if the input is invalid. It's used for scheduling tasks at specific times.
    
    Args:
        prompt_text (str): The instruction to show the user.
                          Should explain what time is being requested.
        default_value (str): The default time in HH:MM format to use if user
                           presses Enter. Should be pre-validated.
    
    Returns:
        str: A validated time string in HH:MM format.
    
    Example:
        >>> start = prompt_time_hhmm("Daily backup time", "09:00")
        Daily backup time [09:00]: 14:30
        >>> print(start)
        '14:30'
        
        >>> time = prompt_time_hhmm("Schedule time", "09:00")
        Schedule time [09:00]: 25:00
        Use 24-hour HH:MM, e.g., 09:00 or 18:30.
        Schedule time [09:00]: 00:00
        >>> print(time)
        '00:00'
    
    Note:
        - Continuously prompts until valid input received
        - Shows helpful error message with format examples
        - Always returns a string in HH:MM format
    """
    while True:
        # Get time input from user
        t = prompt_with_default(prompt_text, default_value)
        
        # Validate the time format and values
        if validate_time_hhmm(t):
            return t
        
        # Invalid input - show examples and loop
        print("Use 24-hour HH:MM, e.g., 09:00 or 18:30.")

def prompt_schedule() -> Tuple[str, str, int]:
    """
    Interactively collect scheduling preferences from the user.
    
    This function guides the user through choosing how often the backup task
    should run, with support for daily, hourly, and minute-based schedules.
    
    Returns:
        Tuple[str, str, int]: A tuple containing:
            - schedule_type (str): One of "DAILY", "HOURLY", or "MINUTE"
            - start_time (str): Time in HH:MM format for task alignment
            - modifier (int): Interval modifier (1 for DAILY, N for HOURLY/MINUTE)
    
    Schedule Types Explained:
        DAILY:
            - Runs once per day at a specific time
            - User specifies the exact time (e.g., 09:00)
            - Modifier is always 1 (ignored by schtasks)
            - Best for: Regular daily backups
        
        HOURLY:
            - Runs every N hours
            - User specifies interval (e.g., every 2 hours)
            - Start time aligns the schedule (first run)
            - Best for: Frequent backups during work hours
        
        MINUTE:
            - Runs every N minutes
            - User specifies interval (e.g., every 30 minutes)
            - Start time provides initial alignment
            - Best for: Critical data or testing
    
    Example Interaction:
        Schedule type options supported here:
          DAILY  - run once each day at the time you choose
          HOURLY - run every N hours
          MINUTE - run every N minutes
        Choose schedule type (DAILY|HOURLY|MINUTE) [DAILY]: HOURLY
        Every how many hours? (modifier /MO) [1]: 4
        Start time (HH:MM) for first run [09:00]: 08:00
        
        Returns: ("HOURLY", "08:00", 4)
    
    Note:
        - Input is case-insensitive (converted to uppercase)
        - Invalid schedule types trigger re-prompt
        - Each schedule type has appropriate follow-up questions
    """
    # Display clear options to the user
    print("\nSchedule type options supported here:")
    print("  DAILY  - run once each day at the time you choose")
    print("  HOURLY - run every N hours")
    print("  MINUTE - run every N minutes")

    # Get and validate schedule type
    while True:
        sched = prompt_with_default("Choose schedule type (DAILY|HOURLY|MINUTE)", DEFAULT_SCHEDULE_TYPE).upper()
        if sched in ("DAILY", "HOURLY", "MINUTE"):
            break
        print("Type DAILY, HOURLY, or MINUTE.")

    # Initialize return values with defaults
    start_time = DEFAULT_START_TIME
    modifier = DEFAULT_MODIFIER

    # Collect schedule-specific parameters
    if sched == "DAILY":
        # DAILY schedule: runs once per day at specified time
        # schtasks command will use: /SC DAILY /ST HH:MM
        start_time = prompt_time_hhmm("Start time (HH:MM) for daily run", DEFAULT_START_TIME)
        modifier = 1  # Modifier is ignored by schtasks for DAILY, but we keep it for consistency
        
    elif sched == "HOURLY":
        # HOURLY schedule: runs every N hours starting at specified time
        # schtasks command will use: /SC HOURLY /MO N /ST HH:MM
        modifier = prompt_int_with_default("Every how many hours? (modifier /MO)", DEFAULT_MODIFIER, min_value=1)
        start_time = prompt_time_hhmm("Start time (HH:MM) for first run", DEFAULT_START_TIME)
        
    elif sched == "MINUTE":
        # MINUTE schedule: runs every N minutes
        # schtasks command will use: /SC MINUTE /MO N /ST HH:MM
        modifier = prompt_int_with_default("Every how many minutes? (modifier /MO)", DEFAULT_MODIFIER, min_value=1)
        # Start time for minute schedules helps with alignment but is optional in schtasks
        start_time = prompt_time_hhmm("Start time (HH:MM) to align first run", DEFAULT_START_TIME)

    return sched, start_time, modifier

# -----------------------------
# PATHS AND ROBUSTNESS HELPERS
# -----------------------------
def onedrive_path() -> Path:
    """
    Intelligently resolve the OneDrive root directory path.
    
    This function handles both personal and business OneDrive installations,
    which may have different folder names (e.g., "OneDrive - CompanyName").
    
    Returns:
        Path: The resolved OneDrive directory path.
    
    Resolution Strategy:
        1. First, check the %OneDrive% environment variable
           - This is set by OneDrive and handles special names
           - Works for "OneDrive - Contoso" business accounts
        2. Fallback to %UserProfile%\\OneDrive if env var not set
           - Standard location for personal OneDrive
           - Usually C:\\Users\\Username\\OneDrive
    
    Example:
        >>> path = onedrive_path()
        >>> print(path)
        WindowsPath('C:/Users/John/OneDrive')
        
        # For business account:
        >>> path = onedrive_path()
        >>> print(path)
        WindowsPath('C:/Users/John/OneDrive - Contoso Corp')
    
    Note:
        - Does not verify the path exists (caller should check)
        - Returns Path object for cross-platform compatibility
        - Handles spaces and special characters in path names
    """
    # Try to get OneDrive path from environment variable (most reliable)
    p = os.getenv("OneDrive")
    
    # Return environment path if available, otherwise use standard location
    return Path(p) if p else Path(os.path.expandvars(r"%UserProfile%\OneDrive"))

def timestamp_stamp() -> str:
    """
    Generate a timestamp string suitable for folder naming.
    
    Creates a timestamp in YYYY-MM-DD_HH-MM format that is:
    - Sortable (chronological when sorted alphabetically)
    - Filesystem-safe (no colons or other problematic characters)
    - Human-readable (clear date and time components)
    
    Returns:
        str: Timestamp string in YYYY-MM-DD_HH-MM format.
    
    Example:
        >>> stamp = timestamp_stamp()
        >>> print(stamp)
        '2025-01-15_14-30'
    
    Format Details:
        - YYYY: 4-digit year
        - MM: 2-digit month (01-12)
        - DD: 2-digit day (01-31)
        - HH: 2-digit hour (00-23, 24-hour format)
        - MM: 2-digit minute (00-59)
        - Underscore separates date from time
        - Hyphens used instead of colons (filesystem compatibility)
    
    Note:
        - Always uses local system time
        - Format matches STAMP_RE regex pattern
        - Leading zeros ensure consistent sorting
    """
    # Generate timestamp using current local time
    # Format: YYYY-MM-DD_HH-MM (e.g., 2025-01-15_09-30)
    return dt.datetime.now().strftime("%Y-%m-%d_%H-%M")

# -----------------------------
# CORE BACKUP AND PRUNING
# -----------------------------
def run_robocopy(src: Path, dst: Path) -> int:
    """
    Execute robocopy to mirror OneDrive to a backup folder.
    
    Robocopy (Robust File Copy) is a Windows command-line tool that efficiently
    copies files and folders with many advanced options. This function uses it
    to create an exact mirror of the OneDrive folder.
    
    Args:
        src (Path): Source path (OneDrive folder) to copy from.
        dst (Path): Destination path (backup folder) to copy to.
    
    Returns:
        int: Robocopy exit code. Codes < 8 indicate success, >= 8 indicate errors.
    
    Robocopy Options Used:
        /MIR (Mirror):
            - Mirrors the source to destination
            - Copies all files and subdirectories
            - Deletes files in destination not present in source
            - Equivalent to /E (copy subdirs) plus /PURGE
        
        /FFT (Fat File Times):
            - Uses 2-second precision for file times
            - Helps with timestamp differences between filesystems
            - Reduces unnecessary copies due to precision mismatches
        
        /R:1 (Retry):
            - Retry failed copies only once
            - Default is 1 million retries (too many for our use)
            - Prevents hanging on locked files
        
        /W:1 (Wait):
            - Wait 1 second between retries
            - Default is 30 seconds (too long for single retry)
            - Keeps the process moving quickly
    
    Exit Code Interpretation:
        0: No files copied, no errors (already in sync)
        1: Files copied successfully
        2: Extra files or directories detected
        3: Some files copied, extra files present  
        4: Mismatched files or directories detected
        5: Some files copied, some mismatched
        6: Additional and mismatched files, nothing copied
        7: Files copied, mismatches and extras present
        8+: Serious error, no files copied
    
    Example:
        >>> src = Path("C:/Users/John/OneDrive")
        >>> dst = Path("D:/Backup/2025-01-15_09-00")
        >>> result = run_robocopy(src, dst)
        Running: robocopy C:/Users/John/OneDrive D:/Backup/2025-01-15_09-00 /MIR /FFT /R:1 /W:1
        ... (robocopy output) ...
        >>> print(f"Exit code: {result}")
        Exit code: 1
    
    Note:
        - Creates destination directory if it doesn't exist
        - Prints the exact command for transparency
        - Captures and displays robocopy output
        - Shows errors to stderr if exit code >= 8
    """
    # Ensure destination directory exists before running robocopy
    dst.mkdir(parents=True, exist_ok=True)
    
    # Build the robocopy command with our chosen options
    cmd = [
        "robocopy",
        str(src),      # Source directory
        str(dst),      # Destination directory  
        "/MIR",        # Mirror source to destination
        "/FFT",        # Use FAT file time (2-second precision)
        "/R:1",        # Retry once on failure
        "/W:1",        # Wait 1 second between retries
    ]
    
    # Print the command for transparency and debugging
    # Users can see exactly what command is being run
    print("\nRunning:", " ".join(cmd))
    
    # Execute robocopy and capture output
    completed = subprocess.run(cmd, capture_output=True, text=True)
    
    # Always show robocopy's standard output (file listing, summary)
    print(completed.stdout)
    
    # Show error output only if there was a serious error (code >= 8)
    if completed.returncode >= 8:
        print(completed.stderr, file=sys.stderr)
    
    return completed.returncode

def prune_old_backups(root: Path, retention_days: int) -> None:
    """
    Delete backup folders older than the retention period.
    
    This function safely removes only the dated backup folders created by this
    script, identified by their YYYY-MM-DD_HH-MM naming pattern. It will never
    delete folders with different naming patterns, protecting any user data
    that might be in the backup root.
    
    Args:
        root (Path): The backup root directory containing dated folders.
        retention_days (int): Number of days of backups to keep.
    
    Pruning Logic:
        1. Calculate cutoff date (now - retention_days)
        2. Scan all folders in backup root
        3. Check if folder name matches our date pattern
        4. Parse the date from folder name
        5. Delete if folder date is before cutoff
    
    Safety Features:
        - Only touches folders matching YYYY-MM-DD_HH-MM pattern
        - Ignores any files (only processes directories)
        - Handles parse errors gracefully
        - Uses ignore_errors=True to continue even if deletion fails
    
    Example:
        Given retention_days=7 and current date 2025-01-15:
        
        D:\\Backup\\
            2025-01-07_09-00\\  # 8 days old - will be deleted
            2025-01-08_09-00\\  # 7 days old - will be deleted  
            2025-01-09_09-00\\  # 6 days old - kept
            2025-01-15_09-00\\  # today - kept
            MyImportantData\\   # doesn't match pattern - ignored
    
    Note:
        - Cutoff is calculated as days, not hours (time component ignored)
        - Deletion errors are suppressed to ensure process continues
        - Prints each folder being deleted for audit trail
        - Silent if no folders need pruning
    """
    # Calculate the cutoff date/time
    # Folders older than this will be deleted
    cutoff = dt.datetime.now() - dt.timedelta(days=retention_days)
    
    # Iterate through all items in the backup root
    for child in root.iterdir():
        # Only process directories that match our naming pattern
        if child.is_dir() and STAMP_RE.match(child.name):
            try:
                # Parse the timestamp from the folder name
                # Format: YYYY-MM-DD_HH-MM
                d = dt.datetime.strptime(child.name, "%Y-%m-%d_%H-%M")
            except ValueError:
                # Name looked like our pattern but didn't parse correctly
                # Skip this folder defensively (don't delete what we don't understand)
                continue
            
            # Check if this backup is older than our retention cutoff
            if d < cutoff:
                # Log the deletion for audit purposes
                print(f"Pruning old backup: {child}")
                
                # Remove the entire directory tree
                # ignore_errors=True ensures we continue even if some files are locked
                shutil.rmtree(child, ignore_errors=True)

def run_once(backup_root: Path, retention_days: int) -> int:
    """
    Execute a complete backup cycle.
    
    This is the main backup logic that:
    1. Verifies OneDrive exists
    2. Creates a new timestamped backup
    3. Prunes old backups if successful
    
    Args:
        backup_root (Path): Root directory where backups are stored.
        retention_days (int): Number of days of backups to retain.
    
    Returns:
        int: Exit code (0 for success, >0 for errors).
    
    Process Flow:
        1. Resolve OneDrive path from environment
        2. Verify OneDrive directory exists
        3. Create backup root if needed
        4. Generate new timestamped folder name
        5. Run robocopy to mirror OneDrive
        6. If successful (code < 8), prune old backups
        7. Return appropriate exit code
    
    Error Handling:
        - Returns 1 if OneDrive path doesn't exist
        - Returns robocopy exit code if >= 8 (serious error)
        - Only prunes if backup was successful
    
    Example:
        >>> result = run_once(Path("D:/Backup"), 30)
        Running: robocopy C:/Users/John/OneDrive D:/Backup/2025-01-15_09-30 /MIR /FFT /R:1 /W:1
        ... (robocopy output) ...
        Pruning old backup: D:/Backup/2024-12-15_09-00
        >>> print(result)
        0
    
    Note:
        - Creates all necessary directories automatically
        - Timestamp includes minutes for multiple daily runs
        - Pruning only happens after successful backup
        - All output goes to stdout except errors
    """
    # Step 1: Resolve and verify OneDrive path
    src = onedrive_path()
    if not src.exists():
        # OneDrive directory not found - can't proceed
        print(f"OneDrive path not found: {src}", file=sys.stderr)
        return 1

    # Step 2: Ensure backup root directory exists
    # Create full path including parents if needed
    backup_root.mkdir(parents=True, exist_ok=True)
    
    # Step 3: Create new timestamped destination folder
    # Format: backup_root/YYYY-MM-DD_HH-MM
    dst = backup_root / timestamp_stamp()
    
    # Step 4: Execute the backup using robocopy
    rc = run_robocopy(src, dst)
    
    # Step 5: If backup successful, prune old backups
    # Only prune if robocopy succeeded (exit code < 8)
    if rc < 8:
        prune_old_backups(backup_root, retention_days)
        return 0  # Return success
    
    # Backup failed - return the robocopy error code
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
    Construct the Windows schtasks command to create a scheduled task.
    
    This function builds the complete command line arguments for schtasks.exe
    to register a scheduled task that runs this script in headless mode.
    
    Args:
        task_name (str): Name for the scheduled task (shown in Task Scheduler).
        schedule_type (str): One of "DAILY", "HOURLY", or "MINUTE".
        start_time_hhmm (str): Start time in HH:MM format.
        modifier (int): Interval modifier for HOURLY/MINUTE schedules.
        python_exe (str): Path to Python interpreter.
        script_path (Path): Path to this script.
        backup_root (Path): Backup destination root directory.
        retention_days (int): Days of backups to retain.
    
    Returns:
        List[str]: Command line arguments for schtasks.exe.
    
    Schtasks Parameters Explained:
        /Create: Create a new task (or update if exists with /F)
        /TN: Task name (appears in Task Scheduler GUI)
        /SC: Schedule type (DAILY, HOURLY, MINUTE)
        /ST: Start time in HH:MM format
        /MO: Modifier (interval for HOURLY/MINUTE)
        /TR: Task to run (our script with arguments)
        /RL: Run level (HIGHEST for admin privileges)
        /F: Force - overwrite existing task without prompting
    
    Task Command Line:
        The task will execute:
        "python.exe" "script.py" --headless-run --backup-root "D:\\Backup" --retention-days 30
        
        This runs the script in headless mode (no prompts) with the specified parameters.
    
    Example:
        >>> cmd = build_schtasks_command(
        ...     "MyBackup", "DAILY", "09:00", 1,
        ...     "C:/Python/python.exe", Path("backup.py"),
        ...     Path("D:/Backup"), 30
        ... )
        >>> print(" ".join(cmd))
        schtasks /Create /TN MyBackup /SC DAILY /TR "..." /RL HIGHEST /F /ST 09:00
    
    Note:
        - Quotes paths to handle spaces
        - Uses HIGHEST run level for file access permissions
        - /F flag overwrites existing tasks automatically
        - Start time included for all schedule types for consistency
    """
    # Build the command line that the scheduled task will execute
    # This runs our script with --headless-run flag and all necessary parameters
    run_args = (
        f'"{python_exe}" "{script_path}" '  # Python interpreter and script path
        f'--headless-run '                   # Flag for non-interactive mode
        f'--backup-root "{backup_root}" '    # Where to store backups
        f'--retention-days {retention_days}' # How many days to keep
    )

    # Start building the schtasks command
    cmd = [
        "schtasks",
        "/Create",         # Create new task
        "/TN", task_name,  # Task name
        "/SC", schedule_type,  # Schedule type (DAILY/HOURLY/MINUTE)
        "/TR", run_args,   # Command to run
        "/RL", "HIGHEST",  # Run with highest privileges available
        "/F"               # Force creation (overwrite if exists)
    ]

    # Add start time if valid (used for schedule alignment)
    # /ST is accepted by all schedule types for initial timing
    if validate_time_hhmm(start_time_hhmm):
        cmd.extend(["/ST", start_time_hhmm])

    # Add modifier for HOURLY and MINUTE schedules
    # /MO specifies the interval (every N hours/minutes)
    # This parameter is ignored for DAILY schedules
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
    Register or update a Windows Scheduled Task for automatic backups.
    
    This function creates a scheduled task that will run this script
    automatically at the specified schedule. The task runs in headless
    mode (no user interaction required).
    
    Args:
        task_name (str): Name for the task in Task Scheduler.
        schedule_type (str): Schedule type (DAILY/HOURLY/MINUTE).
        start_time_hhmm (str): Start time in HH:MM format.
        modifier (int): Interval for HOURLY/MINUTE schedules.
        backup_root (Path): Backup destination directory.
        retention_days (int): Days of backups to keep.
    
    Returns:
        int: Exit code from schtasks (0 = success).
    
    Task Properties:
        - Runs with highest available privileges
        - Can run whether user is logged on or not
        - Will not stop if running on batteries (laptops)
        - Overwrites existing task with same name
    
    Required Permissions:
        - Standard user: Can create tasks for own account
        - Administrator: Can create system-wide tasks
        - Run as admin for best results (file access permissions)
    
    Example:
        >>> result = install_task(
        ...     "DailyBackup", "DAILY", "09:00", 1,
        ...     Path("D:/Backup"), 30
        ... )
        Registering Scheduled Task with command:
        schtasks /Create /TN DailyBackup ...
        SUCCESS: The scheduled task "DailyBackup" has successfully been created.
        >>> print(result)
        0
    
    Troubleshooting:
        - Error 5: Access denied - need admin rights
        - Error 1: Incorrect function - invalid parameters
        - Task appears disabled: Check task conditions in GUI
    
    Note:
        - Task name must be unique
        - Use Task Scheduler GUI to view/modify created tasks
        - Task will start at next scheduled time
        - Check Event Viewer for task execution history
    """
    # Get the Python executable path that's running this script
    python_exe = sys.executable
    
    # Get the absolute path to this script file
    script_path = Path(__file__).resolve()
    
    # Build the complete schtasks command
    cmd = build_schtasks_command(
        task_name, schedule_type, start_time_hhmm, modifier,
        python_exe, script_path, backup_root, retention_days
    )

    # Show the user what command we're running for transparency
    print("\nRegistering Scheduled Task with command:")
    print(" ".join(cmd))
    
    # Execute the schtasks command
    res = subprocess.run(cmd, capture_output=True, text=True)
    
    # Display results
    if res.returncode != 0:
        # Task creation failed - show error message
        print(res.stderr, file=sys.stderr)
    else:
        # Task created successfully - show success message
        print(res.stdout)
    
    return res.returncode

def stop_task(task_name: str) -> None:
    """
    Stop and optionally delete a scheduled task.
    
    This function first attempts to disable the task (keeping its definition),
    then falls back to deleting it entirely if disabling fails.
    
    Args:
        task_name (str): Name of the task to stop.
    
    Process:
        1. Try to disable the task (keeps configuration)
        2. If disable fails, try to delete the task
        3. Report the outcome to user
    
    Disable vs Delete:
        - Disable: Task remains in Task Scheduler but won't run
        - Delete: Task is completely removed from Task Scheduler
    
    Example:
        >>> stop_task("DailyBackup")
        Stopping task 'DailyBackup'...
        Task 'DailyBackup' disabled.
        
        >>> stop_task("NonExistentTask")
        Stopping task 'NonExistentTask'...
        No task found to stop or delete.
    
    Note:
        - Requires appropriate permissions to modify the task
        - Gracefully handles non-existent tasks
        - Uses /F flag to force deletion without confirmation
    """
    print(f"\nStopping task '{task_name}'...")
    
    # First attempt: Try to disable the task
    # This keeps the task definition but prevents it from running
    disable = subprocess.run(["schtasks", "/Change", "/TN", task_name, "/Disable"],
                             capture_output=True, text=True)
    if disable.returncode == 0:
        print(f"Task '{task_name}' disabled.")
        return
    
    # Second attempt: Try to delete the task entirely
    # Use /F to force deletion without confirmation prompt
    delete = subprocess.run(["schtasks", "/Delete", "/TN", task_name, "/F"],
                            capture_output=True, text=True)
    if delete.returncode == 0:
        print(f"Task '{task_name}' deleted.")
    else:
        # Task doesn't exist or we lack permissions
        print("No task found to stop or delete.", file=sys.stderr)

# -----------------------------
# MAIN INTERACTIVE FLOW
# -----------------------------
def interactive_main() -> int:
    """
    Main interactive entry point for the script.
    
    This function provides a user-friendly wizard that:
    1. Collects configuration with sensible defaults
    2. Optionally runs an immediate backup
    3. Optionally creates/updates a scheduled task
    4. Optionally stops an existing task
    
    Returns:
        int: Exit code (always 0 for interactive mode).
    
    User Experience Flow:
        1. Welcome message
        2. Configuration prompts:
           - Retention days (how long to keep backups)
           - Backup destination folder
           - Task name for scheduler
           - Schedule type and timing
        3. Action prompts:
           - Run backup now? (recommended for testing)
           - Create/update scheduled task?
           - Stop existing task?
        4. Completion message
    
    Design Philosophy:
        - Every prompt has a sensible default
        - User can accept all defaults by pressing Enter
        - Clear explanations for each setting
        - Actions are optional and confirmed
        - Non-destructive (only copies data)
    
    Example Session:
        === OneDrive Versioned Backup (Interactive) ===
        Retention in days (how many days of dated backups to keep) [30]: 7
        Backup root folder (should NOT be inside OneDrive) [D:\\OneDriveBackup]: E:\\Backups
        Scheduled Task name [OneDriveVersionedBackup]: <Enter>
        
        Schedule type options supported here:
          DAILY  - run once each day at the time you choose
          HOURLY - run every N hours
          MINUTE - run every N minutes
        Choose schedule type (DAILY|HOURLY|MINUTE) [DAILY]: <Enter>
        Start time (HH:MM) for daily run [09:00]: 06:00
        
        Run a one-time backup now? [Y/n]: y
        Running: robocopy ...
        
        Install or update the Scheduled Task with these settings? [Y/n]: y
        Registering Scheduled Task...
        
        Do you want to stop (disable/delete) the Scheduled Task now? [y/N]: n
        
        Done.
    
    Error Recovery:
        - Failed backup doesn't prevent task scheduling
        - Failed task creation shows helpful error message
        - User can re-run script to try again
        - All errors are non-fatal in interactive mode
    
    Note:
        - Designed for Windows users of all skill levels
        - No command-line arguments needed
        - Settings are not persisted (asked each time)
        - Consider running as administrator for best results
    """
    # Display welcome banner
    print("\n=== OneDrive Versioned Backup (Interactive) ===")

    # STEP 1: Collect configuration settings
    
    # Ask for retention period with explanation
    retention_days = prompt_int_with_default(
        "Retention in days (how many days of dated backups to keep)",
        DEFAULT_RETENTION_DAYS,
        min_value=1
    )

    # Ask for backup destination with warning about OneDrive
    backup_root_str = prompt_with_default(
        "Backup root folder (should NOT be inside OneDrive)",
        DEFAULT_BACKUP_ROOT
    )
    # Expand user home directory notation (~) if present
    backup_root = Path(backup_root_str).expanduser()

    # Ask for task name (shown in Task Scheduler)
    task_name = prompt_with_default(
        "Scheduled Task name",
        DEFAULT_TASK_NAME
    )

    # Ask for schedule configuration
    schedule_type, start_time, modifier = prompt_schedule()

    # STEP 2: Optional immediate backup
    # Useful for testing configuration and permissions
    if prompt_yes_no_default("Run a one-time backup now?", default_yes=True):
        rc = run_once(backup_root, retention_days)
        if rc >= 8:
            # Backup failed with serious error
            # Inform user but continue (they may want to schedule anyway)
            print("Backup failed (robocopy exit code >= 8). Fix issues and try again.", file=sys.stderr)

    # STEP 3: Optional task scheduling
    # Creates or updates the Windows Scheduled Task
    if prompt_yes_no_default("Install or update the Scheduled Task with these settings?", default_yes=True):
        rc = install_task(task_name, schedule_type, start_time, modifier, backup_root, retention_days)
        if rc != 0:
            # Task registration failed
            # Common cause: need administrator privileges
            print("Task registration failed. You may need to run your shell as Administrator.", file=sys.stderr)

    # STEP 4: Optional task stopping
    # Allows user to disable/delete task if needed
    if prompt_yes_no_default("Do you want to stop (disable/delete) the Scheduled Task now?", default_yes=False):
        stop_task(task_name)

    # Display completion message
    print("\nDone.")
    return 0

# -----------------------------
# HEADLESS ENTRY FOR SCHEDULED TASK
# -----------------------------
def headless_run(backup_root: Path, retention_days: int) -> int:
    """
    Execute backup in headless (non-interactive) mode.
    
    This function is called when the script runs from Task Scheduler or
    with the --headless-run flag. It performs exactly one backup cycle
    without any user interaction.
    
    Args:
        backup_root (Path): Directory where backups are stored.
        retention_days (int): Number of days of backups to keep.
    
    Returns:
        int: Exit code (0 = success, >0 = error).
    
    Characteristics:
        - No user prompts or interaction
        - Minimal output (only essential information)
        - Suitable for automated execution
        - Returns meaningful exit codes for monitoring
    
    Use Cases:
        - Called by Windows Task Scheduler
        - Batch scripts or automation tools
        - Testing with specific parameters
    
    Example:
        >>> result = headless_run(Path("D:/Backup"), 7)
        Running: robocopy ...
        Pruning old backup: D:/Backup/2025-01-01_09-00
        >>> print(result)
        0
    
    Note:
        - All output suitable for log files
        - Errors go to stderr for separate capture
        - Exit code can trigger alerts in monitoring systems
    """
    # Simply run one backup cycle with the provided parameters
    return run_once(backup_root, retention_days)

# -----------------------------
# ARG PARSE LITE
# -----------------------------
def parse_args(argv: List[str]) -> dict:
    """
    Lightweight command-line argument parser.
    
    This function provides minimal argument parsing without external dependencies.
    It supports just the essential flags needed for headless operation.
    
    Args:
        argv (List[str]): Command line arguments (typically sys.argv[1:]).
    
    Returns:
        dict: Parsed arguments with keys:
            - "mode": Either "interactive" or "headless"
            - "backup_root": Backup destination path
            - "retention_days": Days to retain backups
    
    Supported Arguments:
        --headless-run:
            Switch to headless mode (no user interaction)
            Used by scheduled tasks
        
        --backup-root <path>:
            Specify backup destination directory
            Required for headless mode
        
        --retention-days <int>:
            Specify retention period in days
            Required for headless mode
    
    Examples:
        Interactive (default):
            python main.py
        
        Headless with parameters:
            python main.py --headless-run --backup-root "D:\\Backup" --retention-days 30
        
        Partial arguments (falls back to interactive):
            python main.py --backup-root "D:\\Backup"
    
    Parsing Rules:
        - Unknown arguments are safely ignored
        - Invalid values use defaults with warning
        - Order of arguments doesn't matter
        - No short flags (only --long-flags)
    
    Note:
        - Intentionally simple (no argparse dependency)
        - Could be extended for more options if needed
        - Defaults ensure script always has valid values
    """
    # Initialize with defaults
    args = {
        "mode": "interactive",
        "backup_root": DEFAULT_BACKUP_ROOT,
        "retention_days": DEFAULT_RETENTION_DAYS,
    }
    
    # Parse arguments using simple position-based approach
    i = 0
    while i < len(argv):
        tok = argv[i]
        
        # Check for headless mode flag
        if tok == "--headless-run":
            args["mode"] = "headless"
            i += 1
            
        # Check for backup root with value
        elif tok == "--backup-root" and i + 1 < len(argv):
            args["backup_root"] = argv[i + 1]
            i += 2
            
        # Check for retention days with value
        elif tok == "--retention-days" and i + 1 < len(argv):
            try:
                # Validate integer value
                args["retention_days"] = int(argv[i + 1])
            except ValueError:
                # Invalid integer - warn and use default
                print("Invalid --retention-days. Using default.", file=sys.stderr)
            i += 2
            
        else:
            # Unknown argument - skip it
            # Interactive mode will ignore it safely
            i += 1
    
    return args

# -----------------------------
# PROGRAM ENTRY
# -----------------------------
if __name__ == "__main__":
    """
    Script entry point - determines mode and executes accordingly.
    
    This block runs when the script is executed directly (not imported).
    It parses command-line arguments and routes to either interactive or
    headless mode.
    
    Exit Codes:
        0: Success
        1: General error (e.g., OneDrive not found)
        8+: Robocopy error (serious backup failure)
    
    Execution Modes:
        1. Interactive (default):
           - No arguments or unrecognized arguments
           - User prompts for all settings
           - Friendly wizard-style interface
        
        2. Headless (automated):
           - Triggered by --headless-run flag
           - No user interaction
           - Used by Task Scheduler
    
    Example Usage:
        Interactive:
            python main.py
        
        Headless:
            python main.py --headless-run --backup-root "D:\\Backup" --retention-days 30
    """
    # Parse command-line arguments
    parsed = parse_args(sys.argv[1:])
    
    # Route to appropriate mode based on parsed arguments
    if parsed["mode"] == "headless":
        # Headless mode - called by Task Scheduler or automation
        # Run one backup cycle and exit with appropriate code
        sys.exit(headless_run(Path(parsed["backup_root"]), int(parsed["retention_days"])))
    else:
        # Interactive mode - normal manual execution
        # Start the interactive wizard for configuration and execution
        sys.exit(interactive_main())
