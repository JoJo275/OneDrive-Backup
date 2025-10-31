# OneDrive-Backup

This project provides a Python script that automates the backup of specified local directories to OneDrive. It supports versioned backups, allowing users to maintain multiple versions of their files based on configurable retention policies.

## Status

    Did not finish
    - Stopped after finding out locked files cannot be copied without third party software or closing the application using the file.
    - Basic functionality works, but needs more features and testing.
    - Also can just copy and paste files over night or smth.

## Changes to Consider

    Add additional documentation after implementing the changes below.

if backup is missed due to PC powered off

- change code to automatically backup when pc is powered on (make this an option the user can choose when running program)

Additionally

- Add options a user can choose from (the user can see the options when they first run the program)

- notify user if modifier is at a certain place sctasks is not used at line: 359 by function prompt_schedule

- Add logging functionality to track backup activities and errors

- Maybe implement email notifications to inform users about backup status and issues
