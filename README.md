# Windows Server Cleanup Script

## Overview

This PowerShell script automates the cleanup process for Ideal Integration's NOC. It removes outdated logs, temporary files, and unnecessary clutter to ensure smooth system performance and clearing disk space alerts.

## Features

- Automates the deletion of old log files.
- Clears temporary directories.
- Ensures system efficiency by removing unnecessary data.
- Provides logging for audit and verification purposes.

## Requirements

- Windows PowerShell 5.1 or later
- Appropriate permissions to delete files and folders
- Execution policy set to allow script execution (use `Set-ExecutionPolicy RemoteSigned` if necessary)

## Installation

1. Download the script `ImprovedNOCCleanup.ps1` to a directory of your choice.
2. Ensure you have the necessary permissions to execute the script.

## Usage

### Flags and Options

-DryRun: Runs the script without making any changes, showing what would be deleted.

-Verbose: Enables detailed output for debugging purposes.

-LogPath <path>: Specifies a custom log file location.

-DaysOld <days>: Specifies the age threshold (in days) for cleaning log files. Default is 3 days.


Run the script manually:

```powershell
powershell -ExecutionPolicy Bypass -File ImprovedNOCCleanup.ps1
```

Or schedule it as a task:

1. Open Task Scheduler.
2. Create a new task.
3. Set the trigger to your desired frequency.
4. Set the action to run PowerShell with the script:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Path\To\ImprovedNOCCleanup.ps1"
```

Save and enable the task.

## Logging

The script generates logs to track the cleanup process. Logs are stored in:

```
C:\Path\To\Logs\NOCCleanup.log
```

## Customization

Modify the script to adjust:

- Retention period for logs and files.

- Additional directories to clean.

- Logging verbosity.