<#
.SYNOPSIS
    Cleans caches and temporary files for various applications and system components.

.DESCRIPTION
    This script cleans caches (e.g. Firefox, Chrome, Edge, Teams, etc.), temporary folders, CBS logs,
    SoftwareDistribution downloads, and additional log files. It exports the list of users, then processes each user.
    Use the -DryRun switch to simulate deletion without actually removing any files.
    
.PARAMETER DryRun
    If specified, the script only logs the actions it would take without deleting files.

.PARAMETER DaysOld
    Specifies the age threshold (in days) for cleaning log files.

.PARAMETER LogFile
    Specifies the location of the log file.

.EXAMPLE
    .\ImprovedNOCCleanup.ps1 -DryRun -DaysOld 3

.NOTES
    Last Revised: 2/12/2024
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [int]$DaysOld = 3,
    [string]$LogFile = "$env:USERPROFILE\cleanup_log.txt"
)

#---------------------------------------------------------------------
# Logging function: Write messages with a timestamp to both the console and log file.
#---------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $Message"
    Add-Content -Path $LogFile -Value $entry
}

#---------------------------------------------------------------------
# Get-FolderSize: Returns the total size (in bytes) of all files under the specified path.
#---------------------------------------------------------------------
function Get-FolderSize {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )
    $size = 0
    if (Test-Path $Path) {
        try {
            $items = Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                if ($item -is [System.IO.FileInfo]) {
                    $size += $item.Length
                }
            }
        } catch {
            Write-Log "Error getting folder size for ${Path}: $_"
        }
    }
    return $size
}

#---------------------------------------------------------------------
# Remove-Items: Remove files/folders matching the given wildcard pattern.
#---------------------------------------------------------------------
function Remove-Items {
    param(
        [Parameter(Mandatory)]
        [string]$PathPattern,
        [switch]$DryRun
    )
    Write-Host "Deleting files matching: $PathPattern" -ForegroundColor Yellow
    Write-Log "Deleting files matching: $PathPattern"
    if (-not $DryRun) {
        try {
            Remove-Item -Path $PathPattern -Recurse -Force -ErrorAction SilentlyContinue -Verbose
        } catch {
            Write-Log "Error deleting items for pattern ${PathPattern}: $_"
        }
    }
}

#---------------------------------------------------------------------
# Clean-LogFiles: Deletes log files older than the specified number of days.
#---------------------------------------------------------------------
function Clean-LogFiles {
    param(
        [Parameter(Mandatory)]
        [string]$TargetFolder,
        [int]$Days,
        [switch]$DryRun
    )
    Write-Host "Cleaning log files in: $TargetFolder" -ForegroundColor Yellow
    Write-Log "Cleaning log files in: $TargetFolder"
    $spaceCleared = 0

    if (Test-Path $TargetFolder) {
        $cutoffDate = (Get-Date).AddDays(-$Days)
        $files = Get-ChildItem $TargetFolder -Recurse -File -ErrorAction SilentlyContinue |
                 Where-Object { $_.Extension -in ".log", ".blg", ".etl" -and $_.LastWriteTime -le $cutoffDate }
        foreach ($file in $files) {
            Write-Host "Deleting file: $($file.FullName)" -ForegroundColor Yellow
            Write-Log "Deleting file: $($file.FullName)"
            $spaceCleared += $file.Length
            if (-not $DryRun) {
                try {
                    Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
                } catch {
                    Write-Log "Error deleting file $($file.FullName): $_"
                }
            }
        }
    } else {
        Write-Host "The folder $TargetFolder doesn't exist!" -ForegroundColor Red
        Write-Log "The folder $TargetFolder doesn't exist!"
    }
    return $spaceCleared
}

#---------------------------------------------------------------------
# Process-UserCache: For a given user, process a set of cache cleanup tasks.
#---------------------------------------------------------------------
function Process-UserCache {
    param(
        [Parameter(Mandatory)]
        [string]$UserName,
        [switch]$DryRun,
        [int]$DaysOld
    )

    # Array of tasks – each task is defined by a description, base path, and an array of wildcard patterns.
    $tasks = @(
        @{
            Description = "Mozilla Firefox Cache"
            BasePath    = "C:\Users\$UserName\AppData\Local\Mozilla\Firefox\Profiles"
            Patterns    = @(
                "*.default\cache\*",
                "*.default\thumbnails\*",
                "*.default\cookies.sqlite",
                "*.default\webappsstore.sqlite",
                "*.default\chromeappsstore.sqlite",
                "*.*\cache2\entries\*"
            )
        },
        @{
            Description = "Google Chrome Cache"
            BasePath    = "C:\Users\$UserName\AppData\Local\Google\Chrome\User Data\Default"
            Patterns    = @(
                "Cache\*",
                "Cache2\entries\*",
                "Cookies",
                "Media Cache",
                "Cookies-Journal"
            )
        },
        @{
            Description = "Toshiba Cache"
            BasePath    = "C:\Users\$UserName\TOSHIBA\eSTUDIOX\UNIDRV\Cache"
            Patterns    = @("*")
        },
        @{
            Description = "Internet Explorer Cache and Temp"
            BasePath    = "C:\Users\$UserName\AppData\Local\Microsoft\Windows"
            Patterns    = @(
                "Temporary Internet Files\*",
                "INetCache\*.*",
                "WER\*",
                "Explorer\*",
                "Temp\*"
            )
        },
        @{
            Description = "Edge Cache"
            BasePath    = "C:\Users\$UserName\AppData\Local\Microsoft\Edge\User Data\Default"
            Patterns    = @(
                "cache\Cache_data\*",
                "Service Worker\CacheStorage\*"
            )
        },
        @{
            Description = "Microsoft Teams Cache"
            BasePath    = "C:\Users\$UserName\AppData\Roaming\Microsoft\Teams"
            Patterns    = @(
                "Cache\*",
                "Service Worker\CacheStorage\*"
            )
        },
        @{
            Description = "Chrome Service Worker Cache"
            BasePath    = "C:\Users\$UserName\AppData\Local\Google\Chrome\User Data\Default\Service Worker\CacheStorage"
            Patterns    = @("*")
        }
    )

    foreach ($task in $tasks) {
        $desc = $task.Description
        $basePath = $task.BasePath
        $patterns = $task.Patterns

        Write-Host "Processing $desc for user: $UserName" -ForegroundColor Green
        Write-Log "Processing $desc for user: $UserName"

        $initialSize = Get-FolderSize -Path $basePath
        if (Test-Path $basePath) {
            foreach ($pattern in $patterns) {
                $fullPattern = Join-Path $basePath $pattern
                Remove-Items -PathPattern $fullPattern -DryRun:$DryRun
            }
        }
        else {
            Write-Host "$desc path does not exist for $UserName" -ForegroundColor Yellow
            Write-Log "$desc path does not exist for $UserName"
        }
        $finalSize = Get-FolderSize -Path $basePath
        $global:TotalSpaceCleared += ($initialSize - $finalSize)
    }

    # --- Special Handling: Outlook Cache (Section 10) ---
    Write-Host "Processing Outlook Cache for user: $UserName" -ForegroundColor Green
    Write-Log "Processing Outlook Cache for user: $UserName"
    $outlookPath = "C:\Users\$UserName\AppData\Local\Microsoft\Outlook"
    $initialSize = Get-FolderSize -Path $outlookPath
    if (Test-Path $outlookPath) {
        $ostFiles = Get-ChildItem $outlookPath -Filter *.ost -File -ErrorAction SilentlyContinue |
                    Where-Object { ($_.LastWriteTime -gt (Get-Date).AddDays(-14)) -and (($_.Length / 1GB) -gt 1) }
        foreach ($file in $ostFiles) {
            Write-Host "Deleting OST file: $($file.FullName)" -ForegroundColor Yellow
            Write-Log "Deleting OST file: $($file.FullName)"
            $global:TotalSpaceCleared += $file.Length
            if (-not $DryRun) {
                try {
                    Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
                } catch {
                    Write-Log "Error deleting OST file $($file.FullName): $_"
                }
            }
        }
    }
    else {
        Write-Host "Outlook path does not exist for user: $UserName" -ForegroundColor Yellow
        Write-Log "Outlook path does not exist for user: $UserName"
    }
    $finalSize = Get-FolderSize -Path $outlookPath
    $global:TotalSpaceCleared += ($initialSize - $finalSize)
}

#=====================================================================
# MAIN SCRIPT EXECUTION
#=====================================================================

# Display header information
Write-Host "#######################################################" -ForegroundColor Yellow
Write-Host "Improved PowerShell Cleanup Script for NOC" -ForegroundColor Green
Write-Host "Last Revised: 7/19/2024" -ForegroundColor Green
Write-Host "#######################################################" -ForegroundColor Yellow
Write-Log "Cleanup Script Started"

if ($DryRun) {
    Write-Host "Dry run mode enabled. No files will be deleted." -ForegroundColor Yellow
    Write-Log "Dry run mode enabled."
}

# --- Section 1: Export list of users ---
$usersCsv = "$env:USERPROFILE\users.csv"
Write-Host "Exporting the list of users to $usersCsv" -ForegroundColor Green
Write-Log "Exporting the list of users to $usersCsv"
Get-ChildItem 'C:\Users' -Directory | Select-Object Name | Export-Csv -Path $usersCsv -NoTypeInformation

if (-not (Test-Path $usersCsv)) {
    Write-Host "User list export failed. Exiting..." -ForegroundColor Red
    Write-Log "User list export failed."
    exit
}

$users = Import-Csv $usersCsv

# Global variable to accumulate total space cleared (in bytes)
$global:TotalSpaceCleared = 0

# --- Sections 3-10: Process user caches ---
foreach ($user in $users) {
    Process-UserCache -UserName $user.Name -DryRun:$DryRun -DaysOld $DaysOld
}

# --- Section 11: System Temp folders ---
Write-Host "Cleaning System Temp folders" -ForegroundColor Green
Write-Log "Cleaning System Temp folders"
$systemTempPaths = @("C:\Windows\Temp", "C:\Windows\Installer", "C:\Windows\Panther")
foreach ($path in $systemTempPaths) {
    $initialSize = Get-FolderSize -Path $path
    if (Test-Path $path) {
        # For Installer folder, remove only *.tmp files; otherwise, all files.
        $pattern = if ($path -eq "C:\Windows\Installer") { "*.tmp" } else { "*.*" }
        Remove-Items -PathPattern (Join-Path $path $pattern) -DryRun:$DryRun
    }
    $finalSize = Get-FolderSize -Path $path
    $global:TotalSpaceCleared += ($initialSize - $finalSize)
}

# --- Section 12: Clearing CBS logs ---
Write-Host "Clearing CBS logs" -ForegroundColor Green
Write-Log "Clearing CBS logs"
$cbsPath = "C:\Windows\Logs\CBS"
$initialSize = Get-FolderSize -Path $cbsPath
if (Test-Path $cbsPath) {
    try {
        Stop-Service -Name TrustedInstaller -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Log "Failed to stop TrustedInstaller: $_"
    }
    Remove-Items -PathPattern (Join-Path $cbsPath "CBS.log") -DryRun:$DryRun
    try {
        Start-Service -Name TrustedInstaller -ErrorAction SilentlyContinue
    } catch {
        Write-Log "Failed to start TrustedInstaller: $_"
    }
}
$finalSize = Get-FolderSize -Path $cbsPath
$global:TotalSpaceCleared += ($initialSize - $finalSize)

# --- Section 13: Clearing SoftwareDistribution\Download ---
Write-Host "Clearing SoftwareDistribution\Download" -ForegroundColor Green
Write-Log "Clearing SoftwareDistribution\Download"
$sdPath = "C:\Windows\SoftwareDistribution\Download"
$initialSize = Get-FolderSize -Path $sdPath
if (Test-Path $sdPath) {
    Remove-Items -PathPattern (Join-Path $sdPath "*.*") -DryRun:$DryRun
}
$finalSize = Get-FolderSize -Path $sdPath
$global:TotalSpaceCleared += ($initialSize - $finalSize)

# --- Section 14: Cleaning additional log files ---
Write-Host "Cleaning additional log files" -ForegroundColor Green
Write-Log "Cleaning additional log files"
$logPaths = @(
    "C:\inetpub\logs\LogFiles",
    "D:\Program Files\Microsoft\Exchange Server\V15\Logging",
    "D:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\ETLTraces",
    "D:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\Logs"
)
foreach ($logPath in $logPaths) {
    $global:TotalSpaceCleared += Clean-LogFiles -TargetFolder $logPath -Days $DaysOld -DryRun:$DryRun
}

# --- Section 15: Running Disk Cleanup ---
Write-Host "Running Disk Cleanup" -ForegroundColor Green
Write-Log "Running Disk Cleanup"
if (-not $DryRun) {
    try {
        Start-Process cleanmgr.exe -ErrorAction SilentlyContinue
    } catch {
        Write-Log "Error starting Disk Cleanup: $_"
    }
}

#--- Final Summary ---
$spaceClearedMB = "{0:N2}" -f ($global:TotalSpaceCleared / 1MB)
Write-Host "All Tasks Done!" -ForegroundColor Green
Write-Host "Total space cleared: $spaceClearedMB MB" -ForegroundColor Green
Write-Log "Total space cleared: $spaceClearedMB MB"
