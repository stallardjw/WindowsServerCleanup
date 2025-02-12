Write-Host -ForegroundColor Yellow "#######################################################"
Write-Host -ForegroundColor Green "Improved Powershell commands For NOC cleanup"
Write-Host -ForegroundColor Green "Last Revised: 7/19/2024"
Write-Host -ForegroundColor Yellow "#######################################################"

Write-Host -ForegroundColor Green "CHANGE_LOG:"
Write-Host -ForegroundColor Green "SECTION 1: Getting the list of users"
Write-Host -ForegroundColor Green "SECTION 2: Beginning Script..."
Write-Host -ForegroundColor Green "SECTION 3: Clearing Mozilla Firefox Caches"
Write-Host -ForegroundColor Green "SECTION 4: Clearing Google Chrome Caches"
Write-Host -ForegroundColor Green "SECTION 5: Clearing Toshiba Cache"
Write-Host -ForegroundColor Green "SECTION 6: Clearing Internet Explorer Caches"
Write-Host -ForegroundColor Green "SECTION 7: Edge Caches"
Write-Host -ForegroundColor Green "SECTION 8: Clearing Microsoft Team Caches"
Write-Host -ForegroundColor Green "SECTION 9: Clearing Chrome Service Worker Cache"
Write-Host -ForegroundColor Green "SECTION 10: Clearing Outlook Cache"
Write-Host -ForegroundColor Green "SECTION 11: System Temp folders"
Write-Host -ForegroundColor Green "SECTION 12: Clearing CBS logs"
Write-Host -ForegroundColor Green "SECTION 13: Clearing softwaredistribution\download"
Write-Host -ForegroundColor Green "SECTION 14: Cleaning Log Files"
Write-Host -ForegroundColor Green "SECTION 15: Running Disk Cleanup"
Write-Host -ForegroundColor Yellow "#######################################################"

$totalSpaceCleared = 0
$logFile = "C:\\users\\$env:USERNAME\\cleanup_log.txt"
$dryRun = $false
$days = 3

# Function to log messages
Function Write-Log {
    param (
        [string]$message
    )
    Add-Content -Path $logFile -Value $message
}

# Function to get folder size
Function Get-FolderSize {
    param ($Path)
    $size = 0
    if (Test-Path $Path) {
        $items = Get-ChildItem -Path $Path -Recurse
        foreach ($item in $items) {
            $size += $item.Length
        }
    }
    return $size
}

# Function to clean log files
Function CleanLogfiles {
    param ($TargetFolder)
    Write-Host -ForegroundColor Yellow -BackgroundColor Cyan $TargetFolder
    Write-Log "Cleaning log files in $TargetFolder"
    $spaceCleared = 0

    if (Test-Path $TargetFolder) {
        $Now = Get-Date
        $LastWrite = $Now.AddDays(-$days)
        $Files = Get-ChildItem $TargetFolder -Recurse | Where-Object { $_.Name -like "*.log" -or $_.Name -like "*.blg" -or $_.Name -like "*.etl" } | Where-Object { $_.LastWriteTime -le $LastWrite } | Select-Object FullName, Length
        foreach ($File in $Files) {
            $FullFileName = $File.FullName
            $spaceCleared += $File.Length
            Write-Host "Deleting file $FullFileName" -ForegroundColor "yellow"
            Write-Log "Deleting file $FullFileName"
            if (-not $dryRun) {
                Remove-Item $FullFileName -ErrorAction SilentlyContinue | Out-Null
            }
        }
    } else {
        Write-Host "The folder $TargetFolder doesn't exist! Check the folder path!" -ForegroundColor "red"
        Write-Log "The folder $TargetFolder doesn't exist! Check the folder path!"
    }
    return $spaceCleared
}

# Check if dry run parameter is passed
if ($args -contains '-dryrun') {
    $dryRun = $true
    Write-Host -ForegroundColor Yellow "Dry run mode enabled. No files will be deleted."
    Write-Log "Dry run mode enabled."
}

# SECTION 1: Getting the list of users
Write-Host -ForegroundColor Green "SECTION 1: Getting the list of users"
Write-Host -ForegroundColor Yellow "Exporting the list of users to C:\\users\\$env:USERNAME\\users.csv"
Write-Log "Exporting the list of users to C:\\users\\$env:USERNAME\\users.csv"
Get-ChildItem 'C:\Users' -Directory | Select-Object Name | Export-Csv -Path "C:\\users\\$env:USERNAME\\users.csv" -NoTypeInformation

$list = Test-Path "C:\\users\\$env:USERNAME\\users.csv"

# SECTION 2: Beginning Script...
Write-Host -ForegroundColor Green "SECTION 2: Beginning Script..."
Write-Log "Beginning script..."

if ($list) {
    $users = Import-Csv "C:\\users\\$env:USERNAME\\users.csv"

    # SECTION 3: Clearing Mozilla Firefox Caches
    Write-Host -ForegroundColor Green "SECTION 3: Clearing Mozilla Firefox Caches"
    Write-Log "SECTION 3: Clearing Mozilla Firefox Caches"
    foreach ($user in $users.Name) {
        $initialSize = 0
        $finalSize = 0
        $initialSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles"
        $folderpath = Test-Path -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles"
        
        if ($folderpath) {
            if (-not $dryRun) {
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles\\*.default\\cache\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles\\*.default\\cache\\*.*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles\\*.*\\cache2\\entries\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles\\*.default\\thumbnails\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles\\*.default\\cookies.sqlite" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles\\*.default\\webappsstore.sqlite" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles\\*.default\\chromeappsstore.sqlite" -Recurse -Force -EA SilentlyContinue -Verbose
            }
            Write-Output "Cleared Mozilla Firefox cache for $user"
            Write-Log "Cleared Mozilla Firefox cache for $user"
        } else {
            Write-Output "Mozilla Firefox cache doesn't exist for $user"
            Write-Log "Mozilla Firefox cache doesn't exist for $user"
        }
        $finalSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Mozilla\\Firefox\\Profiles"
        $totalSpaceCleared += ($initialSize - $finalSize)
    }
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing Mozilla Firefox Caches"

    # SECTION 4: Clearing Google Chrome Caches
    Write-Host -ForegroundColor Green "SECTION 4: Clearing Google Chrome Caches"
    Write-Log "SECTION 4: Clearing Google Chrome Caches"
    foreach ($user in $users.Name) {
        $initialSize = 0
        $finalSize = 0
        $initialSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache"
        $folderpath = Test-Path -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache"
        
        if ($folderpath) {
            if (-not $dryRun) {
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache2\\entries\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cookies" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Media Cache" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cookies-Journal" -Recurse -Force -EA SilentlyContinue -Verbose
                # Comment out the following line to remove the Chrome Write Font Cache too.
                # Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\ChromeDWriteFontCache" -Recurse -Force -EA SilentlyContinue -Verbose
            }
            Write-Output "Cleared Google Chrome cache for $user"
            Write-Log "Cleared Google Chrome cache for $user"
        } else {
            Write-Output "Google Chrome cache doesn't exist for $user"
            Write-Log "Google Chrome cache doesn't exist for $user"
        }
        $finalSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache"
        $totalSpaceCleared += ($initialSize - $finalSize)
    }
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing Google Chrome Caches"

    # SECTION 5: Clearing Toshiba Cache
    Write-Host -ForegroundColor Green "SECTION 5: Clearing Toshiba Cache"
    Write-Log "SECTION 5: Clearing Toshiba Cache"
    foreach ($user in $users.Name) {
        $initialSize = 0
        $finalSize = 0
        $initialSize += Get-FolderSize -Path "C:\\Users\\$user\\TOSHIBA\\eSTUDIOX\\UNIDRV\\Cache"
        $folderpath = Test-Path -Path "C:\\Users\\$user\\TOSHIBA\\eSTUDIOX\\UNIDRV\\Cache"
        
        if ($folderpath) {
            if (-not $dryRun) {
                Remove-Item -Path "C:\\Users\\$user\\TOSHIBA\\eSTUDIOX\\UNIDRV\\Cache\\*" -Recurse -Force -EA SilentlyContinue -Verbose
            }
            Write-Output "Cleared Toshiba cache for $user"
            Write-Log "Cleared Toshiba cache for $user"
        } else {
            Write-Output "Toshiba cache doesn't exist for $user"
            Write-Log "Toshiba cache doesn't exist for $user"
        }
        $finalSize += Get-FolderSize -Path "C:\\Users\\$user\\TOSHIBA\\eSTUDIOX\\UNIDRV\\Cache"
        $totalSpaceCleared += ($initialSize - $finalSize)
    }
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing Toshiba Cache"

    # SECTION 6: Clearing Internet Explorer Caches
    Write-Host -ForegroundColor Green "SECTION 6: Clearing Internet Explorer Caches"
    Write-Log "SECTION 6: Clearing Internet Explorer Caches"
    foreach ($user in $users.Name) {
        $initialSize = 0
        $finalSize = 0
        $initialSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Windows\\Temporary Internet Files"
        $folderpath = Test-Path -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Windows\\Temporary Internet Files"
        
        if ($folderpath) {
            if (-not $dryRun) {
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Windows\\Temporary Internet Files\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Windows\\INetCache\\*.*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Windows\\WER\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Windows\\Explorer\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Temp\\*" -Recurse -Force -EA SilentlyContinue -Verbose
            }
            Write-Output "Cleared Internet Explorer cache for $user"
            Write-Log "Cleared Internet Explorer cache for $user"
        } else {
            Write-Output "Internet Explorer cache doesn't exist for $user"
            Write-Log "Internet Explorer cache doesn't exist for $user"
        }
        $finalSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Windows\\Temporary Internet Files"
        $totalSpaceCleared += ($initialSize - $finalSize)
    }
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing Internet Explorer Caches"

    # SECTION 7: Edge Caches
    Write-Host -ForegroundColor Green "SECTION 7: Edge Caches"
    Write-Log "SECTION 7: Edge Caches"
    foreach ($user in $users.Name) {
        $initialSize = 0
        $finalSize = 0
        $initialSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\cache\\Cache_data"
        $folderpath = Test-Path -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\cache\\Cache_data"
        
        if ($folderpath) {
            if (-not $dryRun) {
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\cache\\Cache_data\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Service Worker\\CacheStorage\\*" -Recurse -Force -EA SilentlyContinue -Verbose
            }
            Write-Output "Cleared Edge cache for $user"
            Write-Log "Cleared Edge cache for $user"
        } else {
            Write-Output "Edge cache doesn't exist for $user"
            Write-Log "Edge cache doesn't exist for $user"
        }
        $finalSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\cache\\Cache_data"
        $totalSpaceCleared += ($initialSize - $finalSize)
    }
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing Edge Caches"

    # SECTION 8: Clearing Microsoft Team Caches
    Write-Host -ForegroundColor Green "SECTION 8: Clearing Microsoft Team Caches"
    Write-Log "SECTION 8: Clearing Microsoft Team Caches"
    foreach ($user in $users.Name) {
        $initialSize = 0
        $finalSize = 0
        $initialSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Roaming\\Microsoft\\Teams\\Cache"
        $folderpath = Test-Path -Path "C:\\Users\\$user\\AppData\\Roaming\\Microsoft\\Teams\\Cache"
        
        if ($folderpath) {
            if (-not $dryRun) {
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Roaming\\Microsoft\\Teams\\Cache\\*" -Recurse -Force -EA SilentlyContinue -Verbose
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Roaming\\Microsoft\\Teams\\Service Worker\\CacheStorage\\*" -Recurse -Force -EA SilentlyContinue -Verbose
            }
            Write-Output "Cleared Microsoft Team cache for $user"
            Write-Log "Cleared Microsoft Team cache for $user"
        } else {
            Write-Output "Microsoft Team cache doesn't exist for $user"
            Write-Log "Microsoft Team cache doesn't exist for $user"
        }
        $finalSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Roaming\\Microsoft\\Teams\\Cache"
        $totalSpaceCleared += ($initialSize - $finalSize)
    }
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing Microsoft Team Caches"

    # SECTION 9: Clearing Chrome Service Worker Cache
    Write-Host -ForegroundColor Green "SECTION 9: Clearing Chrome Service Worker Cache"
    Write-Log "SECTION 9: Clearing Chrome Service Worker Cache"
    foreach ($user in $users.Name) {
        $initialSize = 0
        $finalSize = 0
        $initialSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Service Worker\\CacheStorage"
        $folderpath = Test-Path -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Service Worker\\CacheStorage"
        
        if ($folderpath) {
            if (-not $dryRun) {
                Remove-Item -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Service Worker\\CacheStorage\\*" -Recurse -Force -EA SilentlyContinue -Verbose
            }
            Write-Output "Cleared Chrome Service Worker cache for $user"
            Write-Log "Cleared Chrome Service Worker cache for $user"
        } else {
            Write-Output "Chrome Service Worker cache doesn't exist for $user"
            Write-Log "Chrome Service Worker cache doesn't exist for $user"
        }
        $finalSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Service Worker\\CacheStorage"
        $totalSpaceCleared += ($initialSize - $finalSize)
    }
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing Chrome Service Worker Cache"

    # SECTION 10: Clearing Outlook Cache
    Write-Host -ForegroundColor Green "SECTION 10: Clearing Outlook Cache"
    Write-Log "SECTION 10: Clearing Outlook Cache"
    foreach ($user in $users.Name) {
        $initialSize = 0
        $finalSize = 0
        $initialSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Outlook"
        $folderpath = Test-Path -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Outlook"
        
        if ($folderpath) {
            Get-ChildItem "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Outlook" -Filter *.ost | Where-Object { ($_.LastWriteTime -gt (Get-Date).AddDays(-14)) -and ($_.Length / 1GB -gt 1) } | ForEach-Object { 
                $totalSpaceCleared += $_.Length
                if (-not $dryRun) {
                    Remove-Item $_.FullName -ErrorAction SilentlyContinue
                }
            }
            Write-Output "Deleted OST file for $user"
            Write-Log "Deleted OST file for $user"
        } else {
            Write-Output "OST file doesn't exist or meet criteria for $user"
            Write-Log "OST file doesn't exist or meet criteria for $user"
        }
        $finalSize += Get-FolderSize -Path "C:\\Users\\$user\\AppData\\Local\\Microsoft\\Outlook"
        $totalSpaceCleared += ($initialSize - $finalSize)
    }
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing Outlook Cache"
    
    # SECTION 11: System Temp folders
    Write-Host -ForegroundColor Green "SECTION 11: System Temp folders"
    Write-Log "SECTION 11: System Temp folders"
    Write-Host -ForegroundColor Yellow "Clearing temp folders"
    Write-Log "Clearing temp folders"
    $initialSize = 0
    $finalSize = 0
    $initialSize += Get-FolderSize -Path "C:\\windows\\temp\\"
    if (-not $dryRun) {
        Remove-Item -Path C:\\windows\\temp\\*.* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
    }
    $finalSize += Get-FolderSize -Path "C:\\windows\\temp\\"
    $totalSpaceCleared += ($initialSize - $finalSize)

    $initialSize += Get-FolderSize -Path "C:\\Windows\\Installer\\"
    if (-not $dryRun) {
        Remove-Item -Path C:\\Windows\\Installer\\*.tmp -Recurse -Force -ErrorAction SilentlyContinue -Verbose
    }
    $finalSize += Get-FolderSize -Path "C:\\Windows\\Installer\\"
    $totalSpaceCleared += ($initialSize - $finalSize)

    $initialSize += Get-FolderSize -Path "C:\\Windows\\Panther\\"
    if (-not $dryRun) {
        Remove-Item -Path C:\\Windows\\Panther\\*.tmp -Recurse -Force -ErrorAction SilentlyContinue -Verbose
    }
    $finalSize += Get-FolderSize -Path "C:\\Windows\\Panther\\"
    $totalSpaceCleared += ($initialSize - $finalSize)
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing System Temp folders"

    # SECTION 12: Clearing CBS logs
    Write-Host -ForegroundColor Green "SECTION 12: Clearing CBS logs"
    Write-Log "SECTION 12: Clearing CBS logs"
    Write-Host -ForegroundColor Yellow "Clearing CBS logs"
    Write-Log "Clearing CBS logs"
    Stop-Service TrustedInstaller -Force
    $initialSize = 0
    $finalSize = 0
    $initialSize += Get-FolderSize -Path "C:\\Windows\\Logs\\CBS\\"
    if (-not $dryRun) {
        Remove-Item -Path C:\\Windows\\Logs\\CBS\\CBS.log -Recurse -Force -ErrorAction SilentlyContinue -Verbose
    }
    Start-Service TrustedInstaller
    $finalSize += Get-FolderSize -Path "C:\\Windows\\Logs\\CBS\\"
    $totalSpaceCleared += ($initialSize - $finalSize)
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing CBS logs"
    
    # SECTION 13: Clearing softwaredistribution\download
    Write-Host -ForegroundColor Green "SECTION 13: Clearing softwaredistribution\download"
    Write-Log "SECTION 13: Clearing softwaredistribution\download"
    Write-Host -ForegroundColor Yellow "Clearing softwaredistribution\download"
    Write-Log "Clearing softwaredistribution\download"
    $initialSize = 0
    $finalSize = 0
    $initialSize += Get-FolderSize -Path "C:\\windows\\softwaredistribution\\download\\"
    if (-not $dryRun) {
        Remove-Item -Path C:\\windows\\softwaredistribution\\download\\*.* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
    }
    $finalSize += Get-FolderSize -Path "C:\\windows\\softwaredistribution\\download\\"
    $totalSpaceCleared += ($initialSize - $finalSize)
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done clearing softwaredistribution\download"

    # SECTION 14: Cleaning Log Files
    Write-Host -ForegroundColor Green "SECTION 14: Cleaning Log Files"
    Write-Log "SECTION 14: Cleaning Log Files"

    $IISLogPath = "C:\\inetpub\\logs\\LogFiles\\"
    $ExchangeLoggingPath = "D:\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\"
    $ETLLoggingPath = "D:\\Program Files\\Microsoft\\Exchange Server\\V15\\Bin\\Search\\Ceres\\Diagnostics\\ETLTraces\\"
    $ETLLoggingPath2 = "D:\\Program Files\\Microsoft\\Exchange Server\\V15\\Bin\\Search\\Ceres\\Diagnostics\\Logs\\"

    $totalSpaceCleared += CleanLogfiles $IISLogPath
    $totalSpaceCleared += CleanLogfiles $ExchangeLoggingPath
    $totalSpaceCleared += CleanLogfiles $ETLLoggingPath
    $totalSpaceCleared += CleanLogfiles $ETLLoggingPath2
    Write-Host -ForegroundColor Yellow "Done..."
    Write-Log "Done cleaning log files"

    # SECTION 15: Running Disk Cleanup
    Write-Host -ForegroundColor Green "SECTION 15: Running Disk Cleanup"
    Write-Log "SECTION 15: Running Disk Cleanup"
    Write-Host -ForegroundColor Green "Running Disk Cleanup..."
    Write-Log "Running Disk Cleanup..."
    if (-not $dryRun) {
        Start-Process cleanmgr.exe
    }

    Write-Host -ForegroundColor Green "All Tasks Done!"
    Write-Host -ForegroundColor Green ("Total space cleared: {0:N2} MB" -f ($totalSpaceCleared / 1MB))
    Write-Log ("Total space cleared: {0:N2} MB" -f ($totalSpaceCleared / 1MB))
} else {
    Write-Host -ForegroundColor Yellow "Session Cancelled"
    Write-Log "Session Cancelled"
    Exit
}
