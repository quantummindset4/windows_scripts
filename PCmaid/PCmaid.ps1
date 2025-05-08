
Add-Type -AssemblyName Microsoft.VisualBasic

function Get-UsedSpaceMB {
    $drive = Get-PSDrive -Name C
    return [math]::Round($drive.Used / 1MB, 2)
}

# 0. Confirm intent
$confirmation = [Microsoft.VisualBasic.Interaction]::MsgBox(
    "This will clean ALL system, browser, memory, event, and app garbage - like a digital purge. Proceed?",
    "YesNo,Information", "GOD MODE BOOSTER"
)
if ($confirmation -ne "Yes") { exit }

# 0.1 Initial space usage
$beforeMB = Get-UsedSpaceMB

Write-Host "`n[+] GOD MODE Optimization Starting..." -ForegroundColor Cyan

# Placeholder script header
# 1.1 Prompt to Clear Browser History for Each Browser Individually
$browserActions = @{
    "Chrome" = @{
        Process = "chrome"
        History = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\History"
        Cookies = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cookies"
    }
    "Edge" = @{
        Process = "msedge"
        History = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\History"
        Cookies = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cookies"
    }
    "Firefox" = @{
        Process = "firefox"
        History = "$env:APPDATA\Mozilla\Firefox\Profiles\*\places.sqlite"
    }
    "Brave" = @{
        Process = "brave"
        History = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default\History"
    }
    "Opera" = @{
        Process = "opera"
        History = "$env:APPDATA\Opera Software\Opera Stable\History"
    }
    "Vivaldi" = @{
        Process = "vivaldi"
        History = "$env:LOCALAPPDATA\Vivaldi\User Data\Default\History"
    }
}

foreach ($browser in $browserActions.Keys) {
    $response = [Microsoft.VisualBasic.Interaction]::MsgBox(
        "Do you want to CLOSE and CLEAR $browser history/cookies?",
        "YesNo,Question", "$browser Cleaner"
    )
    if ($response -eq "Yes") {
        Write-Host "  -> Closing $browser and wiping history..."
        $proc = $browserActions[$browser]["Process"]
        Get-Process $proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

        $browserActions[$browser].Keys | ForEach-Object {
            if ($_ -ne "Process") {
                $target = $browserActions[$browser][$_]
                Remove-Item $target -Force -Recurse -ErrorAction SilentlyContinue
            }
        }
    }
}
# Rest of script...

# 2. Flush DNS
Write-Host "[2] Flushing DNS cache..."
ipconfig /flushdns | Out-Null

# 3. Kill background bloatware
$junk = @("OneDrive", "SkypeApp", "YourPhone", "GameBar", "Cortana", "Widgets", "Teams", "Zoom", "Spotify", "SearchApp", "Photos", "Xbox", "EdgeUpdate")
Write-Host "[3] Terminating background bloat..."
foreach ($p in $junk) {
    Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$p*" } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# 4. Stop background services
$servicesToStop = @("SysMain", "WSearch", "DiagTrack", "Fax", "RetailDemo")
Write-Host "[4] Stopping unnecessary services..."
foreach ($svc in $servicesToStop) {
    try {
        Stop-Service -Name $svc -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
    } catch {
        Write-Host "    Could not stop service: $svc (may already be stopped or protected)" -ForegroundColor Yellow
    }
}

# 5. Empty Recycle Bin
Write-Host "[5] Emptying Recycle Bin..."
try {
    $shell = New-Object -ComObject Shell.Application
    $recycleBin = $shell.Namespace(0xA)
    $recycleBin.Items() | ForEach-Object { Remove-Item $_.Path -Recurse -Force -ErrorAction SilentlyContinue }
} catch {
    Write-Host "    Failed to empty recycle bin." -ForegroundColor Yellow
}

# 6. Optional: Wipe Event Logs
$clearLogs = [Microsoft.VisualBasic.Interaction]::MsgBox(
    "Do you want to WIPE all Event Logs? (Saves space, but removes logs)", "YesNo", "Log Cleaner"
)
if ($clearLogs -eq "Yes") {
    Write-Host "[6] Clearing Event Logs..."
    wevtutil el | ForEach-Object {
        try {
            wevtutil cl "$_"
        } catch {
            Write-Host "    Failed to clear log: $_ (likely access denied)" -ForegroundColor DarkYellow
        }
    }
}

# 7. Flush Standby RAM if tool is present
$standbyPath = "$PSScriptRoot\emptystandbylist.exe"
if (Test-Path $standbyPath) {
    Write-Host "[7] Cleaning Standby RAM..."
    Start-Process $standbyPath -ArgumentList workingsets -WindowStyle Hidden
} else {
    $url = "https://wj32.org/wp/software/empty-standby-list/"
    [Microsoft.VisualBasic.Interaction]::MsgBox(
        "To clean RAM: Download 'emptystandbylist.exe' from:`n$url`nand place it next to this EXE.",
        "OKOnly,Information", "RAM Cleanup Suggestion"
    )
}

# 8. Optional Safe Boosters
Write-Host "[8] Running safe additional boosters..."

# Windows Update Cleanup
Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
Start-Service -Name wuauserv -ErrorAction SilentlyContinue

# Thumbnail & Icon Cache
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache_*.db" -Force -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\IconCache.db" -Force -ErrorAction SilentlyContinue

# CBS Logs & Crash Dumps
Remove-Item "C:\Windows\Logs\CBS\CBS*.log" -Force -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\CrashDumps\*" -Force -Recurse -ErrorAction SilentlyContinue

# WER Cleanup
Remove-Item "C:\ProgramData\Microsoft\Windows\WER\*" -Recurse -Force -ErrorAction SilentlyContinue

# Defrag if not SSD
$driveType = (Get-PhysicalDisk | Where-Object { $_.FriendlyName -like "*C*" }).MediaType
if ($driveType -ne "SSD") {
    Write-Host "[8.1] Running defragmentation..."
    defrag C: /O
} else {
    Write-Host "[8.1] Skipped defrag (SSD detected)."
}

# Disk Cleanup (Microsoft Safe Mode)
Start-Process cleanmgr.exe -ArgumentList "/verylowdisk" -WindowStyle Hidden

# 9. Final Disk Usage & Summary
$afterMB = Get-UsedSpaceMB
$freedMB = [math]::Round($beforeMB - $afterMB, 2)
$freedGB = [math]::Round($freedMB / 1024, 2)

# 10. Ask for restart
$reboot = [Microsoft.VisualBasic.Interaction]::MsgBox(
    "GOD MODE Boost Complete!`nDisk space freed: $freedMB MB ($freedGB GB)`n`nDo you want to restart your PC now?",
    "YesNo,Question", "Restart Confirmation"
)

if ($reboot -eq "Yes") {
    Write-Host "[*] Restarting system..." -ForegroundColor Cyan
    shutdown /r /t 5 /f
} else {
    Write-Host "[*] Reboot skipped. You can restart later for full effect." -ForegroundColor Yellow
}
