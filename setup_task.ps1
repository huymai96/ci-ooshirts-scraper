# Setup Scheduled Task for CI/Ooshirts Scraper
# Run this script as Administrator

param(
    [string]$TaskName = "Run Receiving Tool Scrapers",
    [string]$Interval = "1"  # Hours between runs
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatFile = Join-Path $ScriptDir "run_scrapers.bat"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "CI/Ooshirts Scraper - Task Setup" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Script Directory: $ScriptDir"
Write-Host "Batch File: $BatFile"
Write-Host "Task Name: $TaskName"
Write-Host "Interval: Every $Interval hour(s)"
Write-Host ""

# Check if batch file exists
if (-not (Test-Path $BatFile)) {
    Write-Host "ERROR: run_scrapers.bat not found!" -ForegroundColor Red
    Write-Host "Make sure this script is in the same directory as run_scrapers.bat"
    exit 1
}

# Check for existing task
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Task '$TaskName' already exists." -ForegroundColor Yellow
    $confirm = Read-Host "Do you want to replace it? (Y/N)"
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        exit 0
    }
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed existing task." -ForegroundColor Green
}

# Create the action
$Action = New-ScheduledTaskAction -Execute $BatFile -WorkingDirectory $ScriptDir

# Create the trigger (repeat every N hours indefinitely)
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval (New-TimeSpan -Hours $Interval) -RepetitionDuration ([TimeSpan]::MaxValue)

# Create settings
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable:$false `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 72) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1)

# Create principal (run as current user with highest privileges)
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

# Register the task
try {
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $Action `
        -Trigger $Trigger `
        -Settings $Settings `
        -Principal $Principal `
        -Description "Scrapes CustomInk and OOShirts orders hourly and uploads to Promos Ink Supply Chain API."
    
    Write-Host ""
    Write-Host "SUCCESS! Task '$TaskName' created." -ForegroundColor Green
    Write-Host ""
    Write-Host "The scraper will run every $Interval hour(s)."
    Write-Host ""
    Write-Host "To run immediately:" -ForegroundColor Cyan
    Write-Host "  schtasks /run /tn `"$TaskName`""
    Write-Host ""
    Write-Host "To check status:" -ForegroundColor Cyan
    Write-Host "  schtasks /query /tn `"$TaskName`" /v /fo LIST"
    Write-Host ""
    Write-Host "To delete:" -ForegroundColor Cyan
    Write-Host "  schtasks /delete /tn `"$TaskName`" /f"
    Write-Host ""
}
catch {
    Write-Host "ERROR creating task: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Try running PowerShell as Administrator." -ForegroundColor Yellow
    exit 1
}
