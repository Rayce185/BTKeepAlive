if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }


# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator!" -ForegroundColor Red
    Write-Host "Right-click the script and select 'Run with PowerShell'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

$serviceName = "BTKeepAliveService"

Write-Host "Uninstalling Bluetooth Keep-Alive Service..." -ForegroundColor Yellow

# Stop the service if running
if (Get-Service $serviceName -ErrorAction SilentlyContinue) {
    Write-Host "Stopping service..." -ForegroundColor Yellow
    Stop-Service $serviceName -Force
    
    # Use NSSM to remove the service
    $nssmPath = "C:\BTService\nssm.exe"
    if (Test-Path $nssmPath) {
        Write-Host "Removing service..." -ForegroundColor Yellow
        & $nssmPath remove $serviceName confirm
    }
    
    # Clean up files
    Write-Host "Cleaning up files..." -ForegroundColor Yellow
    Remove-Item "C:\BTService" -Recurse -Force -ErrorAction SilentlyContinue
    
    Write-Host "Service removed successfully!" -ForegroundColor Green
} else {
    Write-Host "Service not found!" -ForegroundColor Red
}

Read-Host "Press Enter to exit"