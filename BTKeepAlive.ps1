if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Check for admin rights
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator!" -ForegroundColor Red
    Write-Host "Right-click the script and select 'Run with PowerShell'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

Clear-Host
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "   Bluetooth Keep-Alive Service Configuration" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host

# Config file path
$configPath = Join-Path $PSScriptRoot "BTConfig.json"

# Get all Bluetooth devices
$btDeviceList = Get-PnpDevice | Where-Object {$_.Class -eq "Bluetooth"} | 
    Select-Object Name, InstanceId |
    Where-Object {$_.Name -notlike "*Radio*" -and $_.Name -notlike "*Adapter*" -and $_.Name -notlike "*Controller*"}

if ($btDeviceList.Count -eq 0) {
    Write-Host "No Bluetooth devices found!" -ForegroundColor Red
    Write-Host "Please make sure your Bluetooth devices are paired with Windows." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

# Display devices with numbers
Write-Host "Available Bluetooth Devices:" -ForegroundColor Green
Write-Host "---------------------------------------------------"
for ($i = 0; $i -lt $btDeviceList.Count; $i++) {
    Write-Host ("[{0}] {1}" -f ($i + 1), $btDeviceList[$i].Name)
}
Write-Host "---------------------------------------------------"
Write-Host
Write-Host "Enter the numbers of the devices you want to keep alive." -ForegroundColor Yellow
Write-Host "For multiple devices, separate numbers with commas (e.g., 1,3,4)" -ForegroundColor Yellow
Write-Host "Or press Enter without selection to exit." -ForegroundColor Yellow
Write-Host

$selection = Read-Host "Select devices"

if ([string]::IsNullOrWhiteSpace($selection)) {
    Write-Host "No devices selected. Exiting..." -ForegroundColor Yellow
    exit
}

# Parse selection and validate
try {
    $selectedNumbers = $selection -split '[,\s]+' | 
        Where-Object { $_ -match '^\d+$' } | 
        ForEach-Object { [int]$_ }

    $selectedDevices = @()
    foreach ($num in $selectedNumbers) {
        if ($num -lt 1 -or $num -gt $btDeviceList.Count) {
            Write-Host "Invalid selection: $num" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit
        }
        $selectedDevices += $btDeviceList[$num - 1]
    }
}
catch {
    Write-Host "Invalid input format!" -ForegroundColor Red
    Write-Host "Please use numbers separated by commas." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

if ($selectedDevices.Count -eq 0) {
    Write-Host "No valid devices selected!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# Confirm selection
Write-Host
Write-Host "Selected Devices:" -ForegroundColor Green
Write-Host "---------------------------------------------------"
foreach ($device in $selectedDevices) {
    Write-Host $device.Name -ForegroundColor Cyan
}
Write-Host "---------------------------------------------------"
Write-Host

$confirm = Read-Host "Proceed with these devices? (Y/N)"
if ($confirm -notlike "Y*") {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    exit
}

# Save configuration to JSON file
$config = @{
    devices = @($selectedDevices | ForEach-Object {
        @{
            id = $_.InstanceId
            name = $_.Name
        }
    })
}
$config | ConvertTo-Json | Set-Content -Path $configPath

# Create service directory
$serviceDir = "C:\BTService"
New-Item -Path $serviceDir -ItemType Directory -Force | Out-Null

# Copy files to service directory
Copy-Item -Path $PSScriptRoot\BTConfig.json -Destination $serviceDir -Force
Copy-Item -Path $PSScriptRoot\BTService.ps1 -Destination $serviceDir -Force
Copy-Item -Path $PSScriptRoot\BTUninstallService.ps1 -Destination $serviceDir -Force
Copy-Item -Path $MyInvocation.MyCommand.Path -Destination $serviceDir -Force

# Download and setup NSSM if not present
$nssmPath = "$serviceDir\nssm.exe"
if (!(Test-Path $nssmPath)) {
    Write-Host "Downloading NSSM (Service Manager)..." -ForegroundColor Yellow
    $webClient = New-Object System.Net.WebClient
    $nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
    $nssmZip = "$serviceDir\nssm.zip"
    $webClient.DownloadFile($nssmUrl, $nssmZip)
    Expand-Archive -Path $nssmZip -DestinationPath "$serviceDir\nssm-temp"
    Copy-Item "$serviceDir\nssm-temp\nssm-2.24\win64\nssm.exe" -Destination $nssmPath
    Remove-Item "$serviceDir\nssm-temp" -Recurse -Force
    Remove-Item $nssmZip -Force
}

# Service details
$serviceName = "BTKeepAliveService"
$serviceDescription = "Maintains Bluetooth audio devices in active state and prevents standby mode"

# Check if service exists
$serviceExists = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

if ($serviceExists) {
    Write-Host "Updating existing service..." -ForegroundColor Yellow
    Stop-Service $serviceName -Force
    # Add a small delay to ensure service is fully stopped
    Start-Sleep -Seconds 2
    $service = Get-Service $serviceName
    if ($service.Status -ne 'Stopped') {
        Write-Host "Failed to stop service completely. Current status: $($service.Status)" -ForegroundColor Red
        exit
    }
    
    # Clean up old log files after stopping the service
    if (Test-Path "$serviceDir\service.log") { Remove-Item "$serviceDir\service.log" -Force }
    if (Test-Path "$serviceDir\error.log") { Remove-Item "$serviceDir\error.log" -Force }
} else {
    Write-Host "Installing new service..." -ForegroundColor Yellow
    
    # For new installations, just clean up any leftover logs if they exist
    if (Test-Path "$serviceDir\service.log") { Remove-Item "$serviceDir\service.log" -Force }
    if (Test-Path "$serviceDir\error.log") { Remove-Item "$serviceDir\error.log" -Force }
}

    # Construct the proper command line
    $powershellPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $serviceScriptPath = "$serviceDir\BTService.ps1"
    
    Write-Host "Creating service..." -ForegroundColor Yellow
    
	# Install service with proper parameters
	& $nssmPath install $serviceName $powershellPath
	& $nssmPath set $serviceName AppDirectory $serviceDir
	& $nssmPath set $serviceName AppParameters "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$serviceScriptPath`""
	& $nssmPath set $serviceName Description $serviceDescription
	& $nssmPath set $serviceName DisplayName "Bluetooth Keep-Alive Service"
	& $nssmPath set $serviceName ObjectName "LocalSystem"
	& $nssmPath set $serviceName Start SERVICE_AUTO_START
	& $nssmPath set $serviceName AppStdout "$serviceDir\service.log"
	& $nssmPath set $serviceName AppStderr "$serviceDir\error.log"
	& $nssmPath set $serviceName AppRotateFiles 1
	& $nssmPath set $serviceName AppRotateBytes 1048576
    
    Write-Host "Service installation completed" -ForegroundColor Green

# Start the service
try {
    Start-Service $serviceName
    $service = Get-Service $serviceName
    Write-Host "`nService Status: $($service.Status)" -ForegroundColor Green
} catch {
    Write-Host "`nError starting service: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Check logs at: $serviceDir\error.log" -ForegroundColor Yellow
}

Write-Host "`nYou can manage the service in several ways:"
Write-Host "1. Services.msc (Windows Services)" -ForegroundColor Yellow
Write-Host "2. PowerShell commands:" -ForegroundColor Yellow
Write-Host "   - Start-Service $serviceName"
Write-Host "   - Stop-Service $serviceName"
Write-Host "   - Restart-Service $serviceName"

Read-Host "`nPress Enter to exit"