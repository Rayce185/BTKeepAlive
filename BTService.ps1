if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit 
}

# Enable detailed error logging
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

try {
    Write-Host "Service script starting..."
    Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-Host "Running as user: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Host "Script Path: $PSCommandPath"
    Write-Host "Working Directory: $PWD"

    # Setup logging
    if (![System.Diagnostics.EventLog]::SourceExists("BTKeepAlive")) {
        New-EventLog -LogName "Application" -Source "BTKeepAlive"
    }

    function Write-ServiceLog {
    param(
        [string]$Message,
        [string]$Type = "Information"
    )
    # Write to Event Log
    Write-EventLog -LogName "Application" -Source "BTKeepAlive" -EventId 1 -EntryType $Type -Message $Message
    
    # Write to file with timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    Add-Content -Path "$PSScriptRoot\service.log" -Value $logMessage
    
    # If it's an error or warning, also log to error.log
    if ($Type -in @("Error", "Warning")) {
        Add-Content -Path "$PSScriptRoot\error.log" -Value $logMessage
    }
}

# Load configuration
$configPath = Join-Path $PSScriptRoot "BTConfig.json"
try {
    $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
    Write-ServiceLog "Configuration loaded successfully"
} catch {
    Write-ServiceLog "Failed to load configuration: $_" -Type "Error"
    exit 1
}

# Import required DLL
$source = @"
using System;
using System.Runtime.InteropServices;

public class BluetoothHelper {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile);

    [DllImport("Kernel32.dll", SetLastError = true)]
    public static extern bool DeviceIoControl(
        IntPtr hDevice,
        uint dwIoControlCode,
        IntPtr lpInBuffer,
        uint nInBufferSize,
        IntPtr lpOutBuffer,
        uint nOutBufferSize,
        ref uint lpBytesReturned,
        IntPtr lpOverlapped);

    // ... other DllImport definitions ...

    // Replace these constants with the new ones
    public const uint DIGCF_PRESENT = 2;
    public const uint DIGCF_DEVICEINTERFACE = 16;
    public const uint FILE_SHARE_READ = 1;
    public const uint FILE_SHARE_WRITE = 2;
    public const uint OPEN_EXISTING = 3;
    public const uint FILE_ATTRIBUTE_NORMAL = 128;
    public const uint ERROR_INSUFFICIENT_BUFFER = 122;
}
"@

Add-Type -TypeDefinition $source

function Get-BluetoothDevicePath {
    param (
        [string]$DeviceId
    )
    
    try {
        Write-ServiceLog "Starting device path lookup for ID: $DeviceId"
        
        # Convert to device interface path format
        $devicePath = $DeviceId.Replace('\', '#')
        $fullPath = "\\.\$devicePath"
        Write-ServiceLog "Created device path: $fullPath"
        
        return $fullPath
        
    } catch {
        Write-ServiceLog "Error getting device path: $($_.Exception.Message)" -Type "Error"
        Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        return $null
    }
}

# First, add the Send-BluetoothKeepAlive function
function Send-BluetoothKeepAlive {
    param (
        [string]$DeviceId,
        [string]$DeviceName
    )
    
    try {
        Write-ServiceLog "Attempting to send keep-alive to device: $DeviceName"
        
        try {
            $btDevice = Get-PnpDevice -InstanceId $DeviceId -ErrorAction Stop
            Write-ServiceLog "Found device via PnP: $($btDevice.FriendlyName)"
            
            # Query power management properties with detailed decoding
            Write-ServiceLog "Querying power management properties..."
            $powerData = Get-PnpDeviceProperty -InstanceId $DeviceId -KeyName DEVPKEY_Device_PowerData
            if ($powerData) {
                $bytes = $powerData.Data
                Write-ServiceLog "Power State: $($bytes[4])"
            }
            
            # Get Bluetooth specific properties
            $btProps = @{
                Address = Get-PnpDeviceProperty -InstanceId $DeviceId -KeyName DEVPKEY_Bluetooth_DeviceAddress
                Flags = Get-PnpDeviceProperty -InstanceId $DeviceId -KeyName DEVPKEY_Bluetooth_DeviceFlags
                Class = Get-PnpDeviceProperty -InstanceId $DeviceId -KeyName DEVPKEY_Bluetooth_ClassOfDevice
            }
            
            if ($btProps.Flags) {
                $flags = $btProps.Flags.Data
                $connected = ($flags -band 0x20) -ne 0
                Write-ServiceLog "Connection Status: $(if ($connected) { 'Connected' } else { 'Disconnected' })"
                
                if ($connected) {
                    # More aggressive keep-alive approach
                    Write-ServiceLog "Sending active keep-alive signal..."
                    
                    # Temporarily disable and quickly re-enable to force a connection refresh
                    Disable-PnpDevice -InstanceId $DeviceId -Confirm:$false
                    Start-Sleep -Milliseconds 100
                    Enable-PnpDevice -InstanceId $DeviceId -Confirm:$false
                    Start-Sleep -Milliseconds 100
                    
                    # Verify device is back online
                    $status = (Get-PnpDevice -InstanceId $DeviceId).Status
                    Write-ServiceLog "Device status after refresh: $status"
                    
                    if ($status -eq "OK") {
                        Write-ServiceLog "Keep-alive refresh successful"
                        return $true
                    } else {
                        Write-ServiceLog "Device not in OK state after refresh" -Type "Warning"
                        return $false
                    }
                }
                else {
                    Write-ServiceLog "Device not connected - no keep-alive sent" -Type "Warning"
                    return $false
                }
            }
            
        } catch {
            Write-ServiceLog "Error accessing device: $($_.Exception.Message)" -Type "Error"
            return $false
        }
    }
    catch {
        Write-ServiceLog "Error in Send-BluetoothKeepAlive: $($_.Exception.Message)" -Type "Error"
        Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        return $false
    }
}

# Then add the Test-StandbyTimeout function
function Test-StandbyTimeout {
    param (
        [string]$DeviceId,
        [string]$DeviceName
    )
    
    try {
        Write-ServiceLog "Starting standby timeout detection for: $DeviceName"
        
        # Wait for device to be inactive (no audio playing)
        Write-ServiceLog "Waiting for device to become inactive (no audio playing)..."
        Start-Sleep -Seconds 30  # Initial wait to ensure no audio is playing
        
        # Now start monitoring for standby
        $startTime = Get-Date
        $lastStatus = "OK"
        $standbyDetected = $false
        
        Write-ServiceLog "Beginning standby monitoring..."
        while (-not $standbyDetected -and ((Get-Date) - $startTime).TotalMinutes -lt 30) {
            $device = Get-PnpDevice -InstanceId $DeviceId -ErrorAction Stop
            $btProps = Get-PnpDeviceProperty -InstanceId $DeviceId -KeyName DEVPKEY_Bluetooth_DeviceFlags
            $connected = ($btProps.Data -band 0x20) -ne 0
            
            if (-not $connected -or $device.Status -ne "OK") {
                if ($lastStatus -eq "OK") {
                    $timeToStandby = ((Get-Date) - $startTime).TotalMinutes
                    Write-ServiceLog "Standby detected after $timeToStandby minutes"
                    $standbyDetected = $true
                    
                    # Calculate keep-alive interval (80% of standby time)
                    $keepAliveInterval = [Math]::Floor($timeToStandby * 0.8)
                    Write-ServiceLog "Setting keep-alive interval to $keepAliveInterval minutes"
                    
                    # Store the calculated interval
                    $config | Add-Member -NotePropertyName "KeepAliveInterval" -NotePropertyValue $keepAliveInterval -Force
                    $config | ConvertTo-Json | Set-Content -Path $configPath
                    
                    return $keepAliveInterval
                }
            }
            
            $lastStatus = $device.Status
            Write-ServiceLog "Status check: $($device.Status), Connected: $connected"
            Start-Sleep -Seconds 30
        }
        
        if (-not $standbyDetected) {
            Write-ServiceLog "No standby detected after 30 minutes - using default interval" -Type "Warning"
            return 5  # Default 5-minute interval
        }
        
    } catch {
        Write-ServiceLog "Error in standby detection: $($_.Exception.Message)" -Type "Error"
        Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        return 5  # Default 5-minute interval on error
    }
}

# Set default interval if not configured
if (-not $config.KeepAliveInterval) {
    Write-ServiceLog "No keep-alive interval configured - using default 2 minutes"
    $interval = 2
    
    # Store the default interval
    $config | Add-Member -NotePropertyName "KeepAliveInterval" -NotePropertyValue $interval -Force
    $config | ConvertTo-Json | Set-Content -Path $configPath
} else {
    $interval = $config.KeepAliveInterval
    Write-ServiceLog "Using configured keep-alive interval: $interval minutes"
}

# Convert interval to seconds for the sleep timer
$sleepSeconds = $interval * 60

# Initialize device tracking
$deviceStatus = @{}
foreach ($device in $config.devices) {
    $deviceStatus[$device.id] = @{
        lastStatus = $false
        lastKeepAlive = [DateTime]::MinValue
    }
}

Write-ServiceLog "Service started, monitoring devices:`n$($config.devices | ForEach-Object { $_.name } | Out-String)"

# Main service loop
try {
    Write-ServiceLog "Entering main service loop"
    while ($true) {
        foreach ($device in $config.devices) {
            try {
                $currentDevice = Get-PnpDevice -InstanceId $device.id -ErrorAction Stop
                Write-ServiceLog "Device status for $($device.name): $($currentDevice.Status)"
                
                if ($currentDevice.Status -eq "OK") {
                    $result = Send-BluetoothKeepAlive -DeviceId $device.id -DeviceName $device.name
                    if ($result) {
                        Write-ServiceLog "Keep-alive signal successfully sent to $($device.name)"
                    }
                }
                else {
                    Write-ServiceLog "Device $($device.name) is not in OK state. Current status: $($currentDevice.Status)" -Type "Warning"
                }
            }
            catch {
                Write-ServiceLog "Error processing device $($device.name): $($_ | Out-String)" -Type "Error"
                Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
                # Don't exit on device error, try again next time
            }
        }
        Write-ServiceLog "Waiting $interval minutes before next keep-alive signal"
        Start-Sleep -Seconds $sleepSeconds
    }
}
catch {
    Write-ServiceLog "Critical error in main service loop: $($_ | Out-String)" -Type "Error"
    Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
    throw  # Re-throw to ensure service shows as failed
}

} catch {
    Write-Host "Critical error in service script: $_"
    Write-Host "Stack trace: $($_.ScriptStackTrace)"
    throw
} finally {
    Stop-Transcript
}