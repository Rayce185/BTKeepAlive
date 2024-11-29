if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Service script for Bluetooth Keep-Alive
$ErrorActionPreference = "Stop"

# Initial startup logging
$startupLog = Join-Path $PSScriptRoot "startup.log"
try {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Service script starting" | Out-File -FilePath $startupLog
    "PowerShell Version: $($PSVersionTable.PSVersion)" | Out-File -FilePath $startupLog -Append
    "Script Path: $PSCommandPath" | Out-File -FilePath $startupLog -Append
    "Working Directory: $PWD" | Out-File -FilePath $startupLog -Append
} catch {
    # If we can't write to startup.log, try writing to a temp file
    $tempLog = Join-Path $env:TEMP "BTService_startup.log"
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Failed to write to primary log: $_" | Out-File -FilePath $tempLog
}

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

# Import required DLLs
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

    [DllImport("Kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr hObject);
}

public class AudioEndpoint {
    [DllImport("ole32.dll")]
    public static extern int CoCreateInstance(ref Guid clsid, 
        IntPtr pUnkOuter, uint dwClsContext, ref Guid iid, out IntPtr ppv);

    public static readonly Guid CLSID_MMDeviceEnumerator = 
        new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E");
    
    public static readonly Guid IID_IMMDeviceEnumerator = 
        new Guid("A95664D2-9614-4F35-A746-DE8DB63617E6");

    [DllImport("avrt.dll")]
    public static extern IntPtr AvSetMmThreadCharacteristics(string TaskName, ref int TaskIndex);

    [DllImport("avrt.dll")]
    public static extern bool AvRevertMmThreadCharacteristics(IntPtr Handle);

    [DllImport("winmm.dll")]
    public static extern uint timeBeginPeriod(uint uPeriod);

    [DllImport("winmm.dll")]
    public static extern uint timeEndPeriod(uint uPeriod);

    [DllImport("winmm.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern uint waveOutGetNumDevs();

    [DllImport("winmm.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern uint waveOutGetVolume(uint deviceID, out uint volume);

    [DllImport("winmm.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern uint waveOutSetVolume(uint deviceID, uint volume);
}
"@

Add-Type -TypeDefinition $source

# Load configuration
$configPath = Join-Path $PSScriptRoot "BTConfig.json"
try {
    $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
    Write-ServiceLog "Configuration loaded successfully"
    
    # Validate configuration
    if (-not $config.devices -or $config.devices.Count -eq 0) {
        Write-ServiceLog "No devices configured" -Type "Warning"
    } else {
        Write-ServiceLog "Found $($config.devices.Count) configured device(s)"
    }
} catch {
    Write-ServiceLog "Failed to load configuration: $_" -Type "Error"
    exit 1
}

# Test audio device access
try {
    $numDevs = [AudioEndpoint]::waveOutGetNumDevs()
    Write-ServiceLog "Found $numDevs audio output device(s)"
    
    $currentVolume = 0
    $result = [AudioEndpoint]::waveOutGetVolume(0, [ref]$currentVolume)
    Write-ServiceLog "Audio device access test: OK"
} catch {
    Write-ServiceLog "Error accessing audio devices: $($_.Exception.Message)" -Type "Error"
    Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
}

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
            
            # Get Bluetooth properties
            $btProps = Get-PnpDeviceProperty -InstanceId $DeviceId -KeyName DEVPKEY_Bluetooth_DeviceFlags
            if ($btProps) {
                $flags = $btProps.Data
                $connected = ($flags -band 0x20) -ne 0
                Write-ServiceLog "Connection Status: $(if ($connected) { 'Connected' } else { 'Disconnected' })"
                
                if ($connected) {
                    # Initialize audio endpoint
                    $taskIndex = 0
                    $handle = [AudioEndpoint]::AvSetMmThreadCharacteristics("Audio", [ref]$taskIndex)
                    [AudioEndpoint]::timeBeginPeriod(1)
                    
                    try {
                        # Keep audio endpoint active
                        $mmde = [IntPtr]::Zero
                        $hr = [AudioEndpoint]::CoCreateInstance(
                            [ref][AudioEndpoint]::CLSID_MMDeviceEnumerator,
                            [IntPtr]::Zero, 1,
                            [ref][AudioEndpoint]::IID_IMMDeviceEnumerator,
                            [ref]$mmde)
                        
                        if ($hr -eq 0) {
                            Write-ServiceLog "Audio endpoint initialized"
                            
                            # Method 1: Audio endpoint manipulation
                            $currentVolume = 0
                            [AudioEndpoint]::waveOutGetVolume(0, [ref]$currentVolume)
                            
                            # Minimal volume change to keep endpoint active
                            $newVolume = $currentVolume
                            [AudioEndpoint]::waveOutSetVolume(0, $newVolume)
                            Start-Sleep -Milliseconds 10
                            [AudioEndpoint]::waveOutSetVolume(0, $currentVolume)
                            
                            # Method 2: Occasional device refresh (less frequent)
                            if ((Get-Random -Minimum 1 -Maximum 30) -eq 1) {
                                Write-ServiceLog "Performing maintenance cycle..."
                                Disable-PnpDevice -InstanceId $DeviceId -Confirm:$false
                                Start-Sleep -Milliseconds 20
                                Enable-PnpDevice -InstanceId $DeviceId -Confirm:$false
                                Start-Sleep -Milliseconds 20
                            }
                        }
                        
                        # Verify final state
                        $status = (Get-PnpDevice -InstanceId $DeviceId).Status
                        Write-ServiceLog "Device status after audio init: $status"
                        
                        if ($status -eq "OK") {
                            Write-ServiceLog "Keep-alive signal successfully sent"
                            return $true
                        }
                    }
                    finally {
                        if ($handle -ne [IntPtr]::Zero) {
                            [AudioEndpoint]::AvRevertMmThreadCharacteristics($handle)
                        }
                        [AudioEndpoint]::timeEndPeriod(1)
                    }
                }
                else {
                    Write-ServiceLog "Device not connected - attempting reconnection" -Type "Warning"
                    Enable-PnpDevice -InstanceId $DeviceId -Confirm:$false
                    Start-Sleep -Seconds 1
                    return $false
                }
            }
            
        } catch {
            Write-ServiceLog "Error in device communication: $($_.Exception.Message)" -Type "Error"
            Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
            return $false
        }
    }
    catch {
        Write-ServiceLog "Critical error in keep-alive function: $($_.Exception.Message)" -Type "Error"
        Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        return $false
    }
    
    return $false
}

# Set default interval if not configured
if (-not $config.KeepAliveInterval) {
    Write-ServiceLog "No keep-alive interval configured - using default 30 seconds"
    $interval = 0.5  # 30 seconds
    
    # Store the default interval
    $config | Add-Member -NotePropertyName "KeepAliveInterval" -NotePropertyValue $interval -Force
    $config | ConvertTo-Json | Set-Content -Path $configPath
} else {
    $interval = $config.KeepAliveInterval
    Write-ServiceLog "Using configured keep-alive interval: $interval minutes"
}

# Convert interval to seconds for sleep timer
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
        }
    }
    Write-ServiceLog "Waiting $interval minutes before next keep-alive signal"
    Start-Sleep -Seconds $sleepSeconds
}