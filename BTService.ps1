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

    [DllImport("Kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr hObject);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr SetupDiGetClassDevs(
        ref Guid ClassGuid,
        IntPtr Enumerator,
        IntPtr hwndParent,
        uint Flags);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern bool SetupDiEnumDeviceInterfaces(
        IntPtr DeviceInfoSet,
        IntPtr DeviceInfoData,
        ref Guid InterfaceClassGuid,
        uint MemberIndex,
        ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern bool SetupDiGetDeviceInterfaceDetail(
        IntPtr DeviceInfoSet,
        ref SP_DEVICE_INTERFACE_DATA DeviceInterfaceData,
        IntPtr DeviceInterfaceDetailData,
        uint DeviceInterfaceDetailDataSize,
        ref uint RequiredSize,
        IntPtr DeviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true)]
    public static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

    [StructLayout(LayoutKind.Sequential)]
    public struct SP_DEVICE_INTERFACE_DATA
    {
        public uint cbSize;
        public Guid InterfaceClassGuid;
        public uint Flags;
        public IntPtr Reserved;
    }

    public const uint DIGCF_PRESENT = 0x2;
    public const uint DIGCF_DEVICEINTERFACE = 0x10;
    public const uint GENERIC_READ = 0x80000000;
    public const uint GENERIC_WRITE = 0x40000000;
    public const uint FILE_SHARE_READ = 0x00000001;
    public const uint FILE_SHARE_WRITE = 0x00000002;
    public const uint OPEN_EXISTING = 3;
    public const uint FILE_ATTRIBUTE_NORMAL = 0x80;
    public const uint ERROR_INSUFFICIENT_BUFFER = 122;
}
"@

Add-Type -TypeDefinition $source

# Add this function after the DLL imports and before the main loop
function Get-BluetoothDevicePath {
    param (
        [string]$DeviceId
    )
    
    try {
        Write-ServiceLog "Starting device path lookup for ID: $DeviceId"
        
        # Convert device ID to path format using Global prefix
        $devicePath = $DeviceId.Replace('\', '#')
        $fullPath = "\\.\Global\$devicePath"
        Write-ServiceLog "Created device path: $fullPath"
        
        return $fullPath
        
    } catch {
        Write-ServiceLog "Error getting device path: $($_.Exception.Message)" -Type "Error"
        Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        return $null
    }
}

# Initialize device tracking
$deviceStatus = @{}
foreach ($device in $config.devices) {
    $deviceStatus[$device.id] = @{
        lastStatus = $false
        lastKeepAlive = [DateTime]::MinValue
    }
}

Write-ServiceLog "Service started, monitoring devices:`n$($config.devices | ForEach-Object { $_.name } | Out-String)"

while ($true) {
    foreach ($device in $config.devices) {
        try {
            $currentDevice = Get-PnpDevice -InstanceId $device.id -ErrorAction Stop
            Write-ServiceLog "Device status for $($device.name): $($currentDevice.Status)"
            
            if ($currentDevice.Status -eq "OK") {
                $devicePath = Get-BluetoothDevicePath $device.id
				if ($null -eq $devicePath) {
					Write-ServiceLog "Could not get device interface path for $($device.name)" -Type "Warning"
					continue
				}
				Write-ServiceLog "Attempting to open device at path: $devicePath"
                $handle = [BluetoothHelper]::CreateFile(
                    $devicePath,
                    [BluetoothHelper]::GENERIC_READ -bor [BluetoothHelper]::GENERIC_WRITE,
                    [BluetoothHelper]::FILE_SHARE_READ -bor [BluetoothHelper]::FILE_SHARE_WRITE,
                    [IntPtr]::Zero,
                    [BluetoothHelper]::OPEN_EXISTING,
                    [BluetoothHelper]::FILE_ATTRIBUTE_NORMAL,
                    [IntPtr]::Zero
                )
                
                if ($handle -ne -1) {
                    $IoControlCode = 0x41248C
                    Write-ServiceLog "Attempting DeviceIoControl with code: 0x$($IoControlCode.ToString('X'))"
                    $bytes = 0
                    $result = [BluetoothHelper]::DeviceIoControl(
                        $handle,
                        $IoControlCode,
                        [IntPtr]::Zero,
                        0,
                        [IntPtr]::Zero,
                        0,
                        [ref]$bytes,
                        [IntPtr]::Zero
                    )
                    Write-ServiceLog "DeviceIoControl result for $($device.name): $result"
                    
                    if ($result) {
                        Write-ServiceLog "Keep-alive signal successfully sent to $($device.name)"
                    } else {
                        $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                        Write-ServiceLog "Failed to send keep-alive signal to $($device.name). Error code: $errorCode" -Type "Warning"
                    }
                    
                    [BluetoothHelper]::CloseHandle($handle)
                } else {
                    $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    Write-ServiceLog "Failed to open device handle for $($device.name). Error code: $errorCode" -Type "Warning"
                }
            } else {
                Write-ServiceLog "Device $($device.name) is not in OK state. Current status: $($currentDevice.Status)" -Type "Warning"
            }
        } catch {
            Write-ServiceLog "Error processing device $($device.name): $($_ | Out-String)" -Type "Error"
            Write-ServiceLog "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        }
    }
    Start-Sleep -Seconds 15
}