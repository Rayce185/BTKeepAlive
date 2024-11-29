**Version 0.7.2**

A Windows service designed to prevent Bluetooth audio devices from entering standby mode when no audio is playing.

**Overview**

Many Bluetooth speakers and audio devices automatically enter standby mode after a period of inactivity to conserve power. While this is generally beneficial, it can be inconvenient when you want the device to remain readily available. This service maintains an active connection by sending periodic keep-alive signals to your selected Bluetooth devices.

**Features**

* 🔌 Keeps Bluetooth audio devices active and connected
* 🔄 Automatic service installation and configuration
* 📊 Real-time connection status monitoring
* 📝 Detailed logging of device status and operations
* 🛠️ Easy device selection through interactive setup
* ⚙️ Configurable keep-alive intervals
* 🔍 Device status verification before each keep-alive signal

**Requirements**

* Windows 10/11
* PowerShell 5.1 or later
* Administrator privileges
* Bluetooth-capable system
* Pre-paired Bluetooth audio devices

**Installation**

1. Clone or download this repository
2. Run PowerShell as Administrator
3. Navigate to the downloaded directory
4. Execute BTKeepAlive.ps1

```txt
.\BTKeepAlive.ps1
```

### Setup Process

1. The script will display a list of available Bluetooth audio devices
2. Select your device(s) by entering the corresponding number(s)
3. Confirm your selection
4. The service will be installed and started automatically

**Configuration**

The service creates these files in C:\BTService:

* BTConfig.json - Device and interval configuration
* service.log - Operational logs
* error.log - Error and warning messages

**Default Settings**

* Keep-alive interval: 2 minutes
* Log rotation: Enabled (1MB limit)
* Service account: LocalSystem
* Startup type: Automatic

**How It Works**

The service:

1. Monitors the connection status of configured devices
2. Sends keep-alive signals at configured intervals
3. Verifies device power state and connectivity
4. Logs all operations and any issues encountered
5. Automatically recovers from connection interruptions

### Advanced Configuration

You can modify BTConfig.json to adjust settings:

```txt
{
    "devices": [
        {
            "id": "DEVICE_ID",
            "name": "DEVICE_NAME"
        }
    ],
    "KeepAliveInterval": 5
}
```

### Known Limitations

* Currently optimized for audio devices
* May not work with all Bluetooth device types
* Some devices might require specific keep-alive intervals
* Service requires system restart after major Windows updates

**Troubleshooting**

Check the following logs for issues:

* C:\BTService\service.log - General operation logs
* C:\BTService\error.log - Error details
* Windows Event Viewer > Application > BTKeepAlive source

**Work in Progress**

This project is under active development. Current focus areas:

* ⏳ Automatic standby timeout detection
* 🔄 Dynamic interval adjustment
* 🔧 Device-specific optimization
* 🛡️ Enhanced error recovery
* ⚡ Power consumption optimization

**Disclaimer**

This software is provided "as is", without warranty of any kind. While efforts have been made to ensure reliability, the authors are not responsible for any device issues or unintended consequences. Always ensure your devices are compatible and test thoroughly before extended use.

**Version History**

* 0.7.2 - Current development version
  * Enhanced logging
  * Improved error handling
  * Basic keep-alive functionality
* 0.7.1 - Initial public release
  * Basic service functionality
  * Device selection
  * Service installation
