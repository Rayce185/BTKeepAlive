# BTKeepAlive
## Bluetooth & Audio Keep-Alive Service
### Version 0.8.0

A Windows service designed to prevent Bluetooth audio devices from entering standby mode and maintain active audio endpoints to eliminate connection gaps.

## Overview
Many Bluetooth speakers and audio devices automatically enter standby mode after a period of inactivity to conserve power. This can cause a noticeable delay (~1 second) when audio playback resumes. This service maintains both the Bluetooth connection and audio system activity to ensure instant audio playback.

## Features
- 🔌 Keeps Bluetooth audio devices active and connected
- 🎵 Maintains active audio endpoints to prevent standby
- 🔄 Automatic service installation and configuration
- 📊 Real-time connection status monitoring
- 📝 Detailed logging of device status and operations
- 🛠️ Easy device selection through interactive setup
- ⚙️ Configurable keep-alive intervals
- 🔍 Device status verification before each keep-alive signal
- 🎧 Active audio stack management

## Requirements
- Windows 10/11
- PowerShell 5.1 or later
- Administrator privileges
- Bluetooth-capable system
- Pre-paired Bluetooth audio devices

## Installation
1. Clone or download this repository
2. Run PowerShell as Administrator
3. Navigate to the downloaded directory
4. Execute BTKeepAlive.ps1

```txt
.\BTKeepAlive.ps1
```

## Uninstallation
1. Navigate to the installation directory
2. Run BTUninstallService.ps1 as Administrator

```txt
.\BTUninstallService.ps1

```

This will:
- Stop the service if running
- Remove the service from Windows
- Clean up service files
- Remove event log sources

## How It Works
The service uses a dual-approach strategy to maintain device availability:

1. **Bluetooth Connection Management**
   - Monitors device connection status
   - Sends keep-alive signals to prevent disconnection
   - Automatically recovers from connection interruptions

2. **Audio System Management**
   - Maintains active audio endpoints
   - Prevents audio stack from entering power-saving mode
   - Ensures immediate audio playback availability

## Configuration
The service creates these files in C:\BTService:
- BTConfig.json - Device configuration (automatically managed by install script)
- service.log - Operational logs
- error.log - Error and warning messages

### Default Settings
- Keep-alive interval: 30 seconds (automatically configured)
- Log rotation: Enabled (1MB limit)
- Service account: LocalSystem
- Startup type: Automatic

### Configuration Process
1. Run BTKeepAlive.ps1 as administrator
2. Select your Bluetooth devices from the displayed list
3. The script will automatically:
   - Create/update the configuration file
   - Configure optimal settings for selected devices
   - Install/update the service

**How It Works**

The service:
1. Monitors the connection status of configured devices
2. Sends keep-alive signals at configured intervals
3. Verifies device power state and connectivity
4. Logs all operations and any issues encountered
5. Automatically recovers from connection interruptions

### Known Limitations

* Optimized for audio devices
* May not work with all Bluetooth device types
* Some devices might require specific keep-alive intervals
* Service requires system restart after major Windows updates
* Power consumption may be slightly higher due to prevented standby
* ⚠️ **Current Status**: The service reduces but does not completely eliminate the standby gap when audio playback starts. This is being actively worked on.

**Troubleshooting**

Check the following logs for issues:
* C:\BTService\service.log - General operation logs
* C:\BTService\error.log - Error details
* Windows Event Viewer > Application > BTKeepAlive source

**Work in Progress**

This project is under active development. Current focus areas:
* 🚧 **Priority**: Eliminating the audio playback gap
* 🎵 Optimizing audio system integration
* 🔋 Fine-tuning power management settings
* ⚡ Performance improvements for different device types
* 🛡️ Enhanced error recovery and resilience
* 🔄 Automatic device reconnection improvements
* 🎚️ Advanced audio endpoint management
* 📊 Enhanced connection status monitoring

**Disclaimer**

This software is provided "as is", without warranty of any kind. While efforts have been made to ensure reliability, the authors are not responsible for any device issues or unintended consequences. Always ensure your devices are compatible and test thoroughly before extended use.

**Version History**

* 0.8.0 - Current development version
  * Added audio system keep-alive functionality
  * Implemented dual-approach connection management
  * Enhanced power management prevention
  * Reduced keep-alive intervals for better responsiveness
  * Improved logging and diagnostics
* 0.7.2 - Previous version
  * Enhanced logging
  * Improved error handling
  * Basic keep-alive functionality
* 0.7.1 - Initial public release
  * Basic service functionality
  * Device selection
  * Service installation
