# Windows Microphone Troubleshooter v2.0

A comprehensive Windows batch script to diagnose and fix common microphone issues, particularly when Windows recognizes the microphone but other applications don't.

<p align="center">
  <img src="banner.png" alt="MICTFIX Banner" style="width:250px;">
</p>

## Features

### Diagnostics
- Check microphone status
- Run advanced diagnostic tests
- View system audio information

### Fixes
- Fix microphone privacy settings
- Reset audio services
- Fix microphone permissions for apps
- Check default microphone settings
- Fix exclusive mode issues
- Run comprehensive fix (recommended)

### Tools
- Test microphone (audio recording test)
- Check for audio driver updates
- Create system restore point
- View troubleshooting log

## Installation

1. Download `MicrophoneFixer.bat`
2. Right-click the file and select "Run as administrator"

## Usage

Run the script and select from the available options in the menu. For most cases, the "Run comprehensive fix" option (9) is recommended as it applies all fixes in sequence.

### Example Code

Here's a snippet of the batch script:

```batch
@echo off
color 0A
title Windows Microphone Troubleshooter v2.0
setlocal enabledelayedexpansion

:: MicrophoneFixer v2.0
:: Created by imaginesamurai
:: https://github.com/imaginesamurai

:: Set up log file
set "logfile=%temp%\MicrophoneFixer_log.txt"
```

## What This Tool Fixes

- Windows privacy settings blocking microphone access
- Incorrect default device settings
- Audio service issues
- App permission problems
- Exclusive mode control issues
- USB microphone detection problems
- Windows Store app microphone access
- Background app microphone access
- Audio driver issues

## Logging

The tool creates a detailed log file of all operations performed during the troubleshooting session. The log file is saved at:
`%temp%\MicrophoneFixer_log.txt`

## System Requirements

- Windows 10 or Windows 11
- Administrator privileges
- PowerShell 5.1 or later

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

Created by [imaginesamurai](https://github.com/imaginesamurai)

## Acknowledgments

- Thanks to all the users who provided feedback and helped improve this tool
- Special thanks to the Windows PowerShell community for their documentation

## Support

If you're having issues:

1. Check the troubleshooting log file
2. Try the comprehensive fix option
3. Create an issue on GitHub with your log file
4. Make sure your microphone is properly connected
5. Update your audio drivers

## Disclaimer

This tool makes changes to your Windows registry and system settings. While it's designed to be safe, it's recommended to create a system restore point (option 12) before making changes to your system.



## Changelog

### v2.0
- Added comprehensive diagnostic tests
- Added microphone testing capability
- Added system restore point creation
- Added troubleshooting log
- Enhanced privacy settings fixes
- Added USB device reset functionality
- Improved error handling
- Added exclusive mode fixes

### v1.0 (i've deleted this version by mistake and i dont have it anymore)
- Initial release
- Basic troubleshooting functionality
- Privacy settings fixes
- Audio service reset
- App permissions management 
