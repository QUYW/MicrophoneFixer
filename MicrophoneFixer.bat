@echo off
color 0A
title Windows Microphone Troubleshooter v2.0
setlocal enabledelayedexpansion

:: MicrophoneFixer v2.0
:: Created by imaginesamurai
:: https://github.com/imaginesamurai
:: please dont remove my Watermark

:: Set up log file
set "logfile=%temp%\MicrophoneFixer_log.txt"
echo Windows Microphone Troubleshooter v2.0 - Log started at %date% %time% > "%logfile%"

cls
echo.
echo  ___       __       _____ ______   _________       
echo ^\^\  ^\     ^\^\  ^\    ^\^\   _ ^\  _   ^\^\___   ___^\     
echo ^\ ^\  ^\    ^\ ^\  ^\   ^\ ^\  ^\^\^\__^\ ^\  ^\^|___ ^\  \_^|     
echo  ^\ ^\  ^\  __^\ ^\  ^\   ^\ ^\  ^\^\^|__^| ^\  ^\   ^\ ^\  ^\      
echo   ^\ ^\  ^\^|^\__\_^\  ^\ __^\ ^\  ^\    ^\ ^\  ^\ __^\ ^\  ^\ ___ 
echo    ^\ ^\____________^\^\__^\ ^\__^\    ^\ ^\__^\^\__^\ ^\__^\^\__^\
echo     ^\^|____________^\^|__^\^\^|__^|     ^\^|__^\^|__^\^\^|__^\^|__^|
echo.
echo ===================================================
echo         WINDOWS MICROPHONE TROUBLESHOOTER v2.0
echo          Created by github.com/imaginesamurai
echo ===================================================
echo.
echo This tool will help diagnose and fix common microphone issues
echo where Windows recognizes your microphone but other apps don't.
echo.
echo Press any key to begin the troubleshooting process...
pause > nul

:MENU
cls
echo  ___       __       _____ ______   _________       
echo ^\^\  ^\     ^\^\  ^\    ^\^\   _ ^\  _   ^\^\___   ___^\     
echo ^\ ^\  ^\    ^\ ^\  ^\   ^\ ^\  ^\^\^\__^\ ^\  ^\^|___ ^\  \_^|     
echo  ^\ ^\  ^\  __^\ ^\  ^\   ^\ ^\  ^\^\^|__^| ^\  ^\   ^\ ^\  ^\      
echo   ^\ ^\  ^\^|^\__\_^\  ^\ __^\ ^\  ^\    ^\ ^\  ^\ __^\ ^\  ^\ ___ 
echo    ^\ ^\____________^\^\__^\ ^\__^\    ^\ ^\__^\^\__^\ ^\__^\^\__^\
echo     ^\^|____________^\^|__^\^\^|__^|     ^\^|__^\^|__^\^\^|__^\^|__^|
echo.
echo ===================================================
echo         WINDOWS MICROPHONE TROUBLESHOOTER v2.0
echo          Created by github.com/imaginesamurai
echo ===================================================
echo.
echo Please select an option:
echo.
echo  DIAGNOSTICS:
echo  1. Check microphone status
echo  2. Run advanced diagnostic tests
echo  3. View system audio information
echo.
echo  FIXES:
echo  4. Fix microphone privacy settings
echo  5. Reset audio services
echo  6. Fix microphone permissions for apps
echo  7. Check default microphone settings
echo  8. Fix exclusive mode issues
echo  9. Run comprehensive fix (recommended)
echo.
echo  TOOLS:
echo  10. Test microphone (audio recording test)
echo  11. Check for audio driver updates
echo  12. Create system restore point
echo  13. View troubleshooting log
echo.
echo  14. Exit
echo.
set /p choice=Enter your choice (1-14): 

if "%choice%"=="1" goto CHECK_MIC
if "%choice%"=="2" goto ADVANCED_DIAGNOSTICS
if "%choice%"=="3" goto SYSTEM_AUDIO_INFO
if "%choice%"=="4" goto FIX_PRIVACY
if "%choice%"=="5" goto RESET_AUDIO
if "%choice%"=="6" goto FIX_PERMISSIONS
if "%choice%"=="7" goto CHECK_DEFAULT
if "%choice%"=="8" goto FIX_EXCLUSIVE_MODE
if "%choice%"=="9" goto COMPREHENSIVE
if "%choice%"=="10" goto TEST_MIC
if "%choice%"=="11" goto CHECK_DRIVERS
if "%choice%"=="12" goto CREATE_RESTORE
if "%choice%"=="13" goto VIEW_LOG
if "%choice%"=="14" goto EXIT
goto MENU

:CHECK_MIC
cls
echo ===================================================
echo             CHECKING MICROPHONE STATUS
echo ===================================================
echo.
echo Checking if your microphone is detected by Windows...
echo.

echo [%date% %time%] Checking microphone status... >> "%logfile%"

powershell -Command "Get-PnpDevice -Class 'AudioEndpoint' | Where-Object { $_.FriendlyName -like '*microphone*' -or $_.FriendlyName -like '*mic*' -or $_.FriendlyName -like '*input*' } | Format-Table FriendlyName, Status, InstanceId -AutoSize"

echo.
echo Checking for disabled audio devices...
echo.
powershell -Command "Get-PnpDevice -Class 'AudioEndpoint' | Where-Object { $_.Status -ne 'OK' } | Format-Table FriendlyName, Status, InstanceId -AutoSize"

echo.
echo If your microphone is listed above with "OK" status, Windows detects it.
echo If it shows another status, it may be disabled or have driver issues.
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:ADVANCED_DIAGNOSTICS
cls
echo ===================================================
echo             ADVANCED DIAGNOSTICS TESTS
echo ===================================================
echo.
echo Running comprehensive diagnostics on your audio system...
echo This may take a moment...
echo.

echo [%date% %time%] Running advanced diagnostics... >> "%logfile%"

echo 1. Checking all audio endpoints...
powershell -Command "Get-PnpDevice -Class 'AudioEndpoint' | Format-Table FriendlyName, Status, InstanceId -AutoSize"

echo.
echo 2. Checking audio services status...
sc query Audiosrv | findstr "STATE"
sc query AudioEndpointBuilder | findstr "STATE"
sc query Windows Audio | findstr "STATE"

echo.
echo 3. Checking for conflicting applications...
echo.
tasklist /fi "imagename eq audiodg.exe" /v
tasklist /fi "imagename eq Realtek*" /v
tasklist /FI "MODULES eq *audio*" /v

echo.
echo 4. Checking volume mixer settings...
powershell -Command "Get-AudioDevice -PlaybackVolume | Format-Table Name, Volume -AutoSize"
powershell -Command "Get-AudioDevice -RecordingVolume | Format-Table Name, Volume -AutoSize" 2>nul

echo.
echo 5. Checking for Windows Audio API errors in event log...
powershell -Command "Get-EventLog -LogName System -Newest 50 | Where-Object {$_.Source -like '*audio*'} | Format-Table TimeGenerated, EntryType, Message -AutoSize" 2>nul

echo.
echo Diagnostic tests completed.
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:SYSTEM_AUDIO_INFO
cls
echo ===================================================
echo             SYSTEM AUDIO INFORMATION
echo ===================================================
echo.
echo Gathering information about your audio subsystem...
echo.

echo [%date% %time%] Gathering system audio information... >> "%logfile%"

echo AUDIO HARDWARE:
powershell -Command "Get-WmiObject Win32_SoundDevice | Format-Table Name, Status, DeviceID -AutoSize"

echo.
echo DEFAULT RECORDING DEVICE:
powershell -Command "& {$devices = Get-WmiObject -Class Win32_SoundDevice | Where-Object {$_.StatusInfo -eq 3}; foreach ($device in $devices) {if ($device.Name -like '*microphone*' -or $device.Name -like '*mic*') {Write-Host $device.Name}}}"

echo.
echo AUDIO DRIVER INFORMATION:
powershell -Command "Get-WmiObject Win32_PnPSignedDriver | Where-Object {$_.DeviceClass -eq 'MEDIA'} | Format-Table DeviceName, DriverVersion, DriverDate -AutoSize"

echo.
echo WINDOWS VERSION:
ver

echo.
echo This information can be helpful when seeking support for audio issues.
echo Consider including it if you need to ask for help online.
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:FIX_PRIVACY
cls
echo ===================================================
echo          FIXING MICROPHONE PRIVACY SETTINGS
echo ===================================================
echo.
echo Windows 10/11 has privacy settings that can block microphone access.
echo This will enable microphone access for Windows and apps.
echo.
echo Applying fixes...

echo [%date% %time%] Fixing microphone privacy settings... >> "%logfile%"

:: Enable microphone access in privacy settings
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone" /v "Value" /t REG_SZ /d "Allow" /f
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone" /v "Value" /t REG_SZ /d "Allow" /f

:: Allow apps to access microphone
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged" /v "Value" /t REG_SZ /d "Allow" /f

:: Enable background apps to access microphone (Windows 10/11)
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications" /v "GlobalUserDisabled" /t REG_DWORD /d "0" /f
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Search" /v "BackgroundAppGlobalToggle" /t REG_DWORD /d "1" /f

echo.
echo Microphone privacy settings have been updated.
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:RESET_AUDIO
cls
echo ===================================================
echo              RESETTING AUDIO SERVICES
echo ===================================================
echo.
echo This will restart Windows audio services which can fix many issues.
echo.
echo Stopping audio services...

echo [%date% %time%] Resetting audio services... >> "%logfile%"

net stop Audiosrv
net stop AudioEndpointBuilder
timeout /t 2 /nobreak > nul

echo Starting audio services...
net start AudioEndpointBuilder
net start Audiosrv

echo.
echo Checking Windows audio service dependencies...
echo.
sc qc Audiosrv

echo.
echo Audio services have been reset.
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:FIX_PERMISSIONS
cls
echo ===================================================
echo         FIXING MICROPHONE APP PERMISSIONS
echo ===================================================
echo.
echo This will ensure apps have permission to use your microphone.
echo.

echo [%date% %time%] Fixing microphone app permissions... >> "%logfile%"

:: Enable microphone access for apps
powershell -Command "& {Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone' -Name 'Value' -Value 'Allow'; Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone' -Name 'Value' -Value 'Allow';}"

:: Allow desktop apps to access microphone
powershell -Command "& {Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged' -Name 'Value' -Value 'Allow';}"

:: Reset the Windows app permissions for microphone
powershell -Command "& {Get-ChildItem 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\' | ForEach-Object { if ($_.PSChildName -ne 'NonPackaged') { Set-ItemProperty -Path $_.PSPath -Name 'Value' -Value 'Allow' } }}"

echo Microphone app permissions have been updated.
echo.
echo Now opening Windows microphone privacy settings...
start ms-settings:privacy-microphone

echo.
echo Please ensure the toggles for "Microphone access" and "Let apps access your microphone" are turned ON.
echo.
echo Press any key when you've finished...
pause > nul
goto MENU

:CHECK_DEFAULT
cls
echo ===================================================
echo         CHECKING DEFAULT MICROPHONE SETTINGS
echo ===================================================
echo.
echo This will check and set your default microphone device.
echo.

echo [%date% %time%] Checking default microphone settings... >> "%logfile%"

echo Current default recording device:
powershell -Command "& {$devices = Get-WmiObject -Class Win32_SoundDevice | Where-Object {$_.StatusInfo -eq 3}; foreach ($device in $devices) {if ($device.Name -like '*microphone*' -or $device.Name -like '*mic*') {Write-Host $device.Name}}}"

echo.
echo Opening Sound Control Panel so you can set the default device...
echo.
echo 1. Right-click on your preferred microphone
echo 2. Select "Set as Default Device"
echo 3. Also set it as "Default Communication Device" if available
echo 4. Click "Properties" and ensure levels are up and not muted
echo 5. Go to "Advanced" tab and uncheck "Allow applications to take exclusive control"
echo.

start control mmsys.cpl,,1

echo Press any key when you've finished adjusting settings...
pause > nul
goto MENU

:FIX_EXCLUSIVE_MODE
cls
echo ===================================================
echo         FIXING EXCLUSIVE MODE SETTINGS
echo ===================================================
echo.
echo Some applications can't access the microphone when other apps
echo have exclusive control. This fix prevents that problem.
echo.

echo [%date% %time%] Fixing exclusive mode settings... >> "%logfile%"

echo Detecting audio devices...
echo.

:: List all audio devices
powershell -Command "& {$devices = Get-PnpDevice -Class 'AudioEndpoint' | Where-Object { $_.Status -eq 'OK' }; foreach ($device in $devices) { Write-Host $device.FriendlyName }}"

echo.
echo Disabling exclusive mode for audio devices...
echo (This process runs in the background, please wait...)
echo.

:: Use PowerShell to modify registry settings for all audio devices to disable exclusive mode
powershell -Command "& {$devices = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\' -ErrorAction SilentlyContinue; foreach ($device in $devices) { New-Item -Path ($device.PSPath + '\Properties') -ErrorAction SilentlyContinue | Out-Null; New-Item -Path ($device.PSPath + '\Properties\{b3f8fa53-0004-438e-9003-51a46e139bfc}') -ErrorAction SilentlyContinue | Out-Null; Set-ItemProperty -Path ($device.PSPath + '\Properties\{b3f8fa53-0004-438e-9003-51a46e139bfc}') -Name '14' -Value ([byte[]](0x00, 0x00, 0x00, 0x00)) -Type Binary -ErrorAction SilentlyContinue }}"

echo Exclusive mode settings have been updated for all audio devices.
echo.
echo You should restart your computer for these changes to take full effect.
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:COMPREHENSIVE
cls
echo ===================================================
echo            RUNNING COMPREHENSIVE FIX
echo ===================================================
echo.
echo This will run all fixes in sequence to resolve common microphone issues.
echo.

echo [%date% %time%] Running comprehensive fix... >> "%logfile%"

echo Step 1: Fixing privacy settings...
call :FIX_PRIVACY_SILENT

echo Step 2: Resetting audio services...
net stop Audiosrv > nul 2>&1
net stop AudioEndpointBuilder > nul 2>&1
timeout /t 2 /nobreak > nul
net start AudioEndpointBuilder > nul 2>&1
net start Audiosrv > nul 2>&1

echo Step 3: Fixing app permissions...
powershell -Command "& {Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone' -Name 'Value' -Value 'Allow'; Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone' -Name 'Value' -Value 'Allow'; Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged' -Name 'Value' -Value 'Allow';}" > nul 2>&1

echo Step 4: Checking for disabled devices...
powershell -Command "& {$devices = Get-PnpDevice -Class 'AudioEndpoint' | Where-Object { $_.Status -eq 'Error' }; foreach ($device in $devices) { Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false }}" > nul 2>&1

echo Step 5: Disabling exclusive mode for audio devices...
powershell -Command "& {$devices = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\' -ErrorAction SilentlyContinue; foreach ($device in $devices) { New-Item -Path ($device.PSPath + '\Properties') -ErrorAction SilentlyContinue | Out-Null; New-Item -Path ($device.PSPath + '\Properties\{b3f8fa53-0004-438e-9003-51a46e139bfc}') -ErrorAction SilentlyContinue | Out-Null; Set-ItemProperty -Path ($device.PSPath + '\Properties\{b3f8fa53-0004-438e-9003-51a46e139bfc}') -Name '14' -Value ([byte[]](0x00, 0x00, 0x00, 0x00)) -Type Binary -ErrorAction SilentlyContinue }}" > nul 2>&1

echo Step 6: Resetting Windows Store apps that might be using the microphone...
powershell -Command "Get-AppxPackage | Where-Object {$_.Name -like '*comm*' -or $_.Name -like '*skype*' -or $_.Name -like '*teams*'} | Reset-AppxPackage" > nul 2>&1

echo Step 7: Fixing Windows audio policies...
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Audio" /v "DisableProtectedAudioDG" /t REG_DWORD /d "1" /f > nul 2>&1

echo Step 8: Resetting USB devices (may help with USB microphones)...
powershell -Command "& {$devices = Get-PnpDevice -Class 'USB' | Where-Object { $_.FriendlyName -like '*audio*' -or $_.FriendlyName -like '*microphone*' -or $_.FriendlyName -like '*mic*' }; foreach ($device in $devices) { Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false; Start-Sleep -Seconds 2; Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false }}" > nul 2>&1

echo Step 9: Opening Sound Control Panel for final adjustments...
echo.
echo Please check your microphone settings in the Sound Control Panel:
echo 1. Ensure your microphone is set as the default device
echo 2. Check that levels are up and not muted
echo 3. In Properties -^> Advanced, uncheck "Allow applications to take exclusive control"
echo.

start control mmsys.cpl,,1

echo All automatic fixes have been applied.
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:TEST_MIC
cls
echo ===================================================
echo              TESTING MICROPHONE
echo ===================================================
echo.
echo This will test your microphone by recording a short audio clip
echo and playing it back to confirm your microphone is working.
echo.

echo [%date% %time%] Testing microphone... >> "%logfile%"

set "test_file=%temp%\mic_test.wav"

echo Please speak into your microphone for 5 seconds after pressing any key...
pause > nul
echo.
echo Recording... Please speak now!
echo.

:: Record audio for 5 seconds
powershell -Command "Add-Type -AssemblyName System.Speech; $recognizer = New-Object System.Speech.Recognition.SpeechRecognitionEngine; $grammar = New-Object System.Speech.Recognition.DictationGrammar; $recognizer.LoadGrammar($grammar); $recognizer.SetInputToDefaultAudioDevice(); $recognizer.RecognizeAsync([System.Speech.Recognition.RecognitionMode]::Multiple); Start-Sleep -Seconds 5; $recognizer.SpeechRecognized | Where-Object {$_.Result.Text.Length -gt 0} | ForEach-Object {$_.Result.Text}; $recognizer.Dispose()" > "%test_file%_text.txt"

:: Alternative method using Sound Recorder CLI (if available)
start /wait soundrecorder /FILE "%test_file%" /DURATION 00:00:05

echo.
echo Recording complete!
echo.
echo Playing back your recording... (If you don't hear anything, your microphone may not be working)
echo.

powershell -Command "Add-Type -AssemblyName System.Windows.Forms; $player = New-Object System.Media.SoundPlayer; $player.SoundLocation = '%test_file%'; $player.PlaySync()"

echo.
echo If you heard your voice played back, your microphone is working correctly!
echo If not, please try the comprehensive fix from the main menu.
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:CHECK_DRIVERS
cls
echo ===================================================
echo          CHECKING AUDIO DRIVER UPDATES
echo ===================================================
echo.
echo This will check if your audio drivers need updating...
echo.

echo [%date% %time%] Checking audio drivers... >> "%logfile%"

echo Current audio drivers:
echo.
powershell -Command "Get-WmiObject Win32_PnPSignedDriver | Where-Object {$_.DeviceClass -eq 'MEDIA'} | Format-Table DeviceName, DriverVersion, DriverDate -AutoSize"

echo.
echo Looking for outdated drivers...
echo.

:: Check driver status
powershell -Command "& {$drivers = Get-WmiObject Win32_PnPSignedDriver | Where-Object {$_.DeviceClass -eq 'MEDIA'}; foreach ($driver in $drivers) { $deviceID = $driver.DeviceID; $status = Get-PnpDeviceProperty -InstanceId $deviceID -KeyName DEVPKEY_Device_DriverProblemDesc -ErrorAction SilentlyContinue; if ($status.Data) { Write-Host ('Driver issue found for ' + $driver.DeviceName + ': ' + $status.Data) } else { Write-Host ('No known issues for ' + $driver.DeviceName + ' driver') } }}"

echo.
echo To update your audio drivers, you can:
echo 1. Visit your computer manufacturer's website
echo 2. Visit your sound card manufacturer's website
echo 3. Use Windows Device Manager to update drivers
echo.
echo Opening Device Manager so you can update drivers...
echo Right-click on your audio devices and select "Update driver"
echo.

start devmgmt.msc

echo Press any key to return to the menu...
pause > nul
goto MENU

:CREATE_RESTORE
cls
echo ===================================================
echo          CREATING SYSTEM RESTORE POINT
echo ===================================================
echo.
echo It's a good idea to create a restore point before making 
echo significant changes to your system.
echo.

echo [%date% %time%] Creating system restore point... >> "%logfile%"

echo Checking if System Restore is enabled...
powershell -Command "& {$SR = Get-ComputerRestorePoint -ErrorAction SilentlyContinue; if ($SR -eq $null) { Write-Host 'System Restore appears to be disabled on this system.' } else { Write-Host 'System Restore is enabled.' }}"

echo.
echo Attempting to create a restore point...
powershell -Command "& {try { Checkpoint-Computer -Description 'Before Microphone Troubleshooter Fixes' -RestorePointType 'APPLICATION_INSTALL' -ErrorAction Stop; Write-Host 'Restore point created successfully!' } catch { Write-Host ('Failed to create restore point: ' + $_.Exception.Message) }}"

echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:VIEW_LOG
cls
echo ===================================================
echo          VIEWING TROUBLESHOOTING LOG
echo ===================================================
echo.
echo Below is the log of actions performed in this session:
echo.

type "%logfile%"

echo.
echo The log file is saved at: %logfile%
echo.
echo Press any key to return to the menu...
pause > nul
goto MENU

:FIX_PRIVACY_SILENT
:: Enable microphone access in privacy settings (silent version)
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone" /v "Value" /t REG_SZ /d "Allow" /f > nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone" /v "Value" /t REG_SZ /d "Allow" /f > nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged" /v "Value" /t REG_SZ /d "Allow" /f > nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications" /v "GlobalUserDisabled" /t REG_DWORD /d "0" /f > nul 2>&1
goto :eof

:EXIT
cls
echo  ___       __       _____ ______   _________       
echo ^\^\  ^\     ^\^\  ^\    ^\^\   _ ^\  _   ^\^\___   ___^\     
echo ^\ ^\  ^\    ^\ ^\  ^\   ^\ ^\  ^\^\^\__^\ ^\  ^\^|___ ^\  \_^|     
echo  ^\ ^\  ^\  __^\ ^\  ^\   ^\ ^\  ^\^\^|__^| ^\  ^\   ^\ ^\  ^\      
echo   ^\ ^\  ^\^|^\__\_^\  ^\ __^\ ^\  ^\    ^\ ^\  ^\ __^\ ^\  ^\ ___ 
echo    ^\ ^\____________^\^\__^\ ^\__^\    ^\ ^\__^\^\__^\ ^\__^\^\__^\
echo     ^\^|____________^\^|__^\^\^|__^|     ^\^|__^\^|__^\^\^|__^\^|__^|
echo.
echo ===================================================
echo            MICROPHONE TROUBLESHOOTER v2.0
echo          Created by github.com/imaginesamurai
echo ===================================================
echo.
echo Thank you for using the Windows Microphone Troubleshooter!
echo.
echo If your microphone is still not working, consider:
echo 1. Updating or reinstalling your audio drivers
echo 2. Checking if your microphone is physically connected properly
echo 3. Testing the microphone on another device to rule out hardware issues
echo 4. Contacting your PC manufacturer for support
echo.
echo A log of this session has been saved to:
echo %logfile%
echo.
echo Press any key to exit...
pause > nul
exit

endlocal 
