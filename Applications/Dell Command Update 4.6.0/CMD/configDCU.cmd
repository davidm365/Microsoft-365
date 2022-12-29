@Echo off
echo Configure DCU after install
echo key to prevent setuppopup
reg add HKLM\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\CFG\ /v ShowSetupPopup /t REG_DWORD /d 0 /f
if errorlevel 1 (
echo Error installing reg key for ShowSetupPopup
exit /b 1
) else (
echo ShowSetupPopup key set
)
echo Set key configured
reg add HKLM\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\CFG\ /v DCUconfigured /t REG_DWORD /d 1 /f
if errorlevel 1 (
echo Error installing reg key for configured
exit /b 1
) else (
echo Configured key set
)
if exist "%PROGRAMFILES(x86)%\Dell\CommandUpdate\dcu-cli.exe" (
 echo NON Universal DCU
 "%PROGRAMFILES(x86)%\Dell\CommandUpdate\dcu-cli.exe" /configure -importSettings=MySettings.xml
 "%PROGRAMFILES(x86)%\Dell\CommandUpdate\dcu-cli.exe" /configure -lockSettings=enable
 exit /b 0
) 
if exist "%PROGRAMFILES%\Dell\CommandUpdate\dcu-cli.exe" (
 echo Universal DCU
 "%PROGRAMFILES%\Dell\CommandUpdate\dcu-cli.exe" /configure -importSettings=MySettings.xml
 "%PROGRAMFILES%\Dell\CommandUpdate\dcu-cli.exe" /configure -lockSettings=enable
 exit /b 0
) else (
 echo dcu-cli.exe doesn't exist
 exit /b 1
)