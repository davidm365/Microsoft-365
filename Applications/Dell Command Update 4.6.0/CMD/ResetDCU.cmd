@Echo off
echo Configure DCU to defaults
echo Set key unconfigured
reg add HKLM\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\CFG\ /v DCUconfigured /t REG_DWORD /d 0 /f
if errorlevel 1 (
echo Error setting DCU back to unconfigured
exit /b 1
) else (
echo UnConfigured key set
)
)
if exist "%PROGRAMFILES(x86)%\Dell\CommandUpdate\dcu-cli.exe" (
 echo NONUniversal DCU
 "%PROGRAMFILES(x86)%\Dell\CommandUpdate\dcu-cli.exe" /configure -restoreDefaults
 "%PROGRAMFILES(x86)%\Dell\CommandUpdate\dcu-cli.exe" /configure -lockSettings=disable
 exit /b 0
) 
if exist "%PROGRAMFILES%\Dell\CommandUpdate\dcu-cli.exe" (
 echo Universal DCU
 "%PROGRAMFILES%\Dell\CommandUpdate\dcu-cli.exe" /configure -restoreDefaults
 "%PROGRAMFILES%\Dell\CommandUpdate\dcu-cli.exe" /configure -lockSettings=disable
 exit /b 0
) else (
 echo dcu-cli.exe doesn't exist
 exit /b 1
)