if not exist "C:\ProgramData\AutoPilotConfig" md "C:\ProgramData\AutoPilotConfig"
if not exist "C:\ProgramData\AutoPilotConfig\Icons" md "C:\ProgramData\AutoPilotConfig\Icons"
md c:\Users\Public\Desktop\Shortcuts
xcopy "Shortcuts.ps1" "C:\ProgramData\AutoPilotConfig" /Y
xcopy "ADP.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "Athena.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "Dentrix.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "healthstream.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "incidentresponse.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "lippincott.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "outlook.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "Qminder.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "Workflow.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
xcopy "Printers.ico" "C:\ProgramData\AutoPilotConfig\Icons" /Y
Powershell.exe -Executionpolicy bypass -File "C:\ProgramData\AutoPilotConfig\Shortcuts.ps1"