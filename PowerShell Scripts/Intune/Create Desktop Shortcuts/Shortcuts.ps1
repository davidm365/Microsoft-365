if (-not (Test-Path "C:\Users\Public\Desktop\Shortcuts\ADP.url"))
{
$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\ADP.url"
$targetpath = "https://clock.adp.com"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\ADP.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\AthenaWorkflows.url"
$targetpath = "\\cchs-prtfs01\athena workflows$"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\Workflow.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\AthenaNet.url"
$targetpath = "https://athenanet.athenahealth.com/1/2/login.esp"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\Athena.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\Dentrix Share.url"
$targetpath = "\\cchs-sql01\cchs shared\all dental users"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\Dentrix.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\HealthStream.url"
$targetpath = "https://www.healthstream.com/hlc/choptank"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\healthstream.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\Incident Response.url"
$targetpath = "\\cchs-prtfs01\incident response guide - red books"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\incidentresponse.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\Lippincott.url"
$targetpath = "https://procedures.lww.com"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\lippincott.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\Outlook Web Mail.url"
$targetpath = "https://office.com"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\outlook.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\Qminder.url"
$targetpath = "https://dashboard.qminder.com"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\qminder.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

$null = $WshShell = New-Object -comObject WScript.Shell
$path = "C:\Users\Public\Desktop\Shortcuts\CCHS Printers.url"
$targetpath = "\\cchs-prtfs01"
$iconlocation = "C:\ProgramData\AutoPilotConfig\Icons\Printers.ico"
$iconfile = "IconFile=" + $iconlocation
$Shortcut = $WshShell.CreateShortcut($path)
$Shortcut.TargetPath = $targetpath
$Shortcut.Save()

Add-Content $path "HotKey=0"
Add-Content $path "$iconfile"
Add-Content $path "IconIndex=0"
}