$wshshell = New-Object -ComObject WScript.Shell

$lnk = $wshshell.CreateShortcut("C:\Users\Public\Desktop\CCHS-RDS01.lnk")

$lnk.TargetPath = "%windir%\system32\mstsc.exe"

$lnk.Arguments = "/v:cchs-rds01.cchs.local "

$lnk.Description = "AdminBase"

$lnk.Save()

$wshshell = New-Object -ComObject WScript.Shell

$lnk = $wshshell.CreateShortcut("C:\Users\Public\Desktop\CCHS-RDS02.lnk")

$lnk.TargetPath = "%windir%\system32\mstsc.exe"

$lnk.Arguments = "/v:cchs-rds02.cchs.local "

$lnk.Description = "AdminBase"

$lnk.Save()

$wshshell = New-Object -ComObject WScript.Shell

$lnk = $wshshell.CreateShortcut("C:\Users\Public\Desktop\CCHS-TESTRDS.lnk")

$lnk.TargetPath = "%windir%\system32\mstsc.exe"

$lnk.Arguments = "/v:cchs-testrds.cchs.local "

$lnk.Description = "AdminBase"

$lnk.Save()

$wshshell = New-Object -ComObject WScript.Shell

$lnk = $wshshell.CreateShortcut("C:\Users\Public\Desktop\CCHS-SQL01.lnk")

$lnk.TargetPath = "%windir%\system32\mstsc.exe"

$lnk.Arguments = "/v:cchs-SQL01.cchs.local "

$lnk.Description = "AdminBase"

$lnk.Save()

mkdir C:\Users\Public\Desktop\RDP
Move C:\Users\Public\Desktop\CCHS-RDS01.lnk C:\Users\Public\Desktop\RDP
Move C:\Users\Public\Desktop\CCHS-RDS02.lnk C:\Users\Public\Desktop\RDP
Move C:\Users\Public\Desktop\CCHS-SQL01.lnk C:\Users\Public\Desktop\RDP
Move C:\Users\Public\Desktop\CCHS-TESTRDS.lnk C:\Users\Public\Desktop\RDP