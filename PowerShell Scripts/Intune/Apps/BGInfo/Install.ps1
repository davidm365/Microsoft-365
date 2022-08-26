New-item -itemtype directory -force -path “c:\Program Files\BGInfo”

Copy-item -path “$psscriptroot\bginfo64.exe” -destination “C:\Program Files\BGInfo\bginfo64.exe”

Copy-item -path “$psscriptroot\custom.bgi ” -destination “C:\Program Files\BGInfo\custom.bgi”

Copy-item -path “$psscriptroot\bginfo.lnk” -destination “C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\bginfo.lnk”

Start-process “C:\Program Files\BGInfo\Bginfo64.exe” -ArgumentList “`”C:\Program Files\BgInfo\custom.bgi`””,”/timer:0″,”/silent”,”/nolicprompt”

Return 0