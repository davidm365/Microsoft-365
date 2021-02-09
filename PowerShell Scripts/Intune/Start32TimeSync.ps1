Set-Service -Name W32Time -StartupType Automatic
net start w32time
w32tm /resync /force