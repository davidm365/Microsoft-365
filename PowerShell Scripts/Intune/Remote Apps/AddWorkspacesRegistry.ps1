# Set the location to the registry
Set-Location -Path 'HKCU:\Software\Policies\Microsoft'

# Create a new Key
Get-Item -Path 'HKCU:\Software\Policies\Microsoft' | New-Item -Name 'Workspaces' -Force

# Create new items with values
New-ItemProperty -Path 'HKCU:\Software\Policies\Microsoft\Workspaces\' -Name 'DefaultConnectionURL' -Value "https://cchs-rds01.cchs.local/rdweb/feed/webfeed.aspx" -PropertyType String

Set-ItemProperty -Path 'HKCU:\Software\Policies\Microsoft\Workspaces\' -Name '(Default)' -Value 0

# Get out of the Registry
Pop-Location