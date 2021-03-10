# Set the location to the registry
Set-Location -Path 'HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services'

# Create new items with values
New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services' -Name 'TrustedCertThumbprints' -Value "7F98F4500F5DFBD7505E00B7C730855E2365716E,4C83ED92EDD626C3AF51FFE1BA0B97E63CC6245F" -PropertyType String -Force

# Get out of the Registry
Pop-Location