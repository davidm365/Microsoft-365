function Connect-ExchangeOnline {
 
# .SYNOPSIS
# Connect-ExchangeOnline connects you to your Office365 Exchange Online portal.
 
# .PARAMETER
# None
 
# .EXAMPLE
# Connect-ExchangeOnline
 
# .NOTES
# Author: Patrick Gruenauer
# Web: https://sid-500.com
 
$userCredential = Get-Credential
$session = New-PSSession -ConfigurationName Microsoft.Exchange `
-ConnectionUri https://outlook.office365.com/powershell-liveid/ `
-Credential $userCredential -Authentication Basic -AllowRedirection
Import-Module (Import-PSSession $session -DisableNameChecking) -Global `
-WarningAction SilentlyContinue
$domain=Get-AcceptedDomain | Where-Object Default -EQ 'True'
""
Write-Output "***** Welcome to Exchange Online for the domain $domain *****"
""
}