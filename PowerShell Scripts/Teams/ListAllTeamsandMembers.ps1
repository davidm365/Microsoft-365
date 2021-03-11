<#
 
.SYNOPSIS 
Lists all Microsoft Teams by name, team members and team channels.
  
.DESCRIPTION 
Copy this code into an editor or your choosing 
(recommended: PowerShell ISE or VS Code)
  
#>
 
# Check if Teams Module is installed. If not, it will be installed
 
 
If (-not (Get-Module -ListAvailable | Where-Object Name -eq 'MicrosoftTeams' ))
 
{
 
  Install-Module -Name MicrosoftTeams -Force -AllowClobber
 
}
 
# Connect to Microsoft Teams
 
Connect-MicrosoftTeams
 
# List all Teams and all Channels
 
$ErrorAction = "SilentlyContinue"
 
$allteams = Get-Team
$object = @()
 
foreach ($t in $allteams) {
 
    $members = Get-TeamUser -GroupId $t.GroupId
 
    $owner = Get-TeamUser -GroupId $t.GroupId -Role Owner
 
    $channels = Get-TeamChannel -GroupId $t.GroupId 
 
    $object += New-Object -TypeName PSObject -Property ([ordered]@{
 
        'Team'= $t.DisplayName
        'GroupId' = $t.GroupId
        'Owner' = $owner.User
        'Members' = $members.user -join "`r`n"
        'Channels' = $channels.displayname -join "`r`n"
     
        })
         
}
Write-Output $object