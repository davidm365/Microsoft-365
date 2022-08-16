#Set Parameters
$AdminSiteURL="https://squirrelhillhealthcenter-admin.sharepoint.com/"
$SiteCollAdmin="dtheilman@squirrelhillhealthcenter.onmicrosoft.com"
  
#Connect to PnP Online to the Tenant Admin Site
Connect-PnPOnline -Url $AdminSiteURL -Interactive
  
#Get All OneDrive Sites
$OneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'"
 
#Loop through each site
ForEach($Site in $OneDriveSites)
{
    #Add Site collection Admin
    Set-PnPTenantSite -Url $Site.URL -Owners $SiteCollAdmin
    Write-Host -f Green "Added Site Collection Admin to: "$Site.URL
}