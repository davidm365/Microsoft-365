Function Add-OnedriveSecondaryAdmin($AdminURL,$SecondaryAdmin)
{
    #connect Spo service.
    Connect-SPOService -Url $https://Tenantname-admin.sharepoint.com
    #Get all Onedrive URL's.
    $OneDriveURLs = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'"
    foreach($OneDriveURL in $OneDriveURLs)
    {
        #Add Secondary administrator to Onedrive Site.
        Set-SPOUser -Site $OneDriveURL.URL -LoginName $SecondaryAdmin -IsSiteCollectionAdmin $True -ErrorAction SilentlyContinue
        Write-Host "Added secondary admin to the site $($OneDriveURL.URL)" 
    }
}
Add-OnedriveSecondaryAdmin -SecondaryAdmin "EMAIL OF SECOND ADMIN" -AdminURL "https://Tenantname-admin.sharepoint.com"
                    