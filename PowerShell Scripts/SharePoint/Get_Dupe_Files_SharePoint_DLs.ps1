#Parameters
$SiteURL = "https://artissl.sharepoint.com/sites/ArtisLegalDepartment"
$Pagesize = 2000
$ReportOutput = "C:\Temp\Duplicates.csv"
 
#Connect to SharePoint Online site
Connect-PnPOnline $SiteURL -Interactive
  
#Array to store results
$DataCollection = @()
 
#Get all Document libraries
$DocumentLibraries = Get-PnPList | Where-Object {$_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false -and $_.ItemCount -gt 0 -and $_.Title -Notin("Site Pages","Style Library", "Preservation Hold Library")}
 
#Iterate through each document library
ForEach($Library in $DocumentLibraries)
{   
    #Get All documents from the library
    $global:counter = 0;
    $Documents = Get-PnPListItem -List $Library -PageSize $Pagesize -Fields ID, File_x0020_Type -ScriptBlock `
        { Param($items) $global:counter += $items.Count; Write-Progress -PercentComplete ($global:Counter / ($Library.ItemCount) * 100) -Activity `
             "Getting Documents from Library '$($Library.Title)'" -Status "Getting Documents data $global:Counter of $($Library.ItemCount)";} | Where {$_.FileSystemObjectType -eq "File"}
   
    $ItemCounter = 0
    #Iterate through each document
    Foreach($Document in $Documents)
    {
        #Get the File from Item
        $File = Get-PnPProperty -ClientObject $Document -Property File
 
        #Get The File Hash
        $Bytes = $File.OpenBinaryStream()
        Invoke-PnPQuery
        $MD5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
        $HashCode = [System.BitConverter]::ToString($MD5.ComputeHash($Bytes.Value))
  
        #Collect data       
        $Data = New-Object PSObject
        $Data | Add-Member -MemberType NoteProperty -name "FileName" -value $File.Name
        $Data | Add-Member -MemberType NoteProperty -Name "HashCode" -value $HashCode
        $Data | Add-Member -MemberType NoteProperty -Name "URL" -value $File.ServerRelativeUrl
        $Data | Add-Member -MemberType NoteProperty -Name "FileSize" -value $File.Length       
        $DataCollection += $Data
        $ItemCounter++
        Write-Progress -PercentComplete ($ItemCounter / ($Library.ItemCount) * 100) -Activity "Collecting data from Documents $ItemCounter of $($Library.ItemCount) from $($Library.Title)" `
                     -Status "Reading Data from Document '$($Document['FileLeafRef']) at '$($Document['FileRef'])"
    }
}
#Get Duplicate Files by Grouping Hash code
$Duplicates = $DataCollection | Group-Object -Property HashCode | Where {$_.Count -gt 1}  | Select -ExpandProperty Group
Write-host "Duplicate Files Based on File Hashcode:"
$Duplicates | Format-table -AutoSize
 
#Export the duplicates results to CSV
$Duplicates | Export-Csv -Path $ReportOutput -NoTypeInformation


#Read more: https://www.sharepointdiary.com/2019/04/sharepoint-online-find-duplicate-files-using-powershell.html#ixzz7Zh9KRUEB