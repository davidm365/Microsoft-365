# Silently install Adobe Reader DC with Microsoft Intune
# In order to distribute Adobe Acrobat Reader DC software you need to have 
# a valid Adobe Acrobat Reader DC Distribution Agreement in place.
# See http://www.adobe.com/products/acrobat/distribute.html?readstep for details.

# Check if Software is installed already in registry.
$CheckADCReg = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | where {$_.DisplayName -like "Adobe Acrobat Reader DC*"}
# If Adobe Reader is not installed continue with script. If it's istalled already script will exit.
If ($CheckADCReg -eq $null) {

# Path for the temporary downloadfolder. Script will run as system so no issues here
$Installdir = "c:\temp\install_adobe"
New-Item -Path $Installdir  -ItemType directory

# Download the installer from the Adobe website. Always check for new versions!!
$source = "ftp://ftp.adobe.com/pub/adobe/reader/win/AcrobatDC/2001320064/AcroRdrDC2001320064_en_US.exe"
$destination = "$Installdir\AcroRdrDC2001320064_en_US.exe"
Invoke-WebRequest $source -OutFile $destination

# Start the installation when download is finished
Start-Process -FilePath "$Installdir\AcroRdrDC2001320064_en_US.exe" -ArgumentList "/sAll /rs /rps /msi /norestart /quiet EULA_ACCEPT=YES"

# Wait for the installation to finish. Test the installation and time it yourself. I've set it to 240 seconds.
Start-Sleep -s 240

# Finish by cleaning up the download. I choose to leave c:\temp\ for future installations.
rm -Force $Installdir\AcroRdrDC*
}