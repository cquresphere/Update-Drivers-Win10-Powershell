# Load assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName PresentationFramework

#Check if script is running as Administrator. If it is not break/stop.
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
 {    
    Write-Host "This script needs to be run As Admin!!!" -ForegroundColor Red
    $oReturn=[System.Windows.Forms.Messagebox]::Show("This script needs to be run As Admin!!!", "Error")
  
  #Echo "This script needs to be run As Admin!!!"
    Break
 }

#set Window Title 
$host.ui.RawUI.WindowTitle = "Driver Updater by harrymc updated by cquresphere" 


#Install and Import Windows Powershell Update Module
#$cred = Get-Credential
Install-Module PSWindowsUpdate -Force #-credentials $cred

Import-Module PSWindowsUpdate 


# Add ServiceID for Windows Update
Add-WUServiceManager -ServiceID 7971f918-a847-4430-9279-4a52d1efe18d -Confirm:$false 

Start-Service wuauserv 
 
# Pause and give the service time to update
Start-Sleep 30

#search and list all missing Drivers

$Session = New-Object -ComObject Microsoft.Update.Session           
$Searcher = $Session.CreateUpdateSearcher() 

$Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
$Searcher.SearchScope =  1 # MachineOnly
$Searcher.ServerSelection = 3 # Third Party

$Criteria = "IsInstalled=0 and Type='Driver' and ISHidden=0"
Write-Host('Searching Driver-Updates...') -ForegroundColor Green  
$SearchResult = $Searcher.Search($Criteria)          
$Updates = $SearchResult.Updates

#Show available Drivers

$Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | Format-List

#Download the Drivers from Microsoft

$UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
$updates | % { $UpdatesToDownload.Add($_) | out-null }
Write-Host('Downloading Drivers...')  -ForegroundColor Green  
$UpdateSession = New-Object -Com Microsoft.Update.Session
$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$Downloader.Download()

#Check if the Drivers are all downloaded and trigger the Installation

$UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
$updates | % { if($_.IsDownloaded) { $UpdatesToInstall.Add($_) | out-null } }

Write-Host('Installing Drivers...')  -ForegroundColor Green  
$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToInstall
$InstallationResult = $Installer.Install()
if($InstallationResult.RebootRequired) {  
    Write-Host('Reboot required! please reboot now..') -ForegroundColor Red  
} else { Write-Host('Done..') -ForegroundColor Green }

# Pause and give the service time to update
Start-Sleep 30

#Stop Windows Update Service
Stop-Service wuauserv