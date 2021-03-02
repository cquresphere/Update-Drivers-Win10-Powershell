# Load assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName PresentationFramework

#Check if script is Running as Administrator. If it is not break/stop.
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

#Create log
$date = (Get-Date -format "yyyyMMdd_hhmmss")
$compname = $env:COMPUTERNAME
$logname = $compname + "_" + $date + "_UpdateScanScript.log"
$scanlog = "c:\temp\logs\" + $logname 
new-item -path $scanlog -ItemType File -Force

Write-Host "New log file created $scanlog"

$WUAver = (Get-Item C:\Windows\System32\wuaueng.dll).VersionInfo.ProductVersion
Write-Host "Windows Update Agent version $WUAver" -ForegroundColor Yellow

$dateLog = (Get-Date -format "HH:mm:ss dd.MM.yyyy")

Add-content $scanlog -value "`n$dateLog `t `t Windows Update Agent version $WUAver `n `n "

Set-WindowsUpdateAgent -ResetToDefaults

Stop-Service wuauserv -Force

Stop-Service cryptSvc -Force
Stop-Service bits -Force
Stop-Service msiserver -Force

#Remove-Item HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate -Recurse

#Resolving problem with HRESULT: 0X8024001E
try {
    Rename-Item -Path C:\Windows\SoftwareDistribution SoftwareDistribution.old -Force
}
catch {
    Rename-Item -Path C:\Windows\SoftwareDistribution SoftwareDistribution.old1 -Force
}
finally {
    
}

try {
    Rename-Item -Path C:\Windows\System32\catroot2 catroot2.old -Force}
catch {
    Rename-Item -Path C:\Windows\System32\catroot2 catroot2.old1 -Force}
finally {
    
}
# Add ServiceID for Windows Update
Add-WUServiceManager -ServiceID 7971f918-a847-4430-9279-4a52d1efe18d -Confirm:$false 

Start-Service wuauserv 
Start-Service cryptSvc
Start-Service bits
Start-Service msiserver 

Start-Sleep 12

Do {
    $IsWUAsvc = (Get-Service -Name wuauserv).Status     
}
Until($IsWUAsvc -eq 'Running')

if($IsWUAsvc -ne 'stopped'){
    Write-Host "Windows Update Service is Running" -ForegroundColor Green
    $dateLog = (Get-Date -format "HH:mm:ss dd.MM.yyyy")
    Add-content $scanlog -value "`n$dateLog `t `t Windows Update Service started `n `n "
}

#search and list all missing Drivers
try {
    $Session = New-Object -ComObject Microsoft.Update.Session           
    $Searcher = $Session.CreateUpdateSearcher() 

    $Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
    $Searcher.SearchScope =  1 # MachineOnly
    $Searcher.ServerSelection = 3 # Third Party

    $Criteria = "IsInstalled=0 and Type='Driver' and ISHidden=0"
    $dateLog = (Get-Date -format "HH:mm:ss dd.MM.yyyy")
    Add-content $scanlog -value "$dateLog `t `tStart Searching Driver-Updates..."
    Write-Host("$dateLog `t `tStart Searching Driver-Updates...") -ForegroundColor Green  
    $StartSearchTime = (Get-Date).Second
    $SearchResult = $Searcher.Search($Criteria)          
    $Updates = $SearchResult.Updates
    $EndSearchTime = (Get-Date).Second
}
catch {
    $Error[0]
    $dateLog = (Get-Date -format "HH:mm:ss dd.MM.yyyy")
    Add-content $scanlog -value "`n$dateLog `t `t `n $Error[0] `n"
}
finally {
    
}


#Show available Drivers
$dateLog = (Get-Date -format "HH:mm:ss dd.MM.yyyy")
Write-Host "$dateLog `t `tScan took $($EndSearchTime - $StartSearchTime) seconds and here are results:" -ForegroundColor Yellow
$UpdatesList = $Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | ft | out-string
$Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | Format-List
Add-content $scanlog -value "$dateLog `n  `n $UpdatesList"

#Download the Drivers from Microsoft
try {
    $UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
    $updates | % { $UpdatesToDownload.Add($_) | out-null }
    Write-Host('Downloading Drivers...')  -ForegroundColor Green  
    $UpdateSession = New-Object -Com Microsoft.Update.Session
    $StartDownloadTime = (Get-Date).Second
    $Downloader = $UpdateSession.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToDownload
    $Downloader.Download()
    $EndDownloadTime = (Get-Date).Second
    Write-Host "$dateLog `t `tDownload took $($EndDownloadTime - $StartDownloadTime) seconds." -ForegroundColor Yellow
}
catch {
    $Error[0]
    $dateLog = (Get-Date -format "HH:mm:ss dd.MM.yyyy")
    Add-content $scanlog -value "`n$dateLog `t `t `n $Error[0] `n"
}
finally {
    
}

#Check if the Drivers are all downloaded and trigger the Installation
try {
    $UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
    $updates | % { if($_.IsDownloaded) { $UpdatesToInstall.Add($_) | out-null } }

    Write-Host('Installing Drivers...')  -ForegroundColor Green 
    $StartInstallTime = (Get-Date).Second
    $Installer = $UpdateSession.CreateUpdateInstaller()
    $Installer.Updates = $UpdatesToInstall
    $InstallationResult = $Installer.Install()
    $EndInstallTime = (Get-Date).Second
    if($InstallationResult.RebootRequired) {  
        Write-Host('Reboot required! please reboot now..') -ForegroundColor Red  
    } else { Write-Host('Done..') -ForegroundColor Green }
}
catch {
    $Error[0]
    $dateLog = (Get-Date -format "HH:mm:ss dd.MM.yyyy")
    Add-content $scanlog -value "`n$dateLog `t `t `n $Error[0] `n"
}
finally {
    
}


# Pause and give the service time to update
Start-Sleep 30

#Stop Windows Update Service
Stop-Service wuauserv

Do {
    $IsWUAsvc = (Get-Service -Name wuauserv).Status     
}
Until($IsWUAsvc -eq 'stopped')

if($IsWUAsvc -eq 'stopped'){
    Write-Host "Windows Update Service is stopped" -ForegroundColor Green
    $dateLog = (Get-Date -format "HH:mm:ss dd.MM.yyyy")
    Add-content $scanlog -value "`n$dateLog `t `t Windows Update Service stopped `n"
}
