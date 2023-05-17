$report =@()
 
[System.Collections.ArrayList]$SingleReport = @()


function check-PendingReboot
{
 if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { return $true }
 if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { return $true }
 
 try { 
   $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
   $status = $util.DetermineIfRebootPending()
   if(($status -ne $null) -and $status.RebootPending){
     return $true
   }
 }catch{}
 
  return $false
}



function get-lastinstalledDate
{
 $Retvalue=$false;
$Session = New-Object -ComObject Microsoft.Update.Session            
$Searcher = $Session.CreateUpdateSearcher()            
$HistoryCount = $Searcher.GetTotalHistoryCount()            
$history =  $Searcher.QueryHistory(0,$HistoryCount) | ForEach-Object -Process {            
    $Title = $null            
            
        $Title = $_.Title            
             
     $Result = $null            
    Switch ($_.ResultCode)            
    {            
        0 { $Result = 'NotStarted'}            
        1 { $Result = 'InProgress' }            
        2 { $Result = 'Succeeded' }            
        3 { $Result = 'SucceededWithErrors' }            
        4 { $Result = 'Failed' }            
        5 { $Result = 'Aborted' }            
        default { $Result = $_ }            
    }            
    New-Object -TypeName PSObject -Property @{            
        InstalledOn = Get-Date -Date $_.Date;            
        Title = $Title;            
        Name = $_.Title;            
        Status = $Result            
    }            
            
} | Sort-Object -Descending:$true -Property InstalledOn |             

Select-Object -Property * -ExcludeProperty Name  
 

$history | Select-Object -First 1| Sort-Object InstallOn –Descending

 
}
function set_objects
{
Param ($Computer,$Title,$KB,$IsDownloaded,$Notes,$OSVersion,$LastScanTime,$LastRebootTime,$LastPatchInstalledTime)
$ListObject = [PSCustomObject]@{
            Computer       = $Computer
            Title        = $Title
            KB                  = $KB
            IsDownloaded        = $IsDownloaded
            Notes              = $Notes
            OSVersion          = $OSVersion
            LastScanTime        = $LastScanTime
            LastRebootTime               = $LastRebootTime
            LastPatchInstalledTime = $LastPatchInstalledTime
               
        }
return $ListObject 

}


try{
$Computer = $env:computername

$sOS=(Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Name
$sOS= $sOS.Substring(0,$sOS.IndexOf("|"))
$LASTREBOOT = Get-CimInstance -ComputerName $Computer -ClassName win32_operatingsystem | select csname, lastbootuptime
$lastboouptime=$LASTREBOOT.lastbootuptime
$date=Get-Date
$lastpatchinstalled=get-lastinstalledDate

#foreach ($Computer in $Servers){

if (Test-Connection -ComputerName $Computer -Count 1 -Quiet){
## Log file path
$outfilePath =  "c:\$($Computer)_patchlog.csv"

if([System.IO.File]::Exists($outfilePath)){ 

  Remove-Item $outfilePath -force -recurse

 }
  

if($Computer -eq "$env:computername")
				{
					$UpdateSession = New-Object -ComObject Microsoft.Update.Session
				}
else { $UpdateSession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$Computer)) }

Write-Host("Searching for Any pending reboots...")   (Get-Date) -Fore Green
$checkreboot = check-PendingReboot 

if($checkreboot)
{
 $report=set_objects -Computer $Computer -Title "RebootPending" -KB "NA" -IsDownloaded "NA"  -Notes "Installed and Reboot Pending" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
 $SingleReport.Add($report) | Out-Null

}
else{

$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

 
Write-Host("Searching for applicable updates...")   (Get-Date) -Fore Green
$SearchResult = $UpdateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0")
$cnt = $SearchResult.Updates.Count

Write-Host "There are " $SearchResult.Updates.Count ": TOTAL updates available." 
Write-Host("List of applicable items on the machine:") -Fore Green

For ($X = 0; $X -lt $SearchResult.Updates.Count; $X++){
    $Update = $SearchResult.Updates.Item($X)
    Write-Host( ($X + 1).ToString() + "> " + $Update.Title)
}

 
If ($SearchResult.Updates.Count -eq 0) {
    $report=set_objects -Computer $Computer -Title "NoPendingUpdates" -KB "NA" -IsDownloaded "NA"   -Notes "There are no pending updates on this Machine." -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
    $SingleReport.Add($report) | Out-Null
                         
}
if($SearchResult.Updates.Count -gt 0)
{ 
Write-Host("")
Write-Host("Creating collection of updates to download:") -Fore Green

if($Computer -eq "$env:computername")
				{
					$UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
				}
else { $UpdatesToDownload = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.UpdateColl",$Computer)) }
 

 
For ($X = 0; $X -lt $SearchResult.Updates.Count; $X++){
    $Update = $SearchResult.Updates.Item($X)
    Write-Host( ($X + 1).ToString() + "&gt; Adding: " + $Update.Title)
    $Null = $UpdatesToDownload.Add($Update)
}
 

Write-Host("Downloading Updates...")  -ForegroundColor Green
 
$Downloader = $UpdateSession.CreateUpdateDownloader()
Write-Host("Downloading Updates.1..")  -ForegroundColor Green

$Downloader.Updates = $UpdatesToDownload

Write-Host("Downloading Updates.2..")  -ForegroundColor Green


$Null = $Downloader.Download()
 
 Write-Host("Downloading Updates.3..")  -ForegroundColor Green
 
Write-Host("List of Downloaded Updates...") -ForegroundColor Green

if($Computer -eq "$env:computername")
				{
					$UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
				}
else { $UpdatesToInstall = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.UpdateColl",$Computer)) }


 
For ($X = 0; $X -lt $SearchResult.Updates.Count; $X++){
    $Update = $SearchResult.Updates.Item($X)
    If ($Update.IsDownloaded) {
        Write-Host( ($X + 1).ToString() + "&gt; " + $Update.Title)
        $Null = $UpdatesToInstall.Add($Update)        
    }
}
 
$Install = [System.String]$Args[0]
$Reboot  = [System.String]$Args[1]
 
If (!$Install){
    $Install = "Y"
}
 
If ($Install.ToUpper() -eq "Y" -or $Install.ToUpper() -eq "YES"){
    Write-Host("")
    Write-Host("Installing Updates...")  "-" (Get-Date) -Fore Green
 
    $Installer = $UpdateSession.CreateUpdateInstaller()
    $Installer.Updates = $UpdatesToInstall
 
    $InstallationResult = $Installer.Install()
 
    Write-Host("")
    write-host "List of Updates Installed with Results:" (Get-Date) -ForegroundColor Green
  
    For ($X = 0; $X -lt $UpdatesToInstall.Count; $X++){

        $resultcode += $UpdatesToInstall.Item($X).Title
        #$InstallationResult.GetUpdateResult($X).ResultCode

    
    }
    
     $report=set_objects -Computer $Computer -Title "InstallationCompleted" -KB "NA" -IsDownloaded "NA"   -Notes "Installation is completed for - $resultcode " -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
     $SingleReport.Add($report) | Out-Null

     
    If ($InstallationResult.RebootRequire -eq $True){
        If (!$Reboot){
            $Reboot = "N"
        }
 
        If ($Reboot.ToUpper() -eq "Y" -or $Reboot.ToUpper() -eq "YES"){
            Write-Host("")
            Write-Host("Rebooting...") -Fore Green
            (Get-WMIObject -Class Win32_OperatingSystem).Reboot()
        }
      }
    }
  }
 }
}

}catch [Exception]{
          
           $sOS=(Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Name
            $sOS= $sOS.Substring(0,$sOS.IndexOf("|"))
            $LASTREBOOT = Get-CimInstance -ComputerName $Computer -ClassName win32_operatingsystem | select csname, lastbootuptime
            $lastboouptime=$LASTREBOOT.lastbootuptime
            
          
            $report=set_objects -Computer $Computer -Title "ERROR:$_.Exception.Message" -KB "NA" -IsDownloaded "NA"   -Notes "ERROR:$_.Exception.Message" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $null 
            
            $SingleReport.Add($report) | Out-Null
}

 if($SingleReport.Count -gt 0)
 {
 $SingleReport|Export-CSV $outfilePath -noTypeInformation -Append
 }