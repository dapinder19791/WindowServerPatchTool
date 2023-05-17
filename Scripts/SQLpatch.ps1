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

function checkSqlServerExists
{

Param ($computername)

$val=$false

$SQLInstances = Invoke-Command -ComputerName $computername -ErrorAction Ignore{
Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Microsoft SQL Server'
} 

$counter=@($SQLInstances).Count

if ($counter -gt 0) {
$val= $true
 } Else {
 $val= $false
 }
  

 return $val
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

$Computer = $env:computername

$sOS=(Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Name
$sOS= $sOS.Substring(0,$sOS.IndexOf("|"))
$LASTREBOOT = Get-CimInstance -ComputerName $Computer -ClassName win32_operatingsystem | select csname, lastbootuptime
$lastboouptime=$LASTREBOOT.lastbootuptime
$date=Get-Date
$lastpatchinstalled=get-lastinstalledDate

   
 
if (Test-Connection -ComputerName $Computer -Count 1 -Quiet){
## Log file path
$outfilePath =  "\\$Computer\c$\$($Computer)_SQLPatchLog.csv"

if([System.IO.File]::Exists($outfilePath)){ 

  Remove-Item $outfilePath -force -recurse

 }
 
$checkreboot = check-PendingReboot 
$status=checkSqlServerExists -computername $Computer
if($status -eq $false)
{
$report=set_objects -Computer $Computer -Title "SQLDoesnotExists" -KB "NA" -IsDownloaded "NA"  -Notes "SQL Server Does not Exists in this server." -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
 $SingleReport.Add($report) | Out-Null
}
else{
if($checkreboot)
{
 $report=set_objects -Computer $Computer -Title "RebootPending" -KB "NA" -IsDownloaded "NA"  -Notes "This Server is pending Reboot" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
 $SingleReport.Add($report) | Out-Null

}
else{
    
    \\itdsl\msnplat\gold\msnpatch\msnpatch.exe /CORP /NOVALIDATION:WindowsClusterCheck  /OnlyQFE:3194725,3194724,3194722,3194718,3194717,3194719,3194721,3194720,3194714,3194716 | Out-File $outfilePath -Append
 
 $report=set_objects -Computer $Computer -Title "Installed" -KB "NA" -IsDownloaded "NA"  -Notes "Installed SQL Patches" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
 $SingleReport.Add($report) | Out-Null

      }
   }
}
else{
$report=set_objects -Computer $Computer -Title "Offline" -KB "NA" -IsDownloaded "NA"  -Notes "Server is Offline." -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
 $SingleReport.Add($report) | Out-Null
}

   
 if($SingleReport.Count -gt 0)
 {
 $SingleReport|Export-CSV $outfilePath -noTypeInformation -Append
 }