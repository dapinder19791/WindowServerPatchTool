$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            
                            $securestring  =convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass

$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))
 
 function CreatePs
{
  [cmdletbinding()]
    param ($computername,$Srcpath,$copylocally,$Type)

    $Destpath = ""
    $retvalue=$false
    if($Type -eq "Custom")
    {
        if($copylocally -eq $true)
        {
        $Destpath = " \\$computername\C$\temp"

            If(!(test-path $Destpath))
            {
                  New-Item -ItemType Directory -Force -Path $Destpath
            }
          net use \\$computername  $password /USER:$userid $powerShellArgs 2> $null

          Copy-Item  -destination $Destpath -Recurse  -Path $Srcpath -Force -ErrorAction SilentlyContinue
          $retvalue =$true
       }
   
   }
   return $retvalue
} 

function createps1file
{
Param ($Srcpath,$cpylocally,$Type)

$retvalue=$false
$path=""

$createfile=@()
if($Type -eq "Custom")
{
if($cpylocally -eq $true)
{
$filename= [System.IO.Directory]::GetFiles("c:\temp\", "*.exe")
$path = $filename[0]  +  " /q /action=patch /allinstances /IAcceptSQLServerLicenseTerms "
write-host $path -ForegroundColor Green

}
else{
$path = $Srcpath  +  " /q /action=patch /allinstances /IAcceptSQLServerLicenseTerms "
write-host $path -ForegroundColor Green

}

$createfile +=
@'
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
'@

$createfile += $path + "           " 

$createfile += @'
    
   
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

}
'@

$retvalue=$true

}

if($Type -eq "StayCurrent")
{
 
$createfile +=
@'
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
'@

$createfile += ' \\ITDSL\MSNPLAT\GOLD\MSNPATCH\MSNPATCH.exe /CORP /NOVALIDATION:WindowsClusterCheck  /OnlyQFE:4019090,4019091,4032542,4019092,4019093,4036996,4019086,4019088  '


$createfile += @'
    
   
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

}
'@

$retvalue=$true
}

$createfile | Out-File -FilePath .\Scripts\SQLExecuteFile.ps1

return $retvalue
}

Function enableWinRM {
Param([string]$computerName)

    $checkresult=$false
      net use \\$computerName  $password /USER:$userid $powerShellArgs 2> $null

	$result = winrm id -r:$computerName 2>$null

	 
	if ($LastExitCode -eq 0) {
		Write-Host "WinRM already enabled on" $computerName "..." -ForegroundColor green
	} else {
		Write-Host "Enabling WinRM on" $computerName "..." -ForegroundColor Blue
		 C:\psscripts\PsExec.exe  \\$computerName -s C:\Windows\system32\winrm.cmd qc -quiet

		if ($LastExitCode -eq 0) {
            Write-Host "Resetting window service - WINRM on" $computerName "..." -ForegroundColor Yellow

			 C:\psscripts\PsExec.exe \\$computerName restart WinRM
			$result = winrm id -r:$computerName 2>$null
			
			if ($LastExitCode -eq 0) {
                $checkresult=$true
            Write-Host 'WinRM successfully enabled!' -ForegroundColor green}
			 
		}
        else{
        Write-Host "Unable to reset window service - WINRM on" $computerName "..." -ForegroundColor Red
        } 
		 
	}

return $checkresult
}

Function Copy-Packages {
Param ($computername,$Srcpath,$cpylocally,$Type)

    
   
    
    $flag=$false
    $retryAttempt=$false
    $message=""

    try{
     write-host "[Information] : Attempting to connect to Server - $computername with userid - [$userid]" -ForegroundColor White
     Enter-PSSession –ComputerName $computername -Credential $cred  -ErrorAction Stop
     $flag=$true
     write-host "[Information]: Connection is Established with Server - $computername with userid - [$userid]" -ForegroundColor Green
    }
    catch{
     write-host "[Error]: Error Connecting Server - $computername with userid - [$userid]" -ForegroundColor Red
     write-host "[Information]: Retry Attempt connecting to Server - $computername with userid - [$userid]" -ForegroundColor Yellow
     
     $checkresults = enableWinRM -computerName $computername
     if($checkresults){
     $retryAttempt=$true
     $flag=$true
     write-host "[Information]: Retry Attempt connecting to Server - $computername with userid - [$userid] is successfull" -ForegroundColor Green
     }
     else{$retryAttempt=$false
     $flag=$false
     write-host "[Error]: Retry Attempt Failed connecting to Server - $computername with userid - [$userid] is successfull" -ForegroundColor Red
     }

    try{
      

               if($flag)
               {
                Remove-PSSession -ComputerName $computername -ErrorAction SilentlyContinue
                
                $getstatus=CreatePs -computername $computername -Srcpath $Srcpath -copylocally  $cpylocally -Type $Type
                $createscript=createps1file -Srcpath $Srcpath -cpylocally $cpylocally -Type $Type

                net use \\$computername  $password /USER:$userid $powerShellArgs 2> $null
                if($Type -eq "Custom")
                { 
                Copy-Item  -destination \\$computername\c$ -Recurse  -Path .\Scripts\Runupdates_SQL.bat -Force -ErrorAction SilentlyContinue
                Copy-Item -destination \\$computername\c$ -Recurse -Path  .\Scripts\SQLExecuteFile.ps1 -Force -ErrorAction SilentlyContinue
                Invoke-Command -ComputerName $computername  -ScriptBlock { Remove-ItemProperty -Path 'HKLM:System\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue} -ErrorAction SilentlyContinue 
                }
                else{
                Copy-Item  -destination \\$computername\c$ -Recurse  -Path .\Scripts\Runupdates_SQL.bat -Force -ErrorAction SilentlyContinue
                Copy-Item -destination \\$computername\c$ -Recurse -Path  .\Scripts\SQLExecuteFile.ps1 -Force -ErrorAction SilentlyContinue
                Invoke-Command -ComputerName $computername  -ScriptBlock { Remove-ItemProperty -Path 'HKLM:System\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue} -ErrorAction SilentlyContinue 
                
                }
               
  
                $message="SUCCESS"
               }
              else{
                $message="UnAuthorized"
                 write-host "[Information]: Your Userid - [$userid] is UnAuthorized to connect to Server - $computername " -ForegroundColor Yellow  
              }
              
               
        }
        catch [Exception]
        {
                        $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                        $report.Computer = $Computername
                        $report.Title = "ERROR:$_.Exception.Message"
                        $report.KB = "ERROR"
                        $report.IsDownloaded = "ERROR"
                        $report.LastRebootTime = "ERROR"
                        $report.LastPatchInstalledTime = "ERROR"
                        $report.Notes = "ERROR:$_.Exception.Message" 
                        $report.OSVersion = "ERROR:$_.Exception.Message" 
                        $message="ERROR:$_.Exception.Message"    
                         write-host "[Error]: Error Occured in Server - $computername , [Error Message :] " $_.Exception.Message -ForegroundColor Red     
        
        }
        
        return $message   
   
    
  
}


Function Format-InstallPatchLog {
    [cmdletbinding()]
    param ($computername)
     
      $installreport = @()
    #Check for logfile
    If (Test-Path "\\$computername\c$\$($computername)_SQLPatchLog.csv") {
        #Retrieve the logfile from remote server
        
         
        $file = Import-Csv "\\$computername\c$\$($computername)_SQLPatchLog.csv" 
        $installreport=$file
             
        }
    Else {

    write-host "[Error]: File Path Error Occured in Server - $computername " -ForegroundColor Red  
       $temp = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
        
        $temp.Computer = $computername
        $temp.Title = "ERROR"
        $temp.KB = "File Path Error \\$computername\c$\$($computername)_SQLPatchLog.csv"
        $temp.IsDownloaded = "File Path Error \\$computername\c$\$($computername)_SQLPatchLog.csv"
        $temp.LastRebootTime = "File Path Error \\$computername\c$\$($computername)_SQLPatchLog.csv"
        $temp.LastPatchInstalledTime = "File Path Error \\$computername\c$\$($computername)_SQLPatchLog.csv"
        $temp.Notes = "File Path Error \\$computername\c$\$($computername)_SQLPatchLog.csv"
        $temp.OSVersion  = "File Path Error \\$computername\c$\$($computername)_SQLPatchLog.csv"
        $Updatereport = $temp      
        }

    Write-Output $installreport

  
    
}

function SQL-Patch
{
     
    Param($Computername,$Srcpath,$cpylocally,$Type)
   
    

if(Test-Connection -ComputerName $computername -Count 1 -Quiet)
   {
    
    try{
        
    If (Test-Path C:\psscripts\psexec.exe) {
        
        $returnmessage = Copy-Packages -computer $Computername -Srcpath $Srcpath -cpylocally $cpylocally -Type $Type
       

        if($returnmessage -eq "SUCCESS")
        { 
          C:\psscripts\PsExec.exe -s \\$computername   C:\Runupdates_SQL.bat /quiet /norestart /accepteula -h $powerShellArgs 2> $null  


         $path1 ="\\$computername\c$\Runupdates_SQL.bat"
                $path2 = "\\$computername\c$\SQLpatch.ps1"

                if([System.IO.File]::Exists($path1)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path1   $powerShellArgs 2> $null  
                        }
 
                if([System.IO.File]::Exists($path2)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path2   $powerShellArgs 2> $null  
                        }
               
                    If ($LASTEXITCODE -eq 0) {
                        Format-InstallPatchLog -computer $Computername
                        }            
                    Else {
                        #$host.ui.WriteLine("Unsuccessful run of install script!")
                         $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                        $report.Computer = $Computername
                        $report.Title = "ERROR"
                        $report.KB = "ERROR"
                        $report.IsDownloaded = "ERROR"
                        $report.LastRebootTime = "ERROR"
                        $report.LastPatchInstalledTime = "ERROR"
                         
                        $report.Notes = "ERROR:Unsuccessful PsExec.exe run of install script" 
                        $report.OSVersion = "ERROR:Unsuccessful PsExec.exe run of install script" 
                        Write-Output $report                 
                        }
                    }
                     elseif ($returnmessage -eq "UnAuthorized")
                   {
                   $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                    $report.Computer = $Computername
                    $report.Title = "UnAuthorizedAccess"
                    $report.KB = "UnAuthorizedAccess"
                    $report.IsDownloaded = "UnAuthorizedAccess"
                    $report.LastRebootTime = "UnAuthorizedAccess"
                    $report.LastPatchInstalledTime = "UnAuthorizedAccess"
                    $report.Notes = "UnAuthorizedAccess" 
                    $report.OSVersion ="UnAuthorizedAccess" 
                    Write-Output $report          

                   
                   }
                   else 
                   {
                    $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                        $report.Computer = $Computername
                        $report.Title = "$returnmessage"
                        $report.KB = "ERROR"
                        $report.IsDownloaded = "ERROR"
                        $report.LastRebootTime = "ERROR"
                        $report.LastPatchInstalledTime = "ERROR"
                         
                        $report.Notes = "$returnmessage" 
                        $report.OSVersion = "ERROR" 
                        Write-Output $report   
                   
                   }
                }      
            Else {

                        $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                        $report.Computer = $Computername
                        $report.Title = "PSExec not found ! please download it first."
                        $report.KB = "PSExec not found ! please download it first."
                        $report.IsDownloaded = "PSExec not found ! please download it first."
                        $report.LastRebootTime = "PSExec not found ! please download it first."
                        $report.LastPatchInstalledTime = "PSExec not found ! please download it first."
                        $report.Notes = "PSExec not found ! please download it first." 
                        $report.OSVersion = "PSExec not found ! please download it first." 
                        Write-Output $report

             
                }
     }
     catch{
             $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
            $report.Computer = $Computername
            $report.Title = "UnAuthorizedAccess"
            $report.KB = "UnAuthorizedAccess"
            $report.IsDownloaded = "UnAuthorizedAccess"
            $report.LastRebootTime = "UnAuthorizedAccess"
            $report.LastPatchInstalledTime = "UnAuthorizedAccess"
            $report.Notes = "UnAuthorizedAccess" 
            $report.OSVersion ="UnAuthorizedAccess" 
            Write-Output $report  

       }
   }
   else{
           $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
            $report.Computer = $Computername
            $report.Title = "Offline"
            $report.KB = "Offline"
            $report.IsDownloaded = "Offline"
            $report.LastRebootTime = "Offline"
            $report.LastPatchInstalledTime = "Offline"
            $report.Notes = "Offline" 
            $report.OSVersion ="Offline"
            Write-Output $report          

   } 
    
    
     
      
} 
 