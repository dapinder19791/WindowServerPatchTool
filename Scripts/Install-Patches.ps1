$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            $securestring  =convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass

$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))
  
  

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


Function Create-UpdateVBS {
Param ($computername)

    
   
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
     

    }

    try{
      

               if($flag)
               {
                Remove-PSSession -ComputerName $computername -ErrorAction SilentlyContinue
              
                net use \\$computername  $password /USER:$userid $powerShellArgs 2> $null
                Copy-Item  -destination \\$computername\c$ -Recurse  -Path .\scripts\InstallWindowsUpdates.bat -Force -ErrorAction SilentlyContinue
                Copy-Item -destination \\$computername\c$ -Recurse -Path   .\scripts\InstallWindowsUpdates.ps1 -Force -ErrorAction SilentlyContinue
               
  

                $message="SUCCESS"
               }
              else{
                $message="UnAuthorized"
                 write-host "[Information]: Your Userid - [$userid] is UnAuthorized to connect to Server - $computername " -ForegroundColor Yellow  
              }
              
               
        }
        catch [Exception]
        {
                    write-host "[Error]: Error Occured in Server - $computername , [Error Message :] " $_.Exception.Message -ForegroundColor Red 

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
        
        }
        
        return $message   
   
    
  
}

function CreatePs
{
  [cmdletbinding()]
    param ($computername)

   $batfilename="InstallWindowsUpdates.bat"
   $Psfilename="InstallWindowsUpdates.ps1"

   $sessions = New-PSSession -ComputerName $computername -Credential $cred
   Enter-PSSession -Session $sessions
   $remotejob = Invoke-Command -Session $sessions -ScriptBlock { param($batfilename) & cmd.exe /c "C:\$batfilename" } -ArgumentList $batfilename -AsJob
   $remotejob | Wait-Job #wait for the remote job to complete 
   Remove-Item -Path "\\$computername\C$\$batfilename" -Force 
   Remove-Item -Path "\\$computername\C$\$Psfilename" -Force 

   Remove-PSSession -Session $sessions #remove the PSSession once it is done


}

Function Format-InstallPatchLog {
    [cmdletbinding()]
    param ($computername)
     
      $installreport = @()
    #Check for logfile
    If (Test-Path "\\$computername\c$\$($computername)_patchlog.csv") {
        #Retrieve the logfile from remote server
        
         
        $file = Import-Csv "\\$computername\c$\$($computername)_patchlog.csv" 
        $installreport=$file
             
        }
    Else {

    write-host "[Error]: File Path Error Occured in Server - $computername " -ForegroundColor Red 
       $temp = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
        
        $temp.Computer = $computername
        $temp.Title = "ERROR"
        $temp.KB = "File Path Error \\$computername\c$\$($computername)_patchlog.csv"
        $temp.IsDownloaded = "File Path Error \\$computername\c$\$($computername)_patchlog.csv"
        $temp.LastRebootTime = "File Path Error \\$computername\c$\$($computername)_patchlog.csv"
        $temp.LastPatchInstalledTime = "File Path Error \\$computername\c$\$($computername)_patchlog.csv"
        $temp.Notes = "File Path Error \\$computername\c$\$($computername)_patchlog.csv"
        $temp.OSVersion  = "File Path Error \\$computername\c$\$($computername)_patchlog.csv"
        $Updatereport = $temp      
        }

    Write-Output $installreport

  
    
}

function Install-Patches
{
     
    Param($Computername)
   
    

if(Test-Connection -ComputerName $computername -Count 1 -Quiet)
   {
    
    try{
        
    If (Test-Path C:\psscripts\psexec.exe) {
        
        $returnmessage = Create-UpdateVBS -computer $Computername
       

        if($returnmessage -eq "SUCCESS")
        {
        
        C:\psscripts\PsExec.exe -s \\$computername   C:\InstallWindowsUpdates.bat /quiet /norestart /accepteula -h $powerShellArgs 2> $null 


        $path1 ="\\$computername\c$\InstallWindowsUpdates.bat"
        $path2 = "\\$computername\c$\InstallWindowsUpdates.ps1"
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
                        CreatePs -computername $Computername
                        Format-InstallPatchLog -computer $Computername
                        
                         If (Test-Path "\\$computername\c$\$($computername)_patchlog.csv") {
                         ##just check if file is there else throw error
                         }
                        else{
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
 