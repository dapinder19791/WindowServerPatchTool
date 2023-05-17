 
$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
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
			 C:\psscripts\PsExec.exe \\$computerName restart WinRM
			$result = winrm id -r:$computerName 2>$null
			
			if ($LastExitCode -eq 0) {
                $checkresult=$true
            Write-Host 'WinRM successfully enabled!' -ForegroundColor green}
			 
		} 
		 
	}

return $checkresult
}
   
Function Create-KBsVBS {
Param ($computername)
    #Create Here-String of vbscode to create file on remote system
       
   
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
     if($checkresults){$retryAttempt=$true
     $flag=$true
     }
     else{$retryAttempt=$false
     $flag=$false
     }

    }

    try{
      

               if($flag)
               {
                Remove-PSSession -ComputerName $computername -ErrorAction SilentlyContinue
                net use \\$computername  $password /USER:$userid $powerShellArgs 2> $null
                Copy-Item  -destination \\$computername\c$ -Recurse  -Path .\scripts\Get-KBDetails.bat -Force -ErrorAction SilentlyContinue
                Copy-Item -destination \\$computername\c$ -Recurse -Path   .\scripts\Get-KBDetails.ps1 -Force -ErrorAction SilentlyContinue
                Copy-Item -destination \\$computername\c$ -Recurse -Path .\scripts\KBListFile.txt -Force -ErrorAction SilentlyContinue

                $message="SUCCESS"
               }
              else{
                $message="UnAuthorized"
                 write-host "[Information]: Your Userid - [$userid] is UnAuthorized to connect to Server - $computername " -ForegroundColor Yellow  
              }
              
               
        }
        catch [Exception]
        {
                        $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
                        $report.Computer = $Computername
                        $report.Title = "ERROR:$_.Exception.Message"
                        $report.KB = "ERROR"
                        $report.IsDownloaded = "ERROR"
                        $report.LastRebootTime = "ERROR"
                        $report.LastPatchInstalledTime = "ERROR"
                        $report.Notes = "ERROR:$_.Exception.Message" 
                        $report.OSVersion = "ERROR:$_.Exception.Message" 
                        $report.SqlVersion= "ERROR:$_.Exception.Message" 
                        $message="ERROR:$_.Exception.Message"        
        
        }
        
        return $message   
   
  
}

function CreatePs
{
  [cmdletbinding()]
    param ($computername)

   $batfilename="Get-KBDetails.bat"
   $Psfilename="Get-KBDetails.ps1"
   $KBFile = "KBListFile.txt"
   $sessions = New-PSSession -ComputerName $computername -Credential $cred
   Enter-PSSession -Session $sessions
   $remotejob = Invoke-Command -Session $sessions -ScriptBlock { param($batfilename) & cmd.exe /c "C:\$batfilename" } -ArgumentList $batfilename -AsJob
   $remotejob | Wait-Job #wait for the remote job to complete 
   Remove-Item -Path "\\$computername\C$\$batfilename" -Force 
   Remove-Item -Path "\\$computername\C$\$Psfilename" -Force 
   Remove-Item -Path "\\$computername\C$\$KBFile" -Force 

   Remove-PSSession -Session $sessions #remove the PSSession once it is done


}

Function Format-Log {
    [cmdletbinding()]
    param ($computername)
     
    #Create empty collection
    $Updatereport = @()
    #Check for logfile
    If (Test-Path "\\$computername\c$\$($computername)_KBDetailslog.csv") {
        #Retrieve the logfile from remote server
        
         
        $file = Import-Csv "\\$computername\c$\$($computername)_KBDetailslog.csv" 
        $Updatereport=$file
             
        }
    Else {
       $temp = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
        
        $temp.Computer = $computername
        $temp.Title = "ERROR"
        $temp.KB = "File Path Error \\$computername\c$\$($computername)_KBDetailslog.csv"
        $temp.IsDownloaded = "File Path Error \\$computername\c$\$($computername)_KBDetailslog.csv"
        $temp.LastRebootTime = "File Path Error \\$computername\c$\$($computername)_KBDetailslog.csv"
        $temp.LastPatchInstalledTime = "File Path Error \\$computername\c$\$($computername)_KBDetailslog.csv"
        $temp.Notes = "File Path Error \\$computername\c$\$($computername)_KBDetailslog.csv"
        $temp.OSVersion  = "File Path Error \\$computername\c$\$($computername)_KBDetailslog.csv"
        $temp.SqlVersion= "File Path Error \\$computername\c$\$($computername)_KBDetailslog.csv"
        $Updatereport = $temp      
        }

    Write-Output $Updatereport
    
}

function GetKB
{
    [cmdletbinding()]
    Param($Computername)
   
    

if(Test-Connection -ComputerName $computername -Count 1 -Quiet)
   {
    
    try{
        
    If (Test-Path C:\psscripts\psexec.exe) {
        
        $returnmessage = Create-KBsVBS -computer $Computername  
       

        if($returnmessage -eq "SUCCESS")
        { 
        C:\psscripts\PsExec.exe -s \\$computername   C:\Get-KBDetails.bat /quiet /norestart /accepteula -h $powerShellArgs 2> $null 
        $path1 ="\\$computername\c$\Get-KBDetails.bat"
        $path2 = "\\$computername\c$\Get-KBDetails.ps1"
        $path3 = "\\$computername\c$\KBListFile.txt"
                if([System.IO.File]::Exists($path1)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path1   $powerShellArgs 2> $null  
                        }
 
                if([System.IO.File]::Exists($path2)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path2   $powerShellArgs 2> $null  
                        }
              
                if([System.IO.File]::Exists($path3)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path3   $powerShellArgs 2> $null  
                        }
                         
                    If ($LASTEXITCODE -eq 0) {
                        Format-Log -computer $computername
                        }            
                     Else {
                        CreatePs -computername $computername
                        Format-Log -computer $computername
                        
                         If (Test-Path "\\$computername\c$\$($computername)_KBDetailslog.csv") {
                         ##just check if file is there else throw error
                         }
                        Else {
                            #$host.ui.WriteLine("Unsuccessful run of install script!")
                             $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
                            $report.Computer = $computername
                            $report.Title = "ERROR"
                            $report.KB = "ERROR"
                            $report.IsDownloaded = "ERROR"
                            $report.LastRebootTime = "ERROR"
                            $report.LastPatchInstalledTime = "ERROR"
                            $report.SqlVersion= "ERROR"
                            $report.Notes = "ERROR:Unsuccessful PsExec.exe run of install script" 
                            $report.OSVersion = "ERROR:Unsuccessful PsExec.exe run of install script" 
                            Write-Output $report                 
                            }
                        }
                    }
                     elseif ($returnmessage -eq "UnAuthorized")
                   {
                   $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
                    $report.Computer = $computername
                    $report.Title = "UnAuthorizedAccess"
                    $report.KB = "UnAuthorizedAccess"
                    $report.IsDownloaded = "UnAuthorizedAccess"
                    $report.LastRebootTime = "UnAuthorizedAccess"
                    $report.LastPatchInstalledTime = "UnAuthorizedAccess"
                    $report.Notes = "UnAuthorizedAccess" 
                    $report.OSVersion ="UnAuthorizedAccess" 
                    $report.SqlVersion="UnAuthorizedAccess" 
                    Write-Output $report          

                   
                   }
                   else 
                   {
                    $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
                        $report.Computer = $computername
                        $report.Title = "$returnmessage"
                        $report.KB = "ERROR"
                        $report.IsDownloaded = "ERROR"
                        $report.LastRebootTime = "ERROR"
                        $report.LastPatchInstalledTime = "ERROR"
                        $report.SqlVersion= "ERROR"
                        $report.Notes = "$returnmessage" 
                        $report.OSVersion = "ERROR" 
                        Write-Output $report   
                   
                   }
                }      
            Else {

                        $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
                        $report.Computer = $computername
                        $report.Title = "PSExec not found ! please download it first."
                        $report.KB = "PSExec not found ! please download it first."
                        $report.IsDownloaded = "PSExec not found ! please download it first."
                        $report.LastRebootTime = "PSExec not found ! please download it first."
                        $report.LastPatchInstalledTime = "PSExec not found ! please download it first."
                        $report.Notes = "PSExec not found ! please download it first." 
                        $report.OSVersion = "PSExec not found ! please download it first." 
                        $report.SqlVersion= "PSExec not found ! please download it first." 
                        Write-Output $report

             
                }
     }
     catch{
             $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
            $report.Computer = $computername
            $report.Title = "UnAuthorizedAccess"
            $report.KB = "UnAuthorizedAccess"
            $report.IsDownloaded = "UnAuthorizedAccess"
            $report.LastRebootTime = "UnAuthorizedAccess"
            $report.LastPatchInstalledTime = "UnAuthorizedAccess"
            $report.Notes = "UnAuthorizedAccess" 
            $report.OSVersion ="UnAuthorizedAccess" 
            $report.SqlVersion="UnAuthorizedAccess" 
            Write-Output $report  

       }
   }
   else{
           $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
            $report.Computer = $computername
            $report.Title = "Offline"
            $report.KB = "Offline"
            $report.IsDownloaded = "Offline"
            $report.LastRebootTime = "Offline"
            $report.LastPatchInstalledTime = "Offline"
            $report.Notes = "Offline" 
            $report.OSVersion ="Offline"
            $report.SqlVersion="Offline" 
            Write-Output $report          

   } 
      
}
 
  
 