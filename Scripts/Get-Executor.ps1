$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
$logfilename =  $Optionshash['DiagPath'] + "\" + $Optionshash['DiagName']

Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}
function global:accessibility{
param($userid,$pass)
 
                            if($userid -ne $null)
                            {
                            $userid =$userid 
                            $securestring  = convertto-securestring -string $pass 
                            }

                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass

                            $password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))

return $cred
}
function global:getpass{
param($pass)

return $password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($pass))

}
function global:getCredentials{
param($userid,$password)
 
$global:cred=accessibility -userid $userid -pass $password
return $cred 
}
function global:createSession{
param($userid,$password,$computer)
try{
  $global:cred=getCredentials -userid $userid -password $password
  $global:targetSession = New-PSSession  –ComputerName $computer -Credential $cred  -ErrorAction Stop
  }catch{
 
  $targetSession=$null
 
  }
  return $targetSession    
}
Function enableWinRM {
Param([string]$computerName,[string] $userid)
$checkresult=$false

try{
    
    Write-Host "[Server: $computerName] Enabling WinRm with user id : $userid"
    Write-Log -Message "[Server: $computerName] Enabling WinRm with user id : $userid" -logfile $logfilename -Level INFO 


    $passwordtxt=getpass -pass $cred.Password
    net use \\$computerName  $passwordtxt /USER:$userid $powerShellArgs 2> $null
	$result = winrm id -r:$computerName 2>$null

	 
	if ($LastExitCode -eq 0) {
		Write-Host "[Server: $computerName] WinRM already enabled on" $computerName "..." -ForegroundColor green
        Write-Log -Message "[Server: $computerName] WinRM already enabled on" $computerName "..." -logfile $logfilename -Level INFO 


	} else {
		Write-Host "[Server: $computerName] Enabling WinRM on" $computerName "..." -ForegroundColor Blue
		 C:\psscripts\PsExec.exe  \\$computerName -s C:\Windows\system32\winrm.cmd qc -quiet

		if ($LastExitCode -eq 0) {
            Write-Host "[Server: $computerName] Resetting window service - WINRM on" $computerName "..." -ForegroundColor Yellow
             Write-Log -Message "[Server: $computerName] Resetting window service - WINRM on" $computerName "..." -logfile $logfilename -Level INFO 

			 C:\psscripts\PsExec.exe \\$computerName restart WinRM
			$result = winrm id -r:$computerName 2>$null
			
			if ($LastExitCode -eq 0) {
                $checkresult=$true
            Write-Host "[Server: $computerName] WinRM successfully enabled!" -ForegroundColor green
            Write-Log -Message "[Server: $computerName] WinRM successfully enabled!" -logfile $logfilename -Level INFO
            }
			 
		}
        else{
        Write-Host "[Server: $computerName] Unable to reset window service - WINRM on" $computerName "..." -ForegroundColor Red
        Write-Log -Message "[Server: $computerName] Unable to reset window service - WINRM on" $computerName "..." -logfile $logfilename -Level INFO
        } 
		 
	}
}
catch{

}
return $checkresult
}   
Function Create-UpdateVBS {
Param ($computername,$SourceBat,$SourceScript,$log)
    #Create Here-String of vbscode to create file on remote system
    $flag=$false
    $retryAttempt=$false
    $message=""
    $global:userid=""
    $global:password=""

    try{
    
    $userid =$Optionshash['Userid']
    $password=$Optionshash['password']
    $targetSession = createSession -userid $userid -password  $password -computer $computername
    write-host "[Information][Server: $computername] : Attempting to connect to Server - $computername with userid - $userid" -ForegroundColor White
    Write-Log -Message "[Information][Server: $computername] : Attempting to connect to Server - $computername with userid - $userid" -logfile $logfilename -Level INFO

     if($targetSession -eq $null){
     if($Optionshash['Userid1'] -ne $null)
        {
         $userid = $Optionshash['Userid1']
          $password=$Optionshash['password1']

        $targetSession = createSession -userid $userid -password $password -computer $computername
         write-host "[Information][Server: $computername]: Attempting to connect to Server - $computername with userid - $userid" -ForegroundColor White
         Write-Log -Message "[Information][Server: $computername]: Attempting to connect to Server - $computername with userid - $userid" -logfile $logfilename -Level INFO
          
        }
     }
     else{$flag=$true}
     if($targetSession -eq $null)
     {
      if($Optionshash['Userid2'] -ne $null)
        {
      $userid = $Optionshash['Userid2']
      $password=$Optionshash['password2']
        $targetSession = createSession -userid $userid -password $password -computer $computername
        write-host "[Information][Server: $computername]: Attempting to connect to Server - $computername with userid - $userid" -ForegroundColor White
        Write-Log -Message "[Information][Server: $computername]: Attempting to connect to Server - $computername with userid - $userid" -logfile $logfilename -Level INFO
        }
     }else{
     $flag=$true
     
     }
     if($targetSession -ne $null)
     {
        write-host "[Information][Server: $computername] : Connection is Established with Server - $computername with userid - [$userid]" -ForegroundColor Green
        Write-Log -Message "[Information][Server: $computername] : Connection is Established with Server - $computername with userid - [$userid]" -logfile $logfilename -Level INFO
     }
     else{
     write-host "[Warning][Server: $computername]: All Attempts to connect to Server - $computername failed.The connection is not established." -ForegroundColor Yellow
     Write-Log -Message "[Warning][Server: $computername]: All Attempts to connect to Server - $computername failed.The connection is not established." -logfile $logfilename -Level WARN
     }

    }
    catch{
     write-host "[Error][Server: $computername]: Error Connecting Server - $computername with userid - [$userid]" -ForegroundColor Red
     Write-Log -Message "[Error][Server: $computername]: Error Connecting Server - $computername with userid - [$userid]" -logfile $logfilename -Level ERROR

     write-host "[Information][Server: $computername]: Retry Attempt connecting to Server - $computername with alternating approach" -ForegroundColor Yellow
      Write-Log -Message "[Information][Server: $computername]: Retry Attempt connecting to Server - $computername with alternating approach" -logfile $logfilename -Level INFO

     write-host "[Information][Server: $computername]: Enabling winRM in - $computername with alternating approach" -ForegroundColor Yellow
     Write-Log -Message "[Information][Server: $computername]: Enabling winRM in - $computername with alternating approach"  -logfile $logfilename -Level INFO

     $checkresults = enableWinRM -computerName $computername -userid $userid

     if($checkresults){
     $retryAttempt=$true
     $flag=$true
     $targetSession =$null
     write-host "[Information][Server: $computername]: Retry Attempt connecting to Server - $computername with alternating Approach is successfull" -ForegroundColor Green
     Write-Log -Message "[Information][Server: $computername]: Retry Attempt connecting to Server - $computername with alternating Approach is successfull"  -logfile $logfilename -Level INFO

     }
     else{
     $retryAttempt=$false
     $flag=$false
     write-host "[Error][Server: $computername]: Retry Attempt Failed connecting to Server - $computername with alternating Approach is successfull" -ForegroundColor Red
     Write-Log -Message "[Error][Server: $computername]: Retry Attempt Failed connecting to Server - $computername with alternating Approach is successfull"  -logfile $logfilename -Level ERROR

     }

    }

    try{
      $message=""

               if($flag -eq $true -and ($targetSession -ne $null))
               {
                    $message="Initiating Copy process by Network Session" 
                    $source1=".\scripts\$SourceBat"
                    $target1="C:\"
                    Copy-Item -Force -Recurse  -Path $source1 -Destination $target1 -ToSession $targetSession  -ErrorAction Stop
                    $source2=".\scripts\$SourceScript"
                    $target2="C:\"
                    Copy-Item -Force -Recurse  -Path $source2 -Destination $target2 -ToSession $targetSession -ErrorAction Stop
                    write-host "[Information][Server: $computername] : $message" -ForegroundColor White       
                    Write-Log -Message "[Information][Server: $computername] : $message"  -logfile $logfilename -Level INFO

                    write-host "[Information][Server: $computername]: Copy process to Server - $computername Completed" -ForegroundColor White
                    Write-Log -Message "[Information][Server: $computername]: Copy process to Server - $computername Completed" -logfile $logfilename -Level INFO

                    $message="SUCCESS"

               }
               elseif($flag -eq $true -and ($targetSession -eq $null)) 
               {
                  
                  
                    Write-Host "[Information][Server: $computername]: Attempting file access with userid : $userid" -ForegroundColor White
                    Write-Log -Message "[Information][Server: $computername]: Attempting file access with userid : $userid" -logfile $logfilename  -Level INFO

                    $message="Initiating Copy process by Network share"
              
                    $passwordtxt=getpass -pass $cred.Password
                    net use \\$computername  $passwordtxt /USER:$userid $powerShellArgs 2> $null
                    Copy-Item  -destination \\$computername\c$ -Recurse  -Path .\scripts\$SourceBat -Force -ErrorAction SilentlyContinue
                    Copy-Item -destination \\$computername\c$ -Recurse -Path   .\scripts\$SourceScript -Force -ErrorAction SilentlyContinue
                     write-host "[Information][Server: $computername] : $message" -ForegroundColor White       
                     Write-Log -Message "[Information][Server: $computername] : $message" -logfile $logfilename  -Level INFO

                    write-host "[Information][Server: $computername]: Copy process to Server - $computername Completed" -ForegroundColor White
                    Write-Log -Message "[Information][Server: $computername]: Copy process to Server - $computername Completed" -logfile $logfilename  -Level INFO

                    $message="SUCCESS"
                   }
                
              else{
                $message="UnAuthorized"
                 write-host "[Warning][Server: $computername]: Userid - [$userid] is UnAuthorized to connect to Server - $computername " -ForegroundColor Yellow 
                 Write-Log -Message "[Warning][Server: $computername]: Userid - [$userid] is UnAuthorized to connect to Server - $computername "  -logfile $logfilename  -Level WARN
                  
              }
              if($targetSession -ne $null){
                Remove-PSSession  -Session  $targetSession -ErrorAction SilentlyContinue
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
                        write-host "[Error][Server: $computername]: Error Occured in Server - $computername , [Error Message :] " $_.Exception.Message -ForegroundColor Red  
                        Write-Log -Message "[Error][Server: $computername]: Error Occured in Server - $computername , [Error Message :] " $_.Exception.Message -logfile $logfilename  -Level ERROR
        }
        
        return $message   
   
  
}
function CreatePs{
  [cmdletbinding()]
    param ($computername,$batfilename,$Psfilename)

   $BATPATH="C:\$batfilename"
   $PSPATH="C:\$Psfilename"
 
   $targetSession = New-PSSession  –ComputerName $computername -Credential $cred  -ErrorAction Stop

   $remotejob = Invoke-Command -Session $targetSession -ScriptBlock { param($BATPATH) & cmd.exe /c  $BATPATH } -ArgumentList $BATPATH -AsJob
   $remotejob | Wait-Job #wait for the remote job to complete 

   Invoke-Command -Session $targetSession -ScriptBlock { 
   
   Remove-Item -Path  $using:BATPATH -Force 
   Remove-Item -Path  $using:PSPATH -Force 

   } -ArgumentList $BATPATH,$PSPATH 
   

  
   Remove-PSSession -Session $targetSession #remove the PSSession once it is done


}
Function Format-InstallPatchLog {
    [cmdletbinding()]
    param ($computername,$mode,$log)
     
    #Create empty collection
    $Updatereport = @()
    $LOGPATH="\\$computername\c$\$($computername)_$log"
    $SESSIONLOGPATH="c:\$($computername)_$log"
  

    if($mode -eq 1)
    {
    If (Test-Path $LOGPATH) {
        #Retrieve the logfile from remote server
        $file = Import-Csv $LOGPATH
        $Updatereport=$file
          Write-Output $Updatereport     
        }
        Else {

       write-host "[Error]: File Path Error Occured in Server - $computername " -ForegroundColor Red  
       Write-Log -Message "[Error]: File Path Error Occured in Server - $computername "  -logfile $logfilename  -Level ERROR

       $temp = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
        
        $temp.Computer = $computername
        $temp.Title = "ERROR"
        $temp.KB = "File Path Error $LOGPATH"
        $temp.IsDownloaded = "File Path Error $LOGPATH"
        $temp.LastRebootTime = "File Path Error $LOGPATH"
        $temp.LastPatchInstalledTime = "File Path Error $LOGPATH"
        $temp.Notes = "File Path Error $LOGPATH"
        $temp.OSVersion  = "File Path Error $LOGPATH"
        $temp.SqlVersion= "File Path Error $LOGPATH"
        $Updatereport = $temp      
          Write-Output $Updatereport
        }
    }
    else{
   
       $sessions = New-PSSession -ComputerName $computername -Credential $cred
        Invoke-Command -Session $sessions -ScriptBlock {
     
        #Retrieve the logfile from remote server
        $validate = Test-Path  $using:SESSIONLOGPATH
       
        if($validate -eq $true)
            { 
            $file = Import-Csv $using:SESSIONLOGPATH
            $Updatereport=$file
            Write-Output $Updatereport

	 	    }
            else{
                $Updatereport=$null
            }
            if($Updatereport -eq $null)
                {

                   write-host "[Error][Server: $computername]: File Path Error Occured in Server - $computername " -ForegroundColor Red  
                   Write-Log -Message "[Error][Server: $computername]: File Path Error Occured in Server - $computername "  -logfile $logfilename  -Level ERROR

                   $temp = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
        
                    $temp.Computer = $computername
                    $temp.Title = "ERROR"
                    $temp.KB = "File Path Error $LOGPATH"
                    $temp.IsDownloaded = "File Path Error $LOGPATH"
                    $temp.LastRebootTime = "File Path Error $LOGPATH"
                    $temp.LastPatchInstalledTime = "File Path Error $LOGPATH"
                    $temp.Notes = "File Path Error $LOGPATH"
                    $temp.OSVersion  = "File Path Error $LOGPATH"
                    $temp.SqlVersion= "File Path Error $LOGPATH"
                    $Updatereport = $temp     
                     Write-Output $Updatereport
          
                    } 

		} -ArgumentList $computername,$SESSIONLOGPATH

        Remove-PSSession -Session $sessions #remove the PSSession once it is done
        }

    }
function Getupdate-Patches{
    [cmdletbinding()]
    Param($Computername,$batfilename,$Psfilename,$log)
   
    

if(Test-Connection -ComputerName $computername -Count 1 -Quiet)
   {
    
    try{
        
    If (Test-Path C:\psscripts\psexec.exe) {
        
        $returnmessage = Create-UpdateVBS -computer $Computername -SourceBat $batfilename -SourceScript $Psfilename -log $log
        

        if($returnmessage -eq "SUCCESS"){ 
        
        write-host "[Information][Server: $Computername]: Attempting to Execute the package files in Server - $Computername " -ForegroundColor White 
        Write-Log -Message "[Information][Server: $Computername]: Attempting to Execute the package files in Server - $Computername " -logfile $logfilename  -Level INFO

        $path="C:\$batfilename"

        C:\psscripts\PsExec.exe -s \\$computername $path  /quiet /norestart /accepteula -h $powerShellArgs 2> $null 
        $path1 ="\\$computername\c$\$batfilename"
        $path2 = "\\$computername\c$\$Psfilename"
                if([System.IO.File]::Exists($path1)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path1   $powerShellArgs 2> $null  
                        }
 
                if([System.IO.File]::Exists($path2)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path2   $powerShellArgs 2> $null  
                        }
                
                    If ($LASTEXITCODE -eq 0) {
        write-host "[Information][Server: $Computername]: Formatting Files for Display - $Computername " -ForegroundColor White 
        Write-Log -Message "[Information][Server: $Computername]: Formatting Files for Display - $Computername " -logfile $logfilename  -Level INFO

                        Format-InstallPatchLog -computer $Computername -mode 1 -log $log
                        }            
                     Else {
        write-host "[Information][Server: $Computername]: Formatting Files for Display - $Computername " -ForegroundColor White 
          Write-Log -Message "[Information][Server: $Computername]: Formatting Files for Display - $Computername " -logfile $logfilename  -Level INFO             
                       CreatePs -computername $Computername -batfilename $batfilename -Psfilename $Psfilename
                        Format-InstallPatchLog -computer $Computername -mode 2 -log $log
                        
                        
                    }
           
           }  
        elseif ($returnmessage -eq "UnAuthorized"){
                   $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
                    $report.Computer = $Computername
                    $report.Title = "UnAuthorizedAccess"
                    $report.KB = "UnAuthorizedAccess"
                    $report.IsDownloaded = "UnAuthorizedAccess"
                    $report.LastRebootTime = "UnAuthorizedAccess"
                    $report.LastPatchInstalledTime = "UnAuthorizedAccess"
                    $report.Notes = "UnAuthorizedAccess" 
                    $report.OSVersion ="UnAuthorizedAccess" 
                    $report.SqlVersion="UnAuthorizedAccess" 
                    write-host "[Warning][Server: $Computername]: You are Not authorized to access the server - $Computername " -ForegroundColor Yellow
                     Write-Log -Message "[Warning][Server: $Computername]: You are Not authorized to access the server - $Computername " -logfile $logfilename  -Level WARN 
                    Write-Output $report          

                   
                   }
        else {
                    $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
                        $report.Computer = $Computername
                        $report.Title = "$returnmessage"
                        $report.KB = "ERROR"
                        $report.IsDownloaded = "ERROR"
                        $report.LastRebootTime = "ERROR"
                        $report.LastPatchInstalledTime = "ERROR"
                        $report.SqlVersion= "ERROR"
                        $report.Notes = "$returnmessage" 
                        $report.OSVersion = "ERROR" 
                        write-host "[Error][Server: $Computername]: Error has occured in the server - $Computername; Error Message : $returnmessage " -ForegroundColor Red
                        Write-Log -Message "[Error][Server: $Computername]: Error has occured in the server - $Computername; Error Message : $returnmessage " -logfile $logfilename  -Level ERROR
                        Write-Output $report   
                   
                   }
                }      
    Else {

                        $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
                        $report.Computer = $Computername
                        $report.Title = "PSExec not found ! please download it first."
                        $report.KB = "PSExec not found ! please download it first."
                        $report.IsDownloaded = "PSExec not found ! please download it first."
                        $report.LastRebootTime = "PSExec not found ! please download it first."
                        $report.LastPatchInstalledTime = "PSExec not found ! please download it first."
                        $report.Notes = "PSExec not found ! please download it first." 
                        $report.OSVersion = "PSExec not found ! please download it first." 
                        $report.SqlVersion= "PSExec not found ! please download it first." 
                        write-host "[Error][Server: $Computername]: Error has occured in the server - $Computername; Error Message : PSExec not found ! please download it first. " -ForegroundColor Red
                         Write-Log -Message "[Error][Server: $Computername]: Error has occured in the server - $Computername; Error Message : PSExec not found ! please download it first. " -logfile $logfilename  -Level ERROR 
                        Write-Output $report

             
                }
     
     }
     catch{
             $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
            $report.Computer = $Computername
            $report.Title = "UnAuthorizedAccess"
            $report.KB = "UnAuthorizedAccess"
            $report.IsDownloaded = "UnAuthorizedAccess"
            $report.LastRebootTime = "UnAuthorizedAccess"
            $report.LastPatchInstalledTime = "UnAuthorizedAccess"
            $report.Notes = "UnAuthorizedAccess" 
            $report.OSVersion ="UnAuthorizedAccess" 
            $report.SqlVersion="UnAuthorizedAccess" 
            write-host "[Warning][Server: $Computername]: You are Not authorized to access the server - $Computername " -ForegroundColor Yellow
             Write-Log -Message "[Warning][Server: $Computername]: You are Not authorized to access the server - $Computername " -logfile $logfilename  -Level WARN 

            Write-Output $report  

       }
   }
   else{
           $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion,SqlVersion
            $report.Computer = $Computername
            $report.Title = "Offline"
            $report.KB = "Offline"
            $report.IsDownloaded = "Offline"
            $report.LastRebootTime = "Offline"
            $report.LastPatchInstalledTime = "Offline"
            $report.Notes = "Offline" 
            $report.OSVersion ="Offline"
            $report.SqlVersion="Offline" 
            write-host "[Warning][Server: $Computername]: The server - $Computername is Offline. Please Ensure the Server should be in Running State first." -ForegroundColor Yellow
            Write-Log -Message "[Warning][Server: $Computername]: The server - $Computername is Offline. Please Ensure the Server should be in Running State first." -logfile $logfilename  -Level WARN

            Write-Output $report          

   } 
      
}
 