
 $Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass
 
$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))

 



Function Create-UpdateVBS {
Param ($computername)
    
    $flag=$false
    $message=""

    try{
     write-host "[Information] : Attempting to connect to Server - $computername with userid - [$userid]" -ForegroundColor White
     Enter-PSSession –ComputerName $computername -Credential $cred  -ErrorAction Stop
     $flag=$true
     write-host "[Information]: Connection is Established with Server - $computername with userid - [$userid]" -ForegroundColor Green
    }
    catch{
     write-host "[Error]: Error Connecting Server - $computername with userid - [$userid]" -ForegroundColor Red

     $flag=$false
    }

    try{
      

               if($flag)
               {
                Remove-PSSession -ComputerName $computername -ErrorAction SilentlyContinue
                net use \\$computername  $password /USER:$userid $powerShellArgs 2> $null
                Copy-Item  -destination \\$computername\c$ -Recurse  -Path .\scripts\Get-VulnPatchScan.bat  -Force -ErrorAction SilentlyContinue 
                Copy-Item -destination \\$computername\c$ -Recurse -Path   .\scripts\Get-VulnPatchScan.ps1  -Force -ErrorAction SilentlyContinue
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
        
        }
        
        return $message   
  
  
}


Function Format-InstallPatchLog {
    [cmdletbinding()]
    param ($computername)
     
    #Create empty collection
    $Updatereport = @()
    #Check for logfile
    If (Test-Path "\\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv") {
        #Retrieve the logfile from remote server
          
            $file = Import-Csv "\\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv" 
            $Updatereport=$file
           
        }
    Else {
        $temp = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
        $temp.Computer = $computername
        $temp.Title = "ERROR"
        $temp.KB = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"
        $temp.IsDownloaded = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"
        $temp.LastRebootTime = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"
        $temp.LastPatchInstalledTime = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"

        $temp.Notes = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"
        $temp.OSVersion  = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"
        $Updatereport = $temp      
        }

    Write-Output $Updatereport
    
}

function Getupdate-VulnPatches
{
    [cmdletbinding()]
    Param($Computername)
     
 
if(Test-Connection -ComputerName $Computername -Count 1 -Quiet)
   {
     
    try{
     
     
       
    If (Test-Path C:\psscripts\psexec.exe) {
    
      
        $returnmessage = Create-UpdateVBS -computer $Computername
       

        if($returnmessage -eq "SUCCESS")
        { 
        C:\psscripts\PsExec.exe -s \\$computername  C:\Get-VulnPatchScan.bat  -u $userid -p $password  /quiet /norestart /accepteula -h $powerShellArgs 2> $null 

        
        $path1 ="\\$computername\c$\Get-VulnPatchScan.bat"
        $path2 = "\\$computername\c$\Get-VulnPatchScan.ps1"
        
       
                if([System.IO.File]::Exists($path1)){
                    C:\psscripts\PsExec.exe \\$computername -u $userid -p $password cmd /c del $path1   $powerShellArgs 2> $null  
                        }
 
                if([System.IO.File]::Exists($path2)){
                    C:\psscripts\PsExec.exe \\$computername -u $userid -p $password cmd /c del $path2   $powerShellArgs 2> $null  
                        }
       
                    $ccheckpath=Test-Path "\\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"

                    If ($LASTEXITCODE -eq 0 -or $ccheckpath -eq $true) {
        
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

                Write-Verbose "PSExec not found ! please download it first."        
                Write-Host "PSExec not found ! please download it first."
             
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

 