
 
                            $userid = "redmond\gspperf"
                            $securestring  = convertto-securestring -string  (Get-Content "C:\PatchingToolKit\securestring.txt")
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass
 
$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))


Function Create-UpdateVBS {
Param ($computername)
    

   
                net use \\$computername  $password /USER:$userid $powerShellArgs 2> $null
                Copy-Item  -destination \\$computername\c$ -Recurse  -Path .\scripts\checkupdate.bat -Force -ErrorAction SilentlyContinue
                Copy-Item -destination \\$computername\c$ -Recurse -Path   .\scripts\checkupdate.ps1 -Force -ErrorAction SilentlyContinue
           
  
  
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
        $temp = "" | Select Computer, Title, KB,IsDownloaded,Notes,OSVersion
        $temp.Computer = $computername
        $temp.Title = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"
        $temp.KB = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"
        $temp.IsDownloaded = "File Path Error \\azmarprd06\ZeroVulnPatch\$($computername)_Zeropatchlog.csv"
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
     
     # $pass = cat .\securestring.txt | ConvertTo-SecureString
      #$cred = new-object -typename System.Management.Automation.PSCredential `
       #                                 -argumentlist $username, $pass
      #$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))
       
      net use \\$Computername  $password /USER:$userid $powerShellArgs 2> $null

       
    If (Test-Path C:\psscripts\psexec.exe) {
    
      
        Create-UpdateVBS -computer $Computername
         
        C:\psscripts\PsExec.exe -s \\$computername   C:\checkupdate.bat /quiet /norestart /accepteula -h $powerShellArgs 2> $null 

        
                  }
           Else {

                Write-Verbose "PSExec not found ! please download it first."        
                Write-Host "PSExec not found ! please download it first."
             
                        $report = "" | Select Computer,Title,KB,IsDownloaded,Notes,OSVersion
                        $report.Computer = $Computername
                        $report.Title = "PSExec not found ! please download it first."
                        $report.KB = "PSExec not found ! please download it first."
                        $report.IsDownloaded = "PSExec not found ! please download it first."
                        $report.Notes = "PSExec not found ! please download it first." 
                        $report.OSVersion = "PSExec not found ! please download it first." 
                        Write-Output $report

             
                }
     }
     catch{
            $report = "" | Select Computer,Title,KB,IsDownloaded,Notes,OSVersion
            $report.Computer = $Computername
            $report.Title = "UnAuthorizedAccess"
            $report.KB = "UnAuthorizedAccess"
            $report.IsDownloaded = "UnAuthorizedAccess"
            $report.Notes = "UnAuthorizedAccess" 
            $report.OSVersion ="UnAuthorizedAccess" 
            Write-Output $report          

       }
   }
   else{
     
            $report = "" | Select Computer,Title,KB,IsDownloaded,Notes,OSVersion
            $report.Computer = $Computername
            $report.Title = "Offline"
            $report.KB = "Offline"
            $report.IsDownloaded = "Offline"
            $report.Notes = "Offline" 
            $report.OSVersion ="Offline"
            Write-Output $report          

   } 
      
}

  Getupdate-VulnPatches -computername "AZOEMCRMTOOLS"
