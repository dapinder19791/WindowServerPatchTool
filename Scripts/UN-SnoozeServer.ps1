 $report =@()
 
$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass

$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))
  
  

 
Function UnSnoozeServer {
Param ($computername)
            

$computername="$computername" 
$list=@($computername)
    
    #Create Here-String of vbscode to create file on remote system
                try{
             
             $snooze = @{
                 Stop=$false;
                 TargetedVMs=$list;
                 }
          

                  $json = $snooze | ConvertTo-Json 
                  write-host "Calling ... http://sdoazureutil/api/ACRPUtil/PostStartSnoozeJob/"
                  write-host "Passing JSON object ......"

                  write-host $json 
                 
                 

                  $response = Invoke-RestMethod  -Method POST -Uri 'http://sdoazureutil/api/ACRPUtil/PostStartSnoozeJob/ ' -Body $json   -ContentType 'application/json'  -Credential $cred  
                   
                 $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                        $report.Computer = $computername
                        $report.Title = "Un-Snooze Task"
                        $report.KB = "NA"
                        $report.IsDownloaded = "NA"
                        $report.LastRebootTime = "NA"
                        $report.LastPatchInstalledTime = "NA"
                        $report.Notes = "Completed" 
                        $report.OSVersion = "NA" 
                              
         Write-Output $report

                  
                   
                 }Catch [Exception]
                 {
                 $message=$_.Exception.Message

                        $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                        $report.Computer = $computername
                        $report.Title = "Error"
                        $report.KB = "ERROR"
                        $report.IsDownloaded = "ERROR"
                        $report.LastRebootTime = "ERROR"
                        $report.LastPatchInstalledTime = "ERROR"
                        $report.Notes =  "ERROR:$message" 
                        $report.OSVersion = "ERROR" 
        Write-Output $report
    
                 }
 
             
    }   
 