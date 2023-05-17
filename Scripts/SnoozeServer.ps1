 $report =@()
 
$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass

$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))
  
   


Function SnoozeServer {
Param ($computername)
$computername="$computername"  

$list=@($computername)

  if (Test-Connection -ComputerName $computername -Count 1 -Quiet){
    #Create Here-String of vbscode to create file on remote system
                try{
            
            
             $snooze = @{
                 Stop=$true;
                 TargetedVMs=$list;
                 }
                 
                  $json = $snooze | ConvertTo-Json
                  $response = Invoke-RestMethod  -Method POST -Uri 'http://sdoazureutil/api/ACRPUtil/PostStartSnoozeJob/ ' -Body $json   -ContentType 'application/json' -Credential $cred
                     
                 $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                        $report.Computer = $computername
                        $report.Title = "Snooze Task"
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
 
            }else{
           
                        $report = "" | Select Computer, Title, KB,IsDownloaded,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion
                        $report.Computer = $computername
                        $report.Title = "Offline"
                        $report.KB = "Offline"
                        $report.IsDownloaded = "Offline"
                        $report.LastRebootTime = "Offline"
                        $report.LastPatchInstalledTime = "Offline"
                        $report.Notes =  "Offline" 
                        $report.OSVersion = "Offline" 
                        Write-Output $report

            
            }
    }   

     