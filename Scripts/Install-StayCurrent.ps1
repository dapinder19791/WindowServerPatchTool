$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass

$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))
  

Function Copy-Packages
{
Param ($computername)

 if (Test-Connection -ComputerName $computername -Count 1 -Quiet){
  
  
   
                net use \\$computername   $password /USER:$userid $powerShellArgs 2> $null
                Copy-Item  -destination \\$computername\c$ -Recurse  -Path .\Scripts\RunStayCurrentupdates.bat -Force -ErrorAction SilentlyContinue
                Copy-Item -destination \\$computername\c$ -Recurse -Path  .\Scripts\RunStayCurrentupdates.ps1 -Force -ErrorAction SilentlyContinue
                Invoke-Command -ComputerName $computername  -ScriptBlock { Remove-ItemProperty -Path 'HKLM:System\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue} -ErrorAction SilentlyContinue 
                
 

 }

}


   Function Format-InstallPatchLog {
    [cmdletbinding()]
    param ($computername)
     
    #Create empty collection
    $installreport = @()
    #Check for logfile
    If (Test-Path "\\$computername\c$\$($computername)_StayCurrentPatchLog.csv") {
        #Retrieve the logfile from remote server
        $CSVreport = Import-Csv "\\$computername\c$\$($computername)_StayCurrentPatchLog.csv" -Delimiter " "
        #Iterate through all items in patchlog
        ForEach ($log in $CSVreport) {
             
            $temp = "" | Select Computer,Title,KB,IsDownloaded,Notes
            $temp.Computer = $computername
            $statuschk1 =  @($log | Where {$_.Title -eq "successfully"})
            $statuschk2 =  @($log | Where {$_.KB -eq "completed"})
            $statuschk3 =  @($log | Where {$_.Title -eq "reboot"})
            $statuschk4 =  @($log | Where {$_.KB -eq "reboot"})
            $statuschk5 =  @($log | Where {$_.Title -eq "errors"})
            $statuschk6 =  @($log | Where {$_.KB -eq "errors"})
            $statuschk7 =  @($log | Where {$_.Title -eq "failed"})
            $statuschk8 =  @($log | Where {$_.KB -eq "failed"})

            $cnt=0;
            if($statuschk1.Title -eq "successfully")
            {
            $cnt=1
            }
            if($statuschk2.KB -eq "completed")
            {
            $cnt=1
            }
            if($statuschk3.Title -eq "reboot")
            {
             $cnt=0
            }

            if($statuschk4.KB -eq "reboot")
            {
             $cnt=0
            }

             if($statuschk5.Title -eq "errors")
            {
             $cnt=3
            }

            if($statuschk6.KB -eq "errors")
            {
             $cnt=3
            }

             if($statuschk7.Title -eq "failed")
            {
             $cnt=4
            }

            if($statuschk8.KB -eq "failed")
            {
             $cnt=4
            }


            Switch ($cnt) {
                0 {
                $temp.Notes = "Reboot is required"
                $temp.Title= "Reboot is required"

                }
                1 {$temp.Notes = "Patching Completed Successfully"
                $temp.Title= "Patching Completed Successfully"
                }
                2 {$temp.Notes = "No Updates Available"
                $temp.Title= "No Updates Available"
                }
                3 {$temp.Notes = "Error with Audit"
                $temp.Title= "Error with Audit"
                
                }
                
                4 {$temp.Notes = "failed"
                $temp.Title= "failed"
                
                }
                Default {$temp.Notes = "Unable to determine Result Code"}            
                }
            $installreport = $temp
            }
        }
    Else {
        $temp = "" | Select Computer, Title, KB,IsDownloaded,Notes
        $temp.Computer = $computername
        $temp.Title = "offline"
        $temp.KB = "offline"
        $temp.IsDownloaded = "offline"
        $temp.Notes = "offline"  
        $installreport = $temp      
        }

 #   Write-Host $installreport
Write-Output $installreport
    
}


Function Stay-Current
{
   [cmdletbinding()]
    param ($computername)
 
    
$count=0;
 
try{
 

        if (Test-Connection -ComputerName $computername -Count 1 -Quiet){
                If (Test-Path C:\psscripts\psexec.exe) {

                Copy-Packages -computer $Computername

                C:\psscripts\PsExec.exe -s \\$computername   C:\RunStayCurrentupdates.bat /quiet /norestart /accepteula -h $powerShellArgs 2> $null  

                $path1 ="\\$computername\c$\RunStayCurrentupdates.bat"
                $path2 = "\\$computername\c$\RunStayCurrentupdates.ps1"

                if([System.IO.File]::Exists($path1)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path1   $powerShellArgs 2> $null  
                        }
 
                if([System.IO.File]::Exists($path2)){
                    C:\psscripts\PsExec.exe \\$computername cmd /c del $path2   $powerShellArgs 2> $null  
                        }
               } 
                   If ($LASTEXITCODE -eq 0) {
                #$host.ui.WriteLine("Successful run of install script!")
                Format-InstallPatchLog -computer $computername
                }            
            Else {
                #$host.ui.WriteLine("Unsuccessful run of install script!")
                $report = "" | Select Computer,Title,KB,IsDownloaded,Notes
                $report.Computer = $computername
                $report.Title = "ERROR"
                $report.KB = "ERROR"
                $report.IsDownloaded = "ERROR"
                $report.Notes = "ERROR" 
                Write-Output $report
                }
 
                 
        }
        else {
          $report = "" | Select Computer,Title,KB,IsDownloaded,Notes
                $report.Computer = $computername
                $report.Title = "offline"
                $report.KB = "offline"
                $report.IsDownloaded = "offline"
                $report.Notes = "offline" 
                Write-Output $report

         }
         
     }catch{


     $ErrorMessage = $_.Exception.Message
     $FailedItem = $_.Exception.ItemName

     write-host $ErrorMessage

       $report = "" | Select Computer,Title,KB,IsDownloaded,Notes
                $report.Computer = $computername
                $report.Title = "ERROR"
                $report.KB = "ERROR"
                $report.IsDownloaded = "ERROR"
                $report.Notes = "ERROR" 
                Write-Output $report
     }
   } 


 