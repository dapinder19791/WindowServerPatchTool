$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass

$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))
  



Function  Reboot-Server {
    [cmdletbinding()]
    param ($computername)
  $Rebootreport = @()
  try{
if(Test-Connection -ComputerName $computername)
{
   

  net use \\$computername  $password /USER:$userid $powerShellArgs 2> $null
  psexec -s \\$computername  C:\windows\system32\shutdown.exe -r 
  $temp = "" | Select Computer, Title ,Notes
  $temp.Computer = $computername
  $temp.Title = "RebootInitiated"
  $temp.Notes = "RebootInitiated"
  $Rebootreport += $temp      
  Write-Host $Rebootreport
  Write-Output $Rebootreport
 }
 else{
  $temp = "" | Select Computer, Title ,Notes
  $temp.Computer = $computername
  $temp.Title = "offline"
  $temp.Notes = "offline"
  $Rebootreport += $temp      
  Write-Host $Rebootreport
  Write-Output $Rebootreport
 
 } 
  }catch{
  
  
  $temp = "" | Select Computer, Title ,Notes
  $temp.Computer = $computername
  $temp.Title = "ERROR"
  $temp.Notes = "ERROR"
  $Rebootreport += $temp      
  Write-Host $Rebootreport
  Write-Output $Rebootreport
 
  }
}