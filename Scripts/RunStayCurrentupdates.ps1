function check-PendingReboot
{
 if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { return $true }
 if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { return $true }
 #if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA Ignore) { return $true }
 try { 
   $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
   $status = $util.DetermineIfRebootPending()
   if(($status -ne $null) -and $status.RebootPending){
     return $true
   }
 }catch{}
 
 return $false
}



$Servers = $env:computername
foreach ($Computer in $Servers){
if (Test-Connection -ComputerName $Computer -Count 1 -Quiet){
## Log file path
$outfilePath =  "\\$Computer\c$\$($Computer)_StayCurrentPatchLog.csv"

if([System.IO.File]::Exists($outfilePath)){ 

  Remove-Item $outfilePath -force -recurse

 }
$txt ="Computer" + " " + "Title" + " " + "KB" + " " + "IsDownloaded" + " " + "Notes" |Out-File $outfilePath -Append
 
$checkreboot = check-PendingReboot 

if($checkreboot)
{
    $txt=$Computer + " " +"RebootRequired" + " " + "NA" + " " + "NA" + " " + "0"  | Out-File $outfilePath -Append
    Exit
}
else{

       # \\gme.gbl\cdds\MSNPLT\GOLD\MSNPATCH\msnpatch.exe | Out-File $outfilePath -Append

       \\ITDSL\MSNPLAT\GOLD\MSNPATCH\MSNPATCH.exe /CORP | Out-File $outfilePath -Append

        }
    }
}