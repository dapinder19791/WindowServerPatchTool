$report =@()
 
[System.Collections.ArrayList]$SingleReport = @()


function check-PendingReboot
{
 if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { return $true }
 if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { return $true }
  
 try { 
   $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
   $status = $util.DetermineIfRebootPending()
   if(($status -ne $null) -and $status.RebootPending){
     return $true
   }
 }catch{}
 
 return $false
}



function get-lastinstalledDate
{
 $Retvalue=$false;
$Session = New-Object -ComObject Microsoft.Update.Session            
$Searcher = $Session.CreateUpdateSearcher()            
$HistoryCount = $Searcher.GetTotalHistoryCount()            
$history =  $Searcher.QueryHistory(0,$HistoryCount) | ForEach-Object -Process {            
    $Title = $null            
            
        $Title = $_.Title            
             
     $Result = $null            
    Switch ($_.ResultCode)            
    {            
        0 { $Result = 'NotStarted'}            
        1 { $Result = 'InProgress' }            
        2 { $Result = 'Succeeded' }            
        3 { $Result = 'SucceededWithErrors' }            
        4 { $Result = 'Failed' }            
        5 { $Result = 'Aborted' }            
        default { $Result = $_ }            
    }            
    New-Object -TypeName PSObject -Property @{            
        InstalledOn = Get-Date -Date $_.Date;            
        Title = $Title;            
        Name = $_.Title;            
        Status = $Result            
    }            
            
} | Sort-Object -Descending:$true -Property InstalledOn |             

Select-Object -Property * -ExcludeProperty Name  
 

$history | Select-Object -First 1| Sort-Object InstallOn –Descending

 
}
function getsqlversion
{
$version = Invoke-Sqlcmd -Query "SELECT @@VERSION;" -QueryTimeout 3
return $version
}

function checkSqlServerExists
{

Param ($computername)

$val=$false

$SQLInstances = Invoke-Command -ComputerName $computername -ErrorAction Ignore{
Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Microsoft SQL Server'
} 

$counter=@($SQLInstances).Count

if ($counter -gt 0) {
$val= $true
 } Else {
 $val= $false
 }
  

 return $val
}

 
function getSQLserverfrmRegistry{

param(
	  	$Computers,
	$Date = $(Get-Date )
	
)
 
$i = 1
$ret = @()
foreach ($server in $Computers)
{
	Write-Progress -Activity "Gathering information" -Status "Current Server $($Server.name) ($i of $($computers.count))" -PercentComplete (($i/$Computers.Count) * 100) -ID 1
	$i++
	
	if (Test-Connection -ComputerName $server -Count 1 -Quiet)
	{
		$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $server)
		$SqlKey = $regKey.OpenSubKey("SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL")
		Write-Verbose "Processing Server $server"
		if ($SqlKey -ne $null)
		{
			Foreach ($instance in $SqlKey.GetValueNames())
			{
				$SQLServerObject = New-Object PSObject
				$SQLServerObject | Add-Member -MemberType NoteProperty -Name ServerName -Value $server
				$InstanceName = $SqlKey.GetValue("$instance")
				Write-Progress -Activity "Parsing SQL-Server Information" -Status "Current Instance $instancename" -ParentID 1
				$InstanceKey = $regKey.OpenSubkey("SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceName\Setup")
				 
				$SQLServerObject | Add-Member -MemberType NoteProperty -Name InstanceName -Value $InstanceName
				
				$SQLServerObject | Add-Member -MemberType NoteProperty -Name Edition -Value $InstanceKey.GetValue("Edition")
				$SQLServerObject | Add-Member -MemberType NoteProperty -Name Version -Value $InstanceKey.GetValue("Version")
				 
				 
			}
			Write-Progress -Activity "Parsing SQL-Server Information" -Completed
			
		}
		else
		{
			Write-Warning -Message "The Server $server hasn´t SQL Server installed"
		}
	}
	else
	{
		Write-Error -Message "$server is not available"
	}
}

$retvalue = "Instance Name : "  + $SQLServerObject.InstanceName + " , Edition : " + $SQLServerObject.Edition  + " , Version : "  + $SQLServerObject.Version
return $retvalue

}

function checkSQLService
{
$servicestatus = Get-service  -Name 'MSSQLSERVER' 
return $servicestatus

}
function set_objects
{
Param ($Computer,$Title,$KB,$IsDownloaded,$Notes,$OSVersion,$SqlVersion,$LastScanTime,$LastRebootTime,$LastPatchInstalledTime)
$ListObject = [PSCustomObject]@{
            Computer       = $Computer
            Title        = $Title
            KB                  = $KB
            IsDownloaded        = $IsDownloaded
            Notes              = $Notes
            OSVersion          = $OSVersion
            SqlVersion         = $SqlVersion
            LastScanTime        = $LastScanTime
            LastRebootTime               = $LastRebootTime
            LastPatchInstalledTime = $LastPatchInstalledTime
               
        }
return $ListObject 

}
 
$Computer = $env:computername

try{


 
if (Test-Connection -ComputerName $Computer -Count 1 -Quiet){

$sOS=(Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Name
$sOS= $sOS.Substring(0,$sOS.IndexOf("|"))
$LASTREBOOT = Get-CimInstance -ComputerName $Computer -ClassName win32_operatingsystem | select csname, lastbootuptime
$lastboouptime=$LASTREBOOT.lastbootuptime
$date=Get-Date
$lastpatchinstalled=get-lastinstalledDate
$status=checkSqlServerExists -computername $Computer
$getversion="Not Installed"
$getservicestatus="SQL Server Stopped"

if($status -eq $true)
{
$getservicestatus = checkSQLService
$getservicestatus=$getservicestatus.Status

if($getservicestatus -eq "Stopped")
{
$getsqlregistry = getSQLserverfrmRegistry -Computers $Computer
$getversion = $getsqlregistry + ", Status : " + $getservicestatus


}
else{
$getversion = getsqlversion 
$getversion = $getversion.Column1 -split " "
$getversion = $getversion[0] + " " + $getversion[1]  + " " + $getversion[2] + " " + [string] $getversion[3] + " " + [string]$getversion[4]  + " " + [string]$getversion[5] + " " +  [string]$getversion[6] + [string]$getversion[7] + " " + [string]$getversion[8] 

}
}
## Log file path
$outfilePath =  "c:\$($Computer)_Updatespatchlog.csv"
 
if([System.IO.File]::Exists($outfilePath)){ 

  Remove-Item $outfilePath -force -recurse

 }
 

#check if the server is the same where this script is running
if($Computer -eq "$env:computername")
				{
					$UpdateSession = New-Object -ComObject Microsoft.Update.Session
				}
else { $UpdateSession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$Computer)) }

Write-Host("Searching for Any pending reboots...")   (Get-Date) -Fore Green
$checkreboot = check-PendingReboot 

if($checkreboot)
{             
               if($status -eq $false)
                    {
                    $report=set_objects -Computer $Computer -Title "RebootPending" -KB "NA" -IsDownloaded "NA" -SqlVersion "Not Installed"  -Notes "Installed and Reboot Pending" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
                    
                    }
                    else{
                    $report=set_objects -Computer $Computer -Title "RebootPending" -KB "NA" -IsDownloaded "NA" -SqlVersion $getversion  -Notes "Installed and Reboot Pending" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
                     
                    }
                    $SingleReport.Add($report) | Out-Null  
}
else{

write-host "Creating update searcher session at "  (Get-Date)
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$searchresult = $updatesearcher.Search("IsInstalled=0")    
        
                #Verify if Updates need installed
                Write-host "Verifing that updates are available to install"
                If ($searchresult.Updates.Count -gt 0) {
                    #Updates are waiting to be installed
                    Write-host "Found $($searchresult.Updates.Count) update\s!"
                    #Cache the count to make the For loop run faster
                    $count = $searchresult.Updates.Count
                    
                    #Begin iterating through Updates available for installation
                    Write-host "Iterating through list of updates"
                   $Export = @()
                    
                   
                    For ($i=0; $i -lt $Count; $i++) {
                        #Create object holding update
                        $Update = $searchresult.Updates.Item($i)
                        $kbid=$($Update.KBArticleIDs)
                        $title=$Update.Title

                        if($status -eq $false)
                        {
                        $report=set_objects -Computer $Computer -Title $title -KB $kbid -IsDownloaded "NA" -SqlVersion "Not Installed"  -Notes "There are total of $Count Updates Found in this machine." -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
                        
                         }
                     else{
                         $report=set_objects -Computer $Computer -Title $title -KB $kbid -IsDownloaded "NA" -SqlVersion $getversion  -Notes "There are total of $Count Updates Found in this machine." -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
                         
                         }   
            
                     $SingleReport.Add($report) | Out-Null
                       
                        }
                       
                          

                    }
                Else {
                    #Create Temp collection for report
                    Write-host "Creating report"
           if($status -eq $false)
                        {
                        $report=set_objects -Computer $Computer -Title "NoPendingUpdates" -SqlVersion "Not Installed" -KB "NA" -IsDownloaded "NA"   -Notes "There are no pending updates on this Machine." -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
                        }
                        else{
                        $report=set_objects -Computer $Computer -Title "NoPendingUpdates" -SqlVersion $getversion -KB "NA" -IsDownloaded "NA"   -Notes "There are no pending updates on this Machine." -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.InstalledOn 
                        
                        }
                        
                        $SingleReport.Add($report) | Out-Null
                         
                          
                    }              
                }
            
            }
        Else {
            #Nothing to install at this time
            Write-Warning "$($Computer): Offline"
             $report=set_objects -Computer $Computer -Title "Offline" -KB "NA" -IsDownloaded "NA" -SqlVersion "Offline"  -Notes "Server is Offline" -OSVersion "Offline" -LastScanTime "Offline" -LastRebootTime "Offline" -LastPatchInstalledTime "Offline"
             $SingleReport.Add($report) | Out-Null
             
             
        } 
 }catch [Exception]
 {
    #Create Temp collection for report
          
           $sOS=(Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Name
            $sOS= $sOS.Substring(0,$sOS.IndexOf("|"))
            $LASTREBOOT = Get-CimInstance -ComputerName $Computer -ClassName win32_operatingsystem | select csname, lastbootuptime
            $lastboouptime=$LASTREBOOT.lastbootuptime
            
          
            $report=set_objects -Computer $Computer -Title "ERROR:$_.Exception.Message" -SqlVersion "Error" -KB "NA" -IsDownloaded "NA"   -Notes "ERROR:$_.Exception.Message" -OSVersion "Error" -LastScanTime "Error" -LastRebootTime "Error" -LastPatchInstalledTime "Error" 
            
            $SingleReport.Add($report) | Out-Null
 }

 
 if($SingleReport.Count -gt 0)
 {
 $SingleReport|Export-CSV $outfilePath -noTypeInformation -Append
 }