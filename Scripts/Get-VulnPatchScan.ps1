$report =@()
[System.Collections.ArrayList]$ReportArray = @()
[System.Collections.ArrayList]$SingleReport = @()

function check-registry
{
 $value=$false

 
$listRegistry = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\" | Select-Object -ExpandProperty Property
if($listRegistry.Contains("FeatureSettingsOverride") -or $listRegistry.Contains("FeatureSettingsOverrideMask"))
{
return $true
}

 
 return $value
}


function check-history
{
$Retvalue=$false;
$Session = New-Object -ComObject Microsoft.Update.Session            
$Searcher = $Session.CreateUpdateSearcher()            
$HistoryCount = $Searcher.GetTotalHistoryCount()            
# http://msdn.microsoft.com/en-us/library/windows/desktop/aa386532%28v=vs.85%29.aspx            
$history =  $Searcher.QueryHistory(0,$HistoryCount) | ForEach-Object -Process {            
    $Title = $null            
            
        $Title = $_.Title            
             
    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa387095%28v=vs.85%29.aspx            
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

$ifexists1 = $history |  Select-String -Pattern "KB4056898" -CaseSensitive
$ifexists2 = $history |  Select-String -Pattern "KB4054519" -CaseSensitive
$ifexists3 = $history |  Select-String -Pattern "KB4056890" -CaseSensitive
$ifexists4 = $history |  Select-String -Pattern "KB4056897" -CaseSensitive




if(($ifexists1.Count -gt 0) -or ($ifexists2.Count -gt 0) -or ($ifexists3.Count -gt 0) -or ($ifexists4.Count -gt 0))
{
$Retvalue=$true
}

return $Retvalue

}


function get-lastinstalledDate
{
 $Retvalue=$false;
$Session = New-Object -ComObject Microsoft.Update.Session            
$Searcher = $Session.CreateUpdateSearcher()            
$HistoryCount = $Searcher.GetTotalHistoryCount()            
# http://msdn.microsoft.com/en-us/library/windows/desktop/aa386532%28v=vs.85%29.aspx            
$history =  $Searcher.QueryHistory(0,$HistoryCount) | ForEach-Object -Process {            
    $Title = $null            
            
        $Title = $_.Title            
             
    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa387095%28v=vs.85%29.aspx            
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

$ifexists1 = $history |  Select-String -Pattern "KB4056898" -CaseSensitive
$ifexists2 = $history |  Select-String -Pattern "KB4054519" -CaseSensitive
$ifexists3 = $history |  Select-String -Pattern "KB4056890" -CaseSensitive
$ifexists4 = $history |  Select-String -Pattern "KB4056897" -CaseSensitive

if($ifexists1.count -gt 0)
{
$ifexists1 | Select-Object -First 1| Sort-Object InstallOn –Descending
}
elseif($ifexists2.count -gt 0)
{
$ifexists2 | Select-Object -First 1| Sort-Object InstallOn –Descending
}
elseif($ifexists3.count -gt 0)
{
$ifexists3 | Select-Object -First 1| Sort-Object InstallOn –Descending
}
elseif($ifexists4.count -gt 0)
{
$ifexists4 | Select-Object -First 1| Sort-Object InstallOn –Descending
} 
else{
$history | Select-Object -First 1| Sort-Object InstallOn –Descending

}  
}

function set_objects
{
Param ($Computer,$Title,$KB,$IsDownloaded,$Notes,$OSVersion,$LastScanTime,$LastRebootTime,$LastPatchInstalledTime)
$ListObject = [PSCustomObject]@{
            Computer       = $Computer
            Title        = $Title
            KB                  = $KB
            IsDownloaded        = $IsDownloaded
            Notes              = $Notes
            OSVersion          = $OSVersion
            LastScanTime        = $LastScanTime
            LastRebootTime               = $LastRebootTime
            LastPatchInstalledTime = $LastPatchInstalledTime
               
        }
return $ListObject 

}

 
#$file = Get-Content -Path $outfileCommonPath
$Computer = $env:computername
try{
 
$sOS=(Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Name
$sOS= $sOS.Substring(0,$sOS.IndexOf("|"))
$LASTREBOOT = Get-CimInstance -ComputerName $Computer -ClassName win32_operatingsystem | select csname, lastbootuptime
$lastboouptime=$LASTREBOOT.lastbootuptime
$date=Get-Date

if (Test-Connection -ComputerName $Computer -Count 1 -Quiet){
                ## Log file path
                $outfilePath =  "C:\$($Computer)_Zeropatchlog.csv"
                if([System.IO.File]::Exists($outfilePath)){ 

  Remove-Item $outfilePath -force -recurse

 }
                 #check if the server is the same where this script is running
                 if($Computer -eq "$env:computername")
		           {
					$UpdateSession = New-Object -ComObject Microsoft.Update.Session
				}
                 else { 
                    $UpdateSession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$Computer)) 
            }
                 Write-Host("Searching for Zero Patch Registry...")   (Get-Date) -Fore Green
                 $checkregistry = check-registry 
                 $checkhistory=check-history
                if($checkregistry -or $checkhistory)
                {
                         
                        $key = “SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install” 
                        $keytype = [Microsoft.Win32.RegistryHive]::LocalMachine 
                        $RemoteBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($keytype,$Server) 
                        $regKey = $RemoteBase.OpenSubKey($key) 
                        

                if($checkregistry -and $checkhistory)
                {
                 if($regKey -ne $null){
                        $KeyValue = $regkey.GetValue(”LastSuccessTime”) 
                        }
                        else{
                        $lastinstalltimeinfo=get-lastinstalledDate 
                        $KeyValue=$lastinstalltimeinfo.InstalledOn
                        }
                 
                $report=set_objects -Computer $Computer -Title "Installed" -KB "NA" -IsDownloaded "NA"  -Notes "Registry and Window Patch exist" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $KeyValue 
                $SingleReport.Add($report) | Out-Null
                $ReportArray.Add($report) | Out-Null

                }
                elseif($checkregistry -and !$checkhistory)
                {
                 if($regKey -ne $null){
                        $KeyValue = $regkey.GetValue(”LastSuccessTime”) 
                        }
                 
                $report=set_objects -Computer $Computer -Title "Installed" -KB "NA" -IsDownloaded "NA"   -Notes "Registry Exist,But Window Patch Does not exist" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $KeyValue 
                $SingleReport.Add($report) | Out-Null
                $ReportArray.Add($report) | Out-Null  
                }
                elseif(!$checkregistry -and $checkhistory)
                {
                $lastinstallfunc=get-lastinstalledDate
                $lastinstall=$lastinstallfunc.InstalledOn
                $SingleReport.Add($report) | Out-Null
                $report=set_objects -Computer $Computer -Title "Installed" -KB "NA" -IsDownloaded "NA"   -Notes "Window Patch Exists,But Registry Key is Missing" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastinstall 
                $ReportArray.Add($report) | Out-Null
                
                }
                 
                }
                else{
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                else{ 
                    write-host "The Registry and Windows History for  KB4056898 and KB4054519 does not exists"
                    write-host "Searching for KB ID KB4056898 and KB4054519 for Zero Patch Vulnerability"  (Get-Date)
        
                    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
                    $searchresult = $updatesearcher.Search("IsInstalled=0 AND Type='Software'")    
        
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
                       $checkexists=$false

                        For ($i=0; $i -lt $Count; $i++) {
                            #Create object holding update
                            $update = $searchresult.Updates.Item($i)
                            
                                 if($update.Title.Contains("4056898") -OR $update.Title.Contains("4056897")-OR $update.Title.Contains("4056890")-OR $update.Title.Contains("KB4054519"))
                            {
                                 
                                $checkexists=$true 
                                
                            }
                         }
                              
                             if($checkexists)
                             {
                             $report=set_objects -Computer $Computer -Title "NotInstalled" -KB "NA" -IsDownloaded "NA"   -Notes "$TITLE  - Patch available to be downloaded" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $null 
                             $ReportArray.Add($report) | Out-Null
                             $SingleReport.Add($report) | Out-Null
                              }
                             else{
                             $report=set_objects -Computer $Computer -Title "NotInstalled" -KB "NA" -IsDownloaded "NA"   -Notes "$TITLE  - Patch not available in windows updates" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $null 
                             $ReportArray.Add($report) | Out-Null
                             $SingleReport.Add($report) | Out-Null
                             }

                    }
                Else {
                          
                            $report=set_objects -Computer $Computer -Title "NotInstalled" -KB "NA" -IsDownloaded "NA"   -Notes "Patch is not available for Window Updates" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $null 
                            $ReportArray.Add($report) | Out-Null
                            $SingleReport.Add($report) | Out-Null
 
                    }              
                }
            
                }
            }
            else {
             
                            
                            $report=set_objects -Computer $Computer -Title "Offline" -KB "NA" -IsDownloaded "NA"   -Notes "Server is Offline" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $null 
                            $ReportArray.Add($report) | Out-Null
                            $SingleReport.Add($report) | Out-Null
                             
        

                            
            }
 
 }
catch [Exception]
 {
    
            $sOS=(Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Name
            $sOS= $sOS.Substring(0,$sOS.IndexOf("|"))
            $LASTREBOOT = Get-CimInstance -ComputerName $Computer -ClassName win32_operatingsystem | select csname, lastbootuptime
            $lastboouptime=$LASTREBOOT.lastbootuptime
            
          
                            $report=set_objects -Computer $Computer -Title "ERROR:$_.Exception.Message" -KB "NA" -IsDownloaded "NA"   -Notes "ERROR:$_.Exception.Message" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $null 
                            $ReportArray.Add($report) | Out-Null
                            $SingleReport.Add($report) | Out-Null
 
 }
 
 if($SingleReport.Count -gt 0)
 {
 $SingleReport|Export-CSV $outfilePath -noTypeInformation -Append
 }