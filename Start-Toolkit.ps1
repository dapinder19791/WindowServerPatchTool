#region Synchronized Collections
$uiHash = [hashtable]::Synchronized(@{})

$runspaceHash = [hashtable]::Synchronized(@{})
$jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$OutputArray = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))

$jobCleanup = [hashtable]::Synchronized(@{})
$Global:updateAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

$Global:KBAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

$Global:updateVulnAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:installAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:servicesAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:servicesAuditStart = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))


$Global:SnoozeAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:UnSnoozeAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))



$Global:JobsAudit = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$Global:installedUpdates = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
#endregion

#region Startup Checks and configurations
#Determine if running from ISE
Write-Verbose "Checking to see if running from console"
If ($Host.name -eq "Windows PowerShell ISE Host") {
    Write-Warning "Unable to run this from the PowerShell ISE due to issues with PSexec!`nPlease run from console."
    Break
}

#Validate user is an Administrator
Write-Verbose "Checking Administrator credentials"
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You are not running this as an Administrator!`nRe-running script and will prompt for administrator credentials."
    Start-Process -Verb "Runas" -File PowerShell.exe -Argument "-STA -noprofile -file $($myinvocation.mycommand.definition)"
    Break
}

#Ensure that we are running the GUI from the correct location
Set-Location $(Split-Path $MyInvocation.MyCommand.Path)
$Global:Path = $(Split-Path $MyInvocation.MyCommand.Path)
Write-Debug "Current location: $Path"

#Check for PSExec
Write-Verbose "Checking for psexec.exe"
If (-Not (Test-Path C:\psscripts\PsExec.exe)) {
    Write-Warning ("Psexec.exe missing from {0}!`n Please place file in the path so UI can work properly" -f (Split-Path $MyInvocation.MyCommand.Path))
    Break
}

#Determine if this instance of PowerShell can run WPF 
Write-Verbose "Checking the apartment state"
If ($host.Runspace.ApartmentState -ne "STA") {
    Write-Warning "This script must be run in PowerShell started using -STA switch!`nScript will attempt to open PowerShell in STA and run re-run script."
    Start-Process -File PowerShell.exe -Argument "-STA -noprofile -WindowStyle hidden -file $($myinvocation.mycommand.definition)"
    Break
}

#Load Required Assemblies
Add-Type –assemblyName PresentationFramework
Add-Type –assemblyName PresentationCore
Add-Type –assemblyName WindowsBase
Add-Type –assemblyName Microsoft.VisualBasic
Add-Type –assemblyName System.Windows.Forms

#DotSource Help script
. ".\HelpFiles\HelpOverview.ps1"

#DotSource About script
. ".\HelpFiles\About.ps1"
#endregion

Function Set-PoshPAIGOption {
    If (Test-Path (Join-Path $Path 'options.xml')) {
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        $Global:maxConcurrentJobs = $Optionshash['MaxJobs']
        $Global:MaxRebootJobs = $Optionshash['MaxRebootJobs']
     
        If ($Optionshash['ReportPath']) {
            
            If (Test-Path ($Optionshash['ReportPath']))
            {
            $Global:reportpath = $Optionshash['ReportPath']
            }
            else{
             $Optionshash['ReportPath'] = $Global:reportpath = (Join-Path $Home 'Desktop')
              $Global:reportpath = (Join-Path $env:USERPROFILE 'Desktop')
            }
    
        } Else {
            $Optionshash['ReportPath'] = $Global:reportpath = (Join-Path $Home 'Desktop')
        }

        If ($Optionshash['DiagPath']) {
            
            If (Test-Path ($Optionshash['DiagPath']))
            {
            $Global:DiagPath = $Optionshash['DiagPath']
            }
            else{
             $Optionshash['DiagPath'] = $Global:DiagPath = (Join-Path $Home 'Desktop')
              $Global:DiagPath = (Join-Path $env:USERPROFILE 'Desktop')
            }
    
        } Else {
            $Optionshash['DiagPath'] = $Global:DiagPath = (Join-Path $Home 'Desktop')
        }


    }
     Else {
        #Default Options
        $optionshash = @{
            MaxJobs = 5
            MaxRebootJobs = 5
            ReportPath = (Join-Path $env:USERPROFILE 'Desktop')
            DiagPath= (Join-Path $env:USERPROFILE 'Desktop')
        }


        $Global:maxConcurrentJobs = 5
        $Global:MaxRebootJobs = 5
        $Global:reportpath = (Join-Path $env:USERPROFILE 'Desktop')
        $Global:DiagPath= (Join-Path $env:USERPROFILE 'Desktop')
    }
    $optionshash | Export-Clixml -Path (Join-Path $pwd 'options.xml') -Force
}

#Function for Debug output
Function Global:Show-DebugState {
    Write-Debug ("Number of Items: {0}" -f $uiHash.Listview.ItemsSource.count)
    Write-Debug ("First Item: {0}" -f $uiHash.Listview.ItemsSource[0].Computer)
    Write-Debug ("Last Item: {0}" -f $uiHash.Listview.ItemsSource[$($uiHash.Listview.ItemsSource.count) -1].Computer)
    Write-Debug ("Max Progress Bar: {0}" -f $uiHash.ProgressBar.Maximum)
}

#Reboot Warning Message
Function Show-RebootWarning {
    $title = "Reboot Server Warning"
    $message = "You are about to reboot servers which can affect the environment! `nAre you sure you want to do this?"
    $button = [System.Windows.Forms.MessageBoxButtons]::YesNo
    $icon = [Windows.Forms.MessageBoxIcon]::Warning
    [windows.forms.messagebox]::Show($message,$title,$button,$icon)
}

#Format and display errors
Function Get-Error {
    Process {
        ForEach ($err in $error) {
            Switch ($err) {
                {$err -is [System.Management.Automation.ErrorRecord]} {
                        $hash = @{
                        Category = $err.categoryinfo.Category
                        Activity = $err.categoryinfo.Activity
                        Reason = $err.categoryinfo.Reason
                        Type = $err.GetType().ToString()
                        Exception = ($err.exception -split ": ")[1]
                        QualifiedError = $err.FullyQualifiedErrorId
                        CharacterNumber = $err.InvocationInfo.OffsetInLine
                        LineNumber = $err.InvocationInfo.ScriptLineNumber
                        Line = $err.InvocationInfo.Line
                        TargetObject = $err.TargetObject
                        }
                    }               
                Default {
                    $hash = @{
                        Category = $err.errorrecord.categoryinfo.category
                        Activity = $err.errorrecord.categoryinfo.Activity
                        Reason = $err.errorrecord.categoryinfo.Reason
                        Type = $err.GetType().ToString()
                        Exception = ($err.errorrecord.exception -split ": ")[1]
                        QualifiedError = $err.errorrecord.FullyQualifiedErrorId
                        CharacterNumber = $err.errorrecord.InvocationInfo.OffsetInLine
                        LineNumber = $err.errorrecord.InvocationInfo.ScriptLineNumber
                        Line = $err.errorrecord.InvocationInfo.Line                    
                        TargetObject = $err.errorrecord.TargetObject
                    }               
                }                        
            }
        $object = New-Object PSObject -Property $hash
        $object.PSTypeNames.Insert(0,'ErrorInformation')
        $object
        }
    }
}

#Add new server to GUI
Function Add-Server {
    $computer = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a server name or names. Separate servers with a comma (,) or semi-colon (;).", "Add Server/s")
    If (-Not [System.String]::IsNullOrEmpty($computer)) {
        [string[]]$computername = $computer -split ",|;"
        ForEach ($computer in $computername) { 
            If (-NOT [System.String]::IsNullOrEmpty($computer)) {
                $clientObservable.Add((
                    New-Object PSObject -Property @{
                        Computer = ($computer).Trim()
                        Audited = 0 -as [int]
                        Installed = 0 -as [int]
                        InstallErrors = 0 -as [int]
                        Offline=0 -as [int]
                        UnAuthorized=0 -as [int]
                        LastRebootTime="" -as [string]
                        LastInstalledTime="" -as [string]
                        OsVersion="" -as [string]
                        SqlVersion = "" -as [string]
                        Services = 0 -as [int]
                        SQLJobs= 0 -as [int]
                        StatusTxt="" -as [string]
                        Notes = $Null
                        Status=$Null
                    }
                ))     
                Show-DebugState
            }
        }
    } 
}

#Remove server from GUI
Function Remove-Server {
    $Servers = @($uiHash.Listview.SelectedItems)
    ForEach ($server in $servers) {
        $clientObservable.Remove($server)
    }
    $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
    Show-DebugState  
}
 



#Report Generation function
Function Start-Report {
    Write-Debug ("Data: {0}" -f $uiHash.ReportComboBox.SelectedItem.Text)
    Switch ($uiHash.ReportComboBox.SelectedItem.Text) {
         
        "Check Window Updates - Report" {
            
            If ($updateAudit.count -gt 0) {
                    $updateAudit | Where {$_.Notes -ne $null} |  Select Computer,Title,KB,IsDownloaded,Offline,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion | Out-GridView -Title 'Check Window Updates - Report'
            
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }
        }
        "Check Zero Vulnerability - Report" {
            
            If ($updateVulnAudit.count -gt 0) {
                    $updateVulnAudit | Where {$_.Notes -ne $null} | Select Computer,Title,KB,IsDownloaded,Offline,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion | Out-GridView -Title 'Check Zero Vulnerability - Report'
            
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }
        }

        "Search Installed KBs - Report"{

         If ($KBAudit.count -gt 0) {
                    $KBAudit | Where {$_.Notes -ne $null} | Select Computer,Title,KB,IsDownloaded,Offline,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion | Out-GridView -Title 'KB Detail - Report'
            
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }
             

        }
        


        "Install Window Patches - Report" {
            If ($installAudit.count -gt 0) {
            $installAudit | Where {$_.Notes -ne $null} | Select Computer,Title,KB,IsDownloaded,Offline,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion | Out-GridView -Title 'Install Window Patches - Report'
             
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }        
        }
        "Install SQL Patches - Report"{
         If ($installAudit.count -gt 0) {
            $installAudit | Where {$_.Notes -ne $null} | Select Computer,Title,KB,IsDownloaded,Offline,LastRebootTime,LastPatchInstalledTime,Notes,OSVersion | Out-GridView -Title 'Install SQL Patches - Report'
             
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            } 
        }
        "Install Stay Current Patches - Report"{}
        "Reboot Server - Report"{}
        "Snooze Server (Classic) - Report"{
           If ($SnoozeAudit.count -gt 0) {
                $SnoozeAudit | Out-GridView -Title 'Snooze Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }        
        }
        "Un-Snooze Server (Classic) - Report"{
           If ($UnSnoozeAudit.count -gt 0) {
                $UnSnoozeAudit | Out-GridView -Title 'Un-Snooze Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"         
            }        
        }
        "Service Check (Stopped) - Report" {
           
            If (@($servicesAudit).count -gt 0) {
                $servicesAudit | Select @{L='Computername';E={$_.__Server}},Name,DisplayName,State,StartMode,ExitCode,Status | Out-GridView -Title 'Services Stopped Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"             
            }
        }
        "Service Action (Start) - Report" {
           
            If (@($servicesAuditStart).count -gt 0) {
                $servicesAuditStart | Select @{L='Computername';E={$_.__Server}},Name,DisplayName,State,StartMode,ExitCode,Status | Out-GridView -Title 'Services Start Report'
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"             
            }
        }
        "SQL Jobs Check - Report" {
          
            If (@($JobsAudit).count -gt 0) {
              $JobsAudit |select Computer,name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,Notes | Out-GridView -Wait -Title 'Job Report'
              
            } Else {
                $uiHash.StatusTextBox.Foreground = "Red"
                $uiHash.StatusTextBox.Text = "No report to create!"             
            }
        }

         
    }
}

 
Function generateExcelReport{

 If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}
         }

      $report = $reportpath + "\" + "serverlist.csv"
      $savedreport =  $report

     
     #$savedreport = Join-Path (Join-Path $home Desktop) "serverlist.csv"
     write-host "Checking Reports $savedreport"

        if([System.IO.File]::Exists($savedreport)){ 

          Remove-Item $savedreport -force -recurse
          Write-Host "Removing Old Reports"

         }

         Write-Host "Creating Latest Reports"

        $uiHash.Listview.ItemsSource | Export-Csv -NoTypeInformation $savedreport
}

#start-RunJob function
Function Start-RunJob {    
    Write-Host ("ComboBox {0}" -f $uiHash.RunOptionComboBox.Text)
    Write-Debug ("TextBox {0}" -f $uiHash.ServicesInputBox.Text)

    $selectedItems = $uiHash.Listview.SelectedItems
    $global:ServiceText = $uiHash.ServicesInputBox.Text

    If ($selectedItems.Count -gt 0) {
        $uiHash.ProgressBar.Maximum = $selectedItems.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}        
        If ($uiHash.RunOptionComboBox.Text -eq 'Install Window Patches') {             
            #region Install Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Installing Window Patches for all servers...Please Wait"              
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $installAudit,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Installing Patches"
                    $Computer.colorchange = "Black" 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  

                

                Set-Location -Path $Path 
                 . .\Scripts\Get-Executor.ps1  
                 
                 $Computer.Status= "$pwd\Images\progress.png"
                
                #$installAudit.AddRange($clientInstall) | Out-Null

                try{
                $clientInstall = @(Getupdate-Patches -Computer $computer.computer -batfilename "InstallWindowsUpdates.bat" -Psfilename "InstallWindowsUpdates.ps1" -log "patchlog.csv")
                
                $clientInstalledCount =  @($clientInstall |  Select-String -Pattern "InstallationCompleted" -CaseSensitive).Count
                $clientInstalledErrorCount = @($clientInstall | Select-String -Pattern "ERROR" -CaseSensitive).Count
                $clientUpdateunauthorized =  @($clientInstall |  Select-String -Pattern "UnAuthorizedAccess" -CaseSensitive).Count
                $clientUpdateNotFound =  @($clientInstall |  Select-String -Pattern "NoPendingUpdates" -CaseSensitive).Count
                $clientoffline =  @($clientInstall |  Select-String -Pattern "Offline" -CaseSensitive).Count
                $clientRebootPending=  @($clientInstall |  Select-String -Pattern "RebootPending" -CaseSensitive).Count 
                }
             Catch {
                        $queryError = $_.Exception.Message
                    }
                  
                 If (Test-Connection -ComputerName $computer.computer -Count 1 -Quiet) {
                $sOS=(Get-WmiObject -ComputerName $computer.computer -Class Win32_OperatingSystem).Name
                $LASTREBOOT = Get-CimInstance -ComputerName $computer.computer -ClassName win32_operatingsystem | select csname, lastbootuptime
                
                $lastboouptime=$LASTREBOOT.lastbootuptime
                }
                $date=Get-Date

                 
                 If ($clientInstall.count -gt 0) {
                     $installAudit.AddRange($clientInstall) | Out-Null
                }


                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    

                    $uiHash.Listview.Items.EditItem($Computer)
                    if($clientInstalledCount -gt 0)
                    { 
                     
                    $Computer.Notes = "Installation is Completed Successfully, Please roboot the server!"  
                    $Computer.Audited= $clientInstalledCount.count
                    $Computer.Installed= $clientInstalledCount.count
                    if($clientInstall.LastRebootTime.count -gt 1)
                    {
                    $Computer.LastRebootTime= $clientInstall.LastRebootTime[0]
                    }
                    else{
                    $Computer.LastRebootTime= $clientInstall.LastRebootTime
                    }
                    if($clientInstall.LastPatchInstalledTime.count -gt 1)
                    {
                    $Computer.LastPatchInstalledTime= $clientInstall.LastPatchInstalledTime[0]
                    }
                    else{
                    $Computer.LastPatchInstalledTime= $clientInstall.LastPatchInstalledTime
                    }
                    
                     
                    $Computer.SqlVersion ="NA"
                    $Computer.UnAuthorized=0
                    $Computer.StatusTxt="Installation Completed"
                    $Computer.Offline= 0
                    if($clientInstall.OSVersion.count -gt 0)
                    {
                    $Computer.OSVersion = $clientInstall.OSVersion[0]
                    
                    }
                    else{
                    $Computer.OSVersion = $clientInstall.OSVersion
                    
                    }
                    $Computer.Status= "$pwd\Images\found.png"
                        
                    }
                    if($clientRebootPending -gt 0)
                    {
                    $Computer.Notes = "This server is Pending Reboot!" 
                    $Computer.Audited= $clientRebootPending.count
                    $Computer.Installed= $clientRebootPending.count
                    $Computer.LastRebootTime= $clientInstall.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientInstall.LastPatchInstalledTime
                    $Computer.SqlVersion ="NA" 
                    
                    $Computer.OSVersion = $clientInstall.OSVersion
                    $Computer.UnAuthorized=0
                    $Computer.StatusTxt="Reboot Required"
                    $Computer.Offline= 0
                    $Computer.Status= "$pwd\Images\restart.png"

                    
                    }
                    if($clientUpdateNotFound -gt 0)
                    {
                    $Computer.Notes = "There are no new updates available in this machine!"
                    $Computer.Audited= $clientUpdateNotFound.count
                    $Computer.Audited= $clientUpdateNotFound.count
                    $Computer.Installed= $clientUpdateNotFound.count
                    $Computer.LastRebootTime= $clientInstall.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientInstall.LastPatchInstalledTime
                    $Computer.SqlVersion ="NA"
                    
                    $Computer.OSVersion = $clientInstall.OSVersion
                    $Computer.UnAuthorized=0
                    $Computer.StatusTxt="No Update Found"
                    $Computer.Offline= 0
                    $Computer.Status= "$pwd\Images\ok.png"
                    }

                    if($clientUpdateunauthorized -gt 0)
                    {
                    $Computer.Notes = "Un-Authorized Access!"
                    $Computer.Offline= 0
                    $Computer.Audited= 1
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=$clientUpdateunauthorized.count
                    $Computer.LastRebootTime= "UnAuthorizedAccess"
                    $Computer.OSVersion = "UnAuthorizedAccess"
                    $Computer.SqlVersion ="UnAuthorizedAccess"
                    
                    $Computer.StatusTxt="No Access"
                    $Computer.LastPatchInstalledTime= "UnAuthorizedAccess"
                    $Computer.Status= "$pwd\Images\noaccess.png"
                    
                    }

                    if($clientoffline -gt 0)
                    {
                    $Computer.Notes="Offline"
                    $Computer.colorchange = "Red"
                    $Computer.Offline=$clientoffline.Count
                    $Computer.Audited= 0
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=0
                    $Computer.LastRebootTime= "Offline"
                    $Computer.LastPatchInstalledTime= "Offline"
                    $Computer.StatusTxt="Server Offline"
                    $Computer.OSVersion = "Offline"
                     $Computer.SqlVersion ="Offline"
                    $Computer.Status= "$pwd\Images\offline.png"
                    }

                    If ($queryError) {
                     $Computer.colorchange = "Red"  
                        $Computer.notes = $queryError
                    } Else {
                    
                    if($clientInstalledErrorCount -gt 0)
                     {
                     $notes=$clientInstall.Notes
                     $Computer.notes = "$notes"
                     $Computer.Offline= 0
                     $Computer.Audited= 1
                     $Computer.Installed= 0
                     $Computer.UnAuthorized=0
                     $Computer.InstallErrors= $clientInstalledErrorCount.count
                     $Computer.LastRebootTime= "Error"
                     $Computer.StatusTxt="Error Occured"
                     $Computer.SqlVersion ="Error"
                     $Computer.LastPatchInstalledTime= "Error"
                     $Computer.OSVersion = "Error"
                     $Computer.Status= "$pwd\Images\error.png"
 
                       }                    
                       
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })

                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++  
                })

                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"  
                                    
                    }
                         
                })                  
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()  

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Patch Install"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($installAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion  
        }   
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Check Windows Updates') {
            #region Audit Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Auditing Patches for all servers...Please Wait"    
            $updateInfo = $uiHash.StatusTextBox.Text
            $Global:updatelayout = [Windows.Input.InputEventHandler]{ $uiHash.ProgressBar.UpdateLayout() }
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $updateAudit,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Auditing Patches"
                    $Computer.colorchange = "Black" 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  

                

             
                Set-Location $Path
                . .\Scripts\Get-Executor.ps1         
                
                 $Computer.Status= "$pwd\Images\progress.png"

                     Try {
                
              
                $clientUpdate = @(Getupdate-Patches -Computer $Computer.computer -batfilename "Get-PendingUpdates.bat" -Psfilename "Get-PendingUpdates.ps1" -log "Updatespatchlog.csv")
                 $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                $uiHash.StatusTextBox.Text=$updateInfo + " Scanning Path for server " + $Computer.computer 
                })

                $clientUpdateError =  @($clientUpdate |  Select-String -Pattern "ERROR" -CaseSensitive).Count
                $clientUpdateunauthorized =  @($clientUpdate |  Select-String -Pattern "UnAuthorizedAccess" -CaseSensitive).Count
                $clientUpdateFound =  @($clientUpdate |  Select-String -Pattern "Updates Found" -CaseSensitive).Count
                $clientUpdateNotFound =  @($clientUpdate |  Select-String -Pattern "NoPendingUpdates" -CaseSensitive).Count
                $clientoffline =  @($clientUpdate |  Select-String -Pattern "Offline" -CaseSensitive).Count

                
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                $uiHash.StatusTextBox.Text=$updateInfo + " Validation System Reboot " + $Computer.computer 
                })

                $clientRebootPending=  @($clientUpdate |  Select-String -Pattern "RebootPending" -CaseSensitive).Count 
 
                }
                 Catch {
                        $queryError = $_.Exception.Message
                    }
                  
                
                If (Test-Connection -ComputerName $computer.computer -Count 1 -Quiet) {
                $sOS=(Get-WmiObject -ComputerName $computer.computer -Class Win32_OperatingSystem).Name
                $LASTREBOOT = Get-CimInstance -ComputerName $computer.computer -ClassName win32_operatingsystem | select csname, lastbootuptime
                
                $lastboouptime=$LASTREBOOT.lastbootuptime
                }
                $date=Get-Date

                 
                 If ($clientUpdate.count -gt 0) {
                     $updateAudit.AddRange($clientUpdate) | Out-Null
                }

               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    

                    $uiHash.Listview.Items.EditItem($Computer)
                    if($clientUpdateFound -gt 0)
                    { 
                    if($clientUpdate.Notes.count -gt 1)
                    {
                    $Computer.Notes = $clientUpdate.Notes[0]
                    }
                    else{
                    $Computer.Notes = $clientUpdate.Notes 
                    }
                     
                    $Computer.Audited= $clientUpdateFound.count
                    $Computer.Installed= $clientUpdateFound.count
                  
                    if($clientUpdate.LastRebootTime.count -gt 1)
                    {
                    $Computer.LastRebootTime= $clientUpdate.LastRebootTime[0]
                    }
                    else{
                    $Computer.LastRebootTime= $clientUpdate.LastRebootTime
                    }
                    
                    if($clientUpdate.LastPatchInstalledTime.count -gt 1)
                    {
                    $Computer.LastPatchInstalledTime= $clientUpdate.LastPatchInstalledTime[0]
                    }
                    else{
                    $Computer.LastPatchInstalledTime= $clientUpdate.LastPatchInstalledTime
                    }

                    if($clientUpdate.OSVersion.count -gt 1)
                    {
                    $Computer.OSVersion = $clientUpdate.OSVersion[0]
                    
                    }
                    else{
                    $Computer.OSVersion = $clientUpdate.OSVersion
                    
                    }
                     if($clientUpdate.SqlVersion.count -gt 1)
                    {
                    $Computer.SqlVersion = $clientUpdate.SqlVersion[0]
                    
                    }
                    else{
                    $Computer.SqlVersion = $clientUpdate.SqlVersion
                    
                    }
                     
                    
                    $Computer.UnAuthorized=0
                    $Computer.Offline= 0
                    $Computer.StatusTxt="Updates Found"

                    $Computer.Status= "$pwd\Images\found.png"
                        
                    }
                    if($clientRebootPending -gt 0)
                    {
                     $Computer.Notes = "Updates are installed in this machine, Pending Reboot!" 
                    $Computer.Audited= $clientRebootPending.count
                    $Computer.Installed= $clientRebootPending.count
                    $Computer.LastRebootTime= $clientUpdate.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientUpdate.LastPatchInstalledTime
                    $Computer.OSVersion = $clientUpdate.OSVersion
                    $Computer.UnAuthorized=0
                    $Computer.Offline= 0
                    $Computer.SqlVersion = $clientUpdate.SqlVersion
                     $Computer.StatusTxt="Reboot Required"
                    $Computer.Status= "$pwd\Images\restart.png"

                    
                    }
                    if($clientUpdateNotFound -gt 0)
                    {
                    $Computer.Notes = "There are no new updates available in this machine!"
                    $Computer.Audited= $clientUpdateNotFound.count
                    $Computer.Audited= $clientUpdateNotFound.count
                    $Computer.Installed= $clientUpdateNotFound.count
                    $Computer.LastRebootTime= $clientUpdate.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientUpdate.LastPatchInstalledTime
                    $Computer.OSVersion = $clientUpdate.OSVersion
                    $Computer.UnAuthorized=0
                    $Computer.StatusTxt="No Update Found"
                    $Computer.SqlVersion = $clientUpdate.SqlVersion
                    $Computer.Offline= 0
                    $Computer.Status= "$pwd\Images\ok.png"
                    }

                    if($clientUpdateunauthorized -gt 0)
                    {
                    $Computer.Notes = "Un-Authorized Access!"
                    $Computer.Offline= 0
                    $Computer.Audited= 1
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=$clientUpdateunauthorized.count
                    $Computer.OSVersion = "UnAuthorizedAccess"
                    $Computer.LastRebootTime= "UnAuthorizedAccess"
                    $Computer.StatusTxt="No Access"
                    $Computer.SqlVersion = "UnAuthorizedAccess"
                    $Computer.LastPatchInstalledTime= "UnAuthorizedAccess"
                    $Computer.Status= "$pwd\Images\noaccess.png"
                    
                    }

                    if($clientoffline -gt 0)
                    {
                    $Computer.Notes="Offline"
                    $Computer.colorchange = "Red"
                    $Computer.Offline=$clientoffline.Count
                    $Computer.Audited= 0
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=0
                    $Computer.LastRebootTime= "Offline"
                    $Computer.LastPatchInstalledTime= "Offline"
                    $Computer.OSVersion = "Offline"
                    $Computer.SqlVersion = "Offline"
                    $Computer.StatusTxt="Server Offline"
                    $Computer.Status= "$pwd\Images\offline.png"
                    }

                    If ($queryError) {
                     $Computer.colorchange = "Red"  
                        $Computer.notes = $queryError
                    } Else {
                    
                    if($clientUpdateError -gt 0)
                     {
                     $notes=$clientUpdate.Notes
                     $Computer.notes = "$notes"
                     $Computer.Offline= 0
                     $Computer.Audited= 1
                     $Computer.Installed= 0
                     $Computer.UnAuthorized=0
                     $Computer.InstallErrors= $clientUpdateError.count
                     $Computer.LastRebootTime= "Error"
                     $Computer.StatusTxt="Error Occured"
                     $Computer.SqlVersion = "Error"
                     $Computer.LastPatchInstalledTime= "Error"
                     $Computer.OSVersion = "Error"
                     $Computer.Status= "$pwd\Images\error.png"
 
                       }                    
                       
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })

                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                $uiHash.StatusTextBox.Text=$updateInfo + " Scanning Completed on server " + $Computer.computer 
                })

                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                                     
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"   
                                      
                    }    
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Patch Audit"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($updateAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion                 


        } 
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Check Zero Vulnerability') {
            #region Audit Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Auditing Patches for Zero Vulnerability on all servers...Please Wait"            
            $Global:updatelayout = [Windows.Input.InputEventHandler]{ $uiHash.ProgressBar.UpdateLayout() }
            $uiHash.StartTime = (Get-Date)
            
             [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $updateVulnAudit,
                    $uiHash
                )
                 $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Auditing Zero Vuln Patches"
                    $Computer.colorchange = "Black" 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  


                              
                [System.Collections.ArrayList]$ReportArray = @()
                Set-Location $Path
                . .\Scripts\Get-Executor.ps1   
                
                $Computer.Status= "$pwd\Images\progress.png"

               # $outfileCommonPath =  "\\azmarprd06\ZeroVulnPatch\CommonZeropatchlog.csv"
                

              

                Try {
                       
                $clientUpdate =@(Getupdate-Patches -Computer $Computer.computer -batfilename "Get-VulnPatchScan.bat" -Psfilename "Get-VulnPatchScan.ps1" -log "Zeropatchlog.csv")
                $clientUpdateError =  @($clientUpdate |  Select-String -Pattern "ERROR" -CaseSensitive).Count
                $clientUpdateunauthorized =  @($clientUpdate |  Select-String -Pattern "UnAuthorizedAccess" -CaseSensitive).Count
                $clientUpdateInstalled =  @($clientUpdate |  Select-String -Pattern "Installed" -CaseSensitive).Count
                $clientUpdateNotInstalled =  @($clientUpdate |  Select-String -Pattern "NotInstalled" -CaseSensitive).Count
                $clientoffline =  @($clientUpdate |  Select-String -Pattern "Offline" -CaseSensitive).Count
                  
                }
                 Catch {
                        $queryError = $_.Exception.Message
                    }
                 
                If (Test-Connection -ComputerName $computer.computer -Count 1 -Quiet) {
                $sOS=(Get-WmiObject -ComputerName $computer.computer -Class Win32_OperatingSystem).Name
                $LASTREBOOT = Get-CimInstance -ComputerName $computer.computer -ClassName win32_operatingsystem | select csname, lastbootuptime
                
                $lastboouptime=$LASTREBOOT.lastbootuptime
                }
                $date=Get-Date

                 
                 If ($clientUpdate.count -gt 0) {
                     $updateVulnAudit.AddRange($clientUpdate) | Out-Null
                    
                }

                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    

                    $uiHash.Listview.Items.EditItem($Computer)
                    if($clientUpdateInstalled -gt 0)
                    {
                     
                    $Computer.Notes = "Patch Installed, Please Check Report for more details!"
                    $Computer.Audited= $clientUpdateInstalled.count
                    $Computer.Installed= $clientUpdateInstalled.count
                    $Computer.LastRebootTime= $clientUpdate.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientUpdate.LastPatchInstalledTime
                    $Computer.OSVersion = $clientUpdate.OSVersion
                    $Computer.UnAuthorized=0
                    $Computer.Offline= 0
                    $Computer.StatusTxt="Installation Completed"

                    $Computer.Status= "$pwd\Images\ok.png"

                    }

                    if($clientUpdateunauthorized -gt 0)
                    {
                    $Computer.Notes = "Un-Authorized Access!"
                    $Computer.Offline= 0
                    $Computer.Audited= 1
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=$clientUpdateunauthorized.count
                    $Computer.OSVersion = "UnAuthorizedAccess"
                    $Computer.LastRebootTime= "UnAuthorizedAccess"
                    $Computer.LastPatchInstalledTime= "UnAuthorizedAccess"
                    $Computer.StatusTxt="No Access"
                    $Computer.Status= "$pwd\Images\noaccess.png"
                    
                     }
                    if($clientUpdateNotInstalled -gt 0)
                    {
                    $Computer.Notes = "Patch Not Installed, Please Check Report for more details!"
                    $Computer.LastRebootTime= $clientUpdate.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientUpdate.LastPatchInstalledTime
                     $Computer.OSVersion = $clientUpdate.OSVersion
                    $Computer.Offline= 0
                    $Computer.Audited= 1
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=0
                    $Computer.StatusTxt="Updates Not Installed"
                    $Computer.Status= "$pwd\Images\warning.png"

                     
                    }
                    if($clientoffline -gt 0)
                    {
                    $Computer.Notes="Offline"
                    $Computer.colorchange = "Red"
                    $Computer.Offline=$clientoffline.Count
                    $Computer.Audited= 0
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=0
                    $Computer.LastRebootTime= "Offline"
                    $Computer.LastPatchInstalledTime= "Offline"
                    $Computer.OSVersion = "Offline"
                    $Computer.StatusTxt="Server Offline"
                    $Computer.Status= "$pwd\Images\offline.png"

                     
              
                    }
                    If ($queryError) {
                        $Computer.colorchange = "Red"  
                        $Computer.notes = $queryError
                    } 
                    Else {
                    if($clientUpdateError -gt 0)
                     {
                     $notes=$clientUpdate.Notes
                     $Computer.notes = "$notes"
                     $Computer.Offline= 0
                     $Computer.Audited= 1
                     $Computer.Installed= 0
                     $Computer.UnAuthorized=0
                     $Computer.InstallErrors= $clientUpdateError.count
                     $Computer.LastRebootTime= "Error"
                     $Computer.LastPatchInstalledTime= "Error"
                     $Computer.OSVersion = "Error"
                     $Computer.StatusTxt="Error Occured"
                     $Computer.Status= "$pwd\Images\error.png"
                      
                                       
                        }
                        
                    }
                   
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                
                
                
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                                     
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"    
                                         
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Patch Audit"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($updateVulnAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion                 


        } 
        
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Install SQL Patches') {             
            #region Install Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Installing SQL Patches for listed servers...Please Wait"              
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $installAudit,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Installing Patches"
                    $computer.colorchange = "Black" 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  

                

                Set-Location -Path $Path 
                . .\Scripts\Install-SQLPatches.ps1
                 
                 $Computer.Status= "$pwd\Images\progress.png"


                try{
                 
                if(($uiHash.sqlcustomCheckBox.IsChecked) -and ($uiHash.sqlstaycurrentCheckBox.IsChecked))
                {
                 
                Write-Host "You cannot Run both Custom Patch and Stay Current Patch at the same Time!" -ForegroundColor Red
                exit

                }
                elseif($uiHash.sqlstaycurrentCheckBox.IsChecked)
                {
                
                write-host "Performing Stay Current Patches"
                $pathv=""
                $cpylocally=$false
                $clientInstall = @(SQL-Patch -Computername $computer.computer -Srcpath $pathv -cpylocally $cpylocally -$Type "StayCurrent")  
                }
                elseif ($uiHash.sqlcustomCheckBox.IsChecked)
                {
                 $pathv=$uiHash.CommentTextBox.Text
                 write-host "Performing Custom Patches"
                 $clientInstall = @(SQL-Patch -Computername $computer.computer -Srcpath $pathv -cpylocally $cpylocally -$Type "Custom")  
                }
                else{
                write-host "No action selected"
                }

                #$clientInstall = @(SQL-Patch -Computername $computer.computer -Srcpath $path -cpylocally $true)  
                $clientInstalledCount =  @($clientInstall |  Select-String -Pattern "Installed" -CaseSensitive).Count
                $clientInstalledErrorCount = @($clientInstall | Select-String -Pattern "ERROR" -CaseSensitive).Count
                $clientUpdateunauthorized =  @($clientInstall |  Select-String -Pattern "UnAuthorizedAccess" -CaseSensitive).Count
                $clientsqlNotFound =  @($clientInstall |  Select-String -Pattern "SQLDoesnotExists" -CaseSensitive).Count
                $clientoffline =  @($clientInstall |  Select-String -Pattern "Offline" -CaseSensitive).Count
                $clientRebootPending=  @($clientInstall |  Select-String -Pattern "RebootPending" -CaseSensitive).Count 
                }
             Catch {
                        $queryError = $_.Exception.Message
                        write-host $queryError
                    }

 



      If (Test-Connection -ComputerName $computer.computer -Count 1 -Quiet) {
                $sOS=(Get-WmiObject -ComputerName $computer.computer -Class Win32_OperatingSystem).Name
                $LASTREBOOT = Get-CimInstance -ComputerName $computer.computer -ClassName win32_operatingsystem | select csname, lastbootuptime
                
                $lastboouptime=$LASTREBOOT.lastbootuptime
                }
                $date=Get-Date

                 
                 If ($clientInstall.count -gt 0) {
                     $installAudit.AddRange($clientInstall) | Out-Null
                }


                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    

                    $uiHash.Listview.Items.EditItem($Computer)
                    if($clientInstalledCount -gt 0)
                    { 
                    $Computer.Notes = "Installation is Completed Successfully, Please roboot the server!"  
                    $Computer.Audited= $clientInstalledCount.count
                    $Computer.Installed= $clientInstalledCount.count
                    $Computer.LastRebootTime= $clientInstall.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientInstall.LastPatchInstalledTime
                    $Computer.OSVersion = $clientInstall.OSVersion
                    $Computer.StatusTxt="Installation Completed"
                 
                    $Computer.UnAuthorized=0
                    $Computer.Offline= 0
                    $Computer.Status= "$pwd\Images\found.png"
                        
                    }
                    if($clientRebootPending -gt 0)
                    {
                    $Computer.Notes = "This server is Pending Reboot!" 
                    $Computer.Audited= $clientRebootPending.count
                    $Computer.Installed= $clientRebootPending.count
                    $Computer.LastRebootTime= $clientInstall.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientInstall.LastPatchInstalledTime
                    $Computer.OSVersion = $clientInstall.OSVersion
                    $Computer.UnAuthorized=0
                    $Computer.Offline= 0
                     $Computer.StatusTxt="Reboot Required"
                    $Computer.Status= "$pwd\Images\restart.png"

                    
                    }
                    if($clientsqlNotFound -gt 0)
                    {
                     $Computer.Notes = "SQL Server does not exists in this server!" 
                    $Computer.Audited= $clientsqlNotFound.count
                    $Computer.Installed= $clientsqlNotFound.count
                    $Computer.LastRebootTime= $clientInstall.LastRebootTime
                    $Computer.LastPatchInstalledTime= $clientInstall.LastPatchInstalledTime
                    $Computer.OSVersion = $clientInstall.OSVersion
                    $Computer.UnAuthorized=0
                    $Computer.Offline= 0
                    $Computer.StatusTxt="SQL Not Installed"
                    $Computer.Status= "$pwd\Images\nosql.png"
                    }
                    

                    if($clientUpdateunauthorized -gt 0)
                    {
                    $Computer.Notes = "Un-Authorized Access!"
                    $Computer.Offline= 0
                    $Computer.Audited= 1
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=$clientUpdateunauthorized.count
                    $Computer.LastRebootTime= "UnAuthorizedAccess"
                    $Computer.LastPatchInstalledTime= "UnAuthorizedAccess"
                    $Computer.OSVersion = "UnAuthorizedAccess"
                    $Computer.StatusTxt="No Access"
                    $Computer.Status= "$pwd\Images\noaccess.png"
                    
                    }

                    if($clientoffline -gt 0)
                    {
                    $Computer.Notes="Offline"
                    $Computer.colorchange = "Red"
                    $Computer.Offline=$clientoffline.Count
                    $Computer.Audited= 0
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=0
                    $Computer.LastRebootTime= "Offline"
                    $Computer.LastPatchInstalledTime= "Offline"
                    $Computer.OSVersion = "Offline"
                    $Computer.StatusTxt="Server Offline"
                    $Computer.Status= "$pwd\Images\offline.png"
                    }

                    If ($queryError) {
                     $Computer.colorchange = "Red"  
                        $Computer.notes = $queryError
                    } Else {
                    
                    if($clientInstalledErrorCount -gt 0)
                     {
                     $notes=$clientInstall.Notes
                     $Computer.notes = "$notes"
                     $Computer.Offline= 0
                     $Computer.Audited= 1
                     $Computer.Installed= 0
                     $Computer.UnAuthorized=0
                     $Computer.InstallErrors= $clientInstalledErrorCount.count
                     $Computer.LastRebootTime= "Error"
                     $Computer.LastPatchInstalledTime= "Error"
                     $Computer.OSVersion = "Error"
                     $Computer.StatusTxt="Error Occured"
                     $Computer.Status= "$pwd\Images\error.png"
 
                       }                    
                       
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                }) 
               
              


                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"  
                                           
                    }
                })                  
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()  

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Patch Install"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($installAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion  
       
        }    
            
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Install Stay Current Patches') {             
            #region Install Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Installing Stay Current Patches for all servers...Please Wait"              
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $installAudit,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Installing Patches"
                    $Computer.colorchange = "Black" 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  

                

                Set-Location -Path $Path 
                . .\Scripts\Install-StayCurrent.ps1
                
                $Computer.Status= "$pwd\Images\progress.png"

                $clientInstall = @(Stay-Current -Computername $computer.computer)
                $installAudit.AddRange($clientInstall) | Out-Null
                $clientInstalledCount =  @($clientInstall | Where {$_.Notes -notmatch "Failed to Install Patch|ERROR|No Updates Avaiable|Reboot is required|failed|Offline|Error with Audit"}).Count
                $clientInstalledErrorCount = @($clientInstall | Where {$_.Notes -match "Failed to Install Patch|ERROR|failed|Error with Audit"}).Count
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                $uiHash.Listview.Items.EditItem($Computer)
                
                     
                    If ($clientInstall.Title -eq "Offline") {                        
                        $Computer.Installed = 0     
                        $Computer.colorchange = "Red"   
                        $Computer.Status= "$pwd\Images\offline.png"
                 
                    } 
                    elseif($clientInstall.Title -eq "Reboot is required")
                    {
                        $Computer.Installed = 0        
                         $Computer.colorchange = "Orange"   
                          $Computer.Status= "$pwd\Images\restart.png"         
                    }
                    elseif($clientInstall.Title -eq "Patching Completed Successfully")
                    {
                        $Computer.Installed = $clientInstalledCount
                        $Computer.InstallErrors = $clientInstalledErrorCount    
                        $Computer.colorchange = "Green"             
                         $Computer.Status= "$pwd\Images\ok.png"   
                    }
                    elseif(($clientInstall.Title -eq "ERROR") -or ($clientInstall.Title -eq "failed")-or ($clientInstall.Title -eq "Error with Audit"))
                    {
                        $Computer.Installed = 0
                        $Computer.InstallErrors = $clientInstalledErrorCount
                        
                         $Computer.colorchange = "Red"
                          $Computer.Status= "$pwd\Images\error.png"   
                    }
                    Else {
                        $Computer.Installed = $clientInstalledCount
                        $Computer.InstallErrors = $clientInstalledErrorCount
                        
                         $Computer.colorchange = "Red"
                          $Computer.Status= "$pwd\Images\error.png"   
                    }
                   
                    $Computer.Notes =  $clientInstall.Notes
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"  
                                           
                    }
                })                  
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()  

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Patch Install"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($installAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion  
       
        }       
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Reboot Server') {
            #region Reboot
            If ((Show-RebootWarning) -eq "Yes") {
                $uiHash.RunButton.IsEnabled = $False
                $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
                $uiHash.CancelButton.IsEnabled = $True
                $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
                $uiHash.StatusTextBox.Foreground = "Black"
                $uiHash.StatusTextBox.Text = "Rebooting Servers..."            
                $uiHash.StartTime = (Get-Date)
            
                [Float]$uiHash.ProgressBar.Value = 0
                $scriptBlock = {
                    Param (
                        $Computer,
                        $uiHash,
                        $Path
                    )               
                    $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                        $uiHash.Listview.Items.EditItem($Computer)
                        $computer.Notes = "Rebooting..."
                        $uiHash.Listview.Items.CommitEdit()
                        $uiHash.Listview.Items.Refresh() 
                    })                
                    Set-Location $Path
                    $retvalue=""
                     $Computer.Status= "$pwd\Images\progress.png"

                    If (Test-Connection -Computer $Computer.computer -count 1 -Quiet) {
                        Try {

                            $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
                            $username = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $username, $pass

                             

                            $serverNameOrIp = $Computer.computer
                            Restart-Computer -ComputerName $serverNameOrIp `
                                             -Authentication default `
                                             -Credential $cred -Force -ea stop


                            
                            Do {
                                Start-Sleep -Seconds 2
                                Write-Verbose ("Waiting for {0} to shutdown..." -f $Computer.computer)
                            }
                            While ((Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet))    
                            Do {
                                Start-Sleep -Seconds 5
                                $i++        
                                Write-Verbose ("{0} down...{1}" -f $Computer.computer, $i)
                                If($i -eq 60) {
                                    Write-Warning ("{0} did not come back online from reboot!" -f $Computer.computer)
                                    $connection = $False
                                    $retvalue="$Computer.computer did not come back online from reboot"
                                }
                            }
                            While (-NOT(Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet))
                            Write-Verbose ("{0} is back up" -f $Computer.computer)
                            $connection = $True
                            $retvalue="Online"
                               $Computer.Status= "$pwd\Images\ok.png"
                        }
                         Catch {
                            Write-Warning "$($Error[0])"
                            $connection = $False
                            $retvalue="$($Error[0])"
                        }
                    }
                    else{
                        $connection = $False
                        $retvalue="Offline"
                        }
                    $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{                    
                        $uiHash.Listview.Items.EditItem($Computer)
                        If ($Connection) {
                            $Computer.Notes = $retvalue
                        } ElseIf (-Not $Connection) {
                            $Computer.Notes = $retvalue
                        } Else {
                            $Computer.Notes = "Unknown"
                        } 
                        $uiHash.Listview.Items.CommitEdit()
                        $uiHash.Listview.Items.Refresh()
                    })
                    $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                        $uiHash.ProgressBar.value++  
                    })
                    $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                        #Check to see if find job
                        If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                            $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                            $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                            $uiHash.RunButton.IsEnabled = $True
                            $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                            $uiHash.CancelButton.IsEnabled = $False
                            $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"  
                                               
                        }
                    })  
                
                }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxRebootJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()  

                ForEach ($Computer in $selectedItems) {
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Pending Reboot"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                    #Create the powershell instance and supply the scriptblock with the other parameters 
                    $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                    #Add the runspace into the powershell instance
                    $powershell.RunspacePool = $runspaceHash.runspacepool
           
                    #Create a temporary collection for each runspace
                    $temp = "" | Select-Object PowerShell,Runspace,Computer
                    $Temp.Computer = $Computer.computer
                    $temp.PowerShell = $powershell
           
                    #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                    $temp.Runspace = $powershell.BeginInvoke()
                    Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                    $jobs.Add($temp) | Out-Null
                }                
            }#endregion           
        } 
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Ping Sweep') {
            #region PingSweeps
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking server connection..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Checking connection"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Set-Location $Path
                $Connection = (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet)
                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    If ($Connection) {
                        $Computer.Notes = "Online"
                    } ElseIf (-Not $Connection) {
                        $Computer.Notes = "Offline"
                    } Else {
                        $Computer.Notes = "Unknown"
                    } 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png" 
                                            
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()   

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Network Test"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }
            #endregion           
        }  
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Snooze Server (Classic)') {
            #region Audit Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Snoozing Listed Server(s)..Please Wait"            
            $Global:updatelayout = [Windows.Input.InputEventHandler]{ $uiHash.ProgressBar.UpdateLayout() }
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $SnoozeAudit,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Snoozing in progress.."
                    $Computer.colorchange = "Black" 
                     
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  
                Set-Location $path
                . .\Scripts\SnoozeServer.ps1           
                
                $Computer.Status= "$pwd\Images\progress.png"
                    
                $clientSnooze = @(SnoozeServer -Computer $computer.computer)
                 
                   
                $clientUpdateError =  @($clientSnooze |  Select-String -Pattern "Error" -CaseSensitive).Count 
                $clientUpdateInstalled =  @($clientSnooze |  Select-String -Pattern "Completed" -CaseSensitive).Count 
                $clientUpdateInstallednotdownloaded=0
                $SnoozeAudit.AddRange($clientSnooze) | Out-Null

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                $uiHash.Listview.Items.EditItem($Computer)
                
                 If ($clientUpdateError-gt 0) {                        
                        $Computer.Audited = 0
                        $Computer.Notes = "Error"
                        $Computer.InstallErrors = $clientUpdateError
                        $Computer.colorchange = "Red" 
                        $Computer.Status= "$pwd\Images\error.png"
                        $Computer.StatusTxt= "Error Occured"
                    }      
                  
                  if($clientUpdateInstalled -gt 0) 
                  {
                        $Computer.Audited = 0
                        $Computer.Notes = "Completed"
                        $Computer.InstallErrors = 0
                        $Computer.colorchange = "Green" 
                         $Computer.StatusTxt= "Error Occured"
                        $Computer.Status= "$pwd\Images\ok.png"
                        $Computer.StatusTxt= "Installation Completed"
                        
                  }
                 

                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++ 
                }) 
                    
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"  
                                           
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Snooze Server"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($SnoozeAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion                 
        } 
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Un-Snooze Server (Classic)') {
            #region Audit Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Un-Snoozing Listed Server(s)..Please Wait"            
            $Global:updatelayout = [Windows.Input.InputEventHandler]{ $uiHash.ProgressBar.UpdateLayout() }
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $UnSnoozeAudit,
                    $uiHash
                )
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $computer.Notes = "Un-Snoozing in progress.."
                    $Computer.colorchange = "Black"  
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  
                Set-Location $path
                . .\Scripts\UN-SnoozeServer.ps1            
                $clientUnSnooze = @(UnSnoozeServer -Computer $computer.computer)
                 
                   
                $clientUpdateError =  @($clientUnSnooze |  Select-String -Pattern "Error" -CaseSensitive).Count 
                $clientUpdateInstalled =  @($clientUnSnooze |  Select-String -Pattern "Completed" -CaseSensitive).Count 
                $clientUpdateInstallednotdownloaded=0
                $UnSnoozeAudit.AddRange($clientUnSnooze) | Out-Null

                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.Listview.Items.EditItem($Computer)
                
                 If ($clientUpdateError -gt 0) {                        
                        $Computer.Audited = 0
                        $Computer.Notes = "Error"
                        $Computer.InstallErrors = $clientUpdateError
                        $Computer.colorchange = "Red"
                        $Computer.StatusTxt = "Error Occured"
                        $Computer.Status= "$pwd\Images\error.png"
                         
                    }      
                  if($clientUpdateInstalled -gt 0) 
                  {
                        $Computer.Audited = 0
                        $Computer.Notes = "Completed"
                        $Computer.InstallErrors = 0
                        $Computer.colorchange = "Green" 
                        $Computer.StatusTxt = "Installation Completed"
                        $Computer.Status= "$pwd\Images\ok.png"
                  }
                  

                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{
                    $uiHash.ProgressBar.value++ 
                }) 
                    
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"      
                                       
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Un-Snooze Server"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($UnSnoozeAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion                 
        } 
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Services Check (Stopped)') {
            #region Check Services
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Checking for Non-Running Automatic Services..."            
            $uiHash.StartTime = (Get-Date)
             

            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $servicesAudit,
                    $ServiceText
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.colorchange = "Black" 
                    $computer.Notes = "Checking for non-running services set to Auto"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Clear-Variable queryError -ErrorAction SilentlyContinue
                Set-Location $Path

                $Computer.Status= "$pwd\Images\progress.png"
               
                If (Test-Connection -ComputerName $computer.computer -Count 1 -Quiet) {
                    Try {
                    
                   
                   
                     if($ServiceText -ne "")
                        {
                      
                       $getservices=$ServiceText -split ","
                       
                       $A=""
                       $cnt=0
                       $cntservice=@($getservices).Count
                       if($cntservice -gt 1)
                        {
                                for($cnt -eq 0;$cnt -lt $cntservice;$cnt++)
                                  {
                                  if($cnt -ne 0)
                                    {
                                     
                                       $value=$getservices[$cnt]
                                        if($cnt  -eq $cntservice-1)
                                       {
                                        $A +=" OR Name Like '%$value%') "   
                                        }else{
                                        
                                         $A +=" OR Name Like '%$value%' "   
                                        }
                                    }
                                    elseif($cnt -eq 0)
                                    {
                                     $value=$getservices[$cnt]
                                     $Q="Select __Server,Name,DisplayName,State,StartMode,ExitCode,Status FROM Win32_Service WHERE (StartMode='Auto' AND State!='Running') AND (Name Like '%$value%'"
                                    }
                                  }
                                }else{
                                
                                 $Q="Select __Server,Name,DisplayName,State,StartMode,ExitCode,Status FROM Win32_Service WHERE StartMode='Auto' AND State!='Running' AND Name Like '%$ServiceText%'"
                                }
                     
                                    $wmi = @{
                                ErrorAction = 'Stop'
                                Computername = $computer.computer
                                Query = $Q + $A
                            }
             
                        
                        }
                        
                        else{
                        
                                $wmi = @{
                            ErrorAction = 'Stop'
                            Computername = $computer.computer
                            Query = "Select __Server,Name,DisplayName,State,StartMode,ExitCode,Status FROM Win32_Service WHERE StartMode='Auto' AND State!='Running'"
                        }
                
                        }

                        
                            

                            $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
                            $username = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $username, $pass


                        $services = @(Get-WmiObject @wmi -Credential $cred )
 
                    } Catch {
                        $queryError = $_.Exception.Message
                    }
                } Else {
                    $queryError = "Offline"
                }
                If ($services.count -gt 0) {
                    $servicesAudit.AddRange($services) | Out-Null
                }

                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.Services = $services.count
                    If ($queryError) {
                     $Computer.colorchange = "Red"  
                        $Computer.notes = $queryError
                        $Computer.Status= "$pwd\Images\error.png"

                    } Else {
                     $Computer.colorchange = "Green"  
                        $Computer.notes = 'Completed'
                        $Computer.Status= "$pwd\Images\ok.png"
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"   
                                          
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Service Check"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($servicesAudit).AddArgument($ServiceText)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion
          }
           ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Services Action (Start)') {
            #region Check Services
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Action for Starting Services..."            
            $uiHash.StartTime = (Get-Date)
             

            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $servicesAuditStart,
                    $ServiceText
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.colorchange = "Black" 
                    $computer.Notes = "Performing Action to start the Mentioned Services.."
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Clear-Variable queryError -ErrorAction SilentlyContinue
                Set-Location $Path

                $Computer.Status= "$pwd\Images\progress.png"
               
                If (Test-Connection -ComputerName $computer.computer -Count 1 -Quiet) {
                    Try {
                    
                   
                   
                     if($ServiceText -ne "")
                        {
                      
                       $getservices=$ServiceText -split ","
                       
                       $A=""
                       $cnt=0
                       $cntservice=@($getservices).Count
                       if($cntservice -gt 1)
                        {
                                for($cnt -eq 0;$cnt -lt $cntservice;$cnt++)
                                  {
                                  if($cnt -ne 0)
                                    {
                                     
                                       $value=$getservices[$cnt]
                                        if($cnt  -eq $cntservice-1)
                                       {
                                        $A +=" OR Name Like '%$value%') "   
                                        }else{
                                        
                                         $A +=" OR Name Like '%$value%' "   
                                        }
                                    }
                                    elseif($cnt -eq 0)
                                    {
                                     $value=$getservices[$cnt]
                                     $Q="Select __Server,Name,DisplayName,State,StartMode,ExitCode,Status  FROM Win32_Service WHERE (StartMode='Auto' AND State!='Running') AND (Name Like '%$value%'"
                                    }
                                  }
                                }else{
                                
                                 $Q="Select __Server,Name,DisplayName,State,StartMode,ExitCode,Status FROM Win32_Service WHERE StartMode='Auto' AND State!='Running' AND Name Like '%$ServiceText%'"
                                }
                     
                                    $wmi = @{
                                ErrorAction = 'Stop'
                                Computername = $computer.computer
                                Query = $Q + $A
                            }
             
                        
                        }
                        
                        else{
                        
                                $wmi = @{
                            ErrorAction = 'Stop'
                            Computername = $computer.computer
                            Query = "Select  __Server,Name,DisplayName,State,StartMode,ExitCode,Status FROM Win32_Service WHERE StartMode='Auto' AND State!='Running'"
                        }
                
                        }

                        
                            

                            $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
                            $username = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $username, $pass


                        $services = @(Get-WmiObject @wmi -Credential $cred)
                        $Computer.notes="[Information] : Starting Service Please Wait .."
                       
                        foreach($service in $services)
                        {
                        $servname=$service.Name
                         Write-Host "[Information] : Starting Service [$servname] Please Wait"
                       
                         $service = Get-WmiObject -ComputerName $service.__Server -Class Win32_Service -Filter "Name='$servname'" -Credential $cred
                         $service.StartService()
                        }
                         
 
                    } Catch {
                        $queryError = $_.Exception.Message
                    }
                } Else {
                    $queryError = "Offline"
                }
                If ($services.count -gt 0) {
                    $servicesAuditStart.AddRange($services) | Out-Null
                }

                $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.Services = $services.count
                    If ($queryError) {
                     $Computer.colorchange = "Red"  
                        $Computer.notes = $queryError
                        $Computer.Status= "$pwd\Images\error.png"

                    } Else {
                     $Computer.colorchange = "Green"  
                        $Computer.notes = 'Completed'
                        $Computer.Status= "$pwd\Images\ok.png"
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"   
                                          
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Performing Service Start"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($servicesAuditStart).AddArgument($ServiceText)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion
          }

        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'SQL Jobs Check') {
            #region Check Services
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Auditing SQL Jobs..."            
            $uiHash.StartTime = (Get-Date)
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Computer,
                    $uiHash,
                    $Path,
                    $JobsAudit
                )               
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                    $Computer.colorchange = "Black" 
                    $computer.Notes = "Auditing SQL Jobs"
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })                
                Clear-Variable queryError -ErrorAction SilentlyContinue
                Set-Location $Path
                $Computer.Status= "$pwd\Images\progress.png"

                If (Test-Connection -ComputerName $computer.computer -Count 1 -Quiet) {
                    Try {
                        
                Set-Location $path
                . .\Scripts\SQlServerJob.ps1            
                $clientJob = @(Get-Jobstatus -Computer $computer.computer|select Computer,name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,Notes | where {$_.Computer -ne $null})
                $clientJobInstalledorNot =  @($clientJob | Where {$_.Notes -eq "SQL Server Not Installed"}).Count
                $clientRunning =  @($clientJob | Where {$_.CurrentRunStatus -eq "Executing"}) 
               
                 
                  
                    } Catch {
                        $queryError = $_.Exception.Message
                    }
                } Else {
                    $queryError = "Offline"
                }

                 
                 
                 If ($clientJob.count -gt 0) {
                    $JobsAudit.AddRange($clientJob) | Out-Null
                }


               $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                    $uiHash.Listview.Items.EditItem($Computer)
                    if($clientJobInstalledorNot -gt 0)
                    {
                    $Computer.SQLJobs = 0
                    $Computer.RunningSQLJobs= 0
                    
                    }
                    else{
                    $Computer.SQLJobs = $clientJob.count
                    $Computer.RunningSQLJobs= $clientRunning.count
                    }

                    If ($queryError) {
                     $Computer.colorchange = "Red"  
                        $Computer.notes = $queryError
                        $Computer.Status= "$pwd\Images\error.png"
                    } Else {
                     
                     if($clientJobInstalledorNot -gt 0)
                     {
                      $Computer.colorchange = "Red" 
                      $Computer.notes = 'SQL Server Not Installed In This Machine'
                      $Computer.Status= "$pwd\Images\warning.png"  
                        }
                        else{
                         $Computer.colorchange = "Green" 
                        $Computer.notes = 'Completed'
                        $Computer.Status= "$pwd\Images\ok.png"
                        }
                    }
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh()
                })
                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"     
                                        
                    }
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Job Check"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($JobsAudit)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion
        }   
        ElseIf ($uiHash.RunOptionComboBox.Text -eq 'Search Installed KBs') {
            #region Audit Patches
            $uiHash.RunButton.IsEnabled = $False
            $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
            $uiHash.CancelButton.IsEnabled = $True
            $uiHash.CancelImage.Source = "$pwd\Images\stop.png"            
            $uiHash.StatusTextBox.Foreground = "Black"
            $uiHash.StatusTextBox.Text = "Searching Installed KBs...Please Wait"            
            $Global:updatelayout = [Windows.Input.InputEventHandler]{ $uiHash.ProgressBar.UpdateLayout() }
            $uiHash.StartTime = (Get-Date)
             
            
            [Float]$uiHash.ProgressBar.Value = 0
            $scriptBlock = {
                Param (
                    $Path,
                    $Computer,
                    $KBAudit,
                    $uiHash
                )
                
                $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                    $uiHash.Listview.Items.EditItem($Computer)
                   
                    if(@($uiHash.KBInputBox.Text -split ",").Count -gt 1)
                    {
                     $computer.Notes = "Searching KB from list provided ..."
                    }
                    else{
                    $computer.Notes = "Searching KB - " + $uiHash.KBInputBox.Text
                    }

                    $Computer.colorchange = "Black" 
                    $uiHash.Listview.Items.CommitEdit()
                    $uiHash.Listview.Items.Refresh() 
                })  


             
                Set-Location $Path
                . .\Scripts\Get-KBs.ps1          
                
                 $Computer.Status= "$pwd\Images\progress.png"

                     Try {
               
               write-host "Attempting to connect to computer " $Computer.computer

                 
                        $clientKB = @(GetKB -Computer $Computer.computer)
                        $clientKBInstalled =  @($clientKB |  Select-String -Pattern "KB is Installed" -CaseSensitive).Count
                        $clientKBNotInstalled =  @($clientKB |  Select-String -Pattern "KB is not Installed" -CaseSensitive).Count 
                        $clientoffline =  @($clientKB |  Select-String -Pattern "Offline" -CaseSensitive).Count
                        $clientUpdateunauthorized =  @($clientKB |  Select-String -Pattern "UnAuthorizedAccess" -CaseSensitive).Count
                        $clientUpdateError =  @($clientKB |  Select-String -Pattern "ERROR" -CaseSensitive).Count
                     

                }
                 Catch {
                        $queryError = $_.Exception.Message

                         write-host $queryError
                    }
                  
                
               If (Test-Connection -ComputerName $Computer.computer -Count 1 -Quiet) {
                
                        $sOS=(Get-WmiObject -ComputerName $computer.computer -Class Win32_OperatingSystem).Name
                        $LASTREBOOT = Get-CimInstance -ComputerName $computer.computer -ClassName win32_operatingsystem | select csname, lastbootuptime
                        $lastboouptime=$LASTREBOOT.lastbootuptime
                        $date=Get-Date

               
                 
                                 If ($clientKB.count -gt 0) {
                                     $KBAudit.AddRange($clientKB) | Out-Null
                                }
    
                                    $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    

                                    $uiHash.Listview.Items.EditItem($Computer)
                  

                            if($clientoffline -gt 0)
                       {
                            $Computer.Notes="Offline"
                            $Computer.colorchange = "Red"
                            $Computer.Offline=$clientoffline.Count
                            $Computer.Audited= 0
                            $Computer.Installed= 0
                            $Computer.UnAuthorized=0
                            $Computer.LastRebootTime= "Offline"
                            $Computer.LastPatchInstalledTime= "Offline"
                            $Computer.OSVersion = "Offline"
                            $Computer.SqlVersion = "Offline"
                            $Computer.StatusTxt="Server Offline"
                            $Computer.Status= "$pwd\Images\offline.png"
                    }
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    else{
                  
                            if($clientKB.LastRebootTime.count -gt 1)
                            {
                            $Computer.LastRebootTime= $clientKB.LastRebootTime[0]
                            }
                            else{
                            $Computer.LastRebootTime= $clientKB.LastRebootTime
                            }
                    
                            if($clientKB.LastPatchInstalledTime.count -gt 1)
                            {
                            $Computer.LastPatchInstalledTime= $clientKB.LastPatchInstalledTime[0]
                            }
                            else{
                            $Computer.LastPatchInstalledTime= $clientKB.LastPatchInstalledTime
                            }

                            if($clientKB.OSVersion.count -gt 1)
                            {
                            $Computer.OSVersion = $clientKB.OSVersion[0]
                    
                            }
                            else{
                            $Computer.OSVersion = $clientKB.OSVersion
                    
                            }
                             if($clientKB.SqlVersion.count -gt 1)
                            {
                            $Computer.SqlVersion = $clientKB.SqlVersion[0]
                    
                            }
                            else{
                            $Computer.SqlVersion = $clientKB.SqlVersion
                    
                            }
                     


                            if($clientKBInstalled -gt 0)
                            { 
                       
                            $Computer.UnAuthorized=0
                            $Computer.Offline= 0
                            $Computer.StatusTxt="KB Installed"
                            $Computer.Notes="KB is found in this machine"
                            $Computer.Status= "$pwd\Images\found.png"
                        
                            }

                            else
                            {
                    
                            $Computer.UnAuthorized=0
                            $Computer.Offline= 0
                            $Computer.StatusTxt="KB Not Installed"
                            $Computer.Notes="KB is not found in this machine"
                            $Computer.Status= "$pwd\Images\found.png"
                    
                            }

                             If ($queryError) {
                             $Computer.colorchange = "Red"  
                                $Computer.notes = $queryError
                            }  
                     
                       }

                   
                            $uiHash.Listview.Items.CommitEdit()
                            $uiHash.Listview.Items.Refresh()
                        })

                }
               else{
                 $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    

                     $uiHash.Listview.Items.EditItem($Computer)

                    $Computer.Notes="Offline"
                    $Computer.colorchange = "Red"
                    $Computer.Offline=1
                    $Computer.Audited= 0
                    $Computer.Installed= 0
                    $Computer.UnAuthorized=0
                    $Computer.LastRebootTime= "Offline"
                    $Computer.LastPatchInstalledTime= "Offline"
                    $Computer.OSVersion = "Offline"
                    $Computer.StatusTxt="Server Offline"
                    $Computer.SqlVersion = "Offline"
                    $Computer.Status= "$pwd\Images\offline.png"

                       $uiHash.Listview.Items.CommitEdit()
                            $uiHash.Listview.Items.Refresh()
                        })
                }
                

                $uiHash.ProgressBar.Dispatcher.Invoke("Normal",[action]{   
                    $uiHash.ProgressBar.value++  
                })
                                     
                $uiHash.Window.Dispatcher.Invoke("Normal",[action]{
                    #Check to see if find job
                    If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) {    
                        $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)                                     
                        $uiHash.RunButton.IsEnabled = $True
                        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                        $uiHash.CancelButton.IsEnabled = $False
                        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"   
                                      
                    }    
                })  
                
            }

            Write-Verbose ("Creating runspace pool and session states")
            $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
            $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
            $runspaceHash.runspacepool.Open()    

            ForEach ($Computer in $selectedItems) {
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Pending Patch Audit"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
                #Create the powershell instance and supply the scriptblock with the other parameters 
                $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Path).AddArgument($computer).AddArgument($KBAudit).AddArgument($uiHash)
           
                #Add the runspace into the powershell instance
                $powershell.RunspacePool = $runspaceHash.runspacepool
           
                #Create a temporary collection for each runspace
                $temp = "" | Select-Object PowerShell,Runspace,Computer
                $Temp.Computer = $Computer.computer
                $temp.PowerShell = $powershell
           
                #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                $temp.Runspace = $powershell.BeginInvoke()
                Write-Verbose ("Adding {0} collection" -f $temp.Computer)
                $jobs.Add($temp) | Out-Null                
            }#endregion                 


        } 

                                       
      Else {
        $uiHash.StatusTextBox.Foreground = "Red"
        $uiHash.StatusTextBox.Text = "No server/s selected!"
    }  
    }


}

Function Open-FileDialog {
    $dlg = new-object microsoft.win32.OpenFileDialog
    $dlg.DefaultExt = "*.txt"
    $dlg.Filter = "Text Files |*.txt;*.log"    
    $dlg.InitialDirectory = $path
    [void]$dlg.showdialog()
    Write-Output $dlg.FileName
}

Function Open-FileDialogexe {
    $dlg = new-object microsoft.win32.OpenFileDialog
    $dlg.DefaultExt = "*.txt"
    $dlg.Filter = "EXE Files |*.exe;*.msi"    
    $dlg.InitialDirectory = $path
    [void]$dlg.showdialog()
    Write-Output $dlg.FileName
}

Function Open-DomainDialog {
    $domain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the LDAP path for the Domain or press OK to use the default domain.", 
    "Domain Query", "$(([adsisearcher]'').SearchRoot.distinguishedName)")  
    If (-Not [string]::IsNullOrEmpty($domain)) {
        Write-Output $domain
    }
}


#Build the GUI
[xml]$xaml = @"
<Window  
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    xmlns:local="clr-namespace:Mm.Wpf.Controls"
    x:Name='Window' Title='PowerShell - Patching Tool Utility' WindowStartupLocation = 'CenterScreen' 
    Width = '1080' Height = '675' ShowInTaskbar = 'True' Icon="$Pwd\Images\windows.png" Style="{DynamicResource MetroWindowStyle}">
  
   


    <Window.Background>
        <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
            <LinearGradientBrush.GradientStops> <GradientStop Color='#b9c3c9' Offset='0' /> <GradientStop Color='#b9c3c9' Offset='0.2' /> 
            <GradientStop Color='#b9c3c9' Offset='0.9' /> <GradientStop Color='#b9c3c9' Offset='1' /> </LinearGradientBrush.GradientStops>
        </LinearGradientBrush>
    </Window.Background> 
    <Window.Resources> 
  
   

       <!-- Flat ComboBox -->
  <SolidColorBrush x:Key="ComboBoxNormalBorderBrush" Color="#b9c3c9" />
  <SolidColorBrush x:Key="ComboBoxNormalBackgroundBrush" Color="#ffffff" />
  <SolidColorBrush x:Key="ComboBoxDisabledForegroundBrush" Color="#888" />
  <SolidColorBrush x:Key="ComboBoxDisabledBackgroundBrush" Color="#eee" />
  <SolidColorBrush x:Key="ComboBoxDisabledBorderBrush" Color="#888" />

  <ControlTemplate TargetType="ToggleButton" x:Key="ComboBoxToggleButtonTemplate">
    <Grid VerticalAlignment="Stretch">
      <Grid.ColumnDefinitions>
        <ColumnDefinition />
        <ColumnDefinition Width="20" />
      </Grid.ColumnDefinitions>
      <Border Grid.ColumnSpan="2" Name="Border"
              BorderBrush="{StaticResource ComboBoxNormalBorderBrush}" 
              CornerRadius="0" BorderThickness="1, 1, 1, 1" 
              Background="{StaticResource ComboBoxNormalBackgroundBrush}" />
      <Border Grid.Column="1" Margin="1, 1, 1, 1" BorderBrush="#444" Name="ButtonBorder"
              CornerRadius="0, 0, 0, 0" BorderThickness="0, 0, 0, 0" 
              Background="{StaticResource ComboBoxNormalBackgroundBrush}" />

      <Path Name="Arrow" Grid.Column="1" 
            Data="M0,0 L0,2 L4,6 L8,2 L8,0 L4,4 z"
            HorizontalAlignment="Center" Fill="#1a85ff"
            VerticalAlignment="Center" />
    </Grid>
    <ControlTemplate.Triggers>
      <Trigger Property="UIElement.IsMouseOver" Value="True">
        <Setter Property="Panel.Background" TargetName="ButtonBorder" Value="WhiteSmoke"/>
      </Trigger>
      <Trigger Property="ToggleButton.IsChecked" Value="True">
        <Setter Property="Panel.Background" TargetName="ButtonBorder" Value="WhiteSmoke"/>
        <Setter Property="Shape.Fill" TargetName="Arrow" Value="#FF8D979E"/>
      </Trigger>
      <Trigger Property="UIElement.IsEnabled" Value="False">
        <Setter Property="Panel.Background" TargetName="Border" Value="{StaticResource ComboBoxDisabledBackgroundBrush}"/>
        <Setter Property="Panel.Background" TargetName="ButtonBorder" Value="{StaticResource ComboBoxDisabledBackgroundBrush}"/>
        <Setter Property="Border.BorderBrush" TargetName="ButtonBorder" Value="{StaticResource ComboBoxDisabledBorderBrush}"/>
        <Setter Property="TextElement.Foreground" Value="{StaticResource ComboBoxDisabledForegroundBrush}"/>
        <Setter Property="Shape.Fill" TargetName="Arrow" Value="#999"/>
      </Trigger>
    </ControlTemplate.Triggers>
  </ControlTemplate>

  <Style x:Key="SimpleComboBox"  TargetType="{x:Type ComboBox}">
    <Setter Property="UIElement.SnapsToDevicePixels" Value="True"/>
    <Setter Property="FrameworkElement.OverridesDefaultStyle" Value="True"/>
    <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
    <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
    <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
    <Setter Property="TextElement.Foreground" Value="Black"/>
    <Setter Property="FrameworkElement.FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Control.Template">
      <Setter.Value>
        <ControlTemplate TargetType="ComboBox">
          <Grid VerticalAlignment="Stretch">
            <ToggleButton Name="ToggleButton" Grid.Column="2"
                ClickMode="Press" Focusable="False"
                IsChecked="{Binding Path=IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}"
                Template="{StaticResource ComboBoxToggleButtonTemplate}"/>

            <ContentPresenter Name="ContentSite" Margin="5, 3, 23, 3" IsHitTestVisible="False"
                              HorizontalAlignment="Left" VerticalAlignment="Center"                              
                              Content="{TemplateBinding ComboBox.SelectionBoxItem}" 
                              ContentTemplate="{TemplateBinding ComboBox.SelectionBoxItemTemplate}"
                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"/>
            <TextBox Name="PART_EditableTextBox" Margin="3, 3, 23, 3"                     
                     IsReadOnly="{TemplateBinding IsReadOnly}"
                     Visibility="Hidden" Background="Transparent"
                     HorizontalAlignment="Left" VerticalAlignment="Center"
                     Focusable="True" >
              <TextBox.Template>
                <ControlTemplate TargetType="TextBox" >
                  <Border Name="PART_ContentHost" Focusable="False" />
                </ControlTemplate>
              </TextBox.Template>
            </TextBox>
            <!-- Popup showing items -->
            <Popup Name="Popup" Placement="Bottom"
                   Focusable="False" AllowsTransparency="True"
                   IsOpen="{TemplateBinding ComboBox.IsDropDownOpen}"
                   PopupAnimation="Slide">
              <Grid VerticalAlignment="Stretch" Name="DropDown" SnapsToDevicePixels="True"
                    MinWidth="{TemplateBinding FrameworkElement.ActualWidth}"
                    MaxHeight="{TemplateBinding ComboBox.MaxDropDownHeight}">
                <Border Name="DropDownBorder" Background="White" Margin="0, 1, 0, 0"
                        CornerRadius="0" BorderThickness="1,1,1,1" 
                        BorderBrush="{StaticResource ComboBoxNormalBorderBrush}"/>
                <ScrollViewer Margin="4" SnapsToDevicePixels="True">
                  <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained" />
                </ScrollViewer>
              </Grid>
            </Popup>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="ItemsControl.HasItems" Value="False">
              <Setter Property="FrameworkElement.MinHeight" TargetName="DropDownBorder" Value="95"/>
            </Trigger>
            <Trigger Property="UIElement.IsEnabled" Value="False">
              <Setter Property="TextElement.Foreground" Value="{StaticResource ComboBoxDisabledForegroundBrush}"/>
            </Trigger>
            <Trigger Property="ItemsControl.IsGrouping" Value="True">
              <Setter Property="ScrollViewer.CanContentScroll" Value="False"/>
            </Trigger>
            <Trigger Property="ComboBox.IsEditable" Value="True">
              <Setter Property="KeyboardNavigation.IsTabStop" Value="False"/>
              <Setter Property="UIElement.Visibility" TargetName="PART_EditableTextBox" Value="Visible"/>
              <Setter Property="UIElement.Visibility" TargetName="ContentSite" Value="Hidden"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <!-- End of Flat ComboBox -->
  
  

        
        <DataTemplate x:Key="HeaderTemplate">
            <DockPanel>
                <TextBlock FontSize="10" Foreground="darkblue" FontWeight="Bold" >
                    <TextBlock.Text>
                        <Binding/>
                    </TextBlock.Text>
                </TextBlock>
            </DockPanel>
        </DataTemplate> 
        
         
               
    </Window.Resources>   
    
     
    <Grid VerticalAlignment="Stretch" x:Name = 'Grid' ShowGridLines = 'false'>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = '*'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
        </Grid.RowDefinitions>    
        <Menu Width = 'Auto' HorizontalAlignment = 'Stretch' Grid.Row = '0'>
        <Menu.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#b9c3c9' Offset='0' /> <GradientStop Color='#b9c3c9' Offset='0.2' /> 
                <GradientStop Color='#b9c3c9' Offset='0.9' /> <GradientStop Color='#b9c3c9' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>
        </Menu.Background>
            <MenuItem x:Name = 'FileMenu' Header = '_File'>
                <MenuItem x:Name = 'RunMenu' Header = '_Run' ToolTip = 'Initiate Run operation' InputGestureText ='F5'> </MenuItem>
                <MenuItem x:Name = 'GenerateReportMenu' Header = 'Generate R_eport' ToolTip = 'Generate Report' InputGestureText ='F8'/>
                <Separator />            
                <MenuItem x:Name = 'OptionMenu' Header = '_Options' ToolTip = 'Open up options window.' InputGestureText ='Ctrl+O'/>
                <Separator />
                <MenuItem x:Name = 'ExitMenu' Header = 'E_xit' ToolTip = 'Exits the utility.' InputGestureText ='Ctrl+E'/>
            </MenuItem>
            <MenuItem x:Name = 'EditMenu' Header = '_Edit'>
                <MenuItem x:Name = 'SelectAllMenu' Header = 'Select _All' ToolTip = 'Selects all rows.' InputGestureText ='Ctrl+A'/>               
                <Separator />
                <MenuItem x:Name = 'ClearErrorMenu' Header = 'Clear ErrorLog' ToolTip = 'Clears error log.'> </MenuItem>                
                <MenuItem x:Name = 'ClearAllMenu' Header = 'Clear All' ToolTip = 'Clears everything on the WSUS utility.'/>
            </MenuItem>
            <MenuItem x:Name = 'ActionMenu' Header = '_Action'>
                <MenuItem Header = 'Reports'>
                    <MenuItem x:Name = 'ClearAuditReportMenu' Header = 'Clear Audit Report' ToolTip = 'Clears the current report.'/>
                    <MenuItem x:Name = 'ClearInstallReportMenu' Header = 'Clear Install Report' ToolTip = 'Clears the current report.'/>                   
                    <MenuItem x:Name = 'ClearInstalledUpdateMenu' Header = 'Clear Installed Update Report' ToolTip = 'Clears the installed update report.'/>
                </MenuItem>
                <MenuItem Header = 'Server List'>
                    <MenuItem x:Name = 'ClearServerListMenu' Header = 'Clear Server List' ToolTip = 'Clears the server list.'/>
                    <MenuItem x:Name = 'ClearServerListNotesMenu' Header = 'Clear Server List Notes' ToolTip = 'Clears the server list notes column.'/>
                    <MenuItem x:Name = 'OfflineHostsMenu' Header = 'Remove Offline Servers' ToolTip = 'Removes all Offline hosts from Server List'/>                   
                    <MenuItem x:Name = 'RebootHostsMenu' Header = 'List Reboot Required Servers' ToolTip = 'List Reboot Required Servers'/>      
                    <MenuItem x:Name = 'UnauthorizedHostsMenu' Header = 'List Unathorized Servers' ToolTip = 'List Unauthorized Servers'/>   
                    <MenuItem x:Name = 'UpdatesNeededHostsMenu' Header = 'List Update Needed Servers' ToolTip = 'List Update Needed Servers'/>   
                    <MenuItem x:Name = 'ResetDataMenu' Header = 'Reset Computer List Data' ToolTip = 'Resets the audit and patch data on Server List'/>
                </MenuItem> 
            <Separator />                           
            <MenuItem x:Name = 'HostListMenu' Header = 'Create Host List' ToolTip = 'Creates a list of all servers and saves to a text file.'/>
                <MenuItem x:Name = 'ServerListReportMenu' Header = 'Create Server List Report' 
                ToolTip = 'Creates a CSV file listing the current Server List.'/>
                <Separator/>
                <MenuItem x:Name = 'ViewErrorMenu' Header = 'View ErrorLog' ToolTip = 'Clears error log.'/>            
            </MenuItem>            
            <MenuItem x:Name = 'HelpMenu' Header = '_Help'>
                <MenuItem x:Name = 'AboutMenu' Header = '_About' ToolTip = 'Show the current version and other information.'> </MenuItem>
                <MenuItem x:Name = 'HelpFileMenu' Header = 'WSUS Utility _Help' 
                ToolTip = 'Displays a help file to use the WSUS Utility.' InputGestureText ='F1'> </MenuItem>
            </MenuItem>            
        </Menu>
        <ToolBarTray Grid.Row = '1' Grid.Column = '0'>
        
        <ToolBarTray.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#b9c3c9' Offset='0' /> <GradientStop Color='#b9c3c9' Offset='0.2' /> 
                <GradientStop Color='#b9c3c9' Offset='0.9' /> <GradientStop Color='#b9c3c9' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>        
        </ToolBarTray.Background>
          
            <ToolBar Background = 'Silver' Band = '1' BandIndex = '1'>
        
          
                <Button  Cursor="Hand" x:Name = 'RunButton' Width = 'Auto' ToolTip = 'Performs action against all servers in the server list based on checked radio button.'>
                    <Image x:Name = 'StartImage' Source = '$Pwd\Images\Start.png'/>
                </Button>         
                <Separator Background = 'Black'/>   
                <Button Cursor="Hand" x:Name = 'CancelButton' Width = 'Auto' ToolTip = 'Cancels currently running operations.' IsEnabled = 'False'>
                    <Image x:Name = 'CancelImage' Source = '$pwd\Images\Stop_locked.png' />
                </Button>
                <Separator Background = 'Black'/>
                <ComboBox Style="{StaticResource SimpleComboBox}" x:Name = 'RunOptionComboBox' Width = 'Auto' IsReadOnly = 'True'
                SelectedIndex = '0'>
                
        
                    <TextBlock Tag="Hide"> Check Windows Updates </TextBlock>
                    <TextBlock Tag="Show2"> Search Installed KBs </TextBlock>
                    <TextBlock Tag="Hide"> Check Zero Vulnerability </TextBlock>
                    <Separator />
                    <TextBlock Tag="Hide"> Install Window Patches </TextBlock>
                    <TextBlock Tag="Show"> Install SQL Patches (Under Construction)</TextBlock>
                    <TextBlock Tag="Hide"> Install Stay Current Patches </TextBlock>
                    <Separator />
                    <TextBlock Tag="Hide"> Reboot Server </TextBlock>
                    <TextBlock Tag="Show3"> Services Check (Stopped) </TextBlock>
                    <TextBlock Tag="Show3"> Services Action (Start) </TextBlock>
                    <TextBlock Tag="Hide"> SQL Jobs Check </TextBlock>
                    <Separator />
                    <TextBlock Tag="Hide"> Snooze Server (Classic) </TextBlock>
                    <TextBlock Tag="Hide"> Un-Snooze Server (Classic) </TextBlock>
                     <Separator  />
                    <TextBlock Tag="Hide"> Check Server Health (Under Construction)</TextBlock>

                    
                </ComboBox>                
           </ToolBar>
           <ToolBar Background = 'Silver' Band = '1' BandIndex = '1'>

               <Button Cursor="Hand" x:Name = 'BrowseFileButton' Width = 'Auto' 
                ToolTip = 'Open a file dialog to select a host file. Upon selection, the contents will be loaded into Server list.'>
                    <Image Source = '$pwd\Images\BrowseFile.png' />
                </Button>  

             </ToolBar>
           
           <ToolBar Background = 'Silver' Band = '1' BandIndex = '1'>
                    
                    
                <ComboBox Style="{StaticResource SimpleComboBox}" x:Name = 'ReportComboBox' Width = 'Auto' IsReadOnly = 'True' SelectedIndex = '0'>
                    <TextBlock>Check Window Updates - Report</TextBlock>
                    <TextBlock>Check Zero Vulnerability - Report</TextBlock>
                    <TextBlock>Search Installed KBs - Report</TextBlock>
                    <Separator  />
                    <TextBlock>Install Window Patches - Report</TextBlock>
                    <TextBlock>Install SQL Patches - Report (Under Construction)</TextBlock>
                    <TextBlock>Install Stay Current Patches - Report</TextBlock>
                    <Separator  />
                    <TextBlock>Reboot Server - Report</TextBlock>
                    <TextBlock>Service Check (Stopped) - Report</TextBlock>
                    <TextBlock>Service Action (Start) - Report</TextBlock>
                    <TextBlock>SQL Jobs Check - Report</TextBlock>
                    <Separator  />
                    <TextBlock>Snooze Server (Classic) - Report</TextBlock>
                    <TextBlock>Un-Snooze Server (Classic) - Report</TextBlock>
                     <Separator  />
                    <TextBlock> Check Server Health - Report (Under Construction)</TextBlock>

                </ComboBox>              
 
 
            <Button Cursor="Hand" x:Name = 'GenerateReportButton' Width = 'Auto' ToolTip = 'Generates a report based on user selection.'>
                    <Image Source = '$pwd\Images\Gen_Report.png' />
                </Button>  

                <Separator Background = 'Black'/>

              
            
            <ComboBox Style="{StaticResource SimpleComboBox}" x:Name = 'filterComboBox' Width = 'Auto' IsReadOnly = 'True' SelectedIndex = '0'>
                    <TextBlock>Select All</TextBlock>
                    <TextBlock>Error Occured</TextBlock>
                    <TextBlock>Server Offline</TextBlock>
                    <TextBlock>No Access</TextBlock>
                    <TextBlock>No Update Found</TextBlock>
                    <TextBlock>Reboot Required</TextBlock>
                    <TextBlock>SQL Not Installed</TextBlock> 
                    <TextBlock>Installation Completed</TextBlock> 
                    <TextBlock>Updates Not Installed</TextBlock> 
                    <TextBlock>Updates Found</TextBlock> 

                </ComboBox>              
                
              <Button Cursor="Hand" x:Name = 'FilterButton' Width = 'Auto' 
                ToolTip = 'Select Filter option to filter from the list of results file. Upon selection, the contents will be loaded into Server list as filtered records.'>
                    <Image Source = '$pwd\Images\filter.png' />
                </Button>
                
                <Separator Background = 'Black'/>

                 <Button Cursor="Hand" x:Name = 'ExcelButton' Width = 'Auto' 
                ToolTip = 'Select Excel Button to generate the Excel file for the listed results.'>
                    <Image Source = '$pwd\Images\excel.png' /> 
                </Button> 

                <Separator Background = 'Black'/>

                 <Button Cursor="Hand" x:Name = 'ClearButton' Width = 'Auto' 
                ToolTip = 'This will clear the selection.'>
                    <Image Source = '$pwd\Images\clear.png' /> 
                </Button> 

                 <Separator Background = 'Black'/>

                 <Button Cursor="Hand" x:Name = 'resetButton' Width = 'Auto' 
                ToolTip = 'This will Reset the Table selections.'>
                    <Image Source = '$pwd\Images\reset.png' /> 
                </Button> 

                <Separator Background = 'Black'/>

                <Button Cursor="Hand" x:Name = 'graphButton' Width = 'Auto' 
                ToolTip = 'This will show telemetry graphs.'>
                    <Image Source = '$pwd\Images\chart.png' /> 
                </Button>
                

                <ToolBar Background = 'Silver' Band = '1' BandIndex = '1'>
               
  <StackPanel Orientation="Horizontal">
                      
                   <Label Content="Filter By Server Name :" VerticalAlignment="Center" Width="Auto"/>
                 <TextBox x:Name="InputBox" Height = "40"  Width = '160' ToolTip='You can search Servers by entering Server names here.' />    
    </StackPanel>            

  <Separator Background = 'Black'>
    <Separator.Style>
            <Style TargetType="{x:Type Separator}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show3">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </Separator.Style>

  </Separator>

             
   <StackPanel Orientation="Horizontal">
   
  <Label Content="Filter By Service Name :" VerticalAlignment="Center" Width="Auto">
  <Label.Style>
            <Style TargetType="{x:Type Label}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show3">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </Label.Style>

  </Label> 

 <TextBox x:Name="ServicesInputBox"  Height = "40"  Width = '280' Text="RemoteRegistry"  ToolTip='You can search Stopped Services by entering the service name here.For Mutiple Service ensure Service Names with (,) comma Seperated.'>
 <TextBox.Style>
            <Style TargetType="{x:Type TextBox}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show3">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </TextBox.Style>
</TextBox>
</StackPanel>


<Separator Background = 'Black'>
    <Separator.Style>
            <Style TargetType="{x:Type Separator}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show2">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </Separator.Style>

  </Separator>

   <StackPanel Orientation="Horizontal">
  
   <Label Margin="10"  VerticalAlignment="Top"/>
             <Label Content="Search KB by ID:" VerticalAlignment="Center" Width="Auto">
            <Label.Style>
            <Style TargetType="{x:Type Label}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show2">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </Label.Style>
        
     </Label>
   <TextBox   x:Name="KBInputBox_new" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "40"  Width = '140' Text="KB"  ToolTip='You can search Installed KBs by entering KB Names here'>
    
        <TextBox.Style>
            <Style TargetType="{x:Type TextBox}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show2">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </TextBox.Style>
    </TextBox>

       <Label Margin="10"  VerticalAlignment="Top"/>
             <Label Content="From Date :" VerticalAlignment="Center" Width="Auto" HorizontalAlignment="Left">
            <Label.Style>
            <Style TargetType="{x:Type Label}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show2">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </Label.Style>
       
     </Label>
     <DatePicker x:Name="txtfromdate" HorizontalAlignment="Left" VerticalAlignment="Center" Width="106" >
     
      <DatePicker.Style>
            <Style TargetType="{x:Type DatePicker}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show2">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </DatePicker.Style>
     
     </DatePicker>
      <Label  Content="To Date :" VerticalAlignment="Center" Width="Auto" HorizontalAlignment="Left">
       <Label.Style>
            <Style TargetType="{x:Type Label}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show2">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </Label.Style>
       
     </Label>
     <DatePicker x:Name="txtTodate" HorizontalAlignment="Left" VerticalAlignment="Center" Width="106">
      <DatePicker.Style>
            <Style TargetType="{x:Type DatePicker}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show2">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </DatePicker.Style>
     
     </DatePicker>
     
      </StackPanel>
    
      
 

   <StackPanel>

    <CheckBox Content="SQL Stay Current" x:Name="sqlstaycurrentCheckBox">
    
        <CheckBox.Style>
            <Style TargetType="{x:Type CheckBox}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </CheckBox.Style>
    </CheckBox>


    <CheckBox Content="SQL Custom Patch" x:Name="sqlcustomCheckBox">

   
        <CheckBox.Style>
            <Style TargetType="{x:Type CheckBox}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </CheckBox.Style>
    </CheckBox>

    <CheckBox Content="Copy Package to Destination Machine ?" x:Name="sqlcopyCheckBox">

   
        <CheckBox.Style>
            <Style TargetType="{x:Type CheckBox}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}"  Value="Show">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </CheckBox.Style>
    </CheckBox>


</StackPanel>
              
  <StackPanel Orientation="Horizontal">
    
    <TextBox ToolTip='Please Provide Custom Patch .exe path.'  MinWidth="220" MinHeight="40" Text="{Binding Comment, UpdateSourceTrigger=PropertyChanged}" x:Name="CommentTextBox">
        <TextBox.Style>
            <Style TargetType="{x:Type TextBox}">
                <Setter Property="Visibility" Value="Hidden"/>
                <Style.Triggers>
                    <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}" Value="Show">
                        <Setter Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>
        </TextBox.Style>
    </TextBox>
</StackPanel>


 <Button Cursor="Hand" x:Name = 'patchButton' Width = 'Auto' 
                ToolTip = 'Upload the custom patch path'>
                <Image Source = '$pwd\Images\BrowseFile.png' /> 

                    <Button.Style>

                    <Style TargetType="{x:Type Button}">
                        <Setter Property="Visibility" Value="Hidden"/>
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding ElementName=RunOptionComboBox, Path=SelectedItem.Tag}" Value="Show">
                                <Setter Property="Visibility" Value="Visible"/>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </Button.Style>

                </Button>

</ToolBar>
  
            </ToolBar>
        </ToolBarTray>
        <Grid VerticalAlignment="Stretch" Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
            <Grid.Resources>
                <Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
                    <Setter Property="Background" Value="#b9c3c9"/>
                    <Setter Property="Foreground" Value="Black"/>
                    <Style.Triggers>
                        <Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
                            <Setter Property="Background" Value="white"/>
                            <Setter Property="Foreground" Value="Black"/>                                
                        </Trigger>                            
                    </Style.Triggers>
                </Style>                    
            </Grid.Resources>                  
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = '*'/>
                <RowDefinition Height = '*'/>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = 'Auto'/>
            </Grid.RowDefinitions> 
            <GroupBox Header = "Server List" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
                <Grid VerticalAlignment="Stretch" Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
                
                <ListView x:Name = 'Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
                ToolTip = 'Server List that displays all information regarding statuses of servers and patches.'>
                <ListView.Resources>
                <DataTemplate x:Key="Templ">
                <TextBlock HorizontalAlignment="Left" Text="{Binding}"/>
                </DataTemplate>
                

                
                <Style x:Key="HeaderStyle" TargetType="GridViewColumnHeader">
                <Setter Property="HorizontalContentAlignment" Value="Left"/>
                <Setter Property="Foreground" Value="darkblue" />
                 <Setter Property="FontSize" Value="10"/>
                 <Setter Property="FontFamily" Value="Verdana"/>

                </Style>


                </ListView.Resources>

                    <ListView.View>
                        <GridView  x:Name = 'GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
                            <GridViewColumn HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'ComputerColumn' Width = '110' DisplayMemberBinding = '{Binding Path = Computer}' Header='Computer'/>
                            <GridViewColumn Width="50" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'AuditedColumn'   DisplayMemberBinding = '{Binding Path = Audited}' Header='Audited'/>                    
                            <GridViewColumn Width="50" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'InstalledColumn'   DisplayMemberBinding = '{Binding Path = Installed}' Header='Installed' />                    
                            <GridViewColumn Width="70" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'InstallErrorColumn'  DisplayMemberBinding = '{Binding Path = InstallErrors}' Header='InstallErrors'/>
                            <GridViewColumn Width="40" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'OfflineColumn'  DisplayMemberBinding = '{Binding Path = Offline}' Header='Offline'/>
                            <GridViewColumn Width="75" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'UnAuthorizedColumn'   DisplayMemberBinding = '{Binding Path = UnAuthorized}' Header='UnAuthorized'/>
                            <GridViewColumn Width="95" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'ServicesColumn'   DisplayMemberBinding = '{Binding Path = Services}' Header='StoppedServices'/>                                                
                            <GridViewColumn Width="85" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'SQLJobsColumn'   DisplayMemberBinding = '{Binding Path = SQLJobs}' Header='Total SQL Jobs'/>                                                
                            <GridViewColumn Width="100" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'SQLJobsRunningColumn'   DisplayMemberBinding = '{Binding Path = RunningSQLJobs}' Header='Running SQL Jobs'/>                                                
                            <GridViewColumn Width="90" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'LastRebootColumn'  DisplayMemberBinding = '{Binding Path = LastRebootTime}' Header='LastRebootTime'/>
                            <GridViewColumn Width="125" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'LastInstallColumn'   DisplayMemberBinding = '{Binding Path = LastPatchInstalledTime}' Header='LastPatchInstalledTime'/>
                            <GridViewColumn Width="170" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'OsVersionColumn'   DisplayMemberBinding = '{Binding Path = OsVersion}' Header='OsVersion'/>
                            <GridViewColumn Width="170" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'SqlVersionColumn'   DisplayMemberBinding = '{Binding Path = SqlVersion}' Header='SqlVersion'/>
                            <GridViewColumn Width="110" HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}"  x:Name = 'StatVersionColumn'   DisplayMemberBinding = '{Binding Path = StatusTxt}' Header='Status'/>
                            
                   

                   
                    <GridViewColumn HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}" x:Name = 'StatusColumn'>
                            <GridViewColumnHeader Content="" Width="30"/>
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <Image Source="{Binding Path=Status}"  />
                                                       
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>



                    <GridViewColumn HeaderTemplate="{StaticResource Templ}" HeaderContainerStyle="{StaticResource HeaderStyle}" x:Name = 'NotesColumn'>
                            <GridViewColumnHeader Content="Notes" Width="Auto" MinWidth="280px"/>
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Path=Notes}" TextAlignment="Left" Width="Auto" Foreground="{Binding Path=colorchange}" />
                                                       
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>

                        


                        </GridView>
                    </ListView.View>
                    <ListView.ContextMenu>
                        <ContextMenu x:Name = 'ListViewContextMenu'>
                            <MenuItem x:Name = 'AddServerMenu' Header = 'Add Server' InputGestureText ='Ctrl+S'/>               
                            <MenuItem x:Name = 'RemoveServerMenu' Header = 'Remove Server' InputGestureText ='Ctrl+D'/>
                            <Separator />
                            <MenuItem x:Name = 'WindowsUpdateServiceMenu' Header = 'Windows Update Service' > 
                                <MenuItem x:Name = 'WUStopServiceMenu' Header = 'Stop Service' />
                                <MenuItem x:Name = 'WUStartServiceMenu' Header = 'Start Service' />
                                <MenuItem x:Name = 'WURestartServiceMenu' Header = 'Restart Service' />
                            </MenuItem>                            
                            <MenuItem x:Name = 'WindowsUpdateLogMenu' Header = 'WindowsUpdateLog' > 
                                <MenuItem x:Name = 'EntireLogMenu' Header = 'View Entire Log'/>
                                <MenuItem x:Name = 'Last25LogMenu' Header = 'View Last 25' />
                                <MenuItem x:Name = 'Last50LogMenu' Header = 'View Last 50'/>
                                <MenuItem x:Name = 'Last100LogMenu' Header = 'View Last 100'/>
                            </MenuItem>
                            <MenuItem x:Name = 'InstalledUpdatesMenu' Header = 'Installed Updates' >
                                <MenuItem x:Name = 'GUIInstalledUpdatesMenu' Header = 'Get Installed Updates'/>
                            </MenuItem>
                        </ContextMenu>
                    </ListView.ContextMenu>            
                </ListView>                
                </Grid>
            </GroupBox>                                    
        </Grid>        
        <ProgressBar x:Name = 'ProgressBar' Grid.Row = '3' Height = '20' ToolTip = 'Displays progress of current action via a graphical progress bar.'/>   
        <TextBox x:Name = 'StatusTextBox' Grid.Row = '4' ToolTip = 'Displays current status of operation'> Waiting for Action... </TextBox>                           
    </Grid>   
</Window>
"@ 

#region Load XAML into PowerShell
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$uiHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
#endregion

 
 
  
#region Background runspace to clean up jobs
$jobCleanup.Flag = $True
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"          
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)          
$newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)     
$newRunspace.SessionStateProxy.SetVariable("jobs",$jobs) 
$jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
    #Routine to handle completed runspaces
    Do {    
        Foreach($runspace in $jobs) {
            If ($runspace.Runspace.isCompleted) {
                $runspace.powershell.EndInvoke($runspace.Runspace) | Out-Null
                $runspace.powershell.dispose()
                $runspace.Runspace = $null
                $runspace.powershell = $null               
            } 
        }
        #Clean out unused runspace jobs
        $temphash = $jobs.clone()
        $temphash | Where {
            $_.runspace -eq $Null
        } | ForEach {
            Write-Verbose ("Removing {0}" -f $_.computer)
            $jobs.remove($_)
        }        
        Start-Sleep -Seconds 1     
    } while ($jobCleanup.Flag)
})
$jobCleanup.PowerShell.Runspace = $newRunspace
$jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  
#endregion

#region Connect to all controls
$uiHash.GenerateReportMenu = $uiHash.Window.FindName("GenerateReportMenu")
 

$uiHash.ClearAuditReportMenu = $uiHash.Window.FindName("ClearAuditReportMenu")
$uiHash.ClearInstallReportMenu = $uiHash.Window.FindName("ClearInstallReportMenu")
$uiHash.SelectAllMenu = $uiHash.Window.FindName("SelectAllMenu")
$uiHash.OptionMenu = $uiHash.Window.FindName("OptionMenu")
$uiHash.WUStopServiceMenu = $uiHash.Window.FindName("WUStopServiceMenu")
$uiHash.WUStartServiceMenu = $uiHash.Window.FindName("WUStartServiceMenu")
$uiHash.WURestartServiceMenu = $uiHash.Window.FindName("WURestartServiceMenu")
$uiHash.WindowsUpdateServiceMenu = $uiHash.Window.FindName("WindowsUpdateServiceMenu")
$uiHash.GenerateReportButton = $uiHash.Window.FindName("GenerateReportButton")
$uiHash.ReportComboBox = $uiHash.Window.FindName("ReportComboBox")
$uiHash.StartImage = $uiHash.Window.FindName("StartImage")
$uiHash.CancelImage = $uiHash.Window.FindName("CancelImage")
$uiHash.RunOptionComboBox = $uiHash.Window.FindName("RunOptionComboBox")
$uiHash.filterComboBox = $uiHash.Window.FindName("filterComboBox")
$uiHash.KBInputBox = $uiHash.Window.FindName("KBInputBox_new")
$uiHash.ServicesInputBox = $uiHash.Window.FindName("ServicesInputBox")
$uiHash.txtfromdate= $uiHash.Window.FindName("txtfromdate")
$uiHash.txtTodate= $uiHash.Window.FindName("txtTodate")


 

$uiHash.ClearErrorMenu = $uiHash.Window.FindName("ClearErrorMenu")
$uiHash.ViewErrorMenu = $uiHash.Window.FindName("ViewErrorMenu")
$uiHash.EntireLogMenu = $uiHash.Window.FindName("EntireLogMenu")
$uiHash.Last25LogMenu = $uiHash.Window.FindName("Last25LogMenu")
$uiHash.Last50LogMenu = $uiHash.Window.FindName("Last50LogMenu")
$uiHash.Last100LogMenu = $uiHash.Window.FindName("Last100LogMenu")
$uiHash.ResetDataMenu = $uiHash.Window.FindName("ResetDataMenu")
$uiHash.ResetAuthorizationMenu = $uiHash.Window.FindName("ResetAuthorizationMenu")
$uiHash.ClearServerListNotesMenu = $uiHash.Window.FindName("ClearServerListNotesMenu")
$uiHash.ServerListReportMenu = $uiHash.Window.FindName("ServerListReportMenu")
$uiHash.OfflineHostsMenu = $uiHash.Window.FindName("OfflineHostsMenu")
$uiHash.RebootHostsMenu = $uiHash.Window.FindName("RebootHostsMenu")
$uiHash.UnauthorizedHostsMenu = $uiHash.Window.FindName("UnauthorizedHostsMenu")
$uiHash.UpdatesNeededHostsMenu = $uiHash.Window.FindName("UpdatesNeededHostsMenu")


$uiHash.HostListMenu = $uiHash.Window.FindName("HostListMenu")
$uiHash.InstalledUpdatesMenu = $uiHash.Window.FindName("InstalledUpdatesMenu")
$uiHash.DetectNowMenu = $uiHash.Window.FindName("DetectNowMenu")
$uiHash.WindowsUpdateLogMenu = $uiHash.Window.FindName("WindowsUpdateLogMenu")
#$uiHash.WUAUCLTMenu = $uiHash.Window.FindName("WUAUCLTMenu")
$uiHash.GUIInstalledUpdatesMenu = $uiHash.Window.FindName("GUIInstalledUpdatesMenu")
$uiHash.AddServerMenu = $uiHash.Window.FindName("AddServerMenu")
$uiHash.RemoveServerMenu = $uiHash.Window.FindName("RemoveServerMenu")
$uiHash.ListviewContextMenu = $uiHash.Window.FindName("ListViewContextMenu")
$uiHash.ExitMenu = $uiHash.Window.FindName("ExitMenu")
$uiHash.ClearInstalledUpdateMenu = $uiHash.Window.FindName("ClearInstalledUpdateMenu")
$uiHash.RunMenu = $uiHash.Window.FindName('RunMenu')
$uiHash.ClearAllMenu = $uiHash.Window.FindName("ClearAllMenu")
$uiHash.ClearServerListMenu = $uiHash.Window.FindName("ClearServerListMenu")
$uiHash.AboutMenu = $uiHash.Window.FindName("AboutMenu")
$uiHash.HelpFileMenu = $uiHash.Window.FindName("HelpFileMenu")
$uiHash.Listview = $uiHash.Window.FindName("Listview")
$uiHash.LoadFileButton = $uiHash.Window.FindName("LoadFileButton")
$uiHash.BrowseFileButton = $uiHash.Window.FindName("BrowseFileButton")
$uiHash.patchButton= $uiHash.Window.FindName("patchButton")
$uiHash.CommentTextBox= $uiHash.Window.FindName("CommentTextBox")
$uiHash.sqlstaycurrentCheckBox= $uiHash.Window.FindName("sqlstaycurrentCheckBox")
$uiHash.sqlcustomCheckBox= $uiHash.Window.FindName("sqlcustomCheckBox")
$uiHash.sqlcopyCheckBox= $uiHash.Window.FindName("sqlcopyCheckBox")
$uiHash.FilterButton = $uiHash.Window.FindName("FilterButton")
$uiHash.ExcelButton = $uiHash.Window.FindName("ExcelButton")
$uiHash.ClearButton = $uiHash.Window.FindName("ClearButton")
$uiHash.resetButton = $uiHash.Window.FindName("resetButton")
$uiHash.graphButton = $uiHash.Window.FindName("graphButton")
$uiHash.LoadADButton = $uiHash.Window.FindName("LoadADButton")
$uiHash.StatusTextBox = $uiHash.Window.FindName("StatusTextBox")
$uiHash.ProgressBar = $uiHash.Window.FindName("ProgressBar")
$uiHash.RunButton = $uiHash.Window.FindName("RunButton")
$uiHash.CancelButton = $uiHash.Window.FindName("CancelButton")
$uiHash.GridView = $uiHash.Window.FindName("GridView")
$uiHash.InputBox = $uiHash.Window.FindName('InputBox')




 #endregion
 #region Event Handlers 

#Window Load Events
$uiHash.Window.Add_SourceInitialized({
    #Configure Options
    Write-Verbose 'Updating configuration based on options'
    Set-PoshPAIGOption 
    Write-Debug ("maxConcurrentJobs: {0}" -f $maxConcurrentJobs)
    Write-Debug ("MaxRebootJobs: {0}" -f $MaxRebootJobs)
    Write-Debug ("reportpath: {0}" -f $reportpath)
    Write-Debug ("DiagPath: {0}" -f $DiagPath)

    #Define hashtable of settings
    $Script:SortHash = @{}


     
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.Listview.Items.Refresh()   
            }             
        }
    }
 
   
   $uiHash.Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ColumnSortHandler)
    
         
    #Create and bind the observable collection to the GridView
    $Script:clientObservable = New-Object System.Collections.ObjectModel.ObservableCollection[object]    
    $uiHash.ListView.ItemsSource = $clientObservable
    $Global:Clients = $clientObservable | Select -Expand Computer
})    


 $uiHash.InputBox.Add_TextChanged({
     try{
    
     $view = [System.Windows.Data.CollectionViewSource]:: GetDefaultView($clientObservable)
     #
        $filter =  $uiHash.InputBox.Text

        $view.Filter = {param ($item) $item -match $filter} 
       
        $view.Refresh()
        $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
        }catch{}
})


#Window Close Events
$uiHash.Window.Add_Closed({
    #Halt job processing
    $jobCleanup.Flag = $False

    #Stop all runspaces
    $jobCleanup.PowerShell.Dispose()
    
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()    
})
 

$uiHash.graphButton.Add_Click({


If ($uiHash.Listview.Items.count -gt 0) {
        #$uiHash.StatusTextBox.Foreground = "Black"
       

         If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}
         }

        $report = $reportpath + "\" + "serverlist.csv"

        $savedreport =  $report
        $uiHash.Listview.ItemsSource | Export-Csv -NoTypeInformation $savedreport
        #$uiHash.StatusTextBox.Text = "Report saved to $savedreport"        
    }

.\Charts.ps1

})
#Cancel Button Event
$uiHash.CancelButton.Add_Click({
    $runspaceHash.runspacepool.Dispose()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Action cancelled" 
    [Float]$uiHash.ProgressBar.Value = 0
    $uiHash.RunButton.IsEnabled = $True
    $uiHash.StartImage.Source = "$pwd\Images\Start.png"
    $uiHash.CancelButton.IsEnabled = $False
    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"    
         
})

#EntireUpdateLog Event
$uiHash.EntireLogMenu.Add_Click({
    If ($uiHash.Listview.Items.count -eq 1) {
        $selectedItem = $uiHash.Listview.SelectedItem
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\stop.png"         
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Retrieving Windows Update log from Server..."            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0

        $uiHash.Listview.Items.EditItem($selectedItem)
        $selectedItem.Notes = "Retrieving Update Log"
        $uiHash.Listview.Items.CommitEdit()
        $uiHash.Listview.Items.Refresh() 
        . .\Scripts\Get-UpdateLog.ps1
        Try {
            $log = Get-UpdateLog -Computername $selectedItem.computer 
            If ($log) {
                $log | Out-GridView -Title ("{0} Update Log" -f $selectedItem.computer)
                $selectedItem.Notes = "Completed"
            }
        } Catch {
            $selectedItem.notes = $_.Exception.Message
        }
        $uiHash.ProgressBar.value++ 
        $End = New-Timespan $uihash.StartTime (Get-Date) 
        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
        $uiHash.RunButton.IsEnabled = $True
        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
        $uiHash.CancelButton.IsEnabled = $False
        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"
    }  
}) 


 
#Last100UpdateLog Event
$uiHash.Last100LogMenu.Add_Click({
    If ($uiHash.Listview.Items.count -eq 1) {
        $selectedItem = $uiHash.Listview.SelectedItem
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\stop.png"         
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Retrieving Windows Update log from Server..."            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0

        $uiHash.Listview.Items.EditItem($selectedItem)
        $selectedItem.Notes = "Retrieving Update Log"
        $uiHash.Listview.Items.CommitEdit()
        $uiHash.Listview.Items.Refresh() 
        . .\Scripts\Get-UpdateLog.ps1
        Try {
            $log = Get-UpdateLog -Last 100 -Computername $selectedItem.computer 
            If ($log) {
                $log | Out-GridView -Title ("{0} Update Log" -f $selectedItem.computer)
                $selectedItem.Notes = "Completed"
            }
        } Catch {
            $selectedItem.notes = $_.Exception.Message
        }
        $uiHash.ProgressBar.value++ 
        $End = New-Timespan $uihash.StartTime (Get-Date) 
        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
        $uiHash.RunButton.IsEnabled = $True
        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
        $uiHash.CancelButton.IsEnabled = $False
        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"
    }       
})

#Last50UpdateLog Event
$uiHash.Last50LogMenu.Add_Click({
    If ($uiHash.Listview.Items.count -eq 1) {
        $selectedItem = $uiHash.Listview.SelectedItem
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\stop.png"         
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Retrieving Windows Update log from Server..."            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0

        $uiHash.Listview.Items.EditItem($selectedItem)
        $selectedItem.Notes = "Retrieving Update Log"
        $uiHash.Listview.Items.CommitEdit()
        $uiHash.Listview.Items.Refresh() 
        . .\Scripts\Get-UpdateLog.ps1
        Try {
            $log = Get-UpdateLog -Last 50 -Computername $selectedItem.computer 
            If ($log) {
                $log | Out-GridView -Title ("{0} Update Log" -f $selectedItem.computer)
                $selectedItem.Notes = "Completed"
            }
        } Catch {
            $selectedItem.notes = $_.Exception.Message
        }
        $uiHash.ProgressBar.value++ 
        $End = New-Timespan $uihash.StartTime (Get-Date) 
        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
        $uiHash.RunButton.IsEnabled = $True
        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
        $uiHash.CancelButton.IsEnabled = $False
        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"
    }                  
})

#Last25UpdateLog Event
$uiHash.Last25LogMenu.Add_Click({
    If ($uiHash.Listview.Items.count -eq 1) {
        $selectedItem = $uiHash.Listview.SelectedItem
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\stop.png"         
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Retrieving Windows Update log from Server..."            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0

        $uiHash.Listview.Items.EditItem($selectedItem)
        $selectedItem.Notes = "Retrieving Update Log"
        $uiHash.Listview.Items.CommitEdit()
        $uiHash.Listview.Items.Refresh() 
        . .\Scripts\Get-UpdateLog.ps1
        Try {
            $log = Get-UpdateLog -Last 25 -Computername $selectedItem.computer 
            If ($log) {
                $log | Out-GridView -Title ("{0} Update Log" -f $selectedItem.computer)
                $selectedItem.Notes = "Completed"
            }
        } Catch {
            $selectedItem.notes = $_.Exception.Message
        }
        $uiHash.ProgressBar.value++ 
        $End = New-Timespan $uihash.StartTime (Get-Date) 
        $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
        $uiHash.RunButton.IsEnabled = $True
        $uiHash.StartImage.Source = "$pwd\Images\Start.png"
        $uiHash.CancelButton.IsEnabled = $False
        $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"             
    }    
})

#Offline server removal
$uiHash.OfflineHostsMenu.Add_Click({
    Write-Verbose "Removing any server that is shown as Offline"
    $Offline = @($uiHash.Listview.Items | Where {$_.Notes -eq "Offline"})
    $Offline | ForEach {
        Write-Verbose ("Removing {0}" -f $_.Computer)
        $clientObservable.Remove($_)
        }
    $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
})


#Offline server removal
$uiHash.RebootHostsMenu.Add_Click({
    Write-Verbose "Filtering Reboot Needed Servers"
    $reboot = @($uiHash.Listview.Items | Where {$_.Notes -ne "Reboot is required"})
    $reboot | ForEach {
        Write-Verbose ("Listing Reboot {0}" -f $_.Computer)
        $clientObservable.Remove($_)
        }
    $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
})

 


$uiHash.UnauthorizedHostsMenu.Add_Click({

    Write-Verbose "Filtering UnAuthorized Servers"
    $reboot = @($uiHash.Listview.Items | Where {$_.Notes -ne "UnAuthorized Access"})
    $reboot | ForEach {
        Write-Verbose ("Listing Unauthorized {0}" -f $_.Computer)
        $clientObservable.Remove($_)
        }
    $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
})

$uiHash.UpdatesNeededHostsMenu.Add_Click({
    Write-Verbose "Filtering Update Needed Servers"
    $reboot = @($uiHash.Listview.Items | Where {($_.Notes -ne "Updates Found and Downloaded") -and ($_.Notes -ne "Updates Found and Not Downloaded")})
    $reboot | ForEach {
        Write-Verbose ("Listing Update Needed Servers {0}" -f $_.Computer)
        $clientObservable.Remove($_)
        }
    $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
})

 
 
#Rightclick Event
$uiHash.Listview.Add_MouseRightButtonUp({
    Write-Debug "$($This.SelectedItem.Row.Computer)"
    If ($uiHash.Listview.SelectedItems.count -eq 0) {
        $uiHash.RemoveServerMenu.IsEnabled = $False
        $uiHash.InstalledUpdatesMenu.IsEnabled = $False
        $uiHash.WindowsUpdateLogMenu.IsEnabled = $False
        $uiHash.WindowsUpdateServiceMenu.IsEnabled = $False
        #$uiHash.WUAUCLTMenu.IsEnabled = $False
        } ElseIf ($uiHash.Listview.SelectedItems.count -eq 1) {
        $uiHash.RemoveServerMenu.IsEnabled = $True
        $uiHash.InstalledUpdatesMenu.IsEnabled = $True
        $uiHash.WindowsUpdateLogMenu.IsEnabled = $True
        $uiHash.WindowsUpdateServiceMenu.IsEnabled = $True
        #$uiHash.WUAUCLTMenu.IsEnabled = $True      
        } Else {
        $uiHash.RemoveServerMenu.IsEnabled = $True
        $uiHash.InstalledUpdatesMenu.IsEnabled = $True
        $uiHash.WindowsUpdateLogMenu.IsEnabled = $False
        #$uiHash.WUAUCLTMenu.IsEnabled = $True     
    }    
})

#ListView drop file Event
$uiHash.Listview.add_Drop({
    $content = Get-Content $_.Data.GetFileDropList()
    $content | ForEach {
        $clientObservable.Add((
            New-Object PSObject -Property @{
                Computer = $_
                Audited = 0 -as [int]
                Installed = 0 -as [int]
                InstallErrors = 0 -as [int]
                Offline = 0 -as [int]
                UnAuthorized=0 -as [int]
                LastRebootTime = "" -as [string]
                LastPatchInstalledTime = "" -as [string]
                OsVersion = "" -as [string]
                SqlVersion = "" -as [string]
                Services = 0 -as [int]
                SQLJobs=0 -as [int]
                RunningSQLJobs=0 -as [int]
                StatusTxt="" -as [string]
                Notes = $Null
                Status=$Null
                colorchange='Black' 
            }
        ))      
    }
    Show-DebugState
})


$uiHash.patchButton.Add_Click({
$File =Open-FileDialogexe
$uiHash.CommentTextBox.Text=$File

})

#FindFile Button
$uiHash.BrowseFileButton.Add_Click({
    $File = Open-FileDialog
    If (-Not ([system.string]::IsNullOrEmpty($File))) {
        
      
      $unique= Get-Content $File | Select-Object -Unique
      $unique | Where {$_ -ne ""} | ForEach {
         
             $clientObservable.Add((
                New-Object PSObject -Property @{
                    Computer = $_
                    Audited = 0 -as [int]
                    Installed = 0 -as [int]
                    InstallErrors = 0 -as [int]
                    Offline = 0 -as [int]
                    UnAuthorized=0 -as [int]
                    LastRebootTime = "" -as [string]
                    LastPatchInstalledTime = "" -as [string]
                    OsVersion = "" -as [string]
                    SqlVersion = "" -as [string]
                    Services = 0 -as [int]
                    SQLJobs=0 -as [int]
                    RunningSQLJobs=0 -as [int]
                    StatusTxt="" -as [string]
                    Notes = $Null
                    Status=$Null
                    colorchange='Black'
                }
            ))       
        }

         $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Waiting for action..."
        Show-DebugState     
    }        
})

$uiHash.ClearButton.Add_Click({ 
$clientObservable.Clear()
    $content = $Null
    [Float]$uiHash.ProgressBar.value = 0
    $uiHash.StatusTextBox.Foreground = "Black"
    $Global:updateAudit.Clear()
    $uiHash.StatusTextBox.Text = "Waiting for action..." 

})
$uiHash.resetButton.Add_Click({ 
   $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null;$_.Status = $Null;$_.StatusTxt = "";$_.SqlVersion="";$_.Audited = 0;$_.Installed = 0;$_.InstallErrors = 0;$_.Offline = 0;$_.UnAuthorized = 0;$_.LastRebootTime = "";$_.LastPatchInstalledTime = "";$_.OsVersion = "";$_.Services = 0;$_.SQLJobs = 0;$_.RunningSQLJobs = 0;$_.colorchange='Black';}
    $uiHash.Listview.Items.Refresh() 
})


$uiHash.ExcelButton.Add_Click({ 

If ($uiHash.Listview.Items.count -gt 0) {
        $uiHash.StatusTextBox.Foreground = "Black"
       

         If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}
         }

      $report = $reportpath + "\" + "serverlist.csv"

        $savedreport =  $report
        $uiHash.Listview.ItemsSource | Export-Csv -NoTypeInformation $savedreport
        $uiHash.StatusTextBox.Text = "Report saved to $savedreport"        
    } Else {
        $uiHash.StatusTextBox.Foreground = "Red"
        $uiHash.StatusTextBox.Text = "No report to create!"         
    }   

})
#Filter
$uiHash.FilterButton.Add_Click({
     
      

      
if($uiHash.filterComboBox.Text -ne 'Select All'){
    $updfound = @($uiHash.Listview.Items | Where {($_.Notes -eq $uiHash.filterComboBox.Text)})
         
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($clientObservable)
        $filter = $uiHash.filterComboBox.Text

        $view.Filter = {param ($item) $item -match $filter}
       
        $view.Refresh()
        $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count
        }
     elseif($uiHash.filterComboBox.Text -eq 'Select All')
     {
        $selectall = @($uiHash.Listview.Items)
         
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($clientObservable)
        $filter = ""
        $view.Filter = {param ($item) $item -match $filter}
        $view.Refresh()
        $uiHash.ProgressBar.Maximum = $uiHash.Listview.ItemsSource.count

     }
     
             
})


#LoadADButton Events    
 #RunButton Events    
$uiHash.RunButton.add_Click({
        

    
 if (Test-Path  .\scripts\KBListFile.txt) 
    {
        Remove-Item .\scripts\KBListFile.txt
    }

    $getKBValues=$uiHash.KBInputBox.Text
    $fromdate=$uiHash.txtfromdate.Text
    $todate=$uiHash.txtTodate.Text
    $splitKB=$getKBValues -split ","
    $madaterrror=$false

    if(@($splitKB).Count -eq 1 -and @($splitKB).Length -le 2 -and $uiHash.RunOptionComboBox.Text.trim() -eq "Search Installed KBs")
    {
    $ListAdd = @()
            if($fromdate -eq "")
            {
            Write-Host "[Warning]: From Date is Mandatory for KB search"
            $madaterrror=$true

            }
            elseif($todate -eq "")
            {
            Write-Host "[Warning]: To Date is Mandatory for KB search"
            $madaterrror=$true
            }
        else{
            
            $fromdate = [datetime]::ParseExact($fromdate, "M/dd/yyyy", $null)
            $fromdate=$fromdate.ToString("dd/MM/yyyy")
             
            
            
            $todate = [datetime]::ParseExact($todate, "M/dd/yyyy", $null)
            $todate=$todate.ToString("dd/MM/yyyy")
             
               
            #Use New-Object and Add-Member to create an object
            $Kb = New-Object System.Object
            $Kb | Add-Member -MemberType NoteProperty -Name "KB" -Value "$getKBValues" 
            $Kb | Add-Member -MemberType NoteProperty -Name "FromDate" -Value "$fromdate"
            $Kb | Add-Member -MemberType NoteProperty -Name "ToDate" -Value "$todate"

            #Add newly created object to the array
            $ListAdd += $Kb
            $outfile=".\scripts\KBListFile.txt"
            $ListAdd |select KB,FromDate,ToDate |Select-Object -Skip 0| Out-File -FilePath  $outfile -Append 
            (gc $outFile) | ? {$_.trim() -ne "" } | select -Skip 2 | set-content $OutFile
        }

    }
    elseif($uiHash.RunOptionComboBox.Text.trim() -eq "Search Installed KBs")
    {
    #Finally, use Export-Csv to export the data to a csv file
 

        foreach($item in $splitKB)
        {
        $item | Out-File -FilePath .\scripts\KBListFile.txt -Append
        }
    }

    

    if(($uiHash.RunOptionComboBox.Text.trim() -eq "Search Installed KBs") -and ($madaterrror -eq $false))
    {
     Start-RunJob 
    }
    elseif($uiHash.RunOptionComboBox.Text.trim() -ne "Search Installed KBs"){
    Start-RunJob 
    
    }
})



#region Client WSUS Service Action
#Stop Service
$uiHash.WUStopServiceMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $Servers = $uiHash.Listview.SelectedItems | Select -ExpandProperty Computer
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\stop.png"
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Stopping WSUS Client service on selected servers"            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0
        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Stopping Update Client Service"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            Try {
                $updateClient = Get-Service -ComputerName $computer.computer -Name wuauserv -ErrorAction Stop
                Stop-Service -inputObject $updateClient -ErrorAction Stop
                $result = $True
            } Catch {
                $updateClient = $_.Exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $Computer.Notes = "Service Stopped"
                } Else {
                    $computer.notes = ("Issue Occurred: {0}" -f $updateClient)
                } 
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
            })
            $uiHash.ProgressBar.Dispatcher.Invoke("Background",[action]{   
                $uiHash.ProgressBar.value++  
            })
            $uiHash.Window.Dispatcher.Invoke("Background",[action]{                       
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"   
                                      
                }
            })  
                
        }

        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            $uiHash.Listview.Items.EditItem($Computer)
            $computer.Notes = "Pending Stop Service"
            $uiHash.Listview.Items.CommitEdit()
            $uiHash.Listview.Items.Refresh() 
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null
        }     
    }    
})

#Start Service
$uiHash.WUStartServiceMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $Servers = $uiHash.Listview.SelectedItems | Select -ExpandProperty Computer
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\stop.png"
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Starting WSUS Client service on selected servers"            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0
        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Starting Update Client Service"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            Try {
                $updateClient = Get-Service -ComputerName $computer.computer -Name wuauserv -ErrorAction Stop
                Start-Service -inputObject $updateClient -ErrorAction Stop
                $result = $True
            } Catch {
                $updateClient = $_.Exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $Computer.Notes = "Service Started"
                } Else {
                    $computer.notes = ("Issue Occurred: {0}" -f $updateClient)
                } 
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
            })
            $uiHash.ProgressBar.Dispatcher.Invoke("Background",[action]{   
                $uiHash.ProgressBar.value++  
            })
            $uiHash.Window.Dispatcher.Invoke("Background",[action]{           
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"   
                                      
                }
            })  
                
        }

        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            $uiHash.Listview.Items.EditItem($Computer)
            $computer.Notes = "Pending Start Service"
            $uiHash.Listview.Items.CommitEdit()
            $uiHash.Listview.Items.Refresh()
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null
        } 
    
    }    
})

#Restart Service
$uiHash.WURestartServiceMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $Servers = $uiHash.Listview.SelectedItems | Select -ExpandProperty Computer
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\stop.png"
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Restarting WSUS Client service on selected servers"            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0
        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Restarting Update Client Service"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            Try {
                $updateClient = Get-Service -ComputerName $computer.computer -Name wuauserv -ErrorAction Stop
                Restart-Service -inputObject $updateClient -ErrorAction Stop
                $result = $True
            } Catch {
                $updateClient = $_.Exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $Computer.Notes = "Service Restarted"
                } Else {
                    $computer.notes = ("Issue Occurred: {0}" -f $updateClient)
                } 
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
            })
            $uiHash.ProgressBar.Dispatcher.Invoke("Background",[action]{   
                $uiHash.ProgressBar.value++  
            })
            $uiHash.Window.Dispatcher.Invoke("Background",[action]{           
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png" 
                                        
                }
            })  
                
        }

        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            $uiHash.Listview.Items.EditItem($Computer)
            $computer.Notes = "Pending Restart Service"
            $uiHash.Listview.Items.CommitEdit()
            $uiHash.Listview.Items.Refresh()
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null
        } 
    
    }    
})
#endregion

#View Installed Update Event
$uiHash.GUIInstalledUpdatesMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $selectedItems = $uiHash.Listview.SelectedItems
        [Float]$uiHash.ProgressBar.Maximum = $servers.count
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        $uiHash.RunButton.IsEnabled = $False
        $uiHash.StartImage.Source = "$pwd\Images\Start_locked.png"
        $uiHash.CancelButton.IsEnabled = $True
        $uiHash.CancelImage.Source = "$pwd\Images\stop.png"        
        $uiHash.StatusTextBox.Foreground = "Black"
        $uiHash.StatusTextBox.Text = "Gathering all installed patches on Servers"            
        $uiHash.StartTime = (Get-Date)        
        [Float]$uiHash.ProgressBar.Value = 0             

        $scriptBlock = {
            Param (
                $Computer,
                $uiHash,
                $Path,
                $installedUpdates
            )               
            $uiHash.ListView.Dispatcher.Invoke("Normal",[action]{ 
                $uiHash.Listview.Items.EditItem($Computer)
                $computer.Notes = "Querying Installed Updates"
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh() 
            })                
            Set-Location $Path
            Try {
                $updates = Get-HotFix -ComputerName $computer.computer -ErrorAction Stop | Where {$_.Description -ne ""}
                If ($updates) {
                    $installedUpdates.AddRange($updates) | Out-Null
                }
            } Catch {
                $result = $_.exception.Message
            }
            $uiHash.ListView.Dispatcher.Invoke("Background",[action]{                    
                $uiHash.Listview.Items.EditItem($Computer)
                If ($result) {
                    $computer.Notes = $result
                } Else {
                    $computer.Notes = "Completed"
                }
                $uiHash.Listview.Items.CommitEdit()
                $uiHash.Listview.Items.Refresh()
            })
            $uiHash.ProgressBar.Dispatcher.Invoke("Background",[action]{   
                $uiHash.ProgressBar.value++  
            })
            $uiHash.Window.Dispatcher.Invoke("Background",[action]{           
                #Check to see if find job
                If ($uiHash.ProgressBar.value -eq $uiHash.ProgressBar.Maximum) { 
                    $End = New-Timespan $uihash.StartTime (Get-Date)                                             
                    $uiHash.StatusTextBox.Text = ("Completed in {0}" -f $end)
                    $uiHash.RunButton.IsEnabled = $True
                    $uiHash.StartImage.Source = "$pwd\Images\Start.png"
                    $uiHash.CancelButton.IsEnabled = $False
                    $uiHash.CancelImage.Source = "$pwd\Images\Stop_locked.png"     
                                 
                }
            })  
                
        }
        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspaceHash.runspacepool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrentJobs, $sessionstate, $Host)
        $runspaceHash.runspacepool.Open()   

        ForEach ($Computer in $selectedItems) {
            #Create the powershell instance and supply the scriptblock with the other parameters 
            $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($uiHash).AddArgument($Path).AddArgument($installedUpdates)
           
            #Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspaceHash.runspacepool
           
            #Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $Temp.Computer = $Computer.computer
            $temp.PowerShell = $powershell
           
            #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $jobs.Add($temp) | Out-Null                
        }
    }
})

#ClearAuditReportMenu Events    
$uiHash.ClearAuditReportMenu.Add_Click({
    $updateAudit.Clear()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Audit Report Cleared!"  
})

#ClearInstallReportMenu Events    
$uiHash.ClearInstallReportMenu.Add_Click({
    Remove-Variable InstallPatchReport -scope Global -force -ea 'silentlycontinue'
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Install Report Cleared!"  
})

#ClearInstalledUpdateMenu
$uiHash.ClearInstalledUpdateMenu.Add_Click({
    Remove-Variable InstalledPatches -scope Global -force -ea 'silentlycontinue'
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Installed Updates Report Cleared!"    
})
    
#ClearServerListMenu Events    
$uiHash.ClearServerListMenu.Add_Click({
    $clientObservable.Clear()
    $uiHash.StatusTextBox.Foreground = "Black"
    $uiHash.StatusTextBox.Text = "Server List Cleared!"  
})    

#AboutMenu Event
$uiHash.AboutMenu.Add_Click({
    Open-PoshPAIGAbout
})

#Options Menu
$uiHash.OptionMenu.Add_Click({
    #Launch options window
    Write-Verbose "Launching Options Menu"
    .\Options.ps1
    #Process Updates Options
    Set-PoshPAIGOption    
})

#Select All
$uiHash.SelectAllMenu.Add_Click({
    $uiHash.Listview.SelectAll()
})

#HelpFileMenu Event
$uiHash.HelpFileMenu.Add_Click({
    Open-PoshPAIGHelp
})

#KeyDown Event
$uiHash.Window.Add_KeyDown({ 
    $key = $_.Key  
    If ([System.Windows.Input.Keyboard]::IsKeyDown("RightCtrl") -OR [System.Windows.Input.Keyboard]::IsKeyDown("LeftCtrl")) {
        Switch ($Key) {
        "E" {$This.Close()}
        "A" {$uiHash.Listview.SelectAll()}
        "O" {
            .\Options.ps1
            #Process Updates Options
            Set-PoshPAIGOption
        }
        "S" {Add-Server}
        "D" {Remove-Server}
        Default {$Null}
        }
    } ElseIf ([System.Windows.Input.Keyboard]::IsKeyDown("LeftShift") -OR [System.Windows.Input.Keyboard]::IsKeyDown("RightShift")) {
        Switch ($Key) {
            "RETURN" {Write-Host "Hit Shift+Return"}
        }
    }   

})

#Key Up Event
$uiHash.Window.Add_KeyUp({
    $Global:Test = $_
    Write-Debug ("Key Pressed: {0}" -f $_.Key)
    Switch ($_.Key) {
        "F1" {Open-PoshPAIGHelp}
        "F5" {Start-RunJob}
        "F8" {Start-Report}
        Default {$Null}
    }

})

#AddServer Menu
$uiHash.AddServerMenu.Add_Click({
    Add-Server   
})

#RemoveServer Menu
$uiHash.RemoveServerMenu.Add_Click({
    Remove-Server 
})  

#Run Menu
$uiHash.RunMenu.Add_Click({
    Start-RunJob
})      
      
#Report Menu
$uiHash.GenerateReportMenu.Add_Click({
    Start-Report
})       
      
#Exit Menu
$uiHash.ExitMenu.Add_Click({
    $uiHash.Window.Close()
})

#ClearAll Menu
$uiHash.ClearAllMenu.Add_Click({
    $clientObservable.Clear()
    $content = $Null
    [Float]$uiHash.ProgressBar.value = 0
    $uiHash.StatusTextBox.Foreground = "Black"
    $Global:updateAudit.Clear()
    $uiHash.StatusTextBox.Text = "Waiting for action..."    
})

#Clear Server List notes
$uiHash.ClearServerListNotesMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null}
        }
})

#Save Server List
$uiHash.ServerListReportMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) {
        $uiHash.StatusTextBox.Foreground = "Black"
        $savedreport = Join-Path (Join-Path $home Desktop) "serverlist.csv"
        $uiHash.Listview.ItemsSource | Export-Csv -NoTypeInformation $savedreport
        $uiHash.StatusTextBox.Text = "Report saved to $savedreport"        
    } Else {
        $uiHash.StatusTextBox.Foreground = "Red"
        $uiHash.StatusTextBox.Text = "No report to create!"         
    }         
})
     
#HostListMenu
$uiHash.HostListMenu.Add_Click({
    If ($uiHash.Listview.Items.count -gt 0) { 
        $uiHash.StatusTextBox.Foreground = "Black"
        $savedreport = Join-Path $reportpath "hosts.txt"
        $uiHash.Listview.DataContext | Select -Expand Computer | Out-File $savedreport
        $uiHash.StatusTextBox.Text = "Report saved to $savedreport"
        } Else {
        $uiHash.StatusTextBox.Foreground = "Red"
        $uiHash.StatusTextBox.Text = "No report to create!"         
    }         
})     

#Report Generation
$uiHash.GenerateReportButton.Add_Click({
    Start-Report
})

#Clear Error log
$uiHash.ClearErrorMenu.Add_Click({
    Write-Verbose "Clearing error log"
    $Error.Clear()
})

#View Error Event
$uiHash.ViewErrorMenu.Add_Click({
    Get-Error | Out-GridView
})

#ResetServerListData Event
$uiHash.ResetDataMenu.Add_Click({
    Write-Verbose "Resetting Server List data"
    $uiHash.Listview.ItemsSource | ForEach {$_.Notes = $Null;$_.Status = $Null;$_.StatusTxt = "";$_.SqlVersion="";$_.Audited = 0;$_.Installed = 0;$_.InstallErrors = 0;$_.Offline = 0;$_.UnAuthorized = 0;$_.LastRebootTime = "";$_.LastPatchInstalledTime = "";$_.OsVersion = "";$_.Services = 0;$_.SQLJobs = 0;$_.RunningSQLJobs = 0;$_.colorchange='Black';}
})
#endregion        

#Start the GUI
$uiHash.Window.ShowDialog() | Out-Null