$report =@()
[System.Collections.ArrayList]$SingleReport = @()
$Computer = $env:computername
$outfilePath =  "c:\$($Computer)_KBDetailslog.csv"
 
if([System.IO.File]::Exists($outfilePath)){ 

  Remove-Item $outfilePath -force -recurse

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

function Convert-WuaResultCodeToName
{
param( [Parameter(Mandatory=$true)]
[int] $ResultCode
)
$Result = $ResultCode
switch($ResultCode)
{
2
{
$Result = "Succeeded"
}
3
{
$Result = "Succeeded With Errors"
}
4
{
$Result = "Failed"
}
}
return $Result
}


function get-lastinstalledDate
{
 $Retvalue=$false;
 
 # Get a WUA Session
$session = (New-Object -ComObject 'Microsoft.Update.Session')
# Query the latest 1000 History starting with the first recordp
$history = $session.QueryHistory("",0,15000) | ForEach-Object {
$Result = Convert-WuaResultCodeToName -ResultCode $_.ResultCode
# Make the properties hidden in com properties visible.
$_ | Add-Member -MemberType NoteProperty -Value $Result -Name Result
$Product = $_.Categories | Where-Object {$_.Type -eq 'Product'} | Select-Object -First 1 -ExpandProperty Name
$_ | Add-Member -MemberType NoteProperty -Value $_.UpdateIdentity.UpdateId -Name UpdateId
$_ | Add-Member -MemberType NoteProperty -Value $_.UpdateIdentity.RevisionNumber -Name RevisionNumber
$_ | Add-Member -MemberType NoteProperty -Value $Product -Name Product -PassThru
Write-Output $_
}
#Remove null records and only return the fields we want
$histo=$history |
Where-Object {![String]::IsNullOrWhiteSpace($_.title)} |
Select-Object Result, Date, Title, SupportUrl, Product, UpdateId, RevisionNumber


 

$histo| Select-Object -First 1| Sort-Object Date –Descending
 
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

function get-KBInstalled
{ 
param([String] $SearchKbnames,$fromdate,$todate)

 
$Retvalue=$false;


 
 # Get a WUA Session
$session = (New-Object -ComObject 'Microsoft.Update.Session')
# Query the latest 1000 History starting with the first recordp
$history = $session.QueryHistory("",0,15000) | ForEach-Object {
$Result = Convert-WuaResultCodeToName -ResultCode $_.ResultCode
# Make the properties hidden in com properties visible.
$_ | Add-Member -MemberType NoteProperty -Value $Result -Name Result
$Product = $_.Categories | Where-Object {$_.Type -eq 'Product'} | Select-Object -First 1 -ExpandProperty Name
$_ | Add-Member -MemberType NoteProperty -Value $_.UpdateIdentity.UpdateId -Name UpdateId
$_ | Add-Member -MemberType NoteProperty -Value $_.UpdateIdentity.RevisionNumber -Name RevisionNumber
$_ | Add-Member -MemberType NoteProperty -Value $Product -Name Product -PassThru
Write-Output $_
}
#Remove null records and only return the fields we want
$histo=$history |
Where-Object {![String]::IsNullOrWhiteSpace($_.title)} |
Select-Object Result, Date, Title, SupportUrl, Product, UpdateId, RevisionNumber



 if($fromdate -ne "" -and $Todate -ne "")
 {
 $fromdate = [datetime]::parseexact($fromdate, 'dd/MM/yyyy', $null)
 $Todate =[datetime]::parseexact($todate, 'dd/MM/yyyy', $null)
 $histo =  $histo | where {$_.Title -like "*$SearchKbnames*" -and $_.Date -gt $fromdate -and $_.Date -le $Todate} 
 }


$checkKB = @($histo |  Select-String -Pattern $SearchKbnames -CaseSensitive).Count
$checkKBDetails = $histo |  Select-String -Pattern $SearchKbnames -CaseSensitive


$sOS=(Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Name
$sOS= $sOS.Substring(0,$sOS.IndexOf("|"))
$LASTREBOOT = Get-CimInstance -ComputerName $Computer -ClassName win32_operatingsystem | select csname, lastbootuptime
$lastboouptime=$LASTREBOOT.lastbootuptime
$date=Get-Date
$lastpatchinstalled=get-lastinstalledDate
$checksql=checkSqlServerExists -computername  $Computer
 
if($checksql)
{
$sql="Installed"
}
else{
$sql="Not Installed"
}

if($checkKB -gt 0)
{
    foreach($item in $histo)
    {
    $title=$item.Title

    $iteminstalleddate = $item.Date
    $report=set_objects -Computer $Computer -Title "$title KB is Installed" -KB "$title" -IsDownloaded "NA" -SqlVersion $sql  -Notes "$title is Installed" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $iteminstalleddate 
    $SingleReport.Add($report) | Out-Null  

    }

}
else{

 $report=set_objects -Computer $Computer -Title "$SearchKbnames KB is not Installed" -KB "$SearchKbnames" -IsDownloaded "NA" -SqlVersion $sql -Notes "$SearchKbnames is not Installed" -OSVersion $sOS -LastScanTime $date -LastRebootTime $lastboouptime -LastPatchInstalledTime $lastpatchinstalled.Date 
 $SingleReport.Add($report) | Out-Null  
 
}

 
                   

}



if(Test-Path ("C:\KBListFile.txt"))
{
$getKB= Get-Content -Path "C:\KBListFile.txt"
}

if (@($getKB).Count -gt 1)
{
    foreach($itemkb in $getKB)
    {
    get-KBInstalled -SearchKbname $itemkb -fromdate "" -todate ""
    }

}
else{


foreach ($line in $getKB) {
    $fields = $line.Split(" ")
    get-KBInstalled -SearchKbname $fields[0] -fromdate $fields[1] -todate $fields[2]
      
}


}

 if($SingleReport.Count -gt 0)
 {
 $SingleReport|Export-CSV $outfilePath -noTypeInformation -Append
 }
