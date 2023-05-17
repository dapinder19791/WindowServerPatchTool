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
$hist=$ifexists1 | Select-Object -First 1| Sort-Object InstallOn –Descending
}
elseif($ifexists2.count -gt 0)
{
$hist=$ifexists2 | Select-Object -First 1| Sort-Object InstallOn –Descending
}
elseif($ifexists3.count -gt 0)
{
$hist=$ifexists3 | Select-Object -First 1| Sort-Object InstallOn –Descending
}
elseif($ifexists4.count -gt 0)
{
$hist=$ifexists4 | Select-Object -First 1| Sort-Object InstallOn –Descending
} 
else{
$hist=$history | Select-Object -First 1| Sort-Object InstallOn –Descending

}
  
  Write-Host $hist
}

get-lastinstalledDate

  