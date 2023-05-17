[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo")  
function checkSQLJobs
{
Param ($computername,$cmd)
	$srv=NEW-OBJECT ('MICROSOFT.SQLSERVER.MANAGEMENT.SMO.SERVER') $computername

	###########################################################################################################################
	if($cmd -EQ "JA") ## List Of All Jobs
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*"} | sort-object -descending {$_.lastrundate} | select name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}} | format-table -AUTOSIZE | Out-String -Width 5096
	}
	###########################################################################################################################
	elseif($cmd -EQ "JR") ## List Of Running Jobs
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.currentrunstatus -eq 1 } | sort-object -descending {$_.lastrundate} | select name,currentrunstatus,currentrunstep,lastrundate,lastrunoutcome,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}} | format-table -AUTOSIZE | Out-String -Width 5096
	}
	###########################################################################################################################
	elseif($cmd -EQ "JS") ## List Of Jobs with status "Succeeded" 
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.lastrunoutcome -eq "succeeded" -and $_.currentrunstatus -ne 1} | sort-object -descending {$_.lastrundate} | select name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}}  | format-table -AUTOSIZE | Out-String -Width 5096
	}
	###########################################################################################################################
	elseif($cmd -EQ "JF") ## List Of Jobs with status "Failed"
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.lastrunoutcome -eq "failed" -and $_.currentrunstatus -ne 1} | sort-object -descending {$_.lastrundate} | select name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}}  | format-table -AUTOSIZE | Out-String -Width 5096
	}
	###########################################################################################################################
	elseif($cmd -EQ "JC") ## List Of Jobs with status "Cancelled"
	{
	$srv.jobserver.jobs| where-object {$_.name -like "**" -and $_.lastrunoutcome -eq "cancelled" -and $_.currentrunstatus -ne 1} | sort-object -descending {$_.lastrundate} | select name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}}| format-table -AUTOSIZE | Out-String -Width 5096
	}
	###########################################################################################################################
	elseif($cmd -EQ "JD") ## List Of Jobs with status "Disabled"
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.isenabled -eq 0 } | sort-object -descending {$_.lastrundate} | select name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}} | format-table -AUTOSIZE | Out-String -Width 5096
	}
	###########################################################################################################################
	elseif($cmd -EQ "JNS") ## List Of Jobs which are not scheduled"
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.isenabled -eq 1 -and $_.nextrundate -eq "1/1/0001 12:00:00 AM" } | sort-object -descending {$_.lastrundate} | select name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}} | format-table -AUTOSIZE | Out-String -Width 5096
	}
	###########################################################################################################################
	elseif($cmd -EQ "jnxtrun") ## Jobs Next Run date and time
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.isenabled -eq 1 -and $_.nextrundate -ne "1/1/0001 12:00:00 AM" } | sort-object  {$_.nextrundate} | select name,currentrunstatus,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}} | format-table -AUTOSIZE | Out-String -Width 5096
	}
	###########################################################################################################################
}


function checkSqlServerExists
{
$val=$false

if (Test-Path “HKLM:\Software\Microsoft\Microsoft SQL Server\Instance Names\SQL”) {
$val= $true
 } Else {
 $val= $false
 }
 return $val
}

 
function execute{


Param ($Servers)

 
#        $Servers = $env:computername
 
        ## Log file path
        $outfilePath =  "\\$Servers\c$\$($Servers)_SQLJoblog.csv"

        if([System.IO.File]::Exists($outfilePath)){ 

          Remove-Item $outfilePath -force -recurse

         }

         $Jobsstatus=""
         $status=checkSqlServerExists
  
        if($status -eq $true)
        {
 
        $Jobsstatus=checkSQLJobs -computername $Servers -cmd "JA"  | Out-File $outfilePath -Append

        }
        else{
        $temp = "" |  select Computer, name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,Notes
                                    $temp.Computer = $Servers
                                    $temp.name = "SQL Server Not Installed"
                                    $temp.lastrundate = "SQL Server Not Installed"
                                    $temp.lastrunoutcome = "SQL Server Not Installed"
                                    $temp.currentrunstatus = "SQL Server Not Installed"   
                                    $temp.currentrunstep = "SQL Server Not Installed"   
                                    $temp.isenabled = "SQL Server Not Installed"   
                                    $temp.nextrundate = "SQL Server Not Installed" 
                                    $temp.Notes = "SQL Server Not Installed"
                                    $Jobsstatus = $temp

          $Jobsstatus = $Jobsstatus | format-table -AUTOSIZE | Out-String -Width 5096
          $Jobsstatus  | Out-File $outfilePath -Append
        }
 
} 


execute -Servers $env:computername