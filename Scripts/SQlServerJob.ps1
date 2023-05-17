$Optionshash = Import-Clixml -Path (Join-Path $Path '.\options.xml')
                            $userid = $Optionshash['Userid']
                            $securestring  = convertto-securestring -string  $Optionshash['password']
                            $pass = $securestring
                            $cred = new-object -typename System.Management.Automation.PSCredential `
                                     -argumentlist $userid, $pass

$password=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password))
  
   
 $temp=@()
    
Function Create-UpdateVBS {
Param ($computername)
  if (Test-Connection -ComputerName $computername -Count 1 -Quiet){
    #Create Here-String of vbscode to create file on remote system
                try{
  
                net use \\$computername  $password /USER:$userid $powerShellArgs 2> $null
                execute -Servers $computername
          
                 }Catch
                 {
                 
                 }
  }
  
}

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo")  
function checkSQLJobs
{
Param ($computername,$cmd)
	$srv=NEW-OBJECT ('MICROSOFT.SQLSERVER.MANAGEMENT.SMO.SERVER') $computername

	###########################################################################################################################
	if($cmd -EQ "JA") ## List Of All Jobs
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*"} | sort-object -descending {$_.lastrundate} | Select-Object @{name="Computer";expression={$computername}},name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}}  

	}
	###########################################################################################################################
	elseif($cmd -EQ "JR") ## List Of Running Jobs
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.currentrunstatus -eq 1 } | sort-object -descending {$_.lastrundate} | select @{name="Computer";expression={$computername}},name,currentrunstatus,currentrunstep,lastrundate,lastrunoutcome,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}}  
	}
	###########################################################################################################################
	elseif($cmd -EQ "JS") ## List Of Jobs with status "Succeeded" 
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.lastrunoutcome -eq "succeeded" -and $_.currentrunstatus -ne 1} | sort-object -descending {$_.lastrundate} | select @{name="Computer";expression={$computername}},name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}}  | format-table -AUTOSIZE | Out-String 4059
	}
	###########################################################################################################################
	elseif($cmd -EQ "JF") ## List Of Jobs with status "Failed"
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.lastrunoutcome -eq "failed" -and $_.currentrunstatus -ne 1} | sort-object -descending {$_.lastrundate} | select @{name="Computer";expression={$computername}},name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}}  | format-table -AUTOSIZE | Out-String 4059
	}
	###########################################################################################################################
	elseif($cmd -EQ "JC") ## List Of Jobs with status "Cancelled"
	{
	$srv.jobserver.jobs| where-object {$_.name -like "**" -and $_.lastrunoutcome -eq "cancelled" -and $_.currentrunstatus -ne 1} | sort-object -descending {$_.lastrundate} | select @{name="Computer";expression={$computername}},name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}}| format-table -AUTOSIZE | Out-String 4059
	}
	###########################################################################################################################
	elseif($cmd -EQ "JD") ## List Of Jobs with status "Disabled"
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.isenabled -eq 0 } | sort-object -descending {$_.lastrundate} | select @{name="Computer";expression={$computername}},name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}} | format-table -AUTOSIZE | Out-String 4059
	}
	###########################################################################################################################
	elseif($cmd -EQ "JNS") ## List Of Jobs which are not scheduled"
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.isenabled -eq 1 -and $_.nextrundate -eq "1/1/0001 12:00:00 AM" } | sort-object -descending {$_.lastrundate} | select @{name="Computer";expression={$computername}},name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}} | format-table -AUTOSIZE | Out-String 4059
	}
	###########################################################################################################################
	elseif($cmd -EQ "jnxtrun") ## Jobs Next Run date and time
	{
	$srv.jobserver.jobs| where-object {$_.name -like "*" -and $_.isenabled -eq 1 -and $_.nextrundate -ne "1/1/0001 12:00:00 AM" } | sort-object  {$_.nextrundate} | select @{name="Computer";expression={$computername}},name,currentrunstatus,nextrundate,@{name="Notes";expression={$_.lastrunoutcome}} | format-table -AUTOSIZE | Out-String 4059
	}
	###########################################################################################################################
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

 
function execute{


Param ($Servers)
  
         $Jobsstatus=""
         $status=checkSqlServerExists -computername $Servers
   
        if($status -eq $true)
        {
 
        $Jobsstatus=checkSQLJobs -computername $Servers -cmd "JA"  
        
        }
        else{
        $temp = "" |  Select-Object Computer, name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,Notes
                                    $temp.Computer = $Servers
                                    $temp.name = "SQL Server Not Installed"
                                    $temp.lastrundate = "SQL Server Not Installed"
                                    $temp.lastrunoutcome = "SQL Server Not Installed"
                                    $temp.currentrunstatus = "SQL Server Not Installed"   
                                    $temp.currentrunstep = "SQL Server Not Installed"   
                                    $temp.isenabled = "SQL Server Not Installed"   
                                    $temp.nextrundate = "SQL Server Not Installed" 
                                    $temp.Notes = "SQL Server Not Installed"
                                    $Jobsstatus=$temp

          
           
        }
            
    Write-Output $Jobsstatus 
 
} 

 
 


function Get-Jobstatus
{
    [cmdletbinding()]
    Param($Computername)
     $installreport = @()

 try{  
 if(Test-Connection -ComputerName $Computername -Count 1 -Quiet)
   {
     try{   
     
      
            net use \\$Computername  $password /USER:$userid $powerShellArgs 2> $null
         }
    catch [Exception]{

            $temp = "" | Select-Object Computer,name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,Notes
            $temp.Computer =$computername
            $temp.name ="Un-Authorized"
            $temp.lastrundate ="Un-Authorized"
            $temp.lastrunoutcome ="Un-Authorized"
            $temp.currentrunstatus ="Un-Authorized"
            $temp.currentrunstep ="Un-Authorized"
            $temp.isenabled ="Un-Authorized"
            $temp.nextrundate ="Un-Authorized"
            $temp.Notes ="Un-Authorized"
            $installreport = $temp      
 
            Write-Output $installreport

    }
    finally{
     
      
            Create-UpdateVBS -computer $Computername
           
            }
        }      
   
  
  else{
  
   $temp = "" | Select-Object Computer,name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,Notes
            $temp.Computer =$computername
            $temp.name ="Offline"
            $temp.lastrundate ="Offline"
            $temp.lastrunoutcome ="Offline"
            $temp.currentrunstatus ="Offline"
            $temp.currentrunstep ="Offline"
            $temp.isenabled ="Offline"
            $temp.nextrundate ="Offline"
            $temp.Notes ="Offline"
            $installreport = $temp  

            Write-Output $installreport
  
    }
  }

  catch [Exception]{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName

            $temp = "" | Select-Object Computer,name,lastrundate,lastrunoutcome,currentrunstatus,currentrunstep,isenabled,nextrundate,Notes
            $temp.Computer =$computername
            $temp.name =$ErrorMessage
            $temp.lastrundate =$ErrorMessage
            $temp.lastrunoutcome =$ErrorMessage
            $temp.currentrunstatus =$ErrorMessage
            $temp.currentrunstep =$ErrorMessage
            $temp.isenabled =$ErrorMessage
            $temp.nextrundate =$ErrorMessage
            $temp.Notes ="Error"
            $installreport = $temp  

            Write-Output $installreport

           
        }

 }
      