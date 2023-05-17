 
 Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
 
 Set-Location $(Split-Path $MyInvocation.MyCommand.Path)
$Global:Path = $(Split-Path $MyInvocation.MyCommand.Path)
Write-Debug "Current location: $Path"


[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form1 = New-Object 'System.Windows.Forms.Form'
    $form1.MaximizeBox = $false

 
#######################
function New-Chart
{
    param ([int]$width,[int]$height,[int]$left,[int]$top,$xTitle,$yTitle)
    # create chart object
    $global:Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $global:Chart.Width = $width
    $global:Chart.Height = $height
    $global:Chart.Left = $left
    $global:Chart.Top = $top
   # create a chartarea to draw on and add to chart
    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $global:chart.ChartAreas.Add($chartArea)
  
 
    # change chart area colour
    $global:Chart.BackColor = [System.Drawing.Color]::Transparent
 
} #New-Chart
 
#######################
function New-BarColumnChart
{
    param ([hashtable]$ht, $chartType='Column', $chartTitle,$xTitle,$yTitle, [int]$width,[int]$height,[int]$left,[int]$top,[bool]$asc)
 
    New-Chart -width $width -height $height -left $left -top $top -chartTile $chartTitle
 
    $chart.ChartAreas[0].AxisX.Title = $xTitle
    $chart.ChartAreas[0].AxisY.Title = $yTitle

    
    [void]$global:Chart.Series.Add("Data")
    $global:Chart.Series["Data"].Points.DataBindXY($ht.Keys, $ht.Values)
 
   
    $global:Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::$chartType
  

    if ($asc)
    {
        $global:Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "Y")
    }
    else
    {
        $global:Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Descending, "Y")
    }
 
    $global:Chart.Series["Data"]["DrawingStyle"] = "Cylinder"
    $global:chart.Series["Data"].IsValueShownAsLabel = $true
    $global:chart.Series["Data"]["LabelStyle"] = "Top"
 
} #New-BarColumnChart
 
#######################
function New-LineChart
{
 
    param ([hashtable]$ht,$chartTitle, [int]$width,[int]$height,[int]$left,[int]$top)
 
    New-Chart -width $width -height $height -left $left -top $top -chartTile $chartTitle
 
    [void]$global:Chart.Series.Add("Data")
    #$global:Chart.Series["Data"].Points.AddXY($(get-date), $($ht.Values))
    $global:Chart.Series["Data"].Points.DataBindXY($ht.Keys,$ht.Values)
 
    #$global:Chart.Series["Data"].XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::Time
    #$global:Chart.chartAreas[0].AxisX.LabelStyle.Format = "hh:mm:ss"
    #$global:Chart.chartAreas[0].AxisX.LabelStyle.Interval = 1
    #$global:Chart.chartAreas[0].AxisX.LabelStyle.IntervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Seconds
    $global:Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    #$global:chart.Series["Data"].IsValueShownAsLabel = $false
 
} #New-LineChart
 
#######################
function New-PieChart
{
    param ([hashtable]$ht,$chartTitle, [int]$width,[int]$height,[int]$left,[int]$top)
 
    New-Chart -width $width -height $height -left $left -top $top -chartTile $chartTitle
 
    [void]$global:Chart.Series.Add("Data")
    $global:Chart.Series["Data"].Points.DataBindXY($ht.Keys, $ht.Values)
    $global:Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
    $global:Chart.Series["Data"]["PieLabelStyle"] = "Outside"
    $global:Chart.Series["Data"]["PieLineColor"] = "Black"
    #$global:chart.Series["Data"].IsValueShownAsLabel = $true
    #$global:chart.series["Data"].Label =  "#PERCENT{P1}"
    #$legend = New-object System.Windows.Forms.DataVisualization.Charting.Legend
    #$global:Chart.Legends.Add($legend)
    #$Legend.Name = "Default"
 
} #New-PieChart
 
#######################
function Remove-Points
{
    param($name='Data',[int]$maxPoints=200)
 
    while ($global:chart.Series["$name"].Points.Count > $maxPoints)
    {
        $global:chart.Series["$name"].Points.RemoveAT(0)
    }
 
} #Remove-Points
 
 
   
  
 
 
  #Out-Form

function Out-Form
{
    param($interval,$scriptBlock,$xField,$yField,$total)
 
    # display the chart on a form
    $global:Chart.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor
                    [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
   
    $form1.Text = 'PowerCharts'
    $form1.Width = 1027
    $form1.Height = 860
    $form1.backcolor = "#b0b3b7"
    $form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $form1.controls.add($global:Chart)
    $form1.MaximizeBox =$false
    $form1 = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Bold)
 
    
     if ($scriptBlock -is [ScriptBlock])
    {
        if (!($xField -or $yField))
        {
            throw 'xField and yField required with scriptBlock parameter.'
        }
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = $interval
        $timer.add_Tick(
        {
            $ht = &$scriptBlock | ConvertTo-HashTable $xField $yField
            if ($global:Chart.Series["Data"].ChartTypeName -eq 'Line')
            {
                Remove-Points
                $ht | foreach { $global:Chart.Series["Data"].Points.AddXY($($_.Keys), $($_.Values)) }
            }
            else
            {
                $global:Chart.Series["Data"].Points.DataBindXY($ht.Keys, $ht.Values)

                
        

            }
            $global:chart.ResetAutoValues()
            $global:chart.Invalidate()
        })
        $timer.Enabled = $true
        $timer.Start()
    }
 
  

   


   
} 
 
#######################
function Out-ImageFile
{
    param($fileName,$fileformat)
 
    $global:Chart.SaveImage($fileName, $fileformat)
}
 
#######################
 
#######################
function ConvertTo-Hashtable
{
    param([string]$key, $value)
 
    Begin
    {
        $hash = @{}
    }
    Process
    {
        $thisKey = $_.$Key
        $hash.$thisKey = $_.$Value
    }
    End
    {
        Write-Output $hash
    }
 
} #ConvertTo-Hashtable
 
#######################
function Out-Chart
{
    param(  $xField=$(throw 'Out-Chart:xField is required'),
            $yField=$(throw 'Out-Chart:yField is required'),
            $chartType='Column',
            $chartTitle,
            [int]$width=700,
            [int]$height=700,
            [int]$left=40,
            [int]$top=30,
            $filename,
            $fileformat='png',
            [int]$interval=1000,
            $scriptBlock,
            [switch]$asc,
            [int]$total
        )
 
    Begin
    {
        $ht = @{}
    }
    Process
    {
        if ($_)
        {
            $thisKey = $_.$xField
            $ht.$thisKey = $_.$yField
        }
    }
    End
    {
        if ($scriptBlock)
        {
            $ht = &$scriptBlock | ConvertTo-HashTable $xField $yField
        }
 
        switch ($chartType)
        {
            'Bar' {New-BarColumnChart -ht $ht -chartType $chartType -chartTitle $chartTitle -width $width -height $height -left $left -top $top -asc $($asc.IsPresent)}
            'Column' {New-BarColumnChart -ht $ht -chartType $chartType -chartTitle $chartTitle -width $width -height $height -left $left -top $top -asc $($asc.IsPresent)}
            'Pie' {New-PieChart -chartType -ht $ht  -chartTitle $chartTitle -width $width -height $height -left $left -top $top }
            'Line' {New-LineChart -chartType -ht $ht -chartTitle $chartTitle -width $width -height $height -left $left -top $top }
        }
 
        if ($filename)
        {
            Out-ImageFile  $filename $fileformat  
        }
        elseif ($scriptBlock )
        {
            Out-Form -interval $interval -scriptBlock $scriptBlock -xField $xField -yField $yField -total $total
        }
        else
        {
            Out-Form -total $total
        }
      
    }
 
} #Out-Chart
 

  
function getdata1{
 
 
 If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}
      
       $report = $reportpath + "\" + "serverlist.csv"

       if (Test-Path  ($report))
       {
       $data = import-csv $report
       $totalservers=$data.Count
       $data | Sort-Object -Property Computer|Group-Object Statustxt  | Select Count,Name   | out-chart -xField 'Name' -yField 'Count' -total $totalservers 
       }
 
      else{
      
      Write-Host "The Report Path is Invalid , please check the File > Options > Settings"

      }
  }
      
     
      


}
function getdata2{
 
 
 If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}
    
      $report = $reportpath + "\" + "serverlist.csv"
      
       if (Test-Path  ($report))
       {
      $data = import-csv $report
      $totalservers=$data.Count
      

      $data | Sort-Object -Property Computer|Group-Object Statustxt  | Select Count,Name   | out-chart -xField 'Name' -yField 'Count' -chartType "Pie" -total $totalservers 
      }
      
 else{
      
      Write-Host "The Report Path is Invalid , please check the File > Options > Settings"

      }
       

    }
}

function getdata_os1{
 
 
 If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}
      $report = $reportpath + "\" + "serverlist.csv"
      
       if (Test-Path  ($report))
       {
      $data = import-csv $report
      $totalservers=$data.Count
       
      $data | Sort-Object -Property Computer|Group-Object OsVersion  | Select Count,Name   | out-chart -xField 'Name' -yField 'Count' -chartType "Pie"  -total $totalservers 
     
      }

      
      
 else{
      
      Write-Host "The Report Path is Invalid , please check the File > Options > Settings"

      }
     
   }

}
function getdata_os2{
 
 
 If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}
      $report = $reportpath + "\" + "serverlist.csv"
   
       if (Test-Path  ($report))
       {
      
      $data = import-csv $report
      $totalservers=$data.Count
       
      $data | Sort-Object -Property Computer|Group-Object OsVersion  | Select Count,Name   | out-chart -xField 'Name' -yField 'Count' -chartType "Bar"  -total $totalservers 

      }

      
      
 else{
      
      Write-Host "The Report Path is Invalid , please check the File > Options > Settings"

      }
     
 }


}
function getdata_sql1{
If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}

      $report = $reportpath + "\" + "serverlist.csv"
  
      if (Test-Path  ($report))
       {
       
      $data = import-csv $report
      $totalservers=$data.Count
      

      $data | Sort-Object -Property Computer|Group-Object SqlVersion  | Select Count,Name   | out-chart -xField 'Name' -yField 'Count' -chartType "Bar"  -total $totalservers 

      }


 else{
      
      Write-Host "The Report Path is Invalid , please check the File > Options > Settings"

      }
  }

}
function getdata_sql2{
If (Test-Path (Join-Path $Path 'options.xml')) {
 
        $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
        If ($Optionshash['ReportPath']) {
            $reportpath = $Optionshash['ReportPath']}

      $report = $reportpath + "\" + "serverlist.csv"
       if (Test-Path  ($report))
       {
      
      $data = import-csv $report
      $totalservers=$data.Count
      

      $data | Sort-Object -Property Computer|Group-Object SqlVersion  | Select Count,Name   | out-chart -xField 'Name' -yField 'Count' -chartType "Column"  -total $totalservers 


      }

 else{
      
      Write-Host "The Report Path is Invalid , please check the File > Options > Settings"

      }
 }

}

function Show-tabcontrol_psf {


	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------

	$tabcontrol1 = New-Object 'System.Windows.Forms.TabControl'
	$tabpage1 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage2 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage3 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage4 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage5 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage6 = New-Object 'System.Windows.Forms.TabPage'
	 
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$form1_Load={
                        getdata1
                        $tabpage1.Controls.add($global:Chart)
		 
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
    
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form1.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$form1.remove_Load($form1_Load)
			$form1.remove_Load($Form_StateCorrection_Load)
			$form1.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null  }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$form1.SuspendLayout()
	$tabcontrol1.SuspendLayout()
	#
	# form1
	#

  $KayakoTabWasLoaded = 0
    # I set the counter to 0 as the tab is not yet been selected

            $handler_tabpage_SelectedIndexChanged_KayakoTab = {
                    if($KayakoTabWasLoaded -lt 1){ # if the tab has not yet been selected, include the file with the needed functions
                    if ($tabcontrol1.SelectedTab.Name -eq "ScanStatus1")  {
                       getdata1
                        $tabpage1.Controls.add($global:Chart)
                    }
                    elseif ($tabcontrol1.SelectedTab.Name -eq "ScanStatus2")  {
                       getdata2
                        $tabpage2.Controls.add($global:Chart)
                    }
                    elseif ($tabcontrol1.SelectedTab.Name -eq "OsDetails1")  {
                     
                       getdata_os1
                        $tabpage3.Controls.add($global:Chart)
                        
                    }
                    elseif ($tabcontrol1.SelectedTab.Name -eq "OsDetails2")  {
                     
                       getdata_os2
                        $tabpage4.Controls.add($global:Chart)
                        
                    }
                     elseif ($tabcontrol1.SelectedTab.Name -eq "SqlDetails1")  {
                     
                       getdata_sql1
                        $tabpage5.Controls.add($global:Chart)
                        
                    }
                     elseif ($tabcontrol1.SelectedTab.Name -eq "SqlDetails2")  {
                     
                       getdata_sql2
                        $tabpage6.Controls.add($global:Chart)
                        
                    }
                }else{ } # the tab was selected before so do nothing
        }
  

  $tabcontrol1.add_SelectedIndexChanged($handler_tabpage_SelectedIndexChanged_KayakoTab)

	$form1.Controls.Add($tabcontrol1)
	$form1.AutoScaleDimensions = '6, 13'
	$form1.AutoScaleMode = 'Font'
	$form1.ClientSize = '1000, 800'
	$form1.Name = 'form1'
	$form1.Text = 'PowerChart'
	$form1.add_Load($form1_Load)
	#
	# tabcontrol1
	#
	$tabcontrol1.Controls.Add($tabpage1)
	$tabcontrol1.Controls.Add($tabpage2)
	$tabcontrol1.Controls.Add($tabpage3)
	$tabcontrol1.Controls.Add($tabpage4)
    $tabcontrol1.Controls.Add($tabpage5)
    $tabcontrol1.Controls.Add($tabpage6)
 
	$tabcontrol1.Alignment = 'Left'
	$tabcontrol1.Location = '12, 12'
	$tabcontrol1.Multiline = $True
	$tabcontrol1.Name = 'tabcontrol1'
	$tabcontrol1.SelectedIndex = 0
	$tabcontrol1.Size = '980, 780'
	$tabcontrol1.TabIndex = 0
    $tabcontrol1.ForeColor= "#8c0c35"
	#
	# tabpage1
	#
	$tabpage1.Location = '42, 4'
	$tabpage1.Name = 'ScanStatus1'
	$tabpage1.Padding = '3, 3, 3, 3'
	$tabpage1.Size = '800, 700'
	$tabpage1.TabIndex = 0
	$tabpage1.Text = 'ScanStatus1'
	$tabpage1.UseVisualStyleBackColor = $True
      
   


	#
	# tabpage2
	#
	$tabpage2.Location = '23, 4'
	$tabpage2.Name = 'ScanStatus2'
	$tabpage2.Padding = '3, 3, 3, 3'
	$tabpage2.Size = '800, 700'
	$tabpage2.TabIndex = 1
	$tabpage2.Text = 'ScanStatus2'
	$tabpage2.UseVisualStyleBackColor = $True

 
   

	#
	# tabpage3
	#
	$tabpage3.Location = '23, 4'
	$tabpage3.Name = 'OsDetails1'
	$tabpage3.Padding = '3, 3, 3, 3'
	$tabpage3.Size = '800, 700'
	$tabpage3.TabIndex = 2
	$tabpage3.Text = 'OsDetails1'
	$tabpage3.UseVisualStyleBackColor = $True
	#
 $tabpage4.Location = '23, 4'
	$tabpage4.Name = 'OsDetails2'
	$tabpage4.Padding = '3, 3, 3, 3'
	$tabpage4.Size = '800, 700'
	$tabpage4.TabIndex = 2
	$tabpage4.Text = 'OsDetails2'
	$tabpage4.UseVisualStyleBackColor = $True
	

$tabpage5.Location = '23, 4'
	$tabpage5.Name = 'SqlDetails1'
	$tabpage5.Padding = '3, 3, 3, 3'
	$tabpage5.Size = '800, 700'
	$tabpage5.TabIndex = 2
	$tabpage5.Text = 'SqlDetails1'
	$tabpage5.UseVisualStyleBackColor = $True
	

$tabpage6.Location = '23, 4'
	$tabpage6.Name = 'SqlDetails2'
	$tabpage6.Padding = '3, 3, 3, 3'
	$tabpage6.Size = '800, 700'
	$tabpage6.TabIndex = 2
	$tabpage6.Text = 'SqlDetails'
	$tabpage6.UseVisualStyleBackColor = $True
	
	$tabcontrol1.ResumeLayout()
	$form1.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form1.add_Load($Form_StateCorrection_Load)
     
   
  
	#Clean up the control events
	$form1.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $form1.ShowDialog()

} #End Function

   
#Call the form
Show-tabcontrol_psf | Out-Null




