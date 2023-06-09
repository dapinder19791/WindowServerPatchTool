Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
$VerbosePreference = 'silentlycontinue'
$DebugPreference = 'silentlycontinue'
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name = 'OptionsWindow' Title="Tool Options" Height="565" Width="600"  
    WindowStartupLocation="CenterScreen" WindowStyle="ToolWindow" FontWeight="Bold">
        <Window.Background>
        <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
            <LinearGradientBrush.GradientStops> <GradientStop Color='#99c9ff' Offset='0' /> <GradientStop Color='#99c9ff' Offset='0.2' /> 
            <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
        </LinearGradientBrush>
    </Window.Background>     
      <Window.Resources>
        <Style x:Key="myStyle" TargetType="Button">
           <Setter Property="Background" Value="#c2cad6" />
           
        </Style>

      
    </Window.Resources>
     
    <Grid Name="Grid1" ShowGridLines="False">
        <Grid.RowDefinitions >
            <RowDefinition Height="100" />
            <RowDefinition Height="30" />
            <RowDefinition Height="30" />
            <RowDefinition Height="30" />
            <RowDefinition Height="30" />
            <RowDefinition Height="30" />
            <RowDefinition Height="30" />
            <RowDefinition Height="30" />
             <RowDefinition Height="30" />
              <RowDefinition Height="30" />
               <RowDefinition Height="30" />
                <RowDefinition Height="30" />
                 <RowDefinition Height="30" />
                  <RowDefinition Height="30" />
                   <RowDefinition Height="30" />

        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150" />
            <ColumnDefinition Width="10" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>
        <Label HorizontalContentAlignment="Left" Grid.ColumnSpan="3" Name="OptionsLabel" FontSize="24" VerticalAlignment="Top">Patching Tool Options</Label>
        <Label HorizontalContentAlignment="Left" Grid.Row="1" Name="MAxJobs_lbl" VerticalAlignment="Top" HorizontalAlignment="Left">MaxJobs</Label>
        <TextBox MinWidth="15"  MaxWidth="300" Grid.Column="2" Grid.Row="1" Name="MaxJobs_txtBx" VerticalAlignment="Top" />
        <TextBox MinWidth="15"  MaxWidth="300" Name="MaxRebootJobs_txtbx" Grid.Column="2" Grid.Row="2" VerticalAlignment="Top" />
        <Label HorizontalContentAlignment="Left" Name="MaxRebootJobs_lbl" Grid.Row="2" HorizontalAlignment="Left" VerticalAlignment="Top">MaxRebootJobs</Label>
        <Label HorizontalContentAlignment="Left" Name="ReportPath_lbl" Grid.Row="3" HorizontalAlignment="Left" VerticalAlignment="Top">Report Path</Label>
        <TextBox MinWidth="15"  MaxWidth="300" Name="ReportPath_txtbx" Grid.Column="2" Grid.Row="3" VerticalAlignment="Top" />

         <Label HorizontalContentAlignment="Left" Name="DiagPath_lbl" Grid.Row="4" HorizontalAlignment="Left" VerticalAlignment="Top">Diagnostic Log Path</Label>
        <TextBox MinWidth="15"  MaxWidth="300" Name="DiagPath_txtbx" Grid.Column="2" Grid.Row="4" VerticalAlignment="Top" />


             <Label HorizontalContentAlignment="Left" Name="DiagName_lbl" Grid.Row="5" HorizontalAlignment="Left" VerticalAlignment="Top">Diagnostic Log File Name</Label>
        <TextBox MinWidth="15"  MaxWidth="300" Name="DiagName_txtbx" Grid.Column="2" Grid.Row="5" VerticalAlignment="Top" />



        <Label HorizontalContentAlignment="Left" Name="user_lbl" Grid.Row="6" HorizontalAlignment="Left" VerticalAlignment="Top">Option 1 : User Id</Label>
        <TextBox MinWidth="15"  MaxWidth="300" Name="user_txtbx" Grid.Column="2" Grid.Row="6" VerticalAlignment="Top" />

        <Label HorizontalContentAlignment="Left" Name="pass_lbl" Grid.Row="7" HorizontalAlignment="Left" VerticalAlignment="Top">Option 1 : Password</Label>
        <PasswordBox MinWidth="15"  MaxWidth="300" PasswordChar="*" Name="pass_txtbx" Grid.Column="2" Grid.Row="7" VerticalAlignment="Top" />

          <Label HorizontalContentAlignment="Left" Name="user_lbl2" Grid.Row="8" HorizontalAlignment="Left" VerticalAlignment="Top">Option 2 : User Id</Label>
        <TextBox MinWidth="15"  MaxWidth="300" Name="user_txtbx2" Grid.Column="2" Grid.Row="8" VerticalAlignment="Top" />

        <Label HorizontalContentAlignment="Left" Name="pass_lbl2" Grid.Row="9" HorizontalAlignment="Left" VerticalAlignment="Top">Option 2 : Password</Label>
        <PasswordBox MinWidth="15"  MaxWidth="300" PasswordChar="*" Name="pass_txtbx2" Grid.Column="2" Grid.Row="9" VerticalAlignment="Top" />

          <Label HorizontalContentAlignment="Left" Name="user_lbl3" Grid.Row="10" HorizontalAlignment="Left" VerticalAlignment="Top">Option 3 : User Id</Label>
        <TextBox MinWidth="15"  MaxWidth="300" Name="user_txtbx3" Grid.Column="2" Grid.Row="10" VerticalAlignment="Top" />

        <Label HorizontalContentAlignment="Left" Name="pass_lbl3" Grid.Row="11" HorizontalAlignment="Left" VerticalAlignment="Top">Option 3 : Password</Label>
        <PasswordBox MinWidth="15"  MaxWidth="300" PasswordChar="*" Name="pass_txtbx3" Grid.Column="2" Grid.Row="11" VerticalAlignment="Top" />

 


        <Grid Grid.Column="2" Grid.Row="12" Name="Grid2" ShowGridLines="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="120" />
                <ColumnDefinition Width="10" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>

            
            <Button Cursor="Hand" Style="{StaticResource myStyle}" HorizontalAlignment="Right"  Name="Cancel_btn" Grid.Column="0" VerticalAlignment="Center" Width = "80">Cancel</Button>
            <Button Cursor="Hand" Style="{StaticResource myStyle}" HorizontalAlignment="Left"  Name="Save_btn" Grid.Column="2" VerticalAlignment="Center" Width = "80">Save</Button>
        </Grid>
    </Grid>
</Window>
"@

$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Global:Window=[Windows.Markup.XamlReader]::Load( $reader )

##Connect to Controls
$MaxJobs_txtBx = $Window.FindName('MaxJobs_txtBx')
$ReportPath_txtbx = $Window.FindName('ReportPath_txtbx')
$DiagPath_txtbx= $Window.FindName('DiagPath_txtbx')
$DiagName_txtbx =$Window.FindName('DiagName_txtbx')
$MaxRebootJobs_txtbx = $Window.FindName('MaxRebootJobs_txtbx')
$user_txtbx = $Window.FindName('user_txtbx')
$pass_txtbx = $Window.FindName('pass_txtbx')

$user_txtbx2 = $Window.FindName('user_txtbx2')
$pass_txtbx2 = $Window.FindName('pass_txtbx2')

$user_txtbx3 = $Window.FindName('user_txtbx3')
$pass_txtbx3 = $Window.FindName('pass_txtbx3')
 
$Cancel_btn = $Window.FindName('Cancel_btn')
$Save_btn = $Window.FindName('Save_btn')

##Event Handlers
#Cancel Button
$Cancel_btn.Add_Click({
    $Window.Close()
})

$Window.Add_Loaded({
    $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
    $MaxRebootJobs_txtbx.Text = $Optionshash['MaxRebootJobs']
    $MaxJobs_txtBx.Text = $Optionshash['MaxJobs']
    $ReportPath_txtbx.Text = $Optionshash['ReportPath']
    if($Optionshash['DiagPath'] -eq $null)
    {
    $DiagPath_txtbx.Text="C:\"
    }
    else{
    $DiagPath_txtbx.Text = $Optionshash['DiagPath']
    }
    if($Optionshash['DiagName'] -ne $null)
     {
     $DiagName_txtbx.Text=$Optionshash['DiagName']
     }
     else{
     $DiagName_txtbx.Text="Diagnosticlog.log"
     }
    
    $user_txtbx.Text = $Optionshash['Userid']
    $user_txtbx2.Text = $Optionshash['Userid1']
    $user_txtbx3.Text = $Optionshash['Userid2']

    if($Optionshash['password'] -ne "")
    {
    $securestring  = convertto-securestring -string  $Optionshash['password']
    # retrieve the password.
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securestring)
    $passwsord = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    $pass_txtbx.Password  = $passwsord
    }
   
    if($Optionshash['password1'] -ne $null)
    {
    $securestring  = convertto-securestring -string  $Optionshash['password1']
    # retrieve the password.
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securestring)
    $passwsord2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    $pass_txtbx2.Password  = $passwsord2
    }
    
    
    if($Optionshash['password2'] -ne $null)
    {
    $securestring  = convertto-securestring -string  $Optionshash['password2']
    # retrieve the password.
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securestring)
    $passwsord3 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    $pass_txtbx3.Password  = $passwsord3
    }
    
    
     
    
    
    Write-Verbose ("Current Path: {0}" -f $Path)
})

#Save Button
$Save_btn.Add_Click({
    $i = 0
    $Optionshash = Import-Clixml -Path (Join-Path $Path 'options.xml')
    
    #Validate option data is valid
    If ($MaxRebootJobs_txtbx.Text -notmatch "^\d+$") {
        $MaxRebootJobs_txtbx.ForeGround = 'Red'
        $i++
    } Else {
        $MaxRebootJobs_txtbx.Foreground = 'Black'
        $Optionshash['MaxRebootJobs'] = $MaxRebootJobs_txtbx.Text
    }
    If ($MaxJobs_txtBx.Text -notmatch "^\d+$") {
        $MaxJobs_txtBx.ForeGround = 'Red'
        $i++
    } Else {
        $MaxJobs_txtBx.Foreground = 'Black'
        $Optionshash['MaxJobs'] = $MaxJobs_txtBx.Text
    }
        
    If ($ReportPath_txtbx.Text -notmatch '\w:\\[a-zA-Z0-9\\-_]*') {
        $ReportPath_txtbx.ForeGround = 'Red'
        $i++
    } Else {
        $ReportPath_txtbx.Foreground = 'Black'
        $Optionshash['ReportPath'] = $ReportPath_txtbx.Text
    } 

    
    If ($DiagPath_txtbx.Text -notmatch '\w:\\[a-zA-Z0-9\\-_]*') {
        $DiagPath_txtbx.ForeGround = 'Red'
        $i++
    } Else {
        $DiagPath_txtbx.Foreground = 'Black'
        $Optionshash['DiagPath'] = $DiagPath_txtbx.Text
    } 

    $passflag=$false
    $userflag=$false

    if($user_txtbx.Text.Trim() -eq "")
    {
     $user_txtbx.ForeGround = 'Red'
     $i++
    }
    else{
    $user_txtbx.ForeGround = 'Black'
    $Optionshash['Userid'] = $user_txtbx.Text
    $userflag=$true
    }

    if($pass_txtbx.Password.Trim() -eq "")
    {
     $i++
      
    }
    else{
    $passflag=$true
    $Optionshash['password'] = $pass_txtbx.Password | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    }     

    $Optionshash['Userid1'] = $user_txtbx2.Text
    $Optionshash['Userid2'] = $user_txtbx3.Text
    $Optionshash['DiagName'] =  $DiagName_txtbx.Text
      
   if($pass_txtbx2.Password -ne ""){
   $Optionshash['password1'] = $pass_txtbx2.Password | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString}
   
   if($pass_txtbx3.Password -ne "")
   {
   $Optionshash['password2'] = $pass_txtbx3.Password | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString}
   
 
    #Save update options to XML file
    If ($i -eq 0) {
    write-host "Saving the Option Settings"
        $optionshash | Export-Clixml -Path (Join-Path $pwd 'options.xml') -Force
        $Window.Close()
    }
    else{
    if(!$userflag)
    {
    write-host "User id cannot be empty"
    }
    elseif(!$passflag)
    {
     write-host "Password cannot be empty"
    }
    }
})

#Used for debugging
$Window.Add_KeyUp({
    If ($_.Key -eq 'F5') {
        Write-Verbose ("MaxJobs_txtBx.Text: {0};{1}" -f $MaxJobs_txtBx.Text,($MaxJobs_txtBx.Text -notmatch "^\d+$"))
        Write-Verbose ("MaxRebootJobs_txtbx: {0};{1}" -f $MaxRebootJobs_txtbx.Text,($MaxRebootJobs_txtbx.Text -notmatch "^\d+$"))
        Write-Verbose ("ReportPath_txtbx: {0}" -f $ReportPath_txtbx.Text)
        Write-Verbose ("DiagPath_txtbx: {0}" -f $DiagPath_txtbx.Text)
        Write-Verbose ("I: {0}" -f $i)
    }
})

$Window.Showdialog() | Out-Null