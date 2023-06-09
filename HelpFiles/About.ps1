Function Open-PoshPAIGAbout {
	$rs=[RunspaceFactory]::CreateRunspace()
	$rs.ApartmentState = "STA"
	$rs.ThreadOptions = "ReuseThread"
	$rs.Open()
	$ps = [PowerShell]::Create()
	$ps.Runspace = $rs
    $ps.Runspace.SessionStateProxy.SetVariable("pwd",$pwd)
	[void]$ps.AddScript({ 
[xml]$xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    x:Name='AboutWindow' Title='About' Height = '170' Width = '330' ResizeMode = 'NoResize' WindowStartupLocation = 'CenterScreen' ShowInTaskbar = 'False'>    
        <Window.Background>
        <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
            <LinearGradientBrush.GradientStops> <GradientStop Color='#99c9ff' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
            <GradientStop Color='#99c9ff' Offset='0.9' /> <GradientStop Color='#99c9ff' Offset='1' /> </LinearGradientBrush.GradientStops>
        </LinearGradientBrush>
    </Window.Background>     
    <StackPanel>
            <Label FontWeight = 'Bold' FontSize = '20'>PowerShell Patch Tool </Label>
            <Label FontWeight = 'Bold' FontSize = '16' Padding = '0'> Version: 1.0.0 </Label>
            <Button x:Name = 'CloseButton' Width = '100'> Close </Button>
    </StackPanel>
</Window>
"@
#Load XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$AboutWindow=[Windows.Markup.XamlReader]::Load( $reader )


#Connect to Controls
$CloseButton = $AboutWindow.FindName("CloseButton")
$AuthorLink = $AboutWindow.FindName("AuthorLink")

#PsexecLink Event
$AuthorLink.Add_Click({
    Start-Process "http://learn-powershell.net"
    })

$CloseButton.Add_Click({
    $AboutWindow.Close()
    })

#Show Window
[void]$AboutWindow.showDialog()
}).BeginInvoke()
}