$env:TITLE = " FINDVM GUI v2"

#List of vcenter is needed for run the script, create it in the same forlder of the .ps1 file
$testVCconn = Get-Content .\VCenter.txt -ErrorAction SilentlyContinue | select -First 1

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing 

$prerq = Get-PowerCLIConfiguration | Where-Object{$_.Scope -eq "AllUser"}
if ($prerq.ParticipateInCEIP -ne $false){
Set-PowerCLIConfiguration -InvalidCertificateAction ignore -confirm:$false -ParticipateInCeip:$false
Set-PowerCLIConfiguration -Scope AllUsers -ParticipateInCEIP $false -Confirm:$false
}

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) | Out-Null

function ACCESSPROMPT {
    Param($credTIT,$credText,$credUSR,$credPSW)
    $Host.ui.PromptForCredential("$credTIT","$credText","$credUSR","$credPSW") 
}

if (!$testVCconn){[System.Windows.MessageBox]::Show("VCENTER.TXT not found","ERROR","OK","Error") | Out-Null ; break}
if ($testVCconn){
do{
try{
$cred = ACCESSPROMPT -credTIT "NEED CREDENTIAL" -credText "Please type your userid" -credUSR "" -credPSW "NetBiosUserName" 
$conn = Connect-VIServer $testVCconn -Credential $cred -ErrorAction Stop
}
catch{
$wshell = New-Object -ComObject Wscript.Shell
$ERRMESS = $_.Exception.message
$ans = $wshell.Popup("$ERRMESS 

Do you want to retry?",0," CONNECTION ERROR",4)
if ($ans -eq "7"){exit}
}
} Until ($conn.Name -eq $testVCconn)
} 

$guide = '
Type the name of the <Bold>VM</Bold> and click on <Bold>SEARCH</Bold>  for search</Paragraph>
<Paragraph>the machine in all the virtual centers</Paragraph>
<Paragraph>
</Paragraph>
<Paragraph><Bold>NOW MAY BE USED THE FOLLOWING COMMANDS</Bold></Paragraph>
<Paragraph>
</Paragraph>
<Paragraph>1)  <Bold>CONSOLE</Bold> open the vmware console</Paragraph>
<Paragraph>2)  <Bold>START</Bold> starts the VM directly from VMware</Paragraph>
<Paragraph>3)  <Bold>STOP</Bold> turns off the VM directly from VMware</Paragraph>
<Paragraph>4)  <Bold>RESTART</Bold> restarts the VM directly from VMware</Paragraph>
<Paragraph>5)  <Bold>PING</Bold> performs a continuous ping</Paragraph>
<Paragraph>6)  <Bold>TEST PORT</Bold> perform a check of the server ports state</Paragraph>
<Paragraph>8)  <Bold>REMOTE DESKTOP</Bold> connects to remote desktop</Paragraph>
<Paragraph>9)  <Bold>VMPM</Bold> open a realtime performance monitor</Paragraph>
<Paragraph>10)<Bold>STAT CHART</Bold> you can create an history performance chart
'
$time = "
 <ComboBoxItem>00:00</ComboBoxItem>
 <ComboBoxItem>00:30</ComboBoxItem>
 <ComboBoxItem>01:00</ComboBoxItem>
 <ComboBoxItem>01:30</ComboBoxItem>
 <ComboBoxItem>02:00</ComboBoxItem>
 <ComboBoxItem>02:30</ComboBoxItem>
 <ComboBoxItem>03:00</ComboBoxItem>
 <ComboBoxItem>03:30</ComboBoxItem>
 <ComboBoxItem>04:00</ComboBoxItem>
 <ComboBoxItem>04:30</ComboBoxItem>
 <ComboBoxItem>05:00</ComboBoxItem>
 <ComboBoxItem>05:30</ComboBoxItem>
 <ComboBoxItem>06:00</ComboBoxItem>
 <ComboBoxItem>06:30</ComboBoxItem>
 <ComboBoxItem>07:00</ComboBoxItem>
 <ComboBoxItem>07:30</ComboBoxItem>
 <ComboBoxItem>08:00</ComboBoxItem>
 <ComboBoxItem>08:30</ComboBoxItem>
 <ComboBoxItem>09:00</ComboBoxItem>
 <ComboBoxItem>09:30</ComboBoxItem>
 <ComboBoxItem>10:00</ComboBoxItem>
 <ComboBoxItem>10:30</ComboBoxItem>
 <ComboBoxItem>11:00</ComboBoxItem>
 <ComboBoxItem>11:30</ComboBoxItem>
 <ComboBoxItem>12:00</ComboBoxItem>
 <ComboBoxItem>12:30</ComboBoxItem>
 <ComboBoxItem>13:00</ComboBoxItem>
 <ComboBoxItem>13:30</ComboBoxItem>
 <ComboBoxItem>14:00</ComboBoxItem>
 <ComboBoxItem>14:30</ComboBoxItem>
 <ComboBoxItem>15:00</ComboBoxItem>
 <ComboBoxItem>15:30</ComboBoxItem>
 <ComboBoxItem>16:00</ComboBoxItem>
 <ComboBoxItem>16:30</ComboBoxItem>
 <ComboBoxItem>17:00</ComboBoxItem>
 <ComboBoxItem>17:30</ComboBoxItem>
 <ComboBoxItem>18:00</ComboBoxItem>
 <ComboBoxItem>18:30</ComboBoxItem>
 <ComboBoxItem>19:00</ComboBoxItem>
 <ComboBoxItem>19:30</ComboBoxItem>
 <ComboBoxItem>20:00</ComboBoxItem>
 <ComboBoxItem>20:30</ComboBoxItem>
 <ComboBoxItem>21:00</ComboBoxItem>
 <ComboBoxItem>21:30</ComboBoxItem>
 <ComboBoxItem>22:00</ComboBoxItem>
 <ComboBoxItem>22:30</ComboBoxItem>
 <ComboBoxItem>23:00</ComboBoxItem>
 <ComboBoxItem>23:30</ComboBoxItem>
"

[xml]$xaml = @"
<Window Name="FINDVM"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

        Title=" FINDVM GUI v2" Height="398" Width="505" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" Cursor="Arrow" ShowInTaskbar="True">
    <Window.Background>
        <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
            <GradientStop Color="Black" Offset="0"/>
            <GradientStop Color="#FF8394FF" Offset="1"/>
        </LinearGradientBrush>
    </Window.Background>
    <Grid x:Name="vmgrid" Margin="0,0,0,0">
        <TextBox x:Name="VMNAME" HorizontalAlignment="Left" VerticalAlignment="Top" Width="245" Height="26" Margin="10,24,0,0" TextWrapping="Wrap" Text="VM" FontSize="15" VerticalContentAlignment="Center" Foreground="DarkGray"/>
        <Label x:Name="label1" HorizontalAlignment="Left" Width="259" Content='TYPE YOUR VM' Margin="10,0,0,336" Foreground="White" Padding="5" />
        <RichTextBox x:Name="outputBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="470" Height="223" Margin="10,60,0,0" FontSize="11" Block.LineHeight="1" FontFamily="lucida console" IsReadOnly="true" Padding="0,0,60,0">
            <RichTextBox.Background>
                <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
                    <LinearGradientBrush.RelativeTransform>
                        <TransformGroup>
                            <ScaleTransform CenterY="0.5" CenterX="0.5"/>
                            <SkewTransform CenterY="0.5" CenterX="0.5"/>
                            <RotateTransform Angle="90" CenterY="0.5" CenterX="0.5"/>
                            <TranslateTransform/>
                        </TransformGroup>
                    </LinearGradientBrush.RelativeTransform>
                    <GradientStop Color="#FF0E0E0E" Offset="0"/>
                    <GradientStop Color="White" Offset="0.141"/>
                </LinearGradientBrush>
            </RichTextBox.Background>
            <FlowDocument>
                <Paragraph>
                </Paragraph>
                <Paragraph>
                    $guide
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <GroupBox Header="CRITICAL VM OPERATION" HorizontalAlignment="Left" Height="65" Margin="9,287,0,0" VerticalAlignment="Top" Width="175"/>
        <GroupBox Header="TEST CONNECTION" HorizontalAlignment="Left" Height="65" Margin="189,287,0,0" VerticalAlignment="Top" Width="175"/>
        <Button x:Name ="SEARCH" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "35" Width = "100" Margin="269,15,0,0" Content = 'SEARCH' IsEnabled = "false"/>
        <Button x:Name ="CONSOLE" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "35" Width = "100" Margin="380,15,0,0" Content = 'CONSOLE' IsEnabled = "false"/>
        <Button x:Name ="PING" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "35" Width = "76" Margin="198,308,0,0" Content = 'PING' RenderTransformOrigin="-0.018,0.514" IsEnabled = "false"/>
        <Button x:Name ="START" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "35" Width = "38" Margin="16,308,0,0" Content = 'START' IsEnabled = "false"/>
        <Button x:Name ="RDP" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "35" Width = "100" Margin="375,297,0,0" Content = 'REMOTE&#xD;&#xA;DESKTOP' IsEnabled = "false"/>
        <Button x:Name ="RESTART" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "35" Width = "76" Margin="102,308,0,0" Content = 'RESTART' RenderTransformOrigin="-0.018,0.514" IsEnabled = "false"/>
        <Button x:Name ="STOP" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "35" Width = "38" Margin="59,308,0,0" Content = 'STOP' IsEnabled = "false"/>
        <Button x:Name ="TESTPORT" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "35" Width = "76" Margin="281,308,0,0" Content = 'TEST PORT' RenderTransformOrigin="-0.018,0.514" IsEnabled = "false"/>
        <Label Content="Andrea Springolo" HorizontalAlignment="Left" Margin="375,338,0,0" Foreground="Black" FontSize="9" Height="15" Width="100" FontStretch="Condensed" VerticalAlignment="Top" HorizontalContentAlignment="Center" Padding="0" FontFamily="Arial Rounded MT Bold"/>
        <TextBox x:Name="CPU" HorizontalAlignment="Left" VerticalAlignment="Top" Width="55" Height="26" Margin="420,77,0,0" TextWrapping="Wrap" FontSize="10" VerticalContentAlignment="Center" IsReadOnly="True" HorizontalContentAlignment="Center"/>
        <TextBox x:Name="MEM" HorizontalAlignment="Left" VerticalAlignment="Top" Width="55" Height="26" Margin="420,118,0,0" TextWrapping="Wrap" FontSize="10" VerticalContentAlignment="Center" IsReadOnly="True" HorizontalContentAlignment="Center"/>
        <TextBox x:Name="DISC_C" HorizontalAlignment="Left" VerticalAlignment="Top" Width="55" Height="26" Margin="420,159,0,0" TextWrapping="Wrap" FontSize="10" VerticalContentAlignment="Center" IsReadOnly="True" HorizontalContentAlignment="Center"/>
        <Label x:Name="CPU_L" Content="CPU%" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="421,65,0,0" Height="16" Width="40" FontSize="9" Padding="0"/>
        <Label x:Name="RAM_L" Content="RAM%" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="421,106,0,0" Height="16" Width="40" FontSize="9" Padding="0"/>
        <Label x:Name="DISC_L" Content="DISC C%" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="421,147,0,0" Height="16" Width="40" FontSize="9" Padding="0"/>
        <Button x:Name ="CHART" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "34" Width = "54" Margin="421,243,0,0" Content = 'STAT&#xD;&#xA;CHART' IsEnabled = "false" VerticalContentAlignment="Center" HorizontalContentAlignment="Center" FontSize="9"/>
        <Button x:Name ="VMPM" HorizontalAlignment="Left" VerticalAlignment="Top" Height = "34" Width = "54" Margin="421,205,0,0" Content = 'VMPM' IsEnabled = "false" VerticalContentAlignment="Center" HorizontalContentAlignment="Center" FontSize="9"/>
    </Grid>
</Window>
"@

[xml]$xaml2 = @"
<Window Name="TESTPORT"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

        Title="TESTPORT V2" Height="190" Width="504" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Window.Background>
        <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
            <GradientStop Color="Black" Offset="0"/>
            <GradientStop Color="#FFB1B8FF" Offset="1"/>
        </LinearGradientBrush>
    </Window.Background>
    <Grid x:Name="testconn">
        <TextBox x:Name="SERVER" HorizontalAlignment="Left" Height="21" TextWrapping="Wrap" VerticalAlignment="Top" Width="195" Margin="5,6,0,0"/>
        <TextBox x:Name="PORT" HorizontalAlignment="Left" Height="21" TextWrapping="Wrap" Text="PORT" VerticalAlignment="Top" Width="195" Margin="5,35,0,0" Foreground="DarkGray"/>
        <TextBox x:Name="OUTPUT" HorizontalAlignment="Left" Height="138" TextWrapping="Wrap" VerticalAlignment="Top" Width="277" Margin="205,6,0,0" FontSize="9" IsReadOnly="$true"/>
        <Button x:Name="RUNTEST" Content="RUN TEST" HorizontalAlignment="Left" Height="29" Margin="11,105,0,0" VerticalAlignment="Top" Width="183"/>
        <ComboBox x:Name="CPORT" HorizontalAlignment="Left" Margin="5,64,0,0" VerticalAlignment="Top" Width="195">
            <ComboBoxItem Background="#FFDDE5FF">CUSTOM PORT</ComboBoxItem>
            <ComboBoxItem>HTTP</ComboBoxItem>
            <ComboBoxItem>RDP</ComboBoxItem>
            <ComboBoxItem>SMB</ComboBoxItem>
            <ComboBoxItem>WINRM</ComboBoxItem>
        </ComboBox>
    </Grid>
</Window>
    
"@

[xml]$xaml3 = @"
<Window Name="STATCHART"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:wfi="clr-namespace:System.Windows.Forms.Integration;assembly=WindowsFormsIntegration" 
        xmlns:winformchart="clr-namespace:System.Windows.Forms.DataVisualization.Charting;assembly=System.Windows.Forms.DataVisualization"
        
        Title="STATCHART v1" Height="405" Width="585" ResizeMode="CanResize">
    <Window.Background>
        <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
            <GradientStop Color="Black" Offset="0"/>
            <GradientStop Color="#FFB1B8FF" Offset="1"/>
        </LinearGradientBrush>
    </Window.Background>
    <Grid x:Name="chartback">
        <DatePicker x:Name="startdate" HorizontalAlignment="Left" Margin="350,13,0,0" VerticalAlignment="Top" SelectedDateFormat="Short"/>
        <ComboBox x:Name="HW" HorizontalAlignment="Left" Margin="165,14,0,0" VerticalAlignment="Top" Width="86" SelectedIndex="0">
            <ComboBoxItem Background="#FFBBC1FF">CPU</ComboBoxItem>
            <ComboBoxItem Background="#FFAAC1FF">RAM</ComboBoxItem>
            <ComboBoxItem Background="#FF838FFF">DISK</ComboBoxItem>
            <ComboBoxItem Background="#FF5F5FFF">NETWORK</ComboBoxItem>
        </ComboBox>
        <DatePicker x:Name="enddate" HorizontalAlignment="Left" Margin="350,44,0,0" VerticalAlignment="Top" SelectedDateFormat="Short"/>
        <ComboBox x:Name="starttime" HorizontalAlignment="Left" VerticalAlignment="Top" Width="60" Margin="498,14,0,0" SelectedIndex="0" VerticalContentAlignment="Center">
            $time
        </ComboBox>
        <ComboBox x:Name="endtime" HorizontalAlignment="Left" VerticalAlignment="Top" Width="60" Margin="498,45,0,0" SelectedIndex="0" VerticalContentAlignment="Center">
            $time
        </ComboBox>
        <Button x:Name="load" Content="LOAD" HorizontalAlignment="Left" Margin="10,45,0,0" VerticalAlignment="Top" Width="147" Height="23"/>
        <Button x:Name="saveC" Content="SAVE to PNG" HorizontalAlignment="Left" Margin="165,45,0,0" VerticalAlignment="Top" Width="86" Height="23"/>
        <Button x:Name="saveF" Content="SAVE to CSV" HorizontalAlignment="Left" Margin="256,45,0,0" VerticalAlignment="Top" Width="86" Height="23"/>
        <TextBox x:Name="server" Text="VM" HorizontalAlignment="Left" Height="23" Margin="10,14,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="147" VerticalContentAlignment="Center" IsReadOnly="true"/>
        <DockPanel LastChildFill="true">
            <wfi:WindowsFormsHost x:Name="formhost" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="10,82,10,10">
                <winformchart:Chart x:Name="HWchart" Dock="Fill">
                    <winformchart:Chart.ChartAreas>
                        <winformchart:ChartArea Name="area"/>
                    </winformchart:Chart.ChartAreas>
                    <winformchart:Chart.Series>
                        <winformchart:Series Name="series" ToolTip="#SERIESNAME : #VALY{##}\ntime : #VALX{dd/MM/yy HH:mm}"/>
                    </winformchart:Chart.Series>
                </winformchart:Chart>
            </wfi:WindowsFormsHost>
        </DockPanel>
    </Grid>
</Window>
"@

[xml]$xaml4 = @"
<Window Name="VMPM"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        
        Title="VMPM" Height="310" Width="310" ResizeMode="NoResize" FontSize="9" Topmost="True">
    <Window.Background>
        <SolidColorBrush Color="Black" Opacity="0.8"/>
    </Window.Background>
    <Grid x:Name="chartback">
        <TextBox x:Name="OUTPUT" HorizontalAlignment="Left" Height="246" TextWrapping="Wrap" Text="loading...." VerticalAlignment="Top" Width="294" FontSize="11" Foreground="#FF3AFF00" Padding="5,5,5,5" Margin="0" Background="Black" FontFamily="Lucida Console"/>
        <Button x:Name="RESIZE1" Content="MINIMIZE" HorizontalAlignment="Left" Margin="0,246,0,0" VerticalAlignment="Top" Width="294" Height="25" Background="#FFDDDDDD"/>
        <Button x:Name="RESIZE2" Content="RESTORE" HorizontalAlignment="Left" VerticalAlignment="Top" Width="294" Height="25" Margin="0,0,0,0" Visibility="Hidden"/>
    </Grid>
</Window>
"@

$Reader = (New-Object System.Xml.XmlNodeReader $xaml) 
$FINDVM = [Windows.Markup.XamlReader]::Load($reader) 

$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {New-Variable -Name $_.Name -Value $FINDVM.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

$VMNAME.Add_GotFocus({if ($VMNAME.Text -eq "VM"){$VMNAME.Text = '' ; $VMNAME.Foreground = "black"}})
$VMNAME.Add_LostFocus({if ($VMNAME.text -eq ''){$VMNAME.Text = 'VM' ; $VMNAME.Foreground = "DarkGray"}})
$VMNAME.Add_TextChanged({if ($VMNAME.Text -ne '' -and $VMNAME.Text -ne 'VM'){$SEARCH.IsEnabled = "true"}})
$VMNAME.Add_KeyDown({if ($_.Key -eq 'Enter'){SearchVM}})
$SEARCH.Add_Click({SearchVM})
$CONSOLE.Add_Click({Console})
$START.Add_Click({Starter})
$STOP.Add_Click({Stopper})
$RESTART.Add_Click({Restarter})
$PING.Add_Click({Ping})
$TESTPORT.Add_Click({testconn})
$RDP.Add_Click({rdp})
$CHART.Add_Click({statchart})
$VMPM.Add_Click({VMPM_f})

function SearchVM(){
Disconnect-VIServer * -Confirm:$false

$CPU.Text = ""
$MEM.Text = ""
$DISC_C.Text = ""

$CPU.Background = "white"
$MEM.Background = "white"
$DISC_C.Background = "white"

$outputBox.FontSize = 11
$vm = ($VMNAME.text).trim()
$vmup = $VM.ToUpper()
$VMNAME.text = $vmup

$testlist = Test-Path -Path .\VCenter.txt

$vicenters = get-content .\VCenter.txt -ErrorAction SilentlyContinue

$FINDVM.Title = " SEARCHING FOR $VMup........."

if ($testlist){
ForEach ($vicenter in $vicenters){
$TVCenter = $vicenter.ToUpper()
Connect-VIServer $vicenter -Credential $cred -ErrorAction SilentlyContinue
$findvmout = Get-VM $VM -ErrorAction SilentlyContinue

if ($findvmout){
$outputBox.Document.Blocks.Clear();
$env:vmdom = $findvmout.Guest.HostName
$exp = $findvmout | Format-List -property @{name=”NAME";expression={$findvmout.Guest.HostName}},
@{name=”IP";expression={$findvmout.Guest.IPAddress}},
PowerState,
NumCPU,
MemoryMB,
@{name=”OS";expression={$findvmout.Guest.OSFullName}},
VMHost,
@{name=”Datastore";expression={$findvmout.ExtensionData.Datastore.value}},
Notes | Out-String
$outputBox.AppendText("-----------------------------------------------------------
$VMup located in $TVCenter
-----------------------------------------------------------`n")
$outputBox.AppendText($exp)
$outputBox.Foreground = "#000000"
$FINDVM.Title = $env:TITLE

$CONSOLE.IsEnabled = "true"
$PING.IsEnabled = "true"
$START.IsEnabled = "true"
$RDP.IsEnabled = "true"
$RESTART.IsEnabled = "true"
$STOP.IsEnabled = "true"
$TESTPORT.IsEnabled = "true"
$CHART.IsEnabled = "true"
$VMPM.IsEnabled = "true"

$b= Get-Stat -Entity $findvmout -Realtime -ErrorAction SilentlyContinue -MaxSamples 1
$bcpu = ($b | Where-Object {$_.MetricId -eq "cpu.usage.average"} | select value).value
$bmem = ($b | Where-Object {$_.MetricId -eq "mem.usage.average"} | select value).value

$drive = $findvmout.Guest.Disks | Where-Object{$_.Path -like "C*"}
$Path = $drive.Path
$cap = $drive.CapacityGB
$free = $drive.FreeSpaceGB
$perc = [math]::Round(($free)/($cap)*100,2)
$perutil = 100 - $perc

$CPU.Text = "$bcpu" + " %"
if ($bcpu -lt "85"){$CPU.Background = "#FF63FF6A"}
if ($bcpu -ge "85"){$CPU.Background = "Yellow"}
if ($bcpu -ge "95"){$CPU.Background = "Red"}
$MEM.Text = "$bmem" + " %"
if ($bmem -lt "85"){$MEM.Background = "#FF63FF6A"}
if ($bmem -ge "85"){$MEM.Background = "Yellow"}
if ($bmem -ge "95"){$MEM.Background = "Red"}
$DISC_C.Text =  "$perutil"  + " %"
if ($perutil -lt "85"){$DISC_C.Background = "#FF63FF6A"}
if ($perutil -ge "85"){$DISC_C.Background = "Yellow"}
if ($perutil -ge "95"){$DISC_C.Background = "Red"}

$FINDVM.Title = "$vmup FOUND!"

break
}
}
if (!$findvmout){
$FINDVM.Title = "$VMup NOT FOUND"

$outputBox.Document.Blocks.Clear();
$outputBox.AppendText( "
 $VMup NOT FOUND")

$outputBox.Foreground = "#ff0000"
$outputBox.FontSize = 15

$CONSOLE.IsEnabled = $false
$PING.IsEnabled = $false
$START.IsEnabled = $false
$RDP.IsEnabled = $false
$RESTART.IsEnabled = $false
$STOP.IsEnabled = $false
$TESTPORT.IsEnabled = $false
$CHART.IsEnabled = $false
$VMPM.IsEnabled = $false
}
} else {
$outputBox.Document.Blocks.Clear();
$FINDVM.Title = " $env:TITLE"
$outputBox.AppendText( "
 VCENTER LIST NOT FOUND 
 PLEASE CHECK")
$outputBox.Foreground = "#ff0000"
$outputBox.FontSize = 15

$CONSOLE.IsEnabled = $false
$PING.IsEnabled = $false
$START.IsEnabled = $false
$RDP.IsEnabled = $false
$RESTART.IsEnabled = $false
$STOP.IsEnabled = $false
$TESTPORT.IsEnabled = $false
$CHART.IsEnabled = $false
$VMPM.IsEnabled = $false
}
}

function Console{Get-VM $VMNAME.text | Open-VMConsoleWindow}

function Starter(){
$a = new-object -comobject wscript.shell
$popup = $a.popup("Are you sure to start the VM?",0,"START",4)
if ($popup -eq 6) {Start-VM -VM $VMNAME.text -RunAsync -Confirm:$false}
}

function Stopper(){
$b = new-object -comobject wscript.shell
$popup = $b.popup("Are you sure to stop the VM?",0,"STOP",4)
if ($popup -eq 6) {Stop-VM -VM $VMNAME.text -RunAsync -Confirm:$false}
}

function Restarter(){
$c = new-object -comobject wscript.shell
$popup = $c.popup("Are you sure to restart the VM?",0,"RESTART",4)
if ($popup -eq 6) {Restart-VM -VM $VMNAME.text -RunAsync -Confirm:$false}
}

function ping(){Start-Process -FilePath "ping.exe" -ArgumentList " $env:vmdom -t"}

function rdp(){Start-Process -FilePath "mstsc.exe" -ArgumentList "/V:$env:vmdom"}

function testconn(){
$Reader2 = (New-Object System.Xml.XmlNodeReader $xaml2) 
$TESTPORT = [Windows.Markup.XamlReader]::Load($reader2) 

$xaml2.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {New-Variable -Name $_.Name -Value $TESTPORT.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

$VMup = ($VMNAME.text).ToUpper()
$SERVER.text = $VMup

$PORT.Add_GotFocus({if ($PORT.Text -eq "PORT"){$PORT.Text = ''}})
$PORT.Add_LostFocus({if ($PORT.text -eq ''){$PORT.Text = 'PORT' ; $PORT.Foreground = "DarkGray" ; $CPORT.IsEnabled = "true" ; $CPORT.Foreground = "Black"}})
$PORT.Add_TextChanged({if ($PORT.Text -ne '' -and $PORT.Text -ne 'PORT'){$CPORT.IsEnabled = $false ; $PORT.Foreground = "black" ; $CPORT.Text = 'CUSTOM PORT' ; $CPORT.Foreground = "DarkGray"}})

$CPORT.Add_SelectionChanged({
if ($CPORT.SelectedIndex -ne "0"){$PORT.IsEnabled = $false}
if ($CPORT.SelectedIndex -eq "0"){$PORT.IsEnabled = "true"}
})

$SERVER.Add_KeyDown({if ($_.Key -eq 'Enter'){testport}})
$RUNTEST.Add_Click({testport})

function testport{
$env:servconn = ''
$env:portconn = ''

if ($CPORT.SelectedIndex -ne "0"){
$comp = $SERVER.Text
$env:servconn = [System.Net.Dns]::GetHostByName(("$comp")).Hostname
$env:portconn = $CPORT.text
if ($env:portconn){
$tempconn = Start-Job -ScriptBlock {Test-NetConnection -ComputerName $env:servconn -CommonTCPPort $env:portconn -WarningAction SilentlyContinue -ErrorAction SilentlyContinue}
$ou = Receive-Job -Job $tempconn -Wait

$env:t1 = $SERVER.text
$t2 = $OU.ComputerName
$t3 = $OU.RemoteAddress.IPAddressToString
$t4 = $OU.RemotePort
$t6 = $ou.SourceAddress.IPAddress
$t7 = $OU.TcpTestSucceeded

if($t7 -match "True"){$OUTPUT.Text = "----------------------------------------------------
PORT $t4 OPENED"
$SERVER.Background = "#FF63FF6A"}
if($t7 -match "False"){$OUTPUT.Text = "----------------------------------------------------
PORT $t4 CLOSED"
$SERVER.Background = "Red"}

$OUTPUT.Text += "
Server: $env:t1
----------------------------------------------------
ComputerName            : $t2
RemoteAddress           : $t3
RemotePort              : $t4
SourceAddress           : $t6
TcpTestSucceeded        : $t7
"
}
if (!$env:portconn){$OUTPUT.Text = "Please type the port number"}
}
if ($CPORT.IsEnabled -eq $false){
$comp = $SERVER.Text
$env:servconn = [System.Net.Dns]::GetHostByName(("$comp")).Hostname
$env:portconn = $PORT.Text
if ($env:portconn){
$tempconn = Start-Job -ScriptBlock {Test-NetConnection -ComputerName $env:servconn -Port $env:portconn -WarningAction SilentlyContinue -ErrorAction SilentlyContinue}
$ou1 = Receive-Job -Job $tempconn -Wait

$env:t1 = $SERVER.text
$t2 = $OU1.ComputerName
$t3 = $OU1.RemoteAddress.IPAddressToString
$t4 = $OU1.RemotePort
$t6 = $ou1.SourceAddress.IPAddress
$t7 = $OU1.TcpTestSucceeded

if($t7 -match "True"){$OUTPUT.Text = "----------------------------------------------------
PORT $t4 OPENED"
$SERVER.Background = "#FF63FF6A"}
if($t7 -match "False"){$OUTPUT.Text = "----------------------------------------------------
PORT $t4 CLOSED"
$SERVER.Background = "Red"}

$OUTPUT.Text += "
Server: $env:t1 
----------------------------------------------------
ComputerName            : $t2
RemoteAddress           : $t3
RemotePort              : $t4
SourceAddress           : $t6
TcpTestSucceeded        : $t7
"
}
}
if (!$env:portconn){$OUTPUT.Text = "Please type the port number"}
}
$TESTPORT.ShowDialog() | out-null
}

function statchart{
$Reader3 = (New-Object System.Xml.XmlNodeReader $xaml3) 
$statchart = [Windows.Markup.XamlReader]::Load($reader3)

$xaml3.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {New-Variable -Name $_.Name -Value $statchart.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

$server.text = $VMNAME.text
$today = Get-Date
$startdate.SelectedDate = "$today"
$enddate.SelectedDate = "$today"

$server.Add_GotFocus({if ($server.Text -eq "VM"){$server.Text = '' ; $server.Foreground = "black"}})
$server.Add_LostFocus({if ($server.text -eq ''){$server.Text = 'VM' ; $server.Foreground = "DarkGray"}})

$SERVER.Add_KeyDown({if ($_.Key -eq 'Enter'){fload}})
$load.Add_Click({fload})
$saveC.Add_Click({savechartPNG})
$saveF.Add_Click({savechartCSV})

$startdate.Add_SelectedDateChanged({
if($enddate.SelectedDate -le $startdate.SelectedDate){$enddate.SelectedDate = $startdate.SelectedDate}
})

$starttime.Add_SelectionChanged({
if ($startdate.text -eq $enddate.text){$endtime.SelectedIndex = $starttime.SelectedIndex}
})

function fload{
$Sdata = ($server.text).trim()
$Sdata = $Sdata.ToUpper()
$server.text = $Sdata

$code = 0
if(!$server.text -or $server.Text -eq 'VM'){[System.Windows.MessageBox]::Show("Please type the VM name (the field VM cannot be blank)","VM ERROR","OK","Error") | Out-Null; $code = 1}
if($enddate.SelectedDate -lt $startdate.SelectedDate){[System.Windows.MessageBox]::Show("The End date cannot be before the Start date","DATE ERROR","OK","Error") | Out-Null; $code = 1}
if($enddate.SelectedDate -gt (Get-Date)){[System.Windows.MessageBox]::Show("The End date cannot be greater than today date","DATE ERROR","OK","Error") | Out-Null; $code = 1}
if($enddate.SelectedDate -eq $startdate.SelectedDate){if($endtime.Text -le $starttime.Text){
[System.Windows.MessageBox]::Show("With the same date, the End time cannot be before or equal to the Start time","TIME ERROR","OK","Error") | Out-Null; $code = 1}}

if($code -eq 0){
if ($HW.Text -eq "CPU"){$stat = "cpu.usage.average"}
if ($HW.Text -eq "RAM"){$stat = "mem.usage.average"}
if ($HW.Text -eq "DISK"){$stat = "disk.usage.average"}
if ($HW.Text -eq "NETWORK"){$stat = "net.usage.average"}

$start = $startdate.Text + " " + $starttime.Text + ":00"
$end = $enddate.Text + " " + $endtime.Text + ":00"

$a = Get-Stat -Entity $server.text -Stat "$stat" -Start "$start" -Finish "$end" | Sort-Object Timestamp | select Timestamp,Value,Unit
$global:outstat = $a

$HWkind = $HW.Text
$hwdate = $a.Timestamp
$hwperc = $a.Value
$hwunit = $a.Unit | select -First 1

if(!$HWchart.Titles){
$HWchart.Titles.Add($HWkind)
}
$HWchart.Titles[0].Text = $HWkind + " (" + $stat + ")"
$HWchart.Titles[0].Font = "Arial,10pt"

$HWchart.ChartAreas[0].Axisx.MajorGrid.Enabled = $False
$HWchart.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yy
 HH:mm"

$HWchart.Series[0].Points.DataBindXY($hwdate,$hwperc)
$HWchart.Series[0].XValueType = "DateTime"
$HWchart.Series[0].Name = $stat

$HWchart.ChartAreas[0].AxisY.Title = $HWkind + ' {' + $hwunit + '}'

$HWchart.Refresh()
}
}

function savechartPNG{
$salva1 = New-Object System.Windows.Forms.SaveFileDialog
$salva1.initialDirectory = "$Env:USERPROFILE\Desktop\"
$salva1.filter = "PNG files (*.png)| *.png"
$salva1.ShowDialog() | Out-Null
$HWchart.SaveImage($salva1.filename, "PNG")
}
function savechartCSV{
$salva1 = New-Object System.Windows.Forms.SaveFileDialog
$salva1.initialDirectory = "$Env:USERPROFILE\Desktop\"
$salva1.filter = "CSV files (*.csv)| *.csv"
$salva1.ShowDialog() | Out-Null
$global:outstat | Export-Csv -Path $salva1.filename -NoTypeInformation
}

$statchart.ShowDialog() | out-null
}

function VMPM_f{
$Reader4 = (New-Object System.Xml.XmlNodeReader $xaml4) 
$VMPM = [Windows.Markup.XamlReader]::Load($reader4)

$xaml4.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {New-Variable -Name $_.Name -Value $VMPM.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

$VMPM.title = $VMNAME.text
$VM = $VMNAME.text

function loadtotext {
$a = Get-VM "$VM"
$b= Get-Stat -Entity $a -Realtime -ErrorAction SilentlyContinue -MaxSamples 1

$bcpu = ($b | Where-Object {$_.MetricId -eq "cpu.usage.average"} | select value).value
$bmem = ($b | Where-Object {$_.MetricId -eq "mem.usage.average"} | select value).value

$c = Get-Stat -Entity $a -Realtime -ErrorAction SilentlyContinue -MaxSamples 1 -Stat "sys.uptime.latest"
$cou = New-Timespan -Seconds $c.Value 
$cou1 = "" + $cou.Days + " Days, " + $cou.Hours + " Hours, " + $cou.Minutes + " Minutes"

$ds = $a.Guest.Disks | Sort-Object Path
$ds1 = $ds | Out-String

$pou1 = foreach ($d in $ds){
$capacity = $d.CapacityGB
$freespace = $d.FreeSpaceGB
$percentage = [math]::Round(($freespace)/($capacity)*100,2)
$perutil = 100 - $percentage
$pou = $d.Path + " " + $perutil + '%' + " " + "Utilizzato"
$pou} 

$pou2 = $pou1 | Out-String

$port = Test-NetConnection -ComputerName $VM -Port 3389 -InformationLevel Quiet
if ($port){$ouport = "RAGGIUNGIBILE IN RDP!"} else {$ouport = "NON RAGGIUNGIBILE IN RDP!"}

if (!$port -or $bcpu -gt "95" -or $bmem -gt "95" -or $perutil -gt "95"){
$OUTPUT.Background = "red"; $OUTPUT.Foreground = "White"; $RESIZE2.Background = "red"; $RESIZE1.Background = "red"; $RESIZE1.Foreground = "White"
} else {$OUTPUT.Background = "black"; $OUTPUT.Foreground = "#FF3AFF00"; $RESIZE2.Background = "#d6d6d6"}

$OUTPUT.Text = "$ouport
$a

SYSTEM UPTIME: 
$cou1

CPU%   MEM%
-----  -----
$bcpu   $bmem

$pou2
$ds1"
}

$RESIZE1.Add_Click({minimizef})
$RESIZE2.Add_Click({restoref})

function minimizef{
$VMPM.Height = 60
$VMPM.Width = 310
$RESIZE2.Visibility = "Visible"
$RESIZE1.Visibility = "Hidden"
}

function restoref{
$VMPM.Height = 310
$VMPM.Width = 310
$RESIZE2.Visibility = "Hidden"
$RESIZE1.Visibility = "Visible"
}

$VMPM.Add_ContentRendered({loadtotext})
$VMPM.Add_MouseDoubleClick({loadtotext})
$VMPM.Add_Closing({$timer.Dispose()})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10000
$timer.Add_Tick({loadtotext})
$timer.Start()

$VMPM.ShowDialog() | out-null
}


$FINDVM.ShowDialog() | out-null