#Title: Exchange Manager
#Author: Drew Nash
#Date: 10/17/18
#Version: v1.0
#
#Description: GUI console for managing various Exchange Administrative tasks such as: 1)Email Forwarding, 2)Setting Out of Office, 3)Checking members of a distribution list
#
#Instructions: Before initially running this script, input your Exhange server address in the $server variable on line 17. This will allow the script to run faster than 
#              inputting the server each time. Future versions may instead include a config file for this information. The script is also setup to not prompt for additional
#              credentials, so be sure to run the program with credentials that have administrative priveleges on the serer.
#
#Improvements in future coming versions:
#1) Create error popup if an invalid email or action occurs
#2) Fix checkbox for delivering to both addresses. Currently only allows to deliver to both inboxes.
#3) Create dropdown for email address suggestions 
#4) Set later time to remove forward

###Connect current session to Exchange Server
$server = "" #input your Exchange server address in the quotes so that the script does not need to ask for it each time.
$session = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri $server -Authentication Kerberos
Import-pssession $session

###Functions

function setForward{
    param(
        $userMailbox,
        $targetMailbox,
        $checkbox
    )
    if($checkbox.IsChecked = $true){
        Set-Mailbox $userMailbox -ForwardingAddress $targetMailbox -DeliverToMailboxAndForward $true
    }
    else{
        Set-Mailbox $userMailbox -ForwardingAddress $targetMailbox -DeliverToMailboxAndForward $false
    }
}

function removeForward{
    param(
        $userMailbox
    )
    Set-Mailbox $userMailbox -ForwardingAddress $null
}

###WPF Form Creation
$inputXML = @"
<Window x:Class="MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp4"
        mc:Ignorable="d"
        Title="Exchange Manager" Height="250" Width="375" Background="#4c4c4c" WindowStyle="ThreeDBorderWindow" Icon="c:\Scripts\ExchangeManager\Images\exchange.ico">
    <Grid HorizontalAlignment="Left" Height="100" Margin="62,64,0,0" VerticalAlignment="Top" Width="100">
        <TextBox x:Name="UserMailbox" HorizontalAlignment="Left" Height="23" Margin="-37,-38,-17,0" TextWrapping="NoWrap" Text="User Mailbox" VerticalAlignment="Top" Width="154"/>
        <TextBox x:Name="TargetMailbox" HorizontalAlignment="Left" Height="23" Margin="-37,0,-17,0" TextWrapping="NoWrap" Text="Target Mailbox" VerticalAlignment="Top" Width="154"/>
        <CheckBox x:Name="Checkbox1" Content="Deliver to both inboxes?" HorizontalAlignment="Left" Margin="-37,40,-17,0" VerticalAlignment="Top" Width="154" Foreground="White"/>
        <Image x:Name="ExchangeIcon" HorizontalAlignment="Left" Height="93" Margin="167,-38,-167,0" VerticalAlignment="Top" Width="100" Source="c:\Scripts\ExchangeManager\Images\exchange.png"/>
        <Button x:Name="Button2" Content="Remove Forward" HorizontalAlignment="Left" Height="25" Margin="-37,69,0,0" VerticalAlignment="Top" Width="101"/>
        <Button x:Name="Button1" Content="Create Forward" HorizontalAlignment="Left" Height="24" Margin="166,70,-167,0" VerticalAlignment="Top" Width="101"/>
    </Grid>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}
catch{
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}

$xaml.SelectNodes("//*[@Name]") | %{"trying item $($_.Name)";
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }

###Set default properties/values for form variables

$WPFButton1.ClickMode="Press"
$WPFButton2.ClickMode="Press"
$WPFCheckbox1.IsChecked = $true

###Associate Functions with the form

$WPFButton1.Add_Click({setForward -userMailbox $WPFUserMailbox.Text -targetMailbox $WPFTargetMailbox.Text -checkbox $WPFCheckbox1})
$WPFButton2.Add_Click({removeForward -userMailbox $WPFUserMailbox.Text})

$Form.ShowDialog() | Out-Null