$inputXML = @"
<Window x:Class="PS_API_Test.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PS_API_Test"
        mc:Ignorable="d"
        Title="MainWindow" Height="600" Width="885.883" ScrollViewer.VerticalScrollBarVisibility="Auto">
    <Grid RenderTransformOrigin="0.507,0.583" ScrollViewer.VerticalScrollBarVisibility="Auto">
        <Grid.Background>
            <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
                <GradientStop Color="Black" Offset="0"/>
                <GradientStop Color="#FFF5C0B2" Offset="0.013"/>
            </LinearGradientBrush>
        </Grid.Background>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition/>
            <ColumnDefinition Width="866"/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Label Content="API Type" HorizontalAlignment="Left" Margin="4.667,28,0,0" VerticalAlignment="Top" Height="26" Width="69" Grid.Column="3" FontWeight="Bold"/>
        <ComboBox x:Name="combo_apitype" HorizontalAlignment="Left" Margin="103.667,32,0,0" VerticalAlignment="Top" Width="148" Height="22" Grid.Column="3">
            <ComboBoxItem Background="#FFF4DDDD" Content="GET" IsSelected="True"/>
            <ComboBoxItem Content="POST"/>
        </ComboBox>
        <Label Content="API URL" HorizontalAlignment="Left" Margin="4.667,63,0,0" VerticalAlignment="Top" Height="26" Width="64" Grid.Column="3" FontWeight="Bold"/>
        <TextBox x:Name="txtbox_apiurl" HorizontalAlignment="Left" Height="40" Margin="103.667,63,0,0" TextWrapping="Wrap" Text="https://sit-node1.if-Application.corporate.company.com:81/api/Entity/" VerticalAlignment="Top" Width="722" RenderTransformOrigin="0.5,0.5" Grid.Column="3">
            <TextBox.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform AngleX="-0.002"/>
                    <RotateTransform Angle="0.004"/>
                    <TranslateTransform/>
                </TransformGroup>
            </TextBox.RenderTransform>
        </TextBox>
        <Label Content="PayLoad" HorizontalAlignment="Left" Margin="4.667,100,0,0" VerticalAlignment="Top" Height="26" Width="67" Grid.Column="3" FontWeight="Bold"/>
        <TextBox x:Name="txtbox_payload" HorizontalAlignment="Left" Height="125" Margin="103.667,110,0,0" Text="Provide Payload in JSON/XML (Not applicable for GET)" VerticalAlignment="Top" Width="719" TextWrapping="Wrap" AcceptsReturn="True" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Auto" Grid.Column="3" />
        <Label Content="Response" HorizontalAlignment="Left" Margin="1.667,410,0,0" VerticalAlignment="Top" Height="26" Width="75" Grid.Column="3" FontWeight="Bold"/>
        <TextBox x:Name="txtbox_response" HorizontalAlignment="Left" Height="137" Margin="103.667,419,0,0" Text="Response Body" VerticalAlignment="Top" Width="719" TextWrapping="Wrap" AcceptsReturn="True" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Auto" Grid.Column="3" IsReadOnly="True" />
        <Button x:Name="btn_send" Content="Send" Grid.Column="3" HorizontalAlignment="Left" Margin="344.667,243,0,0" VerticalAlignment="Top" Width="209" RenderTransformOrigin="-1.365,0.632" Height="32"/>
        <Label Content="ResponseMessage" HorizontalAlignment="Left" Margin="199.667,298,0,0" VerticalAlignment="Top" Height="26" Width="115" Grid.Column="3" FontWeight="Bold" />
        <Label Content="ResponseStatus" HorizontalAlignment="Left" Margin="1.667,298,0,0" VerticalAlignment="Top" Height="26" Width="96" Grid.Column="3" FontWeight="Bold"/>
        <Label x:Name="lbl_status" Content="StatusValue" HorizontalAlignment="Left" Margin="111.667,298,0,0" VerticalAlignment="Top" Height="26" Width="80" Grid.Column="3" RenderTransformOrigin="0.5,0.5" Foreground="Blue" >
            <Label.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="0.272"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Label.RenderTransform>
        </Label>
        <Label x:Name="lbl_message" Content="MessageValue" HorizontalAlignment="Left" Margin="319.667,298,0,0" VerticalAlignment="Top" Height="26" Width="488" Grid.Column="3" Foreground="Blue" />
        <Label Content="Body Type" HorizontalAlignment="Left" Margin="371.667,28,0,0" VerticalAlignment="Top" Height="26" Width="81" Grid.Column="3" FontWeight="Bold"/>
        <RadioButton x:Name="radio_JSON" Content="JSON" Grid.Column="3" HorizontalAlignment="Left" Margin="489.667,35,0,0" VerticalAlignment="Top" IsChecked="True" Height="19" Width="52"/>
        <RadioButton x:Name="radio_XML" Content="XML" Grid.Column="3" HorizontalAlignment="Left" Margin="568.667,35,0,0" VerticalAlignment="Top" Height="19" Width="52"/>
        <Label Content="ExecutionMessage" HorizontalAlignment="Left" Margin="199.667,342,0,0" VerticalAlignment="Top" Height="26" Width="115" Grid.Column="3" FontWeight="Bold"/>
        <Label Content="ExecutionStatus" HorizontalAlignment="Left" Margin="1.667,342,0,0" VerticalAlignment="Top" Height="26" Width="112" Grid.Column="3" FontWeight="Bold"/>
        <Label x:Name="lbl_status_exec" Content="N/A" HorizontalAlignment="Left" Margin="111.667,342,0,0" VerticalAlignment="Top" Height="26" Width="80" Grid.Column="3" RenderTransformOrigin="0.5,0.5" Foreground="Blue" >
            <Label.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="0.272"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Label.RenderTransform>
        </Label>
        <Label x:Name="lbl_message_exec" Content="N/A" HorizontalAlignment="Left" Margin="319.667,342,0,0" VerticalAlignment="Top" Height="51" Width="503" Grid.Column="3" Foreground="Blue" />
    </Grid>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
#$inputXML

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')


#===========================================================================
#Read XAML
#===========================================================================
[xml]$XAML = $inputXML

$reader=(New-Object System.Xml.XmlNodeReader $xaml)

try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Error "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Create PowerShell variables for Form Objects
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | %{ 
    Set-Variable -Name "gui_$($_.Name)" -Value $Form.FindName($_.Name) 
    }

#--------------------------------------------------------------------------------------------------------------------------------------------------
Function Get-FormVariables{
        write-host "Here is the list of GUI variables: " -ForegroundColor Cyan
        get-variable gui_*
    }
#--------------------------------------------------------------------------------------------------------------------------------------------------

#To generate LW Token
<#function Get-LWAPIToken2{	 
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", "Basic U0lULldlYkFQSVVzZXI6Zz8jfEEwY1U=")

    $response = Invoke-RestMethod -Uri 'https://if-Application.corporate.company.com:81/api/authentication/login' -Method 'GET' -Headers $headers #-ContentType "application/x-www-form-urlencoded"	
    Return $response
}#>
$creds = Get-Credential $null

$eventHandler = [System.Windows.RoutedEventHandler]{
        try{    
            If ($gui_radio_JSON.IsChecked -eq 'True'){$contenttype = 'application/json'}
            If ($gui_radio_XML.IsChecked -eq 'True'){$contenttype = 'application/xml'}
            
            #$getLWToken = Get-LWAPIToken2
            #$getLWToken = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($getLWToken))
  
            If ($gui_combo_apitype.Text -eq 'GET')
            {$RestRqst = Invoke-WebRequest -Uri $gui_txtbox_apiurl.Text -Method $gui_combo_apitype.Text -ContentType $contenttype -Credential $creds} 
            Else 
            {$body = $gui_txtbox_payload.Text; $RestRqst = Invoke-WebRequest -Uri $gui_txtbox_apiurl.Text -Method $gui_combo_apitype.Text -Body $body -ContentType $contenttype -Credential $creds}

            #$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            #$headers.Add("Authorization", "Basic U0lULldlYkFQSVVzZXI6Zz8jfEEwY1U=")
            $response = $RestRqst

            if ($response.StatusCode -ne "") {$gui_lbl_status.Content = $response.StatusCode}
            if ($response.StatusDescription -ne "") {$gui_lbl_message.Content = $response.StatusDescription}
            $gui_txtbox_response.Text = $response.Content -replace 'ï»¿', ''   
            
            If ($gui_lbl_status.Content -eq '200'){$gui_lbl_status.Foreground = 'Green'; $gui_lbl_message.Foreground = 'Green'}
            Else {$gui_lbl_status.Foreground = 'Red'; $gui_lbl_message.Foreground = 'Red'}

            If ($gui_combo_apitype.Text -ne 'GET'){
                If($gui_radio_JSON.IsChecked -eq 'True'){$responseobj = $response.Content -replace 'ï»¿', '' | Out-String | ConvertFrom-JSON}
                Else{$responseobj = $response.Content}

                $gui_lbl_status_exec.Content = $responseobj.Success
                $m = @()
                $m =  $responseobj.Messages -replace '"','' -replace ']','' -replace '\[','' -split ","
                [string]$messages = $m
                $gui_lbl_message_exec.Content = $messages

                #If ($gui_lbl_status_exec.Content -eq 'True'){$gui_lbl_status_exec.Content = 'Success'; $gui_lbl_status_exec.Foreground = 'Green'; $gui_lbl_message_exec.Foreground = 'Green'}
                #Else {$gui_lbl_status_exec.Content = 'Failure'; $gui_lbl_status_exec.Foreground = 'Red'; $gui_lbl_message_exec.Foreground = 'Red'}
                Switch($gui_lbl_status_exec.Content)
                {
                    "True"{
                        $gui_lbl_status_exec.Content = 'Success'; $gui_lbl_status_exec.Foreground = 'Green'; $gui_lbl_message_exec.Foreground = 'Green'
                    }

                    "False"{
                        $gui_lbl_status_exec.Content = 'Failure'; $gui_lbl_status_exec.Foreground = 'Red'; $gui_lbl_message_exec.Foreground = 'Red' 
                    }

                    Default{
                        $gui_lbl_status_exec.Content = 'N/A'; $gui_lbl_message_exec.Content = 'N/A'; $gui_lbl_status_exec.Foreground = 'BLue'; $gui_lbl_message_exec.Foreground = 'Blue'
                    }
                }
            }

           }

    catch{
        $ErrorActionPreference = 'SilentlyContinue' 
        $gui_lbl_status.Content = 500
        $gui_lbl_message.Content = "Error during API call. Check if the Payload exist and/or correct the other Excel inputs"
        $gui_lbl_status.Foreground = 'Red'
        $gui_lbl_message.Foreground = 'Red'
    }    
}

$gui_btn_send.Add_Click($eventHandler)

#--------------------------------------------------------------------------------------------------------------------------------------------------

#$Form.Controls.Add($gui_lbl_status);
#$Form.Controls.Add($gui_lbl_message);
#$Form.Controls.Add($gui_txtbox_response);
$ret = $form.ShowDialog()

Invoke-RestMethod -Uri "https://if-Application.corporate.company.com:81/api/authentication/logoff" -Method GET -Credential $creds