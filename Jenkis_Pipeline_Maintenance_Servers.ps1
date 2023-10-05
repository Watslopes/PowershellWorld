$AppEnvironment = ${env:AppEnvironment} 
$AppServerNode = ${env:AppServerNode}
$AppOperation = ${env:AppOperation}
$AppComponent = ${env:AppComponent}
$BuildUser = ${env:BUILD_USER}
$BuildUser = $BuildUser.Substring(0, $BuildUser.IndexOf("("))

Switch ($AppEnvironment + "_" + $AppServerNode)
{
    "PROD_LessorPortal-Node1"   {$AppServerFQDN = "Server.Company.com"; $AppIISName = 'LessorPortal.PROD'; $AppPrcName = 'PricingService.PROD'  }
    "PROD_LessorPortal-Node2"   {$AppServerFQDN = "Server.Company.com"; $AppIISName = 'LessorPortal.PROD'; $AppPrcName = 'PricingService.PROD'  }
    "DR_LessorPortal-Node1"   	{$AppServerFQDN = "Server.Company.com"; $AppIISName = 'LessorPortal.DR'; $AppPrcName = 'PricingService.DR'  }
    "DR_LessorPortal-Node2"   	{$AppServerFQDN = "Server.Company.com"; $AppIISName = 'LessorPortal.DR'; $AppPrcName = 'PricingService.DR'  }
    "PROD_LessorPortal-Both"    {$AppServerFQDN = "Server.Company.com", "G2422USPWHEW05V.LOGON.DS.company.com"; $AppIISName = 'LessorPortal.PROD'; $AppPrcName = 'PricingService.PROD' }
    "DR_LessorPortal-Both"    	{$AppServerFQDN = "Server.Company.com", "G2422USRWHEW05V.LOGON.DS.company.com"; $AppIISName = 'LessorPortal.DR'; $AppPrcName = 'PricingService.DR' }
    "PROD_CustomerPortal-Node1" {$AppServerFQDN = "Server.Company.com"; $AppIISName = 'CustomerPortal.PROD'; $AppPrcName = 'PricingService.PROD'  }
    "PROD_CustomerPortal-Node2" {$AppServerFQDN = "Server.Company.com"; $AppIISName = 'CustomerPortal.PROD'; $AppPrcName = 'PricingService.PROD'  }
    "DR_CustomerPortal-Node1" 	{$AppServerFQDN = "Server.Company.com"; $AppIISName = 'CustomerPortal.DR'; $AppPrcName = 'PricingService.DR'  }
    "DR_CustomerPortal-Node2" 	{$AppServerFQDN = "Server.Company.com"; $AppIISName = 'CustomerPortal.DR'; $AppPrcName = 'PricingService.DR'  }
    "PROD_CustomerPortal-Both"  {$AppServerFQDN = "Server.Company.com", "Server.Company.com"; $AppIISName = 'CustomerPortal.PROD'; $AppPrcName = 'PricingService.PROD' }
    "DR_CustomerPortal-Both"  	{$AppServerFQDN = "Server.Company.com", "Server.Company.com"; $AppIISName = 'CustomerPortal.DR'; $AppPrcName = 'PricingService.DR' }
    Default {Write-Host "App server doesn't exist for ($AppEnvironment ""&"" $AppServerNode) combination. Please try again!!!"
         	Exit}
}

Invoke-Command -ComputerName $AppServerFQDN -ScriptBlock {

    $smtpServer = "emailserver.company.com"
    $SubjectDate = (Get-Date).AddDays(0).ToString('dd-MMM-yyyy hh:mm')
    $emailFrom = "Application-maintenance@company.com" 
    $emailTo = @("support@company.com")

  Switch ($Using:AppOperation){
		
        #Stopping Website and AppPool
		"Stop"
            { 
               $Appcmp = $using:AppComponent
               $cnt = ($Appcmp.Split(',').Count) - 1
               Do{
                    If((($Appcmp.Split(',')[$cnt] -eq 'App Job Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'Pricing Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'Shibboleth Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'TValue')) -AND ($Using:AppServerNode -Like '*CustomerPortal*'))
                    {
                        Write-Host "There is no TValue, Shibboleth and App Job Service for CustomerPortal $Using:AppEnvironment-$env:COMPUTERNAME. Please choose the option wisely and retry!!!"
                        Exit
                    }
                 	Switch($Appcmp.Split(',')[$cnt])
                    {
                        "AppPool"
                        {Stop-WebAppPool -Name "$Using:AppIISName"}
						"Website"
                        {Get-Website "$Using:AppIISName" | Stop-Website}
                        "Pricing Service"
                        {Stop-WebAppPool -Name "$Using:AppPrcName"
						Get-Website "$Using:AppPrcName" | Stop-Website}  							
                        "TValue"
                        {Get-Service -Name 'TVA_Service' | Stop-Service}
                        "App Job Service"
                        {Get-Service -Name 'AppJobManager_Service' | Stop-Service}
                        "Shibboleth Service"
                        {Get-Service -Name 'shibd_Default' | Stop-Service}
                        Defaut {Write-Host 'Invalid Choice.. Please try again!!!' Exit}
                    }
                $cnt -= 1
               }While ($cnt -ge 0) 

                Start-Sleep -Seconds 10

                #Email Sending Part
                $ErrorActionPreference = "silentlycontinue"
                $Hostname = ($using:AppServerFQDN).Substring(0,15)
                $body = "<font face = calibri>Greetings,<br><br> Below is the status of Application Services:<br><ul>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> WebSite Status : <B>"+(Get-Website -Name $using:AppIISName).State +"</B></li>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> AppPool Status : <B>"+(Get-WebAppPoolState -Name $using:AppIISName).Value +"</B></li>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> Pricing Service Status : <B>"+(Get-Website -Name $using:AppPrcName).State +"</B></li>				
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> App TVal Service Status : <B>"+(get-service -Name TVA_Service).Status +"</B></li>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> App Job Service Status : <B>"+(get-service -Name AppJobManager_Service).Status +"</B></li>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> Shibboleth Service Status : <B>"+(get-service -Name shibd_Default).Status +"</B></li></ul>
                <br>Regards,<br>Application Support Team" 

                Clear-Variable -Name Hostname
                $subject = $Using:AppEnvironment+": Application Services Stopped: "+$SubjectDate
                Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml                       
            }
        
        #Starting Website and AppPool
		"Start"
            { 
               $Appcmp = $using:AppComponent
               $cnt = ($Appcmp.Split(',').Count) - 1
               Do{
                    If((($Appcmp.Split(',')[$cnt] -eq 'App Job Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'Pricing Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'Shibboleth Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'TValue')) -AND ($Using:AppServerNode -Like '*CustomerPortal*'))
                    {
                        Write-Host "There is no TValue, Shibboleth and App Job Service for CustomerPortal $Using:AppEnvironment-$env:COMPUTERNAME. Please choose the option wisely and retry!!!"
                        Exit
                    }
                 	Switch($Appcmp.Split(',')[$cnt])
                    {                        
                        "Website"
                        {Get-Website "$Using:AppIISName" | Start-Website}
						"AppPool"
                        {Start-WebAppPool -Name "$Using:AppIISName"}
                        "Pricing Service"
                        {Get-Website "$Using:AppPrcName" | Start-Website
						Start-WebAppPool -Name "$Using:AppPrcName"}  						
                        "TValue"
                        {Get-Service -Name 'TVA_Service' | Start-Service}
                        "App Job Service"
                        {Get-Service -Name 'AppJobManager_Service' | Start-Service}
                        "Shibboleth Service"
                        {Get-Service -Name 'shibd_Default' | Start-Service}
                        Defaut {Write-Host 'Invalid Choice.. Please try again!!!' Exit}
                    }
                $cnt -= 1                  
               }While ($cnt -ge 0) 
                Start-Sleep -Seconds 10

                #Email Sending Part
                $ErrorActionPreference = "silentlycontinue"
                $Hostname = ($using:AppServerFQDN).Substring(0,15)
                $body = "<font face = calibri>Greetings,<br><br> Below is the status of Application Services:<br><ul>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> WebSite Status : <B>"+(Get-Website -Name $using:AppIISName).State +"</B></li>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> AppPool Status : <B>"+(Get-WebAppPoolState -Name $using:AppIISName).Value +"</B></li>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> Pricing Service Status : <B>"+(Get-Website -Name $using:AppPrcName).State +"</B></li>								
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> App TVal Service Status : <B>"+(get-service -Name TVA_Service).Status +"</B></li>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> App Job Service Status : <B>"+(get-service -Name AppJobManager_Service).Status +"</B></li>
                <li><U>"+$using:AppEnvironment+"-"+$Hostname+ "</U> Shibboleth Service Status : <B>"+(get-service -Name shibd_Default).Status +"</B></li></ul>
                <br>Regards,<br>Application Support Team" 

                Clear-Variable -Name Hostname
                $subject = $Using:AppEnvironment+": Application Services Started: "+$SubjectDate
                Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml                                    
            }
        
        #Querying Website and AppPool
		"Query"
            { 
               $Appcmp = $using:AppComponent
               $cnt = ($Appcmp.Split(',').Count) - 1
               Do{
                    If((($Appcmp.Split(',')[$cnt] -eq 'App Job Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'Pricing Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'Shibboleth Service') -OR ($Appcmp.Split(',')[$cnt] -eq 'TValue')) -AND ($Using:AppServerNode -Like '*CustomerPortal*'))
                    {
                        Write-Host "There is no TValue, Shibboleth and App Job Service for CustomerPortal $Using:AppEnvironment-$env:COMPUTERNAME. Please choose the option wisely and retry."
                        Exit
                    }
                 	Switch($Appcmp.Split(',')[$cnt])
                    {
                        "Website"
                        {Write-Host "$Using:AppEnvironment-$env:COMPUTERNAME WebSite Status            : " (Get-Website -Name "$Using:AppIISName").State}
                        "AppPool"                                          
                        {Write-Host "$Using:AppEnvironment-$env:COMPUTERNAME AppPool Status            : " (Get-WebAppPoolState -Name "$Using:AppIISName").Value}
                        "Pricing Service"
                        {Write-Host "$Using:AppEnvironment-$env:COMPUTERNAME Pricing Service Status    : " (Get-Website -Name "$Using:AppPrcName").State}						
                        "TValue"                                           
                        {Write-Host "$Using:AppEnvironment-$env:COMPUTERNAME App TVal Service Status    : " (get-service -Name 'TVA_Service').Status}
                        "App Job Service"                                   
                        {Write-Host "$Using:AppEnvironment-$env:COMPUTERNAME App Job Service Status     : " (get-service -Name 'AppJobManager_Service').Status}
                        "Shibboleth Service"                               
                        {Write-Host "$Using:AppEnvironment-$env:COMPUTERNAME Shibboleth Service Status : " (get-service -Name 'shibd_Default').Status}
                        Defaut {Write-Host 'Invalid Choice.. Please try again!!!' Exit}
                    }
                $cnt -= 1
               } While ($cnt -ge 0)           
            }
               		
		Default {
                Write-Host "The ($Using:AppOperation) is invalid operation. Please try again!!!"
                Exit}
	}
}
