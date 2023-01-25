$User = 'logon\lg688521sv'
$PWord = ConvertTo-SecureString -String 'Gr34test_T34m' -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord 

Enter-PSSession -ComputerName SERVERNAME -Credential $Credential

Invoke-Command -FilePath D:\LW_DM_ETL_SSIS\App_Maintenance\COVID19ExcelLoad\COVID19ExcelLoad.ps1 -ComputerName SERVERNAME

Exit-PSSession
