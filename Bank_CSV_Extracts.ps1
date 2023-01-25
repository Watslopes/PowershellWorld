Param(
     [Parameter()]
     [string]$SQLUsername = 'UserName', 
     [Parameter()]
     [string]$SQLPassword = 'Password',#'C@shR3port',
     $SQLServer = 'SQLInstance',#"G2422USQWSQL03P.LOGON.DS.company.com\LWMAIN_UAT",
     $SQLDBName = "SQLDB",
     $ReportMonth = (Get-Date).AddDays(-4).ToString('MMM-yyyy'),
     $FolderPath = "NASPath",
     $SubjectDate = (Get-Date).AddDays(0).ToString('dd-MMM-yyyy hh:mm'),
     $smtpServer = "e2ksmtp01.e2k.ad.company.com",

     #Queries for Autmated Buyer’s Extract
     $SQLQuery1  =  "select * from DTP.ContractInformation_Report where effectivemonth = '"+ $ReportMonth +"'",
     $SQLQuery2  =  "Select * from DTP.AssetInformation_Report where effectivemonth = '"+ $ReportMonth +"'",
     $SQLQuery3  =  "Select * from DTP.CorporateGuarantor_Report where effectivemonth = '"+ $ReportMonth +"'",
     $SQLQuery4  =  "Select * from DTP.PersonalGuarantor_Report where effectivemonth = '"+ $ReportMonth +"'",
     $SQLQuery5  =  "Select * from DTP.GoForwardPayments where effectivemonth = '"+ $ReportMonth +"'",
     $SQLQuery6  =  "Select * from DTP.RemainingPaymentStream where effectivemonth = '"+ $ReportMonth +"'",
     $SQLQuery7  =  "Select * from DTP.AssetMakeAndModel where effectivemonth = '"+ $ReportMonth +"'",

     #Queries for Autmated 8LG NEA Extract-
     $SQLQuery8  =  "select * from DTP.US_NEA_Extract_TIAA where effectivemonth = '"+ $ReportMonth +"'"
     )


Try
{
    $ErrorActionPreference = 'Stop'

    #Generate DB Extract in CSV
    Get-DBExtract $SQLQuery1  $SQLServer $SQLDBName -AttachmentName "GEHFS_DT_Contract" -NeedHTMLObj 2
    Get-DBExtract $SQLQuery2  $SQLServer $SQLDBName -AttachmentName "GEHFS_DT_Asset" -NeedHTMLObj 2                                                    
    Get-DBExtract $SQLQuery3  $SQLServer $SQLDBName -AttachmentName "GEHFS_DT_CorporateGuarantor" -NeedHTMLObj 2                                                    
    Get-DBExtract $SQLQuery4  $SQLServer $SQLDBName -AttachmentName "GEHFS_DT_PersonalGuarantor" -NeedHTMLObj 2                                                    
    Get-DBExtract $SQLQuery5  $SQLServer $SQLDBName -AttachmentName "GEHFS_DT_GoForwardPayments" -NeedHTMLObj 2                                                    
    Get-DBExtract $SQLQuery6  $SQLServer $SQLDBName -AttachmentName "GEHFS_DT_RemainingPaymentStream" -NeedHTMLObj 2                                                    
    Get-DBExtract $SQLQuery7  $SQLServer $SQLDBName -AttachmentName "GEHFS_DT_ AssetMakeAndModel" -NeedHTMLObj 2                                                    
    Get-DBExtract $SQLQuery8  $SQLServer $SQLDBName -AttachmentName "GEHFS_DT_NEA" -NeedHTMLObj 2                                                    

    #Remove from NAS Archive and Move files to NAS folder  
    Remove-Item $FolderPath\Archive\*.csv -EA SilentlyContinue
    Move-Item -Path $FolderPath\*.csv -Destination $FolderPath\Archive -EA SilentlyContinue
    Move-Item -Path D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\*.csv -Destination $FolderPath -EA SilentlyContinue
    (Get-ChildItem $FolderPath -File) | rename-item -newname { $_.name.substring(0,$_.name.length-14)+".csv" }
    $cnt = (Get-ChildItem $FolderPath\*.csv).Count

    If($cnt -ne 8) {throw "MyException"}
    Else{
    #Send email with provided parameters
    $body = "<font face = calibri>Greetings,<br><br>There are <B>$cnt</B> files in TIAA NAS Path for <B>$ReportMonth</B> at below NAS drive :<br><br><B><I>$FolderPath</B></I><br><br>No further action is required.<br><br><br>Regards,<br>Application Support Team" 
    $subject = "TIAA NAS MOnthly CSV files Success: "+$SubjectDate
    Send-MailMessage -smtpserver $smtpserver -from 'if-lw_nonProd-maintenance@company.com' -to 'watson.lopes@company.com' -subject $subject -body $body -bodyashtml -Priority High
    }
} #End Try
Catch
{
    If ($error[0].FullyQualifiedErrorID -eq "MyException")
    {
        $body = "<font face = calibri>Greetings,<br><br>There are more than or less than desired number of files (i.e. 8) generated in NAS. The generated files count is <B>$cnt</B> for <B>$ReportMonth</B> at below NAS drive :<br><br><B><I>$FolderPath</B></I><br><br>Please check urgently and take the necessary action to fix this.<br><br><br>Regards,<br>Application Support Team" 
        $subject = "TIAA NAS MOnthly CSV files Failure: "+$SubjectDate
        Send-MailMessage -smtpserver $smtpserver -from 'if-lw_nonProd-maintenance@company.com' -to 'watson.lopes@company.com' -subject $subject -body $body -bodyashtml -Priority High
    }
    Else
    {
        $body = "<font face = calibri>Greetings,<br><br>There was below error while processing TIAA monthly CSV files for <B>$ReportMonth</B>.<br><br><font color = Red>$Error[0]</font><br><br>Please check urgently and take the necessary action to fix this.<br><br><br>Regards,<br>Application Support Team"
        $subject = "TIAA NAS MOnthly CSV files Failure: "+$SubjectDate
        Send-MailMessage -smtpserver $smtpserver -from 'if-lw_nonProd-maintenance@company.com' -to 'watson.lopes@company.com' -subject $subject -body $body -bodyashtml -Priority High
    }
}
Finally
{
	$ErrorActionPreference = "SilentlyContinue"	
	$SqlConnection.Close()
    Remove-Item –path D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\*.csv
	Clear-Variable SQLPassword
}