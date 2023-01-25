Try
{
$ErrorActionPreference = 'Stop'
$SQLServer = "SQLInstance"
$SQLDBName = "SQLDB"
$SQLQuery = 'SELECT 
	            ''BillToInvoiceFormats'' AS ''SourceTable''
	            ,BillToInvoiceFormats.Id InvoiceFormatId
	            ,BillToInvoiceFormats.ReceivableCategory
	            ,BillToes.Id AS ''BillToId''
	            ,BillToes.Name AS ''BillToName''
	            ,CreatedByUser BillToCreatedBy
	            ,UpdatedByUser BillToUpdatedBy
	            ,BillToes.CreatedTime BillToCreated
	            ,BillToes.UpdatedTime BillToUpdated
             FROM 
	            BillToes
	            JOIN BillToInvoiceFormats ON BillToes.Id = BillToInvoiceFormats.BillToId AND BillToes.Id = BillToInvoiceFormats.BillToId
	            LEFT JOIN (SELECT BillToes.Id BillToId, Users.Id, Users.FullName CreatedByUser FROM Users INNER JOIN BillToes ON BillToes.CreatedById = Users.Id ) CreatedBy ON BillToes.Id = CreatedBy.BillToId
	            LEFT JOIN (SELECT BillToes.Id BillToId, Users.Id, Users.FullName UpdatedByUser FROM Users INNER JOIN BillToes ON BillToes.UpdatedById = Users.Id) UpdatedBy ON BillToes.Id = UpdatedBy.BillToId
             WHERE
	            BillToInvoiceFormats.InvoiceEmailTemplateId IS NULL
	            and DeliverInvoiceViaEmail = 1
	            and BillToInvoiceFormats.IsActive = 1'

$SQLUsername, $SQLPassword  = Get-LWCreds LWDBUser_ETL

$Records = Get-DBExtract $SQLQuery $SQLServer $SQLDBName $SQLUsername $SQLPassword
#$Records = 0

$Subject = 'LeaseWave Bill Tos without Email Template :' + (Get-Date).AddDays(0).ToString('dd-MMM-yy')                 

If ($Records -ne 0){       
    $Body = "<font face = calibri>Greetings,<br><br> There are <B><font color=RED>"+$Records+"</B></font> Bill Tos without Email Template in LeaseWave.<br><br>PFA detailed report for more information.<br><br><br>Regards,<br>LeaseWave Support Team"
    Send-MailMessage -smtpserver "e2ksmtp01.e2k.ad.ge.com" -from "if-leasewave-NAS-monitor@ge.com" -to "watson.lopes@ge.com" -subject $subject -body $Body -attachments "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\*.csv" -bodyashtml}
ELse{
    $Body = "<font face = calibri>Greetings,<br><br> There are <B><font color=RED>No</B></font> Bill Tos without Email Template in LeaseWave.<br><br><br>Regards,<br>LeaseWave Support Team"
    Send-MailMessage -smtpserver "e2ksmtp01.e2k.ad.ge.com" -from "if-leasewave-NAS-monitor@ge.com" -to "watson.lopes@ge.com" -subject $subject -body $Body -bodyashtml}

} #End Try
Catch
{
    $subject = "LeaseWave Bill Tos without Email Template connectivity issue while extracting: "+$SubjectDate
    $body = "<font face = calibri>Greetings,<br><br>Below is the error message while executing this process -<br><font color = Red>$Error[0]</font><br><br>Please check and retrigger the Job Job/Chain in Cronacle if required.<br><br><br>Regards,<br>LeaseWave Support Team"
    Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml
}
Finally
{
	$ErrorActionPreference = "SilentlyContinue"
	Remove-Item –path D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\*.csv
	$SqlConnection.Close()
	Clear-Variable SQLPassword
}