function Populate-COVIDtable ($TargetServer, $TargetDb, $USer, $PWord)
{   
    
    $emailFrom = "if-lw-Covid19Restr-ETL@company.com" 
    $Date = (Get-Date).AddDays(-1).ToString('MM/dd/yyyy')
    [string[]]$emailTo = "Watson.lopes@company.com"#,"Ronald.Thompson1@company.com","Sebastian.John@company.com","Dibya.Nayak@company.com"
    $smtpServer = "e2ksmtp01.e2k.ad.company.com"
    $subject = "Application COVID-19 restructure file: "+$Date
    $PWord = ConvertTo-SecureString -String $PWord -AsPlainText -Force
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord 
    $ConnectedDb = Get-DbaDatabase -SqlInstance $TargetServer -Database $TargetDb -SqlCredential $Credential
    try
    {
        $ConnectedDb.Query("truncate table tbl_covid19restructures")
        $ConnectedDb.Query("CREATE TABLE [dbo].[tbl_tempcovid]([Schedule Number] [varchar](15) NULL, [Customer Number] [varchar](50) NULL, [Request Number] [varchar](15) NULL, [Exposure at time of Restructure] [money] NULL,	[Sales Rep (hierarchy as assigned in SF then LW)] varchar(50) Null,	[Sales Manager] varchar(50) Null, [PM (as assigned in LW)] varchar(50) Null, [RA (as assigned in LW)] varchar(50) Null,	[Pricing Box Folder] varchar(200) Null,	[Pricing Request Date] [nvarchar](20) NULL,	[Pricing Request Status] [varchar](30) NULL, [Pricing Assigned To] [varchar](30) NULL,	[TC Assigned Locating Original Prcing] [varchar](30) NULL,	[TC Status] [varchar](30) NULL,	[Current Step/Assignment] varchar(1000) Null, [Status] [varchar](20) NULL, [Request Date] [nvarchar](20) NULL, [# of Mths] int Null, [Monthly Rent Payment] Money Null, [Total Skip Payments] [money] NULL, [Payment Frequency]  [varchar](15) NULL, [Daily Addition / Withdrawal] [char](1) Null) ON [PRIMARY]")
        #$ConnectedDb.Query("select count(*) from tbl_tempcovid")
   
        # Splat the params (cause it looks nice):
        $ConvertToSqlParams = @{
                    TableName = 'tbl_tempcovid'
                    #$PSScriptRoot
                    Path = "C:\Users\502740204\Documents\Application_Files\CurrentWork\Automations\MyPowerShellAutomations\GERITM8190725_Box_File_Load\CovidData.xlsx"
                    ConvertEmptyStringsToNull = $true
                    UseMSSQLSyntax = $true
                    }
        # Create the insert stmts and save them to variable:
        $SqlInsertStmts = ConvertFrom-ExcelToSqlInsert @ConvertToSqlParams
        #$SqlInsertStmts    
        # Create a connection to the db and save it to a variable:
        # Run the insert stmts:
        $ConnectedDb.Query($SqlInsertStmts)
        $ConnectedDb.Query("Insert into tbl_covid19restructures Select [Schedule Number],[Customer Number],[Request Number],[Exposure at time of Restructure],[Sales Rep (hierarchy as assigned in SF then LW)],[Sales Manager],[PM (as assigned in LW)],[RA (as assigned in LW)],[Pricing Box Folder],CASE WHEN (([Pricing Request Date] is NULL) OR ([Pricing Request Date] = 'NULL')) THEN '1900-01-01' Else CONVERT(DATETIME,CONVERT(INT, [Pricing Request Date])) END,[Pricing Request Status],[Pricing Assigned To],[TC Assigned Locating Original Prcing],[TC Status],[Current Step/Assignment],[Status],CASE WHEN (([Request Date] is NULL) OR ([Request Date] = 'NULL')) THEN '1900-01-01' Else CONVERT(DATETIME,CONVERT(INT, substring([Request Date],1,5))) END,[# of Mths],[Monthly Rent Payment],[Total Skip Payments],[Payment Frequency], [Daily Addition / Withdrawal], Getdate() FROM [tbl_tempcovid]")
        $ConnectedDb.Query("drop table [dbo].[tbl_tempcovid]")
        

        $body = "<font face = calibri>Greetings,<br><br>The COVID-19 restructure file for the date ""<font color = Blue>$Date</Font>"" has been successly loaded into <B>$TargetDb</B>.<br><br><br>Regards,<br>Application Support Team"  
        Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml
    }

    catch
    {
        $body = "<font face = calibri>Greetings,<br><br>There was an error as follows while loading COVID-19 restructure file for the date ""<font color = Blue>$Date</Font>"" into <B>$TargetDb</B>.<br><br><font color = Red>$Error[0]</font><br><br>Please check urgently.<br><br><br>Regards,<br>Application Support Team"
        Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml
    }
}

Populate-COVIDtable 'SQLInstance' 'DB' 'DBUser' 'Password'


#Remove all excel files from the folder
#Get-ChildItem -Path $ExcelPath *.xlsx | foreach { Remove-Item -Path $_.FullName }

