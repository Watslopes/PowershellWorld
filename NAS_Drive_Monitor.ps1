##Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

$FolderPath = "\\NASDrive"
$inputpath = $args[0]
$SubjectDate = (Get-Date).AddDays(0).ToString('dd-MMM-yyyy hh:mm')
$emailFrom = "if-Application-NAS-monitor@company.com" 
$emailTo = @("hef.Application.support@company.com")
[string[]]$emailCc = "charles.nash@company.com","christine.bohte@company.com"
$smtpServer = "e2ksmtp01.e2k.ad.company.com"
$environemnt = "PROD"

Try
{
$ErrorActionPreference = 'Stop'
#Dir $FolderPath"\"$inputpath | Select-Object -Property Name -First 1

Switch ($inputpath)
{
    ##Case 1 LockBox- Check the files for 06:00 AM/05:30 PM ET LockBox job
    "$environemnt\Integrations\Incoming\Cash\LB_Wires"
    {
        #Region 1
        $DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
        $FilePattern = "LB_AutoCash_PNCCUS33XXX_*.csv"
        #Path folder
        $checkForFiles = $FolderPath+"\"+$inputpath+"\"+$FilePattern
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #$fileExistence
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
	        {      
              $L_Msg = ''
              $Files = Get-ChildItem $FolderPath"\"$inputpath -File | where {$_.name -like $FilePattern} -ErrorAction Stop
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $filesize = (Get-Item $FolderPath"\"$inputpath"\"$filename).length/1024
                    $filesize = ([Math]::Round($filesize, 2))
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filesize + '</td><td>' + $filetime + '</td></tr>'
                }

              $body = "<font face = calibri>Greetings,<br><br> Below is/are the file(s) present in <b>LockBox</b> NAS drive location (" + $inputpath +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>File Size (in Kb)</th><th>Creation Time</th></tr>" + $L_Msg +" </table><br>Kindly monitor the today's LockBox job in Application Prod.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS LockBox files received for the day: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
	        } 
              # If no files exist, send missing file email
              else
	        {
                $body = "<font face = calibri>Greetings,<br><br><font color = Red> We have not received the <b>LockBox</b> files for today yet in the NAS location (" + $inputpath +"). </font><br><br>Kindly take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
                $subject = $environemnt+": NAS LockBox missing files: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
	        } 

        #Check for the files without .txt or .csv extensions
        If ((Get-ChildItem $FolderPath"\"$inputpath -Exclude *.txt,*.csv,Archive,Failure,PMSCash,ManualRerun,LockBoxArchieve,LockBoxFailure,Hold -ErrorAction Stop).Count -gt 0)
        {
		$L_Msg = ''
        $Files = Get-ChildItem $FolderPath"\"$inputpath -Exclude *.txt,*.csv,Archive,Failure,PMSCash,ManualRerun,LockBoxArchieve,LockBoxFailure,Hold -ErrorAction Stop
            foreach ($f in $Files)
            {
                $filename = $f.Name
				$filetime = $f.CreationTime
                $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
            }          
            $body = "<font face = calibri>Greetings,<br><br><font color = Red> We have received the <b>LockBox</b> files for today with the Wrong extentions in the NAS location (" + $inputpath +"): </font><br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>Kindly take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
            $subject = $environemnt+": NAS LockBox files with wrong Extentions: "+$SubjectDate
        Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
	    }
    }

    ##Case 2 LockBox\Failure- Check if the any files failed while processing and went to the failure folder at 06:30 PM ET
    "$environemnt\Integrations\Incoming\Cash\LB_Wires\Failure"
    {
    #Region 2
     #Path folder
        $DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
        $FilePattern = "LB_AutoCash_PNCCUS*"+ $DateStr +"*"
         #Path folder
        $checkForFiles = $FolderPath + "\"+$inputpath+"\"+$FilePattern
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #$fileExistence
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
	        {      
              $L_Msg = ''
              $Files = Get-ChildItem $FolderPath"\"$inputpath -File | where {$_.name -like $FilePattern} -ErrorAction Stop
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $Measure = Get-Content $f.FullName | Measure-Object -ErrorAction Stop
                    $LineCount = $Measure.Count
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $LineCount + '</td><td>' + $filetime + '</td></tr>'
                }

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) is/are present in <b>LockBox</b> failure location in the NAS drive location (" + $inputpath +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left  bgcolor = #6495ED><th>File Name</th><th>No. of Records</th><th>Creation Time</th></tr>" + $L_Msg +" </table><br>Kindly take the necessary action.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS LockBox failure file: "+$SubjectDate 
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -cc $emailCc -subject $subject -body $body -bodyashtml -Priority High
	        }
            Else
            {  
              $subject = $environemnt+": NAS LockBox failure empty: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to "watson.lopes@company.com" -subject $subject
            }
    }

    ##Case 3 LockBox\Archive- Check the archive folder for the day and send informative email to team at 06:30 PM ET
    "$environemnt\Integrations\Incoming\Cash\LB_Wires\Archive"
    {
    #Region 3
        $DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
        $FilePattern = 'LB_AutoCash_PNCCUS33XXX_*'+ $DateStr +'*.csv'
         #Path folder
        $checkForFiles = $FolderPath+"\"+$inputpath+"\"+$FilePattern
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #$fileExistence
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
	        {      
              $L_Msg = ''
              $L_Cnt = 0
              $Files = Get-ChildItem $FolderPath"\"$inputpath -File | where {$_.name -like $FilePattern} -ErrorAction Stop                
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $Measure = Get-Content $f.FullName | Measure-Object -ErrorAction Stop
                    $LineCount = $Measure.Count
                    $L_Cnt = $L_Cnt + $LineCount
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $LineCount + '</td></tr>'
                }
                $L_Msg = $L_Msg +'<tr align = "left"  bgcolor = #FED8B1><font color="Blue"><B><td>Total no. of LockBox records processed for the day</td><td>' + $L_Cnt + '</td></font></b></tr>'
        #SQL Part starts
            # This code defines the search string in the IssueTrak database tables
            $SQLServer = "SQLInstance"
            $SQLDBName = "SQLDB"
            $SQLUsername = "Username"
            $SQLPassword = "Password"
            $SQLQuery = "Select
                        Postdate,
                        COUNT(DISTINCT GUID) AS Total_Receipts,
                        SUM(CASE WHEN Balance_Amount <> '0' THEN 0 ELSE 1 END) AS Receipts_with_Zero_Balance_Amount,
                        cast(cast(100.0 * SUM(CASE WHEN Balance_Amount <> '0' THEN 0 ELSE 1 END)/COUNT(DISTINCT GUID) as decimal(18,2)) as varchar(20)) + '%' As [Auto-Cash_Hit_Rate]
                        From Receipts(nolock) 
                        Where GUID Like 'TB%' And postdate = (select CurrentBusinessDate from businessunits(nolock)) --and CAST([createdtime] as time) > CAST('17:30' as time)
                        GROUP BY PostDate"

            # This code connects to the SQL server and retrieves the data
            $SQLConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; uid = $SQLUsername; pwd = $SQLPassword"

            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandText = $SqlQuery
            $SqlCmd.Connection = $SqlConnection

            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SqlAdapter.SelectCommand = $SqlCmd

            $DataSet = New-Object System.Data.DataSet
            $SqlAdapter.Fill($DataSet)
            $SqlConnection.Close()

            # This code outputs the retrieved data
            $table = $DataSet.Tables[0]

            $htmltbl = "<table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr bgcolor = #6495ED><td>Postdate</td><td>Total_Receipts</td><td>Receipts_with_Zero_Balance_Amount</td><td>Auto-Cash_Hit_Rate</td></tr>"
            foreach ($row in $table.Rows)
            {
			    if($row[3] -gt 75)
                    {$color = 'Green'}
                else
                    {$color = 'Red'}
                $htmltbl += "<tr><td>" + $row[0] + "</td><td><font color = Blue>" + $row[1] + "</font></td><td>" + $row[2] + "</td><td><b><font color = $color>" + $row[3] + "</font></b></td></tr>"
            }
            $htmltbl += "</table>"

        #SQL Part ends

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) are present in <b>LockBox</b> Archive path in the NAS drive location (" + $inputpath +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left  bgcolor = #6495ED><th>File Name</th><th>No. of Records</th></tr>" + $L_Msg +"</table><br>Below are the LockBox records got loaded into Application database today :<br><br>"+ $htmltbl +" <br><B>Note : </B> For Monday evening's run, The 'AutoCash HitRate %' will be shown for the receipts processed for the entire day (i.e. combining rececipts for morning's run as well).<br><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS LockBox files Archive analysis: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -cc $emailCc -subject $subject -body $body -bodyashtml
            }

        $FilePattern = "Wire_AutoCash_*"+ $DateStr +"*"
         #Path folder
        $checkForFiles = $FolderPath + "\"+$inputpath+"\"+$FilePattern
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #$fileExistence
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
            {
             $L_Msg = ''
             $Files = Get-ChildItem $FolderPath"\"$inputpath -File | where {$_.name -like $FilePattern} -ErrorAction Stop 
                foreach ($f in $Files)
                {
                    $filename = $f.Name
					$filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }
              $body = "<font face = calibri>Greetings,<br><br> Below file(s) are present in <b>AutoCash Wire</b> Archive path for today in the NAS drive location (" + $inputpath +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS Autocash Wire files Archive analysis: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
            }
    }

    ##Case 4 ACH/PAAP- Check the NAS ACH/PAAP folder for the day and send email to team at 06:00 PM ET
    "$environemnt\Integrations\Outgoing\Cash\ACH"
    {
    #Region 4
        #Path folder
        $checkForFiles = $FolderPath + "\"+$inputpath+"\*.*"
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
            {
                $body = "<font face = calibri>Greetings,<br><br><font color = Red> The <b>ACH/PAAP</b> files are still present in the NAS drive location (" + $inputpath +"). </font><br><br>Kindly check with SOA team and take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
                $subject = $environemnt+": NAS ACH/PaaP files not sent to SOA: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
            }
            else
	        {      
              $DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
              $L_Msg = ''
              $L_Cnt = 0
              $Files = Get-ChildItem $FolderPath"\"$inputpath"\Archive\"$DateStr -File -ErrorAction Stop       
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }
                #$L_Msg = $L_Msg +'<tr align = "left"><font color="Blue"><B><td>Total no. of ACH/PaaP records processed for the day</td><td>' + $L_Cnt + '</td></font></b></tr>'

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) are present in <b>ACH/PaaP</b> Archive path for today in the NAS drive location (" + $inputpath +"\Archive\"+ $DateStr +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS ACH/PaaP files Archive analysis: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
            }
    }

    ##Case 5 PositivePay- Check the NAS PositivePay folder for the day and send email to team at 06:30 PM ET
    "$environemnt\Integrations\Outgoing\Cash\PositivePay"
    {
    #Region 5
        #Path folder
        $checkForFiles = $FolderPath + "\"+$inputpath+"\*.*"
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #$fileExistence
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
            {
                $body = "<font face = calibri>Greetings,<br><br><font color = Red> The <b>PositivePay</b> files are still present in the NAS drive location (" + $inputpath +"). </font><br><br>Kindly check with SOA/Glabalscape team and take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
                $subject = $environemnt+": NAS PositivePay files not sent to SOA: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
            }
            else
	        {      
              $DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
              $L_Msg = ''
              $L_Cnt = 0
              $Files = Get-ChildItem $FolderPath"\"$inputpath"\Archive\"$DateStr -File -ErrorAction Stop        
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }
                #$L_Msg = $L_Msg +'<tr align = "left"><font color="Blue"><B><td>Total no. of ACH/PaaP records processed for the day</td><td>' + $L_Cnt + '</td></font></b></tr>'

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) are present in <b>PositivePay</b> Archive path for today in the NAS drive location (" + $inputpath +"\Archive\"+ $DateStr +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS PositivePay files Archive analysis: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
            }
    }

    ##Case 6 WebCash- Check the NAS WebCash folder for the day and send email to team at 11:20 AM ET, 2:20 PM, 4:50 PM ET. QC- 5:50 PM ET
    "$environemnt\Integrations\Outgoing\Cash\WebCash"
    {
    #Region 6
        #Path folder
        $checkForFiles = $FolderPath + "\"+$inputpath+"\*.*"
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #$fileExistence
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
            {
                $body = "<font face = calibri>Greetings,<br><br><font color = Red> The <b>WebCash</b> files are still in the NAS drive location (" + $inputpath +") : </font><br><br>Kindly check with SOA/Glabalscape team and take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
                $subject = $environemnt+": NAS WebCash files not sent to SOA: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
            }
            else
	        {      
              $DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
              $L_Msg = ''
              $L_Cnt = 0
              $Files = Get-ChildItem $FolderPath"\"$inputpath"\Archive\"$DateStr -File -ErrorAction Stop           
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }
                #$L_Msg = $L_Msg +'<tr align = "left"><font color="Blue"><B><td>Total no. of ACH/PaaP records processed for the day</td><td>' + $L_Cnt + '</td></font></b></tr>'

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) are generated in <b>WebCash</b> Archive path for today in the NAS drive location (" + $inputpath +"\Archive\"+ $DateStr +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS WebCash files Archive analysis: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
            }
    }

    ##Case 7 FiWare- Check the files for 07:10 PM ET FiWare job
    "$environemnt\LWAppFiles\Fiware"
    {
        #Region 7
        #Path folder
		$DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
        $FilePattern = "*"+ $DateStr +"*_Request.xml"
        $checkForFiles = $FolderPath + "\"+$inputpath+"\"+$FilePattern
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #$fileExistence
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
	        {      
              $L_Msg = ''
              $Files = Get-ChildItem $FolderPath"\"$inputpath -File | where {$_.name -like $FilePattern} -ErrorAction Stop
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }
                #$L_Msg = $L_Msg +'<tr align = "left"><font color="Blue"><B><td>Total no. of ACH/PaaP records processed for the day</td><td>' + $L_Cnt + '</td></font></b></tr>'

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) have been generated in <b>FiWare</b> path for today in the NAS drive location (" + $inputpath +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS FiWare files generated for today: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
	        } 
              # If no files exist, send missing file email
              else
	        {
                $body = "<font face = calibri>Greetings,<br><br><font color = Red> The <b>FiWare</b> files have not been generated for today yet in the NAS drive location (" + $inputpath +"). </font><br><br>Kindly take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
                $subject = $environemnt+": NAS FiWare missing files for today: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
	        } 
    }

    ##Case 8 ACHReturns- Check the NAS ACHReturens and ACHReturnsProcess folders for the day and send email to team at 07:10 PM ET
    "$environemnt\Integrations\Incoming\Cash\ACHReturn"
    {
    #Region 8
        #Path folder
        $checkForFiles = $FolderPath + "\"+$inputpath+"\*.*"
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #$fileExistence
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
            {
                $body = "<font face = calibri>Greetings,<br><br><font color = Red> The <b>ACHReturns</b> files are still present in the NAS drive location (" + $inputpath +"). </font><br><br>Kindly check with SOA team and take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
                $subject = $environemnt+": NAS ACHReturns files not sent to SOA: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
            }
            else
	        {      
              $DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
              $L_Msg = ''
              $L_Cnt = 0
              $FilePattern = "CIS_Application_PNC_USD_ACH*"+ $DateStr +"*"
              $Files = Get-ChildItem $FolderPath"\"$inputpath"\ACHReturnsProcess" -File | where {$_.name -like $FilePattern} -ErrorAction Stop         
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }
                #$L_Msg = $L_Msg +'<tr align = "left"><font color="Blue"><B><td>Total no. of ACH/PaaP records processed for the day</td><td>' + $L_Cnt + '</td></font></b></tr>'

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) are present in <b>ACHReturns\ACHReturnsProcess</b> path for today in the NAS drive location (" + $inputpath +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS ACHReturns\ACHReturnsProcess files Archive analysis: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
            }
    }

    ##Case 9 ACHReturns\ACHReturnsFailure- Check if the any files failed while processing and went to the failure folder at 07:10 PM ET
    "$environemnt\Integrations\Incoming\Cash\ACHReturn\ACHReturnsFailure"
    {
    #Region 9
        #Path folder
        $DateStr = (Get-Date).AddDays(0).ToString('yyyyMMdd')
        $FilePattern = "CIS_Application_PNC_USD_ACH*"+ $DateStr +"*"
        $checkForFiles = $FolderPath + "\"+$inputpath+"\"+$FilePattern
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #check for the existence of files in the folder
            if ($fileExistence -eq $true) 
	        {      
              $L_Msg = ''
              $Files = Get-ChildItem $FolderPath"\"$inputpath -File | where {$_.name -like $FilePattern} -ErrorAction Stop
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }

              $body = "<font face = calibri>Greetings,<br><br> Below is the files present in <b>ACHReturnsFailure</b> path in the NAS drive location (" + $inputpath +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +" </table><br>Kindly take the necessary action.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS ACHReturnsFailure file generated: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -cc $emailCc -subject $subject -body $body -bodyashtml -Priority High
	        }
            Else
            {              
              $subject = $environemnt+": NAS ACHReturnsFailure empty: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to "watson.lopes@company.com" -subject $subject
            }
    }

	##Case 10 Bridger Outgoing- Check the NAS Bridger folder for the weekly/monthly and send email to team at Sunday 10:30 PM ET & 1st of EveyMonth at 03:30 AM ET
    "$environemnt\Integrations\Outgoing\Bridger"
    {
    #Region 10
        #Path folder
        $DateStr = (Get-Date).AddDays(0).ToString('MMddyyyy')
        $DateStr1 = (Get-Date).AddDays(0).ToString('yyyyMMdd')
        $DateStr2 = (Get-Date -day 1 -hour 0 -minute 0 -second 0).ToString('yyyyMMdd')
        #Check if today is a 1st day of a month
        if($DateStr1 -eq $DateStr2)
        {
            $FilePattern = "BRIDGER*_M_"+ $DateStr +"*"
            $Tag = "Monthly"
        }
        else
        {
            $FilePattern = "BRIDGER*_W_"+ $DateStr +"*"
            $Tag = "Weekly"
        }
        
        $checkForFiles = $FolderPath + "\"+$inputpath+"\"+$FilePattern
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #check for the existence of files in the folder
            if ($fileExistence -ne $true) 
            {
                $body = "<font face = calibri>Greetings,<br><br><font color = Red> The <b>Bridger Outgoing "+$Tag+" </b> files are not created in the NAS drive location (" + $inputpath +"). </font><br><br>Kindly check the Application job status and take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
                $subject = $environemnt+": NAS Bridger Outgoing "+$Tag+" files not generated: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
            }
            else
	        {      
              $L_Msg = ''
              $L_Cnt = 0
              $Files = Get-ChildItem $FolderPath"\"$inputpath -File  | where {$_.name -like $FilePattern} -ErrorAction Stop           
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) are present in <b>Bridger Outgoing "+$Tag+" </b> Archive path for today in the NAS drive location (" + $inputpath +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS Bridger Outgoing "+$Tag+" files Archive analysis: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
            }
    }

	##Case 11 Bridger Incoming- Check the NAS Bridger folder for the weekly/monthly and send email to team at Sunday 2:00 AM ET & 1st of EveyMonth at 07:00 AM ET
    "$environemnt\Integrations\Incoming\Bridger"
    {
    #Region 11
        #Path folder
        $DateStr = (Get-Date).AddDays(0).ToString('MMddyyyy')
        $DateStr1 = (Get-Date).AddDays(0).ToString('yyyyMMdd')
        $DateStr2 = (Get-Date -day 1 -hour 0 -minute 0 -second 0).ToString('yyyyMMdd')
        #Check if today is a 1st day of a month
        if($DateStr1 -eq $DateStr2)
        {
            $FilePattern = "BRIDGER*_M_"+ $DateStr +"*out"
            $Tag = "Monthly"
        }
        else
        {
            $FilePattern = "BRIDGER*_W_"+ $DateStr +"*out"
            $Tag = "Weekly"
        }
        
        $checkForFiles = $FolderPath + "\"+$inputpath+"\"+$FilePattern
        #Test for the existence of files
        $fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
        #check for the existence of files in the folder
            if ($fileExistence -ne $true) 
            {
                $body = "<font face = calibri>Greetings,<br><br><font color = Red> The <b>Bridger Incoming "+$Tag+" </b> files are not received in the NAS drive location (" + $inputpath +"). </font><br><br>Kindly check the Application job status and take the necessary action on an urgent basis.<br><br><br>Regards,<br>Application Support Team"  
                $subject = $environemnt+": NAS Bridger Incoming "+$Tag+" files not received: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml -Priority High
            }
            else
	        {
              $L_Msg = ''
              $L_Cnt = 0
              $Files = Get-ChildItem $FolderPath"\"$inputpath"\Archive\"$DateStr1 -File -ErrorAction Stop   
                foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filetime + '</td></tr>'
                }

              $body = "<font face = calibri>Greetings,<br><br> Below file(s) are present in <b>Bridger Incoming "+$Tag+" </b> Archive path for today in the NAS drive location (" + $inputpath +"\Archive\"+ $DateStr +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>Creation Time</th></tr>" + $L_Msg +"</table><br>This is an informative email.<br><br><br>Regards,<br>Application Support Team" 
              $subject = $environemnt+": NAS Bridger Incoming "+$Tag+" files Archive analysis: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
            }
    }
	
	##Case 12 9CV8 file monitor ~6:20 and 8 PM ET
    "$environemnt\Integrations\Incoming\Cash\LB_Wires\9CV8"
	{
    #Region 12
		#InputPath Declare
		$inputpath1 = "$environemnt\Integrations\Incoming\Cash\LB_Wires"
		$ArchivePath = "$environemnt\Integrations\Incoming\Cash\LB_Wires\Archive"
		$HoldFolder = "$environemnt\Integrations\Incoming\Cash\LB_Wires\HOLD"
		$FailureFolder = "$environemnt\Integrations\Incoming\Cash\LB_Wires\Failure"
		$DateStr = (Get-Date).AddDays(0).ToString('yyMMdd')
		$FilePattern = "Wire_AutoCash_9CV8_*"
		$FilePattern1 = 'Wire_AutoCash_9CV8_'+ $DateStr +'*.csv'
		$checkForFiles = $FolderPath+"\"+$inputpath1+"\"+$FilePattern
		$checkForFiles1 = $FolderPath+"\"+$ArchivePath+"\"+$FilePattern1
		$checkForFiles2 = $FolderPath+"\"+$FailureFolder+"\"+$FilePattern1
		$X=(Get-ChildItem -File $checkForFiles | Measure-Object).Count
		$X2=(Get-ChildItem -File $checkForFiles1 | Measure-Object).Count
		
		#check if files are exist
		$fileExistence = test-path $checkForFiles -PathType Leaf -ErrorAction Stop
		$fileExistence1 = test-path $checkForFiles1 -PathType Leaf -ErrorAction Stop
		$fileExistence2 = test-path $checkForFiles2 -PathType Leaf -ErrorAction Stop
		
		#Check for files in Failure folder, if found send email and also check if any files present then move that in HOLD and ask to check
		
		if ($fileExistence2 -eq $true)
		{
				$L_Msg = ''
				$Files = Get-ChildItem $FolderPath"\"$FailureFolder -File | where {$_.name -like $FilePattern} -ErrorAction Stop
				foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $filesize = (Get-Item $FolderPath"\"$FailureFolder"\"$filename).length/1024
                    $filesize = ([Math]::Round($filesize, 2))
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filesize + '</td><td>' + $filetime + '</td></tr>'
                }
                $body = "<font face = calibri>Greetings,<br><br><font color = Red>We have found below 9CV8 file in a Failure folder location (" + $FailureFolder +") hence, if any 9CV8 file found in parent folder will be moved to HOLD folder:</font><br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>File Size (in Kb)</th><th>Creation Time</th></tr>" + $L_Msg +" </table><br>Please check Failure/Hold folder immediately and take necessary steps. <br><br>Regards,<br>Application Support Team"
				$subject = $environemnt+": 9CV8 file found in the Failure folder for the day: "+$SubjectDate
				Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml -Priority High
				Move-Item -Path $checkForFiles -Destination $FolderPath"\"$HoldFolder"\" -Force
		}
		else
		{
			Write-Host ("No files found in Failure folder")
		}

		#Check parent and Archive location and if files are found on both location then move all parent folders 9CV8 files to Hold folder and send email
		if ($fileExistence1 -eq $true -and $fileExistence -eq $true )
		{
				$L_Msg = ''
				$Files = Get-ChildItem $FolderPath"\"$inputpath1 -File | where {$_.name -like $FilePattern} -ErrorAction Stop
				foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $filesize = (Get-Item $FolderPath"\"$inputpath1"\"$filename).length/1024
                    $filesize = ([Math]::Round($filesize, 2))
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filesize + '</td><td>' + $filetime + '</td></tr>'
                }
                $body = "<font face = calibri>Greetings,<br><br><font color = Red>We have found 9CV8 file in Archive location (" + $ArchivePath +") as well as LB_Wires location, hence we are moving all the below list of files from parent folder to Hold Folder (" + $HoldFolder +") :</font><br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>File Size (in Kb)</th><th>Creation Time</th></tr>" + $L_Msg +" </table><br>Please check it immediately and find out, why we have received an extra file, if one is already received and processed. <br><br>Regards,<br>Application Support Team"
				$subject = $environemnt+": 9CV8 file found in LB_Wires Parent and Archive location for the day: "+$SubjectDate
				Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml -Priority High
				Move-Item -Path $checkForFiles -Destination $FolderPath"\"$HoldFolder"\" -Force
		}
		#Check if files are greter than one then keep the latest on and move all other 9CV8 files to HoldFolder and send email
		elseif ($X -gt 1)
		{
				$L_Msg = ''
				$Files = Get-ChildItem $FolderPath"\"$inputpath1 -File | where {$_.name -like $FilePattern} -ErrorAction Stop
				foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $filesize = (Get-Item $FolderPath"\"$inputpath1"\"$filename).length/1024
                    $filesize = ([Math]::Round($filesize, 2))
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filesize + '</td><td>' + $filetime + '</td></tr>'
                }
                $body = "<font face = calibri>Greetings,<br><br><font color = Red>Today, we have received more than <b> 1 </b> 9CV8 Files in NAS drive location (" + $inputpath1 +") and we have moved all the files to Hold folder except the latest one for processing :</font><br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>File Size (in Kb)</th><th>Creation Time</th></tr>" + $L_Msg +" </table><br>Please check it Immediately and take required action as per the KEDB Steps. <br><br>Regards,<br>Application Support Team"
				$subject = $environemnt+": Extra 9CV8 file received for the day: "+$SubjectDate
				Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml -Priority High
				$FileNumber = (get-childitem $checkForFiles).count - 1
				get-childitem -path $checkForFiles | sort LastWriteTime -Descending | select -last $FileNumber | Move-Item -destination $FolderPath"\"$HoldFolder"\" -Force
		}
			#Check if there is only one file then send an informative email
        elseif ($X -eq 1)
		{
				$L_Msg = ''
				$Files = Get-ChildItem $FolderPath"\"$inputpath1 -File | where {$_.name -like $FilePattern} -ErrorAction Stop
				foreach ($f in $Files)
                {
                    $filename = $f.Name
                    $filetime = $f.CreationTime
                    $filesize = (Get-Item $FolderPath"\"$inputpath1"\"$filename).length/1024
                    $filesize = ([Math]::Round($filesize, 2))
                    $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filesize + '</td><td>' + $filetime + '</td></tr>'
                }
                $body = "<font face = calibri>Greetings,<br><br> Below is the 9CV8 file details present in NAS drive location (" + $inputpath1 +") :<br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>File Size (in Kb)</th><th>Creation Time</th></tr>" + $L_Msg +" </table><br>This is an informative email.<br><br>Regards,<br>Application Support Team"
				$subject = $environemnt+": 9CV8 file received for the day: "+$SubjectDate
				Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
		}
			# If file is not present on parent/Archive location then, send email saying files are not received.
		elseif ($fileExistence -ne $true -and $fileExistence1 -ne $true)
		{
			$body = "<font face = calibri>Greetings,<br><br><font color = Red>We did not receive today's 9CV8 File, neither in NAS drive parent location (" +  $inputpath1 +") nor in Archive location (" +  $ArchivePath +") :</font><br><br>Please check it Immediately and take action as per the KEDB Steps. <br><br>Regards,<br>Application Support Team"
            $subject = $environemnt+": 9CV8 file did not receive for the day: "+$SubjectDate
            Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml  -Priority High
		}
		else
		{
			Write-Host ("File seems to be processed and present in archive folders")
		}
		
		#Check if more than one 9cv8 files present in Archive, then send email to team to check and update.
		if ($X2 -gt 1)
		{
				$L_Msg = ''
				$Files = Get-ChildItem $FolderPath"\"$ArchivePath -File | where {$_.name -like $FilePattern} -ErrorAction Stop
				foreach ($f in $Files)
				{
                 $filename = $f.Name
                 $filetime = $f.CreationTime
                 $filesize = (Get-Item $FolderPath"\"$ArchivePath"\"$filename).length/1024
                 $filesize = ([Math]::Round($filesize, 2))
                 $L_Msg = $L_Msg +'<tr align = "left"><td>' + $filename + '</td><td>' + $filesize + '</td><td>' + $filetime + '</td></tr>'
				}
             $body = "<font face = calibri>Greetings,<br><br><font color = Red>Today, we have received more than <b> 1 </b> 9CV8 Files in Archive location (" + $ArchivePath +")<br> Below are the file details :</font><br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>File Name</th><th>File Size (in Kb)</th><th>Creation Time</th></tr>" + $L_Msg +" </table><br>Please check it Immediately and take required action as per the KEDB Steps. <br><br>Regards,<br>Application Support Team"
			$subject = $environemnt+": More than 1 9CV8 file seems to be processed and found in Archive: "+$SubjectDate
			Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml -Priority High
         }
		 else
		 {
		 }
	}
}
}
Catch
{
    $subject = "NAS Monitoring Cronacle-NAS connetivity issue: "+$SubjectDate
    $body = "<font face = calibri>Greetings,<br><br>There was a below error in NAS Monitoring while connecting NAS <b>$inputpath</b> from Cronacle.<br><br><font color = Red>$Error[0]</font><br><br>Please retrigger the Job Chain in Cronacle.<br><br><br>Regards,<br>Application Support Team"
    Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to  $emailTo -subject $subject -body $body -bodyashtml
}