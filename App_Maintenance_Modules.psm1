function Get-LWCreds($UserId)
{
     Switch($UserId)
     {
         "LWDBUser"         {$SQLUsername = $UserId; $PasswordFile = "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\Password_RW.txt"}
         "LWDBUser_ETL"     {$SQLUsername = $UserId; $PasswordFile = "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\Password_RO.txt"}
         "LWDMETL_USER"     {$SQLUsername = $UserId; $PasswordFile = "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\Password_DM.txt"}
         "Logon\lg688521sv" {$SQLUsername = $UserId; $PasswordFile = "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\Password_FN.txt"}
         "GECatalog-CiLUz1q6Q13kT70L23i0Y4lO" {$SQLUsername = $UserId; $PasswordFile = "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\Password_StgAkana.txt"}         
 	 }

     $KeyFile = "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\LW_AES.key"
     $key = Get-Content $KeyFile
     $securecred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SQLUsername, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
     $SQLPassword = $securecred.GetNetworkCredential().Password
     Return $SQLUsername, $SQLPassword    
}


function Get-DBExtract ($SQLQuery, $SQLServer, $SQLDBName, $AttachmentName, $needhtmlobj)
{
     if(!$AttachmentName){$AttachmentName = 'LW_Extract'}
     if(!$needhtmlobj){$needhtmlobj = 1}
     $AttachmentPath = "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\"+$AttachmentName+"_"+(Get-Date).AddDays(0).ToString('dd-MMM-yy')+".csv"
     # This code connects to the SQL server and retrieves the data     
     $SQLConnection = New-Object System.Data.SqlClient.SqlConnection
     $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; uid = $SQLUsername; pwd = $SQLPassword"
     $SqlConnection.Open()
     $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
     $SqlCmd.CommandText = $SqlQuery
     $SqlCmd.Connection = $SqlConnection

     $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
     $SqlAdapter.SelectCommand = $SqlCmd

     $DataSet = New-Object System.Data.DataSet
     $Records = $SqlAdapter.Fill($DataSet)
	 $Records | Out-Null
	 $SqlConnection.Close()

	if ($Records -gt 0)
	{
	#Populate Hash Table
	    $objTable = $DataSet.Tables[0]
	    #Export Hash Table to CSV File
	    $objTable | Export-CSV -NoTypeInformation $AttachmentPath
    
        if($needhtmlobj -eq 1)
        {
            [int]$i = 0
            $columnhtml = ''
            While ($i -le ($objTable.Columns.Count - 1))
            {
                $columnhtml += "<td>"+ $objTable.Columns[$i].ColumnName + "</td>"
                $i += 1
            }

            $htmltbl = "<table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr bgcolor = 6495ED>$columnhtml</tr>"
	
            foreach ($row in $objTable.Rows[0..4])
            {
            $i = 0
            $rowhtml=""
                While ($i -le ($objTable.Columns.Count - 1))
                {
                    $rowhtml += "<td>" + $row[$i] + "</td>"
                    $i += 1
                }
            $htmltbl += "<tr>$rowhtml</tr>"
            }
            $htmltbl += "</table>"
            Return $Records, $htmltbl, $needhtmlobj
        }
        Else
        {
            Return $Records, $needhtmlobj
        }
   } Return $Records, $objTable
}

function send-email ($emailfrom, $emailto, $emailcc, $AttachmentName, $Emailtag)
{
    #$Records = 0
    if(!$emailfrom){$emailfrom = 'if-Application-maintenance@company.com'}
    if(!$emailto){$emailto = 'capitaltcslwsupport@company.com'}
    if(!$emailcc){$emailcc = 'capitaltcslwsupport@company.com'}
    if(!$Emailtag){$Emailtag = $AttachmentName}
    $smtp = 'e2ksmtp01.e2k.ad.company.com'
    $pos = $sqlserver.IndexOf("_")
    $env = $sqlserver.Substring($pos+1)   
    $AttachmentPath = "D:\LW_DM_ETL_SSIS\LW_Maintenance\Encrypted_Passwords\"+$AttachmentName+"_"+(Get-Date).AddDays(0).ToString('dd-MMM-yy')+".csv"

    If ($Records -eq 0){
        $Subject = "$env : No "+ "$Emailtag : " + (Get-Date).AddDays(0).ToString('dd-MMM-yy')
        $Body = "<font face = calibri>Greetings,<br><br> There are <B><font color=RED>No</B></font> record(s) found for <U> $Emailtag </U>.<br><br>Regards,<br>Application Support Team"
        Send-MailMessage -smtpserver $smtp -from $emailfrom -to $emailto -subject $subject -body $Body -bodyashtml}
    ELse{
        If($needhtmlobj -eq 1){            
            $Subject = "$env : "+ "$Emailtag : " + (Get-Date).AddDays(0).ToString('dd-MMM-yy')
            $Body = "<font face = calibri>Greetings,<br><br> There are <B><font color=RED>"+$Records+"</B></font> record(s) found for <U>$Emailtag</U>. Below are <B>5</B> sample records for your reference.<br><br> $htmltbl <br>PFA detailed report for more information.<br><br>Regards,<br>Application Support Team"
            Send-MailMessage -smtpserver $smtp -from $emailfrom -to $emailto -cc $emailcc -subject $subject -body $Body -bodyashtml -Attachment $AttachmentPath}
        Else{            
            $Subject = "$env : "+ "$Emailtag : " + (Get-Date).AddDays(0).ToString('dd-MMM-yy')
            $Body = "<font face = calibri>Greetings,<br><br> There are <B><font color=RED>"+$Records+"</B></font> record(s) found for <U>$Emailtag</U>. PFA detailed report for more information.<br><br>Regards,<br>Application Support Team"
            Send-MailMessage -smtpserver $smtp -from $emailfrom -to $emailto -cc $emailcc -subject $subject -body $Body -bodyashtml -Attachment $AttachmentPath}        
        }        
}

function send-erroremail ($emailerror, $emailtag)
{
    $pos = $sqlserver.IndexOf("_")
    $env = $sqlserver.Substring($pos+1)
    if(!$Emailtag){$Emailtag = $AttachmentName}
    $subject = "$env : "+ "$Emailtag connectivity issue while extracting: "+ (Get-Date).AddDays(0).ToString('dd-MMM-yy')
    $body = "<font face = calibri>Greetings,<br><br>Below is the error message while executing <U> $Emailtag </U> process -<br><font color = Red>$emailerror</font><br><br>Please check and retrigger the Job Job/Chain in Cronacle if required.<br><br><br>Regards,<br>Application Support Team"
    Send-MailMessage -smtpserver 'e2ksmtp01.e2k.ad.company.com' -from 'if-Application-maintenance@company.com' -to "capitaltcslwsupport@company.com" -subject $subject -body $body -bodyashtml -Priority High
}

function Get-AkanaToken ($strClientID, $strAuthURL)
{
	$strClientID, $strClientSecret = Get-LWCreds $strClientID
	 
	#GetAkanaTokenStart
	$params = @{
		'client_id' = $strClientID
		'client_secret' = $strClientSecret
		'grant_type' = 'client_credentials'
		'scope' = 'api'
	}            
	$getAuthToken = Invoke-RestMethod -Method Post -Uri $strAuthURL -Body $params -ContentType "application/x-www-form-urlencoded"	 
	$getAuthToken = $getAuthToken.access_token.ToString()
	#GetAkanaTokenEnd	 
	
	Return $getAuthToken    
}

function Get-BoxToken ($strBoxAuthURL, $getAuthToken)
{	 
	#GetBoxTokenStart          
	$getBoxToken = Invoke-RestMethod -Method Get -Uri $strBoxAuthURL -Headers @{Authorization = 'Bearer ' + $getAuthToken}	 
	$getBoxToken = $getBoxToken.accesstoken.ToString()
	#GetBoxTokenEnd	 
	
	Return $getBoxToken    
}

function CreateBoxFolder ($strBoxFolderName, $strParentBoxFolderId, $strZscalarProxy, $strBoxFolderURL, $getBoxToken)
{
	$param = @{
		'name' = $strBoxFolderName
		'parent' = @{'id'=$strParentBoxFolderId}
	} | ConvertTo-Json
	$NewBoxFolderInfo = Invoke-RestMethod -Proxy $strZscalarProxy -Method Post -Uri $strBoxFolderURL -Body $param -Headers @{Authorization = 'Bearer ' + $getBoxToken} -ContentType "application/json"
	
	Return $NewBoxFolderInfo 
}

function UploadFileToBox ($LWFileFullName, $LWFileName, $BoxFolderId, $strZscalarProxy, $strBoxFileURL, $getBoxToken)
{
	$boundary = [guid]::NewGuid().ToString()
	$filebody = [System.IO.File]::ReadAllBytes($LWFileFullName)
	$enc = [System.Text.Encoding]::GetEncoding('utf-8')
	$filebodytemplate = $enc.GetString($filebody)
	
	[System.Text.StringBuilder]$contents = New-Object System.Text.StringBuilder
	[void]$contents.AppendLine()
	[void]$contents.AppendLine("--$boundary")
	[void]$contents.AppendLine("Content-Disposition: form-data; name=""file""; filename=""$LWFileName""")
	[void]$contents.AppendLine()
	[void]$contents.AppendLine($filebodytemplate)
	[void]$contents.AppendLine("--$boundary")
	[void]$contents.AppendLine("Content-Type: text/plain; charset=utf-8")
	[void]$contents.AppendLine("Content-Disposition: form-data; name=""metadata""")
	[void]$contents.AppendLine()
	[void]$contents.AppendLine("{""name"":""$LWFileName"",""parent"":{""id"":""$BoxFolderId""}}")
	[void]$contents.AppendLine()
	[void]$contents.AppendLine("--$boundary--")
	$body = $contents.ToString()
	
	$strUploadedFileInfo = Invoke-RestMethod -Proxy $strZscalarProxy -Method Post -Uri $strBoxFileURL -Body $body -Headers @{Authorization = 'Bearer ' + $getBoxToken} -ContentType "multipart/form-data;boundary=$boundary" -Verbose       

	Return $strUploadedFileInfo 
}

 function Draw-Chart ($chartname, $charttype, $chartcolor, $DBTable, $x, $y){

     $chartarea.AxisX.Title = $X
     $chartarea.AxisY.Title = $Y
     

     $datasource = $DBTable
     [void]$mychart.Series.Add($chartname)
     $mychart.Series[$chartname].ChartType = $charttype
     $mychart.Series[$chartname].IsVisibleInLegend = $true
     $mychart.Series[$chartname].BorderWidth  = 2
     #$mychart.Series[$chartname].chartarea = "ChartArea1"
     $mychart.Series[$chartname].Legend = "Legend1"
     $mychart.Series[$chartname].color = $chartcolor
     $datasource | ForEach-Object  {$mychart.Series[$chartname].Points.addxy($_.$x, $_.$y) }
 }

 function Execute-SQLQuery ($SQLQuery, $needDBTable){

     $SQLConnection = New-Object System.Data.SqlClient.SqlConnection
     $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; uid = $SQLUsername; pwd = $SQLPassword"
     $SqlConnection.Open()
     $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
     $SqlCmd.CommandText = $SqlQuery
     $SqlCmd.Connection = $SqlConnection

     $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
     $SqlAdapter.SelectCommand = $SqlCmd

     $DataSet = New-Object System.Data.DataSet
     $Records = $SqlAdapter.Fill($DataSet)
     $Records | Out-Null
     $SqlConnection.Close()
     If($needDBTable -eq 1){ 
        Return $DataSet.Tables[0]
     }
 }