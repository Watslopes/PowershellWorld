## Export of "larger" Sql Server Blob to file with GetBytes-Stream.         
# Configuration data         
Param (   
$Server = "SQLInstance",         # SQL Server Instance.            
$Database = "DBName",
$Uid = "Username",
$Pwd = "Password",  
$Dest = "D:\BoxSync_Restructure\Box Sync\COVID19DocumentTEST\",             # Path to export to. 
$SubjectDate = (Get-Date).AddDays(0).ToString('dd-MMM-yy hh:mm'),
$emailFrom = "if-lw-RstrBox-upload@company.com",
$emailTo = @("Watson.Lopes@company.com"),
[string[]]$emailCc = @("Watson.Lopes@company.com", "Mohan.Swamy@company.com", "Alexandra.Cunningham@company.com", "Manish.Dhall@company.com"),
$smtpServer = "e2ksmtp01.e2k.ad.company.com",
$bufferSize = 8192,                         # Stream buffer size in bytes.    
$Logflepath = "LogFile\Sequences.txt",
$Archiveflepath = "LogFile\Sequences_Arch.txt"
)

Try
{
	$ErrorActionPreference = "Stop"

    $SkipSeq = Get-Content -Path $Dest$Logflepath;  

    If ($SkipSeq.Count -eq 0) {$SkipSeq = '''000000000'''}

    # Select-Statement for name & blob            
    # with filter.            
    $Sql = "Select AT.File_Source, FS.Content, AC.EntityNaturalId from Attachments(Nolock) AT
            Inner Join ActivityAttachments(Nolock) AAT on AT.Id = AAT.AttachmentID And AAT.IsActive = 1
            Inner Join TransactionInstances(Nolock) TI on TI.EntityId = AAT.Activityid And TI.EntityName = 'Activity'
            Inner Join Activities(Nolock) AC on AC.Id = AAT.ActivityId And AC.IsActive = 1 AND AC.StatusId = 1
            Inner Join ActivityTypes(Nolock) ACT on ACT.Id = AC.ActivityTypeId
		    Inner Join FileStores(Nolock) FS on FS.GUID = Replace(convert(nvarchar(max), AT.File_Content,0),'GUID:','')
			Where ACT.Name in ('Rebooking','Restructure') And
			(AC.Name like '%COVID%' OR AC.Name like '%CV%19%') 
			And ActivityTypeID in (420,421)
            And AC.OwnerId in (select UserId from UserReportingToes 
            Where reportingtoid = (Select Id from Users(Nolock) Where FullName = 'Mary Beres') 
				And UserId <> (Select Id from Users(Nolock) where LoginName = 'Booking_Funding')
        	    And IsActive = 1)
            And AC.EntityNaturalId not in ($SkipSeq)";           
            
    # Open ADO.NET Connection            
    $con = New-Object Data.SqlClient.SqlConnection;            
    $con.ConnectionString = "Data Source=$Server;" +             
                            "Integrated Security=False;" +            
                            "Initial Catalog=$Database;" +   
                            "User ID=$Uid;" +            
                            "Password=$Pwd";          
    $con.Open();            
            
    # New Command and Reader            
    $cmd = New-Object Data.SqlClient.SqlCommand $Sql, $con; 
    $cmd.CommandTimeout=0;
    $rd = $cmd.ExecuteReader();    

    If(!($rd.HasRows)) {throw "MyException"}
           
    # Create a byte array for the stream.            
    $out = [array]::CreateInstance('Byte', $bufferSize)            
            
    # Looping through records            
    While ($rd.Read())            
    {            
        #Write-Output ("Exporting: {0}" -f $rd.GetString(0));                    
        # New BinaryWriter            
        $fs = New-Object System.IO.FileStream ($Dest + $rd.GetString(0)), Create, Write;            
        $bw = New-Object System.IO.BinaryWriter $fs;            
               
        $start = 0;            
        # Read first byte stream            
        $received = $rd.GetBytes(1, $start, $out, 0, $bufferSize - 1);            
        While ($received -gt 0)            
        {            
           $bw.Write($out, 0, $received);            
           $bw.Flush();            
           $start += $received;            
           # Read next byte stream            
           $received = $rd.GetBytes(1, $start, $out, 0, $bufferSize - 1);            
        }       
        
		$bw.Close();            
    	$fs.Close();
    	$fs.Dispose(); 
    
        #File Movements to Seq# wise folders
        $Srcpath = $Dest+$rd.GetValue(0)
        $Destpath = $Dest+$rd.GetValue(2)
        $exists = Test-Path $Destpath
        If (!$exist) 
        {
            #mkdir $Destpath | out-null 
            New-Item -ItemType Directory -Force -Path $Destpath
            Move $Srcpath $Destpath -Force
        }
        Else
        {
            Move $Srcpath $Destpath -Force
        }   
    }    
    # Log File creation part
        (Get-ChildItem -Path $Dest -Attributes D | Where-Object { $_.Name -notmatch '[A-Z]' }).Name | Out-File $Dest$Logflepath
        (Get-ChildItem -Path $Dest -Attributes D | Where-Object { $_.Name -notmatch '[A-Z]' }).Name | Out-File -Append $Dest$Archiveflepath
        Get-Content $Dest$Archiveflepath | Set-Content $Dest$Logflepath
        (Get-Content $Dest$Logflepath) | foreach {",'" + $_ } | foreach { $_ + "'" } | Out-File $Dest$Logflepath
        $SkipSeq = (Get-Content -Path $Dest$Logflepath )
        $SkipSeq[0] = $SkipSeq[0] -replace ',',''
        Set-Content $Dest$Logflepath $SkipSeq          
}

Catch
{

    If ($error[0].FullyQualifiedErrorID -eq "MyException")
    {
        $body = "<font face = calibri>Greetings,<br><br>There are no Records to Process for this run of <B> Covid Restructure File Extract and Box upload</B> from Cronacle.<br><br>No further action is required. Please monitor next load.<br><br><br>Regards,<br>Application Support Team" 
        $subject = "Covid Restructure File Extract and Box upload Status: "+$SubjectDate
        Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -subject $subject -body $body -bodyashtml
    }
    Else
    {
        $body = "<font face = calibri>Greetings,<br><br>There was a below error in <B> Covid Restructure File Extract and Box upload </B> process from Cronacle.<br><br><font color = Red>$Error[0]</font><br><br>Please check urgently and take the necessary action.<br><br><br>Regards,<br>Application Support Team"
        $subject = "Covid Restructure File Extract and Box upload Failure: "+$SubjectDate
        Send-MailMessage -smtpserver $smtpserver -from $emailFrom -to $emailTo -cc $emailCc -subject $subject -body $body -bodyashtml -Priority High
    }
}

Finally
{
	$ErrorActionPreference = "SilentlyContinue"
	
		$bw.Close();            
    	$fs.Close();
    	$fs.Dispose();
    	$rd.Close();            
    	$cmd.Dispose();            
    	$con.Close(); 	
}	