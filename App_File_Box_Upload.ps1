#### POSH Variables ####
$environment=$args[0]
$strLWFileName=$args[1]
$inputpath=$args[2]
$scriptFilePath=$args[3]
$strNASFolderName=$args[4]
$FilePattern=$args[5]
$strParentBoxFolderId=$args[6]
$strBoxFolderName=$args[7]

#### Path Variables ####
$strOutputFilePath="$scriptFilePath\OutputLog_" + (Get-Date).ToString('yyyyMMdd') + ".txt"

#### API Variables ####
$strBoxFolderURL="https://api.box.com/2.0/folders"
$strBoxFileURL="https://upload.box.com/api/2.0/files/content"

#### Email Notification Variables ####
$mailSubjDate=(Get-Date).ToString('dd-MMM-yyyy')
$mailsubject=$environment + " - application $strLWFileName File Upload Status : " + $mailSubjDate
$mailFrom="if-application-" + $environment.ToLower() + "@company.com" 
[string[]]$mailTo=$args[8].Split(',')
[string[]]$mailCc=$args[9].Split(',')
$mailsmtpSrv="e2ksmtp01.e2k.ad.company.com"
$mailbody = "<html><head><style>
			table.gridtable {font-family:Calibri;font-size:15px;border-width:1px;border-color:black;border-collapse:collapse;} 
			table.gridtable th {border-width:1px;border-style:solid;border-color:black;font-weight:bold;background-color:#305588;color:#e9fbfa;} 
			table.gridtable td {border-width:1px;border-style:solid;border-color:black;} 	
			body {font-family:Calibri;font-size:15px;} 
			</style></head>
			<body>Greetings,<br/><br/>application $strLWFileName File Upload Status :</br></br>"

Try
{

	Switch ($environment){
		
		"UAT"{
			$FolderPath="\\NAS Folder"
			$strZscalarProxy="ProxyURL"
			$strAuthURL="AuthURL"
			$strClientID="dsvjksdnvdsnfs"
			$strBoxAuthURL="https://stage.api.company.com/digital/boxapi/v1/token?email=abc@mail.ad.company.com"		
		}
		
		"PROD"{
			$FolderPath="\\NAS Folder"
			$strZscalarProxy="ProxyURL"
			$strAuthURL="AuthURL"
			$strClientID="dsvjksdnvdsnfs"
			$strBoxAuthURL="https://stage.api.company.com/digital/boxapi/v1/token?email=abc@mail.ad.company.com"	
		}		
	}
	
    $ErrorActionPreference = 'Stop'
	Remove-Item "$scriptFilePath\*.txt"
	
	(Get-Date).ToString() + " - application File Upload Automation Started `n" | Out-File $strOutputFilePath

	#GetFolderInfoStart
	$Folder = Get-ChildItem $FolderPath"\"$inputpath | Where {$_.Name -eq $strNASFolderName}
	#GetFolderInfoEnd

	If($Folder.count -gt 0)
	{
		#GetFileInfoStart
		$Files = Get-ChildItem $FolderPath"\"$inputpath"\"$strNASFolderName -File | where {$_.name -like $FilePattern}
		#GetFileInfoEnd

		If ($Files.count -gt 0)
		{    
		   #GetAkanaTokenStart
		    (Get-Date).ToString() + " - Get OAuth 2.0 Token Started" | Out-File $strOutputFilePath -Append
			$getAuthToken = Get-AkanaToken $strClientID $strAuthURL
			If ($getAuthToken.Length -gt 0) {(Get-Date).ToString() + " - OAuth 2.0 Token Generated" | Out-File $strOutputFilePath -Append}
			(Get-Date).ToString() + " - Get OAuth 2.0 Token Completed `n" | Out-File $strOutputFilePath -Append
			#GetAkanaTokenEnd

			#GetBoxTokenStart
			(Get-Date).ToString() + " - Get Box Access Token Started" | Out-File $strOutputFilePath -Append
			$getBoxToken = Get-BoxToken $strBoxAuthURL $getAuthToken
			If ($getBoxToken.Length -gt 0) {(Get-Date).ToString() + " - Box Access Token Generated" | Out-File $strOutputFilePath -Append}
			(Get-Date).ToString() + " - Get Box Access Token Completed `n" | Out-File $strOutputFilePath -Append
			#GetBoxTokenEnd
	
			#GetBoxFolderInfoStart
			(Get-Date).ToString() + " - Get Box Folder Information Started" | Out-File $strOutputFilePath -Append
			$getBoxFolderInfo = Invoke-RestMethod -Proxy $strZscalarProxy -Method Get -Uri ($strBoxFolderURL+"/"+$strParentBoxFolderId) -Headers @{Authorization = 'Bearer ' + $getBoxToken}
			$NewBoxFolderInfo = $getBoxFolderInfo.item_collection.entries | Where-Object {$_.name -eq $strBoxFolderName}
			if($NewBoxFolderInfo.id.Length -gt 0){
				(Get-Date).ToString() + " - Folder with name '" + $strBoxFolderName + "' is already exist. FolderId = " + $NewBoxFolderInfo.id + " and FolderName = " + $NewBoxFolderInfo.name | Out-File $strOutputFilePath -Append
			}
			else{
				(Get-Date).ToString() + " - Creating new folder with name '" + $strBoxFolderName + "'."| Out-File $strOutputFilePath -Append
				#CreateBoxFolderStart
				(Get-Date).ToString() + " - Create Box Folder Started" | Out-File $strOutputFilePath -Append
				$NewBoxFolderInfo = CreateBoxFolder $strBoxFolderName $strParentBoxFolderId $strZscalarProxy $strBoxFolderURL $getBoxToken
				(Get-Date).ToString() + " - New Box folder is created. FolderId = " + $NewBoxFolderInfo.id + " | FolderName = " + $NewBoxFolderInfo.name | Out-File $strOutputFilePath -Append
				(Get-Date).ToString() + " - Create Box Folder Completed" | Out-File $strOutputFilePath -Append
				#CreateBoxFolderEnd
			}	
			(Get-Date).ToString() + " - Get Box Folder Information Completed `n" | Out-File $strOutputFilePath -Append
			#GetBoxFolderInfoEnd
	
			#BoxFileUploadStart
			(Get-Date).ToString() + " - Box File Upload Started" | Out-File $strOutputFilePath -Append
		
			$L_Msg += "<table class='gridtable'><tr><th>Environment</th><th>FileName</th><th>FileSize (kb)</th><th>Status</th></tr>"
			for ($i=0; $i -lt $Files.count; $i++) {
				#Call UploadFileToBox Function				
				$strUploadedFileInfo = UploadFileToBox $($files[$i].FullName) $($files[$i].Name) $($NewBoxFolderInfo.id) $strZscalarProxy $strBoxFileURL $getBoxToken
				(Get-Date).ToString() + " - File - " + $strUploadedFileInfo.entries[0].name + " with Id - " + $strUploadedFileInfo.entries[0].id + " is successfully uploaded to '" + $NewBoxFolderInfo.Name + "' Box folder" | Out-File $strOutputFilePath -Append
				$filesize = ([Math]::Round($files[$i].Length/1024, 2))
				$L_Msg += '<tr><td>' + $environment + '</td><td>' + $files[$i].Name + '</td><td>' + $filesize + '</td><td style="color:green;">Success</td></tr>'
			}
		
			(Get-Date).ToString() + " - Box File Upload Completed `n" | Out-File $strOutputFilePath -Append
			#BoxFileUploadEnd
		
			(Get-Date).ToString() + " - application File Upload Automation Completed" | Out-File $strOutputFilePath -Append
		
			$mailbody += $L_Msg + "</table><br/>Regards,<br/>application Support Team</body>"
			Send-MailMessage -smtpserver $mailsmtpSrv -from $mailFrom -to $mailTo -Cc $mailCc -subject $mailsubject -Attachments $strOutputFilePath -body $mailbody -bodyashtml
		}
		else{	
			Write-Host "Job is unable to find file(s) with pattern '$FilePattern' at path - $inputpath\$strNASFolderName"
			(Get-Date).ToString() + " - Job is unable to find file(s) with pattern '"+ $FilePattern +"' at path - " + $inputpath + "\" + $strNASFolderName + "`n" | Out-File $strOutputFilePath -Append
			(Get-Date).ToString() + " - application File Upload Automation Completed" | Out-File $strOutputFilePath -Append
		}
	}
	else{
		 Write-Host "Job is unable to find folder with name $strNASFolderName at path - $inputpath"
		(Get-Date).ToString() + " - Job is unable to find folder with name '"+ $strNASFolderName +"' at path - " + $inputpath + "`n" | Out-File $strOutputFilePath -Append
		(Get-Date).ToString() + " - application File Upload Automation Completed" | Out-File $strOutputFilePath -Append
	}
} #End Try
Catch
{
    "Exception Occurred - " + $Error[0] | Out-File $strOutputFilePath -Append
	"Exception Message  - " + $_.Exception.Message | Out-File $strOutputFilePath -Append	
	$mailbody += "An error occurred in script execution. Please review attached log and take necessary actions.<font color='red'></br>Error : $_</font><br/><br/>Regards,<br/>application Support Team</body>"    
    Send-MailMessage -smtpserver $mailsmtpSrv -from $mailFrom -to $mailTo -Cc $mailCc -subject $mailsubject -Attachments $strOutputFilePath -body $mailbody -bodyashtml
	Break
}