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
$FolderPath="\\NAS Drive"
$strOutputFilePath="$scriptFilePath\OutputLog_" + (Get-Date).ToString('yyyyMMdd') + ".txt"

#### API Variables ####
$strZscalarProxy="http://PITC-Zscaler-Americas-Alpharetta3PR.proxy.corporate.company.com:80"
$strAuthURL="https://fssfed.stage.company.com/fss/as/token.oauth2"
$strClientID="CiLUz1q6Q13kT70L23i0Y4lO"
$strClientSecret="c86e1eb700ad2ea40f525b200fba3a9570016071"
$strAuthUserUrl="https://stage.api.company.com/digital/boxapi/v1/authurl?email=502794554@mail.ad.company.com"
$strBoxAuthURL="https://stage.api.company.com/digital/boxapi/v1/token?email=502794554@mail.ad.company.com"
$strBoxFolderURL="https://api.box.com/2.0/folders"
$strBoxFileURL="https://upload.box.com/api/2.0/files/content"

#### Email Notification Variables ####
$mailSubjDate=(Get-Date).ToString('dd-MMM-yyyy')
$mailsubject=$environment + " - Application $strLWFileName File Upload Status : " + $mailSubjDate
$mailFrom="if-Application-" + $environment.ToLower() + "@company.com" 
[string[]]$mailTo=$args[8].Split(',')
[string[]]$mailCc=$args[9].Split(',')
$mailsmtpSrv="e2ksmtp01.e2k.ad.company.com"
$mailbody = "<html><head><style>
			table.gridtable {font-family:Calibri;font-size:15px;border-width:1px;border-color:black;border-collapse:collapse;} 
			table.gridtable th {border-width:1px;border-style:solid;border-color:black;font-weight:bold;background-color:#305588;color:#e9fbfa;} 
			table.gridtable td {border-width:1px;border-style:solid;border-color:black;} 	
			body {font-family:Calibri;font-size:15px;} 
			</style></head>
			<body>Greetings,<br/><br/>Application $strLWFileName File Upload Status :</br></br>"

Try
{
    $ErrorActionPreference = 'Stop'
	Remove-Item "$scriptFilePath\*.txt"
	
	(Get-Date).ToString() + " - Application File Upload Automation Started `n" | Out-File $strOutputFilePath

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
			$params = @{
				'client_id' = $strClientID
				'client_secret' = $strClientSecret
				'grant_type' = 'client_credentials'
				'scope' = 'api'
			}            
			$getAuthToken = Invoke-RestMethod -Method Post -Uri $strAuthURL -Body $params -ContentType "application/x-www-form-urlencoded"
			(Get-Date).ToString() + " - OAuth 2.0 Token Generated - " + $getAuthToken | Out-File $strOutputFilePath -Append
			$getAuthToken = $getAuthToken.access_token.ToString()    
			(Get-Date).ToString() + " - Get OAuth 2.0 Token Completed `n" | Out-File $strOutputFilePath -Append
			#GetAkanaTokenEnd

			#GetBoxTokenStart
			(Get-Date).ToString() + " - Get Box Access Token Started" | Out-File $strOutputFilePath -Append
			$getBoxToken = Invoke-RestMethod -Method Get -Uri $strBoxAuthURL -Headers @{Authorization = 'Bearer ' + $getAuthToken}
			(Get-Date).ToString() + " - Box Access Token Generated - " + $getBoxToken | Out-File $strOutputFilePath -Append
			$getBoxToken = $getBoxToken.accesstoken.ToString()
			(Get-Date).ToString() + " - Get Box Access Token Completed `n" | Out-File $strOutputFilePath -Append
			#GetBoxTokenEnd
	
			#GetBoxFolderInfoStart
			(Get-Date).ToString() + " - Get Box Folder Information Started" | Out-File $strOutputFilePath -Append
			$getBoxFolderInfo = Invoke-RestMethod -Proxy $strZscalarProxy -Method Get -Uri ($strBoxFolderURL+"/"+$strParentBoxFolderId) -Headers @{Authorization = 'Bearer ' + $getBoxToken}
			$strFolderInfo = $getBoxFolderInfo.item_collection.entries | Where-Object {$_.name -eq $strBoxFolderName}
			if($strFolderInfo.id.Length -gt 0){
				(Get-Date).ToString() + " - Folder with name '" + $strBoxFolderName + "' is already exist. FolderId = " + $strFolderInfo.id + " and FolderName = " + $strFolderInfo.name | Out-File $strOutputFilePath -Append
			}
			else{
				(Get-Date).ToString() + " - Creating new folder with name '" + $strBoxFolderName + "'."| Out-File $strOutputFilePath -Append
				#CreateBoxFolderStart
				(Get-Date).ToString() + " - Create Box Folder Started" | Out-File $strOutputFilePath -Append
				$param2 = @{
							'name' = $strBoxFolderName
							'parent' = @{'id'=$strParentBoxFolderId}
						   } | ConvertTo-Json
				$strFolderInfo = Invoke-RestMethod -Proxy $strZscalarProxy -Method Post -Uri $strBoxFolderURL -Body $param2 -Headers @{Authorization = 'Bearer ' + $getBoxToken} -ContentType "application/json"
				(Get-Date).ToString() + " - New Box folder is created. FolderId = " + $strFolderInfo.id + " | FolderName = " + $strFolderInfo.name | Out-File $strOutputFilePath -Append
				(Get-Date).ToString() + " - Create Box Folder Completed" | Out-File $strOutputFilePath -Append
				#CreateBoxFolderEnd
			}	
			(Get-Date).ToString() + " - Get Box Folder Information Completed `n" | Out-File $strOutputFilePath -Append
			#GetBoxFolderInfoEnd
	
			#BoxFileUploadStart
			(Get-Date).ToString() + " - Box File Upload Started" | Out-File $strOutputFilePath -Append
		
			$L_Msg += "<table class='gridtable'><tr><th>Environment</th><th>FileName</th><th>FileSize (kb)</th><th>Status</th></tr>"
			for ($i=0; $i -lt $Files.count; $i++) {
		
				$boundary = [guid]::NewGuid().ToString()
				$filebody = [System.IO.File]::ReadAllBytes($files[$i].FullName)
				$enc = [System.Text.Encoding]::GetEncoding('utf-8')
				$filebodytemplate = $enc.GetString($filebody)
	
				[System.Text.StringBuilder]$contents = New-Object System.Text.StringBuilder
				[void]$contents.AppendLine()
				[void]$contents.AppendLine("--$boundary")
				[void]$contents.AppendLine("Content-Disposition: form-data; name=""file""; filename=""$($files[$i].Name)""")
				[void]$contents.AppendLine()
				[void]$contents.AppendLine($filebodytemplate)
				[void]$contents.AppendLine("--$boundary")
				[void]$contents.AppendLine("Content-Type: text/plain; charset=utf-8")
				[void]$contents.AppendLine("Content-Disposition: form-data; name=""metadata""")
				[void]$contents.AppendLine()
				[void]$contents.AppendLine("{""name"":""$($files[$i].Name)"",""parent"":{""id"":""$($strFolderInfo.id)""}}")
				[void]$contents.AppendLine()
				[void]$contents.AppendLine("--$boundary--")
				$body = $contents.ToString()
	
				$strFileInfo = Invoke-RestMethod -Proxy $strZscalarProxy -Method Post -Uri $strBoxFileURL -Body $body -Headers @{Authorization = 'Bearer ' + $getBoxToken} -ContentType "multipart/form-data;boundary=$boundary" -Verbose       
	
				(Get-Date).ToString() + " - File - " + $strFileInfo.entries[0].name + " with Id - " + $strFileInfo.entries[0].id + " is successfully uploaded to '" + $strFolderInfo.Name + "' Box folder" | Out-File $strOutputFilePath -Append
				$filesize = ([Math]::Round($files[$i].Length/1024, 2))
				$L_Msg += '<tr><td>' + $environment + '</td><td>' + $files[$i].Name + '</td><td>' + $filesize + '</td><td style="color:green;">Success</td></tr>'
			}
		
			(Get-Date).ToString() + " - Box File Upload Completed `n" | Out-File $strOutputFilePath -Append
			#BoxFileUploadEnd
		
			(Get-Date).ToString() + " - Application File Upload Automation Completed" | Out-File $strOutputFilePath -Append
		
			$mailbody += $L_Msg + "</table><br/>Regards,<br/>Application Support Team</body>"
			Send-MailMessage -smtpserver $mailsmtpSrv -from $mailFrom -to $mailTo -Cc $mailCc -subject $mailsubject -Attachments $strOutputFilePath -body $mailbody -bodyashtml
		}
		else{	
			Write-Host "Job is unable to find file(s) with pattern '$FilePattern' at path - $inputpath\$strNASFolderName"
			(Get-Date).ToString() + " - Job is unable to find file(s) with pattern '"+ $FilePattern +"' at path - " + $inputpath + "\" + $strNASFolderName + "`n" | Out-File $strOutputFilePath -Append
			(Get-Date).ToString() + " - Application File Upload Automation Completed" | Out-File $strOutputFilePath -Append
		}
	}
	else{
		 Write-Host "Job is unable to find folder with name $strNASFolderName at path - $inputpath"
		(Get-Date).ToString() + " - Job is unable to find folder with name '"+ $strNASFolderName +"' at path - " + $inputpath + "`n" | Out-File $strOutputFilePath -Append
		(Get-Date).ToString() + " - Application File Upload Automation Completed" | Out-File $strOutputFilePath -Append
	}
} #End Try
Catch
{
    "Exception Occurred - " + $Error[0] | Out-File $strOutputFilePath -Append
	"Exception Message  - " + $_.Exception.Message | Out-File $strOutputFilePath -Append	
	$mailbody += "An error occurred in script execution. Please review attached log and take necessary actions.<font color='red'></br>Error : $_</font><br/><br/>Regards,<br/>Application Support Team</body>"    
    Send-MailMessage -smtpserver $mailsmtpSrv -from $mailFrom -to $mailTo -Cc $mailCc -subject $mailsubject -Attachments $strOutputFilePath -body $mailbody -bodyashtml
	Break
}