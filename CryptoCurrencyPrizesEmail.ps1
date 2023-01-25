$MyCrypto1 = (Invoke-WebRequest -uri https://data.messari.io/api/v1/assets/btc/metrics).Content | ConvertFrom-Json
$MyCrypto2 = (Invoke-WebRequest -uri https://data.messari.io/api/v1/assets/eth/metrics).Content | ConvertFrom-Json

$L_Msg = '<tr align = "left"><td>' + $MyCrypto1.data.name + '</td><td><B><font color = Blue> ' + ([Math]::Round($MyCrypto1.data.market_data.price_usd, 4)) + '</font></B></td></tr>'
$L_Msg = $L_Msg +'<tr align = "left"><td>' + $MyCrypto2.data.name + '</td><td> <B><font color = Blue>' + ([Math]::Round($MyCrypto2.data.market_data.price_usd, 4)) + '</font></B></td></tr>'

$body = "<font face = calibri>Greetings,<br><br>Your Crypto price status today: <br><br><table border = 1 cellspacing = 0 cellpadding = 5 bordercolor = black><tr align = left bgcolor = #6495ED><th>Crypto Name</th><th>Crypto Price (USD)</th></tr>" + $L_Msg +"</table><br><br>Regards,<br>Watson Lopes"  
$subject = "Your Crypto Status : "+(Get-Date).AddDays(0).ToString('dd-MMM-yyyy hh:mm')
Send-MailMessage -smtpserver "e2ksmtp01.e2k.ad.ge.com" -from abc@email.com -to  abc@email.com -subject $subject -body $body -bodyashtml -Priority High
