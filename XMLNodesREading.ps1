[xml]$XmlDocument = Get-Content -Path 'c:\Test.xml'
$node = $XmlDocument.SelectNodes("//File/FundsTransfer")
$result = @()

foreach ($row1 in $node)
{    
    If ($row1.HasChildNodes)
    {
    $row2 = $row1.LastChild

        If ($row2.HasChildNodes)
        {            
              If($row2.LastChild.Type -eq '2250')
              {                    
                    <#$result = $row2.LastChild.Type + "|" + $row2.LastChild.Description.Substring($row2.LastChild.Description.Length-22,22) + "|" + $row1.FirstChild.FSource
                    $result#>
                    
                    $details = @{ Type         = $row2.LastChild.Type             
                                  FedRefNum    = $row2.LastChild.Description.Substring($row2.LastChild.Description.Length-22,22)                 
                                  LWVourcherId = $row1.FirstChild.FSource }                           
                    $result += New-Object PSObject -Property $details                  
              }

        }

    }
}
$result
#$result | export-csv -Path 'C:\Users\502740204\Desktop\XML_Ron\Output.csv'-NoTypeInformation

<#
#post title and content 
$params = @{ 
    title = "test Rest API post" 
    content = "test Rest API post content" 
    status = 'publish' 
} 
#change username and password before use 
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("user:pass@23"))) 
$header = @{ 
Authorization=("Basic {0}" -f $base64AuthInfo) 
} 
$params1=$params|ConvertTo-Json 
Invoke-RestMethod -Method post -Uri http://khaoodara.com/wp-json/wp/v2/posts -ContentType "application/json" -Body $params1  -Headers $header -UseBasicParsing 
 
# for deleting post you can use rest DELETE method and just add post id at the end of URI.. 
#Invoke-RestMethod -Method delete -Uri http://khaoodara.com/wp-json/wp/v2/posts/6705 -ContentType "application/json" -headers $header 
 
#for get post data you can use rest GET method .. 
#Invoke-RestMethod -Method get -Uri http://khaoodara.com/wp-json/wp/v2/posts -ContentType "application/json" -headers $header 
#>