#Data variable to hold Json data from https://pomber.github.io/covid19/timeseries.json
# Thanks to Rodrigo Pombo for JSON data
$data=Invoke-WebRequest -UseBasicParsing https://pomber.github.io/covid19/timeseries.json


# Converting inputing Json to a variable called India
$india=ConvertFrom-Json -InputObject $data 

# This link has data upto previous day, so prevday variable will do that conversion
$prevday=(date).AddDays(-1).ToString("yyyy-M-dd")

# I am only looking for India specific data, you can choose your country
$india.US|Where-Object{$_.date -eq $prevday}

Write-Output "Above Data is for India, Please change for your country"