function Get-CounteryData()
{
  [cmdletbinding()]
param
(
[Parameter(Mandatory)]
$CountryName

)

$brdrcntry=@()

$uri=Invoke-WebRequest -UseBasicParsing "https://restcountries.eu/rest/v2/name/$($CountryName)?fullText=true" -ContentType application/json
$data=ConvertFrom-Json $uri

foreach($border in $data.borders)
{

$brdrcntry+=(Invoke-WebRequest -UseBasicParsing "https://restcountries.eu/rest/v2/alpha/$($border)"|ConvertFrom-Json).name
}

$data=[pscustomobject]@{
Country=$data.name
Region=$data.region
Capital=$data.capital  
PopulationInBn=$data.population/1000000000 
Lattiude=$data.latlng[0]
Longitude=$data.latlng[1]
Demonym=$data.demonym 
AreaInSqKm=$data.area
TimeZone=$data.timezones 
BorderCountry=[string]::join(",",$brdrcntry)
currency=$data.currencies.name
Language=[string]::join(",",$data.languages.name)
Flag=$data.flag
Code=$data.cioc
}
$data
}

Get-CounteryData USA
