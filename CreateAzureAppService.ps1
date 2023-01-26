$location="EastUS"
# transform userid to lowercase since some Azure resource names don't like uppercase
$userid=$env:USERNAME.tolower()
$FuncAppName="$userid$(Get-Random)"
$rgname="$FuncAppName-rg"
$storageAccount="$($FuncAppName)stg"
$FunctionName="HttpTriggerCSharp3"

# ---------------------------------------------------------------------------------
# create the resource group
# ---------------------------------------------------------------------------------
New-AzureRmResourceGroup -Name "$rgname" -Location "$location" -force

# ---------------------------------------------------------------------------------
# create a storage account needed for the Function App
# ---------------------------------------------------------------------------------
New-AzureRmStorageAccount -ResourceGroupName "$rgname" -AccountName "$storageAccount" -Location "$location" -SkuName "Standard_LRS"
$keys = Get-AzureRmStorageAccountKey -ResourceGroupName "$rgname" -AccountName "$storageAccount"
$storageAccountConnectionString = 'DefaultEndpointsProtocol=https;AccountName=' + $storageAccount + ';AccountKey=' + $keys[0].Value

# ---------------------------------------------------------------------------------
# create the Function App
# ---------------------------------------------------------------------------------
New-AzureRmResource -ResourceGroupName "$rgname" -ResourceType "Microsoft.Web/Sites" -ResourceName "$FuncAppName" -kind "functionapp" -Location "$location" -Properties @{} -force

$AppSettings = @{'AzureWebJobsDashboard' = $storageAccountConnectionString;
    'AzureWebJobsStorage' = $storageAccountConnectionString;
    'FUNCTIONS_EXTENSION_VERSION' = '~1';
    'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING' = $storageAccountConnectionString;
    'WEBSITE_CONTENTSHARE' = $storageAccount;
}
Set-AzureRMWebApp -Name "$FuncAppName" -ResourceGroupName "$rgname" -AppSettings $AppSettings
