
$appEnvironment=$args[0]
$interfaceType=$args[1]
$appInterface=$args[2]

#$Basepath="D:\LW_DM_ETL_SSIS\LW_Maintenance\LW_Realtime_Interfaces"
$Basepath="\\NASDrive"
$strZscalarProxy="http://Proxy_URL"

	$ErrorActionPreference = 'Stop'
	Write-Host "`nInvoking $appInterface request via $interfaceType endpoint for $appEnvironment environment. Please find below Response status.`n"

Try
{
	Switch ($appEnvironment + "_" + $interfaceType + "_" + $appInterface){
		
		"SIT_Middleware_Bridger"{$endpointURL="http://tst-is.corporate.company.com/ERP/IFLWOSBIntegrationProject/Services/ProxyService/PSLWReq"}
		"SIT_Middleware_CARRA"{$endpointURL="http://tst-is.corporate.company.com/ERP/IFLWOSBIntegrationProject/Services/ProxyService/PSLWReq"}
		"SIT_Middleware_CSC-LienFiling"{$endpointURL="http://tst-is.corporate.company.com/ERP/IFCSC_DiligenzIntegration/Services/ProxyService/PSReadLWCSCReq"}
		"SIT_Middleware_Vertex-LookUpTax"{$endpointURL="http://tst-is.corporate.company.com/IFLWVertexIntegrationProject/Services/Proxy/PSLWVertexLookUpTax"}
		"SIT_Middleware_Vertex-CalculateTax"{$endpointURL="http://tst-is.corporate.company.com/IFLWVertexIntegrationProject/Services/Proxy/PSLWVertexCalculateTax"}
		
		"UAT_Middleware_Bridger"{$endpointURL="http://stg-is.corporate.company.com/ERP/IFLWOSBIntegrationProject/Services/ProxyService/PSLWReq"}
		"UAT_Middleware_CARRA"{$endpointURL="http://stg-is.corporate.company.com/ERP/IFLWOSBIntegrationProject/Services/ProxyService/PSLWReq"}
		"UAT_Middleware_CSC-LienFiling"{$endpointURL="http://stg-is.corporate.company.com/ERP/IFCSC_DiligenzIntegration/Services/ProxyService/PSReadLWCSCReq"}
		"UAT_Middleware_Vertex-LookUpTax"{$endpointURL="http://stg-is.corporate.company.com/IFLWVertexIntegrationProject/Services/Proxy/PSLWVertexLookUpTax"}
		"UAT_Middleware_Vertex-CalculateTax"{$endpointURL="http://stg-is.corporate.company.com/IFLWVertexIntegrationProject/Services/Proxy/PSLWVertexCalculateTax"}

		"SIT_Direct_Bridger"{$endpointURL="https://aml-esb2-test.capital.company.com/FSDMTransformationCASAService1/casaPort1"}
		"SIT_Direct_CARRA"{$endpointURL="https://aml-esb2-test.capital.company.com/FSDMTransformationCASAService1/casaPort1"}
		"SIT_Direct_CSC-LienFiling"{$endpointURL="https://eservices-test.diligenz.com/eservices.asmx"}
		"SIT_Direct_Vertex-LookUpTax"{$endpointURL="https://ge.ondemand.vertexinc.com/vertex-ws/services/LookupTaxAreasString"}
		"SIT_Direct_Vertex-CalculateTax"{$endpointURL="https://ge.ondemand.vertexinc.com/vertex-ws/services/CalculateTaxString"}
		
		"UAT_Direct_Bridger"{$endpointURL="https://aml-esb2-test.capital.company.com/FSDMTransformationCASAService1/casaPort1"}
		"UAT_Direct_CARRA"{$endpointURL="https://aml-esb2-test.capital.company.com/FSDMTransformationCASAService1/casaPort1"}
		"UAT_Direct_CSC-LienFiling"{$endpointURL="https://eservices-test.diligenz.com/eservices.asmx"}
		"UAT_Direct_Vertex-LookUpTax"{$endpointURL="https://ge.ondemand.vertexinc.com/vertex-ws/services/LookupTaxAreasString"}
		"UAT_Direct_Vertex-CalculateTax"{$endpointURL="https://ge.ondemand.vertexinc.com/vertex-ws/services/CalculateTaxString"}

		default {
			Write-Host "WSDL endpoint doesn't exist for $appEnvironment_$interfaceType_$appInterface Interface. Please try again."
			Break
		}
	}

	[xml] $soap = Get-Content $Basepath\$appInterface\$interfaceType"_Request.xml"

	if(($interfaceType+"_"+$appInterface).Contains("Direct_Vertex") -or ($interfaceType+"_"+$appInterface).Contains("Direct_CSC")){
		$result = Invoke-WebRequest -Uri $endpointURL -Method POST -Proxy $strZscalarProxy -Body $soap -ContentType "text/xml" -UseBasicParsing
	}
	else{
		$result = Invoke-WebRequest -Uri $endpointURL -Method POST -Body $soap -ContentType "text/xml" -UseBasicParsing
	}

    $result.Content.Replace("&lt;","<").Replace("&gt;",">") | Out-File -FilePath $Basepath\$appInterface\$interfaceType"_Response.xml"	

    If($result.Content.Contains("ResponseTimeMS") -or $result.Content.Contains("GetJurisdictionsResult") -or $result.Content.Contains("<StatusDesc>Successful</StatusDesc>")){
        $strResponse = "Success"
        Write-Host "Request is processed successfully with statusCode = "$result.StatusCode" and statusDescription = "$result.StatusDescription". Please review the Response.xml file at $Basepath\$appInterface"
    }
	else{
        $strResponse = "Failed"
        Write-Host "An Error occurred while processing the request. Please review the Response.xml file at $Basepath\$appInterface"
	}
	
	New-Object -TypeName PSCustomObject -Property @{
		Environment = $appEnvironment
		InterfaceName = $appInterface
		TestedVia = $interfaceType+" Endpoint"
		Result = $strResponse
	}| Format-Table -AutoSize -Property Environment, InterfaceName, TestedVia, Result

} #End Try
Catch
{
    Write-Host "An Error occurred while processing the request with Error message - "$_.Exception.Message" Please review the Request.xml file at $Basepath\$appInterface"
    "`nException Description - " + ($_.ErrorDetails.Message).ToString().Trim() + "`n"
	Break
}