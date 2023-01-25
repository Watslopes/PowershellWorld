#Variable to hold variable  
$SQLServer = "SQLINstance"  
$SQLDBName = "SQLDB"  
$uid ="UserName"  
$pwd = "Password"   
#SQL Query  
$SqlQuery = "SELECT top 5 AccountSchedule, Customer, CustomerNumber from DTP.ContractInformation_Report where EffectiveMonth = 'Oct-2019';"  
#$SqlQuery = "SELECT * from tableName;"  
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection  
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True;"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand  
$SqlCmd.CommandText = $SqlQuery  
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter  
$SqlAdapter.SelectCommand = $SqlCmd   
#Creating Dataset  
$DataSet = New-Object System.Data.DataSet  
$SqlAdapter.Fill($DataSet)  
$DataSet.Tables[0] | Select "AccountSchedule", "Customer", @{name="CustomerNumber"; expression={"'"+$_.CustomerNumber}} -First 5 | Export-Csv "C:\Users\502740204\Desktop\temp2237\test.csv" -NoTypeInformation
$SqlConnection.Close()


$DataSet.Tables[1] 