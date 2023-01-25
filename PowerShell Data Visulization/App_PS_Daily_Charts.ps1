Param(
     [Parameter()]
     [string]$SQLUsername = 'DBUsername',
     [Parameter()]
     [string]$SQLPassword = 'Password',
     $SQLServer = 'SQLInstance',
     $SQLDBName = 'DBName' 
)     

Try
{
    $ErrorActionPreference = 'Stop'
    Remove-Item –path $PSScriptRoot\*.png -ErrorAction SilentlyContinue

#region 1
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    # chart object
     $mychart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
     $mychart.Width = 1200
     $mychart.Height = 500
     $mychart.BackColor = [System.Drawing.Color]::lightcyan

    # title 
     [void]$mychart.Titles.Add("Data Visualization Chart : Job Steps which takes > 3-5 Mins on an avg.")
     $mychart.Titles[0].Font = "Arial,13pt"
     $mychart.Titles[0].Alignment = "topLeft"

     # chart area 

     $chartarea3D = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea3DStyle
     # $chartarea.Name = "ChartArea1"
     $chartarea3D.Enable3D = $false    
     #$ChartArea.Area3DStyle = $chartarea3D
     $ChartArea = $mychart.ChartAreas.Add('ChartArea');
     $chartarea.AxisX.Interval = 1
     $chartarea.AxisY.Interval = 6

     # legend 
     $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
     $legend.name = "Legend1"
     $mychart.Legends.Add($legend)
    #endregion 

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Income Recognition' order by JobInstanceId desc" 1
    Draw-Chart 'Income Recognition' 'spline' 'Red' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Loan Income Amort' order by JobInstanceId desc" 1
    Draw-Chart 'Loan Income Amort' 'spline' 'Blue' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Sales Tax Assessment' order by JobInstanceId desc" 1
    Draw-Chart 'Sales Tax Assessment' 'spline' 'Green' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Collections' order by JobInstanceId desc" 1
    Draw-Chart 'Collections' 'spline' 'YelloW' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Recurring Sundry Update' order by JobInstanceId desc" 1
    Draw-Chart 'Recurring Sundry Update' 'spline' 'Cyan' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'SYNAM Export' order by JobInstanceId desc" 1
    Draw-Chart 'SYNAM Export' 'spline' 'Black' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'RNI Update' order by JobInstanceId desc" 1
    Draw-Chart 'RNI Update' 'spline' 'FUCHSIA' $DBTable 'Day' 'Time_in_Min'

    $mychart.SaveImage("$PSScriptRoot\LW_Chart1.png","png")

    #Remove-Variable mychart, DBTable, legend, chartarea, chartarea3D
    Remove-Variable -Name mychart, DBTable, legend, chartarea, chartarea3D -ErrorAction SilentlyContinue 
#End Region
    #region 2
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    # chart object
     $mychart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
     $mychart.Width = 1200
     $mychart.Height = 500
     $mychart.BackColor = [System.Drawing.Color]::lightcyan

    # title 
     [void]$mychart.Titles.Add("Data Visualization Chart : Job Steps which takes < 3-5 Mins on an avg.")
     $mychart.Titles[0].Font = "Arial,13pt"
     $mychart.Titles[0].Alignment = "topLeft"

     # chart area 

     $chartarea3D = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea3DStyle
     # $chartarea.Name = "ChartArea1"
     $chartarea3D.Enable3D = $false    
     #$ChartArea.Area3DStyle = $chartarea3D
     $ChartArea = $mychart.ChartAreas.Add('ChartArea');
     $chartarea.AxisX.Interval = 1
     $chartarea.AxisY.Interval = 1

     # legend 
     $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
     $legend.name = "Legend1"
     $mychart.Legends.Add($legend)

    #endregion

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Receivable to GL' order by JobInstanceId desc" 1
    Draw-Chart 'Post Receivable to GL' 'spline' 'Red' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Receivable Tax To GL' order by JobInstanceId desc" 1
    Draw-Chart 'Post Receivable Tax To GL' 'spline' 'Green' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Receipt Post By Lockbox' order by JobInstanceId desc" 1
    Draw-Chart 'Receipt Post By Lockbox' 'spline' 'Purple' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Increment Business Date' order by JobInstanceId desc" 1
    Draw-Chart 'Increment Business Date' 'spline' 'DarkOrange' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Invoice Generation' order by JobInstanceId desc" 1
    Draw-Chart 'Invoice Generation' 'spline' 'Blue' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Lease Extension AR Update' order by JobInstanceId desc" 1
    Draw-Chart 'Lease Extension AR Update' 'spline' 'Yellow' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Lease Float Rate Update' order by JobInstanceId desc" 1
    Draw-Chart 'Lease Float Rate Update' 'spline' 'Teal' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Loan Interim Interest' order by JobInstanceId desc" 1
    Draw-Chart 'Loan Interim Interest' 'spline' 'Black' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Generate Cash Reports' order by JobInstanceId desc" 1
    Draw-Chart 'Generate Cash Reports' 'spline' 'Cyan' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Book Depreciation Amort' order by JobInstanceId desc" 1
    Draw-Chart 'Book Depreciation Amort' 'spline' 'Pink' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Payable To GL' order by JobInstanceId desc" 1
    Draw-Chart 'Post Payable To GL' 'spline' 'FUCHSIA' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Tax Depreciation Amort' order by JobInstanceId desc" 1
    Draw-Chart 'Tax Depreciation Amort' 'spline' 'Violet' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(Date, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Tax Dep Amortization GL' order by JobInstanceId desc" 1
    Draw-Chart 'Tax Dep Amortization GL' 'spline' 'Crimson' $DBTable 'Day' 'Time_in_Min'

    # save chart
    $mychart.SaveImage("$PSScriptRoot\LW_Chart2.png","png")

    $Subject = "SIT: LW JobTimings - Daily Trends :"+(Get-Date).AddDays(0).ToString('dd-MMM-yyyy')
    $Body = @"
    <html>
        <body style="font-family:calibri">Greetings,<br><br>Below id the company.com Daily jobs timings analysis for <b>past 7 days</b> : <br><br>
            <b><u>company.com Daily Jobs Timings Chart 1</u></b>
            <img src='cid:LW_Chart1.png'>
            <br><b><u>company.com Daily Jobs Timings Chart 2</u></b>
            <img src='cid:LW_Chart2.png'>
            <br><br><br>Regards,<br>company.com Support Team
        </body>
    </html>
"@

    Send-MailMessage -To "watson.lopes@company.com" -from "if-lw_nonProd-datavisualization@company.com" -SmtpServer "e2ksmtp01.e2k.ad.company.com" -Subject $Subject -BodyAsHtml -body $body `
        -Attachments "$PSScriptRoot\LW_Chart1.png", "$PSScriptRoot\LW_Chart2.png"

} #End Try
Catch
{
    #Send error email with provided parameters
    Write-Output $Error[0]
    #Send-ErrorEmail -emailerror $Error[0] -Emailtag $Emailtag
}
Finally
{
	$ErrorActionPreference = "SilentlyContinue"	
	Clear-Variable SQLPassword, DBTable, mychart
}