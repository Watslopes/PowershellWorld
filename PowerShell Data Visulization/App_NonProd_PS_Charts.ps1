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
    #$scriptpath = Split-Path -parent $MyInvocation.MyCommand.Definition

    # chart object
     $mychart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
     $mychart.Width = 1400
     $mychart.Height = 600
     $mychart.BackColor = [System.Drawing.Color]::lightcyan

    # title 
     [void]$mychart.Titles.Add("Data Visualization - Application Daily Jobs Timings")
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


    $DBTable = Execute-SQLQuery "select TaskName as JobTaskName, AVG(DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate))) as AvgTime from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' And JobStepStartDate between (GETDATE()-8) and (GETDATE()-1) Group By TaskName" 1
    Draw-Chart 'AVG JobTime in Mins (For last 7 days)' 'Column' 'Green' $DBTable 'JobTaskName' 'AvgTime'
    $mychart.SaveImage("$PSScriptRoot\LW_Chart2.png","png")

    #region 2
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
    #$scriptpath = Split-Path -parent $MyInvocation.MyCommand.Definition

    # chart object
     $mychart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
     $mychart.Width = 1400
     $mychart.Height = 600
     $mychart.BackColor = [System.Drawing.Color]::lemonchiffon

    # title 
     [void]$mychart.Titles.Add("Data Visualization - Application Daily Jobs Timings")
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

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Receivable to GL' order by JobInstanceId desc" 1
    Draw-Chart 'Post Receivable to GL' 'line' 'Silver' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Receivable Tax To GL' order by JobInstanceId desc" 1
    Draw-Chart 'Post Receivable Tax To GL' 'line' 'Brown' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'RNI Update' order by JobInstanceId desc" 1
    Draw-Chart 'RNI Update' 'line' 'SkyBlue' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Collections' order by JobInstanceId desc" 1
    Draw-Chart 'Collections' 'line' 'YelloW' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'SYNAM Export' order by JobInstanceId desc" 1
    Draw-Chart 'SYNAM Export' 'line' 'Black' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Receipt Post By Lockbox' order by JobInstanceId desc" 1
    Draw-Chart 'Receipt Post By Lockbox' 'line' 'Purple' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Increment Business Date' order by JobInstanceId desc" 1
    Draw-Chart 'Increment Business Date' 'line' 'DarkOrange' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Invoice Generation' order by JobInstanceId desc" 1
    Draw-Chart 'Invoice Generation' 'line' 'Gray' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Lease Extension AR Update' order by JobInstanceId desc" 1
    Draw-Chart 'Lease Extension AR Update' 'line' 'Tan' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Lease Float Rate Update' order by JobInstanceId desc" 1
    Draw-Chart 'Lease Float Rate Update' 'line' 'Teal' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Recurring Sundry Update' order by JobInstanceId desc" 1
    Draw-Chart 'Recurring Sundry Update' 'line' 'Cyan' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Loan Interim Interest' order by JobInstanceId desc" 1
    Draw-Chart 'Loan Interim Interest' 'line' 'Gold' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Loan Income Amort' order by JobInstanceId desc" 1
    Draw-Chart 'Loan Income Amort' 'line' 'Blue' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Generate Cash Reports' order by JobInstanceId desc" 1
    Draw-Chart 'Generate Cash Reports' 'line' 'LIME' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Income Recognition' order by JobInstanceId desc" 1
    Draw-Chart 'Income Recognition' 'line' 'Red' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Book Depreciation Amort' order by JobInstanceId desc" 1
    Draw-Chart 'Book Depreciation Amort' 'line' 'Pink' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Payable To GL' order by JobInstanceId desc" 1
    Draw-Chart 'Post Payable To GL' 'line' 'FUCHSIA' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Tax Depreciation Amort' order by JobInstanceId desc" 1
    Draw-Chart 'Tax Depreciation Amort' 'line' 'Violet' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Tax Dep Amortization GL' order by JobInstanceId desc" 1
    Draw-Chart 'Tax Dep Amortization GL' 'line' 'crimson' $DBTable 'Day' 'Time_in_Min'

    $DBTable = Execute-SQLQuery "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Sales Tax Assessment' order by JobInstanceId desc" 1
    Draw-Chart 'Sales Tax Assessment' 'line' 'Green' $DBTable 'Day' 'Time_in_Min'

    # save chart
    $mychart.SaveImage("$PSScriptRoot\LW_Chart.png","png")


    $Body = @"
    <html>
        <body style="font-family:calibri"> 
            <b><u>Application Job timings Overview (Past 7 days)</u></b>
            <img src='cid:LW_Chart.png'>
            <br><b><u>Average Time taken (in Mins) by Daily Job steps (past 7 days)</u></b>
            <img src='cid:LW_Chart2.png'>
        </body>
    </html>
"@

    Send-MailMessage -To "watson.lopes@company.com" -from "LW_NoReply@company.com" -SmtpServer "e2ksmtp01.e2k.ad.company.com" -Subject "SIT: Application Job Timings - Data Visualization" -BodyAsHtml -body $body `
        -Attachments "$PSScriptRoot\LW_Chart.png", "$PSScriptRoot\LW_Chart2.png"


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