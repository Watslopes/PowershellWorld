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
 $ChartArea.Area3DStyle = $chartarea3D
 $ChartArea = $mychart.ChartAreas.Add('ChartArea');
 $chartarea.AxisX.Interval = 1
 $chartarea.AxisY.Interval = 6

 # legend 
 $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
 $legend.name = "Legend1"
 $mychart.Legends.Add($legend)
#endregion 

function Draw-Chart ($chartname, $charttype, $chartcolor, $constr, $x, $y)
{

    $chartarea.AxisY.Title = $Y
    $chartarea.AxisX.Title = $X

    $datasource = $constr
    [void]$mychart.Series.Add($chartname)
    $mychart.Series[$chartname].ChartType = $charttype
    $mychart.Series[$chartname].IsVisibleInLegend = $true
    $mychart.Series[$chartname].BorderWidth  = 2
    #$mychart.Series[$chartname].chartarea = "ChartArea1"
    $mychart.Series[$chartname].Legend = "Legend1"
    $mychart.Series[$chartname].color = $chartcolor
    $datasource | ForEach-Object  {$mychart.Series[$chartname].Points.addxy($_.$x, $_.$y) }
}

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select TaskName as JobTaskName, AVG(DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate))) as AvgTime from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' And JobStepStartDate between (GETDATE()-8) and (GETDATE()-1) Group By TaskName"
Draw-Chart 'AVG JobTime in Mins (For last 7 days)' 'Column' 'Green' $constr 'JobTaskName' 'AvgTime'
$mychart.SaveImage("$PSScriptRoot\LW_Chart2.png","png")

#region 2
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
 $ChartArea.Area3DStyle = $chartarea3D
 $ChartArea = $mychart.ChartAreas.Add('ChartArea');
 $chartarea.AxisX.Interval = 1
 $chartarea.AxisY.Interval = 6

 # legend 
 $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
 $legend.name = "Legend1"
 $mychart.Legends.Add($legend)

#endregion

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Receivable to GL' order by JobInstanceId desc"
Draw-Chart 'Post Receivable to GL' 'line' 'Silver' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Receivable Tax To GL' order by JobInstanceId desc"
Draw-Chart 'Post Receivable Tax To GL' 'line' 'Brown' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'RNI Update' order by JobInstanceId desc"
Draw-Chart 'RNI Update' 'line' 'SkyBlue' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Collections' order by JobInstanceId desc"
Draw-Chart 'Collections' 'line' 'YelloW' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'SYNAM Export' order by JobInstanceId desc"
Draw-Chart 'SYNAM Export' 'line' 'Black' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Receipt Post By Lockbox' order by JobInstanceId desc"
Draw-Chart 'Receipt Post By Lockbox' 'line' 'Purple' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Increment Business Date' order by JobInstanceId desc"
Draw-Chart 'Increment Business Date' 'line' 'DarkOrange' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Invoice Generation' order by JobInstanceId desc"
Draw-Chart 'Invoice Generation' 'line' 'Gray' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Lease Extension AR Update' order by JobInstanceId desc"
Draw-Chart 'Lease Extension AR Update' 'line' 'Tan' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Lease Float Rate Update' order by JobInstanceId desc"
Draw-Chart 'Lease Float Rate Update' 'line' 'Teal' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Recurring Sundry Update' order by JobInstanceId desc"
Draw-Chart 'Recurring Sundry Update' 'line' 'Cyan' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Loan Interim Interest' order by JobInstanceId desc"
Draw-Chart 'Loan Interim Interest' 'line' 'Gold' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Loan Income Amort' order by JobInstanceId desc"
Draw-Chart 'Loan Income Amort' 'line' 'Blue' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Generate Cash Reports' order by JobInstanceId desc"
Draw-Chart 'Generate Cash Reports' 'line' 'LIME' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Income Recognition' order by JobInstanceId desc"
Draw-Chart 'Income Recognition' 'line' 'Red' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Book Depreciation Amort' order by JobInstanceId desc"
Draw-Chart 'Book Depreciation Amort' 'line' 'Pink' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Post Payable To GL' order by JobInstanceId desc"
Draw-Chart 'Post Payable To GL' 'line' 'FUCHSIA' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Tax Depreciation Amort' order by JobInstanceId desc"
Draw-Chart 'Tax Depreciation Amort' 'line' 'Violet' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Tax Dep Amortization GL' order by JobInstanceId desc"
Draw-Chart 'Tax Dep Amortization GL' 'line' 'crimson' $constr 'Day' 'Time_in_Min'

$constr = Invoke-Sqlcmd2 -ServerInstance 'SQLInstance' -Database 'DBName' -Query "select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Sales Tax Assessment' order by JobInstanceId desc"
Draw-Chart 'Sales Tax Assessment' 'line' 'Green' $constr 'Day' 'Time_in_Min'

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

Send-MailMessage -To "watson.lopes@company.com" `
    -from "LW_NoReply@company.com" `
    -SmtpServer "e2ksmtp01.e2k.ad.company.com" `
    -Subject "Application Job Timings - Data Visualization" `
    -BodyAsHtml -body $body `
    -Attachments "$PSScriptRoot\LW_Chart.png", "$PSScriptRoot\LW_Chart2.png"