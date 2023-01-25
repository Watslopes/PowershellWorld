--https://stackoverflow.com/questions/28237117/how-to-send-a-chart-as-an-email-body-using-powershell

select Count(*) from dtp.US_NEA_Extract_TIAA_GL_FORMAT where effectivemonth = 'AUG-2020'

select * from sys.procedures where name = 'SP_ParentDataLoad'

select *  from JobDetailsView WITH (NOLOCK) where JobName like '%DAily%'

select top 7 Convert(DAte, JobStepStartDate), DateDiff(s, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) from JobDetailsView WITH (NOLOCK) where JobName like '%DAily%' and TaskNAme = 'Loan Income Amort' order by JobInstanceId desc

select top 70 * from JobDetailsView WITH (NOLOCK) where JobName like '%DAily%' 
and TaskNAme = 'Loan Interim Interest' order by JobInstanceId desc

sp_helptext JobDetailsView

Select Convert(DAte, JobStepStartDate) from JobDetailsView WITH (NOLOCK)


select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Loan Interim Interest' order by JobInstanceId desc
select top 7 Convert(DAte, JobStepStartDate) As Day, DateDiff(n, CONVERT(DATETIME, JobStepStartDate), CONVERT(DATETIME, JobStepEndDate)) as Time_in_Min from JobDetailsView WITH (NOLOCK) where JobName like '%Daily%' and TaskNAme = 'Loan Income Amort' order by JobInstanceId desc
