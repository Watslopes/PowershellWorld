Select AT.File_Source, FS.Content, AC.EntityNaturalId from Attachments AT
Inner Join ActivityAttachments(Nolock) AAT on AT.Id = AAT.AttachmentID And AAT.IsActive = 1
Inner Join TransactionInstances(Nolock) TI on TI.EntityId = AAT.Activityid And TI.EntityName = 'Activity'
Inner Join Activities(Nolock) AC on AC.Id = AAT.ActivityId And AC.IsActive = 1 AND AC.StatusId = 1
Inner Join ActivityTypes(Nolock) ACT on ACT.Id = AC.ActivityTypeId
Inner Join FileStores(Nolock) FS on FS.GUID = Replace(convert(nvarchar(max), AT.File_Content,0),'GUID:','')
Where ACT.Name in ('Rebooking','Restructure') And
(AC.Name like '%COVID%' OR AC.Name like '%CV%19%')
And ActivityTypeID in (420,421)
And AC.OwnerId in (select UserId from UserReportingToes
Where reportingtoid = (Select Id from USers Where FullName = 'Mary Beres')
And UserId <> (Select Id from Users where LoginName = 'Booking_Funding')
And IsActive = 1)
And And AC.EntityNaturalId not in (<Seq#s from LogFile>)