USE [LW_Reporting]
GO

/****** Object:  Table [dbo].[tbl_covid19restructures]    Script Date: 4/10/2020 9:07:18 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tbl_covid19restructures](
	[SequenceNumber] [varchar](215) NULL,
	[CustomerNumber] [varchar](215) NULL,
	[RequestNumber] [varchar](215) NULL,
	[ExposureAmount] [money] NULL,
	[LegalEntity] [varchar](3) NULL,
	[DealType] [varchar](5) NULL,
	[ProductType] [varchar](10) NULL,
	[FundingSource] [varchar](100) NULL,
	[NextACHDueDate] [date] NULL,
	[ReportStatus] [varchar](20) NULL,
	[Status] [varchar](20) NULL,
	[RequestDate] [date] NULL,
	[NumberofMonths] [int] NULL,
	[MonthlyRentPaymentAmount] [money] NULL,
	[TotalSkipPaymentsAmount] [money] NULL,
	[PaymentFrequency] [varchar](10) NULL,
	[DailyAdditionRejected?] [char](1) NULL
) ON [PRIMARY]

GO

