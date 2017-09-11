USE [Reports]
GO

/****** Object:  Table [dbo].[BillTransactionDetails]    Script Date: 11/09/2017 9:05:11 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BillTransactionDetails](
	[Country] [varchar](5) NULL,
	[LawFirm] [varchar](10) NULL,
	[CurrencyCode] [varchar](10) NULL,
	[ChargedDate] [datetime] NULL,
	[MonInvDate] [datetime] NULL,
	[DebitNoteNo] [varchar](10) NOT NULL,
	[InvoiceNo] [varchar](20) NOT NULL,
	[ClientCode] [varchar](10) NULL,
	[ClientRefNo] [varchar](100) NULL,
	[FileRefNo] [varchar](12) NULL,
	[AttorneyName] [varchar](30) NULL,
	[HourlyRate] [varchar](10) NULL,
	[BillCode] [varchar](2) NULL,
	[ApplicationNo] [varchar](50) NULL,
	[PatentNo] [varchar](50) NULL,
	[BillingCode] [nvarchar](10) NULL,
	[WIPCode] [varchar](10) NOT NULL,
	[NarrativeCode] [varchar](10) NULL,
	[ServiceFee] [decimal](18, 2) NULL,
	[OfficialFee] [decimal](18, 2) NULL,
	[Others] [decimal](18, 2) NULL,
	[TotalAmount] [decimal](18, 2) NULL,
	[Remarks] [varchar](2000) NULL,
	[DocType] [char](1) NULL
) ON [PRIMARY]

GO


-----------------------------------------------------------------------------------------
USE [Reports]
GO

/****** Object:  Table [dbo].[BillTransactionImport]    Script Date: 11/09/2017 9:05:33 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BillTransactionImport](
	[CurrencyCode] [varchar](10) NULL,
	[ChargedDate] [datetime] NULL,
	[InvoiceNo] [varchar](20) NOT NULL,
	[IRN] [nvarchar](20) NULL,
	[FileRefNo] [varchar](20) NULL,
	[FileRefNo_SF] [varchar](20) NULL,
	[AttorneyCode] [varchar](10) NULL,
	[WIPCode] [varchar](10) NOT NULL,
	[NarrativeCode] [varchar](10) NULL,
	[Narrative] [text] NULL,
	[TotalAmount] [decimal](18, 2) NULL,
	[DocType] [char](1) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO


--------------------------------------------------------------------------
USE [Reports]
GO

/****** Object:  Table [dbo].[BillTransactionSummaryRep]    Script Date: 11/09/2017 9:06:18 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BillTransactionSummaryRep](
	[InvoiceNo] [varchar](30) NOT NULL,
	[DebitNoteNo] [varchar](30) NULL,
	[BillCode] [varchar](2) NULL,
	[ServiceFee] [decimal](18, 2) NULL,
	[OfficialFee] [decimal](18, 2) NULL,
	[Others] [decimal](18, 2) NULL,
	[TotalAmount] [decimal](18, 2) NULL
) ON [PRIMARY]

GO

--------------------------------------------------------------------------------------------------

USE [Reports]
GO

/****** Object:  Table [dbo].[BillTransactionError]    Script Date: 11/09/2017 9:08:08 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BillTransactionError](
	[SNo] [bigint] IDENTITY(1,1) NOT NULL,
	[ErrorType] [varchar](100) NULL,
	[ChargedDate] [datetime] NULL,
	[InvoiceNo] [varchar](20) NULL,
	[FileRefNo] [varchar](12) NULL,
	[AttorneyCode] [varchar](30) NULL,
	[WIPCode] [varchar](10) NULL,
	[NarrativeCode] [nvarchar](10) NULL,
	[TotalAmount] [decimal](18, 2) NULL
) ON [PRIMARY]

GO

-------------------------------------------------------------
USE [Reports]
GO

/****** Object:  Table [dbo].[BillDataError]    Script Date: 11/09/2017 9:08:44 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BillDataError](
	[Status] [float] NULL,
	[FileRef] [nvarchar](255) NULL,
	[UploadCode] [nvarchar](255) NULL,
	[TrxID] [float] NULL,
	[Postedsum] [float] NULL,
	[UploadSum] [float] NULL,
	[Billstatus] [float] NULL
) ON [PRIMARY]

GO


----------------------------------------------------------------------------------
