USE [Reports]
GO

/****** Object:  Table [dbo].[BillCode]    Script Date: 11/09/2017 9:08:25 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BillCode](
	[BillCode] [varchar](5) NOT NULL,
	[Description] [varchar](50) NULL
) ON [PRIMARY]

GO

----------------------------------------------------------

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
USE [Reports]
GO

/****** Object:  Table [dbo].[BillingCode]    Script Date: 11/09/2017 9:10:00 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BillingCode](
	[NarrativeCode] [varchar](10) NOT NULL,
	[WIPCode] [varchar](10) NOT NULL,
	[ClientBillingCode] [varchar](10) NULL,
	[BillCode] [varchar](5) NULL,
	[Description] [varchar](250) NULL,
 CONSTRAINT [PK_BillingCode] PRIMARY KEY CLUSTERED 
(
	[NarrativeCode] ASC,
	[WIPCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO



-----------------------------------------------------------------------------


