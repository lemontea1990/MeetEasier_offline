USE [ews]
GO

/****** Object:  Table [dbo].[Meeting_Status_Loop]    Script Date: 6/21/2020 1:31:46 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Meeting_Status_Loop](
	[id] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[Room_Name] [nvarchar](max) NULL,
	[Start_Time] [nvarchar](max) NULL,
	[End_Time] [nvarchar](max) NULL,
	[Organzier] [nvarchar](max) NULL,
	[Location] [nvarchar](max) NULL,
	[timestamp] [nvarchar](max) NULL,
	[Subject] [nvarchar](max) NULL,
	[Opt_att] [nvarchar](max) NULL,
	[Req_att] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
