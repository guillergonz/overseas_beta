USE [overseas_prod]
GO

/****** Object:  Table [dbo].[OIC_PartsPerCategory]    Script Date: 7/15/2014 9:21:36 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[OIC_PartsPerCategory](
	[uniqueidcol] [numeric](18, 0) NOT NULL,
	[partno] [varchar](20) NOT NULL,
	[image] [varchar](100) NOT NULL,
	[link] [varchar](150) NOT NULL,
	[IdCatalog] [int] NOT NULL,
	[categoryid] [int] NOT NULL,
	[titulo] [varchar](50) NULL,
	[nota] [varchar](50) NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

