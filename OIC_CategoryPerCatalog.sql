USE [overseas_prod]
GO

/****** Object:  Table [dbo].[OIC_CategoryPerCatalog]    Script Date: 7/15/2014 9:21:01 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[OIC_CategoryPerCatalog](
	[uniqueidcol] [numeric](18, 0) NOT NULL,
	[IdCatalog] [int] NOT NULL,
	[CategoryId] [int] NOT NULL,
	[orden] [int] NULL,
	[activo] [int] NULL
) ON [PRIMARY]

GO


