USE [overseas_prod]
GO

/****** Object:  Table [dbo].[OIC_Catalog]    Script Date: 7/15/2014 9:20:21 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[OIC_Catalog](
	[IdCatalog] [int] NOT NULL,
	[CatalogName] [varchar](50) NOT NULL,
	[CatalogPdfName] [varchar](50) NULL,
	[orden] [int] NULL,
	[activo] [int] NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

