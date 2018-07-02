USE [AuditCODB]
GO

/****** Object:  Table [dbo].[Hosts]    Script Date: 02/07/2018 10:09:53 ******/
DROP TABLE [dbo].[Hosts]
GO

/****** Object:  Table [dbo].[Hosts]    Script Date: 02/07/2018 10:09:53 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Hosts](
	[HostID] [int] IDENTITY(1,1) NOT NULL,
	[nmhost] [varchar](150) NULL,
	[datacadastro] [date] NULL,
	[nronffiscal] [int] NULL,
 CONSTRAINT [PK_HostID] PRIMARY KEY CLUSTERED 
(
	[HostID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


