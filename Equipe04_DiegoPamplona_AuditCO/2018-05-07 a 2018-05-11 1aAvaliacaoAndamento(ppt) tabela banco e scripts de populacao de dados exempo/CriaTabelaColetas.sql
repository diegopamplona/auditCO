USE [AuditCODB]
GO

/****** Object:  Table [dbo].[Coletas]    Script Date: 01/07/2018 23:41:49 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Coletas](
	[ColetaID] [int] IDENTITY(1,1) NOT NULL,
	[datacoleta] [date] NULL,
	[horacoleta] [time](7) NULL,
	[nmhost] [nvarchar](150) NULL,
	[nmuserlogado] [nvarchar](150) NULL,
	[nmcontdomhost] [nvarchar](150) NULL,
	[uptimehost] [nvarchar](150) NULL,
	[timezonehost] [nvarchar](150) NULL,
	[prochost] [nvarchar](150) NULL,
	[qtdcorhost] [nvarchar](150) NULL,
	[nmsoatualhost] [nvarchar](150) NULL,
	[platsoatualhost] [nvarchar](150) NULL,
	[lingsoatualhost] [nvarchar](150) NULL,
	[qtdmemhost] [nvarchar](150) NULL,
	[qtdmemdisphost] [nvarchar](150) NULL,
	[discosohost] [nvarchar](150) NULL,
	[formatdiscosohost] [nvarchar](150) NULL,
	[tamtotdiscosohost] [nvarchar](150) NULL,
	[tamdispdiscosohost] [nvarchar](150) NULL,
	[nomeniccabo] [nvarchar](150) NULL,
	[descniccabo] [nvarchar](150) NULL,
	[macaddressniccabo] [nvarchar](150) NULL,
	[enderipniccabo] [nvarchar](150) NULL,
	[nomenicsemfio] [nvarchar](150) NULL,
	[descnicsemfio] [nvarchar](150) NULL,
	[macaddressnicsemfio] [nvarchar](150) NULL,
	[enderipnicsemfio] [nvarchar](150) NULL,
	[placavideo] [nvarchar](150) NULL,
	[resolucaoatual] [nvarchar](150) NULL,
	[impressorainst1] [nvarchar](150) NULL,
	[impressorainst2] [nvarchar](150) NULL,
	[impressorainst3] [nvarchar](150) NULL,
	[impressorainst4] [nvarchar](150) NULL,
	[impressorainst5] [nvarchar](150) NULL,
	[impressorainst6] [nvarchar](150) NULL,
	[impressorainst7] [nvarchar](150) NULL,
	[chavewindows] [nvarchar](150) NULL,
	[chaveoffice] [nvarchar](150) NULL,
	[instanciasql] [nvarchar](150) NULL,
	[programinstal1] [nvarchar](150) NULL,
	[programinstal2] [nvarchar](150) NULL,
	[programinstal3] [nvarchar](150) NULL,
	[programinstal4] [nvarchar](150) NULL,
	[programinstal5] [nvarchar](150) NULL,
	[programinstal6] [nvarchar](150) NULL,
	[programinstal7] [nvarchar](150) NULL,
	[programinstal8] [nvarchar](150) NULL,
	[programinstal9] [nvarchar](150) NULL,
	[programinstal10] [nvarchar](150) NULL,
	[programinstal11] [nvarchar](150) NULL,
	[programinstal12] [nvarchar](150) NULL,
	[programinstal13] [nvarchar](150) NULL,
	[programinstal14] [nvarchar](150) NULL,
	[programinstal15] [nvarchar](150) NULL,
	[programinstal16] [nvarchar](150) NULL,
	[programinstal17] [nvarchar](150) NULL,
	[programinstal18] [nvarchar](150) NULL,
	[programinstal19] [nvarchar](150) NULL,
	[programinstal20] [nvarchar](150) NULL,
	[programinstal21] [nvarchar](150) NULL,
	[programinstal22] [nvarchar](150) NULL,
	[programinstal23] [nvarchar](150) NULL,
	[programinstal24] [nvarchar](150) NULL,
	[programinstal25] [nvarchar](150) NULL,
	[programinstal26] [nvarchar](150) NULL,
	[programinstal27] [nvarchar](150) NULL,
	[programinstal28] [nvarchar](150) NULL,
	[programinstal29] [nvarchar](150) NULL,
	[programinstal30] [nvarchar](150) NULL,
	[programinstal31] [nvarchar](150) NULL,
	[programinstal32] [nvarchar](150) NULL,
	[programinstal33] [nvarchar](150) NULL,
	[programinstal34] [nvarchar](150) NULL,
	[programinstal35] [nvarchar](150) NULL,
	[programinstal36] [nvarchar](150) NULL,
	[programinstal37] [nvarchar](150) NULL,
	[programinstal38] [nvarchar](150) NULL,
	[programinstal39] [nvarchar](150) NULL,
	[programinstal40] [nvarchar](150) NULL,
	[programinstal41] [nvarchar](150) NULL,
	[programinstal42] [nvarchar](150) NULL,
	[programinstal43] [nvarchar](150) NULL,
	[programinstal44] [nvarchar](150) NULL,
	[programinstal45] [nvarchar](150) NULL,
	[programinstal46] [nvarchar](150) NULL,
	[programinstal47] [nvarchar](150) NULL,
	[programinstal48] [nvarchar](150) NULL,
	[programinstal49] [nvarchar](150) NULL,
	[programinstal50] [nvarchar](150) NULL,
	[programinstal51] [nvarchar](150) NULL,
	[programinstal52] [nvarchar](150) NULL,
	[programinstal53] [nvarchar](150) NULL,
	[programinstal54] [nvarchar](150) NULL,
	[programinstal55] [nvarchar](150) NULL,
	[programinstal56] [nvarchar](150) NULL,
	[programinstal57] [nvarchar](150) NULL,
	[programinstal58] [nvarchar](150) NULL,
	[programinstal59] [nvarchar](150) NULL,
	[programinstal60] [nvarchar](150) NULL,
	[programinstal61] [nvarchar](150) NULL,
	[programinstal62] [nvarchar](150) NULL,
	[programinstal63] [nvarchar](150) NULL,
	[programinstal64] [nvarchar](150) NULL,
	[programinstal65] [nvarchar](150) NULL,
	[programinstal66] [nvarchar](150) NULL,
	[programinstal67] [nvarchar](150) NULL,
	[programinstal68] [nvarchar](150) NULL,
	[programinstal69] [nvarchar](150) NULL,
	[programinstal70] [nvarchar](150) NULL,
	[programinstal71] [nvarchar](150) NULL,
	[programinstal72] [nvarchar](150) NULL,
	[programinstal73] [nvarchar](150) NULL,
	[programinstal74] [nvarchar](150) NULL,
	[programinstal75] [nvarchar](150) NULL,
	[programinstal76] [nvarchar](150) NULL,
	[programinstal77] [nvarchar](150) NULL,
	[programinstal78] [nvarchar](150) NULL,
	[programinstal79] [nvarchar](150) NULL,
	[programinstal80] [nvarchar](150) NULL,
	[programinstal81] [nvarchar](150) NULL,
	[programinstal82] [nvarchar](150) NULL,
	[programinstal83] [nvarchar](150) NULL,
	[programinstal84] [nvarchar](150) NULL,
	[programinstal85] [nvarchar](150) NULL,
	[programinstal86] [nvarchar](150) NULL,
	[programinstal87] [nvarchar](150) NULL,
	[programinstal88] [nvarchar](150) NULL,
	[programinstal89] [nvarchar](150) NULL,
	[programinstal90] [nvarchar](150) NULL,
	[programinstal91] [nvarchar](150) NULL,
	[programinstal92] [nvarchar](150) NULL,
	[programinstal93] [nvarchar](150) NULL,
	[programinstal94] [nvarchar](150) NULL,
 CONSTRAINT [PK_Coletas] PRIMARY KEY CLUSTERED 
(
	[ColetaID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

