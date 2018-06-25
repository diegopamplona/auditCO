SET QUOTED_IDENTIFIER OFF;
GO
USE [GeraDetiDB];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

IF OBJECT_ID(N'[dbo].[Parametros]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Parametros];
GO

CREATE TABLE [dbo].[Parametros] (
    [ParameterID] int IDENTITY(1,1) NOT NULL,
    [cdparam] nvarchar(50)  NULL,
    [vlparam] nvarchar(50)  NULL,
);
GO

ALTER TABLE [dbo].[Parametros]
ADD CONSTRAINT [PK_Parametros]
    PRIMARY KEY CLUSTERED ([ParameterID] ASC);