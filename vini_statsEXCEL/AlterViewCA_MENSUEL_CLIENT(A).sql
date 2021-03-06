USE [vnc5]
GO

/****** Object:  View [dbo].[CA_MENSUELCLIENT]    Script Date: 09/06/2017 10:36:33 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER VIEW [dbo].[CA_MENSUELCLIENT]
AS
SELECT        TOP (100) PERCENT dbo.CLIENT.CLT_CODE, dbo.CLIENT.CLT_NOM, dbo.CLIENT.CLT_RS, YEAR(dbo.COMMANDE.CMD_DATE) AS annee, MONTH(dbo.COMMANDE.CMD_DATE) AS mois, 
                         SUM(dbo.COMMANDE.CMD_TOTAL_HT) AS CAHT, CMD_ORIGINE
FROM            dbo.CLIENT INNER JOIN
                         dbo.COMMANDE ON dbo.CLIENT.CLT_ID = dbo.COMMANDE.CMD_CLT_ID
WHERE        (YEAR(dbo.COMMANDE.CMD_DATE) >= YEAR(GETDATE()) - 3)
GROUP BY dbo.CLIENT.CLT_CODE, YEAR(dbo.COMMANDE.CMD_DATE), MONTH(dbo.COMMANDE.CMD_DATE), dbo.CLIENT.CLT_NOM, dbo.CLIENT.CLT_RS, CMD_ORIGINE

GO


