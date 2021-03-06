USE [vnc4]
GO
/****** Objet :  StoredProcedure [dbo].[SELECT_SCMD_SANSLIGNE]    Date de génération du script : 07/15/2009 19:01:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	Liste des SousCommande Sans lignes
-- =============================================
CREATE PROCEDURE [dbo].[SELECT_SCMD_SANSLIGNE] 
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
SELECT     DISTINCT COMMANDE.CMD_CODE as 'Code Commande', COMMANDE.CMD_DATE as 'Date Commande', 
                      COMMANDE.CMD_ETAT as 'Etat Comamnde'
FROM         SOUSCOMMANDE INNER JOIN
                      COMMANDE ON SOUSCOMMANDE.SCMD_CMD_ID = COMMANDE.CMD_ID LEFT OUTER JOIN
                      LIGNE_COMMANDE ON SOUSCOMMANDE.SCMD_ID = LIGNE_COMMANDE.LGCM_SCMD_ID
GROUP BY SOUSCOMMANDE.SCMD_ID, SOUSCOMMANDE.SCMD_CODE, COMMANDE.CMD_ID, COMMANDE.CMD_CODE, COMMANDE.CMD_DATE, 
                      COMMANDE.CMD_ETAT
HAVING      (COUNT(LIGNE_COMMANDE.LGCM_ID) = 0) AND CMD_DATE > '01/01/2009'

END

USE [vnc4]
GO
/****** Objet :  StoredProcedure [dbo].[SUPPR_SCMD_CMD]    Date de génération du script : 07/15/2009 19:01:21 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	Suppression des SousCommandes sans lignes d'une commande
-- =============================================
CREATE PROCEDURE [dbo].[SUPPR_SCMD_CMD] 
(@CmdCode varchar(10))
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
DELETE FROM SOUSCOMMANDE WHERE SCMD_CMD_ID = (SELECT CMD_ID FROM COMMANDE WHERE COMMANDE.CMD_CODE = @CmdCode AND COMMANDE.CMD_CODE in (Select COMMANDE.CMD_CODE FROM SOUSCOMMANDE INNER JOIN
                      COMMANDE ON SOUSCOMMANDE.SCMD_CMD_ID = COMMANDE.CMD_ID LEFT OUTER JOIN
                      LIGNE_COMMANDE ON SOUSCOMMANDE.SCMD_ID = LIGNE_COMMANDE.LGCM_SCMD_ID
GROUP BY SOUSCOMMANDE.SCMD_ID, SOUSCOMMANDE.SCMD_CODE, COMMANDE.CMD_ID, COMMANDE.CMD_CODE, COMMANDE.CMD_DATE, 
                      COMMANDE.CMD_ETAT
HAVING      (COUNT(LIGNE_COMMANDE.LGCM_ID) = 0) AND CMD_DATE > '01/01/2009'
))
END











