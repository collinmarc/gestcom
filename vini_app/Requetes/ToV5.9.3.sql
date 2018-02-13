﻿--Ajout de clé Etrangères sur Precommandes
ALTER TABLE [dbo].[PRECOMMANDE]  WITH CHECK ADD  CONSTRAINT [FK_PRECOMMANDE_CLIENT] FOREIGN KEY([PCMD_CLT_ID])
REFERENCES [dbo].[CLIENT] ([CLT_ID]);
GO

ALTER TABLE [dbo].[PRECOMMANDE]  WITH CHECK ADD  CONSTRAINT [FK_PRECOMMANDE_PRODUIT] FOREIGN KEY([PCMD_PRD_ID])
REFERENCES [dbo].[PRODUIT] ([PRD_ID])
GO


-- Ajout de la notion de Dossier
ALTER TABLE dbo.PRODUIT ADD
	PRD_DOSSIER nvarchar(50) NULL;
GO
UPDATE PRODUIT SET PRD_DOSSIER = 'VINICOM';
GO
UPDATE PRODUIT SET PRD_DOSSIER = 'HOBIVIN' WHERE PRD_FRN_ID in (SELECT FRN_ID FROM FOURNISSEUR WHERE FRN_BEXP_INTERNET = '2');
GO

ALTER TABLE [dbo].[LIGNE_COMMANDE] DROP CONSTRAINT [FK_LIGNE_COMMANDE_COMMANDE]
GO
ALTER TABLE [dbo].[LIGNE_COMMANDE]  WITH NOCHECK ADD  CONSTRAINT [FK_LIGNE_V:\V5\vini_app\Requetes\ToV5.9.3.sqlCOMMANDE_COMMANDE] FOREIGN KEY([LGCM_CMD_ID])
REFERENCES [dbo].[COMMANDE] ([CMD_ID])
NOT FOR REPLICATION 
GO

-- TEST UNITAIRES
ALTER TABLE [dbo].[LIGNE_COMMANDE] NOCHECK CONSTRAINT [FK_LIGNE_COMMANDE_COMMANDE]
GO

-- TX COMMISSION
insert into FRN_COMM (FRNC_FRN_ID, FRNC_TYPE_CLIENT, FRNC_TX_COMM) 
SELECT  FRN_ID, 'INT', 12.0 from FOURNISSEUR


-- Etat des Stocks
/****** Object:  View [dbo].[RQ_PRODUIT_COMPLET]    Script Date: 05/02/2018 18:13:45 ******/
DROP VIEW [dbo].[RQ_PRODUIT_COMPLET]
GO

/****** Object:  View [dbo].[RQ_PRODUIT_COMPLET]    Script Date: 05/02/2018 18:13:45 ******/

CREATE VIEW [dbo].[RQ_PRODUIT_COMPLET]
AS
SELECT        dbo.PRODUIT.PRD_ID, dbo.PRODUIT.PRD_CODE, dbo.PRODUIT.PRD_LIBELLE, dbo.PRODUIT.PRD_MOT_CLE, dbo.PRODUIT.PRD_FRN_ID, dbo.PRODUIT.PRD_CONT_ID, dbo.PRODUIT.PRD_COND_ID, 
                         dbo.PRODUIT.PRD_COUL_ID, dbo.PRODUIT.PRD_MIL, dbo.PRODUIT.PRD_RGN_ID, dbo.PRODUIT.PRD_TVA_ID, dbo.PRODUIT.PRD_DATE_DERN_INVENT, dbo.PRODUIT.PRD_QTE_STK, 
                         dbo.PRODUIT.PRD_QTE_STOCK_DERN_INVENT, dbo.PRODUIT.PRD_DISPO, dbo.PRODUIT.PRD_CODE_STAT, dbo.RQ_Couleur.PAR_VALUE AS RQ_Couleur_PAR_VALUE, dbo.RQ_CONDITIONNEMENT.PAR_CODE, 
                         dbo.CONTENANT.CONT_LIBELLE, dbo.RQ_Region.PAR_VALUE AS RQ_Region_PAR_VALUE, dbo.FOURNISSEUR.FRN_NOM, dbo.RQ_QTECMD_PRD.PRD_QTE_COMMANDE, dbo.RQ_Tva.PAR_VALUE AS RQ_Tva_PAR_VALUE, 
                         dbo.PRODUIT.PRD_STOCK, dbo.PRODUIT.PRD_DOSSIER
FROM            dbo.FOURNISSEUR INNER JOIN
                         dbo.CONTENANT INNER JOIN
                         dbo.PRODUIT INNER JOIN
                         dbo.RQ_Couleur ON dbo.PRODUIT.PRD_COUL_ID = dbo.RQ_Couleur.PAR_ID ON dbo.CONTENANT.CONT_ID = dbo.PRODUIT.PRD_CONT_ID INNER JOIN
                         dbo.RQ_Region ON dbo.PRODUIT.PRD_RGN_ID = dbo.RQ_Region.PAR_ID INNER JOIN
                         dbo.RQ_Tva ON dbo.PRODUIT.PRD_TVA_ID = dbo.RQ_Tva.PAR_ID ON dbo.FOURNISSEUR.FRN_ID = dbo.PRODUIT.PRD_FRN_ID INNER JOIN
                         dbo.RQ_CONDITIONNEMENT ON dbo.PRODUIT.PRD_COND_ID = dbo.RQ_CONDITIONNEMENT.PAR_ID LEFT OUTER JOIN
                         dbo.RQ_QTECMD_PRD ON dbo.PRODUIT.PRD_ID = dbo.RQ_QTECMD_PRD.LGCM_PRD_ID

GO

-- UPDATE SOUSCOMMANDE HOBIVIN 2017-2018
UPDATE SOUSCOMMANDE 
SET SCMD_COM_TAUX = 12, SCMD_COM_MONTANT = SCMD_COM_BASE * (0.12)
WHERE SCMD_ID IN (
SELECT         SCMD_ID
FROM            SOUSCOMMANDE INNER JOIN
                         FOURNISSEUR ON SOUSCOMMANDE.SCMD_FRN_ID = FOURNISSEUR.FRN_ID
WHERE        (FOURNISSEUR.FRN_CODE = N'003517') AND YEAR(SCMD_DATE)>=2017
)
GO

-- CREATION D'une Vue intermédiaires
CREATE VIEW [dbo].[RQ_INTERMEDIAIRES]
AS
SELECT        dbo.CLIENT.*, dbo.PARAMETRE.PAR_CODE
FROM            dbo.CLIENT INNER JOIN
                         dbo.PARAMETRE ON dbo.CLIENT.CLT_TYPE_ID = dbo.PARAMETRE.PAR_ID
WHERE        (dbo.PARAMETRE.PAR_CODE = N'INT')

GO

-- Mise à jour des Mvt de stocks HOBIVIN
UPDATE         MVT_STOCK
SET STK_LIB = concat('CMD ' , commande.CMD_CODE , ' - SARL HOBIVIN')
FROM            MVT_STOCK INNER JOIN
                         COMMANDE ON MVT_STOCK.STK_REF_ID = COMMANDE.CMD_ID
WHERE        COMMANDE.CMD_ORIGINE = 'HOBIVIN'
GO

--Numéro de version
UPDATE CONSTANTES SET CST_VERSION_BD='5.9.3';