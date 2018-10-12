
/* Pour éviter les problèmes éventuels de perte de données, passez attentivement ce script en revue avant de l'exécuter en dehors du contexte du Concepteur de bases de données.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.LGFACTCOLISAGE ADD
	LGCOL_PRD_ID int NULL
GO
-- Mettre le champs LGCOL_FACTCOL_ID en int
COMMIT

--Numéro de version
UPDATE CONSTANTES SET CST_VERSION_BD='5.9.7';