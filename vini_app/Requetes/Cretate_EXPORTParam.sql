DELETE From [dbo].[EXPORTPARAM] where Exp_type = 'COM' and EXP_NLIGNE = 4
GO

INSERT INTO [dbo].[EXPORTPARAM]
           ([EXP_TYPE]
           ,[EXP_NLIGNE]
           ,[EXP_NCHAMPS]
           ,[EXP_TYPECHAMPS]
           ,[EXP_VALEUR]
           ,[EXP_LONGUEUR]
           ,[EXP_DESCRIPTION])
     VALUES
          ('COM',4,1,'C','X',1,'Type=X'),
		   ('COM',4,2,'V','CODECOMPTA',0,'Numéro de compte'),
		   ('COM',4,3,'C',';',0,''),
		   ('COM',4,4,'C','ModePaiement=',0,'Champ Mode de reglement sur 4 car'),
		   ('COM',4,5,'V','MODEREGLEMENT4',0,'Valeur Mode de reglement sur 4 car'),
		   ('COM',4,6,'C',';',0,''),
		   ('COM',4,7,'C','DomBanque=',0,'champ banque'),
		   ('COM',4,8,'V','BANQUE',0,'Valeur banque'),
		   ('COM',4,9,'C',';',0,''),
		   ('COM',4,10,'C','Rib=',0,'Champ RIB'),
		   ('COM',4,11,'V','RIB',0,'Valeur RIB')
GO

--INSERT INTO [dbo].[EXPORTPARAM]
--           ([EXP_TYPE]
--           ,[EXP_NLIGNE]
--           ,[EXP_NCHAMPS]
--           ,[EXP_TYPECHAMPS]
--           ,[EXP_VALEUR]
--           ,[EXP_LONGUEUR]
--          ,[EXP_DESCRIPTION])
--     VALUES
--           ('COM',4,1,'C','C',1,'Type=C'),
---		   ('COM',4,2,'V','CODECOMPTA',8,'Numéro de compte'),
--		   ('COM',4,3,'V','LIBELLE',30,'Libellé'),
---		   ('COM',4,4,'C','',7,'Clé alpha'),
---		   ('COM',4,5,'C','',13,'Montant Débit N-1'),
--		   ('COM',4,6,'C','',13,'Montant crédit N-1'),
--		   ('COM',4,5,'C','',13,'Montant Débit N-2'),
--		   ('COM',4,6,'C','',13,'Montant crédit N-2'),
--		   ('COM',4,7,'C','',8,'Compte Collectif'),
--		   ('COM',4,8,'C','',30,'Adresse rue1'),
--		   ('COM',4,8,'C','',30,'Adresse rue2'),
--		   ('COM',4,8,'C','',30,'Adresse Ville'),
--		   ('COM',4,9,'C','',20,'Téléphone'),
--		   ('COM',4,10,'C','',1,'Flag'),
--		   ('COM',4,11,'C','C',1,'Type de Compte'),
--		   ('COM',4,11,'C','N',1,'Compte à centraliser'),
--		   ('COM',4,12,'V','BANQUE',30,'Dommicialiation bancaire'),
--		   ('COM',4,13,'V','RIB',30,'RIB'),
--		   ('COM',4,14,'V','MODEREGLEMENT2',2,'Mode de reglement'),
--		   ('COM',4,14,'C','',2,'Nombre de jours d échéance'),
--		   ('COM',4,14,'C','',2,'Terme échéance'),
--		   ('COM',4,14,'C','',2,'Départ calcul échéance'),
--		   ('COM',4,14,'C','',2,'Code TVA'),
--		   ('COM',4,14,'C','',8,'Compte de contrepartie'),
--		   ('COM',4,14,'C','',3,'Nombre de jours d échéance'),
--		   ('COM',4,14,'C','',1,'Flag TVA'),
--		   ('COM',4,14,'C','',20,'Fax'),
--		   ('COM',4,14,'V','MODEREGLEMENT4',4,'Mode de Reglement 4'),
--		   ('COM',4,14,'C','',8,'Goupe 4'),
--		   ('COM',4,14,'C','',14,'SIRET'),
--		   ('COM',4,14,'C','',1,'Edit M2'),
--		   ('COM',4,14,'C','',30,'Profession'),
--		   ('COM',4,14,'C','',50,'Pays'),
--		   ('COM',4,14,'C','',3,'Code journal d trésorerie'),
--		   ('COM',4,14,'C','',1,'Personne morale'),
--		   ('COM',4,14,'C','',1,'Bon à payer'),
--		   ('COM',4,14,'C','',4,'IBAN'),
--		   ('COM',4,14,'C','',11,'BIC'),
--		   ('COM',4,14,'C','',2,'Code imputation Frais'),
--		   ('COM',4,14,'C','',3,'Numéro du mandat SEPA')
-- GO

