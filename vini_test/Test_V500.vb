'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NUnit.Extensions.Forms
Imports vini_DB
Imports vini_App



<TestClass()> Public Class test_V500
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objFRN2 As Fournisseur
    Private m_objCLT As Client
    Private m_objCLT2 As Client
    Private m_objProduit As Produit
    <TestInitialize()> Public Overrides Sub TestInitialize()

        MyBase.TestInitialize()

        m_objFRN = New Fournisseur("F01", "Nom du fournisseur")
        m_objFRN.rs = "FRNV500"
        m_objFRN.AdresseLivraison.nom = "ADF_Nom"
        m_objFRN.AdresseLivraison.rue1 = "ADF_Rue1"
        m_objFRN.AdresseLivraison.rue2 = "ADF_Rue2"
        m_objFRN.AdresseLivraison.cp = "ADF_cp"
        m_objFRN.AdresseLivraison.ville = "ADF_Ville"
        m_objFRN.AdresseLivraison.tel = "01010101"
        m_objFRN.AdresseLivraison.fax = "02020202"
        m_objFRN.AdresseLivraison.port = "03030303"
        m_objFRN.AdresseLivraison.Email = "04040404"
        m_objFRN.AdresseFacturation.nom = "ADF_Nom"
        m_objFRN.AdresseFacturation.rue1 = "ADF_Rue1"
        m_objFRN.AdresseFacturation.rue2 = "ADF_Rue2"
        m_objFRN.AdresseFacturation.cp = "ADF_cp"
        m_objFRN.AdresseFacturation.ville = "ADF_Ville"
        m_objFRN.AdresseFacturation.tel = "01010101"
        m_objFRN.AdresseFacturation.fax = "02020202"
        m_objFRN.AdresseFacturation.port = "03030303"
        m_objFRN.AdresseFacturation.Email = "04040404"
        Assert.IsTrue(m_objFRN.Save(), "FRN.Create")

        m_objFRN2 = New Fournisseur("F02", "Nom du fournisseur-2")
        m_objFRN2.rs = "FRNV500-2"
        m_objFRN2.AdresseLivraison.nom = "ADF_Nom-2"
        m_objFRN2.AdresseLivraison.rue1 = "ADF_Rue1-2"
        m_objFRN2.AdresseLivraison.rue2 = "ADF_Rue2-2"
        m_objFRN2.AdresseLivraison.cp = "ADF_cp-2"
        m_objFRN2.AdresseLivraison.ville = "ADF_Ville-2"
        m_objFRN2.AdresseLivraison.tel = "01010101-2"
        m_objFRN2.AdresseLivraison.fax = "02020202-2"
        m_objFRN2.AdresseLivraison.port = "03030303-2"
        m_objFRN2.AdresseLivraison.Email = "04040404-1"
        m_objFRN2.AdresseFacturation.nom = "ADF_Nom-2"
        m_objFRN2.AdresseFacturation.rue1 = "ADF_Rue1-2"
        m_objFRN2.AdresseFacturation.rue2 = "ADF_Rue2-2"
        m_objFRN2.AdresseFacturation.cp = "ADF_cp-2"
        m_objFRN2.AdresseFacturation.ville = "ADF_Ville-2"
        m_objFRN2.AdresseFacturation.tel = "01010101-2"
        m_objFRN2.AdresseFacturation.fax = "02020202-2"
        m_objFRN2.AdresseFacturation.port = "03030303-2"
        m_objFRN2.AdresseFacturation.Email = "04040404-1"
        Assert.IsTrue(m_objFRN2.Save(), "FRN2.Create")

        m_objCLT = New Client("CLTV500", "NOMClient de' test")
        m_objCLT.rs = "RSClient de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")

        m_objCLT2 = New Client("CLTV500-2", "NOMClient de' test 2")
        m_objCLT2.rs = "RSClient de test 2"
        Assert.IsTrue(m_objCLT2.save(), "Client.Create")

        m_objPRD = New Produit("PRDT500", m_objFRN, 2011)
        Assert.IsTrue(m_objPRD.save(), "Produit.Create")
        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()


        MyBase.TestCleanup()
    End Sub

    ''' <summary>
    ''' Nom et Raison Sociale du client/ Fournisseur duppliqué dans les tables 
    ''' commandes et Bon Appo
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T01_NOMRSCLIENT_CMDCLT()

        Dim oCmd As CommandeClient

        oCmd = New CommandeClient(m_objCLT)
        oCmd.DuppliqueCaracteristiqueTiers()
        'Vérification de la concordance
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom différents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.rs, "Nom différents")

        'Vérification sur la modif n'est pas reportée
        oCmd.caracteristiqueTiers.nom = "NOM1"
        oCmd.caracteristiqueTiers.rs = "RS1"
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, "NOM1", "Nom différents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, "RS1", "RS différents")
        Assert.AreNotEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom égaux")
        Assert.AreNotEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.nom, "RS égaux")

        'Sauvegarde de la commande
        Assert.IsTrue(oCmd.save(), "SV Commande")
        Dim nCmd As Integer
        nCmd = oCmd.id

        oCmd = Nothing
        oCmd = CommandeClient.createandload(nCmd)
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, "NOM1", "Nom différents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, "RS1", "RS différents")
        Assert.AreNotEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom égaux")
        Assert.AreNotEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.nom, "RS égaux")

        'Vérification de l'Update
        oCmd.caracteristiqueTiers.nom = "NOM2"
        oCmd.caracteristiqueTiers.rs = "RS2"
        Assert.IsTrue(oCmd.save(), "SV Commande")
        nCmd = oCmd.id

        oCmd = Nothing
        oCmd = CommandeClient.createandload(nCmd)
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, "NOM2", "Nom différents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, "RS2", "RS différents")





    End Sub
    ''' <summary>
    ''' Teste la recopie des infos du tiers sur la commande
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T02_COPY_TIERS_TO_CMD()

        Dim oCmd As CommandeClient

        oCmd = New CommandeClient(m_objCLT)
        oCmd.DuppliqueCaracteristiqueTiers()
        'Modification de la commande
        oCmd.caracteristiqueTiers.nom = "NOM1"
        oCmd.caracteristiqueTiers.rs = "RS1"
        'Sauvegarde de la commande
        Assert.IsTrue(oCmd.save(), "SV Commande")
        Dim nCmd As Integer
        nCmd = oCmd.id
        oCmd = CommandeClient.createandload(nCmd)
        'Vérification de la valeur sauvegardée et relues
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, "NOM1", "Nom différents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, "RS1", "RS différents")

        Assert.IsTrue(oCmd.DuppliqueCaracteristiqueTiers())
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom différents aprés recopie")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.rs, "Nom différents après recopie")

        'Changement de tiers
        oCmd.oTiers = m_objCLT2
        'Vérification que les caractéristiques du tiers n'ont pas été recopiées    
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom modifié par la reaffection du tiers")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.rs, "RS modifié par la reaffection du tiers")

        'Changement de tiers après réinitialisation
        oCmd.oTiers = Nothing
        oCmd.oTiers = m_objCLT2
        'Vérification que les caractéristiques du tiers ont été recopiées    
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT2.nom, "Nom non modifié par la reaffection du tiers")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT2.rs, "RS non modifié par la reaffection du tiers")

    End Sub

    <TestMethod()> Public Sub T10_NOMRSFRN_BA()

        Dim oBA As BonAppro

        oBA = New BonAppro(m_objFRN)
        oBA.DuppliqueCaracteristiqueTiers()
        'Vérification de la concordance
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom différents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, m_objFRN.rs, "Nom différents")

        'Vérification sur la modif n'est pas reportée
        oBA.caracteristiqueTiers.nom = "NOM1"
        oBA.caracteristiqueTiers.rs = "RS1"
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, "NOM1", "Nom différents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, "RS1", "RS différents")
        Assert.AreNotEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom égaux")
        Assert.AreNotEqual(oBA.caracteristiqueTiers.rs, m_objFRN.nom, "RS égaux")

        'Sauvegarde de la commande
        Assert.IsTrue(oBA.save(), "SV Commande")
        Dim nCmd As Integer
        nCmd = oBA.id

        oBA = Nothing
        oBA = BonAppro.createandload(nCmd)
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, "NOM1", "Nom différents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, "RS1", "RS différents")
        Assert.AreNotEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom égaux")
        Assert.AreNotEqual(oBA.caracteristiqueTiers.rs, m_objFRN.nom, "RS égaux")

        'Vérification de l'Update
        oBA.caracteristiqueTiers.nom = "NOM2"
        oBA.caracteristiqueTiers.rs = "RS2"
        Assert.IsTrue(oBA.save(), "SV Commande")
        nCmd = oBA.id

        oBA = Nothing
        oBA = BonAppro.createandload(nCmd)
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, "NOM2", "Nom différents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, "RS2", "RS différents")

        oBA = New BonAppro(m_objFRN)
        oBA.DuppliqueCaracteristiqueTiers()
        'Modification de la commande
        oBA.caracteristiqueTiers.nom = "NOM1"
        oBA.caracteristiqueTiers.rs = "RS1"
        'Sauvegarde de la commande
        Assert.IsTrue(oBA.save(), "SV Commande")
        nCmd = oBA.id
        oBA = BonAppro.createandload(nCmd)
        'Vérification de la valeur sauvegardée et relues
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, "NOM1", "Nom différents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, "RS1", "RS différents")

        Assert.IsTrue(oBA.DuppliqueCaracteristiqueTiers())
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom différents aprés recopie")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, m_objFRN.rs, "Nom différents après recopie")

        'Changement de tiers
        oBA.oTiers = m_objFRN2
        'Vérification que les caractéristiques du tiers n'ont pas été recopiées    
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom modifié par la reaffection du tiers")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, m_objFRN.rs, "RS modifié par la reaffection du tiers")

        'Changement de tiers après réinitialisation
        oBA.oTiers = Nothing
        oBA.oTiers = m_objFRN2
        'Vérification que les caractéristiques du tiers ont été recopiées    
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, m_objFRN2.nom, "Nom non modifié par la reaffection du tiers")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, m_objFRN2.rs, "RS non modifié par la reaffection du tiers")


    End Sub

    <TestMethod()> Public Sub T20_SetProperty()
        Dim oCMD As CommandeClient
        Dim oLG As LgCommande

        oCMD = New CommandeClient(m_objCLT)
        oCMD.AjouteLigne("10", m_objPRD, 0, 0)

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_CODE, "35510"))
        Assert.AreEqual(oCMD.code, "35510")
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_CODE), "35510")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_CoutTransport, "355.10"))
        Assert.AreEqual(oCMD.coutTransport, CDec(355.1))
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_CoutTransport), "355.10")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_dateCommande, "25021964"))
        Assert.AreEqual(oCMD.dateCommande, CDate("25/02/64"))
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_dateCommande), "25021964")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_dateEnlevement, "25031964"))
        Assert.AreEqual(oCMD.dateEnlevement, CDate("25/03/64"))
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_dateEnlevement), "25031964")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_dateLivraison, "25041964"))
        Assert.AreEqual(oCMD.dateLivraison, CDate("25/04/64"))
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_dateLivraison), "25041964")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_dateValidation, "26031964"))
        Assert.AreEqual(oCMD.dateValidation, CDate("26/03/64"))
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_dateValidation), "26031964")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_etat, vncEnums.vncEtatCommande.vncEclatee))
        Assert.AreEqual(oCMD.etat.codeEtat, vncEnums.vncEtatCommande.vncEclatee)
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_etat), "3")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_LettreVoiture, "AZERTTYUI"))
        Assert.AreEqual(oCMD.lettreVoiture, "AZERTTYUI")
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_LettreVoiture), "AZERTTYUI")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_MontantTransport, "150.25"))
        Assert.AreEqual(oCMD.montantTransport, CDec("150.25"))
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_MontantTransport), "150.25")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_poids, "151.26"))
        Assert.AreEqual(oCMD.poids, CDec("151.26"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_puPalettesNonPreparees, "152.26"))
        Assert.AreEqual(oCMD.puPalettesNonPreparees, CDec("152.26"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_puPalettesPreparees, "153.26"))
        Assert.AreEqual(oCMD.puPalettesPreparees, CDec("153.26"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_qteColis, "152"))
        Assert.AreEqual(oCMD.qteColis, CDec("152"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_qtePalettesNonPreparees, "252.26"))
        Assert.AreEqual(oCMD.qtePalettesNonPreparees, CDec("252.26"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_qtePalettesPreparees, "253.26"))
        Assert.AreEqual(oCMD.qtePalettesPreparees, CDec("253.26"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_RefFactTrp, "REFFRACTTRP"))
        Assert.AreEqual(oCMD.refFactTRP, "REFFRACTTRP")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_RefLivraison, "REFLIV"))
        Assert.AreEqual(oCMD.refLivraison, "REFLIV")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADFACT1, "ADRFACT1"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseFacturation.rue1, "ADRFACT1")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADFACT2, "ADRFACT2"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseFacturation.rue2, "ADRFACT2")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADFACTCP, "ADRFACTCP"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseFacturation.cp, "ADRFACTCP")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADFACTVILLE, "ADRFACTVILLE"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseFacturation.ville, "ADRFACTVILLE")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADFACTPORT, "0680667189"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseFacturation.port, "0680667189")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADFACTTEL, "0299556339"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseFacturation.tel, "0299556339")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADFACTFAX, "0299556338"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseFacturation.fax, "0299556338")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADFACTEMAIL, "ADRFACTEMAIL"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseFacturation.Email, "ADRFACTEMAIL")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIV1, "ADRLIV1"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.rue1, "ADRLIV1")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIV2, "ADRLIV2"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.rue2, "ADRLIV2")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIVCP, "ADRLIVCP"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.cp, "ADRLIVCP")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIVVILLE, "ADRLIVVILLE"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.ville, "ADRLIVVILLE")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIVPORT, "0680667190"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.port, "0680667190")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIVTEL, "0299556377"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.tel, "0299556377")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIVFAX, "0299556388"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.fax, "0299556388")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIVEMAIL, "ADRLIVEMAIL"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.Email, "ADRLIVEMAIL")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_ADLIVEMAIL, "ADRLIVEMAIL"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.AdresseLivraison.Email, "ADRLIVEMAIL")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_NOM, "tiersnom"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.nom, "tiersnom")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TIERS_RS, "tiersRS"))
        Assert.AreEqual(oCMD.caracteristiqueTiers.rs, "tiersRS")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TotalHT, "546.23"))
        Assert.AreEqual(oCMD.totalHT, CDec(546.23))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TotalTTC, "548.23"))
        Assert.AreEqual(oCMD.totalTTC, CDec(548.23))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TransporteurCODE, "RCT"))
        Assert.AreEqual(oCMD.oTransporteur.code, "RCT")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_TiersCODE, "CLTV500"))
        Assert.AreEqual(oCMD.oTiers.code, "CLTV500")
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_num, "10"))
        Assert.AreEqual(oCMD.colLignes.Count, 1)
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_bGratuit, "TRUE"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.num, 10)
        Assert.AreEqual(oLG.bGratuit, True)
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_MtComm, "123.52"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.MtComm, CDec("123.52"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_poids, "124.52"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.poids, CDec("124.52"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_prixHT, "125.52"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.prixHT, CDec("125.52"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_prixTTC, "126.52"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.prixTTC, CDec("126.52"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_prixU, "127.52"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.prixU, CDec("127.52"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_qteColis, "321.65"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.qteColis, CDec("321.65"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_qteCom, "322.65"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.qteCommande, CDec("322.65"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_qteLiv, "323.65"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.qteLiv, CDec("323.65"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_qteFact, "324.65"))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.qteFact, CDec("324.65"))
        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.LGCMD_PRD_CODE, m_objPRD.code))
        oLG = oCMD.colLignes(1)
        Assert.AreEqual(oLG.oProduit.code, m_objPRD.code)

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_ComentCom, "Commentaire de commande"))
        Assert.AreEqual(oCMD.CommentaireCommandeText, "Commentaire de commande")
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_ComentCom), "Commentaire de commande")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_ComentLiv, "Commentaire de Livraison"))
        Assert.AreEqual(oCMD.CommentaireLivraisonText, "Commentaire de Livraison")
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_ComentLiv), "Commentaire de Livraison")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_ComentFact, "Commentaire de Facturation"))
        Assert.AreEqual(oCMD.CommentaireFacturationText, "Commentaire de Facturation")
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_ComentFact), "Commentaire de Facturation")

        Assert.IsTrue(oCMD.SetProperty(vncEnums.vncObjectProperties.CMD_ComentLibre, "Commentaire Libre"))
        Assert.AreEqual(oCMD.CommentaireLibreText, "Commentaire Libre")
        Assert.AreEqual(oCMD.getProperty(vncEnums.vncObjectProperties.CMD_ComentLibre), "Commentaire Libre")
    End Sub

    <TestMethod()> Public Sub T30_ImportCmd()
        Dim oCmd As CommandeClient
        If My.Computer.FileSystem.FileExists("./T30.txt") Then
            My.Computer.FileSystem.DeleteFile("./T30.txt")
        End If
        My.Computer.FileSystem.WriteAllText("./T30.txt", m_objCLT.code + ";31012005;" + m_objPRD.code + ";2;15.60;commentaire" + vbCrLf, True)
        My.Computer.FileSystem.WriteAllText("./T30.txt", ";;" + m_objPRD.code + ";3;29.70;" + vbCrLf, True)
        oCmd = CommandeClient.importCommandeClient("./T30.txt")
        Assert.IsNotNull(oCmd)
        Assert.AreEqual(oCmd.TiersCode, m_objCLT.code)
        Assert.AreEqual(oCmd.dateCommande, CDate("31/01/2005"))
        Assert.AreEqual(oCmd.CommentaireCommandeText, "commentaire")
        Assert.AreEqual(oCmd.colLignes.Count, 2)
        Dim oLg As LgCommande
        oLg = oCmd.colLignes(1)
        Assert.AreEqual(oLg.oProduit.code, m_objPRD.code)
        Assert.AreEqual(oLg.qteCommande, 2D)
        Assert.AreEqual(oLg.prixU, CDec(15.6))
        oLg = oCmd.colLignes(2)
        Assert.AreEqual(oLg.oProduit.code, m_objPRD.code)
        Assert.AreEqual(oLg.qteCommande, 3D)
        Assert.AreEqual(oLg.prixU, CDec(29.7))
        Assert.IsTrue(oCmd.save())
        Trace.WriteLine("CMDCODE=" + oCmd.code)
    End Sub
End Class


