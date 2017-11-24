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
    ''' Nom et Raison Sociale du client/ Fournisseur duppliqu� dans les tables 
    ''' commandes et Bon Appo
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T01_NOMRSCLIENT_CMDCLT()

        Dim oCmd As CommandeClient

        oCmd = New CommandeClient(m_objCLT)
        oCmd.DuppliqueCaracteristiqueTiers()
        'V�rification de la concordance
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom diff�rents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.rs, "Nom diff�rents")

        'V�rification sur la modif n'est pas report�e
        oCmd.caracteristiqueTiers.nom = "NOM1"
        oCmd.caracteristiqueTiers.rs = "RS1"
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, "NOM1", "Nom diff�rents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, "RS1", "RS diff�rents")
        Assert.AreNotEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom �gaux")
        Assert.AreNotEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.nom, "RS �gaux")

        'Sauvegarde de la commande
        Assert.IsTrue(oCmd.save(), "SV Commande")
        Dim nCmd As Integer
        nCmd = oCmd.id

        oCmd = Nothing
        oCmd = CommandeClient.createandload(nCmd)
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, "NOM1", "Nom diff�rents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, "RS1", "RS diff�rents")
        Assert.AreNotEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom �gaux")
        Assert.AreNotEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.nom, "RS �gaux")

        'V�rification de l'Update
        oCmd.caracteristiqueTiers.nom = "NOM2"
        oCmd.caracteristiqueTiers.rs = "RS2"
        Assert.IsTrue(oCmd.save(), "SV Commande")
        nCmd = oCmd.id

        oCmd = Nothing
        oCmd = CommandeClient.createandload(nCmd)
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, "NOM2", "Nom diff�rents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, "RS2", "RS diff�rents")





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
        'V�rification de la valeur sauvegard�e et relues
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, "NOM1", "Nom diff�rents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, "RS1", "RS diff�rents")

        Assert.IsTrue(oCmd.DuppliqueCaracteristiqueTiers())
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom diff�rents apr�s recopie")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.rs, "Nom diff�rents apr�s recopie")

        'Changement de tiers
        oCmd.setTiers(m_objCLT2)
        'V�rification que les caract�ristiques du tiers n'ont pas �t� recopi�es    
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom modifi� par la reaffection du tiers")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.rs, "RS modifi� par la reaffection du tiers")

        'Changement de tiers apr�s r�initialisation
        oCmd.setTiers(Nothing)
        oCmd.setTiers(m_objCLT2)
        'V�rification que les caract�ristiques du tiers ont �t� recopi�es    
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT2.nom, "Nom non modifi� par la reaffection du tiers")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT2.rs, "RS non modifi� par la reaffection du tiers")

    End Sub

    <TestMethod()> Public Sub T10_NOMRSFRN_BA()

        Dim oBA As BonAppro

        oBA = New BonAppro(m_objFRN)
        oBA.DuppliqueCaracteristiqueTiers()
        'V�rification de la concordance
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom diff�rents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, m_objFRN.rs, "Nom diff�rents")

        'V�rification sur la modif n'est pas report�e
        oBA.caracteristiqueTiers.nom = "NOM1"
        oBA.caracteristiqueTiers.rs = "RS1"
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, "NOM1", "Nom diff�rents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, "RS1", "RS diff�rents")
        Assert.AreNotEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom �gaux")
        Assert.AreNotEqual(oBA.caracteristiqueTiers.rs, m_objFRN.nom, "RS �gaux")

        'Sauvegarde de la commande
        Assert.IsTrue(oBA.save(), "SV Commande")
        Dim nCmd As Integer
        nCmd = oBA.id

        oBA = Nothing
        oBA = BonAppro.createandload(nCmd)
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, "NOM1", "Nom diff�rents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, "RS1", "RS diff�rents")
        Assert.AreNotEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom �gaux")
        Assert.AreNotEqual(oBA.caracteristiqueTiers.rs, m_objFRN.nom, "RS �gaux")

        'V�rification de l'Update
        oBA.caracteristiqueTiers.nom = "NOM2"
        oBA.caracteristiqueTiers.rs = "RS2"
        Assert.IsTrue(oBA.save(), "SV Commande")
        nCmd = oBA.id

        oBA = Nothing
        oBA = BonAppro.createandload(nCmd)
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, "NOM2", "Nom diff�rents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, "RS2", "RS diff�rents")

        oBA = New BonAppro(m_objFRN)
        oBA.DuppliqueCaracteristiqueTiers()
        'Modification de la commande
        oBA.caracteristiqueTiers.nom = "NOM1"
        oBA.caracteristiqueTiers.rs = "RS1"
        'Sauvegarde de la commande
        Assert.IsTrue(oBA.save(), "SV Commande")
        nCmd = oBA.id
        oBA = BonAppro.createandload(nCmd)
        'V�rification de la valeur sauvegard�e et relues
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, "NOM1", "Nom diff�rents")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, "RS1", "RS diff�rents")

        Assert.IsTrue(oBA.DuppliqueCaracteristiqueTiers())
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom diff�rents apr�s recopie")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, m_objFRN.rs, "Nom diff�rents apr�s recopie")

        'Changement de tiers
        oBA.setTiers(m_objFRN2)
        'V�rification que les caract�ristiques du tiers n'ont pas �t� recopi�es    
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, m_objFRN.nom, "Nom modifi� par la reaffection du tiers")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, m_objFRN.rs, "RS modifi� par la reaffection du tiers")

        'Changement de tiers apr�s r�initialisation
        oBA.setTiers(Nothing)
        oBA.setTiers(m_objFRN2)
        'V�rification que les caract�ristiques du tiers ont �t� recopi�es    
        Assert.AreEqual(oBA.caracteristiqueTiers.nom, m_objFRN2.nom, "Nom non modifi� par la reaffection du tiers")
        Assert.AreEqual(oBA.caracteristiqueTiers.rs, m_objFRN2.rs, "RS non modifi� par la reaffection du tiers")


    End Sub


End Class


