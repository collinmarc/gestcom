'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.IO



<TestClass()> Public Class T1110_sscommande
    Inherits test_Base
    Private m_oProduit As Produit
    Private m_oProduit2 As Produit
    Private m_oFourn As Fournisseur
    Private m_oFourn2 As Fournisseur
    Private m_oClient As Client
    Private m_oCmd As CommandeClient
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Dim col As Collection
        Dim oTaux As TauxComm
        Dim strCode As String

        strCode = "FRNT1110"
        m_oFourn = New Fournisseur(strCode, "MonFournisseur")
        m_oFourn.AdresseFacturation.nom = "ADF_Nom"
        m_oFourn.AdresseFacturation.rue1 = "ADF_Nom"
        m_oFourn.AdresseFacturation.rue2 = "ADF_Nom"
        m_oFourn.AdresseFacturation.cp = "ADF_Nom"
        m_oFourn.AdresseFacturation.ville = "ADF_Nom"
        m_oFourn.AdresseFacturation.tel = "01010101"
        m_oFourn.AdresseFacturation.fax = "02020202"
        m_oFourn.AdresseFacturation.port = "03030303"
        m_oFourn.AdresseFacturation.Email = "04040404"
        m_oFourn.Save()

        strCode = "FRN2T1110"
        col = Fournisseur.getListe(strCode)
        For Each m_oFourn2 In col
            m_oFourn2.bDeleted = True
            m_oFourn2.Save()
        Next
        m_oFourn2 = New Fournisseur(strCode, "MonFournisseur")
        m_oFourn2.Save()

        strCode = "TSTPRDT1110"
        col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, strCode)
        For Each m_oProduit In col
            m_oProduit.bDeleted = True
            m_oProduit.save()
        Next
        m_oProduit = New Produit(strCode, m_oFourn, 1990)
        m_oProduit.save()

        strCode = "TSTPRD2T1110"
        col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, strCode)
        For Each m_oProduit2 In col
            m_oProduit2.bDeleted = True
            m_oProduit2.save()
        Next
        m_oProduit2 = New Produit(strCode, m_oFourn2, 1990)
        m_oProduit2.save()

        strCode = "CLTT1110"
        col = Client.getListe(strCode)
        For Each m_oClient In col
            m_oClient.bDeleted = True
            m_oClient.save()
        Next
        m_oClient = New Client(strCode, "MonClient")
        Debug.Assert(m_oClient.save(), "Creation du client")
        '            m_oClient()

        'Creation des Taux de Commissions
        oTaux = New TauxComm(m_oFourn, m_oClient.codeTypeClient, 9.5)
        oTaux.save()
        oTaux = New TauxComm(m_oFourn2, m_oClient.codeTypeClient, 9.5)
        oTaux.save()

        'Suppression des Commandes de 1964
        col = CommandeClient.getListe(#6/2/1964#, #6/2/1964#)
        For Each m_oCmd In col
            m_oCmd.bDeleted = True
            m_oCmd.save()
        Next
        m_oCmd = New CommandeClient(m_oClient)
        m_oCmd.dateCommande = #6/2/1964#
        m_oCmd.save()



    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        MyBase.TestCleanup()

    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objSCMD2 As SousCommande
        Dim objSCMD22 As SousCommande



        objSCMD2 = New SousCommande(m_oCmd, m_oFourn)

        Assert.IsTrue(objSCMD2.Selected)


        objSCMD2.code = "CODE"
        objSCMD2.dateCommande = CDate("06/02/1964")
        objSCMD2.dateLivraison = CDate("06/02/1964")
        objSCMD2.dateEnlevement = CDate("31/07/1964")
        objSCMD2.refLivraison = "BL0003"
        objSCMD2.oTransporteur.AdresseLivraison.nom = "TRANSPORT RAULT"
        objSCMD2.idCommandeClient = m_oCmd.id
        objSCMD2.refFactFournisseur = "FACT1"
        objSCMD2.dateFactFournisseur = "01/08/2004"
        objSCMD2.totalHTFacture = 150
        objSCMD2.totalTTCFacture = 165
        objSCMD2.baseCommission = 155
        objSCMD2.tauxCommission = 12

        Assert.AreEqual(objSCMD2.code, "CODE")
        Assert.AreEqual(objSCMD2.dateCommande, CDate("06/02/1964"))
        Assert.AreEqual(objSCMD2.dateLivraison, CDate("06/02/1964"))
        Assert.AreEqual(objSCMD2.dateEnlevement, CDate("31/07/1964"))
        Assert.AreEqual(objSCMD2.refLivraison, "BL0003")
        Assert.AreEqual(objSCMD2.oTransporteur.AdresseLivraison.nom, "TRANSPORT RAULT")
        Assert.AreEqual(objSCMD2.idCommandeClient, m_oCmd.id)
        Assert.AreEqual(objSCMD2.refFactFournisseur, "FACT1")
        Assert.AreEqual(objSCMD2.dateFactFournisseur, CDate("01/08/2004"))
        Assert.AreEqual(objSCMD2.totalHTFacture, CDec(150))
        Assert.AreEqual(objSCMD2.totalTTCFacture, CDec(165))
        Assert.AreEqual(objSCMD2.baseCommission, CDec(155))
        Assert.AreEqual(objSCMD2.tauxCommission, CDec(12))


        'Test des indicateurs
        Assert.IsTrue(objSCMD2.bNew)
        Assert.IsTrue(objSCMD2.bUpdated)
        Assert.IsFalse(objSCMD2.bDeleted)

        Assert.IsTrue(objSCMD2.Equals(objSCMD2), "Egal à Lui même")
        objSCMD22 = New SousCommande(m_oCmd, m_oFourn)

        objSCMD22.code = "CODE"
        objSCMD22.dateCommande = CDate("06/02/1964")
        objSCMD22.dateLivraison = CDate("06/02/1964")
        objSCMD22.dateEnlevement = CDate("31/07/1964")
        objSCMD22.refLivraison = "BL0003"
        objSCMD22.oTransporteur.AdresseLivraison.nom = "TRANSPORT RAULT"
        objSCMD22.idCommandeClient = m_oCmd.id
        objSCMD22.refFactFournisseur = "FACT1"
        objSCMD22.dateFactFournisseur = "01/08/2004"
        objSCMD22.totalHTFacture = 150
        objSCMD22.totalTTCFacture = 165
        objSCMD22.baseCommission = 155
        objSCMD22.tauxCommission = 12

        Assert.IsTrue(objSCMD2.Equals(objSCMD22), "Egal à un semblable")
        objSCMD2.totalHTFacture = 2
        Assert.IsFalse(objSCMD2.Equals(objSCMD22), "Egal à un Différent")

        Dim obj As Object
        Assert.IsFalse(objSCMD2.Equals(obj), "Egal autrecjhose")


    End Sub
    <TestMethod()> Public Sub T15_DB1()
        Dim objSCMD As SousCommande
        Dim nid As Long

        'I - Création d'une Sous-commande 

        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        objSCMD.code = ""
        objSCMD.dateCommande = CDate("06/02/1964")
        objSCMD.dateLivraison = CDate("06/02/1964")
        objSCMD.refLivraison = "BL0003"
        objSCMD.oTransporteur.AdresseLivraison.nom = "TRANSPORT RAULT"
        objSCMD.idCommandeClient = m_oCmd.id
        objSCMD.refFactFournisseur = "FACT1"
        objSCMD.dateFactFournisseur = "01/08/2004"
        objSCMD.totalHTFacture = 150
        objSCMD.totalTTCFacture = 165
        objSCMD.baseCommission = 155
        objSCMD.tauxCommission = 12
        objSCMD.CommCommande.comment = "CommCommande"
        objSCMD.CommLivraison.comment = "CommLivriaon"
        objSCMD.CommFacturation.comment = "CommFact"
        objSCMD.CommLibre.comment = "Libre"
        'Test des indicateurs Avant le Save
        Assert.IsTrue(objSCMD.bNew)
        Assert.IsTrue(objSCMD.bUpdated)
        Assert.IsFalse(objSCMD.bDeleted)
        'Save
        Assert.IsTrue(objSCMD.Save(), "Insert" & objSCMD.getErreur)
        ' Assert.IsTrue((objSCMD.id <> 0), "Id Apres le Save doit être différent de 0")
        ' 'Test des indicateurs Après le Save
        ' Assert.IsFalse(objSCMD.bNew, "bNew apres insert")
        ' Assert.IsFalse(objSCMD.bUpdated, "bUpdated apres insert")
        ' Assert.IsFalse(objSCMD.bDeleted, "bDeleted apres insert")

        nid = objSCMD.id
        'II - Rechargement d'une Sous - Commande
        '========================================
        'objSCMD2 = New SousCommande(m_oCmd, m_oFourn)
        'Assert.IsTrue(objSCMD2.load(nid), "Load de la Sous- commande Client " & nid & ":" & objSCMD2.getErreur())
        'Assert.IsTrue(objSCMD2.oTiers.load(), "chargement du client")
        'Assert.AreEqual(objSCMD2.refLivraison, "BL0003")
        'Assert.IsTrue(objSCMD.Equals(objSCMD2), "Différents")

        ''III - Modification d'une Sous- commande
        ''=================================
        '' Modification de la commande
        'objSCMD2.refFactFournisseur = "FACT"
        'objSCMD2.refLivraison = "BL0002"

        ''Test des indicateurs Avant le Save
        'Assert.IsFalse(objSCMD2.bNew)
        'Assert.IsTrue(objSCMD2.bUpdated)
        'Assert.IsFalse(objSCMD2.bDeleted)
        ''Save
        'Assert.IsTrue(objSCMD2.Save(), "Update" & objSCMD2.getErreur)
        ''Test des indicateurs Après le Save
        'Assert.IsFalse(objSCMD2.bNew, "bNew apres Update")
        'Assert.IsFalse(objSCMD2.bUpdated, "bUpdated apres Update")
        'Assert.IsFalse(objSCMD2.bDeleted, "bDeleted apres Update")
        ''Rechargement de l'objet
        'nid = objSCMD2.id
        'objSCMD = New SousCommande(m_oCmd, m_oFourn)
        'Assert.IsTrue(objSCMD.load(nid), "Load")
        'Assert.IsTrue(objSCMD.oTiers.load(), "Load De la sous-commande après Update")
        'Assert.IsTrue(objSCMD.Equals(objSCMD2), "Apres Update , Equals")

        ''IV - FAXER LA Sous COMMANDE
        ''==========================
        ''Assert.IsTrue(objSCMD.faxer("01234567"), "Faxer" & objSCMD.getErreur())

        'V - Suppression de la Sous commande
        '=================================
        ' Modification de la commande
        'objSCMD.bDeleted = True
        'Assert.IsTrue(objSCMD.Save(), "Delete" & objSCMD.getErreur())
        ''Rechargement dans un autre objet
        'objSCMD2 = New SousCommande(m_oCmd, m_oFourn)
        'Assert.IsFalse(objSCMD2.load(nid), "Load")
    End Sub
    <TestMethod()> Public Sub T15_DB()
        Dim objSCMD As SousCommande
        Dim objSCMD2 As SousCommande
        Dim nid As Long

        'I - Création d'une Sous-commande 

        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        objSCMD.code = ""
        objSCMD.dateCommande = CDate("06/02/1964")
        objSCMD.dateLivraison = CDate("06/02/1964")
        objSCMD.refLivraison = "BL0003"
        objSCMD.oTransporteur.AdresseLivraison.nom = "TRANSPORT RAULT"
        objSCMD.idCommandeClient = m_oCmd.id
        objSCMD.refFactFournisseur = "FACT1"
        objSCMD.dateFactFournisseur = "01/08/2004"
        objSCMD.totalHTFacture = 150
        objSCMD.totalTTCFacture = 165
        objSCMD.baseCommission = 155
        objSCMD.tauxCommission = 12
        objSCMD.CommCommande.comment = "CommCommande"
        objSCMD.CommLivraison.comment = "CommLivriaon"
        objSCMD.CommFacturation.comment = "CommFact"
        objSCMD.CommLibre.comment = "Libre"
        'Test des indicateurs Avant le Save
        Assert.IsTrue(objSCMD.bNew)
        Assert.IsTrue(objSCMD.bUpdated)
        Assert.IsFalse(objSCMD.bDeleted)
        'Save
        Assert.IsTrue(objSCMD.Save(), "Insert" & objSCMD.getErreur)
        Assert.IsTrue((objSCMD.id <> 0), "Id Apres le Save doit être différent de 0")
        'Test des indicateurs Après le Save
        Assert.IsFalse(objSCMD.bNew, "bNew apres insert")
        Assert.IsFalse(objSCMD.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objSCMD.bDeleted, "bDeleted apres insert")

        nid = objSCMD.id
        'II - Rechargement d'une Sous - Commande
        '========================================
        objSCMD2 = New SousCommande(m_oCmd, m_oFourn)
        Assert.IsTrue(objSCMD2.load(nid), "Load de la Sous- commande Client " & nid & ":" & objSCMD2.getErreur())
        Assert.IsTrue(objSCMD2.oTiers.load(), "chargement du client")
        Assert.AreEqual(objSCMD2.refLivraison, "BL0003")
        Assert.IsTrue(objSCMD.Equals(objSCMD2), "Différents")

        'III - Modification d'une Sous- commande
        '=================================
        ' Modification de la commande
        objSCMD2.refFactFournisseur = "FACT"
        objSCMD2.refLivraison = "BL0002"

        'Test des indicateurs Avant le Save
        Assert.IsFalse(objSCMD2.bNew)
        Assert.IsTrue(objSCMD2.bUpdated)
        Assert.IsFalse(objSCMD2.bDeleted)
        'Save
        Assert.IsTrue(objSCMD2.Save(), "Update" & objSCMD2.getErreur)
        'Test des indicateurs Après le Save
        Assert.IsFalse(objSCMD2.bNew, "bNew apres Update")
        Assert.IsFalse(objSCMD2.bUpdated, "bUpdated apres Update")
        Assert.IsFalse(objSCMD2.bDeleted, "bDeleted apres Update")
        'Rechargement de l'objet
        nid = objSCMD2.id
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.IsTrue(objSCMD.load(nid), "Load")
        Assert.IsTrue(objSCMD.oTiers.load(), "Load De la sous-commande après Update")
        Assert.IsTrue(objSCMD.Equals(objSCMD2), "Apres Update , Equals")

        'IV - FAXER LA Sous COMMANDE
        '==========================
        'Assert.IsTrue(objSCMD.faxer("01234567"), "Faxer" & objSCMD.getErreur())

        'V - Suppression de la Sous commande
        '=================================
        ' Modification de la commande
        objSCMD.bDeleted = True
        Assert.IsTrue(objSCMD.Save(), "Delete" & objSCMD.getErreur())
        'Rechargement dans un autre objet
        objSCMD2 = New SousCommande(m_oCmd, m_oFourn)
        Assert.IsFalse(objSCMD2.load(nid), "Load")
    End Sub
    <TestMethod()> Public Sub T16_DB()
        Dim objLG As LgCommande

        Dim objSCMD As SousCommande
        Dim objLgCMD1 As LgCommande
        Dim objLgCMD2 As LgCommande
        Dim nid As Long

        objLgCMD1 = m_oCmd.AjouteLigne("10", m_oProduit, 10, 12)
        objLgCMD1.calculPrixTotal()
        objLgCMD2 = m_oCmd.AjouteLigne("20", m_oProduit, 20, 12)
        objLgCMD1.calculPrixTotal()
        Assert.IsTrue(m_oCmd.save(), "Sauvegarde de la Commande" & m_oCmd.getErreur())

        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.IsTrue(objSCMD.Save(), "Creation de Sous-Commande" & objSCMD.getErreur())

        'Chargement de la Sous Commande Vide
        Assert.AreEqual(objSCMD.colLignes.Count, 0, "Collection non vide")

        ''Ajout d' 1 ligne de SousCommande
        objLgCMD1 = objSCMD.AjouteLigne(objLgCMD1)
        Assert.IsTrue(Not objLgCMD1 Is Nothing, "Ajout OPRD1")
        Assert.AreEqual(objLgCMD1.num, objSCMD.colLignes.Count * 10, "Num1")
        Assert.IsTrue(objSCMD.totalTTC = objLgCMD1.prixTTC, "Calculm du prix Total")

        ''Sauvegarde de la SousCommande
        objSCMD.totalTTC = 123
        Assert.IsTrue(objSCMD.Save(), "Ocmd.Save" & objSCMD.getErreur())

        ''Rechargement de la SousCommande et de ses lignes
        nid = objSCMD.id
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.IsTrue(objSCMD.load(nid), "ObjSCMD.load")
        Assert.IsTrue(objSCMD.loadcolLignes(), "objSCMD.loadColLignes")
        Assert.AreEqual(objSCMD.colLignes.Count, 1, "colLignes.count ")
        Assert.AreEqual(objSCMD.totalTTC, CDec(123), "Calcul du prix Total")

        ''Ajout d'une ligne de SousCommande
        Assert.IsTrue(Not objSCMD.AjouteLigne(objLgCMD2) Is Nothing, "SCMD.AjouteLg 2" & objSCMD.getErreur())
        ''Sauvegarde de la SousCommande
        Assert.IsTrue(objSCMD.Save(), "SCMD.Save Apres Ajout Ligne 2" & objSCMD.getErreur())
        ''Rechargement de la souscommande
        nid = objSCMD.id
        objSCMD = New SousCommande(m_oCmd, m_oFourn)

        Assert.IsTrue(objSCMD.load(nid), "SCMD.load" & objSCMD.getErreur)
        Assert.IsTrue(objSCMD.loadcolLignes(), "objSCMD.loadColLignes")
        Assert.AreEqual(objSCMD.colLignes.Count, 2, "SCMD.count")

        'Maj d'une ligne de la precommande
        objLG = objSCMD.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteCommande = 150
        objLG.prixU = 15
        ''Sauvegarde de la souscommande
        Assert.IsTrue(objSCMD.Save(), "SCmd.Save" & objSCMD.getErreur())
        'Rechargement de la sousComande
        nid = objSCMD.id
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.IsTrue(objSCMD.load(nid), "SCMD.load" & objSCMD.getErreur())
        Assert.IsTrue(objSCMD.loadcolLignes(), "objSCMD.loadColLignes")
        Assert.AreEqual(objSCMD.colLignes.Count, 2, "SCMD.count ")
        objLG = objSCMD.colLignes(1)
        Assert.AreEqual(objLG.qteCommande, CDec(150))
        Assert.AreEqual(objLG.prixU, CDec(15))

        'Maj d'une ligne de la Commande (Qte Livrée)
        objLG = objSCMD.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteLiv = 100
        objLG.prixU = 16
        ''Sauvegarde de la souscommande
        Assert.IsTrue(objSCMD.Save(), "OSCMD.Save" & objSCMD.getErreur())
        'Rechargement de la sous commande  et de ses lignes
        nid = objSCMD.id
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.IsTrue(objSCMD.load(nid), "OSCMD.load" & objSCMD.getErreur())
        Assert.IsTrue(objSCMD.loadcolLignes(), "objSCMD.loadColLignes")
        Assert.AreEqual(objSCMD.colLignes.Count, 2, "ColLignes.count ")
        objLG = objSCMD.colLignes(1)
        Assert.AreEqual(objLG.qteLiv, CDec(100))
        Assert.AreEqual(objLG.prixU, CDec(16))

        'Maj d'une ligne de la Commande (Qte Facturée)
        objLG = objSCMD.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteLiv = 101
        objLG.prixU = 17
        ''Sauvegarde de la souscommande
        Assert.IsTrue(objSCMD.Save(), "OSCMD.Save" & objSCMD.getErreur())
        'Rechargement de la sous Commande
        nid = objSCMD.id
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.IsTrue(objSCMD.load(nid), "OSCMD.load" & objSCMD.getErreur())
        Assert.IsTrue(objSCMD.loadcolLignes(), "objSCMD.loadColLignes")
        Assert.AreEqual(objSCMD.colLignes.Count, 2, "ColLignes.count ")
        objLG = objSCMD.colLignes(1)
        Assert.AreEqual(objLG.qteLiv, CDec(101))
        Assert.AreEqual(objLG.prixU, CDec(17))

        'Assert.IsTrue(objSCMD.faxer("01234567"), "Faxer" & objSCMD.getErreur())

        'Chargement de la sous-commande

        'Suppression de la sous Commande
        objSCMD.bDeleted = True
        Assert.IsTrue(objSCMD.Save(), "OSCMD.delete" & objSCMD.getErreur())

    End Sub
    <TestMethod()> Public Sub T20_Eclatement()
        Dim objLG As LgCommande

        Dim objSCMD As SousCommande
        Dim objLgCMD1 As LgCommande
        Dim objLgCMD2 As LgCommande
        Dim oFRN2 As Fournisseur
        Dim oPRD2 As Produit
        Dim oCMD As CommandeClient
        Dim nIDCmd As Long
        Dim nidSCMD1 As Long
        Dim nidSCMD2 As Long




        oFRN2 = New Fournisseur("FRNT20" & Now(), "Fournisseur2")
        Assert.IsTrue(oFRN2.Save(), "FRN2.save" & oFRN2.getErreur())
        oPRD2 = New Produit("PRDT20" & Now(), oFRN2, 1990)
        Assert.IsTrue(oPRD2.save(), "OPRD2.Save" & oPRD2.getErreur())

        'Création d'une commande client
        oCMD = New CommandeClient(m_oClient)
        oCMD.save()
        objLgCMD1 = oCMD.AjouteLigne("10", m_oProduit, 10, 12)
        objLgCMD1.calculPrixTotal()
        objLgCMD2 = oCMD.AjouteLigne("20", oPRD2, 20, 12)
        objLgCMD1.calculPrixTotal()
        oCMD.calculPrixTotal()
        oCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        oCMD.oTransporteur.AdresseLivraison.nom = "TestTransporteur"
        oCMD.refLivraison = "REFLEVRAISON"
        oCMD.CommFacturation.comment = "Mon Comment"
        oCMD.dateLivraison = "31/08/2004"
        Assert.IsTrue(oCMD.save(), "Sauvegarde de la Commande" & racine.getErreur())

        'Génération des sous-commandes
        Assert.IsTrue(oCMD.generationSousCommande(), "OCMD.EclatementCommande()" & racine.getErreur())

        Assert.AreEqual(oCMD.colSousCommandes.Count, 2, "OCMD.colSousCommande.Count" & oCMD.colSousCommandes.toString())
        'Vérification de la première sous-commande
        objSCMD = oCMD.colSousCommandes.Item(1)
        'Test du passage d'info entre la commande et la souscommande
        Assert.AreEqual(objSCMD.oFournisseur.id, m_oFourn.id, "Test Fournisseur1" & objSCMD.oFournisseur.toString())
        Assert.AreEqual(objSCMD.totalTTC, objLgCMD1.prixTTC, "Test PrixTTC1" & objSCMD.toString())
        Assert.IsTrue(objSCMD.oTransporteur.Equals(oCMD.oTransporteur), "Transporteur Différents")
        Assert.AreEqual(objSCMD.refLivraison, oCMD.refLivraison, "REfLivraison Différents")
        Assert.AreEqual(objSCMD.CommFacturation.comment, oCMD.CommFacturation.comment, "ComFacturation Différents")
        Assert.AreEqual(objSCMD.dateLivraison, oCMD.dateLivraison, "dateLivraison Différents")

        'Vérification de la seconde sous-commande
        objSCMD = oCMD.colSousCommandes.Item(2)
        Assert.AreEqual(objSCMD.oFournisseur.id, oFRN2.id, "Test Fournisseur2" & objSCMD.oFournisseur.toString())
        Assert.AreEqual(objSCMD.totalTTC, objLgCMD2.prixTTC, "Test PrixTTC2" & objSCMD.toString())
        'Vérification des Lignes de la commande

        '1)dans la Commande , les lignes sont marquées comme éclatées
        For Each objLG In oCMD.colLignes
            Assert.IsTrue(objLG.bLigneEclatee, "objLG.bLigneEclatee = False")
        Next objLG

        '2) dans chaque Sous Commandes
        For Each objSCMD In oCMD.colSousCommandes
            'Assert.IsTrue(objSCMD.loadcolLignes(), "SCMD.loadcolLignes" & objSCMD.getErreur())
            For Each objLG In objSCMD.colLignes
                Assert.IsTrue(objLG.bLigneEclatee, "objLG.bLigneEclatee = False")
                Assert.AreEqual(objLG.idSCmd, objSCMD.id, "objLG.idSCMD<> ")
            Next objLG
        Next objSCMD

        'Sauvegarde de la commande => Sauvegarde des Sous Commandes
        Assert.IsTrue(oCMD.save(), "SaveCommandes" & racine.getErreur())
        nidSCMD1 = CType(oCMD.colSousCommandes.Item(1), SousCommande).id
        Assert.IsTrue(nidSCMD1 <> 0, "idSCMD apres Sauvegarde")
        nidSCMD2 = CType(oCMD.colSousCommandes.Item(2), SousCommande).id
        Assert.IsTrue(nidSCMD1 <> 0, "idSCMD apres Sauvegarde")

        'Rechargement de la commande
        nIDCmd = oCMD.id
        oCMD = New CommandeClient(m_oClient)
        Assert.IsTrue(oCMD.load(nIDCmd), "Load CMD" & racine.getErreur())
        Assert.IsTrue(oCMD.LoadColSousCommande, "Load colSCMD" & racine.getErreur())
        'Vérification des Lignes de commande
        '1)dans la Commande (Les lignes sont éclatées et portent la référence de la sous commande
        For Each objLG In oCMD.colLignes
            Assert.IsTrue(objLG.bLigneEclatee, "objLG.bLigneEclatee = False")
            Assert.IsTrue(objLG.idSCmd <> 0, "objLG.idSCMD= 0")
        Next objLG
        '2) dans chaque Sous Commandes
        ' Les Lignes sont éclatées et portent la référence de la sous commande
        For Each objSCMD In oCMD.colSousCommandes
            For Each objLG In objSCMD.colLignes
                Assert.IsTrue(objLG.bLigneEclatee, "objLG.bLigneEclatee = False")
                Assert.AreEqual(objLG.idSCmd, objSCMD.id, "objLG.idSCMD<> ")
            Next objLG
        Next objSCMD


        'Modification de Commandes
        oCMD.CommCommande.comment = "Modification de Commentaires"
        Assert.IsTrue(oCMD.save())
        'Vérification des Lignes de commande
        '1)dans la Commande
        For Each objLG In oCMD.colLignes
            Assert.IsTrue(objLG.bLigneEclatee, "objLG.bLigneEclatee = False")
            Assert.IsTrue(objLG.idSCmd <> 0, "objLG.idSCMD= 0")
        Next objLG
        '2) dans chaque Sous Commandes
        For Each objSCMD In oCMD.colSousCommandes
            For Each objLG In objSCMD.colLignes
                Assert.IsTrue(objLG.bLigneEclatee, "objLG.bLigneEclatee = False")
                Assert.AreEqual(objLG.idSCmd, objSCMD.id, "objLG.idSCMD<> ")
            Next objLG
        Next objSCMD


        'Rechargement de la commande Client
        nIDCmd = oCMD.id
        oCMD = New CommandeClient(New Client("", ""))
        Assert.IsTrue(oCMD.load(nIDCmd), "Chargement de la commande" & oCMD.getErreur())
        Assert.IsTrue(oCMD.LoadColSousCommande(), "Chargement des SousCommandes" & oCMD.getErreur())
        Assert.IsTrue(oCMD.colSousCommandes.Count = 2, "Chargement des SousCommandes Count" & oCMD.getErreur())

        objSCMD = oCMD.colSousCommandes.Item(1)
        Assert.IsTrue(objSCMD.loadcolLignes(), "objSCMD.loadColLignes()")
        Assert.IsTrue(objSCMD.colLignes.Count = 1, "Lignes de SousCommandes 1 Count" & objSCMD.getErreur())

        Assert.IsTrue(objSCMD.id = nidSCMD2 Or objSCMD.id = nidSCMD1, "Comparaison ID1 " & objSCMD.id & "," & nidSCMD1 & "," & nidSCMD2)
        objSCMD = oCMD.colSousCommandes.Item(2)
        Assert.IsTrue(objSCMD.id = nidSCMD2 Or objSCMD.id = nidSCMD1, "Comparaison ID2 " & objSCMD.id & "," & nidSCMD1 & "," & nidSCMD2)
        Assert.IsTrue(objSCMD.loadcolLignes(), "ObjSCMD.loadColLignes()")
        Assert.IsTrue(objSCMD.colLignes.Count = 1, "Lignes de SousCommandes 2 Count" & objSCMD.getErreur())

        'annulation de la génération des sous commandes
        'oCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        oCMD.annulationSousCommande()
        Assert.IsTrue(oCMD.save(), "oCMD.save" & racine.getErreur())

        'Rechargement de la commande Client
        nIDCmd = oCMD.id
        oCMD = New CommandeClient(New Client("", ""))
        Assert.IsTrue(oCMD.load(nIDCmd), "Chargement de la commande" & oCMD.getErreur())
        Assert.IsTrue(oCMD.LoadColSousCommande(), "Chargement des SousCommandes" & oCMD.getErreur())
        'Vérification des Lignes de commande
        '1)dans la Commande
        For Each objLG In oCMD.colLignes
            Assert.IsFalse(objLG.bLigneEclatee, "objLG.bLigneEclatee = False")
            Assert.AreEqual(0, objLG.idSCmd, "objLG.idSCMD= 0")
        Next objLG
        '2) dans chaque Sous Commandes
        Assert.AreEqual(0, oCMD.colSousCommandes.Count, "colSousCommande.count()")

        oCMD.bDeleted = True
        oCMD.save()

        oPRD2.bDeleted = True
        oFRN2.Save()

        oFRN2.bDeleted = True
        oFRN2.Save()
    End Sub
    <TestMethod()> Public Sub T30_ListeSousCommande()
        Dim colSCMD As List(Of SousCommande)
        Dim objCMDCLT2 As CommandeClient

        'Création de la commande et eclatement
        m_oCmd.dateCommande = "15/09/1984"
        m_oCmd.AjouteLigne("10", m_oProduit, 10, 12.4)
        m_oCmd.AjouteLigne("20", m_oProduit2, 20, 15.6)
        Assert.IsTrue(m_oCmd.save())

        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save())
        m_oCmd.generationSousCommande()
        Assert.IsTrue(m_oCmd.save())

        'Création de seconde commande (Chgmt de date) et eclatement
        objCMDCLT2 = New CommandeClient(m_oClient)
        objCMDCLT2.dateCommande = "15/10/1984"
        objCMDCLT2.AjouteLigne("10", m_oProduit, 10, 12.4)
        objCMDCLT2.AjouteLigne("20", m_oProduit2, 20, 15.6)
        Assert.IsTrue(objCMDCLT2.save())
        objCMDCLT2.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCMDCLT2.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT2.save())
        objCMDCLT2.generationSousCommande()
        Assert.IsTrue(objCMDCLT2.save())


        Persist.shared_connect()

        colSCMD = SousCommande.getListe("15/09/1984", "15/09/1984", "")
        Assert.IsTrue(Not colSCMD Is Nothing, "ListeCMD()" & Persist.getErreur())
        Assert.AreEqual(2, colSCMD.Count, "2 Enregistrements")

        colSCMD = SousCommande.getListe("01/10/1984", "31/10/1984", "")
        Assert.IsTrue(Not colSCMD Is Nothing, "ListeCMD()" & Persist.getErreur())
        Assert.AreEqual(2, colSCMD.Count, "2 Enregistrements")


        colSCMD = SousCommande.getListe("01/09/1984", "31/10/1984", "")
        Assert.IsTrue(Not colSCMD Is Nothing, "ListeCMD()" & Persist.getErreur())
        Assert.AreEqual(4, colSCMD.Count, "4 Enregistrements")

        colSCMD = SousCommande.getListe(, , "FRN%")
        Assert.IsTrue(Not colSCMD Is Nothing, "ListeCMD()" & Persist.getErreur())
        Assert.IsTrue(colSCMD.Count >= 4, "4 Enregistrements")

        objCMDCLT2.bDeleted = True
        Assert.IsTrue(objCMDCLT2.save())

        Persist.shared_disconnect()

    End Sub
    <TestMethod()> Public Sub T40_Etats()
        'Vérification des états de la sous commande
        'Générée --- faxer ---> transmiseFax
        'Générée --- ExportInternet --> exportée
        'TransmiseFax --- RapprochementManuel ---> Rapprochée
        'Exportée --- importInternet ---> Rapprochée
        'Rapprochée --- AnnRapprochée ---> TransmiseFax/exportée
        'Rapprochée --- Provisionnement ---> Provisionnée
        'Provisionnée --- AnnProvisionnement ---> Rapprochée
        'Provisionnée --- Facturation Commission ---> Facturée
        'Facturée --- Ann Facturation Commission ---> Provisisonnée
        Dim objSCMD As SousCommande

        'Tramission par Fax
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.AreEqual(New EtatSSCommandeGeneree(False, False).codeEtat, objSCMD.etat.codeEtat, "Etat sous commande origine = Générée")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
        Assert.AreEqual(New EtatSSCommandeTransmise(False, False).codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = TransmiseFax")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
        Assert.AreEqual(New EtatSSCommandeRapprochee(False, False).codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = Rapprochée")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFacturer)
        Assert.AreEqual(New EtatSSCommandeFacturee().codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = Facturée")

        'Transmission par internet
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.AreEqual(New EtatSSCommandeGeneree(False, False).codeEtat, objSCMD.etat.codeEtat, "Etat sous commande origine = Générée")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
        Assert.AreEqual(New EtatSSCommandeExporteeInt().codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = TransmiseFax")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
        Assert.AreEqual(New EtatSSCommandeRapprocheeInt(False, False).codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = Rapprochée Internet")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFacturer)
        Assert.AreEqual(New EtatSSCommandeFacturee().codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = Facturée")

        'Transmission Annulation de Rapprochement internet
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        Assert.AreEqual(New EtatSSCommandeGeneree(False, False).codeEtat, objSCMD.etat.codeEtat, "Etat sous commande origine = Générée")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
        Assert.AreEqual(New EtatSSCommandeExporteeInt().codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = Exportéeinternet")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
        Assert.AreEqual(New EtatSSCommandeRapprocheeInt().codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = RapprochéeInt")
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDAnnImportInternet)
        Assert.AreEqual(New EtatSSCommandeExporteeInt().codeEtat, objSCMD.etat.codeEtat, "Etat sous commande Attendu = Exportée Int")


        'Transmission Annulation de Facturation
        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFacturer)
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDAnnFacturer)
    End Sub ' T40_Etats
    <TestMethod()> Public Sub T50_TOCSV_Structure()
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim nfile As Integer
        Dim nLineNumber As Integer
        Dim strResult As String
        Dim tabCSV As String() = Nothing
        Dim n As Integer

        'Ajout de 2 Lignes à la commande
        oLgCmd = m_oCmd.AjouteLigne("10", m_oProduit, 10, 15.55)
        oLgCmd.qteLiv = 9

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


        nfile = FreeFile()
        FileOpen(nfile, "./adel.txt", OpenMode.Output, OpenAccess.Write, OpenShare.Shared)
        For Each oSCmd In m_oCmd.colSousCommandes
            strResult = oSCmd.toCSV()
            Print(nfile, strResult)
        Next oSCmd
        FileClose(nfile)

        nfile = FreeFile()
        FileOpen(nfile, "./adel.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        nLineNumber = 0
        While Not EOF(nfile)
            nLineNumber = nLineNumber + 1
            strResult = LineInput(nfile)
            Console.WriteLine(strResult)
            If nLineNumber = 1 Then
                tabCSV = strResult.Split(";")
                oSCmd = m_oCmd.colSousCommandes(1)
                'Vérification du contenu de la ligne
                n = 0
                Assert.AreEqual(Trim(oSCmd.id), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.codeCommandeClient), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(oSCmd.dateCommande, "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.rs), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.rue1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.rue2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.cp), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.ville), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.tel), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.fax), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.port), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.Email), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.rue1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.rue2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.cp), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.ville), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.tel), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.fax), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.port), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.Email), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.rs), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.rue1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.rue2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.cp), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.ville), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.tel), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.fax), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.port), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseLivraison.Email), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.rue1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.rue2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.cp), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.ville), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.tel), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.fax), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.port), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.AdresseFacturation.Email), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.libTypeClient), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oCmd.typeTransport), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.libModeReglement), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oCmd.refLivraison), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.banque), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.rib1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.rib2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.rib3), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.rib4), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(oSCmd.dateLivraison, "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(oSCmd.dateEnlevement, "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual(CDec(139.95), CDec(tabCSV(n)))
                n = n + 1
                Assert.AreEqual(CDec(167.94), CDec(tabCSV(n)))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.CommFacturation.comment), tabCSV(n))
                n = n + 1
                Assert.AreEqual("10", tabCSV(n))
                n = n + 1
                Assert.AreEqual(CDec(9), CDec(tabCSV(n)))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.libCouleur), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.millesime), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.libConditionnement), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.libContenant), tabCSV(n))
                n = n + 1
                Assert.AreEqual(CDec(15.55), CDec(tabCSV(n)))
                n = n + 1
                Assert.AreEqual("", tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(oSCmd.dateFactFournisseur, "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual(oSCmd.totalHTFacture, CDec(tabCSV(n))) ' Total HT Facturé
                n = n + 1
                Assert.AreEqual(oSCmd.totalTTCFacture, CDec(tabCSV(n))) 'Total TTC Facturé
                n = n + 1
                Assert.AreEqual(oSCmd.tauxCommission, CDec(tabCSV(n))) 'Taux de commission
                n = n + 1

            End If

        End While
        FileClose(nfile)

        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")
    End Sub
    ''' <summary>
    ''' Test l'export d'une sous commande avec un intermédiaire
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T50_TOCSV_StructureAvecIntermediaire()
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim nfile As Integer
        Dim nLineNumber As Integer
        Dim strResult As String
        Dim tabCSV As String() = Nothing
        Dim n As Integer
        Dim oIntermediaire As Client

        oIntermediaire = New Client("INTER", "Intermediaire")
        oIntermediaire.setTypeIntermediaire("HOBIVIN")
        oIntermediaire.save()

        'Ajout de 2 Lignes à la commande
        m_oCmd.Origine = "HOBIVIN"
        oLgCmd = m_oCmd.AjouteLigne("10", m_oProduit, 10, 15.55)
        oLgCmd.qteLiv = 9
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")

        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(oIntermediaire), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


        nfile = FreeFile()
        FileOpen(nfile, "./adel.txt", OpenMode.Output, OpenAccess.Write, OpenShare.Shared)
        For Each oSCmd In m_oCmd.colSousCommandes
            strResult = oSCmd.toCSV()
            Print(nfile, strResult)
        Next oSCmd
        FileClose(nfile)

        nfile = FreeFile()
        FileOpen(nfile, "./adel.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        nLineNumber = 0
        While Not EOF(nfile)
            nLineNumber = nLineNumber + 1
            strResult = LineInput(nfile)
            Console.WriteLine(strResult)
            If nLineNumber = 1 Then
                tabCSV = strResult.Split(";")
                oSCmd = m_oCmd.colSousCommandes(1)
                'Vérification du contenu de la ligne
                n = 0
                Assert.AreEqual(Trim(oSCmd.id), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.codeCommandeClient), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(oSCmd.dateCommande, "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.rs), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.rue1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.rue2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.cp), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.ville), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.tel), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.fax), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.port), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oFournisseur.AdresseFacturation.Email), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.rue1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.rue2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.cp), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.ville), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.tel), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.fax), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.port), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oTransporteur.AdresseLivraison.Email), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.rs), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.rue1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.rue2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.cp), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.ville), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.tel), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.fax), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.port), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseLivraison.Email), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.rue1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.rue2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.cp), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.ville), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.tel), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.fax), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.port), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.AdresseFacturation.Email), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.libTypeClient), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oCmd.typeTransport), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.oClient.libModeReglement), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oCmd.refLivraison), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.banque), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.rib1), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.rib2), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.rib3), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(oIntermediaire.rib4), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(oSCmd.dateLivraison, "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(oSCmd.dateEnlevement, "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual(CDec(139.95), CDec(tabCSV(n)))
                n = n + 1
                Assert.AreEqual(CDec(167.94), CDec(tabCSV(n)))
                n = n + 1
                Assert.AreEqual(Trim(oSCmd.CommFacturation.comment), tabCSV(n))
                n = n + 1
                Assert.AreEqual("10", tabCSV(n))
                n = n + 1
                Assert.AreEqual(CDec(9), CDec(tabCSV(n)))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.nom), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.libCouleur), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.millesime), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.libConditionnement), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oProduit.libContenant), tabCSV(n))
                n = n + 1
                Assert.AreEqual(CDec(15.55), CDec(tabCSV(n)))
                n = n + 1
                Assert.AreEqual("", tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(oSCmd.dateFactFournisseur, "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual(oSCmd.totalHTFacture, CDec(tabCSV(n))) ' Total HT Facturé
                n = n + 1
                Assert.AreEqual(oSCmd.totalTTCFacture, CDec(tabCSV(n))) 'Total TTC Facturé
                n = n + 1
                Assert.AreEqual(oSCmd.tauxCommission, CDec(tabCSV(n))) 'Taux de commission
                n = n + 1

            End If

        End While
        FileClose(nfile)

        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")
        oIntermediaire.bDeleted = True
        oIntermediaire.save()
    End Sub
    <TestMethod()> Public Sub T51_TOCSV_Lines()
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim nfile As Integer
        Dim nLineNumber As Integer
        Dim strResult As String
        Dim tabCSV As String() = Nothing

        'Ajout de 2 Lignes à la commande
        oLgCmd = m_oCmd.AjouteLigne("10", m_oProduit, 10, 15.55)
        oLgCmd.qteLiv = 9
        oLgCmd = m_oCmd.AjouteLigne("20", m_oProduit2, 8, 5.77)
        oLgCmd.qteLiv = 7

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


        If System.IO.File.Exists("./adel.txt") Then
            System.IO.File.Delete("./adel.txt")
        End If
        nfile = FreeFile()
        FileOpen(nfile, "./adel.txt", OpenMode.Output, OpenAccess.Write)
        For Each oSCmd In m_oCmd.colSousCommandes
            strResult = oSCmd.toCSV()
            Print(nfile, strResult)
        Next oSCmd
        FileClose(nfile)

        nfile = FreeFile()
        FileOpen(nfile, "./adel.txt", OpenMode.Input, OpenAccess.Read)
        nLineNumber = 0
        While Not EOF(nfile)
            nLineNumber = nLineNumber + 1
            strResult = LineInput(nfile)
        End While
        FileClose(nfile)

        Assert.AreEqual(2, nLineNumber, "Deux Lignes de fichier")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")
        System.IO.File.Delete("./adel.txt")

    End Sub
    <TestMethod()> Public Sub T60_GenerePDF()
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim tabCSV As String() = Nothing

        'Ajout de 2 Lignes à la commande
        oLgCmd = m_oCmd.AjouteLigne("10", m_oProduit, 10, 15.55)
        oLgCmd.qteLiv = 9
        oLgCmd = m_oCmd.AjouteLigne("20", m_oProduit2, 8, 5.77)
        oLgCmd.qteLiv = 7

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


        For Each oSCmd In m_oCmd.colSousCommandes
            If System.IO.File.Exists("./TEMP/" & oSCmd.code & ".PDF") Then
                System.IO.File.Delete("./TEMP/" & oSCmd.code & ".PDF")
            End If
            oSCmd.genererPDF("../../../vini_app/bin/debug/", "./TEMP/" & oSCmd.code & ".PDF")
            Assert.IsTrue(System.IO.File.Exists("./TEMP/" & oSCmd.code & ".PDF"), "Fichier PDF existant")
        Next oSCmd
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


    End Sub
    <TestMethod()> Public Sub T70_FromCSV()
        'Test la mise à jour de la sous commande à partir d'un fichier CSV 
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim nfile As Integer
        Dim nLineNumber As Integer
        Dim strResult As String
        Dim tabCSV As String() = Nothing
        Dim nId As Integer

        '===============================
        ' I - Préparation du fichier CSV 
        '===============================
        'Ajout de 2 Lignes à la commande
        oLgCmd = m_oCmd.AjouteLigne("10", m_oProduit, 10, 15.55)
        oLgCmd.qteLiv = 9
        oLgCmd = m_oCmd.AjouteLigne("20", m_oProduit2, 8, 5.77)
        oLgCmd.qteLiv = 7

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


        'Initialisation de informations de factures fournisseurs pour créer le fichier CSV
        For Each oSCmd In m_oCmd.colSousCommandes
            oSCmd.bExportInternet = True
            oSCmd.Save()
        Next oSCmd
        For Each oSCmd In m_oCmd.colSousCommandes
            oSCmd.refFactFournisseur = "FACTFOURN" & oSCmd.code
            oSCmd.dateFactFournisseur = #2/6/1964#
            oSCmd.totalHTFacture = 155.55
            oSCmd.totalTTCFacture = 190.67
            oSCmd.tauxCommission = 9.99
            oSCmd.MontantCommission = 111.11
            oSCmd.Save()
        Next oSCmd
        'Création du fichier CSV 
        If System.IO.File.Exists("./toViniCom.csv") Then
            System.IO.File.Delete("./toViniCom.csv")
        End If
        nfile = FreeFile()
        FileOpen(nfile, "./toVinicom.csv", OpenMode.Output, OpenAccess.Write)
        For Each oSCmd In m_oCmd.colSousCommandes
            strResult = oSCmd.FTO_toCSVFromInternet
            Print(nfile, strResult)
        Next oSCmd
        FileClose(nfile)


        '=========================================================================
        ' II - Lecture du fichier CSV
        '=========================================================================
        nfile = FreeFile()
        FileOpen(nfile, "./toVinicom.csv", OpenMode.Input, OpenAccess.Read)
        nLineNumber = 0
        While Not EOF(nfile)
            nLineNumber = nLineNumber + 1
            strResult = LineInput(nfile)
            Console.WriteLine(strResult)
            tabCSV = strResult.Split(";")
            nId = tabCSV(0)
            oSCmd = SousCommande.createandload(nId)
            Assert.AreEqual(nId, oSCmd.id, "Id de la sous commande")
            oSCmd.refFactFournisseur = tabCSV(1)
            oSCmd.dateFactFournisseur = Mid(tabCSV(2), 1, 2) & "/" & Mid(tabCSV(2), 3, 2) & "/" & Mid(tabCSV(2), 5, 4)
            oSCmd.totalHTFacture = tabCSV(3)
            oSCmd.totalTTCFacture = tabCSV(4)
            'oSCmd.tauxCommission = tabCSV(5)
            oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
            Assert.AreEqual("FACTFOURN" & oSCmd.code, oSCmd.refFactFournisseur)
            Assert.AreEqual(#2/6/1964#, oSCmd.dateFactFournisseur)
            Assert.AreEqual(CDec(155.55), CDec(oSCmd.totalHTFacture))
            Assert.AreEqual(CDec(190.67), CDec(oSCmd.totalTTCFacture))
            Assert.AreEqual(CDec(9.99), CDec(oSCmd.tauxCommission))
            Assert.AreEqual(vncEnums.vncEtatCommande.vncSCMDRapprocheeInt, oSCmd.etat.codeEtat)
            Assert.AreEqual(True, oSCmd.bExportInternet)
        End While
        FileClose(nfile)

        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")
    End Sub 'T70_FromCSV
    <TestMethod()> Public Sub T71_LoadFromCode()
        'Test de rechargement de la sous commande avec le code

        Dim objSCMD As SousCommande
        Dim objLgCMD1 As LgCommande
        Dim objLgCMD2 As LgCommande
        Dim strCodeSCMD As String

        objLgCMD1 = m_oCmd.AjouteLigne("10", m_oProduit, 10, 12)
        objLgCMD1.calculPrixTotal()
        objLgCMD2 = m_oCmd.AjouteLigne("20", m_oProduit, 20, 12)
        objLgCMD1.calculPrixTotal()
        Assert.IsTrue(m_oCmd.save(), "Sauvegarde de la Commande" & m_oCmd.getErreur())

        objSCMD = New SousCommande(m_oCmd, m_oFourn)
        ''Ajout d' 1 ligne de SousCommande
        objLgCMD1 = objSCMD.AjouteLigne(objLgCMD1)
        objSCMD.totalTTC = 123
        Assert.IsTrue(objSCMD.Save(), "Creation de Sous-Commande" & objSCMD.getErreur())

        strCodeSCMD = objSCMD.code

        ''Rechargement de la SousCommande et de ses lignes
        objSCMD = SousCommande.createandload(strCodeSCMD)
        Assert.IsTrue(objSCMD.loadcolLignes(), "objSCMD.loadColLignes")
        Assert.AreEqual(objSCMD.colLignes.Count, 1, "colLignes.count ")
        Assert.AreEqual(objSCMD.totalTTC, CDec(123), "Calcul du prix Total")


        'Suppression de la sous Commande
        objSCMD.bDeleted = True
        Assert.IsTrue(objSCMD.Save(), "OSCMD.delete" & objSCMD.getErreur())

    End Sub 'T71 LoadFromCode
    'Test de la ganaration de sousCommande avec un intermédiaire
    <TestMethod()> Public Sub T80_EclatementAvecIntermedaire()

        Dim objLG As LgCommande

        Dim objSCMD As SousCommande
        Dim objLgCMD1 As LgCommande
        Dim objLgCMD2 As LgCommande
        Dim oFRN2 As Fournisseur
        Dim oPRD2 As Produit
        Dim oFRN3 As Fournisseur
        Dim oPRD3 As Produit
        Dim oCMD As CommandeClient

        Dim oCltIntermediaire As Client
        oCltIntermediaire = New Client("CLTINTER", "ClientIntermédiaire")
        oCltIntermediaire.setTypeIntermediaire("HOBIVIN")
        oCltIntermediaire.save()


        oFRN2 = New Fournisseur("FRNT20" & Now(), "Fournisseur2")
        oFRN2.bExportInternet = vncTypeExportScmd.vncExportInternet
        Assert.IsTrue(oFRN2.Save(), "FRN2.save" & oFRN2.getErreur())
        oPRD2 = New Produit("PRDT20" & Now(), oFRN2, 1990)
        Assert.IsTrue(oPRD2.save(), "OPRD2.Save" & oPRD2.getErreur())

        oFRN3 = New Fournisseur("FRNT30" & Now(), "Fournisseur3")
        oFRN3.bExportInternet = vncTypeExportScmd.vncExportQuadra
        Assert.IsTrue(oFRN3.Save(), "FRN2.save" & oFRN2.getErreur())
        oPRD3 = New Produit("PRDT30" & Now(), oFRN3, 1990)
        Assert.IsTrue(oPRD3.save(), "OPRD2.Save" & oPRD3.getErreur())


        'Création d'une commande client origine "VINICOM" 
        oCMD = New CommandeClient(m_oClient)
        oCMD.Origine = "Vinicom"
        oCMD.save()
        oCMD.AjouteLigne("10", m_oProduit, 10, 12)
        oCMD.AjouteLigne("20", oPRD2, 20, 22)
        oCMD.AjouteLigne("30", oPRD3, 30, 32)

        Assert.IsTrue(oCMD.save(), "Sauvegarde de la Commande" & racine.getErreur())

        'Génération des sous-commandes (Sans intermédiaire)
        Assert.IsTrue(oCMD.generationSousCommande(), "OCMD.EclatementCommande()" & racine.getErreur())

        Assert.AreEqual(oCMD.colSousCommandes.Count, 3, "OCMD.colSousCommande.Count" & oCMD.colSousCommandes.toString())
        'Vérification de la première sous-commande (Fournisseur Sans Export)

        objSCMD = oCMD.colSousCommandes.Item(1)
        'Test du passage d'info entre la commande et la souscommande
        'Producteur
        Assert.AreEqual(objSCMD.oFournisseur.id, m_oFourn.id, "Test Fournisseur1" & objSCMD.oFournisseur.toString())
        'Client = Client Originel
        Assert.AreEqual(objSCMD.oTiers.id, m_oClient.id, "Le Tiers de la Sous Commande est le client initial" & objSCMD.oTiers.toString())

        'Vérification de la seconde sous-commande
        objSCMD = oCMD.colSousCommandes.Item(2)
        Assert.AreEqual(objSCMD.oFournisseur.id, oFRN2.id, "Test Fournisseur2" & objSCMD.oFournisseur.toString())
        'Client = Client initial
        Assert.AreEqual(objSCMD.oTiers.id, m_oClient.id, "Le Tiers de la Sous Commande devrait être le Client initial" & objSCMD.oTiers.toString())

        'Vérification de la Troisième sous-commande
        objSCMD = oCMD.colSousCommandes.Item(3)
        Assert.AreEqual(objSCMD.oFournisseur.id, oFRN3.id, "Fournisseur3" & objSCMD.oFournisseur.toString())
        'Client = Client initial
        Assert.AreEqual(objSCMD.oTiers.id, m_oClient.id, "Le Tiers de la Sous Commande devrait être le Client initial" & objSCMD.oTiers.toString())

        'Suppression de la Commande
        oCMD.bDeleted = True
        oCMD.save()

        'Création d'une commande client origine AVEC intermédiaire
        oCMD = New CommandeClient(m_oClient)
        oCMD.Origine = "HOBIVIN"
        oCMD.save()
        objLgCMD1 = oCMD.AjouteLigne("10", m_oProduit, 10, 12)
        objLgCMD2 = oCMD.AjouteLigne("20", oPRD2, 20, 12)
        objLgCMD2 = oCMD.AjouteLigne("30", oPRD3, 30, 32)
        Assert.IsTrue(oCMD.save(), "Sauvegarde de la Commande" & racine.getErreur())

        'Génération des sous-commandes avec un intermédiaire
        Assert.IsTrue(oCMD.generationSousCommande(oCltIntermediaire), "OCMD.EclatementCommande()" & racine.getErreur())

        Assert.AreEqual(oCMD.colSousCommandes.Count, 3, "OCMD.colSousCommande.Count" & oCMD.colSousCommandes.toString())

        'Vérification de la première sous-commande (Fournisseur = Pas D'export)
        objSCMD = oCMD.colSousCommandes.Item(1)
        'Test du passage d'info entre la commande et la souscommande
        'Producteur
        Assert.AreEqual(objSCMD.oFournisseur.id, m_oFourn.id, "Test Fournisseur1" & objSCMD.oFournisseur.toString())
        'Client = Client Originel
        Assert.AreEqual(objSCMD.oTiers.id, m_oClient.id, "Le Tiers de la Sous Commande doit être le client Originel" & objSCMD.oTiers.toString())

        'Vérification de la seconde sous-commande (Fournisseur = Export Internet)
        objSCMD = oCMD.colSousCommandes.Item(2)
        Assert.AreEqual(objSCMD.oFournisseur.id, oFRN2.id, "Test Fournisseur2" & objSCMD.oFournisseur.toString())
        'Client = Intermédiaire
        Assert.AreEqual(objSCMD.oTiers.id, oCltIntermediaire.id, "Le Tiers de la Sous Commande devrait être l'intermédiaire" & objSCMD.oTiers.toString())

        'Vérification de la seconde sous-commande (Fournisseur = Export Quadra)
        objSCMD = oCMD.colSousCommandes.Item(3)
        Assert.AreEqual(objSCMD.oFournisseur.id, oFRN3.id, "Test Fournisseur2" & objSCMD.oFournisseur.toString())
        'Client = Intermédiaire
        Assert.AreEqual(objSCMD.oTiers.id, m_oClient.id, "Le Tiers de la Sous Commande devrait être le client d'origine" & objSCMD.oTiers.toString())


        oCMD.bDeleted = True
        oCMD.save()

        oPRD3.bDeleted = True
        oFRN3.Save()

        oFRN3.bDeleted = True
        oFRN3.Save()

        oPRD2.bDeleted = True
        oFRN2.Save()

        oFRN2.bDeleted = True
        oFRN2.Save()
    End Sub
    Private Sub GenereSousCommande(pcmd As CommandeClient)
        pcmd.changeEtat(vncActionEtatCommande.vncActionLivrer)
        pcmd.save()
        pcmd.generationSousCommande()
        pcmd.save()
    End Sub


End Class



