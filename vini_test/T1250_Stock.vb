'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1250_Stock
    Inherits test_Base

    Private m_obj As Produit
    Private objClient As Client
    Private objFournisseur As Fournisseur
    Private objP1 As Produit
    Private objP2 As Produit
    Private objP3 As Produit
    Private objP4 As Produit
    Private objP5 As Produit

    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        m_obj = New Produit("", New Fournisseur, 1990)

        'Création d'un client
        objClient = New Client("CLT01", "CLIENT 01")
        Assert.IsTrue(objClient.save())

        'Création d'un fournisseur
        objFournisseur = New Fournisseur("FRN01", "FOURNISSEUR 01")
        Assert.IsTrue(objFournisseur.Save())

        'Création de 5 Produits
        objP1 = New Produit("PRD01", objFournisseur, 2000)
        objP1.DateDernInventaire = "01/01/1900"
        Assert.IsTrue(objP1.save())
        objP2 = New Produit("PRD02", objFournisseur, 2000)
        objP1.DateDernInventaire = "01/01/1900"
        Assert.IsTrue(objP2.save())
        objP3 = New Produit("PRD03", objFournisseur, 2000)
        objP3.DateDernInventaire = "01/01/1900"
        Assert.IsTrue(objP3.save())
        objP4 = New Produit("PRD04", objFournisseur, 2000)
        objP4.DateDernInventaire = "01/01/1900"
        Assert.IsTrue(objP4.save())
        objP5 = New Produit("PRD05", objFournisseur, 2000)
        objP5.DateDernInventaire = "01/01/1900"
        Assert.IsTrue(objP5.save())

    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        'Suppression des Produits
        objP1.bDeleted = True
        Assert.IsTrue(objP1.save())
        objP2.bDeleted = True
        Assert.IsTrue(objP2.save())
        objP3.bDeleted = True
        Assert.IsTrue(objP3.save())
        objP4.bDeleted = True
        Assert.IsTrue(objP4.save())
        objP5.bDeleted = True
        Assert.IsTrue(objP5.save())

        'Suppression du Client
        objClient.bDeleted = True
        Assert.IsTrue(objClient.save())

        'Suppression du Fournisseur
        objFournisseur.bDeleted = True
        Assert.IsTrue(objFournisseur.Save())

        m_obj.shared_disconnect()
        MyBase.TestCleanup()

    End Sub
    <TestMethod()> Public Sub T01_Objet()
        Dim objMvt As mvtStock
        Dim objmvt2 As mvtStock
        objMvt = New mvtStock(Now(), objP1.id, vncEnums.vncTypeMvt.vncmvtRegul, +10, "Essai")
        '            Assert.IsTrue(objMvt.datemvt.ToShortDateString.Equals(DATE))
        Assert.IsTrue(objMvt.idProduit.Equals(objP1.id))
        Assert.IsTrue(objMvt.typeMvt.Equals(vncTypeMvt.vncmvtRegul))
        Assert.IsTrue(objMvt.idReference = 0)
        Assert.IsTrue(objMvt.libelle = "Essai")
        Assert.IsTrue(objMvt.qte = 10)
        Assert.IsTrue(objMvt.Commentaire = "")

        'Egal à un semblable
        objmvt2 = New mvtStock(Now(), objP1.id, vncEnums.vncTypeMvt.vncmvtRegul, +10, "Essai")
        Assert.IsTrue(objMvt.Equals(objmvt2))

        'Egal à un Différent
        objMvt.idReference = 132
        objMvt.Commentaire = "Mon Commentaire"
        Assert.IsFalse(objMvt.Equals(objmvt2))

        'Egal à Autrechose
        Dim obj As Object
        Assert.IsFalse(objMvt.Equals(obj), "Egal autrecjhose")
    End Sub

    <TestMethod()> Public Sub T15_DB()
        Dim n As Integer
        Dim obj As mvtStock
        Dim obj2 As mvtStock

        'I - Création d'un objet
        '=========================
        obj = New mvtStock(Now(), objP1.id, vncEnums.vncTypeMvt.vncmvtBonAppro, 10, "Test")
        'obj.idReference = 15
        'Test des indicateurs Avant le Save
        Assert.IsTrue(obj.bNew)
        Assert.IsFalse(obj.bUpdated)
        Assert.IsFalse(obj.bDeleted)
        'Save
        obj.idReference = 15
        Assert.IsTrue(obj.save(), "Insert" & obj.getErreur)
        Assert.IsTrue((obj.id <> 0), "Id Apres le Save doit être différent de 0")

        'Test des indicateurs Après le Save
        Assert.IsFalse(obj.bNew, "bNew apres insert")
        Assert.IsFalse(obj.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(obj.bDeleted, "bDeleted apres insert")

        'II - Rechargement d'un objet
        '=========================
        n = obj.id
        obj2 = New mvtStock(Now(), 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 0, "")
        Assert.IsTrue(obj2.load(n), "Load" & obj2.getErreur())
        Assert.IsTrue(obj.Equals(obj2))

        'III - Modification de l'objet
        '=================================
        ' Modification du Client
        obj2.idProduit = objP2.id
        obj2.typeMvt = vncEnums.vncTypeMvt.vncMvtCommandeClient
        obj2.datemvt = "31/12/2004"
        obj2.idReference = 12
        obj2.libelle = "TEST" & Now()
        obj2.qte = 23
        obj2.Commentaire = "Mon commentaire"

        'Test des indicateurs Avant le Save
        Assert.IsFalse(obj2.bNew)
        Assert.IsTrue(obj2.bUpdated)
        Assert.IsFalse(obj2.bDeleted)
        'Save
        Assert.IsTrue(obj2.save(), "Update" & obj.getErreur)
        'Test des indicateurs Après le Save
        Assert.IsFalse(obj2.bNew, "bNew apres insert")
        Assert.IsFalse(obj2.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(obj2.bDeleted, "bDeleted apres insert")
        'Rechargement de l'objet
        n = obj2.id
        obj = New mvtStock(Now(), 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 0, "")
        Assert.IsTrue(obj.load(n), "Load")
        Assert.IsTrue(obj.Equals(obj2), "Apres Update , Equals")

        'IV - Suppression de l'objet
        '=================================
        obj.bDeleted = True
        Assert.IsTrue(obj.save(), "Delete" & obj.getErreur())
        'Rechargement dans un autre objet
        obj2 = New mvtStock(Now(), 0, vncEnums.vncTypeMvt.vncmvtBonAppro, 0, "")
        Assert.IsFalse(obj2.load(n), "Load")
    End Sub


    <TestMethod()> Public Sub T20_Quantite_en_commande()
        Dim objCMD1 As CommandeClient
        Dim objCMD2 As CommandeClient


        'Création de 2 Commandes
        objCMD1 = New CommandeClient(objClient)
        objCMD2 = New CommandeClient(objClient)
        Persist.executeSQLNonQuery("DELETE * FROM LIGNE_COMMANDE WHERE LGCM_PRD_ID = " & objP1.id)
        Persist.executeSQLNonQuery("DELETE * FROM LIGNE_COMMANDE WHERE LGCM_PRD_ID = " & objP2.id)
        Persist.executeSQLNonQuery("DELETE * FROM LIGNE_COMMANDE WHERE LGCM_PRD_ID = " & objP3.id)
        Persist.executeSQLNonQuery("DELETE * FROM LIGNE_COMMANDE WHERE LGCM_PRD_ID = " & objP4.id)
        '1ere Commande avec P1 et P2
        objCMD1.AjouteLigne("10", objP1, 10, 150)
        objCMD1.AjouteLigne("20", objP2, 15, 150)
        Debug.Assert(objCMD1.save)
        '2eme Commande avec P3 et P4
        objCMD2.AjouteLigne("10", objP3, 20, 150)
        objCMD2.AjouteLigne("20", objP4, 25, 150)
        Debug.Assert(objCMD2.save)
        'controle des Quantités Commandée 
        'P1
        objP1.load()
        Assert.AreEqual(CDec(10), objP1.qteCommande, "Qte Commandée P1")
        'P2
        objP2.load()
        Assert.AreEqual(CDec(15), objP2.qteCommande, "Qte Commandée P2")
        'P3
        objP3.load()
        Assert.AreEqual(CDec(20), objP3.qteCommande, "Qte Commandée P3")
        'P4
        objP4.load()
        Assert.AreEqual(CDec(25), objP4.qteCommande, "Qte Commandée P4")
        'P5
        objP5.load()
        Assert.AreEqual(CDec(0), objP5.qteCommande, "Qte Commandée P5")

        'Validation de CMD1
        objCMD1.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(objCMD1.save())
        'controle des Quantités Commandées
        'P1
        objP1.load()
        Assert.AreEqual(CDec(10), objP1.qteCommande, "Qte Commandée P1")
        'P2
        objP2.load()
        Assert.AreEqual(CDec(15), objP2.qteCommande, "Qte Commandée P2")
        'P3
        objP3.load()
        Assert.AreEqual(CDec(20), objP3.qteCommande, "Qte Commandée P3")
        'P4
        objP4.load()
        Assert.AreEqual(CDec(25), objP4.qteCommande, "Qte Commandée P4")
        'P5
        objP5.load()
        Assert.AreEqual(CDec(0), objP5.qteCommande, "Qte Commandée P5")

        'Livraison de CMD1
        objCMD1.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD1.save())
        'controle des Quantités Commandées
        'P1
        objP1.load()
        Assert.AreEqual(CDec(0), objP1.qteCommande, "Qte Commandée P1")
        'P2
        objP2.load()
        Assert.AreEqual(0D, objP2.qteCommande, "Qte Commandée P2")
        'P3
        objP3.load()
        Assert.AreEqual(20D, objP3.qteCommande, "Qte Commandée P3")
        'P4
        objP4.load()
        Assert.AreEqual(25D, objP4.qteCommande, "Qte Commandée P4")
        'P5
        objP5.load()
        Assert.AreEqual(0D, objP5.qteCommande, "Qte Commandée P5")

        'Suppression de CMD2
        objCMD2.bDeleted = True
        Debug.Assert(objCMD2.save())
        'controle des Quantités Commandées
        'P1
        objP1.load()
        Assert.AreEqual(0D, objP1.qteCommande, "Qte Commandée P1")
        'P2
        objP2.load()
        Assert.AreEqual(0D, objP2.qteCommande, "Qte Commandée P2")
        'P3
        objP3.load()
        Assert.AreEqual(0D, objP3.qteCommande, "Qte Commandée P3")
        'P4
        objP4.load()
        Assert.AreEqual(0D, objP4.qteCommande, "Qte Commandée P4")
        'P5
        objP5.load()
        Assert.AreEqual(0D, objP5.qteCommande, "Qte Commandée P5")

        'Suppression de CMD1
        objCMD1.bDeleted = True
        Debug.Assert(objCMD1.save())
        'controle des Quantités Commandées
        'P1
        objP1.load()
        Assert.AreEqual(0D, objP1.qteCommande, "Qte Commandée P1")
        'P2
        objP2.load()
        Assert.AreEqual(0D, objP2.qteCommande, "Qte Commandée P2")
        'P3
        objP3.load()
        Assert.AreEqual(0D, objP3.qteCommande, "Qte Commandée P3")
        'P4
        objP4.load()
        Assert.AreEqual(0D, objP4.qteCommande, "Qte Commandée P4")
        'P5
        objP5.load()
        Assert.AreEqual(0D, objP5.qteCommande, "Qte Commandée P5")
    End Sub
    <TestMethod()> Public Sub T30_colMvtStock()
        Dim objPRD As Produit
        Dim nid As Long
        Dim objMVT As mvtStock

        'Creation d'un Produit
        objPRD = objP1

        'Chargement de la liste des mouvements de stocks
        Assert.IsTrue(objPRD.loadcolmvtStock(), "loadcolmvtStock" & Produit.getErreur)
        Assert.AreEqual(objPRD.colmvtStock.Count, 0, "Collection non vide")

        'Ajout de 3 mvts des stocks
        Assert.AreEqual(0D, objPRD.QteStock, "Qte en Stock intial")
        Assert.IsTrue(Not objPRD.ajouteLigneMvtStock("06/02/1964", vncEnums.vncTypeMvt.vncmvtRegul, 0, "Ajoute 10", 10, "PremierMVT") Is Nothing, "Ajout mvt1")
        Assert.AreEqual(10D, objPRD.QteStock, "Qte en Stock 1mvt ")

        Assert.IsTrue(Not objPRD.ajouteLigneMvtStock("07/02/1964", vncEnums.vncTypeMvt.vncmvtRegul, 0, "Retire 5", -5, "SecondMVT") Is Nothing, "Ajout mvt2")
        Assert.AreEqual(5D, objPRD.QteStock, "Qte en Stock 2mvt")

        Assert.IsTrue(Not objPRD.ajouteLigneMvtStock("08/02/1964", vncEnums.vncTypeMvt.vncmvtRegul, 0, "Ajoute 3", +3, "TroisièmeMVT") Is Nothing, "Ajout mvt3")
        Assert.AreEqual(8D, objPRD.QteStock, "Qte en Stock 3mvt")

        'Sauvegarde du produit
        Assert.IsTrue(objPRD.save(), "objPRD.Save")
        'Rechargement du prodtui et stock
        nid = objPRD.id
        objPRD = Produit.createandload(nid)
        Assert.IsTrue(objPRD.loadcolmvtStock, "OPRD.LoadcolMvtStock")
        Assert.AreEqual(objPRD.colmvtStock.Count, 3, "colmvtStock.count ")
        Assert.AreEqual(8D, objPRD.QteStock, "Qte en Stock apres rechargement")

        'Ajout d'une ligne (Bon Appro)
        Assert.IsTrue(Not objPRD.ajouteLigneMvtStock("09/02/1964", vncEnums.vncTypeMvt.vncmvtBonAppro, 0, "Bon Appro 20", 20, "4eme MVT") Is Nothing, "Ajout MVT4")
        Assert.AreEqual(28D, objPRD.QteStock, "Qte en Stock 4mvt")
        'Sauvegarde de l'objet
        Assert.IsTrue(objPRD.save(), "PRD.Save")
        'Rechargement du client et de sa precommande
        nid = objPRD.id
        objPRD = Produit.createandload(nid)
        Assert.IsTrue(objPRD.loadcolmvtStock, "OPRD.LoadcolMvtStock")
        Assert.AreEqual(objPRD.colmvtStock.Count, 4, "colmvtStock.count ")

        'Suppression d'un mvt de stock (Retire 5)
        objPRD.supprimeLigneMvtStock(2)
        Assert.AreEqual(33D, objPRD.QteStock, "Qte en Stock Apres supreesio de ligne")
        'Sauvegarde de l'objet
        Assert.IsTrue(objPRD.save(), "PRD.Save")
        'Rechargement de l'objet
        nid = objPRD.id
        objPRD = Produit.createandload(nid)
        Assert.IsTrue(objPRD.loadcolmvtStock, "OPRD.LoadcolMvtStock")
        Assert.AreEqual(objPRD.colmvtStock.Count, 3, "colmvtStock.count ")

        'Maj d'un Mouvement de stock
        'Attention les mouvement dont chargé dans l'odre décroissant des dates
        objMVT = objPRD.colmvtStock(1) '(Ajoute 3)
        objMVT.qte = 11
        objMVT.libelle = "Remplace par 11"
        objMVT.save()
        objPRD.recalculStock()
        Assert.AreEqual(41D, objPRD.QteStock, "Qte en Stock Apres remplacement")

        'Sauvegarde de l'objet
        Assert.IsTrue(objPRD.save(), "PRD.Save")
        'Rechargement de l'objet
        nid = objPRD.id
        objPRD = Produit.createandload(nid)
        Assert.IsTrue(objPRD.load(nid), "PRD.load")
        Assert.IsTrue(objPRD.loadcolmvtStock, "OPRD.LoadcolMvtStock")
        Assert.AreEqual(objPRD.colmvtStock.Count, 3, "colmvtStock.count ")
        objMVT = objPRD.colmvtStock(1)
        Assert.AreEqual(11D, objMVT.qte)

        'Suppression du Produit 
        objPRD.bDeleted = True
        nid = objPRD.id
        Assert.IsTrue(objPRD.save(), "objPRD.Delete")
        Persist.shared_connect()
        Dim strReponse As String
        strReponse = Persist.executeSQLQuery("SELECT COUNT(*) FROM MVT_STOCK WHERE STK_PRD_ID =" + nid.ToString())
        Assert.AreEqual(0, CInt(strReponse))

    End Sub
    <TestMethod()> Public Sub T34_GenerationMvtStock()
        Dim objCmd As CommandeClient
        Dim objBA As BonAppro
        Dim nidCmd As Integer
        Dim nidBA As Integer

        'MAJ inventaire au 01/01/2004 P1
        objP1.ajouteLigneMvtStock("01/01/2004", vncEnums.vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/01/2004", 10)
        Assert.IsTrue(objP1.save(), "Save P1")
        'MAJ inventaire au 01/01/2004 P2
        objP2.ajouteLigneMvtStock("01/01/2004", vncEnums.vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/01/2004", 20)
        Assert.IsTrue(objP2.save(), "Save P2")
        'MAJ inventaire au 01/01/2004 P3
        objP3.ajouteLigneMvtStock("01/01/2004", vncEnums.vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/01/2004", 30)
        Assert.IsTrue(objP3.save(), "Save P3")

        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(10D, objP1.QteStock, "P1 Qte en Stock")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(20D, objP2.QteStock, "P2 Qte en Stock")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(30D, objP3.QteStock, "P3 Qte en Stock")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande")

        'Création d'une Commande
        objCmd = New CommandeClient(objClient)
        'Ajout D'une Ligne P1
        objCmd.AjouteLigne("10", objP1, 11, 120)
        'Ajout D'une ligne P2
        objCmd.AjouteLigne("20", objP2, 22, 120)


        Assert.IsTrue(objCmd.colLignes.Count <> 0, "ColScmd (1)-1")
        Assert.IsTrue(objCmd.save(), "Sauvegarde Commande(1)")
        Assert.IsTrue(objCmd.colLignes.Count <> 0, "ColScmd (1)-2")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.IsTrue(objP1.load(), "Chargement P1(1)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(1)")
        Assert.IsTrue(objP1.colmvtStock.Count = 1, "objP1.colmvtStock.Count = 2")
        Assert.IsTrue(objP2.load(), "Chargement P1(1)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P1(1)")
        Assert.AreEqual(10D, objP1.QteStock, "P1 Qte en Stock(1)")
        Assert.AreEqual(11D, objP1.qteCommande, "P1 Qte en commande(1)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(20D, objP2.QteStock, "P2 Qte en Stock(1)")
        Assert.AreEqual(22D, objP2.qteCommande, "P2 Qte en commande(1)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(30D, objP3.QteStock, "P3 Qte en Stock(1)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(1)")

        'Validation de Commande
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(objCmd.colLignes.Count <> 0)
        Assert.IsTrue(objCmd.getActionMvtStock() = vncEnums.vncGenererSupprimer.vncRien)
        Assert.IsTrue(objCmd.save(), "Save Commande (2)")
        Assert.IsTrue(objP1.load(), "Chargement P1(2)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(2)")
        Assert.IsTrue(objP1.colmvtStock.Count = 1, "objP1.colmvtStock.Count = 1")
        Assert.IsTrue(objP2.load(), "Chargement P2(2)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(2)")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(10D, objP1.QteStock, "P1 Qte en Stock(2)")
        Assert.AreEqual(11D, objP1.qteCommande, "P1 Qte en commande(2)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(20D, objP2.QteStock, "P2 Qte en Stock(2)")
        Assert.AreEqual(22D, objP2.qteCommande, "P2 Qte en commande(2)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(30D, objP3.QteStock, "P3 Qte en Stock(2)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(2)")

        'Livraison Commande
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCmd.colLignes.Count <> 0)
        CType(objCmd.colLignes(1), LgCommande).qteLiv = 11
        CType(objCmd.colLignes(2), LgCommande).qteLiv = 22
        objCmd.dateLivraison = "01/05/2004"
        Assert.IsTrue(objCmd.getActionMvtStock() = vncEnums.vncGenererSupprimer.vncGenerer)
        Assert.IsTrue(objCmd.save(), "Save Commande (3)")
        Assert.IsTrue(objP1.load(), "Chargement P1(3)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(3)")
        Assert.IsTrue(objP1.colmvtStock.Count = 2, "objP1.colmvtStock.Count = 2")
        Assert.IsTrue(objP2.load(), "Chargement P2(3)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(3)")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(-1D, objP1.QteStock, "P1 Qte en Stock(3)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(3)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(3)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(3)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(30D, objP3.QteStock, "P3 Qte en Stock(3)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(3)")

        'Annulation de Livraison de commande
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnLivrer)
        Assert.IsTrue(objCmd.getActionMvtStock() = vncEnums.vncGenererSupprimer.vncSupprimer)
        Assert.IsTrue(objCmd.save(), "Save Commande (4)")
        Assert.IsTrue(objP1.load(), "Chargement P1(4)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(4)")
        Assert.IsTrue(objP2.load(), "Chargement P2(4)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(4)")
        Assert.IsTrue(objP1.colmvtStock.Count = 1, "objP1.colmvtStock.Count = 1")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(10D, objP1.QteStock, "P1 Qte en Stock(4)")
        Assert.AreEqual(11D, objP1.qteCommande, "P1 Qte en commande(4)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(20D, objP2.QteStock, "P2 Qte en Stock(4)")
        Assert.AreEqual(22D, objP2.qteCommande, "P2 Qte en commande(4)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(30D, objP3.QteStock, "P3 Qte en Stock(4)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(4)")
        'Assert.IsTrue(MsgBox("Vérifier que'il n'y a plus que des mvt d'INVENTAIRE(1) dans MVTSTOCK pour L'id Produit " & objP1.id & " et " & objP2.id, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

        'Relivraison de la commande
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCmd.save(), "Save Commande(5)")
        Assert.IsTrue(objP1.load(), "Chargement P1(5)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(5)")
        Assert.IsTrue(objP2.load(), "Chargement P2(5)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(5)")
        Assert.IsTrue(objP1.colmvtStock.Count = 2, "objP1.colmvtStock.Count = 2")
        Assert.IsTrue(objP2.colmvtStock.Count = 2, "objP2.colmvtStock.Count = 2")

        'Creation d'un Bon D'appro
        objBA = New BonAppro(objFournisseur)
        objBA.AjouteLigne("10", objP1, 12, 10)
        objBA.AjouteLigne("20", objP3, 32, 10)
        Assert.IsTrue(objBA.save(), "BA.Save")

        Assert.IsTrue(objP1.load(), "Chargement P1(6)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(6)")
        Assert.IsTrue(objP2.load(), "Chargement P2(6)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(6)")
        Assert.IsTrue(objP3.load(), "Chargement P3(6)")
        Assert.IsTrue(objP3.loadcolmvtStock(), "Chargement Stock P3(6)")
        Assert.IsTrue(objP1.colmvtStock.Count = 2, "objP1.colmvtStock.Count = 2")
        Assert.IsTrue(objP2.colmvtStock.Count = 2, "objP2.colmvtStock.Count = 2")
        Assert.IsTrue(objP3.colmvtStock.Count = 1, "objP3.colmvtStock.Count = 1")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(-1D, objP1.QteStock, "P1 Qte en Stock(6)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(6)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(6)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(6)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(30D, objP3.QteStock, "P3 Qte en Stock(6)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(6)")

        'Livraison du Bon D'appro
        Assert.IsTrue(objBA.colLignes.Count <> 0)
        CType(objBA.colLignes(1), LgCommande).qteLiv = 12
        CType(objBA.colLignes(2), LgCommande).qteLiv = 32
        objBA.dateLivraison = "10/05/2004"
        objBA.changeEtat(vncEnums.vncActionEtatCommande.vncActionBALivrer)
        Assert.IsTrue(objBA.save())
        Assert.IsTrue(objP1.load(), "Chargement P1(6)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(6)")
        Assert.IsTrue(objP2.load(), "Chargement P2(6)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(6)")
        Assert.IsTrue(objP3.load(), "Chargement P3(6)")
        Assert.IsTrue(objP3.loadcolmvtStock(), "Chargement Stock P3(6)")
        Assert.IsTrue(objP1.colmvtStock.Count = 3, "objP1.colmvtStock.Count = 3")
        Assert.IsTrue(objP2.colmvtStock.Count = 2, "objP2.colmvtStock.Count = 2")
        Assert.IsTrue(objP3.colmvtStock.Count = 2, "objP3.colmvtStock.Count = 2")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(-1 + 12), objP1.QteStock, "P1 Qte en Stock(3)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(3)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(3)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(3)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(30 + 32), objP3.QteStock, "P3 Qte en Stock(3)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(3)")

        'Annulation de la livriason du bon D'appro
        objBA.changeEtat(vncEnums.vncActionEtatCommande.vncActionBAAnnLivrer)
        Assert.IsTrue(objBA.save())
        Assert.IsTrue(objP1.load(), "Chargement P1(6)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(6)")
        Assert.IsTrue(objP2.load(), "Chargement P2(6)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(6)")
        Assert.IsTrue(objP3.load(), "Chargement P3(6)")
        Assert.IsTrue(objP3.loadcolmvtStock(), "Chargement Stock P3(6)")
        Assert.AreEqual(2, objP1.colmvtStock.Count, "objP1.colmvtStock.Count = 2")
        Assert.AreEqual(2, objP2.colmvtStock.Count, "objP2.colmvtStock.Count = 2")
        Assert.AreEqual(1, objP3.colmvtStock.Count, "objP3.colmvtStock.Count = 1")

        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(-1D, objP1.QteStock, "P1 Qte en Stock(3)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(3)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(3)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(3)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(30D, objP3.QteStock, "P3 Qte en Stock(3)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(3)")

        'Relivraison du BOnAppro
        objBA.changeEtat(vncEnums.vncActionEtatCommande.vncActionBALivrer)
        Assert.IsTrue(objBA.save())
        Assert.IsTrue(objP1.load(), "Chargement P1(6)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(6)")
        Assert.IsTrue(objP2.load(), "Chargement P2(6)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(6)")
        Assert.IsTrue(objP3.load(), "Chargement P3(6)")
        Assert.IsTrue(objP3.loadcolmvtStock(), "Chargement Stock P3(6)")
        Assert.IsTrue(objP1.colmvtStock.Count = 3, "objP1.colmvtStock.Count = 3")
        Assert.IsTrue(objP2.colmvtStock.Count = 2, "objP2.colmvtStock.Count = 2")
        Assert.IsTrue(objP3.colmvtStock.Count = 2, "objP3.colmvtStock.Count = 2")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(-1 + 12), objP1.QteStock, "P1 Qte en Stock(3)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(3)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(3)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(3)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(30 + 32), objP3.QteStock, "P3 Qte en Stock(3)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(3)")

        'Suppression de Commande
        nidCmd = objCmd.id
        objCmd.bDeleted = True
        Assert.IsTrue(objCmd.save(), "delete CMD")

        'Cela ne change rien au stock
        Assert.IsTrue(objP1.load(), "Chargement P1(6)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(6)")
        Assert.IsTrue(objP2.load(), "Chargement P2(6)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(6)")
        Assert.IsTrue(objP3.load(), "Chargement P3(6)")
        Assert.IsTrue(objP3.loadcolmvtStock(), "Chargement Stock P3(6)")
        Assert.IsTrue(objP1.colmvtStock.Count = 3, "objP1.colmvtStock.Count = 3")
        Assert.IsTrue(objP2.colmvtStock.Count = 2, "objP2.colmvtStock.Count = 2")
        Assert.IsTrue(objP3.colmvtStock.Count = 2, "objP3.colmvtStock.Count = 2")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(-1 + 12), objP1.QteStock, "P1 Qte en Stock(3)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(3)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(3)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(3)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(30 + 32), objP3.QteStock, "P3 Qte en Stock(3)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(3)")

        'Suppression du  BA
        nidBA = objBA.id
        objBA.bDeleted = True
        Assert.IsTrue(objBA.save(), "delete BA")
        'Cela ne change rien au stock
        Assert.IsTrue(objP1.load(), "Chargement P1(6)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(6)")
        Assert.IsTrue(objP2.load(), "Chargement P2(6)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(6)")
        Assert.IsTrue(objP3.load(), "Chargement P3(6)")
        Assert.IsTrue(objP3.loadcolmvtStock(), "Chargement Stock P3(6)")
        Assert.IsTrue(objP1.colmvtStock.Count = 3, "objP1.colmvtStock.Count = 3")
        Assert.IsTrue(objP2.colmvtStock.Count = 2, "objP2.colmvtStock.Count = 2")
        Assert.IsTrue(objP3.colmvtStock.Count = 2, "objP3.colmvtStock.Count = 2")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(-1 + 12), objP1.QteStock, "P1 Qte en Stock(3)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(3)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(3)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(3)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(30 + 32), objP3.QteStock, "P3 Qte en Stock(3)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(3)")

        Dim strReponse As String
        strReponse = Persist.executeSQLQuery("SELECT count(*) from mvt_stock where STK_REF_ID = " + nidCmd.ToString() + " or STK_REF_ID = " + nidBA.ToString())

        Assert.IsTrue(CInt(strReponse) > 0)

        'Création d'une Commande Directe (Pas de Mvt de Stocks)
        objCmd = New CommandeClient(objClient)
        objCmd.typeCommande = vncEnums.vncTypeCommande.vncCmdClientDirecte
        'Ajout D'une Ligne P1
        objCmd.AjouteLigne("10", objP1, 11, 120)
        'Ajout D'une ligne P2
        objCmd.AjouteLigne("20", objP2, 22, 120)
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCmd.save())
        'Cela ne change rien au stock
        Assert.IsTrue(objP1.load(), "Chargement P1(6)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(6)")
        Assert.IsTrue(objP2.load(), "Chargement P2(6)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(6)")
        Assert.IsTrue(objP3.load(), "Chargement P3(6)")
        Assert.IsTrue(objP3.loadcolmvtStock(), "Chargement Stock P3(6)")
        Assert.IsTrue(objP1.colmvtStock.Count = 3, "objP1.colmvtStock.Count = 3")
        Assert.IsTrue(objP2.colmvtStock.Count = 2, "objP2.colmvtStock.Count = 2")
        Assert.IsTrue(objP3.colmvtStock.Count = 2, "objP3.colmvtStock.Count = 2")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(-1 + 12), objP1.QteStock, "P1 Qte en Stock(3)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(3)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(3)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(3)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(30 + 32), objP3.QteStock, "P3 Qte en Stock(3)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(3)")

        'Suppression de la commande directe
        objCmd.bDeleted = True
        Assert.IsTrue(objCmd.save())
        'Cela ne change rien au stock
        Assert.IsTrue(objP1.load(), "Chargement P1(6)")
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement Stock P1(6)")
        Assert.IsTrue(objP2.load(), "Chargement P2(6)")
        Assert.IsTrue(objP2.loadcolmvtStock(), "Chargement Stock P2(6)")
        Assert.IsTrue(objP3.load(), "Chargement P3(6)")
        Assert.IsTrue(objP3.loadcolmvtStock(), "Chargement Stock P3(6)")
        Assert.IsTrue(objP1.colmvtStock.Count = 3, "objP1.colmvtStock.Count = 3")
        Assert.IsTrue(objP2.colmvtStock.Count = 2, "objP2.colmvtStock.Count = 2")
        Assert.IsTrue(objP3.colmvtStock.Count = 2, "objP3.colmvtStock.Count = 2")
        'Controle Stock P1 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(-1 + 12), objP1.QteStock, "P1 Qte en Stock(3)")
        Assert.AreEqual(0D, objP1.qteCommande, "P1 Qte en commande(3)")
        'Controle Stock P2 (Qte en stock, Qte en commande)
        Assert.AreEqual(-2D, objP2.QteStock, "P2 Qte en Stock(3)")
        Assert.AreEqual(0D, objP2.qteCommande, "P2 Qte en commande(3)")
        'Controle Stock P3 (Qte en stock, Qte en commande)
        Assert.AreEqual(CDec(30 + 32), objP3.QteStock, "P3 Qte en Stock(3)")
        Assert.AreEqual(0D, objP3.qteCommande, "P3 Qte en commande(3)")

    End Sub
    <TestMethod()> Public Sub T40_TestMvtStock()

        Dim objCommande As CommandeClient
        Dim colMvtStock As ColEventSorted
        Dim colResult As Collection
        Dim str As String
        Dim objMvtStock As mvtStock
        Dim objLgCmd As LgCommande

        'Création d'une commande
        objCommande = New CommandeClient(objClient)

        'Ajout de 2 lignes sur cette commande
        objCommande.AjouteLigne("1", objP1, 10, 10)
        objCommande.AjouteLigne("2", objP2, 20, 20)
        objCommande.AjouteLigne("3", objP2, 30, 30)
        objCommande.AjouteLigne("4", objP3, 40, 40)
        'Sauvegarde de la commande
        Assert.IsTrue(objCommande.save())

        'Génération des mvts de stocks
        objCommande.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCommande.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        For Each objLgCmd In objCommande.colLignes
            objLgCmd.qteLiv = objLgCmd.qteCommande + 1
        Next
        Assert.IsTrue(objCommande.save())
        colMvtStock = mvtStock.getListe(objP1.id, objCommande.id)
        Assert.IsTrue(colMvtStock.Count = 1, "1 Mvt de stock pour P1")
        colMvtStock = mvtStock.getListe(objP2.id, objCommande.id)
        Assert.IsTrue(colMvtStock.Count = 2, "2 Mvt de stock pour P2")

        'Controle des mvts => OK
        colResult = objCommande.controleMvtStock()
        For Each str In colResult
            'Console.Out.WriteLine(str)
        Next
        Assert.IsTrue(colResult.Count = 0, "Pas D'erreur")

        'Supresssion d'une ligne de mvts de stock
        colMvtStock = mvtStock.getListe(objP2.id, objCommande.id)
        objMvtStock = colMvtStock(0)
        objMvtStock.bDeleted = True
        Assert.IsTrue(objMvtStock.save())
        colMvtStock = mvtStock.getListe(objP2.id, objCommande.id)
        Assert.IsTrue(colMvtStock.Count = 1, "plus qu'un seul Mvts de stock pour P2")

        'controle des mvts de Stock =>NOK
        colResult = objCommande.controleMvtStock()
        For Each str In colResult
            'Console.Out.WriteLine(str)
        Next
        Assert.IsTrue(colResult.Count = 1, "1 erreur après la suppression d'un mouvement de stock")

        'Anulation de la livraison
        objCommande.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnLivrer)
        Assert.IsTrue(objCommande.save())
        'Livraison de la commande
        objCommande.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCommande.save())

        'Controle des mvts de stocks OK
        colResult = objCommande.controleMvtStock()
        For Each str In colResult
            'Console.Out.WriteLine(str)
        Next
        Assert.IsTrue(colResult.Count = 0, "Pas D'erreur")

        'Ajout d'une nlle ligne de mvts de stocks avec le produit1 et la même quantité
        Assert.IsTrue(objP1.loadcolmvtStock(), "Chargement de la collection des mouvements de stocks")
        objP1.ajouteLigneMvtStock(objCommande.dateLivraison, vncEnums.vncTypeMvt.vncMvtCommandeClient, objCommande.id, "Ajout manuel", 1, "Ajout Manuel")
        Assert.IsTrue(objP1.save())

        'Controle des mvts de stocks => NOK
        colResult = objCommande.controleMvtStock()
        For Each str In colResult
            'Console.Out.WriteLine(str)
        Next
        Assert.IsTrue(colResult.Count = 1, "1 erreur après l'ajout d'un mouvement de stock")

        'Suppression de la commande
        objCommande.bDeleted = True
        Assert.IsTrue(objCommande.save())

    End Sub

    <TestMethod()> Public Sub T50_RecalculduStock()
        Dim nid As Integer

        'Tous les mvts sont à la même date 
        objP1.ajouteLigneMvtStock("01/02/2005", vncEnums.vncTypeMvt.vncMvtInventaire, 0, "A NOUVEAU", 12)
        objP1.ajouteLigneMvtStock("01/02/2005", vncEnums.vncTypeMvt.vncMvtCommandeClient, 0, "CMD CLIENT", -2)
        objP1.ajouteLigneMvtStock("01/02/2005", vncEnums.vncTypeMvt.vncmvtBonAppro, 0, "BON APPO", +10)
        objP1.save()

        nid = objP1.id

        objP1 = Produit.createandload(nid)
        objP1.loadcolmvtStock()
        objP1.recalculStock()

        Assert.AreEqual(objP1.QteStock, 20D, "Qte En stock")

    End Sub


    <TestMethod()> Public Sub T60_RegenerationDesMvtsdeStocks()
        Dim nidPRD As Long
        Dim nidCMD As Long
        Dim objPRD As Produit
        Dim objCommande As CommandeClient
        Dim objMVTStk As mvtStock
        Dim strCodeCmd As String

        objPRD = New Produit("PRDT60_RGMSTK", objFournisseur, 1990)
        Assert.IsTrue(objPRD.save(), "Sauvegarde du produit")
        nidPRD = objPRD.id

        'Creation d'une commande
        objCommande = New CommandeClient(objClient)
        objCommande.dateCommande = "10/05/2005"
        objCommande.AjouteLigne("10", objPRD, 10, 1.5)
        Assert.IsTrue(objCommande.save(), "Sauvegarde de la commande")
        nidCMD = objCommande.id
        strCodeCmd = objCommande.code

        'Livraison de la commande
        objCommande.dateLivraison = "15/05/2005"
        objCommande.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCommande.save(), "Sauvegarde de la Livraison de commande")

        'Lecture des mouvements de stocks
        Assert.IsTrue(objPRD.loadcolmvtStock(), "Chrgmt des mvts de stocks")
        Assert.IsTrue(objPRD.colmvtStock.Count = 1, "1 mouvement de stock")

        'chgmt de la date de mvt de stock
        objMVTStk = objPRD.colmvtStock(0)
        objMVTStk.datemvt = "31/05/2005"
        Assert.IsTrue(objMVTStk.save(), "Sauvegarde du Mouvement de stock")


        objCommande = CommandeClient.getListe(strCodeCmd, pOrigine:="")(1)
        objCommande.load()
        objCommande.regenereMvtStock()

        objPRD = Produit.createandload(nidPRD)
        objPRD.loadcolmvtStock()
        Assert.IsTrue(objPRD.colmvtStock.Count = 1, "1 mouvement de stock")
        objMVTStk = objPRD.colmvtStock(0)
        Assert.IsTrue(objMVTStk.datemvt.Equals(objCommande.dateLivraison), "Date de Mvt = Date de livraison")



        objCommande = CommandeClient.createandload(nidCMD)
        objCommande.bDeleted = True
        objCommande.save()
        objPRD = Produit.createandload(nidPRD)
        objPRD.bDeleted = True
        objPRD.save()

    End Sub

    <TestMethod(), Ignore()> Public Sub T10_ContoleUniciteMvtStk()
        Dim oCMD As CommandeClient

        oCMD = New CommandeClient(objClient)
        'Ajout de 2 lignes sur cette commande
        oCMD.AjouteLigne("1", objP1, 10, 10)
        oCMD.AjouteLigne("2", objP2, 20, 20)
        oCMD.AjouteLigne("3", objP2, 30, 30)
        oCMD.AjouteLigne("4", objP3, 40, 40)
        'Sauvegarde de la commande
        Assert.IsTrue(oCMD.save())

        'Génération des mvts de stocks
        oCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        oCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        'Sauvegarde de la commande
        Assert.IsTrue(oCMD.etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer, "Sauve = Générer mvtStock")
        Assert.IsTrue(oCMD.save())
        Assert.IsTrue(oCMD.etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncRien, "Sauve = Générer mvtStock")
        'Je repasse le code à générer 
        Console.WriteLine("Les Messages d'ASSERT suivants sont NORMAUX")
        Console.WriteLine("====================================================================")
        oCMD.etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer
        Assert.IsFalse(oCMD.save(), "La sauvegarde de la commande doit échouée")
        Console.WriteLine("====================================================================")

        'Suppression de la commande
        oCMD.bDeleted = True
        Assert.IsTrue(oCMD.save())




    End Sub
    <TestMethod()> Public Sub T60_Champslongs()

        Dim oMvt As mvtStock
        objP1.ajouteLigneMvtStock("06/02/1964", vncTypeMvt.vncmvtRegul, 150, "x".PadRight(150, "x"), 10, "y".PadRight(500, "x"))
        objP1.savecolmvtStock()

        objP1.releasecolmvtStock()
        objP1.loadcolmvtStock()
        oMvt = objP1.colmvtStock(0)
        Assert.AreEqual(50, oMvt.libelle.Length)
        Assert.AreEqual(255, oMvt.Commentaire.Length)

    End Sub

    <TestMethod()> Public Sub T70_GenerationMvtInventaire()

        objP1.ajouteLigneMvtStock("01/01/2001", vncTypeMvt.vncMvtInventaire, 0, "Iventaire au 01/1", 10)
        objP1.ajouteLigneMvtStock("10/01/2001", vncTypeMvt.vncMvtCommandeClient, 0, "CMD 10/1", -5)
        objP1.ajouteLigneMvtStock("20/01/2001", vncTypeMvt.vncmvtBonAppro, 0, "CMD 20/1", +15)
        objP1.ajouteLigneMvtStock("01/02/2001", vncTypeMvt.vncMvtCommandeClient, 0, "CMD 01/02", -1)
        objP1.ajouteLigneMvtStock("05/02/2001", vncTypeMvt.vncMvtCommandeClient, 0, "CMD 05/02", -7)
        objP1.savecolmvtStock()

        Assert.AreEqual(CDec(10 - 5 + 15 - 7 - 1), objP1.QteStock)
        Assert.AreEqual(CDec(10 - 5 + 15), objP1.CalculeStockAu("01/02/2001"))
        objP1.genereMvtInventaire("01/02/2001", objP1.CalculeStockAu("01/02/01"))
        Assert.AreEqual(CDec(10 - 5 + 15 - 7 - 1), objP1.QteStock)
        Assert.AreEqual(CDbl(10 - 5 + 15), objP1.QteStockDernInventaire)


    End Sub


    <TestMethod(), TestCategory("5.9.3")> Public Sub TestGenerationMvtStockHobivin()
        Dim nidPRD As Long
        Dim nidCMD As Long
        Dim objPRD As Produit
        Dim objCommande As CommandeClient
        Dim objMVTStk As mvtStock
        Dim strCodeCmd As String

        Dim oInter As Client
        oInter = Client.getIntermediairePourUneOrigine(Dossier.HOBIVIN)
        If oInter Is Nothing Then
            oInter = New Client
            oInter.setTypeIntermediaire(Dossier.HOBIVIN)
            oInter.rs = "INTER HOBIVIN"
            Assert.IsTrue(oInter.save(), "Sauvegarde du client intermédiaire")
        End If

        objPRD = New Produit("PRDT60_RGMSTK", objFournisseur, 1990)
        Assert.IsTrue(objPRD.save(), "Sauvegarde du produit")
        nidPRD = objPRD.id

        'Creation d'une commande
        objCommande = New CommandeClient(objClient)
        objCommande.Origine = Dossier.HOBIVIN
        objCommande.dateCommande = "10/05/2005"
        objCommande.AjouteLigne("10", objPRD, 10, 1.5)
        Assert.IsTrue(objCommande.save(), "Sauvegarde de la commande")
        nidCMD = objCommande.id
        strCodeCmd = objCommande.code

        'Livraison de la commande
        objCommande.dateLivraison = "15/05/2005"
        objCommande.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCommande.save(), "Sauvegarde de la Livraison de commande")

        'Lecture des mouvements de stocks
        Assert.IsTrue(objPRD.loadcolmvtStock(), "Chrgmt des mvts de stocks")
        Assert.IsTrue(objPRD.colmvtStock.Count = 1, "1 mouvement de stock")

        'chgmt de la date de mvt de stock
        objMVTStk = objPRD.colmvtStock(0)
        Assert.AreEqual("CMD " & objCommande.code & " - " & "10" & " " & oInter.rs, objMVTStk.libelle)

        objCommande = CommandeClient.createandload(nidCMD)
        objCommande.bDeleted = True
        objCommande.save()

        objPRD = Produit.createandload(nidPRD)
        objPRD.bDeleted = True
        objPRD.save()


    End Sub

End Class



