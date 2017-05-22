'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB





<TestClass()> Public Class test_V197
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objCLT As Client
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Persist.shared_connect()
        m_objFRN = New Fournisseur("FRNV197_T10", "FRn de test")
        Assert.IsTrue(m_objFRN.Save(), "FRN.Create")

        m_objPRD = New Produit("PRD197_T10", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objCLT = New Client("CLT001", "Client de test")
        m_objCLT.rs = "Client de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")

        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        Persist.shared_connect()
        m_objPRD.bDeleted = True
        Assert.IsTrue(m_objPRD.save(), "TestCleanup")

        m_objFRN.bDeleted = True
        Assert.IsTrue(m_objFRN.Save())

        m_objCLT.bDeleted = True
        Assert.IsTrue(m_objCLT.save())

        Persist.shared_disconnect()
        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_RecalculduStock()
        'Test pour vérifier que les Mvts de stocks qui ont la même date que l'inventaire soit bien pris en compte.
        Dim nid As Integer

        'Tous les mvts sont à la même date 
        m_objPRD.ajouteLigneMvtStock("01/02/2005", vncEnums.vncTypeMvt.vncMvtInventaire, 0, "A NOUVEAU", 12)
        m_objPRD.ajouteLigneMvtStock("01/02/2005", vncEnums.vncTypeMvt.vncMvtCommandeClient, 0, "CMD CLIENT", -2)
        m_objPRD.ajouteLigneMvtStock("01/02/2005", vncEnums.vncTypeMvt.vncmvtBonAppro, 0, "BON APPO", +10)
        m_objPRD.save()

        nid = m_objPRD.id

        m_objPRD = Produit.createandload(nid)
        m_objPRD.loadcolmvtStock()
        m_objPRD.recalculStock()

        Assert.AreEqual(m_objPRD.QteStock, 20, "Qte En stock")

    End Sub
    <TestMethod()> Public Sub T20_SauvegardePrecommande()
        Dim objPRD As Produit
        Dim objLGPrecom As lgPrecomm
        Dim nid As Integer
        Dim nidLg As Integer

        objPRD = New Produit("PRD197_T20", m_objFRN, 1990)
        Assert.IsTrue(objPRD.save(), "Sauvegarde Produit")
        Assert.IsTrue(m_objCLT.LoadPreCommande(), "LoadPreommande()")
        m_objCLT.ajouteLgPrecom(m_objPRD, 10, 12, 12.5)
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde client")

        nid = m_objCLT.id
        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        nidLg = m_objCLT.getLgPrecom(1).id

        'Mise à jour d'uine ligne => update
        objLGPrecom = m_objCLT.getLgPrecom(1)
        objLGPrecom.refDerniereCommande = "DERNCMD"
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde client")
        nid = m_objCLT.id

        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(nidLg, m_objCLT.getLgPrecom(1).id, "L'id ne change pas")
        Assert.AreEqual(m_objCLT.getlgPrecomCount, 1, "Nbre le ligne de precommande")

        'Ajout D'une ligne => insert delete
        m_objCLT.ajouteLgPrecom(objPRD, 11, 21, 11.5)
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde client")
        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        Assert.IsFalse(nidLg = m_objCLT.getLgPrecom(1).id, "L'id Change")
        Assert.AreEqual(m_objCLT.getlgPrecomCount, 2, "Nbre le ligne de precommande")
        nidLg = m_objCLT.getLgPrecom(1).id

        'Mise à jour d'une ligne => UPDATE
        nidLg = m_objCLT.getLgPrecom(1).id
        Assert.IsTrue(nidLg <> 0)
        objLGPrecom = m_objCLT.getLgPrecom(1)
        objLGPrecom.refDerniereCommande = "UPDATE"
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde client")
        nid = m_objCLT.id
        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(nidLg, m_objCLT.getLgPrecom(1).id, "L'id ne change pas")

        'Suppression d'un ligne => INSERT-DELETE
        m_objCLT.supprimeLgPrecom(2)
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde client")
        nid = m_objCLT.id
        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        Assert.IsFalse(nidLg = m_objCLT.getLgPrecom(1).id, "L'id change ")



        objPRD.bDeleted = True
        Assert.IsTrue(objPRD.save(), "Suppression produit")

    End Sub
    <TestMethod()> Public Sub T30_ReinitPrecommande()
        Dim objPRD1 As Produit
        Dim objPRD2 As Produit
        Dim objCommande As CommandeClient
        Dim nid As Integer

        objPRD1 = New Produit("PRD1", m_objFRN, 1999)
        Assert.IsTrue(objPRD1.save, "Sauvegarde du produit1")
        objPRD2 = New Produit("PRD2", m_objFRN, 1999)
        Assert.IsTrue(objPRD2.save, "Sauvegarde du produit2")


        objCommande = New CommandeClient(m_objCLT)
        objCommande.AjouteLigne("10", m_objPRD, 1, 1)
        Assert.IsTrue(objCommande.save, "Sauvegarde Commande")


        'Controle de la precommande du client
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(1, m_objCLT.getlgPrecomCount(), "1 ligne en Precommande")

        'Ajout d'une ligne de precommande
        m_objCLT.ajouteLgPrecom(objPRD1, 10, 2)
        'Sauvegarde du client
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde Client")
        nid = m_objCLT.id
        'Rechargement du client
        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        'Controle de la precommande du client
        Assert.AreEqual(2, m_objCLT.getlgPrecomCount(), "2 lignes en Precommande")


        'Reinitialisation du client
        Assert.IsTrue(m_objCLT.reinitPrecommande(), "Reinit Precommande")
        'Sauvegarde du client
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde Client")
        nid = m_objCLT.id
        'Rechargement du client
        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        'Controle de la precommande du client
        Assert.AreEqual(1, m_objCLT.getlgPrecomCount(), "1 ligne en Precommande")

        objCommande.bDeleted = True
        Assert.IsTrue(objCommande.save, "Suppression Commande")

        objPRD1.bDeleted = True
        Assert.IsTrue(objPRD1.save(), "Suppression PRD1")
        objPRD2.bDeleted = True
        Assert.IsTrue(objPRD2.save(), "Suppression PRD2")

    End Sub
End Class


