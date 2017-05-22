'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V201
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objCLT As Client
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Persist.shared_connect()
        m_objFRN = New Fournisseur("FRNV200", "FRn de'' test")
        Assert.IsTrue(m_objFRN.Save(), "FRN.Create")

        m_objPRD = New Produit("PRD200", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objCLT = New Client("CLT200", "Client de' test")
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
    <TestMethod()> Public Sub T10_LIBMVTSTK()
        Dim objCmd As CommandeClient
        Dim objLgCmd As LgCommande

        objCmd = New CommandeClient(m_objCLT)
        objCmd.AjouteLigne("10", m_objPRD, 10, 10)
        objCmd.bFactTransport = True
        Assert.IsTrue(objCmd.save(), "Création de la commande")
        ' Livraison de la machandise
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        'Creation des Mouvement s de stock
        objLgCmd = objCmd.colLignes(1)
        objLgCmd.qteLiv = 10
        Assert.IsTrue(objCmd.save(), "Création des mouvements de stocks")

        objCmd.bDeleted = True
        Assert.IsTrue(objCmd.save(), "Suprression de la commande")








    End Sub

End Class


