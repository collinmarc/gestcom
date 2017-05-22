'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V212
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objCLT As Client
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Persist.shared_connect()
        m_objFRN = New Fournisseur("FRNV210", "FRn de'' test")
        Assert.IsTrue(m_objFRN.Save(), "FRN.Create")

        m_objPRD = New Produit("PRD210", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objCLT = New Client("CLT210", "Client de' test")
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
    <TestMethod()> Public Sub T10_ADD_0_ON_FAXNUMBER()

        'Quand le Numéro ne commence pas par 0 il est ajouté automatiquement
        m_objFRN.AdresseFacturation.fax = "123456"
        Assert.AreEqual(m_objFRN.AdresseFacturation.fax, "0123456")
        m_objFRN.AdresseFacturation.tel = "123456"
        Assert.AreEqual(m_objFRN.AdresseFacturation.tel, "0123456")
        m_objFRN.AdresseFacturation.port = "123456"
        Assert.AreEqual(m_objFRN.AdresseFacturation.port, "0123456")
        m_objFRN.AdresseLivraison.fax = "123456"
        Assert.AreEqual(m_objFRN.AdresseLivraison.fax, "0123456")
        m_objFRN.AdresseLivraison.tel = "123456"
        Assert.AreEqual(m_objFRN.AdresseLivraison.tel, "0123456")
        m_objFRN.AdresseLivraison.port = "123456"
        Assert.AreEqual(m_objFRN.AdresseLivraison.port, "0123456")

        'Mais quand il commence par 0, pas d'ajout
        m_objFRN.AdresseFacturation.fax = "0123456"
        Assert.AreEqual(m_objFRN.AdresseFacturation.fax, "0123456")
    End Sub

End Class


