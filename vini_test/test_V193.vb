'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB





<TestClass()> Public Class test_V193
    Inherits test_Base

    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Persist.shared_connect()
        m_objFRN = New Fournisseur("FRNV193_T10", "FRn de test")
        m_objFRN.Save()


        m_objPRD = New Produit("PRDV193_T10", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")
        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        Persist.shared_connect()
        m_objPRD.bDeleted = True
        m_objPRD.save()

        m_objFRN.bDeleted = True
        m_objFRN.Save()
        Persist.shared_disconnect()
        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_LIBMVTSTOCK()
        Dim nidPRD As Integer

        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(Now(), vncEnums.vncTypeMvt.vncmvtRegul, 0, "AAA", 1)
        m_objPRD.ajouteLigneMvtStock(Now(), vncEnums.vncTypeMvt.vncmvtRegul, 0, "", 1)
        Assert.IsTrue(m_objPRD.save())

        nidPRD = m_objPRD.id
        m_objPRD = Produit.createandload(nidPRD)
        m_objPRD.loadcolmvtStock()

        Assert.AreEqual(m_objPRD.colmvtStock.Count, 2, "Il y a 2 enregistrements dans la collection des Mvt de stocks")
        Persist.shared_disconnect()

    End Sub


End Class


