'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB





<TestClass()> Public Class test_V194
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
    <TestMethod()> Public Sub T10_SAVEBA()
        Dim objBA As BonAppro
        Dim id As Integer
        Dim objLg As LgCommande


        objBA = New BonAppro(m_objFRN)
        objBA.AjouteLigne("10", m_objPRD, 10, 0)
        objBA.save()
        id = objBA.id()
        objBA.CommCommande.comment = ""
        objBA.save()
        For Each objLg In objBA.colLignes
            Assert.AreEqual(id, objLg.idBA)
        Next

        objBA.bDeleted = True
        objBA.save()



    End Sub


End Class


