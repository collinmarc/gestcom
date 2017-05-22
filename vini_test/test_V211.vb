'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V211
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
    <TestMethod()> Public Sub T10_ARRONDIRQTECOLIS()
        Dim objcond As Param
        Dim qte As Decimal
        Dim bOk As Boolean
        Dim objCmd As CommandeClient
        For Each objcond In Param.colConditionnement
            If objcond.id = m_objPRD.idConditionnement Then
                bOk = True
                Exit For
            End If
        Next

        Assert.IsTrue(bOk, "Conditionnement non trouvé")

        'Calcul de la qte = le conditionnement du produit * un peu plus que le conditionnement
        qte = CDec(objcond.valeur) * 1.1

        objCmd = New CommandeClient(m_objCLT)
        objCmd.AjouteLigne("10", m_objPRD, qte, 10.63)

        objCmd.CalcPoidsColis()
        Assert.AreEqual(2, objCmd.qteColis, "Il faut 2 colis")

    End Sub

End Class


