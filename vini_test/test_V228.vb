'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V228
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objPRD2 As Produit
    Private m_objPRD3 As Produit
    Private m_objFRN As Fournisseur
    Private m_objFRN2 As Fournisseur
    Private m_objFRN3 As Fournisseur
    Private m_objCLT As Client
    Private m_oCmd As CommandeClient
    <TestInitialize()> Public Overrides Sub TestInitialize()
        Dim col As Collection

        MyBase.TestInitialize()


        m_objFRN = New Fournisseur("FRNV228", "FRn de'' test")
        m_objFRN.rs = "FRNV22"
        m_objFRN.AdresseFacturation.nom = "ADF_Nom"
        m_objFRN.AdresseFacturation.rue1 = "ADF_Nom"
        m_objFRN.AdresseFacturation.rue2 = "ADF_Nom"
        m_objFRN.AdresseFacturation.cp = "ADF_Nom"
        m_objFRN.AdresseFacturation.ville = "ADF_Nom"
        m_objFRN.AdresseFacturation.tel = "01010101"
        m_objFRN.AdresseFacturation.fax = "02020202"
        m_objFRN.AdresseFacturation.port = "03030303"
        m_objFRN.AdresseFacturation.Email = "04040404"
        Assert.IsTrue(m_objFRN.Save(), "FRN.Create")

        m_objPRD = New Produit("PRDV228", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objFRN2 = New Fournisseur("FRNV228-2", "FRn de'' test2")
        m_objFRN2.rs = "FRNV228-2"
        Assert.IsTrue(m_objFRN2.Save(), "FRN.Create")

        m_objPRD2 = New Produit("PRDV228-2", m_objFRN2, 1990)
        Assert.IsTrue(m_objPRD2.save(), "Prod.Create")

        m_objFRN3 = New Fournisseur("FRNV228-3", "FRn de'' test3")
        m_objFRN3.rs = "FRNV228-3"
        Assert.IsTrue(m_objFRN3.Save(), "FRN3.Create")

        m_objPRD3 = New Produit("PRDV228-3", m_objFRN3, 1990)
        Assert.IsTrue(m_objPRD3.save(), "Prod3.Create")

        m_objCLT = New Client("CLTV228", "Client de' test")
        m_objCLT.rs = "Client de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")


        col = CommandeClient.getListe(#6/2/1964#, #6/2/1964#)
        For Each m_oCmd In col
            m_oCmd.bDeleted = True
            m_oCmd.save()
        Next
        m_oCmd = New CommandeClient(m_objCLT)
        m_oCmd.dateCommande = #6/2/1964#
        m_oCmd.save()


        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        Persist.shared_connect()
        m_objPRD.bDeleted = True
        Assert.IsTrue(m_objPRD.save(), "TestCleanup")
        m_objPRD2.bDeleted = True
        Assert.IsTrue(m_objPRD2.save(), "TestCleanup")
        m_objPRD3.bDeleted = True
        Assert.IsTrue(m_objPRD3.save(), "TestCleanup")

        m_objFRN.bDeleted = True
        Assert.IsTrue(m_objFRN.Save())
        m_objFRN2.bDeleted = True
        Assert.IsTrue(m_objFRN2.Save())
        m_objFRN3.bDeleted = True
        Assert.IsTrue(m_objFRN3.Save())

        m_objCLT.bDeleted = True
        Assert.IsTrue(m_objCLT.save())
        '=====
        m_oCmd.bDeleted = True
        m_oCmd.save()

        MyBase.TestCleanup()

    End Sub

End Class


