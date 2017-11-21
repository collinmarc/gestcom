'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports vini_App



<TestClass()> Public Class test_V560
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objCLT As Client
    Private m_objProduit As Produit
    <TestInitialize()> Public Overrides Sub TestInitialize()

        MyBase.TestInitialize()

        m_objFRN = Fournisseur.createandload("F01")
        If m_objFRN IsNot Nothing Then
            m_objFRN.delete()
        End If
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

        m_objCLT = Client.createandload("CLTV500")
        If m_objCLT IsNot Nothing Then
            m_objCLT.delete()
        End If

        m_objCLT = New Client("CLTV500", "NOMClient de' test")
        m_objCLT.rs = "RSClient de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")


        m_objPRD = Produit.createandloadbyKey("PRDT500")
        If m_objPRD IsNot Nothing Then
            m_objPRD.delete()
        End If
        m_objPRD = New Produit("PRDT500", m_objFRN, 2011)
        Assert.IsTrue(m_objPRD.save(), "Produit.Create")
        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        m_objPRD.delete()
        m_objFRN.delete()
        m_objCLT.delete()

        MyBase.TestCleanup()
    End Sub


End Class


