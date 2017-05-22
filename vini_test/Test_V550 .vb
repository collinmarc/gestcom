'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NUnit.Extensions.Forms
Imports vini_DB
Imports vini_App



<TestClass()> Public Class test_V550
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objFRN2 As Fournisseur
    Private m_objCLT As Client
    Private m_objCLT2 As Client
    Private m_objProduit As Produit
    <TestInitialize()> Public Overrides Sub TestInitialize()

        MyBase.TestInitialize()

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

        m_objFRN2 = New Fournisseur("F02", "Nom du fournisseur-2")
        m_objFRN2.rs = "FRNV500-2"
        m_objFRN2.AdresseLivraison.nom = "ADF_Nom-2"
        m_objFRN2.AdresseLivraison.rue1 = "ADF_Rue1-2"
        m_objFRN2.AdresseLivraison.rue2 = "ADF_Rue2-2"
        m_objFRN2.AdresseLivraison.cp = "ADF_cp-2"
        m_objFRN2.AdresseLivraison.ville = "ADF_Ville-2"
        m_objFRN2.AdresseLivraison.tel = "01010101-2"
        m_objFRN2.AdresseLivraison.fax = "02020202-2"
        m_objFRN2.AdresseLivraison.port = "03030303-2"
        m_objFRN2.AdresseLivraison.Email = "04040404-1"
        m_objFRN2.AdresseFacturation.nom = "ADF_Nom-2"
        m_objFRN2.AdresseFacturation.rue1 = "ADF_Rue1-2"
        m_objFRN2.AdresseFacturation.rue2 = "ADF_Rue2-2"
        m_objFRN2.AdresseFacturation.cp = "ADF_cp-2"
        m_objFRN2.AdresseFacturation.ville = "ADF_Ville-2"
        m_objFRN2.AdresseFacturation.tel = "01010101-2"
        m_objFRN2.AdresseFacturation.fax = "02020202-2"
        m_objFRN2.AdresseFacturation.port = "03030303-2"
        m_objFRN2.AdresseFacturation.Email = "04040404-1"
        Assert.IsTrue(m_objFRN2.Save(), "FRN2.Create")

        m_objCLT = New Client("CLTV500", "NOMClient de' test")
        m_objCLT.rs = "RSClient de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")

        m_objCLT2 = New Client("CLTV500-2", "NOMClient de' test 2")
        m_objCLT2.rs = "RSClient de test 2"
        Assert.IsTrue(m_objCLT2.save(), "Client.Create")

        m_objPRD = New Produit("PRDT500", m_objFRN, 2011)
        Assert.IsTrue(m_objPRD.save(), "Produit.Create")
        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()


        MyBase.TestCleanup()
    End Sub

    ''' <summary>
    ''' Nom et Raison Sociale du client/ Fournisseur duppliqué dans les tables 
    ''' commandes et Bon Appo
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T01_NOMRSLIVRAISON_CMDCLT()
        Dim oCmd As CommandeClient

        oCmd = New CommandeClient(m_objCLT)
        oCmd.DuppliqueCaracteristiqueTiers()
        'Vérification de la concordance
        Assert.AreEqual(oCmd.caracteristiqueTiers.nom, m_objCLT.nom, "Nom différents")
        Assert.AreEqual(oCmd.caracteristiqueTiers.rs, m_objCLT.rs, "Nom différents")
        oCmd.NomLivraison = "Marc Collin"
        oCmd.RaisonSocialeLivraison = "MCII"
        Assert.AreEqual(oCmd.NomLivraison, "Marc Collin", "Nom différents")
        Assert.AreEqual(oCmd.RaisonSocialeLivraison, "MCII", "RS différents")

        oCmd.save()
        Dim nId As Integer
        nId = oCmd.id

        oCmd = CommandeClient.createandload(nId)
        Assert.AreEqual(oCmd.NomLivraison, "Marc Collin", "Nom différents")
        Assert.AreEqual(oCmd.RaisonSocialeLivraison, "MCII", "RS différents")

        oCmd.NomLivraison = "Collin Marc"
        oCmd.RaisonSocialeLivraison = "Test"

        oCmd.save()

        nId = oCmd.id

        oCmd = CommandeClient.createandload(nId)
        Assert.AreEqual(oCmd.NomLivraison, "Collin Marc", "Nom différents")
        Assert.AreEqual(oCmd.RaisonSocialeLivraison, "Test", "RS différents")

        oCmd.delete()


    End Sub
    <TestMethod()> Public Sub T10_ExportWEBEDI()
        'Test de l'export avec la zone observation agrandie à 100 car et le nom du transoprteur à la fin sur 50 car
        'Verification que L'export WebEDI supprime bien les retours chariots dans les commantaires de livraison
        Dim objCMD As CommandeClient
        Dim nFile As Integer
        Dim strResult As String
        Dim nLineNumber As Integer

        'Creation d'une Commande
        objCMD = New CommandeClient(m_objCLT)
        objCMD.RaisonSocialeLivraison = "RSLIVRAISON"
        objCMD.NomLivraison = "NOMLIVRAISON"
        objCMD.dateCommande = "06/02/2000"
        objCMD.CommCommande.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommFacturation.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLibre.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLivraison.comment = "123456789012345678901234567890123456789012345678901234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ45678901234567890123456789012345678901234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ45678901234567890123456789012345678901234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        objCMD.oTransporteur.nom = "Nom du transporteurNom du transporteurNom du transporteur"
        objCMD.AjouteLigne("10", m_objPRD, 10, 10)
        If System.IO.File.Exists("adel.txt") Then
            System.IO.File.Delete("adel.txt")
        End If
        objCMD.exporterWebEDI("adel.txt")
        nFile = FreeFile()
        FileOpen(nFile, "adel.txt", OpenMode.Input, OpenAccess.Read)
        nLineNumber = 0
        While Not EOF(nFile)
            nLineNumber = nLineNumber + 1
            strResult = LineInput(nFile)
            Console.WriteLine(strResult)
        End While

        Assert.AreEqual(1, nLineNumber, "Une seule Ligne de fichier")
        Assert.AreEqual(373, strResult.Length, "Longeur = 374-1")
        Assert.AreEqual("NOMLIVRAISON", Trim(Mid(strResult, 41, 30)), "Nom Livraison NOK")
        Assert.AreEqual("RSLIVRAISON", Trim(Mid(strResult, 71, 30)), "RS Livraison NOK")

        FileClose(nFile)
        'Suppression du fichier créé
        'System.IO.File.Delete("adel.txt")

    End Sub
    <TestMethod()> Public Sub T20_duppliqueCaracteristiqueTiers()
        'Test de l'export avec la zone observation agrandie à 100 car et le nom du transoprteur à la fin sur 50 car
        'Verification que L'export WebEDI supprime bien les retours chariots dans les commantaires de livraison
        Dim objCMD As CommandeClient
        Dim nFile As Integer
        Dim strResult As String
        Dim nLineNumber As Integer

        'Creation d'une Commande
        objCMD = New CommandeClient(m_objCLT)
        objCMD.DuppliqueCaracteristiqueTiers()
        Assert.AreEqual(objCMD.RaisonSocialeLivraison, m_objCLT.rs)
        Assert.AreEqual(objCMD.NomLivraison, m_objCLT.AdresseLivraison.nom)

    End Sub

End Class


