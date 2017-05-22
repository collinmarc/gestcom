'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V213
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
    <TestMethod()> Public Sub T10_COMMENTAIRE_50_CAR()
        'Verification que les commentaires le sont pas tronquées
        Dim objCMD As CommandeClient
        Dim nIdCmd As Long
        'Creation d'une Commande
        objCMD = New CommandeClient(m_objCLT)
        objCMD.dateCommande = "06/02/2000"
        objCMD.CommCommande.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommFacturation.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLibre.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLivraison.comment = "123456789012345678901234567890123456789012345678901234567890"
        Assert.IsTrue(objCMD.save())
        nIdCmd = objCMD.id

        objCMD = CommandeClient.createandload(nIdCmd)

        Assert.AreEqual(60, Len(objCMD.CommCommande.comment), "Commentaire de commande tronqué")
        Assert.AreEqual(60, Len(objCMD.CommFacturation.comment), "Commentaire de Facturation tronqué")
        Assert.AreEqual(60, Len(objCMD.CommLibre.comment), "Commentaire libre tronqué")
        Assert.AreEqual(60, Len(objCMD.CommLivraison.comment), "Commentaire de livraison tronqué")

        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save(), "Destruction de la commande")
    End Sub
    <TestMethod()> Public Sub T20_MtTransportBonAppro()
        'Ajout du montant de transport sur les bons Appro
        Dim objCMD As BonAppro
        Dim nIdCmd As Long
        'Creation d'une Commande
        objCMD = New BonAppro(m_objFRN)
        objCMD.dateCommande = "06/02/2000"
        objCMD.montantTransport = 150.55
        Assert.IsTrue(objCMD.save())
        nIdCmd = objCMD.id

        objCMD = BonAppro.createandload(nIdCmd)

        Assert.AreEqual(150.55, objCMD.montantTransport, "Mont de transport non lu")
        objCMD.montantTransport = 250.55
        Assert.IsTrue(objCMD.save())
        objCMD = BonAppro.createandload(nIdCmd)
        Assert.AreEqual(250.55, objCMD.montantTransport, "Mont de transport non lu")

        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save(), "Destruction du bon Appro")
    End Sub

End Class


