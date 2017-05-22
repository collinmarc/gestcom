'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1200_Integrite
    Inherits test_Base

    Private m_objCLT As Client
    Private m_objFRN As Fournisseur
    Private m_objPRD As Produit

    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Persist.shared_connect()
        'Création du client
        m_objCLT = New Client("CLT01", "CLIENT 01")
        m_objCLT.save()
        m_objFRN = New Fournisseur("FRN01", "Fournisseur 1")
        m_objFRN.Save()
        m_objPRD = New Produit("PRD1", m_objFRN, 2004)
        m_objPRD.save()

    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        m_objCLT.bDeleted = True
        m_objCLT.save()
        m_objFRN.bDeleted = True
        m_objFRN.Save()
        m_objPRD.bDeleted = True
        m_objPRD.save()

        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T01_CLient()
        Dim objCMD As CommandeClient

        'Création d'une ligne de Précommande
        m_objCLT.ajouteLgPrecom(m_objPRD, 10)
        Assert.IsTrue(m_objCLT.save())
        Assert.IsTrue(m_objCLT.checkForDelete(), "CheckforDelete CLT 1")

        'Création d'une Commande
        objCMD = New CommandeClient(m_objCLT)
        objCMD.save()
        Assert.IsFalse(m_objCLT.checkForDelete(), "CheckforDelete CLT 2")
        Assert.IsTrue(objCMD.checkForDelete(), "CheckforDelete CMD 1")
        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save(), "Delete CMD")
        Assert.IsTrue(m_objCLT.checkForDelete(), "CheckforDelete CLT 3")


    End Sub

    <TestMethod()> Public Sub T02_Fournisseur()
        Dim objBA As BonAppro



        Assert.IsFalse(m_objFRN.checkForDelete(), "CheckforDelete FRN 1")
        Assert.IsTrue(m_objPRD.checkForDelete(), "CheckforDelete PRD 1")
        m_objPRD.bDeleted = True
        Assert.IsTrue(m_objPRD.save(), "Delete PRD")
        Assert.IsTrue(m_objFRN.checkForDelete(), "CheckforDelete FRN 2")

        'Création d'un BonAppro
        objBA = New BonAppro(m_objFRN)
        Assert.IsTrue(objBA.save())
        Assert.IsFalse(m_objFRN.checkForDelete(), "CheckforDelete FRN BA")
        'Suppresssion du BA
        Assert.IsTrue(objBA.checkForDelete(), "CheckforDelete BA")
        objBA.bDeleted = True
        Assert.IsTrue(objBA.save())

        m_objPRD = New Produit("PRD1", m_objFRN, 2004)
        m_objPRD.save()





    End Sub
    <TestMethod()> Public Sub T03_Produit()
        Dim objCMD As CommandeClient


        Assert.IsTrue(m_objPRD.checkForDelete, "ChesckForDelete PRD1")

        m_objCLT.ajouteLgPrecom(m_objPRD, 10)
        Assert.IsTrue(m_objCLT.save(), "Save objCLT")
        Assert.IsTrue(m_objPRD.checkForDelete, "ChesckForDelete PRD 2")

        m_objPRD.ajouteLigneMvtStock("01/01/2004", vncEnums.vncTypeMvt.vncMvtInventaire, 0, "TEST", 10)
        Assert.IsTrue(m_objPRD.save())
        Assert.IsTrue(m_objPRD.checkForDelete, "ChesckForDelete PRD 3")

        'Creation d'une Commande avec 1 ligne
        objCMD = New CommandeClient(m_objCLT)
        objCMD.AjouteLigne("10", m_objPRD, 10, 12)
        Assert.IsTrue(objCMD.save)
        Assert.IsFalse(m_objPRD.checkForDelete, "ChesckForDelete PRD 4")

        'Suppression de la Commande
        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save())
        Assert.IsTrue(m_objPRD.checkForDelete, "ChesckForDelete PRD 5")


    End Sub
    <TestMethod()> Public Sub T03_Commande()
        Dim objCMD As CommandeClient
        Dim objSCMD As SousCommande

        Dim nId As Integer

        'Création du Fournisseaur

        'Creation d'une Commande avec 1 ligne
        objCMD = New CommandeClient(m_objCLT)
        objCMD.AjouteLigne("10", m_objPRD, 10, 12)
        Assert.IsTrue(objCMD.save)
        Assert.IsTrue(objCMD.checkForDelete, "ChesckForDelete CMD 1")
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(objCMD.save)
        Assert.IsTrue(objCMD.checkForDelete, "ChesckForDelete CMD 2")
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD.save)
        Assert.IsTrue(objCMD.checkForDelete, "ChesckForDelete CMD 3")
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(objCMD.generationSousCommande())
        Assert.IsTrue(objCMD.save)
        'Il devient impossible de supprimer la Commande
        Assert.IsFalse(objCMD.checkForDelete, "ChesckForDelete CMD 4")
        objSCMD = objCMD.colSousCommandes(1)
        objSCMD.bDeleted = True
        Assert.IsTrue(objSCMD.Save())
        'Sans sous commande cela redevient Possible
        Assert.IsTrue(objCMD.checkForDelete, "ChesckForDelete CMD 5")


        'Suppression de la Commande
        nId = objCMD.id
        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save())
        Assert.IsTrue(MsgBox("Vérifier qu'il n'y a plus enregistrement dans la table LGCOMMANDE pour la Commande " & nId, MsgBoxStyle.YesNo))



    End Sub
End Class



