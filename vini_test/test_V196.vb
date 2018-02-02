'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB





<TestClass()> Public Class test_V196
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objCLT As Client
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Persist.shared_connect()
        m_objFRN = New Fournisseur("FRNV193_T10", "FRn de test")
        Assert.IsTrue(m_objFRN.Save(), "FRN.Create")

        m_objPRD = New Produit("PRDV193_T10", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objCLT = New Client("CLT001", "Client de test")
        Assert.IsTrue(m_objCLT.save(), "Client.Create")

        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        Persist.shared_connect()
        m_objPRD.bDeleted = True
        m_objPRD.save()

        m_objFRN.bDeleted = True
        m_objFRN.Save()

        m_objCLT.bDeleted = True
        m_objCLT.save()

        Persist.shared_disconnect()
        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_CMD_AnnulLiv()
        Dim objCMD As CommandeClient
        Dim objLg As LgCommande

        'Creation d'un commande
        objCMD = New CommandeClient(m_objCLT)
        objCMD.AjouteLigne("10", m_objPRD, 10, 0)
        objCMD.save()

        'Livraison de la commande
        For Each objLg In objCMD.colLignes
            objLg.qteLiv = objLg.qteCommande
        Next objLg
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        objCMD.save()

        'Annulation de la livraison
        For Each objLg In objCMD.colLignes
            objLg.qteLiv = 0
        Next objLg
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnLivrer)
        objCMD.save()

        Assert.IsTrue(objCMD.colLignes.Count > 0, "Il n'y a plus de lignes dans la commandes")

        objCMD.loadcolLignes()
        objCMD.bDeleted = True
        objCMD.save()
    End Sub

    <TestMethod()> Public Sub T101_BA_AnnulLiv()
        Dim nidBA As Integer
        Dim objBA As BonAppro
        Dim objLg As LgCommande

        'Creation d'un Bon Appro
        objBA = New BonAppro(m_objFRN)
        objBA.AjouteLigne("10", m_objPRD, 10, 0)
        objBA.save()
        nidBA = objBA.id
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il Doit y avoir 1 lignes dans le BA")

        'Rechargement du bon
        objBA = BonAppro.createandload(nidBA)
        objBA.loadcolLignes()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il Doit y avoir 1 lignes dans le BA")
        'Livraison de la commande
        For Each objLg In objBA.colLignes
            objLg.qteLiv = objLg.qteCommande
        Next objLg
        objBA.changeEtat(vncEnums.vncActionEtatCommande.vncActionBALivrer)
        objBA.save()

        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")
        objBA = BonAppro.createandload(nidBA)
        objBA.loadcolLignes()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")

        'Annulation de la livraison
        For Each objLg In objBA.colLignes
            objLg.qteLiv = 0
        Next objLg
        objBA.changeEtat(vncEnums.vncActionEtatCommande.vncActionBAAnnLivrer)
        objBA.save()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")
        objBA = BonAppro.createandload(nidBA)
        objBA.loadcolLignes()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")


        objBA.bDeleted = True
        objBA.save()
    End Sub

End Class


