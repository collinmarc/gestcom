'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V191
    Inherits test_Base

    '<TestMethod()> Public Sub T10_Fax()
    '    Dim objFax As clsFax

    '    objFax = New clsFax
    '    Persist.shared_connect()
    '    Persist.initFax()
    '    Persist.shared_disconnect()
    '    Assert.IsTrue(Len(objFax.toString) <> 0)


    'End Sub

    <TestMethod()> Public Sub T20_Gratuit()
        Dim objClient As Client
        Dim objFournisseur As Fournisseur
        Dim objP1 As Produit
        Dim objCMD1 As CommandeClient
        Dim objLgCom As LgCommande

        Persist.shared_connect()
        'Cr�ation d'un fournisseur
        objFournisseur = New Fournisseur("FRNT20v191", "FOURNISSEUR T20 V1.9.1")
        Assert.IsTrue(objFournisseur.Save())

        'Cr�ation d'un Client
        objClient = New Client("CLTT20v191", "Client T20 V1.9.1")
        Assert.IsTrue(objClient.save())

        'Cr�ation d'1 Produits
        objP1 = New Produit("PRDT20V191", objFournisseur, 2000)
        Assert.IsTrue(objP1.save())

        'Cr�ation d'1 commandes Client
        objCMD1 = New CommandeClient(objClient)
        objCMD1.dateCommande = "01/01/2001"
        objCMD1.AjouteLigne("10", objP1, 1, 10)
        objCMD1.AjouteLigne("20", objP1, 1, 10, True)
        Assert.IsTrue(objCMD1.totalHT = 10)

        'Passage de la premie�re ligne � gratuit
        objLgCom = objCMD1.colLignes(1)
        objLgCom.bGratuit = True
        'Calcul du prix total
        objCMD1.calculPrixTotal()
        'Le total de la commande est � 0
        Assert.IsTrue(objCMD1.totalHT = 0)

        'Mais la ligne reste � 10
        Assert.IsTrue(objLgCom.prixHT = 10)
        'Calcul du prix de la ligne
        objLgCom.calculPrixTotal()
        'Apr�s recalcul, le total de la ligne passe � 0
        Assert.IsTrue(objLgCom.prixHT = 0)

        objP1.bDeleted = True
        objP1.save()

        objFournisseur.bDeleted = True
        objFournisseur.Save()

        objClient.bDeleted = True
        objClient.save()

    End Sub

End Class



