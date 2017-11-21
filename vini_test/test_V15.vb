'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V15
    Inherits test_Base

    <TestMethod()> Public Sub T10_Commande()

        Dim objCmd As CommandeClient

        objCmd = New CommandeClient(New Client("", ""))
        'La commande par défaut est Plateforme/Départ
        Assert.IsTrue(objCmd.typeCommande = vncEnums.vncTypeCommande.vncCmdClientPlateforme, "Type de Commande = plateforme")
        Assert.IsTrue(objCmd.typeTransport = vncEnums.vncTypeTransport.vncTrpDepart, "Type de transport")
        'La commande directe ne génére pas de mvts de stocks
        objCmd.typeCommande = vncEnums.vncTypeCommande.vncCmdClientDirecte
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCmd.etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncRien)


    End Sub

    <TestMethod()> Public Sub T20_OrdreMvtStock()

        'Les Mouvement de Sotck divent être triès sur leur date descendant
        Dim objProduit As Produit
        Dim objMvtStock As mvtStock
        objProduit = New Produit("", New Fournisseur, 1990)

        objMvtStock = New mvtStock("06/02/1964", 0, vncEnums.vncTypeMvt.vncmvtRegul, 0, "Mvt 02/06/64")
        objProduit.ajouteLigneMvtStock(objMvtStock, False)

        objMvtStock = New mvtStock(Now(), 0, vncEnums.vncTypeMvt.vncmvtRegul, 0, "Mvt du jour")
        objProduit.ajouteLigneMvtStock(objMvtStock, False)

        objMvtStock = objProduit.colmvtStock(0)
        Assert.IsTrue(objMvtStock.libelle = "Mvt du jour")

        objMvtStock = objProduit.colmvtStock(1)
        Assert.IsTrue(objMvtStock.libelle = "Mvt 02/06/64")

    End Sub
    <TestMethod()> Public Sub T30_ListeProduitPlateforme()
        Dim colPRD As Collection
        Dim objPRD As Produit
        Dim objPRD2 As Produit
        Dim objFRN As Fournisseur
        Dim nidPrd1 As Integer
        Dim nidPrd2 As Integer
        Dim nidFRN As Integer

        Persist.shared_connect()
        objFRN = New Fournisseur("FV15T30", "Mon Fournisseur")
        Assert.IsTrue(objFRN.Save())
        nidFRN = objFRN.id

        objPRD2 = New Produit("PRDV15T30-1", objFRN, 2004)
        objPRD2.nom = "MYPRODuct V15 T30-1, Produit stocké"
        objPRD2.bStock = True
        Assert.IsTrue(objPRD2.save())
        nidPrd2 = objPRD2.id

        objPRD = New Produit("PRDV15T30-2", objFRN, 2004)
        objPRD.nom = "MYPRODuct V15 T30 - 2, Produit Nonstocké"
        objPRD.bStock = False
        Assert.IsTrue(objPRD.save())
        nidPrd1 = objPRD.id



        'II- Liste de tous les Produits
        '=============================
        'Console.Out.WriteLine("Liste Tous les produits")
        'a) Caractère générique
        Persist.shared_connect()
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncTous, "PRDV15%")
        Assert.AreEqual(2, colPRD.Count, "Col.count <> 2")
        For Each objPRD In colPRD
            Assert.AreEqual(Left(objPRD.code, 6), "PRDV15", "Mauvais Code " & objPRD.code)
        Next objPRD
        'b) sans Caractère générique
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncTous, "PRDV15T30-2")
        Assert.AreEqual(1, colPRD.Count, "Col.count <> 1")
        For Each objPRD In colPRD
            Assert.AreEqual(objPRD.code, "PRDV15T30-2", "Mauvais Code " & objPRD.code)
        Next objPRD
        Persist.shared_disconnect()

        'III- Liste des Produits Plateforme
        '=============================
        'Console.Out.WriteLine("Liste des produits Plateforme")
        'a) Caractère générique
        Persist.shared_connect()
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncPlateforme, "PRDV15%")
        Assert.AreEqual(1, colPRD.Count, "Col.count <> 1")
        For Each objPRD In colPRD
            Assert.AreEqual(Left(objPRD.code, 6), "PRDV15", "Mauvais Code " & objPRD.code)
        Next objPRD
        'b) sans Caractère générique avec 1 résultat
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncPlateforme, "PRDV15T30-1")
        Assert.AreEqual(1, colPRD.Count, "Col.count <> 1")
        For Each objPRD In colPRD
            Assert.AreEqual(objPRD.code, "PRDV15T30-1", "Mauvais Code " & objPRD.code)
        Next objPRD
        'b) sans Caractère générique sans résultat
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncPlateforme, "PRDV15T30-2")
        Assert.IsTrue(colPRD.Count = 0, "Col.count = 0" & Produit.getErreur)
        Persist.shared_disconnect()


        'III- Liste des Produits Non stocké
        '=============================
        'Console.Out.WriteLine("Liste des produits Non Stocké")
        'a) Caractère générique
        Persist.shared_connect()
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncFournisseur, "PRDV15%")
        Assert.AreEqual(1, colPRD.Count, "Col.count <> 1")
        For Each objPRD In colPRD
            Assert.AreEqual(Left(objPRD.code, 6), "PRDV15", "Mauvais Code " & objPRD.code)
        Next objPRD
        'b) sans Caractère générique avec 1 résultat
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncFournisseur, "PRDV15T30-2")
        Assert.AreEqual(1, colPRD.Count, "Col.count <> 1")
        For Each objPRD In colPRD
            Assert.AreEqual(objPRD.code, "PRDV15T30-2", "Mauvais Code " & objPRD.code)
        Next objPRD
        'b) sans Caractère générique sans résultat
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncFournisseur, "PRDV15T30-1")
        Assert.IsTrue(colPRD.Count = 0, "Col.count = 0" & Produit.getErreur)
        Persist.shared_disconnect()

        Assert.IsTrue(objPRD.load(nidPrd1))
        objPRD.bDeleted = True
        Assert.IsTrue(objPRD.save())

        Assert.IsTrue(objPRD.load(nidPrd2))
        objPRD.bDeleted = True
        Assert.IsTrue(objPRD.save())

        Assert.IsTrue(objFRN.load(nidFRN))
        objFRN.bDeleted = True
        Assert.IsTrue(objFRN.Save())

    End Sub
    <TestMethod()> Public Sub T40_QteCommande()
        'Les Quantité en Commande Directe ne doivent pas être pris en compte dans les stocks

        Dim objCmd As CommandeClient
        Dim objClt As Client
        Dim objFRN As Fournisseur
        Dim objPRD As Produit

        objFRN = New Fournisseur("FRNV15T40", "Fournisseur Test")
        Assert.IsTrue(objFRN.Save())

        objPRD = New Produit("PRDV15T40", objFRN, 2004)
        Assert.IsTrue(objPRD.save())

        objClt = New Client("CLTV15T40", "Client Test")
        Assert.IsTrue(objClt.save())

        objCmd = New CommandeClient(objClt)
        Assert.IsTrue(objCmd.save())

        'Vérification de la Quantité en commande Avant
        objPRD.load()
        Assert.AreEqual(0D, objPRD.qteCommande, "Au début QteEn commande = 0")

        objCmd.typeCommande = vncEnums.vncTypeCommande.vncCmdClientPlateforme
        objCmd.AjouteLigne("10", objPRD, 10, 5.5)
        objCmd.save()

        'Vérification de la Quantité en commande 
        objPRD.load()
        Assert.AreEqual(10D, objPRD.qteCommande, "QteEn commande = 10, car la commande est de type plateforme")

        objCmd.typeCommande = vncEnums.vncTypeCommande.vncCmdClientDirecte
        objCmd.save()

        'Vérification de la Quantité en commande 
        objPRD.load()
        Assert.AreEqual(0D, objPRD.qteCommande, "QteEn commande = 0, car la commande est de type directe")

        objCmd.bDeleted = True
        objCmd.save()
        objPRD.bDeleted = True
        objPRD.save()

        objClt.bDeleted = True
        objClt.save()
        objFRN.bDeleted = True
        objFRN.Save()

    End Sub


End Class



