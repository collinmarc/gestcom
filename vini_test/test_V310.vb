'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports vini_App



<TestClass()> Public Class test_V310
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objPRD2 As Produit
    Private m_objPRD3 As Produit
    Private m_objPRD4 As Produit
    Private m_objFRN As Fournisseur
    Private m_objFRN2 As Fournisseur
    Private m_objFRN3 As Fournisseur
    Private m_objCLT As Client
    Private m_oCmd As CommandeClient
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        m_objFRN = New Fournisseur("F01", "FRn de'' test")
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

        m_objPRD = New Produit("F01001", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objFRN2 = New Fournisseur("F02", "FRn de'' test2")
        m_objFRN2.rs = "FRNV300-2"
        Assert.IsTrue(m_objFRN2.Save(), "FRN.Create")

        m_objPRD2 = New Produit("F02001", m_objFRN2, 1990)
        Assert.IsTrue(m_objPRD2.save(), "Prod.Create")

        m_objFRN3 = New Fournisseur("F03", "FRn de'' test3")
        m_objFRN3.rs = "FRNV300-3"
        Assert.IsTrue(m_objFRN3.Save(), "FRN3.Create")

        m_objPRD3 = New Produit("F03001", m_objFRN3, 1990)
        Assert.IsTrue(m_objPRD3.save(), "Prod3.Create")

        m_objCLT = New Client("CLTV300", "Client de' test")
        m_objCLT.rs = "Client de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")


        m_oCmd = New CommandeClient(m_objCLT)
        m_oCmd.dateCommande = #6/2/1964#
        m_oCmd.save()

        m_objPRD4 = New Produit("F01002", m_objFRN, 2000)
        m_objPRD4.bStock = False
        m_objPRD4.save()

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

        m_objPRD4.bDeleted = True
        m_objPRD4.save()

        MyBase.TestCleanup()

    End Sub
    ''' <summary>
    ''' Test que la liste ds mvts de stocks ne concerne que les Produits en Stocks
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T10_ListeMvtStock()
        Dim oMvtStk As mvtStock
        Dim colMvtStock As Collection

        Dim objPRD4 As Produit
        '' Création d'un Produit qui n'est pas géré sur le stock
        objPRD4 = New Produit("TV310-T10", m_objFRN, 2000)
        objPRD4.bStock = False
        objPRD4.save()

        'Ajout de Stockinitial

        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(CDate("01/01/2000"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/2005", 120)
        m_objPRD.savecolmvtStock()
        objPRD4.ajouteLigneMvtStock(CDate("01/01/2000"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/2005", 100)
        objPRD4.savecolmvtStock()
        'Ajout d'un Second Mvt d'inventaire
        m_objPRD.ajouteLigneMvtStock(CDate("15/01/2000"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 15/01/2005", 60)
        m_objPRD.savecolmvtStock()
        Persist.shared_disconnect()

        colMvtStock = mvtStock.getListe2("01/01/2000", "31/01/2000", m_objFRN)

        Assert.AreEqual(2, colMvtStock.Count, "le Mouvenent de stock du produit 4 ne sont pas pris en cempte")
        For Each oMvtStk In colMvtStock
            Assert.AreEqual(oMvtStk.idProduit, m_objPRD.id)
        Next


        objPRD4.bDeleted = True
        objPRD4.save()

    End Sub



    ''' <summary>
    ''' Test que les produit hors stock ne sont pas pris dans le colisage
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T60_GenerationFactureColisage()

        Dim oCmd As CommandeClient
        Dim oFactCol1 As FactColisage
        Dim oFactcol2 As FactColisage
        Dim oFactCol3 As FactColisage
        Dim oLgFactCol As LgFactColisage
        Dim oLgCmd As LgCommande
        Dim oMvtStk As mvtStock

        '' Création d'un Produit qui n'est pas géré sur le stock



        'Ajout de Stockinitial

        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/2005", 120)
        m_objPRD.savecolmvtStock()
        m_objPRD2.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/2005", 120)
        m_objPRD2.savecolmvtStock()
        m_objPRD3.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/2005", 120)
        m_objPRD3.savecolmvtStock()
        m_objPRD4.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/2005", 100)
        m_objPRD4.savecolmvtStock()
        'Ajout d'un Second Mvt d'inventaire
        m_objPRD.ajouteLigneMvtStock(CDate("15/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 15/01/2005", 60)
        m_objPRD.savecolmvtStock()
        Persist.shared_disconnect()

        oCmd = New CommandeClient(m_objCLT)
        oCmd.dateCommande = CDate("01/02/1964")

        oCmd.AjouteLigne("10", m_objPRD, 12, 10.5)
        oCmd.AjouteLigne("20", m_objPRD2, 24, 10.5)
        oCmd.AjouteLigne("30", m_objPRD3, 6, 10.5)
        oCmd.save()
        oCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        ' Livraison de la commande
        oCmd.changeEtat(vncActionEtatCommande.vncActionLivrer)
        oCmd.dateLivraison = CDate("01/02/1964")
        For Each oLgCmd In oCmd.colLignes
            oLgCmd.qteLiv = oLgCmd.qteCommande
        Next
        oCmd.save()

        ' Création d'une seconde Commande avec le produit Hors Stock
        oCmd = New CommandeClient(m_objCLT)
        oCmd.dateCommande = CDate("01/02/1964")

        oCmd.AjouteLigne("10", m_objPRD4, 12, 10.5)
        oCmd.save()
        oCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        ' Livraison de la commande
        oCmd.changeEtat(vncActionEtatCommande.vncActionLivrer)
        oCmd.dateLivraison = CDate("01/02/1964")
        For Each oLgCmd In oCmd.colLignes
            oLgCmd.qteLiv = oLgCmd.qteCommande
        Next
        oCmd.save()

        ' Fournisseur 1
        oFactCol1 = FactColisage.GenereFacture(CDate("1/01/1964"), CDate("28/02/1964"), m_objFRN)
        Assert.IsNotNull(oFactCol1, "FactCol1 generée")
        ' Fournisseur 2
        oFactcol2 = FactColisage.GenereFacture(CDate("1/01/1964"), CDate("28/02/1964"), m_objFRN2)
        ' Fournisseur 3
        oFactCol3 = FactColisage.GenereFacture(CDate("1/01/1964"), CDate("28/02/1964"), m_objFRN3)


        Assert.AreEqual(0, oFactCol1.id, "fActure non Sauvegardée")
        Assert.AreEqual(2, oFactCol1.colLignes.Count, "1 ligne par mois Facturé")

        oLgFactCol = oFactCol1.colLignes(1)
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.StockInitial, "Stock initial = PRD1")
        Assert.AreEqual(CDate("01/01/1964"), oLgFactCol.dDeb, "Date de debut")
        Assert.AreEqual(CDate("31/1/1964"), oLgFactCol.dFin, "Date de Fin")
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.qte, "Qte = Stock I - Cmd")
        oLgFactCol = oFactCol1.colLignes(2)
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.StockInitial, "Stock initial = 120")
        Assert.AreEqual(CDate("01/02/1964"), oLgFactCol.dDeb, "Date de debut")
        Assert.AreEqual(CDate("29/02/1964"), oLgFactCol.dFin, "Date de Fin")
        Assert.AreEqual(m_objPRD.qteColis(120 - 12), oLgFactCol.qte, "Qte = Stock I - Cmd")

        Assert.AreEqual(2, oFactcol2.colLignes.Count, "1 ligne par mois Facturé")
        oLgFactCol = oFactcol2.colLignes(2)
        Assert.AreEqual(m_objPRD2.qteColis(120 - 24), oLgFactCol.qte, "Qte = Stock I - Cmd")
        Assert.AreEqual(2, oFactCol3.colLignes.Count, "1 ligne par mois Facturé")
        oLgFactCol = oFactCol3.colLignes(2)
        Assert.AreEqual(m_objPRD3.qteColis(120 - 6), oLgFactCol.qte, "Qte = Stock I - Cmd")

        oFactCol1.save()
        ' La sauvegarde met à ajour la liste des mvts de stock (etat et idFactrColisage)
        Dim ocol As Collection
        ocol = mvtStock.getListe2(CDate("01/01/1964"), CDate("01/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        ' La Liste des Mvts de stocks ne concerne que les produits en stocks
        Assert.AreEqual(0, ocol.Count, "Mvt.non facturés")
        ocol = mvtStock.getListe2(CDate("01/01/1964"), CDate("01/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.IsTrue(ocol.Count > 0, "Mvt facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(oFactCol1.id, oMvtStk.idFactColisage, "ID Facture colisage")
        Next


        oFactCol1.bDeleted = True
        oFactCol1.save()
        ocol = mvtStock.getListe2(CDate("01/01/1964"), CDate("01/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        Assert.IsTrue(ocol.Count > 0, "Mvt.non facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(0, oMvtStk.idFactColisage, "ID Facture colisage")
        Next
        ocol = mvtStock.getListe2(CDate("01/01/1964"), CDate("01/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.AreEqual(0, ocol.Count, "Mvt facturés")


        oCmd.bDeleted = True
        oCmd.save()



    End Sub

    ''' <summary>
    ''' Test Les Champs LettreVoiture, CoutTransport, RefFActTRp pour la Commande Client
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T70_REFACTTRP_CMDCLT()

        Dim oCmd As CommandeClient
        Dim nId As Integer


        oCmd = New CommandeClient(m_objCLT)
        oCmd.dateCommande = CDate("01/02/2000")
        oCmd.lettreVoiture = "6138 ZK 35"
        oCmd.coutTransport = 15.66
        oCmd.refFactTRP = "FCT897564"
        Assert.IsTrue(oCmd.save(), "Insertion de commande")

        nId = oCmd.id

        oCmd = CommandeClient.createandload(nId)

        Assert.AreEqual("6138 ZK 35", oCmd.lettreVoiture, "Load LetteVoiture")
        Assert.AreEqual(15.66, oCmd.coutTransport, "Load CoutTransport")
        Assert.AreEqual("FCT897564", oCmd.refFactTRP, "Load refFactTRP")

        oCmd.lettreVoiture = "6138 ZK 35Bis"
        oCmd.coutTransport = 66.15
        oCmd.refFactTRP = "FCT897564Bis"
        Assert.IsTrue(oCmd.save(), "Insertion de commande")

        oCmd = CommandeClient.createandload(nId)
        Assert.AreEqual("6138 ZK 35Bis", oCmd.lettreVoiture, "Load LetteVoiture")
        Assert.AreEqual(66.15, oCmd.coutTransport, "Load CoutTransport")
        Assert.AreEqual("FCT897564Bis", oCmd.refFactTRP, "Load refFactTRP")



        oCmd.bDeleted = True
        oCmd.save()

    End Sub 'T70_REFACTTRP_CMDCLT
    ''' <summary>
    ''' Test Les Champs LettreVoiture, CoutTransport, RefFActTRp pour la bon Appro
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T80_REFACTTRP_BA()

        Dim oBA As BonAppro
        Dim nId As Integer


        oBA = New BonAppro(m_objFRN)
        oBA.dateCommande = CDate("01/02/2000")
        oBA.lettreVoiture = "6138 ZK 35"
        oBA.coutTransport = 15.66
        oBA.refFactTRP = "FCT897564"
        Assert.IsTrue(oBA.save(), "Insertion de commande")

        nId = oBA.id

        oBA = BonAppro.createandload(nId)

        Assert.AreEqual("6138 ZK 35", oBA.lettreVoiture, "Load LetteVoiture")
        Assert.AreEqual(15.66, oBA.coutTransport, "Load CoutTransport")
        Assert.AreEqual("FCT897564", oBA.refFactTRP, "Load refFactTRP")

        oBA.lettreVoiture = "6138 ZK 35Bis"
        oBA.coutTransport = 66.15
        oBA.refFactTRP = "FCT897564Bis"
        Assert.IsTrue(oBA.save(), "Insertion de commande")

        oBA = BonAppro.createandload(nId)
        Assert.AreEqual("6138 ZK 35Bis", oBA.lettreVoiture, "Load LetteVoiture")
        Assert.AreEqual(66.15, oBA.coutTransport, "Load CoutTransport")
        Assert.AreEqual("FCT897564Bis", oBA.refFactTRP, "Load refFactTRP")



        oBA.bDeleted = True
        oBA.save()

    End Sub 'T80_REFACTTRP_BA

    Public Sub New()

    End Sub
End Class


