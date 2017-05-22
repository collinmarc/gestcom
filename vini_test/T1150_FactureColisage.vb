'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NUnit.Extensions.Forms
Imports vini_DB
Imports vini_App
Imports System.IO



<TestClass()> Public Class test1150_FactureColisage
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objPRD2 As Produit
    Private m_objPRD3 As Produit
    Private m_objPRD4HorsStock As Produit
    Private m_objFRN As Fournisseur
    Private m_objFRN2 As Fournisseur
    Private m_objFRN3 As Fournisseur
    Private m_objCLT As Client
    Private m_oCmd As CommandeClient
    <TestInitialize()> Public Overrides Sub TestInitialize()
        Dim col As Collection
        MyBase.TestInitialize()

        Persist.shared_connect()

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
        m_objFRN.CodeCompta = "411001"
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

        m_objCLT = New Client("C01", "Client de' test")
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

        m_objPRD4HorsStock = New Produit("T60", m_objFRN, 2000)
        m_objPRD4HorsStock.bStock = False
        m_objPRD4HorsStock.save()

        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_DB()

        Dim objFACT As FactColisage
        Dim nId As Integer

        'I - Création d'une Facture 
        '=========================
        objFACT = New FactColisage(m_objFRN)
        objFACT.idModeReglement = 1
        objFACT.dEcheance = "01/04/1964"

        'Save
        Assert.IsTrue(objFACT.save(), "Insert" & objFACT.getErreur)
        nId = objFACT.id


        objFACT = FactColisage.createandload(nId)

        Assert.AreEqual(CDate("01/04/1964"), objFACT.dEcheance)
        Assert.AreEqual(1, objFACT.idModeReglement)

        'Modification de la facture
        objFACT.dEcheance = "01/04/1984"
        objFACT.idModeReglement = 2

        Assert.IsTrue(objFACT.save(), "Update" & objFACT.getErreur)
        objFACT = FactColisage.createandload(nId)

        Assert.AreEqual(CDate("01/04/1984"), objFACT.dEcheance)
        Assert.AreEqual(2, objFACT.idModeReglement)

        objFACT.bDeleted = True
        Assert.IsTrue(objFACT.save())
    End Sub


    ''' <summary>
    ''' Test que les produit hors stock ne sont pas pris dans le colisage
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T60_GenerationFactureColisage()

        Dim oCmd As CommandeClient
        Dim oFactCol1 As FactColisage
        'Dim oFactcol2 As FactColisage
        'Dim oFactCol3 As FactColisage
        Dim oLgFactCol As LgFactColisage
        Dim oLgCmd As LgCommande
        Dim oMvtStk As mvtStock
        Dim objProduit As Produit
        Dim oParam As ParamModeReglement

        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "TEST30FDM"
        oParam.dDebutEcheance = "FDM"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())
        m_objFRN.idModeReglement2 = oParam.id
        Assert.IsTrue(m_objFRN.Save())

        'Création d'un mode de reglement comptant
        oParam = New ParamModeReglement()
        oParam.dDebutEcheance = "FACT"
        oParam.code = "TESTCOMPTANT"
        oParam.valeur2 = 0
        Assert.IsTrue(oParam.Save())
        m_objFRN2.idModeReglement2 = oParam.id
        Assert.IsTrue(m_objFRN2.Save())


        'Ajout de Stockinitial

        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/1964", 120)
        m_objPRD.savecolmvtStock()
        'm_objPRD2.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/2005", 120)
        'm_objPRD2.savecolmvtStock()
        'm_objPRD3.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/2005", 120)
        'm_objPRD3.savecolmvtStock()
        m_objPRD4HorsStock.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/1964", 100)
        m_objPRD4HorsStock.savecolmvtStock()
        'Ajout d'un Second Mvt d'inventaire
        m_objPRD.ajouteLigneMvtStock(CDate("15/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 15/01/1964", 60)
        m_objPRD.savecolmvtStock()
        Persist.shared_disconnect()

        oCmd = New CommandeClient(m_objCLT)
        oCmd.dateCommande = CDate("01/02/1964")

        oCmd.AjouteLigne("10", m_objPRD, 12, 10.5)
        'oCmd.AjouteLigne("20", m_objPRD2, 24, 10.5)
        'oCmd.AjouteLigne("30", m_objPRD3, 6, 10.5)
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

        oCmd.AjouteLigne("10", m_objPRD4HorsStock, 12, 10.5)
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
        oFactCol1 = FactColisage.GenereFacture(CDate("1/01/1964"), CDate("29/02/1964"), m_objFRN)
        Assert.IsNotNull(oFactCol1, "FactCol1 generée")
        '' Fournisseur 2
        'oFactcol2 = FactColisage.GenereFacture(CDate("1/01/1964"), CDate("29/02/1964"), m_objFRN2)
        '' Fournisseur 3
        'oFactCol3 = FactColisage.GenereFacture(CDate("1/01/1964"), CDate("29/02/1964"), m_objFRN3)


        Assert.IsTrue(oFactCol1.id = 0, "fActure non Sauvegardée")
        Assert.AreEqual(2, oFactCol1.colLignes.Count, "2 ligne par Facture")
        Assert.AreEqual(m_objFRN.idModeReglement2, oFactCol1.idModeReglement)
        Assert.AreNotEqual(CDate("01/01/2000"), oFactCol1.dEcheance)

        oLgFactCol = oFactCol1.colLignes(1)
        ' la Ligne pour le Produit 4 ne doit pas être prise en compte
        'Le Profuit 4 fait partie du fournisseur1
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.StockInitial, "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.qte, "Premier mois pas de mouvement")
        Assert.AreEqual(CDate("01/01/1964"), oLgFactCol.dDeb, "Date de debut")
        Assert.AreEqual(CDate("31/01/1964"), oLgFactCol.dFin, "Date de Fin")
        oLgFactCol = oFactCol1.colLignes(2)
        Assert.AreEqual(m_objPRD.qteColis(120 - 12), oLgFactCol.qte, "Qte = Stock I - Cmd")
        Assert.AreEqual(CDate("01/02/1964"), oLgFactCol.dDeb, "Date de debut")
        Assert.AreEqual(CDate("29/02/1964"), oLgFactCol.dFin, "Date de Fin")

        'oLgFactCol = oFactcol2.colLignes(1)
        'oLgFactCol = oFactcol2.colLignes(2)
        'Assert.AreEqual(m_objPRD2.qteColis(120 - 24), oLgFactCol.qte, "Qte = Stock I - Cmd")

        'oLgFactCol = oFactCol3.colLignes(1)
        'Assert.AreEqual(m_objPRD3.qteColis(120), oLgFactCol.qte, "Qte = Stock I - Cmd")
        'oLgFactCol = oFactCol3.colLignes(2)
        'Assert.AreEqual(m_objPRD3.qteColis(120 - 6), oLgFactCol.qte, "Qte = Stock I - Cmd")

        Assert.IsTrue(oFactCol1.save())



        ' La sauvegarde met à ajour la liste des mvts de stock (etat et idFactrColisage)
        Dim ocol As Collection
        ocol = mvtStock.getListe2(CDate("01/01/1964"), CDate("01/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        ' La Liste des Mvts de stocks ne concerne que les produits en stocks
        Assert.AreEqual(0, ocol.Count, "Mvt.non facturés")
        For Each oMvtStk In ocol
            objProduit = Produit.createandload(oMvtStk.idProduit)
            Assert.IsFalse(objProduit.bStock, "Produit non stocké")
        Next
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
    ''' Test la génération des factures de colisage sur plusieurs mois
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T70_GenerationFactureColisage2Mois()

        Dim oCmd As CommandeClient
        Dim oFactCol1 As FactColisage
        Dim oLgFactCol As LgFactColisage
        Dim oLgCmd As LgCommande
        Dim oMvtStk As mvtStock

        'Ajout de Stock Initial = 120
        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(CDate("01/01/2000"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/01/2000", 120)
        m_objPRD.savecolmvtStock()
        'Ajout d'un Second Mvt d'inventaire = 60
        m_objPRD.ajouteLigneMvtStock(CDate("05/02/2000"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 15/01/2000", 60)
        m_objPRD.savecolmvtStock()
        'Ajout d'un  d'inventaire = 150
        m_objPRD.ajouteLigneMvtStock(CDate("01/07/2000"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/07/2000", 150)
        m_objPRD.savecolmvtStock()
        Persist.shared_disconnect()

        oCmd = New CommandeClient(m_objCLT)
        oCmd.dateCommande = CDate("03/04/2000")

        oCmd.AjouteLigne("10", m_objPRD, 12, 10.5)
        oCmd.AjouteLigne("20", m_objPRD, 12, 10.5)
        oCmd.save()
        oCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        ' Livraison de la commande
        oCmd.changeEtat(vncActionEtatCommande.vncActionLivrer)
        oCmd.dateLivraison = CDate("03/04/2000")
        For Each oLgCmd In oCmd.colLignes
            oLgCmd.qteLiv = oLgCmd.qteCommande
        Next
        oCmd.save()


        ' Fournisseur 1
        oFactCol1 = FactColisage.GenereFacture(CDate("01/01/2000"), CDate("30/06/2000"), m_objFRN)
        Assert.IsNotNull(oFactCol1, "FactCol1 generée")


        Assert.IsTrue(oFactCol1.id = 0, "fActure non Sauvegardée")
        Assert.AreEqual(6, oFactCol1.colLignes.Count, "6 lignes (1 par mois même le permier mois)")

        oLgFactCol = oFactCol1.colLignes(1)
        ' le mois 1 => Stock initial
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.StockInitial, "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.StockFinal)
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.qte, "Qte = Stock F ")
        Assert.AreEqual(1, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(1, oLgFactCol.dFin.Month, "Date de Fin")
        Assert.AreNotEqual(0, oLgFactCol.MontantHT, "Montant HT <> 0")

        oLgFactCol = oFactCol1.colLignes(2)
        ' le mois 2 
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.StockInitial, "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.qte, "Qte = Stock I - Cmd")
        Assert.AreEqual(2, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(2, oLgFactCol.dFin.Month, "Date de Fin")

        oLgFactCol = oFactCol1.colLignes(3)
        ' 1 Sorties pour le mois 3
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.StockInitial, "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.qte, "Qte = Stock I")
        Assert.AreEqual(3, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(3, oLgFactCol.dFin.Month, "Date de Fin")

        oLgFactCol = oFactCol1.colLignes(4)
        ' 1 Sorties pour le mois 4
        Assert.AreEqual(4, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(4, oLgFactCol.dFin.Month, "Date de Fin")
        Assert.AreEqual(m_objPRD.qteColis(120), oLgFactCol.StockInitial, "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), oLgFactCol.qte, "Qte = Stock I - Cmd")

        oLgFactCol = oFactCol1.colLignes(5)
        ' 1 Sorties pour le mois 5
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), oLgFactCol.StockInitial, "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), oLgFactCol.qte, "Qte = Stock I - Cmd")
        Assert.AreEqual(5, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(5, oLgFactCol.dFin.Month, "Date de Fin")

        oLgFactCol = oFactCol1.colLignes(6)
        ' 1 Sorties pour le mois 6
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), oLgFactCol.StockInitial, "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), oLgFactCol.qte, "Qte = Stock I - Cmd")
        Assert.AreEqual(6, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual("30/06/2000", oLgFactCol.dFin.ToShortDateString(), "Date de Fin")

        oFactCol1.save()



        ' La sauvegarde met à ajour la liste des mvts de stock (etat et idFactrColisage)
        Dim ocol As Collection
        ocol = mvtStock.getListe2(CDate("01/01/2000"), CDate("30/06/2000"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        ' La Liste des Mvts de stocks ne concerne que les produits en stocks
        Assert.AreEqual(0, ocol.Count, "Mvt.non facturés")
        ocol = mvtStock.getListe2(CDate("01/07/2000"), CDate("31/07/2000"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        ' La Liste des Mvts de stocks ne concerne que les produits en stocks
        Assert.AreEqual(1, ocol.Count, "Le Dernier Mouvemetn d'inventaire ne doit pas être facturé")

        ocol = mvtStock.getListe2(CDate("01/01/2000"), CDate("30/06/2000"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.IsTrue(ocol.Count > 0, "Mvt facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(oFactCol1.id, oMvtStk.idFactColisage, "ID Facture colisage")
        Next


        'La Suppression de la facture met à jour les mouvements de stocks
        oFactCol1.bDeleted = True
        oFactCol1.save()
        ocol = mvtStock.getListe2(CDate("01/01/2000"), CDate("30/06/2000"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        Assert.IsTrue(ocol.Count > 0, "Mvt.non facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(0, oMvtStk.idFactColisage, "ID Facture colisage")
        Next
        ocol = mvtStock.getListe2(CDate("01/01/2000"), CDate("30/06/2000"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.AreEqual(0, ocol.Count, "Mvt facturés")


        oCmd.bDeleted = True
        oCmd.save()



    End Sub

    ''' <summary>
    ''' Test la génération du dataset Colisage
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T80_GenereDataSet()

        Dim tabRows As System.Data.DataRow()
        Dim n As Integer

        'Ajout de Stock Initial = 120
        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(CDate("01/12/1999"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/01/2005", 120)
        m_objPRD.savecolmvtStock()
        'Ajout d'un Second Mvt d'inventaire = 60
        m_objPRD.ajouteLigneMvtStock(CDate("05/12/1999"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 15/01/2005", 60)
        m_objPRD.savecolmvtStock()

        'Ajout d'une Commande Décembre = 60
        m_objPRD.ajouteLigneMvtStock(CDate("05/12/1999"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 05/12/1999", -60)
        m_objPRD.savecolmvtStock()
        'Ajout d'une Commande Décembre = 48
        m_objPRD.ajouteLigneMvtStock(CDate("06/12/1999"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 06/12/1999", -48)
        m_objPRD.savecolmvtStock()

        'Ajout d'une Commande Février = 54
        m_objPRD.ajouteLigneMvtStock(CDate("05/02/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 05/02/2000", -54)
        m_objPRD.savecolmvtStock()
        'Ajout d'une Commande Février = 36
        m_objPRD.ajouteLigneMvtStock(CDate("06/02/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 06/02/2000", -36)
        m_objPRD.savecolmvtStock()

        'Ajout d'un Appro Mars = 48
        m_objPRD.ajouteLigneMvtStock(CDate("05/02/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 05/03/2000", 48)
        m_objPRD.savecolmvtStock()
        'Ajout d'une Commande Mars = 42
        m_objPRD.ajouteLigneMvtStock(CDate("06/02/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 06/03/2000", 42)
        m_objPRD.savecolmvtStock()

        'Ajout d'un Appro Avril = 18
        m_objPRD.ajouteLigneMvtStock(CDate("05/04/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 05/04/2000", 18)
        m_objPRD.savecolmvtStock()
        'Ajout d'un Appro  Avril = 12
        m_objPRD.ajouteLigneMvtStock(CDate("06/04/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 06/04/2000", 12)
        m_objPRD.savecolmvtStock()

        'Ajout d'un Mbt Regul Mai  = +12
        m_objPRD.ajouteLigneMvtStock(CDate("05/05/2000"), vncTypeMvt.vncmvtRegul, 0, "REGUL au 05/05/2000", 12)
        m_objPRD.savecolmvtStock()

        'Ajout d'un Mbt Regul Mai  = -6
        m_objPRD.ajouteLigneMvtStock(CDate("06/05/2000"), vncTypeMvt.vncmvtRegul, 0, "REGUL au 06/05/2000", -6)
        m_objPRD.savecolmvtStock()
        Persist.shared_disconnect()


        'Ajout d'Inventaire  Juin = 120
        m_objPRD.ajouteLigneMvtStock(CDate("01/06/2000"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au  01/06/2000", 6)
        m_objPRD.savecolmvtStock()
        'Ajout d'un Appro  Juin = 120
        m_objPRD.ajouteLigneMvtStock(CDate("06/06/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 06/06/2000", 120)
        m_objPRD.savecolmvtStock()
        'Ajout d'une Cmd  Juiller = 66
        m_objPRD.ajouteLigneMvtStock(CDate("06/06/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 06/06/2000", 66)
        m_objPRD.savecolmvtStock()
        'Ajout d'un Appro  Juiller = 120
        m_objPRD.ajouteLigneMvtStock(CDate("06/06/2000"), vncTypeMvt.vncmvtRegul, 0, "APPRO au 06/06/2000", 120)
        m_objPRD.savecolmvtStock()

        'Ajout d'Inventaire  Decembre = 120
        m_objPRD.ajouteLigneMvtStock(CDate("01/10/1999"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au  01/07/1999", 6)
        m_objPRD.savecolmvtStock()
        'Ajout d'un Appro  Decembre = 120
        m_objPRD.ajouteLigneMvtStock(CDate("06/10/1999"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 06/07/1999", 120)
        m_objPRD.savecolmvtStock()
        'Ajout d'une Cmd  Decembre = 66
        m_objPRD.ajouteLigneMvtStock(CDate("06/10/1999"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 06/07/1999", 66)
        m_objPRD.savecolmvtStock()
        'Ajout d'un Appro  Decembre = 120
        m_objPRD.ajouteLigneMvtStock(CDate("06/10/1999"), vncTypeMvt.vncmvtRegul, 0, "APPRO au 06/07/1999", 120)
        m_objPRD.savecolmvtStock()


        Dim oDS As dsVinicom
        Dim oRECAPRow As dsVinicom.RECAP_COLISAGERow
        oDS = FactColisage.GenereDataSetRecapColisage("01/12/1999", "30/05/2000", m_objFRN.code, 0.1)
        For Each oRECAPRow In oDS.RECAP_COLISAGE
            Console.WriteLine(oRECAPRow.RC_DATE.ToShortDateString() + "," + oRECAPRow.RC_TYPE + "," + oRECAPRow.RC_SI.ToString() + " + " + oRECAPRow.RC_ENTREE.ToString() + " " + oRECAPRow.RC_SORTIE.ToString() + " = " + oRECAPRow.RC_SF.ToString() + "/" + oRECAPRow.RC_COUT.ToString())
        Next
        tabRows = oDS.RECAP_COLISAGE.Select("RC_TYPE = '9'")
        Assert.AreEqual(6, tabRows.Length)

        n = 0 ''Janvier
        oRECAPRow = CType(tabRows(n), dsVinicom.RECAP_COLISAGERow)
        Assert.AreEqual(m_objPRD.qteColis(120), oRECAPRow.RC_SI)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48), oRECAPRow.RC_SF)

        n = 1 ''Fevrier
        oRECAPRow = CType(tabRows(n), dsVinicom.RECAP_COLISAGERow)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48), oRECAPRow.RC_SI)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48), oRECAPRow.RC_SF)

        n = 2 'Mars
        oRECAPRow = CType(tabRows(n), dsVinicom.RECAP_COLISAGERow)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48), oRECAPRow.RC_SI)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48 - 54 - 36 + 48 + 42), oRECAPRow.RC_SF)

        n = 3 'Avril
        oRECAPRow = CType(tabRows(n), dsVinicom.RECAP_COLISAGERow)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48 - 54 - 36 + 48 + 42), oRECAPRow.RC_SI)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48 - 54 - 36 + 48 + 42), oRECAPRow.RC_SF)

        n = 4 'Mai
        oRECAPRow = CType(tabRows(n), dsVinicom.RECAP_COLISAGERow)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48 - 54 - 36 + 48 + 42), oRECAPRow.RC_SI)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48 - 54 - 36 + 48 + 42 + 18 + 12), oRECAPRow.RC_SF)

        n = 5 'Juin
        oRECAPRow = CType(tabRows(n), dsVinicom.RECAP_COLISAGERow)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48 - 54 - 36 + 48 + 42 + 18 + 12), oRECAPRow.RC_SI)
        Assert.AreEqual(m_objPRD.qteColis(120 - 60 - 48 - 54 - 36 + 48 + 42 + 18 + 12 + 12 - 6), oRECAPRow.RC_SF)

    End Sub 'T80_GenereDataset

    'Ignore("test sur des valeurs réelles")
    <TestMethod(), Ignore()> Public Sub T90_Test()

        Dim oFact As FactColisage
        Dim oFRN As Fournisseur
        Dim oLG As LgFactColisage

        oFRN = Fournisseur.createandload("002")
        oFact = FactColisage.GenereFacture("01/10/2007", "31/12/2007", oFRN)
        Assert.AreEqual(3, oFact.colLignes.Count)

        oLG = CType(oFact.colLignes(1), LgFactColisage)
        Assert.AreEqual(10, oLG.dDeb.Month)
        Assert.AreEqual(253.4, oLG.MontantHT)
        oLG = CType(oFact.colLignes(2), LgFactColisage)
        Assert.AreEqual(11, oLG.dDeb.Month)
        Assert.AreEqual(303.1, oLG.MontantHT)
        oLG = CType(oFact.colLignes(3), LgFactColisage)
        Assert.AreEqual(12, oLG.dDeb.Month)
        Assert.AreEqual(300, oLG.MontantHT)

    End Sub
    ''' <summary>
    ''' Test l'export vers Quadra
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T100_EXPORT()
        Dim objFact As FactColisage
        Dim strLines As String()
        Dim strLine1 As String
        Dim strLine2 As String
        Dim strLine3 As String

        objFact = New FactColisage(m_objFRN)
        objFact.periode = "1er Timestre 1964"
        objFact.dateFacture = CDate("06/02/1964")
        objFact.totalHT = 150.56
        objFact.totalTTC = 180.89
        objFact.dEcheance = "01/04/1964"

        Assert.IsTrue(objFact.save(), objFact.getErreur())




        'Save
        Assert.IsTrue(objFact.save(), "Insert" & objFact.getErreur)
        If File.Exists("./T20_EXPORT.txt") Then
            File.Delete("./T20_EXPORT.txt")
        End If

        objFact.Exporter("./T20_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T20_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T20_EXPORT.txt")
        Assert.AreEqual(3, strLines.Length, "3 lignes d'export")

        strLine1 = strLines(0)
        strLine2 = strLines(1)
        strLine3 = strLines(2)

        Assert.AreEqual(231, strLine1.Length)
        Assert.AreEqual("M", strLine1.Substring(0, 1))
        Assert.AreEqual(m_objFRN.CodeCompta, Trim(strLine1.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine1.Substring(9, 2)))
        Assert.AreEqual("060264", strLine1.Substring(14, 6))
        Assert.AreEqual(Trim(("F:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("D", strLine1.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine1.Substring(42, 13)))
        Assert.AreEqual("010464", Trim(strLine1.Substring(63, 6)))

        Assert.AreEqual(231, strLine2.Length)
        Assert.AreEqual("M", strLine2.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTETVA")), Trim(strLine2.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine2.Substring(9, 2)))
        Assert.AreEqual("060264", strLine2.Substring(14, 6))
        Assert.AreEqual(Trim(("F:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine2.Substring(41, 1))
        Assert.AreEqual((180.89 - 150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine2.Substring(42, 13)))

        Assert.AreEqual(231, strLine3.Length)
        Assert.AreEqual("M", strLine3.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTEPRODUIT_COL")), Trim(strLine3.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine3.Substring(9, 2)))
        Assert.AreEqual("060264", strLine3.Substring(14, 6))
        Assert.AreEqual(Trim(("F:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine3.Substring(41, 1))
        Assert.AreEqual((150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine3.Substring(42, 13)))

        objFact.bDeleted = True
        Assert.IsTrue(objFact.save())
    End Sub
    <TestMethod()> Public Sub T110_ChampsLongs()

        Dim objFACT As FactColisage

        'I - Création d'une Facture 
        '=========================
        objFACT = New FactColisage(m_objFRN)

        objFACT.periode = "1er Timestre 1964".PadRight(500, "x")
        'Save
        Assert.IsTrue(objFACT.save(), "Insert" & objFACT.getErreur)

        objFACT.load()

        Assert.AreEqual(50, objFACT.periode.Length)
        objFACT.bDeleted = True
        Assert.IsTrue(objFACT.save())
    End Sub
    'Test l'incrémenation des codes
    <TestMethod()> Public Sub T70_GetNextCode()

        Dim obj1 As New FactColisage(m_objFRN)
        Dim obj2 As New FactColisage(m_objFRN)

        obj1.save()
        obj2.save()
        Assert.AreEqual(obj2.code, (obj1.code + 1).ToString())
        Assert.AreNotEqual(0, obj1.code)

        obj1.bDeleted = True
        obj1.save()
        obj2.bDeleted = True
        obj2.save()


    End Sub
End Class


