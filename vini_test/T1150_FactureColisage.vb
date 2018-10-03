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
        m_objPRD.idConditionnement = Param.conditionnementdefaut.id
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


        col = CommandeClient.getListe(#6/2/1964#, #6/2/1964#, "", vncEtatCommande.vncRien, "")
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

        ' Assert.AreEqual(CDate("01/04/1984"), objFACT.dEcheance)
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
        '        Assert.AreNotEqual(CDate("01/01/2000"), oFactCol1.dEcheance)

        oLgFactCol = oFactCol1.colLignes(1)
        ' la Ligne pour le Produit 4 ne doit pas être prise en compte
        'Le Profuit 4 fait partie du fournisseur1
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.StockInitial), "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.qte), "Premier mois pas de mouvement")
        Assert.AreEqual(CDate("01/01/1964"), oLgFactCol.dDeb, "Date de debut")
        Assert.AreEqual(CDate("31/01/1964"), oLgFactCol.dFin, "Date de Fin")
        oLgFactCol = oFactCol1.colLignes(2)
        Assert.AreEqual(m_objPRD.qteColis(120 - 12), CDec(oLgFactCol.qte), "Qte = Stock I - Cmd")
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
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.StockInitial), "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.StockFinal))
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.qte), "Qte = Stock F ")
        Assert.AreEqual(1, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(1, oLgFactCol.dFin.Month, "Date de Fin")
        Assert.AreNotEqual(0, oLgFactCol.MontantHT, "Montant HT <> 0")

        oLgFactCol = oFactCol1.colLignes(2)
        ' le mois 2 
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.StockInitial), "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.qte), "Qte = Stock I - Cmd")
        Assert.AreEqual(2, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(2, oLgFactCol.dFin.Month, "Date de Fin")

        oLgFactCol = oFactCol1.colLignes(3)
        ' 1 Sorties pour le mois 3
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.StockInitial), "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.qte), "Qte = Stock I")
        Assert.AreEqual(3, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(3, oLgFactCol.dFin.Month, "Date de Fin")

        oLgFactCol = oFactCol1.colLignes(4)
        ' 1 Sorties pour le mois 4
        Assert.AreEqual(4, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(4, oLgFactCol.dFin.Month, "Date de Fin")
        Assert.AreEqual(m_objPRD.qteColis(120), CDec(oLgFactCol.StockInitial), "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), CDec(oLgFactCol.qte), "Qte = Stock I - Cmd")

        oLgFactCol = oFactCol1.colLignes(5)
        ' 1 Sorties pour le mois 5
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), CDec(oLgFactCol.StockInitial), "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), CDec(oLgFactCol.qte), "Qte = Stock I - Cmd")
        Assert.AreEqual(5, oLgFactCol.dDeb.Month, "Date de debut")
        Assert.AreEqual(5, oLgFactCol.dFin.Month, "Date de Fin")

        oLgFactCol = oFactCol1.colLignes(6)
        ' 1 Sorties pour le mois 6
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), CDec(oLgFactCol.StockInitial), "Stock initial = PRD1")
        Assert.AreEqual(m_objPRD.qteColis(120 - 12 - 12), CDec(oLgFactCol.qte), "Qte = Stock I - Cmd")
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

        Dim n As Integer

        m_objPRD.loadcolmvtStock()
        m_objPRD.bStock = vncTypeProduit.vncPlateforme
        m_objPRD.colmvtStock.clear()
        m_objPRD.save()
        'Ajout de Stock Initial = 120
        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(CDate("01/12/1999"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/12", 120)
        'Ajout d'un Second Mvt d'inventaire = 60
        m_objPRD.ajouteLigneMvtStock(CDate("05/12/1999"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 05/12", 72)

        'Ajout d'une Commande Décembre = 60
        m_objPRD.ajouteLigneMvtStock(CDate("06/12/1999"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 06/12", -36)

        'Ajout d'une Commande JAnvier
        m_objPRD.ajouteLigneMvtStock(CDate("01/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 05/02/2000", -48)
        m_objPRD.ajouteLigneMvtStock(CDate("02/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 06/02/2000", -12)
        'Ajout d'un Appro JAnvier
        m_objPRD.ajouteLigneMvtStock(CDate("03/01/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 03/01/2000", 48)
        'Commande du 04
        m_objPRD.ajouteLigneMvtStock(CDate("04/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "APPRO au 06/03/2000", -12)

        'Le 10 => un Appro+une Commande

        m_objPRD.ajouteLigneMvtStock(CDate("10/01/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 10/01/2000", 120)
        m_objPRD.ajouteLigneMvtStock(CDate("10/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 10/01/2000", -36)

        m_objPRD.save()


        Dim oDS As dsVinicom
        oDS = FactColisage.GenereDataSetRecapColisage("01/01/2000", "31/01/2000", m_objFRN.code, 0.1)

        Assert.AreEqual(1, oDS.RECAPCOLISAGEJOURN.Count, "1 ligne générée dans le dataset")

        Dim oRow As dsVinicom.RECAPCOLISAGEJOURNRow
        oRow = oDS.RECAPCOLISAGEJOURN(0)

        Assert.AreEqual(m_objPRD.code, oRow.RC_PRD_CODE)
        Assert.AreEqual((36 - 48) / 6D, oRow.RC_S01)
        Assert.AreEqual((36 - 48 - 12) / 6D, oRow.RC_S02)
        Assert.AreEqual((36 - 48 - 12 + 48) / 6D, oRow.RC_S03)
        Assert.AreEqual((36 - 48 - 12 + 48 - 12) / 6D, oRow.RC_S04)
        Assert.AreEqual(oRow.RC_S04, oRow.RC_S05)
        Assert.AreEqual(oRow.RC_S04, oRow.RC_S06)
        Assert.AreEqual(oRow.RC_S04, oRow.RC_S07)
        Assert.AreEqual(oRow.RC_S04, oRow.RC_S08)
        Assert.AreEqual(oRow.RC_S04, oRow.RC_S09)
        Assert.AreEqual((36 - 48 - 12 + 48 - 12 + 120 - 36) / 6D, oRow.RC_S10)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S11)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S12)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S13)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S14)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S15)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S16)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S17)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S18)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S19)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S20)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S21)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S22)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S23)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S24)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S25)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S26)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S27)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S28)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S29)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S30)
        Assert.AreEqual(oRow.RC_S10, oRow.RC_S31)


    End Sub 'T80_GenereDataset

    ''' <summary>
    ''' Test l'export vers Quadra
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T100_EXPORT()
        Dim objFact As FactColisage
        Dim strLines As String()
        Dim strLine As String
        Dim strLine1 As String
        Dim strLine5 As String
        Dim strLine6 As String

        Dim oParam As ParamModeReglement
        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "CHQ30NETS"
        oParam.dDebutEcheance = "FACT"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())

        objFact = New FactColisage(m_objFRN)
        objFact.periode = "1er Timestre 1964"
        objFact.dateFacture = CDate("06/02/1964")
        objFact.totalHT = 150.56
        objFact.totalTTC = 180.89
        objFact.dEcheance = "01/04/1964"
        objFact.idModeReglement = oParam.id

        Assert.IsTrue(objFact.save(), objFact.getErreur())




        'Save
        Assert.IsTrue(objFact.save(), "Insert" & objFact.getErreur)
        If File.Exists("./T20_EXPORT.txt") Then
            File.Delete("./T20_EXPORT.txt")
        End If

        objFact.Exporter("./T20_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T20_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T20_EXPORT.txt")
        Assert.AreEqual(6, strLines.Length, "3 lignes d'export")

        strLine1 = strLines(0)
        strLine5 = strLines(4)
        strLine6 = strLines(5)

        Assert.AreEqual(231, strLine1.Length)
        Assert.AreEqual("M", strLine1.Substring(0, 1))
        Assert.AreEqual(m_objFRN.CodeCompta, Trim(strLine1.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine1.Substring(9, 2)))
        Assert.AreEqual("060264", strLine1.Substring(14, 6))
        Assert.AreEqual(Trim(("F:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("D", strLine1.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine1.Substring(42, 13)))
        Assert.AreEqual("060364", Trim(strLine1.Substring(63, 6)))

        'Test de la 2eme ligne
        strLine = strLines(1)

        Dim unChamp As String()
        Dim ChampVal As String() = strLine.Split(";")
        Assert.AreEqual("Y", ChampVal(0))
        Assert.AreEqual(2, ChampVal.Length)
        unChamp = ChampVal(1).Split("=")
        Assert.AreEqual("ModePaiement", unChamp(0))
        Assert.AreEqual("CHQ", unChamp(1))


        'Test de la 3eme ligne
        strLine = strLines(2)
        Assert.AreEqual(111, strLine.Length)
        Assert.AreEqual("R", strLine.Substring(0, 1))
        Assert.AreEqual("060364", strLine.Substring(1, 6))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine.Substring(7, 13)))

        'Test de la 4eme ligne
        strLine = strLines(3)
        Assert.AreEqual(18, strLine.Length)
        Assert.AreEqual("Z;ModePaiement=", strLine.Substring(0, 15))
        Assert.AreEqual("CHQ", strLine.Substring(15, 3))


        Assert.AreEqual(231, strLine5.Length)
        Assert.AreEqual("M", strLine5.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTETVA")), Trim(strLine5.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine5.Substring(9, 2)))
        Assert.AreEqual("060264", strLine5.Substring(14, 6))
        Assert.AreEqual(Trim(("F:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine5.Substring(41, 1))
        Assert.AreEqual((180.89 - 150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine5.Substring(42, 13)))

        Assert.AreEqual(231, strLine6.Length)
        Assert.AreEqual("M", strLine6.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTEPRODUIT_COL")), Trim(strLine6.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine6.Substring(9, 2)))
        Assert.AreEqual("060264", strLine6.Substring(14, 6))
        Assert.AreEqual(Trim(("F:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine6.Substring(41, 1))
        Assert.AreEqual((150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine6.Substring(42, 13)))

        objFact.bDeleted = True
        Assert.IsTrue(objFact.save())
    End Sub
    ''' <summary>
    ''' Test l'export vers Quadra d'un avoir
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T100_EXPORTAVOIR()
        Dim objFact As FactColisage
        Dim strLines As String()
        Dim strLine As String
        Dim strLine1 As String
        Dim strLine5 As String
        Dim strLine6 As String

        Dim oParam As ParamModeReglement
        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "CHQ30NETS"
        oParam.dDebutEcheance = "FACT"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())

        objFact = New FactColisage(m_objFRN)
        objFact.periode = "1er Timestre 1964"
        objFact.dateFacture = CDate("06/02/1964")
        objFact.totalHT = -150.56
        objFact.totalTTC = -180.89
        objFact.dEcheance = "01/04/1964"
        objFact.idModeReglement = oParam.id

        Assert.IsTrue(objFact.save(), objFact.getErreur())




        'Save
        Assert.IsTrue(objFact.save(), "Insert" & objFact.getErreur)
        If File.Exists("./T20_EXPORT.txt") Then
            File.Delete("./T20_EXPORT.txt")
        End If

        objFact.Exporter("./T20_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T20_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T20_EXPORT.txt")
        Assert.AreEqual(6, strLines.Length, "3 lignes d'export")

        strLine1 = strLines(0)
        strLine5 = strLines(4)
        strLine6 = strLines(5)

        Assert.AreEqual(231, strLine1.Length)
        Assert.AreEqual("M", strLine1.Substring(0, 1))
        Assert.AreEqual(m_objFRN.CodeCompta, Trim(strLine1.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine1.Substring(9, 2)))
        Assert.AreEqual("060264", strLine1.Substring(14, 6))
        Assert.AreEqual(Trim(("A:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine1.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine1.Substring(42, 13)))
        Assert.AreEqual("060364", Trim(strLine1.Substring(63, 6)))

        'Test de la 2eme ligne
        strLine = strLines(1)

        Dim unChamp As String()
        Dim ChampVal As String() = strLine.Split(";")
        Assert.AreEqual("Y", ChampVal(0))
        Assert.AreEqual(2, ChampVal.Length)
        unChamp = ChampVal(1).Split("=")
        Assert.AreEqual("ModePaiement", unChamp(0))
        Assert.AreEqual("CHQ", unChamp(1))

        'Test de la 3eme ligne
        strLine = strLines(2)
        Assert.AreEqual(111, strLine.Length)
        Assert.AreEqual("R", strLine.Substring(0, 1))
        Assert.AreEqual("060364", strLine.Substring(1, 6))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine.Substring(7, 13)))

        'Test de la 4eme ligne
        strLine = strLines(3)
        Assert.AreEqual(18, strLine.Length)
        Assert.AreEqual("Z;ModePaiement=", strLine.Substring(0, 15))
        Assert.AreEqual("CHQ", strLine.Substring(15, 3))



        Assert.AreEqual(231, strLine5.Length)
        Assert.AreEqual("M", strLine5.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTETVA")), Trim(strLine5.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine5.Substring(9, 2)))
        Assert.AreEqual("060264", strLine5.Substring(14, 6))
        Assert.AreEqual(Trim(("A:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("D", strLine5.Substring(41, 1))
        Assert.AreEqual((180.89 - 150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine5.Substring(42, 13)))

        Assert.AreEqual(231, strLine6.Length)
        Assert.AreEqual("M", strLine6.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTEPRODUIT_COL")), Trim(strLine6.Substring(1, 8)))
        Assert.AreEqual("CO", Trim(strLine6.Substring(9, 2)))
        Assert.AreEqual("060264", strLine6.Substring(14, 6))
        Assert.AreEqual(Trim(("A:" + objFact.code + " " + m_objFRN.rs + Space(20)).Substring(0, 19)), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("D", strLine6.Substring(41, 1))
        Assert.AreEqual((150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine6.Substring(42, 13)))

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


