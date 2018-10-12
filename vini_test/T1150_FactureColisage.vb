'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NUnit.Extensions.Forms
Imports vini_DB
Imports vini_App
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared



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

        Dim objFACT As FactColisageJ
        Dim nId As Integer

        'I - Création d'une Facture 
        '=========================
        objFACT = New FactColisageJ(m_objFRN)
        objFACT.idModeReglement = 1
        objFACT.dEcheance = "01/04/1964"

        'Save
        Assert.IsTrue(objFACT.save(), "Insert" & objFACT.getErreur)
        nId = objFACT.id


        objFACT = FactColisageJ.createandload(nId)

        Assert.AreEqual(CDate("01/04/1964"), objFACT.dEcheance)
        Assert.AreEqual(1, objFACT.idModeReglement)

        'Modification de la facture
        objFACT.dEcheance = "01/04/1984"
        objFACT.idModeReglement = 2

        Assert.IsTrue(objFACT.save(), "Update" & objFACT.getErreur)
        objFACT = FactColisageJ.createandload(nId)

        ' Assert.AreEqual(CDate("01/04/1984"), objFACT.dEcheance)
        Assert.AreEqual(2, objFACT.idModeReglement)

        objFACT.bDeleted = True
        Assert.IsTrue(objFACT.save())
    End Sub


    ''' <summary>
    ''' Test que les produit hors stock ne sont pas pris dans le colisage
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T60_GenerationFactureColisagePrdHorsStock()

        Dim oCmd As CommandeClient
        Dim oFactCol1 As FactColisageJ
        'Dim oFactcol2 As FactColisage
        'Dim oFactCol3 As FactColisage
        Dim oLgFactCol As LgFactColisage
        Dim oLgCmd As LgCommande
        Dim oMvtStk As mvtStock
        Dim objProduit As Produit


        'Ajout de Stockinitial

        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/1964", 120)
        m_objPRD.savecolmvtStock()

        m_objPRD4HorsStock.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 31/01/1964", 100)
        m_objPRD4HorsStock.savecolmvtStock()


        m_objPRD4HorsStock.bStock = False
        m_objPRD4HorsStock.idFournisseur = m_objFRN.id

        m_objPRD4HorsStock.save()


        Persist.shared_disconnect()

        oCmd = New CommandeClient(m_objCLT)
        oCmd.dateCommande = CDate("01/02/1964")

        oCmd.AjouteLigne("10", m_objPRD, 12, 10.5)
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
        oFactCol1 = FactColisageJ.GenereFacture(CDate("1/02/1964"), m_objFRN)
        Assert.IsNotNull(oFactCol1, "FactCol1 generée")


        Assert.IsTrue(oFactCol1.id = 0, "fActure non Sauvegardée")
        Assert.AreEqual(1, oFactCol1.colLignes.Count, "1 seule ligne de Facture")

        oLgFactCol = oFactCol1.colLignes(1)
        Assert.AreEqual(m_objPRD.id, oLgFactCol.oProduit.id)


        oFactCol1.save()


        ' La sauvegarde met à ajour la liste des mvts de stock (etat et idFactrColisage)
        Dim ocol As Collection
        ' La Liste des Mvts de stocks ne concerne que les produits en stocks
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("28/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        Assert.AreEqual(0, ocol.Count, "Mvt.non facturés")

        ocol = mvtStock.getListe2(CDate("01/01/1964"), CDate("28/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.IsTrue(ocol.Count = 1, "Mvt facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(oFactCol1.id, oMvtStk.idFactColisage, "ID Facture colisage")
        Next

        'La Suppression de la facture libère les mouvements de stocks
        oFactCol1.bDeleted = True
        oFactCol1.save()

        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("28/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        Assert.IsTrue(ocol.Count = 1, "Mvt.non facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(0, oMvtStk.idFactColisage, "ID Facture colisage")
        Next
        ocol = mvtStock.getListe2(CDate("01/01/1964"), CDate("28/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
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

        m_objPRD.ajouteLigneMvtStock(CDate("10/01/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 10/01/2000", 240)
        m_objPRD.ajouteLigneMvtStock(CDate("10/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 10/01/2000", -36)

        m_objPRD.save()


        Dim oDS As dsVinicom
        oDS = FactColisageJ.GenereDataSetRecapColisage("01/01/2000", "31/01/2000", m_objFRN.code, 0.1)

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
        Assert.AreEqual(oRow.RC_S04 + (240 - 36) / 6D, oRow.RC_S10)
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
    ''' Test la génération du dataset Colisage
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T80_GenereEtatRecapColisage()

        Dim n As Integer


        If System.IO.File.Exists("REcapColisageJ.pdf") Then
            System.IO.File.Delete("REcapColisageJ.pdf")
        End If
        m_objPRD.loadcolmvtStock()
        m_objPRD.bStock = vncTypeProduit.vncPlateforme
        m_objPRD.nom = "PRODUIT DE TEST POUR LE RECAP COLISAGE JOURNALIER"
        m_objPRD.colmvtStock.clear()
        m_objPRD.save()
        Dim objPRD2 As New Produit("PRD002", m_objFRN, 2018)
        objPRD2.idConditionnement = Param.conditionnementdefaut.id
        objPRD2.nom = "PRODUIT2"
        objPRD2.save()


        m_objFRN.rs = "CAVE DES MILLE SOLEILS"
        m_objFRN.Save()

        'Ajout de Stock Initial = 120
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
        m_objPRD.ajouteLigneMvtStock(CDate("11/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 11/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("12/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 12/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("13/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 13/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("14/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 14/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("15/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 15/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("16/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 16/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("17/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 17/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("18/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 18/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("19/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 19/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("20/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 20/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("21/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 21/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("22/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 22/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("23/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 23/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("24/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 24/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("25/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 25/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("26/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 26/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("27/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 27/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("28/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 28/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("29/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 29/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("30/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 30/01/2000", -6)
        m_objPRD.ajouteLigneMvtStock(CDate("31/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 31/01/2000", -6)

        m_objPRD.save()

        'Mouvements de stock pour le Produit2
        '===========================================
        objPRD2.ajouteLigneMvtStock(CDate("01/12/1999"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/12", 1200)
        'Ajout d'un Second Mvt d'inventaire = 60
        objPRD2.ajouteLigneMvtStock(CDate("05/12/1999"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 05/12", 720)

        'Ajout d'une Commande Décembre = 60
        objPRD2.ajouteLigneMvtStock(CDate("06/12/1999"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 06/12", -36)

        'Ajout d'une Commande JAnvier
        objPRD2.ajouteLigneMvtStock(CDate("01/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 05/02/2000", -48)
        objPRD2.ajouteLigneMvtStock(CDate("02/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 06/02/2000", -12)
        'Ajout d'un Appro JAnvier
        objPRD2.ajouteLigneMvtStock(CDate("03/01/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 03/01/2000", 480)
        'Commande du 04
        objPRD2.ajouteLigneMvtStock(CDate("04/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "APPRO au 06/03/2000", -12)

        'Le 10 => un Appro+une Commande

        objPRD2.ajouteLigneMvtStock(CDate("10/01/2000"), vncTypeMvt.vncmvtBonAppro, 0, "APPRO au 10/01/2000", 120)
        objPRD2.ajouteLigneMvtStock(CDate("10/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 10/01/2000", -36)
        objPRD2.ajouteLigneMvtStock(CDate("11/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 11/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("12/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 12/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("13/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 13/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("14/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 14/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("15/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 15/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("16/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 16/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("17/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 17/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("18/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 18/01/2000", -60)
        objPRD2.ajouteLigneMvtStock(CDate("19/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 19/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("20/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 20/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("21/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 21/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("22/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 22/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("23/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 23/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("24/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 24/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("25/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 25/01/2000", -60)
        objPRD2.ajouteLigneMvtStock(CDate("26/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 26/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("27/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 27/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("28/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 28/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("29/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 29/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("30/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 30/01/2000", -6)
        objPRD2.ajouteLigneMvtStock(CDate("31/01/2000"), vncTypeMvt.vncMvtCommandeClient, 0, "CMD au 31/01/2000", -6)

        objPRD2.save()

        Dim oDS As dsVinicom
        Dim dDeb As Date = CDate("01/01/2000")
        Dim dFin As Date = dDeb.AddMonths(1).AddDays(-1)
        oDS = FactColisageJ.GenereDataSetRecapColisage(dDeb, dFin, m_objFRN.code, 0.1)

        Dim objReport As ReportDocument

        objReport = New ReportDocument
        objReport.Load("V:\V5\vini_app/" & "crRecapColisageJournalier.rpt")


        Dim periode As String = dDeb.ToString("MMMM yyyy")
        objReport.SetParameterValue("Periode", periode)

        objReport.SetDataSource(oDS)


        objReport.ExportToDisk(ExportFormatType.PortableDocFormat, "REcapColisageJ.pdf")
        Assert.IsTrue(System.IO.File.Exists("REcapColisageJ.pdf"))

        System.Diagnostics.Process.Start("REcapColisageJ.pdf")


    End Sub 'T80_GenereEtatRecapColisage
    Private Sub setReportConnection(ByVal objReport As ReportDocument)

        Dim myConnectionInfo As ConnectionInfo = New ConnectionInfo()
        myConnectionInfo.ServerName = Persist.oleDBConnection.DataSource
        myConnectionInfo.DatabaseName = Persist.oleDBConnection.Database
        myConnectionInfo.UserID = My.Settings.ReportCnxUser
        myConnectionInfo.Password = My.Settings.ReportCnxPassword

        Dim mySections As Sections = objReport.ReportDefinition.Sections
        For Each mySection As Section In mySections
            Dim myReportObjects As ReportObjects = mySection.ReportObjects
            For Each myReportObject As ReportObject In myReportObjects
                If myReportObject.Kind = ReportObjectKind.SubreportObject Then
                    Dim mySubreportObject As SubreportObject = CType(myReportObject, SubreportObject)
                    Dim subReportDocument As ReportDocument = mySubreportObject.OpenSubreport(mySubreportObject.SubreportName)
                    setReportConnection(subReportDocument)
                End If
            Next
        Next
        'On n'applique pas la connexion sur les tables, car ce rapport est base sur un dataset 

        ''Applique la connection sur les tables du rapport
        'Dim myTables As Tables = objReport.Database.Tables
        'For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
        '    Dim myTableLogonInfo As TableLogOnInfo = myTable.LogOnInfo
        '    myTableLogonInfo.ConnectionInfo = myConnectionInfo
        '    myTable.ApplyLogOnInfo(myTableLogonInfo)
        'Next
    End Sub

    ''' <summary>
    ''' Test l'export vers Quadra
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T100_EXPORT()
        Dim objFact As FactColisageJ
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

        objFact = New FactColisageJ(m_objFRN)
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
        Dim objFact As FactColisageJ
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

        objFact = New FactColisageJ(m_objFRN)
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
     'Test l'incrémenation des codes
    <TestMethod()> Public Sub T70_GetNextCode()

        Dim obj1 As New FactColisageJ(m_objFRN)
        Dim obj2 As New FactColisageJ(m_objFRN)

        obj1.save()
        obj2.save()
        Assert.AreEqual(obj2.code, (obj1.code + 1).ToString())
        Assert.AreNotEqual(0, obj1.code)

        obj1.bDeleted = True
        obj1.save()
        obj2.bDeleted = True
        obj2.save()


    End Sub

    ''' <summary>
    ''' Test the factcolisage Object, Insert, Update,Delete
    ''' </summary>
    ''' <remarks></remarks>

    <TestMethod()> Public Sub T30_TestFactColisage()

        Dim oFactColisage As FactColisageJ
        Dim oFactColisage2 As FactColisageJ
        Dim nId As Integer

        oFactColisage = New FactColisageJ(m_objFRN)
        Assert.AreEqual(vncEtatCommande.vncFactCOLGeneree, oFactColisage.etat.codeEtat, oFactColisage.etat.codeEtat, "Etat Généréée")
        Assert.AreEqual(Now.ToString("MMMM yyyy"), oFactColisage.periode, "Période par defaut")
        Assert.AreEqual(Param.getConstante("CST_FACT_COL_TAXES"), oFactColisage.montantTaxes.ToString(), "Montant des taxes")
        Assert.AreEqual(m_objFRN.idModeReglement2, oFactColisage.idModeReglement, "Mode de reglemement")

        'La prédio doit être au Format MMMM yyyy
        oFactColisage.periode = "TEST"
        Assert.AreEqual(Now.ToString("MMMM yyyy"), oFactColisage.periode)

        oFactColisage.totalHT = 100
        oFactColisage.periode = "octobre 2018"

        oFactColisage.dateFacture = #1/1/2005#
        oFactColisage.montantTaxes = 10.5
        oFactColisage.totalHT = 100
        oFactColisage.totalTTC = 119.6
        oFactColisage.CommFacturation.comment = "TEST COM FACT"
        oFactColisage.montantReglement = 90
        oFactColisage.dateReglement = #1/2/2005#
        oFactColisage.refReglement = "CHQ TEST"

        Assert.IsTrue(oFactColisage.save, "FactCol.insert")
        nId = oFactColisage.id

        oFactColisage2 = FactColisageJ.createandload(nId)

        Assert.IsTrue(oFactColisage.Equals(oFactColisage2), "Equal after create and load")
        oFactColisage2.periode = "octobre 2017"
        oFactColisage2.dateFacture = #1/1/2006#
        oFactColisage2.montantTaxes = 11.5
        oFactColisage2.totalHT = 110
        oFactColisage2.totalTTC = 129.6
        oFactColisage2.CommFacturation.comment = "TEST COM FACT2"
        oFactColisage2.montantReglement = 91
        oFactColisage2.dateReglement = #1/2/2006#
        oFactColisage2.refReglement = "CHQ TEST2"
        Assert.IsTrue(oFactColisage2.save, "Factcolisage.update")

        oFactColisage = FactColisageJ.createandload(nId)

        Assert.AreEqual("octobre 2017", oFactColisage2.periode)
        Assert.AreEqual(#1/1/2006#, oFactColisage2.dateFacture)
        Assert.AreEqual(11.5D, oFactColisage2.montantTaxes)
        Assert.AreEqual(110D, oFactColisage2.totalHT)
        Assert.AreEqual(129.6D, oFactColisage2.totalTTC)
        Assert.AreEqual("TEST COM FACT2", oFactColisage2.CommFacturation.comment)
        Assert.AreEqual(91D, oFactColisage2.montantReglement)
        Assert.AreEqual(#1/2/2006#, oFactColisage2.dateReglement)
        Assert.AreEqual("CHQ TEST2", oFactColisage2.refReglement)

        oFactColisage.bDeleted = True
        oFactColisage.save()

        oFactColisage = FactColisageJ.createandload(nId)
        Assert.IsTrue(oFactColisage.id = 0, "After Delete Id = 0")

    End Sub

    ''' <summary>
    ''' Test the Line factcolisage Object, 
    ''' </summary>
    ''' <remarks></remarks>

    <TestMethod()> Public Sub T40_TestLgFactColisage()
        Dim oFactCol As FactColisageJ
        Dim oLgFactCol As LgFactColisage
        Dim nid As Integer
        Dim strResult As String

        oFactCol = New FactColisageJ(m_objFRN)
        Assert.IsTrue(oFactCol.save())


        oLgFactCol = New LgFactColisage(0)
        oLgFactCol.oProduit = m_objPRD
        oLgFactCol.qteCommande = 5
        oLgFactCol.prixU = 0.15
        oLgFactCol.calculPrixTotal()
        Assert.AreEqual(0.75D, oLgFactCol.prixHT, "Montant HT")

        oFactCol.AjouteLigne(oLgFactCol)

        Assert.AreEqual(0.75D + oFactCol.montantTaxes, oFactCol.totalHT)
        oFactCol.setEtat(vncEtatCommande.vncFactCOLExportee)
        Assert.IsTrue(oFactCol.save(), "Sauvegarde")

        nid = oFactCol.id

        oFactCol = FactColisageJ.createandload(nid)

        oFactCol.loadcolLignes()

        Assert.IsTrue(oFactCol.colLignes.Count = 1, "1 ligne")

        oLgFactCol = CType(oFactCol.colLignes(1), LgFactColisage)
        Assert.AreEqual(oLgFactCol.qteCommande, 5D)
        Assert.AreEqual(oLgFactCol.prixU, 0.15D)
        Assert.AreEqual(0.75D, oLgFactCol.prixHT, "Montant HT")

        oFactCol.bDeleted = True
        oFactCol.save()

        strResult = Persist.executeSQLQuery("SELECT COUNT(*) FROM LGFACTCOLISAGE WHERE LGCOL_FACTCOL_ID = " & nid)
        Assert.AreEqual("0", strResult, "No Lines after delete")

    End Sub
    ''' <summary>
    ''' Test la Liste des factures de colisage
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T50_ListeFactColisage()
        Dim oFactCol1 As FactColisageJ
        Dim oFactCol2 As FactColisageJ
        Dim oFactCol3 As FactColisageJ

        Dim oCol As Collection = New Collection()
        oCol = FactColisageJ.getListe(#1/1/2004#, #12/31/2004#)
        For Each oFactCol1 In oCol
            oFactCol1.bDeleted = True
            oFactCol1.save()
        Next


        oFactCol1 = New FactColisageJ(m_objFRN)
        oFactCol1.dateFacture = #1/1/2004#
        oFactCol1.save()

        oFactCol2 = New FactColisageJ(m_objFRN2)
        oFactCol2.dateFacture = #2/1/2004#
        oFactCol2.save()

        oFactCol3 = New FactColisageJ(m_objFRN3)
        oFactCol3.dateFacture = #3/1/2004#
        oFactCol3.changeEtat(vncActionEtatCommande.vncActionFactCOLExporter)
        oFactCol3.save()

        oCol = FactColisageJ.getListe(#1/1/2004#, #12/31/2004#)
        Assert.AreEqual(3, oCol.Count, "GetListe sur date")
        oCol = FactColisageJ.getListe(#1/1/2004#, #12/31/2004#, m_objFRN2.code)
        Assert.AreEqual(1, oCol.Count, "GetListe sur date+Code Fourn")
        oCol = FactColisageJ.getListe(#1/1/2004#, #12/31/2004#, , vncEtatCommande.vncFactCOLGeneree)
        Assert.AreEqual(2, oCol.Count, "GetListe sur date+Etat")
        oCol = FactColisageJ.getListe(#1/1/2004#, #1/1/2004#, , vncEtatCommande.vncFactCOLGeneree)
        Assert.AreEqual(1, oCol.Count, "GetListe sur date+Etat =>1")
        oCol = FactColisageJ.getListe(#1/1/2004#, #1/1/2004#, , vncEtatCommande.vncFactCOLExportee)
        Assert.AreEqual(0, oCol.Count, "GetListe sur date+Etat =>0")
        oCol = FactColisageJ.getListe(#3/1/2004#, #3/1/2004#, , vncEtatCommande.vncFactCOLExportee)
        Assert.AreEqual(1, oCol.Count, "GetListe sur date+Etat =>1")

        oFactCol1.bDeleted = True
        oFactCol1.save()
        oFactCol2.bDeleted = True
        oFactCol2.save()
        oFactCol3.bDeleted = True
        oFactCol3.save()

        oCol = FactColisageJ.getListe(#1/1/2004#, #12/31/2004#)
        Assert.AreEqual(0, oCol.Count, "GetListe sur date apres delete")

    End Sub

    <TestMethod()> Public Sub T60_GenerationFactureColisage2()

        Dim oCmd As CommandeClient
        Dim oFactCol1 As FactColisageJ
        Dim oFactcol2 As FactColisageJ
        Dim oFactCol3 As FactColisageJ
        Dim oLgFactCol As LgFactColisage
        Dim oLgCmd As LgCommande
        Dim oMvtStk As mvtStock

        'Ajout de Stockinitial

        Persist.shared_connect()
        m_objPRD.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/01/1964", 120)
        m_objPRD.savecolmvtStock()
        m_objPRD2.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/01/1964", 120)
        m_objPRD2.savecolmvtStock()
        m_objPRD3.ajouteLigneMvtStock(CDate("01/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 01/01/1964", 120)
        m_objPRD3.savecolmvtStock()
        'Ajout d'un Second Mvt d'inventaire
        m_objPRD.ajouteLigneMvtStock(CDate("15/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 15/01/1964", 60)
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

        ' Fournisseur 1
        oFactCol1 = FactColisageJ.GenereFacture(CDate("1/02/1964"), m_objFRN)
        Assert.IsNotNull(oFactCol1, "FactCol1 generée")
        ' Fournisseur 2
        oFactcol2 = FactColisageJ.GenereFacture(CDate("1/02/1964"), m_objFRN2)
        ' Fournisseur 3
        oFactCol3 = FactColisageJ.GenereFacture(CDate("1/02/1964"), m_objFRN3)


        Assert.IsTrue(oFactCol1.id = 0, "fActure non Sauvegardée")
        Assert.AreEqual(1, oFactCol1.colLignes.Count, "1 ligne par produit")

        oLgFactCol = oFactCol1.colLignes(1)
        Assert.AreEqual(m_objPRD.qteColis(60 - 12) * 29, CDec(oLgFactCol.qteCommande), "Qte = Stock I - Cmd")

        oLgFactCol = oFactcol2.colLignes(1)
        Assert.AreEqual(m_objPRD2.qteColis(120 - 24) * 29, CDec(oLgFactCol.qteCommande), "Qte = Stock I - Cmd")

        oLgFactCol = oFactCol3.colLignes(1)
        Assert.AreEqual(m_objPRD3.qteColis(120 - 6) * 29, CDec(oLgFactCol.qteCommande), "Qte = Stock I - Cmd")

        oFactCol1.save()
        ' La sauvegarde met à ajour la liste des mvts de stock (etat et idFactrColisage)
        Dim ocol As Collection
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("29/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        Assert.AreEqual(0, ocol.Count, "Mvt.non facturés")
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("29/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.IsTrue(ocol.Count > 0, "Mvt facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(oFactCol1.id, oMvtStk.idFactColisage, "ID Facture colisage")
        Next


        oFactCol1.bDeleted = True
        oFactCol1.save()
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("29/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        Assert.IsTrue(ocol.Count > 0, "Mvt.non facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(0, oMvtStk.idFactColisage, "ID Facture colisage")
        Next
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("29/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.IsTrue(ocol.Count = 0, "Mvt facturés")


        oCmd.bDeleted = True
        oCmd.save()


    End Sub
    ''' <summary>
    ''' Test que les produit hors stock ne sont pas pris dans le colisage
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T60_GenerationFactureColisage()

        Dim oCmd As CommandeClient
        Dim oFactCol1 As FactColisageJ
        Dim oFactcol2 As FactColisageJ
        Dim oFactCol3 As FactColisageJ
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
        'Ajout d'un Second Mvt d'inventaire > le SI est donc de 60 => 10 Colis
        m_objPRD.ajouteLigneMvtStock(CDate("15/01/1964"), vncTypeMvt.vncMvtInventaire, 0, "Inventaire au 15/01/2005", 60)
        m_objPRD.savecolmvtStock()
        Persist.shared_disconnect()

        oCmd = New CommandeClient(m_objCLT)
        oCmd.dateCommande = CDate("01/02/1964")

        oCmd.AjouteLigne("10", m_objPRD, 12, 10.5) 'COMMANDE DE 12 Prouit
        oCmd.AjouteLigne("20", m_objPRD2, 24, 10.5) ' Commande de 24 Produits2
        oCmd.AjouteLigne("30", m_objPRD3, 6, 10.5) 'Commandede 6 Produits 3
        oCmd.save()
        oCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        ' Livraison de la commande
        oCmd.changeEtat(vncActionEtatCommande.vncActionLivrer)
        oCmd.dateLivraison = CDate("01/02/1964")
        For Each oLgCmd In oCmd.colLignes
            oLgCmd.qteLiv = oLgCmd.qteCommande
        Next
        oCmd.save()

       ' Génération de la facture de colisage pour le Fournisseur 1
        oFactCol1 = FactColisageJ.GenereFacture(CDate("1/02/1964"), m_objFRN)
        Assert.IsNotNull(oFactCol1, "FactCol1 generée")
        ' Fournisseur 2
        oFactcol2 = FactColisageJ.GenereFacture(CDate("1/02/1964"), m_objFRN2)
        ' Fournisseur 3
        oFactCol3 = FactColisageJ.GenereFacture(CDate("1/02/1964"), m_objFRN3)

        ' Controle de la facture 1 pour le fourisseur 1
        Assert.AreEqual("février 1964", oFactCol1.periode)
        Assert.AreEqual(0, oFactCol1.id, "fActure non Sauvegardée")
        Assert.AreEqual(1, oFactCol1.colLignes.Count, "1 Seul produit Facturé")

        'Nombre de jour du Mois (28 pour Fevrier 1964)
        Dim njour As Integer = CDate(oFactCol1.periode).AddMonths(1).AddDays(-1).Day
        oLgFactCol = oFactCol1.colLignes(1)
        Assert.AreEqual(m_objPRD.qteColis(60 - 12) * njour, oLgFactCol.qteCommande, "Qte = Stock I - Cmd")

        Assert.AreEqual("février 1964", oFactcol2.periode)
        Assert.AreEqual(1, oFactcol2.colLignes.Count, "1 ligne par mois Facturé")
        oLgFactCol = oFactcol2.colLignes(1)
        Assert.AreEqual(m_objPRD2.qteColis(120 - 24) * njour, oLgFactCol.qteCommande, "Qte = Stock I - Cmd")

        Assert.AreEqual("février 1964", oFactCol3.periode)
        Assert.AreEqual(1, oFactCol3.colLignes.Count, "1 ligne par mois Facturé")
        oLgFactCol = oFactCol3.colLignes(1)
        Assert.AreEqual(m_objPRD3.qteColis(120 - 6) * njour, oLgFactCol.qteCommande, "Qte = Stock I - Cmd")

        'Sauvegarde le la Facture de colisage => mise à jout des mouvements de stock du Fournisseur
        oFactCol1.save()
        ' La sauvegarde met à ajour la liste des mvts de stock (etat et idFactrColisage)
        Dim ocol As Collection
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("28/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        ' La Liste des Mvts de stocks ne concerne que les produits Plateforme
        Assert.AreEqual(0, ocol.Count, "Mvt.non facturés")
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("28/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.IsTrue(ocol.Count > 0, "Mvt facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(oFactCol1.id, oMvtStk.idFactColisage, "ID Facture colisage")
        Next


        oFactCol1.bDeleted = True
        oFactCol1.save()
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("28/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
        Assert.IsTrue(ocol.Count > 0, "Mvt.non facturés")
        For Each oMvtStk In ocol
            Assert.AreEqual(0, oMvtStk.idFactColisage, "ID Facture colisage")
        Next
        ocol = mvtStock.getListe2(CDate("01/02/1964"), CDate("28/02/1964"), m_objFRN, vncEtatMVTSTK.vncMVTSTK_Fact)
        Assert.AreEqual(0, ocol.Count, "Mvt facturés")


        oCmd.bDeleted = True
        oCmd.save()



    End Sub

End Class


