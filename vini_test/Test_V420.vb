'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NUnit.Extensions.Forms
Imports vini_DB
Imports vini_App



<TestClass()> Public Class test_V420
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
        m_objPRD.bDisponible = True
        m_objPRD.bStock = True
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
        m_oCmd.typeCommande = vncTypeCommande.vncCmdClientPlateforme
        m_oCmd.save()

        m_objPRD4 = New Produit("F01002", m_objFRN, 2000)
        m_objPRD4.bStock = False
        m_objPRD4.save()

        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()


        MyBase.TestCleanup()
    End Sub

    ''' <summary>
    ''' Test la génération des factures de commission avec des sous commandes rapprochées
    ''' changement d'état Raprochée -> Facturée
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T00()

        Dim oCol As Collection
        Dim oColScmd As List(Of SousCommande)
        Dim oColEvent As ColEvent
        Dim oFact As FactCom
        Dim objSCMD As SousCommande

        'Ajout de 2 lignes sur la commande
        m_oCmd.AjouteLigne("10", m_objPRD, 1, 150)
        m_oCmd.AjouteLigne("20", m_objPRD2, 2, 160)
        m_oCmd.save()

        m_oCmd.changeEtat(vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save())

        'Génération des Sous-Commandes
        m_oCmd.generationSousCommande()
        Assert.IsTrue(m_oCmd.save(), "Ocmd.Save" & m_oCmd.getErreur())

        ' Rapprochement des Sous Commandes (sans provisionnement)
        ' 1 Sous commande par fax
        ' l'autre par internet

        objSCMD = m_oCmd.colSousCommandes(1)
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
        objSCMD.refFactFournisseur = "MYRef"
        objSCMD.dateFactFournisseur = #6/2/1964#
        objSCMD.totalHTFacture = objSCMD.totalHT
        objSCMD.baseCommission = objSCMD.totalHT
        objSCMD.tauxCommission = 0.08
        objSCMD.MontantCommission = objSCMD.baseCommission * 0.08
        Assert.IsTrue(objSCMD.Save(), objSCMD.getErreur())

        objSCMD = m_oCmd.colSousCommandes(2)
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
        objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
        objSCMD.refFactFournisseur = "MYRef"
        objSCMD.dateFactFournisseur = #6/2/1964#
        objSCMD.totalHTFacture = objSCMD.totalHT
        objSCMD.baseCommission = objSCMD.totalHT
        objSCMD.tauxCommission = 0.08
        objSCMD.MontantCommission = objSCMD.baseCommission * 0.08
        Assert.IsTrue(objSCMD.Save(), objSCMD.getErreur())


        'Récupération de la liste des Sous-Commandes a facturer 
        oColScmd = SousCommande.getListeAFacturer(#6/2/1964#, #6/2/1964#)
        Assert.AreEqual(2, oColScmd.Count)

        'Génération des Factures de commissions
        oColEvent = FactCom.createFactComs(oColScmd, #6/2/1964#, #6/2/1964#, "Test")
        For Each oFact In oColEvent
            Assert.IsTrue(oFact.Save(), oFact.getErreur())
        Next

        'Lectures des factures de commissions
        oCol = FactCom.getListe(#6/2/1964#, #6/2/1964#)
        Assert.AreEqual(2, oCol.Count)
        oFact = oCol(1)
        Assert.AreEqual(CDec(150 * 0.08), oFact.totalHT)
        oFact = oCol(2)
        Assert.AreEqual(CDec(320 * 0.08), oFact.totalHT)

        'Lectures Sous commandes Facturée
        oColScmd = SousCommande.getListe(#6/2/1964#, #6/2/1964#, , vncEtatCommande.vncSCMDFacturee)
        Assert.AreEqual(2, oColScmd.Count)
        'Lectures Sous commandes A Facturer 
        oColScmd = SousCommande.getListeAFacturer(#6/2/1964#, #6/2/1964#)
        Assert.AreEqual(0, oColScmd.Count)




    End Sub
    ''' <summary>
    ''' Test le passge du taux de commission de la commande à la sous commande
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T01()

        Dim oCol As List(Of SousCommande)
        Dim oLG As LgCommande
        Dim objSCMD As SousCommande

        'Ajout de Taux de commission sur les Fournisseur et le Client
        Dim oTx As TauxComm
        oTx = New TauxComm(m_objFRN, m_objCLT.codeTypeClient, 15.0)
        Assert.IsTrue(oTx.save(), oTx.getErreur())
        oTx = New TauxComm(m_objFRN2, m_objCLT.codeTypeClient, 18.0)
        Assert.IsTrue(oTx.save(), oTx.getErreur())

        'Ajout de 2 lignes sur la commande
        oLG = m_oCmd.AjouteLigne("10", m_objPRD, 1, 150)
        Assert.AreEqual(CDec(15.0), oLG.TxComm)
        oLG = m_oCmd.AjouteLigne("20", m_objPRD2, 2, 160)
        Assert.AreEqual(CDec(18.0), oLG.TxComm)
        m_oCmd.save()

        m_oCmd.changeEtat(vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save())

        'Génération des Sous-Commandes
        m_oCmd.generationSousCommande()
        Assert.IsTrue(m_oCmd.save(), "Ocmd.Save" & m_oCmd.getErreur())


        oCol = SousCommande.getListe(#6/2/1964#, #6/2/1964#)
        Assert.AreEqual(2, oCol.Count)
        objSCMD = oCol(0)
        Assert.AreEqual(CDec(15), objSCMD.tauxCommission)
        objSCMD = oCol(1)
        Assert.AreEqual(CDec(18), objSCMD.tauxCommission)






    End Sub

    ''' <summary>
    ''' test le non ecrasement des mouvements de stocks en sauvegarde article
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T02_ColMvtStock()
        Dim oLgCMD As LgCommande
        Dim nMvt As Integer
        Dim oPRD As Produit

        'Ajout d'un ligne de mouvement de stock manuellement
        m_objPRD.loadcolmvtStock()
        nMvt = m_objPRD.colmvtStock.Count
        m_objPRD.ajouteLigneMvtStock(CDate("06/02/1964"), vncTypeMvt.vncmvtRegul, 0, "Test", 10)


        'Génération d'une ligne de mouvement de stock à partir d'une commande
        oPRD = Produit.createandload(m_objPRD.id)
        'Ajout d'un ligne de commande
        m_oCmd.AjouteLigne("64", oPRD, 12, 10.5)
        Assert.IsTrue(m_oCmd.save())
        'Livraison de la commande => generattion de mouvement de stock
        For Each oLgCMD In m_oCmd.colLignes
            oLgCMD.qteLiv = oLgCMD.qteCommande
        Next
        m_oCmd.changeEtat(vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save())

        'Sauvergarde du Produit précédemment Chargé
        Assert.IsTrue(m_objPRD.save())

        'Rechargement du produit
        m_objPRD.load()
        m_objPRD.loadcolmvtStock()
        'Vérification du nombre de mouvement de stock
        Assert.AreEqual(nMvt + 2, m_objPRD.colmvtStock.Count)

    End Sub
End Class


