'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V223
    Inherits test_Base

    Private m_objPRD As Produit
    Private m_objPRD2 As Produit
    Private m_objPRD3 As Produit
    Private m_objFRN As Fournisseur
    Private m_objFRN2 As Fournisseur
    Private m_objFRN3 As Fournisseur
    Private m_objCLT As Client
    Private m_oCmd As CommandeClient
    <TestInitialize()> Public Overrides Sub TestInitialize()
        Dim col As Collection
        Dim oTaux As TauxComm

        MyBase.TestInitialize()

        Persist.shared_connect()
        Param.LoadcolParams()

        col = Fournisseur.getListe("FRNV223")
        If col.Count > 0 Then
            For Each m_objFRN In col
                m_objFRN.bDeleted = True
                m_objFRN.Save()
            Next
        End If
        col = Fournisseur.getListe("FRNV223-2")
        If col.Count > 0 Then
            For Each m_objFRN In col
                m_objFRN.bDeleted = True
                m_objFRN.Save()
            Next
        End If
        col = Fournisseur.getListe("FRNV223-3")
        If col.Count > 0 Then
            For Each m_objFRN In col
                m_objFRN.bDeleted = True
                m_objFRN.Save()
            Next
        End If

        col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, "PRDV223")
        If col.Count > 0 Then
            For Each m_objPRD In col
                m_objPRD.bDeleted = True
                m_objPRD.save()
            Next
        End If

        col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, "PRDV223-2")
        If col.Count > 0 Then
            For Each m_objPRD In col
                m_objPRD.bDeleted = True
                m_objPRD.save()
            Next
        End If

        col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, "PRDV223-3")
        If col.Count > 0 Then
            For Each m_objPRD In col
                m_objPRD.bDeleted = True
                m_objPRD.save()
            Next
        End If
        col = Client.getListe("CLTV223")
        If col.Count > 0 Then
            For Each m_objCLT In col
                m_objCLT.bDeleted = True
                m_objCLT.save()
            Next
        End If
        m_objFRN = New Fournisseur("FRNV223", "FRn de'' test")
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

        m_objPRD = New Produit("PRDV223", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objFRN2 = New Fournisseur("FRNV223-2", "FRn de'' test2")
        m_objFRN2.rs = "FRNV223-2"
        Assert.IsTrue(m_objFRN2.Save(), "FRN.Create")

        m_objPRD2 = New Produit("PRDV223-2", m_objFRN2, 1990)
        Assert.IsTrue(m_objPRD2.save(), "Prod.Create")

        m_objFRN3 = New Fournisseur("FRNV223-3", "FRn de'' test3")
        m_objFRN3.rs = "FRNV223-3"
        Assert.IsTrue(m_objFRN3.Save(), "FRN3.Create")

        m_objPRD3 = New Produit("PRDV223-3", m_objFRN3, 1990)
        Assert.IsTrue(m_objPRD3.save(), "Prod3.Create")

        m_objCLT = New Client("CLTV223", "Client de' test")
        m_objCLT.rs = "Client de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")

        'Creation des Taux de Commissions
        oTaux = New TauxComm(m_objFRN, m_objCLT.codeTypeClient, 9.5)
        oTaux.save()
        oTaux = New TauxComm(m_objFRN2, m_objCLT.codeTypeClient, 9.5)
        oTaux.save()
        oTaux = New TauxComm(m_objFRN3, m_objCLT.codeTypeClient, 9.5)
        oTaux.save()


        col = CommandeClient.getListe(CDate("06/02/1964"), CDate("06/02/1964"))
        For Each m_oCmd In col
            m_oCmd.bDeleted = True
            m_oCmd.save()
        Next
        m_oCmd = New CommandeClient(m_objCLT)
        m_oCmd.dateCommande = CDate("06/02/1964")
        m_oCmd.save()


        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        MyBase.TestCleanup()
    End Sub
    '"Test done TestV300"
    <TestMethod(), Ignore()> Public Sub T10_EXPORT_AVEC_COM_CARSPECIAUX()
        'Verification que L'export WebEDI supprime bien les retours chariots dans les commantaires de livraison
        Dim objCMD As CommandeClient
        Dim nFile As Integer
        Dim strResult As String
        Dim nLineNumber As Integer

        'Creation d'une Commande
        objCMD = New CommandeClient(m_objCLT)
        objCMD.dateCommande = "06/02/2000"
        objCMD.CommCommande.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommFacturation.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLibre.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLivraison.comment = "1234" + vbCrLf + "5678" + vbLf + "90123" + vbCr + "456" + vbNullChar + "7890" + vbTab + "1234" + vbBack + "56789012345678901234567890"
        objCMD.AjouteLigne("10", m_objPRD, 10, 10)
        If System.IO.File.Exists("./adel.txt") Then
            System.IO.File.Delete("./adel.txt")
        End If
        objCMD.exporterWebEDI("./adel.txt")
        nFile = FreeFile()
        FileOpen(nFile, "./adel.txt", OpenMode.Input, OpenAccess.Read)
        nLineNumber = 0
        While Not EOF(nFile)
            nLineNumber = nLineNumber + 1
            strResult = LineInput(nFile)
            Console.WriteLine(strResult)
        End While

        Assert.AreEqual(1, nLineNumber, "Une seule Ligne de fichier")
        Assert.AreEqual("010", Mid(strResult, 242, 3), "Le Numéro lde ligne est sur 249 ")

        FileClose(nFile)
        'Suppression du fichier créé
        System.IO.File.Delete("./adel.txt")

    End Sub


    <TestMethod()> Public Sub T61_bExportInternet()
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim oSCmd2 As SousCommande
        Dim tabCSV As String() = Nothing

        'Ajout de 2 Lignes à la commande
        oLgCmd = m_oCmd.AjouteLigne("10", m_objPRD, 10, 15.55)
        oLgCmd.qteLiv = 9
        oLgCmd = m_oCmd.AjouteLigne("20", m_objPRD2, 8, 5.77)
        oLgCmd.qteLiv = 7

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


        For Each oSCmd In m_oCmd.colSousCommandes
            Assert.IsFalse(oSCmd.bExportInternet, "Export internet à false par défaut")
            oSCmd.bExportInternet = True
            Assert.IsTrue(oSCmd.bExportInternet, "Export internet à true")
            Assert.IsTrue(oSCmd.Save())
            oSCmd2 = SousCommande.createandload(oSCmd.id)
            Assert.IsTrue(oSCmd2.bExportInternet, "Export internet à true après rechargement")
        Next oSCmd

        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")
    End Sub

    <TestMethod()> Public Sub T65_ListeExport()
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim col As List(Of SousCommande)


        'Ajout de 2 Lignes à la commande
        oLgCmd = m_oCmd.AjouteLigne("10", m_objPRD, 10, 15.55)
        oLgCmd.qteLiv = 9
        oLgCmd = m_oCmd.AjouteLigne("20", m_objPRD2, 8, 5.77)
        oLgCmd.qteLiv = 7

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")

        'Lecture des sous commandes exportées => 0 elements
        col = SousCommande.getListeExportee(CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(0, col.Count, "0 Elements dans la Liste")

        'Test du bollean ExportIternet du Fournisseur
        m_objFRN.bExportInternet = 1
        m_objFRN.Save()
        m_objFRN2.bExportInternet = 0
        m_objFRN2.Save()
        'Lecture des sous commandes Non exportées => 1 elements caril  n'y a qu'un fournisseur internet
        col = SousCommande.getListeAExporterInternet(vncOrigineCmd.vncVinicom, CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(1, col.Count, "2 Elements dans la Liste")

        'Test du bollean ExportIternet du Fournisseur
        m_objFRN.bExportInternet = 0
        m_objFRN.Save()
        m_objFRN2.bExportInternet = 0
        m_objFRN2.Save()
        'Lecture des sous commandes Non exportées => 0 elements caril  n'y a qu'un fournisseur internet
        col = SousCommande.getListeAExporterInternet(vncOrigineCmd.vncVinicom, CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(0, col.Count, "0 Elements dans la Liste")

        'Test du bollean ExportIternet du Fournisseur
        m_objFRN.bExportInternet = 1
        m_objFRN.Save()
        m_objFRN2.bExportInternet = 1
        m_objFRN2.Save()
        'Lecture des sous commandes Non exportées => 2 elements 
        col = SousCommande.getListeAExporterInternet(vncOrigineCmd.vncVinicom, CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(2, col.Count, "2 Elements dans la Liste")
        For Each oSCmd In col
            Assert.IsFalse(oSCmd.bExportInternet, "L'export internet doit être à false")
            'Passage du boolean à true et sauvegarde
            oSCmd.load()
            oSCmd.bExportInternet = True
            oSCmd.changeEtat(vncActionEtatCommande.vncActionSCMDExportInternet)
            oSCmd.Save()
        Next

        'Lecture des sous commandes non exportée => 0 elements
        col = SousCommande.getListeAExporterInternet(vncOrigineCmd.vncVinicom, CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(0, col.Count, "0 Elements dans la Liste")

        'Lecture des sous commandes exportées => 2 elements
        col = SousCommande.getListeExportee(CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(2, col.Count, "2 Elements dans la Liste")
        For Each oSCmd In col
            Assert.IsTrue(oSCmd.bExportInternet, "L'export internet doit être à true")
            'Passage du boolean à False et sauvegarde
            oSCmd.load()
            oSCmd.bExportInternet = False
            oSCmd.Save()
        Next


        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livre")

    End Sub 'T65_ListeExport
    <TestMethod()> Public Sub T70_ListeSCMDTrasmises()
        'Test lafonction SousCommde.getListeTransmises = TransmiseFax + ExportéeInternet
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim col As List(Of SousCommande)


        'Ajout de 3 Lignes à la commandes
        oLgCmd = m_oCmd.AjouteLigne("10", m_objPRD, 10, 15.55)
        oLgCmd.qteLiv = 9
        oLgCmd = m_oCmd.AjouteLigne("20", m_objPRD2, 8, 5.77)
        oLgCmd.qteLiv = 7
        oLgCmd = m_oCmd.AjouteLigne("30", m_objPRD3, 8.5, 5.77)
        oLgCmd.qteLiv = 7.5

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")

        'Lecture des sous commandes exportées => 0 elements
        col = SousCommande.getListeTransmises(CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(0, col.Count, "0 Elements dans la Liste")

        'Transmission des SousCommandes
        'La première est faxée
        oSCmd = m_oCmd.colSousCommandes(1)
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
        oSCmd.Save()

        'La Seconde est Exportée par internet
        oSCmd = m_oCmd.colSousCommandes(2)
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
        oSCmd.Save()

        'Lecture des sous commandes transmises => 2 elements
        col = SousCommande.getListeTransmises(CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(2, col.Count, "2 Elements dans la Liste")

        'Transmission des SousCommandes
        'La Troisième est faxée
        oSCmd = m_oCmd.colSousCommandes(3)
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
        oSCmd.Save()

        'Lecture des sous commandes trasmises => 3 elements
        col = SousCommande.getListeTransmises(CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(3, col.Count, "3 Elements dans la Liste")
        For Each oSCmd In col
            Assert.IsTrue((oSCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncSCMDtransmiseFax Or oSCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncSCMDExporteeInt), "Code Etat =TransmiseparFax ou exportée internet")
        Next

        'Lecture des sous commandes trasmises par fax => 2 elements
        col = SousCommande.getListe(CDate("06/02/1964"), CDate("06/02/1964"), , vncEnums.vncEtatCommande.vncSCMDtransmiseFax)
        Assert.AreEqual(2, col.Count, "2 Elements dans la Liste")

        'Lecture des sous commandes trasmises par internet=> 1 elements
        col = SousCommande.getListe(CDate("06/02/1964"), CDate("06/02/1964"), , vncEnums.vncEtatCommande.vncSCMDExporteeInt)
        Assert.AreEqual(1, col.Count, "1 Elements dans la Liste")


        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livre")

    End Sub 'T70_ListSCMDTrasmises
    <TestMethod()> Public Sub T75_ListeSCMDAProvionner()
        'Test lafonction SousCommde.getListeAProvionnner = Rapprochée + RapprochéeInternet
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande
        Dim col As List(Of SousCommande)


        'Ajout de 3 Lignes à la commandes
        oLgCmd = m_oCmd.AjouteLigne("10", m_objPRD, 10, 15.55)
        oLgCmd.qteLiv = 9
        oLgCmd = m_oCmd.AjouteLigne("20", m_objPRD2, 8, 5.77)
        oLgCmd.qteLiv = 7
        oLgCmd = m_oCmd.AjouteLigne("30", m_objPRD3, 8.5, 5.77)
        oLgCmd.qteLiv = 7.5

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


        'Transmission des SousCommandes
        'La première est faxée puis rapprochée
        oSCmd = m_oCmd.colSousCommandes(1)
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
        oSCmd.Save()
        oSCmd.dateFactFournisseur = CDate("06/02/1964")
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
        oSCmd.Save()

        'La Seconde est Exportée par internet puis importée
        oSCmd = m_oCmd.colSousCommandes(2)
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
        oSCmd.Save()
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
        oSCmd.dateFactFournisseur = CDate("06/02/1964")
        oSCmd.Save()

        'Transmission des SousCommandes
        'La Troisième est faxée mias pas rapprochée
        oSCmd = m_oCmd.colSousCommandes(3)
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
        oSCmd.Save()

        'Lecture des sous commandes A provionner => 2 elements
        col = SousCommande.getListeAProvisionner(CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(2, col.Count, "2 Elements dans la Liste")
        'La Troisième est rapprochée
        oSCmd = m_oCmd.colSousCommandes(3)
        oSCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
        oSCmd.Save()
        'Lecture des sous commandes A provionner => 3 elements
        col = SousCommande.getListeAProvisionner(CDate("06/02/1964"), CDate("06/02/1964"))
        Assert.AreEqual(3, col.Count, "2 Elements dans la Liste")
        For Each oSCmd In col
            Assert.IsTrue((oSCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncSCMDRapprochee Or oSCmd.etat.codeEtat = vncEnums.vncEtatCommande.vncSCMDRapprocheeInt), "Code Etat =Rapprochée ou Rapprochée internet")
        Next

        'Lecture des sous commandes Rapprochee => 2 elements
        col = SousCommande.getListe(CDate("06/02/1964"), CDate("06/02/1964"), , vncEnums.vncEtatCommande.vncSCMDRapprochee)
        Assert.AreEqual(2, col.Count, "1 Elements dans la Liste")

        'Lecture des sous commandes Rapprochee Internet=> 1 elements
        col = SousCommande.getListe(CDate("06/02/1964"), CDate("06/02/1964"), , vncEnums.vncEtatCommande.vncSCMDRapprocheeInt)
        Assert.AreEqual(1, col.Count, "1 Elements dans la Liste")


        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livre")

    End Sub 'T75_ListSCMDAProvisionner
    <TestMethod()> Public Sub T80_CalcCommision()
        'Test lafonction SousCommde.CalcCommsionStandard= 
        Dim oLgCmd As LgCommande
        Dim oSCmd As SousCommande



        'Ajout de 3 Lignes à la commandes
        oLgCmd = m_oCmd.AjouteLigne("10", m_objPRD, 10, 15.55)
        oLgCmd.qteLiv = 9
        oLgCmd = m_oCmd.AjouteLigne("20", m_objPRD2, 8, 5.77)
        oLgCmd.qteLiv = 7
        oLgCmd = m_oCmd.AjouteLigne("30", m_objPRD3, 8.5, 5.77)
        oLgCmd.qteLiv = 7.5

        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Encours")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livrée")
        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(m_oCmd.generationSousCommande(), "Génération des sous-commandes")
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Eclater")


        oSCmd = m_oCmd.colSousCommandes(1)
        Assert.AreEqual(oSCmd.totalHT, oSCmd.baseCommission, "Base comm = total HT")
        Assert.AreEqual(CDec(9.5), oSCmd.tauxCommission, "Taux de commission Standard")
        Assert.AreEqual(Math.Round(CDec(oSCmd.totalHT * 9.5 / 100), 2), oSCmd.MontantCommission, "Calcul Commission")

        oSCmd.totalHTFacture = 456.52
        oSCmd.calcCommisionstandard(CalculCommScmd.CALCUL_COMMISSION_HT_FACTURE)
        Assert.AreEqual(oSCmd.totalHTFacture, oSCmd.baseCommission, "Base comm = total HT")
        Assert.AreEqual(CDec(9.5), oSCmd.tauxCommission, "Taux de commission Standard")
        Assert.AreEqual(Math.Round(CDec(oSCmd.totalHTFacture * 9.5 / 100), 2), oSCmd.MontantCommission, "Calcul Commission")


        m_oCmd.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        Assert.IsTrue(m_oCmd.save, "Sauvegarde de la commande client Etat = Livre")

    End Sub 'T80_CalcCommision

End Class


