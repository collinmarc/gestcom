Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.Collections.Generic
Imports System.IO
<TestClass()> Public Class TestExportQuadra
    Inherits test_Base
    Private m_oProduit As Produit
    Private m_oProduit2 As Produit
    Private m_oFourn As Fournisseur
    Private m_oFourn2 As Fournisseur
    Private m_oClient As Client
    Private m_oCmd As CommandeClient
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Dim col As Collection
        Dim oTaux As TauxComm
        Dim strCode As String

        strCode = "FRNT1110"
        m_oFourn = New Fournisseur(strCode, "MonFournisseur")
        m_oFourn.AdresseFacturation.nom = "ADF_Nom"
        m_oFourn.AdresseFacturation.rue1 = "ADF_Nom"
        m_oFourn.AdresseFacturation.rue2 = "ADF_Nom"
        m_oFourn.AdresseFacturation.cp = "ADF_Nom"
        m_oFourn.AdresseFacturation.ville = "ADF_Nom"
        m_oFourn.AdresseFacturation.tel = "01010101"
        m_oFourn.AdresseFacturation.fax = "02020202"
        m_oFourn.AdresseFacturation.port = "03030303"
        m_oFourn.AdresseFacturation.Email = "04040404"
        m_oFourn.Save()

        strCode = "FRN2T1110"
        col = Fournisseur.getListe(strCode)
        For Each m_oFourn2 In col
            m_oFourn2.bDeleted = True
            m_oFourn2.Save()
        Next
        m_oFourn2 = New Fournisseur(strCode, "MonFournisseur")
        m_oFourn2.Save()

        strCode = "TSTPRDT1110"
        col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, strCode)
        For Each m_oProduit In col
            m_oProduit.bDeleted = True
            m_oProduit.save()
        Next
        m_oProduit = New Produit(strCode, m_oFourn, 1990)
        m_oProduit.save()

        strCode = "TSTPRD2T1110"
        col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, strCode)
        For Each m_oProduit2 In col
            m_oProduit2.bDeleted = True
            m_oProduit2.save()
        Next
        m_oProduit2 = New Produit(strCode, m_oFourn2, 1990)
        m_oProduit2.save()

        strCode = "CLTT1110"
        col = Client.getListe(strCode)
        For Each m_oClient In col
            m_oClient.bDeleted = True
            m_oClient.save()
        Next
        m_oClient = New Client(strCode, "MonClient")
        Debug.Assert(m_oClient.save(), "Creation du client")
        '            m_oClient()

        'Creation des Taux de Commissions
        oTaux = New TauxComm(m_oFourn, m_oClient.codeTypeClient, 9.5)
        oTaux.save()
        oTaux = New TauxComm(m_oFourn2, m_oClient.codeTypeClient, 9.5)
        oTaux.save()

        'Suppression des Commandes de 1964
        col = CommandeClient.getListe(#6/2/1964#, #6/2/1964#)
        For Each m_oCmd In col
            m_oCmd.bDeleted = True
            m_oCmd.save()
        Next
        m_oCmd = New CommandeClient(m_oClient)
        m_oCmd.dateCommande = #6/2/1964#
        m_oCmd.save()



    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        MyBase.TestCleanup()

    End Sub
    ''' <summary>
    ''' Test de l'export d'une souscommande vers QuadraFacturation
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T101_ExportQuadra()
        Dim objCMD As CommandeClient
        Dim oSCMD As SousCommande
        Dim oLg As LgCommande
        Dim strFile As String


        'Creation d'une Commande
        objCMD = New CommandeClient(m_oClient)
        oLg = objCMD.AjouteLigne("10", m_oProduit, 15, 15.5)
        oLg.qteLiv = oLg.qteCommande
        objCMD.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMD.save()

        Assert.IsTrue(objCMD.generationSousCommande())
        Assert.AreEqual(1, objCMD.colSousCommandes.Count)
        oSCMD = objCMD.colSousCommandes(1)

        strFile = "./adel.csv"
        If System.IO.File.Exists(strFile) Then
            System.IO.File.Delete(strFile)
        End If

        oSCMD.toCSVQuadraFact(strFile, vncTypeExportQuadra.vncExportBafClient)

        Assert.IsTrue(System.IO.File.Exists(strFile))
        'Il y a Bien 2 lignes dans le fichier( 1 entête et une ligne)
        Assert.AreEqual(2, System.IO.File.ReadAllLines(strFile).Count)
        Assert.IsTrue(oSCMD.ValiderExportQuadra())
        Assert.IsTrue(oSCMD.bExportQuadra)
        '        Assert.AreEqual(vncEnums.vncEtatCommande.vncSCMDFacturee, oSCMD.EtatCode)
        '        oSCMD.delete()
        objCMD.delete()
    End Sub

    ''' <summary>
    ''' Test de l'export d'une souscommande vers QuadraFacturation
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T102_ExportQuadraSCommande()
        Dim objCMD As CommandeClient
        Dim oSCMD As SousCommande
        Dim oLg As LgCommande
        Dim strFile As String

        m_oFourn.bExportInternet = vncTypeExportScmd.vncExportInternet 'Fournisseur Vinicom
        m_oFourn.Save()
        m_oProduit.idFournisseur = m_oFourn.id
        m_oProduit.save()

        m_oFourn2.bExportInternet = vncTypeExportScmd.vncExportQuadra 'Fournisseur HOBIVIN
        m_oFourn2.Save()
        m_oProduit2.idFournisseur = m_oFourn2.id
        m_oProduit2.save()

        'Creation d'une Commande "VINICOM" avec un produit Vinicom et un produit HOBIVIN
        objCMD = New CommandeClient(m_oClient)
        objCMD.Origine = Dossier.VINICOM
        oLg = objCMD.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit Vinicom
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMD.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit HOBIVIN
        oLg.qteLiv = oLg.qteCommande
        objCMD.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMD.changeEtat(vncActionEtatCommande.vncActionLivrer)

        objCMD.save()

        'Commande Vinicom Donc pas d'intermédiaire
        Assert.IsTrue(objCMD.generationSousCommande())
        Assert.AreEqual(2, objCMD.colSousCommandes.Count)  '2 sous Commandes générées


        'Export de la SECONDE SousCommande (CMD VINICOM, PROD HOBIVIN)
        'Export de type BAFClient (Facture à générer par QuadraFact)
        oSCMD = objCMD.colSousCommandes(2)
        Assert.AreEqual(CInt(vncTypeExportScmd.vncExportQuadra), oSCMD.oFournisseur.bExportInternet)

        strFile = "./adel.csv"
        If System.IO.File.Exists(strFile) Then
            System.IO.File.Delete(strFile)
        End If
        oSCMD.toCSVQuadraFact(strFile, vncTypeExportQuadra.vncExportBafClient)

        Assert.IsTrue(System.IO.File.Exists(strFile))
        'Il y a Bien 2 lignes dans le fichier( 1 entête et une ligne)
        Assert.AreEqual(2, System.IO.File.ReadAllLines(strFile).Count)
        Debug.WriteLine(File.ReadAllText(strFile))

        Dim strLine As String = File.ReadAllLines(strFile)(1)
        Dim tab As String()
        tab = strLine.Split(";")
        Assert.AreEqual(tab(0), m_oClient.code)
        Assert.AreEqual(tab(1), oSCMD.codeCommandeClient)
        Assert.AreEqual(tab(2), Format(oSCMD.dateCommande, "yyMMdd"))
        Assert.AreEqual(tab(3), oSCMD.codeCommandeClient)
        Assert.AreEqual(tab(4), m_oProduit2.code)
        Assert.AreEqual(tab(5), "25")
        Assert.AreEqual(tab(6), "25.5")


        'Export de la Première SousCommande (CMD VINICOM, PROD VINICOM)
        'Normalement cette Souscommande n'est pas Exportée sous Quadra
        'même traitement qu'une commande "HOBIVIN" avec un produit VINICOM
        'Export de type BAFFournisseur (Mise à jour des stocks dans QuadraFact)
        m_oCmd.Origine = Dossier.HOBIVIN

        oSCMD = objCMD.colSousCommandes(1)
        strFile = "./adel.csv"
        If System.IO.File.Exists(strFile) Then
            System.IO.File.Delete(strFile)
        End If
        oSCMD.toCSVQuadraFact(strFile, vncTypeExportQuadra.vncExportBaFournisseur)

        Assert.IsTrue(System.IO.File.Exists(strFile))
        'Il y a Bien 2 lignes dans le fichier( 1 entête et une ligne)
        Assert.AreEqual(2, System.IO.File.ReadAllLines(strFile).Count)
        Debug.WriteLine(File.ReadAllText(strFile))

        strLine = File.ReadAllLines(strFile)(1)
        tab = strLine.Split(";")
        Assert.AreEqual(tab(0), oSCMD.oFournisseur.code) 'Cette fois c'est le code fournisseur
        Assert.AreEqual(tab(1), oSCMD.code)
        Assert.AreEqual(tab(2), Format(oSCMD.dateCommande, "yyMMdd"))
        Assert.AreEqual(tab(3), oSCMD.code.Replace("-", ""))
        Assert.AreEqual(tab(4), m_oProduit.code)
        Assert.AreEqual(tab(5), "15")
        Assert.AreEqual(tab(6), "15.5")

        objCMD.delete()
    End Sub

    ''' <summary>
    ''' Test de l'export d'une commande HOBIVIN vers QuadraFacturation
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T101_ExportQuadraCmdHBV()
        Dim objCMD As CommandeClient
        Dim oSCMD As SousCommande
        Dim oLg As LgCommande
        Dim strFile As String
        Dim oCltInter As Client
        Dim strLine As String
        Dim TAB As String()

        m_oFourn2.bExportInternet = vncTypeExportScmd.vncExportQuadra 'Fournisseur HOBIVIN
        m_oFourn2.Save()

        m_oProduit2.idFournisseur = m_oFourn2.id
        m_oProduit2.save()

        oCltInter = New Client("CLTINTER", "Client intermédiare")
        oCltInter.idTypeClient = Param.getTypeClientIntermediaire().id
        oCltInter.Origine = Dossier.HOBIVIN
        oCltInter.save()

        'Creation d'une Commande
        objCMD = New CommandeClient(m_oClient)
        objCMD.Origine = Dossier.HOBIVIN
        oLg = objCMD.AjouteLigne("10", m_oProduit2, 15, 15.5)
        oLg.qteLiv = oLg.qteCommande
        objCMD.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMD.save()

        Assert.IsTrue(objCMD.generationSousCommande())

        strFile = "./adel.csv"
        If System.IO.File.Exists(strFile) Then
            System.IO.File.Delete(strFile)
        End If

        objCMD.toCSVQuadraFact(strFile, vncTypeExportQuadra.vncExportBafClient)

        Assert.IsTrue(System.IO.File.Exists(strFile))
        'Il y a Bien 2 lignes dans le fichier( 1 entête et une ligne)
        Assert.AreEqual(2, System.IO.File.ReadAllLines(strFile).Count)
        Debug.WriteLine(File.ReadAllText(strFile))

        strLine = File.ReadAllLines(strFile)(1)
        TAB = strLine.Split(";")
        Assert.AreEqual(TAB(0), m_oClient.code) 'Cette fois c'est le code Client
        Assert.AreEqual(TAB(1), objCMD.code)
        Assert.AreEqual(TAB(2), Format(objCMD.dateCommande, "yyMMdd"))
        Assert.AreEqual(TAB(3), objCMD.code)
        Assert.AreEqual(TAB(4), m_oProduit2.code)
        Assert.AreEqual(TAB(5), "15")
        Assert.AreEqual(TAB(6), "15.5")


        objCMD.delete()
    End Sub

    ''' <summary>
    ''' Test de la Fonction LoadCmd de ExportQuadra
    ''' </summary>
    <TestMethod()> Public Sub T110_ExportQuadraLoadCmd()
        Dim objCMDV As CommandeClient
        Dim objCMDH As CommandeClient
        Dim oSCMD As SousCommande
        Dim oLg As LgCommande

        'Founisseur Vinicom
        m_oFourn.bExportInternet = vncTypeExportScmd.vncExportInternet 'Fournisseur Vinicom
        m_oFourn.Save()
        m_oProduit.idFournisseur = m_oFourn.id
        m_oProduit.save()

        'Fournisseur HOBIVIN
        m_oFourn2.bExportInternet = vncTypeExportScmd.vncExportQuadra 'Fournisseur HOBIVIN
        m_oFourn2.Save()
        m_oProduit2.idFournisseur = m_oFourn2.id
        m_oProduit2.save()

        'Client intermédiaire
        Dim oCltIntermediaire As Client
        oCltIntermediaire = Client.getIntermediairePourUneOrigine(Dossier.HOBIVIN)
        If oCltIntermediaire Is Nothing Then
            oCltIntermediaire = New Client("CLTINTER", "ClientIntermédiaire")
            oCltIntermediaire.setTypeIntermediaire(Dossier.HOBIVIN)
            oCltIntermediaire.save()
        End If


        'Creation d'une Commande "VINICOM" avec un produit Vinicom et un produit HOBIVIN
        objCMDV = New CommandeClient(m_oClient)
        objCMDV.dateCommande = New Date(1964, 2, 6)
        objCMDV.Origine = Dossier.VINICOM
        oLg = objCMDV.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit Vinicom
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMDV.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit HOBIVIN
        oLg.qteLiv = oLg.qteCommande
        objCMDV.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMDV.changeEtat(vncActionEtatCommande.vncActionLivrer)
        objCMDV.save()
        'Commande Vinicom Donc pas d'intermédiaire
        Assert.IsTrue(objCMDV.generationSousCommande())
        objCMDV.save()

        Assert.AreEqual(2, objCMDV.colSousCommandes.Count)  '2 sous Commandes générées

        'Creation d'une Commande "HOBIVIN" avec un produit Vinicom et un produit HOBIVIN
        objCMDH = New CommandeClient(m_oClient)
        objCMDH.dateCommande = New Date(1964, 2, 7)
        objCMDH.Origine = Dossier.HOBIVIN
        oLg = objCMDH.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit Vinicom
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMDH.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit HOBIVIN
        oLg.qteLiv = oLg.qteCommande
        objCMDH.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMDH.changeEtat(vncActionEtatCommande.vncActionLivrer)
        objCMDH.save()
        'Commande HOBIVIN Donc un intermédiaire
        Assert.IsTrue(objCMDH.generationSousCommande())
        objCMDH.save()
        Assert.AreEqual(2, objCMDH.colSousCommandes.Count)  '2 sous Commandes générées

        Dim oExport As New ExportQuadra(New Date(1964, 2, 1), New Date(1964, 3, 1), vncTypeExportQuadra.vncExportBafClient, ".", True)

        'Export de type Baf Client
        oExport.typeExport = vncTypeExportQuadra.vncExportBafClient
        '==============
        'Chagement de la liste des sousCommandes à Exporter
        '=============
        oExport.LoadListCmd()

        Assert.IsNotNull(oExport.ListCmd)
        Assert.AreEqual(2, oExport.ListCmd.Count)

        'Premier Element à Exporter
        'une Souscommande issu de la commande VINICOM concernant le Fournisseur HOBIVIN
        Assert.IsInstanceOfType(oExport.ListCmd(0), New SousCommande().GetType())
        oSCMD = oExport.ListCmd(0)
        oSCMD.oFournisseur.load()
        Assert.AreEqual(CInt(vncEnums.vncTypeExportScmd.vncExportQuadra), oSCMD.oFournisseur.bExportInternet)
        Assert.AreEqual(Dossier.VINICOM, CommandeClient.createandload(oSCMD.idCommandeClient).Origine)
        Assert.AreEqual(m_oClient.code, oSCMD.TiersCode)

        'Second Element à Exporter
        'une commande HOBIVIN avec les 2 produits
        Assert.IsInstanceOfType(oExport.ListCmd(1), objCMDH.GetType())
        objCMDH = oExport.ListCmd(1)
        Assert.AreEqual(Dossier.HOBIVIN, objCMDH.Origine)
        Assert.AreEqual(2, objCMDH.colLignes.Count)
        Assert.AreEqual(m_oClient.code, oSCMD.TiersCode)

        'Export de type Ba Fournisseur
        oExport.typeExport = vncTypeExportQuadra.vncExportBaFournisseur
        oExport.LoadListCmd()
        Assert.IsNotNull(oExport.ListCmd)

        Assert.AreEqual(1, oExport.ListCmd.Count)
        Assert.IsInstanceOfType(oExport.ListCmd(0), New SousCommande().GetType())
        oSCMD = oExport.ListCmd(0)
        oSCMD.oFournisseur.load()
        Assert.AreEqual(CInt(vncEnums.vncTypeExportScmd.vncExportInternet), oSCMD.oFournisseur.bExportInternet)
        Assert.AreEqual(Dossier.HOBIVIN, CommandeClient.createandload(oSCMD.idCommandeClient).Origine)
        Assert.AreEqual(oCltIntermediaire.code, oSCMD.TiersCode)


        objCMDV.bDeleted = True
        objCMDV.save()

        oCltIntermediaire.bDeleted = True
        oCltIntermediaire.save()
    End Sub

    <TestMethod()>
    Public Sub TestValidationExportBafClient()

        Dim objCMDV As CommandeClient
        Dim objCMDH As CommandeClient
        Dim oSCMD As SousCommande
        Dim oLg As LgCommande

        'Founisseur Vinicom
        m_oFourn.bExportInternet = vncTypeExportScmd.vncExportInternet 'Fournisseur Vinicom
        m_oFourn.Save()
        m_oProduit.idFournisseur = m_oFourn.id
        m_oProduit.save()

        'Fournisseur HOBIVIN
        m_oFourn2.bExportInternet = vncTypeExportScmd.vncExportQuadra 'Fournisseur HOBIVIN
        m_oFourn2.Save()
        m_oProduit2.idFournisseur = m_oFourn2.id
        m_oProduit2.save()

        'Client intermédiaire
        Dim oInter As New Client("CLTINTER", Dossier.HOBIVIN)
        oInter.idTypeClient = Param.getTypeClientIntermediaire().id
        oInter.save()


        'Creation d'une Commande "VINICOM" avec un produit Vinicom et un produit HOBIVIN
        objCMDV = New CommandeClient(m_oClient)
        objCMDV.dateCommande = New Date(1964, 2, 6)
        objCMDV.Origine = Dossier.VINICOM
        oLg = objCMDV.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit Vinicom
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMDV.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit HOBIVIN
        oLg.qteLiv = oLg.qteCommande
        objCMDV.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMDV.changeEtat(vncActionEtatCommande.vncActionLivrer)
        objCMDV.save()
        'Commande Vinicom Donc pas d'intermédiaire
        Assert.IsTrue(objCMDV.generationSousCommande())
        objCMDV.save()

        Assert.AreEqual(2, objCMDV.colSousCommandes.Count)  '2 sous Commandes générées

        'Creation d'une Commande "HOBIVIN" avec un produit Vinicom et un produit HOBIVIN
        objCMDH = New CommandeClient(m_oClient)
        objCMDH.dateCommande = New Date(1964, 2, 7)
        objCMDH.Origine = Dossier.HOBIVIN
        oLg = objCMDH.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit Vinicom
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMDH.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit HOBIVIN
        oLg.qteLiv = oLg.qteCommande
        objCMDH.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMDH.changeEtat(vncActionEtatCommande.vncActionLivrer)
        objCMDH.save()
        'Commande HOBIVIN Donc un intermédiaire
        Assert.IsTrue(objCMDH.generationSousCommande())
        objCMDH.save()
        Assert.AreEqual(2, objCMDH.colSousCommandes.Count)  '2 sous Commandes générées

        Dim oExport As New ExportQuadra(New Date(1964, 2, 1), New Date(1964, 3, 1), vncTypeExportQuadra.vncExportBafClient, ".", True)

        'Export de type Baf Client
        oExport.typeExport = vncTypeExportQuadra.vncExportBafClient
        oExport.LoadListCmd()
        oExport.ExportBaf()

        'Vérification des etats avant la validation
        Dim oCmdClient As CommandeClient
        For Each oCmd As Commande In oExport.ListCmd
            If (TypeOf (oCmd) Is CommandeClient) Then
                'La Commande est Eclatées
                Assert.AreEqual(vncEtatCommande.vncEclatee, oCmd.EtatCode)
                oCmdClient = oCmd
                For Each oSCMD In oCmdClient.colSousCommandes
                    If oSCMD.oFournisseur.bExportInternet = vncTypeExportScmd.vncExportQuadra Then
                        'Les sousCommandes 'HOBOVIN' sont générée
                        Assert.AreEqual(vncEtatCommande.vncSCMDGeneree, oSCMD.EtatCode)
                        Assert.IsFalse(oSCMD.bExportQuadra)
                    End If
                Next
            End If
            If (TypeOf (oCmd) Is SousCommande) Then
                oSCMD = oCmd
                Assert.IsFalse(oSCMD.bExportQuadra)
            End If

        Next
        '======= VALIDATION ===============
        oExport.ValiderExportBaf()
        'Test de la validation

        'Vérification des etats après la validation
        For Each oCmd As Commande In oExport.ListCmd
            If (TypeOf (oCmd) Is CommandeClient) Then
                'La Commande est Transmise 
                Assert.AreEqual(vncEtatCommande.vncTransmiseQuadra, oCmd.EtatCode)
                oCmdClient = oCmd
                For Each oSCMD In oCmdClient.colSousCommandes
                    If oSCMD.oFournisseur.bExportInternet = vncTypeExportScmd.vncExportQuadra Then
                        'Les SousCommandes 'HOBIVIN' sont Facturée
                        Assert.AreEqual(vncEtatCommande.vncSCMDFacturee, oSCMD.EtatCode)
                        Assert.AreEqual("QUADRAFACT", oSCMD.refFactFournisseur)
                        Assert.IsTrue(oSCMD.bExportQuadra)
                    End If
                Next

            End If
            If (TypeOf (oCmd) Is SousCommande) Then
                oSCMD = oCmd
                Assert.IsTrue(oSCMD.bExportQuadra)

            End If

        Next

        'Vérification du Filtre sur LoadListCmd
        oExport.LoadListCmd()
        Assert.AreEqual(0, oExport.ListCmd.Count)


        objCMDV.bDeleted = True
        objCMDV.save()

        oInter.bDeleted = True
        oInter.save()
    End Sub

    <TestMethod()>
    Public Sub TestValidationExportBaFournisseur()

        Dim objCMDV As CommandeClient
        Dim objCMDH As CommandeClient
        Dim oSCMD As SousCommande
        Dim oLg As LgCommande

        'Founisseur Vinicom
        m_oFourn.bExportInternet = vncTypeExportScmd.vncExportInternet 'Fournisseur Vinicom
        m_oFourn.Save()
        m_oProduit.idFournisseur = m_oFourn.id
        m_oProduit.save()

        'Fournisseur HOBIVIN
        m_oFourn2.bExportInternet = vncTypeExportScmd.vncExportQuadra 'Fournisseur HOBIVIN
        m_oFourn2.Save()
        m_oProduit2.idFournisseur = m_oFourn2.id
        m_oProduit2.save()

        'Client intermédiaire
        Dim oInter As New Client("CLTINTER", Dossier.HOBIVIN)
        oInter.idTypeClient = Param.getTypeClientIntermediaire().id
        oInter.save()


        'Creation d'une Commande "VINICOM" avec un produit Vinicom et un produit HOBIVIN
        objCMDV = New CommandeClient(m_oClient)
        objCMDV.dateCommande = New Date(1964, 2, 6)
        objCMDV.Origine = Dossier.VINICOM
        oLg = objCMDV.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit Vinicom
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMDV.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit HOBIVIN
        oLg.qteLiv = oLg.qteCommande
        objCMDV.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMDV.changeEtat(vncActionEtatCommande.vncActionLivrer)
        objCMDV.save()
        'Commande Vinicom Donc pas d'intermédiaire
        Assert.IsTrue(objCMDV.generationSousCommande())
        objCMDV.save()

        Assert.AreEqual(2, objCMDV.colSousCommandes.Count)  '2 sous Commandes générées

        'Creation d'une Commande "HOBIVIN" avec un produit Vinicom et un produit HOBIVIN
        objCMDH = New CommandeClient(m_oClient)
        objCMDH.dateCommande = New Date(1964, 2, 7)
        objCMDH.Origine = Dossier.HOBIVIN
        oLg = objCMDH.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit Vinicom
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMDH.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit HOBIVIN
        oLg.qteLiv = oLg.qteCommande
        objCMDH.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMDH.changeEtat(vncActionEtatCommande.vncActionLivrer)
        objCMDH.save()
        'Commande HOBIVIN Donc un intermédiaire
        Assert.IsTrue(objCMDH.generationSousCommande())
        objCMDH.save()
        Assert.AreEqual(2, objCMDH.colSousCommandes.Count)  '2 sous Commandes générées

        Dim oExport As New ExportQuadra(New Date(1964, 2, 1), New Date(1964, 3, 1), vncTypeExportQuadra.vncExportBafClient, ".", True)

        'Export de type Baf Client
        oExport.typeExport = vncTypeExportQuadra.vncExportBaFournisseur
        oExport.LoadListCmd()
        oExport.ExportBaf()
        'Test de la validation

        For Each oCmd As Commande In oExport.ListCmd
            If (TypeOf (oCmd) Is CommandeClient) Then
                Assert.AreEqual(vncEtatCommande.vncTransmiseQuadra, oCmd.EtatCode)
            End If
            If (TypeOf (oCmd) Is SousCommande) Then
                oSCMD = oCmd
                Assert.IsFalse(oSCMD.bExportQuadra)

            End If

        Next

        oExport.ValiderExportBaf()
        For Each oCmd As Commande In oExport.ListCmd
            If (TypeOf (oCmd) Is CommandeClient) Then
                Assert.AreEqual(vncEtatCommande.vncTransmiseQuadra, oCmd.EtatCode)
            End If
            If (TypeOf (oCmd) Is SousCommande) Then
                oSCMD = oCmd
                Assert.IsTrue(oSCMD.bExportQuadra)

            End If

        Next


        'Vérification du Filtre sur LoadListCmd
        oExport.LoadListCmd()
        Assert.AreEqual(0, oExport.ListCmd.Count)


        objCMDV.bDeleted = True
        objCMDV.save()

        oInter.bDeleted = True
        oInter.save()
    End Sub

    ''' <summary>
    ''' Test que les sousCommandes sont bien Ordonnées dans la liste
    ''' </summary>
    <TestMethod()>
    Public Sub TestOrdreDesSousCommandesHBV()
        Dim objCMD1 As CommandeClient
        Dim objCMD2 As CommandeClient
        Dim oLg As LgCommande

        'Founisseur Hobivin
        m_oFourn.bExportInternet = vncTypeExportScmd.vncExportQuadra   'Fournisseur Hobivin
        m_oFourn.Save()
        m_oProduit.idFournisseur = m_oFourn.id
        m_oProduit.save()

        'Fournisseur HOBIVIN
        m_oFourn2.bExportInternet = vncTypeExportScmd.vncExportQuadra 'Fournisseur HOBIVIN
        m_oFourn2.Save()
        m_oProduit2.idFournisseur = m_oFourn2.id
        m_oProduit2.save()

        'Client intermédiaire
        Dim oInter As New Client("CLTINTER", Dossier.HOBIVIN)
        oInter.idTypeClient = Param.getTypeClientIntermediaire().id
        oInter.save()


        'Creation d'une Commande "HOBIVIN" avec Deux un produit de fournisseurs différents
        objCMD1 = New CommandeClient(m_oClient)
        objCMD1.dateCommande = New Date(1964, 2, 6)
        objCMD1.Origine = Dossier.VINICOM
        oLg = objCMD1.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit1 
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMD1.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit2
        oLg.qteLiv = oLg.qteCommande
        objCMD1.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMD1.changeEtat(vncActionEtatCommande.vncActionLivrer)
        objCMD1.save()

        Assert.IsTrue(objCMD1.generationSousCommande())
        objCMD1.save()
        Assert.AreEqual(2, objCMD1.colSousCommandes.Count)


        'Creation d'une Commande "HOBIVIN" avec Deux un produit de fournisseurs différents
        objCMD2 = New CommandeClient(m_oClient)
        objCMD2.dateCommande = New Date(1964, 2, 7)
        objCMD2.Origine = Dossier.VINICOM
        oLg = objCMD2.AjouteLigne("10", m_oProduit2, 15, 15.5) 'Produit2
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMD2.AjouteLigne("20", m_oProduit, 25, 25.5) 'Produit1
        oLg.qteLiv = oLg.qteCommande
        objCMD2.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMD2.changeEtat(vncActionEtatCommande.vncActionLivrer)
        objCMD2.save()
        'Commande HOBIVIN Donc un intermédiaire
        Assert.IsTrue(objCMD2.generationSousCommande())
        objCMD2.save()
        Assert.AreEqual(2, objCMD2.colSousCommandes.Count)

        Dim oExportQuadra = New ExportQuadra(CDate("6/2/1964"), CDate("7/2/1964"), vncTypeExportQuadra.vncExportBafClient, "")
        oExportQuadra.LoadListCmd()
        Assert.AreEqual(4, oExportQuadra.ListCmd.Count)
        Assert.AreEqual(objCMD1.code, oExportQuadra.ListCmd(0).getCodeCommande()) 'Sous Commande de la Première Commande
        Assert.AreEqual(objCMD1.code, oExportQuadra.ListCmd(1).getCodeCommande()) 'Sous Commande de la Première Commande
        Assert.AreEqual(objCMD2.code, oExportQuadra.ListCmd(2).getCodeCommande()) 'Sous Commande de la Deuxième Commande
        Assert.AreEqual(objCMD2.code, oExportQuadra.ListCmd(3).getCodeCommande()) 'Sous Commande de la Dexième Commande



    End Sub
    ''' <summary>
    ''' Test de l'export d'une SCMD vinicom ave cune reference de plus de 12 car
    ''' </summary>
    <TestMethod()> Public Sub T_ReferenceSouscommandesup12carac()
        Dim objCMD As CommandeClient
        Dim oSCMD As SousCommande
        Dim oLg As LgCommande
        Dim strFile As String

        m_oFourn.bExportInternet = vncTypeExportScmd.vncExportInternet 'Fournisseur Vinicom
        m_oFourn.Save()
        m_oProduit.idFournisseur = m_oFourn.id
        m_oProduit.save()

        m_oFourn2.bExportInternet = vncTypeExportScmd.vncExportQuadra 'Fournisseur HOBIVIN
        m_oFourn2.Save()
        m_oProduit2.idFournisseur = m_oFourn2.id
        m_oProduit2.save()

        'Creation d'une Commande "VINICOM" avec un produit Vinicom et un produit HOBIVIN
        objCMD = New CommandeClient(m_oClient)
        objCMD.Origine = Dossier.VINICOM
        oLg = objCMD.AjouteLigne("10", m_oProduit, 15, 15.5) 'Produit Vinicom
        oLg.qteLiv = oLg.qteCommande
        oLg = objCMD.AjouteLigne("20", m_oProduit2, 25, 25.5) 'Produit HOBIVIN
        oLg.qteLiv = oLg.qteCommande
        objCMD.changeEtat(vncActionEtatCommande.vncActionValider)
        objCMD.changeEtat(vncActionEtatCommande.vncActionLivrer)

        objCMD.save()

        'Commande Vinicom Donc pas d'intermédiaire
        Assert.IsTrue(objCMD.generationSousCommande())
        Assert.AreEqual(2, objCMD.colSousCommandes.Count)  '2 sous Commandes générées


        'Export de la Première SousCommande (CMD VINICOM, PROD VINICOM)
        'Normalement cette Souscommande n'est pas Exportée sous Quadra
        'même traitement qu'une commande "HOBIVIN" avec un produit VINICOM
        'Export de type BAFFournisseur (Mise à jour des stocks dans QuadraFact)
        m_oCmd.Origine = Dossier.HOBIVIN

        oSCMD = objCMD.colSousCommandes(1)
        'Modification du code pour lui affecter plus de 13 carac
        oSCMD.FTO_setCode(oSCMD.code + "99")

        strFile = "./adel.csv"
        If System.IO.File.Exists(strFile) Then
            System.IO.File.Delete(strFile)
        End If
        oSCMD.toCSVQuadraFact(strFile, vncTypeExportQuadra.vncExportBaFournisseur)

        Assert.IsTrue(System.IO.File.Exists(strFile))
        'Il y a Bien 2 lignes dans le fichier( 1 entête et une ligne)
        Assert.AreEqual(2, System.IO.File.ReadAllLines(strFile).Count)
        Debug.WriteLine(File.ReadAllText(strFile))
        Dim strLine As String
        Dim tab As String()
        strLine = File.ReadAllLines(strFile)(1)
        tab = strLine.Split(";")
        Assert.AreEqual(tab(0), oSCMD.oFournisseur.code) 'Cette fois c'est le code fournisseur
        Assert.AreEqual(tab(1), oSCMD.code)
        Assert.AreEqual(tab(2), Format(oSCMD.dateCommande, "yyMMdd"))
        Assert.AreEqual(tab(3), Right(oSCMD.code.Replace("-", ""), 12))
        Assert.AreEqual(tab(4), m_oProduit.code)
        Assert.AreEqual(tab(5), "15")
        Assert.AreEqual(tab(6), "15.5")

        objCMD.delete()
    End Sub


End Class