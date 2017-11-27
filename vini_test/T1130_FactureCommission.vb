Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.IO
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine



<TestClass()> Public Class T1130_FactCommission
    Inherits test_Base

    Private m_oProduit As Produit
    Private m_oProduit2 As Produit
    Private m_oFourn As Fournisseur
    Private m_oFourn2 As Fournisseur
    Private m_oClient As Client
    Dim objProduit1 As Produit
    Dim objProduit2 As Produit
    Dim objProduit3 As Produit
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        m_oFourn = New Fournisseur("ZZ1", "MonFournisseur")
        m_oFourn.CodeCompta = "411001"
        m_oFourn.rs = "Raison Sociale"
        m_oFourn.Save()

        m_oFourn2 = New Fournisseur("ZZ2", "MonFournisseur")
        m_oFourn2.CodeCompta = "411002"
        m_oFourn2.rs = "Raison Sociale2"
        m_oFourn2.Save()

        m_oProduit = New Produit("ZZ1001", m_oFourn, 1990)
        m_oProduit.save()

        m_oProduit2 = New Produit("ZZ2001", m_oFourn2, 1990)
        m_oProduit2.save()

        m_oClient = New Client("XX1", "MonClient")
        m_oClient.rs = "Raison Sociale"
        Debug.Assert(m_oClient.save(), "Creation du client")
        '            m_oClient()
        objProduit1 = New Produit("ZZ1002", m_oFourn, 1994)
        Assert.IsTrue(objProduit1.save())
        objProduit2 = New Produit("ZZ1003", m_oFourn, 1994)
        Assert.IsTrue(objProduit2.save())
        objProduit3 = New Produit("ZZ1004", m_oFourn, 1994)
        Assert.IsTrue(objProduit3.save())



    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objFACT As FactCom
        Dim objFACT2 As FactCom

        objFACT = New FactCom(m_oFourn)

        Assert.AreEqual("", objFACT.periode, "Periode de reference")
        Assert.AreEqual(0D, objFACT.montantReglement, "Montant du reglement")
        Assert.AreEqual(CDate("01/01/2000"), objFACT.dateReglement, "date de reglement")
        Assert.AreEqual("", objFACT.refReglement, "reference du reglement")


        objFACT.dateStatistique = "01/03/2004"
        objFACT.periode = "Periode de reference"
        objFACT.montantReglement = 155.8
        objFACT.dateReglement = "31/07/64"
        objFACT.refReglement = "CA01"

        Assert.AreEqual(objFACT.periode, "Periode de reference")
        Assert.AreEqual(objFACT.montantReglement, 155.8D)
        Assert.AreEqual(objFACT.dateReglement, CDate("31/07/1964"))
        Assert.AreEqual(objFACT.dateStatistique, CDate("01/03/2004"))
        Assert.AreEqual(objFACT.refReglement, "CA01")


        'Test des indicateurs
        Assert.IsTrue(objFACT.bNew)
        Assert.IsTrue(objFACT.bUpdated)
        Assert.IsFalse(objFACT.bDeleted)

        Assert.IsTrue(objFACT.Equals(objFACT), "Egal à Lui même")
        objFACT2 = New FactCom(m_oFourn)


        objFACT2.periode = "Periode de reference"
        objFACT2.montantReglement = 155.8
        objFACT2.dateReglement = "31/07/64"
        objFACT2.refReglement = "CA01"
        objFACT2.dateStatistique = "01/03/2004"

        Assert.IsTrue(objFACT.Equals(objFACT2), "Egal à un semblable")

        objFACT2.dateStatistique = Now()
        Assert.IsFalse(objFACT.Equals(objFACT2), "Egal à un Différent")

        Dim obj As Object
        Assert.IsFalse(objFACT.Equals(obj), "Egal autrecjhose")


    End Sub
    <TestMethod()> Public Sub T15_DB()
        Dim objFACT As FactCom
        Dim objFACT2 As FactCom
        Dim nid As Long

        'I - Création d'une Facture 
        '=========================
        objFACT = New FactCom(m_oFourn)

        objFACT.dateCommande = CDate("06/02/1964")
        objFACT.periode = "1er Timestre 1964"
        objFACT.dateReglement = CDate("31/07/1964")
        objFACT.montantReglement = 321.34
        objFACT.refReglement = "CMB0034"
        objFACT.dateStatistique = "01/03/1964"
        objFACT.dEcheance = "01/04/1964"
        objFACT.idModeReglement = 1



        'Test des indicateurs Avant le Save
        Assert.IsTrue(objFACT.bNew)
        Assert.IsTrue(objFACT.bUpdated)
        Assert.IsFalse(objFACT.bDeleted)
        'Save
        Assert.IsTrue(objFACT.Save(), "Insert" & objFACT.getErreur)
        Assert.IsTrue((objFACT.id <> 0), "Id Apres le Save doit être différent de 0")
        'Test des indicateurs Après le Save
        Assert.IsFalse(objFACT.bNew, "bNew apres insert")
        Assert.IsFalse(objFACT.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objFACT.bDeleted, "bDeleted apres insert")

        nid = objFACT.id
        'II - Rechargement d'une Facture
        '===============================
        objFACT2 = New FactCom(m_oFourn)
        Assert.IsTrue(objFACT2.load(nid), "Load de la Facture Comm  " & nid & ":" & objFACT2.getErreur())
        Assert.IsTrue(objFACT.Equals(objFACT2), "Différents")

        Assert.AreEqual(CDate("01/04/1964"), objFACT2.dEcheance, "Date Echeance")
        Assert.AreEqual(1, objFACT2.idModeReglement, "Mode de reglement")

        'III - Modification d'une Facture de Comm
        '=================================
        ' Modification de la commande
        objFACT2.dateCommande = CDate("06/02/1984")
        objFACT2.dateStatistique = CDate("01/03/1984")
        objFACT2.periode = "1er Trimestre 1984"
        objFACT2.dateReglement = CDate("31/07/1984")
        objFACT2.montantReglement = 789.9
        objFACT2.refReglement = "CMB0035"
        objFACT2.idModeReglement = 2
        objFACT2.dEcheance = "01/04/1984"


        'Test des indicateurs Avant le Save
        Assert.IsFalse(objFACT2.bNew)
        Assert.IsTrue(objFACT2.bUpdated)
        Assert.IsFalse(objFACT2.bDeleted)
        'Save
        Assert.IsTrue(objFACT2.Save(), "Update" & objFACT2.getErreur)
        'Test des indicateurs Après le Save
        Assert.IsFalse(objFACT2.bNew, "bNew apres Update")
        Assert.IsFalse(objFACT2.bUpdated, "bUpdated apres Update")
        Assert.IsFalse(objFACT2.bDeleted, "bDeleted apres Update")
        'Rechargement de l'objet
        nid = objFACT2.id
        objFACT = New FactCom(m_oFourn)
        Assert.IsTrue(objFACT.load(nid), "Load")
        Assert.IsTrue(objFACT.Equals(objFACT2), "Apres Update , Equals")
        Assert.AreEqual(CDate("01/04/1984"), objFACT.dEcheance, "Date Echeance")
        Assert.AreEqual(2, objFACT.idModeReglement, "Mode de reglement")



        'IV - Suppression de la Facture
        '=================================
        ' Modification de la Facture
        objFACT.bDeleted = True
        Assert.IsTrue(objFACT.Save(), "Delete" & objFACT.getErreur())
        'Rechargement dans un autre objet
        objFACT2 = New FactCom(m_oFourn)
        objFACT2.load(nid)
        Assert.IsTrue(objFACT2.id = 0)
        'Assert.IsFalse(objFACT2.load(nid), "Load")
    End Sub
    <TestMethod()> Public Sub T16_DB()

        Dim objCMD As CommandeClient
        Dim nid1 As Integer
        Dim nid2 As Integer
        Dim oFrn2 As Fournisseur
        Dim nidFRN2 As Integer
        Dim oPRD2 As Produit
        Dim nidPRD2 As Integer
        Dim objFact1 As FactCom
        Dim objFact2 As FactCom
        Dim objSCMD As SousCommande
        Dim objLgCom As LgCommande
        Dim nidSCMD As Integer
        Dim col1 As ColEvent
        Dim col2 As ColEvent

        oFrn2 = New Fournisseur("FRN2T16" & CStr(Now()), "FRN2T16")
        Assert.IsTrue(oFrn2.Save())
        nidFRN2 = oFrn2.id

        oPRD2 = New Produit("PRD2T16" & CStr(Now()), oFrn2, 1990)
        Assert.IsTrue(oPRD2.save())
        nidPRD2 = oPRD2.id

        'CREATION D'UNE COMMANDE CLIENT
        objCMD = New CommandeClient(m_oClient)
        'Ajout de 2 lignes de Commandes
        objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 12, 20, True, 120, (120 * 1.196))
        objCMD.AjouteLigne(objCMD.getNextNumLg, oPRD2, 12, 20, True, 125, (125 * 1.196))
        ''Sauvegarde de la commande
        objCMD.totalTTC = 12
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
        'Génération des Sous-Commandes
        objCMD.generationSousCommande()
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())

        objFact1 = New FactCom(m_oFourn)
        objFact1.periode = "Test T16_DB_FRN"

        For Each objSCMD In objCMD.colSousCommandes
            If objSCMD.oFournisseur.id = m_oFourn.id Then
                objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
                objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
                Assert.IsTrue(objSCMD.Save(), "Sauvegarde de la sous Commande")
                objFact1.AjouteLigneFactCom(objSCMD)
            End If
        Next objSCMD

        objFact2 = New FactCom(oFrn2)
        objFact2.periode = "Test T16_DB_FRN2"

        For Each objSCMD In objCMD.colSousCommandes
            If objSCMD.oFournisseur.id = oFrn2.id Then
                objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
                objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
                Assert.IsTrue(objSCMD.Save(), "Sauvegarde de la sous Commande")
                objFact2.AjouteLigneFactCom(objSCMD)
            End If
        Next objSCMD

        'Sauvegarde des factures de commission
        Assert.IsTrue(objFact1.Save(), "Sauvegarde de la facture de comm")
        nid1 = objFact1.id
        Assert.IsTrue(objFact2.Save(), "Sauvegarde de la facture de comm")
        nid2 = objFact2.id

        objFact1 = New FactCom(m_oFourn)
        objFact1.load(nid1)
        Assert.IsTrue(objFact1.colLignes.Count = 1, "1 ligne de FActures de Comm")
        'Check de chaque ligne de Facture 
        For Each objLgCom In objFact1.colLignes
            nidSCMD = objLgCom.idSCmd
            objSCMD = SousCommande.createandload(nidSCMD)
            Assert.IsTrue(objSCMD.id = nidSCMD, "Load de la sousCommande " & nidSCMD)
            'La sous commande est bien rattachée à la facture
            Assert.AreEqual(objFact1.id, objSCMD.idFactCom, "IDFactcom")
            'La Sous Commande est bien facturée
            Assert.IsTrue(objSCMD.etat.codeEtat = vncEnums.vncEtatCommande.vncSCMDFacturee)
        Next objLgCom

        objFact2 = New FactCom(m_oFourn)
        objFact2.load(nid2)
        Assert.IsTrue(objFact2.colLignes.Count = 1, "1 ligne de FActures de Comm")
        'Check de chaque ligne de Facture 
        For Each objLgCom In objFact2.colLignes
            nidSCMD = objLgCom.idSCmd
            objSCMD = SousCommande.createandload(nidSCMD)
            Assert.IsTrue(objSCMD.id = nidSCMD, "Load de la sousCommande " & nidSCMD)
            'La sous commande est bien rattachée à la facture
            Assert.AreEqual(objFact2.id, objSCMD.idFactCom, "IDFactcom")
            'La Sous Commande est bien facturée
            Assert.IsTrue(objSCMD.etat.codeEtat = vncEnums.vncEtatCommande.vncSCMDFacturee)
        Next objLgCom

        col1 = objFact1.colLignes
        'Suppression de la facture de comm
        objFact1.bDeleted = True
        Assert.IsTrue(objFact1.Save(), "Suppression de la facture de commission")
        'Vérification que les sous Commandes soient bien revenues à l'état transmise
        For Each objLgCom In col1
            nidSCMD = objLgCom.idSCmd
            objSCMD = SousCommande.createandload(nidSCMD)
            Assert.IsTrue(objSCMD.id = nidSCMD, "Load de la sousCommande " & nidSCMD)
            'La sous commande n'est pls rattachée à la facture
            Assert.AreEqual(-1, objSCMD.idFactCom, "IDFactcom")
            'La Sous Commande n'est plus facturée
            Assert.IsTrue(objSCMD.etat.codeEtat <> vncEnums.vncEtatCommande.vncSCMDFacturee)
        Next objLgCom

        col2 = objFact2.colLignes
        'Suppression de la facture de comm
        objFact2.bDeleted = True
        Assert.IsTrue(objFact2.Save(), "Suppression de la facture de commission")
        'Vérification que les sous Commandes soient bien revenues à l'état transmise
        For Each objLgCom In col2
            nidSCMD = objLgCom.idSCmd
            objSCMD = SousCommande.createandload(nidSCMD)
            Assert.IsTrue(objSCMD.id = nidSCMD, "Load de la sousCommande " & nidSCMD)
            'La sous commande n'est pls rattachée à la facture
            Assert.AreEqual(-1, objSCMD.idFactCom, "IDFactcom")
            'La Sous Commande n'est plus facturée
            Assert.IsTrue(objSCMD.etat.codeEtat <> vncEnums.vncEtatCommande.vncSCMDFacturee)
        Next objLgCom


        'Suppression des elements créés
        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save(), "Suppression de la commande")

        oPRD2.load(nidPRD2)
        oPRD2.bDeleted = True
        Assert.IsTrue(oPRD2.save(), "Supression du Produit")
        oFrn2.load(nidFRN2)
        oFrn2.bDeleted = True
        Assert.IsTrue(oFrn2.Save(), "Suppression du Fournisseur")

    End Sub

    ''' <summary>
    ''' Test l'export d'une facture de comm
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T20_EXPORT()
        Dim objFACT As FactCom
        Dim strLines As String()
        Dim strLine As String
        Dim strLine1 As String

        'I - Création d'une Facture 
        '=========================
        m_oFourn.banque = "CMB LIFFRE"
        m_oFourn.rib1 = "12"
        m_oFourn.rib2 = "34"
        m_oFourn.rib3 = "56"
        m_oFourn.rib4 = "78"


        Dim oParam As ParamModeReglement
        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "CHQ30NETS"
        oParam.dDebutEcheance = "FACT"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())


        Assert.IsTrue(m_oFourn.Save())
        objFACT = New FactCom(m_oFourn)

        objFACT.dateCommande = CDate("06/02/1964")
        objFACT.periode = "1er Timestre 1964"
        objFACT.dateFacture = CDate("06/02/1964")
        objFACT.dateReglement = CDate("31/07/1964")
        objFACT.montantReglement = 321.34
        objFACT.refReglement = "CMB0034"
        objFACT.dateStatistique = "01/03/1964"
        objFACT.totalHT = 150.56
        objFACT.totalTTC = 180.89
        objFACT.dEcheance = "01/04/1964"
        objFACT.idModeReglement = oParam.id

        'Save
        Assert.IsTrue(objFACT.Save(), "Insert" & objFACT.getErreur)
        If File.Exists("./T20_EXPORT.txt") Then
            File.Delete("./T20_EXPORT.txt")
        End If

        objFACT.Exporter("./T20_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T20_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T20_EXPORT.txt")
        Assert.AreEqual(6, strLines.Length, "6 lignes d'export")


        strLine = strLines(0)
        Assert.AreEqual(231, strLine.Length)
        Assert.AreEqual("M", strLine.Substring(0, 1))
        Assert.AreEqual(m_oFourn.CodeCompta, Trim(strLine.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine.Substring(9, 2)))
        Assert.AreEqual("", Trim(strLine.Substring(11, 3)))
        Assert.AreEqual("060264", strLine.Substring(14, 6))
        Assert.AreEqual("V", strLine.Substring(20, 1))
        Assert.AreEqual("D", strLine.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine.Substring(42, 13)))
        Assert.AreEqual("060364", Trim(strLine.Substring(63, 6)))

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

        'Test de la 5eme ligne
        strLine = strLines(4)
        Assert.AreEqual(231, strLine.Length)
        Assert.AreEqual("M", strLine.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC_COMPTETVA")), Trim(strLine.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine.Substring(9, 2)))
        Assert.AreEqual("060264", strLine.Substring(14, 6))
        Assert.AreEqual(("F:" + objFACT.code + " " + m_oFourn.rs + Space(20)).Substring(0, 20), Trim(strLine.Substring(21, 20)))
        Assert.AreEqual("C", strLine.Substring(41, 1))
        Assert.AreEqual((180.89 - 150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine.Substring(42, 13)))

        'Test de la 6eme Ligne
        strLine = strLines(5)
        Assert.AreEqual(231, strLine.Length)
        Assert.AreEqual("M", strLine.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC_COMPTEPRODUIT")), Trim(strLine.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine.Substring(9, 2)))
        Assert.AreEqual("060264", strLine.Substring(14, 6))
        Assert.AreEqual(("F:" + objFACT.code + " " + m_oFourn.rs + Space(20)).Substring(0, 20), Trim(strLine.Substring(21, 20)))
        Assert.AreEqual("C", strLine.Substring(41, 1))
        Assert.AreEqual((150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine.Substring(42, 13)))

 
        objFACT.bDeleted = True
        objFACT.Save()
    End Sub
    ''' <summary>
    ''' Test l'export d'un AVOIR de comm
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T21_EXPORTAVOIR()
        Dim objFACT As FactCom
        Dim strLines As String()
        Dim strLine As String

        'I - Création d'une Facture 
        '=========================
        Dim oParam As ParamModeReglement
        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "CHQ30NETS"
        oParam.dDebutEcheance = "FACT"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())

        m_oFourn.banque = "CMB LIFFRE"
        m_oFourn.rib1 = "12"
        m_oFourn.rib2 = "34"
        m_oFourn.rib3 = "56"
        m_oFourn.rib4 = "78"

        Assert.IsTrue(m_oFourn.Save())
        objFACT = New FactCom(m_oFourn)

        objFACT.dateCommande = CDate("06/02/1964")
        objFACT.periode = "1er Timestre 1964"
        objFACT.dateFacture = CDate("06/02/1964")
        objFACT.dateReglement = CDate("31/07/1964")
        objFACT.montantReglement = 321.34
        objFACT.refReglement = "CMB0034"
        objFACT.dateStatistique = "01/03/1964"
        objFACT.totalHT = -150.56
        objFACT.totalTTC = -180.89
        objFACT.dEcheance = "01/04/1964"
        objFACT.idModeReglement = oParam.id

        'Save
        Assert.IsTrue(objFACT.Save(), "Insert" & objFACT.getErreur)
        If File.Exists("./T20_EXPORT.txt") Then
            File.Delete("./T20_EXPORT.txt")
        End If

        objFACT.Exporter("./T20_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T20_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T20_EXPORT.txt")
        Assert.AreEqual(6, strLines.Length, "4 lignes d'export")

        '1etre ligne
        strLine = strLines(0)
        Assert.AreEqual(231, strLine.Length)
        Assert.AreEqual("M", strLine.Substring(0, 1))
        Assert.AreEqual(m_oFourn.CodeCompta, Trim(strLine.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine.Substring(9, 2)))
        Assert.AreEqual("", Trim(strLine.Substring(11, 3)))
        Assert.AreEqual("060264", strLine.Substring(14, 6))
        Assert.AreEqual("V", strLine.Substring(20, 1))
        Assert.AreEqual("C", strLine.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine.Substring(42, 13)))
        Assert.AreEqual("060364", Trim(strLine.Substring(63, 6)))

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

        '5eme Ligne
        strLine = strLines(4)
        Assert.AreEqual(231, strLine.Length)
        Assert.AreEqual("M", strLine.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC_COMPTETVA")), Trim(strLine.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine.Substring(9, 2)))
        Assert.AreEqual("060264", strLine.Substring(14, 6))
        Assert.AreEqual(("A:" + objFACT.code + " " + m_oFourn.rs + Space(20)).Substring(0, 20), Trim(strLine.Substring(21, 20)))
        Assert.AreEqual("D", strLine.Substring(41, 1))
        Assert.AreEqual((180.89 - 150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine.Substring(42, 13)))

        '6eme Ligne
        strLine = strLines(5)
        Assert.AreEqual(231, strLine.Length)
        Assert.AreEqual("M", strLine.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC_COMPTEPRODUIT")), Trim(strLine.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine.Substring(9, 2)))
        Assert.AreEqual("060264", strLine.Substring(14, 6))
        Assert.AreEqual(("A:" + objFACT.code + " " + m_oFourn.rs + Space(20)).Substring(0, 20), Trim(strLine.Substring(21, 20)))
        Assert.AreEqual("D", strLine.Substring(41, 1))
        Assert.AreEqual((150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine.Substring(42, 13)))


        objFACT.bDeleted = True
        objFACT.Save()
    End Sub

    ''' <summary>
    ''' Test l'export d'une facture de comm
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T22_EXPORTModeReglement()
        Dim objFACT As FactCom
        Dim strLines As String()
        Dim strLine As String

        'I - Création d'une Facture 
        '=========================
        m_oFourn.banque = "CMB LIFFRE"
        m_oFourn.rib1 = "12"
        m_oFourn.rib2 = "34"
        m_oFourn.rib3 = "56"
        m_oFourn.rib4 = "78"


        Dim oParam As ParamModeReglement
        'Création d'un mode de reglement Prmevement 60 FDM
        oParam = New ParamModeReglement()
        oParam.code = "PRE60J"
        oParam.dDebutEcheance = "FACT"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())


        Assert.IsTrue(m_oFourn.Save())
        objFACT = New FactCom(m_oFourn)

        objFACT.dateCommande = CDate("06/02/1964")
        objFACT.periode = "1er Timestre 1964"
        objFACT.dateFacture = CDate("06/02/1964")
        objFACT.dateReglement = CDate("31/07/1964")
        objFACT.montantReglement = 321.34
        objFACT.refReglement = "CMB0034"
        objFACT.dateStatistique = "01/03/1964"
        objFACT.totalHT = 150.56
        objFACT.totalTTC = 180.89
        objFACT.dEcheance = "01/04/1964"
        objFACT.idModeReglement = oParam.id

        'Save
        Assert.IsTrue(objFACT.Save(), "Insert" & objFACT.getErreur)
        If File.Exists("./T21_EXPORT.txt") Then
            File.Delete("./T21_EXPORT.txt")
        End If

        objFACT.Exporter("./T21_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T21_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T21_EXPORT.txt")

        'Test de la 2eme ligne
        strLine = strLines(1)

        Dim unChamp As String()
        Dim ChampVal As String() = strLine.Split(";")
        Assert.AreEqual("Y", ChampVal(0))
        Assert.AreEqual(2, ChampVal.Length)
        unChamp = ChampVal(1).Split("=")
        Assert.AreEqual("ModePaiement", unChamp(0))
        Assert.AreEqual("PRE", unChamp(1))



        'Test de la 4eme ligne
        strLine = strLines(3)
        Assert.AreEqual(18, strLine.Length)
        Assert.AreEqual("Z;ModePaiement=", strLine.Substring(0, 15))
        Assert.AreEqual("PRE", strLine.Substring(15, 3))


        objFACT.bDeleted = True
        objFACT.Save()
    End Sub

    <TestMethod()> Public Sub T30_ChampsLongs()

        Dim objFACT As FactCom

        'I - Création d'une Facture 
        '=========================
        objFACT = New FactCom(m_oFourn)

        objFACT.periode = "1er Timestre 1964".PadRight(500, "x")
        'Save
        Assert.IsTrue(objFACT.Save(), "Insert" & objFACT.getErreur)

        objFACT.load()

        Assert.AreEqual(50, objFACT.periode.Length)
        objFACT.bDeleted = True
        Assert.IsTrue(objFACT.Save())

    End Sub

    'Test l'incrémenation des codes
    <TestMethod()> Public Sub T70_GetNextCode()

        Dim obj1 As New FactCom(m_oFourn)
        Dim obj2 As New FactCom(m_oFourn)

        obj1.Save()
        obj2.Save()
        Assert.AreEqual(obj2.code, (obj1.code + 1).ToString())
        Assert.AreNotEqual(0, obj1.code)

        obj1.bDeleted = True
        obj1.Save()
        obj2.bDeleted = True
        obj2.Save()


    End Sub
    'Test l'indicateur d'export internet
    <TestMethod()> Public Sub T80_ExportInternet()

        Dim obj1 As New FactCom(m_oFourn)

        'Valeur par defaut
        Assert.AreEqual(False, obj1.bExportInternet)

        'Modif de la valeur False=> true
        obj1.bExportInternet = True
        Assert.AreEqual(True, obj1.bExportInternet)
        Assert.IsTrue(obj1.Save)
        obj1 = FactCom.createandload(obj1.id)
        Assert.AreEqual(True, obj1.bExportInternet)

        'Modif de la valeur True=> False
        obj1.bExportInternet = False
        Assert.IsTrue(obj1.Save)
        obj1 = FactCom.createandload(obj1.id)
        Assert.AreEqual(False, obj1.bExportInternet)


        obj1.bDeleted = True
        Assert.IsTrue(obj1.Save)

    End Sub

    ''' <summary>
    ''' Test la génération de facture de comm, Cette génération se fait sur la date de commande
    ''' </summary>
    ''' <remarks></remarks>


    <TestMethod()> Public Sub T90_GenerationFactureCommissions()
        Dim objCMD As CommandeClient
        Dim objSCMD As SousCommande
        Dim colFactCom As ColEvent
        Dim colSCMD As List(Of SousCommande)
        Dim oFactcom As FactCom
        Dim nIDCmd As String
        Dim oParam As ParamModeReglement


        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "TEST30FDM"
        oParam.dDebutEcheance = "FDM"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())
        m_oFourn.idModeReglement1 = oParam.id
        Assert.IsTrue(m_oFourn.Save())

        'Création d'un mode de reglement comptant
        oParam = New ParamModeReglement()
        oParam.dDebutEcheance = "FACT"
        oParam.code = "TESTCOMPTANT"
        oParam.valeur2 = 0
        Assert.IsTrue(oParam.Save())
        m_oFourn2.idModeReglement1 = oParam.id
        Assert.IsTrue(m_oFourn2.Save())

        'CREATION D'UNE COMMANDE CLIENT
        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = CDate("06/02/1964")
        'Ajout de 2 lignes de Commandes
        objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 12, 20, True, 120, (120 * 1.196))
        objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit2, 12, 20, True, 125, (125 * 1.196))
        ''Sauvegarde de la commande
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
        nIDCmd = objCMD.id
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
        'Génération des Sous-Commandes
        objCMD.generationSousCommande()
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())

        For Each objSCMD In objCMD.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
            objSCMD.refFactFournisseur = "FACTFRN" + Now()
            objSCMD.totalHTFacture = objSCMD.totalHT
            objSCMD.totalTTCFacture = objSCMD.totalTTC
            objSCMD.dateFactFournisseur = CDate("06/02/1968") ' Date facture fournisseur <> date de commande
            Assert.IsTrue(objSCMD.Save(), "Sauvegarde de la sous Commande")
        Next objSCMD


        colSCMD = SousCommande.getListeAFacturer(CDate("01/02/1968"), CDate("28/02/1968"))
        'MCO : le 23/01/09 : demande de mr Fagnou
        '2 Soucommandes , le tri se fait sur la date de facture et non la date de Commande

        Assert.AreEqual(2, colSCMD.Count)

        colFactCom = FactCom.createFactComs(colSCMD, CDate("28/02/1968"), , "FEVRIER 1968")
        Assert.AreEqual(2, colFactCom.Count)

        For Each oFactcom In colFactCom
            'Une ligne sur chaque Facture
            Assert.AreEqual(1, oFactcom.colLignes.Count)
            If oFactcom.oTiers.id = m_oFourn.id Then
                Assert.AreEqual(m_oFourn.idModeReglement1, oFactcom.idModeReglement)
            End If
            If oFactcom.oTiers.id = m_oFourn2.id Then
                Assert.AreEqual(m_oFourn2.idModeReglement1, oFactcom.idModeReglement)
            End If
            Assert.AreNotEqual(CDate("01/01/2000"), oFactcom.dEcheance)
        Next



    End Sub
    ''' <summary>
    ''' Test l'attribut Selected des sousCommandes dans la génération des Fact de comm
    ''' </summary>
    <TestMethod()> Public Sub T90_GenerationFactureCommissionsSelected()
        Dim objCMD As CommandeClient
        Dim objSCMD As SousCommande
        Dim colFactCom As ColEvent
        Dim colSCMD As List(Of SousCommande)
        Dim oFactcom As FactCom
        Dim nIDCmd As String
        Dim oParam As ParamModeReglement


        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "TEST30FDM"
        oParam.dDebutEcheance = "FDM"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())
        m_oFourn.idModeReglement1 = oParam.id
        Assert.IsTrue(m_oFourn.Save())

        'Création d'un mode de reglement comptant
        oParam = New ParamModeReglement()
        oParam.dDebutEcheance = "FACT"
        oParam.code = "TESTCOMPTANT"
        oParam.valeur2 = 0
        Assert.IsTrue(oParam.Save())
        m_oFourn2.idModeReglement1 = oParam.id
        Assert.IsTrue(m_oFourn2.Save())

        'CREATION D'UNE COMMANDE CLIENT
        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = CDate("06/02/1964")
        'Ajout de 2 lignes de Commandes
        objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 12, 20, True, 120, (120 * 1.196))
        objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit2, 12, 20, True, 125, (125 * 1.196))
        ''Sauvegarde de la commande
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
        nIDCmd = objCMD.id
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
        'Génération des Sous-Commandes
        objCMD.generationSousCommande()
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())

        For Each objSCMD In objCMD.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
            objSCMD.refFactFournisseur = "FACTFRN" + Now()
            objSCMD.totalHTFacture = objSCMD.totalHT
            objSCMD.totalTTCFacture = objSCMD.totalTTC
            objSCMD.dateFactFournisseur = CDate("06/02/1968") ' Date facture fournisseur <> date de commande
            Assert.IsTrue(objSCMD.Save(), "Sauvegarde de la sous Commande")
        Next objSCMD

        '==============================================================================
        colSCMD = SousCommande.getListeAFacturer(CDate("01/02/1968"), CDate("28/02/1968"))
        Assert.AreEqual(2, colSCMD.Count)

        colFactCom = FactCom.createFactComs(colSCMD, CDate("28/02/1968"), , "FEVRIER 1968")
        Assert.AreEqual(2, colFactCom.Count)

        'Déselection de la première sous commande
        colSCMD(0).Selected = False
        colFactCom = FactCom.createFactComs(colSCMD, CDate("28/02/1968"), , "FEVRIER 1968")
        Assert.AreEqual(1, colFactCom.Count)





    End Sub
    ''' <summary>
    ''' test de la fonction GetAtributeValue pour l'export QUADRA
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T100_getAttributeValue()

        Dim oFact As FactCom

        Dim oParam As ParamModeReglement
        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "TA30FDM"
        oParam.dDebutEcheance = "FDM"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())

        oFact = New FactCom(m_oFourn)
        oFact.idModeReglement = oParam.id

        Assert.AreEqual("LC", oFact.getAttributeValue("MODEREGLEMENT2", Nothing))
        Assert.AreEqual("LCR", oFact.getAttributeValue("MODEREGLEMENT4", Nothing))

        oParam = New ParamModeReglement()
        oParam.code = "CHQ30NETS"
        oParam.dDebutEcheance = "FACT"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())

        oFact = New FactCom(m_oFourn)
        oFact.idModeReglement = oParam.id

        Assert.AreEqual("CH", oFact.getAttributeValue("MODEREGLEMENT2", Nothing))
        Assert.AreEqual("CHQ", oFact.getAttributeValue("MODEREGLEMENT4", Nothing))

    End Sub


    <TestMethod()>
    Public Sub TestReleveSurPlusieursPage()

        Dim objCMD As CommandeClient
        Dim lstCmd As List(Of CommandeClient)
        Dim nIdCmd As Integer
        Dim objFrn As Fournisseur
        Dim objPrd As Produit
        Dim oFact As FactCom


        objFrn = New Fournisseur("FRNTEST", "FRN TEST")
        objFrn.bExportInternet = vncTypeExportScmd.vncExportInternet
            objfrn.Save()
        objPrd = New Produit("PRDTEST", objFrn, 0)
        objPrd.save()



        'CREATION D'UNE COMMANDE CLIENT
        'Ajout de 40 lignes de Commandes
        For i As Integer = 0 To 60
            objCMD = New CommandeClient(m_oClient)
            objCMD.dateCommande = CDate("06/02/1964")
            objCMD.AjouteLigne(objCMD.getNextNumLg, objPrd, 12, 20)
            ''Sauvegarde de la commande
            Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
            objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
            Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
            'Génération des Sous-Commandes
            objCMD.generationSousCommande()
            Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
            For Each objSCMD In objCMD.colSousCommandes
                objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
                objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
                objSCMD.refFactFournisseur = "FACTFRN" + Now()
                objSCMD.totalHTFacture = objSCMD.totalHT
                objSCMD.totalTTCFacture = objSCMD.totalTTC
                objSCMD.dateFactFournisseur = CDate("06/02/1968") ' Date facture fournisseur <> date de commande
                Assert.IsTrue(objSCMD.Save(), "Sauvegarde de la sous Commande")

            Next objSCMD
        Next
        Dim lst As New List(Of SousCommande)
        lst = SousCommande.getListeAFacturer(,, objFrn.code)
        'Generation de la facture de commission
        Dim colFact As ColEvent
        colFact = FactCom.createFactComs(lst)
        For Each oFact In colFact
            oFact.Save()
        Next

        Dim strReport As String = ""
        Dim objReport As ReportDocument
        Dim tabIds As ArrayList
        strReport = "V:\V5\vini_app/crFactCom_Releve.rpt"

        If strReport = "" Then
            Exit Sub
        End If
        objReport = New ReportDocument
        objReport.Load(strReport)

        tabIds = New ArrayList()

        oFact = colFact(1)
        oFact.load()
            'oFact.loadcolLignes()
            tabIds.Add(oFact.id)
            objReport.SetParameterValue("IdFactures", tabIds.ToArray())
            objReport.SetParameterValue("bEntete", True)
            If (File.Exists("releveFact2.pdf")) Then
                File.Delete("releveFact2.pdf")
            End If
            Persist.setReportConnection(objReport)
            objReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "releveFact2.pdf")
            Assert.IsTrue(File.Exists("releveFact2.pdf"))
            System.Diagnostics.Process.Start("releveFact2.pdf")

        objReport.Dispose()

        For Each oFact In colFact
            oFact.bDeleted = True
            oFact.Save()

        Next

        objCMD.bDeleted = True
        objCMD.save()
    End Sub
End Class



