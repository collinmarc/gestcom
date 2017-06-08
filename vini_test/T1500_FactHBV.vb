Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.IO
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

''' <summary>
''' Test des la classe FactHBV
''' </summary>
''' <remarks></remarks>
<TestClass()> Public Class T1500_FactHBV
    Inherits test_Base

    Private m_oProduit As Produit
    Private m_oFourn As Fournisseur
    Private m_oClient As Client
    Private m_oClient2 As Client
    Private m_oClient3 As Client
    Dim objProduit1 As Produit
    Dim objProduit2 As Produit
    Dim objProduit3 As Produit
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()

        m_oProduit = Produit.createandloadbyKey("PRDTEST1500")
        If m_oProduit IsNot Nothing Then
            m_oProduit.bDeleted = True
            m_oProduit.save()
        End If

        objProduit1 = Produit.createandloadbyKey("PRDTEST1500-1")
        If objProduit1 IsNot Nothing Then
            objProduit1.bDeleted = True
            objProduit1.save()
        End If

        objProduit2 = Produit.createandloadbyKey("PRDTEST1500-2")
        If objProduit2 IsNot Nothing Then
            objProduit2.bDeleted = True
            objProduit2.save()
        End If

        objProduit3 = Produit.createandloadbyKey("PRDTEST1500-3")
        If objProduit3 IsNot Nothing Then
            objProduit3.bDeleted = True
            objProduit3.save()
        End If

        m_oFourn = Fournisseur.createandload("FRNTEST1500")
        If m_oFourn IsNot Nothing Then
            m_oFourn.bDeleted = True
            m_oFourn.Save()
        End If
        m_oFourn = New Fournisseur("FRNTEST1500", "MonFournisseur")
        m_oFourn.Save()

        m_oProduit = New Produit("PRDTEST1500", m_oFourn, 1990)
        m_oProduit.save()
        objProduit1 = New Produit("PRDTEST1500-1", m_oFourn, 1990)
        objProduit1.save()
        objProduit2 = New Produit("PRDTEST1500-2", m_oFourn, 1990)
        objProduit2.save()
        objProduit3 = New Produit("PRDTEST1500-3", m_oFourn, 1990)
        objProduit3.save()


        m_oClient = Client.createandload("CLTTEST1500")
        If m_oClient IsNot Nothing Then
            m_oClient.bDeleted = True
            m_oClient.save()
        End If
        m_oClient = New Client("CLTTEST1500", "MonClient1")
        m_oClient.CodeCompta = "4100001"
        m_oClient.save()


        m_oClient2 = Client.createandload("CLTTEST1500-2")
        If m_oClient2 IsNot Nothing Then
            m_oClient2.bDeleted = True
            m_oClient2.save()
        End If
        m_oClient2 = New Client("CLTTEST1500-2", "MonClient2")
        m_oClient2.rs = "MonClient2RS"
        m_oClient2.CodeCompta = "4100001"
        m_oClient2.save()

        m_oClient2 = Client.createandload("CLTTEST1500-2")
        If m_oClient3 IsNot Nothing Then
            m_oClient3.bDeleted = True
            m_oClient3.save()
        End If
        m_oClient3 = New Client("CLTTEST1500-3", "MonClient3")
        m_oClient3.rs = "MonClient3RS"
        m_oClient3.CodeCompta = "4100001"
        m_oClient3.save()


    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        MyBase.TestCleanup()

    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objFactHBV As FactHBV
        Dim codeEtat As Integer

        objFactHBV = New FactHBV(m_oClient)
        Assert.IsTrue(objFactHBV.oTiers.Equals(m_oClient), "Le Client est identique à celui passé en paramètre à la création")
        Assert.IsTrue(objFactHBV.etat.codeEtat = vncEnums.vncEtatCommande.vncFactHBVGeneree)
        Assert.IsTrue(objFactHBV.code = "")
        Assert.IsTrue(objFactHBV.totalHT = 0)
        Assert.IsTrue(objFactHBV.totalTTC = 0)
        Assert.IsTrue(objFactHBV.id = 0)
        Assert.AreEqual(m_oClient.idModeReglement1, objFactHBV.idModeReglement)


        objFactHBV.dateCommande = CDate("06/02/05")
        Assert.IsTrue(objFactHBV.dateCommande.Equals(CDate("06/02/05")))

        objFactHBV.totalHT = 100
        Assert.IsTrue(objFactHBV.totalHT = 100)
        objFactHBV.totalTTC = 110
        Assert.IsTrue(objFactHBV.totalTTC = 110)


        objFactHBV.idModeReglement = 45
        Assert.IsTrue(objFactHBV.idModeReglement = 45, "Test idModeReglement")

        'Test de l'état de la Facture
        codeEtat = objFactHBV.etat.codeEtat

        objFactHBV.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactHBVGenerer)
        Assert.IsTrue(objFactHBV.etat.codeEtat = codeEtat, "Code Etat ne change pas")


        objFactHBV.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactHBVExporter)
        Assert.AreEqual(vncEnums.vncEtatCommande.vncFactHBVExportee, objFactHBV.etat.codeEtat, "Code Etat ne change pas")
        objFactHBV.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactHBVAnnExporter)
        Assert.AreEqual(vncEnums.vncEtatCommande.vncFactHBVGeneree, objFactHBV.etat.codeEtat, "Code Etat ne change pas")





        objFactHBV.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactHBVExporter)
        Assert.IsTrue(objFactHBV.etat.codeEtat = vncEnums.vncEtatCommande.vncFactHBVExportee, "Code Etat change ")
        objFactHBV.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactHBVGenerer)
        Assert.IsTrue(objFactHBV.etat.codeEtat = vncEnums.vncEtatCommande.vncFactHBVExportee, "Code Etat ne change pas")
        objFactHBV.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactHBVAnnGenerer)
        Assert.IsTrue(objFactHBV.etat.codeEtat = vncEnums.vncEtatCommande.vncFactHBVExportee, "Code Etat ne change pas")

        'IdCommande
        objFactHBV.idCommande = 15060
        Assert.AreEqual(CLng(15060), objFactHBV.idCommande)

    End Sub 'T10_Object
    <TestMethod()> Public Sub T15_GETDERNNUMFACT()

        Dim oFactHBV As FactHBV
        oFactHBV = New FactHBV(m_oClient)
        Dim oFactHBV2 As FactHBV
        oFactHBV2 = New FactHBV(m_oClient)
        Assert.AreEqual("", oFactHBV.code)
        oFactHBV.setNewcode()
        Assert.AreNotEqual("", oFactHBV.code)
        Dim nNumFact As Integer
        nNumFact = CInt(oFactHBV.code)
        oFactHBV2.setNewcode()
        Dim nNumFact2 As Integer
        nNumFact2 = CInt(oFactHBV2.code)
        Assert.AreEqual(nNumFact2, nNumFact + 1)

    End Sub
    <TestMethod()> Public Sub T20_DBLG_OBJECT()
        Dim objFactHBV As FactHBV
        Dim objFactHBV2 As FactHBV
        Dim nId As Integer
        Dim oLg As LgFactHBV

        'Création d'une facture HOBIVIN
        objFactHBV = New FactHBV(m_oClient)
        objFactHBV.dEcheance = "01/04/1964"
        objFactHBV.idModeReglement = 1

        'Ajout D'une Ligne de facture
        oLg = New LgFactHBV()
        oLg.oProduit = m_oProduit
        oLg.qteCommande = 5
        oLg.qteLiv = 7
        oLg.qteFact = 10
        oLg.bGratuit = False
        oLg.prixU = 50.5
        oLg.prixHT = 60.5
        oLg.prixTTC = 70.5
        objFactHBV.AjouteLigneFactHBV(oLg)

        'Sauvegarde de la facture
        Assert.IsTrue(objFactHBV.Save(), objFactHBV.getErreur())

        'Rechargement de la facture
        nId = objFactHBV.id
        objFactHBV2 = FactHBV.createandload(nId)

        'Rechargement des lignes
        objFactHBV2.loadcolLignes()

        oLg = objFactHBV2.colLignes(1)
        Assert.AreEqual(oLg.ProduitCode, m_oProduit.code)
        Assert.AreEqual(oLg.qteCommande, CDec(5))
        Assert.AreEqual(oLg.qteLiv, CDec(7))
        Assert.AreEqual(oLg.qteFact, CDec(10))
        Assert.AreEqual(oLg.prixU, CDec(50.5))
        Assert.AreEqual(oLg.prixHT, CDec(60.5))
        Assert.AreEqual(oLg.prixTTC, CDec(70.5))
        Assert.IsFalse(oLg.bGratuit)

        oLg.oProduit = objProduit3
        oLg.qteCommande = 50
        oLg.qteLiv = 70
        oLg.qteFact = 100
        oLg.prixU = 150.5
        oLg.prixHT = 160.5
        oLg.prixTTC = 170.5
        oLg.bGratuit = False


        'Sauvegarde de la facture
        Assert.IsTrue(objFactHBV2.Save(), objFactHBV2.getErreur())

        'Rechargement de la facture
        nId = objFactHBV.id
        objFactHBV2 = FactHBV.createandload(nId)

        'Rechargement des lignes
        objFactHBV2.loadcolLignes()

        Assert.AreEqual(1, objFactHBV2.colLignes.Count)
        oLg = objFactHBV2.colLignes(1)
        Assert.AreEqual(oLg.ProduitCode, objProduit3.code)
        Assert.AreEqual(oLg.qteCommande, CDec(50))
        Assert.AreEqual(oLg.qteLiv, CDec(70))
        Assert.AreEqual(oLg.qteFact, CDec(100))
        Assert.AreEqual(oLg.prixU, CDec(150.5))
        Assert.AreEqual(oLg.prixHT, CDec(160.5))
        Assert.AreEqual(oLg.prixTTC, CDec(170.5))
        Assert.IsFalse(oLg.bGratuit)



        objFactHBV2.bDeleted = True
        Assert.IsTrue(objFactHBV2.Save())


    End Sub
    ''' <summary>
    ''' Test de la sauvegarde de la  facture
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T20_DB_FACT()
        Dim objFactHBV As FactHBV
        Dim objFactHBV2 As FactHBV
        Dim nId As Integer

        'Création d'une facture HOBIVIN
        objFactHBV = New FactHBV(m_oClient)
        objFactHBV.dEcheance = "01/04/1964"
        objFactHBV.idModeReglement = 1
        objFactHBV.idCommande = 15060


        'Sauvegarde de la facture
        Assert.IsTrue(objFactHBV.Save(), objFactHBV.getErreur())

        'Rechargement de la facture
        nId = objFactHBV.id
        objFactHBV2 = FactHBV.createandload(nId)

        Assert.AreEqual(objFactHBV.id, objFactHBV2.id)
        Assert.AreEqual(objFactHBV.code, objFactHBV2.code)
        Assert.AreEqual(CDate("01/04/1964"), objFactHBV2.dEcheance, "Date Echeance")
        Assert.AreEqual(1, objFactHBV2.idModeReglement, "Mode de reglement")
        Assert.AreEqual(CLng(15060), objFactHBV2.idCommande, "idCommande")



    End Sub
    ''' <summary>
    ''' Test de la sauvegarde des lignes de factures
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T20_DB_LG()
        Dim objFactHBV As FactHBV
        Dim objFactHBV2 As FactHBV
        Dim nId As Integer
        Dim oLg As LgFactHBV

        'Création d'une facture HOBIVIN
        objFactHBV = New FactHBV(m_oClient)
        objFactHBV.dEcheance = "01/04/1964"
        objFactHBV.idModeReglement = 1

        'Ajout D'une Ligne de facture
        oLg = New LgFactHBV()
        oLg.oProduit = m_oProduit
        oLg.qteFact = 10
        oLg.prixU = m_oProduit.TarifA
        objFactHBV.AjouteLigneFactHBV(oLg)
        Assert.AreEqual(1, objFactHBV.colLignes.Count())

        'Ajout d'une sconde Ligne de facture
        oLg = New LgFactHBV()
        oLg.oProduit = m_oProduit
        oLg.qteFact = 20
        oLg.prixU = m_oProduit.TarifA
        objFactHBV.AjouteLigneFactHBV(oLg)
        Assert.AreEqual(2, objFactHBV.colLignes.Count())

        'Sauvegarde de la facture
        Assert.IsTrue(objFactHBV.Save(), objFactHBV.getErreur())

        'Rechargement de la facture
        nId = objFactHBV.id
        objFactHBV2 = FactHBV.createandload(nId)

        Assert.AreEqual(objFactHBV.id, objFactHBV2.id)
        Assert.AreEqual(objFactHBV.code, objFactHBV2.code)
        Assert.AreEqual(CDate("01/04/1964"), objFactHBV2.dEcheance, "Date Echeance")
        Assert.AreEqual(1, objFactHBV2.idModeReglement, "Mode de reglement")

        'Rechargement des lignes
        objFactHBV2.loadcolLignes()
        Assert.AreEqual(2, objFactHBV2.colLignes.Count())

        oLg = objFactHBV2.colLignes(1)
        Assert.AreEqual(oLg.ProduitCode, m_oProduit.code)
        Assert.AreEqual(oLg.qteFact, CDec(10))
        Assert.AreEqual(oLg.prixU, m_oProduit.TarifA)

        oLg = objFactHBV2.colLignes(2)
        Assert.AreEqual(oLg.ProduitCode, m_oProduit.code)
        Assert.AreEqual(oLg.qteFact, CDec(20))
        Assert.AreEqual(oLg.prixU, m_oProduit.TarifA)

        objFactHBV2.bDeleted = True
        Assert.IsTrue(objFactHBV2.Save())


    End Sub
    ''' <summary>
    ''' Test de la sauvegarde des lignes de factures en ajout ou en suppression
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T20_DB_LG_ADD_DELETE()
        Dim objFactHBV As FactHBV
        Dim objFactHBV2 As FactHBV
        Dim nId As Integer
        Dim oLg As LgFactHBV

        'Création d'une facture HOBIVIN
        objFactHBV = New FactHBV(m_oClient)
        objFactHBV.dEcheance = "01/04/1964"
        objFactHBV.idModeReglement = 1

        'Ajout D'une Ligne de facture
        oLg = New LgFactHBV()
        oLg.oProduit = m_oProduit
        oLg.qteFact = 10
        oLg.prixU = m_oProduit.TarifA
        objFactHBV.AjouteLigneFactHBV(oLg)
        Assert.AreEqual(1, objFactHBV.colLignes.Count())

        'Ajout d'une seconde Ligne de facture
        oLg = New LgFactHBV()
        oLg.oProduit = m_oProduit
        oLg.qteFact = 20
        oLg.prixU = m_oProduit.TarifA
        objFactHBV.AjouteLigneFactHBV(oLg)
        Assert.AreEqual(2, objFactHBV.colLignes.Count())

        'Sauvegarde de la facture
        Assert.IsTrue(objFactHBV.Save(), objFactHBV.getErreur())

        'Rechargement de la facture
        nId = objFactHBV.id
        objFactHBV2 = FactHBV.createandload(nId)

        'Rechargement des lignes
        objFactHBV2.loadcolLignes()
        Assert.AreEqual(2, objFactHBV2.colLignes.Count())

        'Ajout d'une Troisième Ligne de facture
        oLg = New LgFactHBV()
        oLg.oProduit = m_oProduit
        oLg.qteFact = 30
        oLg.prixU = m_oProduit.TarifA
        objFactHBV2.AjouteLigneFactHBV(oLg)

        'Sauvegarde de la facture
        Assert.IsTrue(objFactHBV2.Save(), objFactHBV2.getErreur())

        'Rechargement de la facture
        nId = objFactHBV.id
        objFactHBV = FactHBV.createandload(nId)

        'Rechargement des lignes
        objFactHBV.loadcolLignes()
        Assert.AreEqual(3, objFactHBV.colLignes.Count())


        'Suppression de la Première ligne
        oLg = objFactHBV.colLignes(1)
        oLg.bDeleted = True

        'Sauvegarde de la facture
        Assert.IsTrue(objFactHBV.Save(), objFactHBV.getErreur())

        'Rechargement de la facture
        nId = objFactHBV.id
        objFactHBV = FactHBV.createandload(nId)

        'Rechargement des lignes
        objFactHBV.loadcolLignes()
        Assert.AreEqual(2, objFactHBV.colLignes.Count())

        'Vérification de la suppression de la Ligne 1

        oLg = objFactHBV.colLignes(1)
        Assert.AreEqual(oLg.ProduitCode, m_oProduit.code)
        Assert.AreEqual(oLg.qteFact, CDec(20))
        Assert.AreEqual(oLg.prixU, m_oProduit.TarifA)

        oLg = objFactHBV.colLignes(2)
        Assert.AreEqual(oLg.ProduitCode, m_oProduit.code)
        Assert.AreEqual(oLg.qteFact, CDec(30))
        Assert.AreEqual(oLg.prixU, m_oProduit.TarifA)

        objFactHBV.bDeleted = True
        Assert.IsTrue(objFactHBV.Save())


    End Sub

    <TestMethod()> Public Sub T50_LISTE()
        Dim objFact As FactHBV
        Dim objFact2 As FactHBV
        Dim objFact3 As FactHBV
        Dim ocol As Collection
        Dim strCodeFact As String

        objFact = New FactHBV(m_oClient)
        objFact.dateCommande = "06/02/1964"
        objFact.totalHT = 110
        Assert.IsTrue(objFact.Save(), "Sauvegarde de la facture")
        strCodeFact = objFact.code

        objFact2 = New FactHBV(m_oClient2)
        objFact2.dateCommande = "07/02/1964"
        objFact2.totalHT = 120
        Assert.IsTrue(objFact2.Save(), "Sauvegarde de la facture")

        objFact3 = New FactHBV(m_oClient3)
        objFact3.dateCommande = "08/02/1964"
        objFact3.totalHT = 130
        Assert.IsTrue(objFact3.Save(), "Sauvegarde de la facture")


        ocol = FactHBV.getListe(strCodeFact)
        Assert.IsTrue(ocol.Count = 1, "1 elements dans la liste sur le code facture")

        ocol = FactHBV.getListe(, "MonClient2%")
        Assert.IsTrue(ocol.Count = 1, "1 elements dans la liste sur la code Client")

        ocol = FactHBV.getListe(CDate("01/02/64"), CDate("28/02/64"), , vncEnums.vncEtatCommande.vncFactHBVGeneree)
        Assert.IsTrue(ocol.Count >= 2, "au moins 2 elements dans la liste sur etat = Générée")


        ocol = FactHBV.getListe()
        Assert.IsTrue(ocol.Count >= 3, "3 elements dans la liste")

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), "Suppression de facture1")

        objFact2.bDeleted = True
        Assert.IsTrue(objFact2.Save(), "Suppression de facture2")

        objFact3.bDeleted = True
        Assert.IsTrue(objFact3.Save(), "Suppression de facture3")

    End Sub
    <TestMethod(), Ignore()> Public Sub T60_EXPORTCOMPTA()
        Dim objFact As FactHBV
        Dim objLgFact As LgFactHBV
        Dim nFile As Integer
        Dim v As String
        Dim strFile As String = "F:\temp\export.txt"

        objFact = New FactHBV(m_oClient)
        Assert.IsTrue(objFact.Save(), "Sauvegarde de la facture")

        objLgFact = New LgFactHBV
        objLgFact.prixHT = 150.55

        objFact.AjouteLigneFactHBV(objLgFact)
        If System.IO.File.Exists(strFile) Then
            System.IO.File.Delete(strFile)
        End If
        objFact.Exporter(strFile)
        Try
            nFile = FreeFile()
            FileOpen(nFile, strFile, OpenMode.Input, OpenAccess.Read)
            While Not EOF(nFile)
                v = LineInput(nFile)
                Assert.IsTrue(MsgBox(v, MsgBoxStyle.YesNo, "Ligne du fichier d'export") = MsgBoxResult.Yes)
            End While
            FileClose(nFile)
        Catch ex As Exception
            Assert.IsTrue(False, ex.ToString)
        End Try

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), "Suppression de la facture")

    End Sub


    ''' <summary>
    ''' Test l'export vers Quadra
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod(), Ignore()> Public Sub T100_EXPORT()
        Dim objFact As FactHBV
        Dim strLines As String()
        Dim strLine1 As String
        Dim strLine2 As String
        Dim strLine3 As String

        objFact = New FactHBV(m_oClient)
        objFact.dateFacture = CDate("06/02/1964")
        objFact.totalHT = 150.56
        objFact.totalTTC = 180.89
        objFact.dEcheance = "01/04/1964"


        Assert.IsTrue(objFact.Save(), objFact.getErreur())




        'Save
        Assert.IsTrue(objFact.Save(), "Insert" & objFact.getErreur)
        If File.Exists("./T20_EXPORT.txt") Then
            File.Delete("./T20_EXPORT.txt")
        End If

        objFact.exporter("./T20_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T20_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T20_EXPORT.txt")
        Assert.AreEqual(3, strLines.Length, "3 lignes d'export")

        strLine1 = strLines(0)
        strLine2 = strLines(1)
        strLine3 = strLines(2)

        Assert.AreEqual(231, strLine1.Length)
        Assert.AreEqual("M", strLine1.Substring(0, 1))
        Assert.AreEqual(m_oClient.CodeCompta, Trim(strLine1.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine1.Substring(9, 2)))
        Assert.AreEqual("060264", strLine1.Substring(14, 6))
        Assert.AreEqual(("F:" + objFact.code + " " + m_oClient.rs + Space(20)).Substring(0, 19), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("D", strLine1.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine1.Substring(42, 13)))
        Assert.AreEqual("010464", Trim(strLine1.Substring(63, 6)))

        Assert.AreEqual(231, strLine2.Length)
        Assert.AreEqual("M", strLine2.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTETVA")), Trim(strLine2.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine2.Substring(9, 2)))
        Assert.AreEqual("060264", strLine2.Substring(14, 6))
        Assert.AreEqual(("F:" + objFact.code + " " + m_oClient.rs + Space(20)).Substring(0, 19), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine2.Substring(41, 1))
        Assert.AreEqual((180.89 - 150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine2.Substring(42, 13)))

        Assert.AreEqual(231, strLine3.Length)
        Assert.AreEqual("M", strLine3.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTEPRODUIT")), Trim(strLine3.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine3.Substring(9, 2)))
        Assert.AreEqual("060264", strLine3.Substring(14, 6))
        Assert.AreEqual(("F:" + objFact.code + " " + m_oClient.rs + Space(20)).Substring(0, 19), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine3.Substring(41, 1))
        Assert.AreEqual((150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine3.Substring(42, 13)))

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save())
    End Sub

    ''' <summary>
    ''' Création depuis une commande Client
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T110_CREATEFROMCOMMANDECLT()
        Dim objCMD As CommandeClient
        Dim oFactHBV As FactHBV

        'CREATION D'UNE COMMANDE CLIENT
        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = CDate("06/02/1964")
        'Ajout de 2 lignes de Commandes
        objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 12, 20, True, 120, (120 * 1.196))
        objCMD.AjouteLigne(objCMD.getNextNumLg, objProduit2, 12, 20, True, 125, (125 * 1.196))

        Assert.IsTrue(objCMD.save())

        oFactHBV = New FactHBV(objCMD)

        Assert.AreEqual(CLng(objCMD.id), oFactHBV.idCommande)
        Assert.AreEqual(objCMD.dateCommande, oFactHBV.dateCommande)
        Assert.AreEqual(objCMD.colLignes.Count, oFactHBV.colLignes.Count)
        Dim oLgCMD As LgCommande
        Dim oLgHBV As LgFactHBV

        oLgCMD = objCMD.colLignes(1)
        oLgHBV = oFactHBV.colLignes(1)

        Assert.AreEqual(oLgCMD.oProduit.id, oLgHBV.oProduit.id)
        Assert.AreEqual(oLgCMD.qteCommande, oLgHBV.qteCommande)
        Assert.AreEqual(oLgCMD.qteLiv, oLgHBV.qteLiv)
        Assert.AreEqual(oLgCMD.qteFact, oLgHBV.qteFact)
        Assert.AreEqual(oLgCMD.prixU, oLgHBV.prixU)
        Assert.AreEqual(oLgCMD.prixHT, oLgHBV.prixHT)
        Assert.AreEqual(oLgCMD.prixTTC, oLgHBV.prixTTC)
        Assert.AreEqual(oLgCMD.bGratuit, oLgHBV.bGratuit)

        oLgCMD = objCMD.colLignes(2)
        oLgHBV = oFactHBV.colLignes(2)

        Assert.AreEqual(oLgCMD.oProduit.id, oLgHBV.oProduit.id)
        Assert.AreEqual(oLgCMD.qteCommande, oLgHBV.qteCommande)
        Assert.AreEqual(oLgCMD.qteLiv, oLgHBV.qteLiv)
        Assert.AreEqual(oLgCMD.qteFact, oLgHBV.qteFact)
        Assert.AreEqual(oLgCMD.prixU, oLgHBV.prixU)
        Assert.AreEqual(oLgCMD.prixHT, oLgHBV.prixHT)
        Assert.AreEqual(oLgCMD.prixTTC, oLgHBV.prixTTC)
        Assert.AreEqual(oLgCMD.bGratuit, oLgHBV.bGratuit)
    End Sub
    <TestMethod()> Public Sub T120_CALCULPRIXLG()
        m_oProduit.idTVA = Param.TVAdefaut.id

        Dim oLg As New LgFactHBV()
        oLg.oProduit = m_oProduit
        oLg.qteCommande = 5
        oLg.qteLiv = 7
        oLg.qteFact = 10
        oLg.prixU = 15.5

        Assert.IsTrue(oLg.calculPrixTotal())
        Assert.AreEqual(CDec(10 * 15.5), oLg.prixHT)
        Assert.AreEqual(CDec(10 * 15.5 * (1 + (Param.TVAdefaut.valeur / 100))), oLg.prixTTC)

        oLg.bGratuit = True
        Assert.IsTrue(oLg.calculPrixTotal())
        Assert.AreEqual(CDec(0), oLg.prixHT)
        Assert.AreEqual(CDec(0), oLg.prixTTC)


    End Sub
    <TestMethod()> Public Sub T120_CALCULPRIXFACT()

        Dim oFact As FactHBV

        
        m_oProduit.idTVA = Param.TVAdefaut.id
        objProduit2.idTVA = Param.TVAdefaut.id
        oFact = New FactHBV(m_oClient)

        Dim oLg As New LgFactHBV()
        oLg.oProduit = m_oProduit
        oLg.qteCommande = 5
        oLg.qteLiv = 7
        oLg.qteFact = 10
        oLg.prixU = 15.5

        oFact.AjouteLigneFactHBV(oLg)

        oLg = New LgFactHBV()
        oLg.oProduit = objProduit2
        oLg.qteCommande = 50
        oLg.qteLiv = 70
        oLg.qteFact = 100
        oLg.prixU = 150.5
        oFact.AjouteLigneFactHBV(oLg)


        Assert.IsTrue(oFact.calculPrixTotal())

        Assert.AreEqual(CDec(10 * 15.5) + CDec(100 * 150.5), oFact.totalHT)
        Assert.AreEqual(CDec(oFact.totalHT * (1 + (Param.TVAdefaut.valeur / 100))), oFact.totalTTC)



    End Sub
    ''' <summary>
    ''' Test du chargement depuis  L'id d'uine commmande
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T130_LOADFROMCMD()

        Dim objCMD As CommandeClient
        Dim oFactHBV As FactHBV

        'CREATION D'UNE COMMANDE CLIENT
        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = CDate("06/02/1964")
        'Ajout de 2 lignes de Commandes
        objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 12, 20, True, 120, (120 * 1.196))
        objCMD.AjouteLigne(objCMD.getNextNumLg, objProduit2, 12, 20, True, 125, (125 * 1.196))

        Assert.IsTrue(objCMD.save())

        Dim nId As Long
        nId = objCMD.id

        'Chargement avant la génération
        oFactHBV = FactHBV.createandloadFromCmd(nId)
        Assert.AreEqual(0, oFactHBV.id)

        oFactHBV = New FactHBV(objCMD)
        oFactHBV.Save()
        Dim nIdFact As Integer
        nIdFact = oFactHBV.id

        'Chargement après génération
        oFactHBV = FactHBV.createandloadFromCmd(nId)
        Assert.AreEqual(nIdFact, oFactHBV.id)
        Assert.AreEqual(nId, oFactHBV.idCommande)

        oFactHBV.bDeleted = True
        oFactHBV.Save()


    End Sub
    ''' <summary>
    ''' Test du chargement depuis  L'id d'uine commmande
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T150_GENERATIONETAT()

        Dim objCMD As CommandeClient
        Dim oFactHBV As FactHBV

        'Initialisation des Données pour la facture
        m_oClient.rs = "CLIENT RS"
        m_oClient.AdresseLivraison.rue1 = "AdresseLivraison Rue1"
        m_oClient.AdresseLivraison.rue2 = "AdresseLivraison Rue2"
        m_oClient.AdresseLivraison.cp = "35000"
        m_oClient.AdresseLivraison.ville = "Adresse Livraison Ville"
        m_oClient.AdresseLivraison.tel = "0299559955"
        m_oClient.AdresseLivraison.fax = "0102030405"
        m_oClient.AdresseLivraison.port = "0605040302"
        m_oClient.AdresseLivraison.Email = "liv@client.Com"

        m_oClient.AdresseFacturation.rue1 = "AdresseFacturation Rue1"
        m_oClient.AdresseFacturation.rue2 = "AdresseFacturation Rue2"
        m_oClient.AdresseFacturation.cp = "36000"
        m_oClient.AdresseFacturation.ville = "Adresse Facturation Ville"
        m_oClient.AdresseFacturation.tel = "0399559955"
        m_oClient.AdresseFacturation.fax = "0402030405"
        m_oClient.AdresseFacturation.port = "0505040302"
        m_oClient.AdresseFacturation.Email = "fact@client.Com"
        m_oClient.save()


        'CREATION D'UNE COMMANDE CLIENT
        objCMD = New CommandeClient(m_oClient)
        objCMD.DuppliqueCaracteristiqueTiers()
        objCMD.dateCommande = CDate("06/02/1964")
        objCMD.dateEnlevement = CDate("08/02/1964")
        objCMD.dateLivraison = CDate("10/02/1964")
        objCMD.NamePrestashop = "ASXEDC"

        'Ajout de 2 lignes de Commandes
        objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 12, 20)
        objCMD.AjouteLigne(objCMD.getNextNumLg, objProduit2, 5, 25.5)


        'On Simule la Livraison
        For Each oLg As LgCommande In objCMD.colLignes
            oLg.qteLiv = oLg.qteCommande
            oLg.qteFact = oLg.qteCommande
        Next

        Assert.IsTrue(objCMD.save())

        Dim nId As Long
        nId = objCMD.id

        oFactHBV = New FactHBV(objCMD)
        oFactHBV.calculPrixTotal()
        oFactHBV.Save()
        Dim nIdFact As Integer
        nIdFact = oFactHBV.id

        Dim objReport As ReportDocument
        Dim strReportName As String

        objReport = New ReportDocument
        strReportName = "V:\V5\vini_app\bin\Debug/crFactHobivin.rpt"
        objReport.Load(strReportName)
        objReport.SetParameterValue("IDCOMMANDE", oFactHBV.id)
        Persist.setReportConnection(objReport)

        objReport.ExportToDisk(ExportFormatType.PortableDocFormat, "FactHBV.pdf")

        Assert.IsTrue(File.Exists("FactHBV.pdf"))

        System.Diagnostics.Process.Start("FactHBV.pdf")

        oFactHBV.bDeleted = True
        oFactHBV.Save()


    End Sub
End Class



