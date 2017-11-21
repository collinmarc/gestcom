'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NUnit.Extensions.Forms
Imports vini_DB
Imports vini_App
Imports System.IO



<TestClass()> Public Class test4000_Reglement
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
        m_objFRN = New Fournisseur("FRNV4000", "FRn de'' test")
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

        m_objPRD = New Produit("PRDV4000", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objFRN2 = New Fournisseur("FRNV4000-2", "FRn de'' test2")
        m_objFRN2.rs = "FRNV4000-2"
        Assert.IsTrue(m_objFRN2.Save(), "FRN.Create")

        m_objPRD2 = New Produit("PRDV4000-2", m_objFRN2, 1990)
        Assert.IsTrue(m_objPRD2.save(), "Prod.Create")

        m_objFRN3 = New Fournisseur("FRNV4000-3", "FRn de'' test3")
        m_objFRN3.rs = "FRNV4000-3"
        Assert.IsTrue(m_objFRN3.Save(), "FRN3.Create")

        m_objPRD3 = New Produit("PRDV4000-3", m_objFRN3, 1990)
        Assert.IsTrue(m_objPRD3.save(), "Prod3.Create")

        m_objCLT = New Client("CLTV4000", "Client de' test")
        m_objCLT.rs = "Client de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")


        m_oCmd = New CommandeClient(m_objCLT)
        m_oCmd.dateCommande = #6/2/1964#
        m_oCmd.save()

        m_objPRD4 = New Produit("T60", m_objFRN, 2000)
        m_objPRD4.bStock = False
        m_objPRD4.save()

        Persist.shared_disconnect()
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

        Dim oTA As dsVinicomTableAdapters.REGLEMENTTableAdapter
        oTA = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
        oTA.DeleteBy_IDFACT(15)
        oTA.DeleteBy_IDFACT(16)

        MyBase.TestCleanup()
        Persist.shared_disconnect()
    End Sub


    ''' <summary>
    ''' Teste Les acesseurs de l'objet ainsi que les méthodes de DB
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T00_Object()

        Dim obj As Reglement
        Dim nId As Integer
        Dim oTA As dsVinicomTableAdapters.REGLEMENTTableAdapter
        Dim oDT As dsVinicom.REGLEMENTDataTable

        obj = New Reglement()
        obj.IdFact = 15
        obj.TypeFact = vncEnums.vncTypeDonnee.FACTCOL
        obj.Montant = 15.56
        obj.DateReglement = CDate("31/01/2008")
        obj.Reference = "CB123564"
        obj.Commentaire = "No comment"

        Assert.AreEqual(obj.IdFact, 15)
        Assert.AreEqual(obj.TypeFact, CInt(vncEnums.vncTypeDonnee.FACTCOL))
        Assert.AreEqual(obj.Montant, CDec(15.56))
        Assert.AreEqual(obj.DateReglement, CDate("31/01/2008"))
        Assert.AreEqual(obj.Reference, "CB123564")
        Assert.AreEqual(obj.Commentaire, "No comment")
        Assert.AreEqual(vncEnums.vncEtatReglement.vncRGLMT_Saisi, obj.Etat.codeEtat)

        obj.Montant = 16.66
        obj.DateReglement = CDate("29/02/2008")
        obj.Reference = "CHQ654231"
        obj.Commentaire = "My comment"
        obj.changeEtat(vncActionReglement.vncActionExporter)

        Assert.AreEqual(obj.Montant, CDec(16.66))
        Assert.AreEqual(obj.DateReglement, CDate("29/02/2008"))
        Assert.AreEqual(obj.Reference, "CHQ654231")
        Assert.AreEqual(obj.Commentaire, "My comment")
        Assert.AreEqual(vncEnums.vncEtatReglement.vncRGLMT_Export, obj.Etat.codeEtat)
        obj.changeEtat(vncActionReglement.vncActionAnnExporter)
        Assert.AreEqual(vncEnums.vncEtatReglement.vncRGLMT_Saisi, obj.Etat.codeEtat)


        'Save the Obj
        Assert.IsTrue(obj.save())
        Assert.AreNotEqual(0, obj.id)
        nId = obj.id

        'Reload It and check the value
        obj = New Reglement()
        obj.load(nId)
        Assert.AreEqual(obj.IdFact, 15)
        Assert.AreEqual(obj.TypeFact, CInt(vncEnums.vncTypeDonnee.FACTCOL))
        Assert.AreEqual(obj.Montant, CDec(16.66))
        Assert.AreEqual(obj.DateReglement, CDate("29/02/2008"))
        Assert.AreEqual(obj.Reference, "CHQ654231")
        Assert.AreEqual(obj.Commentaire, "My comment")
        Assert.AreEqual(vncEnums.vncEtatReglement.vncRGLMT_Saisi, obj.Etat.codeEtat)
        'Change som values
        obj.Montant = 17.77
        obj.DateReglement = CDate("31/03/2008")
        obj.Reference = "ESP"
        obj.Commentaire = "No comment"
        obj.changeEtat(vncActionReglement.vncActionExporter)

        'Update the obj
        Assert.IsTrue(obj.save(), obj.getErreur())
        'Reload It and check the value
        obj = New Reglement()
        obj.load(nId)
        Assert.AreEqual(obj.IdFact, 15)
        Assert.AreEqual(obj.TypeFact, CInt(vncEnums.vncTypeDonnee.FACTCOL))
        Assert.AreEqual(obj.Montant, CDec(17.77))
        Assert.AreEqual(obj.DateReglement, CDate("31/03/2008"))
        Assert.AreEqual(obj.Reference, "ESP")
        Assert.AreEqual(obj.Commentaire, "No comment")
        Assert.AreEqual(vncEnums.vncEtatReglement.vncRGLMT_Export, obj.Etat.codeEtat)


        oTA = New dsVinicomTableAdapters.REGLEMENTTableAdapter()
        oDT = oTA.GetDataBy_ID(nId)
        '        Assert.AreEqual(1, oDT.Count, "1 Row in datatable")

        'Delete 
        obj.bDeleted = True
        Assert.IsTrue(obj.save(), obj.getErreur())

        oDT = oTA.GetDataBy_ID(nId)
        Assert.AreEqual(0, oDT.Count, "No Rows in datatable")


    End Sub

    <TestMethod()> Public Sub T10_Liste()

        Dim oRglmt1 As Reglement
        Dim oRglmt2 As Reglement
        Dim oRglmt3 As Reglement
        Dim oRglmt4 As Reglement




        oRglmt1 = New Reglement()
        oRglmt1.IdFact = 15
        oRglmt1.TypeFact = vncEnums.vncTypeDonnee.FACTCOMM
        oRglmt1.DateReglement = CDate("15/01/1964")
        Assert.IsTrue(oRglmt1.save(), oRglmt1.getErreur())

        oRglmt2 = New Reglement()
        oRglmt2.IdFact = 15
        oRglmt2.TypeFact = vncEnums.vncTypeDonnee.FACTCOMM
        oRglmt2.DateReglement = CDate("16/01/1964")
        Assert.IsTrue(oRglmt2.save(), oRglmt2.getErreur())

        oRglmt3 = New Reglement()
        oRglmt3.IdFact = 16
        oRglmt3.TypeFact = vncEnums.vncTypeDonnee.FACTCOMM
        oRglmt3.DateReglement = CDate("17/01/1964")
        Assert.IsTrue(oRglmt3.save(), oRglmt3.getErreur())

        oRglmt4 = New Reglement()
        oRglmt4.IdFact = 16
        oRglmt4.TypeFact = vncEnums.vncTypeDonnee.FACTCOMM
        oRglmt4.DateReglement = CDate("18/01/1964")
        Assert.IsTrue(oRglmt4.save(), oRglmt4.getErreur())



        Dim oCol As Collection
        oCol = Reglement.getListe()
        Assert.IsTrue(oCol.Count > 4, "4 Reglements")
        oCol = Reglement.getListe(CDate("01/01/1964"), CDate("31/01/1964"))
        Assert.AreEqual(4, oCol.Count, "4 Reglements")
        oCol = Reglement.getListe(CDate("01/02/1964"), CDate("29/02/1964"))
        Assert.AreEqual(0, oCol.Count, "0 Reglements")
        oCol = Reglement.getListe(15, vncEnums.vncTypeDonnee.FACTCOMM, CDate("01/01/1964"), CDate("31/01/1964"))
        Assert.AreEqual(2, oCol.Count, "2 Reglements")
        oCol = Reglement.getListe(15, vncEnums.vncTypeDonnee.FACTCOMM, CDate("17/01/1964"), CDate("31/01/1964"))
        Assert.AreEqual(0, oCol.Count, "0 Reglements")

        oCol = Reglement.getListe(16, vncEnums.vncTypeDonnee.FACTCOMM, CDate("01/01/1964"), CDate("31/01/1964"))
        Assert.AreEqual(2, oCol.Count, "2 Reglements")
        oCol = Reglement.getListe(16, vncEnums.vncTypeDonnee.FACTCOMM, CDate("01/01/1964"), CDate("16/01/1964"))
        Assert.AreEqual(0, oCol.Count, "0 Reglements")


    End Sub

    <TestMethod()> Public Sub T10_Solde()

        Dim objFact As FactTRP
        Dim objReglement As Reglement
        objFact = New FactTRP(m_objCLT)
        objFact.totalTTC = 115.56
        objFact.totalHT = 100
        Assert.IsTrue(objFact.Save(), objFact.getErreur())

        'Solde = Total TTC avant Reglement
        Assert.AreEqual(objFact.totalTTC, objFact.getSolde(), "Solde = Total TTc")

        'Ajout d'un reglement
        objReglement = New Reglement(objFact)
        objReglement.DateReglement = DateTime.Now.ToShortDateString()
        objReglement.Montant = 10
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())

        'Chargement des reglements
        objFact.loadReglements()
        Assert.AreEqual(1, objFact.colReglements.Count, "1 reglement Chargé")

        'Solde = Monant TTC - Reglement
        Assert.AreEqual(objFact.totalTTC - objReglement.Montant, objFact.getSolde(), "Solde = Total TTc")

        'Ajout d'un Second Reglement
        objReglement = New Reglement(objFact)
        objReglement.DateReglement = DateTime.Now.ToShortDateString()
        objReglement.Montant = 10
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())

        objFact.loadReglements()

        'Solde = Monant TTC - 2 Reglements
        Assert.AreEqual(2, objFact.colReglements.Count, "1 reglement Chargé")

        Assert.AreEqual(objFact.totalTTC - (objReglement.Montant * 2), objFact.getSolde(), "Solde = Total TTc")


        objReglement.bDeleted = True
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), objFact.getErreur())

    End Sub
    <TestMethod()> Public Sub T20_Liste()

        Dim obj As Reglement
        Dim oCol As Collection

        obj = New Reglement()
        obj.IdFact = 15
        obj.TypeFact = vncEnums.vncTypeDonnee.FACTCOL
        obj.Montant = 15.56
        obj.DateReglement = CDate("31/01/1998")
        Assert.IsTrue(obj.save(), obj.getErreur())

        obj = New Reglement()
        obj.IdFact = 15
        obj.TypeFact = vncEnums.vncTypeDonnee.FACTCOMM
        obj.Montant = 15.56
        obj.DateReglement = CDate("31/01/1998")

        Assert.IsTrue(obj.save(), obj.getErreur())
        obj = New Reglement()
        obj.IdFact = 15
        obj.TypeFact = vncEnums.vncTypeDonnee.FACTTRP
        obj.Montant = 15.56
        obj.DateReglement = CDate("31/01/1998")
        Assert.IsTrue(obj.save(), obj.getErreur())

        oCol = Reglement.getListe(CDate("01/01/1998"), CDate("31/01/1998"), vncTypeDonnee.FACTCOL, vncEtatReglement.vncRGLMT_Tous)
        Assert.AreEqual(1, oCol.Count)

        oCol = Reglement.getListe(CDate("01/01/1998"), CDate("31/01/1998"), vncTypeDonnee.FACTCOMM, vncEtatReglement.vncRGLMT_Saisi)
        Assert.AreEqual(1, oCol.Count)

        oCol = Reglement.getListe(CDate("01/01/1998"), CDate("31/01/1998"), vncTypeDonnee.FACTTRP, vncEtatReglement.vncRGLMT_Saisi)
        Assert.AreEqual(1, oCol.Count)

        oCol = Reglement.getListe(CDate("01/01/1998"), CDate("31/01/1998"), vncTypeDonnee.FACTTRP, vncEtatReglement.vncRGLMT_Export)
        Assert.AreEqual(0, oCol.Count, "Pas de reglement Exporté")

        obj.changeEtat(vncActionReglement.vncActionExporter)
        Assert.IsTrue(obj.save(), obj.getErreur())

        oCol = Reglement.getListe(CDate("01/01/1998"), CDate("31/01/1998"), vncTypeDonnee.FACTTRP, vncEtatReglement.vncRGLMT_Export)
        Assert.AreEqual(1, oCol.Count, "1 reglement Exporté")

        oCol = Reglement.getListe(CDate("01/01/1998"), CDate("31/01/1998"), vncTypeDonnee.FACTTRP, vncEtatReglement.vncRGLMT_Saisi)
        Assert.AreEqual(0, oCol.Count, "0 reglement nonExporté")

        oCol = Reglement.getListe(CDate("01/01/1998"), CDate("31/01/1998"), vncTypeDonnee.FACTTRP, vncEtatReglement.vncRGLMT_Tous)
        Assert.AreEqual(1, oCol.Count, "1 reglement Tous")
    End Sub

    ''' <summary>
    ''' Test l'export vers Quadra
    ''' </summary>
    ''' <remarks>on ne travaille plus avec les reglement en gestcom</remarks>
    <TestMethod(), Ignore()> Public Sub T100_EXPORT()
        Dim objFact As FactColisage
        Dim strLines As String()
        Dim strLine1 As String
        Dim strLine2 As String
        Dim objReglement As Reglement

        objFact = New FactColisage(m_objFRN)
        objFact.periode = "1er Timestre 1964"
        objFact.dateFacture = CDate("06/02/1964")
        objFact.totalHT = 150.56
        objFact.totalTTC = 180.89
        Assert.IsTrue(objFact.save(), objFact.getErreur())

        objReglement = objFact.AddReglement(CDate("06/02/1964"), 180.89, "CHQ2005")
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())

        If File.Exists("./T100_EXPORT.txt") Then
            File.Delete("./T100_EXPORT.txt")
        End If

        objReglement.Exporter("./T100_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T100_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T100_EXPORT.txt")
        Assert.AreEqual(2, strLines.Length, "2 lignes d'export")

        strLine1 = strLines(0)
        strLine2 = strLines(1)

        Assert.AreEqual(231, strLine1.Length)
        Assert.AreEqual("M", strLine1.Substring(0, 1))
        Assert.AreEqual(m_objFRN.CodeCompta, Trim(strLine1.Substring(1, 8)))
        Assert.AreEqual("BA", Trim(strLine1.Substring(9, 2)))
        Assert.AreEqual("060264", strLine1.Substring(14, 6))
        Assert.AreEqual(m_objFRN.rs, Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine1.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine1.Substring(42, 13)))

        Assert.AreEqual(231, strLine2.Length)
        Assert.AreEqual("M", strLine2.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTEBANQUE")), Trim(strLine2.Substring(1, 8)))
        Assert.AreEqual("BA", Trim(strLine2.Substring(9, 2)))
        Assert.AreEqual("060264", strLine2.Substring(14, 6))
        Assert.AreEqual(m_objFRN.rs, Trim(strLine2.Substring(21, 20)))
        Assert.AreEqual("D", strLine2.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine2.Substring(42, 13)))


        objFact.bDeleted = True
        Assert.IsTrue(objFact.save())
    End Sub
    <TestMethod()> Public Sub T110_LsttFactureCOMnonReglées()

        Dim objFact As FactCom
        Dim objReglement As Reglement
        Dim colFact As Collection

        objFact = New FactCom(m_objFRN)
        objFact.totalTTC = 115.56
        objFact.totalHT = 100
        Assert.IsTrue(objFact.Save(), objFact.getErreur())

        colFact = FactCom.getListeNonReglee(, m_objFRN.rs)
        Assert.AreEqual(1, colFact.Count)

        'Ajout d'un reglement partiel
        objReglement = New Reglement(objFact)
        objReglement.DateReglement = DateTime.Now.ToShortDateString()
        objReglement.Montant = 10
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())


        colFact = FactCom.getListeNonReglee(, m_objFRN.rs)
        Assert.AreEqual(1, colFact.Count)

        'Ajout d'un reglement du solde
        objReglement = New Reglement(objFact)
        objReglement.DateReglement = DateTime.Now.ToShortDateString()
        objReglement.Montant = objFact.totalTTC - 10
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())

        colFact = FactCom.getListeNonReglee(, m_objFRN.rs)
        Assert.AreEqual(0, colFact.Count)

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), objFact.getErreur())



    End Sub

    <TestMethod()> Public Sub T110_LsttFactureCOLnonReglées()

        Dim objFact As FactColisage
        Dim objReglement As Reglement
        Dim colFact As Collection

        objFact = New FactColisage(m_objFRN)
        objFact.totalTTC = 115.56
        objFact.totalHT = 100
        Assert.IsTrue(objFact.save(), objFact.getErreur())

        colFact = FactColisage.getListeNonReglee(, m_objFRN.rs)
        Assert.AreEqual(1, colFact.Count)

        'Ajout d'un reglement partiel
        objReglement = New Reglement(objFact)
        objReglement.DateReglement = DateTime.Now.ToShortDateString()
        objReglement.Montant = 10
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())


        colFact = FactColisage.getListeNonReglee(, m_objFRN.rs)
        Assert.AreEqual(1, colFact.Count)

        'Ajout d'un reglement du solde
        objReglement = New Reglement(objFact)
        objReglement.DateReglement = DateTime.Now.ToShortDateString()
        objReglement.Montant = objFact.totalTTC - 10
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())

        colFact = FactColisage.getListeNonReglee(, m_objFRN.rs)
        Assert.AreEqual(0, colFact.Count)

        objFact.bDeleted = True
        Assert.IsTrue(objFact.save(), objFact.getErreur())



    End Sub

    <TestMethod()> Public Sub T110_LsttFactureTRPnonReglées()

        Dim objFact As FactTRP
        Dim objReglement As Reglement
        Dim colFact As Collection

        objFact = New FactTRP(m_objCLT)
        objFact.totalTTC = 115.56
        objFact.totalHT = 100
        Assert.IsTrue(objFact.Save(), objFact.getErreur())

        colFact = FactTRP.getListeNonReglee(, m_objCLT.rs)
        Assert.AreEqual(1, colFact.Count)

        'Ajout d'un reglement partiel
        objReglement = New Reglement(objFact)
        objReglement.DateReglement = DateTime.Now.ToShortDateString()
        objReglement.Montant = 10
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())


        colFact = FactTRP.getListeNonReglee(, m_objCLT.rs)
        Assert.AreEqual(1, colFact.Count)

        'Ajout d'un reglement du solde
        objReglement = New Reglement(objFact)
        objReglement.DateReglement = DateTime.Now.ToShortDateString()
        objReglement.Montant = objFact.totalTTC - 10
        Assert.IsTrue(objReglement.save(), objReglement.getErreur())

        colFact = FactTRP.getListeNonReglee(, m_objCLT.rs)
        Assert.AreEqual(0, colFact.Count)

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), objFact.getErreur())



    End Sub
End Class


