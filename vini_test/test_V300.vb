'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NUnit.Extensions.Forms
Imports vini_DB
Imports vini_App



<TestClass()> Public Class test_V300
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


        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
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

        MyBase.TestCleanup()

    End Sub
    <TestMethod()> Public Sub T10_ExportWEBEDI()
        'Test de l'export avec la zone observation agrandie à 100 car et le nom du transoprteur à la fin sur 50 car
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
        objCMD.CommLivraison.comment = "123456789012345678901234567890123456789012345678901234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ45678901234567890123456789012345678901234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ45678901234567890123456789012345678901234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        objCMD.oTransporteur.nom = "Nom du transporteurNom du transporteurNom du transporteur"
        objCMD.AjouteLigne("10", m_objPRD, 10, 10)
        If System.IO.File.Exists("C:/temp/adel.txt") Then
            System.IO.File.Delete("C:/temp/adel.txt")
        End If
        objCMD.exporterWebEDI("C:/temp/adel.txt")
        nFile = FreeFile()
        FileOpen(nFile, "C:/temp/adel.txt", OpenMode.Input, OpenAccess.Read)
        nLineNumber = 0
        While Not EOF(nFile)
            nLineNumber = nLineNumber + 1
            strResult = LineInput(nFile)
            Console.WriteLine(strResult)
        End While

        Assert.AreEqual(1, nLineNumber, "Une seule Ligne de fichier")
        Assert.AreEqual(373, strResult.Length, "Longeur = 374-1")
        Assert.AreEqual("010", Mid(strResult, 292, 3), "Le Numéro lde ligne est sur 292-3 ")
        Assert.AreEqual(Left(objCMD.CommLivraison.comment, 100), Mid(strResult, 192, 100), "Le Commentaire sur 100 c")
        Assert.AreEqual(Left("Nom du transporteurNom du transporteurNom du transporteur", 50), Mid(strResult, 324, 50), "Le Nom du transporteur sur 50 c")

        FileClose(nFile)
        'Suppression du fichier créé
        System.IO.File.Delete("C:/temp/adel.txt")

    End Sub

    <TestMethod()> Public Sub T20_MVTSTK()
        'Test La structure mvtStock avec les informations de colisage

        Dim objMvt As mvtStock
        Dim nId As Integer

        objMvt = New mvtStock(#6/2/1964#, 1, vncTypeMvt.vncmvtRegul, 11, "Test")

        'Test des Etats
        Assert.AreEqual(vini_DB.vncEtatMVTSTK.vncMVTSTK_nFact, objMvt.Etat.codeEtat, "Etat par defaut")
        Assert.AreEqual(0, objMvt.idFactColisage, "ref Fact colisage par defaut")

        objMvt.changeEtat(vncActionFactColisage.vncActionAnnFacturer)
        Assert.AreEqual(vncEtatMVTSTK.vncMVTSTK_nFact, objMvt.Etat.codeEtat, "Etat par defaut")

        objMvt.changeEtat(vncActionFactColisage.vncActionFacturer)
        Assert.AreEqual(vncEtatMVTSTK.vncMVTSTK_Fact, objMvt.Etat.codeEtat, "Etat par defaut")

        objMvt.changeEtat(vncActionFactColisage.vncActionFacturer)
        Assert.AreEqual(vncEtatMVTSTK.vncMVTSTK_Fact, objMvt.Etat.codeEtat, "Etat par defaut")
        objMvt.changeEtat(vncActionFactColisage.vncActionAnnFacturer)
        Assert.AreEqual(vncEtatMVTSTK.vncMVTSTK_nFact, objMvt.Etat.codeEtat, "Etat par defaut")

        'IdFactColisage
        objMvt.idFactColisage = 123456
        Assert.AreEqual(123456, objMvt.idFactColisage, "IdFactColisage")

        Assert.IsTrue(objMvt.save(), "save")
        nId = objMvt.id

        objMvt = New mvtStock(#6/2/1999#, 15, vncTypeMvt.vncmvtBonAppro, 2, "Rien")

        objMvt.load(nId)
        Assert.AreEqual(nId, objMvt.id, "Id")
        Assert.AreEqual(#6/2/1964#, objMvt.datemvt, "dateMvt")
        Assert.AreEqual(1, objMvt.idProduit, "idPrd")
        Assert.AreEqual(vncTypeMvt.vncmvtRegul, objMvt.typeMvt, "idPrd")
        Assert.AreEqual(CDec(11), objMvt.qte, "Qte")
        Assert.AreEqual("Test", objMvt.libelle, "Libelle")
        Assert.AreEqual(vncEtatMVTSTK.vncMVTSTK_nFact, objMvt.Etat.codeEtat, "Etat")
        Assert.AreEqual(123456, objMvt.idFactColisage, "idFactColisage")

        objMvt.changeEtat(vncActionFactColisage.vncActionFacturer)
        objMvt.idFactColisage = 789456
        Assert.IsTrue(objMvt.save(), "save")
        nId = objMvt.id
        objMvt = New mvtStock(#6/2/1999#, 15, vncTypeMvt.vncmvtBonAppro, 2, "Rien")

        objMvt.load(nId)
        Assert.AreEqual(vncEtatMVTSTK.vncMVTSTK_Fact, objMvt.Etat.codeEtat, "Etat")
        Assert.AreEqual(789456, objMvt.idFactColisage, "idFactColisage")

        objMvt.bDeleted = True
        objMvt.save()

        Assert.IsFalse(objMvt.load(nId), "Load apres suppression")

    End Sub


    '<Test(), Ignore("done in T1060.T15_BD")> Public Sub T70_TarifProduit()

    '    Dim objTest As T1060_Produit
    '    objTest = New T1060_Produit()
    '    objTest.TestInitialize()
    '    objTest.T15_DB()
    '    objTest.TestCleanup()




    'End Sub

    '<Test(), Ignore("A Tester en Debug: NunitForm test")> Public Sub T80_FrmProduit()

    '    'Trjs tester en Debug !!!
    '    Dim ofrm As vini_app.frmProduit
    '    Dim otbTA As TextBoxTester
    '    Dim otbTB As TextBoxTester
    '    Dim otbTC As TextBoxTester

    '    m_objPRD.TarifA = 11.5
    '    m_objPRD.TarifB = 12.5
    '    m_objPRD.TarifC = 13.5

    '    ofrm = New frmProduit()
    '    ofrm.bTesting = True


    '    ofrm.Show()

    '    ofrm.setElementCourant2(m_objPRD)
    '    ofrm.AfficheElementCourant()

    '    otbTA = New TextBoxTester("tbTarifA", ofrm)
    '    otbTB = New TextBoxTester("tbTarifB", ofrm)
    '    otbTC = New TextBoxTester("tbTarifC", ofrm)
    '    Assert.AreEqual(m_objPRD.TarifA.ToString(), otbTA.Properties.Text)
    '    Assert.AreEqual(m_objPRD.TarifB.ToString(), otbTB.Properties.Text)
    '    Assert.AreEqual(m_objPRD.TarifC.ToString(), otbTC.Properties.Text)

    '    otbTA.Enter("15.15")
    '    otbTB.Enter("16.16")
    '    otbTC.Enter("16.16")
    '    ofrm.TST_frmSave()

    '    m_objPRD = Produit.createandload(m_objPRD.id)
    '    Assert.AreEqual(15.15, m_objPRD.TarifA)
    '    Assert.AreEqual(16.16, m_objPRD.TarifB)
    '    Assert.AreEqual(16.16, m_objPRD.TarifC)

    'End Sub
    '<Test(), Ignore("Fait dans le 1070.T15_DB")> Public Sub T90_TarifClient()

    '    Dim objTest As T1070_client
    '    objTest = New T1070_client()
    '    objTest.TestInitialize()
    '    objTest.T15_DB()
    '    objTest.TestCleanup()

    'End Sub
    '<Test(), Ignore("Test frm")> Public Sub T100_FrmClient()

    '    'Trjs tester en Debug !!!
    '    Dim ofrm As vini_app.frmClient
    '    Dim ocbxTarif As ComboBoxTester

    '    m_objCLT.CodeTarif = "C"

    '    ofrm = New frmClient()
    '    ofrm.bTesting = True


    '    ofrm.Show()

    '    ofrm.setElementCourant2(m_objCLT)
    '    ofrm.AfficheElementCourant()

    '    ocbxTarif = New ComboBoxTester("cbxCodeTarif", ofrm)
    '    Assert.AreEqual(m_objCLT.CodeTarif.ToString(), ocbxTarif.Text)

    '    ocbxTarif.Enter("B")
    '    ofrm.TST_frmSave()

    '    m_objCLT = Client.createandload(m_objCLT.id)
    '    Assert.AreEqual("B", m_objCLT.CodeTarif)

    'End Sub
    '<Test(), Ignore("test frm")> Public Sub T110_FrmLigneCommande()

    '    'Trjs tester en Debug !!!
    '    Dim ofrm As vini_app.dlgLgCommande
    '    Dim objLG As LgCommande
    '    Dim otb As TextBoxTester
    '    Dim ocb As ButtonTester

    '    m_objPRD.TarifA = 11.11
    '    m_objPRD.TarifB = 12.12
    '    m_objPRD.TarifC = 13.13
    '    m_objPRD.save()

    '    m_objCLT.CodeTarif = "B"

    '    objLG = New LgCommande(m_oCmd.id)
    '    objLG.num = m_oCmd.getNextNumLg()
    '    ofrm = New dlgLgCommande()
    '    ofrm.setElementCourant(objLG, m_objCLT)
    '    ofrm.cbProduit.Visible = True

    '    otb = New TextBoxTester("tbCodeProduit", ofrm)
    '    ocb = New ButtonTester("cbProduit", ofrm)
    '    otb.Enter(m_objPRD.code)
    '    ocb.Click()

    '    otb = New TextBoxTester("tbPrixUHT", ofrm)
    '    Assert.AreEqual(m_objPRD.TarifB.ToString(), otb.Text)
    '    m_objCLT.CodeTarif = "C"
    '    ofrm.setElementCourant(objLG, m_objCLT)
    '    otb = New TextBoxTester("tbCodeProduit", ofrm)
    '    ocb = New ButtonTester("cbProduit", ofrm)
    '    otb.Enter(m_objPRD.code)
    '    ocb.Click()
    '    otb = New TextBoxTester("tbPrixUHT", ofrm)
    '    Assert.AreEqual(m_objPRD.TarifC.ToString(), otb.Text)


    'End Sub
End Class


