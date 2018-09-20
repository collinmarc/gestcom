Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.Collections.Generic
Imports System.IO

<TestClass()> Public Class TestMvtEDI
    Inherits test_Base
    Private m_oProduit As Produit
    Private m_oFourn As Fournisseur
    Private m_oClient As Client
    Dim objProduit1 As Produit
    Dim objProduit2 As Produit
    Dim objProduit3 As Produit
    Dim oCont As contenant
    Dim oCond As New Param()
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()

        Persist.shared_connect()


        m_oFourn = New Fournisseur("FRNT1100", "MonFournisseur")
        m_oFourn.Save()



        oCont = New contenant()
        oCont.code = "BLLE"
        oCont.poids = 1020
        oCont.defaut = False
        Assert.IsTrue(oCont.Save())

        oCond.code = "CAISS12"
        oCond.type = PAR_CONDITIONNEMENT
        oCond.valeur = 12
        oCond.defaut = False
        Assert.IsTrue(oCond.Save())

        Param.LoadcolParams()
        contenant.LoadcolContenants()

        m_oProduit = New Produit("TSTPRDT1100", m_oFourn, 1990)
        m_oProduit.idConditionnement = oCond.id
        m_oProduit.idContenant = oCont.id
        m_oProduit.save()

        m_oClient = New Client("CLTT1100", "MonClient")
        Debug.Assert(m_oClient.save(), "Creation du client")
        '            m_oClient()
        objProduit1 = New Produit("PRD1T1100", m_oFourn, 1994)
        Assert.IsTrue(objProduit1.save())
        objProduit2 = New Produit("PRD2T1100", m_oFourn, 1994)
        Assert.IsTrue(objProduit2.save())
        objProduit3 = New Produit("PRD3T1100", m_oFourn, 1994)
        Assert.IsTrue(objProduit3.save())

    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()

        objProduit1.bDeleted = True
        Assert.IsTrue(objProduit1.save())

        objProduit2.bDeleted = True
        Assert.IsTrue(objProduit2.save())

        objProduit3.bDeleted = True
        Assert.IsTrue(objProduit3.save())
        m_oProduit.bDeleted = True
        m_oProduit.save()
        TauxComm.deleteTauxComms(m_oFourn.id)

        m_oFourn.bDeleted = True
        m_oFourn.Save()
        m_oClient.bDeleted = True
        m_oClient.save()
        Persist.shared_disconnect()

        oCont.bDeleted = True
        oCont.Save()

        oCond.bDeleted = True
        oCond.Save()

        Param.LoadcolParams()
        MyBase.TestCleanup()
    End Sub

    <TestMethod()> Public Sub TestLireFichier()
        'Fichier issu de Groussard
        Dim strFileName As String = "TestFile.csv"
        CreerFichierCSV(strFileName)

        Dim olst As List(Of mvtEDI)
        olst = mvtEDI.getListFromFile(strFileName)
        Assert.IsNotNull(olst, "la liste est vide")
        Assert.AreEqual(14, olst.Count)
        Dim omvt As mvtEDI
        omvt = olst(0)
        Assert.AreEqual("47651", omvt.NumCMD)
        Assert.AreEqual("0", omvt.Entree)
        Assert.AreEqual("2", omvt.Sortie)

        omvt = olst(13)
        Assert.AreEqual("47674", omvt.NumCMD)
        Assert.AreEqual("0", omvt.Entree)
        Assert.AreEqual("8", omvt.Sortie)


    End Sub

    <TestMethod()> Public Sub TestRenommerLeFichier()
        'Fichier issu de Groussard
        Dim strFileName As String = "TestFile.csv"
        CreerFichierCSV(strFileName)
        Dim olst As List(Of mvtEDI)
        olst = mvtEDI.getListFromFile(strFileName)
        Assert.IsNotNull(olst, "la liste est vide")
        Assert.AreEqual(14, olst.Count)

        'Le Fichier est bien détruit
        Assert.IsFalse(System.IO.File.Exists(strFileName))
        Assert.IsTrue(System.IO.File.Exists(strFileName & ".TRT"))


    End Sub
    <TestMethod()> Public Sub TestLectureDuFichierSansleRenommer()
        'Fichier issu de Groussard
        Dim strFileName As String = "TestFile.csv"
        CreerFichierCSV(strFileName)
        Dim olst As List(Of mvtEDI)
        olst = mvtEDI.getListFromFile(strFileName, False)
        Assert.IsNotNull(olst, "la liste est vide")
        Assert.AreEqual(14, olst.Count)

        'Le Fichier n'est pas détruit
        Assert.IsTrue(System.IO.File.Exists(strFileName))


    End Sub
    <TestMethod()> Public Sub TestCumuls()

        Dim omvt As mvtEDI
        Dim oList As New List(Of mvtEDI)
        Dim oListCumuls As New List(Of mvtEDI)

        omvt = New mvtEDI("12456", "0", "6")
        oList.Add(omvt)
        omvt = New mvtEDI("12456", "0", "5")
        oList.Add(omvt)
        omvt = New mvtEDI("21789", "0", "15")
        oList.Add(omvt)
        omvt = New mvtEDI("21789", "0", "15")
        oList.Add(omvt)
        omvt = New mvtEDI("12789", "0", "25")
        oList.Add(omvt)
        omvt = New mvtEDI("12789", "0", "25")
        oList.Add(omvt)

        oListCumuls = mvtEDI.getListCumuls(oList)
        Assert.AreEqual(3, oListCumuls.Count)
        omvt = oListCumuls(0)
        Assert.AreEqual("12456", omvt.NumCMD)
        Assert.AreEqual("11", omvt.Sortie)
        omvt = oListCumuls(1)
        Assert.AreEqual("12789", omvt.NumCMD)
        Assert.AreEqual("50", omvt.Sortie)
        omvt = oListCumuls(2)
        Assert.AreEqual("21789", omvt.NumCMD)
        Assert.AreEqual("30", omvt.Sortie)

    End Sub
    <TestMethod()> Public Sub TestDownLoad()
        Dim oFTP As clsFTPVinicom
        oFTP = New clsFTPVinicom("ftp.cluster002.ovh.net:21", "vinicomwgs-vinidis", "Vinidis04092018", "EDI")
        Dim strFileNAme As String = "TSTEDI/testFile.csv"
        If System.IO.Directory.Exists("TSTEDI") Then
            System.IO.Directory.Delete("TSTEDI", True)
        End If
        System.IO.Directory.CreateDirectory("TSTEDI")
        CreerFichierCSV(strFileNAme)
        'Chargement du Fichier sur le FTP
        oFTP.uploadFromDir("TSTEDI")

        If System.IO.Directory.Exists("TSTEDI") Then
            System.IO.Directory.Delete("TSTEDI", True)
        End If
        System.IO.Directory.CreateDirectory("TSTEDI")

        mvtEDI.getFilesFromFTP("ftp.cluster002.ovh.net", "21", "vinicomwgs-vinidis", "Vinidis04092018", "EDI", "TSTEDI")

        'Le fichier existe Bien en local
        Assert.IsTrue(System.IO.File.Exists(strFileNAme))

        'Le Fichier n'existe plus en remote
        Assert.IsFalse(oFTP.ftpConnection.FtpFileExists("EDI/testFile.csv"), " Le Fichier distant n'a pas été écrasé")

    End Sub


    Private Sub CreerFichierCSV(pstrFileName As String)
        'Fichier issu de Groussard
        If System.IO.File.Exists(pstrFileName) Then
            System.IO.File.Delete(pstrFileName)
        End If
        System.IO.File.AppendAllText(pstrFileName, "Produit;Désignation produit;Date du mouvement;Numéro du document;Total U.C. entrée;Total U.C. sortie;Numéro de commande" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "001000H;AC CHAMPAGNE PANNIER SANS ETUI;03/07/2018;5344;0;2;00047651" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "001000H;AC CHAMPAGNE PANNIER SANS ETUI;18/07/2018;5486;0;1;00047820" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "001000H;AC CHAMPAGNE PANNIER SANS ETUI;19/07/2018;5492;0;6;00047828" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "001000H;AC CHAMPAGNE PANNIER SANS ETUI;23/07/2018;5519;0;1;00047861" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "001000H;AC CHAMPAGNE PANNIER SANS ETUI;26/07/2018;5544;0;8;00047890" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;02/07/2018;5317;0;6;00047615" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;02/07/2018;5322;0;4;00047623" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;04/07/2018;5366;0;10;00047675" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;04/07/2018;5364;0;4;00047670" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;04/07/2018;5355;0;13;00047657" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;04/07/2018;5365;0;12;00047672" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;05/07/2018;5367;0;6;00047676" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;05/07/2018;5367;0;18;00047676" & vbCrLf)
        System.IO.File.AppendAllText(pstrFileName, "002001;BIB 5L VDP MERLOT GALLICIAN;05/07/2018;5368;0;8;00047674" & vbCrLf)

        Assert.IsTrue(System.IO.File.Exists(pstrFileName))
    End Sub

    <TestMethod()> Public Sub T_VerifInfosLivraisonsOK()
        Dim objCmd As CommandeClient
        Dim objLgCmd1 As LgCommande
        Dim objLgCmd2 As LgCommande
        Dim objLgCmd3 As LgCommande
        Dim nidCmd As Integer

        'GIVEN j'ai un produit avec un contenant de 1020 et un conditionnement par 12 
        oCont.poids = 1020
        oCond.valeur = 12
        Assert.IsTrue(oCont.Save)
        Assert.IsTrue(oCond.Save)
        m_oProduit.idContenant = oCont.id
        m_oProduit.idConditionnement = oCond.id
        m_oProduit.save()
        'AND une Commande Client créée
        objCmd = New CommandeClient(m_oClient)


        'WHEN j'ajoute une ligne de 11 produit payant
        objLgCmd1 = objCmd.AjouteLigne("10", m_oProduit, 11, 10.5, False)
        'AND un produit Gratuit
        objLgCmd2 = objCmd.AjouteLigne("20", m_oProduit, 1, 0, True)
        'AND j'ajoute une autre Ligne gratuite
        objLgCmd3 = objCmd.AjouteLigne("30", m_oProduit, 1, 0, True)
        'AND je sauvegarde ma commande
        objCmd.save()
        objCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        objCmd.save()
        Assert.AreEqual(vncEtatCommande.vncValidee, objCmd.etat.codeEtat)
        Assert.AreEqual(CDec(2), objCmd.qteColis)
        'And le Fichier FTP contient  2 Colis
        'Fichier issu de Groussard
        If System.IO.File.Exists("TSTEDI/testGroussard.csv") Then
            System.IO.File.Delete("TSTEDI//testGroussard.csv")
        End If
        Dim strCode As String = objCmd.code
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "Produit;Désignation produit;Date du mouvement;Numéro du document;Total U.C. entrée;Total U.C. sortie;Numéro de commande" & vbCrLf)
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "xxxxxxx;-----xxxxxxxxxxxxxx;03/07/2018;5344;0;2;000" & strCode & vbCrLf)
        Assert.IsTrue(System.IO.File.Exists("TSTEDI//testGroussard.csv"))

        'THEN
        mvtEDI.VerificationCommandes("TSTEDI//testGroussard.csv")
        objCmd = CommandeClient.getListe(strCode, "", vncEtatCommande.vncRien, "")(1)
        objCmd.load()
        'L'état de la Ccommande est Livrée
        Assert.AreEqual(vncEtatCommande.vncLivree, objCmd.etat.codeEtat)
        'And les lignes sont entèrement livrée
        For Each olg As LgCommande In objCmd.colLignes
            Assert.AreEqual(olg.qteCommande, olg.qteLiv)
        Next
    End Sub

    <TestMethod()> Public Sub T_VerifInfosLivraisonsNOK()
        Dim objCmd As CommandeClient
        Dim objLgCmd1 As LgCommande
        Dim objLgCmd2 As LgCommande
        Dim objLgCmd3 As LgCommande
        Dim nidCmd As Integer

        'GIVEN j'ai un produit avec un contenant de 1020 et un conditionnement par 12 
        oCont.poids = 1020
        oCond.valeur = 12
        Assert.IsTrue(oCont.Save)
        Assert.IsTrue(oCond.Save)
        m_oProduit.idContenant = oCont.id
        m_oProduit.idConditionnement = oCond.id
        m_oProduit.save()
        'AND une Commande Client créée
        objCmd = New CommandeClient(m_oClient)


        'WHEN j'ajoute une ligne de 11 produit payant
        objLgCmd1 = objCmd.AjouteLigne("10", m_oProduit, 11, 10.5, False)
        'AND un produit Gratuit
        objLgCmd2 = objCmd.AjouteLigne("20", m_oProduit, 1, 0, True)
        'AND j'ajoute une autre Ligne gratuite
        objLgCmd3 = objCmd.AjouteLigne("30", m_oProduit, 1, 0, True)
        'AND je sauvegarde ma commande
        objCmd.save()
        objCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        objCmd.save()
        Assert.AreEqual(vncEtatCommande.vncValidee, objCmd.etat.codeEtat)
        Assert.AreEqual(CDec(2), objCmd.qteColis)
        'And le Fichier FTP contient  1 Colis
        'Fichier issu de Groussard
        If System.IO.File.Exists("TSTEDI/testGroussard.csv") Then
            System.IO.File.Delete("TSTEDI//testGroussard.csv")
        End If
        Dim strCode As String = objCmd.code
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "Produit;Désignation produit;Date du mouvement;Numéro du document;Total U.C. entrée;Total U.C. sortie;Numéro de commande" & vbCrLf)
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "xxxxxxx;-----xxxxxxxxxxxxxx;03/07/2018;5344;0;1;000" & strCode & vbCrLf)
        Assert.IsTrue(System.IO.File.Exists("TSTEDI//testGroussard.csv"))

        'THEN
        mvtEDI.VerificationCommandes("TSTEDI//testGroussard.csv")
        objCmd = CommandeClient.getListe(strCode, "", vncEtatCommande.vncRien, "")(1)
        objCmd.load()
        'L'état de la Ccommande est restée à l'état Validée
        Assert.AreEqual(vncEtatCommande.vncValidee, objCmd.etat.codeEtat)
        'And les lignes ne sont pas livrées
        For Each olg As LgCommande In objCmd.colLignes
            Assert.AreEqual(0D, olg.qteLiv)
        Next
    End Sub
    ''' <summary>
    ''' test de la vérification des infos de commandes pour 2 commandes
    ''' La premieère est OK pas la seconde
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T_VerifInfosLivraisons2Commandes()
        Dim objCmd As CommandeClient
        Dim objLgCmd1 As LgCommande
        Dim objLgCmd2 As LgCommande
        Dim objLgCmd3 As LgCommande
        Dim nidCmd As Integer
        Dim strCode As String
        Dim strCode2 As String

        'GIVEN j'ai un produit avec un contenant de 1020 et un conditionnement par 12 
        oCont.poids = 1020
        oCond.valeur = 12
        Assert.IsTrue(oCont.Save)
        Assert.IsTrue(oCond.Save)
        m_oProduit.idContenant = oCont.id
        m_oProduit.idConditionnement = oCond.id
        m_oProduit.save()
        'AND une Commande Client créée
        objCmd = New CommandeClient(m_oClient)


        'WHEN j'ajoute une ligne de 11 produit payant
        objLgCmd1 = objCmd.AjouteLigne("10", objProduit1, 11, 10.5, False)
        'AND un produit 
        objLgCmd2 = objCmd.AjouteLigne("20", objProduit2, 1, 10, False)
        'AND j'ajoute une autre Ligne 
        objLgCmd3 = objCmd.AjouteLigne("30", objProduit3, 1, 10, False)
        'AND je sauvegarde ma commande
        objCmd.save()
        objCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        objCmd.save()
        Assert.AreEqual(vncEtatCommande.vncValidee, objCmd.etat.codeEtat)
        Assert.AreEqual(CDec(4), objCmd.qteColis)

        'AND une Commande Client créée
        objCmd = New CommandeClient(m_oClient)


        'WHEN j'ajoute une ligne de 11 produit payant
        objLgCmd1 = objCmd.AjouteLigne("10", objProduit1, 11, 10.5, False)
        'AND un produit Gratuit
        objLgCmd2 = objCmd.AjouteLigne("20", objProduit2, 1, 10, False)
        'AND j'ajoute une autre Ligne gratuite
        objLgCmd3 = objCmd.AjouteLigne("30", objProduit3, 1, 10, False)
        'AND je sauvegarde ma commande
        objCmd.save()
        objCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        objCmd.save()
        Assert.AreEqual(vncEtatCommande.vncValidee, objCmd.etat.codeEtat)
        Assert.AreEqual(CDec(4), objCmd.qteColis)
        strCode = objCmd.code

        'AND une Commande Client créée
        objCmd = New CommandeClient(m_oClient)
        'WHEN j'ajoute une ligne de 11 produit payant
        objLgCmd1 = objCmd.AjouteLigne("10", objProduit1, 11, 10.5, False)
        'AND un produit Gratuit
        objLgCmd2 = objCmd.AjouteLigne("20", objProduit2, 1, 10, False)
        'AND j'ajoute une autre Ligne gratuite
        objLgCmd3 = objCmd.AjouteLigne("30", objProduit3, 1, 10, False)
        'AND je sauvegarde ma commande
        objCmd.save()
        objCmd.changeEtat(vncActionEtatCommande.vncActionValider)
        objCmd.save()
        Assert.AreEqual(vncEtatCommande.vncValidee, objCmd.etat.codeEtat)
        Assert.AreEqual(CDec(4), objCmd.qteColis)
        strCode2 = objCmd.code


        'And le Fichier FTP contient  1 Colis
        'Fichier issu de Groussard
        If System.IO.File.Exists("TSTEDI/testGroussard.csv") Then
            System.IO.File.Delete("TSTEDI//testGroussard.csv")
        End If
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "Produit;Désignation produit;Date du mouvement;Numéro du document;Total U.C. entrée;Total U.C. sortie;Numéro de commande" & vbCrLf)
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "xxxxxxx;-----xxxxxxxxxxxxxx;03/07/2018;5344;0;1;000" & strCode & vbCrLf)
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "xxxxxxx;-----xxxxxxxxxxxxxx;03/07/2018;5344;0;1;000" & strCode & vbCrLf)
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "xxxxxxx;-----xxxxxxxxxxxxxx;03/07/2018;5344;0;0;000" & strCode & vbCrLf)
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "xxxxxxx;-----xxxxxxxxxxxxxx;03/07/2018;5344;0;2;000" & strCode2 & vbCrLf)
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "xxxxxxx;-----xxxxxxxxxxxxxx;03/07/2018;5344;0;1;000" & strCode2 & vbCrLf)
        System.IO.File.AppendAllText("TSTEDI//testGroussard.csv", "xxxxxxx;-----xxxxxxxxxxxxxx;03/07/2018;5344;0;1;000" & strCode2 & vbCrLf)
        Assert.IsTrue(System.IO.File.Exists("TSTEDI//testGroussard.csv"))

        'AND on lance la vérifcation de commande
        mvtEDI.VerificationCommandes("TSTEDI//testGroussard.csv")

        'La Première commande n'est pas Livrée
        objCmd = CommandeClient.getListe(strCode, "", vncEtatCommande.vncRien, "")(1)
        objCmd.load()
        'L'état de la Ccommande est restée à l'état Validée
        Assert.AreEqual(vncEtatCommande.vncValidee, objCmd.etat.codeEtat)
        'And les lignes ne sont pas livrées
        For Each olg As LgCommande In objCmd.colLignes
            Assert.AreEqual(0D, olg.qteLiv)
        Next
        'La Seconde commande 'est Livrée
        objCmd = CommandeClient.getListe(strCode2, "", vncEtatCommande.vncRien, "")(1)
        objCmd.load()
        'L'état de la Ccommande est restée à l'état Validée
        Assert.AreEqual(vncEtatCommande.vncLivree, objCmd.etat.codeEtat)
        'And les lignes ne sont pas livrées
        For Each olg As LgCommande In objCmd.colLignes
            Assert.AreEqual(olg.qteCommande, olg.qteLiv)
        Next
    End Sub

    <TestMethod()> Public Sub T_getParam()
        Dim strTest As String
        strTest = Param.getConstante("CST_FTPEDI_REPLOCAL")
        Assert.AreEqual("TSTEDI", strTest)
    End Sub
    <TestMethod()> Public Sub T_NumCMD()
        Dim omvt As New mvtEDI
        omvt.NumCMD = "00046800"
        Assert.AreEqual("46800", omvt.NumCMD)
    End Sub

End Class