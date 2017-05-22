'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V22
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objCLT As Client
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()

        Dim col As Collection
        Persist.shared_connect()
        Param.LoadcolParams()

        col = Fournisseur.getListe("FRNV22")
        If col.Count > 0 Then
            For Each m_objFRN In col
                m_objFRN.bDeleted = True
                m_objFRN.Save()
            Next
        End If

        col = Produit.getListe(vncEnums.vncTypeProduit.vncTous, "PRDV22")
        If col.Count > 0 Then
            For Each m_objPRD In col
                m_objPRD.bDeleted = True
                m_objPRD.save()
            Next
        End If

        col = Client.getListe("CLTV22")
        If col.Count > 0 Then
            For Each m_objCLT In col
                m_objCLT.bDeleted = True
                m_objCLT.save()
            Next
        End If
        m_objFRN = New Fournisseur("FRNV22", "FRn de'' test")
        m_objFRN.rs = "FRNV22"
        Assert.IsTrue(m_objFRN.Save(), "FRN.Create")

        m_objPRD = New Produit("PRDV22", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objCLT = New Client("CLTV22", "Client de' test")
        m_objCLT.rs = "Client de test"
        Assert.IsTrue(m_objCLT.save(), "Client.Create")

        Persist.shared_disconnect()
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        Persist.shared_connect()
        m_objPRD.bDeleted = True
        Assert.IsTrue(m_objPRD.save(), "TestCleanup")

        m_objFRN.bDeleted = True
        Assert.IsTrue(m_objFRN.Save())

        m_objCLT.bDeleted = True
        Assert.IsTrue(m_objCLT.save())

        Persist.shared_disconnect()
        MyBase.TestCleanup()
    End Sub
    '<Test(), Ignore("obsolete")> Public Sub T10_EXPORT_AVEC_COM2LIGNES()
    '    'Verification que L'export WebEDI supprime bien les retours chariots dans les commantaires de livraison
    '    Dim objCMD As CommandeClient
    '    Dim nFile As Integer
    '    Dim strResult As String
    '    Dim nLineNumber As Integer

    '    'Creation d'une Commande
    '    objCMD = New CommandeClient(m_objCLT)
    '    objCMD.dateCommande = "06/02/2000"
    '    objCMD.CommCommande.comment = "123456789012345678901234567890123456789012345678901234567890"
    '    objCMD.CommFacturation.comment = "123456789012345678901234567890123456789012345678901234567890"
    '    objCMD.CommLibre.comment = "123456789012345678901234567890123456789012345678901234567890"
    '    objCMD.CommLivraison.comment = "12345678901233" + vbCrLf + "45678901234567890123456789012345678901234567890"
    '    objCMD.AjouteLigne("10", m_objPRD, 10, 10)
    '    If System.IO.File.Exists("./adel.txt") Then
    '        System.IO.File.Delete("./adel.txt")
    '    End If
    '    objCMD.exporterWebEDI("./adel.txt")
    '    nFile = FreeFile()
    '    FileOpen(nFile, "./adel.txt", OpenMode.Input, OpenAccess.Read)
    '    nLineNumber = 0
    '    While Not EOF(nFile)
    '        nLineNumber = nLineNumber + 1
    '        strResult = LineInput(nFile)
    '        Console.WriteLine(strResult)
    '    End While

    '    Assert.AreEqual(1, nLineNumber, "Une seule Ligne de fichier")
    '    Assert.AreEqual("010", Mid(strResult, 242, 3), "Le Numéro lde ligne est sur 249 ")

    '    FileClose(nFile)
    '    'Suppression du fichier créé
    '    System.IO.File.Delete("./adel.txt")

    'End Sub

    <TestMethod()> Public Sub T20_SCMD_NOM_FOURNISSEUR()
        'Verification du chargement du fournisseur lors du chargement de la collection des souscommandes
        Dim objCMD As CommandeClient
        Dim objSCMD As SousCommande

        'Creation d'une Commande
        objCMD = New CommandeClient(m_objCLT)
        objCMD.dateCommande = "06/02/2000"
        objCMD.CommCommande.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommFacturation.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLibre.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLivraison.comment = "1234567890123456789012345678" + vbCrLf + "90123456789012345678901234567890"
        objCMD.AjouteLigne("10", m_objPRD, 10, 10)
        Assert.IsTrue(objCMD.save, "Sauvegarde de la commmande")
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD.save, "Sauvegarde de la commmande etat livrée")
        objCMD.generationSousCommande()
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(objCMD.save, "Sauvegarde de la commmande etat Eclatée")
        Assert.IsTrue(objCMD.LoadColSousCommande(), "Lecture des souscommandes")
        Assert.AreEqual(1, objCMD.colSousCommandes.Count, "1 sous-commandes")
        For Each objSCMD In objCMD.colSousCommandes
            Assert.IsTrue(objSCMD.oFournisseur.rs <> "", "RS Fournisseur non vide")
        Next

        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save, "Suppression de la commande")


    End Sub
    <TestMethod()> Public Sub T30_GEN_FACTRTRP_LOADLIGNES()
        'Apres la génération des factures de transport, vérifier le bon rechargement des lignes
        Dim objCMD As CommandeClient
        Dim colCom As New Collection
        Dim colFactTRP As ColEvent
        Dim idFactTRP As Integer
        Dim objFactTRP As FactTRP

        'Creation d'une Commande
        objCMD = New CommandeClient(m_objCLT)
        objCMD.bFactTransport = True
        objCMD.dateCommande = "06/02/2000"
        objCMD.CommCommande.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommFacturation.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLibre.comment = "123456789012345678901234567890123456789012345678901234567890"
        objCMD.CommLivraison.comment = "1234567890123456789012345678" + vbCrLf + "90123456789012345678901234567890"
        objCMD.AjouteLigne("10", m_objPRD, 10, 10)
        Assert.IsTrue(objCMD.save, "Sauvegarde de la commmande")
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        objCMD.dateLivraison = "06/02/1964"
        objCMD.dateEnlevement = "06/02/1964"
        Assert.IsTrue(objCMD.save, "Sauvegarde de la commmande etat livrée")
        objCMD.generationSousCommande()
        objCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        Assert.IsTrue(objCMD.save, "Sauvegarde de la commmande etat Eclatée")

        colCom = New Collection
        colCom.Add(objCMD)

        colFactTRP = FactTRP.createFactTRPs(colCom, "06/02/1964", "06/02/1964", "Période de test")
        objFactTRP = CType(colFactTRP.Item(1), FactTRP)
        'Assert.AreEqual(1, objFactTRP.colLignes.Count, "Avant rechargement La Facture contient 1 ligne")
        objFactTRP.Save()
        idFactTRP = objFactTRP.id
        Assert.AreEqual(1, colFactTRP.Count, "1 facture de transport généré")
        Assert.IsTrue(0 <> idFactTRP, "Id de la facture Non Nul")
        'Console.Out.WriteLine("IdFactTRp=" & idFactTRP)

        objFactTRP = FactTRP.createandload(idFactTRP)
        objFactTRP.loadcolLignes()
        'Assert.AreEqual(1, objFactTRP.colLignes.Count, "Apres rechargement La Facture contient 1 ligne")

        objFactTRP.bDeleted = True
        objFactTRP.Save()

        objCMD.bDeleted = True
        objCMD.save()

    End Sub
    <TestMethod()> Public Sub T40_SCMD_TXCOM_DECIMAL()
        Dim objCMD As CommandeClient
        Dim objSCMD As SousCommande
        Dim objSCMD2 As SousCommande
        Dim nid As Long

        'I - Création d'une Sous-commande 

        objCMD = New CommandeClient(m_objCLT)
        Assert.IsTrue(objCMD.save())

        objSCMD = New SousCommande(objCMD, m_objFRN)
        objSCMD.code = ""
        objSCMD.dateFactFournisseur = "01/08/2004"
        objSCMD.totalHTFacture = 150
        objSCMD.totalTTCFacture = 165
        objSCMD.baseCommission = 150
        objSCMD.tauxCommission = 1.5
        objSCMD.CommCommande.comment = "CommCommande"
        objSCMD.CommLivraison.comment = "CommLivriaon"
        objSCMD.CommFacturation.comment = "CommFact"
        objSCMD.CommLibre.comment = "Libre"
        'Save
        Assert.IsTrue(objSCMD.Save(), "Insert" & objSCMD.getErreur)
        nid = objSCMD.id
        'II - Rechargement d'une Sous - Commande
        '========================================
        objSCMD2 = SousCommande.createandload(nid)
        Assert.AreEqual(CDec(1.5), objSCMD2.tauxCommission, "Tax de comm Différents")
        Assert.IsTrue(objSCMD.Equals(objSCMD2), "Différents")

        'Suppression
        objSCMD2.bDeleted = True
        Assert.IsTrue(objSCMD2.Save(), "Delet SCMD" & Persist.getErreur)

        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save(), "Delete CMD" & Persist.getErreur)
    End Sub
    '<TestMethod(), Ignore("obsolete")> Public Sub T50_EXPORT_AVEC_FICHIER_EXISTANT()
    '    'Verification que L'export WebEDI supprime bien le fichier généré
    '    Dim objCMD As CommandeClient
    '    Dim nFile As Integer
    '    Dim strResult As String
    '    Dim nLineNumber As Integer

    '    'Creation d'une Commande
    '    objCMD = New CommandeClient(m_objCLT)
    '    objCMD.dateCommande = "06/02/2000"
    '    objCMD.CommCommande.comment = "123456789012345678901234567890123456789012345678901234567890"
    '    objCMD.CommFacturation.comment = "123456789012345678901234567890123456789012345678901234567890"
    '    objCMD.CommLibre.comment = "123456789012345678901234567890123456789012345678901234567890"
    '    objCMD.CommLivraison.comment = "12345678901233" + vbCrLf + "45678901234567890123456789012345678901234567890"
    '    objCMD.AjouteLigne("10", m_objPRD, 10, 10)
    '    If System.IO.File.Exists("./adel.txt") Then
    '        System.IO.File.Delete("./adel.txt")
    '    End If
    '    'Création du fichier avec 2 lignes blanches
    '    nFile = FreeFile()
    '    FileOpen(nFile, "./adel.txt", OpenMode.Output, OpenAccess.Write)
    '    WriteLine(nFile, "First Line")
    '    WriteLine(nFile, "Second Line")
    '    FileClose(nFile)

    '    objCMD.exporterWebEDI("./adel.txt")
    '    nFile = FreeFile()
    '    FileOpen(nFile, "./adel.txt", OpenMode.Input, OpenAccess.Read)
    '    nLineNumber = 0
    '    While Not EOF(nFile)
    '        nLineNumber = nLineNumber + 1
    '        strResult = LineInput(nFile)
    '        Console.WriteLine(strResult)
    '    End While

    '    Assert.AreEqual(1, nLineNumber, "Une seule Ligne de fichier")
    '    Assert.AreEqual("010", Mid(strResult, 242, 3), "Le Numéro lde ligne est sur 249 ")

    '    FileClose(nFile)
    '    'Suppression du fichier créé
    '    System.IO.File.Delete("./adel.txt")

    'End Sub

End Class


