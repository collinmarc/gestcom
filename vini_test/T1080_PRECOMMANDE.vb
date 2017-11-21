'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1080_PRECOMMANDE
    Inherits test_Base

    Private m_objPreCommande As preCommande
    Private m_objCLT As Client
    Private m_oPrd1 As Produit
    Private m_oPrd2 As Produit
    Private m_oPrd3 As Produit
    Private m_oFRN1 As Fournisseur
    Private m_oFRN2 As Fournisseur
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()

        'Creation de 2 Fournisseurs
        m_oFRN1 = New Fournisseur("T60FRN1TEST" & Now(), "")
        Assert.IsTrue(m_oFRN1.Save(), "oFRN1.save")
        m_oFRN2 = New Fournisseur("T60FRN2TEST" & Now(), "")
        Assert.IsTrue(m_oFRN2.Save(), "oFRN2.save")

        'Creation de 3 Produits
        m_oPrd1 = New Produit("T60PRD1" & Now(), m_oFRN1, 97)
        Assert.IsTrue(m_oPrd1.save(), "oPRD1.save")
        Assert.IsTrue(m_oPrd1.id <> 0, "oPRD1<>0")
        m_oPrd2 = New Produit("T60PRD2" & Now(), m_oFRN2, 98)
        Assert.IsTrue(m_oPrd2.save(), "oPRD2.save")
        m_oPrd3 = New Produit("T60PRD3" & Now(), m_oFRN2, 99)
        Assert.IsTrue(m_oPrd3.save(), "oPRD3.save")

        'Creation d'un client
        m_objCLT = New Client("T60CLT1" & Now(), "T60")
        Assert.IsTrue(m_objCLT.save(), "oCLT.Insert" & Client.getErreur)
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objLgPRecom As lgPrecomm
        Dim objLgPRecom1 As lgPrecomm
        Dim objLgPRecom2 As lgPrecomm
        Dim sCode As String


        m_objPreCommande = preCommande.createandload(m_objCLT.id)
        Debug.Assert(m_objPreCommande.id = m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objPreCommande.getlgPrecomCount, 0, "Collection non vide")

        'Ajout de 2 ligne de Precommande
        objLgPRecom1 = m_objPreCommande.ajouteLgPrecom(m_oPrd1, 12, 20, 11, "31/07/1964")
        Assert.IsTrue(Not objLgPRecom1 Is Nothing, "Ajout OPRD1")
        objLgPRecom2 = m_objPreCommande.ajouteLgPrecom(m_oPrd3, 12, 20, 11)
        Assert.IsTrue(Not objLgPRecom2 Is Nothing, "Ajout OPRD3")
        'L'ajout de la 3eme rend False car il s'agit du même code Produit
        Assert.IsNull(m_objPreCommande.ajouteLgPrecom(m_oPrd3.id, m_oPrd3.code, m_oPrd3.nom, 12, 20, 11), "Ajout OPRD3bis")
        'Sauvegarde de la precommande
        Assert.IsTrue(m_objPreCommande.save(), "Oclt.Save")

        'Rechargement de la precommande
        m_objPreCommande = preCommande.createandload(m_objCLT.id)
        Debug.Assert(m_objPreCommande.id = m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objPreCommande.getlgPrecomCount, 2, "Precommande.count ")
        'controle du contenu de daate d derniere commande
        objLgPRecom = m_objPreCommande.colLignes(1)
        Assert.IsTrue(objLgPRecom.Equals(objLgPRecom1), "LgPrecom1 sont différents apres rechargement")
        objLgPRecom = m_objPreCommande.colLignes(2)
        Assert.IsTrue(objLgPRecom.Equals(objLgPRecom2), "LgPrecom2 sont différents apres rechargement")

        'Ajout d'une ligne de precommande
        Assert.IsTrue(Not m_objPreCommande.ajouteLgPrecom(m_oPrd2.id, m_oPrd2.code, m_oPrd2.nom, 113, 20, 11.0) Is Nothing, "oclt.AjouteLg 3")
        'Sauvegarde de la precommande
        Assert.IsTrue(m_objPreCommande.save(), "Oclt.Save")

        'Rechargement de la precommande
        m_objPreCommande = preCommande.createandload(m_objCLT.id)
        Debug.Assert(m_objPreCommande.id = m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objPreCommande.getlgPrecomCount, 3, "Precommande.count ")

        'Suppression d'une ligne de la precommande
        m_objPreCommande.supprimeLgPrecom(3)
        'Sauvegarde de la precommande
        Assert.IsTrue(m_objPreCommande.save(), "Oclt.Save")

        'Rechargement de la precommande
        m_objPreCommande = preCommande.createandload(m_objCLT.id)
        Debug.Assert(m_objPreCommande.id = m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objPreCommande.getlgPrecomCount, 2, "Precommande.count ")

        'Maj d'une ligne de la precommande
        objLgPRecom = m_objPreCommande.colLignes(1)
        Assert.IsFalse(objLgPRecom Is Nothing)
        sCode = objLgPRecom.codeProduit
        objLgPRecom.qteHab = 150
        objLgPRecom.qteDern = 5
        objLgPRecom.prixU = 33

        'Sauvegarde de la precommande
        Assert.IsTrue(m_objPreCommande.save(), "Oclt.Save")

        'Rechargement de la precommande
        m_objPreCommande = preCommande.createandload(m_objCLT.id)
        Debug.Assert(m_objPreCommande.id = m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objPreCommande.getlgPrecomCount, 2, "Precommande.count ")
        objLgPRecom = m_objPreCommande.colLignes(1)
        Assert.AreEqual(CDec(150), objLgPRecom.qteHab)
        Assert.AreEqual(CDec(5), objLgPRecom.qteDern)
        Assert.AreEqual(CDbl(33), objLgPRecom.prixU)

    End Sub
    <TestMethod()> Public Sub T20_InterfaceClient()
        Dim objLgPRecom As lgPrecomm
        Dim objLgPRecom1 As lgPrecomm
        Dim objLgPRecom2 As lgPrecomm
        Dim sCode As String


        m_objPreCommande = m_objCLT.oPReCommande
        Assert.AreEqual(m_objPreCommande.id, m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objPreCommande.getlgPrecomCount, 0, "Collection non vide")

        'Ajout de 2 ligne de Precommande
        objLgPRecom1 = m_objCLT.ajouteLgPrecom(m_oPrd1, 12, 20, 11, "31/07/1964")
        Assert.IsTrue(Not objLgPRecom1 Is Nothing, "Ajout OPRD1")
        objLgPRecom2 = m_objCLT.ajouteLgPrecom(m_oPrd3, 12, 20, 11)
        Assert.IsTrue(Not objLgPRecom2 Is Nothing, "Ajout OPRD3")
        'L'ajout de la 3eme rend False car il s'agit du même code Produit
        Assert.IsNull(m_objCLT.ajouteLgPrecom(m_oPrd3.id, m_oPrd3.code, m_oPrd3.nom, 12, 20, 11), "Ajout OPRD3bis")
        'Sauvegarde de la precommande
        Assert.IsTrue(m_objCLT.save(), "Oclt.Save")

        'Rechargement de la precommande
        m_objCLT = Client.createandload(m_objCLT.id)
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(m_objCLT.getlgPrecomCount, 2, "Precommande.count ")
        'controle du contenu de daate d derniere commande
        objLgPRecom = m_objCLT.oPrecommande.colLignes(1)
        Assert.IsTrue(objLgPRecom.Equals(objLgPRecom1), "LgPrecom1 sont différents apres rechargement")
        objLgPRecom = m_objCLT.oPrecommande.colLignes(2)
        Assert.IsTrue(objLgPRecom.Equals(objLgPRecom2), "LgPrecom2 sont différents apres rechargement")

        'Ajout d'une ligne de precommande
        Assert.IsTrue(Not m_objCLT.ajouteLgPrecom(m_oPrd2.id, m_oPrd2.code, m_oPrd2.nom, 113, 20, 11.0) Is Nothing, "oclt.AjouteLg 3")
        'Sauvegarde de la precommande
        Assert.IsTrue(m_objCLT.save(), "Oclt.Save")

        'Rechargement de la precommande
        m_objCLT = Client.createandload(m_objCLT.id)
        Debug.Assert(m_objPreCommande.id = m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objCLT.oPrecommande.getlgPrecomCount, 3, "Precommande.count ")

        'Suppression d'une ligne de la precommande
        m_objCLT.oPrecommande.supprimeLgPrecom(3)
        'Sauvegarde de la precommande
        Assert.IsTrue(m_objCLT.save(), "Oclt.Save")

        'Rechargement de la precommande
        m_objCLT = Client.createandload(m_objCLT.id)
        Debug.Assert(m_objPreCommande.id = m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objCLT.getlgPrecomCount, 2, "Precommande.count ")

        'Maj d'une ligne de la precommande
        objLgPRecom = m_objCLT.oPrecommande.colLignes(1)
        sCode = objLgPRecom.codeProduit
        objLgPRecom.qteHab = 150
        objLgPRecom.qteDern = 5
        objLgPRecom.prixU = 33

        'Sauvegarde de la precommande
        Assert.IsTrue(m_objCLT.save(), "Oclt.Save")

        'Rechargement de la precommande
        m_objCLT = Client.createandload(m_objCLT.id)
        Debug.Assert(m_objCLT.oPrecommande.id = m_objCLT.id, "Les Ids de precommande et de Client doit être égaux")
        Assert.AreEqual(m_objCLT.getlgPrecomCount, 2, "Precommande.count ")
        objLgPRecom = m_objCLT.oPrecommande.colLignes(1)
        Assert.AreEqual(CDec(150), objLgPRecom.qteHab)
        Assert.AreEqual(CDec(5), objLgPRecom.qteDern)
        Assert.AreEqual(CDbl(33), objLgPRecom.prixU)

    End Sub

    <TestMethod()> Public Sub T30_ReinitPrecommande()
        Dim objCommande As CommandeClient
        Dim nid As Integer

        objCommande = New CommandeClient(m_objCLT)
        objCommande.AjouteLigne("10", m_oPrd1, 1, 1)
        Assert.IsTrue(objCommande.save, "Sauvegarde Commande")


        'Controle de la precommande du client
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(1, m_objCLT.getlgPrecomCount(), "1 ligne en Precommande")

        'Ajout d'une ligne de precommande
        m_objCLT.ajouteLgPrecom(m_oPrd2, 10, 2)
        'Sauvegarde du client
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde Client")
        nid = m_objCLT.id
        'Rechargement du client
        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        'Controle de la precommande du client
        Assert.AreEqual(2, m_objCLT.getlgPrecomCount(), "2 lignes en Precommande")


        'Reinitialisation du client
        Assert.IsTrue(m_objCLT.reinitPrecommande(), "Reinit Precommande")
        'Sauvegarde du client
        Assert.IsTrue(m_objCLT.save(), "Sauvegarde Client")

        'Rechargement du client
        m_objCLT = Client.createandload(nid)
        m_objCLT.LoadPreCommande()
        'Controle de la precommande du client
        Assert.AreEqual(1, m_objCLT.getlgPrecomCount(), "1 ligne en Precommande")

    End Sub
    <TestMethod()> Public Sub T50_TOCSV()
        Dim nfile As Integer
        Dim nLineNumber As Integer
        Dim strResult As String
        Dim tabCSV As String() = Nothing
        Dim n As Integer

        'Ajout de 2 Lignes à la commande

        m_objCLT.ajouteLgPrecom(m_oPrd1, 10, 15, 15.55, "31/07/1964", "CMD001")
        m_objCLT.ajouteLgPrecom(m_oPrd2, 20, 25, 25.55, "31/08/1964", "CMD002")
        Assert.IsTrue(m_objCLT.save)



        nfile = FreeFile()
        FileOpen(nfile, "./adel.txt", OpenMode.Output, OpenAccess.Write, OpenShare.Shared)
        strResult = m_objCLT.oPrecommande.toCSV()
        Print(nfile, strResult)
        FileClose(nfile)

        nfile = FreeFile()
        FileOpen(nfile, "./adel.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        nLineNumber = 0
        While Not EOF(nfile)
            nLineNumber = nLineNumber + 1
            strResult = LineInput(nfile)
            Console.WriteLine(strResult)
            If nLineNumber = 1 Then
                tabCSV = strResult.Split(";")
                n = 0
                Assert.AreEqual(Trim(m_objCLT.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oPrd1.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual("15", tabCSV(n))
                n = n + 1
                Assert.AreEqual("15.55", tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(CDate("31/07/1964"), "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual("CMD001", tabCSV(n))

            End If
            If nLineNumber = 2 Then
                tabCSV = strResult.Split(";")
                n = 0
                Assert.AreEqual(Trim(m_objCLT.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(m_oPrd2.code), tabCSV(n))
                n = n + 1
                Assert.AreEqual("25", tabCSV(n))
                n = n + 1
                Assert.AreEqual("25.55", tabCSV(n))
                n = n + 1
                Assert.AreEqual(Trim(Format(CDate("31/08/1964"), "ddMMyyyy")), tabCSV(n))
                n = n + 1
                Assert.AreEqual("CMD002", tabCSV(n))

            End If

        End While
        Assert.AreEqual(2, nLineNumber)
        FileClose(nfile)

    End Sub
    ''' <summary>
    ''' Test de la purge des précommandes
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T60_Purge()

        'Ajout de 2 Lignes à la commande

        m_objCLT.ajouteLgPrecom(m_oPrd1, 10, 15, 15.55, "31/07/1964", "CMD001")
        m_objCLT.ajouteLgPrecom(m_oPrd2, 20, 25, 25.55, "31/08/1964", "CMD002")
        m_objCLT.ajouteLgPrecom(m_oPrd3, 20, 25, 25.55, "01/03/2016", "CMD003")
        Assert.IsTrue(m_objCLT.save)
        m_objCLT.LoadPreCommande()

        Assert.AreEqual(3, m_objCLT.getlgPrecomCount())
        Assert.IsTrue(preCommande.Purge("01/08/1964"))
        m_objCLT.oPrecommande.colLignes.clear()
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(2, m_objCLT.getlgPrecomCount())

        Assert.IsTrue(lgPrecomm.PurgeLgPrecommande("01/08/1965"))

        m_objCLT.oPrecommande.colLignes.clear()
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(1, m_objCLT.getlgPrecomCount())

        Assert.IsTrue(lgPrecomm.PurgeLgPrecommande("03/01/2016"))
        m_objCLT.oPrecommande.colLignes.clear()
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(1, m_objCLT.getlgPrecomCount())

        Assert.IsTrue(lgPrecomm.PurgeLgPrecommande("02/03/2016"))
        m_objCLT.oPrecommande.colLignes.clear()
        m_objCLT.LoadPreCommande()
        Assert.AreEqual(0, m_objCLT.getlgPrecomCount())

    End Sub
End Class



