'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1060_Produit
    Inherits test_Base

    Private m_obj As Produit
    Private m_idPrd As Integer
    Private m_oFRN As Fournisseur
    Private m_idFRN As Integer

    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()

        '            m_obj.shared_connect()
        m_oFRN = New Fournisseur("T1060PRD", "Test Produit")
        m_oFRN.Save()
        m_idFRN = m_oFRN.id
        m_obj = New Produit("", m_oFRN, 1990)
        m_obj.save()
        m_idPrd = m_obj.id
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        '           m_obj.shared_disconnect()

        m_obj = Produit.createandload(m_idPrd)
        If (Not m_obj Is Nothing) Then
            m_obj.bDeleted = True
            m_obj.save()
        End If
        m_oFRN = Fournisseur.createandload("T1060PRD")
        If (Not m_oFRN Is Nothing) Then

            m_oFRN.bDeleted = True
            m_oFRN.Save()
        End If

        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objPRD As Produit

        'Test des Valeur par defaut
        Assert.AreEqual(m_obj.idConditionnement, Param.conditionnementdefaut.id)
        Assert.AreEqual(m_obj.idContenant, contenant.contenantDefaut.id)
        Assert.AreEqual(m_obj.idCouleur, Param.couleurdefaut.id)
        Assert.AreEqual(m_obj.idRegion, Param.regiondefaut.id)
        Assert.AreEqual(m_obj.idTVA, Param.TVAdefaut.id)
        m_obj.code = "CODE"
        m_obj.bDisponible = True
        m_obj.codeStat = "CODESTAT"
        m_obj.DateDernInventaire = Now()
        m_obj.millesime = 98
        m_obj.motcle = "MTCLE"
        m_obj.QteStock = 10
        m_obj.QteStockDernInventaire = 150
        m_obj.idCouleur = Param.colCouleur(Param.colCouleur.Count).id


        Assert.IsTrue(m_obj.Equals(m_obj), "Egal à Lui même")
        objPRD = New Produit("", m_oFRN, 1990)
        objPRD.code = "CODE"
        objPRD.bDisponible = True
        objPRD.codeStat = "CODESTAT"
        objPRD.DateDernInventaire = m_obj.DateDernInventaire
        objPRD.millesime = 98
        objPRD.motcle = "MTCLE"
        objPRD.QteStock = 10
        objPRD.QteStockDernInventaire = 150
        objPRD.idCouleur = Param.colCouleur(Param.colCouleur.Count).id


        ' Test non valide car l'obhet est sauvegardé dans le TestInitialize
        'Assert.IsTrue(m_obj.Equals(objPRD), "Egal à un semblable")

        objPRD.motcle = "MOTCLE2"
        Assert.IsFalse(m_obj.Equals(objPRD), "Egal à un Différent")
        Dim obj As Object
        Assert.IsFalse(m_obj.Equals(obj), "Egal autrecjhose")


    End Sub
    <TestMethod()> Public Sub T15_DB()
        Dim objPRD As Produit
        Dim objPRD2 As Produit
        Dim n As Integer

        'I - Création d'un Produit
        '=========================
        objPRD = New Produit("", New Fournisseur, 1990)
        Assert.AreEqual(0, objPRD.TarifA, "Tarif A")
        Assert.AreEqual(0, objPRD.TarifB, "Tarif B")
        Assert.AreEqual(0, objPRD.TarifC, "Tarif B")
        objPRD.code = "FTEST" & Now()
        objPRD.nom = "Produit de test'fklfdkl&é#'(-è_çà)=}}]@^\`|[{#~"
        objPRD.idFournisseur = Fournisseur.getListe()(1).id
        objPRD.bDisponible = False
        objPRD.bStock = False
        objPRD.codeStat = "CODESTAT"
        objPRD.DateDernInventaire = Now()
        objPRD.millesime = 98
        objPRD.motcle = "MOTCLE"
        objPRD.QteStock = 14
        objPRD.QteStockDernInventaire = 28
        objPRD.idCouleur = Param.colCouleur(Param.colCouleur.Count).id
        objPRD.TarifA = 11.5
        objPRD.TarifB = 12.5
        objPRD.TarifC = 13.5
        Assert.AreEqual(11.5, CDec(objPRD.TarifA), " TarifA")
        Assert.AreEqual(12.5, objPRD.TarifB, " TarifB")
        Assert.AreEqual(13.5, objPRD.TarifC, "TarifC")


        'Test des indicateurs Avant le Save
        Assert.IsTrue(objPRD.bNew)
        Assert.IsTrue(objPRD.bUpdated)
        Assert.IsFalse(objPRD.bDeleted)
        'Save
        Assert.IsTrue(objPRD.save(), "Insert" & objPRD.getErreur)
        Assert.IsTrue((objPRD.id <> 0), "Id Apres le Save doit être différent de 0")
        'Test des indicateurs Après le Save
        Assert.IsFalse(objPRD.bNew, "bNew apres insert")
        Assert.IsFalse(objPRD.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objPRD.bDeleted, "bDeleted apres insert")

        'II - Rechargement d'un Produit
        '=========================
        n = objPRD.id
        objPRD2 = Produit.createandload(n)
        Assert.AreEqual(CDec(11.5), objPRD2.TarifA, "Load TarifA")
        Assert.AreEqual(CDec(12.5), objPRD2.TarifB, "Load TarifB")
        Assert.AreEqual(CDec(13.5), objPRD2.TarifC, "Load TarifC")
        Assert.IsTrue(objPRD.Equals(objPRD2))

        'III - Modification du Produit
        '=================================
        ' Modification du Produit
        objPRD2.nom = objPRD.nom & "2"
        objPRD2.bStock = True
        objPRD2.QteStock = 30
        objPRD2.idCouleur = Param.colCouleur(Param.colCouleur.Count - 1).id
        objPRD2.TarifA = 15.15
        objPRD2.TarifB = 16.16
        objPRD2.TarifC = 17.17

        'Test des indicateurs Avant le Save
        Assert.IsFalse(objPRD2.bNew)
        Assert.IsTrue(objPRD2.bUpdated)
        Assert.IsFalse(objPRD2.bDeleted)
        'Save
        Assert.IsTrue(objPRD2.save(), "Update" & objPRD.getErreur)
        'Test des indicateurs Après le Save
        Assert.IsFalse(objPRD2.bNew, "bNew apres insert")
        Assert.IsFalse(objPRD2.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objPRD2.bDeleted, "bDeleted apres insert")

        'Rechargement de l'objet
        n = objPRD2.id
        objPRD = Produit.createandload(n)
        Assert.AreEqual(15.15, objPRD.TarifA, "Load TarifA")
        Assert.AreEqual(16.16, objPRD.TarifB, "Load TarifB")
        Assert.AreEqual(17.17, objPRD.TarifC, "Load TarifC")
        Assert.IsTrue(objPRD.Equals(objPRD2))

        'IV - Suppression du Produit
        '=================================
        ' Modification du Produit
        objPRD.bDeleted = True
        Assert.IsTrue(objPRD.save(), "Delete" & objPRD.getErreur())
    End Sub
    <TestMethod()> Public Sub T51_ListeCriteres()
        Dim colPRD As Collection
        Dim oPRD As Produit
        Dim objPRD2 As Produit
        Dim objFRN As Fournisseur
        Dim objCLT As Client

        Persist.shared_connect()
        objFRN = New Fournisseur("FPRDT51" & Now(), "Mon Fournisseur")
        Assert.IsTrue(objFRN.Save())

        objPRD2 = New Produit("PRDT51", objFRN, 2004)
        objPRD2.nom = "MYPRODuct"
        objPRD2.motcle = "c'est"
        Assert.IsTrue(objPRD2.save())

        objCLT = New Client("CLPRDTT51" & Now(), "Monclient")
        Assert.IsTrue(objCLT.save())
        objCLT.ajouteLgPrecom(objPRD2, 10)
        Assert.IsTrue(objCLT.save())
        Persist.shared_disconnect()

        'I - Liste Simple
        '================
        Persist.shared_connect()
        colPRD = Produit.getListe(vncTypeProduit.vncTous, )
        Persist.shared_disconnect()
        Assert.IsTrue(colPRD.Count > 0, "Col.count > 0" & Persist.getErreur)
        For Each oPRD In colPRD
            Assert.IsTrue(oPRD.id <> 0, "ID <> 0")
            Assert.IsTrue(oPRD.bResume, "bResume = True")
            Persist.shared_connect()
            '               oPRD = Produit.createandload(oPRD.id)
        Next oPRD
        Persist.shared_disconnect()

        'II- Liste sur le code 
        '=============================
        'Console.Out.WriteLine("Liste sur le code")
        'a) Caractère générique
        Persist.shared_connect()
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncTous, "PRD%")
        Assert.IsTrue(colPRD.Count > 0, "Col.count > 0")
        For Each oPRD In colPRD
            Assert.AreEqual(Left(oPRD.code, 3), "PRD", "Mauvais Code " & oPRD.code)
        Next oPRD
        'b) sans Caractère générique
        colPRD = Produit.getListe(vncEnums.vncTypeProduit.vncTous, "PRDT51")
        Assert.IsTrue(colPRD.Count > 0, "Col.count > 0" & Produit.getErreur)
        For Each oPRD In colPRD
            Assert.AreEqual(oPRD.code, "PRDT51", "Mauvais Code " & oPRD.code)
        Next oPRD
        Persist.shared_disconnect()

        'III- Liste sur le nom
        '=============================
        Persist.shared_connect()
        'Console.Out.WriteLine("Liste sur le nom")
        'a) Caractère générique
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , "MYPR%")
        Assert.IsTrue(colPRD.Count > 0, "Nom1:Col.count > 0" & Produit.getErreur)
        For Each oPRD In colPRD
            Assert.AreEqual(Left(oPRD.nom, 4), "MYPR")
        Next oPRD
        'b) sans Caractère générique
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , "MYPRODuct")
        Assert.IsTrue(colPRD.Count > 0, "Nom2:Col.count > 0")
        For Each oPRD In colPRD
            Assert.AreEqual(oPRD.nom, "MYPRODuct")
        Next oPRD
        Persist.shared_disconnect()

        'IV- Liste sur le Mot clé
        '=============================
        'Console.Out.WriteLine("Liste sur le motcle")
        'a) Caractère générique
        Persist.shared_connect()
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , , "c'e%")
        Assert.IsTrue(colPRD.Count > 0, "MotClé1: Col.count > 0" & Produit.getErreur)
        For Each oPRD In colPRD
            objPRD2 = Produit.createandload(oPRD.id)
            Assert.AreEqual(Left(objPRD2.motcle, 3), "c'e")
        Next oPRD
        'b) sans Caractère générique
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , , "c'est")
        Assert.IsTrue(colPRD.Count > 0, "Motclé2: Col.count > 0")
        For Each oPRD In colPRD
            objPRD2 = Produit.createandload(oPRD.id)
            Assert.AreEqual(objPRD2.motcle, "c'est")
        Next oPRD
        Persist.shared_disconnect()

        'V - Fournisseur
        '====================
        'Console.Out.WriteLine("Liste sur le Fourn")
        'a) Avec Résultat
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , , , objFRN.id)
        Assert.IsTrue(colPRD.Count > 0, "Col.count > 0")
        For Each oPRD In colPRD
            objPRD2 = Produit.createandload(oPRD.id)
            Assert.AreEqual(objPRD2.idFournisseur, objFRN.id)
        Next oPRD
        'b) sans résultat
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , , , 99999)
        Assert.IsTrue(colPRD.Count = 0, "Col.count = 0")

        'V - Client
        '====================
        'Console.Out.WriteLine("Liste sur le Client")
        'a) Avec Résultat
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , , , , objCLT.id)
        'Assert.IsTrue(colPRD.Count > 0, "Col.count > 0")
        For Each oPRD In colPRD
            objPRD2 = Produit.createandload(oPRD.id)
            'Console.Out.WriteLine(oPRD.toString())
        Next oPRD
        'b) sans résultat
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , , , , 334532)
        Assert.IsTrue(colPRD.Count = 0, "Col.count = 0")
        'VI - Critères Croisés
        '====================
        'a) Avec Résultat
        colPRD = Produit.getListe(vncTypeProduit.vncTous, "PRD%", "", "c'est%")
        'Assert.IsTrue(colPRD.Count > 0, "Col.count > 0")
        For Each oPRD In colPRD
            objPRD2 = Produit.createandload(oPRD.id)
            Assert.AreEqual(Left(oPRD.code, 3), "PRD")
            'Assert.AreEqual(Left(oPRD.libelle, 3) = "TEST", "Mauvais Nom " & oPRD.libelle)
            Assert.IsTrue(objPRD2.load(oPRD.id), "Load False id=" & oPRD.id)
            Assert.AreEqual(Left(objPRD2.motcle, 3), "c'e")
        Next oPRD

        'b) sans résultat
        colPRD = Produit.getListe(vncTypeProduit.vncTous, "AQWZSX%", "BIB%", "BERTI%")
        Assert.IsTrue(colPRD.Count = 0, "Col.count = 0")


        objCLT.bDeleted = True
        Assert.IsTrue(objCLT.save())

        objPRD2.bDeleted = True
        Assert.IsTrue(objPRD2.save())

        objFRN.bDeleted = True
        Assert.IsTrue(objFRN.Save())


    End Sub
    <TestMethod()> Public Sub T60_ListePrecommande()
        Dim objP1 As Produit
        Dim objP2 As Produit
        Dim objP3 As Produit
        Dim objP4 As Produit
        Dim objFRN As Fournisseur
        Dim objCLT As Client
        Dim colPRD As Collection


        objFRN = New Fournisseur("FRNTEST", "Fournisseur de Test")
        Assert.IsTrue(objFRN.Save())

        objP1 = New Produit("P1", objFRN, 2004)
        objP1.nom = "AAAAA"
        objP1.motcle = "CLE1"

        Assert.IsTrue(objP1.save())
        objP2 = New Produit("P2", objFRN, 2004)
        objP2.nom = "BBBB"
        objP2.motcle = "CLE1"
        Assert.IsTrue(objP2.save())
        objP3 = New Produit("P3", objFRN, 2004)
        objP3.nom = "CCCC"
        objP3.motcle = "CLE2"
        Assert.IsTrue(objP3.save())
        objP4 = New Produit("P4", objFRN, 2004)
        objP4.nom = "DDDD"
        Assert.IsTrue(objP4.save())

        objCLT = New Client("CLTEST", "Client Test")
        Assert.IsTrue(objCLT.save())

        'Ajout des produits dans la Precommande du client
        objCLT.ajouteLgPrecom(objP1)
        objCLT.ajouteLgPrecom(objP2)
        objCLT.ajouteLgPrecom(objP3)
        Assert.IsTrue(objCLT.save())

        'Recherche des produit de la preccommande
        colPRD = Produit.getListe(vncTypeProduit.vncTous, , , , , objCLT.id)
        Assert.AreEqual(3, colPRD.Count, "getListe(id)")
        Assert.IsTrue(Not colPRD("P1") Is Nothing, "P1=")
        Assert.IsTrue(Not colPRD("P2") Is Nothing, "P2=")
        Assert.IsTrue(Not colPRD("P3") Is Nothing, "P3=")

        colPRD = Produit.getListe(vncTypeProduit.vncTous, "P%", , , , objCLT.id)
        Assert.AreEqual(3, colPRD.Count, "getListe(Code,,id)")
        Assert.IsTrue(Not colPRD("P1") Is Nothing, "P1=")
        Assert.IsTrue(Not colPRD("P2") Is Nothing, "P2=")
        Assert.IsTrue(Not colPRD("P3") Is Nothing, "P3=")

        colPRD = Produit.getListe(vncTypeProduit.vncTous, , "A%", , , objCLT.id)
        Assert.AreEqual(1, colPRD.Count, "getListe(,nom,,id)")
        Assert.IsTrue(Not colPRD("P1") Is Nothing, "P1=")

        colPRD = Produit.getListe(vncTypeProduit.vncTous, , , "CLE1", , objCLT.id)
        Assert.AreEqual(2, colPRD.Count, "getListe(,,motCle,,id)")
        Assert.IsTrue(Not colPRD("P1") Is Nothing, "P1=")
        Assert.IsTrue(Not colPRD("P2") Is Nothing, "P2=")

        colPRD = Produit.getListe(vncTypeProduit.vncTous, "P1%", , "CLE1", , objCLT.id)
        Assert.AreEqual(1, colPRD.Count, "getListe(,,motCle,,id)")
        Assert.IsTrue(Not colPRD("P1") Is Nothing, "P1=")

        colPRD = Produit.getListe(vncTypeProduit.vncTous, , "B%", "CLE1", , objCLT.id)
        Assert.AreEqual(1, colPRD.Count, "getListe(,,motCle,,id)")
        Assert.IsTrue(Not colPRD("P2") Is Nothing, "P2=")

        objCLT.bDeleted = True
        Assert.IsTrue(objCLT.save())

        objP1.bDeleted = True
        Assert.IsTrue(objP1.save)
        objP2.bDeleted = True
        Assert.IsTrue(objP2.save)
        objP3.bDeleted = True
        Assert.IsTrue(objP3.save)
        objP4.bDeleted = True
        Assert.IsTrue(objP4.save)

        objFRN.bDeleted = True
        Assert.IsTrue(objFRN.Save)

    End Sub
    'Test la sauvegarde des champs long
    <TestMethod()> Public Sub T60_Champslongs()

        Dim obj As Produit
        obj = New Produit("T60", m_oFRN, 1990)
        obj.nom = "TEST".PadRight(500, "x")

        Assert.IsTrue(obj.Save())
        obj.motcle = "TEST2".PadRight(500, "y")
        Assert.IsTrue(obj.Save())
        obj.load()
        Assert.AreEqual(50, obj.nom.Length)
        Assert.AreEqual(50, obj.motcle.Length)

        obj.bDeleted = True
        obj.Save()

    End Sub
End Class



