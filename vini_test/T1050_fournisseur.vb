'<<TestMethod()>()> de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1050_fournisseur
    Inherits test_Base

    Private m_obj As Fournisseur
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        'm_obj.shared_connect()
        m_obj = New Fournisseur("", "")
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        MyBase.TestCleanup()
        'm_obj.shared_disconnect()
    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objFRN As Fournisseur

        Assert.IsTrue(m_obj.bExportInternet = 0, "Par defaut l'exportinternet est à false")

        m_obj.code = "CODE"
        m_obj.nom = "MonFournisseur"
        m_obj.rs = "RS"
        m_obj.banque = "BANQUE"
        m_obj.rib1 = "RIB1"
        m_obj.rib2 = "RIB2"
        m_obj.rib3 = "RIB3"
        m_obj.rib4 = "RIB4"
        m_obj.siret = "1234567890"
        m_obj.numtvaintracom = "0987654321"
        m_obj.AdresseLivraison.nom = "Marc Collin"
        m_obj.AdresseLivraison.rue1 = "La Mettrie"
        m_obj.AdresseLivraison.rue2 = "2eme Etage"
        m_obj.AdresseLivraison.cp = "35250"
        m_obj.AdresseLivraison.ville = "chasné sur illet"
        m_obj.AdresseLivraison.tel = "0299555299"
        m_obj.AdresseLivraison.fax = "0299555277"
        m_obj.AdresseLivraison.port = "0680667189"
        m_obj.AdresseLivraison.Email = "contact@marccollin.com"
        m_obj.CommCommande.comment = "Essai"
        m_obj.bExportInternet = 1

        Assert.AreEqual(m_obj.code, "CODE")
        Assert.AreEqual(m_obj.nom, "MonFournisseur")
        Assert.AreEqual(m_obj.rs, "RS")
        Assert.AreEqual(m_obj.banque, "BANQUE")
        Assert.AreEqual(m_obj.rib1, "RIB1")
        Assert.AreEqual(m_obj.rib2, "RIB2")
        Assert.AreEqual(m_obj.rib3, "RIB3")
        Assert.AreEqual(m_obj.rib4, "RIB4")
        Assert.AreEqual(m_obj.siret, "1234567890")
        Assert.AreEqual(m_obj.numtvaintracom, "0987654321")
        Assert.AreEqual(m_obj.CommCommande.comment, "Essai")
        Assert.AreEqual(m_obj.bExportInternet, 1)


        '<<TestMethod()>()> des indicateurs
        Assert.IsTrue(m_obj.bNew)
        Assert.IsTrue(m_obj.bUpdated)
        Assert.IsFalse(m_obj.bDeleted)

        Assert.IsTrue(m_obj.Equals(m_obj), "Egal à Lui même")
        m_obj.bExportInternet = 0
        objFRN = New Fournisseur("CODE", "MonFournisseur")
        objFRN.rs = "RS"
        objFRN.banque = "BANQUE"
        objFRN.rib1 = "RIB1"
        objFRN.rib2 = "RIB2"
        objFRN.rib3 = "RIB3"
        objFRN.rib4 = "RIB4"
        objFRN.numtvaintracom = "0987654321"
        objFRN.siret = "1234567890"
        objFRN.AdresseLivraison.nom = "Marc Collin"
        objFRN.AdresseLivraison.rue1 = "La Mettrie"
        objFRN.AdresseLivraison.rue2 = "2eme Etage"
        objFRN.AdresseLivraison.cp = "35250"
        objFRN.AdresseLivraison.ville = "chasné sur illet"
        objFRN.AdresseLivraison.tel = "0299555299"
        objFRN.AdresseLivraison.fax = "0299555277"
        objFRN.AdresseLivraison.port = "0680667189"
        objFRN.AdresseLivraison.Email = "contact@marccollin.com"
        objFRN.CommCommande.comment = "Essai"
        Assert.IsTrue(m_obj.Equals(objFRN), "Egal à un semblable")
        objFRN.bExportInternet = 1
        Assert.IsFalse(m_obj.Equals(objFRN), "Egal à un Différent")
        objFRN.bExportInternet = 0
        objFRN.CommFacturation.comment = "Essai"
        Assert.IsFalse(m_obj.Equals(objFRN), "Egal à un Différent")
        Dim obj As Object
        Assert.IsFalse(m_obj.Equals(obj), "Egal autrecjhose")


    End Sub
    <TestMethod()> Public Sub T15_DB()
        Dim objfrn As Fournisseur
        Dim objfrn2 As Fournisseur
        Dim n As Integer
        Dim objParam As ParamModeReglement

        'I - Création d'un Fournisseur
        '=========================
        objfrn = New Fournisseur("FTEST" & Now, "FTEST")
        objfrn.rs = "RSTEST"
        objfrn.banque = "TST_BANK"
        objfrn.rib1 = 12345
        objfrn.rib2 = 67890
        objfrn.rib3 = 1234567890
        objfrn.rib4 = 99
        objfrn.siret = ""
        objfrn.idRegion = 19
        objfrn.idModeReglement = 49
        objfrn.AdresseLivraison.nom = "Marc Collin"
        objfrn.AdresseLivraison.rue1 = "La Mettrie"
        objfrn.AdresseLivraison.rue2 = "2eme Etage"
        objfrn.AdresseLivraison.cp = "35250"
        objfrn.AdresseLivraison.ville = "chasné sur illet"
        objfrn.AdresseLivraison.tel = "0299555299"
        objfrn.AdresseLivraison.fax = "0299555277"
        objfrn.AdresseLivraison.port = "0680667189"
        objfrn.AdresseLivraison.Email = "contact@marccollin.com"
        objfrn.CommCommande.comment = "Commentaire de Commande"
        objfrn.CommLivraison.comment = "Commentaire de Livraison"
        objfrn.CommFacturation.comment = "Commentaire de Facturation"
        objfrn.CommLibre.comment = "Commentaire Libre"
        objfrn.bExportInternet = 1
        objParam = Param.colModeReglement(1)
        objfrn.idModeReglement1 = objParam.id
        objParam = Param.colModeReglement(2)
        objfrn.idModeReglement2 = objParam.id
        objParam = Param.colModeReglement(3)
        objfrn.idModeReglement3 = objParam.id

        '            objfrn.bAdressesIdentiques = True
        '<<TestMethod()>()> des indicateurs Avant le Save
        Assert.IsTrue(objfrn.bNew)
        Assert.IsTrue(objfrn.bUpdated)
        Assert.IsFalse(objfrn.bDeleted)
        'Save
        Assert.IsTrue(objfrn.Save(), "Insert" & objfrn.getErreur)
        Assert.IsTrue((objfrn.id <> 0), "Id Apres le Save doit être différent de 0")
        '<<TestMethod()>()> des indicateurs Après le Save
        Assert.IsFalse(objfrn.bNew, "bNew apres insert")
        Assert.IsFalse(objfrn.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objfrn.bDeleted, "bDeleted apres insert")

        'II - Rechargement d'un Fournisseur
        '=========================
        n = objfrn.id
        objfrn2 = New Fournisseur("", "")
        Assert.IsTrue(objfrn2.load(n), "Load" & Fournisseur.getErreur())
        Assert.IsTrue(objfrn.Equals(objfrn2), "L'objet rechargé est égal l'original")
        objParam = Param.colModeReglement(1)
        Assert.AreEqual(objfrn.idModeReglement1, objParam.id)
        objParam = Param.colModeReglement(2)
        Assert.AreEqual(objfrn.idModeReglement2, objParam.id)
        objParam = Param.colModeReglement(3)
        Assert.AreEqual(objfrn.idModeReglement3, objParam.id)

        'III - Modification du Fournisseur
        '=================================
        ' Modification du Fournisseur
        objfrn2.nom = objfrn2.nom + "Updated"
        '            objfrn.code = "FTEST" & Now()
        objfrn2.rs = objfrn2.rs + "Updated"
        objfrn2.banque = "TST_BANK2"
        objfrn2.rib1 = "12342"
        objfrn2.rib2 = "67892"
        objfrn2.rib3 = "1234567892"
        objfrn2.rib4 = "92"
        objfrn2.idRegion = Param.colRegion(2).id
        objfrn2.idModeReglement = Param.colModeReglement(1).id
        objfrn2.CommCommande.comment = "Commentaire Modifié"
        objfrn2.AdresseFacturation.fax = "211223344"
        objfrn2.bAdressesIdentiques = True
        objfrn2.idModeReglement2 = objfrn.idModeReglement3

        '<<TestMethod()>()> des indicateurs Avant le Save
        Assert.IsFalse(objfrn2.bNew)
        Assert.IsTrue(objfrn2.bUpdated)
        Assert.IsFalse(objfrn2.bDeleted)
        'Save
        Assert.IsTrue(objfrn2.Save(), "Update" & objfrn.getErreur)
        '<<TestMethod()>()> des indicateurs Après le Save
        Assert.IsFalse(objfrn2.bNew, "bNew apres insert")
        Assert.IsFalse(objfrn2.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objfrn2.bDeleted, "bDeleted apres insert")
        'Rechargement de l'objet
        n = objfrn2.id
        objfrn = New Fournisseur("", "")
        Assert.IsTrue(objfrn.load(n), "Load")
        Assert.IsTrue(objfrn.Equals(objfrn2))
        Assert.AreEqual(objfrn.idModeReglement2, objfrn.idModeReglement3)

        'IV - Suppression du Fournisseur
        '=================================
        ' Modification du Fournisseur
        objfrn.bDeleted = True
        Assert.IsTrue(objfrn.Save(), "Delete" & objfrn.getErreur())
        'Rechargement dans un autre objet
        objfrn2 = New Fournisseur("", "")
        '            Assert.IsFalse(objfrn2.load(n), "Load")
    End Sub
    <TestMethod()> Public Sub T51_ListeCriteres()
        Dim colFRN As Collection
        Dim oFRN As Fournisseur
        Dim objFRN2 As Fournisseur
        Dim nidFrn As Integer


        oFRN = New Fournisseur("FRNT51", "FRNT51NOM")
        oFRN.rs = "FRNT51RS"
        Assert.IsTrue(oFRN.Save(), oFRN.getErreur())
        nidFrn = oFRN.id

        'I - Liste Simple
        '================
        colFRN = Fournisseur.getListe()
        Assert.IsTrue(colFRN.Count > 0, "Col.count > 0")
        For Each oFRN In colFRN
            Assert.IsTrue(oFRN.id <> 0, "ID <> 0")
            Assert.IsTrue(oFRN.bResume, "bResume = True")
            objFRN2 = New Fournisseur("", "")
            Assert.IsTrue(objFRN2.load(oFRN.id), "Load" & objFRN2.getErreur())
            Assert.IsTrue(oFRN.code.Equals(objFRN2.code), "Codes" & oFRN.code & "<>" & objFRN2.code)
            Assert.IsTrue(oFRN.nom.Equals(objFRN2.nom), "Nom" & oFRN.nom & "<>" & objFRN2.nom)
            Assert.IsTrue(oFRN.id = objFRN2.id, "ids" & oFRN.id & "<>" & objFRN2.id)
        Next oFRN

        'II- Liste sur le code 
        '=============================
        'a) Caractère générique
        colFRN = Fournisseur.getListe("F%")
        Assert.IsTrue(colFRN.Count > 0, "Col.count > 0")
        For Each oFRN In colFRN
            Assert.IsTrue(Left(oFRN.code, 1) = "F", "Mauvais Code " & oFRN.code)
        Next oFRN
        'b) sans Caractère générique
        colFRN = Fournisseur.getListe("FRNT51")
        Assert.IsTrue(colFRN.Count > 0, "Col.count > 0" & Fournisseur.getErreur)
        For Each oFRN In colFRN
            Assert.IsTrue(oFRN.code = "FRNT51", "Mauvais Code " & oFRN.code)
        Next oFRN

        'III- Liste sur le nom
        '=============================
        'a) Caractère générique
        colFRN = Fournisseur.getListe(, "FRNT51%")
        Assert.IsTrue(colFRN.Count > 0, "Nom1:Col.count > 0" & Fournisseur.getErreur)
        For Each oFRN In colFRN
            Assert.IsTrue(Left(oFRN.nom, 6) = "FRNT51", "Mauvais Nom " & oFRN.nom)
        Next oFRN
        'b) sans Caractère générique
        colFRN = Fournisseur.getListe(, "FRNT51NOM")
        Assert.IsTrue(colFRN.Count > 0, "Nom2:Col.count > 0")
        For Each oFRN In colFRN
            Assert.IsTrue(oFRN.nom = "FRNT51NOM", "Mauvais Nom " & oFRN.nom)
        Next oFRN

        'IV- Liste sur la Raison Sociale
        '=============================
        'a) Caractère générique
        colFRN = Fournisseur.getListe(, , "FRNT51%")
        Assert.IsTrue(colFRN.Count > 0, "MotClé1: Col.count > 0" & Fournisseur.getErreur)
        objFRN2 = New Fournisseur("", "")
        For Each oFRN In colFRN
            Assert.IsTrue(objFRN2.load(oFRN.id), "Load False id=" & oFRN.id & oFRN.getErreur)
            Assert.AreEqual(Left(objFRN2.rs, 6), "FRNT51", "Mausvaise RS " & objFRN2.rs)
        Next oFRN
        'b) sans Caractère générique
        colFRN = Fournisseur.getListe(, , "FRNT51RS")
        objFRN2 = New Fournisseur("", "")
        Assert.IsTrue(colFRN.Count > 0, "Motclé2: Col.count > 0")
        For Each oFRN In colFRN
            Assert.IsTrue(objFRN2.load(oFRN.id), "Load False id=" & oFRN.id)
            Assert.AreEqual(objFRN2.rs, "FRNT51RS", "Mausvaise RS " & objFRN2.rs)
        Next oFRN
        'V - Critères Croisés
        '====================
        'a) Avec Résultat
        colFRN = Fournisseur.getListe("FRN%", "FRNT51%", "FRNT51%")
        Assert.IsTrue(colFRN.Count > 0, "Col.count > 0")
        objFRN2 = New Fournisseur("", "")
        For Each oFRN In colFRN
            Assert.IsTrue(Left(oFRN.code, 3) = "FRN", "Mauvais Code " & oFRN.code)
            Assert.IsTrue(Left(oFRN.nom, 6) = "FRNT51", "Mauvais Nom " & oFRN.nom)
            Assert.IsTrue(objFRN2.load(oFRN.id), "Load False id=" & oFRN.id)
            Assert.IsTrue(Left(objFRN2.rs, 6) = "FRNT51", "Mauvaise RS" & objFRN2.rs)
        Next oFRN
        'b) sans résultat
        colFRN = Fournisseur.getListe("0%", "<<TestMethod()>()>%", "XXX%")
        Assert.IsTrue(colFRN.Count = 0, "Col.count = 0")


        Assert.IsTrue(oFRN.load(nidFrn))
        oFRN.bDeleted = True
        Assert.IsTrue(oFRN.Save())



    End Sub
    '<<TestMethod()>()> la sauvegarde des champs long
    <TestMethod()> Public Sub T60_Champslongs()

        Dim obj As Fournisseur
        obj = New Fournisseur("T60", "Fournisseur")
        obj.AdresseLivraison.rue1 = "<<TestMethod()>()>".PadRight(500, "x")

        Assert.IsTrue(obj.Save())
        obj.AdresseFacturation.rue2 = "TEST2".PadRight(500, "y")
        Assert.IsTrue(obj.Save())
        obj.load()
        Assert.AreEqual(50, obj.AdresseLivraison.rue1.Length)
        Assert.AreEqual(50, obj.AdresseFacturation.rue2.Length)

        obj.bDeleted = True
        obj.Save()

    End Sub
    <TestMethod()> Public Sub T70_IdPrestashop()

        Dim obj As Fournisseur
        obj = New Fournisseur("T70", "ClientPrestashop")
        obj.idPrestashop = 15250
        Assert.IsTrue(obj.save())

        obj = Fournisseur.createandload(obj.id)
        Assert.AreEqual(15250L, obj.idPrestashop)
        obj.idPrestashop = 5263
        Assert.IsTrue(obj.save())
        obj.load()
        Assert.AreEqual(5263L, obj.idPrestashop)

        obj.bDeleted = True
        obj.save()

    End Sub

End Class



