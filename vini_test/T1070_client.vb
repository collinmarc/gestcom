'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1070_client
    Inherits test_Base

    Private m_obj As Client
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        m_obj.shared_connect()
        m_obj = New Client("", "")
    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        m_obj.shared_disconnect()
        MyBase.TestCleanup()
    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objCLT As Client

        m_obj.code = "CODE"
        m_obj.nom = "NOM"
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
        m_obj.CodeTarif = "B"

        Assert.AreEqual(m_obj.code, "CODE")
        Assert.AreEqual(m_obj.nom, "NOM")
        Assert.AreEqual(m_obj.rs, "RS")
        Assert.AreEqual(m_obj.banque, "BANQUE")
        Assert.AreEqual(m_obj.rib1, "RIB1")
        Assert.AreEqual(m_obj.rib2, "RIB2")
        Assert.AreEqual(m_obj.rib3, "RIB3")
        Assert.AreEqual(m_obj.rib4, "RIB4")
        Assert.AreEqual(m_obj.siret, "1234567890")
        Assert.AreEqual(m_obj.numtvaintracom, "0987654321")
        Assert.AreEqual(m_obj.CodeTarif, "B")


        'Test des indicateurs
        Assert.IsTrue(m_obj.bNew)
        Assert.IsTrue(m_obj.bUpdated)
        Assert.IsFalse(m_obj.bDeleted)

        Assert.IsTrue(m_obj.Equals(m_obj), "Egal à Lui même")
        objCLT = New Client("CODE", "NOM")
        objCLT.rs = "RS"
        objCLT.banque = "BANQUE"
        objCLT.rib1 = "RIB1"
        objCLT.rib2 = "RIB2"
        objCLT.rib3 = "RIB3"
        objCLT.rib4 = "RIB4"
        objCLT.numtvaintracom = "0987654321"
        objCLT.siret = "1234567890"
        objCLT.AdresseLivraison.nom = "Marc Collin"
        objCLT.AdresseLivraison.rue1 = "La Mettrie"
        objCLT.AdresseLivraison.rue2 = "2eme Etage"
        objCLT.AdresseLivraison.cp = "35250"
        objCLT.AdresseLivraison.ville = "chasné sur illet"
        objCLT.AdresseLivraison.tel = "0299555299"
        objCLT.AdresseLivraison.fax = "0299555277"
        objCLT.AdresseLivraison.port = "0680667189"
        objCLT.AdresseLivraison.Email = "contact@marccollin.com"
        objCLT.CodeTarif = "B"
        Assert.IsTrue(m_obj.Equals(objCLT), "Egal à un semblable")
        objCLT.CommFacturation.comment = "Mon commentaire"
        Assert.IsFalse(m_obj.Equals(objCLT), "Egal à un Différent")
        Dim obj As Object
        Assert.IsFalse(m_obj.Equals(obj), "Egal autrecjhose")





    End Sub
    <TestMethod()> Public Sub T15_DB()
        Dim objCLT As Client
        Dim objCLT2 As Client
        Dim n As Integer
        Dim objParam1 As ParamModeReglement
        Dim objParam2 As ParamModeReglement
        Dim objParam3 As ParamModeReglement

        objParam1 = Param.colModeReglement(1)
        objParam2 = Param.colModeReglement(2)
        objParam3 = Param.colModeReglement(3)


        'I - Création d'un Client
        '=========================
        objCLT = New Client("FTEST" & Now(), "")
        Assert.AreEqual("A", objCLT.CodeTarif, "Code Tarif par Defaut")
        objCLT.nom = "TEST"
        objCLT.rs = "RSTEST"
        objCLT.banque = "TST_BANK"
        objCLT.rib1 = 12345
        objCLT.rib2 = 67890
        objCLT.rib3 = 1234567890
        objCLT.rib4 = 99
        objCLT.siret = ""
        objCLT.idTypeClient = 8
        objCLT.idModeReglement = 49
        objCLT.AdresseLivraison.nom = "Marc Collin"
        objCLT.AdresseLivraison.rue1 = "La Mettrie"
        objCLT.AdresseLivraison.rue2 = "2eme Etage"
        objCLT.AdresseLivraison.cp = "35250"
        objCLT.AdresseLivraison.ville = "chasné sur illet"
        objCLT.AdresseLivraison.tel = "0299555299"
        objCLT.AdresseLivraison.fax = "0299555277"
        objCLT.AdresseLivraison.port = "0680667189"
        objCLT.AdresseLivraison.Email = "contact@marccollin.com"
        objCLT.CommCommande.comment = "Commentaire de Commande"
        objCLT.CommLivraison.comment = "Commentaire de Livraison"
        objCLT.CommFacturation.comment = "Commentaire de Facturation"
        objCLT.CommLibre.comment = "Commentaire Libre"
        objCLT.CodeTarif = "B"
        objCLT.idModeReglement1 = objParam1.id
        objCLT.idModeReglement2 = objParam2.id
        objCLT.idModeReglement3 = objParam3.id

        'Test des indicateurs Avant le Save
        Assert.IsTrue(objCLT.bNew)
        Assert.IsTrue(objCLT.bUpdated)
        Assert.IsFalse(objCLT.bDeleted)
        'Save
        Assert.IsTrue(objCLT.save(), "Insert" & objCLT.getErreur)
        Assert.IsTrue((objCLT.id <> 0), "Id Apres le Save doit être différent de 0")
        'Test des indicateurs Après le Save
        Assert.IsFalse(objCLT.bNew, "bNew apres insert")
        Assert.IsFalse(objCLT.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objCLT.bDeleted, "bDeleted apres insert")

        'II - Rechargement d'un Client
        '=========================
        n = objCLT.id
        objCLT2 = Client.createandload(n)
        Assert.AreEqual("B", objCLT2.CodeTarif, "Code Tarif après rechargement")
        Assert.AreEqual(objCLT.idModeReglement1, objParam1.id)
        Assert.AreEqual(objCLT.idModeReglement2, objParam2.id)
        Assert.AreEqual(objCLT.idModeReglement3, objParam3.id)

        Assert.IsTrue(objCLT.Equals(objCLT2))

        'III - Modification du Client
        '=================================
        ' Modification du Client
        objCLT2.nom = objCLT2.nom + "Updated"
        objCLT2.code = "FTEST2" & Now()
        objCLT2.AdresseLivraison.nom = "toto"
        objCLT2.rs = objCLT2.rs + "Updated"
        objCLT2.banque = "TST_BANK2"
        objCLT2.rib1 = "12342"
        objCLT2.rib2 = "67892"
        objCLT2.rib3 = "1234567892"
        objCLT2.rib4 = "92"
        objCLT2.idTypeClient = Param.typeclientdefaut.id
        objCLT2.idModeReglement = Param.ModeReglementdefaut.id
        objCLT2.bAdressesIdentiques = False
        objCLT2.CodeTarif = "C"
        objCLT2.idModeReglement2 = objCLT.idModeReglement3

        'Test des indicateurs Avant le Save
        Assert.IsFalse(objCLT2.bNew)
        Assert.IsTrue(objCLT2.bUpdated)
        Assert.IsFalse(objCLT2.bDeleted)
        'Save
        Assert.IsTrue(objCLT2.save(), "Update" & objCLT.getErreur)
        'Test des indicateurs Après le Save
        Assert.IsFalse(objCLT2.bNew, "bNew apres insert")
        Assert.IsFalse(objCLT2.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objCLT2.bDeleted, "bDeleted apres insert")
        'Rechargement de l'objet
        n = objCLT2.id
        objCLT = New Client("", "")
        Assert.IsTrue(objCLT.load(n), "Load")
        Assert.AreEqual("C", objCLT.CodeTarif)
        Assert.AreEqual(objCLT.idModeReglement2, objCLT.idModeReglement3)
        Assert.IsTrue(objCLT.Equals(objCLT2), "Apres Update , Equals")

        'Modification d'un element inclus
        Assert.IsFalse(objCLT.bUpdated, "Updated = false avant")
        objCLT.AdresseLivraison.cp = "45678"
        Assert.IsTrue(objCLT.bUpdated, "Updated = true apres")


        'IV - Suppression du Client
        '=================================
        ' Modification du Client
        objCLT.bDeleted = True
        Assert.IsTrue(objCLT.save(), "Delete" & objCLT.getErreur())
        'Rechargement dans un autre objet
        objCLT2 = New Client("", "")
        Assert.IsFalse(objCLT2.load(n), "Load")
    End Sub
    <TestMethod()> Public Sub T51_ListeCriteres()
        Dim colCLT As Collection
        Dim oCLT As Client
        Dim objCLT2 As Client
        Dim nidClient1 As Integer
        Dim nidClient2 As Integer
        Dim nidClient3 As Integer

        oCLT = New Client("T51TESTCODE1-1", "T51TESTNOM1")
        oCLT.rs = "MCII"
        oCLT.idModeReglement = Param.ModeReglementdefaut.id
        oCLT.idTypeClient = Param.typeclientdefaut.id
        Assert.IsTrue(oCLT.save(), "Insert" & Client.getErreur)
        nidClient1 = oCLT.id
        oCLT = New Client("T51TESTCODE2-1", "T51TESTNOM2")
        oCLT.idModeReglement = Param.ModeReglementdefaut.id
        oCLT.idTypeClient = Param.typeclientdefaut.id
        oCLT.rs = "MCOLL"
        Assert.IsTrue(oCLT.save(), "Insert" & Client.getErreur)
        nidClient2 = oCLT.id
        oCLT = New Client("T51TESTCODE3-1", "T51TESTNOM3")
        oCLT.idModeReglement = Param.ModeReglementdefaut.id
        oCLT.idTypeClient = Param.typeclientdefaut.id
        Assert.IsTrue(oCLT.save(), "Insert" & Client.getErreur)
        oCLT.rs = "RIEN"
        nidClient3 = oCLT.id

        'I - Liste Simple
        '================
        colCLT = Client.getListe()
        Assert.IsTrue(colCLT.Count > 0, "Liste Complête Col.count > 0")
        For Each oCLT In colCLT
            Assert.IsTrue(oCLT.id <> 0, "ID <> 0")
            Assert.IsTrue(oCLT.bResume, "bResume = True")
            objCLT2 = New Client("", "")
            Assert.IsTrue(oCLT.load(), "Load(" & oCLT.id & ")" & oCLT.getErreur())
            'Assert.IsTrue(oCLT.code.Equals(objCLT2.code), "Codes" & oCLT.code & "<>" & objCLT2.code)
            'Assert.IsTrue(oCLT.nom.Equals(objCLT2.nom), "Nom" & oCLT.nom & "<>" & objCLT2.nom)
            'Assert.IsTrue(oCLT.id = objCLT2.id, "ids" & oCLT.id & "<>" & objCLT2.id)
        Next oCLT

        'II- Liste sur le code 
        '=============================
        'a) Caractère générique
        colCLT = Client.getListe("T51%")
        Assert.IsTrue(colCLT.Count > 0, "Liste sur le code Col.count > 0")
        For Each oCLT In colCLT
            Assert.IsTrue(Left(oCLT.code, 3) = "T51", "ListeCode Mauvais Code " & oCLT.code)
        Next oCLT
        'b) sans Caractère générique
        colCLT = Client.getListe("T51TESTCODE1-1")
        Assert.IsTrue(colCLT.Count > 0, "Col.count > 0" & Client.getErreur)
        For Each oCLT In colCLT
            Assert.IsTrue(oCLT.code = "T51TESTCODE1-1", "ListeCode Mauvais Code " & oCLT.code)
        Next oCLT

        'III- Liste sur le nom
        '=============================
        'a) Caractère générique
        colCLT = Client.getListe(, "T51%")
        Assert.IsTrue(colCLT.Count > 0, "Nom1:Col.count > 0" & Client.getErreur)
        For Each oCLT In colCLT
            Assert.IsTrue(Left(oCLT.nom, 3) = "T51", "ListeNom Mauvais Nom " & oCLT.nom)
        Next oCLT
        'b) sans Caractère générique
        colCLT = Client.getListe(, "T51TESTNOM1")
        Assert.IsTrue(colCLT.Count > 0, "Nom2:Col.count > 0")
        For Each oCLT In colCLT
            Assert.IsTrue(oCLT.nom = "T51TESTNOM1", "ListeNom Mauvais Nom " & oCLT.nom)
        Next oCLT

        'IV- Liste sur la raison sociale
        '=============================
        'a) Caractère générique
        colCLT = Client.getListe(, , "MC%")
        Assert.IsTrue(colCLT.Count > 0, "Liste MotClé: Col.count > 0" & Client.getErreur)
        objCLT2 = New Client("", "")
        For Each oCLT In colCLT
            Assert.IsTrue(objCLT2.load(oCLT.id), "ListeRS Load False id=" & oCLT.id & Client.getErreur())
            Assert.IsTrue(Left(objCLT2.rs, 2) = "MC", "ListeRS Load False id=" & oCLT.id & Client.getErreur())
        Next oCLT
        'b) sans Caractère générique
        colCLT = Client.getListe(, , "MCII")
        objCLT2 = New Client("", "")
        Assert.IsTrue(colCLT.Count > 0, "Motclé2: Col.count > 0")
        For Each oCLT In colCLT
            Assert.IsTrue(objCLT2.load(oCLT.id), "Load False id=" & oCLT.id)
            Assert.IsTrue(objCLT2.rs = "MCII", "ListeRS Load False id=" & oCLT.id & Client.getErreur())
        Next oCLT
        'V - Critères Croisés
        '====================
        'a) Avec Résultat
        colCLT = Client.getListe("T51%", "T51%", "MC%")
        Assert.IsTrue(colCLT.Count > 0, "Col.count > 0")
        objCLT2 = New Client("", "")
        For Each oCLT In colCLT
            Assert.IsTrue(Left(oCLT.code, 3) = "T51", "Mauvais Code " & oCLT.code)
            Assert.IsTrue(Left(oCLT.nom, 3) = "T51", "Mauvais Nom " & oCLT.nom)
            Assert.IsTrue(objCLT2.load(oCLT.id), "Load False id=" & oCLT.id)
            Assert.IsTrue(Left(objCLT2.rs, 2) = "MC", "Mauvaise RS" & objCLT2.rs)
        Next oCLT
        'b) sans résultat
        colCLT = Client.getListe("T51%", "T51T%", "XXX%")
        Assert.IsTrue(colCLT.Count = 0, "Col.count = 0")

        'Suppression des Clients
        Assert.IsTrue(oCLT.load(nidClient1))
        oCLT.bDeleted = True
        Assert.IsTrue(oCLT.save())
        Assert.IsTrue(oCLT.load(nidClient2))
        oCLT.bDeleted = True
        Assert.IsTrue(oCLT.save())
        Assert.IsTrue(oCLT.load(nidClient3))
        oCLT.bDeleted = True
        Assert.IsTrue(oCLT.save())

    End Sub
    <TestMethod()> Public Sub T60_Precommande()
        Dim oCLT As Client
        Dim oPrd1 As Produit
        Dim oPrd2 As Produit
        Dim oPrd3 As Produit
        Dim oFRN1 As Fournisseur
        Dim oFRN2 As Fournisseur
        Dim nIDClient As Long
        Dim objLgPRecom As lgPrecomm
        Dim objLgPRecom1 As lgPrecomm
        Dim objLgPRecom2 As lgPrecomm
        Dim sCode As String
        Dim objCMD As CommandeClient


        'Creation de 2 Fournisseurs
        oFRN1 = New Fournisseur("T60FRN1TEST" & Now(), "")
        Assert.IsTrue(oFRN1.Save(), "oFRN1.save")
        oFRN2 = New Fournisseur("T60FRN2TEST" & Now(), "")
        Assert.IsTrue(oFRN2.Save(), "oFRN2.save")

        'Creation de 3 Produits
        oPrd1 = New Produit("T60PRD1" & Now(), oFRN1, 97)
        Assert.IsTrue(oPrd1.save(), "oPRD1.save")
        Assert.IsTrue(oPrd1.id <> 0, "oPRD1<>0")
        oPrd2 = New Produit("T60PRD2" & Now(), oFRN2, 98)
        Assert.IsTrue(oPrd2.save(), "oPRD2.save")
        oPrd3 = New Produit("T60PRD3" & Now(), oFRN2, 99)
        Assert.IsTrue(oPrd3.save(), "oPRD3.save")

        'Creation d'un client
        oCLT = New Client("T60CLT1" & Now(), "T60")
        Assert.IsTrue(oCLT.save(), "oCLT.Insert" & Client.getErreur)

        'Chargement de la Precommande vide
        Assert.IsTrue(oCLT.LoadPreCommande(), "LoadPrecommande" & Client.getErreur)
        Assert.AreEqual(oCLT.getlgPrecomCount, 0, "Collection non vide")

        'Ajout de 2 ligne de Precommande
        objLgPRecom1 = oCLT.ajouteLgPrecom(oPrd1, 12, 20, 11, "31/07/1964")
        Assert.IsTrue(Not objLgPRecom1 Is Nothing, "Ajout OPRD1")
        objLgPRecom2 = oCLT.ajouteLgPrecom(oPrd3, 12, 20, 11)
        Assert.IsTrue(Not objLgPRecom2 Is Nothing, "Ajout OPRD3")
        'L'ajout de la 3eme rend False car il s'agit du même code Produit
        Assert.IsNull(oCLT.ajouteLgPrecom(oPrd3.id, oPrd3.code, oPrd3.nom, 12, 20, 11), "Ajout OPRD3bis")
        'Sauvegarde de la precommande
        Assert.IsTrue(oCLT.save(), "Oclt.Save")
        'Rechargement du client et de sa precommande
        nIDClient = oCLT.id
        oCLT = New Client("", "")
        Assert.IsTrue(oCLT.load(nIDClient), "OCLT.load")
        Assert.IsTrue(oCLT.LoadPreCommande(), "OCLT.LoadPrecommande")
        Assert.AreEqual(oCLT.getlgPrecomCount, 2, "Precommande.count ")
        'controle du contenu de daate d derniere commande
        objLgPRecom = oCLT.getLgPrecomByProductId(oPrd1.id)
        Assert.IsTrue(objLgPRecom.Equals(objLgPRecom1), "LgPrecom1 sont différents apres rechargement")
        objLgPRecom = oCLT.getLgPrecomByProductId(oPrd3.id)
        Assert.IsTrue(objLgPRecom.Equals(objLgPRecom2), "LgPrecom2 sont différents apres rechargement")
        'Ajout d'une ligne de precommande
        Assert.IsTrue(Not oCLT.ajouteLgPrecom(oPrd2.id, oPrd2.code, oPrd2.nom, 113, 20, 11.0) Is Nothing, "oclt.AjouteLg 3")
        'Sauvegarde de la precommande
        Assert.IsTrue(oCLT.save(), "Oclt.Save")
        'Rechargement du client et de sa precommande
        nIDClient = oCLT.id
        oCLT = New Client("", "")
        Assert.IsTrue(oCLT.load(nIDClient), "OCLT.load")
        Assert.IsTrue(oCLT.LoadPreCommande(), "OCLT.LoadPrecommande")
        Assert.AreEqual(oCLT.getlgPrecomCount, 3, "Precommande.count ")
        'Suppression d'une ligne de la precommande
        oCLT.supprimeLgPrecom(3)
        'Sauvegarde de la precommande
        Assert.IsTrue(oCLT.save(), "Oclt.Save")
        'Rechargement du client et de sa precommande
        nIDClient = oCLT.id
        oCLT = New Client("", "")
        Assert.IsTrue(oCLT.load(nIDClient), "OCLT.load")
        Assert.IsTrue(oCLT.LoadPreCommande(), "OCLT.LoadPrecommande")
        Assert.AreEqual(oCLT.getlgPrecomCount, 2, "Precommande.count ")

        'Maj d'une ligne de la precommande
        objLgPRecom = oCLT.getLgPrecomByProductId(oPrd2.id)
        sCode = objLgPRecom.codeProduit
        objLgPRecom.qteHab = 150
        objLgPRecom.qteDern = 5
        objLgPRecom.prixU = 33
        'Sauvegarde de la precommande
        Assert.IsTrue(oCLT.save(), "Oclt.Save")
        'Rechargement du client et de sa precommande
        nIDClient = oCLT.id
        oCLT = New Client("", "")
        Assert.IsTrue(oCLT.load(nIDClient), "OCLT.load")
        Assert.IsTrue(oCLT.LoadPreCommande(), "OCLT.LoadPrecommande")
        Assert.AreEqual(oCLT.getlgPrecomCount, 2, "Precommande.count ")
        objLgPRecom = oCLT.getLgPrecomByProductId(oPrd2.id)
        Assert.AreEqual(objLgPRecom.qteHab, CDec(150))
        Assert.AreEqual(objLgPRecom.qteDern, CDec(5))
        Assert.AreEqual(objLgPRecom.prixU, CDbl(33))

        oCLT = New Client("", "")
        Assert.IsTrue(oCLT.load(nIDClient), "OCLT.load")

        'Creation d'une commande client
        objCMD = New CommandeClient(oCLT)
        objCMD.AjouteLigne("10", oPrd1, 5, 34.5)
        objCMD.AjouteLigne("20", oPrd2, 10, 53.5)
        Assert.IsTrue(objCMD.save())
        'Rechargement du client et de sa precommande
        oCLT = New Client("", "")
        Assert.IsTrue(oCLT.load(nIDClient), "OCLT.load")
        Assert.IsTrue(oCLT.LoadPreCommande(), "OCLT.LoadPrecommande")
        Assert.AreEqual(oCLT.getlgPrecomCount, 2, "Precommande.count ")
        objLgPRecom = oCLT.getLgPrecomByProductId(oPrd1.id)
        Assert.AreEqual(objLgPRecom.qteHab, CDec(12))
        Assert.AreEqual(objLgPRecom.qteDern, CDec(5))
        Assert.AreEqual(objLgPRecom.prixU, CDbl(34.5))

        '        Assert.IsTrue(oCLT.lgPrecomExists(oPrd1.code))

        'Suppression du Client 
        oCLT.bDeleted = True
        nIDClient = oCLT.id
        Assert.IsTrue(oCLT.save(), "OCLT.Delete")
        Assert.IsTrue(MsgBox("Vérifier que'il n'y a plus de enr dans LGPRECOM pour L'id Client " & nIDClient, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)
        'Suppression des 3 Produits
        oPrd1.bDeleted = True
        Assert.IsTrue(oPrd1.save(), "oPRD1.delete")
        oPrd2.bDeleted = True
        Assert.IsTrue(oPrd2.save(), "oPRD2.delete")
        oPrd3.bDeleted = True
        Assert.IsTrue(oPrd3.save(), "oPRD3.delete")

        'Supression des 2 fournisseurs
        oFRN1.bDeleted = True
        Assert.IsTrue(oFRN1.Save, "oFRN1.Delete")
        oFRN2.bDeleted = True
        Assert.IsTrue(oFRN2.Save, "oFRN2.Delete")
    End Sub

    <TestMethod()> Public Sub T70_ReinitPrecommande()
        Dim objPRD1 As Produit
        Dim objPRD2 As Produit
        Dim objCommande As CommandeClient
        Dim nid As Integer
        Dim objFRN As Fournisseur
        Dim objPRD As Produit
        Dim objCLT As Client

        objFRN = New Fournisseur("FRNT701", "Fournisseur de Test")
        Assert.IsTrue(objFRN.Save(), "Sauvegarde du fournisseur")

        objPRD = New Produit("PRD", objFRN, 1999)
        Assert.IsTrue(objPRD.save, "Sauvegarde du produit")
        objPRD1 = New Produit("PRD1", objFRN, 1999)
        Assert.IsTrue(objPRD1.save, "Sauvegarde du produit1")
        objPRD2 = New Produit("PRD2", objFRN, 1999)
        Assert.IsTrue(objPRD2.save, "Sauvegarde du produit2")


        objCLT = New Client("CLTT701", "client de test T70")
        Assert.IsTrue(objCLT.save(), "Sauvegarde Client")
        objCommande = New CommandeClient(objCLT)
        objCommande.AjouteLigne("10", objPRD, 1, 1)
        Assert.IsTrue(objCommande.save, "Sauvegarde Commande")


        'Controle de la precommande du client
        objCLT.LoadPreCommande()
        Assert.AreEqual(1, objCLT.getlgPrecomCount(), "1 ligne en Precommande")

        'Ajout d'une ligne de precommande
        objCLT.ajouteLgPrecom(objPRD1, 10, 2)
        'Sauvegarde du client
        Assert.IsTrue(objCLT.save(), "Sauvegarde Client")
        nid = objCLT.id
        'Rechargement du client
        objCLT = Client.createandload(nid)
        objCLT.LoadPreCommande()
        'Controle de la precommande du client
        Assert.AreEqual(2, objCLT.getlgPrecomCount(), "2 lignes en Precommande")


        'Reinitialisation du client
        Assert.IsTrue(objCLT.reinitPrecommande(), "Reinit Precommande")
        'Sauvegarde du client
        Assert.IsTrue(objCLT.save(), "Sauvegarde Client")
        nid = objCLT.id
        'Rechargement du client
        objCLT = Client.createandload(nid)
        objCLT.LoadPreCommande()
        'Controle de la precommande du client
        Assert.AreEqual(1, objCLT.getlgPrecomCount(), "1 ligne en Precommande")

        objCommande.bDeleted = True
        Assert.IsTrue(objCommande.save, "Suppression Commande")

        objPRD.bDeleted = True
        Assert.IsTrue(objPRD.save(), "Suppression PRD")
        objPRD1.bDeleted = True
        Assert.IsTrue(objPRD1.save(), "Suppression PRD1")
        objPRD2.bDeleted = True
        Assert.IsTrue(objPRD2.save(), "Suppression PRD2")

        objFRN.bDeleted = True
        Assert.IsTrue(objFRN.Save(), "Suppression FRN")

        objCLT.bDeleted = True
        Assert.IsTrue(objCLT.save(), "Suppression client")


    End Sub

    'Test la sauvegarde des champs long
    <TestMethod()> Public Sub T60_Champslongs()

        Dim obj As Client
        obj = New Client("T60", "Fournisseur")
        obj.AdresseLivraison.rue1 = "TEST".PadRight(500, "x")

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

        Dim obj As Client
        obj = New Client("T70", "ClientPrestashop")
        obj.idPrestashop = 15250
        Assert.IsTrue(obj.save())

        obj = Client.createandload(obj.id)
        Assert.AreEqual(15250L, obj.idPrestashop)
        obj.idPrestashop = 5263
        Assert.IsTrue(obj.save())
        obj.load()
        Assert.AreEqual(5263L, obj.idPrestashop)

        obj.bDeleted = True
        obj.save()

    End Sub
End Class



