'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class T1120_BA
    Inherits test_Base

    Private m_oProduit As Produit
    Private m_oFourn As Fournisseur
    Private m_oClient As Client
    <TestInitialize()> Public Overrides Sub TestInitialize()

        MyBase.TestInitialize()
        m_oFourn = New Fournisseur("FRN" & Now(), "MonFournisseur")
        m_oFourn.Save()

        m_oProduit = New Produit("PRD" & Now(), m_oFourn, 1990)
        m_oProduit.save()

        m_oClient = New Client("CLT" & Now(), "MonClient")
        Debug.Assert(m_oClient.save(), "Creation du client")
        '            m_oClient()

    End Sub
    <TestCleanup()> Public Overrides Sub TestCleanup()
        m_oProduit.bDeleted = True
        m_oProduit.save()
        m_oFourn.bDeleted = True
        m_oFourn.Save()
        m_oClient.bDeleted = True
        m_oClient.save()

        MyBase.TestCleanup()


    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objCMD As BonAppro
        Dim objCMD2 As BonAppro

        objCMD = New BonAppro(m_oFourn)

        objCMD.code = "CODE"
        objCMD.dateCommande = CDate("06/02/1964")
        objCMD.caracteristiqueTiers.banque = "BANQUE"
        objCMD.caracteristiqueTiers.rib1 = "RIB1"
        objCMD.caracteristiqueTiers.rib2 = "RIB2"
        objCMD.caracteristiqueTiers.rib3 = "RIB3"
        objCMD.caracteristiqueTiers.rib4 = "RIB4"
        objCMD.caracteristiqueTiers.AdresseLivraison.nom = "Marc Collin"
        objCMD.caracteristiqueTiers.AdresseLivraison.rue1 = "La Mettrie"
        objCMD.caracteristiqueTiers.AdresseLivraison.rue2 = "2eme Etage"
        objCMD.caracteristiqueTiers.AdresseLivraison.cp = "35250"
        objCMD.caracteristiqueTiers.AdresseLivraison.ville = "chasné sur illet"
        objCMD.caracteristiqueTiers.AdresseLivraison.tel = "0299555299"
        objCMD.caracteristiqueTiers.AdresseLivraison.fax = "0299555277"
        objCMD.caracteristiqueTiers.AdresseLivraison.port = "0680667189"
        objCMD.caracteristiqueTiers.AdresseLivraison.Email = "contact@marccollin.com"
        objCMD.caracteristiqueTiers.AdresseFacturation.nom = "Marc Collin" & "Fact"
        objCMD.caracteristiqueTiers.AdresseFacturation.rue1 = "La Mettrie" & "Fact"
        objCMD.caracteristiqueTiers.AdresseFacturation.rue2 = "2eme Etage" & "Fact"
        objCMD.caracteristiqueTiers.AdresseFacturation.cp = "35250" & "Fact"
        objCMD.caracteristiqueTiers.AdresseFacturation.ville = "chasné sur illet" & "Fact"
        objCMD.caracteristiqueTiers.AdresseFacturation.tel = "0299555299" & "Fact"
        objCMD.caracteristiqueTiers.AdresseFacturation.fax = "0299555277" & "Fact"
        objCMD.caracteristiqueTiers.AdresseFacturation.port = "0680667189" & "Fact"
        objCMD.caracteristiqueTiers.AdresseFacturation.Email = "contact@marccollin.com" & "Fact"
        objCMD.dateLivraison = CDate("06/02/1964")
        objCMD.dateEnlevement = CDate("31/07/1964")
        objCMD.refLivraison = "BL0003"

        Assert.AreEqual(objCMD.code, "CODE")
        Assert.AreEqual(objCMD.caracteristiqueTiers.banque, "BANQUE")
        Assert.AreEqual(objCMD.caracteristiqueTiers.rib1, "RIB1")
        Assert.AreEqual(objCMD.caracteristiqueTiers.rib2, "RIB2")
        Assert.AreEqual(objCMD.caracteristiqueTiers.rib3, "RIB3")
        Assert.AreEqual(objCMD.caracteristiqueTiers.rib4, "RIB4")
        Assert.AreEqual(objCMD.dateLivraison, CDate("06/02/1964"))
        Assert.AreEqual(objCMD.dateEnlevement, CDate("31/07/1964"))
        Assert.AreEqual(objCMD.refLivraison, "BL0003")


        'Test des indicateurs
        Assert.IsTrue(objCMD.bNew)
        Assert.IsTrue(objCMD.bUpdated)
        Assert.IsFalse(objCMD.bDeleted)

        Assert.IsTrue(objCMD.Equals(objCMD), "Egal à Lui même")
        objCMD2 = New BonAppro(m_oFourn)

        objCMD2.code = "CODE"
        objCMD2.dateCommande = CDate("06/02/1964")
        objCMD2.caracteristiqueTiers.banque = "BANQUE"
        objCMD2.caracteristiqueTiers.rib1 = "RIB1"
        objCMD2.caracteristiqueTiers.rib2 = "RIB2"
        objCMD2.caracteristiqueTiers.rib3 = "RIB3"
        objCMD2.caracteristiqueTiers.rib4 = "RIB4"
        objCMD2.caracteristiqueTiers.AdresseLivraison.nom = "Marc Collin"
        objCMD2.caracteristiqueTiers.AdresseLivraison.rue1 = "La Mettrie"
        objCMD2.caracteristiqueTiers.AdresseLivraison.rue2 = "2eme Etage"
        objCMD2.caracteristiqueTiers.AdresseLivraison.cp = "35250"
        objCMD2.caracteristiqueTiers.AdresseLivraison.ville = "chasné sur illet"
        objCMD2.caracteristiqueTiers.AdresseLivraison.tel = "0299555299"
        objCMD2.caracteristiqueTiers.AdresseLivraison.fax = "0299555277"
        objCMD2.caracteristiqueTiers.AdresseLivraison.port = "0680667189"
        objCMD2.caracteristiqueTiers.AdresseLivraison.Email = "contact@marccollin.com"
        objCMD2.caracteristiqueTiers.AdresseFacturation.nom = "Marc Collin" & "Fact"
        objCMD2.caracteristiqueTiers.AdresseFacturation.rue1 = "La Mettrie" & "Fact"
        objCMD2.caracteristiqueTiers.AdresseFacturation.rue2 = "2eme Etage" & "Fact"
        objCMD2.caracteristiqueTiers.AdresseFacturation.cp = "35250" & "Fact"
        objCMD2.caracteristiqueTiers.AdresseFacturation.ville = "chasné sur illet" & "Fact"
        objCMD2.caracteristiqueTiers.AdresseFacturation.tel = "0299555299" & "Fact"
        objCMD2.caracteristiqueTiers.AdresseFacturation.fax = "0299555277" & "Fact"
        objCMD2.caracteristiqueTiers.AdresseFacturation.port = "0680667189" & "Fact"
        objCMD2.caracteristiqueTiers.AdresseFacturation.Email = "contact@marccollin.com" & "Fact"
        objCMD2.dateLivraison = CDate("06/02/1964")
        objCMD2.dateEnlevement = CDate("31/07/1964")
        objCMD2.refLivraison = "BL0003"

        Assert.IsTrue(objCMD.Equals(objCMD2), "Egal à un semblable")
        objCMD2.caracteristiqueTiers.banque = "Deuxième banque"
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.caracteristiqueTiers.banque = objCMD.caracteristiqueTiers.banque
        objCMD2.dateLivraison = CDate("07/02/1964")
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.dateEnlevement = CDate("07/02/1964")
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.refLivraison = "BL0002"
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        Dim obj As Object
        Assert.IsFalse(objCMD.Equals(obj), "Egal autrecjhose")


    End Sub
    <TestMethod()> Public Sub T12_GestLignes()
        Dim objComm As BonAppro
        Dim objLg As LgCommande



        objComm = New BonAppro(m_oFourn)

        objLg = objComm.AjouteLigne("10", m_oProduit, 10, 10)
        Assert.IsTrue(objComm.colLignes.Count = 1)
        Assert.IsTrue(objComm.totalHT = 10 * 10)
        Assert.AreEqual(objComm.totalTTC, CDec(objComm.totalHT * 1.196))
        Assert.IsTrue(Not objLg Is Nothing, "Ajout de Ligne impossible")
        Assert.IsTrue(objLg.num = objComm.colLignes.Count * 10)
        objLg = objComm.AjouteLigne("20", m_oProduit, 20, 10)
        Assert.IsTrue(objComm.colLignes.Count = 2)
        Assert.IsTrue(objComm.totalHT = (10 * 10) + (20 * 10))
        Assert.AreEqual(objComm.totalTTC, CDec(objComm.totalHT * 1.196))
        Assert.IsTrue(Not objLg Is Nothing, "Ajout de Ligne impossible")
        Assert.IsTrue(objLg.num = objComm.colLignes.Count * 10)
        objLg = objComm.AjouteLigne("30", m_oProduit, 30, 10)
        Assert.IsTrue(objComm.colLignes.Count = 3)
        Assert.IsTrue(objComm.totalHT = (10 * 10) + (20 * 10) + (30 * 10))
        Assert.IsTrue(objComm.totalTTC = (objComm.totalHT * 1.196))
        Assert.IsTrue(Not objLg Is Nothing, "Ajout de Ligne impossible")
        Assert.IsTrue(objLg.num = objComm.colLignes.Count * 10)

        Assert.IsTrue(Not objComm.estEntierementLivree())
        For Each objLg In objComm.colLignes
            objLg.qteLiv = objLg.qteCommande
        Next objLg
        Assert.IsTrue(objComm.estEntierementLivree())


    End Sub 'T12
    <TestMethod()> Public Sub T15_DB()
        Dim objBA As BonAppro
        Dim objBA2 As BonAppro
        Dim nid As Long

        'I - Création d'une commande Client
        '=========================
        objBA = New BonAppro(m_oFourn)
        objBA.oTransporteur.nom = "Transport Rault"
        objBA.oTransporteur.rs = "Transport Rault"
        objBA.oTransporteur.AdresseLivraison.nom = "Transport Rault"
        objBA.oTransporteur.AdresseLivraison.rue1 = "Penhouet Maro"
        objBA.oTransporteur.AdresseLivraison.rue2 = "Neulliac"
        objBA.oTransporteur.AdresseLivraison.cp = "56300"
        objBA.oTransporteur.AdresseLivraison.ville = "Penhouet Maro"
        objBA.oTransporteur.AdresseLivraison.tel = "123456"
        objBA.oTransporteur.AdresseLivraison.fax = "789012"
        objBA.oTransporteur.AdresseLivraison.port = "345678"
        objBA.oTransporteur.AdresseLivraison.Email = "Rault@wanandoo.fr"

        objBA.dateCommande = CDate("06/02/1964")
        objBA.caracteristiqueTiers.banque = "BANQUE"
        objBA.caracteristiqueTiers.rib1 = "15589"
        objBA.caracteristiqueTiers.rib2 = "35148"
        objBA.caracteristiqueTiers.rib3 = "04128612144"
        objBA.caracteristiqueTiers.rib4 = "05"
        objBA.caracteristiqueTiers.AdresseLivraison.nom = "Marc Collin"
        objBA.caracteristiqueTiers.AdresseLivraison.rue1 = "La Mettrie"
        objBA.caracteristiqueTiers.AdresseLivraison.rue2 = "2eme Etage"
        objBA.caracteristiqueTiers.AdresseLivraison.cp = "35250"
        objBA.caracteristiqueTiers.AdresseLivraison.ville = "chasné sur illet"
        objBA.caracteristiqueTiers.AdresseLivraison.tel = "0299555299"
        objBA.caracteristiqueTiers.AdresseLivraison.fax = "0299555277"
        objBA.caracteristiqueTiers.AdresseLivraison.port = "0680667189"
        objBA.caracteristiqueTiers.AdresseLivraison.Email = "contact@marccollin.com"
        objBA.caracteristiqueTiers.AdresseFacturation.nom = "Marc Collin" & "Fact"
        objBA.caracteristiqueTiers.AdresseFacturation.rue1 = "La Mettrie" & "Fact"
        objBA.caracteristiqueTiers.AdresseFacturation.rue2 = "2eme Etage" & "Fact"
        objBA.caracteristiqueTiers.AdresseFacturation.cp = "35250" & "Fact"
        objBA.caracteristiqueTiers.AdresseFacturation.ville = "chasné sur illet" & "Fact"
        objBA.caracteristiqueTiers.AdresseFacturation.tel = "0299555299" & "Fact"
        objBA.caracteristiqueTiers.AdresseFacturation.fax = "0299555277" & "Fact"
        objBA.caracteristiqueTiers.AdresseFacturation.port = "0680667189" & "Fact"
        objBA.caracteristiqueTiers.AdresseFacturation.Email = "contact@marccollin.com" & "Fact"
        objBA.caracteristiqueTiers.bAdressesIdentiques = False

        objBA.refLivraison = "BL0001"
        'Test des indicateurs Avant le Save
        Assert.IsTrue(objBA.bNew)
        Assert.IsTrue(objBA.bUpdated)
        Assert.IsFalse(objBA.bDeleted)
        'Save
        Assert.IsTrue(objBA.save(), "Insert" & objBA.getErreur)
        Assert.IsTrue((objBA.id <> 0), "Id Apres le Save doit être différent de 0")
        'Test des indicateurs Après le Save
        Assert.IsFalse(objBA.bNew, "bNew apres insert")
        Assert.IsFalse(objBA.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objBA.bDeleted, "bDeleted apres insert")

        nid = objBA.id
        'II - Rechargement d'une Commande
        '===============================
        objBA2 = New BonAppro(m_oFourn)
        Assert.IsTrue(objBA2.load(nid), "Load de la commande Client " & nid & ":" & objBA2.getErreur())
        Assert.IsTrue(objBA2.oTiers.load(), "chargement du client")
        Assert.AreEqual("BL0001", objBA2.refLivraison)
        Assert.IsTrue(objBA.Equals(objBA2), "Différents")
        'III - Modification d'une commande
        '=================================
        ' Modification de la commande
        objBA2.oTransporteur.nom = "LE CALVEZ"
        objBA2.oTransporteur.rs = "LE CALVEZ"
        objBA2.oTransporteur.AdresseLivraison.nom = "LE CALVEZ"
        objBA2.oTransporteur.AdresseLivraison.ville = "CHATEAU GIRON"
        objBA2.oTransporteur.AdresseLivraison.cp = "35"
        objBA2.changeEtat(vncEnums.vncActionEtatCommande.vncActionBAAnnLivrer)
        Assert.IsTrue(objBA2.etat.codeEtat = vncEnums.vncEtatCommande.vncBAEnCours)
        'objBA2.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        'Assert.IsTrue(objBA2.etat.codeEtat = vncEnums.vncEtatCommande.vncBAEnCours)
        objBA2.changeEtat(vncEnums.vncActionEtatCommande.vncActionBALivrer)
        Assert.IsTrue(objBA2.etat.codeEtat = vncEnums.vncEtatCommande.vncBALivre)
        objBA2.dateValidation = "10/07/2004"
        objBA2.caracteristiqueTiers.bAdressesIdentiques = True
        objBA2.refLivraison = "BL0002"

        'Test des indicateurs Avant le Save
        Assert.IsFalse(objBA2.bNew)
        Assert.IsTrue(objBA2.bUpdated)
        Assert.IsFalse(objBA2.bDeleted)
        'Save
        Assert.IsTrue(objBA2.save(), "Update" & objBA2.getErreur)
        'Test des indicateurs Après le Save
        Assert.IsFalse(objBA2.bNew, "bNew apres Update")
        Assert.IsFalse(objBA2.bUpdated, "bUpdated apres Update")
        Assert.IsFalse(objBA2.bDeleted, "bDeleted apres Update")
        'Rechargement de l'objet
        nid = objBA2.id
        objBA = New BonAppro(m_oFourn)
        Assert.IsTrue(objBA.load(nid), "Load")
        Assert.IsTrue(objBA.oTiers.load(), "Load Du client après Update")
        Assert.IsTrue(objBA.Equals(objBA2), "Apres Update , Equals")

        'IV - Suppression de la commande
        '=================================
        ' Modification de la commande
        objBA.bDeleted = True
        Assert.IsTrue(objBA.save(), "Delete" & objBA.getErreur())
        'Rechargement dans un autre objet
        objBA2 = New BonAppro(m_oFourn)
        'Assert.IsFalse(objBA2.load(nid), "Load")
    End Sub
    <TestMethod()> Public Sub T16_DB()
        Dim objLG As LgCommande

        Dim objBA As BonAppro
        Dim objLgCMD As LgCommande
        Dim nid As Long
        Dim strNumLg As String

        objBA = New BonAppro(m_oFourn)
        objBA.DuppliqueCaracteristiqueTiers()
        Assert.IsTrue(objBA.save(), "Creation de BA" & objBA.getErreur())

        ''Chargement de la collection des lignes vide
        Assert.AreEqual(objBA.colLignes.Count, 0, "Collection non vide")

        ''Ajout de 2 lignes
        objLgCMD = objBA.AjouteLigne(objBA.getNextNumLg, m_oProduit, 12, 20, True, 120, (120 * 1.196))

        Assert.IsTrue(Not objLgCMD Is Nothing, "Ajout OPRD1")
        Assert.AreEqual(objLgCMD.num, objBA.colLignes.Count * 10, "Num1")
        objLgCMD = objBA.AjouteLigne(objBA.getNextNumLg, m_oProduit, 12, 20, True, 125, (125 * 1.196))
        Assert.IsTrue(Not objLgCMD Is Nothing, "Ajout OPRD2")
        Assert.IsTrue(objLgCMD.num.Equals(objBA.colLignes.Count * 10), "Num2")
        ''Sauvegarde du BA
        objBA.totalTTC = 12
        Assert.IsTrue(objBA.save(), "objBA.Save" & objBA.getErreur())

        ''Rechargement de la commande et de ses lignes
        nid = objBA.id
        objBA = New BonAppro(m_oFourn)
        Assert.IsTrue(objBA.load(nid), "ObjBA.load")
        Assert.IsTrue(objBA.loadcolLignes(), "objBA.loadcolLignes")
        Assert.AreEqual(objBA.colLignes.Count, 2, "colLignes.count ")
        Assert.AreEqual(objBA.totalTTC, 12, "ObjCMD.totalTTC")

        ''Ajout d'une ligne 
        Assert.IsTrue(Not objBA.AjouteLigne(objBA.getNextNumLg, m_oProduit, 113, 20) Is Nothing, "oclt.AjouteLg 3")
        objLgCMD = objBA.colLignes(3)
        strNumLg = objLgCMD.num
        ''Sauvegarde du BA
        Assert.IsTrue(objBA.save(), "ObjBA.Save")
        ''Rechargement du BA
        nid = objBA.id
        objBA = New BonAppro(m_oFourn)

        Assert.IsTrue(objBA.load(nid), "objBA.load")
        Assert.AreEqual(objBA.colLignes.Count, 3, "collignes.count ")
        'Suppression d'une ligne 
        objBA.supprimeLigne(strNumLg)
        ''Sauvegarde du BA
        Assert.IsTrue(objBA.save(), "BA.Save")
        ''Rechargement du BA
        nid = objBA.id
        objBA = New BonAppro(m_oFourn)
        Assert.IsTrue(objBA.load(nid), "BA.load")
        Assert.AreEqual(objBA.colLignes.Count, 2, "BA.count ")

        'Maj d'une ligne de la precommande
        objLG = objBA.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteCommande = 150
        objLG.prixU = 15
        ''Sauvegarde de la precommande
        Assert.IsTrue(objBA.save(), "OBA.Save")
        'Rechargement du client et de sa precommande
        nid = objBA.id
        objBA = New BonAppro(m_oFourn)
        Assert.IsTrue(objBA.load(nid), "OCLT.load")
        Assert.AreEqual(objBA.colLignes.Count, 2, "Precommande.count ")
        objLG = objBA.colLignes(1)
        Assert.AreEqual(objLG.qteCommande, 150)
        Assert.AreEqual(objLG.prixU, 15)

        'Maj d'une ligne de la Commande (Qte Livrée)
        objLG = objBA.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteLiv = 100
        objLG.prixU = 16
        ''Sauvegarde 
        Assert.IsTrue(objBA.save(), "OCMD.Save")
        'Rechargement 
        nid = objBA.id
        objBA = New BonAppro(m_oFourn)
        Assert.IsTrue(objBA.load(nid), "OCMD.load")
        Assert.AreEqual(objBA.colLignes.Count, 2, "ColLignes.count ")
        objLG = objBA.colLignes(1)
        Assert.AreEqual(objLG.qteLiv, 100)
        Assert.AreEqual(objLG.prixU, 16)

        'Maj d'une ligne de la Commande (Qte Facturée)
        objLG = objBA.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteLiv = 101
        objLG.prixU = 17
        ''Sauvegarde 
        Assert.IsTrue(objBA.save(), "OCMD.Save")
        'Rechargement 
        nid = objBA.id
        objBA = New BonAppro(m_oFourn)
        Assert.IsTrue(objBA.load(nid), "OCMD.load")
        Assert.AreEqual(objBA.colLignes.Count, 2, "ColLignes.count ")
        objLG = objBA.colLignes(1)
        Assert.AreEqual(objLG.qteLiv, 101)
        Assert.AreEqual(objLG.prixU, 17)

        objBA.bDeleted = True
        Assert.IsTrue(objBA.save(), "BA.delete")

    End Sub
    <TestMethod()> Public Sub T51_ListeCriteres()
        Dim colBA As Collection
        Dim objBA As BonAppro
        Dim objBA1 As BonAppro
        Dim objBA2 As BonAppro
        Dim objFRN1 As Fournisseur
        Dim objFRN2 As Fournisseur
        Dim strCodeBA As String

        objFRN1 = New Fournisseur("FRN1" & Now(), "Albert")
        objFRN1.rs = objFRN1.nom
        Assert.IsTrue(objFRN1.Save())
        objFRN2 = New Fournisseur("FRN2" & Now(), "Arthur")
        objFRN2.rs = objFRN2.nom
        Assert.IsTrue(objFRN2.Save())


        objBA1 = New BonAppro(objFRN1)
        Assert.IsTrue(objBA1.save())
        objBA2 = New BonAppro(objFRN2)
        Assert.IsTrue(objBA2.save())
        strCodeBA = objBA2.code

        'I - Liste Simple
        '================
        colBA = BonAppro.getListe()
        Assert.IsTrue(colBA.Count >= 2, "Liste Complête Col.count >=2")
        For Each objBA In colBA
            Assert.IsTrue(objBA.oTiers.id <> 0, "oTiers.ID <> 0")
            Assert.IsTrue(objBA.id <> 0, "ID <> 0")
            Assert.IsTrue(objBA.bResume, "bResume = True")
            Assert.IsTrue(objBA.load(), "Load(" & objBA.id & ")" & objBA.getErreur())
        Next objBA

        'II- Liste sur le code 
        '=============================
        'a) Caractère générique
        colBA = BonAppro.getListe(Left(strCodeBA, 2) & "%")
        Assert.IsTrue(colBA.Count > 0, "Liste sur le code Col.count > 0")
        For Each objBA In colBA
            Assert.IsTrue(Left(objBA.code, 2) = Left(strCodeBA, 2), "ListeCode Mauvais Code " & objBA.code)
        Next objBA

        'b) sans Caractère générique
        colBA = BonAppro.getListe(strCodeBA)
        Assert.IsTrue(colBA.Count > 0, "Col.count > 0" & Client.getErreur)
        For Each objBA In colBA
            Assert.IsTrue(objBA.code = strCodeBA, "ListeCode Mauvais Code " & objBA.code)
        Next objBA

        'III- Liste sur le nom
        '=============================
        'a) Caractère générique
        colBA = BonAppro.getListe(, "A%")
        Assert.IsTrue(colBA.Count > 0, "Nom1:Col.count > 0" & Client.getErreur)
        For Each objBA In colBA
            objBA.load()
            Assert.IsTrue(Left(objBA.oTiers.nom, 1) = "A", "ListeNom Mauvais Nom " & objBA.oTiers.nom)
        Next objBA
        'b) sans Caractère générique
        colBA = BonAppro.getListe(, "Albert")
        Assert.IsTrue(colBA.Count > 0, "Nom2:Col.count > 0")
        For Each objBA In colBA
            Assert.IsTrue(objBA.load)
            Assert.IsTrue(objBA.oTiers.nom = "Albert", "ListeNom Mauvais Nom " & objBA.oTiers.nom)
        Next objBA

        'V - Critères Croisés
        '====================
        'a) Avec Résultat
        colBA = BonAppro.getListe(Left(strCodeBA, 2) & "%", "A%")
        Assert.IsTrue(colBA.Count > 0, "Col.count > 0")
        For Each objBA In colBA
            objBA.load()
            Assert.IsTrue(Left(objBA.code, 2) = Left(strCodeBA, 2), "Mauvais Code " & objBA.code)
            Assert.IsTrue(Left(objBA.oTiers.nom, 1) = "A", "Mauvais Nom " & objBA.oTiers.nom)
        Next objBA
        'b) sans résultat
        colBA = BonAppro.getListe(strCodeBA, "XXX%")
        Assert.IsTrue(colBA.Count = 0, "Col.count = 0")

        'Nettoyage


        objBA1.bDeleted = True
        Assert.IsTrue(objBA1.save())
        objBA2.bDeleted = True
        Assert.IsTrue(objBA2.save())

        objFRN1.bDeleted = True
        Assert.IsTrue(objFRN1.Save())
        objFRN2.bDeleted = True
        Assert.IsTrue(objFRN2.Save())

    End Sub

    <TestMethod()> Public Sub T60_BA_AnnulLiv()
        Dim nidBA As Integer
        Dim objBA As BonAppro
        Dim objLg As LgCommande


        'Creation d'un commande
        objBA = New BonAppro(m_oFourn)
        objBA.AjouteLigne("10", m_oProduit, 10, 0)
        objBA.save()
        nidBA = objBA.id
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")

        objBA.createandload(nidBA)
        objBA.loadcolLignes()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")
        'Livraison de la commande
        For Each objLg In objBA.colLignes
            objLg.qteLiv = objLg.qteCommande
        Next objLg
        objBA.changeEtat(vncEnums.vncActionEtatCommande.vncActionBALivrer)
        objBA.save()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")
        objBA.createandload(nidBA)
        objBA.loadcolLignes()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")

        'Annulation de la livraison
        For Each objLg In objBA.colLignes
            objLg.qteLiv = 0
        Next objLg
        objBA.changeEtat(vncEnums.vncActionEtatCommande.vncActionBAAnnLivrer)
        objBA.save()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")
        objBA.createandload(nidBA)
        objBA.loadcolLignes()
        Assert.IsTrue(objBA.colLignes.Count = 1, "Il n'y a plus de lignes dans la commandes")


        objBA.bDeleted = True
        objBA.save()
    End Sub

    'Test l'incrémenation des codes
    <TestMethod()> Public Sub T70_GetNextCode()

        Dim obj1 As New BonAppro(m_oFourn)
        Dim obj2 As New BonAppro(m_oFourn)

        obj1.save()
        obj2.save()
        Assert.AreEqual(obj2.code, (obj1.code + 1).ToString())
        Assert.AreNotEqual(0, obj1.code)

        obj1.bDeleted = True
        obj1.save()
        obj2.bDeleted = True
        obj2.save()


    End Sub

End Class



