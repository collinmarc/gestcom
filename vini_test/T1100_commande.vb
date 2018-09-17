'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.IO


<TestClass()> Public Class T1100_commande
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
    <TestMethod()> Public Sub T010_Object()
        Dim objCMD As CommandeClient
        Dim objCMD2 As CommandeClient

        objCMD = New CommandeClient(m_oClient)
        'Initialisation par défaut
        Assert.IsTrue(objCMD.dateCommande.ToShortDateString.Equals(Now().ToShortDateString), "Date de commande")
        Assert.IsTrue(objCMD.dateLivraison.ToShortDateString.Equals(DateAdd(DateInterval.Day, 1, Now()).ToShortDateString), "Date de commande")
        Assert.AreEqual(Transporteur.TransporteurDefault.nom, objCMD.oTransporteur.nom, "Nom du transporteur")
        Assert.AreEqual(Transporteur.TransporteurDefault.AdresseLivraison.rue1, objCMD.oTransporteur.AdresseLivraison.rue1, "Rue1 du transporteur")
        Assert.AreEqual(Transporteur.TransporteurDefault.AdresseLivraison.rue2, objCMD.oTransporteur.AdresseLivraison.rue2, "Rue2 du transporteur")
        Assert.AreEqual(Transporteur.TransporteurDefault.AdresseLivraison.cp, objCMD.oTransporteur.AdresseLivraison.cp, "cp du transporteur")
        Assert.AreEqual(Transporteur.TransporteurDefault.AdresseLivraison.ville, objCMD.oTransporteur.AdresseLivraison.ville, "Ville du transporteur")
        Assert.AreEqual(Transporteur.TransporteurDefault.AdresseLivraison.tel, objCMD.oTransporteur.AdresseLivraison.tel, "Tel du transporteur")
        Assert.AreEqual(Transporteur.TransporteurDefault.AdresseLivraison.fax, objCMD.oTransporteur.AdresseLivraison.fax, "Fax du transporteur")
        Assert.AreEqual(Transporteur.TransporteurDefault.AdresseLivraison.port, objCMD.oTransporteur.AdresseLivraison.port, "port du transporteur")
        Assert.AreEqual(Transporteur.TransporteurDefault.AdresseLivraison.Email, objCMD.oTransporteur.AdresseLivraison.Email, "email du transporteur")
        Assert.AreEqual(0D, objCMD.qteColis, "quantite de colis")
        Assert.AreEqual(0D, objCMD.qtePalettesNonPreparees, "quantite palettes non preparées")
        Assert.AreEqual(0D, objCMD.qtePalettesPreparees, "quantite palettes preparées")
        Assert.AreEqual(0D, objCMD.poids, "poids")
        Assert.AreEqual(CDec(Param.getConstante("CST_PU_PALL_NONPREP")), CDec(objCMD.puPalettesNonPreparees), "PU palettes non preparées")
        Assert.AreEqual(CDec(Param.getConstante("CST_PU_PALL_PREP")), CDec(objCMD.puPalettesPreparees), "PU palettes preparées")
        Assert.AreEqual(True, objCMD.bFactTransport, "Facture de transport")
        Assert.AreEqual(0L, objCMD.idFactTransport, "idFacture de transport")

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
        objCMD.qteColis = 10.5
        objCMD.qtePalettesNonPreparees = 11.5
        objCMD.qtePalettesPreparees = 12.5
        objCMD.poids = 13.5
        objCMD.montantTransport = 1234.56
        objCMD.puPalettesNonPreparees = 14.5
        objCMD.puPalettesPreparees = 15.5
        objCMD.bFactTransport = True
        objCMD.idFactTransport = 1234

        Assert.AreEqual(objCMD.code, "CODE")
        Assert.AreEqual(objCMD.caracteristiqueTiers.banque, "BANQUE")
        Assert.AreEqual(objCMD.caracteristiqueTiers.rib1, "RIB1")
        Assert.AreEqual(objCMD.caracteristiqueTiers.rib2, "RIB2")
        Assert.AreEqual(objCMD.caracteristiqueTiers.rib3, "RIB3")
        Assert.AreEqual(objCMD.caracteristiqueTiers.rib4, "RIB4")
        Assert.AreEqual(objCMD.dateLivraison, CDate("06/02/1964"))
        Assert.AreEqual(objCMD.dateEnlevement, CDate("31/07/1964"))
        Assert.AreEqual(objCMD.refLivraison, "BL0003")
        Assert.AreEqual(objCMD.qteColis, 10.5D)
        Assert.AreEqual(objCMD.qtePalettesNonPreparees, 11.5D)
        Assert.AreEqual(objCMD.qtePalettesPreparees, 12.5D)
        Assert.AreEqual(objCMD.poids, 13.5D)
        Assert.AreEqual(objCMD.montantTransport, 1234.56D)
        Assert.AreEqual(objCMD.puPalettesNonPreparees, 14.5D)
        Assert.AreEqual(objCMD.puPalettesPreparees, 15.5D)
        Assert.AreEqual(objCMD.bFactTransport, True)
        Assert.AreEqual(objCMD.idFactTransport, 1234L)


        'Test des indicateurs
        Assert.IsTrue(objCMD.bNew)
        Assert.IsTrue(objCMD.bUpdated)
        Assert.IsFalse(objCMD.bDeleted)

        Assert.IsTrue(objCMD.Equals(objCMD), "Egal à Lui même")
        objCMD2 = New CommandeClient(m_oClient)

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
        objCMD2.qteColis = 10.5
        objCMD2.qtePalettesNonPreparees = 11.5
        objCMD2.qtePalettesPreparees = 12.5
        objCMD2.poids = 13.5
        objCMD2.montantTransport = 1234.56
        objCMD2.puPalettesNonPreparees = 14.5
        objCMD2.puPalettesPreparees = 15.5
        objCMD2.bFactTransport = True
        objCMD2.idFactTransport = 1234

        Assert.IsTrue(objCMD.Equals(objCMD2), "Egal à un semblable")
        objCMD2.caracteristiqueTiers.banque = "Deuxième banque"
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.caracteristiqueTiers.banque = objCMD.caracteristiqueTiers.banque
        objCMD2.dateLivraison = CDate("07/02/1964")
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.dateLivraison = objCMD.dateLivraison
        objCMD2.refLivraison = "BL0002"
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.refLivraison = objCMD.refLivraison
        objCMD2.montantTransport = 9876.67
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.montantTransport = objCMD.montantTransport
        objCMD2.qteColis = 9876.67
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.qteColis = objCMD.qteColis
        objCMD2.qtePalettesNonPreparees = 9876.67
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.qtePalettesNonPreparees = objCMD.qtePalettesNonPreparees
        objCMD2.qtePalettesPreparees = 9876.67
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.qtePalettesPreparees = objCMD.qtePalettesPreparees
        objCMD2.poids = 9876.67
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.poids = objCMD.poids
        objCMD2.puPalettesNonPreparees = 9876.67
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.puPalettesNonPreparees = objCMD.puPalettesNonPreparees
        objCMD2.puPalettesPreparees = 9876.67
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.puPalettesPreparees = objCMD.puPalettesPreparees
        objCMD2.bFactTransport = False
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        objCMD2.bFactTransport = objCMD.bFactTransport
        objCMD2.idFactTransport = 0
        Assert.IsFalse(objCMD.Equals(objCMD2), "Egal à un Différent")
        Dim obj As Object
        Assert.IsFalse(objCMD.Equals(obj), "Egal autrecjhose")


    End Sub
    <TestMethod()> Public Sub T011_ObjectLigneCommande()
        Dim objLgCMD As LgCommande
        Dim objLgCMD2 As LgCommande
        Dim nTaux As Double
        Dim oParam As Param


        objLgCMD = New LgCommande(15)
        objLgCMD.num = 10
        objLgCMD.oProduit = m_oProduit
        objLgCMD.qteCommande = 10.5
        objLgCMD.qteLiv = 11.5
        objLgCMD.qteFact = 12.5
        objLgCMD.prixU = 15
        objLgCMD.prixHT = 20
        objLgCMD.bGratuit = True

        Assert.AreEqual(objLgCMD.num, 10)
        Assert.AreEqual(objLgCMD.oProduit, m_oProduit)
        Assert.AreEqual(objLgCMD.qteCommande, 10.5D)
        Assert.AreEqual(objLgCMD.qteLiv, 11.5D)
        Assert.AreEqual(objLgCMD.qteFact, 12.5D)
        Assert.AreEqual(objLgCMD.prixU, 15D)
        Assert.AreEqual(objLgCMD.prixHT, 20D)
        Assert.AreEqual(objLgCMD.bGratuit, True)


        'Test des indicateurs
        Assert.IsTrue(objLgCMD.bNew)
        Assert.IsTrue(objLgCMD.bUpdated)
        Assert.IsFalse(objLgCMD.bDeleted)

        Assert.IsTrue(objLgCMD.Equals(objLgCMD), "Egal à Lui même")

        objLgCMD2 = New LgCommande(15)
        objLgCMD2.num = 10
        objLgCMD2.oProduit = m_oProduit
        objLgCMD2.qteCommande = 10.5
        objLgCMD2.qteLiv = 11.5
        objLgCMD2.qteFact = 12.5
        objLgCMD2.prixU = 15
        objLgCMD2.prixHT = 20
        objLgCMD2.bGratuit = True

        Assert.IsTrue(objLgCMD.Equals(objLgCMD2), "Egal à un semblable")
        objLgCMD.bGratuit = False
        Assert.IsFalse(objLgCMD.Equals(objLgCMD2), "Egal à un Différent")
        Dim obj As Object
        Assert.IsFalse(objLgCMD.Equals(obj), "Egal autrechose")

        oParam = Param.TVAdefaut
        m_oProduit.idTVA = oParam.id
        m_oProduit.save()
        nTaux = CDbl(oParam.valeur)
        Console.Out.WriteLine("Ntaux =" & nTaux)
        objLgCMD2 = New LgCommande(15)
        objLgCMD2.num = 10
        objLgCMD2.oProduit = m_oProduit
        objLgCMD2.qteCommande = 15
        objLgCMD2.prixU = 20
        objLgCMD2.prixHT = 0
        objLgCMD2.bGratuit = False
        Assert.IsTrue(objLgCMD2.calculPrixTotal(), "Calcul du prix Total")
        Assert.IsTrue(objLgCMD2.prixHT = objLgCMD2.prixU * objLgCMD2.qteCommande)
        Assert.AreEqual(CDec(objLgCMD2.prixHT * (1 + (nTaux / 100))), CDec(objLgCMD2.prixTTC), "Calcul Prix TTC")
        objLgCMD2.bGratuit = True
        Assert.IsTrue(objLgCMD2.calculPrixTotal(), "Calcul du prix Total gratuit")
        Assert.IsTrue(objLgCMD2.prixHT = 0)
        'Calcul avec un Taux de TVA bas
        objLgCMD2.oProduit.idTVA() = Param.colTVA("BAS").id
        Dim nTaux2 As Decimal = CType(Param.colTVA("BAS"), Param).valeur
        Assert.IsTrue(objLgCMD2.calculPrixTotal(), "Calcul du prix Total")
        Assert.IsTrue(objLgCMD2.prixHT = objLgCMD2.prixU * objLgCMD2.qteCommande)
        Assert.AreEqual(CDec(objLgCMD2.prixHT * (1 + (nTaux2 / 100))), CDec(objLgCMD2.prixTTC), "Calcul Prix TTC avec TVA BAS")

        'Calcul du prix avec la qte Livrée
        objLgCMD2.qteLiv = 10
        objLgCMD2.prixU = 10
        objLgCMD2.bGratuit = False
        Assert.IsTrue(objLgCMD2.calculPrixTotal(), "Calcul du prix Total")
        Assert.IsTrue(objLgCMD2.prixHT = objLgCMD2.prixU * objLgCMD2.qteLiv)

        'Calcul du prix avec la qte Livrée
        objLgCMD2.qteFact = 20
        objLgCMD2.prixU = 30
        objLgCMD2.bGratuit = False
        Assert.IsTrue(objLgCMD2.calculPrixTotal(), "Calcul du prix Total")
        Assert.IsTrue(objLgCMD2.prixHT = objLgCMD2.prixU * objLgCMD2.qteFact)

        objLgCMD2.qteCommande = 1
        objLgCMD2.qteLiv = 0
        objLgCMD2.qteFact = 0
        Assert.IsTrue(Not objLgCMD2.estLivree())
        objLgCMD2.qteCommande = 1
        objLgCMD2.qteLiv = 2
        objLgCMD2.qteFact = 0
        Assert.IsTrue(objLgCMD2.estLivree())
        objLgCMD2.qteCommande = 1
        objLgCMD2.qteLiv = 2
        objLgCMD2.qteFact = 3
        Assert.IsTrue(objLgCMD2.estLivree())
        objLgCMD2.qteCommande = 1
        objLgCMD2.qteLiv = 0
        objLgCMD2.qteFact = 3
        Assert.IsTrue(objLgCMD2.estLivree())
        objLgCMD2.qteCommande = 0
        objLgCMD2.qteLiv = 0
        objLgCMD2.qteFact = 0
        Assert.IsTrue(objLgCMD2.estLivree())


    End Sub
    <TestMethod()> Public Sub T012_GestLignes()
        Dim objComm As CommandeClient
        Dim objLg As LgCommande



        objComm = New CommandeClient(m_oClient)

        objLg = objComm.AjouteLigne("10", m_oProduit, 10, 10)
        Assert.IsTrue(objComm.colLignes.Count = 1)
        Assert.IsTrue(objComm.totalHT = 10 * 10)
        Assert.AreEqual(objComm.totalTTC, CDec(objComm.totalHT * 1.2))
        Assert.IsTrue(Not objLg Is Nothing, "Ajout de Ligne impossible")
        Assert.IsTrue(objLg.num = objComm.colLignes.Count * 10)
        Assert.AreEqual(0, objLg.idSCmd, "SousCommande est null")

        objLg = objComm.AjouteLigne("20", m_oProduit, 20, 10)
        Assert.IsTrue(objComm.colLignes.Count = 2)
        Assert.IsTrue(objComm.totalHT = (10 * 10) + (20 * 10))
        Assert.AreEqual(objComm.totalTTC, CDec(objComm.totalHT * 1.2))
        Assert.IsTrue(Not objLg Is Nothing, "Ajout de Ligne impossible")
        Assert.IsTrue(objLg.num = objComm.colLignes.Count * 10)
        objLg = objComm.AjouteLigne("30", m_oProduit, 30, 10)
        Assert.IsTrue(objComm.colLignes.Count = 3)
        Assert.IsTrue(objComm.totalHT = (10 * 10) + (20 * 10) + (30 * 10))
        Assert.IsTrue(objComm.totalTTC = (objComm.totalHT * 1.2))
        Assert.IsTrue(Not objLg Is Nothing, "Ajout de Ligne impossible")
        Assert.IsTrue(objLg.num = objComm.colLignes.Count * 10)

        Assert.IsTrue(Not objComm.estEntierementLivree())
        For Each objLg In objComm.colLignes
            objLg.qteLiv = objLg.qteCommande
        Next objLg
        Assert.IsTrue(objComm.estEntierementLivree())




    End Sub 'T12
    <TestMethod()> Public Sub T013_commClient()
        Dim objComm As CommandeClient

        objComm = New CommandeClient(m_oClient)
        objComm.dateCommande = CDate("06/02/1964")
        objComm.oTransporteur.AdresseLivraison.nom = "Transport Rault"
        objComm.caracteristiqueTiers.AdresseLivraison.nom = "AdresseLivr"
        objComm.typeCommande = vncEnums.vncTypeCommande.vncCmdClientDirecte
        objComm.typeTransport = vncEnums.vncTypeTransport.vncTrpFranco

        Assert.AreEqual(objComm.oTransporteur.AdresseLivraison.nom, "Transport Rault")
        Assert.AreEqual(objComm.caracteristiqueTiers.AdresseLivraison.nom, "AdresseLivr")
        Assert.AreEqual(objComm.typeCommande, vncEnums.vncTypeCommande.vncCmdClientDirecte)
        Assert.AreEqual(objComm.typeTransport, vncEnums.vncTypeTransport.vncTrpFranco)



    End Sub 'T13
    <TestMethod()> Public Sub T015_DB()
        Dim objCMDCLT As CommandeClient
        Dim objCMDCLT2 As CommandeClient
        Dim nid As Long

        'I - Création d'une commande Client
        '=========================
        objCMDCLT = New CommandeClient(m_oClient)
        objCMDCLT.oTransporteur.nom = "Transport Rault"
        objCMDCLT.oTransporteur.rs = "Transport Rault"
        objCMDCLT.oTransporteur.AdresseLivraison.nom = "Transport Rault"
        objCMDCLT.oTransporteur.AdresseLivraison.rue1 = "Penhouet Maro"
        objCMDCLT.oTransporteur.AdresseLivraison.rue2 = "Neulliac"
        objCMDCLT.oTransporteur.AdresseLivraison.cp = "56300"
        objCMDCLT.oTransporteur.AdresseLivraison.ville = "Penhouet Maro"
        objCMDCLT.oTransporteur.AdresseLivraison.tel = "123456"
        objCMDCLT.oTransporteur.AdresseLivraison.fax = "789012"
        objCMDCLT.oTransporteur.AdresseLivraison.port = "345678"
        objCMDCLT.oTransporteur.AdresseLivraison.Email = "Rault@wanandoo.fr"

        objCMDCLT.dateCommande = CDate("06/02/1964")
        objCMDCLT.caracteristiqueTiers.banque = "BANQUE"
        objCMDCLT.caracteristiqueTiers.rib1 = "15589"
        objCMDCLT.caracteristiqueTiers.rib2 = "35148"
        objCMDCLT.caracteristiqueTiers.rib3 = "04128612144"
        objCMDCLT.caracteristiqueTiers.rib4 = "05"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.nom = "Marc Collin"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.rue1 = "La Mettrie"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.rue2 = "2eme Etage"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.cp = "35250"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.ville = "chasné sur illet"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.tel = "0299555299"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.fax = "0299555277"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.port = "0680667189"
        objCMDCLT.caracteristiqueTiers.AdresseLivraison.Email = "contact@marccollin.com"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.nom = "Marc Collin" & "Fact"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.rue1 = "La Mettrie" & "Fact"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.rue2 = "2eme Etage" & "Fact"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.cp = "35250" & "Fact"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.ville = "chasné sur illet" & "Fact"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.tel = "0299555299" & "Fact"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.fax = "0299555277" & "Fact"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.port = "0680667189" & "Fact"
        objCMDCLT.caracteristiqueTiers.AdresseFacturation.Email = "contact@marccollin.com" & "Fact"
        objCMDCLT.caracteristiqueTiers.idModeReglement = 33
        objCMDCLT.caracteristiqueTiers.bAdressesIdentiques = False
        objCMDCLT.etat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncEnCoursSaisie)
        objCMDCLT.typeCommande = vncEnums.vncTypeCommande.vncCmdClientDirecte
        objCMDCLT.typeTransport = vncEnums.vncTypeTransport.vncTrpFranco
        objCMDCLT.refLivraison = "BL0001"
        objCMDCLT.montantTransport = 12345.67
        objCMDCLT.qteColis = 123.1
        objCMDCLT.qtePalettesNonPreparees = 234.2
        objCMDCLT.qtePalettesPreparees = 345.3
        objCMDCLT.puPalettesNonPreparees = 12.1
        objCMDCLT.puPalettesPreparees = 23.1
        objCMDCLT.poids = 3.1
        objCMDCLT.bFactTransport = True
        objCMDCLT.idFactTransport = 1234
        'Test des indicateurs Avant le Save
        Assert.IsTrue(objCMDCLT.bNew)
        Assert.IsTrue(objCMDCLT.bUpdated)
        Assert.IsFalse(objCMDCLT.bDeleted)
        'Save
        Assert.IsTrue(objCMDCLT.save(), "Insert" & objCMDCLT.getErreur)
        Assert.IsTrue((objCMDCLT.id <> 0), "Id Apres le Save doit être différent de 0")
        'Test des indicateurs Après le Save
        Assert.IsFalse(objCMDCLT.bNew, "bNew apres insert")
        Assert.IsFalse(objCMDCLT.bUpdated, "bUpdated apres insert")
        Assert.IsFalse(objCMDCLT.bDeleted, "bDeleted apres insert")

        nid = objCMDCLT.id
        'II - Rechargement d'une Commande
        '===============================
        objCMDCLT2 = CommandeClient.createandload(nid)
        Assert.IsTrue(objCMDCLT2.load(nid), "Load de la commande Client " & nid & ":" & objCMDCLT2.getErreur())
        Assert.IsTrue(objCMDCLT2.oTiers.load(), "chargement du client")
        Assert.AreEqual(objCMDCLT2.refLivraison, "BL0001")
        Assert.IsTrue(objCMDCLT.Equals(objCMDCLT2), "Différents")

        'III - Modification d'une commande
        '=================================
        ' Modification de la commande
        objCMDCLT2.oTransporteur.AdresseLivraison.nom = "LE CALVEZ"
        objCMDCLT2.oTransporteur.nom = "LE CALVEZ"
        objCMDCLT2.oTransporteur.rs = "LE CALVEZ"
        objCMDCLT2.oTransporteur.AdresseLivraison.nom = "LE CALVEZ"
        objCMDCLT2.oTransporteur.AdresseLivraison.ville = "CHATEAU GIRON"
        objCMDCLT2.oTransporteur.AdresseLivraison.cp = "35"
        Assert.IsTrue(objCMDCLT2.etat.codeEtat = vncEnums.vncEtatCommande.vncEnCoursSaisie)
        objCMDCLT2.changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnValider)
        Assert.IsTrue(objCMDCLT2.etat.codeEtat = vncEnums.vncEtatCommande.vncEnCoursSaisie)
        objCMDCLT2.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        Assert.IsTrue(objCMDCLT2.etat.codeEtat = vncEnums.vncEtatCommande.vncValidee)
        objCMDCLT2.dateValidation = "10/07/2004"
        objCMDCLT2.caracteristiqueTiers.bAdressesIdentiques = True
        objCMDCLT2.refLivraison = "BL0002"
        objCMDCLT2.bFactTransport = False
        objCMDCLT2.idFactTransport = 5678


        'Test des indicateurs Avant le Save
        Assert.IsFalse(objCMDCLT2.bNew)
        Assert.IsTrue(objCMDCLT2.bUpdated)
        Assert.IsFalse(objCMDCLT2.bDeleted)
        'Save
        Assert.IsTrue(objCMDCLT2.save(), "Update" & objCMDCLT2.getErreur)
        'Test des indicateurs Après le Save
        Assert.IsFalse(objCMDCLT2.bNew, "bNew apres Update")
        Assert.IsFalse(objCMDCLT2.bUpdated, "bUpdated apres Update")
        Assert.IsFalse(objCMDCLT2.bDeleted, "bDeleted apres Update")
        'Rechargement de l'objet
        nid = objCMDCLT2.id
        objCMDCLT = New CommandeClient(m_oClient)
        Assert.IsTrue(objCMDCLT.load(nid), "Load")
        Assert.IsTrue(objCMDCLT.oTiers.load(), "Load Du client après Update")
        Assert.IsTrue(objCMDCLT.Equals(objCMDCLT2), "Apres Update , Equals")

        'IV - Suppression de la commande
        '=================================
        ' Modification de la commande
        objCMDCLT.bDeleted = True
        Assert.IsTrue(objCMDCLT.save(), "Delete" & objCMDCLT.getErreur())
    End Sub
    <TestMethod()> Public Sub T016_DB()
        Dim objLG As LgCommande

        Dim objCMD As CommandeClient
        Dim objLgCMD As LgCommande
        Dim nid As Long
        Dim strNumLg As String

        objCMD = New CommandeClient(m_oClient)
        objCMD.setTiers(m_oClient)
        objCMD.DuppliqueCaracteristiqueTiers()
        Assert.IsTrue(objCMD.save(), "Creation de Commande" & objCMD.getErreur())

        ''Chargement de la commande vide
        Assert.AreEqual(objCMD.colLignes.Count, 0, "Collection non vide")

        ''Ajout de 2 ligne de Commande
        objLgCMD = objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 12, 20, True, 120, (120 * 1.2))
        Assert.IsTrue(Not objLgCMD Is Nothing, "Ajout OPRD1")
        Assert.AreEqual(objLgCMD.num, objCMD.colLignes.Count * 10, "Num1")
        objLgCMD = objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 12, 20, True, 125, (125 * 1.2))
        Assert.IsTrue(Not objLgCMD Is Nothing, "Ajout OPRD2")
        Assert.IsTrue(objLgCMD.num.Equals(objCMD.colLignes.Count * 10), "Num2")
        ''Sauvegarde de la commande
        objCMD.totalTTC = 12
        Assert.IsTrue(objCMD.save(), "Ocmd.Save" & objCMD.getErreur())
        Assert.AreEqual(0, objLgCMD.idSCmd, "IDSouscommande")

        ''Rechargement de la commande et de ses lignes
        nid = objCMD.id
        objCMD = New CommandeClient(m_oClient)
        Assert.IsTrue(objCMD.load(nid), "ObjCMD.load")
        Assert.AreEqual(2, objCMD.colLignes.Count, "2 Lignes attendues")
        Assert.AreEqual(12D, objCMD.totalTTC, "ObjCMD.totalTTC")

        ''Ajout d'une ligne de commande
        Assert.IsTrue(Not objCMD.AjouteLigne(objCMD.getNextNumLg, m_oProduit, 113, 20) Is Nothing, "oclt.AjouteLg 3")
        objLG = objCMD.colLignes(3)
        strNumLg = objLG.num
        ''Sauvegarde de la commande
        Assert.IsTrue(objCMD.save(), "Oclt.Save")
        ''Rechargement de la commande
        nid = objCMD.id
        objCMD = New CommandeClient(m_oClient)

        Assert.IsTrue(objCMD.load(nid), "OCLT.load")
        Assert.AreEqual(objCMD.colLignes.Count, 3, "colLignes.count ")
        'Suppression d'une ligne de la commande
        objCMD.supprimeLigne(strNumLg)
        ''Sauvegarde de la pcommande
        Assert.IsTrue(objCMD.save(), "Oclt.Save")
        ''Rechargement de la commande
        nid = objCMD.id
        objCMD = New CommandeClient(m_oClient)
        Assert.IsTrue(objCMD.load(nid), "OCLT.load")
        Assert.AreEqual(objCMD.colLignes.Count, 2, "colLignes.count ")

        'Maj d'une ligne de la commande
        objLG = objCMD.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteCommande = 150
        objLG.prixU = 15.5
        ''Sauvegarde de la commande
        Assert.IsTrue(objCMD.save(), "Oclt.Save")
        'Rechargement de la commande
        nid = objCMD.id
        objCMD = New CommandeClient(m_oClient)
        Assert.IsTrue(objCMD.load(nid), "OCLT.load")
        Assert.AreEqual(objCMD.colLignes.Count, 2, "colLignes.count ")
        objLG = objCMD.colLignes(1)
        Assert.AreEqual(objLG.qteCommande, 150D)
        Assert.AreEqual(objLG.prixU, 15.5D)

        'Maj d'une ligne de la Commande (Qte Livrée)
        objLG = objCMD.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteLiv = 100
        objLG.prixU = 16
        ''Sauvegarde de la commande
        Assert.IsTrue(objCMD.save(), "OCMD.Save")
        'Rechargement de la commande
        nid = objCMD.id
        objCMD = New CommandeClient(m_oClient)
        Assert.IsTrue(objCMD.load(nid), "OCMD.load")
        Assert.AreEqual(objCMD.colLignes.Count, 2, "ColLignes.count ")
        objLG = objCMD.colLignes(1)
        Assert.AreEqual(objLG.qteLiv, 100D)
        Assert.AreEqual(objLG.prixU, 16D)

        'Maj d'une ligne de la Commande (Qte Facturée)
        objLG = objCMD.colLignes.Item(1)
        'sCode = objLgPRecom.codeProduit
        objLG.qteLiv = 101
        objLG.prixU = 17
        ''Sauvegarde de la commande
        Assert.IsTrue(objCMD.save(), "OCMD.Save")
        'Rechargement 
        nid = objCMD.id
        objCMD = New CommandeClient(m_oClient)
        Assert.IsTrue(objCMD.load(nid), "OCMD.load")
        Assert.AreEqual(objCMD.colLignes.Count, 2, "ColLignes.count ")
        objLG = objCMD.colLignes(1)
        Assert.AreEqual(objLG.qteLiv, 101D)
        Assert.AreEqual(objLG.prixU, 17D)

    End Sub

    <TestMethod()> Public Sub T020_MAJPRECOMMANDE()
        Dim objCmd As CommandeClient
        Dim objLGPRECOM As lgPrecomm



        'Creation de la precommande
        'Ajout 1 ligne Produit1
        m_oClient.ajouteLgPrecom(objProduit1, 10, 11)

        'Ajout 1 ligne Produit3
        objLGPRECOM = m_oClient.ajouteLgPrecom(objProduit3, 30, 31)
        objLGPRECOM.prixU = 32
        objLGPRECOM.dateDerniereCommande = "15/09/2004"
        objLGPRECOM.refDerniereCommande = "MAREF"
        Assert.IsTrue(m_oClient.save())

        'Creation d'une Commande
        objCmd = New CommandeClient(m_oClient)
        objCmd.dateCommande = "06/02/2000"
        objCmd.AjouteLigne("10", objProduit1, 21, 22, False)
        objCmd.AjouteLigne("20", objProduit2, 31, 32, False)
        objCmd.AjouteLigne("30", objProduit3, 41, 42, False)
        Assert.IsTrue(objCmd.save())

        'Rechargement de La PRecommande
        Assert.IsTrue(m_oClient.load())
        Assert.IsTrue(m_oClient.LoadPreCommande())


        '1re ligne de precommande
        objLGPRECOM = m_oClient.getLgPrecomByProductId(objProduit1.id)
        Assert.AreEqual(21D, objLGPRECOM.qteDern, "Derniere Qte commandée")
        Assert.AreEqual(CDbl(22D), objLGPRECOM.prixU, "Dern Prix U")
        Assert.AreEqual(objCmd.dateCommande, objLGPRECOM.dateDerniereCommande, "Date dernière commande")
        Assert.AreEqual(objCmd.code, objLGPRECOM.refDerniereCommande, "reference dernière commande")

        '2eme ligne de precommande
        Assert.IsTrue(m_oClient.lgPrecomExists(objProduit2.id), "La Ligne de Precommande n'a pas été ajoutée")
        objLGPRECOM = m_oClient.getLgPrecomByProductId(objProduit2.id)
        Assert.AreEqual(31D, objLGPRECOM.qteDern, "Derniere Qte commandée")
        Assert.AreEqual(CDbl(32D), objLGPRECOM.prixU, "Dern Prix U")
        Assert.AreEqual(objCmd.dateCommande, objLGPRECOM.dateDerniereCommande, "Date dernière commande")
        Assert.AreEqual(objCmd.code, objLGPRECOM.refDerniereCommande, "reference dernière commande")

        '3eme ligne de precommande 
        'Elle n'a pas changé car la date de dern commande est supérieure à la date de commande
        Assert.IsTrue(m_oClient.lgPrecomExists(objProduit3.id), "La Ligne de Precommande n'a pas été ajoutée")
        objLGPRECOM = m_oClient.getLgPrecomByProductId(objProduit3.id)
        Assert.AreEqual(31D, objLGPRECOM.qteDern, "Derniere Qte commandée")
        Assert.AreEqual(CDbl(32), objLGPRECOM.prixU, "Dern Prix U")
        Assert.AreEqual(CDate("15/09/2004"), objLGPRECOM.dateDerniereCommande, "Date dernière commande")
        Assert.AreEqual("MAREF", objLGPRECOM.refDerniereCommande, "reference dernière commande")

        objCmd.bDeleted = True
        Assert.IsTrue(objCmd.save())



    End Sub

    <TestMethod()> Public Sub T070_Gratuit()
        Dim objCMD As CommandeClient
        Dim objLg As LgCommande
        Dim nIdCmd As Long

        'Creation d'une Commande
        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = "06/02/2000"
        'Ajout d'une ligne gratuite
        objCMD.AjouteLigne("10", objProduit1, 21, 22, True)
        Assert.IsTrue(objCMD.save())
        nIdCmd = objCMD.id

        objCMD = CommandeClient.createandload(nIdCmd)
        objCMD.loadcolLignes()
        objLg = objCMD.colLignes(1)
        Assert.IsTrue(objLg.bGratuit, "La Ligne est gratuite")

        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save(), "Destruction de la commande")

    End Sub
    'Test la sauvegarde des champs long
    <TestMethod()> Public Sub T060_Champslongs()

        Dim obj As CommandeClient
        obj = New CommandeClient(m_oClient)
        obj.caracteristiqueTiers.AdresseLivraison.rue1 = "TEST".PadRight(500, "x")
        Assert.IsTrue(obj.save())

        obj.lettreVoiture = "TEST2".PadRight(500, "x")
        Assert.IsTrue(obj.save())

        obj.load()
        Assert.AreEqual(50, obj.lettreVoiture.Length)
        Assert.AreEqual("TEST2".PadRight(500, "x").Substring(0, 50), obj.lettreVoiture)

        obj.bDeleted = True
        obj.save()

    End Sub

    'Test l'incrémenation des codes
    <TestMethod()> Public Sub T070_GetNextCode()

        Dim obj1 As New CommandeClient(m_oClient)
        Dim obj2 As New CommandeClient(m_oClient)

        obj1.save()
        obj2.save()
        Assert.AreEqual(obj2.code, (obj1.code + 1).ToString())
        Assert.AreNotEqual(0, obj1.code)

        obj1.bDeleted = True
        obj1.save()
        obj2.bDeleted = True
        obj2.save()


    End Sub
    ''' <summary>
    ''' Test les champs commision dans les lignes de commandes
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T080_COMMISSION()

        Dim objCMD As CommandeClient
        Dim objLg As LgCommande
        Dim nIdCmd As Long

        'Creation d'une Commande
        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = "06/02/2000"
        'Ajout d'une ligne 
        objLg = objCMD.AjouteLigne("10", objProduit1, 21, 22, True)
        'Mise à jour des informations de commission
        objLg.TxComm = 8.53
        objLg.MtComm = 156.36
        Assert.IsTrue(objCMD.save())
        nIdCmd = objCMD.id

        objCMD = CommandeClient.createandload(nIdCmd)
        objCMD.loadcolLignes()
        objLg = objCMD.colLignes(1)
        Assert.AreEqual(8.53D, objLg.TxComm, "Tx de commssion différent")
        Assert.AreEqual(156.36D, objLg.MtComm, "Mt de commssion différent")

        objCMD.bDeleted = True
        Assert.IsTrue(objCMD.save(), "Destruction de la commande")
    End Sub
    ''' <summary>
    ''' Test les champs commision dans les lignes de commandes
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T080_CALCULCOMMISSION()

        Dim objCMD As CommandeClient
        Dim objLg As LgCommande
        Dim objParam As Param
        Dim objTaux As TauxComm


        'Creation d'un Taux de commission

        objParam = New Param
        objParam.load(m_oClient.idTypeClient)

        objTaux = New TauxComm(m_oFourn.id, objParam.code, 12.8)
        objTaux.save()
        objTaux = New TauxComm(m_oFourn.id, "CG", 12.5)
        objTaux.save()
        objTaux = New TauxComm(m_oFourn.id, "AE", 13.0)
        objTaux.save()


        'Creation d'une Commande
        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = "06/02/2000"
        'Ajout d'une ligne 
        objLg = objCMD.AjouteLigne("10", objProduit1, 21, 22)
        'Mise à jour des informations de commission
        objLg.CalculCommission(objCMD.Origine, CalculCommQte.CALCUL_COMMISSION_QTE_CMDE)
        Assert.AreEqual(CDec(12.8), objLg.TxComm, "Taux de la commission")
        Assert.AreEqual(Math.Round(CDec((21 * 22 * 12.8) / 100), 2), objLg.MtComm, "Montant de la commission")

        TauxComm.deleteTauxComms(m_oFourn.id)
    End Sub

    <TestMethod()> Public Sub T090_IDPRESTASHOP()
        Dim objCMD As CommandeClient
        Dim objCMD2 As CommandeClient

        objCMD = New CommandeClient(m_oClient)

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
        objCMD.qteColis = 10.5
        objCMD.qtePalettesNonPreparees = 11.5
        objCMD.qtePalettesPreparees = 12.5
        objCMD.poids = 13.5
        objCMD.montantTransport = 1234.56
        objCMD.puPalettesNonPreparees = 14.5
        objCMD.puPalettesPreparees = 15.5
        objCMD.bFactTransport = True
        objCMD.idFactTransport = 1234

        objCMD.IDPrestashop = 4125
        objCMD.NamePrestashop = "AQWZSX"

        Assert.AreEqual(4125L, objCMD.IDPrestashop)
        Assert.AreEqual("AQWZSX", objCMD.NamePrestashop)

        Assert.IsTrue(objCMD.save())

        objCMD2 = CommandeClient.createandload(objCMD.id)
        Assert.AreEqual(4125L, objCMD2.IDPrestashop)
        Assert.AreEqual("AQWZSX", objCMD2.NamePrestashop)

        objCMD2.IDPrestashop = 85296
        objCMD2.NamePrestashop = "EDCRFV"

        Assert.IsTrue(objCMD2.save())

        objCMD = CommandeClient.createandload(objCMD2.id)
        Assert.AreEqual(85296L, objCMD2.IDPrestashop)
        Assert.AreEqual("EDCRFV", objCMD2.NamePrestashop)

    End Sub
    <TestMethod()> Public Sub T095_ORIGINEHOBIVIN()
        Dim objCMD As CommandeClient
        Dim objCMD2 As CommandeClient

        objCMD = New CommandeClient(m_oClient)

        Assert.AreEqual(Dossier.VINICOM, objCMD.Origine)

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
        objCMD.qteColis = 10.5
        objCMD.qtePalettesNonPreparees = 11.5
        objCMD.qtePalettesPreparees = 12.5
        objCMD.poids = 13.5
        objCMD.montantTransport = 1234.56
        objCMD.puPalettesNonPreparees = 14.5
        objCMD.puPalettesPreparees = 15.5
        objCMD.bFactTransport = True
        objCMD.idFactTransport = 1234
        objCMD.IDPrestashop = 4125
        objCMD.NamePrestashop = "AQWZSX"

        objCMD.Origine = Dossier.HOBIVIN


        Assert.IsTrue(objCMD.save())

        objCMD2 = CommandeClient.createandload(objCMD.id)
        Assert.AreEqual(Dossier.HOBIVIN, objCMD2.Origine)

        objCMD2.Origine = Dossier.VINICOM

        Assert.IsTrue(objCMD2.save())

        objCMD = CommandeClient.createandload(objCMD2.id)
        Assert.AreEqual(Dossier.VINICOM, objCMD2.Origine)

    End Sub
    ''' <summary>
    ''' Test de la fonction de recherche sur l'origine
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T100_RECHERCHEORIGINE()
        Dim objCMD As CommandeClient
        Dim nId As Integer
        Dim cmdCode As String
        Dim oCol As Collection

        objCMD = New CommandeClient(m_oClient)
        'objCMD.code = "CMDTEST"
        objCMD.dateCommande = CDate("06/02/1964")
        objCMD.Origine = Dossier.HOBIVIN
        Assert.IsTrue(objCMD.save())

        nId = objCMD.id
        cmdCode = objCMD.code

        'Vérification avec la bonne origine
        oCol = CommandeClient.getListe(cmdCode, , , Dossier.HOBIVIN)
        Assert.AreEqual(1, oCol.Count)
        objCMD = oCol(1)
        Assert.AreEqual(nId, objCMD.id)

        'Vérification avec la Mauvaise origine
        oCol = CommandeClient.getListe(cmdCode, , , Dossier.VINICOM)
        Assert.AreEqual(0, oCol.Count)

        'Vérification sans origine
        oCol = CommandeClient.getListe(cmdCode, , , "")
        Assert.AreEqual(1, oCol.Count)
        objCMD = oCol(1)
        Assert.AreEqual(nId, objCMD.id)

        'Vérification avec les dates et la bonne origine
        oCol = CommandeClient.getListe(CDate("05/02/1964"), CDate("07/02/1964"), , , Dossier.HOBIVIN)
        Assert.AreEqual(1, oCol.Count)
        objCMD = oCol(1)
        Assert.AreEqual(nId, objCMD.id)

        'Vérification avec les dates et sans la bonne origine
        oCol = CommandeClient.getListe(CDate("05/02/1964"), CDate("07/02/1964"), , , Dossier.VINICOM)
        Assert.AreEqual(0, oCol.Count)

        'Vérification avec les dates et Sans l'origine
        oCol = CommandeClient.getListe(CDate("05/02/1964"), CDate("07/02/1964"), , , "")
        Assert.AreEqual(1, oCol.Count)
        objCMD = oCol(1)
        Assert.AreEqual(nId, objCMD.id)

    End Sub

    ''' <summary>
    ''' Test la récupération du Tx de commission dur une ligne de commande
    ''' </summary>
    <TestMethod()> Public Sub T593_GETTXCOMM()
        Dim objCMD As CommandeClient
        Dim objLg As LgCommande

        Dim oParam As Param
        oParam = Param.colTypeClient(1)
        m_oClient.idTypeClient = oParam.id

        Assert.IsTrue(m_oClient.save())
        'Ajout du Tx de commission 
        Dim oTx As New TauxComm(m_oFourn, oParam.code, 15.0)
        Assert.IsTrue(oTx.save)

        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = CDate("06/02/1964")


        objLg = objCMD.AjouteLigne("10", m_oProduit, 10, 15.0)

        Assert.AreEqual(15D, objLg.TxComm)



    End Sub
    ''' <summary>
    ''' Verification du Recalcul du Tax de comm uniquement si RAZ
    ''' </summary>
    <TestMethod()> Public Sub T593_GETTXCOMM2()
        Dim objCMD As CommandeClient
        Dim objLg As LgCommande

        Dim oParam As Param
        oParam = Param.colTypeClient(1)
        m_oClient.idTypeClient = oParam.id

        Assert.IsTrue(m_oClient.save())
        'Ajout du Tx de commission 
        Dim oTx As New TauxComm(m_oFourn, oParam.code, 15.0)
        Assert.IsTrue(oTx.save)

        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = CDate("06/02/1964")
        objLg = objCMD.AjouteLigne("10", m_oProduit, 10, 15.0)
        Assert.AreEqual(15D, objLg.TxComm)

        oTx.TauxComm = 11
        Assert.IsTrue(oTx.save)

        objLg.CalculCommission(objCMD.Origine, CalculCommQte.CALCUL_COMMISSION_QTE_CMDE)
        'Le Tx de comm n'ayant pas été remis à 0 , il n'est pas recalculé
        Assert.AreEqual(15D, objLg.TxComm)

        objLg.TxComm = 0
        objLg.CalculCommission(objCMD.Origine, CalculCommQte.CALCUL_COMMISSION_QTE_CMDE)
        'Le Tx ayant été remis à 0 => Relecture du taux
        Assert.AreEqual(11D, objLg.TxComm)



    End Sub

    <TestMethod()> Public Sub T593_GETTXCOMMHBV()
        Dim objCMD As CommandeClient
        Dim objLg As LgCommande

        Dim oParam As Param
        oParam = Param.colTypeClient(1)
        m_oClient.idTypeClient = oParam.id
        Assert.IsTrue(m_oClient.save())

        m_oProduit.Dossier = Dossier.VINICOM
        Assert.IsTrue(m_oProduit.save())

        'Ajout du Tx de commission 
        Dim oTx As New TauxComm(m_oFourn, oParam.code, 15.0)
        Assert.IsTrue(oTx.save)
        oTx = New TauxComm(m_oFourn, "INT", 10.0)
        Assert.IsTrue(oTx.save)

        Dim oCltIntermediaire As Client
        oCltIntermediaire = Client.getIntermediairePourUneOrigine(Dossier.HOBIVIN)
        If oCltIntermediaire Is Nothing Then
            oCltIntermediaire = New Client("CLTINTER", "ClientIntermédiaire")
            oCltIntermediaire.setTypeIntermediaire(Dossier.HOBIVIN)
            oCltIntermediaire.save()
        End If




        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = CDate("06/02/1964")

        'Command HOBIVIN Produit VINCOM
        objCMD.Origine = Dossier.HOBIVIN
        m_oProduit.Dossier = Dossier.VINICOM
        Assert.IsTrue(m_oProduit.save())
        objLg = objCMD.AjouteLigne("10", m_oProduit, 10, 15.0)
        'Le Tx Est Celui du client intermédiaire
        Assert.AreEqual(10D, objLg.TxComm)

        'Command HOBIVIN Produit HOBIVIN
        m_oProduit.Dossier = Dossier.HOBIVIN
        Assert.IsTrue(m_oProduit.save())
        objLg = objCMD.AjouteLigne("20", m_oProduit, 10, 15.5)
        'Produit HOBIVIN ur Comomande HOBIVIN => 0
        Assert.AreEqual(0D, objLg.TxComm)


    End Sub

    <TestMethod()> Public Sub T593_GETTXCOMMVNC()
        Dim objCMD As CommandeClient
        Dim objLg As LgCommande

        Dim oParam As Param
        oParam = Param.colTypeClient(1)
        m_oClient.idTypeClient = oParam.id
        Assert.IsTrue(m_oClient.save())

        m_oProduit.Dossier = Dossier.VINICOM
        Assert.IsTrue(m_oProduit.save())

        'Ajout du Tx de commission 
        Dim oTx As New TauxComm(m_oFourn, oParam.code, 15.0)
        Assert.IsTrue(oTx.save)
        oTx = New TauxComm(m_oFourn, "INT", 10.0)
        Assert.IsTrue(oTx.save)


        objCMD = New CommandeClient(m_oClient)
        objCMD.dateCommande = CDate("06/02/1964")

        'Command Vinicom Produit VINCOM
        objCMD.Origine = Dossier.VINICOM
        m_oProduit.Dossier = Dossier.VINICOM
        Assert.IsTrue(m_oProduit.save())
        objLg = objCMD.AjouteLigne("10", m_oProduit, 10, 15.0)
        'Commande VINICOM, Produit VINICOM => Tx = Client
        Assert.AreEqual(15D, objLg.TxComm)

        'Command VINICOM Produit HOBIVIN
        m_oProduit.Dossier = Dossier.HOBIVIN
        Assert.IsTrue(m_oProduit.save())
        objLg = objCMD.AjouteLigne("20", m_oProduit, 10, 15.5)
        'Produit HOBIVIN ur Comomande VINICOM=> 0
        Assert.AreEqual(0D, objLg.TxComm)


    End Sub
    ''' <summary>
    ''' Test de Prise en compte de produits gratuits dans le calcul des nombre de colis
    ''' </summary>
    ''' <remarks></remarks>

    <TestMethod()> Public Sub T_PECGratuitNbreColis()
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

        'THEN le nombre de colis de la premier Ligne est égal à 1
        'Les qte de colis ne sont uniquement que sur les Payants
        Assert.AreEqual(CDec(1), objLgCmd1.qteColis, "QteColis de la ligne 1")
        'AND le nombre de colis de la Ligne gratuite est égale à 0
        Assert.AreEqual(CDec(0), objLgCmd2.qteColis, "QteColis de la ligne 2")

        'AND le nombre de colis de la commande est égale à 1
        Assert.AreEqual(1D, objCmd.qteColis, "QteColis de la Commande ")
        'AND le poids de la commande est égale à 12* poids unitaire
        Assert.AreEqual(CDec(12 * oCont.poids), objCmd.poids, "Poids de la Commande ")

        'WHEN j'ajoute une autre Ligne gratuite
        objLgCmd3 = objCmd.AjouteLigne("30", m_oProduit, 1, 0, True)

        'THEN le nombre de colis de la premier Ligne est égal à 2
        Assert.AreEqual(CDec(2), objLgCmd1.qteColis, "QteColis de la ligne 1")
        'AND le nombre de colis des Lignes gratuites est égale à 0
        Assert.AreEqual(CDec(0), objLgCmd2.qteColis, "QteColis de la ligne 2")
        Assert.AreEqual(CDec(0), objLgCmd3.qteColis, "QteColis de la ligne 3")

        'AND Le nombre de colis de la commande est égale à 2
        Assert.AreEqual(2D, objCmd.qteColis, "QteColis de la Commande ")
        'AND le poids est de 13
        Assert.AreEqual(CDec(13 * oCont.poids), objCmd.poids, "Poids de la Commande ")

    End Sub
    ''' <summary>
    ''' Test de Prise en compte de produits gratuits dans le calcul des nombre de colis
    ''' </summary>
    ''' <remarks></remarks>

    <TestMethod()> Public Sub T_PECGratuitWebEDI()
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
        'And le Fichier d'export n'existe pas
        If System.IO.File.Exists("./exportWebEDI.txt") Then
            System.IO.File.Delete("./exportWebEDI.txt")
        End If


        'WHEN j'ajoute une ligne de 11 produit payant
        objLgCmd1 = objCmd.AjouteLigne("10", m_oProduit, 11, 10.5, False)
        'AND un produit Gratuit
        objLgCmd2 = objCmd.AjouteLigne("20", m_oProduit, 1, 0, True)
        'AND j'ajoute une autre Ligne gratuite
        objLgCmd3 = objCmd.AjouteLigne("30", m_oProduit, 1, 0, True)
        'AND J'export en WebEDI
        objCmd.exporterWebEDI("./exportEDI.txt")

        'THEN 
        'Le Fichier d'export est bien créé
        Assert.IsTrue(System.IO.File.Exists("./ExportEDI.txt"))
        Dim oSr As StreamReader = System.IO.File.OpenText("./ExportEDI.txt")
        Dim nLineNumber As Integer = 0
        Dim strResult As String
        While Not oSr.EndOfStream

            nLineNumber = nLineNumber + 1
            strResult = oSr.ReadLine()
        End While

        Assert.AreEqual(1, nLineNumber, "Une seule Ligne de fichier")
        Assert.AreEqual("0000002", Mid(strResult, 310, 7), "qte Colis en WebEDI")
        Assert.AreEqual("0000013", Mid(strResult, 317, 7), "qte Commandée en WebEDI")

    End Sub
    <TestMethod()> Public Sub T_ToutLivrerOK()
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
        'And le Fichier d'export n'existe pas
        If System.IO.File.Exists("./exportWebEDI.txt") Then
            System.IO.File.Delete("./exportWebEDI.txt")
        End If


        'WHEN j'ajoute une ligne de 11 produit payant
        objLgCmd1 = objCmd.AjouteLigne("10", m_oProduit, 11, 10.5, False)
        'AND un produit Gratuit
        objLgCmd2 = objCmd.AjouteLigne("20", m_oProduit, 1, 0, True)
        'AND j'ajoute une autre Ligne gratuite
        objLgCmd3 = objCmd.AjouteLigne("30", m_oProduit, 1, 0, True)
        'AND je valide ma commande
        objCmd.save()
        'And je Livre ma Commande
        objCmd.LivrerToutOK()


        'THEN 
        'Le nombre de Colis est de 2
        Assert.AreEqual(vncEnums.vncEtatCommande.vncLivree, objCmd.etat.codeEtat)
        'Pour Chaque Ligne la qte Livrée est égale à la Qte Commandées
        For Each objLg As LgCommande In objCmd.colLignes
            Assert.AreEqual(objLg.qteCommande, objLg.qteLiv)
        Next

    End Sub

End Class



