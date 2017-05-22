'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V200
    Inherits test_Base


    Private m_objPRD As Produit
    Private m_objFRN As Fournisseur
    Private m_objCLT As Client
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
        Persist.shared_connect()
        m_objFRN = New Fournisseur("FRNV200", "FRn de test")
        Assert.IsTrue(m_objFRN.Save(), "FRN.Create")

        m_objPRD = New Produit("PRD200", m_objFRN, 1990)
        Assert.IsTrue(m_objPRD.save(), "Prod.Create")

        m_objCLT = New Client("CLT200", "Client de test")
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
    <TestMethod()> Public Sub T10_CalculPoidsetColis()
        Dim objCmd As CommandeClient
        Dim objLgCmd As LgCommande
        Dim nidCmd As Integer

        Dim poidsCont As Decimal
        Dim nbreCond As Decimal

        poidsCont = contenant.contenantDefaut.poids
        nbreCond = Param.conditionnementdefaut.valeur


        '            For Each objCont In contenant.colContenant
        '            If objCont.id = m_objPRD.idContenant Then
        '            objCont.poids = 1.5
        '            End If
        '            Next

        'For Each objCond In Param.colConditionnement
        'If objCond.id = m_objPRD.idConditionnement Then
        'objCond.valeur = 6
        'End If
        'Next



        objCmd = New CommandeClient(m_objCLT)
        'Prix U des pall est en constante
        Assert.IsTrue(objCmd.puPalettesPreparees = Param.getConstante("CST_PU_PALL_PREP"), "Pu des pallettes préparées")
        Assert.IsTrue(objCmd.puPalettesNonPreparees = Param.getConstante("CST_PU_PALL_NONPREP"), "Pu des pallettes nonpréparées")

        Assert.IsTrue(objCmd.save(), "Création de la commande")

        objLgCmd = objCmd.AjouteLigne("10", m_objPRD, 18, 10.5)

        objLgCmd.calcPoidsColis()
        Assert.AreEqual(objLgCmd.poids, 18 * poidsCont, "Poids de la ligne")
        Assert.AreEqual(objLgCmd.qteColis, 18 / nbreCond, "QteColis de la ligne")

        objLgCmd = objCmd.AjouteLigne("20", m_objPRD, 24, 10.5)

        Assert.IsTrue(objCmd.CalcPoidsColis(), objCmd.getErreur())
        Assert.AreEqual(objCmd.poids, (18 + 24) * poidsCont, "Poids de la Commande ")
        Assert.AreEqual(CDec(objCmd.qteColis), CDec((18 + 24) / nbreCond), "QteColis de la Commande ")

        Assert.IsTrue(objCmd.save(), "Sauvegarde de la commande")
        nidCmd = objCmd.id

        objCmd = CommandeClient.createandload(nidCmd)
        objCmd.loadcolLignes()

        objLgCmd = objCmd.colLignes(1)
        Assert.AreEqual(objLgCmd.poids, 18 * poidsCont, "Poids de la ligne Apres rechargement")
        Assert.AreEqual(objLgCmd.qteColis, 18 / nbreCond, "QteColis de la ligne apres rechargement")
        Assert.AreEqual(objCmd.poids, (18 + 24) * poidsCont, "Poids de la Commande Apres rechargement")
        Assert.AreEqual(CDec(objCmd.qteColis), CDec((18 + 24) / nbreCond), "QteColis de la Commande Apres rechargement")



        objCmd.bDeleted = True
        Assert.IsTrue(objCmd.save(), "Suppression de la commande")


    End Sub
    <TestMethod()> Public Sub T20_Transporteur()
        Dim oCol As Collection
        Dim oTrp As Transporteur
        Dim oTrp1 As Transporteur
        Dim oCmd As CommandeClient
        Dim oCmd2 As CommandeClient
        Dim nidCommande As Long

        oCol = Transporteur.colTransporteur

        'Recherche du premier transporteur qui n'est defaut
        For Each oTrp In oCol
            'Console.Out.WriteLine(oTrp.shortResume)
            If Not oTrp.bDefaut Then
                oTrp1 = oTrp
            End If
        Next



        oCmd = New CommandeClient(m_objCLT)
        oCmd.oTransporteur = oTrp1

        Assert.IsTrue(oCmd.save(), "Sauvegarde de la commande")

        nidCommande = oCmd.id
        oCmd2 = CommandeClient.createandload(nidCommande)

        Assert.IsTrue(oCmd.oTransporteur.Equals(oCmd.oTransporteur), "Les transporteurs sont égaux ")

        For Each oTrp In oCol
            'Console.Out.WriteLine(oTrp.shortResume)
        Next


        oCmd2.bDeleted = True
        Assert.IsTrue(oCmd2.save, "Suppression de la commande")
    End Sub
    <TestMethod()> Public Sub T30_BONappro()
        Dim objBA1 As BonAppro
        Dim objBA2 As BonAppro
        Dim nidCmd As Integer
        Dim oCol As Collection
        Dim objTRp As Transporteur


        oCol = Transporteur.colTransporteur


        objBA1 = New BonAppro(m_objFRN)
        'Prix U des pall est en constante
        Assert.IsTrue(objBA1.puPalettesPreparees = Param.getConstante("CST_PU_PALL_PREP"), "Pu des pallettes préparées")
        Assert.IsTrue(objBA1.puPalettesNonPreparees = Param.getConstante("CST_PU_PALL_NONPREP"), "Pu des pallettes nonpréparées")


        objBA1.qteColis = 123
        objBA1.qtePalettesNonPreparees = 456
        objBA1.qtePalettesNonPreparees = 789
        objBA1.poids = 1011

        objTRp = oCol(1)
        objBA1.oTransporteur = objTRp
        objBA1.oTransporteur.nom = "Mon Trap"


        Assert.IsTrue(objBA1.save(), "Création du bon Appro")
        nidCmd = objBA1.id



        objBA2 = BonAppro.createandload(nidCmd)

        Assert.AreEqual(objBA1.puPalettesNonPreparees, objBA2.puPalettesNonPreparees, "puPalettesNonPreparees")
        Assert.AreEqual(objBA1.puPalettesPreparees, objBA2.puPalettesPreparees, "puPalettesPreparees")
        Assert.AreEqual(objBA1.qteColis, objBA2.qteColis, "qteColis")
        Assert.AreEqual(objBA1.qtePalettesNonPreparees, objBA2.qtePalettesNonPreparees, "qtePalettesNonPreparees")
        Assert.AreEqual(objBA1.qtePalettesPreparees, objBA2.qtePalettesPreparees, "qtePalettesPreparees")
        Assert.AreEqual(objBA1.poids, objBA2.poids, "poids")
        Assert.AreEqual(objBA1.poids, objBA2.poids, "poids")

        Assert.AreEqual(objBA1.oTransporteur.id, objBA2.oTransporteur.id, "TRP_ID")
        Assert.AreEqual(objBA1.oTransporteur.code, objBA2.oTransporteur.code, "TRP_CODE")
        Assert.AreEqual(objBA1.oTransporteur.nom, objBA2.oTransporteur.nom, "TRP_NOM")
        Assert.AreEqual(objBA1.oTransporteur.AdresseLivraison.rue1, objBA2.oTransporteur.AdresseLivraison.rue1, "TRP_LIV_RUE1")
        Assert.AreEqual(objBA1.oTransporteur.AdresseLivraison.rue2, objBA2.oTransporteur.AdresseLivraison.rue2, "TRP_LIV_RUE2")
        Assert.AreEqual(objBA1.oTransporteur.AdresseLivraison.cp, objBA2.oTransporteur.AdresseLivraison.cp, "TRP_LIV_CP")
        Assert.AreEqual(objBA1.oTransporteur.AdresseLivraison.ville, objBA2.oTransporteur.AdresseLivraison.ville, "TRP_LIV_VILLE")
        Assert.AreEqual(objBA1.oTransporteur.AdresseLivraison.tel, objBA2.oTransporteur.AdresseLivraison.tel, "TRP_LIV_TEL")
        Assert.AreEqual(objBA1.oTransporteur.AdresseLivraison.fax, objBA2.oTransporteur.AdresseLivraison.fax, "TRP_LIV_FAX")
        Assert.AreEqual(objBA1.oTransporteur.AdresseLivraison.port, objBA2.oTransporteur.AdresseLivraison.port, "TRP_LIV_PORT")
        Assert.AreEqual(objBA1.oTransporteur.AdresseLivraison.Email, objBA2.oTransporteur.AdresseLivraison.Email, "TRP_LIV_EMAIL")


        objBA2.bDeleted = True
        Assert.IsTrue(objBA2.save(), "Suppression du bon Appro")


    End Sub
    <TestMethod()> Public Sub T30_CommandeTransporteur()
        Dim objCmd As CommandeClient
        Dim oSCMD1 As SousCommande
        Dim oSCMD2 As SousCommande

        objCmd = New CommandeClient(m_objCLT)
        oSCMD1 = New SousCommande(objCmd, m_objFRN)
        oSCMD2 = New SousCommande(objCmd, m_objFRN)

        oSCMD1.oTransporteur.nom = "TOTO"
        Assert.IsFalse(oSCMD2.bUpdated)

    End Sub

    <TestMethod()> Public Sub T40_Ligne_Gratuite()
        Dim objCmd As CommandeClient
        Dim objLgCmd As LgCommande
        Dim objLgCmd2 As LgCommande
        Dim nidCmd As Integer


        objCmd = New CommandeClient(m_objCLT)
        objLgCmd = objCmd.AjouteLigne("10", m_objPRD, 18, 10.5, True)
        objLgCmd2 = objCmd.AjouteLigne("20", m_objPRD, 18, 10.5, False)

        Assert.IsTrue(objCmd.save(), "Suppression de la commande")
        nidCmd = objCmd.id

        objCmd = CommandeClient.createandload(nidCmd)
        objCmd.loadcolLignes()
        objLgCmd = objCmd.colLignes(1)
        objLgCmd2 = objCmd.colLignes(2)

        Assert.IsTrue(objLgCmd.bGratuit, "Ligne1")
        Assert.IsFalse(objLgCmd2.bGratuit, "Ligne2")

        objCmd.bDeleted = True
        Assert.IsTrue(objCmd.save(), "Suppression de la commande")


    End Sub

End Class


