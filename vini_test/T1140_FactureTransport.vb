Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB
Imports System.IO


<TestClass()> Public Class T1140_FactureTransport
    Inherits test_Base

    Private m_nTXGO As Decimal
    Private m_oProduit As Produit
    Private m_oFourn As Fournisseur
    Private m_oClient As Client
    Private m_oClient2 As Client
    Private m_oClient3 As Client
    Dim objProduit1 As Produit
    Dim objProduit2 As Produit
    Dim objProduit3 As Produit
    <TestInitialize()> Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()

        Persist.shared_connect()
        m_oFourn = New Fournisseur("FRN" & Now(), "MonFournisseur")
        m_oFourn.Save()

        m_oProduit = New Produit("TSTPRD" & Now(), m_oFourn, 1990)
        m_oProduit.save()

        m_oClient = New Client("CLT1-" & Now(), "MonClient1")
        m_oClient.rs = "MonClient1RS"
        m_oClient2 = New Client("CLT2-" & Now(), "MonClient2")
        m_oClient2.rs = "MonClient2RS"
        m_oClient3 = New Client("CLT3-" & Now(), "MonClient3")
        m_oClient3.rs = "MonClient3RS"
        m_oClient.CodeCompta = "4100001"

        Debug.Assert(m_oClient.save(), "Creation du client")
        Debug.Assert(m_oClient2.save(), "Creation du client2")
        Debug.Assert(m_oClient3.save(), "Creation du client3")
        '            m_oClient()
        objProduit1 = New Produit("PRD1T20", m_oFourn, 1994)
        Assert.IsTrue(objProduit1.save())
        objProduit2 = New Produit("PRD2T20", m_oFourn, 1994)
        Assert.IsTrue(objProduit2.save())
        objProduit3 = New Produit("PRD3T20", m_oFourn, 1994)
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
        m_oFourn.bDeleted = True
        m_oFourn.Save()
        m_oClient.bDeleted = True
        m_oClient.save()
        m_oClient2.bDeleted = True
        m_oClient2.save()
        m_oClient3.bDeleted = True
        m_oClient3.save()

        MyBase.TestCleanup()

    End Sub
    <TestMethod()> Public Sub T10_Object()
        Dim objFactTRP As FactTRP
        Dim codeEtat As Integer

        objFactTRP = New FactTRP(m_oClient)
        Assert.IsTrue(objFactTRP.oTiers.Equals(m_oClient), "Le Client est identique à celui passé en paramètre à la création")
        Assert.IsTrue(objFactTRP.periode = "")
        Assert.IsTrue(objFactTRP.montantReglement = 0)
        Assert.IsTrue(objFactTRP.etat.codeEtat = vncEnums.vncEtatCommande.vncFactTRPGeneree)
        Assert.IsTrue(objFactTRP.code = "")
        Assert.IsTrue(objFactTRP.totalHT = 0)
        Assert.IsTrue(objFactTRP.totalTTC = 0)
        Assert.IsTrue(objFactTRP.id = 0)
        Assert.AreEqual(m_oClient.idModeReglement1, objFactTRP.idModeReglement)


        objFactTRP.dateCommande = CDate("06/02/05")
        Assert.IsTrue(objFactTRP.dateCommande.Equals(CDate("06/02/05")))

        objFactTRP.totalHT = 100
        Assert.IsTrue(objFactTRP.totalHT = 100)
        objFactTRP.totalTTC = 110
        Assert.IsTrue(objFactTRP.totalTTC = 110)

        objFactTRP.periode = "Septembre"
        Assert.IsTrue(objFactTRP.periode.Equals("Septembre"))

        objFactTRP.idModeReglement = 45
        Assert.IsTrue(objFactTRP.idModeReglement = 45, "Test idModeReglement")

        'Test de l'état de la Facture
        codeEtat = objFactTRP.etat.codeEtat

        objFactTRP.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactTRPGenerer)
        Assert.IsTrue(objFactTRP.etat.codeEtat = codeEtat, "Code Etat ne change pas")


        objFactTRP.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactTRPExporter)
        Assert.AreEqual(vncEnums.vncEtatCommande.vncFactTRPExportee, objFactTRP.etat.codeEtat, "Code Etat ne change pas")
        objFactTRP.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactTRPAnnExporter)
        Assert.AreEqual(vncEnums.vncEtatCommande.vncFactTRPGeneree, objFactTRP.etat.codeEtat, "Code Etat ne change pas")





        objFactTRP.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactTRPExporter)
        Assert.IsTrue(objFactTRP.etat.codeEtat = vncEnums.vncEtatCommande.vncFactTRPExportee, "Code Etat change ")
        objFactTRP.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactTRPGenerer)
        Assert.IsTrue(objFactTRP.etat.codeEtat = vncEnums.vncEtatCommande.vncFactTRPExportee, "Code Etat ne change pas")
        objFactTRP.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactTRPAnnGenerer)
        Assert.IsTrue(objFactTRP.etat.codeEtat = vncEnums.vncEtatCommande.vncFactTRPExportee, "Code Etat ne change pas")

    End Sub 'T10_Object
    <TestMethod()> Public Sub T20_DB1()
        Dim objFactTRP As FactTRP
        Dim objFactTRP2 As FactTRP
        Dim nId As Integer

        objFactTRP = New FactTRP(m_oClient)
        objFactTRP.dEcheance = "01/04/1964"
        objFactTRP.idModeReglement = 1

        Assert.IsTrue(objFactTRP.Save(), objFactTRP.getErreur())

        Assert.IsTrue(objFactTRP.id <> 0)
        Assert.IsTrue(objFactTRP.code <> "")
        nId = objFactTRP.id

        objFactTRP2 = FactTRP.createandload(nId)

        Assert.IsTrue(objFactTRP.Equals(objFactTRP2))
        Assert.AreEqual(CDate("01/04/1964"), objFactTRP2.dEcheance, "Date Echeance")
        Assert.AreEqual(1, objFactTRP2.idModeReglement, "Mode de reglement")

        objFactTRP2.CommFacturation.comment = "Mon Commentaire"
        objFactTRP2.dEcheance = "01/04/1984"
        objFactTRP2.idModeReglement = 2
        Assert.IsTrue(objFactTRP2.Save())
        objFactTRP2 = FactTRP.createandload(nId)
        Assert.IsTrue(objFactTRP2.CommFacturation.comment = "Mon Commentaire")

        Assert.AreEqual(CDate("01/04/1984"), objFactTRP2.dEcheance, "Date Echeance")
        Assert.AreEqual(2, objFactTRP2.idModeReglement, "Mode de reglement")


        objFactTRP2.bDeleted = True
        Assert.IsTrue(objFactTRP2.Save())


    End Sub
    <TestMethod()> Public Sub T21_DB_LG()
        Dim objFactTRP As FactTRP
        Dim objFactTRP2 As FactTRP
        Dim objCMDCLT As CommandeClient
        Dim objLgFactTRP As LgFactTRP
        Dim objLgFactTRP2 As LgFactTRP
        Dim nidCMDCLT As Long
        Dim nidFactTRP As Long

        setTaxeGOEqualZero()

        objCMDCLT = New CommandeClient(m_oClient)

        objCMDCLT.montantTransport = 50
        objCMDCLT.oTransporteur.nom = "TRP TEST"
        objCMDCLT.dateCommande = CDate("06/02/05")
        objCMDCLT.dateLivraison = CDate("07/02/05")
        objCMDCLT.refLivraison = "BLRTH87UJ"
        objCMDCLT.bFactTransport = True
        objCMDCLT.qteColis = 234.1
        objCMDCLT.qtePalettesNonPreparees = 345.2
        objCMDCLT.qtePalettesPreparees = 456.3
        objCMDCLT.puPalettesNonPreparees = 12.1
        objCMDCLT.puPalettesPreparees = 13.2
        objCMDCLT.poids = 678.4
        objCMDCLT.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)

        Assert.IsTrue(objCMDCLT.save, "Sauvegarde de la commande")
        nidCMDCLT = objCMDCLT.id

        objFactTRP = New FactTRP(m_oClient)
        Assert.IsTrue(objFactTRP.Save(), "Sauvegarde de la facture de transport")
        nidFactTRP = objFactTRP.id

        'Premiere Ligne
        objLgFactTRP = objFactTRP.AjouteLigneFactTRP(objCMDCLT, True)

        Assert.IsTrue(objFactTRP.bcolLignesLoaded = True)
        Assert.IsTrue(Not objLgFactTRP Is Nothing, "Ligne de facture crée")
        'Test de la dupplication des éléments
        Assert.IsTrue(objLgFactTRP.idCmdCLT = objCMDCLT.id)
        Assert.IsTrue(objLgFactTRP.idFactTRP = objFactTRP.id)
        Assert.IsTrue(objLgFactTRP.nomTransporteur = objCMDCLT.oTransporteur.nom)
        Assert.IsTrue(objLgFactTRP.dateCommande = objCMDCLT.dateCommande)
        Assert.IsTrue(objLgFactTRP.dateLivraison = objCMDCLT.dateLivraison)
        Assert.IsTrue(objLgFactTRP.refCommande = objCMDCLT.code)
        Assert.IsTrue(objLgFactTRP.referenceLivraison = objCMDCLT.refLivraison)
        Assert.IsTrue(objLgFactTRP.qteColis = objCMDCLT.qteColis)
        Assert.IsTrue(objLgFactTRP.qtePalettesPreparees = objCMDCLT.qtePalettesPreparees)
        Assert.IsTrue(objLgFactTRP.qtePalettesNonPreparees = objCMDCLT.qtePalettesNonPreparees)
        Assert.IsTrue(objLgFactTRP.puPalettesPreparees = objCMDCLT.puPalettesPreparees)
        Assert.IsTrue(objLgFactTRP.puPalettesNonPreparees = objCMDCLT.puPalettesNonPreparees)
        Assert.IsTrue(objLgFactTRP.poids = objCMDCLT.poids)

        Assert.IsTrue(objCMDCLT.idFactTransport = objFactTRP.id)

        Assert.IsTrue(objFactTRP.colLignes.Count = 1)
        Assert.IsTrue(objLgFactTRP.Equals(objFactTRP.colLignes(1)))
        Assert.IsTrue(objFactTRP.bcolLignesLoaded = True)

        Assert.IsTrue(objFactTRP.Save(), "Sauvegarde de la facture avec ses lignes")

        'Rechargement de la facture
        objFactTRP2 = FactTRP.createandload(objFactTRP.id)
        objFactTRP2.loadcolLignes()

        'Test de la Collection des lignes lues
        Assert.IsTrue(objFactTRP2.colLignes.Count = 1, "Une ligne de transport chargée")
        objLgFactTRP2 = objFactTRP2.colLignes(1)
        Assert.IsTrue(objLgFactTRP.Equals(objLgFactTRP2), "Les lignes sont identiques")

        'Modification d'une ligne
        objLgFactTRP2.nomTransporteur = "TRP2"
        objLgFactTRP2.puPalettesNonPreparees = 123.56
        objLgFactTRP2.puPalettesPreparees = 234.56
        Assert.IsTrue(objFactTRP2.Save(), "Update des lignes de factures")

        'Rechargement de la facture
        objFactTRP2 = FactTRP.createandload(objFactTRP.id)
        objFactTRP2.loadcolLignes()

        'Test de la Collection des lignes lues
        Assert.IsTrue(objFactTRP2.colLignes.Count = 1, "Une ligne de transport chargée")
        objLgFactTRP = objFactTRP2.colLignes(1)
        Assert.IsTrue(objLgFactTRP.Equals(objLgFactTRP2), "Les lignes dont identiques")

        'Ajout d'une nouvelle ligne de transport (Sans commande)
        objLgFactTRP = New LgFactTRP
        objLgFactTRP.idFactTRP = objFactTRP2.id
        objLgFactTRP.nomTransporteur = "Moi"
        objLgFactTRP.prixHT = 100
        objLgFactTRP.puPalettesNonPreparees = 123.34
        objLgFactTRP.puPalettesPreparees = 234.56

        objFactTRP2.AjouteLigneFactTRP(objLgFactTRP)
        Assert.IsTrue(objFactTRP2.Save(), "Sauvegarde de la facture de transport avec 2 lignes")
        'Relecture de la facture
        objFactTRP2 = FactTRP.createandload(objFactTRP.id)
        objFactTRP2.loadcolLignes()

        'Test de la Collection des lignes lues
        Assert.AreEqual(2, objFactTRP2.colLignes.Count, "2 lignes de transport attendues")
        objLgFactTRP2 = objFactTRP2.colLignes(2)
        Assert.IsTrue(objLgFactTRP.Equals(objLgFactTRP2), "Les lignes dont identiques")

        objCMDCLT = CommandeClient.createandload(nidCMDCLT)
        Assert.IsTrue(objCMDCLT.idFactTransport = nidFactTRP, "La commande a été mise à jour avec l'id de la facture de transport")


        objFactTRP.bDeleted = True
        objFactTRP.Save()

        objCMDCLT.bDeleted = True
        objCMDCLT.save()

        resetTaxeGO()

    End Sub
    <TestMethod()> Public Sub T30_LISTE_CMDCLT_TRP()
        Dim objCMDCLT As CommandeClient
        Dim nidCMDCLT As Long
        Dim objCMDCLT2 As CommandeClient
        Dim nidCMDCLT2 As Long
        Dim nidCMDCLT3 As Long
        Dim oCol As Collection
        Dim oFactTRP As FactTRP
        Dim strNum As String
        Dim oLgFactTRP As LgFactTRP

        setTaxeGOEqualZero()
        'Creation de 2 commandes Transport
        objCMDCLT = New CommandeClient(m_oClient)
        objCMDCLT.bFactTransport = True
        objCMDCLT.montantTransport = 100.3
        objCMDCLT.dateCommande = "06/02/1964"
        objCMDCLT.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        objCMDCLT.dateLivraison = "06/02/1964"
        Assert.IsTrue(objCMDCLT.save(), "Sauavegarde de la Commande Client avec trp")
        nidCMDCLT = objCMDCLT.id

        objCMDCLT = New CommandeClient(m_oClient)
        objCMDCLT.bFactTransport = True
        objCMDCLT.montantTransport = 110.4
        objCMDCLT.dateCommande = "06/02/1964"
        objCMDCLT.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        objCMDCLT.dateLivraison = "06/02/1964"
        Assert.IsTrue(objCMDCLT.save(), "Sauavegarde de la Commande Client avec trp")
        nidCMDCLT3 = objCMDCLT.id

        'Creation de 1 commandes sans Transport
        objCMDCLT2 = New CommandeClient(m_oClient)
        objCMDCLT2.bFactTransport = False
        objCMDCLT2.montantTransport = 0.0
        objCMDCLT2.dateCommande = "06/02/1964"
        objCMDCLT2.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        objCMDCLT2.dateLivraison = "06/02/1964"
        Assert.IsTrue(objCMDCLT2.save(), "Sauavegarde de la Commande Client avec trp")
        nidCMDCLT2 = objCMDCLT2.id

        oCol = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.IsTrue(oCol.Count = 2, "2 Elements dans la Liste")

        'Creation d'une facture de trabsport
        oFactTRP = New FactTRP(m_oClient)
        Assert.IsTrue(oFactTRP.Save(), "Sauvegarde de la facture de trabsport")
        'Ajout d'une ligne de facture de transport
        oLgFactTRP = oFactTRP.AjouteLigneFactTRP(objCMDCLT)
        Assert.IsTrue(oFactTRP.Save(), "Sauvegarde de la facture de trabsport")
        strNum = oLgFactTRP.num

        oCol = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.IsTrue(oCol.Count = 1, "1 Seul Element dans la Liste")
        objCMDCLT = oCol(1)
        Assert.IsTrue(objCMDCLT.id = nidCMDCLT)

        'Suppression de la ligne de la facture de transport
        oFactTRP.supprimeLigne(strNum)
        oFactTRP.Save()
        oCol = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.IsTrue(oCol.Count = 2, "2 Elements dans la Liste")

        'Ajout d'une ligne de facture de transport
        objCMDCLT = CommandeClient.createandload(nidCMDCLT3)
        oLgFactTRP = oFactTRP.AjouteLigneFactTRP(objCMDCLT)
        Assert.IsTrue(oFactTRP.Save(), "Sauvegarde de la facture de trabsport")
        oCol = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.IsTrue(oCol.Count = 1, "1 Seul Element dans la Liste")
        objCMDCLT = oCol(1)
        Assert.IsTrue(objCMDCLT.id = nidCMDCLT)


        'Suppression de la facture de Transport
        oFactTRP.bDeleted = True
        oFactTRP.Save()

        'La commande a été libérée
        oCol = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.IsTrue(oCol.Count = 2, "2 Elements dans la Liste")


        objCMDCLT = CommandeClient.createandload(nidCMDCLT)
        objCMDCLT.bDeleted = True
        objCMDCLT.save()

        objCMDCLT = CommandeClient.createandload(nidCMDCLT2)
        objCMDCLT.bDeleted = True
        objCMDCLT.save()
        resetTaxeGO()

    End Sub
    <TestMethod(), Ignore()> Public Sub T40_GENERATION_FACTTRP()
        Dim oCLT1 As Client
        Dim oCLT2 As Client
        Dim nIdCLT1 As Long
        Dim nIdCLT2 As Long
        Dim oCLT3 As Client
        Dim nIdCLT3 As Long
        Dim objCMDCLT1 As CommandeClient
        Dim objCMDCLT2 As CommandeClient
        Dim objCMDCLT3 As CommandeClient
        Dim objCMDCLT4 As CommandeClient
        Dim objCMDCLT5 As CommandeClient
        Dim objCMDCLT6 As CommandeClient
        Dim objCMDCLT7 As CommandeClient
        Dim nIdCMDCLT1 As Long
        Dim nIdCMDCLT2 As Long
        Dim nIdCMDCLT3 As Long
        Dim nIdCMDCLT4 As Long
        Dim nIdCMDCLT5 As Long
        Dim nIdCMDCLT6 As Long
        Dim nIdCMDCLT7 As Long
        Dim oColCMD As Collection
        Dim oColFactTRP As ColEvent
        Dim oFactTRP1 As FactTRP
        Dim oLgFactTRP As LgFactTRP
        Dim oCMD As CommandeClient
        Dim oParam As ParamModeReglement





        setTaxeGOEqualZero()
        'Creation des Clients
        oCLT1 = New Client("CLT1T40", "Mon Client1T401")
        Assert.IsTrue(oCLT1.save(), "Sauvegarde du 1er Client")
        nIdCLT1 = oCLT1.id

        oCLT2 = New Client("CLT2T40", "Mon Client2T401")
        Assert.IsTrue(oCLT2.save(), "Sauvegarde du 1er Client")
        nIdCLT2 = oCLT2.id

        oCLT3 = New Client("CLT3T40", "Mon Client3T401")
        Assert.IsTrue(oCLT3.save(), "Sauvegarde du 3eme Client")
        nIdCLT3 = oCLT3.id

        'Création d'un mode de reglement 30 fin de mois
        oParam = New ParamModeReglement()
        oParam.code = "TEST30FDM"
        oParam.dDebutEcheance = "FDM"
        oParam.valeur2 = 30
        Assert.IsTrue(oParam.Save())
        oCLT1.idModeReglement1 = oParam.id
        Assert.IsTrue(oCLT1.save())

        'Création d'un mode de reglement comptant
        oParam = New ParamModeReglement()
        oParam.dDebutEcheance = "FACT"
        oParam.code = "TESTCOMPTANT"
        oParam.valeur2 = 0
        Assert.IsTrue(oParam.Save())
        oCLT2.idModeReglement1 = oParam.id
        Assert.IsTrue(oCLT2.save())

        'Creation des Commandes

        '2+1 Commandes pour le Client1
        objCMDCLT1 = New CommandeClient(oCLT1)
        objCMDCLT1.dateLivraison = "06/02/1964"
        objCMDCLT1.bFactTransport = True
        objCMDCLT1.montantTransport = 101.5
        objCMDCLT1.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT1.save(), "Sauvegarde de la premieère commande")
        nIdCMDCLT1 = objCMDCLT1.id

        objCMDCLT2 = New CommandeClient(oCLT1)
        objCMDCLT2.dateLivraison = "06/02/1964"
        objCMDCLT2.bFactTransport = True
        objCMDCLT2.montantTransport = 201.5
        objCMDCLT2.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT2.save(), "Sauvegarde de la 2eme commande")
        nIdCMDCLT2 = objCMDCLT2.id

        objCMDCLT3 = New CommandeClient(oCLT1)
        objCMDCLT3.dateLivraison = "06/02/1964"
        objCMDCLT3.bFactTransport = False
        objCMDCLT3.montantTransport = 301.5
        objCMDCLT3.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT3.save(), "Sauvegarde de la 3eme commande")
        nIdCMDCLT3 = objCMDCLT3.id

        '2 Commandes pour le Client2
        objCMDCLT4 = New CommandeClient(oCLT2)
        objCMDCLT4.dateLivraison = "06/02/1964"
        objCMDCLT4.bFactTransport = True
        objCMDCLT4.montantTransport = 401.5
        objCMDCLT4.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT4.save(), "Sauvegarde de la 4eme commande")
        nIdCMDCLT4 = objCMDCLT4.id

        objCMDCLT5 = New CommandeClient(oCLT2)
        objCMDCLT5.dateLivraison = "06/02/1964"
        objCMDCLT5.bFactTransport = True
        objCMDCLT5.montantTransport = 501.5
        objCMDCLT5.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT5.save(), "Sauvegarde de la 5eme commande")
        nIdCMDCLT5 = objCMDCLT5.id

        '1 Commande pour le Client3
        objCMDCLT6 = New CommandeClient(oCLT3)
        objCMDCLT6.dateLivraison = "06/02/1964"
        objCMDCLT6.bFactTransport = True
        objCMDCLT6.montantTransport = 601.5
        objCMDCLT6.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT6.save(), "Sauvegarde de la 5eme commande")
        nIdCMDCLT6 = objCMDCLT6.id

        '1 Commande Non Livrée pour le Client3
        objCMDCLT7 = New CommandeClient(oCLT3)
        objCMDCLT7.dateLivraison = "06/02/1964"
        objCMDCLT7.bFactTransport = True
        objCMDCLT7.montantTransport = 601.5
        Assert.IsTrue(objCMDCLT7.save(), "Sauvegarde de la 5eme commande")
        nIdCMDCLT7 = objCMDCLT7.id

        'Recuperation des commande avec transport 
        oColCMD = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.IsTrue(oColCMD.Count = 5, "5 elements sont Attendus dans la liste->" & oColCMD.Count)
        For Each oCMD In oColCMD
            Console.WriteLine(oCMD.shortResume)
        Next

        'Génération des gfactures de transport
        oColFactTRP = FactTRP.createFactTRPs(oColCMD, "29/02/1964", , "FEVRIER 64")
        Assert.IsTrue(oColFactTRP.Count = 3, "3 Factures de Générées")


        'Controle des factures de transport générées
        oFactTRP1 = oColFactTRP(oCLT1.code)
        Assert.IsTrue(Not oFactTRP1 Is Nothing, "Facture du client1")
        Assert.IsTrue(oFactTRP1.colLignes.Count = 2, "2 lignes de factures dans la Commande 1")
        Assert.IsTrue(oFactTRP1.montantTaxes = Param.getConstante("CST_TAXES_TRP") * 2, "Montant des taxes ")

        Assert.AreEqual(objCMDCLT1.montantTransport + objCMDCLT2.montantTransport + oFactTRP1.montantTaxes, oFactTRP1.totalHT, "Montant HT de la commande")
        Assert.AreEqual(oFactTRP1.idModeReglement, oCLT1.idModeReglement1)
        Assert.AreNotEqual(CDate("01/01/2000"), oFactTRP1.dEcheance)

        oLgFactTRP = oFactTRP1.colLignes(1)
        Assert.IsTrue(Not oLgFactTRP Is Nothing)
        Assert.IsTrue(oLgFactTRP.idCmdCLT = objCMDCLT1.id Or oLgFactTRP.idCmdCLT = objCMDCLT2.id, "Id de commande")

        oLgFactTRP = oFactTRP1.colLignes(2)
        Assert.IsTrue(Not oLgFactTRP Is Nothing)
        Assert.IsTrue(oLgFactTRP.idCmdCLT = objCMDCLT1.id Or oLgFactTRP.idCmdCLT = objCMDCLT2.id, "Id de commande")


        'Controle des factures de transport générées
        oFactTRP1 = oColFactTRP(oCLT2.code)
        Assert.IsTrue(Not oFactTRP1 Is Nothing, "Facture du client2")
        Assert.IsTrue(oFactTRP1.colLignes.Count = 2, "2 lignes de factures dans la Facture 2")
        Assert.IsTrue(oFactTRP1.totalHT = objCMDCLT4.montantTransport + objCMDCLT5.montantTransport + oFactTRP1.montantTaxes, "Montant HT de la commande")
        Assert.IsTrue(oFactTRP1.montantTaxes = Param.getConstante("CST_TAXES_TRP") * 2, "Montant des taxes ")
        Assert.AreEqual(oFactTRP1.idModeReglement, oCLT2.idModeReglement1)
        Assert.AreNotEqual(CDate("01/01/2000"), oFactTRP1.dEcheance)

        oLgFactTRP = oFactTRP1.colLignes(1)
        Assert.IsTrue(Not oLgFactTRP Is Nothing)
        Assert.IsTrue(oLgFactTRP.idCmdCLT = objCMDCLT4.id Or oLgFactTRP.idCmdCLT = objCMDCLT5.id, "Id de commande")

        oLgFactTRP = oFactTRP1.colLignes(2)
        Assert.IsTrue(Not oLgFactTRP Is Nothing)
        Assert.IsTrue(oLgFactTRP.idCmdCLT = objCMDCLT4.id Or oLgFactTRP.idCmdCLT = objCMDCLT5.id, "Id de commande")

        'Controle des factures de transport générées CLT3
        oFactTRP1 = oColFactTRP(oCLT3.code)
        Assert.IsTrue(Not oFactTRP1 Is Nothing, "Facture du client3")
        Assert.IsTrue(oFactTRP1.colLignes.Count = 1, "1 lignes de factures dans la Facture 2")
        Assert.IsTrue(oFactTRP1.totalHT = objCMDCLT6.montantTransport + oFactTRP1.montantTaxes, "Montant HT de la commande")
        Assert.IsTrue(oFactTRP1.montantTaxes = Param.getConstante("CST_TAXES_TRP"), "Montant des taxes ")

        oLgFactTRP = oFactTRP1.colLignes(1)
        Assert.IsTrue(Not oLgFactTRP Is Nothing)
        Assert.IsTrue(oLgFactTRP.idCmdCLT = objCMDCLT6.id, "Id de commande")



        'Sauvegarde des factures de transports
        For Each oFactTRP1 In oColFactTRP
            Assert.IsTrue(oFactTRP1.Save, "Sauvegarde des factures")
        Next

        'Vérification qu'il n'y a plus de commandes à facturer 
        oColCMD = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.AreEqual(0, oColCMD.Count, "0 elements dans la liste")

        'Suppression des factures de transport
        For Each oFactTRP1 In oColFactTRP
            oFactTRP1.bDeleted = True
            Assert.IsTrue(oFactTRP1.Save, "Suppression  des factures")
        Next

        'Vérification que les commandes soient redevenues facturables 
        oColCMD = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.IsTrue(oColCMD.Count = 5, "5 elements dans la liste")

        'Suppression des commandes
        objCMDCLT1 = CommandeClient.createandload(nIdCMDCLT1)
        objCMDCLT1.bDeleted = True
        Assert.IsTrue(objCMDCLT1.save(), "Supression de la commande")
        objCMDCLT1 = CommandeClient.createandload(nIdCMDCLT2)
        objCMDCLT1.bDeleted = True
        Assert.IsTrue(objCMDCLT1.save(), "Supression de la commande")
        objCMDCLT1 = CommandeClient.createandload(nIdCMDCLT3)
        objCMDCLT1.bDeleted = True
        Assert.IsTrue(objCMDCLT1.save(), "Supression de la commande")
        objCMDCLT1 = CommandeClient.createandload(nIdCMDCLT4)
        objCMDCLT1.bDeleted = True
        Assert.IsTrue(objCMDCLT1.save(), "Supression de la commande")
        objCMDCLT1 = CommandeClient.createandload(nIdCMDCLT5)
        objCMDCLT1.bDeleted = True
        Assert.IsTrue(objCMDCLT1.save(), "Supression de la commande")
        objCMDCLT1 = CommandeClient.createandload(nIdCMDCLT6)
        objCMDCLT1.bDeleted = True
        Assert.IsTrue(objCMDCLT1.save(), "Supression de la commande")
        objCMDCLT1 = CommandeClient.createandload(nIdCMDCLT7)
        objCMDCLT1.bDeleted = True
        Assert.IsTrue(objCMDCLT1.save(), "Supression de la commande")

        'Suppression des Clients
        oCLT1 = Client.createandload(nIdCLT1)
        oCLT1.bDeleted = True
        oCLT1.save()
        oCLT1 = Client.createandload(nIdCLT2)
        oCLT1.bDeleted = True
        oCLT1.save()
        oCLT1 = Client.createandload(nIdCLT3)
        oCLT1.bDeleted = True
        oCLT1.save()

        resetTaxeGO()


    End Sub
    <TestMethod()> Public Sub T50_LISTE()
        Dim objFact As FactTRP
        Dim objFact2 As FactTRP
        Dim objFact3 As FactTRP
        Dim ocol As Collection
        Dim strCodeFact As String

        objFact = New FactTRP(m_oClient)
        objFact.dateCommande = "06/02/1964"
        objFact.totalHT = 110
        Assert.IsTrue(objFact.Save(), "Sauvegarde de la facture")
        strCodeFact = objFact.code

        objFact2 = New FactTRP(m_oClient2)
        objFact2.dateCommande = "07/02/1964"
        objFact2.totalHT = 120
        Assert.IsTrue(objFact2.Save(), "Sauvegarde de la facture")

        objFact3 = New FactTRP(m_oClient3)
        objFact3.dateCommande = "08/02/1964"
        objFact3.totalHT = 130
        Assert.IsTrue(objFact3.Save(), "Sauvegarde de la facture")


        ocol = FactTRP.getListe(strCodeFact)
        Assert.IsTrue(ocol.Count = 1, "1 elements dans la liste sur le code facture")

        ocol = FactTRP.getListe(, "MonClient2%")
        Assert.IsTrue(ocol.Count = 1, "1 elements dans la liste sur la code Client")

        ocol = FactTRP.getListe(CDate("01/02/64"), CDate("28/02/64"), , vncEnums.vncEtatCommande.vncFactTRPGeneree)
        Assert.IsTrue(ocol.Count >= 2, "au moins 2 elements dans la liste sur etat = Générée")


        ocol = FactTRP.getListe()
        Assert.IsTrue(ocol.Count >= 3, "3 elements dans la liste")

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), "Suppression de facture1")

        objFact2.bDeleted = True
        Assert.IsTrue(objFact2.Save(), "Suppression de facture2")

        objFact3.bDeleted = True
        Assert.IsTrue(objFact3.Save(), "Suppression de facture3")

    End Sub
    <TestMethod()> Public Sub T60_EXPORTCOMPTA()
        Dim objFact As FactTRP
        Dim objLgFact As LgFactTRP
        Dim nFile As Integer
        Dim v As String
        Dim strFile As String = "temp\export.txt"

        objFact = New FactTRP(m_oClient)
        Assert.IsTrue(objFact.Save(), "Sauvegarde de la facture")

        objLgFact = New LgFactTRP
        objLgFact.dateLivraison = "06/02/1964"
        objLgFact.nomTransporteur = "Rault"
        objLgFact.referenceLivraison = "AZERRTTY"
        objLgFact.prixHT = 150.55

        objFact.AjouteLigneFactTRP(objLgFact)
        If System.IO.File.Exists(strFile) Then
            System.IO.File.Delete(strFile)
        End If
        objFact.Exporter(strFile)
        Try
            nFile = FreeFile()
            FileOpen(nFile, strFile, OpenMode.Input, OpenAccess.Read)
            While Not EOF(nFile)
                v = LineInput(nFile)
                Assert.IsTrue(MsgBox(v, MsgBoxStyle.YesNo, "Ligne du fichier d'export") = MsgBoxResult.Yes)
            End While
            FileClose(nFile)
        Catch ex As Exception
            Assert.IsTrue(False, ex.ToString)
        End Try

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), "Suppression de la facture")

    End Sub
    <TestMethod(), Ignore()> Public Sub T70_CALCULTAXEGO()
        Dim objFact As FactTRP
        Dim objLgFact1 As LgFactTRP
        Dim objLgFact2 As LgFactTRP
        Dim objLgFact3 As LgFactTRP
        Dim objLgFact As LgFactTRP
        Dim nidFact As Integer

        setTaxeGO(51 / 10)

        objFact = New FactTRP(m_oClient)
        Assert.IsTrue(objFact.Save(), "Sauvegarde de la facture")

        objLgFact1 = New LgFactTRP
        objLgFact1.dateLivraison = "06/02/1964"
        objLgFact1.nomTransporteur = "Rault"
        objLgFact1.referenceLivraison = "AZERRTTY"
        objLgFact1.prixHT = 150.55
        objFact.AjouteLigneFactTRP(objLgFact1, False)

        objLgFact2 = New LgFactTRP
        objLgFact2.dateLivraison = "07/02/1964"
        objLgFact2.nomTransporteur = "Rault"
        objLgFact2.referenceLivraison = "AZERRTTY"
        objLgFact2.prixHT = 315.55

        objFact.AjouteLigneFactTRP(objLgFact2, False)

        Assert.IsTrue(objFact.Save(), "Sauvegarde de la facture")
        Assert.IsTrue(objFact.colLignes.Count = 2, "Avant le calcul ,2 lignes sur la facture")

        objFact.calculPrixTotal()
        Assert.IsTrue(objFact.colLignes.Count = 3, "Après le calcul ,3 lignes sur la facture")

        objLgFact = objFact.colLignes(3)
        '            Assert.IsTrue(Decimal.Round(objLgFact.prixHT, 2) = Decimal.Round((objLgFact1.prixHT + objLgFact2.prixHT) * (Param.getConstante("CST_TRP_TXGAZOLE") / 100), 2), "Montant Gasoil")
        Assert.IsTrue(Decimal.Round(objLgFact.prixHT, 2) = Decimal.Round(CDec((objLgFact1.prixHT + objLgFact2.prixHT) * (Param.getConstante("CST_TRP_TXGAZOLE") / 100)), 2), "Montant Gasoil")

        'Ajout d'une Troisième ligne avec recalcul automatique
        objLgFact3 = New LgFactTRP
        objLgFact3.dateLivraison = "06/02/1964"
        objLgFact3.nomTransporteur = "Rault"
        objLgFact3.referenceLivraison = "AZERRTTY"
        objLgFact3.prixHT = 150.55
        objFact.AjouteLigneFactTRP(objLgFact3)

        Assert.IsTrue(objFact.colLignes.Count = 4, "Après le calcul ,4 lignes sur la facture")

        objLgFact = objFact.colLignes(3)
        Assert.IsTrue(Decimal.Round(objLgFact.prixHT, 2) = Decimal.Round(CDec((objLgFact1.prixHT + objLgFact2.prixHT + objLgFact3.prixHT) * (Param.getConstante("CST_TRP_TXGAZOLE") / 100)), 2), "Montant Gasoil")


        'Sauvegarde de la facture avec ses lignes
        Assert.IsTrue(objFact.Save(), "Sauvegarde de la facture avec 4 lignes")
        nidFact = objFact.id
        objFact = FactTRP.createandload(nidFact)
        Assert.IsTrue(objFact.loadcolLignes(), "Chargement des lignes de la facture")
        Assert.IsTrue(objFact.colLignes.Count = 4, "Après le Chargement ,4 lignes sur la facture")
        'La Ligne de GO est passée en Dernière
        objLgFact = objFact.colLignes(4)
        Assert.AreEqual(objLgFact.num, CInt("999999"), "Numero de ligne GO")

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), "Suppression de la commande")

        '=================================================================================
        ' Test avec la Constante = 0
        '=================================================================================
        setTaxeGOEqualZero()
        objFact = New FactTRP(m_oClient)
        Assert.IsTrue(objFact.Save(), "Sauvegarde de la facture")

        objLgFact1 = New LgFactTRP
        objLgFact1.dateLivraison = "06/02/1964"
        objLgFact1.nomTransporteur = "Rault"
        objLgFact1.referenceLivraison = "AZERRTTY"
        objLgFact1.prixHT = 150.55
        objFact.AjouteLigneFactTRP(objLgFact1, False)

        objLgFact2 = New LgFactTRP
        objLgFact2.dateLivraison = "07/02/1964"
        objLgFact2.nomTransporteur = "Rault"
        objLgFact2.referenceLivraison = "AZERRTTY"
        objLgFact2.prixHT = 315.55

        objFact.AjouteLigneFactTRP(objLgFact2, False)

        objFact.calculPrixTotal()
        Assert.IsTrue(objFact.colLignes.Count = 2, "Après le calcul ,2 lignes sur la facture")

        Assert.IsTrue(Decimal.Round(objFact.totalHT, 2) = Decimal.Round((objLgFact1.prixHT + objLgFact2.prixHT) + objFact.montantTaxes, 2), "Montant Facture = Montants Lignes + taxes")

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save(), "Suppression de la commande")
        resetTaxeGO()

        '======================================================================
        ' Test calsult Taxe Gazole avec génération
        '======================================================================


        'Creation des Commandes
        Dim objCMDCLT1 As CommandeClient
        Dim objCMDCLT2 As CommandeClient
        Dim nidCMDCLT1 As Integer
        Dim nidCMDCLT2 As Integer
        Dim oColCMD As Collection
        Dim oColFactTRP As ColEvent

        '2+1 Commandes pour le Client1
        objCMDCLT1 = New CommandeClient(m_oClient)
        objCMDCLT1.dateLivraison = "06/02/1964"
        objCMDCLT1.bFactTransport = True
        objCMDCLT1.montantTransport = 101.5
        objCMDCLT1.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT1.save(), "Sauvegarde de la premieère commande")
        nidCMDCLT1 = objCMDCLT1.id

        objCMDCLT2 = New CommandeClient(m_oClient)
        objCMDCLT2.dateLivraison = "06/02/1964"
        objCMDCLT2.bFactTransport = True
        objCMDCLT2.montantTransport = 201.5
        objCMDCLT2.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMDCLT2.save(), "Sauvegarde de la 2eme commande")
        nidCMDCLT2 = objCMDCLT2.id


        'Recuperation des commande avec transport 
        oColCMD = CommandeClient.getListe_TRP("06/02/1964", "06/02/1964")
        Assert.IsTrue(oColCMD.Count = 2, "2 elements sont Attendus dans la liste->" & oColCMD.Count)

        'Génération des gfactures de transport
        oColFactTRP = FactTRP.createFactTRPs(oColCMD, "29/02/1964", , "FEVRIER 64")
        Assert.IsTrue(oColFactTRP.Count = 1, "1 Facture Générée")


        'Controle des factures de transport générées
        objFact = oColFactTRP(m_oClient.code)
        Assert.IsTrue(Not objFact Is Nothing, "Facture du client1")
        Assert.IsTrue(objFact.colLignes.Count = 3, "3 lignes de factures dans la Commande 1")
        Assert.IsTrue(objFact.montantTaxes = Param.getConstante("CST_TAXES_TRP") * 2, "Montant des taxes ")

        'Controle de la ligne GO
        objLgFact = objFact.colLignes("999999")
        Assert.AreEqual(Decimal.Round(objLgFact.prixHT, 2), Decimal.Round(CDec((objCMDCLT1.montantTransport + objCMDCLT2.montantTransport) * (Param.getConstante("CST_TRP_TXGAZOLE") / 100)), 2), "Montant Gasoil")
        Assert.AreEqual(Decimal.Round(objFact.totalHT, 2), Decimal.Round(objCMDCLT1.montantTransport + objCMDCLT2.montantTransport + objFact.montantTaxes + objLgFact.prixHT, 2), "Montant HT de la commande")


        'Suppression des commandes
        objCMDCLT1 = CommandeClient.createandload(nidCMDCLT1)
        objCMDCLT1.bDeleted = True
        Assert.IsTrue(objCMDCLT1.save(), "Supression de la commande")
        objCMDCLT1 = CommandeClient.createandload(nidCMDCLT2)
        objCMDCLT1.bDeleted = True

        resetTaxeGO()

        objFact.bDeleted = True
        objFact.Save()



    End Sub

    Private Sub setTaxeGOEqualZero()
        m_nTXGO = Param.getConstante("CST_TRP_TXGAZOLE")
        'CONSTANTE GO = 0
        Persist.shared_connect()
        Persist.executeSQLNonQuery("UPDATE CONSTANTES SET CST_TRP_TXGAZOLE = 0")
        Persist.shared_disconnect()
        Param.LoadcolParams()
    End Sub
    Private Sub setTaxeGO(ByVal pValue As Decimal)
        m_nTXGO = Param.getConstante("CST_TRP_TXGAZOLE")
        'CONSTANTE GO = 0
        Persist.shared_connect()
        Dim strSQL As String
        Dim strValue As String
        strValue = pValue
        strValue = strValue.Replace(",", ".")

        strSQL = "UPDATE CONSTANTES SET CST_TRP_TXGAZOLE = " & strValue

        Call Persist.executeSQLNonQuery(strSQL)

        Persist.shared_disconnect()
        Param.LoadcolParams()
    End Sub
    Private Sub resetTaxeGO()
        'CONSTANTE GO = 0
        Dim strSQL As String
        Dim strValue As String
        strValue = m_nTXGO
        strValue = strValue.Replace(",", ".")

        strSQL = "UPDATE CONSTANTES SET CST_TRP_TXGAZOLE = " & strValue
        Persist.shared_connect()
        Persist.executeSQLNonQuery(strSQL)
        Persist.shared_disconnect()
        Param.LoadcolParams()
    End Sub

    ''' <summary>
    ''' Test l'export vers Quadra
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> Public Sub T100_EXPORT()
        Dim objFact As FactTRP
        Dim strLines As String()
        Dim strLine1 As String
        Dim strLine2 As String
        Dim strLine3 As String

        objFact = New FactTRP(m_oClient)
        objFact.periode = "1er Timestre 1964"
        objFact.dateFacture = CDate("06/02/1964")
        objFact.totalHT = 150.56
        objFact.totalTTC = 180.89
        objFact.dEcheance = "01/04/1964"


        Assert.IsTrue(objFact.Save(), objFact.getErreur())




        'Save
        Assert.IsTrue(objFact.Save(), "Insert" & objFact.getErreur)
        If File.Exists("./T20_EXPORT.txt") Then
            File.Delete("./T20_EXPORT.txt")
        End If

        objFact.exporter("./T20_EXPORT.txt")

        Assert.IsTrue(File.Exists("./T20_EXPORT.txt"), "le fichier d'export n'existe pas")
        strLines = File.ReadAllLines("./T20_EXPORT.txt")
        Assert.AreEqual(3, strLines.Length, "3 lignes d'export")

        strLine1 = strLines(0)
        strLine2 = strLines(1)
        strLine3 = strLines(2)

        Assert.AreEqual(231, strLine1.Length)
        Assert.AreEqual("M", strLine1.Substring(0, 1))
        Assert.AreEqual(m_oClient.CodeCompta, Trim(strLine1.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine1.Substring(9, 2)))
        Assert.AreEqual("060264", strLine1.Substring(14, 6))
        Assert.AreEqual(("F:" + objFact.code + " " + m_oClient.rs + Space(21)).Substring(0, 20), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("D", strLine1.Substring(41, 1))
        Assert.AreEqual((180.89).ToString("0000000000.00").Replace(".", ""), Trim(strLine1.Substring(42, 13)))
        Assert.AreEqual("010464", Trim(strLine1.Substring(63, 6)))

        Assert.AreEqual(231, strLine2.Length)
        Assert.AreEqual("M", strLine2.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTETVA")), Trim(strLine2.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine2.Substring(9, 2)))
        Assert.AreEqual("060264", strLine2.Substring(14, 6))
        Assert.AreEqual(("F:" + objFact.code + " " + m_oClient.rs + Space(21)).Substring(0, 20), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine2.Substring(41, 1))
        Assert.AreEqual((180.89 - 150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine2.Substring(42, 13)))

        Assert.AreEqual(231, strLine3.Length)
        Assert.AreEqual("M", strLine3.Substring(0, 1))
        Assert.AreEqual(Trim(Param.getConstante("CST_SOC2_COMPTEPRODUIT")), Trim(strLine3.Substring(1, 8)))
        Assert.AreEqual("VE", Trim(strLine3.Substring(9, 2)))
        Assert.AreEqual("060264", strLine3.Substring(14, 6))
        Assert.AreEqual(("F:" + objFact.code + " " + m_oClient.rs + Space(21)).Substring(0, 20), Trim(strLine1.Substring(21, 20)))
        Assert.AreEqual("C", strLine3.Substring(41, 1))
        Assert.AreEqual((150.56).ToString("0000000000.00").Replace(".", ""), Trim(strLine3.Substring(42, 13)))

        objFact.bDeleted = True
        Assert.IsTrue(objFact.Save())
    End Sub
    <TestMethod()> Public Sub T30_ChampsLongs()

        Dim objFACT As FactTRP

        'I - Création d'une Facture 
        '=========================
        objFACT = New FactTRP(m_oClient)

        objFACT.periode = "1er Timestre 1964".PadRight(500, "x")
        'Save
        Assert.IsTrue(objFACT.Save(), "Insert" & objFACT.getErreur)

        objFACT.load()

        Assert.AreEqual(50, objFACT.periode.Length)
        objFACT.bDeleted = True
        Assert.IsTrue(objFACT.Save())
    End Sub
    'Test l'incrémenation des codes
    <TestMethod()> Public Sub T70_GetNextCode()

        Dim obj1 As New FactTRP(m_oClient)
        Dim obj2 As New FactTRP(m_oClient)

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



