'Test de la classe dbConnection
Imports System
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vini_DB



<TestClass()> Public Class test_V17
    Inherits test_Base
    'Ignore("A tester une une base propre")
    <TestMethod(), Ignore()> Public Sub T10_PurgeStock()
        Dim objFournisseur As Fournisseur
        Dim objP1 As Produit

        Persist.shared_connect()
        'Création d'un fournisseur
        objFournisseur = New Fournisseur("FRNT10v17", "FOURNISSEUR T10 V1.7")
        Assert.IsTrue(objFournisseur.Save())

        'Création de 5 Produits
        objP1 = New Produit("PRDT10V17", objFournisseur, 2000)
        Assert.IsTrue(objP1.save())

        objP1.ajouteLigneMvtStock("01/01/2001", vncEnums.vncTypeMvt.vncMvtInventaire, 0, "INVENTAIRE AU 1/1/2001", 12)
        objP1.ajouteLigneMvtStock("15/01/2001", vncEnums.vncTypeMvt.vncMvtCommandeClient, 0, "Commande de 10 Pièces", -10)
        objP1.ajouteLigneMvtStock("31/01/2001", vncEnums.vncTypeMvt.vncmvtBonAppro, 0, "Appro de 10 Pièces", 10)
        Assert.IsTrue(objP1.save())
        Assert.AreEqual(12, objP1.QteStock)
        objP1.ajouteLigneMvtStock("01/02/2001", vncEnums.vncTypeMvt.vncMvtInventaire, 0, "INVENTAIRE AU 1/1/2001", 15)
        objP1.ajouteLigneMvtStock("15/02/2001", vncEnums.vncTypeMvt.vncMvtCommandeClient, 0, "Commande de 100 Pièces", -100)
        objP1.ajouteLigneMvtStock("28/02/2001", vncEnums.vncTypeMvt.vncmvtBonAppro, 0, "Appro de 100 Pièces", 100)
        Assert.IsTrue(objP1.save())
        Assert.IsTrue(objP1.colmvtStock.Count = 6, "Il y a 6 mvts de Stocks")
        Assert.AreEqual(15, objP1.QteStock)
        Assert.IsTrue(objP1.purgeMvtStock("01/02/2001"))
        Assert.IsTrue(objP1.load())
        Assert.IsTrue(objP1.loadcolmvtStock())
        Assert.IsTrue(objP1.colmvtStock.Count = 3, "Il y a 3 mvts de stocks")
        Assert.AreEqual(15, objP1.QteStock)

        objP1.bDeleted = False
        Assert.IsTrue(objP1.save())
        objFournisseur.bDeleted = True
        Assert.IsTrue(objFournisseur.Save())
    End Sub
    'Ignore("A tester une une base propre")
    <TestMethod(), Ignore()> Public Sub T20_PurgeFactCom()
        Dim objClient As Client
        Dim objFournisseur As Fournisseur
        Dim objP1 As Produit
        Dim objCMD1 As CommandeClient
        Dim objCMD2 As CommandeClient
        Dim objCMD3 As CommandeClient
        Dim nid1 As Integer
        Dim nid2 As Integer
        Dim nid3 As Integer
        Dim nidFactCom As Integer
        Dim nidFactCom3 As Integer

        Dim objSCMD As SousCommande
        Dim objCol As Collection
        Dim objColF As ColEvent
        Dim objFactCom1 As FactCom
        Dim objFactCom2 As FactCom
        Dim objFactCom3 As FactCom
        Dim colFact As Collection

        Persist.shared_connect()
        'Création d'un fournisseur
        objFournisseur = New Fournisseur("FRNT20v17", "FOURNISSEUR T20 V1.7")
        Assert.IsTrue(objFournisseur.Save())

        'Création d'un Client
        objClient = New Client("CLTT20v17", "FOURNISSEUR T20 V1.7")
        Assert.IsTrue(objClient.save())

        'Création de 5 Produits
        objP1 = New Produit("PRDT20V17", objFournisseur, 2000)
        Assert.IsTrue(objP1.save())

        'Création de 3 commandes Client
        objCMD1 = New CommandeClient(objClient)
        objCMD1.dateCommande = "01/01/2001"
        objCMD1.AjouteLigne("10", objP1, 1, 1)
        Assert.IsTrue(objCMD1.save())
        nid1 = objCMD1.id

        objCMD2 = New CommandeClient(objClient)
        objCMD2.dateCommande = "01/01/2001"
        objCMD2.AjouteLigne("20", objP1, 2, 2)
        Assert.IsTrue(objCMD2.save())
        nid2 = objCMD2.id

        objCMD3 = New CommandeClient(objClient)
        objCMD3.dateCommande = "01/01/2004"
        objCMD3.AjouteLigne("30", objP1, 3, 3)
        Assert.IsTrue(objCMD3.save())
        nid3 = objCMD3.id

        'Livraison des 2 commandes
        objCMD1.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCMD1.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD1.save())
        objCMD2.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCMD2.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD2.save())
        objCMD3.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCMD3.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD3.save())

        'Eclatement des 2 Commandes
        Assert.IsTrue(objCMD1.generationSousCommande())
        Assert.IsTrue(objCMD1.save())

        Assert.IsTrue(objCMD2.generationSousCommande())
        Assert.IsTrue(objCMD2.save())

        Assert.IsTrue(objCMD3.generationSousCommande())
        Assert.IsTrue(objCMD3.save())

        'Transmission des sous-commandes de la 1ere commande
        Assert.IsTrue(objCMD1.LoadColSousCommande())
        For Each objSCMD In objCMD1.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
            Assert.IsTrue(objSCMD.Save())
        Next
        Assert.IsTrue(objCMD2.LoadColSousCommande())
        For Each objSCMD In objCMD2.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
            Assert.IsTrue(objSCMD.Save())
        Next

        Assert.IsTrue(objCMD3.LoadColSousCommande())
        For Each objSCMD In objCMD3.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
            Assert.IsTrue(objSCMD.Save())
        Next

        'Rapprochement des sous-commandes de la 1ere commande
        For Each objSCMD In objCMD1.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
            Assert.IsTrue(objSCMD.Save())
        Next
        For Each objSCMD In objCMD2.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
            Assert.IsTrue(objSCMD.Save())
        Next
        For Each objSCMD In objCMD3.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
            Assert.IsTrue(objSCMD.Save())
        Next



        'Creation des Factures de Commissions
        objCol = New Collection
        For Each objSCMD In objCMD1.colSousCommandes
            objCol.Add(objSCMD)
        Next
        objColF = FactCom.createFactComs(objCol, objCMD1.dateCommande, objCMD1.dateCommande, objCMD1.dateCommande.ToString)
        objFactCom1 = objColF(1)
        Assert.IsTrue(objFactCom1.Save())

        'Creation des Factures de Commissions
        objCol = New Collection
        For Each objSCMD In objCMD2.colSousCommandes
            objCol.Add(objSCMD)
        Next
        objColF = FactCom.createFactComs(objCol, objCMD2.dateCommande, objCMD2.dateCommande, objCMD2.dateCommande.ToString)
        objFactCom2 = objColF(1)
        Assert.IsTrue(objFactCom2.Save())

        'Creation des Factures de Commissions
        objCol = New Collection
        For Each objSCMD In objCMD3.colSousCommandes
            objCol.Add(objSCMD)
        Next
        objColF = FactCom.createFactComs(objCol, objCMD3.dateCommande, objCMD3.dateCommande, objCMD3.dateCommande.ToString)
        objFactCom3 = objColF(1)
        Assert.IsTrue(objFactCom3.Save())


        'Rapprochement Des Factures de commission
        objFactCom1.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactComExporter)
        Assert.IsTrue(objFactCom1.Save())
        nidFactCom = objFactCom1.id

        objFactCom3.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactComExporter)
        Assert.IsTrue(objFactCom3.Save())
        nidFactCom3 = objFactCom3.id

        colFact = FactCom.getListe("01/01/2000", "01/01/2001", , vncEnums.vncEtatCommande.vncFactComExportee)
        Assert.IsTrue(colFact.Count = 1)
        Assert.AreEqual(colFact(1).code, objFactCom1.code)

        Assert.IsTrue(objFactCom1.purge())
        Assert.IsTrue(MsgBox("Vérifier qu'il n'y ait plus d'enr dans la table FACTCOM pour L'id " & nidFactCom, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)
        Assert.IsTrue(MsgBox("Vérifier qu'il n'y ait plus d'enr dans la table SOUSCOMMANDE pour L'idCommande " & nid1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)
        Assert.IsTrue(MsgBox("Vérifier qu'il y ait TOUJOURS 1 enr dans la table COMMANDE pour L'id " & nid1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)
        Assert.IsTrue(MsgBox("Vérifier qu'il ait TOUJOURS des enrs dans la table LIGNE_COMMANDE pour L'LGCM_CMD_ID " & nid1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

        Assert.IsTrue(MsgBox("Vérifier qu'il y ait 1 enr dans la table FACTCOM pour L'id " & nidFactCom3, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

        objFactCom2.bDeleted = True
        objFactCom2.Save()

        objFactCom3.bDeleted = True
        objFactCom3.Save()

        objCMD3.bDeleted = True
        objCMD3.save()
        objCMD2.bDeleted = True
        objCMD2.save()
        objCMD1.bDeleted = True
        objCMD1.save()

        objP1.bDeleted = True
        objP1.save()

        objFournisseur.bDeleted = True
        objFournisseur.Save()

        objClient.bDeleted = True
        objClient.save()

    End Sub
    '"A tester une une base propre"
    <TestMethod(), Ignore()> Public Sub T30_PurgeCommande()
        Dim objClient As Client
        Dim objFournisseur As Fournisseur
        Dim objP1 As Produit
        Dim objCMD1 As CommandeClient
        Dim objCMD2 As CommandeClient
        Dim objCMD3 As CommandeClient
        Dim nid1 As Integer
        Dim nid2 As Integer
        Dim nid3 As Integer
        Dim nidFactCom As Integer

        Dim objSCMD As SousCommande
        Dim objCol As Collection
        Dim objColF As ColEvent
        Dim objFactCom1 As FactCom
        Dim objFactCom2 As FactCom
        Dim colCmd As Collection

        Persist.shared_connect()
        'Création d'un fournisseur
        objFournisseur = New Fournisseur("FRNT20v17", "FOURNISSEUR T20 V1.7")
        Assert.IsTrue(objFournisseur.Save())

        'Création d'un Client
        objClient = New Client("CLTT20v17", "FOURNISSEUR T20 V1.7")
        Assert.IsTrue(objClient.save())

        'Création de 5 Produits
        objP1 = New Produit("PRDT20V17", objFournisseur, 2000)
        Assert.IsTrue(objP1.save())

        'Création de 3 commandes Client
        objCMD1 = New CommandeClient(objClient)
        objCMD1.dateCommande = "01/01/2001"
        objCMD1.AjouteLigne("10", objP1, 1, 1)
        Assert.IsTrue(objCMD1.save())
        nid1 = objCMD1.id

        objCMD2 = New CommandeClient(objClient)
        objCMD2.dateCommande = "01/01/2001"
        objCMD2.AjouteLigne("20", objP1, 2, 2)
        Assert.IsTrue(objCMD2.save())
        nid2 = objCMD2.id

        objCMD3 = New CommandeClient(objClient)
        objCMD3.dateCommande = "01/01/2001"
        objCMD3.AjouteLigne("30", objP1, 3, 3)
        Assert.IsTrue(objCMD3.save())
        nid3 = objCMD3.id

        'Livraison des 2 commandes
        objCMD1.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCMD1.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD1.save())
        objCMD2.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCMD2.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD2.save())
        objCMD3.changeEtat(vncEnums.vncActionEtatCommande.vncActionValider)
        objCMD3.changeEtat(vncEnums.vncActionEtatCommande.vncActionLivrer)
        Assert.IsTrue(objCMD3.save())

        'Eclatement des 2 Commandes
        Assert.IsTrue(objCMD1.generationSousCommande())
        Assert.IsTrue(objCMD1.save())

        Assert.IsTrue(objCMD2.generationSousCommande())
        Assert.IsTrue(objCMD2.save())

        Assert.IsTrue(objCMD3.generationSousCommande())
        Assert.IsTrue(objCMD3.save())

        'Transmission des sous-commandes de la 1ere commande
        Assert.IsTrue(objCMD1.LoadColSousCommande())
        For Each objSCMD In objCMD1.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
            Assert.IsTrue(objSCMD.Save())
        Next
        Assert.IsTrue(objCMD2.LoadColSousCommande())
        For Each objSCMD In objCMD2.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
            Assert.IsTrue(objSCMD.Save())
        Next

        Assert.IsTrue(objCMD3.LoadColSousCommande())
        For Each objSCMD In objCMD3.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
            Assert.IsTrue(objSCMD.Save())
        Next

        'Rapprochement des sous-commandes de la 1ere commande
        For Each objSCMD In objCMD1.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
            Assert.IsTrue(objSCMD.Save())
        Next
        For Each objSCMD In objCMD2.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
            Assert.IsTrue(objSCMD.Save())
        Next
        For Each objSCMD In objCMD3.colSousCommandes
            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
            Assert.IsTrue(objSCMD.Save())
        Next



        'Creation des Factures de Commissions
        objCol = New Collection
        For Each objSCMD In objCMD1.colSousCommandes
            objCol.Add(objSCMD)
        Next
        objColF = FactCom.createFactComs(objCol, objCMD1.dateCommande, objCMD1.dateCommande, objCMD1.dateCommande.ToString)
        objFactCom1 = objColF(1)
        Assert.IsTrue(objFactCom1.Save())

        'Creation des Factures de Commissions
        objCol = New Collection
        For Each objSCMD In objCMD2.colSousCommandes
            objCol.Add(objSCMD)
        Next
        objColF = FactCom.createFactComs(objCol, objCMD2.dateCommande, objCMD2.dateCommande, objCMD2.dateCommande.ToString)
        objFactCom2 = objColF(1)
        Assert.IsTrue(objFactCom2.Save())

        'Creation des Factures de Commissions
        'objCol = New Collection
        'For Each objSCMD In objCMD3.colSousCommandes
        'objCol.Add(objSCMD)
        'Next
        'objColF = FactCom.createFactComs(objCol, objCMD3.dateCommande, objCMD3.dateCommande, objCMD3.dateCommande.ToString)
        'objFactCom3 = objColF(1)
        'Assert.IsTrue(objFactCom3.Save())


        'Rapprochement Des Factures de commission
        objFactCom1.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactComExporter)
        Assert.IsTrue(objFactCom1.Save())
        nidFactCom = objFactCom1.id

        'objFactCom3.changeEtat(vncEnums.vncActionEtatCommande.vncActionFactComRapprocher)
        'Assert.IsTrue(objFactCom3.Save())
        'nidFactCom3 = objFactCom3.id

        colCmd = CommandeClient.getListe("01/01/2000", "01/01/2001", , vncEnums.vncEtatCommande.vncEclatee)
        Try
            objCMD1 = colCmd(objCMD1.code)
        Catch ex As Exception
            Assert.IsTrue(False, "la Commande" & objCMD1.code & "n'est pas dans la collection")
        End Try
        'La Commande 3 doit être dans la liste
        Try
            objCMD3 = colCmd(objCMD3.code)
        Catch ex As Exception
            Assert.IsTrue(False, "la Commande" & objCMD3.code & "n'est pas dans la collection")
        End Try

        Assert.IsTrue(objCMD1.purge())

        Assert.IsTrue(objCMD3.purge())
        Assert.IsTrue(objCMD3.id <> 0, "La commande " & objCMD3.id & "n'est pas purgeable")
        Assert.IsTrue(MsgBox("Vérifier qu'il n'y ait plus d' enr dans la table COMMANDE pour L'id " & nid1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)
        Assert.IsTrue(MsgBox("Vérifier qu'il n'y ait plus des enrs dans la table LIGNE_COMMANDE pour L'LGCM_CMD_ID " & nid1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)
        Assert.IsTrue(MsgBox("Vérifier qu'il y ait TOUJOURS des enrs dans la table SOUS_COMMANDE pour L'SCMD_CMD_ID " & nid1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

        objCMD2.bDeleted = True
        objCMD2.save()

        objCMD3.bDeleted = True
        objCMD3.save()

        objP1.bDeleted = True
        objP1.save()

        objFournisseur.bDeleted = True
        objFournisseur.Save()

        objClient.bDeleted = True
        objClient.save()

    End Sub
    '"A tester une une base propre"
    <TestMethod(), Ignore()> Public Sub T40_PurgeBA()
        Dim objClient As Client
        Dim objFournisseur As Fournisseur
        Dim objP1 As Produit
        Dim objCMD1 As BonAppro
        Dim objCMD2 As BonAppro
        Dim objCMD3 As BonAppro
        Dim nid1 As Integer
        Dim nid2 As Integer
        Dim nid3 As Integer
        Dim colCmd As Collection

        Persist.shared_connect()
        'Création d'un fournisseur
        objFournisseur = New Fournisseur("FRNT40v17", "FOURNISSEUR T40 V1.7")
        Assert.IsTrue(objFournisseur.Save())

        'Création d'un Client
        objClient = New Client("CLTT40v17", "FOURNISSEUR T40 V1.7")
        Assert.IsTrue(objClient.save())

        'Création de 5 Produits
        objP1 = New Produit("PRDT40V17", objFournisseur, 2000)
        Assert.IsTrue(objP1.save())

        'Création de 3 BA
        objCMD1 = New BonAppro(objFournisseur)
        objCMD1.dateCommande = "01/01/2001"
        objCMD1.AjouteLigne("10", objP1, 1, 1)
        Assert.IsTrue(objCMD1.save())
        nid1 = objCMD1.id

        objCMD2 = New BonAppro(objFournisseur)
        objCMD2.dateCommande = "01/01/2001"
        objCMD2.AjouteLigne("20", objP1, 2, 2)
        Assert.IsTrue(objCMD2.save())
        nid2 = objCMD2.id

        objCMD3 = New BonAppro(objFournisseur)
        objCMD3.dateCommande = "01/01/2004"
        objCMD3.AjouteLigne("30", objP1, 3, 3)
        Assert.IsTrue(objCMD3.save())
        nid3 = objCMD3.id

        'Livraison des 2 BA
        objCMD1.changeEtat(vncEnums.vncActionEtatCommande.vncActionBALivrer)
        Assert.IsTrue(objCMD1.save())
        '            objCMD2.changeEtat(vncEnums.vncActionEtatCommande.vncActionBALivrer)
        '            Assert.IsTrue(objCMD2.save())
        objCMD3.changeEtat(vncEnums.vncActionEtatCommande.vncActionBALivrer)
        Assert.IsTrue(objCMD3.save())

        colCmd = BonAppro.getListe("01/01/2000", "01/01/2001", , vncEnums.vncEtatCommande.vncBALivre)
        Assert.IsTrue(colCmd.Count = 1)
        Try
            objCMD1 = colCmd(objCMD1.code)
        Catch ex As Exception
            Assert.IsTrue(False, "la Commande" & objCMD1.code & "n'est pas dans la collection")
        End Try

        Assert.IsTrue(objCMD1.purge())
        Assert.IsTrue(MsgBox("Vérifier qu'il n'y ait plus d' enr dans la table BONAPPRO pour L'id " & nid1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)
        Assert.IsTrue(MsgBox("Vérifier qu'il n'y ait plus des enrs dans la table LIGNE_COMMANDE pour L'ID_BA " & nid1, MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

        objCMD2.bDeleted = True
        objCMD2.save()

        objCMD3.bDeleted = True
        objCMD3.save()

        objP1.bDeleted = True
        objP1.save()

        objFournisseur.bDeleted = True
        objFournisseur.Save()

        objClient.bDeleted = True
        objClient.save()

    End Sub
End Class



