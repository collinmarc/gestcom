'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : BonAppro   
' Description : Bon D'approvisionnement
'               Hérite de Commande
'===================================================================================================================================
'Membres de Classes
'==================
'Public
'Protected
'Private
'Membres d'instances
'==================
'Public
'   toString()      : Rend 'objet sous forme de chaine
'   Equals()        : Test l'équivalence d'instance
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================Public MustInherit Class Persist
Public Class BonAppro
    Inherits Commande
#Region "Membres Privés"
    '=================== MEMBRES PRIVES ====================================
#End Region
#Region "Accesseurs"
    Public Sub New(ByVal poFournisseur As Fournisseur)
        MyBase.New(poFournisseur, New EtatBAEnCoursSaisie)
        m_typedonnee = vncEnums.vncTypeDonnee.BA
        Debug.Assert(Not m_oTransporteur Is Nothing)
        Debug.Assert(Not etat Is Nothing)
        majBooleenAlaFinDuNew()
    End Sub
    Friend Sub New()
        MyBase.New(New Fournisseur, New EtatBAEnCoursSaisie)
        m_typedonnee = vncEnums.vncTypeDonnee.BA
        Debug.Assert(Not m_oTransporteur Is Nothing)
        Debug.Assert(Not etat Is Nothing)
        majBooleenAlaFinDuNew()
    End Sub

#End Region
#Region "Méthodes de Classe"
    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListe 
    'Description : Liste des Bons d'Approvisionnement
    'Retour : Rend une collection de Clients
    '=======================================================================
    Public Shared Function getListe(Optional ByVal strCode As String = "", Optional ByVal strRSFourn As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection
        Dim objBidon As New BonAppro(New Fournisseur)

        Persist.shared_connect()
        colReturn = ListeBA(strCode, strRSFourn, pEtat)
        Persist.shared_disconnect()
        Return colReturn
    End Function
    Public Shared Function getListe(ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pRSFourn As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeBAEtat(pddeb, pdfin, pRSFourn, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "BaonAppro.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

    End Function 'getListe
    Public Shared Function createandload(ByVal pid As Integer) As BonAppro
        '=======================================================================
        ' Constructeur pour chargement
        '=======================================================================
        Dim objBA As BonAppro
        objBA = New BonAppro
        Try
            If pid <> 0 Then
                objBA.load(pid)
            End If
        Catch ex As Exception
            setError("CommandeClient.createAndLoad", ex.ToString)
        End Try
        Debug.Assert(objBA.id = pid, "Bon Appro " & pid & " non chargé")
        Return objBA
    End Function 'createandload

#End Region
#Region "Interface Persist"
    '=======================================================================
    'Fonction : DBLoad()
    'Description : Chargement de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If

        Debug.Assert(id <> 0, "idCommande <> 0")
        bReturn = LoadBA()
        If bReturn Then
            'Chargement des caractéristiques du client
            bReturn = oTiers.load()
            Debug.Assert(bReturn, Tiers.getErreur())
        End If
        m_colLignes = New ColEvent
        m_bcolLgLoaded = False
        bReturn = LoadcolLGCMD()
        m_bcolLgLoaded = bReturn
        Return bReturn
    End Function 'DBLoad
    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = MyBase.Save()
        If m_bcolLgLoaded Then
            bReturn = savecolLignes()
        End If
        'MVTS STOCKS
        If getActionMvtStock() = vncEnums.vncGenererSupprimer.vncGenerer Then
            bReturn = genereMvtStock()
            Debug.Assert(bReturn, "Erreur en generemvtStock" & getErreur())
        End If
        If getActionMvtStock() = vncEnums.vncGenererSupprimer.vncSupprimer Then
            bReturn = supprimeMvtStock()
            Debug.Assert(bReturn, "Erreur en generemvtStock" & getErreur())
        End If
        shared_disconnect()
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : delete()
    'Description : Suppression de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai si l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Debug.Assert(id <> 0, "idBA <> 0")
        Dim bReturn As Boolean
        'Pas de suppression des Mvts de stocks

        'Suppression des Lignes de commandes
        bReturn = deleteBAcolLgCMD()
        If bReturn Then
            bReturn = deleteBA
            m_bcolLgLoaded = False
        End If
        Return bReturn

    End Function ' delete
    '=======================================================================
    'Fonction : checkFordelete
    'description : Controle si l'élément est supprimable
    '                Existance de sous-commandes
    '=======================================================================
    Public Overrides Function checkForDelete() As Boolean
        Dim bReturn As Boolean
        Try
            shared_connect()
            bReturn = True
            '            If existeSousCommandeCommande() Then
            '            bReturn = False
            '            End If
            shared_disconnect()
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'checkForDelete

    '=======================================================================
    'Fonction : insert()
    'Description : insertion de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Fournisseur n'est pas Renseigné")
        Debug.Assert(code.Equals(""), "Le Code doit être nul")
        Debug.Assert(oTiers.id <> 0, "Le Fournisseur n'est pas sauvegardé")
        Debug.Assert(id = 0, "idBA = 0")

        Dim bReturn As Boolean
        If setNewcode() Then
            bReturn = insertBA()
        End If
        Return bReturn
    End Function 'insert
    '=======================================================================
    'Fonction : Update()
    'Description : Mise à jour de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Fournisseur n'est pas Renseigné")
        Debug.Assert(code <> "", "Le Code ne doit pas être nul")
        Debug.Assert(oTiers.id <> 0, "Le Fournisseur n'est pas sauvegardé")
        Debug.Assert(id <> 0, "idBA <> 0")
        Dim bReturn As Boolean
        bReturn = UpdateBA
        Return bReturn

    End Function 'Update

#End Region
#Region "Méthodes Overrides"
    Public Overrides Sub Exporter(ByVal pfileNAme As String)
        ' no export for bonAppro
    End Sub

    Friend Overrides Function setNewcode() As Boolean
        Dim str As String = ""
        Dim ncode As Integer
        Dim breturn As Boolean

        shared_connect()
        ncode = getNumeroBonAppro()
        shared_disconnect()
        If ncode = -1 Then
            breturn = False
        Else
            code = str & CStr(ncode)
            breturn = True
        End If
        Return breturn
    End Function 'setnewCode

    Public Overrides Function changeEtat(ByVal p_action As vncActionEtatCommande) As Boolean
        Debug.Assert(p_action >= vncActionEtatCommande.vncActionMinBA And p_action <= vncActionEtatCommande.vncActionMaxBA)

        Dim bReturn As Boolean = False
        Try
            If p_action >= vncActionEtatCommande.vncActionMinBA And p_action <= vncActionEtatCommande.vncActionMaxBA Then
                MyBase.changeEtat(p_action)
                RaiseUpdated()
                bReturn = True
            Else
                setError("BonAppro.changeEtat", "Code Action incorrect :" + p_action)
                bReturn = False
            End If

        Catch ex As Exception
            setError("BonAppro.changeEtat", ex.ToString)
            bReturn = False
        End Try

        Return bReturn
    End Function 'ChangeEtat


    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return MyBase.toString()
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objCommande As BonAppro
        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn

    End Function 'Equals
    '=======================================================================
    'Function : GenereMvtStock
    'Description : Génération des mouvement de Stock
    '=======================================================================
    Public Overrides Function genereMvtStock() As Boolean
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer)

        Dim bReturn As Boolean
        Dim objLgCom As LgCommande
        Dim objProduit As Produit
        Dim objmvtStock As mvtStock
        Dim bcolADecharger As Boolean

        'Chargement de la collection des lignes si elle n'est pas chargée
        bcolADecharger = False
        If Not m_bcolLgLoaded Then
            bReturn = loadcolLignes()
            bcolADecharger = True
            Debug.Assert(bReturn, getErreur())
        End If


        Try
            For Each objLgCom In m_colLignes
                objProduit = objLgCom.oProduit
                objProduit.load()
                objProduit.loadcolmvtStock()
                'Ajout de la ligne de stock avec recalcul du stock
                objmvtStock = objProduit.ajouteLigneMvtStock(Me.dateLivraison, vncEnums.vncTypeMvt.vncmvtBonAppro, id, "BA " & code & " BL:" & refLivraison, objLgCom.qteLiv * +1, "Bon Appro N° " & code & vbCrLf & "Fournisseur : " & oTiers.code & vbCrLf & "Date Commande" & dateCommande.ToShortDateString & vbCrLf & "Ref Livraison : " & refLivraison, True)
                objmvtStock.save()
                'Recalcul du stock produit
                objProduit.releasecolmvtStock()
                'Sauvegarde du produit
                objProduit.save()

            Next objLgCom
            etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncRien

            'Déchargement de la collection si elle n'était pas chargée à l'entrée
            If bcolADecharger Then
                releasecolLignes()
            End If

            bReturn = True

        Catch ex As Exception
            bReturn = False
            setError("genereMvtStock", ex.ToString)
        End Try

        Return bReturn
    End Function 'genereMvtStock


#End Region

    '=======================================================================
    'Function : SuppressionMvtStock
    'Description : Suppression des mouvements de Stock
    '=======================================================================
    Public Overrides Function supprimeMvtStock() As Boolean
        Debug.Assert(etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncSupprimer)

        Dim bReturn As Boolean
        Dim objLgCom As LgCommande
        Dim objProduit As Produit
        Dim objmvtStock As mvtStock

        Dim bcolADecharger As Boolean

        'Chargement de la collection des lignes si elle n'est pas chargée
        bcolADecharger = False
        If Not m_bcolLgLoaded Then
            bReturn = loadcolLignes()
            bcolADecharger = True
            Debug.Assert(bReturn, getErreur())
        End If


        Try
            'Pour chaque ligne de commande
            For Each objLgCom In m_colLignes
                objProduit = objLgCom.oProduit
                'Décrément de la qte en stock
                objmvtStock = New mvtStock(Me.dateLivraison, objProduit.id, vncEnums.vncTypeMvt.vncMvtCommandeClient, objLgCom.qteLiv * -1, "Mouvement créé juste pour recalculer le stock")
                objProduit.calculStock(objmvtStock)
                objProduit.releasecolmvtStock() 'On ne sauvegarde pas les mouvements de stocks
                objProduit.save()
            Next objLgCom

            'Suppression en bloc des lignes de Mvts de Stocks
            shared_connect()
            bReturn = deleteBA_MVTSTK()
            etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncRien
            shared_disconnect()

            'Déchargement de la collection si elle n'était pas chargée à l'entrée
            If bcolADecharger Then
                releasecolLignes()
            End If

        Catch ex As Exception
            bReturn = False
            setError("BonAppro.supprimeMvtStock", ex.ToString)
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn




    End Function 'supprimeMvtStock
    Public Function purge() As Boolean
        '================================================================================
        ' Fonction : purge 
        ' Description : Supression du Bon D'appro
        '               Le Bon d'appro est purgeable s'il est dans l'état Livré
        '================================================================================
        Dim bReturn As Boolean

        bReturn = True
        'Test sur l'état du Bon Appro = Livré
        bReturn = Me.etat.codeEtat = vncEnums.vncEtatCommande.vncBALivre
        If bReturn Then
            Me.bDeleted = True
            bReturn = Me.save()
        End If

        Return bReturn
    End Function 'purge

    Public Overrides Function controleMvtStock() As Collection

        '=====================================================================================================
        ' Function : controleMvtStock
        ' Description : parcours de chaque ligne de commande pour vérifier l'existence d'un mouvement de stock
        '=====================================================================================================
        Dim objLgCMD As LgCommande = Nothing
        Dim colReturn As New Collection
        Dim strErreurLg As String
        Dim nidProduitProduit As Integer
        Dim nLigneMemeProduit As Integer
        Dim colMvtStock As ColEventSorted = New ColEventSorted()
        Dim strErreur As String
        Dim objMvtStock As mvtStock
        Dim bTrouve As Boolean

        strErreur = "BA Num: " & code() & "(" & id & ")"

        nidProduitProduit = 0
        'chargement de la collection des lignes de produits
        'Elles sont triées par id produit
        If Not m_bcolLgLoaded Then
            loadcolLignes()
        End If
        For Each objLgCMD In colLignes
            strErreurLg = ""
            If nidProduitProduit <> objLgCMD.oProduit.id Then
                'changement de produit
                If nidProduitProduit <> 0 Then
                    '2) Controle des les mvts ne correspondant pas à une ligne de commande
                    For Each objMvtStock In colMvtStock
                        strErreurLg = ""
                        If Not objMvtStock.bControle Then
                            strErreurLg = "Produit Code " & objLgCMD.oProduit.code & "(" & objLgCMD.oProduit.id & ")" & " pas de lignes de Commandes pour le mouvement de stock " & objMvtStock.toString & "(" & objMvtStock.id & ")"
                            objMvtStock.bControle = True
                        End If
                        If strErreurLg <> "" Then
                            colReturn.Add(strErreur & strErreurLg)
                            strErreurLg = ""
                        End If
                    Next

                End If
                colMvtStock = mvtStock.getListe(objLgCMD.oProduit.id, , m_id)
                nidProduitProduit = objLgCMD.oProduit.id
            End If
            nLigneMemeProduit = nLigneMemeProduit + 1

            '1) Controle de l'existence d'un Mvt de stock pour chaque ligne
            ' Parcours de la collection des mvts de stocks pour en trouver un de la même quantité
            Debug.Assert(Not colMvtStock Is Nothing)
            bTrouve = False
            For Each objMvtStock In colMvtStock
                'Si une ligne à déjà été controlée il ne faut la prendre en compte une seconde fois
                'Probleme des doublons
                If Not objMvtStock.bControle Then
                    If objMvtStock.qte = (objLgCMD.qteLiv * +1) Then
                        bTrouve = True
                        objMvtStock.bControle = True
                    End If
                End If
                If bTrouve Then
                    'On a trouve une ligne de mvt de stock correspondant au produit et à la quantité 
                    'donc on sort
                    Exit For
                End If

            Next
            If Not bTrouve Then
                strErreurLg = "Produit Code " & objLgCMD.oProduit.code & "(" & objLgCMD.oProduit.id & ")" & " pas de lignes de mouvements de stocks trouvée pour la quantité " & objLgCMD.qteLiv
            End If


            If strErreurLg <> "" Then
                colReturn.Add(strErreur & strErreurLg)
                strErreurLg = ""
            End If
        Next
        If Not colMvtStock Is Nothing Then
            'Pour le dernier produit chargé
            '2) Controle des les mvts ne correspondant pas à une ligne de commande
            For Each objMvtStock In colMvtStock
                strErreurLg = ""
                If Not objMvtStock.bControle Then
                    strErreurLg = "Produit Code " & objLgCMD.oProduit.code & "(" & objLgCMD.oProduit.id & ")" & " pas de lignes de Commandes pour le mouvement de stock " & objMvtStock.toString & "(" & objMvtStock.id & ")"
                    objMvtStock.bControle = True
                End If
                If strErreurLg <> "" Then
                    colReturn.Add(strErreur & strErreurLg)
                    strErreurLg = ""
                End If
            Next
        End If


        'Suppression de la collection des lignes pour gagner de la place
        releasecolLignes()

        Debug.Assert(Not colReturn Is Nothing)
        Return colReturn

    End Function ' controleMvtStock

End Class ' BonAppro
