Imports System.IO
'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Commande Client   
' Description : Commande Client
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
'15/09/05 : Ajout des factures de transport
'1/10/05 : Ajout des qteColis, Poids et Palettes
'
'===================================================================================================================================Public MustInherit Class Persist
Public Class CommandeClient
    Inherits Commande
#Region "Membres Privés"
    '=================== MEMBRES PRIVES ====================================
    Private m_typeCommande As vncTypeCommande       'Type de Commande
    Private m_typeTransport As vncTypeTransport      'Type de Transport
    Private m_colSousCommande As ColEvent                       'Collection de sous-commandes
    Private m_bcolSousCommandeLoaded As Boolean
    'Ajout Logistique
    Private m_bFactTransport As Boolean
    Private m_idFactTransport As Long
    'Ajout du nom et raison Sociale de livraison
    Private m_RaisonSocialeLivraison As String
    Private m_NomLivraison As String
    Protected m_bUpdatePrecommande As Boolean = True     'Mise à jour des précommandes (par defaut = vrai)
    Protected m_IDPrestashop As Long
    Protected m_NamePrestashop As String
    Protected m_Origine As String
    Protected m_idFHBV As Long 'Identifiant de la facture Hobivin
#End Region
#Region "Accesseurs"
    Public Sub New(ByVal poClient As Client)
        MyBase.New(poClient, New EtatCommandeEnCoursSaisie)
        m_typedonnee = vncEnums.vncTypeDonnee.COMMANDECLIENT
        m_typeCommande = vncEnums.vncTypeCommande.vncCmdClientPlateforme
        m_typeTransport = vncEnums.vncTypeTransport.vncTrpDepart
        m_colSousCommande = New ColEvent
        m_bcolSousCommandeLoaded = False
        m_bFactTransport = True
        m_idFactTransport = 0
        m_NomLivraison = ""
        m_RaisonSocialeLivraison = ""
        m_bUpdatePrecommande = True
        m_NamePrestashop = ""
        m_idFHBV = 0
        m_Origine = "VINICOM"
        Debug.Assert(Not m_oTransporteur Is Nothing)
        Debug.Assert(Not etat Is Nothing)
        majBooleenAlaFinDuNew()
    End Sub
    Public Shared Function createandload(ByVal pid As Integer) As CommandeClient
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim objCommandeClient As CommandeClient
        Dim bReturn As Boolean
        objCommandeClient = New CommandeClient(New Client("", ""))
        Try
            If pid <> 0 Then
                bReturn = objCommandeClient.load(pid)
                If Not bReturn Then
                    setError("CommandeClient.createAndLoad", getErreur())
                Else
                    objCommandeClient.oTiers.load(objCommandeClient.oTiers.id)
                End If

            End If
        Catch ex As Exception
            setError("CommandeClient.createAndLoad", ex.ToString)
        End Try
        Debug.Assert(objCommandeClient.id = pid, "Commande Client " & pid & " non chargée")
        Return objCommandeClient
    End Function 'createandload

    Public Property typeCommande() As vncTypeCommande
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_typeCommande
        End Get
        Set(ByVal Value As vncTypeCommande)
            If m_typeCommande <> Value Then
                m_typeCommande = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property typeTransport() As vncTypeTransport
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_typeTransport
        End Get
        Set(ByVal Value As vncTypeTransport)
            If m_typeTransport <> Value Then
                m_typeTransport = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public ReadOnly Property colSousCommandes() As ColEvent
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_colSousCommande
        End Get
    End Property

    Public Property bFactTransport() As Boolean
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_bFactTransport
        End Get
        Set(ByVal Value As Boolean)
            If (Value <> m_bFactTransport) Then
                m_bFactTransport = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property bUpdatePrecommande() As Boolean
        Get
            Return m_bUpdatePrecommande
        End Get
        Set(ByVal Value As Boolean)
            m_bUpdatePrecommande = Value
        End Set
    End Property
    Public Property idFactTransport() As Long

        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_idFactTransport
        End Get
        Set(ByVal Value As Long)
            If (Value <> m_idFactTransport) Then
                m_idFactTransport = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Raison Sociale de Livraison (#702)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RaisonSocialeLivraison As String
        Get
            Return m_RaisonSocialeLivraison
        End Get
        Set(value As String)
            If value <> m_RaisonSocialeLivraison Then
                m_RaisonSocialeLivraison = value
                RaiseUpdated()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Nom de la personne pour la livraison (#702)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property NomLivraison As String
        Get
            Return m_NomLivraison
        End Get
        Set(value As String)
            If value <> m_RaisonSocialeLivraison Then
                m_NomLivraison = value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property IDPrestashop As Long
        Get
            Return m_IDPrestashop
        End Get
        Set(value As Long)
            If value <> m_IDPrestashop Then
                m_IDPrestashop = value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property NamePrestashop As String
        Get
            Return m_NamePrestashop
        End Get
        Set(value As String)
            If value <> m_NamePrestashop Then
                m_NamePrestashop = value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property Origine As String
        Get
            Return m_Origine
        End Get
        Set(value As String)
            If value <> m_Origine Then
                m_Origine = value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property idFactHobivin As Long
        Get
            Return m_idFHBV
        End Get
        Set(value As Long)
            If value <> m_idFHBV Then
                m_idFHBV = value
                RaiseUpdated()
            End If
        End Set
    End Property
#End Region
#Region "Méthodes de Classe"
    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListe 
    'Description : Liste des Clients
    'Retour : Rend une collection de Clients
    '=======================================================================
    Public Shared Function getListe(Optional ByVal strCode As String = "", Optional ByVal strNomClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien, Optional pOrigine As String = "VINICOM") As Collection
        Dim colReturn As Collection

        Persist.shared_connect()
        colReturn = ListeCMDCLT(strCode, strNomClient, pEtat, pOrigine)
        Persist.shared_disconnect()
        Return colReturn
    End Function
    Public Shared Function getListe(ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pNomClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien, Optional pOrigine As String = "VINICOM") As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeCMDCLTEtat(pddeb, pdfin, pNomClient, pEtat, pOrigine)
        Debug.Assert(Not colReturn Is Nothing, "CommandeClient.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

    End Function

    Public Shared Function getListe_TRP(ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pNomClient As String = "", Optional ByVal pCodeClient As String = "") As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeCMDCLT_TRP(pddeb, pdfin, pNomClient, pCodeClient)
        Debug.Assert(Not colReturn Is Nothing, "FactCom.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

    End Function

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
        bReturn = LoadCMDCLT()
        If bReturn Then
            'Chargement des caractéristiques du client
            bReturn = oTiers.loadLight()
            'Debug.Assert(bReturn, Tiers.getErreur())
        End If
        If Commande.bChargerColLignes Then
            bReturn = loadcolLignes()
        End If
        m_colSousCommande.clear()
        m_bcolLgLoaded = bReturn
        Return bReturn
    End Function 'DBLoad
    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = MyBase.Save()
        If m_bcolLgLoaded Then
            bReturn = bReturn And savecolLignes()
        End If
        If m_id > 0 Then
            If m_bcolLgLoaded And bUpdatePrecommande Then
                bReturn = bReturn And updatePrecommande()
                Debug.Assert(bReturn, "Erreur en updatePrecommande" & getErreur())
            End If
            'MVTS STOCKS
            If getActionMvtStock() = vncEnums.vncGenererSupprimer.vncGenerer Then
                bReturn = bReturn And genereMvtStock()
                Debug.Assert(bReturn, "Erreur en generemvtStock" & getErreur())
            End If
            If getActionMvtStock() = vncEnums.vncGenererSupprimer.vncSupprimer Then
                bReturn = bReturn And supprimeMvtStock()
                Debug.Assert(bReturn, "Erreur en supprimeMvtStock" & getErreur())
            End If
            'Sous-Commandes
            Select Case getActionSousCommande()
                Case vncGenererSupprimer.vncGenerer
                    bReturn = bReturn And insertcolSCMD()
                    Debug.Assert(bReturn, "Erreur en insertcolSCMD" & getErreur())
                Case vncGenererSupprimer.vncSupprimer
                    bReturn = bReturn And deletecolSCMD()
                    Debug.Assert(bReturn, "Erreur en deletecolSCMD" & getErreur())
                Case vncGenererSupprimer.vncRien
                    bReturn = bReturn And updatecolSCMD()
                    Debug.Assert(bReturn, "Erreur en updatecolSCMD" & getErreur())
            End Select
        End If
        shared_disconnect()
        Return bReturn
    End Function ' Save
    Private Function updatePrecommande() As Boolean
        '=================================================================================
        'Fonction : updatePrecommande
        'Description : Mise à jour de la precommande du client avec le prix de la commande
        'Détails    :  
        'Retour : Vrai si l'opération s'est bien déroulée
        '================================================================================
        Dim bReturn As Boolean
        Dim objClient As Client
        Dim oSW As New Stopwatch()

        Try
            objClient = oTiers
            oSW.Start()
            objClient.updatePrecommande(Me)
            oSW.Stop()
            oSW.Reset()
            oSW.Start()
            bReturn = objClient.save
            oSW.Stop()

        Catch ex As Exception
            bReturn = False
            setError("CommandeClient.updatePrecommande", ex.ToString)
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'updatePrecommande
    Friend Overrides Function delete() As Boolean
        '=======================================================================
        'Fonction : delete()
        'Description : Suppression de l'objet dans la base de l'objet
        'Détails    :  
        'Retour : Vrai si l'opération s'est bien déroulée
        '=======================================================================
        Debug.Assert(id <> 0, "idCommande <> 0")
        Dim bReturn As Boolean
        'On ne supprime plus automtiquement les mouvements de stocks
        '   La commande peut-être supprimée sans modifier le stock
        '        etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncSupprimer 'Il faut supprimer les mvts de stocks
        '       supprimeMvtStock()
        bReturn = deletecolLgCMD()
        If bReturn Then
            m_bcolLgLoaded = False
            m_colLignes.clear()
            bReturn = deletecolSCMD()
            'Les sous commandes sont Supprimées par les factures de commissions
            If bReturn Then
                bReturn = deleteCMDCLT()
            End If
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
            If existeSousCommandeCommande() Then
                bReturn = False
            End If
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
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas Renseigné")
        Debug.Assert(code.Equals(""), "Le Code doit être nul")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
        Debug.Assert(id = 0, "idCommande = 0")

        Dim bReturn As Boolean
        If setNewcode() Then
            bReturn = insertCMDCLT()
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
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas Renseigné")
        Debug.Assert(code <> "", "Le Code ne doit pas être nul")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
        Debug.Assert(id <> 0, "idCommande <> 0")
        Dim bReturn As Boolean
        bReturn = updateCMDCLT()
        Return bReturn

    End Function 'Update

    '======================================================================
    'Fonction : LoadcolSousCommande
    'Description : Chargement de la collection des sous-commandes
    '======================================================================
    Public Function LoadColSousCommande() As Boolean
        Debug.Assert(m_id <> 0, "La commande doit être sauvegardée au Préalable")
        Dim bReturn As Boolean
        bReturn = False
        shared_connect()
        If m_bcolLgLoaded Then
            m_colSousCommande.clear()
        End If
        bReturn = LoadColSCMD()
        Debug.Assert(bReturn, racine.getErreur())
        m_bcolSousCommandeLoaded = bReturn ' Les Sous commandes sont chargées
        shared_disconnect()
        Return bReturn
    End Function 'LoadColSousCommande
#End Region
#Region "Méthodes Overrides"
    Public Overrides Sub Exporter(ByVal pfileNAme As String)
        ' no export for Commandeclient
    End Sub

    Friend Overrides Function setNewcode() As Boolean
        Dim str As String
        Dim ncode As Integer
        Dim breturn As Boolean

        shared_connect()
        str = ""
        ncode = getNumeroCommandeClient()
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
        '        etat = etat.changeEtat(p_action)
        '4/10/04 Les Commandes Directe ne générent pas de mvts de stocks

        Debug.Assert(p_action >= vncActionEtatCommande.vncActionMinCommande And p_action <= vncActionEtatCommande.vncActionMaxCommande)
        Dim bReturn As Boolean = False
        Try
            If p_action >= vncActionEtatCommande.vncActionMinCommande And p_action <= vncActionEtatCommande.vncActionMaxCommande Then
                MyBase.changeEtat(p_action)
                If typeCommande = vncEnums.vncTypeCommande.vncCmdClientDirecte Then
                    etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncRien
                End If
                RaiseUpdated()
                bReturn = True
            Else
                setError("Commandeclient.changeEtat", "Code Action incorrect :" + p_action)
                bReturn = False
            End If

        Catch ex As Exception
            setError("CommandeClient.changeEtat", ex.ToString)
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
        Dim objCommande As CommandeClient
        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj
            bReturn = bReturn And (typeCommande.Equals(objCommande.typeCommande))
            bReturn = bReturn And (typeTransport.Equals(objCommande.typeTransport))
            bReturn = bReturn And (m_bFactTransport.Equals(objCommande.bFactTransport))
            bReturn = bReturn And (m_idFactTransport.Equals(objCommande.idFactTransport))
            bReturn = bReturn And (m_NomLivraison.Equals(objCommande.NomLivraison))
            bReturn = bReturn And (m_RaisonSocialeLivraison.Equals(objCommande.RaisonSocialeLivraison))

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn

    End Function 'Equals
#End Region
    ''Methode : EstEntierementLivree
    ''Description : rend Vrai si tous les lignes on été Livrées
    'Public Function estEntierementLivree() As Boolean
    '    Dim olgCom As LgCommande
    '    Dim bReturn As Boolean

    '    bReturn = True
    '    For Each olgCom In colLignes
    '        bReturn = bReturn And olgCom.estLivree()
    '    Next

    '    Return bReturn

    'End Function 'estEntierementLivree

    ''=======================================================================
    ''Fonction : AjouteLigne()
    ''Description : Ajoute une ligne sur une commandeClient
    ''Détails    :  
    ''Retour : une ligne de commande ou nothing si l'ajout échoue
    ''=======================================================================
    ''    Public Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrixCommande As Boolean = True) As LgCommande
    'Public Overloads Overrides Function AjouteLigne(ByVal pobjLgCMD As LgCommande, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
    '    Debug.Assert(Not m_colLignes Is Nothing)
    '    Debug.Assert(Not pobjLgCMD Is Nothing)
    '    Dim oReturn As LgCommande

    '    Try
    '        If p_bCalculPrix Then
    '            pobjLgCMD.calculPrixTotal()
    '        End If
    '        m_colLignes.Add(pobjLgCMD, CStr(pobjLgCMD.num))
    '        If p_bCalculPrix Then
    '            calculPrixTotal()
    '        End If
    '        oReturn = pobjLgCMD
    '        m_bcolLgLoaded = True
    '    Catch ex As Exception
    '        setError("Commande.AjouteLigne", "Ajout de ligne impossible key = " & pobjLgCMD.num)
    '        oReturn = Nothing
    '    End Try
    '    '        Debug.Assert(m_ocolLignes.Count = n + 1, "Le nombre d'élement dans la collection est incrémenté de 1")
    '    Return oReturn
    'End Function 'AjouteLigne

    ''=======================================================================
    ''Fonction : AjouteLigne()
    ''Description : Créé une ligne de commande et l'ajoute à la collection via AjouteLigne
    ''Détails    :   Appelle la Fonction AjoutLigne a
    ''Retour : une ligne de commande ou nothing si l'ajout échoue
    ''=======================================================================
    'Public Overloads Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
    '    Debug.Assert(Not m_colLignes Is Nothing)
    '    Debug.Assert(Not p_oProduit Is Nothing)

    '    Dim oLgCmd As LgCommande

    '    oLgCmd = New LgCommande(m_id)
    '    oLgCmd.num = p_strNum
    '    oLgCmd.idCmd = id
    '    oLgCmd.idSCmd = 0 ' La sous commande n'est pas connue
    '    oLgCmd.oProduit = p_oProduit
    '    oLgCmd.qteCommande = p_qteCmd
    '    oLgCmd.prixU = p_prixU
    '    oLgCmd.bGratuit = p_bGratuit
    '    oLgCmd.prixHT = p_prixHT
    '    oLgCmd.prixTTC = p_prixTTC

    '    oLgCmd = AjouteLigne(oLgCmd, p_bCalculPrix)
    '    Return oLgCmd
    'End Function 'AjouteLigne

    Public Function generationSousCommande(Optional pIntermediaire As Client = Nothing) As Boolean
        '=======================================================================
        'Fonction : generationSousCommande
        'Description : Création des sousCommandes : 1 Souscommande = 1 Fournisseur 
        '=======================================================================
        Debug.Assert(Not m_colLignes Is Nothing, "m_colLignes is Nothing")
        Debug.Assert(Not m_colSousCommande Is Nothing, "m_colSouscommande is Nothing")
        Debug.Assert(m_colSousCommande.Count() = 0, "m_colSouscommande non vide")

        Dim bReturn As Boolean
        Dim objLGCMD As LgCommande
        Dim bFini As Boolean
        Dim idFRN As Integer
        Dim oFRN As Fournisseur
        Dim oSCMD As SousCommande


        oSCMD = Nothing
        bReturn = False
        bFini = False
        'Reinitialisation des lignes 
        For Each objLGCMD In m_colLignes
            objLGCMD.idSCmd = 0
            objLGCMD.bLigneEclatee = False
        Next
        'Tant que l'on a pas terminer on boucle sur la collection de lignes
        While Not bFini
            bFini = True
            idFRN = 0 'on reinitialise le fournisseur courrant avant de commencer la boucle
            For Each objLGCMD In m_colLignes
                'on ne traitre que les lignes qui ne sont pas éclatées
                If Not objLGCMD.bLigneEclatee Then
                    bFini = False
                    If idFRN = 0 Then
                        'S'il n'y a pas de fournisseur courant on prend celui là
                        idFRN = objLGCMD.oProduit.idFournisseur
                        oFRN = Fournisseur.createandload(idFRN)
                        ' et on crée la sous-commande avec ce fournisseur
                        oSCMD = New SousCommande(Me, oFRN)
                        oSCMD.setNewcode()
                        If pIntermediaire IsNot Nothing Then
                            oSCMD.oTiers = pIntermediaire
                        End If
                        'on ajoute la sous-commande à la collection
                        Me.colSousCommandes.Add(oSCMD, oSCMD.code)
                    End If
                    If objLGCMD.oProduit.idFournisseur = idFRN Then
                        'La Ligne correspond au fournisseur courant
                        ' On l'ajoute à la sous-commande courante 
                        oSCMD.AjouteLigne(objLGCMD, True)
                        'On la marque comme traitée
                        objLGCMD.bLigneEclatee = True
                        objLGCMD.idSCmd = oSCMD.id ' 0 car la sous commande n'est pas encore sauvegardée
                    End If
                End If
            Next
        End While
        'Calcul de la commission pour chaque Sous commande
        For Each oSCMD In m_colSousCommande
            oSCMD.calcCommisionstandard(CalculCommScmd.CALCUL_COMMISSION_HT_CALCULE)
        Next


        bReturn = bFini
        If bReturn Then
            changeEtat(vncEnums.vncActionEtatCommande.vncActionEclater)
        End If
        Return bReturn
    End Function 'generationSousCommande
    ''' <summary>
    ''' Annulation des sousCommandes 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function annulationSousCommande() As Boolean
        '=======================================================================
        'Fonction : generationSousCommande
        'Description : Création des sousCommandes : 1 Souscommande = 1 Fournisseur 
        '=======================================================================
        Debug.Assert(Not m_colLignes Is Nothing, "m_colLignes is Nothing")
        Debug.Assert(Not m_colSousCommande Is Nothing, "m_colSouscommande is Nothing")
        Debug.Assert(m_bcolSousCommandeLoaded, "Collection des sous-commandes non chargée")

        Dim bReturn As Boolean
        Dim objLGCMD As LgCommande

        bReturn = False
        'Nettoyage de la collection des Sous-commandes
        colSousCommandes.clear()
        'Reinitialisation des lignes 
        For Each objLGCMD In m_colLignes
            objLGCMD.bLigneEclatee = False
            objLGCMD.idSCmd = 0
        Next
        bReturn = True
        'changement D'état de la commande
        If bReturn Then
            changeEtat(vncEnums.vncActionEtatCommande.vncActionAnnEclater)
        End If
        Return bReturn
    End Function 'annulationSousCommande
    ' <summary>mise à jour de la colection des sous-commandes</summary>
    Private Function updatecolSCMD() As Boolean
        '=======================================================================
        'Fonction : updatecolSCMD
        'Description : Mise à jour de la collection des sous-comandes
        '               Boucle sur les sous-commandes et sauvegarde de chacune
        '=======================================================================
        Dim bReturn As Boolean
        Dim objScmd As SousCommande
        Try
            For Each objScmd In colSousCommandes
                bReturn = objScmd.Save()
                Debug.Assert(bReturn, "Sauvegarde de Sous Commandes")
            Next objScmd
            'Le Travail sur les sous-commandes à été fait
            etat.actionSousCommande = vncEnums.vncGenererSupprimer.vncRien
            bReturn = True
        Catch ex As Exception
            bReturn = False
            setError("SaveSousCommande", ex.ToString)
        End Try
        Return bReturn
    End Function 'saveSousCommandes
    Private Function insertcolSCMD() As Boolean
        '=======================================================================
        'Fonction : insertcolSCMD
        'Description : insertion  des sous-commandes en base de données
        '=======================================================================
        Dim bReturn As Boolean
        Dim objScmd As SousCommande
        Try
            For Each objScmd In colSousCommandes
                objScmd.bNew = True
                bReturn = objScmd.Save()
                Debug.Assert(bReturn, "insertcolSCMD" & SousCommande.getErreur())
            Next objScmd
            'Le Travail sur les sous-commandes à été fait
            etat.actionSousCommande = vncEnums.vncGenererSupprimer.vncRien
            bReturn = True
        Catch ex As Exception
            bReturn = False
            setError("insertcolSCMD", ex.ToString)
        End Try
        Return bReturn
    End Function 'insertSousCommandes
    Public Function deletecolSCMD() As Boolean
        '=======================================================================
        'Fonction : deletecolSCMD
        'Description : Suppression en base des sous-commandes
        '               Passage de chaque Sous-commande à bdeleted = True
        '=======================================================================
        Dim bReturn As Boolean
        Try
            'Il n'y a plus de sous commande dans la collection
            'Donc la suppression est plus bestiale : Suppression des tous les enregistrements Sous-commandes reliés à une commande
            'Seules les sousCommandes n'ayant pas été affectée à une facture de comm sont supprimées
            bReturn = deleteCMDCLT_SCMD()
            Debug.Assert(bReturn, "deleteCMDCLT_SCMD" & getErreur())

            'Le Travail sur les sous-commandes à été fait
            etat.actionSousCommande = vncEnums.vncGenererSupprimer.vncRien
            m_bcolSousCommandeLoaded = False
            m_colSousCommande.clear()
            bReturn = True
        Catch ex As Exception
            bReturn = False
            setError("deletecolSCMD", ex.ToString)
        End Try
        Return bReturn
    End Function 'deletecolSCMD
    '=======================================================================
    'Function : GenereMvtStock
    'Description : Génération des mouvement de Stock
    '                   Uniquement pour les commande Plateforme !!!
    '=======================================================================
    Public Overrides Function genereMvtStock() As Boolean
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer)


        Dim bReturn As Boolean
        Dim objLgCom As LgCommande
        Dim objProduit As Produit
        Dim objmvtStock As mvtStock
        Dim bcolADecharger As Boolean
        Dim strLib As String

        'Chargement de la collection des lignes si elle n'est pas chargée
        bcolADecharger = False
        If Not m_bcolLgLoaded Then
            bReturn = loadcolLignes()
            bcolADecharger = True
            Debug.Assert(bReturn, getErreur())
        End If

        Try
            If m_typeCommande = vncEnums.vncTypeCommande.vncCmdClientPlateforme Then
                'Pour chaque ligne de commande
                For Each objLgCom In m_colLignes
                    objProduit = objLgCom.oProduit
                    objProduit.load()
                    objProduit.loadcolmvtStock()
                    strLib = "CMD " & code & " - " & objLgCom.num & " " & oTiers.rs
                    'Controle de la non-exitence d'une ligne de mvt de stock pour cette ligne de commande
                    If existeMvtSocklib(strLib) Then
                        Debug.Assert(False, "Il y a déja un mvt de stock pour cette ligne")
                        bReturn = False
                        setError("CommandeClient.genereMvtStock()", "Il y a déja un mvt de stock pour cette ligne")
                        Return bReturn
                        Persist.shared_connect()
                        Exit Function
                    End If
                    'Ajout de la ligne de stock avec recalcul du stock
                    objmvtStock = objProduit.ajouteLigneMvtStock(Me.dateLivraison, vncEnums.vncTypeMvt.vncMvtCommandeClient, id, strLib, objLgCom.qteLiv * -1, "Commande N° " & code & vbCrLf & " Client : " & oTiers.code & " " & oTiers.rs & vbCrLf & "Date Commande" & dateCommande.ToShortDateString & vbCrLf & "Ref Livraison : " & refLivraison, True)
                    objmvtStock.save()
                    'Liberation des mouvement de stock produits pour libérer la mémoire
                    objProduit.releasecolmvtStock()
                    'Sauvegarde du produit car le stock a été recalculé
                    objProduit.save()
                Next objLgCom
                etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncRien
            End If

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

    '=======================================================================
    'Function : SuppressionMvtStock
    'Description : Suppression des mouvements de Stock
    '=======================================================================
    Public Overrides Function supprimeMvtStock() As Boolean
        Debug.Assert(etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncSupprimer)

        Dim bReturn As Boolean
        Dim objLgCom As LgCommande
        Dim objProduit As Produit
        Dim bcolADecharger As Boolean

        'Chargement de la collection des lignes si elle n'est pas chargée
        bcolADecharger = False
        If Not m_bcolLgLoaded Then
            bReturn = loadcolLignes()
            bcolADecharger = True
            Debug.Assert(bReturn, getErreur())
        End If

        Try
            shared_connect()
            'Recalcul du stock de la commande
            If m_typeCommande = vncEnums.vncTypeCommande.vncCmdClientPlateforme Then
                'Suppression en bloc des lignes de Mvts de Stocks
                bReturn = deleteCMDCLT_MVTSTK()
                Debug.Assert(bReturn, "CommandeClient.supprimeMvtStock" & getErreur())
                'Pour chaque ligne de commande
                For Each objLgCom In m_colLignes
                    objProduit = objLgCom.oProduit
                    objProduit.releasecolmvtStock()
                    objProduit.loadcolmvtStock()
                    objProduit.recalculStock()
                    objProduit.releasecolmvtStock()
                    objProduit.save()
                Next objLgCom
                etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncRien
                'Déchargement de la collection si elle n'était pas chargée à l'entrée
                If bcolADecharger Then
                    releasecolLignes()
                End If
            End If
            shared_disconnect()

        Catch ex As Exception
            bReturn = False
            setError("Commandeclient.supprimeMvtStock", ex.ToString)
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'supprimeMvtStock

    Public Overrides Function calculPrixTotal() As Boolean
        '========================================================================
        ' CalculPrixTotal : Calcul du prix de lacommande, du poids et des qte de colis
        '       Utilise le fonction CalculPtixTotal de la classe Commande pour calculer les prixs
        '========================================================================
        Dim bReturn As Boolean
        Try
            bReturn = MyBase.calculPrixTotal()
            If bReturn Then
                bReturn = CalcPoidsColis()
            End If
        Catch ex As Exception
            setError("CommandeClient.CalculPrixTotal", ex.ToString)
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'calculPrixTotal


    Public Function purge() As Boolean
        '================================================================================
        ' Fonction : purge 
        ' Description : Supression de la Commmande Client
        '           La Commande est purgeable si tous ses sous-commandes ont été facturées
        '================================================================================
        Dim bReturn As Boolean
        Dim objSCMD As SousCommande
        Dim bPurgeable As Boolean

        bReturn = True
        'Test sur l'état de la commande = Eclatée
        bReturn = Me.etat.codeEtat = vncEnums.vncEtatCommande.vncEclatee
        If bReturn Then
            'Chargement des lignes de factures de comm (Sous-Commandes)
            If Not m_bcolSousCommandeLoaded Then
                bReturn = LoadColSousCommande()
            End If
            'Test si la commande est 'Purgeable'
            'CAD toutes les sous commandes doivent être facturées
            bPurgeable = True
            For Each objSCMD In m_colSousCommande
                bPurgeable = bPurgeable And ( _
                        objSCMD.etat.codeEtat = vncEnums.vncEtatCommande.vncSCMDFacturee)
            Next objSCMD

            'Si elle est purgeable => Suppression de la commande et de ses lignes de commandes
            If bPurgeable Then
                'Pour chaque ligne de commande création d'une ligne de statistiques
                'For Each objLgCmd In m_colLignes
                'objlgStat = New LgStat(Me, objLgCmd)
                'objlgStat.Save()
                'Next
                Me.bDeleted = True
                bReturn = Me.save()
            End If
            bReturn = True
        End If

        Return bReturn
    End Function 'purge

    Public Overrides Function controleMvtStock() As Collection

        '=====================================================================================================
        ' Function : controleMvtStock
        ' Description : parcours de chaque ligne de commande pour vérifier l'existence d'un mouvement de stock
        '=====================================================================================================
        Dim objLgCMD As LgCommande
        Dim colReturn As New Collection
        Dim strErreurLg As String
        Dim nidProduitProduit As Integer
        Dim nLigneMemeProduit As Integer
        Dim colMvtStock As ColEventSorted
        Dim strErreur As String
        Dim objMvtStock As mvtStock
        Dim bTrouve As Boolean
        Dim strCodePrd As String
        Dim nidPrd As Long

        colMvtStock = Nothing
        strCodePrd = ""
        objLgCMD = Nothing
        If typeCommande = vncEnums.vncTypeCommande.vncCmdClientPlateforme Then
            'On ne controle que les commande plateforme

            strErreur = "Commande Num: " & code() & "(" & id & ")"

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
                                strErreurLg = "Produit Code " & strCodePrd & "(" & nidPrd & ")" & " pas de lignes de Commandes pour le mouvement de stock " & objMvtStock.toString & "(" & objMvtStock.id & ")"
                                objMvtStock.bControle = True
                            End If
                            If strErreurLg <> "" Then
                                colReturn.Add(strErreur & strErreurLg)
                                strErreurLg = ""
                            End If
                        Next

                    End If
                    'Chargement des Mvt de Stok pour ce produit et cette commande
                    colMvtStock = mvtStock.getListe(objLgCMD.oProduit.id, m_id)
                    strCodePrd = objLgCMD.oProduit.code
                    nidPrd = objLgCMD.oProduit.id

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
                        If objMvtStock.qte = (objLgCMD.qteLiv * -1) Then
                            bTrouve = True
                            objMvtStock.bControle = True
                            Exit For
                        End If
                    End If
                Next
                If Not bTrouve Then
                    strErreurLg = "Produit Code " & objLgCMD.oProduit.code & "(" & objLgCMD.oProduit.id & ")" & " pas de lignes de mouvements de stocks trouvée pour la quantité " & objLgCMD.qteLiv
                End If


                If strErreurLg <> "" Then
                    colReturn.Add(strErreur & strErreurLg)
                    strErreurLg = ""
                End If
            Next objLgCMD
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

        End If 'Commnande.type = plateforme
        Debug.Assert(Not colReturn Is Nothing)
        Return colReturn

    End Function ' controleMvtStock

    Public ReadOnly Property shortResumeTRP() As String
        Get
            Return code & "|" & oTiers.code & "|" & oTiers.rs & "|" & dateLivraison & "|" & Format(montantTransport, "C")
        End Get
    End Property


    '========================================================================
    'fonction : exporter
    'DEscription : Exporte la facture de transport dans un format WEBEDI
    'Retour : une chaine
    '=========================================================================
    Public Function exporterWebEDI(ByVal strFileName As String) As Boolean

        Debug.Assert(bcolLignesLoaded, "Les lignes doivent être chargées")

        'Dim nvcEntete As System.Collections.Specialized.NameValueCollection
        'Dim nvcLigne As System.Collections.Specialized.NameValueCollection
        Dim bReturn As Boolean
        Dim nFile As Integer
        Dim oLg As LgCommande
        Dim strResult As String
        Try
            nFile = FreeFile()
            FileOpen(nFile, strFileName, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
            For Each oLg In m_colLignes
                '1
                strResult = Right("00000000" + Trim(Me.code), 8)
                '9
                strResult = strResult + Format(Now(), "ddMMyyyy")
                '17
                strResult = strResult + Format(Me.dateEnlevement, "ddMMyyyy")
                '25
                strResult = strResult + Format(Me.dateLivraison, "ddMMyyyy")
                '33
                strResult = strResult + Left(Me.oTiers.code + Space(8), 8)
                '41
                strResult = strResult + Left(Me.NomLivraison + Space(30), 30)
                '71
                strResult = strResult + Left(Me.RaisonSocialeLivraison + Space(30), 30)
                '101
                strResult = strResult + Left(Me.caracteristiqueTiers.AdresseLivraison.rue1 + Space(30), 30)
                '131
                strResult = strResult + Left(Me.caracteristiqueTiers.AdresseLivraison.rue2 + Space(30), 30)
                '161
                strResult = strResult + Left(Me.caracteristiqueTiers.AdresseLivraison.cp + Space(5), 5)
                '166
                strResult = strResult + Left(Me.caracteristiqueTiers.AdresseLivraison.ville + Space(26), 26)
                '192
                strResult = strResult + Left(Me.CommLivraison.comment + Space(100), 100)
                '292
                strResult = strResult + Format(oLg.num, "000")
                '295
                strResult = strResult + Left(oLg.oProduit.code + Space(15), 15)
                '310
                strResult = strResult + Format(oLg.qteColis, "0000000")
                '317
                strResult = strResult + Format(oLg.qteCommande, "0000000")
                '324
                strResult = strResult + Left(Me.oTransporteur.nom + Space(50), 50)
                '374
                strResult = Replace(strResult, vbCrLf, "--")
                strResult = Replace(strResult, vbCr, "-")
                strResult = Replace(strResult, vbLf, "-")
                strResult = Replace(strResult, vbNullChar, "-")
                strResult = Replace(strResult, vbTab, "-")
                strResult = Replace(strResult, vbBack, "-")
                PrintLine(nFile, strResult)
            Next oLg
            FileClose(nFile)


            bReturn = True
        Catch ex As Exception
            bReturn = False
            setError("CommandeClient.exporter", ex.ToString())
        End Try
        Return bReturn
    End Function 'exporterWEBEDI

    'Exportation des sous commandes d'ue commande
    'Public Function faxerTout(ByVal pFAX_Path As String, ByVal pPathToReports As String, ByVal PFax_Nom_Interlocuteur As String, ByVal pFAX_Tel_Interlocuteur As String, ByVal pFAX_Subject As String, ByVal pFAX_Notes As String, ByVal pFAX_BSENDCOVERPAGE As Boolean) As Boolean
    '    Dim objSCMD As SousCommande
    '    Dim strFileName As String
    '    Dim bToutOK As Boolean
    '    For Each objSCMD In colSousCommandes
    '        objSCMD.oFournisseur.load()
    '        'on ne faxe pas automatiquement pour les fourisseurs internet
    '        If Not objSCMD.oFournisseur.bExportInternet Then
    '            If objSCMD.oFournisseur.AdresseFacturation.fax <> "" Then
    '                strFileName = pFAX_Path & objSCMD.code & ".doc"
    '                If objSCMD.faxer(pPathToReports, PFax_Nom_Interlocuteur, pFAX_Tel_Interlocuteur, pFAX_Subject, pFAX_Subject, pFAX_BSENDCOVERPAGE, strFileName, objSCMD.oFournisseur.AdresseFacturation.fax) Then
    '                    objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
    '                Else
    '                    bToutOK = False  ' Il y a au moins une erreur
    '                End If

    '                If File.Exists(strFileName) Then
    '                    File.Delete(strFileName)
    '                End If
    '            Else
    '                bToutOK = False  ' Il y a au moins une erreur
    '            End If
    '        End If
    '    Next objSCMD

    '    Return bToutOK

    'End Function 'faxerTout

    Public Overrides Function DuppliqueCaracteristiqueTiers() As Boolean
        Debug.Assert(Not oTiers Is Nothing)
        MyBase.DuppliqueCaracteristiqueTiers()
        NomLivraison = oTiers.AdresseLivraison.nom
        RaisonSocialeLivraison = oTiers.rs
        Return True
    End Function 'DuppliqueCaracteristiqueTiers
    ''' <summary>
    ''' Controle du stock disponible pour chaque ligne de commande
    ''' met à jour l'indicateur bStockDispo sur la ligne de commande
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function controleStockDispo() As Boolean
        Dim bReturn As Boolean
        Try
            For Each olg As LgCommande In colLignes
                olg.ControleStockdispo()
            Next
            bReturn = True
        Catch ex As Exception
            setError("CommandeClient.ControleStockdispo ERR : " & ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'controleStockDispo

End Class ' CommandeClient
