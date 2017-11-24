'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : EtatCommande
' Description : Etat générique de la commande
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
'===================================================================================================================================
Public MustInherit Class EtatCommande
    Inherits racine
    Protected m_codeEtat As vncEtatCommande
    Protected m_libelle As String
    Protected m_actionMvtStock As vncGenererSupprimer
    Protected m_actionSousCommande As vncGenererSupprimer

#Region "Accesseurs"

    Public ReadOnly Property libelle() As String
        Get
            Return m_libelle
        End Get
    End Property 'Libellé
    Public Property codeEtat() As vncEtatCommande
        Get
            Return m_codeEtat
        End Get

        Set(ByVal Value As vncEtatCommande)
            m_codeEtat = Value
        End Set
    End Property
    Public Property actionMvtStock() As vncGenererSupprimer
        Get
            Return m_actionMvtStock
        End Get

        Set(ByVal Value As vncGenererSupprimer)
            m_actionMvtStock = Value
        End Set
    End Property
    Public Property actionSousCommande() As vncGenererSupprimer
        Get
            Return m_actionSousCommande
        End Get

        Set(ByVal Value As vncGenererSupprimer)
            m_actionSousCommande = Value
        End Set
    End Property

#End Region
#Region "MéthodesDeClasse"
    Public Shared Function createEtat(ByVal pcodeEtat As vncEtatCommande, Optional ByVal pActionMvtStock As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien, Optional ByVal pActionSousCommande As vncGenererSupprimer = vncEnums.vncGenererSupprimer.vncRien) As EtatCommande
        Select Case pcodeEtat
            Case vncEnums.vncEtatCommande.vncRien
                Return New EtatCommandeRien(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncEnCoursSaisie
                Return New EtatCommandeEnCoursSaisie(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncValidee
                Return New EtatCommandeValidee(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncEclatee
                Return New EtatCommandeEclatee(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncTransmiseQuadra
                Return New EtatCommandeTransmiseQuadra(vncGenererSupprimer.vncRien, vncGenererSupprimer.vncRien)
            Case vncEnums.vncEtatCommande.vncRapprochee
                Return New EtatCommandeRapprochee(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncLivree
                Return New EtatCommandeLivree(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncSCMDGeneree
                Return New EtatSSCommandeGeneree(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncSCMDtransmiseFax
                Return New EtatSSCommandeTransmise(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncSCMDRapprochee
                Return New EtatSSCommandeRapprochee(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncSCMDRapprocheeInt
                Return New EtatSSCommandeRapprocheeInt(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncSCMDFacturee
                Return New EtatSSCommandeFacturee(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncSCMDExporteeInt
                Return New EtatSSCommandeExporteeInt(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncBAEnCours
                Return New EtatBAEnCoursSaisie(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncBALivre
                Return New EtatBALivrer(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncFactComGeneree
                Return New EtatFactCommGeneree(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncFactComExportee
                Return New EtatFactCommExportee(pActionMvtStock, pActionSousCommande)
            Case vncEnums.vncEtatCommande.vncFactTRPGeneree
                Return New EtatFactTRPGeneree
            Case vncEnums.vncEtatCommande.vncFactTRPExportee
                Return New EtatFactTRPExportee
            Case vncEnums.vncEtatCommande.vncFactCOLGeneree
                Return New EtatFactColGeneree
            Case vncEnums.vncEtatCommande.vncFactCOLExportee
                Return New EtatFactcolExportee
            Case vncEnums.vncEtatCommande.vncFactHBVGeneree
                Return New EtatFactHBVGeneree
            Case vncEnums.vncEtatCommande.vncFactHBVExportee
                Return New EtatFactHBVExportee
            Case Else
                Debug.Assert(False, "Pas de création d'état pour cet etat " & CStr(pcodeEtat))
                Return Nothing
        End Select
    End Function
#End Region
        '=======================================================================
        'Fonction : ToString()
        'Description : Rend l'objet sous forme de String
        'Détails    :  
        'Retour : une Chaine
        '=======================================================================
    Public Overrides Function toString() As String
        Return "ETA =(" & m_codeEtat & ")"
    End Function 'ToString
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        bReturn = obj.GetType.Name.Equals(Me.GetType().Name)
        Return bReturn
    End Function
    Protected Sub New(ByVal pactionMvtStock As vncGenererSupprimer, ByVal pactionSousCommande As vncGenererSupprimer)
        m_actionMvtStock = pactionMvtStock
        m_actionSousCommande = pactionSousCommande
    End Sub

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_codeEtat
        End Get
    End Property

    Protected MustOverride Function action(ByVal vncAction As vncActionEtatCommande) As vncEtatCommande

    Friend Overridable Function changeEtat(ByVal p_vncAction As vncActionEtatCommande) As EtatCommande
        Dim nEtat As vncEnums.vncActionEtatCommande
        nEtat = action(p_vncAction)
        Return createEtat(nEtat, m_actionMvtStock, m_actionSousCommande)
    End Function




End Class
