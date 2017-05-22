'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : EtatReglement
' Description : Etat générique d'un reglement
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
Public MustInherit Class EtatReglement
    Inherits racine
    Protected m_codeEtat As vncEtatReglement
    Protected m_libelle As String

#Region "Accesseurs"

    Public ReadOnly Property libelle() As String
        Get
            Return m_libelle
        End Get
    End Property 'Libellé
    Public Property codeEtat() As vncEtatReglement
        Get
            Return m_codeEtat
        End Get

        Set(ByVal Value As vncEtatReglement)
            m_codeEtat = Value
        End Set
    End Property

#End Region
#Region "MéthodesDeClasse"
    Public Shared Function createEtat(ByVal pcodeEtat As vncEtatReglement) As EtatReglement
        Select Case pcodeEtat
            Case vncEnums.vncEtatReglement.vncRGLMT_Saisi
                Return New EtatReglement_Saisi()
            Case vncEnums.vncEtatReglement.vncRGLMT_Export
                Return New EtatReglement_Exporte()
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
    Protected Sub New()
    End Sub

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_codeEtat
        End Get
    End Property

    Protected MustOverride Function action(ByVal vncAction As vncActionReglement) As vncEtatReglement

    Friend Overridable Function changeEtat(ByVal p_vncAction As vncActionReglement) As EtatReglement
        Dim nEtat As vncEnums.vncEtatReglement
        nEtat = action(p_vncAction)
        Return createEtat(nEtat)
    End Function
End Class
