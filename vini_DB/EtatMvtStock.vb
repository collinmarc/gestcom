'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Cr�ation : 23/06/04
'===================================================================================================================================
' Classe : EtatCommande
' Description : Etat g�n�rique de la commande
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
'   Equals()        : Test l'�quivalence d'instance
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================
Public MustInherit Class EtatMvtStock
    Inherits racine
    Protected m_codeEtat As vncEtatMVTSTK
    Protected m_libelle As String

#Region "Accesseurs"

    Public ReadOnly Property libelle() As String
        Get
            Return m_libelle
        End Get
    End Property 'Libell�
    Public Property codeEtat() As vncEtatMVTSTK
        Get
            Return m_codeEtat
        End Get

        Set(ByVal Value As vncEtatMVTSTK)
            m_codeEtat = Value
        End Set
    End Property

#End Region
#Region "M�thodesDeClasse"
    Public Shared Function createEtat(ByVal pcodeEtat As vncEtatMVTSTK) As EtatMvtStock
        Select Case pcodeEtat
            Case vncEnums.vncEtatMVTSTK.vncMVTSTK_nFact
                Return New EtatMvtStock_nFact()
            Case vncEnums.vncEtatMVTSTK.vncMVTSTK_Fact
                Return New EtatMvtStock_Fact()
            Case Else
                Debug.Assert(False, "Pas de cr�ation d'�tat pour cet etat " & CStr(pcodeEtat))
                Return Nothing
        End Select
    End Function
#End Region
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'D�tails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "ETA =(" & m_codeEtat & ")"
    End Function 'ToString
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'�quivalence entre 2 objets
    'D�tails    :  
    'Retour : Vari si l'objet pass� en param�tre est equivalent � 
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

    Protected MustOverride Function action(ByVal vncAction As vncActionFactColisage) As vncEtatMVTSTK

    Friend Overridable Function changeEtat(ByVal p_vncAction As vncActionFactColisage) As EtatMvtStock
        Dim nEtat As vncEnums.vncEtatMVTSTK
        nEtat = action(p_vncAction)
        Return createEtat(nEtat)
    End Function
End Class
