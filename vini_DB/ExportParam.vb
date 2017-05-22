'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : ExportParam
' Description : Element de la table Export_param
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
Public Class ExportParam
    Inherits Persist
    Private m_exptype As String
    Private m_Ligne As Integer
    Private m_Champs As Integer
    Private m_TypeChamps As String
    Private m_Valeur As String
    Private m_Longueur As Integer
    Private m_description As String
#Region "Accesseurs"
    Private Sub New()
        m_exptype = ""
    End Sub
    Public Sub New(pExpType As String, pLigne As Integer, pChamps As Integer, pTypeChamps As String, pValeur As String, pLongueur As Integer, pDescription As String)
        m_exptype = pExpType
        m_Ligne = pLigne
        m_Champs = pChamps
        m_Valeur = pValeur
        m_Longueur = pLongueur
        m_description = pDescription
    End Sub

    Public Property exptype As String
        Get
            Return m_exptype
        End Get
        Set(ByVal Value As String)
            If Value <> m_exptype Then
                RaiseUpdated()
                m_exptype = Value
            End If
        End Set
    End Property
    Public Property Ligne As Integer
        Get
            Return m_Ligne
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_Ligne Then
                RaiseUpdated()
                m_Ligne = Value
            End If
        End Set
    End Property
    Public Property champs As Integer
        Get
            Return m_Champs
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_Champs Then
                RaiseUpdated()
                m_Champs = Value
            End If
        End Set
    End Property

    Public Property TypeChamps As String
        Get
            Return m_TypeChamps
        End Get
        Set(ByVal Value As String)
            If Value <> m_TypeChamps Then
                RaiseUpdated()
                m_TypeChamps = Value
            End If
        End Set
    End Property
    Public Property Valeur As String
        Get
            Return m_Valeur
        End Get
        Set(ByVal Value As String)
            If Value <> m_Valeur Then
                RaiseUpdated()
                m_Valeur = Value
            End If
        End Set
    End Property
    Public Property Longueur As String
        Get
            Return m_Longueur
        End Get
        Set(ByVal Value As String)
            If Value <> m_Longueur Then
                RaiseUpdated()
                m_Longueur = Value
            End If
        End Set
    End Property
    Public Property Description As String
        Get
            Return m_description
        End Get
        Set(ByVal Value As String)
            If Value <> m_description Then
                RaiseUpdated()
                m_description = Value
            End If
        End Set
    End Property
    'Private m_description As String
#End Region

#Region "Interface Persist"
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Return False
    End Function 'DBLoad
    Public Overrides Function save() As Boolean
        Return False
    End Function
    Friend Overrides Function delete() As Boolean
        Return False

    End Function ' delete
    Public Overrides Function checkForDelete() As Boolean
        Return False
    End Function 'checkForDelete

    Friend Overrides Function insert() As Boolean
        Return False
    End Function 'insert
    Friend Overrides Function update() As Boolean
        Return False
    End Function 'Update
#End Region
    Public Shared Function GetListe(pExpType As String) As Collection
        Dim ocolReturn As New Collection
        Try
            ocolReturn = ListeExportParam(pExpType)

        Catch ex As Exception
            setError(ex.Message)
            ocolReturn = Nothing
        End Try
        Return ocolReturn
    End Function

End Class
