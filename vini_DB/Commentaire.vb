'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Commentaire
' Description : Commentaire
'===================================================================================================================================
'Membres de Classes
'==================
'Public
'Protected
'Private
'Membres d'instances
'==================
'Public
'   Code            : Code Tiers
'   comment         : Commentaire
'   toString()      : Rend 'objet sous forme de chaine
'   Equals()        : Test l'équivalence d'instance
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================Public MustInherit Class Persist
Public Class Commentaire
    Inherits racine
    Private m_code As String
    Private m_comment As String
    Public Property code() As String
        Get
            Return m_code
        End Get
        Set(ByVal Value As String)
            m_code = Value
        End Set
    End Property

    Public Property comment() As String
        Get
            Return m_comment
        End Get
        Set(ByVal Value As String)
            If Value <> m_comment Then
                m_comment = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Overrides Function toString() As String
        Return "COMM =(" & m_code & "," & m_comment & ")"
    End Function

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_comment
        End Get
    End Property
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objComm As Commentaire
        Dim bReturn As Boolean
        Try
            bReturn = True
            objComm = CType(obj, Commentaire)

            bReturn = bReturn And (code.Equals(objComm.code))
            bReturn = bReturn And (comment.Equals(objComm.comment))

            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Sub New()
        m_code = ""
        m_comment = ""
    End Sub

End Class
