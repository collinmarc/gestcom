'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Racine
' Description : classe de base à toutes les classes
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
Public MustInherit Class racine
    '    Inherits SingleADOConnection
    Private Shared m_ErreurSource As String
    Private Shared m_ErreurMessage As String
    Public MustOverride Shadows Function toString() As String
    Public MustOverride ReadOnly Property shortResume() As String

    Public Event Updated()

    Protected Shared Sub setError(ByVal strSource As String, ByVal strMessage As String)
        m_ErreurSource = strSource
        m_ErreurMessage = strMessage
        If (m_ErreurMessage <> "") Then
            Trace.WriteLine(strSource & ":" & strMessage)
        End If
    End Sub
    Protected Shared Sub setError(ByVal strMessage As String)
        Try

            m_ErreurSource = ""
            m_ErreurMessage = strMessage
            If (m_ErreurMessage <> "") Then
                Trace.WriteLine(m_ErreurSource & ":" & m_ErreurMessage)
            End If
        Catch ex As Exception
            If (m_ErreurMessage <> "") Then
                Trace.WriteLine("???:" & m_ErreurMessage)
            End If

        End Try
    End Sub

    Public Shared Function bErreur() As Boolean
        Return (Len(m_ErreurMessage) <> 0)
    End Function

    Public Shared Function getErreur() As String
        Return "Erreur :" & m_ErreurSource & "|" & m_ErreurMessage
    End Function

    Public Shared Sub cleanErreur()
        m_ErreurMessage = ""
        m_ErreurSource = ""
    End Sub

    Public Overridable Shadows Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        bReturn = obj.GetType.Name.Equals(Me.GetType().Name)  ' Meme Classe D'objet
        Return bReturn
    End Function

    Protected Overridable Sub RaiseUpdated()
        RaiseEvent Updated()
    End Sub
End Class
