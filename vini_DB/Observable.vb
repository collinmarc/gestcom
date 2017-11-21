Imports System.Collections.Generic
Public MustInherit Class Observable
    Inherits racine

    Private m_ColObs As New List(Of IObservateur)

    Public Sub AjouteObservateur(pObservateur As IObservateur)
        m_ColObs.Add(pObservateur)
    End Sub
    Protected Sub Notifier()
        For Each obs As IObservateur In m_ColObs
            obs.Actualiser(Me)
        Next
    End Sub

    Public Overrides Function ToString() As String
        Return Me.GetType().ToString()
    End Function
End Class
