Imports System.Windows.Forms
Imports vini_DB

Public Class dlgReglements

    Public Sub setFact(ByVal objFact As Facture)
        Debug.Assert(objFact IsNot Nothing)

        Dim objReglement As Reglement
        objFact.loadReglements()


        For Each objReglement In objFact.colReglements
            m_bsReglement.Add(objReglement)
        Next
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim objReglement As Reglement
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        For Each objReglement In m_bsReglement
            objReglement.save()
        Next
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
