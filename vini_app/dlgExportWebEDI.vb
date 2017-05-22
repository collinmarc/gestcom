Imports System.Windows.Forms

Public Class dlgExportWebEDI

    Private m_objCmd As vini_DB.CommandeClient

    Public Sub Setcommande(pobjCmd As vini_DB.CommandeClient)
        m_bsrcCommandeClt.Clear()
        m_bsrcCommandeClt.Add(pobjCmd)
        m_objCmd = pobjCmd
    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        m_objCmd.save()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
