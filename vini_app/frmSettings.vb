Public Class frmSettings

    Private Sub cbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSave.Click
        My.Settings.Save()
        My.Settings.Reload()
        Me.Close()
    End Sub
End Class