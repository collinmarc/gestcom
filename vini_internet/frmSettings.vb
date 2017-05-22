Public Class frmSettings

    Private Sub cbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSave.Click
        My.Settings.urlImportInternet = New System.Uri(TextBox4.Text)
        My.Settings.Save()
        My.Settings.Reload()
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub frmSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox4.Text = Global.vini_internet.My.MySettings.Default.urlImportInternet.ToString()
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub
End Class