Public Class dlgInternet

    Private Sub dlgInternet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        WebBrowser4.Navigate(My.Settings.urlImportInternet)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class