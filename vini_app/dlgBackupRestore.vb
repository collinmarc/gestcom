Imports System.Windows.Forms
Imports vini_db

Public Class dlgBackupRestore

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub cbParcourirBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbParcourirBackup.Click
        m_dlgSaveFile.InitialDirectory = tbPath.Text
        If (m_dlgSaveFile.ShowDialog(Me) = Windows.Forms.DialogResult.OK) Then
            tbPath.Text = m_dlgSaveFile.FileName
        End If
    End Sub

    Private Sub cbParcourirRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbParcourirRestore.Click
        m_dlgOpenFile.InitialDirectory = tbPathRestore.Text
        If (m_dlgOpenFile.ShowDialog(Me) = Windows.Forms.DialogResult.OK) Then
            tbPathRestore.Text = m_dlgOpenFile.FileName
        End If

    End Sub

    Private Sub cbBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBackup.Click
        Dim bReturn As Boolean
        Me.Cursor = Cursors.WaitCursor
        bReturn = Persist.dbConnection.Backup(tbPath.Text)
        If bReturn Then
            MessageBox.Show("Sauvegarde terminée avec succès")
        Else
            MessageBox.Show("Erreur durant la sauvegarde : " + Persist.getErreur())
        End If
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cbRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRestore.Click
        Dim bReturn As Boolean

        Me.Cursor = Cursors.WaitCursor
        bReturn = Persist.dbConnection.Restore(tbPathRestore.Text)
        If bReturn Then
            MessageBox.Show("Restauration terminée avec succès")
        Else
            MessageBox.Show("Erreur durant la sauvegarde : " + Persist.getErreur())
        End If
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub tbPathRestore_TextChanged(sender As Object, e As EventArgs) Handles tbPathRestore.TextChanged
        If Not String.IsNullOrEmpty(tbPathRestore.Text) Then
            cbRestore.Enabled = True
        End If
    End Sub
End Class
