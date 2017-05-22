Public Class frmLibere

    Private Sub cbAffiche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAffiche.Click
        m_taLOCK.Connection = vini_DB.Persist.oleDBConnection
        If tbMachineName.Text.Length = 0 Then
            m_taLOCK.Fill(DsVinicom1.LOCK)
        Else
            m_taLOCK.FillByMachineName(DsVinicom1.LOCK, tbMachineName.Text)
        End If
    End Sub

    Private Sub cbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSave.Click
        m_taLOCK.Update(DsVinicom1.LOCK)
    End Sub
End Class