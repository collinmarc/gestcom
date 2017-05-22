Imports vini_DB

Public Class frmCheckDatabase


    Private Sub cbExecuter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExecuter.Click
        execute()
    End Sub
    Private Sub execute()
        Me.Cursor = Cursors.WaitCursor
        Try
            m_oTAFactComError.Connection = Persist.oleDBConnection
            m_oTAFactComError.Fill(m_dsError.FACTCOMERROR, System.Convert.ToInt32(tbAnnee.Text))
            m_oTACmdErreur.Connection = Persist.oleDBConnection
            m_oTACmdErreur.Fill(m_dsError.COMMANDE_ERROR, System.Convert.ToInt32(tbAnnee.Text))
            m_oTACMD_ERROR3.Connection = Persist.oleDBConnection
            m_oTACMD_ERROR3.Fill(m_dsError.COMMANDE_SCMD_NULL, System.Convert.ToInt32(tbAnnee.Text))
            m_oTABASansMVTSTK.Fill(m_dsError.BASansMVTSTK, System.Convert.ToInt32(tbAnnee.Text), System.Convert.ToInt32(tbAnnee.Text))
        Catch ex As Exception

        End Try
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub frmFactComError_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        tbAnnee.Text = DateAndTime.Now.Year.ToString()
        execute()
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub
End Class