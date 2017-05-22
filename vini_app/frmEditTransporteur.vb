Imports vini_DB
Public Class frmEditTransporteur
    Inherits FrmVinicom


    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        m_datagrid.Enabled = True
    End Sub
    Private Sub frmEditTransporteur_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim oTa As vini_DB.dsVinicomTableAdapters.TRANSPORTEURTableAdapter

        oTa = New vini_DB.dsVinicomTableAdapters.TRANSPORTEURTableAdapter()
        oTa.Connection = Persist.oleDBConnection
        oTa.Fill(m_dsVinicom.TRANSPORTEUR)
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub frmEditTransporteur_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Dim oTa As vini_DB.dsVinicomTableAdapters.TRANSPORTEURTableAdapter

        oTa = New vini_DB.dsVinicomTableAdapters.TRANSPORTEURTableAdapter()
        oTa.Connection = Persist.oleDBConnection
        oTa.Update(m_dsVinicom.TRANSPORTEUR)
    End Sub
End Class