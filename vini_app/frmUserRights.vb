Imports vini_DB
Public Class FrmUserRights
    Inherits FrmVinicom
    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        DataGridView2.Enabled = True
    End Sub

    Private Sub FrmUserRights_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
 
    End Sub

    Private Sub FromUserRights_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim oTa As vini_DB.dsVinicomTableAdapters.USERSRIGHTSTableAdapter

        oTa = New vini_DB.dsVinicomTableAdapters.USERSRIGHTSTableAdapter()
        oTa.Connection = Persist.oleDBConnection
        oTa.Fill(m_dsVinicom.USERSRIGHTS)
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub cbValider_Click(sender As System.Object, e As System.EventArgs) Handles cbValider.Click
        Dim oTa As vini_DB.dsVinicomTableAdapters.USERSRIGHTSTableAdapter

        oTa = New vini_DB.dsVinicomTableAdapters.USERSRIGHTSTableAdapter()
        oTa.Connection = Persist.oleDBConnection
        oTa.Update(m_dsVinicom.USERSRIGHTS)

    End Sub
End Class
