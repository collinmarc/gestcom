Imports vini_DB
Public Class frmGestParamContenant
    Private m_DeletedRows As Collection

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        DataGridView1.Enabled = True
    End Sub
    Private Sub frmEditTransporteur_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objParam As contenant

        DataGridView1.DataSource = Nothing

        For Each objParam In contenant.colContenant
            m_bsrcParam.Add(objParam)
        Next
        DataGridView1.DataSource = m_bsrcParam
    End Sub

    Private Sub cbCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCancel.Click
        Me.Close()
    End Sub

    Private Sub cbOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOK.Click
        Dim objParam As contenant

        'Sauvegarde des lignes "supprimées"
        For Each objParam In m_DeletedRows
            objParam.Save()
        Next
        'Sauvegarde des autres lignes 
        For Each objParam In m_bsrcParam
            objParam.Save()
        Next
        contenant.LoadcolContenants()
        Me.Close()
    End Sub

    Private Sub m_bsrcParamModeReglement_AddingNew(ByVal sender As System.Object, ByVal e As System.ComponentModel.AddingNewEventArgs) Handles m_bsrcParam.AddingNew
        Try
            e.NewObject = New contenant
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataGridView1_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DataGridView1.UserDeletingRow
        Dim objParam As Param
        objParam = m_bsrcParam.Current
        If Not objParam Is Nothing Then
            objParam.bDeleted = True
            m_DeletedRows.Add(objParam)
        End If

    End Sub

    Public Sub New()

        ' Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        m_DeletedRows = New Collection()

    End Sub
End Class
