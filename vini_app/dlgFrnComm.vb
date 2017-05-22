Imports System.Windows.Forms
Imports vini_DB

Public Class dlgFrnComm

    Private mcol As Collection
    Private m_FRNid As Integer

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        TauxComm.deleteTauxComms(m_FRNid)
        For Each oTaux As TauxComm In m_bsrcFRNCOM.List
            oTaux.bNew = True
            oTaux.save()
        Next
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub setFRNID(ByVal pid As Integer)
        m_FRNid = pid

        For Each oParam As Param In Param.colTypeClient
            m_bsrcTypeClient.Add(oParam)
        Next
        mcol = TauxComm.getListe(m_FRNid)
        For Each oTaux As TauxComm In mcol
            m_bsrcFRNCOM.Add(oTaux)
        Next
        'm_bsrcFRNCOM.Filter = "bDeleted = false"
    End Sub

    Private Sub m_bsrcFRNCOM_AddingNew(ByVal sender As System.Object, ByVal e As System.ComponentModel.AddingNewEventArgs) Handles m_bsrcFRNCOM.AddingNew
        'Récupération du premier Type non utilisé
        Dim strTypeClient As String
        Dim bCodeUtilise As Boolean
        Dim oTauxComm As TauxComm
        bCodeUtilise = False
        strTypeClient = String.Empty
        For Each oParam As Param In m_bsrcTypeClient.List
            bCodeUtilise = False
            strTypeClient = oParam.code
            For Each oTaux As TauxComm In m_bsrcFRNCOM.List
                If oTaux.TypeClt.Equals(strTypeClient) Then
                    ' Le code est utilisé => on sort
                    bCodeUtilise = True
                    Exit For
                End If
            Next
            'sur le premier code non utilisé on sort
            If Not bCodeUtilise Then
                Exit For
            End If
        Next
        oTauxComm = New TauxComm(m_FRNid)
        If Not bCodeUtilise And Not String.IsNullOrEmpty(strTypeClient) Then
            oTauxComm.TypeClt = strTypeClient
        End If
        e.NewObject = oTauxComm
    End Sub
End Class
