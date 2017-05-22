Imports vini_DB
Public Class frmExportDossier

    Private Sub tbBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbBrowse.Click
        selectionneDossier()
    End Sub

    Private Sub selectionneDossier()
        If m_dlgDossier.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            tbDossier.Text = m_dlgDossier.SelectedPath()
        End If
    End Sub
    Private Shadows Sub DisplayStatus(ByVal strMessage As String)

        Dim oStatus As clsExportstatus
        oStatus = New clsExportstatus()
        oStatus.statusDate = Now()
        oStatus.statusMessage = strMessage
        m_bsrcStatus.Add(oStatus)
        dgvStatus.Refresh()
        If dgvStatus.Rows.Count > dgvStatus.DisplayedRowCount(True) Then
            dgvStatus.FirstDisplayedScrollingRowIndex = dgvStatus.Rows.Count - dgvStatus.DisplayedRowCount(True) + 1
        End If
        System.IO.File.AppendAllText("./vini_internet.trace", Now().ToShortTimeString() + strMessage + vbCrLf)
        Trace.WriteLine(Now().ToShortTimeString() + strMessage)
        'lbStatus.Items.Add(Now() + ":" + strMessage)
        'lbStatus.SetSelected(lbStatus.Items.Count - 1, True)
        'lbStatus.Refresh()
    End Sub

    Private Sub tbExporter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbExporter.Click
        exporter()
    End Sub

    Private Sub exporter()

        Dim bExportOK As Boolean = True
        Dim bReturn As Boolean
        Dim strFolder As String
        Dim oFTPvinicom As clsFTPVinicom
        Dim nFichiersAExporter As Integer
        Dim nSousCommandesExportees As Integer

        Me.Cursor = Cursors.WaitCursor
        Try
            strFolder = tbDossier.Text
            nFichiersAExporter = System.IO.Directory.GetFiles(strFolder).Length
            dgvStatus.Rows.Clear()
            bReturn = True
            DisplayStatus("Nombre de Fichiers à exporter: " & nFichiersAExporter)
            DisplayStatus("Transferts des fichiers vers " + Param.getConstante("FTP_HOSTNAME"))
            oFTPvinicom = New clsFTPVinicom 'Création avec les paramètres par defaut

            If True Then
                If (Not oFTPvinicom.IsLockFrom()) Then
                    nSousCommandesExportees = oFTPvinicom.uploadFromDir(strFolder)
                    If Not String.IsNullOrEmpty(oFTPvinicom.ErrorDescription) Then
                        DisplayStatus(oFTPvinicom.ErrorDescription())
                    Else
                        DisplayStatus("Fin de transfert des fichiers ")
                        DisplayStatus("Nombre de fichiers exportés : " & (nSousCommandesExportees - 1) & "+1")
                        bReturn = True
                    End If
                Else
                    DisplayStatus("Serveur internet vérrouillé")
                End If
            Else
                DisplayStatus("Connexion impossible (" + Param.getConstante("FTP_USERNAME") + " /" + Param.getConstante("FTP_PASSWORD") + ")")
            End If

            WaitnSeconds(10)
            ActiverImportBAF()




        Catch ex As Exception
            MsgBox("Erreur" + ex.Message)
        End Try
        Me.Cursor = Cursors.Default


    End Sub
    Private Sub ActiverImportBAF()
        Dim odlg As dlgInternet = New dlgInternet()
        dlgInternet.ShowDialog()
    End Sub


    Private Sub frmExportDossier_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub
    Private Sub initFenetre()
        Dim oFTPvinicom As clsFTPVinicom
        oFTPvinicom = New clsFTPVinicom 'Création avec les paramètres par defaut
        'If oFTPvinicom.connect() Then
        If True Then
            cbDeverrouillageFrom.Enabled = oFTPvinicom.IsLockFrom()
            cbDeverrouillageTo.Enabled = oFTPvinicom.IsLockTo()
        End If

    End Sub

    Private Sub cbDeverrouillageFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDeverrouillageFrom.Click
        Me.Cursor = Cursors.WaitCursor
        deverouillageFrom()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub deverouillageFrom()
        Dim oFTPvinicom As clsFTPVinicom
        oFTPvinicom = New clsFTPVinicom 'Création avec les paramètres par defaut
        'If oFTPvinicom.connect() Then
        If True Then
            oFTPvinicom.unlockfrom()
            cbDeverrouillageFrom.Enabled = oFTPvinicom.IsLockFrom()
        End If
    End Sub
    Private Sub deverouillageTo()
        Dim oFTPvinicom As clsFTPVinicom
        oFTPvinicom = New clsFTPVinicom 'Création avec les paramètres par defaut
        'If oFTPvinicom.connect() Then
        If True Then
            oFTPvinicom.unlockTo()
            cbDeverrouillageTo.Enabled = oFTPvinicom.IsLockTo()
        End If
    End Sub

    Private Sub cbDeverrouillageTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDeverrouillageTo.Click
        Me.Cursor = Cursors.WaitCursor
        deverouillageTo()
        Me.Cursor = Cursors.Default
    End Sub
End Class