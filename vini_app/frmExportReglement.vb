Imports System.IO
Imports vini_DB

Public Class frmExportReglement
    Public Overrides Function getResume() As String
        Return "Export Facture vers comptabilité"
    End Function

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub


    Private Sub frmExportReglement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO : cette ligne de code charge les données dans la table 'DsVinicom.CONSTANTES'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
        Me.CONSTANTESTableAdapter.Connection = Persist.oleDBConnection
        Me.CONSTANTESTableAdapter.Fill(Me.DsVinicom.CONSTANTES)

    End Sub

    Private Sub cbParcourir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbParcourir.Click
        Me.FolderBrowserDialog1.Description = My.Resources.STR_EXPORTCOMPTA_PATH
        Me.FolderBrowserDialog1.SelectedPath = Me.tbPath.Text
        Me.FolderBrowserDialog1.ShowDialog()
        tbPath.Text = Me.FolderBrowserDialog1.SelectedPath

    End Sub

    Private Sub cbExporter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExporter.Click
        exporter()
    End Sub
    ''' <summary>
    ''' Exporte les Reglements vers le répertoire indiqué
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub exporter()
        Dim strPath As String
        Dim strFile As String
        Dim dateDeb As Date
        Dim dateFin As Date
        Dim bExportFactureCommision As Boolean
        Dim bExportFactureColisage As Boolean
        Dim bExportFactureTransport As Boolean
        Dim bReturn As Boolean
        Dim colFact As Collection

        Try


            'Lecture des valeurs

            strPath = tbPath.Text
            dateDeb = dtdeb.Value.ToShortDateString()
            dateFin = dtFin.Value.ToShortDateString()
            bExportFactureCommision = rbFactCom.Checked
            bExportFactureColisage = rbFactColisage.Checked
            bExportFactureTransport = rbFactTRP.Checked

            lbMessages.Items.Clear()
            bReturn = True
            If bExportFactureCommision Then
                displayMessage("== EXPORTATION DES REGLEMENTS de FACTURES DE COMMISSIONS ==")
                strFile = strPath + "/REGLMT_COM" + Format(Now(), "yyyyMMdd")
                colFact = Reglement.getListe(dateDeb, dateFin, vncEnums.vncTypeDonnee.FACTCOMM, vncEtatReglement.vncRGLMT_Saisi)
                bReturn = bReturn And exportReglement(strFile, colFact)
            End If
            If bExportFactureTransport Then
                displayMessage("== EXPORTATION DES REGLEMENTS de FACTURES DE TRANSPORT ==")
                strFile = strPath + "/REGLMT_TRP" + Format(Now(), "yyyyMMdd")
                colFact = Reglement.getListe(dateDeb, dateFin, vncEnums.vncTypeDonnee.FACTTRP, vncEtatReglement.vncRGLMT_Saisi)
                bReturn = bReturn And exportReglement(strFile, colFact)
            End If
            If bExportFactureColisage Then
                displayMessage("== EXPORTATION DES REGLEMENTS de FACTURES DE COLISAGE ==")
                strFile = strPath + "/REGLMT_COL" + Format(Now(), "yyyyMMdd")
                colFact = Reglement.getListe(dateDeb, dateFin, vncEnums.vncTypeDonnee.FACTCOL, vncEtatReglement.vncRGLMT_Saisi)
                bReturn = bReturn And exportReglement(strFile, colFact)
            End If


            If bReturn Then
                MessageBox.Show(My.Resources.STR_EXPORTCOMPTA_OK)
                If (MessageBox.Show(My.Resources.STR_VALIDATION_EXPORT_REGLEMENT, "Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes) Then
                    ValiderExport()
                End If
            Else
                MessageBox.Show(My.Resources.STR_EXPORTCOMPTA_NOK, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            displayMessage(ex.Message)
        End Try


    End Sub

    ''' <summary>
    ''' Exporte les factures vers le répertoire indiqué
    ''' </summary>
    ''' <remarks></remarks>
    Private Function ValiderExport() As Boolean
        Dim strPath As String
        Dim dateDeb As Date
        Dim dateFin As Date
        Dim bExportFactureCommision As Boolean
        Dim bExportFactureColisage As Boolean
        Dim bExportFactureTransport As Boolean
        Dim bReturn As Boolean
        Dim col As Collection

        Try


            'Lecture des valeurs

            strPath = tbPath.Text
            dateDeb = dtdeb.Value.ToShortDateString()
            dateFin = dtFin.Value.ToShortDateString()
            bExportFactureCommision = rbFactCom.Checked
            bExportFactureColisage = rbFactColisage.Checked
            bExportFactureTransport = rbFactTRP.Checked

            bReturn = True
            If bExportFactureCommision Then
                displayMessage("== Validation Export Facture de commisions ==")
                col = Reglement.getListe(dateDeb, dateFin, vncTypeDonnee.FACTCOMM, vncEtatReglement.vncRGLMT_Saisi)
                For Each obj As Reglement In col
                    obj.changeEtat(vncActionReglement.vncActionExporter)
                    obj.save()
                Next
            End If
            If bExportFactureTransport Then
                displayMessage("== Validation Export Facture de transport ==")
                col = Reglement.getListe(dateDeb, dateFin, vncTypeDonnee.FACTTRP, vncEtatReglement.vncRGLMT_Saisi)
                For Each obj As Reglement In col
                    obj.changeEtat(vncActionReglement.vncActionExporter)
                    obj.save()
                Next
            End If
            If bExportFactureColisage Then
                displayMessage("== Validation Export Facture de colisage ==")
                col = Reglement.getListe(dateDeb, dateFin, vncTypeDonnee.FACTCOL, vncEtatReglement.vncRGLMT_Saisi)
                For Each obj As Reglement In col
                    obj.changeEtat(vncActionReglement.vncActionExporter)
                    obj.save()
                Next
            End If
            displayMessage("Validation terminée")
            bReturn = True
        Catch ex As Exception
            displayMessage(ex.Message)
            bReturn = False
        End Try

        Return bReturn

    End Function
    ''' <summary>
    ''' Export les Reglements vers le répertoire indiqué
    ''' </summary>
    ''' <param name="pstrFile">répertoire d'exportation</param>
    ''' <param name="pcol">Collection à exporter</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function exportReglement(ByVal pstrFile As String, ByVal pcol As Collection) As Boolean
        Dim bReturn As Boolean
        Dim bExport As Boolean
        Dim bFileOk As Boolean = False
        Dim objReglement As Reglement
        Dim nTotal As Decimal = 0
        Try
            bReturn = True
            If File.Exists(pstrFile) Then
                bFileOk = False
                If MsgBox("Le Fichier " & pstrFile & " existe voulez-vous l'écraser?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    File.Delete(pstrFile)
                    bFileOk = True
                End If
            Else
                bFileOk = True
            End If

            If (bFileOk) Then

                For Each objReglement In pcol
                    displayMessage("Export Reglement " + objReglement.DateReglement + " " + objReglement.Montant.ToString("c") + " " + objReglement.Reference)
                    bExport = objReglement.Exporter(pstrFile)
                    If bExport Then
                        nTotal = nTotal + objReglement.Montant
                    Else
                        displayMessage("Erreur pendant l'export de " & objReglement.DateReglement.ToShortDateString() & " - " & objReglement.Montant & ":" & Reglement.getErreur())
                    End If
                    bReturn = bReturn + bExport
                Next
            End If

            displayMessage("Total exporté : " + nTotal.ToString("c"))
        Catch ex As Exception
            displayMessage(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'exportReglement
    Private Sub displayMessage(ByVal strMessage As String)
        lbMessages.Items.Add(DateAndTime.Now.ToLongTimeString() + ":" + strMessage)
        lbMessages.SetSelected(lbMessages.Items.Count - 1, True)
        lbMessages.Refresh()
    End Sub
End Class
