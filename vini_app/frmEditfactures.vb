Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Windows.Forms.Cursors
Imports vini_DB

Public Class frmEditfactures
    Inherits FrmVinicom
    Public Overrides Function getResume() As String
        Return "Edition des factures"
    End Function

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub
    Private Sub frmEditfactures_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dtDateDeb.Value = CDate("01/" + DateAndTime.Now.Month.ToString + "/" + DateAndTime.Now.Year.ToString)
        ' Fin du mois = Debutdu moi suivant -1 jour
        dtDateFin.Value = CDate("01/" + DateAndTime.Now.AddMonths(1).Month.ToString() + "/" + DateAndTime.Now.AddMonths(1).Year.ToString()).AddDays(-1)
        rbFactcom.Checked = True
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        Me.Cursor = WaitCursor
        afficheRapport()
        Me.Cursor = [Default]
    End Sub

    Private Sub afficheRapport()
        Dim strReport As String
        Dim objReport As ReportDocument

        strReport = String.Empty
        If rbFactcom.Checked Then
            If rbCourier.Checked Then
                strReport = "crFactCom_courrier.rpt"
            End If
            If rbFacture.Checked Then
                strReport = "crFactCom.rpt"
            End If
            If rbReleve.Checked Then
                strReport = "crFactCom_Releve.rpt"
            End If
        End If

        If rbFactTRP.Checked Then
            strReport = "crFactureTRP.rpt"
        End If
        If rbFactCol.Checked Then
            strReport = "crFactureColisage.rpt"
        End If
        objReport = New ReportDocument
        objReport.Load(strReport)

        '============== RECUPERATION DES INSTANCES A EDITER ===========================
        Dim oCol As Collection
        Dim ddeb As Date
        Dim dfin As Date
        Dim tabIds As ArrayList
        Dim objFact As Facture

        ddeb = CDate(dtDateDeb.Value.ToShortDateString())
        dfin = CDate(dtDateFin.Value.ToShortDateString())
        oCol = Nothing
        If rbFactcom.Checked Then
            oCol = FactCom.getListe(ddeb, dfin, tbCodeTiers.Text)
        End If
        If rbFactTRP.Checked Then
            oCol = FactTRP.getListe(ddeb, dfin, tbCodeTiers.Text)
        End If
        If rbFactCol.Checked Then
            oCol = FactColisageJ.getListe(ddeb, dfin, tbCodeTiers.Text)
        End If
        tabIds = New ArrayList()
        For Each objFact In oCol
            tabIds.Add(objFact.id)
        Next

        If tabIds.Count > 0 Then
            '==================== PASSAGE des Paramètres ====================================
            If rbFactcom.Checked Then
                objReport.SetParameterValue("IdFactures", tabIds.ToArray())
                objReport.SetParameterValue("bEntete", ckEntete.Checked)
            End If

            If rbFactTRP.Checked Then
                objReport.SetParameterValue("IdFactures", tabIds.ToArray())
                objReport.SetParameterValue("bEntete", ckEntete.Checked)
                objReport.SetParameterValue("LGNUMGAZOLE", Param.LGNUM_GAZOLE)
            End If
            If rbFactCol.Checked Then
                objReport.SetParameterValue("IdFactures", tabIds.ToArray())
                objReport.SetParameterValue("bEntete", ckEntete.Checked)
            End If

            Persist.setReportConnection(objReport)
            CrystalReportViewer1.ReportSource = objReport
        End If


    End Sub

    Private Sub rbFactcom_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFactcom.CheckedChanged
        grpTypeEditFactcom.Visible = rbFactcom.Checked
        rbFacture.Checked = True
    End Sub

    Private Sub rbCourier_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbCourier.CheckedChanged

    End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class