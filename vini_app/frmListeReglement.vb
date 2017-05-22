Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Imports System.IO
Public Class frmListeReglement
    Inherits frmStatistiques
    Public Overrides Function getResume() As String
        Return "Liste des Reglements"
    End Function

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        Dim str As String

        Dim objReport As ReportDocument

        objReport = New ReportDocument
        If File.Exists(PATHTOREPORTS & "crListeReglements.rpt") Then

            objReport.Load(PATHTOREPORTS & "crListeReglements.rpt")


            objReport.SetParameterValue("ddeb", Me.dtDeb.Value.ToShortDateString())
            objReport.SetParameterValue("dfin", Me.dtFin.Value.ToShortDateString())


            str = tbCodeTiers.Text
            If String.IsNullOrEmpty(str) Then
                str = "*"
            End If
            str = Replace(str, "%", "*")
            objReport.SetParameterValue("codeTiers", Trim(str))

            If rbFactCommision.Checked Then
                objReport.SetParameterValue("TypeFact", 10)
            End If
            If rbFactTransport.Checked Then
                objReport.SetParameterValue("TypeFact", 12)
            End If
            If rbFactColisage.Checked Then
                objReport.SetParameterValue("TypeFact", 14)
            End If
            Persist.setReportConnection(objReport)
            CrystalReportViewer1.ReportSource = objReport
        End If

    End Sub

    Private Sub cbRechercher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRechercher.Click
        Dim frm As frmRechercheDB
        Dim oTiers As Tiers

        frm = New frmRechercheDB()
        If rbFactTransport.Checked Then
            frm.setTypeDonnees(vncTypeDonnee.CLIENT)
        Else
            frm.setTypeDonnees(vncTypeDonnee.FOURNISSEUR)
        End If
        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
            oTiers = frm.getElementSelectionne()
            Me.tbCodeTiers.Text = oTiers.code
        End If

    End Sub
End Class
