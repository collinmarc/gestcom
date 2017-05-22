Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Imports System.IO
Public Class frmTableauBordFacture
    Inherits frmStatistiques

    Public Overrides Function getResume() As String
        Return "Tableau de bord Factures"
    End Function

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        Dim objReport As ReportDocument
        Dim r1 As New crTableauBordFacture
        Dim str As String
        Dim strReportName As String = r1.ResourceName


        objReport = New ReportDocument

        objReport.Load(PATHTOREPORTS & strReportName)


        objReport.SetParameterValue("ddeb", Me.dtDeb.Value.ToShortDateString())
        objReport.SetParameterValue("dfin", Me.dtFin.Value.ToShortDateString())


        objReport.SetParameterValue("bNonSoldee", ckFactureSoldee.Checked)

        If rbFacturecommission.Checked Then
            objReport.SetParameterValue("TypeFacture", vncEnums.vncTypeDonnee.FACTCOMM)
        End If
        If rbFactureTransport.Checked Then
            objReport.SetParameterValue("TypeFacture", vncEnums.vncTypeDonnee.FACTTRP)
        End If
        If rbFactureColisage.Checked Then
            objReport.SetParameterValue("TypeFacture", vncEnums.vncTypeDonnee.FACTCOL)
        End If

        If String.IsNullOrEmpty(tbCodeTiers.Text) Then
            str = "*"
        Else
            str = tbCodeTiers.Text + "*"
        End If

        str = str.Replace("%", "*")
        objReport.SetParameterValue("codetiers", str)


        objReport.SetParameterValue("bDetail", ckDetail.Checked)



        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport

    End Sub

    Private Sub rbFacturecommission_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFacturecommission.CheckedChanged, rbFactureTransport.CheckedChanged, rbFactureColisage.CheckedChanged
        AffichelaCodeTiers()
    End Sub

    Private Sub AffichelaCodeTiers()
        If rbFacturecommission.Checked Then
            laCodetiers.Text = "Code Fournisseur: "
        End If
        If rbFactureTransport.Checked Then
            laCodetiers.Text = "Code Client: "
        End If
        If rbFactureColisage.Checked Then
            laCodetiers.Text = "Code Fournisseur: "
        End If
    End Sub

    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        Dim frm As frmRechercheDB
        Dim oTiers As Tiers
        frm = New frmRechercheDB


        If rbFacturecommission.Checked Then
            frm.setTypeDonnees(vncTypeDonnee.FOURNISSEUR)
        End If
        If rbFactureTransport.Checked Then
            frm.setTypeDonnees(vncTypeDonnee.CLIENT)
        End If
        If rbFactureColisage.Checked Then
            frm.setTypeDonnees(vncTypeDonnee.FOURNISSEUR)
        End If
        If (frm.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            oTiers = CType(frm.getElementSelectionne(), Tiers)
            tbCodeTiers.Text = oTiers.code
        End If

    End Sub

    Private Sub frmLstFacturesSolde_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        rbFacturecommission.Checked = True
    End Sub
End Class
