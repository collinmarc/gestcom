Option Strict Off
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmLstSCMDLitige
    Public Overrides Function getResume() As String
        Return "Liste des Commissions en litige"
    End Function


    ''' <summary>
    ''' Recherche d'un fournisseur
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub rechercheFournisseur()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.FOURNISSEUR, tbCodeFournisseur)

        If Not objTiers Is Nothing Then
            tbCodeFournisseur.Text = objTiers.code
        End If
    End Sub 'rechercheFournisseur

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click

        Dim objReport As ReportDocument
        Dim str As String

        objReport = New ReportDocument
        objReport.Load(PATHTOREPORTS & "crLstSousCommandesLitige.rpt")


        objReport.SetParameterValue("ddeb", Me.dtDateDeb.Value.ToShortDateString())
        objReport.SetParameterValue("dfin", Me.dtDateFin.Value.ToShortDateString())

        str = tbCodeFournisseur.Text + "%"
        str = Replace(str, "%", "*")
        objReport.SetParameterValue("codeFourn", Trim(str))

        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport
        CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.GroupTree
    End Sub

    Private Sub frmListeSousCommandes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tbCodeFournisseur.Text = "%"
    End Sub

    Private Sub cbRechercher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRechercher.Click
        rechercheFournisseur()
    End Sub


End Class
