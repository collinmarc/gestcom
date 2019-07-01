Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmcrRecapColisage
    Inherits frmStatistiques

#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents dtMois As System.Windows.Forms.DateTimePicker
    Friend WithEvents tbCodeFourn As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbxDossier As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbIdFacture As System.Windows.Forms.TextBox
    Friend WithEvents lblFournisseur As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtMois = New System.Windows.Forms.DateTimePicker()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.tbCodeFourn = New System.Windows.Forms.TextBox()
        Me.lblFournisseur = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxDossier = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbIdFacture = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(202, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Mois :"
        '
        'dtMois
        '
        Me.dtMois.CustomFormat = "MMMM yyyy"
        Me.dtMois.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtMois.Location = New System.Drawing.Point(292, 7)
        Me.dtMois.Name = "dtMois"
        Me.dtMois.Size = New System.Drawing.Size(130, 20)
        Me.dtMois.TabIndex = 1
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.Location = New System.Drawing.Point(820, 6)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 23)
        Me.cbAfficher.TabIndex = 3
        Me.cbAfficher.Text = "Afficher"
        '
        'tbCodeFourn
        '
        Me.tbCodeFourn.Location = New System.Drawing.Point(558, 7)
        Me.tbCodeFourn.Name = "tbCodeFourn"
        Me.tbCodeFourn.Size = New System.Drawing.Size(100, 20)
        Me.tbCodeFourn.TabIndex = 2
        '
        'lblFournisseur
        '
        Me.lblFournisseur.AutoSize = True
        Me.lblFournisseur.Location = New System.Drawing.Point(463, 9)
        Me.lblFournisseur.Name = "lblFournisseur"
        Me.lblFournisseur.Size = New System.Drawing.Size(89, 13)
        Me.lblFournisseur.TabIndex = 8
        Me.lblFournisseur.Text = "Code Fournisseur"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Dossier :"
        '
        'cbxDossier
        '
        Me.cbxDossier.FormattingEnabled = True
        Me.cbxDossier.Location = New System.Drawing.Point(68, 6)
        Me.cbxDossier.Name = "cbxDossier"
        Me.cbxDossier.Size = New System.Drawing.Size(121, 21)
        Me.cbxDossier.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(202, 31)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 20)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Code Facture : "
        '
        'tbIdFacture
        '
        Me.tbIdFacture.Location = New System.Drawing.Point(292, 31)
        Me.tbIdFacture.Name = "tbIdFacture"
        Me.tbIdFacture.Size = New System.Drawing.Size(130, 20)
        Me.tbIdFacture.TabIndex = 11
        '
        'frmcrRecapColisage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1000, 678)
        Me.Controls.Add(Me.tbIdFacture)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbxDossier)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblFournisseur)
        Me.Controls.Add(Me.tbCodeFourn)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.dtMois)
        Me.Name = "frmcrRecapColisage"
        Me.Text = "Récapitulatif Colisage"
        Me.Controls.SetChildIndex(Me.dtMois, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.tbCodeFourn, 0)
        Me.Controls.SetChildIndex(Me.lblFournisseur, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.cbxDossier, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.tbIdFacture, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Overrides Function getResume() As String
        Return "Recapitulatif colisage"
    End Function



    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click

        Dim objReport As ReportDocument
        Dim oDS As vini_DB.dsVinicom
        Dim r1 As New crRecapColisageJournalier()
        Dim strReportName As String = r1.ResourceName
        Dim strCodeFourn As String
        Dim nCout As Decimal
        Dim pNumFactCol As String = ""
        Dim oFacCol As FactColisageJ
        Dim periode As String
        Dim nbJour As Integer

        Me.Cursor = Cursors.WaitCursor
        objReport = New ReportDocument
        objReport.Load(PATHTOREPORTS & strReportName)
        strCodeFourn = tbCodeFourn.Text.Replace("*", "%")
        nCout = CDec(Param.getConstante("CST_FACT_COL_PU_COLIS"))
        If tbIdFacture.Text <> "" Then
            pNumFactCol = tbIdFacture.Text
        End If

        debAffiche()
        If String.IsNullOrEmpty(pNumFactCol) Then
            Dim dDeb As Date = CDate("01/" & Me.dtMois.Value.Month & "/" & Me.dtMois.Value.Year)
            periode = dDeb.ToString("MMMM yyyy")
            Dim dFin As Date = dDeb.AddMonths(1).AddDays(-1)
            nbJour = dFin.Day

            oDS = FactColisageJ.GenereDataSetRecapColisage(dDeb, dFin, strCodeFourn, nCout, cbxDossier.Text)
        Else
            oFacCol = FactColisageJ.createandload(pNumFactCol)
            periode = oFacCol.periode
            nbJour = CDate(periode).AddMonths(1).AddDays(-1).Day
            oDS = FactColisageJ.GenereDataSetRecapColisage(oFacCol.id, nCout, oFacCol.dossierFact)
        End If
        setReportConnection(objReport)
        objReport.SetDataSource(oDS)
        'Les paramètres sont passé juste pour informations car ils ne sont pas utilisé

        objReport.SetParameterValue("Periode", periode)
        objReport.SetParameterValue("NbJour", nbJour)
        objReport.SetParameterValue("NbJour", nbJour)

        CrystalReportViewer1.ReportSource = objReport
        finAffiche()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub setReportConnection(ByVal objReport As ReportDocument)

        Dim myConnectionInfo As ConnectionInfo = New ConnectionInfo()
        myConnectionInfo.ServerName = Persist.oleDBConnection.DataSource
        myConnectionInfo.DatabaseName = Persist.oleDBConnection.Database
        myConnectionInfo.UserID = My.Settings.ReportCnxUser
        myConnectionInfo.Password = My.Settings.ReportCnxPassword

        Dim mySections As Sections = objReport.ReportDefinition.Sections
        For Each mySection As Section In mySections
            Dim myReportObjects As ReportObjects = mySection.ReportObjects
            For Each myReportObject As ReportObject In myReportObjects
                If myReportObject.Kind = ReportObjectKind.SubreportObject Then
                    Dim mySubreportObject As SubreportObject = CType(myReportObject, SubreportObject)
                    Dim subReportDocument As ReportDocument = mySubreportObject.OpenSubreport(mySubreportObject.SubreportName)
                    setReportConnection(subReportDocument)
                End If
            Next
        Next
        'On n'applique pas la connexion sur les tables, car ce rapport est base sur un dataset 

        ''Applique la connection sur les tables du rapport
        'Dim myTables As Tables = objReport.Database.Tables
        'For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
        '    Dim myTableLogonInfo As TableLogOnInfo = myTable.LogOnInfo
        '    myTableLogonInfo.ConnectionInfo = myConnectionInfo
        '    myTable.ApplyLogOnInfo(myTableLogonInfo)
        'Next
    End Sub

    Private Sub frmcrRecapColisage_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        cbxDossier.Items.Add(Dossier.VINICOM)
        cbxDossier.Items.Add(Dossier.HOBIVIN)
        cbxDossier.Text = Dossier.VINICOM
        dtMois.Value = DateTime.Now


    End Sub

    Private Sub cbxDossier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxDossier.SelectedIndexChanged
        If cbxDossier.Text = Dossier.VINICOM Then
            lblFournisseur.Visible = True
            tbCodeFourn.Visible = True
        End If
        If cbxDossier.Text = Dossier.HOBIVIN Then
            lblFournisseur.Visible = False
            tbCodeFourn.Visible = False
        End If
    End Sub
End Class
