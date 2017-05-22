Option Explicit On 
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmCumulVentesArticles
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents dtdeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents tbCodeFourn As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtdeb = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtFin = New System.Windows.Forms.DateTimePicker
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.cbRecherche = New System.Windows.Forms.Button
        Me.tbCodeFourn = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(9, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Date de Début"
        '
        'dtdeb
        '
        Me.dtdeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdeb.Location = New System.Drawing.Point(128, 8)
        Me.dtdeb.Name = "dtdeb"
        Me.dtdeb.Size = New System.Drawing.Size(136, 20)
        Me.dtdeb.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(270, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 24)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Date de Fin"
        '
        'dtFin
        '
        Me.dtFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFin.Location = New System.Drawing.Point(360, 8)
        Me.dtFin.Name = "dtFin"
        Me.dtFin.Size = New System.Drawing.Size(104, 20)
        Me.dtFin.TabIndex = 4
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(789, 5)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 23)
        Me.cbAfficher.TabIndex = 5
        Me.cbAfficher.Text = "Afficher"
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(686, 5)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(80, 24)
        Me.cbRecherche.TabIndex = 8
        Me.cbRecherche.Text = "Recherche"
        '
        'tbCodeFourn
        '
        Me.tbCodeFourn.Location = New System.Drawing.Point(576, 8)
        Me.tbCodeFourn.Name = "tbCodeFourn"
        Me.tbCodeFourn.Size = New System.Drawing.Size(104, 20)
        Me.tbCodeFourn.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(474, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Code Fournisseur"
        '
        'frmCumulVentesArticles
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(984, 678)
        Me.Controls.Add(Me.cbRecherche)
        Me.Controls.Add(Me.tbCodeFourn)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.dtFin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtdeb)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmCumulVentesArticles"
        Me.Text = "Cumul des ventes par article"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtdeb, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.dtFin, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.tbCodeFourn, 0)
        Me.Controls.SetChildIndex(Me.cbRecherche, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub initFenetre()
        dtdeb.Value = DateAdd(DateInterval.Year, -1, Now())
    End Sub
    Private Sub afficheEtat()
        Dim str As String
        Dim objReport As ReportDocument

        objReport = New ReportDocument
        objReport.Load(PATHTOREPORTS & "crCumulVentesArticles.rpt")


        objReport.SetParameterValue("ddeb", Me.dtdeb.Value.ToShortDateString())
        objReport.SetParameterValue("dfin", Me.dtFin.Value.ToShortDateString())

        str = tbCodeFourn.Text
        str = Replace(str, "%", "*")
        objReport.SetParameterValue("CodeFourn", Trim(str))

        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport
    End Sub
    Private Sub rechercheFournisseur()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.FOURNISSEUR, tbCodeFourn)

        If Not objTiers Is Nothing Then
            tbCodeFourn.Text = objTiers.code
        End If
    End Sub 'RechercheFournisseur

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        afficheEtat()
    End Sub

    Private Sub cbRecherche_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        rechercheFournisseur()
    End Sub

    Private Sub frmCumulVentesArticles_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Public Overrides Function getResume() As String
        Return "Cumul des ventes par articles"
    End Function
End Class
