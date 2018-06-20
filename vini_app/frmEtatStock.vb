Option Explicit On 
Option Strict Off
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmEtatStock
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
    Friend WithEvents tbCodeFourn As System.Windows.Forms.TextBox
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents Label2 As Label
    Friend WithEvents cbxDossier As ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbCodeFourn = New System.Windows.Forms.TextBox()
        Me.cbRecherche = New System.Windows.Forms.Button()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxDossier = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(9, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Code Fournisseur"
        '
        'tbCodeFourn
        '
        Me.tbCodeFourn.Location = New System.Drawing.Point(112, 8)
        Me.tbCodeFourn.Name = "tbCodeFourn"
        Me.tbCodeFourn.Size = New System.Drawing.Size(104, 20)
        Me.tbCodeFourn.TabIndex = 0
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(222, 5)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(80, 24)
        Me.cbRecherche.TabIndex = 1
        Me.cbRecherche.Text = "Recherche"
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(441, 31)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 24)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "Afficher"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Dossier :"
        '
        'cbxDossier
        '
        Me.cbxDossier.FormattingEnabled = True
        Me.cbxDossier.Location = New System.Drawing.Point(112, 35)
        Me.cbxDossier.Name = "cbxDossier"
        Me.cbxDossier.Size = New System.Drawing.Size(104, 21)
        Me.cbxDossier.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(222, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(82, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Date de calcul :"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(311, 35)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(100, 20)
        Me.DateTimePicker1.TabIndex = 3
        '
        'frmEtatStock
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(952, 646)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbxDossier)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.cbRecherche)
        Me.Controls.Add(Me.tbCodeFourn)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmEtatStock"
        Me.Text = "Etat du stock"
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.tbCodeFourn, 0)
        Me.Controls.SetChildIndex(Me.cbRecherche, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.cbxDossier, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.DateTimePicker1, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub afficheEtat()

        Dim objReport As ReportDocument
        Dim str As String

        setcursorWait()
        objReport = New ReportDocument
        objReport.Load(PATHTOREPORTS & "crEtatStock.rpt")
        str = tbCodeFourn.Text
        str = Replace(str, "%", "*")
        objReport.SetParameterValue("CodeFourn", Trim(str))
        objReport.SetParameterValue("Dossier", cbxDossier.Text)
        objReport.SetParameterValue("DateCalcul", DateTimePicker1.Value)
        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport
        restoreCursor()

    End Sub
    Private Sub rechercheFournisseur()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.FOURNISSEUR, tbCodeFourn)

        If Not objTiers Is Nothing Then
            tbCodeFourn.Text = objTiers.code
        End If
    End Sub 'RechercheFournisseur


    Private Sub cbAfficher_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        afficheEtat()
    End Sub

    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        rechercheFournisseur()
    End Sub

    Private Sub frmEtatStock_Load(sender As Object, e As EventArgs) Handles Me.Load
        cbxDossier.Items.Clear()
        cbxDossier.Items.Add(Dossier.VINICOM)
        cbxDossier.Items.Add(Dossier.HOBIVIN)
        DateTimePicker1.Value = Now
    End Sub
End Class
