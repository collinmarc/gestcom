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
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbCodeFourn = New System.Windows.Forms.TextBox()
        Me.cbRecherche = New System.Windows.Forms.Button()
        Me.cbAfficher = New System.Windows.Forms.Button()
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
        Me.tbCodeFourn.TabIndex = 2
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(222, 5)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(80, 24)
        Me.cbRecherche.TabIndex = 3
        Me.cbRecherche.Text = "Recherche"
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(328, 5)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 24)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "Afficher"
        '
        'frmEtatStock
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(952, 646)
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
End Class
