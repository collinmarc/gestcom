Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmListeFactColisage
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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbxEtat As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents m_bsrcEtat As System.Windows.Forms.BindingSource
    Friend WithEvents rbtriEtatFacture As System.Windows.Forms.RadioButton
    Friend WithEvents rbTrieParNumero As System.Windows.Forms.RadioButton
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbCodeClient As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtdeb = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtFin = New System.Windows.Forms.DateTimePicker()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbCodeClient = New System.Windows.Forms.TextBox()
        Me.cbxEtat = New System.Windows.Forms.ComboBox()
        Me.m_bsrcEtat = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label4 = New System.Windows.Forms.Label()
        Me.rbtriEtatFacture = New System.Windows.Forms.RadioButton()
        Me.rbTrieParNumero = New System.Windows.Forms.RadioButton()
        Me.Label5 = New System.Windows.Forms.Label()
        CType(Me.m_bsrcEtat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(5, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Date de début :"
        '
        'dtdeb
        '
        Me.dtdeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdeb.Location = New System.Drawing.Point(94, 8)
        Me.dtdeb.Name = "dtdeb"
        Me.dtdeb.Size = New System.Drawing.Size(104, 20)
        Me.dtdeb.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(224, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Date de fin :"
        '
        'dtFin
        '
        Me.dtFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFin.Location = New System.Drawing.Point(318, 8)
        Me.dtFin.Name = "dtFin"
        Me.dtFin.Size = New System.Drawing.Size(120, 20)
        Me.dtFin.TabIndex = 4
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(864, 8)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 23)
        Me.cbAfficher.TabIndex = 5
        Me.cbAfficher.Text = "Afficher"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(451, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(103, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Code Fournisseur :"
        '
        'tbCodeClient
        '
        Me.tbCodeClient.Location = New System.Drawing.Point(584, 8)
        Me.tbCodeClient.Name = "tbCodeClient"
        Me.tbCodeClient.Size = New System.Drawing.Size(120, 20)
        Me.tbCodeClient.TabIndex = 7
        '
        'cbxEtat
        '
        Me.cbxEtat.DataSource = Me.m_bsrcEtat
        Me.cbxEtat.DisplayMember = "libelle"
        Me.cbxEtat.FormattingEnabled = True
        Me.cbxEtat.Location = New System.Drawing.Point(583, 34)
        Me.cbxEtat.Name = "cbxEtat"
        Me.cbxEtat.Size = New System.Drawing.Size(121, 21)
        Me.cbxEtat.TabIndex = 8
        Me.cbxEtat.ValueMember = "codeEtat"
        '
        'm_bsrcEtat
        '
        Me.m_bsrcEtat.DataSource = GetType(vini_DB.EtatCommande)
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(454, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(26, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Etat"
        '
        'rbtriEtatFacture
        '
        Me.rbtriEtatFacture.AutoSize = True
        Me.rbtriEtatFacture.Checked = True
        Me.rbtriEtatFacture.Location = New System.Drawing.Point(191, 32)
        Me.rbtriEtatFacture.Name = "rbtriEtatFacture"
        Me.rbtriEtatFacture.Size = New System.Drawing.Size(83, 17)
        Me.rbtriEtatFacture.TabIndex = 16
        Me.rbtriEtatFacture.TabStop = True
        Me.rbtriEtatFacture.Text = "Etat Facture"
        Me.rbtriEtatFacture.UseVisualStyleBackColor = True
        '
        'rbTrieParNumero
        '
        Me.rbTrieParNumero.AutoSize = True
        Me.rbTrieParNumero.Location = New System.Drawing.Point(72, 32)
        Me.rbTrieParNumero.Name = "rbTrieParNumero"
        Me.rbTrieParNumero.Size = New System.Drawing.Size(113, 17)
        Me.rbTrieParNumero.TabIndex = 15
        Me.rbTrieParNumero.TabStop = True
        Me.rbTrieParNumero.Text = "Numéro de facture"
        Me.rbTrieParNumero.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 36)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Triée par :"
        '
        'frmListeFactColisage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1000, 678)
        Me.Controls.Add(Me.rbtriEtatFacture)
        Me.Controls.Add(Me.rbTrieParNumero)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbxEtat)
        Me.Controls.Add(Me.tbCodeClient)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.dtFin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtdeb)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmListeFactColisage"
        Me.Text = "Liste des Factures de colisage"
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtdeb, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.dtFin, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.tbCodeClient, 0)
        Me.Controls.SetChildIndex(Me.cbxEtat, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.rbTrieParNumero, 0)
        Me.Controls.SetChildIndex(Me.rbtriEtatFacture, 0)
        CType(Me.m_bsrcEtat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Overrides Function getResume() As String
        Return "Liste des factures de Colisage"
    End Function

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        Dim str As String

        Dim objReport As New ReportDocument()
        If rbtriEtatFacture.Checked Then
            Using r1 As New crListeFactColisage()
                Dim strReportName As String = r1.ResourceName
                objReport.Load(PATHTOREPORTS & strReportName)
            End Using
        Else
            Using r1 As New crListeFactColisageParNumero()
                Dim strReportName As String = r1.ResourceName
                objReport.Load(PATHTOREPORTS & strReportName)
            End Using
        End If


        objReport.SetParameterValue("ddeb", Me.dtdeb.Value.ToShortDateString())
        objReport.SetParameterValue("dfin", Me.dtFin.Value.ToShortDateString())
        objReport.SetParameterValue("codeEtat", Me.cbxEtat.SelectedValue)

        str = tbCodeClient.Text + "%"
        str = Replace(str, "%", "*")
        objReport.SetParameterValue("CODEFRN", Trim(str))

        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport
    End Sub

    Private Sub frmListeFactColisage_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        m_bsrcEtat.Add(New EtatFactColGeneree())
        m_bsrcEtat.Add(New EtatFactcolExportee())
        m_bsrcEtat.Add(New EtatTous())
        cbxEtat.SelectedIndex = cbxEtat.Items.Count - 1
        tbCodeClient.Text = "%"

    End Sub
End Class
