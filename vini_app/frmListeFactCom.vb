Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmListeFactCom
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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbxEtat As System.Windows.Forms.ComboBox
    Friend WithEvents m_bsrcEtat As System.Windows.Forms.BindingSource
    Friend WithEvents tbCodeClient As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents rbTrieParNumero As System.Windows.Forms.RadioButton
    Friend WithEvents rbtriCodeFounisseur As System.Windows.Forms.RadioButton
    Friend WithEvents dtFin As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtdeb = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtFin = New System.Windows.Forms.DateTimePicker()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbxEtat = New System.Windows.Forms.ComboBox()
        Me.m_bsrcEtat = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbCodeClient = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.rbTrieParNumero = New System.Windows.Forms.RadioButton()
        Me.rbtriCodeFounisseur = New System.Windows.Forms.RadioButton()
        CType(Me.m_bsrcEtat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Date de début"
        '
        'dtdeb
        '
        Me.dtdeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdeb.Location = New System.Drawing.Point(128, 8)
        Me.dtdeb.Name = "dtdeb"
        Me.dtdeb.Size = New System.Drawing.Size(136, 20)
        Me.dtdeb.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(272, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Date de fin"
        '
        'dtFin
        '
        Me.dtFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFin.Location = New System.Drawing.Point(360, 8)
        Me.dtFin.Name = "dtFin"
        Me.dtFin.Size = New System.Drawing.Size(104, 20)
        Me.dtFin.TabIndex = 3
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.Location = New System.Drawing.Point(868, 9)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 23)
        Me.cbAfficher.TabIndex = 5
        Me.cbAfficher.Text = "Afficher"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(477, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(26, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Etat"
        '
        'cbxEtat
        '
        Me.cbxEtat.DataSource = Me.m_bsrcEtat
        Me.cbxEtat.DisplayMember = "libelle"
        Me.cbxEtat.FormattingEnabled = True
        Me.cbxEtat.Location = New System.Drawing.Point(606, 34)
        Me.cbxEtat.Name = "cbxEtat"
        Me.cbxEtat.Size = New System.Drawing.Size(121, 21)
        Me.cbxEtat.TabIndex = 6
        Me.cbxEtat.ValueMember = "codeEtat"
        '
        'm_bsrcEtat
        '
        Me.m_bsrcEtat.DataSource = GetType(vini_DB.EtatCommande)
        '
        'tbCodeClient
        '
        Me.tbCodeClient.Location = New System.Drawing.Point(607, 8)
        Me.tbCodeClient.Name = "tbCodeClient"
        Me.tbCodeClient.Size = New System.Drawing.Size(120, 20)
        Me.tbCodeClient.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(474, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(103, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Code Fournisseur :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 42)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Triée par :"
        '
        'rbTrieParNumero
        '
        Me.rbTrieParNumero.AutoSize = True
        Me.rbTrieParNumero.Location = New System.Drawing.Point(72, 38)
        Me.rbTrieParNumero.Name = "rbTrieParNumero"
        Me.rbTrieParNumero.Size = New System.Drawing.Size(113, 17)
        Me.rbTrieParNumero.TabIndex = 7
        Me.rbTrieParNumero.TabStop = True
        Me.rbTrieParNumero.Text = "Numéro de facture"
        Me.rbTrieParNumero.UseVisualStyleBackColor = True
        '
        'rbtriCodeFounisseur
        '
        Me.rbtriCodeFounisseur.AutoSize = True
        Me.rbtriCodeFounisseur.Checked = True
        Me.rbtriCodeFounisseur.Location = New System.Drawing.Point(191, 38)
        Me.rbtriCodeFounisseur.Name = "rbtriCodeFounisseur"
        Me.rbtriCodeFounisseur.Size = New System.Drawing.Size(104, 17)
        Me.rbtriCodeFounisseur.TabIndex = 8
        Me.rbtriCodeFounisseur.TabStop = True
        Me.rbtriCodeFounisseur.Text = "Code fournisseur"
        Me.rbtriCodeFounisseur.UseVisualStyleBackColor = True
        '
        'frmListeFactCom
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1000, 678)
        Me.Controls.Add(Me.rbtriCodeFounisseur)
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
        Me.Name = "frmListeFactCom"
        Me.Text = "Liste des factures de commission"
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
        Me.Controls.SetChildIndex(Me.rbtriCodeFounisseur, 0)
        CType(Me.m_bsrcEtat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click

        Dim objReport As ReportDocument
        Dim str As String

        objReport = New ReportDocument
        If rbtriCodeFounisseur.Checked Then
            objReport.Load(PATHTOREPORTS & "crListeFactcom.rpt")
        Else
            objReport.Load(PATHTOREPORTS & "crListeFactcomParNumero.rpt")
        End If



        objReport.SetParameterValue("ddeb", Me.dtdeb.Value)
        objReport.SetParameterValue("dfin", Me.dtFin.Value)
        str = tbCodeClient.Text + "%"
        str = Replace(str, "%", "*")
        objReport.SetParameterValue("CODEFRN", str)
        objReport.SetParameterValue("codeEtat", cbxEtat.SelectedValue)

        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport
    End Sub

    Private Sub frmListeFactCom_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        m_bsrcEtat.Add(New EtatFactCommGeneree())
        m_bsrcEtat.Add(New EtatFactCommExportee())
        m_bsrcEtat.Add(New EtatTous())
        cbxEtat.SelectedIndex = cbxEtat.Items.Count - 1
        tbCodeClient.Text = "%"

    End Sub
End Class
