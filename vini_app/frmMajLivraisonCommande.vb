Imports vini_DB
Public Class frmMajLivraisonCommande
    Inherits frmDonBase
    Protected m_lstCommandes As List(Of CommandeClient)
    Friend WithEvents m_bsrcCommandeClient As System.Windows.Forms.BindingSource
    Friend WithEvents dgvCommandes As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents grpSCMD As System.Windows.Forms.GroupBox
    Friend WithEvents tbMontantTTCFactCourante As vini_app.textBoxCurrency
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbMontantHTFactCourante As vini_app.textBoxCurrency
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnSuivant As System.Windows.Forms.Button
    Friend WithEvents btnSauve As System.Windows.Forms.Button
    Friend WithEvents tbRefFact As System.Windows.Forms.TextBox
    Friend WithEvents tbLettreVoiture As System.Windows.Forms.TextBox
    Friend WithEvents tbRefLivraison As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblRSClient As System.Windows.Forms.Label
    Friend WithEvents lblCode As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents tbCoutTransport As vini_app.textBoxCurrency
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dateCommande As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalTTC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EtatCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Protected m_colFact As ColEvent
    'Protected getElementCourant() As FactTRP

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
    Friend WithEvents dtdateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtDatedeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents tbRSClient As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.m_bsrcCommandeClient = New System.Windows.Forms.BindingSource(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.dgvCommandes = New System.Windows.Forms.DataGridView()
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dateCommande = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TiersRS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.totalTTC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TiersCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EtatCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cbRecherche = New System.Windows.Forms.Button()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.tbRSClient = New System.Windows.Forms.TextBox()
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.grpSCMD = New System.Windows.Forms.GroupBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.tbCoutTransport = New vini_app.textBoxCurrency()
        Me.btnSuivant = New System.Windows.Forms.Button()
        Me.btnSauve = New System.Windows.Forms.Button()
        Me.tbRefFact = New System.Windows.Forms.TextBox()
        Me.tbLettreVoiture = New System.Windows.Forms.TextBox()
        Me.tbRefLivraison = New System.Windows.Forms.TextBox()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.lblRSClient = New System.Windows.Forms.Label()
        Me.lblCode = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbMontantTTCFactCourante = New vini_app.textBoxCurrency()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbMontantHTFactCourante = New vini_app.textBoxCurrency()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.m_bsrcCommandeClient, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvCommandes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpSCMD.SuspendLayout()
        Me.SuspendLayout()
        '
        'm_bsrcCommandeClient
        '
        Me.m_bsrcCommandeClient.DataSource = GetType(vini_DB.CommandeClient)
        Me.m_bsrcCommandeClient.Filter = ""
        Me.m_bsrcCommandeClient.Sort = "dateCommande"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvCommandes)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cbRecherche)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cbAfficher)
        Me.SplitContainer1.Panel1.Controls.Add(Me.tbRSClient)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtdateFin)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtDatedeb)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.grpSCMD)
        Me.SplitContainer1.Size = New System.Drawing.Size(998, 633)
        Me.SplitContainer1.SplitterDistance = 332
        Me.SplitContainer1.TabIndex = 129
        '
        'dgvCommandes
        '
        Me.dgvCommandes.AllowUserToAddRows = False
        Me.dgvCommandes.AllowUserToDeleteRows = False
        Me.dgvCommandes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCommandes.AutoGenerateColumns = False
        Me.dgvCommandes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCommandes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCommandes.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.dateCommande, Me.TiersRS, Me.DataGridViewTextBoxColumn2, Me.totalTTC, Me.TiersCode, Me.EtatCode})
        Me.dgvCommandes.DataSource = Me.m_bsrcCommandeClient
        Me.dgvCommandes.Location = New System.Drawing.Point(5, 118)
        Me.dgvCommandes.Name = "dgvCommandes"
        Me.dgvCommandes.ReadOnly = True
        Me.dgvCommandes.RowHeadersVisible = False
        Me.dgvCommandes.RowHeadersWidth = 20
        Me.dgvCommandes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCommandes.Size = New System.Drawing.Size(322, 510)
        Me.dgvCommandes.TabIndex = 127
        Me.ToolTip1.SetToolTip(Me.dgvCommandes, "Liste des Commandes validées")
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "CMD"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn.Width = 56
        '
        'dateCommande
        '
        Me.dateCommande.DataPropertyName = "dateCommande"
        DataGridViewCellStyle4.Format = "d"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.dateCommande.DefaultCellStyle = DataGridViewCellStyle4
        Me.dateCommande.HeaderText = "date"
        Me.dateCommande.Name = "dateCommande"
        Me.dateCommande.ReadOnly = True
        Me.dateCommande.Width = 53
        '
        'TiersRS
        '
        Me.TiersRS.DataPropertyName = "TiersRS"
        Me.TiersRS.HeaderText = "Client"
        Me.TiersRS.Name = "TiersRS"
        Me.TiersRS.ReadOnly = True
        Me.TiersRS.Width = 58
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "totalHT"
        DataGridViewCellStyle5.Format = "C2"
        Me.DataGridViewTextBoxColumn2.DefaultCellStyle = DataGridViewCellStyle5
        Me.DataGridViewTextBoxColumn2.HeaderText = "HT"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 47
        '
        'totalTTC
        '
        Me.totalTTC.DataPropertyName = "totalTTC"
        DataGridViewCellStyle6.Format = "C2"
        DataGridViewCellStyle6.NullValue = Nothing
        Me.totalTTC.DefaultCellStyle = DataGridViewCellStyle6
        Me.totalTTC.HeaderText = "TTC"
        Me.totalTTC.Name = "totalTTC"
        Me.totalTTC.ReadOnly = True
        Me.totalTTC.Width = 53
        '
        'TiersCode
        '
        Me.TiersCode.DataPropertyName = "TiersCode"
        Me.TiersCode.HeaderText = "Code CLT"
        Me.TiersCode.Name = "TiersCode"
        Me.TiersCode.ReadOnly = True
        Me.TiersCode.Width = 80
        '
        'EtatCode
        '
        Me.EtatCode.DataPropertyName = "EtatCode"
        Me.EtatCode.HeaderText = "EtatCode"
        Me.EtatCode.Name = "EtatCode"
        Me.EtatCode.ReadOnly = True
        Me.EtatCode.Width = 76
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(160, 50)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(75, 23)
        Me.cbRecherche.TabIndex = 3
        Me.cbRecherche.Text = "Recherche"
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.BackColor = System.Drawing.SystemColors.Control
        Me.cbAfficher.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAfficher.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAfficher.Location = New System.Drawing.Point(3, 79)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAfficher.Size = New System.Drawing.Size(326, 23)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "A&fficher"
        Me.ToolTip1.SetToolTip(Me.cbAfficher, "Afficher la liste des Commandes validées")
        Me.cbAfficher.UseVisualStyleBackColor = False
        '
        'tbRSClient
        '
        Me.tbRSClient.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbRSClient.Location = New System.Drawing.Point(241, 51)
        Me.tbRSClient.Name = "tbRSClient"
        Me.tbRSClient.Size = New System.Drawing.Size(86, 20)
        Me.tbRSClient.TabIndex = 2
        '
        'dtdateFin
        '
        Me.dtdateFin.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtdateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdateFin.Location = New System.Drawing.Point(240, 29)
        Me.dtdateFin.Name = "dtdateFin"
        Me.dtdateFin.Size = New System.Drawing.Size(86, 20)
        Me.dtdateFin.TabIndex = 1
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(4, 9)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(176, 16)
        Me.Label14.TabIndex = 106
        Me.Label14.Text = "date de début de Commande"
        '
        'dtDatedeb
        '
        Me.dtDatedeb.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtDatedeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDatedeb.Location = New System.Drawing.Point(240, 9)
        Me.dtDatedeb.Name = "dtDatedeb"
        Me.dtDatedeb.Size = New System.Drawing.Size(86, 20)
        Me.dtDatedeb.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(4, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 108
        Me.Label1.Text = "RS Client"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(4, 33)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(176, 16)
        Me.Label8.TabIndex = 107
        Me.Label8.Text = "date de fin de Commande"
        '
        'grpSCMD
        '
        Me.grpSCMD.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpSCMD.Controls.Add(Me.Label11)
        Me.grpSCMD.Controls.Add(Me.tbCoutTransport)
        Me.grpSCMD.Controls.Add(Me.btnSuivant)
        Me.grpSCMD.Controls.Add(Me.btnSauve)
        Me.grpSCMD.Controls.Add(Me.tbRefFact)
        Me.grpSCMD.Controls.Add(Me.tbLettreVoiture)
        Me.grpSCMD.Controls.Add(Me.tbRefLivraison)
        Me.grpSCMD.Controls.Add(Me.DateTimePicker1)
        Me.grpSCMD.Controls.Add(Me.lblRSClient)
        Me.grpSCMD.Controls.Add(Me.lblCode)
        Me.grpSCMD.Controls.Add(Me.Label10)
        Me.grpSCMD.Controls.Add(Me.Label9)
        Me.grpSCMD.Controls.Add(Me.Label7)
        Me.grpSCMD.Controls.Add(Me.Label6)
        Me.grpSCMD.Controls.Add(Me.Label5)
        Me.grpSCMD.Controls.Add(Me.Label4)
        Me.grpSCMD.Controls.Add(Me.tbMontantTTCFactCourante)
        Me.grpSCMD.Controls.Add(Me.Label3)
        Me.grpSCMD.Controls.Add(Me.tbMontantHTFactCourante)
        Me.grpSCMD.Controls.Add(Me.Label2)
        Me.grpSCMD.Location = New System.Drawing.Point(3, 9)
        Me.grpSCMD.Name = "grpSCMD"
        Me.grpSCMD.Size = New System.Drawing.Size(645, 611)
        Me.grpSCMD.TabIndex = 0
        Me.grpSCMD.TabStop = False
        '
        'Label11
        '
        Me.Label11.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(13, 170)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(83, 13)
        Me.Label11.TabIndex = 142
        Me.Label11.Text = "Coût Transport :"
        '
        'tbCoutTransport
        '
        Me.tbCoutTransport.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "coutTransport", True))
        Me.tbCoutTransport.Location = New System.Drawing.Point(140, 167)
        Me.tbCoutTransport.Name = "tbCoutTransport"
        Me.tbCoutTransport.Size = New System.Drawing.Size(96, 20)
        Me.tbCoutTransport.TabIndex = 3
        Me.tbCoutTransport.Text = "0"
        '
        'btnSuivant
        '
        Me.btnSuivant.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSuivant.Location = New System.Drawing.Point(534, 550)
        Me.btnSuivant.Name = "btnSuivant"
        Me.btnSuivant.Size = New System.Drawing.Size(75, 23)
        Me.btnSuivant.TabIndex = 9
        Me.btnSuivant.Text = "Suivant"
        Me.btnSuivant.UseVisualStyleBackColor = True
        '
        'btnSauve
        '
        Me.btnSauve.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSauve.Location = New System.Drawing.Point(417, 551)
        Me.btnSauve.Name = "btnSauve"
        Me.btnSauve.Size = New System.Drawing.Size(110, 23)
        Me.btnSauve.TabIndex = 8
        Me.btnSauve.Text = "Sauver && Suivant"
        Me.btnSauve.UseVisualStyleBackColor = True
        '
        'tbRefFact
        '
        Me.tbRefFact.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbRefFact.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "refFactTRP", True))
        Me.tbRefFact.Location = New System.Drawing.Point(140, 193)
        Me.tbRefFact.Name = "tbRefFact"
        Me.tbRefFact.Size = New System.Drawing.Size(499, 20)
        Me.tbRefFact.TabIndex = 4
        '
        'tbLettreVoiture
        '
        Me.tbLettreVoiture.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbLettreVoiture.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "lettreVoiture", True))
        Me.tbLettreVoiture.Location = New System.Drawing.Point(140, 141)
        Me.tbLettreVoiture.Name = "tbLettreVoiture"
        Me.tbLettreVoiture.Size = New System.Drawing.Size(499, 20)
        Me.tbLettreVoiture.TabIndex = 2
        '
        'tbRefLivraison
        '
        Me.tbRefLivraison.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbRefLivraison.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "refLivraison", True))
        Me.tbRefLivraison.Location = New System.Drawing.Point(140, 109)
        Me.tbRefLivraison.Name = "tbRefLivraison"
        Me.tbRefLivraison.Size = New System.Drawing.Size(499, 20)
        Me.tbRefLivraison.TabIndex = 1
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "dateLivraison", True))
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(137, 77)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(96, 20)
        Me.DateTimePicker1.TabIndex = 0
        '
        'lblRSClient
        '
        Me.lblRSClient.AutoSize = True
        Me.lblRSClient.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "TiersRS", True))
        Me.lblRSClient.Location = New System.Drawing.Point(55, 52)
        Me.lblRSClient.Name = "lblRSClient"
        Me.lblRSClient.Size = New System.Drawing.Size(48, 13)
        Me.lblRSClient.TabIndex = 1
        Me.lblRSClient.Text = "RSClient"
        '
        'lblCode
        '
        Me.lblCode.AutoSize = True
        Me.lblCode.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "code", True))
        Me.lblCode.Location = New System.Drawing.Point(52, 20)
        Me.lblCode.Name = "lblCode"
        Me.lblCode.Size = New System.Drawing.Size(88, 13)
        Me.lblCode.TabIndex = 0
        Me.lblCode.Text = "Code Commande"
        '
        'Label10
        '
        Me.Label10.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(13, 196)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(117, 13)
        Me.Label10.TabIndex = 140
        Me.Label10.Text = "Ref Fact Transporteur :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(13, 148)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(76, 13)
        Me.Label9.TabIndex = 139
        Me.Label9.Text = "Lettre Voiture :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(13, 116)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(123, 13)
        Me.Label7.TabIndex = 138
        Me.Label7.Text = "Reférence de Livraison :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(13, 84)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 13)
        Me.Label6.TabIndex = 137
        Me.Label6.Text = "Date de Livraison :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 52)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 136
        Me.Label5.Text = "Client :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(38, 13)
        Me.Label4.TabIndex = 135
        Me.Label4.Text = "Code :"
        '
        'tbMontantTTCFactCourante
        '
        Me.tbMontantTTCFactCourante.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMontantTTCFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "totalTTC", True))
        Me.tbMontantTTCFactCourante.Location = New System.Drawing.Point(543, 247)
        Me.tbMontantTTCFactCourante.Name = "tbMontantTTCFactCourante"
        Me.tbMontantTTCFactCourante.ReadOnly = True
        Me.tbMontantTTCFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantTTCFactCourante.TabIndex = 7
        Me.tbMontantTTCFactCourante.Text = "0"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(439, 250)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 17)
        Me.Label3.TabIndex = 134
        Me.Label3.Text = "Montant TTC"
        '
        'tbMontantHTFactCourante
        '
        Me.tbMontantHTFactCourante.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMontantHTFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommandeClient, "totalHT", True))
        Me.tbMontantHTFactCourante.Location = New System.Drawing.Point(337, 247)
        Me.tbMontantHTFactCourante.Name = "tbMontantHTFactCourante"
        Me.tbMontantHTFactCourante.ReadOnly = True
        Me.tbMontantHTFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantHTFactCourante.TabIndex = 6
        Me.tbMontantHTFactCourante.Text = "0"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(231, 250)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 132
        Me.Label2.Text = "Montant HT"
        '
        'frmMajLivraisonCommande
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(998, 633)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "frmMajLivraisonCommande"
        Me.Text = "Mise à jour des informations de livraison"
        CType(Me.m_bsrcCommandeClient, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvCommandes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpSCMD.ResumeLayout(False)
        Me.grpSCMD.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Méthodes Vinicom"
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False
        m_ToolBarSaveEnabled = isfrmUpdated()
    End Sub

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)
        SplitContainer1.Enabled = True
        SplitContainer1.Panel1.Enabled = True
        SplitContainer1.Panel2.Enabled = True
        dtDatedeb.Enabled = True
        dtdateFin.Enabled = True
        tbRSClient.Enabled = True
        cbAfficher.Enabled = True
        cbRecherche.Enabled = True
        grpSCMD.Enabled = True

    End Sub

    Protected Overrides Function creerElement() As Boolean
        Return False
    End Function
    Protected Shadows Function getElementCourant() As FactTRP
        Return Nothing
    End Function
    ''' <summary>
    ''' Sauvegarde des commandes et des sous commandes modifiées
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function frmSave() As Boolean
        Debug.Assert(m_lstCommandes IsNot Nothing)
        Debug.Assert(m_lstCommandes.Count > 0)
        'Sauvegarde des commandes et des sousCommandes

        For Each oCmd As CommandeClient In m_lstCommandes
            If oCmd.bUpdated Then
                oCmd.bUpdatePrecommande = False 'Pas de Mise à jour des précommandes
                oCmd.save()
            End If
        Next
        setfrmNotUpdated()
        'm_bsrcScmd.Clear()

    End Function

#End Region

#Region "Methodes privées"
    Private Sub initFenetre()
        Dim MoisSuivant As Date
        Dim premierMoisSuivant As Date
        Dim dernierMoisCourant As Date
        'Date de Début = 01 du mois Courant
        dtDatedeb.Value = "01/" & Now.Month() & "/" & Now.Year
        MoisSuivant = DateAdd(DateInterval.Month, +1, Now())
        premierMoisSuivant = "01/" & MoisSuivant.Month() & "/" & MoisSuivant.Year()
        dernierMoisCourant = DateAdd(DateInterval.Day, -1, premierMoisSuivant)

        dtdateFin.Value = dernierMoisCourant
        'dtDateFacture.Value = Now()
        m_lstCommandes = New List(Of CommandeClient)

    End Sub
    Private Sub afficheListeCommande()

        setcursorWait()
        debAffiche()
        m_bsrcCommandeClient.Clear()

        Debug.Assert(Not m_lstCommandes Is Nothing)
        Dim obj As CommandeClient

        For Each obj In m_lstCommandes
            m_bsrcCommandeClient.Add(obj)
        Next obj


        finAffiche()
        restoreCursor()

    End Sub 'AfficheListeCommande

    ''' <summary>
    ''' Récupération de la liste des commandes
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getListeCommandes() As Boolean
        Dim ddeb As Date
        Dim dfin As Date
        Dim RSClient As String
        Dim col As Collection
        Dim bReturn As Boolean
        debAffiche()
        setcursorWait()
        Try

            ddeb = dtDatedeb.Value.ToShortDateString
            dfin = dtdateFin.Value.ToShortDateString
            RSClient = tbRSClient.Text
            'Recupération de la liste des commandes Validées sans charger les lignes de commandes
            CommandeClient.bChargerColLignes = False
            col = CommandeClient.getListe(ddeb, dfin, RSClient, vncEtatCommande.vncEclatee)
            If col Is Nothing Then
                bReturn = False
            Else
                m_lstCommandes.Clear()
                For Each oCmd As CommandeClient In col
                    m_lstCommandes.Add(oCmd)
                Next
                bReturn = True
            End If
            CommandeClient.bChargerColLignes = True
        Catch ex As Exception
            bReturn = False
            Debug.Assert(bReturn, ex.ToString)
        End Try
        finAffiche()
        restoreCursor()
        Return bReturn
    End Function
    Private Function sauver() As Boolean
        Dim bReturn As Boolean
        Try

            Dim oCmd As CommandeClient
            If m_bsrcCommandeClient.Current IsNot Nothing Then
                oCmd = CType(m_bsrcCommandeClient.Current, CommandeClient)
                If (oCmd.save()) Then
                    bReturn = True
                    m_bsrcCommandeClient.MoveNext()
                End If
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Function suivant() As Boolean
        Dim bReturn As Boolean
        Try

            m_bsrcCommandeClient.MoveNext()
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Function elementCourantSansModif() As Boolean
        Return True
    End Function
    Private Sub rechercheClient()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.CLIENT, tbRSClient)

        If Not objTiers Is Nothing Then
            tbRSClient.Text = objTiers.code
        End If
    End Sub 'rechercheClient

#End Region
#Region "Gestion des Evenements"
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click

        If getListeCommandes() Then
            afficheListeCommande()
        End If
        dgvCommandes.Enabled = True

    End Sub





#End Region

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Génerer déclenche L'evenement Updated
        'AddHandler cbAppliquer.Click, AddressOf ControlUpdated
        'AddHandler cbGenerer.Click, AddressOf ControlUpdated
    End Sub

    'Private Sub cbSelectionnerTout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSelectionnerTout.Click
    '    Call selectionnerTout()
    'End Sub


    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        rechercheClient()
    End Sub



    Private Sub btnSauve_Click(sender As Object, e As EventArgs) Handles btnSauve.Click
        sauver()
    End Sub

    Private Sub btnSuivant_Click(sender As Object, e As EventArgs) Handles btnSuivant.Click
        Suivant()
    End Sub
End Class
