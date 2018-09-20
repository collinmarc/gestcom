Imports vini_DB
Public Class frmEclatementCommande
    Inherits frmDonBase
    Protected m_lstCommandes As List(Of CommandeClient)
    Friend WithEvents m_bsrcCommandeClient As System.Windows.Forms.BindingSource
    Friend WithEvents dgvCommandes As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcScmd As System.Windows.Forms.BindingSource
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cbEclatement As System.Windows.Forms.Button
    Friend WithEvents grpSCMD As System.Windows.Forms.GroupBox
    Friend WithEvents tbMontantTTCFactCourante As vini_app.textBoxCurrency
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbMontantHTFactCourante As vini_app.textBoxCurrency
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents liFournisseur As System.Windows.Forms.LinkLabel
    Friend WithEvents liScmd As System.Windows.Forms.LinkLabel
    Friend WithEvents dgvSousCommandes As System.Windows.Forms.DataGridView
    Friend WithEvents LiClient As System.Windows.Forms.LinkLabel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents dbSelectAll As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FournisseurRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalHT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Selected As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dateCommande As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalTTC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents qteColis As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents tbCodeClient As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.m_bsrcScmd = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_bsrcCommandeClient = New System.Windows.Forms.BindingSource(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.dbSelectAll = New System.Windows.Forms.Button()
        Me.dgvCommandes = New System.Windows.Forms.DataGridView()
        Me.cbRecherche = New System.Windows.Forms.Button()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.tbCodeClient = New System.Windows.Forms.TextBox()
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cbEclatement = New System.Windows.Forms.Button()
        Me.grpSCMD = New System.Windows.Forms.GroupBox()
        Me.LiClient = New System.Windows.Forms.LinkLabel()
        Me.tbMontantTTCFactCourante = New vini_app.textBoxCurrency()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbMontantHTFactCourante = New vini_app.textBoxCurrency()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.liFournisseur = New System.Windows.Forms.LinkLabel()
        Me.liScmd = New System.Windows.Forms.LinkLabel()
        Me.dgvSousCommandes = New System.Windows.Forms.DataGridView()
        Me.code = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FournisseurRS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.totalHT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Selected = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dateCommande = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TiersRS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.totalTTC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.qteColis = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.m_bsrcScmd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcCommandeClient, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvCommandes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpSCMD.SuspendLayout()
        CType(Me.dgvSousCommandes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'm_bsrcScmd
        '
        Me.m_bsrcScmd.DataSource = GetType(vini_DB.SousCommande)
        '
        'm_bsrcCommandeClient
        '
        Me.m_bsrcCommandeClient.DataSource = GetType(vini_DB.CommandeClient)
        Me.m_bsrcCommandeClient.Filter = "EtatCode=6"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.dbSelectAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvCommandes)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cbRecherche)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cbAfficher)
        Me.SplitContainer1.Panel1.Controls.Add(Me.tbCodeClient)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtdateFin)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtDatedeb)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.cbEclatement)
        Me.SplitContainer1.Panel2.Controls.Add(Me.grpSCMD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvSousCommandes)
        Me.SplitContainer1.Size = New System.Drawing.Size(998, 633)
        Me.SplitContainer1.SplitterDistance = 332
        Me.SplitContainer1.TabIndex = 129
        '
        'dbSelectAll
        '
        Me.dbSelectAll.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dbSelectAll.Location = New System.Drawing.Point(5, 607)
        Me.dbSelectAll.Name = "dbSelectAll"
        Me.dbSelectAll.Size = New System.Drawing.Size(316, 21)
        Me.dbSelectAll.TabIndex = 128
        Me.dbSelectAll.Text = "Sélectionner Tout"
        Me.dbSelectAll.UseVisualStyleBackColor = True
        '
        'dgvCommandes
        '
        Me.dgvCommandes.AllowUserToAddRows = False
        Me.dgvCommandes.AllowUserToDeleteRows = False
        Me.dgvCommandes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCommandes.AutoGenerateColumns = False
        Me.dgvCommandes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvCommandes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCommandes.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Selected, Me.CodeDataGridViewTextBoxColumn, Me.dateCommande, Me.TiersRS, Me.DataGridViewTextBoxColumn2, Me.totalTTC, Me.qteColis})
        Me.dgvCommandes.DataSource = Me.m_bsrcCommandeClient
        Me.dgvCommandes.Location = New System.Drawing.Point(5, 118)
        Me.dgvCommandes.Name = "dgvCommandes"
        Me.dgvCommandes.RowHeadersVisible = False
        Me.dgvCommandes.RowHeadersWidth = 20
        Me.dgvCommandes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCommandes.Size = New System.Drawing.Size(322, 483)
        Me.dgvCommandes.TabIndex = 127
        Me.ToolTip1.SetToolTip(Me.dgvCommandes, "Liste des Commandes validées")
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
        'tbCodeClient
        '
        Me.tbCodeClient.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbCodeClient.Location = New System.Drawing.Point(241, 51)
        Me.tbCodeClient.Name = "tbCodeClient"
        Me.tbCodeClient.Size = New System.Drawing.Size(86, 20)
        Me.tbCodeClient.TabIndex = 2
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
        Me.Label1.Text = "Client"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(4, 33)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(176, 16)
        Me.Label8.TabIndex = 107
        Me.Label8.Text = "date de fin de Commande"
        '
        'cbEclatement
        '
        Me.cbEclatement.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbEclatement.Location = New System.Drawing.Point(3, 27)
        Me.cbEclatement.Name = "cbEclatement"
        Me.cbEclatement.Size = New System.Drawing.Size(645, 24)
        Me.cbEclatement.TabIndex = 129
        Me.cbEclatement.Text = "Genération Bons à Facturer"
        '
        'grpSCMD
        '
        Me.grpSCMD.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpSCMD.Controls.Add(Me.LiClient)
        Me.grpSCMD.Controls.Add(Me.tbMontantTTCFactCourante)
        Me.grpSCMD.Controls.Add(Me.Label3)
        Me.grpSCMD.Controls.Add(Me.tbMontantHTFactCourante)
        Me.grpSCMD.Controls.Add(Me.Label2)
        Me.grpSCMD.Controls.Add(Me.liFournisseur)
        Me.grpSCMD.Controls.Add(Me.liScmd)
        Me.grpSCMD.Location = New System.Drawing.Point(3, 522)
        Me.grpSCMD.Name = "grpSCMD"
        Me.grpSCMD.Size = New System.Drawing.Size(645, 98)
        Me.grpSCMD.TabIndex = 15
        Me.grpSCMD.TabStop = False
        '
        'LiClient
        '
        Me.LiClient.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LiClient.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcScmd, "TiersRS", True))
        Me.LiClient.Location = New System.Drawing.Point(6, 56)
        Me.LiClient.Name = "LiClient"
        Me.LiClient.Size = New System.Drawing.Size(633, 16)
        Me.LiClient.TabIndex = 137
        Me.LiClient.TabStop = True
        Me.LiClient.Text = "RS Client"
        '
        'tbMontantTTCFactCourante
        '
        Me.tbMontantTTCFactCourante.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMontantTTCFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcScmd, "totalTTC", True))
        Me.tbMontantTTCFactCourante.Location = New System.Drawing.Point(543, 72)
        Me.tbMontantTTCFactCourante.Name = "tbMontantTTCFactCourante"
        Me.tbMontantTTCFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantTTCFactCourante.TabIndex = 6
        Me.tbMontantTTCFactCourante.Text = "0"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(439, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 17)
        Me.Label3.TabIndex = 134
        Me.Label3.Text = "Montant TTC"
        '
        'tbMontantHTFactCourante
        '
        Me.tbMontantHTFactCourante.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMontantHTFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcScmd, "totalHT", True))
        Me.tbMontantHTFactCourante.Location = New System.Drawing.Point(337, 72)
        Me.tbMontantHTFactCourante.Name = "tbMontantHTFactCourante"
        Me.tbMontantHTFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantHTFactCourante.TabIndex = 5
        Me.tbMontantHTFactCourante.Text = "0"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(231, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 132
        Me.Label2.Text = "Montant HT"
        '
        'liFournisseur
        '
        Me.liFournisseur.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.liFournisseur.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcScmd, "FournisseurRS", True))
        Me.liFournisseur.Location = New System.Drawing.Point(6, 40)
        Me.liFournisseur.Name = "liFournisseur"
        Me.liFournisseur.Size = New System.Drawing.Size(633, 16)
        Me.liFournisseur.TabIndex = 1
        Me.liFournisseur.TabStop = True
        Me.liFournisseur.Text = "RS Fournisseur"
        '
        'liScmd
        '
        Me.liScmd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.liScmd.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcScmd, "shortResume", True))
        Me.liScmd.DataBindings.Add(New System.Windows.Forms.Binding("Tag", Me.m_bsrcScmd, "id", True))
        Me.liScmd.Location = New System.Drawing.Point(6, 16)
        Me.liScmd.Name = "liScmd"
        Me.liScmd.Size = New System.Drawing.Size(639, 24)
        Me.liScmd.TabIndex = 0
        Me.liScmd.TabStop = True
        Me.liScmd.Text = "Reference Sous Commande"
        '
        'dgvSousCommandes
        '
        Me.dgvSousCommandes.AllowUserToAddRows = False
        Me.dgvSousCommandes.AllowUserToDeleteRows = False
        Me.dgvSousCommandes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSousCommandes.AutoGenerateColumns = False
        Me.dgvSousCommandes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvSousCommandes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSousCommandes.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.code, Me.FournisseurRS, Me.DataGridViewTextBoxColumn1, Me.totalHT, Me.DataGridViewTextBoxColumn3})
        Me.dgvSousCommandes.DataSource = Me.m_bsrcScmd
        Me.dgvSousCommandes.Location = New System.Drawing.Point(3, 57)
        Me.dgvSousCommandes.Name = "dgvSousCommandes"
        Me.dgvSousCommandes.ReadOnly = True
        Me.dgvSousCommandes.RowHeadersVisible = False
        Me.dgvSousCommandes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSousCommandes.Size = New System.Drawing.Size(645, 459)
        Me.dgvSousCommandes.TabIndex = 128
        '
        'code
        '
        Me.code.DataPropertyName = "code"
        Me.code.FillWeight = 10.0!
        Me.code.HeaderText = "code"
        Me.code.Name = "code"
        Me.code.ReadOnly = True
        '
        'FournisseurRS
        '
        Me.FournisseurRS.DataPropertyName = "FournisseurRS"
        Me.FournisseurRS.FillWeight = 33.0!
        Me.FournisseurRS.HeaderText = "Producteur"
        Me.FournisseurRS.Name = "FournisseurRS"
        Me.FournisseurRS.ReadOnly = True
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "TiersRS"
        Me.DataGridViewTextBoxColumn1.FillWeight = 33.0!
        Me.DataGridViewTextBoxColumn1.HeaderText = "Client"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'totalHT
        '
        Me.totalHT.DataPropertyName = "totalHT"
        DataGridViewCellStyle4.Format = "C2"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.totalHT.DefaultCellStyle = DataGridViewCellStyle4
        Me.totalHT.FillWeight = 10.0!
        Me.totalHT.HeaderText = "HT"
        Me.totalHT.Name = "totalHT"
        Me.totalHT.ReadOnly = True
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "totalTTC"
        DataGridViewCellStyle5.Format = "C2"
        DataGridViewCellStyle5.NullValue = Nothing
        Me.DataGridViewTextBoxColumn3.DefaultCellStyle = DataGridViewCellStyle5
        Me.DataGridViewTextBoxColumn3.FillWeight = 10.0!
        Me.DataGridViewTextBoxColumn3.HeaderText = "TTC"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.ReadOnly = True
        '
        'Selected
        '
        Me.Selected.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Selected.DataPropertyName = "Selected"
        Me.Selected.HeaderText = "Sel"
        Me.Selected.Name = "Selected"
        Me.Selected.Width = 20
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "CMD"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn.Width = 50
        '
        'dateCommande
        '
        Me.dateCommande.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.dateCommande.DataPropertyName = "dateCommande"
        DataGridViewCellStyle1.Format = "d"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.dateCommande.DefaultCellStyle = DataGridViewCellStyle1
        Me.dateCommande.HeaderText = "date"
        Me.dateCommande.Name = "dateCommande"
        Me.dateCommande.ReadOnly = True
        Me.dateCommande.Width = 70
        '
        'TiersRS
        '
        Me.TiersRS.DataPropertyName = "TiersRS"
        Me.TiersRS.HeaderText = "Client"
        Me.TiersRS.Name = "TiersRS"
        Me.TiersRS.ReadOnly = True
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "totalHT"
        DataGridViewCellStyle2.Format = "C2"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.DataGridViewTextBoxColumn2.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewTextBoxColumn2.HeaderText = "HT"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        '
        'totalTTC
        '
        Me.totalTTC.DataPropertyName = "totalTTC"
        DataGridViewCellStyle3.Format = "C2"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.totalTTC.DefaultCellStyle = DataGridViewCellStyle3
        Me.totalTTC.HeaderText = "TTC"
        Me.totalTTC.Name = "totalTTC"
        Me.totalTTC.ReadOnly = True
        '
        'qteColis
        '
        Me.qteColis.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.qteColis.DataPropertyName = "qteColis"
        Me.qteColis.HeaderText = "qteColis"
        Me.qteColis.Name = "qteColis"
        Me.qteColis.ReadOnly = True
        Me.qteColis.Width = 50
        '
        'frmEclatementCommande
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(998, 633)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "frmEclatementCommande"
        Me.Text = "Eclatement de commandes clients"
        CType(Me.m_bsrcScmd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcCommandeClient, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvCommandes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpSCMD.ResumeLayout(False)
        Me.grpSCMD.PerformLayout()
        CType(Me.dgvSousCommandes, System.ComponentModel.ISupportInitialize).EndInit()
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
        SplitContainer1.Panel2.Enabled = False
        dtDatedeb.Enabled = True
        dtdateFin.Enabled = True
        tbCodeClient.Enabled = True
        cbAfficher.Enabled = True
        '        dtDateFacture.Enabled = True
        cbRecherche.Enabled = True

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
        setcursorWait()
        For Each oCmd As CommandeClient In m_lstCommandes
            If oCmd.bUpdated Then
                oCmd.bUpdatePrecommande = False 'Pas de Mise à jour des précommandes
                oCmd.save()
            End If
        Next
        setfrmNotUpdated()
        'm_bsrcScmd.Clear()
        restoreCursor()
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

        'Trie sur la date de commande
        m_lstCommandes.Sort()
        For Each obj In m_lstCommandes
            obj.Selected = False
            m_bsrcCommandeClient.Add(obj)
        Next obj

        m_bsrcScmd.Clear()

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
        Dim codeClient As String
        Dim col As Collection
        Dim bReturn As Boolean
        debAffiche()
        setcursorWait()
        Try

            ddeb = dtDatedeb.Value.ToShortDateString
            dfin = dtdateFin.Value.ToShortDateString
            codeClient = tbCodeClient.Text
            'Recupération de la liste des commandes Livrées sans charger les lignes de commandes
            CommandeClient.bChargerColLignes = False
            col = CommandeClient.getListe(ddeb, dfin, codeClient, vncEtatCommande.vncLivree, "")
            If col Is Nothing Then
                bReturn = False
            Else
                m_lstCommandes.Clear()
                For Each oCmd As CommandeClient In col
                    m_lstCommandes.Add(oCmd)
                Next
                m_lstCommandes.Sort()
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
    Private Sub genererFactures()
    End Sub 'genererFactures

    Private Sub rechercheClient()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.CLIENT, tbCodeClient)

        If Not objTiers Is Nothing Then
            tbCodeClient.Text = objTiers.code
        End If
    End Sub 'rechercheClient

#End Region
#Region "Gestion des Evenements"
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        If m_bsrcScmd.Count > 0 And isfrmUpdated() Then
            If MsgBox("Vous avez déjà généré des Bons à Facturer, voulez-vous les sauvegarder?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                frmSave()
            End If
        End If

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

    Private Sub cbLivraison_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LivraisonEclatement()
    End Sub


    Private Sub liTiers_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liFournisseur.LinkClicked
        afficheFenetreClient(liFournisseur.Tag)
    End Sub


    Private Sub liFactTRP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liScmd.LinkClicked
        afficheFenetreFactTrp(liScmd.Tag)
    End Sub

    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        rechercheClient()
    End Sub



    Private Sub dbSelectAll_Click(sender As Object, e As EventArgs) Handles dbSelectAll.Click
        For Each oCmd As CommandeClient In m_bsrcCommandeClient
            oCmd.Selected = True
        Next
        m_bsrcCommandeClient.ResetBindings(False)
    End Sub

    Private Sub dgvCommandes_SelectionChanged(sender As Object, e As EventArgs) Handles dgvCommandes.SelectionChanged
        If dgvCommandes.SelectedRows.Count > 0 Then
            SplitContainer1.Panel2.Enabled = True
        Else
            If dgvSousCommandes.Rows.Count = 0 Then
                SplitContainer1.Panel2.Enabled = False
            Else
                SplitContainer1.Panel2.Enabled = True
            End If
        End If
    End Sub

    Private Sub cbEclatement_Click(sender As Object, e As EventArgs) Handles cbEclatement.Click
        setcursorWait()
        LivraisonEclatement()
        restoreCursor()
    End Sub
    ''' <summary>
    ''' Livraison des commandes sélectionnées
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LivraisonEclatement()
        Dim oCmd As CommandeClient
        For Each oCmd In m_bsrcCommandeClient
            If oCmd.Selected Then
                oCmd.load()
                If oCmd.EtatCode = vncEtatCommande.vncLivree Then
                    oCmd.refLivraison = oCmd.code

                    oCmd.generationSousCommande()
                    For Each oScmd As SousCommande In oCmd.colSousCommandes
                        m_bsrcScmd.Add(oScmd)
                    Next

                End If
            End If
            'Raffraichissement de la liste des commandes
        Next
        setfrmUpdated()
    End Sub

End Class
