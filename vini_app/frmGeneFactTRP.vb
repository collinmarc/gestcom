Imports vini_DB
Public Class frmGeneFactTRP
    Inherits frmDonBase
    Protected m_colCommandes As Collection
    Friend WithEvents m_bsrcCommandeClient As System.Windows.Forms.BindingSource
    Friend WithEvents dgvCommandes As System.Windows.Forms.DataGridView
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dateLivraison As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents montantTransport As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents m_bsrcFactTRP As System.Windows.Forms.BindingSource
    Friend WithEvents dgvFactTRP As System.Windows.Forms.DataGridView
    Friend WithEvents TiersRSDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dateFacture As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalHT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents cbSelectionnerTout As System.Windows.Forms.Button
    Friend WithEvents cbGenerer As System.Windows.Forms.Button
    Friend WithEvents cbAnnGenerer As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtDateFacture As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbPeriode As System.Windows.Forms.TextBox
    Friend WithEvents grpFact As System.Windows.Forms.GroupBox
    Friend WithEvents dtDateFactCourante As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbAppliquer As System.Windows.Forms.Button
    Friend WithEvents tbPeriodeFactCourante As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents tbMontantTTCFactCourante As textBoxCurrency
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbMontantHTFactCourante As textBoxCurrency
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents cbSauvegarder As System.Windows.Forms.Button
    Friend WithEvents tbCodeClient As System.Windows.Forms.TextBox
    Friend WithEvents liFacture As System.Windows.Forms.LinkLabel
    Friend WithEvents liTiers As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.tbCodeClient = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.cbGenerer = New System.Windows.Forms.Button
        Me.cbSelectionnerTout = New System.Windows.Forms.Button
        Me.cbAnnGenerer = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtDateFacture = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.tbPeriode = New System.Windows.Forms.TextBox
        Me.grpFact = New System.Windows.Forms.GroupBox
        Me.dtDateFactCourante = New System.Windows.Forms.DateTimePicker
        Me.m_bsrcFactTRP = New System.Windows.Forms.BindingSource(Me.components)
        Me.cbAppliquer = New System.Windows.Forms.Button
        Me.tbPeriodeFactCourante = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.tbMontantTTCFactCourante = New vini_app.textBoxCurrency
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbMontantHTFactCourante = New vini_app.textBoxCurrency
        Me.Label2 = New System.Windows.Forms.Label
        Me.liTiers = New System.Windows.Forms.LinkLabel
        Me.liFacture = New System.Windows.Forms.LinkLabel
        Me.cbRecherche = New System.Windows.Forms.Button
        Me.cbSauvegarder = New System.Windows.Forms.Button
        Me.dgvCommandes = New System.Windows.Forms.DataGridView
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersRS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dateLivraison = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.montantTransport = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcCommandeClient = New System.Windows.Forms.BindingSource(Me.components)
        Me.dgvFactTRP = New System.Windows.Forms.DataGridView
        Me.TiersRSDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dateFacture = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.totalHT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.grpFact.SuspendLayout()
        CType(Me.m_bsrcFactTRP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvCommandes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcCommandeClient, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvFactTRP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbCodeClient
        '
        Me.tbCodeClient.Location = New System.Drawing.Point(232, 56)
        Me.tbCodeClient.Name = "tbCodeClient"
        Me.tbCodeClient.Size = New System.Drawing.Size(88, 20)
        Me.tbCodeClient.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 108
        Me.Label1.Text = "Client"
        '
        'dtdateFin
        '
        Me.dtdateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdateFin.Location = New System.Drawing.Point(232, 32)
        Me.dtdateFin.Name = "dtdateFin"
        Me.dtdateFin.Size = New System.Drawing.Size(88, 20)
        Me.dtdateFin.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(176, 16)
        Me.Label8.TabIndex = 107
        Me.Label8.Text = "date de fin de Livraison"
        '
        'dtDatedeb
        '
        Me.dtDatedeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDatedeb.Location = New System.Drawing.Point(232, 8)
        Me.dtDatedeb.Name = "dtDatedeb"
        Me.dtDatedeb.Size = New System.Drawing.Size(88, 20)
        Me.dtDatedeb.TabIndex = 0
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 8)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(176, 16)
        Me.Label14.TabIndex = 106
        Me.Label14.Text = "date de début de Livraison"
        '
        'cbAfficher
        '
        Me.cbAfficher.BackColor = System.Drawing.SystemColors.Control
        Me.cbAfficher.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAfficher.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAfficher.Location = New System.Drawing.Point(8, 80)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAfficher.Size = New System.Drawing.Size(352, 20)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "A&fficher"
        Me.cbAfficher.UseVisualStyleBackColor = False
        '
        'cbGenerer
        '
        Me.cbGenerer.Location = New System.Drawing.Point(368, 160)
        Me.cbGenerer.Name = "cbGenerer"
        Me.cbGenerer.Size = New System.Drawing.Size(80, 24)
        Me.cbGenerer.TabIndex = 8
        Me.cbGenerer.Text = "- Générer ->"
        '
        'cbSelectionnerTout
        '
        Me.cbSelectionnerTout.Location = New System.Drawing.Point(368, 120)
        Me.cbSelectionnerTout.Name = "cbSelectionnerTout"
        Me.cbSelectionnerTout.Size = New System.Drawing.Size(80, 32)
        Me.cbSelectionnerTout.TabIndex = 6
        Me.cbSelectionnerTout.Text = "Sélectionner Tout"
        Me.cbSelectionnerTout.Visible = False
        '
        'cbAnnGenerer
        '
        Me.cbAnnGenerer.Location = New System.Drawing.Point(368, 192)
        Me.cbAnnGenerer.Name = "cbAnnGenerer"
        Me.cbAnnGenerer.Size = New System.Drawing.Size(80, 24)
        Me.cbAnnGenerer.TabIndex = 16
        Me.cbAnnGenerer.Text = "<- Annule -<"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(480, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 119
        Me.Label4.Text = "Date de Facture"
        '
        'dtDateFacture
        '
        Me.dtDateFacture.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFacture.Location = New System.Drawing.Point(584, 8)
        Me.dtDateFacture.Name = "dtDateFacture"
        Me.dtDateFacture.Size = New System.Drawing.Size(104, 20)
        Me.dtDateFacture.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(480, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 16)
        Me.Label6.TabIndex = 123
        Me.Label6.Text = "Période"
        '
        'tbPeriode
        '
        Me.tbPeriode.Location = New System.Drawing.Point(584, 56)
        Me.tbPeriode.Name = "tbPeriode"
        Me.tbPeriode.Size = New System.Drawing.Size(176, 20)
        Me.tbPeriode.TabIndex = 12
        '
        'grpFact
        '
        Me.grpFact.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpFact.Controls.Add(Me.dtDateFactCourante)
        Me.grpFact.Controls.Add(Me.cbAppliquer)
        Me.grpFact.Controls.Add(Me.tbPeriodeFactCourante)
        Me.grpFact.Controls.Add(Me.Label10)
        Me.grpFact.Controls.Add(Me.Label7)
        Me.grpFact.Controls.Add(Me.tbMontantTTCFactCourante)
        Me.grpFact.Controls.Add(Me.Label3)
        Me.grpFact.Controls.Add(Me.tbMontantHTFactCourante)
        Me.grpFact.Controls.Add(Me.Label2)
        Me.grpFact.Controls.Add(Me.liTiers)
        Me.grpFact.Controls.Add(Me.liFacture)
        Me.grpFact.Location = New System.Drawing.Point(456, 352)
        Me.grpFact.Name = "grpFact"
        Me.grpFact.Size = New System.Drawing.Size(536, 248)
        Me.grpFact.TabIndex = 15
        Me.grpFact.TabStop = False
        '
        'dtDateFactCourante
        '
        Me.dtDateFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcFactTRP, "dateFacture", True))
        Me.dtDateFactCourante.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFactCourante.Location = New System.Drawing.Point(112, 72)
        Me.dtDateFactCourante.Name = "dtDateFactCourante"
        Me.dtDateFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.dtDateFactCourante.TabIndex = 2
        '
        'm_bsrcFactTRP
        '
        Me.m_bsrcFactTRP.DataSource = GetType(vini_DB.FactTRP)
        '
        'cbAppliquer
        '
        Me.cbAppliquer.Location = New System.Drawing.Point(424, 216)
        Me.cbAppliquer.Name = "cbAppliquer"
        Me.cbAppliquer.Size = New System.Drawing.Size(104, 24)
        Me.cbAppliquer.TabIndex = 7
        Me.cbAppliquer.Text = "Appliquer"
        '
        'tbPeriodeFactCourante
        '
        Me.tbPeriodeFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactTRP, "periode", True))
        Me.tbPeriodeFactCourante.Location = New System.Drawing.Point(112, 109)
        Me.tbPeriodeFactCourante.Name = "tbPeriodeFactCourante"
        Me.tbPeriodeFactCourante.Size = New System.Drawing.Size(200, 20)
        Me.tbPeriodeFactCourante.TabIndex = 4
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 112)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(88, 24)
        Me.Label10.TabIndex = 140
        Me.Label10.Text = "Periode"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 136
        Me.Label7.Text = "DateFacture"
        '
        'tbMontantTTCFactCourante
        '
        Me.tbMontantTTCFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactTRP, "totalTTC", True))
        Me.tbMontantTTCFactCourante.Location = New System.Drawing.Point(112, 183)
        Me.tbMontantTTCFactCourante.Name = "tbMontantTTCFactCourante"
        Me.tbMontantTTCFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantTTCFactCourante.TabIndex = 6
        Me.tbMontantTTCFactCourante.Text = "0"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 186)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 17)
        Me.Label3.TabIndex = 134
        Me.Label3.Text = "Montant TTC"
        '
        'tbMontantHTFactCourante
        '
        Me.tbMontantHTFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactTRP, "totalHT", True))
        Me.tbMontantHTFactCourante.Location = New System.Drawing.Point(112, 146)
        Me.tbMontantHTFactCourante.Name = "tbMontantHTFactCourante"
        Me.tbMontantHTFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantHTFactCourante.TabIndex = 5
        Me.tbMontantHTFactCourante.Text = "0"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 149)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 132
        Me.Label2.Text = "Montant HT"
        '
        'liTiers
        '
        Me.liTiers.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactTRP, "TiersRS", True))
        Me.liTiers.Location = New System.Drawing.Point(8, 48)
        Me.liTiers.Name = "liTiers"
        Me.liTiers.Size = New System.Drawing.Size(520, 16)
        Me.liTiers.TabIndex = 1
        Me.liTiers.TabStop = True
        Me.liTiers.Text = "RS Tiers"
        '
        'liFacture
        '
        Me.liFacture.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactTRP, "shortResume", True))
        Me.liFacture.DataBindings.Add(New System.Windows.Forms.Binding("Tag", Me.m_bsrcFactTRP, "id", True))
        Me.liFacture.Location = New System.Drawing.Point(8, 24)
        Me.liFacture.Name = "liFacture"
        Me.liFacture.Size = New System.Drawing.Size(520, 24)
        Me.liFacture.TabIndex = 0
        Me.liFacture.TabStop = True
        Me.liFacture.Text = "Reference Facture"
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(328, 56)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(75, 23)
        Me.cbRecherche.TabIndex = 3
        Me.cbRecherche.Text = "Recherche"
        '
        'cbSauvegarder
        '
        Me.cbSauvegarder.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSauvegarder.Location = New System.Drawing.Point(688, 323)
        Me.cbSauvegarder.Name = "cbSauvegarder"
        Me.cbSauvegarder.Size = New System.Drawing.Size(80, 23)
        Me.cbSauvegarder.TabIndex = 126
        Me.cbSauvegarder.Text = "Sauvegarder"
        '
        'dgvCommandes
        '
        Me.dgvCommandes.AllowUserToAddRows = False
        Me.dgvCommandes.AllowUserToDeleteRows = False
        Me.dgvCommandes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvCommandes.AutoGenerateColumns = False
        Me.dgvCommandes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCommandes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCommandes.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.TiersCode, Me.TiersRS, Me.dateLivraison, Me.montantTransport})
        Me.dgvCommandes.DataSource = Me.m_bsrcCommandeClient
        Me.dgvCommandes.Location = New System.Drawing.Point(8, 120)
        Me.dgvCommandes.Name = "dgvCommandes"
        Me.dgvCommandes.ReadOnly = True
        Me.dgvCommandes.RowHeadersVisible = False
        Me.dgvCommandes.Size = New System.Drawing.Size(352, 480)
        Me.dgvCommandes.TabIndex = 127
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "CMD"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn.Width = 56
        '
        'TiersCode
        '
        Me.TiersCode.DataPropertyName = "TiersCode"
        Me.TiersCode.HeaderText = "Code CLT"
        Me.TiersCode.Name = "TiersCode"
        Me.TiersCode.ReadOnly = True
        Me.TiersCode.Width = 80
        '
        'TiersRS
        '
        Me.TiersRS.DataPropertyName = "TiersRS"
        Me.TiersRS.HeaderText = "RS Client"
        Me.TiersRS.Name = "TiersRS"
        Me.TiersRS.ReadOnly = True
        Me.TiersRS.Width = 76
        '
        'dateLivraison
        '
        Me.dateLivraison.DataPropertyName = "dateLivraison"
        Me.dateLivraison.HeaderText = "date Liv."
        Me.dateLivraison.Name = "dateLivraison"
        Me.dateLivraison.ReadOnly = True
        Me.dateLivraison.Width = 73
        '
        'montantTransport
        '
        Me.montantTransport.DataPropertyName = "montantTransport"
        Me.montantTransport.HeaderText = "Mt TRP"
        Me.montantTransport.Name = "montantTransport"
        Me.montantTransport.ReadOnly = True
        Me.montantTransport.Width = 69
        '
        'm_bsrcCommandeClient
        '
        Me.m_bsrcCommandeClient.DataSource = GetType(vini_DB.CommandeClient)
        '
        'dgvFactTRP
        '
        Me.dgvFactTRP.AllowUserToAddRows = False
        Me.dgvFactTRP.AllowUserToDeleteRows = False
        Me.dgvFactTRP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvFactTRP.AutoGenerateColumns = False
        Me.dgvFactTRP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvFactTRP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFactTRP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.TiersRSDataGridViewTextBoxColumn, Me.dateFacture, Me.totalHT, Me.DataGridViewTextBoxColumn1})
        Me.dgvFactTRP.DataSource = Me.m_bsrcFactTRP
        Me.dgvFactTRP.Location = New System.Drawing.Point(483, 120)
        Me.dgvFactTRP.Name = "dgvFactTRP"
        Me.dgvFactTRP.ReadOnly = True
        Me.dgvFactTRP.RowHeadersVisible = False
        Me.dgvFactTRP.Size = New System.Drawing.Size(509, 150)
        Me.dgvFactTRP.TabIndex = 128
        '
        'TiersRSDataGridViewTextBoxColumn
        '
        Me.TiersRSDataGridViewTextBoxColumn.DataPropertyName = "TiersRS"
        Me.TiersRSDataGridViewTextBoxColumn.HeaderText = "Client"
        Me.TiersRSDataGridViewTextBoxColumn.Name = "TiersRSDataGridViewTextBoxColumn"
        Me.TiersRSDataGridViewTextBoxColumn.ReadOnly = True
        Me.TiersRSDataGridViewTextBoxColumn.Width = 58
        '
        'dateFacture
        '
        Me.dateFacture.DataPropertyName = "dateFacture"
        Me.dateFacture.HeaderText = "date Facture"
        Me.dateFacture.Name = "dateFacture"
        Me.dateFacture.ReadOnly = True
        Me.dateFacture.Width = 92
        '
        'totalHT
        '
        Me.totalHT.DataPropertyName = "totalHT"
        Me.totalHT.HeaderText = "HT"
        Me.totalHT.Name = "totalHT"
        Me.totalHT.ReadOnly = True
        Me.totalHT.Width = 47
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "totalTTC"
        Me.DataGridViewTextBoxColumn1.HeaderText = "TTC"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 53
        '
        'frmGeneFactTRP
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(997, 630)
        Me.Controls.Add(Me.dgvFactTRP)
        Me.Controls.Add(Me.dgvCommandes)
        Me.Controls.Add(Me.cbSauvegarder)
        Me.Controls.Add(Me.cbRecherche)
        Me.Controls.Add(Me.grpFact)
        Me.Controls.Add(Me.tbPeriode)
        Me.Controls.Add(Me.tbCodeClient)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dtDateFacture)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbAnnGenerer)
        Me.Controls.Add(Me.cbSelectionnerTout)
        Me.Controls.Add(Me.cbGenerer)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtdateFin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtDatedeb)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cbAfficher)
        Me.Name = "frmGeneFactTRP"
        Me.Text = "Génération de factures de Transport"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grpFact.ResumeLayout(False)
        Me.grpFact.PerformLayout()
        CType(Me.m_bsrcFactTRP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvCommandes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcCommandeClient, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvFactTRP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Méthodes Vinicom"
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False
        m_ToolBarSaveEnabled = False
    End Sub

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        Dim objControl As Control
        MyBase.EnableControls(bEnabled)
        dtDatedeb.Enabled = True
        dtdateFin.Enabled = True
        tbCodeClient.Enabled = True
        cbAfficher.Enabled = True
        dtDateFacture.Enabled = True
        tbPeriode.Enabled = True
        cbRecherche.Enabled = True
        grpFact.Enabled = True
        For Each objControl In grpFact.Controls
            objControl.Enabled = True
        Next objControl

    End Sub

    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenetre n'est pas libre")
        Dim bReturn As Boolean
        bReturn = setElementCourant2(New FactTRP(New Client("", "")))

        Return bReturn
    End Function
    Protected Shadows Function getElementCourant() As FactTRP
        Return CType(m_ElementCourant, FactTRP)
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
        dtDateFacture.Value = Now()
        m_colCommandes = New Collection
    End Sub
    Private Sub afficheListeCommande()

        setcursorWait()
        debAffiche()
        m_bsrcCommandeClient.Clear()

        Debug.Assert(Not m_colCommandes Is Nothing)
        Dim obj As CommandeClient


        For Each obj In m_colCommandes
            m_bsrcCommandeClient.Add(obj)
        Next obj

        cbGenerer.Enabled = True

        finAffiche()
        restoreCursor()

    End Sub 'AfficheListeCommande


    Private Function setListeCommandes() As Boolean
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
            'Recupération de la liste des commandes avec Transport
            col = CommandeClient.getListe_TRP(ddeb, dfin, , codeClient)
            If col Is Nothing Then
                bReturn = False
            Else
                m_colCommandes = col
                bReturn = True
            End If
        Catch ex As Exception
            bReturn = False
            Debug.Assert(bReturn, ex.ToString)
        End Try
        finAffiche()
        restoreCursor()
        Return bReturn
    End Function
    Private Sub genererFactures()
        '===============================================================
        ' Function : genererFactures
        'Description : Génere les facures de transport à partir des commandes sélectionnées
        '               une fois les factures générées elles sont affichées
        '===============================================================
        Dim colFact As ColEvent

        Me.Cursor = Cursors.WaitCursor


        colFact = FactTRP.createFactTRPs(m_colCommandes, dtDateFacture.Value, , tbPeriode.Text)


        If Not colFact Is Nothing Then

            m_colFact = colFact
            afficheListeFactures()

            afficheListeCommande()
            setElementCourant2(Nothing)
        End If

        Me.Cursor = Cursors.Default

    End Sub 'genererFactures
    Private Function annuleGeneration() As Boolean
        '===============================================================
        ' Function : annuleGeneration
        'Description : Annule la génération de factures préalablement générées
        '               Si la Facture de com a été sauvegardé, elle est supprimée 
        '               Netoyage de la collection et de la liste Box
        '               Reaffichage de la liste des sous-commandes
        '===============================================================
        Dim objFactTRP As FactTRP
        Dim bReturn As Boolean

        setcursorWait()
        bReturn = True
        For Each objFactTRP In m_colFact
            If objFactTRP.id <> 0 Then
                'La Facture a été Sauvegardée => Destruction
                objFactTRP.bDeleted = True
                bReturn = bReturn And objFactTRP.Save()
                Debug.Assert(bReturn, FactTRP.getErreur())
            End If
        Next

        m_colFact = New ColEvent
        m_bsrcFactTRP.Clear()
        If setListeCommandes() Then
            afficheListeCommande()
        End If
        restoreCursor()

        Return bReturn
    End Function 'annuleGeneration

    Private Sub afficheListeFactures()
        Debug.Assert(Not m_colFact Is Nothing, "pColFact = Nothing")
        Dim objFactTRP As FactTRP



        debAffiche()
        setcursorWait()

        m_bsrcFactTRP.Clear()
        For Each objFactTRP In m_colFact
            m_bsrcFactTRP.Add(objFactTRP)
        Next objFactTRP

        dgvFactTRP.Enabled = True
        restoreCursor()
        finAffiche()

    End Sub
    Private Function appliqueModifications() As Boolean
        Dim objFact As FactTRP
        If m_bsrcFactTRP.Current IsNot Nothing Then
            objFact = CType(m_bsrcFactTRP.Current, FactTRP)
            objFact.Save()
        End If
    End Function
    Private Function elementCourantSansModif() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
            If Not getElementCourant() Is Nothing Then
                bReturn = bReturn And dtDateFactCourante.Value = getElementCourant().dateCommande
                bReturn = bReturn And tbPeriodeFactCourante.Text = getElementCourant().periode
                bReturn = bReturn And tbMontantHTFactCourante.Text = getElementCourant().totalHT
                bReturn = bReturn And tbMontantTTCFactCourante.Text = getElementCourant().totalTTC
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Sub rechercheClient()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.CLIENT, tbCodeClient)

        If Not objTiers Is Nothing Then
            tbCodeClient.Text = objTiers.code
        End If
    End Sub 'rechercheClient

    Private Function sauvegarderFactures() As Boolean
        Dim objFactTRP As FactTRP
        Dim bReturn As Boolean

        bReturn = False
        If Not m_colFact Is Nothing Then
            bReturn = True
            For Each objFactTRP In m_colFact
                bReturn = bReturn And objFactTRP.Save()
                Debug.Assert(bReturn, FactTRP.getErreur())
            Next
        End If

        Return bReturn
    End Function 'sauvegarderFactures
#End Region
#Region "Gestion des Evenements"
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        sauvegardeElementCourant()
        If setListeCommandes() Then
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

    Private Sub cbGenerer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbGenerer.Click
        Call genererFactures()
        If m_bsrcFactTRP.Count > 0 Then
            dgvFactTRP.Enabled = True
            cbSauvegarder.Enabled = True
            cbAnnGenerer.Enabled = True
            cbGenerer.Enabled = False
        End If
    End Sub





    Private Sub cbAppliquer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAppliquer.Click
        appliqueModifications()
    End Sub

    Private Sub liTiers_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liTiers.LinkClicked
        afficheFenetreClient(liTiers.Tag)
    End Sub

    Private Sub cbAnnGenerer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnGenerer.Click
        annuleGeneration()

        dgvFactTRP.Enabled = False
        cbSauvegarder.Enabled = False
        cbAnnGenerer.Enabled = False
        cbGenerer.Enabled = True
    End Sub

    Private Sub liFactTRP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liFacture.LinkClicked
        afficheFenetreFactTrp(liFacture.Tag)
    End Sub

    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        rechercheClient()
    End Sub


    Private Sub cbSauvegarder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSauvegarder.Click
        If sauvegarderFactures() Then
            afficheListeFactures()
        End If
    End Sub

End Class
