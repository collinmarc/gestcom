Option Strict Off
Imports vini_DB
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
'Imports FAXCOMLib
Imports System.Windows.Forms.Cursors

Public Class frmGestFactColisage
    Inherits frmDonBase
    Private Enum vncColLigneCommande
        COL_NUM = 0
        COL_DATEDEB = 1
        COL_DATEFIN = 2
        COL_STOCKINI = 3
        COL_ENTREES = 4
        COL_SORTIES = 5
        COL_STOCKFINAL = 6
        COL_PU = 7
        COL_MONTANTHT = 8
        COL_NBRECOL = 9
    End Enum

#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()
        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()


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
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents cbSupprimerLigne As System.Windows.Forms.Button
    Public WithEvents cbAjouterLigne As System.Windows.Forms.Button
    Public WithEvents tbTotalHT As textBoxCurrency
    Public WithEvents Label26 As System.Windows.Forms.Label
    Public WithEvents tpLignes As System.Windows.Forms.TabPage
    Public WithEvents tpCommentaires As System.Windows.Forms.TabPage
    Friend WithEvents tpValidCmd As System.Windows.Forms.TabPage
    Friend WithEvents laEtatCommande As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents tbCode As System.Windows.Forms.TextBox
    Public WithEvents SSTabCommandeClient As System.Windows.Forms.TabControl
    Friend WithEvents tbCommentaireFacturation As System.Windows.Forms.TextBox
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Public WithEvents tbTotalTTC As textBoxCurrency
    Friend WithEvents grpEntete As System.Windows.Forms.GroupBox
    Friend WithEvents dtDateFact As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbPeriode As System.Windows.Forms.TextBox
    Friend WithEvents cbAfficherEtat As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents cbRecalcTotaux As System.Windows.Forms.Button
    Friend WithEvents liTiers As System.Windows.Forms.LinkLabel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbMontantTaxes As textBoxCurrency
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboModeReglement As System.Windows.Forms.ComboBox
    Friend WithEvents m_bsrcFactColisage As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcModeReglement As System.Windows.Forms.BindingSource
    Friend WithEvents ckEntete As System.Windows.Forms.CheckBox
    Friend WithEvents cbAnnExport As System.Windows.Forms.Button
    Friend WithEvents cbBrowse As System.Windows.Forms.Button
    Friend WithEvents tbDossierStockage As System.Windows.Forms.TextBox
    Friend WithEvents cbStocker As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcLgFact As System.Windows.Forms.BindingSource
    Friend WithEvents DDebDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DFinDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StockInitialDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EntreesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SortiesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StockFinalDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PuDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MontantHTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents cbReglement As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.grpEntete = New System.Windows.Forms.GroupBox()
        Me.tbPeriode = New System.Windows.Forms.TextBox()
        Me.m_bsrcFactColisage = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtDateFact = New System.Windows.Forms.DateTimePicker()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.liTiers = New System.Windows.Forms.LinkLabel()
        Me.laEtatCommande = New System.Windows.Forms.Label()
        Me.tbCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbAnnExport = New System.Windows.Forms.Button()
        Me.SSTabCommandeClient = New System.Windows.Forms.TabControl()
        Me.tpLignes = New System.Windows.Forms.TabPage()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DDebDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DFinDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StockInitialDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EntreesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SortiesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StockFinalDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PuDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MontantHTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_bsrcLgFact = New System.Windows.Forms.BindingSource(Me.components)
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboModeReglement = New System.Windows.Forms.ComboBox()
        Me.m_bsrcModeReglement = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbMontantTaxes = New vini_app.textBoxCurrency()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbRecalcTotaux = New System.Windows.Forms.Button()
        Me.tbTotalTTC = New vini_app.textBoxCurrency()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.cbSupprimerLigne = New System.Windows.Forms.Button()
        Me.cbAjouterLigne = New System.Windows.Forms.Button()
        Me.tbTotalHT = New vini_app.textBoxCurrency()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tpCommentaires = New System.Windows.Forms.TabPage()
        Me.tbCommentaireFacturation = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.tpValidCmd = New System.Windows.Forms.TabPage()
        Me.cbBrowse = New System.Windows.Forms.Button()
        Me.tbDossierStockage = New System.Windows.Forms.TextBox()
        Me.cbStocker = New System.Windows.Forms.Button()
        Me.ckEntete = New System.Windows.Forms.CheckBox()
        Me.cbAfficherEtat = New System.Windows.Forms.Button()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.cbReglement = New System.Windows.Forms.Button()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.grpEntete.SuspendLayout()
        CType(Me.m_bsrcFactColisage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSTabCommandeClient.SuspendLayout()
        Me.tpLignes.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcLgFact, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcModeReglement, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpCommentaires.SuspendLayout()
        Me.tpValidCmd.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpEntete
        '
        Me.grpEntete.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpEntete.Controls.Add(Me.tbPeriode)
        Me.grpEntete.Controls.Add(Me.Label3)
        Me.grpEntete.Controls.Add(Me.dtDateFact)
        Me.grpEntete.Controls.Add(Me.Label29)
        Me.grpEntete.Controls.Add(Me.liTiers)
        Me.grpEntete.Controls.Add(Me.laEtatCommande)
        Me.grpEntete.Controls.Add(Me.tbCode)
        Me.grpEntete.Controls.Add(Me.Label1)
        Me.grpEntete.Location = New System.Drawing.Point(8, 0)
        Me.grpEntete.Name = "grpEntete"
        Me.grpEntete.Size = New System.Drawing.Size(992, 64)
        Me.grpEntete.TabIndex = 0
        Me.grpEntete.TabStop = False
        Me.grpEntete.Text = "Facture"
        '
        'tbPeriode
        '
        Me.tbPeriode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPeriode.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactColisage, "periode", True))
        Me.tbPeriode.Location = New System.Drawing.Point(659, 16)
        Me.tbPeriode.Name = "tbPeriode"
        Me.tbPeriode.Size = New System.Drawing.Size(184, 20)
        Me.tbPeriode.TabIndex = 2
        '
        'm_bsrcFactColisage
        '
        Me.m_bsrcFactColisage.DataSource = GetType(vini_DB.FactColisage)
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(592, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 16)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Période"
        '
        'dtDateFact
        '
        Me.dtDateFact.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcFactColisage, "dateFacture", True))
        Me.dtDateFact.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFact.Location = New System.Drawing.Point(280, 16)
        Me.dtDateFact.Name = "dtDateFact"
        Me.dtDateFact.Size = New System.Drawing.Size(104, 20)
        Me.dtDateFact.TabIndex = 1
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(168, 16)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(104, 16)
        Me.Label29.TabIndex = 12
        Me.Label29.Text = "Date de Facture"
        '
        'liTiers
        '
        Me.liTiers.Location = New System.Drawing.Point(8, 40)
        Me.liTiers.Name = "liTiers"
        Me.liTiers.Size = New System.Drawing.Size(712, 16)
        Me.liTiers.TabIndex = 8
        '
        'laEtatCommande
        '
        Me.laEtatCommande.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.laEtatCommande.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.laEtatCommande.ForeColor = System.Drawing.Color.Red
        Me.laEtatCommande.Location = New System.Drawing.Point(856, 16)
        Me.laEtatCommande.Name = "laEtatCommande"
        Me.laEtatCommande.Size = New System.Drawing.Size(128, 20)
        Me.laEtatCommande.TabIndex = 7
        '
        'tbCode
        '
        Me.tbCode.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactColisage, "code", True))
        Me.tbCode.Location = New System.Drawing.Point(48, 16)
        Me.tbCode.Name = "tbCode"
        Me.tbCode.Size = New System.Drawing.Size(112, 20)
        Me.tbCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "N°"
        '
        'cbAnnExport
        '
        Me.cbAnnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAnnExport.Location = New System.Drawing.Point(876, 72)
        Me.cbAnnExport.Name = "cbAnnExport"
        Me.cbAnnExport.Size = New System.Drawing.Size(124, 24)
        Me.cbAnnExport.TabIndex = 1
        Me.cbAnnExport.Text = "Annulation Export"
        Me.cbAnnExport.UseVisualStyleBackColor = True
        '
        'SSTabCommandeClient
        '
        Me.SSTabCommandeClient.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SSTabCommandeClient.Controls.Add(Me.tpLignes)
        Me.SSTabCommandeClient.Controls.Add(Me.tpCommentaires)
        Me.SSTabCommandeClient.Controls.Add(Me.tpValidCmd)
        Me.SSTabCommandeClient.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTabCommandeClient.Location = New System.Drawing.Point(0, 93)
        Me.SSTabCommandeClient.Name = "SSTabCommandeClient"
        Me.SSTabCommandeClient.SelectedIndex = 0
        Me.SSTabCommandeClient.Size = New System.Drawing.Size(1000, 571)
        Me.SSTabCommandeClient.TabIndex = 0
        '
        'tpLignes
        '
        Me.tpLignes.Controls.Add(Me.DataGridView1)
        Me.tpLignes.Controls.Add(Me.DateTimePicker1)
        Me.tpLignes.Controls.Add(Me.Label5)
        Me.tpLignes.Controls.Add(Me.cboModeReglement)
        Me.tpLignes.Controls.Add(Me.Label4)
        Me.tpLignes.Controls.Add(Me.tbMontantTaxes)
        Me.tpLignes.Controls.Add(Me.Label2)
        Me.tpLignes.Controls.Add(Me.cbRecalcTotaux)
        Me.tpLignes.Controls.Add(Me.tbTotalTTC)
        Me.tpLignes.Controls.Add(Me.Label51)
        Me.tpLignes.Controls.Add(Me.cbSupprimerLigne)
        Me.tpLignes.Controls.Add(Me.cbAjouterLigne)
        Me.tpLignes.Controls.Add(Me.tbTotalHT)
        Me.tpLignes.Controls.Add(Me.Label10)
        Me.tpLignes.Location = New System.Drawing.Point(4, 22)
        Me.tpLignes.Name = "tpLignes"
        Me.tpLignes.Size = New System.Drawing.Size(992, 545)
        Me.tpLignes.TabIndex = 1
        Me.tpLignes.Text = "Lignes"
        Me.tpLignes.Visible = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DDebDataGridViewTextBoxColumn, Me.DFinDataGridViewTextBoxColumn, Me.StockInitialDataGridViewTextBoxColumn, Me.EntreesDataGridViewTextBoxColumn, Me.SortiesDataGridViewTextBoxColumn, Me.StockFinalDataGridViewTextBoxColumn, Me.PuDataGridViewTextBoxColumn, Me.MontantHTDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcLgFact
        Me.DataGridView1.Location = New System.Drawing.Point(8, 14)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(968, 412)
        Me.DataGridView1.TabIndex = 0
        '
        'DDebDataGridViewTextBoxColumn
        '
        Me.DDebDataGridViewTextBoxColumn.DataPropertyName = "dDeb"
        Me.DDebDataGridViewTextBoxColumn.HeaderText = "Date début"
        Me.DDebDataGridViewTextBoxColumn.Name = "DDebDataGridViewTextBoxColumn"
        Me.DDebDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DFinDataGridViewTextBoxColumn
        '
        Me.DFinDataGridViewTextBoxColumn.DataPropertyName = "dFin"
        Me.DFinDataGridViewTextBoxColumn.HeaderText = "Date fin"
        Me.DFinDataGridViewTextBoxColumn.Name = "DFinDataGridViewTextBoxColumn"
        Me.DFinDataGridViewTextBoxColumn.ReadOnly = True
        '
        'StockInitialDataGridViewTextBoxColumn
        '
        Me.StockInitialDataGridViewTextBoxColumn.DataPropertyName = "StockInitial"
        Me.StockInitialDataGridViewTextBoxColumn.HeaderText = "Stock Initial"
        Me.StockInitialDataGridViewTextBoxColumn.Name = "StockInitialDataGridViewTextBoxColumn"
        Me.StockInitialDataGridViewTextBoxColumn.ReadOnly = True
        '
        'EntreesDataGridViewTextBoxColumn
        '
        Me.EntreesDataGridViewTextBoxColumn.DataPropertyName = "Entrees"
        Me.EntreesDataGridViewTextBoxColumn.HeaderText = "Entrées"
        Me.EntreesDataGridViewTextBoxColumn.Name = "EntreesDataGridViewTextBoxColumn"
        Me.EntreesDataGridViewTextBoxColumn.ReadOnly = True
        '
        'SortiesDataGridViewTextBoxColumn
        '
        Me.SortiesDataGridViewTextBoxColumn.DataPropertyName = "Sorties"
        Me.SortiesDataGridViewTextBoxColumn.HeaderText = "Sorties"
        Me.SortiesDataGridViewTextBoxColumn.Name = "SortiesDataGridViewTextBoxColumn"
        Me.SortiesDataGridViewTextBoxColumn.ReadOnly = True
        '
        'StockFinalDataGridViewTextBoxColumn
        '
        Me.StockFinalDataGridViewTextBoxColumn.DataPropertyName = "StockFinal"
        Me.StockFinalDataGridViewTextBoxColumn.HeaderText = "Stock Final"
        Me.StockFinalDataGridViewTextBoxColumn.Name = "StockFinalDataGridViewTextBoxColumn"
        Me.StockFinalDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PuDataGridViewTextBoxColumn
        '
        Me.PuDataGridViewTextBoxColumn.DataPropertyName = "pu"
        DataGridViewCellStyle1.Format = "c2"
        Me.PuDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle1
        Me.PuDataGridViewTextBoxColumn.HeaderText = "Prix unitaire"
        Me.PuDataGridViewTextBoxColumn.Name = "PuDataGridViewTextBoxColumn"
        Me.PuDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MontantHTDataGridViewTextBoxColumn
        '
        Me.MontantHTDataGridViewTextBoxColumn.DataPropertyName = "MontantHT"
        DataGridViewCellStyle2.Format = "c2"
        Me.MontantHTDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle2
        Me.MontantHTDataGridViewTextBoxColumn.HeaderText = "Montant HT"
        Me.MontantHTDataGridViewTextBoxColumn.Name = "MontantHTDataGridViewTextBoxColumn"
        Me.MontantHTDataGridViewTextBoxColumn.ReadOnly = True
        '
        'm_bsrcLgFact
        '
        Me.m_bsrcLgFact.DataSource = GetType(vini_DB.LgFactColisage)
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DateTimePicker1.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactColisage, "dEcheance", True))
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(136, 506)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(90, 20)
        Me.DateTimePicker1.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 510)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(95, 13)
        Me.Label5.TabIndex = 67
        Me.Label5.Text = "Date d'échéance :"
        '
        'cboModeReglement
        '
        Me.cboModeReglement.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboModeReglement.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcFactColisage, "idModeReglement", True))
        Me.cboModeReglement.DataSource = Me.m_bsrcModeReglement
        Me.cboModeReglement.DisplayMember = "valeur"
        Me.cboModeReglement.Location = New System.Drawing.Point(136, 469)
        Me.cboModeReglement.Name = "cboModeReglement"
        Me.cboModeReglement.Size = New System.Drawing.Size(224, 21)
        Me.cboModeReglement.TabIndex = 3
        Me.cboModeReglement.ValueMember = "id"
        '
        'm_bsrcModeReglement
        '
        Me.m_bsrcModeReglement.DataSource = GetType(vini_DB.Param)
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(8, 472)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(122, 18)
        Me.Label4.TabIndex = 65
        Me.Label4.Text = "Mode de réglement :"
        '
        'tbMontantTaxes
        '
        Me.tbMontantTaxes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMontantTaxes.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactColisage, "montantTaxes", True))
        Me.tbMontantTaxes.Location = New System.Drawing.Point(864, 463)
        Me.tbMontantTaxes.Name = "tbMontantTaxes"
        Me.tbMontantTaxes.Size = New System.Drawing.Size(112, 20)
        Me.tbMontantTaxes.TabIndex = 6
        Me.tbMontantTaxes.Text = "0"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(711, 472)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 16)
        Me.Label2.TabIndex = 63
        Me.Label2.Text = "Frais d'enregistrement"
        '
        'cbRecalcTotaux
        '
        Me.cbRecalcTotaux.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbRecalcTotaux.Location = New System.Drawing.Point(585, 467)
        Me.cbRecalcTotaux.Name = "cbRecalcTotaux"
        Me.cbRecalcTotaux.Size = New System.Drawing.Size(75, 23)
        Me.cbRecalcTotaux.TabIndex = 5
        Me.cbRecalcTotaux.Text = "Re&Calcul"
        '
        'tbTotalTTC
        '
        Me.tbTotalTTC.AcceptsReturn = True
        Me.tbTotalTTC.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTotalTTC.BackColor = System.Drawing.SystemColors.Window
        Me.tbTotalTTC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbTotalTTC.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactColisage, "totalTTC", True))
        Me.tbTotalTTC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbTotalTTC.Location = New System.Drawing.Point(863, 515)
        Me.tbTotalTTC.MaxLength = 0
        Me.tbTotalTTC.Name = "tbTotalTTC"
        Me.tbTotalTTC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbTotalTTC.Size = New System.Drawing.Size(113, 20)
        Me.tbTotalTTC.TabIndex = 8
        Me.tbTotalTTC.Text = "0"
        '
        'Label51
        '
        Me.Label51.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label51.Location = New System.Drawing.Point(711, 515)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(128, 16)
        Me.Label51.TabIndex = 62
        Me.Label51.Text = "Total TTC"
        '
        'cbSupprimerLigne
        '
        Me.cbSupprimerLigne.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSupprimerLigne.BackColor = System.Drawing.SystemColors.Control
        Me.cbSupprimerLigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbSupprimerLigne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbSupprimerLigne.Location = New System.Drawing.Point(903, 432)
        Me.cbSupprimerLigne.Name = "cbSupprimerLigne"
        Me.cbSupprimerLigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbSupprimerLigne.Size = New System.Drawing.Size(73, 25)
        Me.cbSupprimerLigne.TabIndex = 2
        Me.cbSupprimerLigne.Text = "&Supprimer"
        Me.cbSupprimerLigne.UseVisualStyleBackColor = False
        '
        'cbAjouterLigne
        '
        Me.cbAjouterLigne.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAjouterLigne.BackColor = System.Drawing.SystemColors.Control
        Me.cbAjouterLigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAjouterLigne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAjouterLigne.Location = New System.Drawing.Point(824, 432)
        Me.cbAjouterLigne.Name = "cbAjouterLigne"
        Me.cbAjouterLigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAjouterLigne.Size = New System.Drawing.Size(73, 25)
        Me.cbAjouterLigne.TabIndex = 1
        Me.cbAjouterLigne.Text = "A&jouter"
        Me.cbAjouterLigne.UseVisualStyleBackColor = False
        '
        'tbTotalHT
        '
        Me.tbTotalHT.AcceptsReturn = True
        Me.tbTotalHT.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTotalHT.BackColor = System.Drawing.SystemColors.Window
        Me.tbTotalHT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbTotalHT.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactColisage, "totalHT", True))
        Me.tbTotalHT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbTotalHT.Location = New System.Drawing.Point(863, 489)
        Me.tbTotalHT.MaxLength = 0
        Me.tbTotalHT.Name = "tbTotalHT"
        Me.tbTotalHT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbTotalHT.Size = New System.Drawing.Size(113, 20)
        Me.tbTotalHT.TabIndex = 7
        Me.tbTotalHT.Text = "0"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(711, 496)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(136, 17)
        Me.Label10.TabIndex = 60
        Me.Label10.Text = "Total H.T."
        '
        'tpCommentaires
        '
        Me.tpCommentaires.Controls.Add(Me.tbCommentaireFacturation)
        Me.tpCommentaires.Controls.Add(Me.Label26)
        Me.tpCommentaires.Location = New System.Drawing.Point(4, 22)
        Me.tpCommentaires.Name = "tpCommentaires"
        Me.tpCommentaires.Size = New System.Drawing.Size(992, 545)
        Me.tpCommentaires.TabIndex = 3
        Me.tpCommentaires.Text = "Commentaires"
        Me.tpCommentaires.Visible = False
        '
        'tbCommentaireFacturation
        '
        Me.tbCommentaireFacturation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbCommentaireFacturation.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactColisage, "CommentaireFacturationText", True))
        Me.tbCommentaireFacturation.Location = New System.Drawing.Point(96, 8)
        Me.tbCommentaireFacturation.Multiline = True
        Me.tbCommentaireFacturation.Name = "tbCommentaireFacturation"
        Me.tbCommentaireFacturation.Size = New System.Drawing.Size(880, 120)
        Me.tbCommentaireFacturation.TabIndex = 2
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.SystemColors.Control
        Me.Label26.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label26.Location = New System.Drawing.Point(8, 16)
        Me.Label26.Name = "Label26"
        Me.Label26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label26.Size = New System.Drawing.Size(64, 25)
        Me.Label26.TabIndex = 74
        Me.Label26.Text = "Facturation"
        '
        'tpValidCmd
        '
        Me.tpValidCmd.Controls.Add(Me.CrystalReportViewer1)
        Me.tpValidCmd.Controls.Add(Me.cbBrowse)
        Me.tpValidCmd.Controls.Add(Me.tbDossierStockage)
        Me.tpValidCmd.Controls.Add(Me.cbStocker)
        Me.tpValidCmd.Controls.Add(Me.ckEntete)
        Me.tpValidCmd.Controls.Add(Me.cbAfficherEtat)
        Me.tpValidCmd.Location = New System.Drawing.Point(4, 22)
        Me.tpValidCmd.Name = "tpValidCmd"
        Me.tpValidCmd.Size = New System.Drawing.Size(992, 545)
        Me.tpValidCmd.TabIndex = 4
        Me.tpValidCmd.Text = "Editions"
        '
        'cbBrowse
        '
        Me.cbBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbBrowse.Location = New System.Drawing.Point(826, 2)
        Me.cbBrowse.Name = "cbBrowse"
        Me.cbBrowse.Size = New System.Drawing.Size(40, 24)
        Me.cbBrowse.TabIndex = 135
        Me.cbBrowse.Text = "..."
        '
        'tbDossierStockage
        '
        Me.tbDossierStockage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbDossierStockage.Location = New System.Drawing.Point(400, 3)
        Me.tbDossierStockage.Name = "tbDossierStockage"
        Me.tbDossierStockage.Size = New System.Drawing.Size(420, 20)
        Me.tbDossierStockage.TabIndex = 134
        '
        'cbStocker
        '
        Me.cbStocker.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbStocker.Location = New System.Drawing.Point(872, 2)
        Me.cbStocker.Name = "cbStocker"
        Me.cbStocker.Size = New System.Drawing.Size(104, 24)
        Me.cbStocker.TabIndex = 133
        Me.cbStocker.Text = "Stocker"
        '
        'ckEntete
        '
        Me.ckEntete.AutoSize = True
        Me.ckEntete.Location = New System.Drawing.Point(10, 8)
        Me.ckEntete.Name = "ckEntete"
        Me.ckEntete.Size = New System.Drawing.Size(57, 17)
        Me.ckEntete.TabIndex = 132
        Me.ckEntete.Text = "Entete"
        Me.ckEntete.UseVisualStyleBackColor = True
        '
        'cbAfficherEtat
        '
        Me.cbAfficherEtat.Location = New System.Drawing.Point(222, 3)
        Me.cbAfficherEtat.Name = "cbAfficherEtat"
        Me.cbAfficherEtat.Size = New System.Drawing.Size(80, 24)
        Me.cbAfficherEtat.TabIndex = 127
        Me.cbAfficherEtat.Text = "Afficher"
        '
        'Label34
        '
        Me.Label34.Location = New System.Drawing.Point(104, 24)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(128, 16)
        Me.Label34.TabIndex = 68
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(496, 232)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 24)
        Me.Button1.TabIndex = 67
        Me.Button1.Text = "Editer BaF"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(368, 232)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 24)
        Me.Button3.TabIndex = 10
        Me.Button3.Text = "Valider BaF"
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(8, 56)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(64, 16)
        Me.Label33.TabIndex = 9
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Location = New System.Drawing.Point(80, 48)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(488, 176)
        Me.RichTextBox2.TabIndex = 8
        Me.RichTextBox2.Text = ""
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(424, 136)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(88, 24)
        Me.Button4.TabIndex = 7
        Me.Button4.Text = "Appliquer"
        '
        'Label35
        '
        Me.Label35.Location = New System.Drawing.Point(8, 24)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(88, 16)
        Me.Label35.TabIndex = 0
        '
        'cbReglement
        '
        Me.cbReglement.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbReglement.Location = New System.Drawing.Point(795, 72)
        Me.cbReglement.Name = "cbReglement"
        Me.cbReglement.Size = New System.Drawing.Size(75, 24)
        Me.cbReglement.TabIndex = 0
        Me.cbReglement.Text = "Règlements"
        Me.cbReglement.UseVisualStyleBackColor = True
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.CrystalReportViewer1.DisplayStatusBar = False
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(10, 32)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(966, 510)
        Me.CrystalReportViewer1.TabIndex = 136
        Me.CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'frmGestFactColisage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.ClientSize = New System.Drawing.Size(1016, 686)
        Me.Controls.Add(Me.cbReglement)
        Me.Controls.Add(Me.cbAnnExport)
        Me.Controls.Add(Me.SSTabCommandeClient)
        Me.Controls.Add(Me.grpEntete)
        Me.Name = "frmGestFactColisage"
        Me.ShowInTaskbar = False
        Me.Text = "Facture de colisage"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grpEntete.ResumeLayout(False)
        Me.grpEntete.PerformLayout()
        CType(Me.m_bsrcFactColisage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSTabCommandeClient.ResumeLayout(False)
        Me.tpLignes.ResumeLayout(False)
        Me.tpLignes.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcLgFact, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcModeReglement, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpCommentaires.ResumeLayout(False)
        Me.tpCommentaires.PerformLayout()
        Me.tpValidCmd.ResumeLayout(False)
        Me.tpValidCmd.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Méthodes Vinicom"

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)
        Me.tbCode.Enabled = False
    End Sub
    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenetre n'est pas libre")
        Dim bReturn As Boolean
        bReturn = setElementCourant2(New FactColisage(New Fournisseur()))

        Return bReturn
    End Function
    Protected Shadows Function getElementCourant() As FactColisage
        Return CType(m_ElementCourant, FactColisage)
    End Function
    Public Overrides Function setElementCourant2(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        If Not pElement Is Nothing Then
            pElement.load()
            bReturn = MyBase.setElementCourant2(pElement)
            getElementCourant().loadcolLignes()
        Else
            bReturn = MyBase.setElementCourant2(pElement)
        End If
        m_bsrcFactColisage.Clear()
        If bReturn Then
            If Not getElementCourant() Is Nothing Then
                m_bsrcFactColisage.Add(getElementCourant())
                cbAjouterLigne.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
                cbSupprimerLigne.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
                dtDateFact.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
                tbPeriode.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
                tbTotalHT.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
                tbTotalTTC.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
                cbRecalcTotaux.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
                tbMontantTaxes.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
                cboModeReglement.Enabled = (getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactCOLExportee)
            End If
            setToolbarButtons()
        End If
        Return bReturn

    End Function 'SetElement_Specifique
    'Interface ElementCourant -> Ecran
    Public Overrides Function AfficheElement() As Boolean

        Dim objParam As Param


        'Affichage des caractéristiques de la commande


        'Etat de la commande
        AfficheEtat()

        'Fournisseur
        liTiers.Text = getElementCourant().oTiers.shortResume
        liTiers.Tag = getElementCourant().oTiers.id


        'Commentaires
        tbCommentaireFacturation.Text = getElementCourant().CommFacturation.comment

        'Lignes de Commandes
        affichecolLignes()

        'Mode de reglement
        For Each objParam In cboModeReglement.Items
            If objParam.id = getElementCourant().idModeReglement Then
                cboModeReglement.SelectedItem = objParam
                Exit For
            End If
        Next

        CrystalReportViewer1.ReportSource = Nothing
        CrystalReportViewer1.Refresh()


        Return True
    End Function 'AfficheElement
    Public Overrides Function MAJElement() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Dim bReturn As Boolean

        Try
            bReturn = True


        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn
    End Function 'MAJElement
    Public Overrides Function ControleAvantSauvegarde() As Boolean
        Dim bReturn As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing, "Element Courant Requis")

        Try
            bReturn = True
        Catch ex As Exception
            bReturn = False
            DisplayError("ControleAvantSauvegarde", ex.ToString)
        End Try
        Return bReturn
    End Function
    Public Overrides Function getResume() As String 'Rend le caption de la fenêtre
        Dim str As String
        str = "FCTCOL"
        If Not getElementCourant() Is Nothing Then
            str = str & getElementCourant().shortResume()
        End If
        Return str
    End Function 'getResume
    Protected Overrides Function frmNew() As Boolean
        Dim breturn As Boolean
        breturn = False
        Return breturn
    End Function
    Protected Overrides Function frmDel() As Boolean
        Dim bReturn As Boolean
        MyBase.frmDel()
        bReturn = setElementCourant2(Nothing)
        Return bReturn
    End Function ' frmDel

    Protected Overrides Sub setfrmUpdated()
        If Not getElementCourant() Is Nothing Then
            MAJElementCourant()
            MyBase.setfrmUpdated()
        End If
    End Sub
    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        MyBase.AddHandlerValidated(ocol)
        RemoveHandler cbAfficherEtat.Click, AddressOf ControlUpdated
    End Sub


    '==============================================================================================
    '==============================================================================================
    '==============================================================================================
#End Region
#Region "Methodes Privées"
    Protected Sub initFenetre()
        debAffiche()
        m_TypeDonnees = vncEnums.vncTypeDonnee.FACTCOL
        CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        CrystalReportViewer1.DisplayToolbar = True
        Dim objModeReglement As Param
        For Each objModeReglement In Param.colModeReglement
            m_bsrcModeReglement.Add(objModeReglement)
        Next
        finAffiche()
    End Sub 'initFenetre

    'Affiche la boite de dialogue d'ajout de ligne
    Protected Function ajouteruneLigne() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Debug.Assert(Not getElementCourant().oTiers Is Nothing)
        Dim bReturn As Boolean
        Dim odlg As dlgLgFactColisage
        Dim objLg As LgFactColisage


        objLg = New LgFactColisage
        objLg.num = getElementCourant().getNextNumLg()
        odlg = New dlgLgFactColisage
        odlg.setElementCourant(objLg)
        odlg.setTiersCourant(getElementCourant().oTiers)
        If odlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            objLg = getElementCourant().AjouteLigneFactColisage(objLg, True)
            m_bsrcLgFact.Add(objLg)
        End If
        bReturn = True

        Return bReturn

    End Function 'Ajouteruneligne
    'Affiche la boite de dialogue de modification de ligne
    Protected Function supprimeruneLigne() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Dim objLg As LgFactColisage
        Dim bReturn As Boolean

        If Not (m_bsrcLgFact.Current Is Nothing) Then
            Try
                objLg = m_bsrcLgFact.Current
                If MsgBox("Etes-vous sur de vouloir supprimer la ligne " & objLg.num, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    getElementCourant().supprimeLigne(objLg.num)
                    setfrmUpdated()
                    bReturn = True
                End If
            Catch ex As Exception
                bReturn = False
            End Try
            getElementCourant().calculPrixTotal()
            debAffiche()
            affichecolLignes()
            finAffiche()
        End If


        bReturn = True

        Return bReturn
    End Function 'Ajouteruneligne

    'Affiche une ligne de commande dans le tableau 
    Protected Function affichecolLignes() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Debug.Assert(Not getElementCourant().colLignes Is Nothing)

        Dim objLg As LgFactColisage
        debAffiche()
        m_bsrcLgFact.Clear()
        For Each objLg In getElementCourant().colLignes
            m_bsrcLgFact.Add(objLg)
        Next objLg
        m_bsrcLgFact.ResetBindings(False)
        finAffiche()
    End Function ' Affichage de des lignes de commandes

    Private Sub afficherRapport()
        Dim objReport As New ReportDocument()
        Dim tabIds As ArrayList
        Dim strReport As String


        setcursorWait()
        tabIds = New ArrayList()
        tabIds.Add(getElementCourant().id)
        'Choix de l'état
        '        If rbFacture.Checked Then
        strReport = PATHTOREPORTS & "crFactureColisage.rpt"
        objReport.Load(strReport)
        objReport.SetParameterValue("IdFactures", tabIds.ToArray())
        objReport.SetParameterValue("bEntete", ckEntete.Checked)
        'End If
        'If rbReleve.Checked Then
        '    Dim r2 As New crReleveColisage()
        '    Dim strReportName2 As String = r2.ResourceName
        '    objReport.Load(PATHTOREPORTS & strReportName2)
        '    objReport.SetParameterValue("IDFACT", getElementCourant().id)
        '    objReport.SetParameterValue("BENTETE", ckEntete.Checked)
        '    objReport.SetParameterValue("BDETAIL", False)
        'End If



        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport
        restoreCursor()

    End Sub
    Private Sub AfficheEtat()
        Debug.Assert(getElementCourant() IsNot Nothing)
        debAffiche()
        Me.laEtatCommande.Text = getElementCourant().etat.libelle
        laEtatCommande.Tag = getElementCourant().etat.codeEtat
        cbAnnExport.Visible = (getElementCourant.etat.codeEtat = vncEtatCommande.vncFactCOLExportee)
        finAffiche()
    End Sub 'AfficheEtat

    '***********************************************************************************************************************************************************************
#End Region 'Méthodes Privées
#Region "Gestion Evenements"
    'Affichage de l'entete de la liste des lignes
    Private Sub cbAjouterLigne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAjouterLigne.Click
        If ajouteruneLigne() Then
            recalcul()
            setfrmUpdated()
        End If
    End Sub

    Private Sub liTiers_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liTiers.LinkClicked
        afficheFenetreFournisseur(liTiers.Tag)
    End Sub 'liNomClient_LinkClicked



    Private Sub cbSupprimerLigne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSupprimerLigne.Click
        If Not m_bsrcLgFact Is Nothing Then
            If supprimeruneLigne() Then
                recalcul()
                setfrmUpdated()
            End If
        End If

    End Sub



    Private Sub cbRecalcTotaux_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecalcTotaux.Click
        recalcul()
    End Sub

    Private Sub recalcul()
        Debug.Assert(Not getElementCourant() Is Nothing, "Facture courante non renseignée")
        MAJElement()
        getElementCourant().calculPrixTotal()
        affichecolLignes()
        m_bsrcFactColisage.ResetCurrentItem()
        setfrmUpdated()
    End Sub


    Private Sub frmGestFactTRP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub
    Private Sub dbAfficherEtat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficherEtat.Click
        afficherRapport()
    End Sub




#End Region



    Private Sub frmGestFactTRP_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If isfrmUpdated() Then
            If getElementCourant().bcolLignesUpated Then
                MsgBox("Si vous avez ajouté ou supprimé des lignes, la sauvegarde de la facture est automatique")
                getElementCourant().save()
            End If
        End If
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub


    Private Sub ModifieuneLigne()
        Dim odlg As dlgLgFactColisage
        Dim objLg As LgFactColisage
        Try
            objLg = m_bsrcLgFact.Current
            odlg = New dlgLgFactColisage
            odlg.setElementCourant(objLg)
            odlg.setTiersCourant(getElementCourant().oTiers)
            If odlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
                affichecolLignes()
            End If

        Catch
        End Try

    End Sub


    Private Sub dbAnnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnExport.Click
        getElementCourant().changeEtat(vncActionEtatCommande.vncActionFactCOLAnnExporter)
        AfficheEtat()
        setfrmUpdated()
    End Sub

    Private Sub cbBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBrowse.Click
        If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.tbDossierStockage.Text = FolderBrowserDialog1.SelectedPath
        End If

    End Sub

    Private Sub cbStocker_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStocker.Click
        StockerFactCol()
        AfficheEtat()
    End Sub
    '==================================================================================================
    'Function : StockerFactCom
    'Description : Socke les Trois états dans le répertoire spécifié
    '===================================================================================================
    Private Sub StockerFactCol()
        Dim objReport As ReportDocument
        Dim diskOpts As DiskFileDestinationOptions
        Try
            setcursorWait()
            DisplayStatus("Exportation du courrier")
            objReport = New ReportDocument
            diskOpts = New DiskFileDestinationOptions
            DisplayStatus("Exportation de la facture")
            Dim r1 As New crFactureColisage()
            Dim strReportName As String = r1.ResourceName
            objReport.Load(PATHTOREPORTS & strReportName)
            objReport.SetParameterValue("IDFACT", getElementCourant().id)
            objReport.SetParameterValue("BENTETE", ckEntete.Checked)
            objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
            objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            diskOpts.DiskFileName = tbDossierStockage.Text & "\FCOL" & Format(CInt(getElementCourant().code), "000000") & ".doc"
            objReport.ExportOptions.DestinationOptions = diskOpts
            Persist.setReportConnection(objReport)
            objReport.Export()
            objReport.Close()

            DisplayStatus("Exportation du relevé")
            Dim r2 As New crReleveColisage()
            Dim strReportName2 As String = r2.ResourceName
            objReport = New ReportDocument
            diskOpts = New DiskFileDestinationOptions
            objReport.Load(PATHTOREPORTS & strReportName2)
            objReport.SetParameterValue("IDFACT", getElementCourant().id)
            objReport.SetParameterValue("BENTETE", ckEntete.Checked)
            objReport.SetParameterValue("BDETAIL", False)
            objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
            objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            diskOpts.DiskFileName = tbDossierStockage.Text & "\FCOL" & Format(CInt(getElementCourant().code), "000000") & "_ReleveColisage.doc"
            objReport.ExportOptions.DestinationOptions = diskOpts
            Persist.setReportConnection(objReport)
            objReport.Export()
            objReport.Close()

            DisplayStatus("")

            setfrmUpdated()

            restoreCursor()
        Catch ex As Exception
            MsgBox("erreur en exportation de Facture de commission, essayer en manuel" & ex.ToString)
        End Try
    End Sub

    Private Sub cbReglement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbReglement.Click
        Dim odlg As dlgReglements
        odlg = New dlgReglements
        odlg.setFact(Me.getElementCourant())
        odlg.ShowDialog(Me)

    End Sub
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = True
        m_ToolBarLoadEnabled = True
        If isfrmUpdated() Then
            m_ToolBarSaveEnabled = True
        Else
            m_ToolBarSaveEnabled = False
        End If
        m_ToolBarDelEnabled = True
        m_ToolBarRefreshEnabled = True

    End Sub
    Protected Overrides Function isfrmUpdated() As Boolean
        If getElementCourant() Is Nothing Then
            Return False
        End If
        If getElementCourant().etat.codeEtat = vncEtatCommande.vncFactComExportee Then
            Return False
        Else
            Return MyBase.isfrmUpdated()
        End If

    End Function

    Private Sub DataGridView1_RowHeaderMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseDoubleClick
        ModifieuneLigne()
    End Sub
End Class
