Imports vini_DB
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
'Imports FAXCOMLib
Imports System.Windows.Forms.Cursors

Public Class frmGestFactCom
    Inherits frmDonBase
    Private Enum vncColLigneCommande
        COL_NUM = 0
        COL_CODECMDCLIENT = 1
        '        COL_DATECMDCLIENT = 1
        COL_RSCLIENT = 2
        COL_CODEFACTFOURN = 3
        COL_DATEFACTFOURN = 4
        COL_MONTANTFACTFOURN = 5
        COL_BASECOMMISSION = 6
        COL_MONTANTCOMMISSION = 7
        COL_CODESCMD = 8
        COL_NBRECOL = 9
    End Enum
    'Private getElementCourant() As FactCom

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
    Friend WithEvents tbTotalReglements As System.Windows.Forms.GroupBox
    Friend WithEvents dtDateFact As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtDateStat As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbPeriode As System.Windows.Forms.TextBox
    Friend WithEvents liFournisseur As System.Windows.Forms.LinkLabel
    Friend WithEvents cbBrowse As System.Windows.Forms.Button
    Friend WithEvents tbDossierStockage As System.Windows.Forms.TextBox
    Friend WithEvents cbStocker As System.Windows.Forms.Button
    Friend WithEvents obReleve As System.Windows.Forms.RadioButton
    Friend WithEvents obFacture As System.Windows.Forms.RadioButton
    Friend WithEvents obCourrier As System.Windows.Forms.RadioButton
    Friend WithEvents cbAfficherEtat As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents cbRecalcTotaux As System.Windows.Forms.Button
    Friend WithEvents cbAnnExport As System.Windows.Forms.Button
    Friend WithEvents cbReglement As System.Windows.Forms.Button
    Friend WithEvents tbSolder As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents m_bsModeReglement As System.Windows.Forms.BindingSource
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtEcheance As System.Windows.Forms.DateTimePicker
    Friend WithEvents m_bsrcFacture As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcLgFactures As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LGFACTSCMDRSClientDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents tbSolde As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents ckEntete As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.tbTotalReglements = New System.Windows.Forms.GroupBox()
        Me.tbPeriode = New System.Windows.Forms.TextBox()
        Me.m_bsrcFacture = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtDateStat = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtDateFact = New System.Windows.Forms.DateTimePicker()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.liFournisseur = New System.Windows.Forms.LinkLabel()
        Me.laEtatCommande = New System.Windows.Forms.Label()
        Me.tbCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbAnnExport = New System.Windows.Forms.Button()
        Me.SSTabCommandeClient = New System.Windows.Forms.TabControl()
        Me.tpLignes = New System.Windows.Forms.TabPage()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LGFACTSCMDRSClientDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_bsrcLgFactures = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label6 = New System.Windows.Forms.Label()
        Me.dtEcheance = New System.Windows.Forms.DateTimePicker()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.m_bsModeReglement = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label5 = New System.Windows.Forms.Label()
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
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.ckEntete = New System.Windows.Forms.CheckBox()
        Me.cbAfficherEtat = New System.Windows.Forms.Button()
        Me.cbBrowse = New System.Windows.Forms.Button()
        Me.tbDossierStockage = New System.Windows.Forms.TextBox()
        Me.cbStocker = New System.Windows.Forms.Button()
        Me.obReleve = New System.Windows.Forms.RadioButton()
        Me.obFacture = New System.Windows.Forms.RadioButton()
        Me.obCourrier = New System.Windows.Forms.RadioButton()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.cbReglement = New System.Windows.Forms.Button()
        Me.tbSolder = New System.Windows.Forms.Button()
        Me.tbSolde = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbTotalReglements.SuspendLayout()
        CType(Me.m_bsrcFacture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSTabCommandeClient.SuspendLayout()
        Me.tpLignes.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcLgFactures, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsModeReglement, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpCommentaires.SuspendLayout()
        Me.tpValidCmd.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbTotalReglements
        '
        Me.tbTotalReglements.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTotalReglements.Controls.Add(Me.tbPeriode)
        Me.tbTotalReglements.Controls.Add(Me.Label3)
        Me.tbTotalReglements.Controls.Add(Me.dtDateStat)
        Me.tbTotalReglements.Controls.Add(Me.Label2)
        Me.tbTotalReglements.Controls.Add(Me.dtDateFact)
        Me.tbTotalReglements.Controls.Add(Me.Label29)
        Me.tbTotalReglements.Controls.Add(Me.liFournisseur)
        Me.tbTotalReglements.Controls.Add(Me.laEtatCommande)
        Me.tbTotalReglements.Controls.Add(Me.tbCode)
        Me.tbTotalReglements.Controls.Add(Me.Label1)
        Me.tbTotalReglements.Location = New System.Drawing.Point(8, 0)
        Me.tbTotalReglements.Name = "tbTotalReglements"
        Me.tbTotalReglements.Size = New System.Drawing.Size(992, 62)
        Me.tbTotalReglements.TabIndex = 0
        Me.tbTotalReglements.TabStop = False
        Me.tbTotalReglements.Text = "Facture"
        '
        'tbPeriode
        '
        Me.tbPeriode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPeriode.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "periode", True))
        Me.tbPeriode.Location = New System.Drawing.Point(656, 16)
        Me.tbPeriode.Name = "tbPeriode"
        Me.tbPeriode.Size = New System.Drawing.Size(173, 20)
        Me.tbPeriode.TabIndex = 3
        '
        'm_bsrcFacture
        '
        Me.m_bsrcFacture.DataSource = GetType(vini_DB.FactCom)
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(590, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 16)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Période"
        '
        'dtDateStat
        '
        Me.dtDateStat.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "dateStatistique", True))
        Me.dtDateStat.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateStat.Location = New System.Drawing.Point(496, 16)
        Me.dtDateStat.Name = "dtDateStat"
        Me.dtDateStat.Size = New System.Drawing.Size(88, 20)
        Me.dtDateStat.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(390, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 16)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Date statistique"
        '
        'dtDateFact
        '
        Me.dtDateFact.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "dateFacture", True))
        Me.dtDateFact.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFact.Location = New System.Drawing.Point(280, 16)
        Me.dtDateFact.Name = "dtDateFact"
        Me.dtDateFact.Size = New System.Drawing.Size(104, 20)
        Me.dtDateFact.TabIndex = 1
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(170, 19)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(104, 16)
        Me.Label29.TabIndex = 12
        Me.Label29.Text = "Date de facture"
        '
        'liFournisseur
        '
        Me.liFournisseur.Location = New System.Drawing.Point(8, 40)
        Me.liFournisseur.Name = "liFournisseur"
        Me.liFournisseur.Size = New System.Drawing.Size(712, 16)
        Me.liFournisseur.TabIndex = 8
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
        Me.laEtatCommande.Text = "Etat"
        '
        'tbCode
        '
        Me.tbCode.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "code", True))
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
        Me.cbAnnExport.Location = New System.Drawing.Point(872, 68)
        Me.cbAnnExport.Name = "cbAnnExport"
        Me.cbAnnExport.Size = New System.Drawing.Size(124, 23)
        Me.cbAnnExport.TabIndex = 2
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
        Me.SSTabCommandeClient.Location = New System.Drawing.Point(0, 73)
        Me.SSTabCommandeClient.Name = "SSTabCommandeClient"
        Me.SSTabCommandeClient.SelectedIndex = 0
        Me.SSTabCommandeClient.Size = New System.Drawing.Size(1000, 496)
        Me.SSTabCommandeClient.TabIndex = 0
        '
        'tpLignes
        '
        Me.tpLignes.Controls.Add(Me.DataGridView1)
        Me.tpLignes.Controls.Add(Me.Label6)
        Me.tpLignes.Controls.Add(Me.dtEcheance)
        Me.tpLignes.Controls.Add(Me.ComboBox1)
        Me.tpLignes.Controls.Add(Me.Label5)
        Me.tpLignes.Controls.Add(Me.cbRecalcTotaux)
        Me.tpLignes.Controls.Add(Me.tbTotalTTC)
        Me.tpLignes.Controls.Add(Me.Label51)
        Me.tpLignes.Controls.Add(Me.cbSupprimerLigne)
        Me.tpLignes.Controls.Add(Me.cbAjouterLigne)
        Me.tpLignes.Controls.Add(Me.tbTotalHT)
        Me.tpLignes.Controls.Add(Me.Label10)
        Me.tpLignes.Location = New System.Drawing.Point(4, 22)
        Me.tpLignes.Name = "tpLignes"
        Me.tpLignes.Size = New System.Drawing.Size(992, 470)
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
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn, Me.LGFACTSCMDRSClientDataGridViewTextBoxColumn, Me.LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn, Me.LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn, Me.LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn, Me.LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn, Me.LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcLgFactures
        Me.DataGridView1.Location = New System.Drawing.Point(4, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(972, 381)
        Me.DataGridView1.TabIndex = 0
        '
        'LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn
        '
        Me.LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn.DataPropertyName = "LGFACT_SCMD_CodeCommandeClient"
        Me.LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn.HeaderText = "Cmd"
        Me.LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn.Name = "LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn"
        Me.LGFACTSCMDCodeCommandeClientDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LGFACTSCMDRSClientDataGridViewTextBoxColumn
        '
        Me.LGFACTSCMDRSClientDataGridViewTextBoxColumn.DataPropertyName = "LGFACT__SCMD_RSClient"
        Me.LGFACTSCMDRSClientDataGridViewTextBoxColumn.FillWeight = 300.0!
        Me.LGFACTSCMDRSClientDataGridViewTextBoxColumn.HeaderText = "Client"
        Me.LGFACTSCMDRSClientDataGridViewTextBoxColumn.Name = "LGFACTSCMDRSClientDataGridViewTextBoxColumn"
        Me.LGFACTSCMDRSClientDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn
        '
        Me.LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn.DataPropertyName = "LGFACT__SCMD_CodeFactFournisseur"
        Me.LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn.HeaderText = "Fact"
        Me.LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn.Name = "LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn"
        Me.LGFACTSCMDCodeFactFournisseurDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn
        '
        Me.LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn.DataPropertyName = "LGFACT__SCMD_DateFactFournisseur"
        Me.LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn.HeaderText = "Date Fact"
        Me.LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn.Name = "LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn"
        Me.LGFACTSCMDDateFactFournisseurDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn
        '
        Me.LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn.DataPropertyName = "LGFACT__SCMD_MontantFactFournisseur"
        DataGridViewCellStyle1.Format = "C2"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle1
        Me.LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn.HeaderText = "HT Fact"
        Me.LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn.Name = "LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn"
        Me.LGFACTSCMDMontantFactFournisseurDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn
        '
        Me.LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn.DataPropertyName = "LGFACT__SCMD_BaseCommission"
        DataGridViewCellStyle2.Format = "C2"
        Me.LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle2
        Me.LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn.HeaderText = "Base Comm"
        Me.LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn.Name = "LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn"
        Me.LGFACTSCMDBaseCommissionDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn
        '
        Me.LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn.DataPropertyName = "LGFACT__SCMD_MontantCommission"
        DataGridViewCellStyle3.Format = "C2"
        Me.LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle3
        Me.LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn.HeaderText = "Montant Comm"
        Me.LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn.Name = "LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn"
        Me.LGFACTSCMDMontantCommissionDataGridViewTextBoxColumn.ReadOnly = True
        '
        'm_bsrcLgFactures
        '
        Me.m_bsrcLgFactures.DataSource = GetType(vini_DB.LgCommande)
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 436)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(87, 13)
        Me.Label6.TabIndex = 66
        Me.Label6.Text = "Date écheance :"
        '
        'dtEcheance
        '
        Me.dtEcheance.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtEcheance.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "dEcheance", True))
        Me.dtEcheance.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtEcheance.Location = New System.Drawing.Point(122, 431)
        Me.dtEcheance.Name = "dtEcheance"
        Me.dtEcheance.Size = New System.Drawing.Size(93, 20)
        Me.dtEcheance.TabIndex = 2
        '
        'ComboBox1
        '
        Me.ComboBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ComboBox1.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcFacture, "idModeReglement", True))
        Me.ComboBox1.DataSource = Me.m_bsModeReglement
        Me.ComboBox1.DisplayMember = "valeur"
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(122, 407)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(266, 21)
        Me.ComboBox1.TabIndex = 1
        Me.ComboBox1.ValueMember = "id"
        '
        'm_bsModeReglement
        '
        Me.m_bsModeReglement.DataSource = GetType(vini_DB.Param)
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 410)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(104, 13)
        Me.Label5.TabIndex = 63
        Me.Label5.Text = "Mode de réglement :"
        '
        'cbRecalcTotaux
        '
        Me.cbRecalcTotaux.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbRecalcTotaux.Location = New System.Drawing.Point(712, 443)
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
        Me.tbTotalTTC.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "totalTTC", True))
        Me.tbTotalTTC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbTotalTTC.Location = New System.Drawing.Point(863, 447)
        Me.tbTotalTTC.MaxLength = 0
        Me.tbTotalTTC.Name = "tbTotalTTC"
        Me.tbTotalTTC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbTotalTTC.Size = New System.Drawing.Size(113, 20)
        Me.tbTotalTTC.TabIndex = 7
        Me.tbTotalTTC.Text = "0"
        '
        'Label51
        '
        Me.Label51.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label51.Location = New System.Drawing.Point(793, 450)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(64, 16)
        Me.Label51.TabIndex = 62
        Me.Label51.Text = "Total TTC"
        '
        'cbSupprimerLigne
        '
        Me.cbSupprimerLigne.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSupprimerLigne.BackColor = System.Drawing.SystemColors.Control
        Me.cbSupprimerLigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbSupprimerLigne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbSupprimerLigne.Location = New System.Drawing.Point(903, 390)
        Me.cbSupprimerLigne.Name = "cbSupprimerLigne"
        Me.cbSupprimerLigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbSupprimerLigne.Size = New System.Drawing.Size(73, 25)
        Me.cbSupprimerLigne.TabIndex = 4
        Me.cbSupprimerLigne.Text = "&Supprimer"
        Me.cbSupprimerLigne.UseVisualStyleBackColor = False
        '
        'cbAjouterLigne
        '
        Me.cbAjouterLigne.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAjouterLigne.BackColor = System.Drawing.SystemColors.Control
        Me.cbAjouterLigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAjouterLigne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAjouterLigne.Location = New System.Drawing.Point(743, 390)
        Me.cbAjouterLigne.Name = "cbAjouterLigne"
        Me.cbAjouterLigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAjouterLigne.Size = New System.Drawing.Size(73, 25)
        Me.cbAjouterLigne.TabIndex = 3
        Me.cbAjouterLigne.Text = "A&jouter"
        Me.cbAjouterLigne.UseVisualStyleBackColor = False
        '
        'tbTotalHT
        '
        Me.tbTotalHT.AcceptsReturn = True
        Me.tbTotalHT.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTotalHT.BackColor = System.Drawing.SystemColors.Window
        Me.tbTotalHT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbTotalHT.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "totalHT", True))
        Me.tbTotalHT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbTotalHT.Location = New System.Drawing.Point(863, 421)
        Me.tbTotalHT.MaxLength = 0
        Me.tbTotalHT.Name = "tbTotalHT"
        Me.tbTotalHT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbTotalHT.Size = New System.Drawing.Size(113, 20)
        Me.tbTotalHT.TabIndex = 6
        Me.tbTotalHT.Text = "0"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(793, 425)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(65, 17)
        Me.Label10.TabIndex = 60
        Me.Label10.Text = "Total HT"
        '
        'tpCommentaires
        '
        Me.tpCommentaires.Controls.Add(Me.tbCommentaireFacturation)
        Me.tpCommentaires.Controls.Add(Me.Label26)
        Me.tpCommentaires.Location = New System.Drawing.Point(4, 22)
        Me.tpCommentaires.Name = "tpCommentaires"
        Me.tpCommentaires.Size = New System.Drawing.Size(992, 470)
        Me.tpCommentaires.TabIndex = 3
        Me.tpCommentaires.Text = "Commentaires"
        Me.tpCommentaires.Visible = False
        '
        'tbCommentaireFacturation
        '
        Me.tbCommentaireFacturation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbCommentaireFacturation.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "CommentaireFacturationText", True))
        Me.tbCommentaireFacturation.Location = New System.Drawing.Point(96, 8)
        Me.tbCommentaireFacturation.Multiline = True
        Me.tbCommentaireFacturation.Name = "tbCommentaireFacturation"
        Me.tbCommentaireFacturation.Size = New System.Drawing.Size(892, 120)
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
        Me.tpValidCmd.Controls.Add(Me.ckEntete)
        Me.tpValidCmd.Controls.Add(Me.cbAfficherEtat)
        Me.tpValidCmd.Controls.Add(Me.cbBrowse)
        Me.tpValidCmd.Controls.Add(Me.tbDossierStockage)
        Me.tpValidCmd.Controls.Add(Me.cbStocker)
        Me.tpValidCmd.Controls.Add(Me.obReleve)
        Me.tpValidCmd.Controls.Add(Me.obFacture)
        Me.tpValidCmd.Controls.Add(Me.obCourrier)
        Me.tpValidCmd.Location = New System.Drawing.Point(4, 22)
        Me.tpValidCmd.Name = "tpValidCmd"
        Me.tpValidCmd.Size = New System.Drawing.Size(992, 470)
        Me.tpValidCmd.TabIndex = 4
        Me.tpValidCmd.Text = "Editions"
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
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(9, 39)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(979, 428)
        Me.CrystalReportViewer1.TabIndex = 129
        Me.CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'ckEntete
        '
        Me.ckEntete.Location = New System.Drawing.Point(208, 8)
        Me.ckEntete.Name = "ckEntete"
        Me.ckEntete.Size = New System.Drawing.Size(88, 24)
        Me.ckEntete.TabIndex = 128
        Me.ckEntete.Text = "Entete"
        '
        'cbAfficherEtat
        '
        Me.cbAfficherEtat.Location = New System.Drawing.Point(304, 8)
        Me.cbAfficherEtat.Name = "cbAfficherEtat"
        Me.cbAfficherEtat.Size = New System.Drawing.Size(80, 24)
        Me.cbAfficherEtat.TabIndex = 127
        Me.cbAfficherEtat.Text = "Afficher"
        '
        'cbBrowse
        '
        Me.cbBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbBrowse.Location = New System.Drawing.Point(826, 7)
        Me.cbBrowse.Name = "cbBrowse"
        Me.cbBrowse.Size = New System.Drawing.Size(40, 24)
        Me.cbBrowse.TabIndex = 126
        Me.cbBrowse.Text = "..."
        '
        'tbDossierStockage
        '
        Me.tbDossierStockage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbDossierStockage.Location = New System.Drawing.Point(400, 8)
        Me.tbDossierStockage.Name = "tbDossierStockage"
        Me.tbDossierStockage.Size = New System.Drawing.Size(420, 20)
        Me.tbDossierStockage.TabIndex = 125
        '
        'cbStocker
        '
        Me.cbStocker.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbStocker.Location = New System.Drawing.Point(872, 7)
        Me.cbStocker.Name = "cbStocker"
        Me.cbStocker.Size = New System.Drawing.Size(104, 24)
        Me.cbStocker.TabIndex = 124
        Me.cbStocker.Text = "Stocker"
        '
        'obReleve
        '
        Me.obReleve.Location = New System.Drawing.Point(144, 8)
        Me.obReleve.Name = "obReleve"
        Me.obReleve.Size = New System.Drawing.Size(64, 24)
        Me.obReleve.TabIndex = 123
        Me.obReleve.Text = "Releve"
        '
        'obFacture
        '
        Me.obFacture.Location = New System.Drawing.Point(80, 8)
        Me.obFacture.Name = "obFacture"
        Me.obFacture.Size = New System.Drawing.Size(64, 24)
        Me.obFacture.TabIndex = 122
        Me.obFacture.Text = "Facture"
        '
        'obCourrier
        '
        Me.obCourrier.Location = New System.Drawing.Point(16, 8)
        Me.obCourrier.Name = "obCourrier"
        Me.obCourrier.Size = New System.Drawing.Size(64, 24)
        Me.obCourrier.TabIndex = 121
        Me.obCourrier.Text = "Courrier"
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
        Me.cbReglement.Location = New System.Drawing.Point(710, 68)
        Me.cbReglement.Name = "cbReglement"
        Me.cbReglement.Size = New System.Drawing.Size(75, 23)
        Me.cbReglement.TabIndex = 0
        Me.cbReglement.Text = "Réglements"
        Me.cbReglement.UseVisualStyleBackColor = True
        '
        'tbSolder
        '
        Me.tbSolder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbSolder.Location = New System.Drawing.Point(791, 68)
        Me.tbSolder.Name = "tbSolder"
        Me.tbSolder.Size = New System.Drawing.Size(75, 23)
        Me.tbSolder.TabIndex = 1
        Me.tbSolder.Text = "Solder"
        Me.tbSolder.UseVisualStyleBackColor = True
        '
        'tbSolde
        '
        Me.tbSolde.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbSolde.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFacture, "montantReglement", True))
        Me.tbSolde.Location = New System.Drawing.Point(629, 70)
        Me.tbSolde.Name = "tbSolde"
        Me.tbSolde.Size = New System.Drawing.Size(75, 20)
        Me.tbSolde.TabIndex = 21
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(527, 73)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 13)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Total Réglements :"
        '
        'frmGestFactCom
        '
        Me.ClientSize = New System.Drawing.Size(1016, 577)
        Me.Controls.Add(Me.tbSolder)
        Me.Controls.Add(Me.tbSolde)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbReglement)
        Me.Controls.Add(Me.cbAnnExport)
        Me.Controls.Add(Me.SSTabCommandeClient)
        Me.Controls.Add(Me.tbTotalReglements)
        Me.Name = "frmGestFactCom"
        Me.Text = "Facture de commission"
        Me.tbTotalReglements.ResumeLayout(False)
        Me.tbTotalReglements.PerformLayout()
        CType(Me.m_bsrcFacture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSTabCommandeClient.ResumeLayout(False)
        Me.tpLignes.ResumeLayout(False)
        Me.tpLignes.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcLgFactures, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsModeReglement, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpCommentaires.ResumeLayout(False)
        Me.tpCommentaires.PerformLayout()
        Me.tpValidCmd.ResumeLayout(False)
        Me.tpValidCmd.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
#Region "Méthodes Vinicom"
    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)
        Me.tbCode.Enabled = False
    End Sub
    Public Overrides Function setElementCourant2(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        bReturn = MyBase.setElementCourant2(pElement)
        If Not getElementCourant() Is Nothing And bReturn Then
            m_bsrcFacture.Clear()
            m_bsrcFacture.Add(pElement)
            'Desactivation des possibilités de modification si la facture est exportée
            cbAjouterLigne.Enabled = getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactComExportee
            cbSupprimerLigne.Enabled = getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactComExportee
            dtDateFact.Enabled = getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactComExportee
            dtDateStat.Enabled = getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactComExportee
            tbPeriode.Enabled = getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactComExportee
            tbTotalHT.Enabled = getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactComExportee
            tbTotalTTC.Enabled = getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactComExportee
            cbRecalcTotaux.Enabled = getElementCourant().etat.codeEtat <> vncEtatCommande.vncFactComExportee
        End If
        setToolbarButtons()
        Return bReturn
    End Function 'SetElement_Specifique

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

    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenetre n'est pas libre")
        Dim bReturn As Boolean
        bReturn = setElementCourant2(New FactCom(New Client("", "")))

        Return bReturn
    End Function
    Protected Shadows Function getElementCourant() As FactCom
        Return CType(m_ElementCourant, FactCom)
    End Function

    'Interface ElementCourant -> Ecran
    Public Overrides Function AfficheElement() As Boolean


        Debug.Assert(Not getElementCourant() Is Nothing)

        'Affichage des caractéristiques de la commande (Binding)
        m_bsrcFacture.ResetBindings(False)

        'Etat de la commande
        AfficheEtat()

        'Fournisseur
        liFournisseur.Text = getElementCourant().oTiers.shortResume
        liFournisseur.Tag = getElementCourant().oTiers.id



        'Lignes de Commandes
        affichecolLignes()

        CrystalReportViewer1.ReportSource = Nothing
        CrystalReportViewer1.Refresh()

        Return True
    End Function 'AfficheElement
    Public Overrides Function MAJElement() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Dim bReturn As Boolean

        bReturn = True

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
        str = "FCT"
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
        MyBase.frmDel()
        Dim bReturn As Boolean
        bReturn = False
        If getElementCourant() Is Nothing Then
            bReturn = setElementCourant2(Nothing)
        End If
        Return bReturn
    End Function ' frmDel

    Protected Overrides Sub setfrmUpdated()
        If Not getElementCourant() Is Nothing Then
            MAJElementCourant()
            MyBase.setfrmUpdated()
        End If
    End Sub
    Public Overrides Function SauveElement() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Dim bReturn As Boolean
        bReturn = getElementCourant().Save
        If bReturn Then
            tbCode.Text = getElementCourant().code
            MyBase.getResume()
        End If
        Debug.Assert(bReturn, "Erreur en sauvegarde")
        Return bReturn
    End Function
    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        MyBase.AddHandlerValidated(ocol)
        RemoveHandler tbDossierStockage.Validated, AddressOf ControlUpdated
        RemoveHandler cbStocker.Click, AddressOf ControlUpdated
        RemoveHandler cbAfficherEtat.Click, AddressOf ControlUpdated
        RemoveHandler obCourrier.CheckedChanged, AddressOf ControlUpdated
        RemoveHandler obFacture.CheckedChanged, AddressOf ControlUpdated
        RemoveHandler obReleve.CheckedChanged, AddressOf ControlUpdated
    End Sub


    '==============================================================================================
    '==============================================================================================
    '==============================================================================================
#End Region
#Region "Methodes Privées"
    Protected Sub initFenetre()
        debAffiche()
        m_TypeDonnees = vncEnums.vncTypeDonnee.FACTCOMM
        tbDossierStockage.Text = Param.getConstante("CST_PATH_FACTCOM")
        CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        CrystalReportViewer1.DisplayToolbar = True
        Dim objModeReglement As Param
        For Each objModeReglement In Param.colModeReglement
            m_bsModeReglement.Add(objModeReglement)
        Next



        finAffiche()
    End Sub 'initFenetre

    'Affiche la boite de dialogue d'ajout de ligne
    Protected Function ajouteruneLigne() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Debug.Assert(Not getElementCourant().oTiers Is Nothing)
        Dim bReturn As Boolean
        Dim odlg As dlgLgFactCom
        Dim objLg As LgCommande
        Dim objSCMD As SousCommande

        Me.sauvegardeElementCourant()

        objLg = New LgCommande(getElementCourant().id)
        objLg.num = getElementCourant().getNextNumLg()
        odlg = New dlgLgFactCom
        odlg.setElementCourant(objLg)
        odlg.setFournisseurCourant(getElementCourant().oTiers)
        If odlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If objLg.idSCmd <> 0 Then
                objSCMD = SousCommande.createandload(objLg.idSCmd)
                objLg = getElementCourant().AjouteLigneFactCom(objSCMD, True)
                m_bsrcLgFactures.Add(objLg)
            End If
        End If
        bReturn = True

        Return bReturn

    End Function 'Ajouteruneligne
    'Affiche la boite de dialogue de modification de ligne
    Protected Function supprimeruneLigne() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Dim objLg As LgCommande
        Dim bReturn As Boolean
        Dim objSCMD As SousCommande

        Try
            objLg = m_bsrcLgFactures.Current
            objSCMD = SousCommande.createandload(objLg.idSCmd)
            If objSCMD.id <> 0 Then
                If MsgBox("Etes-vous sur de vouloir supprimer la ligne " & objSCMD.shortResume, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDAnnFacturer)
                    objSCMD.idFactCom = 0
                    objSCMD.Save()
                    getElementCourant().loadcolLignes()
                End If
                setfrmUpdated()
                bReturn = True
            End If
        Catch ex As Exception
            bReturn = False
            DisplayError("frmGestFactCom.SupprimerUneLigne", "Impossible de charger la sousCommande")
        End Try
        getElementCourant().calculPrixTotal()
        debAffiche()
        affichecolLignes()
        finAffiche()


        bReturn = True

        Return bReturn
    End Function 'Ajouteruneligne
    Protected Function affichecolLignes() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Debug.Assert(Not getElementCourant().colLignes Is Nothing)

        Dim objLg As LgCommande
        Dim nRow As Integer
        debAffiche()
        nRow = 0
        m_bsrcLgFactures.Clear()
        For Each objLg In getElementCourant().colLignes
            nRow = nRow + 1
            m_bsrcLgFactures.Add(objLg)
        Next objLg
        m_bsrcLgFactures.ResetBindings(False)
        finAffiche()
    End Function ' Affichage de des lignes de commandes

    Private Sub afficherRapport()
        Dim strReport As String = ""
        Dim objReport As ReportDocument
        Dim tabIds As ArrayList



        'Choix de l'état
        If obCourrier.Checked Then
            strReport = PATHTOREPORTS & "crFactCom_courrier.rpt"
        End If
        If obFacture.Checked Then
            strReport = PATHTOREPORTS & "crFactCom.rpt"
        End If
        If obReleve.Checked Then
            strReport = PATHTOREPORTS & "crFactCom_Releve.rpt"
        End If

        If strReport = "" Then
            Exit Sub
        End If
        setcursorWait()
        objReport = New ReportDocument
        objReport.Load(strReport)

        tabIds = New ArrayList()
        tabIds.Add(getElementCourant().id)
        objReport.SetParameterValue("IdFactures", tabIds.ToArray())
        objReport.SetParameterValue("bEntete", ckEntete.Checked)

        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport

        restoreCursor()

    End Sub
    '==================================================================================================
    'Function : StockerFactCom
    'Description : Socke les Trois états dans le répertoire spécifié
    '===================================================================================================
    Private Sub StockerFactCom()
        Dim objReport As ReportDocument
        Dim diskOpts As DiskFileDestinationOptions
        Dim tabIds As ArrayList
        Try
            setcursorWait()
            tabIds = New ArrayList()
            tabIds.Add(getElementCourant().id)
            DisplayStatus("Exportation du courrier")
            objReport = New ReportDocument
            diskOpts = New DiskFileDestinationOptions
            objReport.Load(PATHTOREPORTS & "crFactCom_courrier.rpt")
            objReport.SetParameterValue("IdFacture", tabIds.ToArray())
            objReport.SetParameterValue("bEntete", True)
            objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
            objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            If Not System.IO.Directory.Exists(tbDossierStockage.Text) Then
                System.IO.Directory.CreateDirectory(tbDossierStockage.Text)
            End If
            diskOpts.DiskFileName = tbDossierStockage.Text & "\F" & Format(CInt(getElementCourant().code), "000000") & "_Courrier.doc"
            objReport.ExportOptions.DestinationOptions = diskOpts
            Persist.setReportConnection(objReport)
            objReport.Export()
            objReport.Close()

            DisplayStatus("Exportation de la facture")
            objReport = New ReportDocument
            diskOpts = New DiskFileDestinationOptions
            objReport.Load(PATHTOREPORTS & "crFactCom.rpt")
            objReport.SetParameterValue("IdFacture", tabIds.ToArray())
            objReport.SetParameterValue("bEntete", True)
            objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
            objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            diskOpts.DiskFileName = tbDossierStockage.Text & "\F" & Format(CInt(getElementCourant().code), "000000") & ".doc"
            objReport.ExportOptions.DestinationOptions = diskOpts
            Persist.setReportConnection(objReport)
            objReport.Export()
            objReport.Close()

            DisplayStatus("Exportation du relevé")
            objReport = New ReportDocument
            diskOpts = New DiskFileDestinationOptions
            objReport.Load(PATHTOREPORTS & "crReleveCommission.rpt")
            objReport.SetParameterValue("IdFacture", tabIds.ToArray())
            objReport.SetParameterValue("bEntete", True)
            objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.Excel
            objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            diskOpts.DiskFileName = tbDossierStockage.Text & "\F" & Format(CInt(getElementCourant().code), "000000") & "_ReleveCommission.xls"
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
    Private Sub AfficheEtat()
        debAffiche()
        Me.laEtatCommande.Text = getElementCourant().etat.libelle
        laEtatCommande.Tag = getElementCourant().etat.codeEtat


        cbAnnExport.Visible = (getElementCourant.etat.codeEtat = vncEtatCommande.vncFactComExportee)
        finAffiche()
    End Sub 'AfficheEtat

    '***********************************************************************************************************************************************************************
#End Region 'Méthodes Privées
#Region "Gestion Evenements"
    'Affichage de l'entete de la liste des lignes
    Private Sub cbAjouterLigne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAjouterLigne.Click
        If ajouteruneLigne() Then
            getElementCourant().calculPrixTotal()
            m_bsrcFacture.ResetCurrentItem()
            setfrmUpdated()
        End If
    End Sub

    Private Sub liFournisseur_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liFournisseur.LinkClicked
        afficheFenetreFournisseur(liFournisseur.Tag)
    End Sub 'liNomClient_LinkClicked



    Private Sub cbSupprimerLigne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSupprimerLigne.Click
        If Not m_bsrcLgFactures.Current Is Nothing Then
            If supprimeruneLigne() Then
                getElementCourant().calculPrixTotal()
                m_bsrcFacture.ResetCurrentItem()
                setfrmUpdated()
            End If
        End If

    End Sub



    Private Sub cbRecalcTotaux_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecalcTotaux.Click
        Debug.Assert(Not getElementCourant() Is Nothing, "Facture courante non renseignée")
        getElementCourant().calculPrixTotal()
        m_bsrcFacture.ResetBindings(False)
        setfrmUpdated()
    End Sub


    Private Sub frmFestFactCom_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub
    Private Sub dbAfficherEtat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficherEtat.Click
        afficherRapport()
    End Sub

    Private Sub cbBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBrowse.Click
        If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.tbDossierStockage.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub cbStocker_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStocker.Click
        StockerFactCom()
        AfficheEtat()
    End Sub


#End Region



    Private Sub frmGestFactCom_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        'If Not CrystalReportViewer1.ReportSource Is Nothing Then
        '    CrystalReportViewer1.ReportSource.Dispose()
        'End If
    End Sub

    Private Sub cbAnnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnExport.Click
        getElementCourant().changeEtat(vncActionEtatCommande.vncActionFactComAnnExporter)
        setfrmUpdated()
        AfficheEtat()
    End Sub

    Private Sub cbReglement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbReglement.Click
        Dim odlg As dlgReglements
        odlg = New dlgReglements
        odlg.setFact(Me.getElementCourant())
        odlg.ShowDialog(Me)
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


End Class
