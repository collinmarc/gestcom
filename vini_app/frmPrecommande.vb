Imports vini_DB
Public Class frmPrecommande
    Inherits vini_app.frmDonBase


    Private Const COL_CODEPRODUIT As Integer = 0
    Private Const COL_DESIGNATIONPRODUIT As Integer = 1
    Private Const COL_MILLESIME As Integer = 2
    Private Const COL_CONDITIONNEMENT As Integer = 3
    Private Const COL_QTE_HAB As Integer = 4
    Private Const COL_QTE_DERN As Integer = 5
    Private Const COL_PRIX As Integer = 6
    Private Const COL_IDPRODUIT As Integer = 7
    Private Const COL_DATEDERNCMD As Integer = 8
    Private Const COL_NBRECOL As Integer = 9

    Private m_produit As Produit
    Friend WithEvents m_bsrcLgPrecom As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcClient As System.Windows.Forms.BindingSource
    Private m_objClientCourant As Client
    Friend WithEvents CodeProduitDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibProduitDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MillesimeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibConditionnementDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteHabDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteDernDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PrixUDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateDerniereCommandeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cbExporter As System.Windows.Forms.Button


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
    Friend WithEvents tbNom As System.Windows.Forms.TextBox
    Public WithEvents tbCode As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents tbAdrLivNom As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Public WithEvents tbAdrLivPortable As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivRue1 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivRue2 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivCP As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivVille As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivTel As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivFax As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivEmail As System.Windows.Forms.TextBox
    Friend WithEvents cbVisualiser As System.Windows.Forms.Button
    Friend WithEvents cbPCAjouter As System.Windows.Forms.Button
    Friend WithEvents cbPCModifier As System.Windows.Forms.Button
    Friend WithEvents cbPCSuppr As System.Windows.Forms.Button
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents cbRechProduit As System.Windows.Forms.Button
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpLgPreCom As System.Windows.Forms.GroupBox
    Friend WithEvents laLibProduit As System.Windows.Forms.LinkLabel
    Friend WithEvents laCodeProduit As System.Windows.Forms.LinkLabel
    Friend WithEvents tbQteDern As textBoxNumeric
    Friend WithEvents tbQteHab As textBoxNumeric
    Friend WithEvents laMillesimeProduit As System.Windows.Forms.Label
    Friend WithEvents laCondProduit As System.Windows.Forms.Label
    Friend WithEvents tbCodeProduit As System.Windows.Forms.TextBox
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MnuAjouterUneLigne As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSupprimerUneLigne As System.Windows.Forms.MenuItem
    Friend WithEvents mnuVoir As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModifierUneLigne As System.Windows.Forms.MenuItem
    Friend WithEvents grpAdresse As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbPrixU As textBoxCurrency
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtDateDernCmd As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbReinit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.tbNom = New System.Windows.Forms.TextBox
        Me.m_bsrcClient = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.grpAdresse = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.tbAdrLivNom = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.tbAdrLivPortable = New System.Windows.Forms.TextBox
        Me.tbAdrLivRue1 = New System.Windows.Forms.TextBox
        Me.tbAdrLivRue2 = New System.Windows.Forms.TextBox
        Me.tbAdrLivCP = New System.Windows.Forms.TextBox
        Me.tbAdrLivVille = New System.Windows.Forms.TextBox
        Me.tbAdrLivTel = New System.Windows.Forms.TextBox
        Me.tbAdrLivFax = New System.Windows.Forms.TextBox
        Me.tbAdrLivEmail = New System.Windows.Forms.TextBox
        Me.cbVisualiser = New System.Windows.Forms.Button
        Me.cbPCAjouter = New System.Windows.Forms.Button
        Me.cbPCModifier = New System.Windows.Forms.Button
        Me.cbPCSuppr = New System.Windows.Forms.Button
        Me.grpLgPreCom = New System.Windows.Forms.GroupBox
        Me.dtDateDernCmd = New System.Windows.Forms.DateTimePicker
        Me.m_bsrcLgPrecom = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbPrixU = New vini_app.textBoxCurrency
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbCodeProduit = New System.Windows.Forms.TextBox
        Me.laCondProduit = New System.Windows.Forms.Label
        Me.laMillesimeProduit = New System.Windows.Forms.Label
        Me.tbQteDern = New vini_app.textBoxNumeric
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbQteHab = New vini_app.textBoxNumeric
        Me.Label30 = New System.Windows.Forms.Label
        Me.laLibProduit = New System.Windows.Forms.LinkLabel
        Me.laCodeProduit = New System.Windows.Forms.LinkLabel
        Me.cbRechProduit = New System.Windows.Forms.Button
        Me.Label27 = New System.Windows.Forms.Label
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MnuAjouterUneLigne = New System.Windows.Forms.MenuItem
        Me.mnuModifierUneLigne = New System.Windows.Forms.MenuItem
        Me.mnuSupprimerUneLigne = New System.Windows.Forms.MenuItem
        Me.mnuVoir = New System.Windows.Forms.MenuItem
        Me.cbReinit = New System.Windows.Forms.Button
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.CodeProduitDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibProduitDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MillesimeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibConditionnementDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QteHabDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QteDernDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PrixUDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateDerniereCommandeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.cbExporter = New System.Windows.Forms.Button
        CType(Me.m_bsrcClient, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAdresse.SuspendLayout()
        Me.grpLgPreCom.SuspendLayout()
        CType(Me.m_bsrcLgPrecom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbNom
        '
        Me.tbNom.BackColor = System.Drawing.SystemColors.Window
        Me.tbNom.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "nom", True))
        Me.tbNom.Enabled = False
        Me.tbNom.Location = New System.Drawing.Point(232, 8)
        Me.tbNom.Name = "tbNom"
        Me.tbNom.Size = New System.Drawing.Size(640, 20)
        Me.tbNom.TabIndex = 1
        '
        'm_bsrcClient
        '
        Me.m_bsrcClient.DataSource = GetType(vini_DB.Client)
        '
        'tbCode
        '
        Me.tbCode.AcceptsReturn = True
        Me.tbCode.BackColor = System.Drawing.SystemColors.Window
        Me.tbCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCode.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "code", True))
        Me.tbCode.Enabled = False
        Me.tbCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCode.Location = New System.Drawing.Point(88, 8)
        Me.tbCode.MaxLength = 0
        Me.tbCode.Name = "tbCode"
        Me.tbCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCode.Size = New System.Drawing.Size(129, 20)
        Me.tbCode.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(0, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(57, 17)
        Me.Label2.TabIndex = 71
        Me.Label2.Text = "Code"
        '
        'grpAdresse
        '
        Me.grpAdresse.Controls.Add(Me.Label3)
        Me.grpAdresse.Controls.Add(Me.Label7)
        Me.grpAdresse.Controls.Add(Me.Label8)
        Me.grpAdresse.Controls.Add(Me.Label9)
        Me.grpAdresse.Controls.Add(Me.Label15)
        Me.grpAdresse.Controls.Add(Me.Label20)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivNom)
        Me.grpAdresse.Controls.Add(Me.Label21)
        Me.grpAdresse.Controls.Add(Me.Label22)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivPortable)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivRue1)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivRue2)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivCP)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivVille)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivTel)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivFax)
        Me.grpAdresse.Controls.Add(Me.tbAdrLivEmail)
        Me.grpAdresse.Location = New System.Drawing.Point(8, 32)
        Me.grpAdresse.Name = "grpAdresse"
        Me.grpAdresse.Size = New System.Drawing.Size(928, 120)
        Me.grpAdresse.TabIndex = 2
        Me.grpAdresse.TabStop = False
        Me.grpAdresse.Text = "Adresse de Livraison"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(528, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 97
        Me.Label3.Text = "Email"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(528, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 96
        Me.Label7.Text = "Portable"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(528, 40)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(80, 16)
        Me.Label8.TabIndex = 95
        Me.Label8.Text = "Fax"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(528, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(80, 16)
        Me.Label9.TabIndex = 94
        Me.Label9.Text = "Téléphone"
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 88)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(80, 16)
        Me.Label15.TabIndex = 93
        Me.Label15.Text = "CP/Ville"
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(8, 16)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(80, 24)
        Me.Label20.TabIndex = 92
        Me.Label20.Text = "Nom"
        '
        'tbAdrLivNom
        '
        Me.tbAdrLivNom.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "AdresseLivraisonNom", True))
        Me.tbAdrLivNom.Enabled = False
        Me.tbAdrLivNom.Location = New System.Drawing.Point(96, 16)
        Me.tbAdrLivNom.Name = "tbAdrLivNom"
        Me.tbAdrLivNom.Size = New System.Drawing.Size(416, 20)
        Me.tbAdrLivNom.TabIndex = 0
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(8, 64)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(72, 16)
        Me.Label21.TabIndex = 90
        Me.Label21.Text = "Adresse2"
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(8, 40)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(56, 16)
        Me.Label22.TabIndex = 89
        Me.Label22.Text = "Adresse1"
        '
        'tbAdrLivPortable
        '
        Me.tbAdrLivPortable.AcceptsReturn = True
        Me.tbAdrLivPortable.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivPortable.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivPortable.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "AdresseLivraisonPort", True))
        Me.tbAdrLivPortable.Enabled = False
        Me.tbAdrLivPortable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivPortable.Location = New System.Drawing.Point(648, 64)
        Me.tbAdrLivPortable.MaxLength = 0
        Me.tbAdrLivPortable.Name = "tbAdrLivPortable"
        Me.tbAdrLivPortable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivPortable.Size = New System.Drawing.Size(128, 20)
        Me.tbAdrLivPortable.TabIndex = 7
        '
        'tbAdrLivRue1
        '
        Me.tbAdrLivRue1.AcceptsReturn = True
        Me.tbAdrLivRue1.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivRue1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivRue1.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "AdresseLivraisonRue1", True))
        Me.tbAdrLivRue1.Enabled = False
        Me.tbAdrLivRue1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivRue1.Location = New System.Drawing.Point(96, 40)
        Me.tbAdrLivRue1.MaxLength = 0
        Me.tbAdrLivRue1.Name = "tbAdrLivRue1"
        Me.tbAdrLivRue1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivRue1.Size = New System.Drawing.Size(417, 20)
        Me.tbAdrLivRue1.TabIndex = 1
        '
        'tbAdrLivRue2
        '
        Me.tbAdrLivRue2.AcceptsReturn = True
        Me.tbAdrLivRue2.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivRue2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivRue2.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "AdresseLivraisonRue2", True))
        Me.tbAdrLivRue2.Enabled = False
        Me.tbAdrLivRue2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivRue2.Location = New System.Drawing.Point(96, 64)
        Me.tbAdrLivRue2.MaxLength = 0
        Me.tbAdrLivRue2.Name = "tbAdrLivRue2"
        Me.tbAdrLivRue2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivRue2.Size = New System.Drawing.Size(417, 20)
        Me.tbAdrLivRue2.TabIndex = 2
        '
        'tbAdrLivCP
        '
        Me.tbAdrLivCP.AcceptsReturn = True
        Me.tbAdrLivCP.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivCP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivCP.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "AdresseLivraisonCP", True))
        Me.tbAdrLivCP.Enabled = False
        Me.tbAdrLivCP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivCP.Location = New System.Drawing.Point(96, 88)
        Me.tbAdrLivCP.MaxLength = 0
        Me.tbAdrLivCP.Name = "tbAdrLivCP"
        Me.tbAdrLivCP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivCP.Size = New System.Drawing.Size(73, 20)
        Me.tbAdrLivCP.TabIndex = 3
        '
        'tbAdrLivVille
        '
        Me.tbAdrLivVille.AcceptsReturn = True
        Me.tbAdrLivVille.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivVille.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivVille.DataBindings.Add(New System.Windows.Forms.Binding("Tag", Me.m_bsrcClient, "AdresseLivraisonVille", True))
        Me.tbAdrLivVille.Enabled = False
        Me.tbAdrLivVille.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivVille.Location = New System.Drawing.Point(176, 88)
        Me.tbAdrLivVille.MaxLength = 0
        Me.tbAdrLivVille.Name = "tbAdrLivVille"
        Me.tbAdrLivVille.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivVille.Size = New System.Drawing.Size(337, 20)
        Me.tbAdrLivVille.TabIndex = 4
        '
        'tbAdrLivTel
        '
        Me.tbAdrLivTel.AcceptsReturn = True
        Me.tbAdrLivTel.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivTel.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivTel.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "AdresseLivraisonTel", True))
        Me.tbAdrLivTel.Enabled = False
        Me.tbAdrLivTel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivTel.Location = New System.Drawing.Point(648, 16)
        Me.tbAdrLivTel.MaxLength = 0
        Me.tbAdrLivTel.Name = "tbAdrLivTel"
        Me.tbAdrLivTel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivTel.Size = New System.Drawing.Size(129, 20)
        Me.tbAdrLivTel.TabIndex = 5
        '
        'tbAdrLivFax
        '
        Me.tbAdrLivFax.AcceptsReturn = True
        Me.tbAdrLivFax.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivFax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivFax.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "AdresseLivraisonFax", True))
        Me.tbAdrLivFax.Enabled = False
        Me.tbAdrLivFax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivFax.Location = New System.Drawing.Point(648, 40)
        Me.tbAdrLivFax.MaxLength = 0
        Me.tbAdrLivFax.Name = "tbAdrLivFax"
        Me.tbAdrLivFax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivFax.Size = New System.Drawing.Size(129, 20)
        Me.tbAdrLivFax.TabIndex = 6
        '
        'tbAdrLivEmail
        '
        Me.tbAdrLivEmail.AcceptsReturn = True
        Me.tbAdrLivEmail.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivEmail.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivEmail.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcClient, "AdresseLivraisonEmail", True))
        Me.tbAdrLivEmail.Enabled = False
        Me.tbAdrLivEmail.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivEmail.Location = New System.Drawing.Point(648, 88)
        Me.tbAdrLivEmail.MaxLength = 0
        Me.tbAdrLivEmail.Name = "tbAdrLivEmail"
        Me.tbAdrLivEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivEmail.Size = New System.Drawing.Size(201, 20)
        Me.tbAdrLivEmail.TabIndex = 8
        '
        'cbVisualiser
        '
        Me.cbVisualiser.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbVisualiser.Location = New System.Drawing.Point(824, 464)
        Me.cbVisualiser.Name = "cbVisualiser"
        Me.cbVisualiser.Size = New System.Drawing.Size(112, 24)
        Me.cbVisualiser.TabIndex = 4
        Me.cbVisualiser.Text = "V&oir"
        '
        'cbPCAjouter
        '
        Me.cbPCAjouter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbPCAjouter.Location = New System.Drawing.Point(632, 416)
        Me.cbPCAjouter.Name = "cbPCAjouter"
        Me.cbPCAjouter.Size = New System.Drawing.Size(96, 24)
        Me.cbPCAjouter.TabIndex = 5
        Me.cbPCAjouter.Text = "A&jouter"
        '
        'cbPCModifier
        '
        Me.cbPCModifier.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbPCModifier.Location = New System.Drawing.Point(744, 416)
        Me.cbPCModifier.Name = "cbPCModifier"
        Me.cbPCModifier.Size = New System.Drawing.Size(96, 24)
        Me.cbPCModifier.TabIndex = 6
        Me.cbPCModifier.Text = "&Modifier"
        '
        'cbPCSuppr
        '
        Me.cbPCSuppr.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbPCSuppr.Location = New System.Drawing.Point(856, 416)
        Me.cbPCSuppr.Name = "cbPCSuppr"
        Me.cbPCSuppr.Size = New System.Drawing.Size(80, 24)
        Me.cbPCSuppr.TabIndex = 7
        Me.cbPCSuppr.Text = "&Supprimer"
        '
        'grpLgPreCom
        '
        Me.grpLgPreCom.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpLgPreCom.Controls.Add(Me.dtDateDernCmd)
        Me.grpLgPreCom.Controls.Add(Me.Label5)
        Me.grpLgPreCom.Controls.Add(Me.tbPrixU)
        Me.grpLgPreCom.Controls.Add(Me.Label4)
        Me.grpLgPreCom.Controls.Add(Me.tbCodeProduit)
        Me.grpLgPreCom.Controls.Add(Me.laCondProduit)
        Me.grpLgPreCom.Controls.Add(Me.laMillesimeProduit)
        Me.grpLgPreCom.Controls.Add(Me.tbQteDern)
        Me.grpLgPreCom.Controls.Add(Me.Label1)
        Me.grpLgPreCom.Controls.Add(Me.tbQteHab)
        Me.grpLgPreCom.Controls.Add(Me.Label30)
        Me.grpLgPreCom.Controls.Add(Me.laLibProduit)
        Me.grpLgPreCom.Controls.Add(Me.laCodeProduit)
        Me.grpLgPreCom.Controls.Add(Me.cbRechProduit)
        Me.grpLgPreCom.Controls.Add(Me.Label27)
        Me.grpLgPreCom.Location = New System.Drawing.Point(8, 448)
        Me.grpLgPreCom.Name = "grpLgPreCom"
        Me.grpLgPreCom.Size = New System.Drawing.Size(804, 128)
        Me.grpLgPreCom.TabIndex = 8
        Me.grpLgPreCom.TabStop = False
        '
        'dtDateDernCmd
        '
        Me.dtDateDernCmd.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgPrecom, "dateDerniereCommande", True))
        Me.dtDateDernCmd.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateDernCmd.Location = New System.Drawing.Point(328, 48)
        Me.dtDateDernCmd.Name = "dtDateDernCmd"
        Me.dtDateDernCmd.Size = New System.Drawing.Size(96, 20)
        Me.dtDateDernCmd.TabIndex = 3
        '
        'm_bsrcLgPrecom
        '
        Me.m_bsrcLgPrecom.DataSource = GetType(vini_DB.lgPrecomm)
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(240, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 23)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Date Dern Cmd"
        '
        'tbPrixU
        '
        Me.tbPrixU.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgPrecom, "prixU", True))
        Me.tbPrixU.Location = New System.Drawing.Point(328, 96)
        Me.tbPrixU.Name = "tbPrixU"
        Me.tbPrixU.Size = New System.Drawing.Size(96, 20)
        Me.tbPrixU.TabIndex = 5
        Me.tbPrixU.Text = "0"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(240, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 24)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "PrixU Dern Cmd"
        '
        'tbCodeProduit
        '
        Me.tbCodeProduit.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgPrecom, "codeProduit", True))
        Me.tbCodeProduit.Location = New System.Drawing.Point(104, 16)
        Me.tbCodeProduit.Name = "tbCodeProduit"
        Me.tbCodeProduit.Size = New System.Drawing.Size(144, 20)
        Me.tbCodeProduit.TabIndex = 0
        '
        'laCondProduit
        '
        Me.laCondProduit.Location = New System.Drawing.Point(800, 16)
        Me.laCondProduit.Name = "laCondProduit"
        Me.laCondProduit.Size = New System.Drawing.Size(56, 24)
        Me.laCondProduit.TabIndex = 16
        '
        'laMillesimeProduit
        '
        Me.laMillesimeProduit.Location = New System.Drawing.Point(712, 16)
        Me.laMillesimeProduit.Name = "laMillesimeProduit"
        Me.laMillesimeProduit.Size = New System.Drawing.Size(80, 24)
        Me.laMillesimeProduit.TabIndex = 15
        '
        'tbQteDern
        '
        Me.tbQteDern.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgPrecom, "qteDern", True))
        Me.tbQteDern.Location = New System.Drawing.Point(328, 72)
        Me.tbQteDern.Name = "tbQteDern"
        Me.tbQteDern.Size = New System.Drawing.Size(96, 20)
        Me.tbQteDern.TabIndex = 4
        Me.tbQteDern.Text = "0"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(240, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 24)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Qte Dern Cmd"
        '
        'tbQteHab
        '
        Me.tbQteHab.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgPrecom, "qteHab", True))
        Me.tbQteHab.Location = New System.Drawing.Point(104, 56)
        Me.tbQteHab.Name = "tbQteHab"
        Me.tbQteHab.Size = New System.Drawing.Size(96, 20)
        Me.tbQteHab.TabIndex = 2
        Me.tbQteHab.Text = "0"
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(8, 59)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(88, 24)
        Me.Label30.TabIndex = 9
        Me.Label30.Text = "Quantité Habituelle"
        '
        'laLibProduit
        '
        Me.laLibProduit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.laLibProduit.AutoSize = True
        Me.laLibProduit.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgPrecom, "libProduit", True))
        Me.laLibProduit.Location = New System.Drawing.Point(464, 16)
        Me.laLibProduit.Name = "laLibProduit"
        Me.laLibProduit.Size = New System.Drawing.Size(83, 13)
        Me.laLibProduit.TabIndex = 7
        Me.laLibProduit.TabStop = True
        Me.laLibProduit.Text = "libellé du produit"
        '
        'laCodeProduit
        '
        Me.laCodeProduit.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgPrecom, "codeProduit", True))
        Me.laCodeProduit.Location = New System.Drawing.Point(384, 16)
        Me.laCodeProduit.Name = "laCodeProduit"
        Me.laCodeProduit.Size = New System.Drawing.Size(72, 24)
        Me.laCodeProduit.TabIndex = 6
        Me.laCodeProduit.TabStop = True
        Me.laCodeProduit.Text = "Code Produit"
        '
        'cbRechProduit
        '
        Me.cbRechProduit.Location = New System.Drawing.Point(256, 16)
        Me.cbRechProduit.Name = "cbRechProduit"
        Me.cbRechProduit.Size = New System.Drawing.Size(80, 24)
        Me.cbRechProduit.TabIndex = 1
        Me.cbRechProduit.Text = "Rechercher"
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(8, 16)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(88, 16)
        Me.Label27.TabIndex = 0
        Me.Label27.Text = "Code produit"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuAjouterUneLigne, Me.mnuModifierUneLigne, Me.mnuSupprimerUneLigne, Me.mnuVoir})
        '
        'MnuAjouterUneLigne
        '
        Me.MnuAjouterUneLigne.Index = 0
        Me.MnuAjouterUneLigne.Text = "A&jouter une ligne"
        '
        'mnuModifierUneLigne
        '
        Me.mnuModifierUneLigne.Index = 1
        Me.mnuModifierUneLigne.Text = "&Modifier une ligne"
        '
        'mnuSupprimerUneLigne
        '
        Me.mnuSupprimerUneLigne.Index = 2
        Me.mnuSupprimerUneLigne.Text = "&Supprimer une ligne"
        '
        'mnuVoir
        '
        Me.mnuVoir.Index = 3
        Me.mnuVoir.Text = "&Voir le bon de précommande"
        '
        'cbReinit
        '
        Me.cbReinit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbReinit.Location = New System.Drawing.Point(824, 501)
        Me.cbReinit.Name = "cbReinit"
        Me.cbReinit.Size = New System.Drawing.Size(112, 23)
        Me.cbReinit.TabIndex = 72
        Me.cbReinit.Text = "&Réinitialiser"
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
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeProduitDataGridViewTextBoxColumn, Me.LibProduitDataGridViewTextBoxColumn, Me.MillesimeDataGridViewTextBoxColumn, Me.LibConditionnementDataGridViewTextBoxColumn, Me.QteHabDataGridViewTextBoxColumn, Me.QteDernDataGridViewTextBoxColumn, Me.PrixUDataGridViewTextBoxColumn, Me.DateDerniereCommandeDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcLgPrecom
        Me.DataGridView1.Location = New System.Drawing.Point(8, 158)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(928, 252)
        Me.DataGridView1.TabIndex = 73
        '
        'CodeProduitDataGridViewTextBoxColumn
        '
        Me.CodeProduitDataGridViewTextBoxColumn.DataPropertyName = "codeProduit"
        Me.CodeProduitDataGridViewTextBoxColumn.FillWeight = 2.0!
        Me.CodeProduitDataGridViewTextBoxColumn.HeaderText = "code Produit"
        Me.CodeProduitDataGridViewTextBoxColumn.Name = "CodeProduitDataGridViewTextBoxColumn"
        Me.CodeProduitDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LibProduitDataGridViewTextBoxColumn
        '
        Me.LibProduitDataGridViewTextBoxColumn.DataPropertyName = "libProduit"
        Me.LibProduitDataGridViewTextBoxColumn.FillWeight = 14.0!
        Me.LibProduitDataGridViewTextBoxColumn.HeaderText = "Désignation Produit"
        Me.LibProduitDataGridViewTextBoxColumn.Name = "LibProduitDataGridViewTextBoxColumn"
        Me.LibProduitDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MillesimeDataGridViewTextBoxColumn
        '
        Me.MillesimeDataGridViewTextBoxColumn.DataPropertyName = "millesime"
        Me.MillesimeDataGridViewTextBoxColumn.FillWeight = 2.0!
        Me.MillesimeDataGridViewTextBoxColumn.HeaderText = "Mill."
        Me.MillesimeDataGridViewTextBoxColumn.Name = "MillesimeDataGridViewTextBoxColumn"
        Me.MillesimeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LibConditionnementDataGridViewTextBoxColumn
        '
        Me.LibConditionnementDataGridViewTextBoxColumn.DataPropertyName = "libConditionnement"
        Me.LibConditionnementDataGridViewTextBoxColumn.FillWeight = 2.0!
        Me.LibConditionnementDataGridViewTextBoxColumn.HeaderText = "Cond."
        Me.LibConditionnementDataGridViewTextBoxColumn.Name = "LibConditionnementDataGridViewTextBoxColumn"
        Me.LibConditionnementDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QteHabDataGridViewTextBoxColumn
        '
        Me.QteHabDataGridViewTextBoxColumn.DataPropertyName = "qteHab"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.QteHabDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle4
        Me.QteHabDataGridViewTextBoxColumn.FillWeight = 4.0!
        Me.QteHabDataGridViewTextBoxColumn.HeaderText = "qte Hab"
        Me.QteHabDataGridViewTextBoxColumn.Name = "QteHabDataGridViewTextBoxColumn"
        Me.QteHabDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QteDernDataGridViewTextBoxColumn
        '
        Me.QteDernDataGridViewTextBoxColumn.DataPropertyName = "qteDern"
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.QteDernDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle5
        Me.QteDernDataGridViewTextBoxColumn.FillWeight = 4.0!
        Me.QteDernDataGridViewTextBoxColumn.HeaderText = "qteDern"
        Me.QteDernDataGridViewTextBoxColumn.Name = "QteDernDataGridViewTextBoxColumn"
        Me.QteDernDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PrixUDataGridViewTextBoxColumn
        '
        Me.PrixUDataGridViewTextBoxColumn.DataPropertyName = "prixU"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.Format = "C2"
        DataGridViewCellStyle6.NullValue = Nothing
        Me.PrixUDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle6
        Me.PrixUDataGridViewTextBoxColumn.FillWeight = 4.0!
        Me.PrixUDataGridViewTextBoxColumn.HeaderText = "prixU"
        Me.PrixUDataGridViewTextBoxColumn.Name = "PrixUDataGridViewTextBoxColumn"
        Me.PrixUDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DateDerniereCommandeDataGridViewTextBoxColumn
        '
        Me.DateDerniereCommandeDataGridViewTextBoxColumn.DataPropertyName = "dateDerniereCommande"
        Me.DateDerniereCommandeDataGridViewTextBoxColumn.FillWeight = 4.0!
        Me.DateDerniereCommandeDataGridViewTextBoxColumn.HeaderText = "Dern. Cmd"
        Me.DateDerniereCommandeDataGridViewTextBoxColumn.Name = "DateDerniereCommandeDataGridViewTextBoxColumn"
        Me.DateDerniereCommandeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'cbExporter
        '
        Me.cbExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbExporter.Location = New System.Drawing.Point(824, 539)
        Me.cbExporter.Name = "cbExporter"
        Me.cbExporter.Size = New System.Drawing.Size(112, 23)
        Me.cbExporter.TabIndex = 74
        Me.cbExporter.Text = "Exporter"
        Me.cbExporter.UseVisualStyleBackColor = True
        '
        'frmPrecommande
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(942, 583)
        Me.Controls.Add(Me.cbExporter)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.cbReinit)
        Me.Controls.Add(Me.cbVisualiser)
        Me.Controls.Add(Me.cbPCAjouter)
        Me.Controls.Add(Me.cbPCModifier)
        Me.Controls.Add(Me.cbPCSuppr)
        Me.Controls.Add(Me.grpLgPreCom)
        Me.Controls.Add(Me.grpAdresse)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbNom)
        Me.Controls.Add(Me.tbCode)
        Me.Name = "frmPrecommande"
        Me.Text = "Pre-Commande"
        CType(Me.m_bsrcClient, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAdresse.ResumeLayout(False)
        Me.grpAdresse.PerformLayout()
        Me.grpLgPreCom.ResumeLayout(False)
        Me.grpLgPreCom.PerformLayout()
        CType(Me.m_bsrcLgPrecom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
#Region "Méthodes Vinicom"
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        'm_ToolBarSaveEnabled = False
        m_ToolBarDelEnabled = False
        'm_ToolBarRefreshEnabled = False

    End Sub

    Protected Overrides Function frmNew() As Boolean
        'No new Precommande Allowed
        Return False
    End Function
    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenetre n'est pas libre")
        Dim bReturn As Boolean
        bReturn = setElementCourant2(New Client("", ""))

        Return bReturn
    End Function
    Public Overrides Function setElementCourant2(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        Debug.Assert(pElement.GetType().Name.Equals("Client"))
        m_objClientCourant = CType(pElement, Client)
        bReturn = m_objClientCourant.LoadPreCommande()
        Debug.Assert(bReturn, "LoadPrecommande")
        bReturn = MyBase.setElementCourant2(pElement)
        Return bReturn
    End Function 'SetElement
    Public Overrides Function AfficheElement() As Boolean
        Debug.Assert(Not m_objClientCourant Is Nothing)

        Dim objLgPrecom As lgPrecomm
        Dim i As Integer


        'Affichage des caractéristiques du client
        m_bsrcClient.Clear()
        m_bsrcClient.Add(m_objClientCourant)

        'Affichage de la liste des lignes
        i = 0
        m_bsrcLgPrecom.Clear()
        For i = 1 To m_objClientCourant.getlgPrecomCount
            objLgPrecom = m_objClientCourant.oPrecommande.colLignes(i)
            m_bsrcLgPrecom.Add(objLgPrecom)
        Next i
        Return True
    End Function 'AfficheElement
    Public Overrides Function MAJElement() As Boolean
        Debug.Assert(Not m_objClientCourant Is Nothing)
        Return True
    End Function 'MAJElement
    Public Overrides Function sauveElement() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing, "Element courant isNothing")
        Debug.Assert(getElementCourant().GetType().Name.Equals("Client"), "Element courant de type Client requis")

        Dim bReturn As Boolean
        Dim objClient As Client
        objClient = getElementCourant()
        bReturn = objClient.savePrecommande()
        If bReturn Then
            'Mise à jour de l'état de la fenêtre
            setfrmNotUpdated()
        Else
            DisplayError("Erreur en Sauvegarde", Client.getErreur())
        End If

        Debug.Assert(Not getElementCourant() Is Nothing)
        Debug.Assert(Not isfrmUpdated()) ' La fenêtre doit être non modifiée
        Return bReturn

    End Function 'SaveElement
    Public Overrides Function ControleAvantSauvegarde() As Boolean
        Return True
    End Function
    Public Overrides Function getResume() As String 'Rend le caption de la fenêtre
        If getElementCourant() Is Nothing Then
            Return "Gestion de Précommande"
        Else
            Return "Précommande du client " & getElementCourant.shortResume
        End If
    End Function 'getResume
    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        tbCode.Enabled = False
        tbNom.Enabled = False
        grpAdresse.Enabled = False
        cbVisualiser.Enabled = bEnabled
        cbPCAjouter.Enabled = bEnabled
        cbPCModifier.Enabled = bEnabled
        cbPCSuppr.Enabled = bEnabled
        grpLgPreCom.Enabled = bEnabled
    End Sub
#End Region

    Private Sub initFenetre()
        m_TypeDonnees = vncEnums.vncTypeDonnee.CLIENT
    End Sub
    Private Sub cbVisualiser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVisualiser.Click, mnuVoir.Click
        Dim odlg As dlgVisuPrecommande

        If isfrmUpdated() Then
            If MsgBox("Voulez-vous sauvegarder l'élement courant", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                frmSave()
            End If
        End If
        odlg = New dlgVisuPrecommande
        odlg.Show()
        odlg.setClient(getElementCourant())
    End Sub
    Private Sub cbPCAjouter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPCAjouter.Click, MnuAjouterUneLigne.Click
        Dim oLG As lgPrecomm
        oLG = m_objClientCourant.ajouteLgPrecom()
        m_bsrcLgPrecom.Add(oLG)
        m_bsrcLgPrecom.MoveLast()
        tbCodeProduit.Focus()
    End Sub
    Private Sub cbPCModifier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPCModifier.Click, mnuModifierUneLigne.Click
    End Sub
    Private Sub cbRechProduit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbRechProduit.Click
        rechercheProduit()
    End Sub
    '=========================================================================
    'Fontion : CodeProduitExistantDansTableau
    'Description : Recherche la ligne contenant le code Produit dans le tableau
    'Retour : Rend le numéro de ligne dans le tableau ou -1 s'il n'existepas
    '========================================================================
    Private Function codeProduitExistantDansTableau(ByVal strCode As String) As Integer
        Dim nReturn As Integer
        Dim i As Integer
        Dim objLgPrecom As lgPrecomm

        nReturn = -1
        For i = 1 To m_objClientCourant.getlgPrecomCount
            objLgPrecom = m_objClientCourant.getLgPrecom(i)
            If objLgPrecom.codeProduit.Equals(strCode) Then
                nReturn = i
                Exit For
            End If
        Next i

        Return nReturn
    End Function 'codeProduitExistantDansTableau
     '==========================================================================
    ' Fonction : ActiverSaisieLigne
    ' Desciption : ActiverSaisieLigne
    '
    ' Retour
    '=========================================================================

    Private Sub cbPCSuppr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPCSuppr.Click, mnuSupprimerUneLigne.Click
        supprimerLigne()
    End Sub
    Private Function supprimerLigne() As Boolean
        Dim oLg As lgPrecomm
        Try
            If Not m_bsrcLgPrecom.Current Is Nothing Then
                oLg = m_bsrcLgPrecom.Current
                m_objClientCourant.supprimeLgPrecom(m_bsrcLgPrecom.Position + 1)
                m_bsrcLgPrecom.RemoveCurrent()
                oLg.bDeleted = True
            End If
            setfrmUpdated()
            Return True

        Catch ex As Exception
            Debug.Assert(False, ex.ToString())
            Return False
        End Try
    End Function 'SupprimerLigne

    Private Sub rechercheProduit()
        Dim colProduit As Collection
        Dim objProduit As Produit = Nothing
        Dim frm As frmRechercheDB
        Dim objLgPrecom As lgPrecomm

        If codeProduitExistantDansTableau(tbCodeProduit.Text) <> -1 Then
            MsgBox("Une ligne a déjà été saisie pour ce produit")
            Exit Sub
        End If

        If tbCodeProduit.Text <> "" Then
            colProduit = Produit.getListe(vncEnums.vncTypeProduit.vncTous, tbCodeProduit.Text)
        Else
            colProduit = New Collection
        End If
        If colProduit.Count <> 1 Then
            'Création de la fenêtre de recherche
            frm = New frmRechercheDB
            frm.setTypeDonnees(vncEnums.vncTypeDonnee.PRODUIT)
            frm.setListe(colProduit)
            frm.displayListe()
            'Affichage de la fenêtre
            If frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                'Si on sort par OK
                objProduit = frm.getElementSelectionne()
            End If
        Else
            objProduit = colProduit(1)
        End If
        If Not objProduit Is Nothing Then
            objLgPrecom = m_bsrcLgPrecom.Current
            objLgPrecom.idProduit = objProduit.id
            objLgPrecom.codeProduit = objProduit.code
            objLgPrecom.libProduit = objProduit.nom
            tbQteHab.Focus()
        End If

    End Sub 'RechercheProduit

    Private Sub tbCodeProduit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbCodeProduit.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            rechercheProduit()
        End If
    End Sub

    Private Sub laCodeProduit_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles laCodeProduit.LinkClicked
        afficheFenetreProduit(laCodeProduit.Tag)
    End Sub

    Private Sub laLibProduit_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles laLibProduit.LinkClicked
        afficheFenetreProduit(laCodeProduit.Tag)
    End Sub

    Private Sub frmPrecommande_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
        ' Cette fenêtre ne bloque pas l'élément courant
        m_BloquageElementCourant = False
    End Sub



    Private Sub cbReinit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbReinit.Click
        Debug.Assert(Not m_objClientCourant Is Nothing, "Client Courant non initialisé")
        If MsgBox("Etes-vous sur de vouloir réinitialiser la précommande d ce client", MsgBoxStyle.YesNo, "Réinitialisetion Précommande") = MsgBoxResult.Yes Then
            If m_objClientCourant.reinitPrecommande() Then
                AfficheElement()
                setfrmUpdated()
            End If

        End If
    End Sub

    Private Sub cbExporter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExporter.Click
        Call exporter()
    End Sub

    Private Sub exporter()

        Dim objPreCommande As preCommande
        Dim bExportOK As Boolean = True
        Dim bReturn As Boolean
        Dim strFile As String
        Dim strCSV As String
        Dim nFile As Integer
        Dim strFolder As String
        Dim oFTPvinicom As clsFTPVinicom

        'Suppression - creation du répertoire temporaire
        Try
            setcursorWait()
            DisplayStatus("Création du répertoire Temporaire utilisateur")
            strFolder = My.MySettings.Default.PreCommandeFolder + "/" + currentuser.code
            If System.IO.Directory.Exists(strFolder) Then
                System.IO.Directory.Delete(strFolder, True)
            End If
            System.IO.Directory.CreateDirectory(strFolder)

            strFile = strFolder + "/PC" + Format(Now(), "yyyyMMdd") + ".csv"

            'Génération des fichiers dans le répertoire temporaire
            nFile = FreeFile()
            FileOpen(nFile, strFile, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
            objPreCommande = m_objClientCourant.oPrecommande
            strCSV = objPreCommande.toCSV()
            Print(nFile, strCSV)
            FileClose(nFile)

            bReturn = True
            DisplayStatus("Transferts des fichiers ")
            'Exporter les fichiers générés
            oFTPvinicom = New clsFTPVinicom 'Création avec les paramètres par defaut
            'If oFTPvinicom.connect() Then
            If True Then
                oFTPvinicom.uploadFromDir(strFolder)
                DisplayStatus("Fin de transfert des fichiers ")
            Else
                DisplayStatus("Impossible de se connecter")
            End If



        Catch ex As Exception
            MsgBox("Erreur" + ex.Message)

        End Try
        Me.Cursor = Cursors.Default

    End Sub 'exporter

End Class
