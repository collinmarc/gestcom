Imports vini_DB
Public Class frmProduit
    Inherits frmDonBase
    Private Const COL_STKDATE As Integer = 0
    Private Const COL_STKQTE As Integer = 1
    Private Const COL_STKLIBELLE As Integer = 2
    Private Const COL_STKINDEX As Integer = 3
    Private Const COL_NBRECOL As Integer = 4

    Private m_objProduitCourant As Produit
    Private m_objmvtCourant As mvtStock
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents tbTarifA As vini_app.textBoxCurrency
    Friend WithEvents tbTarifB As vini_app.textBoxCurrency
    Friend WithEvents tbTarifC As vini_app.textBoxCurrency
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents m_bsrcProduit As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcRegion As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcCouleur As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcContenant As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcConditionnement As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcMvtStock As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcTypeMvt As System.Windows.Forms.BindingSource
    Friend WithEvents DatemvtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents typeMvt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibelleDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Private m_bAjoutmvt As Boolean

#Region "Code généré par le Concepteur Windows Form "
    Public Sub New()
        MyBase.New()
        m_TypeDonnees = vncEnums.vncTypeDonnee.PRODUIT
        InitializeComponent()
    End Sub
    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer
    Public WithEvents tbDesignation As System.Windows.Forms.TextBox
    Public WithEvents tbCode As System.Windows.Forms.TextBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    Public WithEvents tbMotDirecteur As System.Windows.Forms.TextBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents tbMillesime As textBoxNumeric
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cbRechercher As System.Windows.Forms.Button
    Friend WithEvents tbCodeFourn As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents tbCodeStat As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cbSTKSupprimer As System.Windows.Forms.Button
    Friend WithEvents cbStkModifier As System.Windows.Forms.Button
    Friend WithEvents cbAjouter As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cbCalculStock As System.Windows.Forms.Button
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cboConditionnement As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents liNomFournisseur As System.Windows.Forms.LinkLabel
    Friend WithEvents laDateDernInventaire As System.Windows.Forms.Label
    Friend WithEvents laQteStock As System.Windows.Forms.Label
    Friend WithEvents laQteDernierInventaire As System.Windows.Forms.Label
    Friend WithEvents laQteCommande As System.Windows.Forms.Label
    Friend WithEvents laQteStockTheorique As System.Windows.Forms.Label
    Public WithEvents cboCouleur As System.Windows.Forms.ComboBox
    Public WithEvents cboRegion As System.Windows.Forms.ComboBox
    Public WithEvents cboContenant As System.Windows.Forms.ComboBox
    Friend WithEvents ckStock As System.Windows.Forms.CheckBox
    Friend WithEvents ckDispo As System.Windows.Forms.CheckBox
    'Friend WithEvents flxLstMvtStock As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents cbAppliquer As System.Windows.Forms.Button
    Friend WithEvents tbmvtQuantite As textBoxNumeric
    Friend WithEvents cbomvtType As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents tbmvtLibelle As System.Windows.Forms.TextBox
    Friend WithEvents dtmvtDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents tbmvtCommentaiire As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.tbDesignation = New System.Windows.Forms.TextBox
        Me.m_bsrcProduit = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbCode = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbMotDirecteur = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbMillesime = New vini_app.textBoxNumeric
        Me.Label7 = New System.Windows.Forms.Label
        Me.cboCouleur = New System.Windows.Forms.ComboBox
        Me.m_bsrcCouleur = New System.Windows.Forms.BindingSource(Me.components)
        Me.cboRegion = New System.Windows.Forms.ComboBox
        Me.m_bsrcRegion = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboContenant = New System.Windows.Forms.ComboBox
        Me.m_bsrcContenant = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label8 = New System.Windows.Forms.Label
        Me.liNomFournisseur = New System.Windows.Forms.LinkLabel
        Me.cbRechercher = New System.Windows.Forms.Button
        Me.tbCodeFourn = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ckStock = New System.Windows.Forms.CheckBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.tbCodeStat = New System.Windows.Forms.TextBox
        Me.laDateDernInventaire = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.tbmvtLibelle = New System.Windows.Forms.TextBox
        Me.m_bsrcMvtStock = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.cbomvtType = New System.Windows.Forms.ComboBox
        Me.m_bsrcTypeMvt = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbmvtCommentaiire = New System.Windows.Forms.RichTextBox
        Me.cbAppliquer = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.tbmvtQuantite = New vini_app.textBoxNumeric
        Me.Label11 = New System.Windows.Forms.Label
        Me.dtmvtDate = New System.Windows.Forms.DateTimePicker
        Me.Label10 = New System.Windows.Forms.Label
        Me.cbSTKSupprimer = New System.Windows.Forms.Button
        Me.cbStkModifier = New System.Windows.Forms.Button
        Me.cbAjouter = New System.Windows.Forms.Button
        Me.laQteStock = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.cbCalculStock = New System.Windows.Forms.Button
        Me.Label16 = New System.Windows.Forms.Label
        Me.cboConditionnement = New System.Windows.Forms.ComboBox
        Me.m_bsrcConditionnement = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label17 = New System.Windows.Forms.Label
        Me.laQteDernierInventaire = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.laQteCommande = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.laQteStockTheorique = New System.Windows.Forms.Label
        Me.ckDispo = New System.Windows.Forms.CheckBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.tbTarifA = New vini_app.textBoxCurrency
        Me.tbTarifB = New vini_app.textBoxCurrency
        Me.tbTarifC = New vini_app.textBoxCurrency
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.DatemvtDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.typeMvt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibelleDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.m_bsrcProduit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcCouleur, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcRegion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcContenant, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.m_bsrcMvtStock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcTypeMvt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcConditionnement, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbDesignation
        '
        Me.tbDesignation.AcceptsReturn = True
        Me.tbDesignation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbDesignation.BackColor = System.Drawing.SystemColors.Window
        Me.tbDesignation.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbDesignation.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "nom", True))
        Me.tbDesignation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbDesignation.Location = New System.Drawing.Point(104, 56)
        Me.tbDesignation.MaxLength = 0
        Me.tbDesignation.Name = "tbDesignation"
        Me.tbDesignation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbDesignation.Size = New System.Drawing.Size(560, 20)
        Me.tbDesignation.TabIndex = 7
        '
        'm_bsrcProduit
        '
        Me.m_bsrcProduit.DataSource = GetType(vini_DB.Produit)
        '
        'tbCode
        '
        Me.tbCode.AcceptsReturn = True
        Me.tbCode.BackColor = System.Drawing.SystemColors.Window
        Me.tbCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCode.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "code", True))
        Me.tbCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCode.Location = New System.Drawing.Point(104, 8)
        Me.tbCode.MaxLength = 0
        Me.tbCode.Name = "tbCode"
        Me.tbCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCode.Size = New System.Drawing.Size(128, 20)
        Me.tbCode.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(8, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(81, 17)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Désignation"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(57, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Code"
        '
        'tbMotDirecteur
        '
        Me.tbMotDirecteur.AcceptsReturn = True
        Me.tbMotDirecteur.BackColor = System.Drawing.SystemColors.Window
        Me.tbMotDirecteur.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbMotDirecteur.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "motcle", True))
        Me.tbMotDirecteur.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbMotDirecteur.Location = New System.Drawing.Point(312, 8)
        Me.tbMotDirecteur.MaxLength = 0
        Me.tbMotDirecteur.Name = "tbMotDirecteur"
        Me.tbMotDirecteur.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbMotDirecteur.Size = New System.Drawing.Size(153, 20)
        Me.tbMotDirecteur.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(240, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(64, 17)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Mot clé"
        '
        'tbMillesime
        '
        Me.tbMillesime.AcceptsReturn = True
        Me.tbMillesime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMillesime.BackColor = System.Drawing.SystemColors.Window
        Me.tbMillesime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbMillesime.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "millesime", True))
        Me.tbMillesime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbMillesime.Location = New System.Drawing.Point(824, 56)
        Me.tbMillesime.MaxLength = 0
        Me.tbMillesime.Name = "tbMillesime"
        Me.tbMillesime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbMillesime.Size = New System.Drawing.Size(48, 20)
        Me.tbMillesime.TabIndex = 8
        Me.tbMillesime.Text = "0"
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(680, 56)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(56, 17)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Millésime"
        '
        'cboCouleur
        '
        Me.cboCouleur.BackColor = System.Drawing.SystemColors.Window
        Me.cboCouleur.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCouleur.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcProduit, "idCouleur", True))
        Me.cboCouleur.DataSource = Me.m_bsrcCouleur
        Me.cboCouleur.DisplayMember = "valeur"
        Me.cboCouleur.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCouleur.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCouleur.Location = New System.Drawing.Point(312, 32)
        Me.cboCouleur.Name = "cboCouleur"
        Me.cboCouleur.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCouleur.Size = New System.Drawing.Size(104, 21)
        Me.cboCouleur.TabIndex = 4
        Me.cboCouleur.ValueMember = "id"
        '
        'm_bsrcCouleur
        '
        Me.m_bsrcCouleur.DataSource = GetType(vini_DB.Param)
        '
        'cboRegion
        '
        Me.cboRegion.BackColor = System.Drawing.SystemColors.Window
        Me.cboRegion.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboRegion.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcProduit, "idRegion", True))
        Me.cboRegion.DataSource = Me.m_bsrcRegion
        Me.cboRegion.DisplayMember = "valeur"
        Me.cboRegion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRegion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboRegion.Location = New System.Drawing.Point(104, 32)
        Me.cboRegion.Name = "cboRegion"
        Me.cboRegion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboRegion.Size = New System.Drawing.Size(129, 21)
        Me.cboRegion.TabIndex = 3
        Me.cboRegion.ValueMember = "id"
        '
        'm_bsrcRegion
        '
        Me.m_bsrcRegion.DataSource = GetType(vini_DB.Param)
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(240, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(64, 17)
        Me.Label4.TabIndex = 32
        Me.Label4.Text = "Couleur"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(8, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(64, 17)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "Région"
        '
        'cboContenant
        '
        Me.cboContenant.BackColor = System.Drawing.SystemColors.Window
        Me.cboContenant.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboContenant.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcProduit, "idContenant", True))
        Me.cboContenant.DataSource = Me.m_bsrcContenant
        Me.cboContenant.DisplayMember = "libelle"
        Me.cboContenant.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboContenant.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboContenant.Location = New System.Drawing.Point(464, 32)
        Me.cboContenant.Name = "cboContenant"
        Me.cboContenant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboContenant.Size = New System.Drawing.Size(200, 21)
        Me.cboContenant.TabIndex = 5
        Me.cboContenant.ValueMember = "id"
        '
        'm_bsrcContenant
        '
        Me.m_bsrcContenant.DataSource = GetType(vini_DB.contenant)
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(424, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(32, 17)
        Me.Label8.TabIndex = 35
        Me.Label8.Text = "Cont"
        '
        'liNomFournisseur
        '
        Me.liNomFournisseur.Location = New System.Drawing.Point(400, 88)
        Me.liNomFournisseur.Name = "liNomFournisseur"
        Me.liNomFournisseur.Size = New System.Drawing.Size(264, 52)
        Me.liNomFournisseur.TabIndex = 11
        Me.liNomFournisseur.TabStop = True
        Me.liNomFournisseur.Text = "NomFournisseur"
        '
        'cbRechercher
        '
        Me.cbRechercher.Location = New System.Drawing.Point(216, 88)
        Me.cbRechercher.Name = "cbRechercher"
        Me.cbRechercher.Size = New System.Drawing.Size(72, 24)
        Me.cbRechercher.TabIndex = 10
        Me.cbRechercher.Text = "Rechercher"
        '
        'tbCodeFourn
        '
        Me.tbCodeFourn.Location = New System.Drawing.Point(104, 88)
        Me.tbCodeFourn.Name = "tbCodeFourn"
        Me.tbCodeFourn.Size = New System.Drawing.Size(104, 20)
        Me.tbCodeFourn.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "Fournisseur"
        '
        'ckStock
        '
        Me.ckStock.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckStock.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Me.m_bsrcProduit, "bStock", True))
        Me.ckStock.Location = New System.Drawing.Point(8, 120)
        Me.ckStock.Name = "ckStock"
        Me.ckStock.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ckStock.Size = New System.Drawing.Size(176, 24)
        Me.ckStock.TabIndex = 12
        Me.ckStock.Text = "Produit Plateforme"
        '
        'Label15
        '
        Me.Label15.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.Location = New System.Drawing.Point(600, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(128, 16)
        Me.Label15.TabIndex = 43
        Me.Label15.Text = "Code Statistique"
        '
        'tbCodeStat
        '
        Me.tbCodeStat.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbCodeStat.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "codeStat", True))
        Me.tbCodeStat.Location = New System.Drawing.Point(736, 8)
        Me.tbCodeStat.Name = "tbCodeStat"
        Me.tbCodeStat.Size = New System.Drawing.Size(136, 20)
        Me.tbCodeStat.TabIndex = 2
        '
        'laDateDernInventaire
        '
        Me.laDateDernInventaire.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.laDateDernInventaire.ForeColor = System.Drawing.Color.Red
        Me.laDateDernInventaire.Location = New System.Drawing.Point(536, 152)
        Me.laDateDernInventaire.Name = "laDateDernInventaire"
        Me.laDateDernInventaire.Size = New System.Drawing.Size(128, 16)
        Me.laDateDernInventaire.TabIndex = 53
        Me.laDateDernInventaire.Text = "Date_dern_inventaire"
        '
        'Label13
        '
        Me.Label13.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label13.Location = New System.Drawing.Point(368, 152)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(160, 16)
        Me.Label13.TabIndex = 52
        Me.Label13.Text = "Date dernier inventaire"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.tbmvtLibelle)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.cbomvtType)
        Me.GroupBox1.Controls.Add(Me.tbmvtCommentaiire)
        Me.GroupBox1.Controls.Add(Me.cbAppliquer)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.tbmvtQuantite)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.dtmvtDate)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Location = New System.Drawing.Point(520, 280)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(376, 248)
        Me.GroupBox1.TabIndex = 18
        Me.GroupBox1.TabStop = False
        '
        'tbmvtLibelle
        '
        Me.tbmvtLibelle.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcMvtStock, "libelle", True))
        Me.tbmvtLibelle.Location = New System.Drawing.Point(80, 80)
        Me.tbmvtLibelle.Name = "tbmvtLibelle"
        Me.tbmvtLibelle.Size = New System.Drawing.Size(288, 20)
        Me.tbmvtLibelle.TabIndex = 2
        '
        'm_bsrcMvtStock
        '
        Me.m_bsrcMvtStock.DataSource = GetType(vini_DB.mvtStock)
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(8, 80)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(56, 23)
        Me.Label18.TabIndex = 22
        Me.Label18.Text = "Libellé"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 48)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(56, 24)
        Me.Label14.TabIndex = 21
        Me.Label14.Text = "Type"
        '
        'cbomvtType
        '
        Me.cbomvtType.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcMvtStock, "typeMvtSTR", True))
        Me.cbomvtType.DataSource = Me.m_bsrcTypeMvt
        Me.cbomvtType.DisplayMember = "valeur"
        Me.cbomvtType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbomvtType.Location = New System.Drawing.Point(80, 48)
        Me.cbomvtType.Name = "cbomvtType"
        Me.cbomvtType.Size = New System.Drawing.Size(288, 21)
        Me.cbomvtType.TabIndex = 1
        Me.cbomvtType.ValueMember = "code"
        '
        'm_bsrcTypeMvt
        '
        Me.m_bsrcTypeMvt.DataSource = GetType(vini_DB.Param)
        '
        'tbmvtCommentaiire
        '
        Me.tbmvtCommentaiire.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcMvtStock, "Commentaire", True))
        Me.tbmvtCommentaiire.Location = New System.Drawing.Point(80, 144)
        Me.tbmvtCommentaiire.Name = "tbmvtCommentaiire"
        Me.tbmvtCommentaiire.Size = New System.Drawing.Size(240, 56)
        Me.tbmvtCommentaiire.TabIndex = 4
        Me.tbmvtCommentaiire.Text = ""
        '
        'cbAppliquer
        '
        Me.cbAppliquer.Location = New System.Drawing.Point(288, 208)
        Me.cbAppliquer.Name = "cbAppliquer"
        Me.cbAppliquer.Size = New System.Drawing.Size(80, 24)
        Me.cbAppliquer.TabIndex = 5
        Me.cbAppliquer.Text = "Appliquer"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 144)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(72, 16)
        Me.Label12.TabIndex = 13
        Me.Label12.Text = "Commentaire"
        '
        'tbmvtQuantite
        '
        Me.tbmvtQuantite.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcMvtStock, "qte", True))
        Me.tbmvtQuantite.Location = New System.Drawing.Point(80, 112)
        Me.tbmvtQuantite.Name = "tbmvtQuantite"
        Me.tbmvtQuantite.Size = New System.Drawing.Size(104, 20)
        Me.tbmvtQuantite.TabIndex = 3
        Me.tbmvtQuantite.Text = "0"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 112)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 24)
        Me.Label11.TabIndex = 11
        Me.Label11.Text = "Quantité"
        '
        'dtmvtDate
        '
        Me.dtmvtDate.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcMvtStock, "datemvt", True))
        Me.dtmvtDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtmvtDate.Location = New System.Drawing.Point(80, 16)
        Me.dtmvtDate.Name = "dtmvtDate"
        Me.dtmvtDate.Size = New System.Drawing.Size(104, 20)
        Me.dtmvtDate.TabIndex = 0
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 24)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "Date"
        '
        'cbSTKSupprimer
        '
        Me.cbSTKSupprimer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSTKSupprimer.Location = New System.Drawing.Point(416, 360)
        Me.cbSTKSupprimer.Name = "cbSTKSupprimer"
        Me.cbSTKSupprimer.Size = New System.Drawing.Size(96, 24)
        Me.cbSTKSupprimer.TabIndex = 17
        Me.cbSTKSupprimer.Text = "Suppprimer"
        '
        'cbStkModifier
        '
        Me.cbStkModifier.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbStkModifier.Location = New System.Drawing.Point(416, 328)
        Me.cbStkModifier.Name = "cbStkModifier"
        Me.cbStkModifier.Size = New System.Drawing.Size(96, 24)
        Me.cbStkModifier.TabIndex = 16
        Me.cbStkModifier.Text = "Modifier"
        '
        'cbAjouter
        '
        Me.cbAjouter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAjouter.Location = New System.Drawing.Point(416, 296)
        Me.cbAjouter.Name = "cbAjouter"
        Me.cbAjouter.Size = New System.Drawing.Size(96, 24)
        Me.cbAjouter.TabIndex = 15
        Me.cbAjouter.Text = "Ajouter"
        '
        'laQteStock
        '
        Me.laQteStock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.laQteStock.ForeColor = System.Drawing.Color.Red
        Me.laQteStock.Location = New System.Drawing.Point(536, 192)
        Me.laQteStock.Name = "laQteStock"
        Me.laQteStock.Size = New System.Drawing.Size(128, 16)
        Me.laQteStock.TabIndex = 46
        Me.laQteStock.Text = "Qte_en_stock"
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.Location = New System.Drawing.Point(368, 192)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(160, 16)
        Me.Label9.TabIndex = 45
        Me.Label9.Text = "Stock Réel"
        '
        'cbCalculStock
        '
        Me.cbCalculStock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbCalculStock.Location = New System.Drawing.Point(696, 208)
        Me.cbCalculStock.Name = "cbCalculStock"
        Me.cbCalculStock.Size = New System.Drawing.Size(80, 24)
        Me.cbCalculStock.TabIndex = 19
        Me.cbCalculStock.Text = "Recalculer"
        '
        'Label16
        '
        Me.Label16.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label16.Location = New System.Drawing.Point(680, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(40, 23)
        Me.Label16.TabIndex = 56
        Me.Label16.Text = "Cond"
        '
        'cboConditionnement
        '
        Me.cboConditionnement.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboConditionnement.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcProduit, "idConditionnement", True))
        Me.cboConditionnement.DataSource = Me.m_bsrcConditionnement
        Me.cboConditionnement.DisplayMember = "valeur"
        Me.cboConditionnement.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboConditionnement.Location = New System.Drawing.Point(736, 32)
        Me.cboConditionnement.Name = "cboConditionnement"
        Me.cboConditionnement.Size = New System.Drawing.Size(136, 21)
        Me.cboConditionnement.TabIndex = 6
        Me.cboConditionnement.ValueMember = "id"
        '
        'm_bsrcConditionnement
        '
        Me.m_bsrcConditionnement.DataSource = GetType(vini_DB.Param)
        '
        'Label17
        '
        Me.Label17.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label17.Location = New System.Drawing.Point(368, 168)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(160, 16)
        Me.Label17.TabIndex = 58
        Me.Label17.Text = "Qte au dernier inventaire"
        '
        'laQteDernierInventaire
        '
        Me.laQteDernierInventaire.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.laQteDernierInventaire.ForeColor = System.Drawing.Color.Red
        Me.laQteDernierInventaire.Location = New System.Drawing.Point(536, 168)
        Me.laQteDernierInventaire.Name = "laQteDernierInventaire"
        Me.laQteDernierInventaire.Size = New System.Drawing.Size(100, 23)
        Me.laQteDernierInventaire.TabIndex = 59
        Me.laQteDernierInventaire.Text = "qte_dern_invent"
        '
        'Label19
        '
        Me.Label19.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label19.Location = New System.Drawing.Point(368, 216)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(160, 23)
        Me.Label19.TabIndex = 60
        Me.Label19.Text = "Quantité commandée"
        '
        'laQteCommande
        '
        Me.laQteCommande.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.laQteCommande.ForeColor = System.Drawing.Color.Green
        Me.laQteCommande.Location = New System.Drawing.Point(536, 216)
        Me.laQteCommande.Name = "laQteCommande"
        Me.laQteCommande.Size = New System.Drawing.Size(100, 16)
        Me.laQteCommande.TabIndex = 61
        Me.laQteCommande.Text = "Qte_commandée"
        '
        'Label21
        '
        Me.Label21.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label21.Location = New System.Drawing.Point(368, 248)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(160, 16)
        Me.Label21.TabIndex = 62
        Me.Label21.Text = "Stock Théorique"
        '
        'laQteStockTheorique
        '
        Me.laQteStockTheorique.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.laQteStockTheorique.ForeColor = System.Drawing.Color.Red
        Me.laQteStockTheorique.Location = New System.Drawing.Point(536, 248)
        Me.laQteStockTheorique.Name = "laQteStockTheorique"
        Me.laQteStockTheorique.Size = New System.Drawing.Size(128, 16)
        Me.laQteStockTheorique.TabIndex = 63
        Me.laQteStockTheorique.Text = "qte_stock_theorique"
        '
        'ckDispo
        '
        Me.ckDispo.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckDispo.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Me.m_bsrcProduit, "bDisponible", True))
        Me.ckDispo.Location = New System.Drawing.Point(224, 120)
        Me.ckDispo.Name = "ckDispo"
        Me.ckDispo.Size = New System.Drawing.Size(128, 24)
        Me.ckDispo.TabIndex = 13
        Me.ckDispo.Text = "Produit Disponible"
        '
        'Label20
        '
        Me.Label20.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(683, 101)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(38, 13)
        Me.Label20.TabIndex = 64
        Me.Label20.Text = "Tarif A"
        '
        'tbTarifA
        '
        Me.tbTarifA.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTarifA.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "TarifA", True))
        Me.tbTarifA.Location = New System.Drawing.Point(736, 98)
        Me.tbTarifA.Name = "tbTarifA"
        Me.tbTarifA.Size = New System.Drawing.Size(136, 20)
        Me.tbTarifA.TabIndex = 65
        Me.tbTarifA.Text = "0"
        '
        'tbTarifB
        '
        Me.tbTarifB.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTarifB.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "TarifB", True))
        Me.tbTarifB.Location = New System.Drawing.Point(736, 124)
        Me.tbTarifB.Name = "tbTarifB"
        Me.tbTarifB.Size = New System.Drawing.Size(136, 20)
        Me.tbTarifB.TabIndex = 66
        Me.tbTarifB.Text = "0"
        '
        'tbTarifC
        '
        Me.tbTarifC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTarifC.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "TarifC", True))
        Me.tbTarifC.Location = New System.Drawing.Point(736, 149)
        Me.tbTarifC.Name = "tbTarifC"
        Me.tbTarifC.Size = New System.Drawing.Size(136, 20)
        Me.tbTarifC.TabIndex = 67
        Me.tbTarifC.Text = "0"
        '
        'Label22
        '
        Me.Label22.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(683, 127)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(38, 13)
        Me.Label22.TabIndex = 68
        Me.Label22.Text = "Tarif B"
        '
        'Label23
        '
        Me.Label23.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(683, 152)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(38, 13)
        Me.Label23.TabIndex = 69
        Me.Label23.Text = "Tarif C"
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DatemvtDataGridViewTextBoxColumn, Me.typeMvt, Me.QteDataGridViewTextBoxColumn, Me.LibelleDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcMvtStock
        Me.DataGridView1.Location = New System.Drawing.Point(11, 151)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 10
        Me.DataGridView1.Size = New System.Drawing.Size(351, 377)
        Me.DataGridView1.TabIndex = 70
        '
        'DatemvtDataGridViewTextBoxColumn
        '
        Me.DatemvtDataGridViewTextBoxColumn.DataPropertyName = "datemvt"
        Me.DatemvtDataGridViewTextBoxColumn.FillWeight = 4.0!
        Me.DatemvtDataGridViewTextBoxColumn.HeaderText = "Date"
        Me.DatemvtDataGridViewTextBoxColumn.Name = "DatemvtDataGridViewTextBoxColumn"
        '
        'typeMvt
        '
        Me.typeMvt.DataPropertyName = "typeMvtSTR"
        Me.typeMvt.FillWeight = 1.0!
        Me.typeMvt.HeaderText = "T"
        Me.typeMvt.Name = "typeMvt"
        '
        'QteDataGridViewTextBoxColumn
        '
        Me.QteDataGridViewTextBoxColumn.DataPropertyName = "qte"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle1.Format = "N0"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.QteDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle1
        Me.QteDataGridViewTextBoxColumn.FillWeight = 4.0!
        Me.QteDataGridViewTextBoxColumn.HeaderText = "Qte"
        Me.QteDataGridViewTextBoxColumn.Name = "QteDataGridViewTextBoxColumn"
        '
        'LibelleDataGridViewTextBoxColumn
        '
        Me.LibelleDataGridViewTextBoxColumn.DataPropertyName = "libelle"
        Me.LibelleDataGridViewTextBoxColumn.FillWeight = 10.0!
        Me.LibelleDataGridViewTextBoxColumn.HeaderText = "Designation"
        Me.LibelleDataGridViewTextBoxColumn.Name = "LibelleDataGridViewTextBoxColumn"
        '
        'frmProduit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(915, 546)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.tbTarifC)
        Me.Controls.Add(Me.tbTarifB)
        Me.Controls.Add(Me.tbTarifA)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.tbCodeStat)
        Me.Controls.Add(Me.tbMillesime)
        Me.Controls.Add(Me.tbMotDirecteur)
        Me.Controls.Add(Me.tbDesignation)
        Me.Controls.Add(Me.tbCode)
        Me.Controls.Add(Me.tbCodeFourn)
        Me.Controls.Add(Me.laQteStockTheorique)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.laQteCommande)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.laQteDernierInventaire)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.cboConditionnement)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.cbCalculStock)
        Me.Controls.Add(Me.laDateDernInventaire)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cbSTKSupprimer)
        Me.Controls.Add(Me.cbStkModifier)
        Me.Controls.Add(Me.cbAjouter)
        Me.Controls.Add(Me.laQteStock)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.ckDispo)
        Me.Controls.Add(Me.ckStock)
        Me.Controls.Add(Me.cboContenant)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.cboCouleur)
        Me.Controls.Add(Me.cboRegion)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.liNomFournisseur)
        Me.Controls.Add(Me.cbRechercher)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmProduit"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Type"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.m_bsrcProduit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcCouleur, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcRegion, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcContenant, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.m_bsrcMvtStock, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcTypeMvt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcConditionnement, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Méthodes Privées"
    Private Function initFenetre() As Boolean
        Dim bReturn As Boolean
        Dim objParam As Param
        Dim objContenant As contenant
        debAffiche()
        Try

            For Each objParam In Param.colRegion
                m_bsrcRegion.Add(objParam)
            Next
            For Each objParam In Param.colCouleur
                m_bsrcCouleur.Add(objParam)
            Next
            For Each objParam In Param.colConditionnement
                m_bsrcConditionnement.Add(objParam)
            Next
            For Each objContenant In contenant.colContenant
                m_bsrcContenant.Add(objContenant)
            Next
            initcboTypeMvt(cbomvtType)
            If currentuser.role = vncEnums.userRole.ADMIN Or currentuser.role = vncEnums.userRole.COMPTABILITE Then
                tbCodeFourn.Enabled = True
                cbRechercher.Enabled = True
            Else
                tbCodeFourn.Enabled = False
                cbRechercher.Enabled = False
            End If


            bReturn = True
        Catch ex As Exception
            bReturn = False
            Debug.Assert(bReturn, "initFenetre" & ex.ToString)
            Throw ex
        End Try
        finAffiche()
        Return bReturn
    End Function 'initFenetre
    Private Sub initcboTypeMvt(ByVal cbo As ComboBox)
        m_bsrcTypeMvt.Clear()
        Dim objParam As Param
        objParam = New Param("X")
        objParam.code = vncTypeMvt.vncMvtInventaire
        objParam.valeur = "1-Inventaire"
        m_bsrcTypeMvt.Add(objParam)
        objParam = New Param("X")
        objParam.code = "2"
        objParam.valeur = "2-CommandeClient"
        m_bsrcTypeMvt.Add(objParam)
        objParam = New Param("X")
        objParam.code = "3"
        objParam.valeur = "3-BonAppro"
        m_bsrcTypeMvt.Add(objParam)
        objParam = New Param("X")
        objParam.code = "4"
        objParam.valeur = "4-Regul"
        m_bsrcTypeMvt.Add(objParam)
    End Sub
    '================================================================
    'Function : ModifierMvtStock()
    'Description : 
    '================================================================
    Private Function ModifierMvtStock() As Boolean
        Debug.Assert(Not m_objProduitCourant Is Nothing, "Pas de Produit courant")
        Dim nIndex As Integer
        Dim bReturn As Boolean
        Try
            'Récupération du mouvement de stock concerné
            'nIndex = flxLstMvtStock.get_TextMatrix(flxLstMvtStock.RowSel, COL_STKINDEX)
            m_objmvtCourant = m_objProduitCourant.colmvtStock(nIndex)
            If (Not m_objmvtCourant.bDeleted) Then
                m_bAjoutmvt = False
                'Affichage du mouvement de stock
                initSaisieMvt()
                activerSaisieMvt(True)
                bReturn = True
            Else
                bReturn = False
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'ModifierMvtStock
    '================================================================
    'Function :         SupprimerMvtStock()
    'Description : créer un mvt de stock vide et active la saisie du mvt de stock
    '================================================================
    Private Function SupprimerMvtStock() As Boolean
        Debug.Assert(Not m_objProduitCourant Is Nothing, "Pas de Produit courant")
        Dim nIndex As Integer
        Dim bReturn As Boolean
        Try
            'Récupération du mouvement de stock concerné
            nIndex = m_bsrcMvtStock.Position
            m_objmvtCourant = m_bsrcMvtStock.Current
            If Not m_objmvtCourant.bDeleted Then
                Debug.Assert(Not m_objmvtCourant Is Nothing, "Pas de Mouvement Courant")
                m_bAjoutmvt = False

                If MsgBox("Etes-vous sûr de vouloir supprimer cet élément " & m_objmvtCourant.shortResume, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    m_bsrcMvtStock.RemoveCurrent()
                    m_objProduitCourant.supprimeLigneMvtStock(nIndex)
                    setfrmUpdated()      'MAJ etat de la fenetre : afaire manuellement car les boutons ne générent pas d'événement et le bouton Supprimer ne génére pas d'affichage
                    afficheStock()          'Affichage des Zones de Stocks
                    activerSaisieMvt(False) 'Désactivation des Zones de Saisie
                End If


                bReturn = True
            Else
                bReturn = False
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'SupprimermvtStock

    '================================================================
    'Function :         AjouterMvtStock()
    'Description : créer un mvt de stock vide et active la saisie du mvt de stock
    '================================================================
    Private Function AjouterMvtStock() As Boolean
        Dim bReturn As Boolean
        Try
            m_objmvtCourant = New mvtStock(Now(), getElementCourant().id, vncEnums.vncTypeMvt.vncmvtRegul, 0, "")
            m_bsrcMvtStock.Add(m_objmvtCourant)
            m_bsrcMvtStock.MoveLast()
            m_bAjoutmvt = True
            initSaisieMvt()
            activerSaisieMvt(True)
            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'ajouterMvtStock

    '================================================================
    'Function : ActiverSaisieMvt()
    'Description : Enable/Disable les controle de saisie d'un mouvement
    '================================================================
    Private Function activerSaisieMvt(ByVal pbEnabled As Boolean) As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
            dtmvtDate.Enabled = pbEnabled
            cbomvtType.Enabled = pbEnabled
            tbmvtLibelle.Enabled = pbEnabled
            tbmvtQuantite.Enabled = pbEnabled
            tbmvtCommentaiire.Enabled = pbEnabled
            dtmvtDate.Enabled = pbEnabled
            cbAppliquer.Enabled = pbEnabled
            If pbEnabled = True Then
                dtmvtDate.Focus()
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'activerSaisieMvt
    '=====================================================================
    'Function : initSaisieMvt
    'Description : initialisation des zones de saisie d'un mouvement
    '=====================================================================
    Private Function initSaisieMvt() As Boolean
        Dim bReturn As Boolean
        'Try
        debAffiche()
        'dtmvtDate.Text = m_objmvtCourant.datemvt.ToShortDateString
        'cbomvtType.SelectedIndex = m_objmvtCourant.typeMvt - 1
        'tbmvtLibelle.Text = m_objmvtCourant.libelle
        'tbmvtQuantite.Text = m_objmvtCourant.qte
        'tbmvtCommentaiire.Text = m_objmvtCourant.Commentaire
        bReturn = True
        finAffiche()
        'Catch ex As Exception
        '    bReturn = False
        '    End Try
        Return bReturn
    End Function 'initSaisiemvt

    '====================================================================
    'Function : AppliquerSaisieMvt
    'Description : Met à jour le mvt Courant à partir de la saisie
    '               Ajoute le mouvement à la collection ou Met à jour l'élément
    '               Affiche le stock
    '====================================================================
    Private Function appliquerSaisieMvt() As Boolean
        Debug.Assert(Not m_objProduitCourant Is Nothing, "Pas de Produit Courant")
        Dim bReturn As Boolean
        Try
            'traitement du mouvement de stock
            If m_bAjoutmvt Then
                m_objmvtCourant = m_objProduitCourant.ajouteLigneMvtStock(m_objmvtCourant)
                Debug.Assert(Not m_objmvtCourant Is Nothing, mvtStock.getErreur())
                m_bAjoutmvt = False
            Else
                m_objProduitCourant.recalculStock()
            End If
            m_objProduitCourant.setcolMvtStockUpdated()
            afficheStock()          'Affichage des Zones de Stocks
            activerSaisieMvt(False) 'Désactivation des Zones de Saisie
            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'appliquerSaisieMvt


    Private Sub rechercheFournisseur()
        Dim objFournisseur As Fournisseur

        'If clientdejaCharge() Then
        '    If MsgBox("Cette commande est déjà affecté à cette commande. Souhaitez-vous le remplacer", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
        '        Exit Sub
        '    End If
        'End If

        objFournisseur = rechercheDonnee(vncEnums.vncTypeDonnee.FOURNISSEUR, tbCodeFourn)
        If Not objFournisseur Is Nothing Then
            If objFournisseur.bResume Then
                objFournisseur.load()
            End If

            If Not objFournisseur Is Nothing Then
                afficheFournisseur(objFournisseur)
            End If
        End If
    End Sub 'RechercheFournisseur
    Private Sub afficheFournisseur(ByVal objFournisseur As Fournisseur)
        liNomFournisseur.Text = objFournisseur.nom
        liNomFournisseur.Tag = objFournisseur.id
        tbCodeFourn.Text = objFournisseur.code
    End Sub
    Private Sub afficheStock()
        Debug.Assert(Not m_objProduitCourant Is Nothing, "Pas de Produit Courant")
        afficheQteStock()
        afficheLstMvtStock()
    End Sub
    '==============================================================
    'Function : AfficheQteStock
    'Description : Affichage des quantité en Stocks
    '===============================================================
    Private Function afficheQteStock() As Boolean
        Debug.Assert(Not m_objProduitCourant Is Nothing, "Pas de Produit Courant")
        Dim bReturn As Boolean
        bReturn = True
        laDateDernInventaire.Text = m_objProduitCourant.DateDernInventaire
        laQteDernierInventaire.Text = m_objProduitCourant.QteStockDernInventaire
        laQteStock.Text = m_objProduitCourant.QteStock
        laQteCommande.Text = m_objProduitCourant.qteCommande
        laQteStockTheorique.Text = m_objProduitCourant.QteStock - m_objProduitCourant.qteCommande
        Return bReturn
    End Function
    Private Function afficheLstMvtStock() As Boolean
        Debug.Assert(Not m_objProduitCourant Is Nothing, "Pas de Produit Courant")
        Dim bReturn As Boolean
        Dim objMvt As mvtStock

        bReturn = True
        'Affichage de la liste des lignes
        m_bsrcMvtStock.Clear()
        For Each objMvt In m_objProduitCourant.colmvtStock
            If Not objMvt.bDeleted Then
                m_bsrcMvtStock.Add(objMvt)
            End If
        Next objMvt
        '
        Return bReturn
    End Function

    Private Function RecalculStock() As Boolean

        Dim bReturn As Boolean
        Debug.Assert(Not m_objProduitCourant Is Nothing, "Pas de produit courant")
        Try
            bReturn = m_objProduitCourant.recalculStock()
            If bReturn Then
                afficheStock()
            End If

        Catch ex As Exception
            bReturn = False
            DisplayError("RecalculStock", ex.ToString)
        End Try
        Return bReturn
    End Function
#End Region
#Region "Méthodes Redefines"
    Protected Overrides Sub setToolbarButtons()
        'm_ToolBarNewEnabled = False
        'm_ToolBarLoadEnabled = False
        'm_ToolBarSaveEnabled = False
        'm_ToolBarDelEnabled = False
        'm_ToolBarRefreshEnabled = False

    End Sub
    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenetre n'est pas libre")
        Dim bReturn As Boolean
        bReturn = setElementCourant2(New Produit("", New Fournisseur, 1990))

        Return bReturn
    End Function

    Public Overrides Function setElementCourant2(ByVal pElement As Persist) As Boolean
        Dim bReturn As Boolean
        m_objProduitCourant = CType(pElement, Produit)
        bReturn = MyBase.setElementCourant2(pElement)
        If bReturn Then
            m_bsrcProduit.Clear()
            m_bsrcProduit.Add(m_objProduitCourant)
        End If
        Return bReturn
    End Function
    Public Overrides Function getResume() As String 'Rend le caption de la fenêtre
        If getElementCourant() Is Nothing Then
            Return "Gestion des Produits"
        Else
            Return "Produit : " & getElementCourant.shortResume
        End If
    End Function 'getResume

    Public Overrides Function AfficheElement() As Boolean
        Dim objFournisseur As Fournisseur
        Dim bReturn As Boolean


        'Chargement des collection de mvts de stocks
        bReturn = m_objProduitCourant.loadcolmvtStock()
        Debug.Assert(bReturn, "Chargement de la collection des mvts de stocks")

        'Chargement du fournisseur
        objFournisseur = New Fournisseur
        If m_objProduitCourant.idFournisseur <> 0 Then
            objFournisseur.load(m_objProduitCourant.idFournisseur)
        End If

        afficheFournisseur(objFournisseur)

        afficheStock()
        activerSaisieMvt(False)
        EnableControls(True)
    End Function 'AfficheElementCourant

    Public Overrides Function MAJElement() As Boolean
        Dim bReturn As Boolean
        Dim objProduit As Produit


        objProduit = CType(getElementCourant(), Produit)
        Try
            objProduit.idFournisseur = liNomFournisseur.Tag ' l'id du fournisseur est dans le tag du nom (det de lien pour l'affichage de la fenêtre)

            bReturn = True
        Catch ex As Exception
            DisplayStatus("")
            bReturn = False
        End Try
        Return bReturn
    End Function 'MAJElementCourant
    Public Overrides Function ControleAvantSauvegarde() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
            If getElementCourant().bNew = True Then
                'controle de l'unicité du code
                If Produit.getListe(vncEnums.vncTypeProduit.vncTous, tbCode.Text).Count > 0 Then
                    DisplayError("Controleavantsauvegarde", "Le code produit doit être unique")
                    bReturn = False
                End If
            End If
            If Not (liNomFournisseur.Tag <> 0 And IsNumeric(liNomFournisseur.Tag)) Then
                bReturn = False
                DisplayError("Controleavantsauvegarde", "Le Fournisseur n'est pas renseigné")
            End If
        Catch ex As Exception
            bReturn = False
            Throw ex
        End Try
        Return bReturn
    End Function
    Public Overrides Function SauveElement() As Boolean
        Debug.Assert(Not m_objProduitCourant Is Nothing)
        Dim bReturn As Boolean
        bReturn = m_objProduitCourant.save
        Debug.Assert(bReturn, "Erreur en sauvegarde")
        Return bReturn
    End Function
#End Region


    Private Sub liNomFournisseur_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liNomFournisseur.LinkClicked
        afficheFenetreFournisseur(liNomFournisseur.Tag)
    End Sub

    Private Sub cbRechercher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRechercher.Click
        rechercheFournisseur()
    End Sub


    Private Sub cbAppliquer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAppliquer.Click
        appliquerSaisieMvt()
    End Sub

    Private Sub cbAjouter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAjouter.Click
        AjouterMvtStock()
    End Sub


    Private Sub cbSTKSupprimer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSTKSupprimer.Click
        'If flxLstMvtStock.RowSel <> 0 Then
        SupprimerMvtStock()
        'End If
    End Sub

    Private Sub cbCalculStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCalculStock.Click
        RecalculStock()
        setfrmUpdated()
    End Sub

    Protected Overrides Function frmNew() As Boolean
        MyBase.frmNew()
        tbCode.Focus()
    End Function

    Private Sub frmProduit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled) ' Activation de tous les contrôles de la fenêtre
        If m_action = vncEnums.vncfrmAction.FRMLOAD Then
            'Modifiation d'un Element
            Me.tbCode.Enabled = currentuser.aLeDroitdeModifierleFourniseurProduit()
            Me.tbCodeFourn.Enabled = currentuser.aLeDroitdeModifierleFourniseurProduit()
            Me.cbRechercher.Enabled = currentuser.aLeDroitdeModifierleFourniseurProduit()
        End If
    End Sub

    Private Sub cbStkModifier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStkModifier.Click
        activerSaisieMvt(True)
    End Sub
End Class