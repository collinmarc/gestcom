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
    Friend WithEvents Label24 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents m_bsrcDossier As BindingSource
    Friend WithEvents LogoList As ImageList
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents dtpStockAu As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnCalcStockAu As System.Windows.Forms.Button
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents tbStockAu As System.Windows.Forms.TextBox
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProduit))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.tbDesignation = New System.Windows.Forms.TextBox()
        Me.m_bsrcProduit = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbCode = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbMotDirecteur = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbMillesime = New vini_app.textBoxNumeric()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboCouleur = New System.Windows.Forms.ComboBox()
        Me.m_bsrcCouleur = New System.Windows.Forms.BindingSource(Me.components)
        Me.cboRegion = New System.Windows.Forms.ComboBox()
        Me.m_bsrcRegion = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboContenant = New System.Windows.Forms.ComboBox()
        Me.m_bsrcContenant = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label8 = New System.Windows.Forms.Label()
        Me.liNomFournisseur = New System.Windows.Forms.LinkLabel()
        Me.cbRechercher = New System.Windows.Forms.Button()
        Me.tbCodeFourn = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ckStock = New System.Windows.Forms.CheckBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.tbCodeStat = New System.Windows.Forms.TextBox()
        Me.laDateDernInventaire = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tbmvtLibelle = New System.Windows.Forms.TextBox()
        Me.m_bsrcMvtStock = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cbomvtType = New System.Windows.Forms.ComboBox()
        Me.m_bsrcTypeMvt = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbmvtCommentaiire = New System.Windows.Forms.RichTextBox()
        Me.cbAppliquer = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.tbmvtQuantite = New vini_app.textBoxNumeric()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.dtmvtDate = New System.Windows.Forms.DateTimePicker()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cbSTKSupprimer = New System.Windows.Forms.Button()
        Me.cbStkModifier = New System.Windows.Forms.Button()
        Me.cbAjouter = New System.Windows.Forms.Button()
        Me.laQteStock = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cbCalculStock = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cboConditionnement = New System.Windows.Forms.ComboBox()
        Me.m_bsrcConditionnement = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label17 = New System.Windows.Forms.Label()
        Me.laQteDernierInventaire = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.laQteCommande = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.laQteStockTheorique = New System.Windows.Forms.Label()
        Me.ckDispo = New System.Windows.Forms.CheckBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.tbTarifA = New vini_app.textBoxCurrency()
        Me.tbTarifB = New vini_app.textBoxCurrency()
        Me.tbTarifC = New vini_app.textBoxCurrency()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DatemvtDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.typeMvt = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LibelleDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.m_bsrcDossier = New System.Windows.Forms.BindingSource(Me.components)
        Me.LogoList = New System.Windows.Forms.ImageList(Me.components)
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.dtpStockAu = New System.Windows.Forms.DateTimePicker()
        Me.btnCalcStockAu = New System.Windows.Forms.Button()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.tbStockAu = New System.Windows.Forms.TextBox()
        CType(Me.m_bsrcProduit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcCouleur, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcRegion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcContenant, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.m_bsrcMvtStock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcTypeMvt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcConditionnement, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcDossier, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbDesignation
        '
        Me.tbDesignation.AcceptsReturn = True
        resources.ApplyResources(Me.tbDesignation, "tbDesignation")
        Me.tbDesignation.BackColor = System.Drawing.SystemColors.Window
        Me.tbDesignation.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbDesignation.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "nom", True))
        Me.tbDesignation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbDesignation.Name = "tbDesignation"
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
        resources.ApplyResources(Me.tbCode, "tbCode")
        Me.tbCode.Name = "tbCode"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'tbMotDirecteur
        '
        Me.tbMotDirecteur.AcceptsReturn = True
        Me.tbMotDirecteur.BackColor = System.Drawing.SystemColors.Window
        Me.tbMotDirecteur.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbMotDirecteur.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "motcle", True))
        Me.tbMotDirecteur.ForeColor = System.Drawing.SystemColors.WindowText
        resources.ApplyResources(Me.tbMotDirecteur, "tbMotDirecteur")
        Me.tbMotDirecteur.Name = "tbMotDirecteur"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'tbMillesime
        '
        Me.tbMillesime.AcceptsReturn = True
        resources.ApplyResources(Me.tbMillesime, "tbMillesime")
        Me.tbMillesime.BackColor = System.Drawing.SystemColors.Window
        Me.tbMillesime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbMillesime.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "millesime", True))
        Me.tbMillesime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbMillesime.Name = "tbMillesime"
        '
        'Label7
        '
        resources.ApplyResources(Me.Label7, "Label7")
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Name = "Label7"
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
        resources.ApplyResources(Me.cboCouleur, "cboCouleur")
        Me.cboCouleur.Name = "cboCouleur"
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
        resources.ApplyResources(Me.cboRegion, "cboRegion")
        Me.cboRegion.Name = "cboRegion"
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
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
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
        resources.ApplyResources(Me.cboContenant, "cboContenant")
        Me.cboContenant.Name = "cboContenant"
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
        resources.ApplyResources(Me.Label8, "Label8")
        Me.Label8.Name = "Label8"
        '
        'liNomFournisseur
        '
        resources.ApplyResources(Me.liNomFournisseur, "liNomFournisseur")
        Me.liNomFournisseur.Name = "liNomFournisseur"
        Me.liNomFournisseur.TabStop = True
        '
        'cbRechercher
        '
        resources.ApplyResources(Me.cbRechercher, "cbRechercher")
        Me.cbRechercher.Name = "cbRechercher"
        '
        'tbCodeFourn
        '
        resources.ApplyResources(Me.tbCodeFourn, "tbCodeFourn")
        Me.tbCodeFourn.Name = "tbCodeFourn"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'ckStock
        '
        resources.ApplyResources(Me.ckStock, "ckStock")
        Me.ckStock.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Me.m_bsrcProduit, "bStock", True))
        Me.ckStock.Name = "ckStock"
        '
        'Label15
        '
        resources.ApplyResources(Me.Label15, "Label15")
        Me.Label15.Name = "Label15"
        '
        'tbCodeStat
        '
        resources.ApplyResources(Me.tbCodeStat, "tbCodeStat")
        Me.tbCodeStat.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "codeStat", True))
        Me.tbCodeStat.Name = "tbCodeStat"
        '
        'laDateDernInventaire
        '
        resources.ApplyResources(Me.laDateDernInventaire, "laDateDernInventaire")
        Me.laDateDernInventaire.ForeColor = System.Drawing.Color.Red
        Me.laDateDernInventaire.Name = "laDateDernInventaire"
        '
        'Label13
        '
        resources.ApplyResources(Me.Label13, "Label13")
        Me.Label13.Name = "Label13"
        '
        'GroupBox1
        '
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
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
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'tbmvtLibelle
        '
        Me.tbmvtLibelle.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcMvtStock, "libelle", True))
        resources.ApplyResources(Me.tbmvtLibelle, "tbmvtLibelle")
        Me.tbmvtLibelle.Name = "tbmvtLibelle"
        '
        'm_bsrcMvtStock
        '
        Me.m_bsrcMvtStock.DataSource = GetType(vini_DB.mvtStock)
        '
        'Label18
        '
        resources.ApplyResources(Me.Label18, "Label18")
        Me.Label18.Name = "Label18"
        '
        'Label14
        '
        resources.ApplyResources(Me.Label14, "Label14")
        Me.Label14.Name = "Label14"
        '
        'cbomvtType
        '
        Me.cbomvtType.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcMvtStock, "typeMvtSTR", True))
        Me.cbomvtType.DataSource = Me.m_bsrcTypeMvt
        Me.cbomvtType.DisplayMember = "valeur"
        Me.cbomvtType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.cbomvtType, "cbomvtType")
        Me.cbomvtType.Name = "cbomvtType"
        Me.cbomvtType.ValueMember = "code"
        '
        'm_bsrcTypeMvt
        '
        Me.m_bsrcTypeMvt.DataSource = GetType(vini_DB.Param)
        '
        'tbmvtCommentaiire
        '
        Me.tbmvtCommentaiire.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcMvtStock, "Commentaire", True))
        resources.ApplyResources(Me.tbmvtCommentaiire, "tbmvtCommentaiire")
        Me.tbmvtCommentaiire.Name = "tbmvtCommentaiire"
        '
        'cbAppliquer
        '
        resources.ApplyResources(Me.cbAppliquer, "cbAppliquer")
        Me.cbAppliquer.Name = "cbAppliquer"
        '
        'Label12
        '
        resources.ApplyResources(Me.Label12, "Label12")
        Me.Label12.Name = "Label12"
        '
        'tbmvtQuantite
        '
        Me.tbmvtQuantite.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcMvtStock, "qte", True))
        resources.ApplyResources(Me.tbmvtQuantite, "tbmvtQuantite")
        Me.tbmvtQuantite.Name = "tbmvtQuantite"
        '
        'Label11
        '
        resources.ApplyResources(Me.Label11, "Label11")
        Me.Label11.Name = "Label11"
        '
        'dtmvtDate
        '
        Me.dtmvtDate.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcMvtStock, "datemvt", True))
        Me.dtmvtDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        resources.ApplyResources(Me.dtmvtDate, "dtmvtDate")
        Me.dtmvtDate.Name = "dtmvtDate"
        '
        'Label10
        '
        resources.ApplyResources(Me.Label10, "Label10")
        Me.Label10.Name = "Label10"
        '
        'cbSTKSupprimer
        '
        resources.ApplyResources(Me.cbSTKSupprimer, "cbSTKSupprimer")
        Me.cbSTKSupprimer.Name = "cbSTKSupprimer"
        '
        'cbStkModifier
        '
        resources.ApplyResources(Me.cbStkModifier, "cbStkModifier")
        Me.cbStkModifier.Name = "cbStkModifier"
        '
        'cbAjouter
        '
        resources.ApplyResources(Me.cbAjouter, "cbAjouter")
        Me.cbAjouter.Name = "cbAjouter"
        '
        'laQteStock
        '
        resources.ApplyResources(Me.laQteStock, "laQteStock")
        Me.laQteStock.ForeColor = System.Drawing.Color.Red
        Me.laQteStock.Name = "laQteStock"
        '
        'Label9
        '
        resources.ApplyResources(Me.Label9, "Label9")
        Me.Label9.Name = "Label9"
        '
        'cbCalculStock
        '
        resources.ApplyResources(Me.cbCalculStock, "cbCalculStock")
        Me.cbCalculStock.Name = "cbCalculStock"
        '
        'Label16
        '
        resources.ApplyResources(Me.Label16, "Label16")
        Me.Label16.Name = "Label16"
        '
        'cboConditionnement
        '
        resources.ApplyResources(Me.cboConditionnement, "cboConditionnement")
        Me.cboConditionnement.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcProduit, "idConditionnement", True))
        Me.cboConditionnement.DataSource = Me.m_bsrcConditionnement
        Me.cboConditionnement.DisplayMember = "valeur"
        Me.cboConditionnement.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboConditionnement.Name = "cboConditionnement"
        Me.cboConditionnement.ValueMember = "id"
        '
        'm_bsrcConditionnement
        '
        Me.m_bsrcConditionnement.DataSource = GetType(vini_DB.Param)
        '
        'Label17
        '
        resources.ApplyResources(Me.Label17, "Label17")
        Me.Label17.Name = "Label17"
        '
        'laQteDernierInventaire
        '
        resources.ApplyResources(Me.laQteDernierInventaire, "laQteDernierInventaire")
        Me.laQteDernierInventaire.ForeColor = System.Drawing.Color.Red
        Me.laQteDernierInventaire.Name = "laQteDernierInventaire"
        '
        'Label19
        '
        resources.ApplyResources(Me.Label19, "Label19")
        Me.Label19.Name = "Label19"
        '
        'laQteCommande
        '
        resources.ApplyResources(Me.laQteCommande, "laQteCommande")
        Me.laQteCommande.ForeColor = System.Drawing.Color.Green
        Me.laQteCommande.Name = "laQteCommande"
        '
        'Label21
        '
        resources.ApplyResources(Me.Label21, "Label21")
        Me.Label21.Name = "Label21"
        '
        'laQteStockTheorique
        '
        resources.ApplyResources(Me.laQteStockTheorique, "laQteStockTheorique")
        Me.laQteStockTheorique.ForeColor = System.Drawing.Color.Red
        Me.laQteStockTheorique.Name = "laQteStockTheorique"
        '
        'ckDispo
        '
        resources.ApplyResources(Me.ckDispo, "ckDispo")
        Me.ckDispo.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Me.m_bsrcProduit, "bDisponible", True))
        Me.ckDispo.Name = "ckDispo"
        '
        'Label20
        '
        resources.ApplyResources(Me.Label20, "Label20")
        Me.Label20.Name = "Label20"
        '
        'tbTarifA
        '
        resources.ApplyResources(Me.tbTarifA, "tbTarifA")
        Me.tbTarifA.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "TarifA", True))
        Me.tbTarifA.Name = "tbTarifA"
        '
        'tbTarifB
        '
        resources.ApplyResources(Me.tbTarifB, "tbTarifB")
        Me.tbTarifB.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "TarifB", True))
        Me.tbTarifB.Name = "tbTarifB"
        '
        'tbTarifC
        '
        resources.ApplyResources(Me.tbTarifC, "tbTarifC")
        Me.tbTarifC.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "TarifC", True))
        Me.tbTarifC.Name = "tbTarifC"
        '
        'Label22
        '
        resources.ApplyResources(Me.Label22, "Label22")
        Me.Label22.Name = "Label22"
        '
        'Label23
        '
        resources.ApplyResources(Me.Label23, "Label23")
        Me.Label23.Name = "Label23"
        '
        'DataGridView1
        '
        resources.ApplyResources(Me.DataGridView1, "DataGridView1")
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DatemvtDataGridViewTextBoxColumn, Me.typeMvt, Me.QteDataGridViewTextBoxColumn, Me.LibelleDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcMvtStock
        Me.DataGridView1.Name = "DataGridView1"
        '
        'DatemvtDataGridViewTextBoxColumn
        '
        Me.DatemvtDataGridViewTextBoxColumn.DataPropertyName = "datemvt"
        Me.DatemvtDataGridViewTextBoxColumn.FillWeight = 4.0!
        resources.ApplyResources(Me.DatemvtDataGridViewTextBoxColumn, "DatemvtDataGridViewTextBoxColumn")
        Me.DatemvtDataGridViewTextBoxColumn.Name = "DatemvtDataGridViewTextBoxColumn"
        '
        'typeMvt
        '
        Me.typeMvt.DataPropertyName = "typeMvtSTR"
        Me.typeMvt.FillWeight = 1.0!
        resources.ApplyResources(Me.typeMvt, "typeMvt")
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
        resources.ApplyResources(Me.QteDataGridViewTextBoxColumn, "QteDataGridViewTextBoxColumn")
        Me.QteDataGridViewTextBoxColumn.Name = "QteDataGridViewTextBoxColumn"
        '
        'LibelleDataGridViewTextBoxColumn
        '
        Me.LibelleDataGridViewTextBoxColumn.DataPropertyName = "libelle"
        Me.LibelleDataGridViewTextBoxColumn.FillWeight = 10.0!
        resources.ApplyResources(Me.LibelleDataGridViewTextBoxColumn, "LibelleDataGridViewTextBoxColumn")
        Me.LibelleDataGridViewTextBoxColumn.Name = "LibelleDataGridViewTextBoxColumn"
        '
        'Label24
        '
        resources.ApplyResources(Me.Label24, "Label24")
        Me.Label24.Name = "Label24"
        '
        'ComboBox1
        '
        resources.ApplyResources(Me.ComboBox1, "ComboBox1")
        Me.ComboBox1.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcProduit, "DossierProduit", True))
        Me.ComboBox1.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.m_bsrcProduit, "DossierProduit", True))
        Me.ComboBox1.DataSource = Me.m_bsrcDossier
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Name = "ComboBox1"
        '
        'm_bsrcDossier
        '
        Me.m_bsrcDossier.DataMember = "String"
        '
        'LogoList
        '
        Me.LogoList.ImageStream = CType(resources.GetObject("LogoList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.LogoList.TransparentColor = System.Drawing.Color.Transparent
        Me.LogoList.Images.SetKeyName(0, "LogoVinicom.jpg")
        Me.LogoList.Images.SetKeyName(1, "LogoHobivin.jpg")
        '
        'PictureBox1
        '
        resources.ApplyResources(Me.PictureBox1, "PictureBox1")
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.TabStop = False
        '
        'dtpStockAu
        '
        resources.ApplyResources(Me.dtpStockAu, "dtpStockAu")
        Me.dtpStockAu.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStockAu.Name = "dtpStockAu"
        '
        'btnCalcStockAu
        '
        resources.ApplyResources(Me.btnCalcStockAu, "btnCalcStockAu")
        Me.btnCalcStockAu.Name = "btnCalcStockAu"
        Me.btnCalcStockAu.UseVisualStyleBackColor = True
        '
        'Label25
        '
        resources.ApplyResources(Me.Label25, "Label25")
        Me.Label25.Name = "Label25"
        '
        'tbStockAu
        '
        resources.ApplyResources(Me.tbStockAu, "tbStockAu")
        Me.tbStockAu.Name = "tbStockAu"
        Me.tbStockAu.ReadOnly = True
        '
        'frmProduit
        '
        resources.ApplyResources(Me, "$this")
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.tbStockAu)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.btnCalcStockAu)
        Me.Controls.Add(Me.dtpStockAu)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label24)
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
        Me.Name = "frmProduit"
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
        CType(Me.m_bsrcDossier, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
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

            m_bsrcDossier.Clear()
            m_bsrcDossier.Add(Dossier.VINICOM)
            m_bsrcDossier.Add(Dossier.HOBIVIN)

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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If (ComboBox1.SelectedIndex >= 0) Then
            PictureBox1.Image = LogoList.Images(ComboBox1.SelectedIndex)
        End If
    End Sub

    Private Sub btnCalcStockAu_Click(sender As Object, e As EventArgs) Handles btnCalcStockAu.Click
        Dim dStockAu As Date
        dStockAu = dtpStockAu.Value
        tbStockAu.Text = CalculStockAu(dStockAu)
    End Sub

    Public Function CalculStockAu(pDate As Date) As Decimal
        Dim nReturn As Decimal
        nReturn = m_objProduitCourant.CalculeStockAu(pDate.AddDays(-1))
        Return nReturn
    End Function
End Class