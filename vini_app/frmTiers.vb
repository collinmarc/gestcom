Imports vini_DB
Public Class frmTiers
    Inherits frmDonBase


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
    Public WithEvents tbRaisonSociale As System.Windows.Forms.TextBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbNom As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents tbCode As System.Windows.Forms.TextBox
    Friend WithEvents tpAddresses As System.Windows.Forms.TabPage
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents tbBanque As System.Windows.Forms.TextBox
    Public WithEvents tbRib4 As System.Windows.Forms.TextBox
    Public WithEvents tbRib3 As System.Windows.Forms.TextBox
    Public WithEvents tbRib2 As System.Windows.Forms.TextBox
    Public WithEvents tbRib1 As System.Windows.Forms.TextBox
    Friend WithEvents RIB As System.Windows.Forms.Label
    Private WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents tbTVAIntra As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents tbSIRET As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents ckAdrIdentiques As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Public WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents tbAdrFactNom As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents tbAdrFactPortable As TextBox
    Public WithEvents tbAdrFactRue1 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrFactRue2 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrFactCP As textBoxNumeric
    Public WithEvents tbAdrFactVille As System.Windows.Forms.TextBox
    Public WithEvents tbAdrFactTel As textBoxStrNumeric
    Public WithEvents tbAdrFactFax As textBoxStrNumeric
    Public WithEvents tbAdrFactEmail As System.Windows.Forms.TextBox
    Public WithEvents Label17 As System.Windows.Forms.Label
    Public WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cboModeRglmt As System.Windows.Forms.ComboBox
    Friend WithEvents tpComs As System.Windows.Forms.TabPage
    Friend WithEvents tabTiers As System.Windows.Forms.TabControl
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents tbAdrLivNom As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Public WithEvents tbAdrLivPortable As TextBox
    Public WithEvents tbAdrLivRue1 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivRue2 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivCP As textBoxNumeric
    Public WithEvents tbAdrLivVille As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLivTel As textBoxStrNumeric
    Public WithEvents tbAdrLivFax As textBoxStrNumeric
    Public WithEvents tbAdrLivEmail As System.Windows.Forms.TextBox
    Friend WithEvents lartbCom4 As System.Windows.Forms.Label
    Friend WithEvents lartbCom3 As System.Windows.Forms.Label
    Friend WithEvents lartbCom2 As System.Windows.Forms.Label
    Friend WithEvents lartbcom1 As System.Windows.Forms.Label
    Public WithEvents rtbCom1 As System.Windows.Forms.RichTextBox
    Public WithEvents rtbCom2 As System.Windows.Forms.RichTextBox
    Public WithEvents rtbCom3 As System.Windows.Forms.RichTextBox
    Friend WithEvents Label400 As System.Windows.Forms.Label
    Friend WithEvents tbCodeCompta As System.Windows.Forms.TextBox
    Friend WithEvents laModeRglmt1 As System.Windows.Forms.Label
    Friend WithEvents laModeReglmt As System.Windows.Forms.Label
    Friend WithEvents cboModeReglement1 As System.Windows.Forms.ComboBox
    Friend WithEvents LaModeReglmt2 As System.Windows.Forms.Label
    Friend WithEvents cboModeReglement2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbIDPrestashop As System.Windows.Forms.TextBox
    Public WithEvents rtbCom4 As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbRaisonSociale = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbNom = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbCode = New System.Windows.Forms.TextBox()
        Me.tabTiers = New System.Windows.Forms.TabControl()
        Me.tpAddresses = New System.Windows.Forms.TabPage()
        Me.cboModeReglement2 = New System.Windows.Forms.ComboBox()
        Me.LaModeReglmt2 = New System.Windows.Forms.Label()
        Me.laModeRglmt1 = New System.Windows.Forms.Label()
        Me.laModeReglmt = New System.Windows.Forms.Label()
        Me.cboModeReglement1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.tbAdrLivNom = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.tbAdrLivPortable = New System.Windows.Forms.TextBox()
        Me.tbAdrLivRue1 = New System.Windows.Forms.TextBox()
        Me.tbAdrLivRue2 = New System.Windows.Forms.TextBox()
        Me.tbAdrLivCP = New vini_app.textBoxNumeric()
        Me.tbAdrLivVille = New System.Windows.Forms.TextBox()
        Me.tbAdrLivTel = New vini_app.textBoxStrNumeric()
        Me.tbAdrLivFax = New vini_app.textBoxStrNumeric()
        Me.tbAdrLivEmail = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.tbBanque = New System.Windows.Forms.TextBox()
        Me.tbRib4 = New System.Windows.Forms.TextBox()
        Me.tbRib3 = New System.Windows.Forms.TextBox()
        Me.tbRib2 = New System.Windows.Forms.TextBox()
        Me.tbRib1 = New System.Windows.Forms.TextBox()
        Me.RIB = New System.Windows.Forms.Label()
        Me.tbTVAIntra = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.tbSIRET = New System.Windows.Forms.TextBox()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.ckAdrIdentiques = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.tbAdrFactNom = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbAdrFactPortable = New System.Windows.Forms.TextBox()
        Me.tbAdrFactRue1 = New System.Windows.Forms.TextBox()
        Me.tbAdrFactRue2 = New System.Windows.Forms.TextBox()
        Me.tbAdrFactCP = New vini_app.textBoxNumeric()
        Me.tbAdrFactVille = New System.Windows.Forms.TextBox()
        Me.tbAdrFactTel = New vini_app.textBoxStrNumeric()
        Me.tbAdrFactFax = New vini_app.textBoxStrNumeric()
        Me.tbAdrFactEmail = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cboModeRglmt = New System.Windows.Forms.ComboBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.tpComs = New System.Windows.Forms.TabPage()
        Me.lartbCom4 = New System.Windows.Forms.Label()
        Me.lartbCom3 = New System.Windows.Forms.Label()
        Me.lartbCom2 = New System.Windows.Forms.Label()
        Me.lartbcom1 = New System.Windows.Forms.Label()
        Me.rtbCom1 = New System.Windows.Forms.RichTextBox()
        Me.rtbCom2 = New System.Windows.Forms.RichTextBox()
        Me.rtbCom3 = New System.Windows.Forms.RichTextBox()
        Me.rtbCom4 = New System.Windows.Forms.RichTextBox()
        Me.Label400 = New System.Windows.Forms.Label()
        Me.tbCodeCompta = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbIDPrestashop = New System.Windows.Forms.TextBox()
        Me.tabTiers.SuspendLayout()
        Me.tpAddresses.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.tpComs.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbRaisonSociale
        '
        Me.tbRaisonSociale.AcceptsReturn = True
        Me.tbRaisonSociale.BackColor = System.Drawing.SystemColors.Window
        Me.tbRaisonSociale.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRaisonSociale.Enabled = False
        Me.tbRaisonSociale.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRaisonSociale.Location = New System.Drawing.Point(88, 56)
        Me.tbRaisonSociale.MaxLength = 0
        Me.tbRaisonSociale.Name = "tbRaisonSociale"
        Me.tbRaisonSociale.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRaisonSociale.Size = New System.Drawing.Size(640, 20)
        Me.tbRaisonSociale.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(8, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(80, 17)
        Me.Label6.TabIndex = 68
        Me.Label6.Text = "Raison Sociale"
        '
        'tbNom
        '
        Me.tbNom.BackColor = System.Drawing.SystemColors.Window
        Me.tbNom.Enabled = False
        Me.tbNom.Location = New System.Drawing.Point(88, 32)
        Me.tbNom.Name = "tbNom"
        Me.tbNom.Size = New System.Drawing.Size(640, 20)
        Me.tbNom.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 16)
        Me.Label1.TabIndex = 65
        Me.Label1.Text = "Nom"
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
        Me.Label2.TabIndex = 64
        Me.Label2.Text = "Code"
        '
        'tbCode
        '
        Me.tbCode.AcceptsReturn = True
        Me.tbCode.BackColor = System.Drawing.SystemColors.Window
        Me.tbCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCode.Enabled = False
        Me.tbCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCode.Location = New System.Drawing.Point(88, 8)
        Me.tbCode.MaxLength = 0
        Me.tbCode.Name = "tbCode"
        Me.tbCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCode.Size = New System.Drawing.Size(129, 20)
        Me.tbCode.TabIndex = 0
        '
        'tabTiers
        '
        Me.tabTiers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabTiers.Controls.Add(Me.tpAddresses)
        Me.tabTiers.Controls.Add(Me.tpComs)
        Me.tabTiers.Enabled = False
        Me.tabTiers.Location = New System.Drawing.Point(11, 95)
        Me.tabTiers.Name = "tabTiers"
        Me.tabTiers.SelectedIndex = 0
        Me.tabTiers.Size = New System.Drawing.Size(952, 557)
        Me.tabTiers.TabIndex = 3
        '
        'tpAddresses
        '
        Me.tpAddresses.Controls.Add(Me.cboModeReglement2)
        Me.tpAddresses.Controls.Add(Me.LaModeReglmt2)
        Me.tpAddresses.Controls.Add(Me.laModeRglmt1)
        Me.tpAddresses.Controls.Add(Me.laModeReglmt)
        Me.tpAddresses.Controls.Add(Me.cboModeReglement1)
        Me.tpAddresses.Controls.Add(Me.GroupBox2)
        Me.tpAddresses.Controls.Add(Me.Label33)
        Me.tpAddresses.Controls.Add(Me.tbBanque)
        Me.tpAddresses.Controls.Add(Me.tbRib4)
        Me.tpAddresses.Controls.Add(Me.tbRib3)
        Me.tpAddresses.Controls.Add(Me.tbRib2)
        Me.tpAddresses.Controls.Add(Me.tbRib1)
        Me.tpAddresses.Controls.Add(Me.RIB)
        Me.tpAddresses.Controls.Add(Me.tbTVAIntra)
        Me.tpAddresses.Controls.Add(Me.Label28)
        Me.tpAddresses.Controls.Add(Me.tbSIRET)
        Me.tpAddresses.Controls.Add(Me.Label31)
        Me.tpAddresses.Controls.Add(Me.ckAdrIdentiques)
        Me.tpAddresses.Controls.Add(Me.GroupBox1)
        Me.tpAddresses.Controls.Add(Me.Label17)
        Me.tpAddresses.Controls.Add(Me.Label16)
        Me.tpAddresses.Controls.Add(Me.cboModeRglmt)
        Me.tpAddresses.Controls.Add(Me.Label25)
        Me.tpAddresses.Location = New System.Drawing.Point(4, 22)
        Me.tpAddresses.Name = "tpAddresses"
        Me.tpAddresses.Size = New System.Drawing.Size(944, 531)
        Me.tpAddresses.TabIndex = 0
        Me.tpAddresses.Text = "Adresses"
        Me.tpAddresses.UseVisualStyleBackColor = True
        '
        'cboModeReglement2
        '
        Me.cboModeReglement2.FormattingEnabled = True
        Me.cboModeReglement2.Location = New System.Drawing.Point(510, 489)
        Me.cboModeReglement2.Name = "cboModeReglement2"
        Me.cboModeReglement2.Size = New System.Drawing.Size(312, 21)
        Me.cboModeReglement2.TabIndex = 9
        '
        'LaModeReglmt2
        '
        Me.LaModeReglmt2.AutoSize = True
        Me.LaModeReglmt2.Location = New System.Drawing.Point(342, 489)
        Me.LaModeReglmt2.Name = "LaModeReglmt2"
        Me.LaModeReglmt2.Size = New System.Drawing.Size(85, 13)
        Me.LaModeReglmt2.TabIndex = 108
        Me.LaModeReglmt2.Text = "LaModeReglmt2"
        '
        'laModeRglmt1
        '
        Me.laModeRglmt1.AutoSize = True
        Me.laModeRglmt1.Location = New System.Drawing.Point(342, 464)
        Me.laModeRglmt1.Name = "laModeRglmt1"
        Me.laModeRglmt1.Size = New System.Drawing.Size(75, 13)
        Me.laModeRglmt1.TabIndex = 107
        Me.laModeRglmt1.Text = "laModeRglmt1"
        '
        'laModeReglmt
        '
        Me.laModeReglmt.AutoSize = True
        Me.laModeReglmt.Location = New System.Drawing.Point(342, 440)
        Me.laModeReglmt.Name = "laModeReglmt"
        Me.laModeReglmt.Size = New System.Drawing.Size(75, 13)
        Me.laModeReglmt.TabIndex = 106
        Me.laModeReglmt.Text = "laModeReglmt"
        '
        'cboModeReglement1
        '
        Me.cboModeReglement1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboModeReglement1.Location = New System.Drawing.Point(510, 461)
        Me.cboModeReglement1.Name = "cboModeReglement1"
        Me.cboModeReglement1.Size = New System.Drawing.Size(312, 21)
        Me.cboModeReglement1.TabIndex = 8
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivNom)
        Me.GroupBox2.Controls.Add(Me.Label21)
        Me.GroupBox2.Controls.Add(Me.Label22)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivPortable)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivRue1)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivRue2)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivCP)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivVille)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivTel)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivFax)
        Me.GroupBox2.Controls.Add(Me.tbAdrLivEmail)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(928, 168)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Adresse de Livraison"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(248, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 97
        Me.Label3.Text = "Email"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(248, 112)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 96
        Me.Label7.Text = "Portable"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 136)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(80, 16)
        Me.Label8.TabIndex = 95
        Me.Label8.Text = "Fax"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(8, 112)
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
        Me.tbAdrLivPortable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivPortable.Location = New System.Drawing.Point(368, 112)
        Me.tbAdrLivPortable.MaxLength = 0
        Me.tbAdrLivPortable.Name = "tbAdrLivPortable"
        Me.tbAdrLivPortable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivPortable.Size = New System.Drawing.Size(145, 20)
        Me.tbAdrLivPortable.TabIndex = 7
        '
        'tbAdrLivRue1
        '
        Me.tbAdrLivRue1.AcceptsReturn = True
        Me.tbAdrLivRue1.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivRue1.Cursor = System.Windows.Forms.Cursors.IBeam
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
        Me.tbAdrLivCP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivCP.Location = New System.Drawing.Point(96, 88)
        Me.tbAdrLivCP.MaxLength = 0
        Me.tbAdrLivCP.Name = "tbAdrLivCP"
        Me.tbAdrLivCP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivCP.Size = New System.Drawing.Size(73, 20)
        Me.tbAdrLivCP.TabIndex = 3
        Me.tbAdrLivCP.Text = "0"
        '
        'tbAdrLivVille
        '
        Me.tbAdrLivVille.AcceptsReturn = True
        Me.tbAdrLivVille.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivVille.Cursor = System.Windows.Forms.Cursors.IBeam
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
        Me.tbAdrLivTel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivTel.Location = New System.Drawing.Point(96, 112)
        Me.tbAdrLivTel.MaxLength = 0
        Me.tbAdrLivTel.Name = "tbAdrLivTel"
        Me.tbAdrLivTel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivTel.Size = New System.Drawing.Size(129, 20)
        Me.tbAdrLivTel.TabIndex = 5
        Me.tbAdrLivTel.Text = "0"
        '
        'tbAdrLivFax
        '
        Me.tbAdrLivFax.AcceptsReturn = True
        Me.tbAdrLivFax.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivFax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivFax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivFax.Location = New System.Drawing.Point(96, 136)
        Me.tbAdrLivFax.MaxLength = 0
        Me.tbAdrLivFax.Name = "tbAdrLivFax"
        Me.tbAdrLivFax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivFax.Size = New System.Drawing.Size(129, 20)
        Me.tbAdrLivFax.TabIndex = 6
        Me.tbAdrLivFax.Text = "0"
        '
        'tbAdrLivEmail
        '
        Me.tbAdrLivEmail.AcceptsReturn = True
        Me.tbAdrLivEmail.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLivEmail.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLivEmail.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLivEmail.Location = New System.Drawing.Point(368, 136)
        Me.tbAdrLivEmail.MaxLength = 0
        Me.tbAdrLivEmail.Name = "tbAdrLivEmail"
        Me.tbAdrLivEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLivEmail.Size = New System.Drawing.Size(201, 20)
        Me.tbAdrLivEmail.TabIndex = 8
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(342, 384)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(160, 24)
        Me.Label33.TabIndex = 104
        Me.Label33.Text = "Banque"
        '
        'tbBanque
        '
        Me.tbBanque.Location = New System.Drawing.Point(510, 384)
        Me.tbBanque.Name = "tbBanque"
        Me.tbBanque.Size = New System.Drawing.Size(312, 20)
        Me.tbBanque.TabIndex = 2
        '
        'tbRib4
        '
        Me.tbRib4.AcceptsReturn = True
        Me.tbRib4.BackColor = System.Drawing.SystemColors.Window
        Me.tbRib4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRib4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRib4.Location = New System.Drawing.Point(766, 408)
        Me.tbRib4.MaxLength = 0
        Me.tbRib4.Name = "tbRib4"
        Me.tbRib4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRib4.Size = New System.Drawing.Size(24, 20)
        Me.tbRib4.TabIndex = 6
        '
        'tbRib3
        '
        Me.tbRib3.AcceptsReturn = True
        Me.tbRib3.BackColor = System.Drawing.SystemColors.Window
        Me.tbRib3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRib3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRib3.Location = New System.Drawing.Point(614, 408)
        Me.tbRib3.MaxLength = 0
        Me.tbRib3.Name = "tbRib3"
        Me.tbRib3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRib3.Size = New System.Drawing.Size(144, 20)
        Me.tbRib3.TabIndex = 5
        '
        'tbRib2
        '
        Me.tbRib2.AcceptsReturn = True
        Me.tbRib2.BackColor = System.Drawing.SystemColors.Window
        Me.tbRib2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRib2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRib2.Location = New System.Drawing.Point(558, 408)
        Me.tbRib2.MaxLength = 0
        Me.tbRib2.Name = "tbRib2"
        Me.tbRib2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRib2.Size = New System.Drawing.Size(48, 20)
        Me.tbRib2.TabIndex = 4
        '
        'tbRib1
        '
        Me.tbRib1.AcceptsReturn = True
        Me.tbRib1.BackColor = System.Drawing.SystemColors.Window
        Me.tbRib1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRib1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRib1.Location = New System.Drawing.Point(510, 408)
        Me.tbRib1.MaxLength = 0
        Me.tbRib1.Name = "tbRib1"
        Me.tbRib1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRib1.Size = New System.Drawing.Size(40, 20)
        Me.tbRib1.TabIndex = 3
        '
        'RIB
        '
        Me.RIB.Location = New System.Drawing.Point(342, 416)
        Me.RIB.Name = "RIB"
        Me.RIB.Size = New System.Drawing.Size(104, 16)
        Me.RIB.TabIndex = 98
        Me.RIB.Text = "RIB"
        '
        'tbTVAIntra
        '
        Me.tbTVAIntra.Location = New System.Drawing.Point(112, 408)
        Me.tbTVAIntra.Name = "tbTVAIntra"
        Me.tbTVAIntra.Size = New System.Drawing.Size(216, 20)
        Me.tbTVAIntra.TabIndex = 1
        '
        'Label28
        '
        Me.Label28.Location = New System.Drawing.Point(8, 408)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(96, 24)
        Me.Label28.TabIndex = 94
        Me.Label28.Text = "N° TVA Intra"
        '
        'tbSIRET
        '
        Me.tbSIRET.Location = New System.Drawing.Point(112, 384)
        Me.tbSIRET.Name = "tbSIRET"
        Me.tbSIRET.Size = New System.Drawing.Size(216, 20)
        Me.tbSIRET.TabIndex = 0
        '
        'Label31
        '
        Me.Label31.Location = New System.Drawing.Point(8, 384)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(96, 24)
        Me.Label31.TabIndex = 92
        Me.Label31.Text = "N° SIRET"
        '
        'ckAdrIdentiques
        '
        Me.ckAdrIdentiques.Location = New System.Drawing.Point(8, 184)
        Me.ckAdrIdentiques.Name = "ckAdrIdentiques"
        Me.ckAdrIdentiques.Size = New System.Drawing.Size(200, 16)
        Me.ckAdrIdentiques.TabIndex = 87
        Me.ckAdrIdentiques.Text = "Identiques"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactNom)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactPortable)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactRue1)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactRue2)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactCP)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactVille)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactTel)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactFax)
        Me.GroupBox1.Controls.Add(Me.tbAdrFactEmail)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 208)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(928, 168)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Adresse de Facturation"
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(248, 136)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(112, 16)
        Me.Label19.TabIndex = 97
        Me.Label19.Text = "Email"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(248, 112)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(112, 16)
        Me.Label18.TabIndex = 96
        Me.Label18.Text = "Portable"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 136)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(80, 16)
        Me.Label14.TabIndex = 95
        Me.Label14.Text = "Fax"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 112)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(80, 16)
        Me.Label13.TabIndex = 94
        Me.Label13.Text = "Téléphone"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 88)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 16)
        Me.Label12.TabIndex = 93
        Me.Label12.Text = "CP/Ville"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 24)
        Me.Label11.TabIndex = 92
        Me.Label11.Text = "Nom"
        '
        'tbAdrFactNom
        '
        Me.tbAdrFactNom.Location = New System.Drawing.Point(96, 16)
        Me.tbAdrFactNom.Name = "tbAdrFactNom"
        Me.tbAdrFactNom.Size = New System.Drawing.Size(416, 20)
        Me.tbAdrFactNom.TabIndex = 0
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 64)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 16)
        Me.Label10.TabIndex = 90
        Me.Label10.Text = "Adresse2"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 89
        Me.Label5.Text = "Adresse1"
        '
        'tbAdrFactPortable
        '
        Me.tbAdrFactPortable.AcceptsReturn = True
        Me.tbAdrFactPortable.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFactPortable.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFactPortable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFactPortable.Location = New System.Drawing.Point(368, 112)
        Me.tbAdrFactPortable.MaxLength = 0
        Me.tbAdrFactPortable.Name = "tbAdrFactPortable"
        Me.tbAdrFactPortable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFactPortable.Size = New System.Drawing.Size(145, 20)
        Me.tbAdrFactPortable.TabIndex = 7
        '
        'tbAdrFactRue1
        '
        Me.tbAdrFactRue1.AcceptsReturn = True
        Me.tbAdrFactRue1.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFactRue1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFactRue1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFactRue1.Location = New System.Drawing.Point(96, 40)
        Me.tbAdrFactRue1.MaxLength = 0
        Me.tbAdrFactRue1.Name = "tbAdrFactRue1"
        Me.tbAdrFactRue1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFactRue1.Size = New System.Drawing.Size(417, 20)
        Me.tbAdrFactRue1.TabIndex = 1
        '
        'tbAdrFactRue2
        '
        Me.tbAdrFactRue2.AcceptsReturn = True
        Me.tbAdrFactRue2.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFactRue2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFactRue2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFactRue2.Location = New System.Drawing.Point(96, 64)
        Me.tbAdrFactRue2.MaxLength = 0
        Me.tbAdrFactRue2.Name = "tbAdrFactRue2"
        Me.tbAdrFactRue2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFactRue2.Size = New System.Drawing.Size(417, 20)
        Me.tbAdrFactRue2.TabIndex = 2
        '
        'tbAdrFactCP
        '
        Me.tbAdrFactCP.AcceptsReturn = True
        Me.tbAdrFactCP.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFactCP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFactCP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFactCP.Location = New System.Drawing.Point(96, 88)
        Me.tbAdrFactCP.MaxLength = 0
        Me.tbAdrFactCP.Name = "tbAdrFactCP"
        Me.tbAdrFactCP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFactCP.Size = New System.Drawing.Size(73, 20)
        Me.tbAdrFactCP.TabIndex = 3
        Me.tbAdrFactCP.Text = "0"
        '
        'tbAdrFactVille
        '
        Me.tbAdrFactVille.AcceptsReturn = True
        Me.tbAdrFactVille.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFactVille.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFactVille.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFactVille.Location = New System.Drawing.Point(176, 88)
        Me.tbAdrFactVille.MaxLength = 0
        Me.tbAdrFactVille.Name = "tbAdrFactVille"
        Me.tbAdrFactVille.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFactVille.Size = New System.Drawing.Size(337, 20)
        Me.tbAdrFactVille.TabIndex = 4
        '
        'tbAdrFactTel
        '
        Me.tbAdrFactTel.AcceptsReturn = True
        Me.tbAdrFactTel.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFactTel.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFactTel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFactTel.Location = New System.Drawing.Point(96, 112)
        Me.tbAdrFactTel.MaxLength = 0
        Me.tbAdrFactTel.Name = "tbAdrFactTel"
        Me.tbAdrFactTel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFactTel.Size = New System.Drawing.Size(129, 20)
        Me.tbAdrFactTel.TabIndex = 5
        Me.tbAdrFactTel.Text = "0"
        '
        'tbAdrFactFax
        '
        Me.tbAdrFactFax.AcceptsReturn = True
        Me.tbAdrFactFax.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFactFax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFactFax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFactFax.Location = New System.Drawing.Point(96, 136)
        Me.tbAdrFactFax.MaxLength = 0
        Me.tbAdrFactFax.Name = "tbAdrFactFax"
        Me.tbAdrFactFax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFactFax.Size = New System.Drawing.Size(129, 20)
        Me.tbAdrFactFax.TabIndex = 6
        Me.tbAdrFactFax.Text = "0"
        '
        'tbAdrFactEmail
        '
        Me.tbAdrFactEmail.AcceptsReturn = True
        Me.tbAdrFactEmail.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFactEmail.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFactEmail.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFactEmail.Location = New System.Drawing.Point(368, 136)
        Me.tbAdrFactEmail.MaxLength = 0
        Me.tbAdrFactEmail.Name = "tbAdrFactEmail"
        Me.tbAdrFactEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFactEmail.Size = New System.Drawing.Size(201, 20)
        Me.tbAdrFactEmail.TabIndex = 8
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(696, 344)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(41, 17)
        Me.Label17.TabIndex = 82
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label16.Location = New System.Drawing.Point(696, 320)
        Me.Label16.Name = "Label16"
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.Size = New System.Drawing.Size(57, 17)
        Me.Label16.TabIndex = 81
        '
        'cboModeRglmt
        '
        Me.cboModeRglmt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboModeRglmt.Location = New System.Drawing.Point(510, 434)
        Me.cboModeRglmt.Name = "cboModeRglmt"
        Me.cboModeRglmt.Size = New System.Drawing.Size(312, 21)
        Me.cboModeRglmt.TabIndex = 7
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(8, 440)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(129, 42)
        Me.Label25.TabIndex = 96
        Me.Label25.Text = "Modes de règlement"
        '
        'tpComs
        '
        Me.tpComs.Controls.Add(Me.lartbCom4)
        Me.tpComs.Controls.Add(Me.lartbCom3)
        Me.tpComs.Controls.Add(Me.lartbCom2)
        Me.tpComs.Controls.Add(Me.lartbcom1)
        Me.tpComs.Controls.Add(Me.rtbCom1)
        Me.tpComs.Controls.Add(Me.rtbCom2)
        Me.tpComs.Controls.Add(Me.rtbCom3)
        Me.tpComs.Controls.Add(Me.rtbCom4)
        Me.tpComs.Location = New System.Drawing.Point(4, 22)
        Me.tpComs.Name = "tpComs"
        Me.tpComs.Size = New System.Drawing.Size(944, 531)
        Me.tpComs.TabIndex = 2
        Me.tpComs.Text = "Commentaires"
        Me.tpComs.UseVisualStyleBackColor = True
        Me.tpComs.Visible = False
        '
        'lartbCom4
        '
        Me.lartbCom4.Location = New System.Drawing.Point(8, 256)
        Me.lartbCom4.Name = "lartbCom4"
        Me.lartbCom4.Size = New System.Drawing.Size(88, 32)
        Me.lartbCom4.TabIndex = 63
        Me.lartbCom4.Text = "Libre"
        '
        'lartbCom3
        '
        Me.lartbCom3.Location = New System.Drawing.Point(8, 176)
        Me.lartbCom3.Name = "lartbCom3"
        Me.lartbCom3.Size = New System.Drawing.Size(80, 40)
        Me.lartbCom3.TabIndex = 62
        Me.lartbCom3.Text = "Facture"
        '
        'lartbCom2
        '
        Me.lartbCom2.Location = New System.Drawing.Point(8, 96)
        Me.lartbCom2.Name = "lartbCom2"
        Me.lartbCom2.Size = New System.Drawing.Size(80, 24)
        Me.lartbCom2.TabIndex = 61
        Me.lartbCom2.Text = "BL"
        '
        'lartbcom1
        '
        Me.lartbcom1.Location = New System.Drawing.Point(8, 8)
        Me.lartbcom1.Name = "lartbcom1"
        Me.lartbcom1.Size = New System.Drawing.Size(88, 32)
        Me.lartbcom1.TabIndex = 60
        Me.lartbcom1.Text = "Commande"
        '
        'rtbCom1
        '
        Me.rtbCom1.Location = New System.Drawing.Point(104, 8)
        Me.rtbCom1.Name = "rtbCom1"
        Me.rtbCom1.Size = New System.Drawing.Size(824, 73)
        Me.rtbCom1.TabIndex = 52
        Me.rtbCom1.Text = ""
        '
        'rtbCom2
        '
        Me.rtbCom2.Location = New System.Drawing.Point(104, 88)
        Me.rtbCom2.Name = "rtbCom2"
        Me.rtbCom2.Size = New System.Drawing.Size(824, 73)
        Me.rtbCom2.TabIndex = 53
        Me.rtbCom2.Text = ""
        '
        'rtbCom3
        '
        Me.rtbCom3.Location = New System.Drawing.Point(104, 168)
        Me.rtbCom3.Name = "rtbCom3"
        Me.rtbCom3.Size = New System.Drawing.Size(824, 73)
        Me.rtbCom3.TabIndex = 54
        Me.rtbCom3.Text = ""
        '
        'rtbCom4
        '
        Me.rtbCom4.Location = New System.Drawing.Point(104, 248)
        Me.rtbCom4.Name = "rtbCom4"
        Me.rtbCom4.Size = New System.Drawing.Size(824, 73)
        Me.rtbCom4.TabIndex = 55
        Me.rtbCom4.Text = ""
        '
        'Label400
        '
        Me.Label400.AutoSize = True
        Me.Label400.Location = New System.Drawing.Point(224, 10)
        Me.Label400.Name = "Label400"
        Me.Label400.Size = New System.Drawing.Size(52, 13)
        Me.Label400.TabIndex = 69
        Me.Label400.Text = "Compta : "
        '
        'tbCodeCompta
        '
        Me.tbCodeCompta.Location = New System.Drawing.Point(283, 8)
        Me.tbCodeCompta.Name = "tbCodeCompta"
        Me.tbCodeCompta.Size = New System.Drawing.Size(60, 20)
        Me.tbCodeCompta.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(352, 11)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 13)
        Me.Label4.TabIndex = 71
        Me.Label4.Text = "ID Prestashop :"
        '
        'tbIDPrestashop
        '
        Me.tbIDPrestashop.Location = New System.Drawing.Point(438, 8)
        Me.tbIDPrestashop.Name = "tbIDPrestashop"
        Me.tbIDPrestashop.Size = New System.Drawing.Size(45, 20)
        Me.tbIDPrestashop.TabIndex = 2
        '
        'frmTiers
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.tbIDPrestashop)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbCodeCompta)
        Me.Controls.Add(Me.Label400)
        Me.Controls.Add(Me.tbRaisonSociale)
        Me.Controls.Add(Me.tbNom)
        Me.Controls.Add(Me.tbCode)
        Me.Controls.Add(Me.tabTiers)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Name = "frmTiers"
        Me.Text = "frmTiers"
        Me.tabTiers.ResumeLayout(False)
        Me.tpAddresses.ResumeLayout(False)
        Me.tpAddresses.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.tpComs.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region



    Public Overrides Function AfficheElement() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing, "Element courant is nothing")

        Dim objTiers As Tiers
        objTiers = getElementCourant()

        Me.tbCode.Text = objTiers.code
        Me.tbCodeCompta.Text = objTiers.CodeCompta
        Me.tbNom.Text = objTiers.nom
        Me.tbRaisonSociale.Text = objTiers.rs
        Me.tbIDPrestashop.Text = objTiers.idPrestashop

        Me.tbAdrLivNom.Text = objTiers.AdresseLivraison.nom
        Me.tbAdrLivRue1.Text = objTiers.AdresseLivraison.rue1
        Me.tbAdrLivRue2.Text = objTiers.AdresseLivraison.rue2
        Me.tbAdrLivCP.Text = objTiers.AdresseLivraison.cp
        Me.tbAdrLivVille.Text = objTiers.AdresseLivraison.ville
        Me.tbAdrLivTel.Text = objTiers.AdresseLivraison.tel
        Me.tbAdrLivFax.Text = objTiers.AdresseLivraison.fax
        Me.tbAdrLivPortable.Text = objTiers.AdresseLivraison.port
        Me.tbAdrLivEmail.Text = objTiers.AdresseLivraison.Email
        Me.ckAdrIdentiques.Checked = objTiers.bAdressesIdentiques
        Me.tbAdrFactNom.Text = objTiers.AdresseFacturation.nom
        Me.tbAdrFactRue1.Text = objTiers.AdresseFacturation.rue1
        Me.tbAdrFactRue2.Text = objTiers.AdresseFacturation.rue2
        Me.tbAdrFactCP.Text = objTiers.AdresseFacturation.cp
        Me.tbAdrFactVille.Text = objTiers.AdresseFacturation.ville
        Me.tbAdrFactTel.Text = objTiers.AdresseFacturation.tel
        Me.tbAdrFactFax.Text = objTiers.AdresseFacturation.fax
        Me.tbAdrFactPortable.Text = objTiers.AdresseFacturation.port
        Me.tbAdrFactEmail.Text = objTiers.AdresseFacturation.Email

        Me.tbSIRET.Text = objTiers.siret
        Me.tbTVAIntra.Text = objTiers.numtvaintracom

        Me.tbRib1.Text = objTiers.rib1
        Me.tbRib2.Text = objTiers.rib2
        Me.tbRib3.Text = objTiers.rib3
        Me.tbRib4.Text = objTiers.rib4
        Me.tbBanque.Text = objTiers.banque


        'Mode de reglement
        Dim objParam As Param
        For Each objParam In cboModeRglmt.Items
            If objParam.id = objTiers.idModeReglement Then
                cboModeRglmt.SelectedItem = objParam
                Exit For
            End If
        Next
        For Each objParam In cboModeReglement1.Items
            If objParam.id = objTiers.idModeReglement1 Then
                cboModeReglement1.SelectedItem = objParam
                Exit For
            End If
        Next
        For Each objParam In cboModeReglement2.Items
            If objParam.id = objTiers.idModeReglement2 Then
                cboModeReglement2.SelectedItem = objParam
                Exit For
            End If
        Next
        'Commentaires
        rtbCom1.Text = objTiers.CommCommande.comment
        rtbCom2.Text = objTiers.CommLivraison.comment
        rtbCom3.Text = objTiers.CommFacturation.comment
        rtbCom4.Text = objTiers.CommLibre.comment

        EnableControls(True)

        Return True
    End Function

    Public Overrides Function SauveElement() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing)
        Dim bReturn As Boolean
        bReturn = getElementCourant.Save
        Debug.Assert(bReturn, "Erreur en sauvegarde")
        Return bReturn
    End Function
    'Doit être appellé par méthode fille
    Public Overrides Function MAJElement() As Boolean
        Dim objTiers As Tiers

        objTiers = CType(getElementCourant(), Tiers)

        objTiers.code = Me.tbCode.Text
        objTiers.CodeCompta = Me.tbCodeCompta.Text
        objTiers.nom = Me.tbNom.Text
        objTiers.rs = Me.tbRaisonSociale.Text
        objTiers.idPrestashop = Me.tbIDPrestashop.Text

        objTiers.AdresseLivraison.nom = Me.tbAdrLivNom.Text
        objTiers.AdresseLivraison.rue1 = Me.tbAdrLivRue1.Text
        objTiers.AdresseLivraison.rue2 = Me.tbAdrLivRue2.Text
        objTiers.AdresseLivraison.cp = Me.tbAdrLivCP.Text
        objTiers.AdresseLivraison.ville = Me.tbAdrLivVille.Text
        objTiers.AdresseLivraison.tel = Me.tbAdrLivTel.Text
        objTiers.AdresseLivraison.fax = Me.tbAdrLivFax.Text
        objTiers.AdresseLivraison.port = Me.tbAdrLivPortable.Text
        objTiers.AdresseLivraison.Email = Me.tbAdrLivEmail.Text
        objTiers.bAdressesIdentiques = ckAdrIdentiques.Checked
        If objTiers.bAdressesIdentiques Then
            objTiers.AdresseFacturation.nom = Me.tbAdrLivNom.Text
            objTiers.AdresseFacturation.rue1 = Me.tbAdrLivRue1.Text
            objTiers.AdresseFacturation.rue2 = Me.tbAdrLivRue2.Text
            objTiers.AdresseFacturation.cp = Me.tbAdrLivCP.Text
            objTiers.AdresseFacturation.ville = Me.tbAdrLivVille.Text
            objTiers.AdresseFacturation.tel = Me.tbAdrLivTel.Text
            objTiers.AdresseFacturation.fax = Me.tbAdrLivFax.Text
            objTiers.AdresseFacturation.port = Me.tbAdrLivPortable.Text
            objTiers.AdresseFacturation.Email = Me.tbAdrLivEmail.Text
        Else
            objTiers.AdresseFacturation.nom = Me.tbAdrFactNom.Text
            objTiers.AdresseFacturation.rue1 = Me.tbAdrFactRue1.Text
            objTiers.AdresseFacturation.rue2 = Me.tbAdrFactRue2.Text
            objTiers.AdresseFacturation.cp = Me.tbAdrFactCP.Text
            objTiers.AdresseFacturation.ville = Me.tbAdrFactVille.Text
            objTiers.AdresseFacturation.tel = Me.tbAdrFactTel.Text
            objTiers.AdresseFacturation.fax = Me.tbAdrFactFax.Text
            objTiers.AdresseFacturation.port = Me.tbAdrFactPortable.Text
            objTiers.AdresseFacturation.Email = Me.tbAdrFactEmail.Text
        End If

        objTiers.siret = Me.tbSIRET.Text
        objTiers.numtvaintracom = Me.tbTVAIntra.Text

        objTiers.rib1 = Me.tbRib1.Text
        objTiers.rib2 = Me.tbRib2.Text
        objTiers.rib3 = Me.tbRib3.Text
        objTiers.rib4 = Me.tbRib4.Text
        objTiers.banque = Me.tbBanque.Text
        objTiers.idModeReglement = cboModeRglmt.SelectedItem.id
        objTiers.idModeReglement1 = cboModeReglement1.SelectedItem.id
        objTiers.idModeReglement2 = cboModeReglement2.SelectedItem.id

        'Commentaires
        objTiers.CommCommande.comment = rtbCom1.Text
        objTiers.CommLivraison.comment = rtbCom2.Text
        objTiers.CommFacturation.comment = rtbCom3.Text
        objTiers.CommLibre.comment = rtbCom4.Text
        Return True
    End Function 'MAJElementCourant


    Private Sub ckAdrIdentiques_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckAdrIdentiques.CheckedChanged
        If ckAdrIdentiques.Checked Then
            'Dupplication de l'adresse
            DuppliqueAdresses()
            tbAdrFactNom.Enabled = False
            tbAdrFactRue1.Enabled = False
            tbAdrFactRue2.Enabled = False
            tbAdrFactCP.Enabled = False
            tbAdrFactVille.Enabled = False
            tbAdrFactTel.Enabled = False
            tbAdrFactPortable.Enabled = False
            tbAdrFactFax.Enabled = False
            tbAdrFactEmail.Enabled = False
        Else
            tbAdrFactNom.Enabled = True
            tbAdrFactRue1.Enabled = True
            tbAdrFactRue2.Enabled = True
            tbAdrFactCP.Enabled = True
            tbAdrFactVille.Enabled = True
            tbAdrFactTel.Enabled = True
            tbAdrFactPortable.Enabled = True
            tbAdrFactFax.Enabled = True
            tbAdrFactEmail.Enabled = True
        End If
    End Sub

    Private Sub DuppliqueAdresses()
        tbAdrFactNom.Text = tbAdrLivNom.Text
        tbAdrFactRue1.Text = tbAdrLivRue1.Text
        tbAdrFactRue2.Text = tbAdrLivRue2.Text
        tbAdrFactCP.Text = tbAdrLivCP.Text
        tbAdrFactVille.Text = tbAdrLivVille.Text
        tbAdrFactTel.Text = tbAdrLivTel.Text
        tbAdrFactFax.Text = tbAdrLivFax.Text
        tbAdrFactPortable.Text = tbAdrLivPortable.Text
        tbAdrFactEmail.Text = tbAdrLivEmail.Text
    End Sub

    Public Overrides Function ControleAvantSauvegarde() As Boolean
        Dim bReturn As Boolean
        bReturn = True
        If Len(tbNom.Text) = 0 Then
            MsgBox("Le Nom doit être renseigné")
            tbNom.Focus()
            bReturn = False
        End If
        If Len(tbCode.Text) = 0 Then
            MsgBox("Le Code doit être renseigné")
            tbCode.Focus()
            bReturn = False
        End If
        If Len(tbRib1.Text) > 5 Then
            MsgBox("Le code banque RIB doit être sur 5 caractères maxi")
            tbRib1.Focus()
            bReturn = False
        End If
        If Len(tbRib2.Text) > 5 Then
            MsgBox("Le code Guichet RIB doit être sur 5 caractères maxi")
            tbRib2.Focus()
            bReturn = False
        End If
        If Len(tbRib3.Text) > 11 Then
            MsgBox("Le N° compte  RIB doit être sur 11 caractères maxi")
            tbRib3.Focus()
            bReturn = False
        End If
        If Len(tbRib4.Text) > 2 Then
            MsgBox("La clé RIB doit être sur 2 caractères maxi ")
            tbRib4.Focus()
            bReturn = False
        End If


        'If Not Controle_Clé_Rib(tbRib1.Text, tbRib2.Text, tbRib3.Text, tbRib4.Text) Then
        ' MsgBox("RIB Incorrect")
        ' tbRib1.Focus()
        ' bReturn = False
        'End If
        Return bReturn
    End Function
    Private Sub initFenetre()
        debAffiche()
        InitComboModeReglement(cboModeRglmt)
        InitComboModeReglement(cboModeReglement1)
        InitComboModeReglement(cboModeReglement2)
        finAffiche()
    End Sub 'initFenetre

    Private Sub frmTiers_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (Not DesignMode) Then
            initFenetre()
        End If
    End Sub

    Protected Overrides Function frmNew() As Boolean
        MyBase.frmNew()
        tbCode.Focus()
    End Function

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)
        If m_action = vncEnums.vncfrmAction.FRMLOAD Then
            'Modification 
            Me.tbCode.Enabled = False
        End If
    End Sub
End Class
