Option Strict Off
Imports vini_DB
Public Class frmGestionSCMD
    Inherits frmDonBase
    'Protected getElementCourant() As SousCommande
    Protected m_colSousCommandes As Collection

#Region "Code généré par le Concepteur Windows Form "
    Public Sub New()
        MyBase.New()
        'Cet appel est requis par le Concepteur Windows Form.
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
    Public ToolTip1 As System.Windows.Forms.ToolTip
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    Public WithEvents tbDateFactFourn As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cbAppliquer As System.Windows.Forms.Button
    Friend WithEvents liFournisseur As System.Windows.Forms.LinkLabel
    Friend WithEvents liClient As System.Windows.Forms.LinkLabel
    Friend WithEvents liCommande As System.Windows.Forms.LinkLabel
    Friend WithEvents ckRapprochee As System.Windows.Forms.CheckBox
    Public WithEvents tbFACTtotalHT As textBoxCurrency
    Public WithEvents tbFACTreference As System.Windows.Forms.TextBox
    Public WithEvents tbFACTTotalTTC As textBoxCurrency
    Public WithEvents tbCMDTotalHT As textBoxCurrency
    Public WithEvents tbCMDtotalTTC As textBoxCurrency
    Friend WithEvents dtFACTdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents grpDetailSousCommande As System.Windows.Forms.GroupBox
    Friend WithEvents tbComfournisseur As System.Windows.Forms.RichTextBox
    Friend WithEvents ckFacturee As System.Windows.Forms.CheckBox
    Friend WithEvents label18 As System.Windows.Forms.Label
    Friend WithEvents liFactCom As System.Windows.Forms.LinkLabel
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents cbTransmettreFax As System.Windows.Forms.Button
    Friend WithEvents ckTransmiseFax As System.Windows.Forms.CheckBox
    Friend WithEvents ckExporteeInternet As System.Windows.Forms.CheckBox
    Friend WithEvents cbExportInternet As System.Windows.Forms.Button
    Friend WithEvents ckImporteeInternet As System.Windows.Forms.CheckBox
    Friend WithEvents m_bsrcSousCommandes As System.Windows.Forms.BindingSource
    Friend WithEvents m_dgvScmd As System.Windows.Forms.DataGridView
    Friend WithEvents code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dateCommande As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalHT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalTTC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FournisseurRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents montantCommission As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FournisseurCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateFactureDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EtatLibelleDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents tbBaseComm As vini_app.textBoxCurrency
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cbCalc As System.Windows.Forms.Button
    Public WithEvents tbMontantCommision As vini_app.textBoxCurrency
    Public WithEvents tbTxCommision As vini_app.textBoxNumeric
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents grpDateScmd As System.Windows.Forms.GroupBox
    Friend WithEvents tbCodeFournisseur As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtdateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateFin As System.Windows.Forms.Label
    Friend WithEvents dtDatedeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateDebut As System.Windows.Forms.Label
    Friend WithEvents rbDateSCMD As System.Windows.Forms.RadioButton
    Friend WithEvents rbCodeScmd As System.Windows.Forms.RadioButton
    Friend WithEvents rbIDScmd As System.Windows.Forms.RadioButton
    Friend WithEvents tbCodeScmd As System.Windows.Forms.TextBox
    Friend WithEvents tbIDScmd As System.Windows.Forms.TextBox
    Friend WithEvents tbCommentaire As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpDetailSousCommande = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbCommentaire = New System.Windows.Forms.TextBox()
        Me.m_bsrcSousCommandes = New System.Windows.Forms.BindingSource(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tbBaseComm = New vini_app.textBoxCurrency()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.cbCalc = New System.Windows.Forms.Button()
        Me.tbMontantCommision = New vini_app.textBoxCurrency()
        Me.tbTxCommision = New vini_app.textBoxNumeric()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ckImporteeInternet = New System.Windows.Forms.CheckBox()
        Me.cbTransmettreFax = New System.Windows.Forms.Button()
        Me.ckTransmiseFax = New System.Windows.Forms.CheckBox()
        Me.ckExporteeInternet = New System.Windows.Forms.CheckBox()
        Me.cbExportInternet = New System.Windows.Forms.Button()
        Me.liFactCom = New System.Windows.Forms.LinkLabel()
        Me.label18 = New System.Windows.Forms.Label()
        Me.ckFacturee = New System.Windows.Forms.CheckBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.ckRapprochee = New System.Windows.Forms.CheckBox()
        Me.tbComfournisseur = New System.Windows.Forms.RichTextBox()
        Me.liFournisseur = New System.Windows.Forms.LinkLabel()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.liClient = New System.Windows.Forms.LinkLabel()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.liCommande = New System.Windows.Forms.LinkLabel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.dtFACTdate = New System.Windows.Forms.DateTimePicker()
        Me.tbFACTtotalHT = New vini_app.textBoxCurrency()
        Me.tbFACTreference = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.tbFACTTotalTTC = New vini_app.textBoxCurrency()
        Me.tbCMDtotalTTC = New vini_app.textBoxCurrency()
        Me.tbCMDTotalHT = New vini_app.textBoxCurrency()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbAppliquer = New System.Windows.Forms.Button()
        Me.tbDateFactFourn = New System.Windows.Forms.Label()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.cbRecherche = New System.Windows.Forms.Button()
        Me.m_dgvScmd = New System.Windows.Forms.DataGridView()
        Me.code = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TiersRS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dateCommande = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.totalHT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.totalTTC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FournisseurRS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.montantCommission = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FournisseurCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EtatLibelleDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TiersCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.grpDateScmd = New System.Windows.Forms.GroupBox()
        Me.tbCodeFournisseur = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker()
        Me.lblDateFin = New System.Windows.Forms.Label()
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker()
        Me.lblDateDebut = New System.Windows.Forms.Label()
        Me.rbDateSCMD = New System.Windows.Forms.RadioButton()
        Me.rbCodeScmd = New System.Windows.Forms.RadioButton()
        Me.rbIDScmd = New System.Windows.Forms.RadioButton()
        Me.tbCodeScmd = New System.Windows.Forms.TextBox()
        Me.tbIDScmd = New System.Windows.Forms.TextBox()
        Me.grpDetailSousCommande.SuspendLayout()
        CType(Me.m_bsrcSousCommandes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.m_dgvScmd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDateScmd.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpDetailSousCommande
        '
        Me.grpDetailSousCommande.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpDetailSousCommande.Controls.Add(Me.Label6)
        Me.grpDetailSousCommande.Controls.Add(Me.tbCommentaire)
        Me.grpDetailSousCommande.Controls.Add(Me.GroupBox1)
        Me.grpDetailSousCommande.Controls.Add(Me.ckImporteeInternet)
        Me.grpDetailSousCommande.Controls.Add(Me.cbTransmettreFax)
        Me.grpDetailSousCommande.Controls.Add(Me.ckTransmiseFax)
        Me.grpDetailSousCommande.Controls.Add(Me.ckExporteeInternet)
        Me.grpDetailSousCommande.Controls.Add(Me.cbExportInternet)
        Me.grpDetailSousCommande.Controls.Add(Me.liFactCom)
        Me.grpDetailSousCommande.Controls.Add(Me.label18)
        Me.grpDetailSousCommande.Controls.Add(Me.ckFacturee)
        Me.grpDetailSousCommande.Controls.Add(Me.Label12)
        Me.grpDetailSousCommande.Controls.Add(Me.ckRapprochee)
        Me.grpDetailSousCommande.Controls.Add(Me.tbComfournisseur)
        Me.grpDetailSousCommande.Controls.Add(Me.liFournisseur)
        Me.grpDetailSousCommande.Controls.Add(Me.Label16)
        Me.grpDetailSousCommande.Controls.Add(Me.liClient)
        Me.grpDetailSousCommande.Controls.Add(Me.Label11)
        Me.grpDetailSousCommande.Controls.Add(Me.liCommande)
        Me.grpDetailSousCommande.Controls.Add(Me.Label10)
        Me.grpDetailSousCommande.Controls.Add(Me.dtFACTdate)
        Me.grpDetailSousCommande.Controls.Add(Me.tbFACTtotalHT)
        Me.grpDetailSousCommande.Controls.Add(Me.tbFACTreference)
        Me.grpDetailSousCommande.Controls.Add(Me.Label2)
        Me.grpDetailSousCommande.Controls.Add(Me.Label9)
        Me.grpDetailSousCommande.Controls.Add(Me.tbFACTTotalTTC)
        Me.grpDetailSousCommande.Controls.Add(Me.tbCMDtotalTTC)
        Me.grpDetailSousCommande.Controls.Add(Me.tbCMDTotalHT)
        Me.grpDetailSousCommande.Controls.Add(Me.Label7)
        Me.grpDetailSousCommande.Controls.Add(Me.Label3)
        Me.grpDetailSousCommande.Controls.Add(Me.cbAppliquer)
        Me.grpDetailSousCommande.Controls.Add(Me.tbDateFactFourn)
        Me.grpDetailSousCommande.Location = New System.Drawing.Point(433, 60)
        Me.grpDetailSousCommande.Name = "grpDetailSousCommande"
        Me.grpDetailSousCommande.Size = New System.Drawing.Size(552, 550)
        Me.grpDetailSousCommande.TabIndex = 7
        Me.grpDetailSousCommande.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(5, 428)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(74, 13)
        Me.Label6.TabIndex = 116
        Me.Label6.Text = "Commentaire :"
        '
        'tbCommentaire
        '
        Me.tbCommentaire.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "CommentaireLibreText", True))
        Me.tbCommentaire.Location = New System.Drawing.Point(150, 425)
        Me.tbCommentaire.Multiline = True
        Me.tbCommentaire.Name = "tbCommentaire"
        Me.tbCommentaire.Size = New System.Drawing.Size(377, 89)
        Me.tbCommentaire.TabIndex = 115
        '
        'm_bsrcSousCommandes
        '
        Me.m_bsrcSousCommandes.DataSource = GetType(vini_DB.SousCommande)
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.tbBaseComm)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.cbCalc)
        Me.GroupBox1.Controls.Add(Me.tbMontantCommision)
        Me.GroupBox1.Controls.Add(Me.tbTxCommision)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Location = New System.Drawing.Point(280, 281)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(251, 94)
        Me.GroupBox1.TabIndex = 114
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Commission"
        '
        'tbBaseComm
        '
        Me.tbBaseComm.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "baseCommission", True))
        Me.tbBaseComm.Location = New System.Drawing.Point(124, 17)
        Me.tbBaseComm.Name = "tbBaseComm"
        Me.tbBaseComm.Size = New System.Drawing.Size(120, 20)
        Me.tbBaseComm.TabIndex = 96
        Me.tbBaseComm.Text = "0"
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(18, 20)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(40, 17)
        Me.Label15.TabIndex = 102
        Me.Label15.Text = "Base : "
        '
        'cbCalc
        '
        Me.cbCalc.Location = New System.Drawing.Point(70, 66)
        Me.cbCalc.Name = "cbCalc"
        Me.cbCalc.Size = New System.Drawing.Size(48, 19)
        Me.cbCalc.TabIndex = 98
        Me.cbCalc.Text = "calcul"
        '
        'tbMontantCommision
        '
        Me.tbMontantCommision.AcceptsReturn = True
        Me.tbMontantCommision.BackColor = System.Drawing.SystemColors.Window
        Me.tbMontantCommision.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbMontantCommision.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "montantCommission", True))
        Me.tbMontantCommision.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbMontantCommision.Location = New System.Drawing.Point(124, 65)
        Me.tbMontantCommision.MaxLength = 0
        Me.tbMontantCommision.Name = "tbMontantCommision"
        Me.tbMontantCommision.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbMontantCommision.Size = New System.Drawing.Size(121, 20)
        Me.tbMontantCommision.TabIndex = 99
        Me.tbMontantCommision.Text = "0"
        '
        'tbTxCommision
        '
        Me.tbTxCommision.AcceptsReturn = True
        Me.tbTxCommision.BackColor = System.Drawing.SystemColors.Window
        Me.tbTxCommision.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbTxCommision.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "tauxCommission", True))
        Me.tbTxCommision.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbTxCommision.Location = New System.Drawing.Point(124, 41)
        Me.tbTxCommision.MaxLength = 0
        Me.tbTxCommision.Name = "tbTxCommision"
        Me.tbTxCommision.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbTxCommision.Size = New System.Drawing.Size(121, 20)
        Me.tbTxCommision.TabIndex = 97
        Me.tbTxCommision.Text = "0"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(18, 69)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(51, 17)
        Me.Label5.TabIndex = 101
        Me.Label5.Text = "Montant : "
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(18, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(51, 17)
        Me.Label4.TabIndex = 100
        Me.Label4.Text = "Taux : "
        '
        'ckImporteeInternet
        '
        Me.ckImporteeInternet.AutoSize = True
        Me.ckImporteeInternet.Location = New System.Drawing.Point(179, 236)
        Me.ckImporteeInternet.Name = "ckImporteeInternet"
        Me.ckImporteeInternet.Size = New System.Drawing.Size(106, 17)
        Me.ckImporteeInternet.TabIndex = 113
        Me.ckImporteeInternet.Text = "Importée Internet"
        Me.ckImporteeInternet.UseVisualStyleBackColor = True
        '
        'cbTransmettreFax
        '
        Me.cbTransmettreFax.Location = New System.Drawing.Point(150, 155)
        Me.cbTransmettreFax.Name = "cbTransmettreFax"
        Me.cbTransmettreFax.Size = New System.Drawing.Size(121, 23)
        Me.cbTransmettreFax.TabIndex = 111
        Me.cbTransmettreFax.Text = "Transmettre"
        Me.cbTransmettreFax.UseVisualStyleBackColor = True
        '
        'ckTransmiseFax
        '
        Me.ckTransmiseFax.AutoSize = True
        Me.ckTransmiseFax.Location = New System.Drawing.Point(6, 158)
        Me.ckTransmiseFax.Name = "ckTransmiseFax"
        Me.ckTransmiseFax.Size = New System.Drawing.Size(94, 17)
        Me.ckTransmiseFax.TabIndex = 110
        Me.ckTransmiseFax.Text = "Transmise Fax"
        Me.ckTransmiseFax.UseVisualStyleBackColor = True
        '
        'ckExporteeInternet
        '
        Me.ckExporteeInternet.AutoSize = True
        Me.ckExporteeInternet.Location = New System.Drawing.Point(6, 188)
        Me.ckExporteeInternet.Name = "ckExporteeInternet"
        Me.ckExporteeInternet.Size = New System.Drawing.Size(107, 17)
        Me.ckExporteeInternet.TabIndex = 109
        Me.ckExporteeInternet.Text = "Exportée Internet"
        Me.ckExporteeInternet.UseVisualStyleBackColor = True
        '
        'cbExportInternet
        '
        Me.cbExportInternet.Location = New System.Drawing.Point(150, 184)
        Me.cbExportInternet.Name = "cbExportInternet"
        Me.cbExportInternet.Size = New System.Drawing.Size(120, 23)
        Me.cbExportInternet.TabIndex = 108
        Me.cbExportInternet.Text = "Exporter"
        Me.cbExportInternet.UseVisualStyleBackColor = True
        '
        'liFactCom
        '
        Me.liFactCom.Location = New System.Drawing.Point(176, 376)
        Me.liFactCom.Name = "liFactCom"
        Me.liFactCom.Size = New System.Drawing.Size(348, 43)
        Me.liFactCom.TabIndex = 17
        '
        'label18
        '
        Me.label18.Location = New System.Drawing.Point(5, 376)
        Me.label18.Name = "label18"
        Me.label18.Size = New System.Drawing.Size(129, 16)
        Me.label18.TabIndex = 16
        Me.label18.Text = "Facturée : "
        '
        'ckFacturee
        '
        Me.ckFacturee.Location = New System.Drawing.Point(150, 375)
        Me.ckFacturee.Name = "ckFacturee"
        Me.ckFacturee.Size = New System.Drawing.Size(16, 16)
        Me.ckFacturee.TabIndex = 107
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(6, 128)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(122, 16)
        Me.Label12.TabIndex = 106
        Me.Label12.Text = "Total TTC commande:"
        '
        'ckRapprochee
        '
        Me.ckRapprochee.Location = New System.Drawing.Point(6, 237)
        Me.ckRapprochee.Name = "ckRapprochee"
        Me.ckRapprochee.Size = New System.Drawing.Size(169, 16)
        Me.ckRapprochee.TabIndex = 5
        Me.ckRapprochee.Text = "Rapprochée manuellement"
        '
        'tbComfournisseur
        '
        Me.tbComfournisseur.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.tbComfournisseur.Location = New System.Drawing.Point(392, 104)
        Me.tbComfournisseur.Name = "tbComfournisseur"
        Me.tbComfournisseur.Size = New System.Drawing.Size(136, 136)
        Me.tbComfournisseur.TabIndex = 103
        Me.tbComfournisseur.Text = ""
        '
        'liFournisseur
        '
        Me.liFournisseur.Location = New System.Drawing.Point(89, 40)
        Me.liFournisseur.Name = "liFournisseur"
        Me.liFournisseur.Size = New System.Drawing.Size(457, 24)
        Me.liFournisseur.TabIndex = 1
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(6, 40)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(72, 24)
        Me.Label16.TabIndex = 97
        Me.Label16.Text = "Fournisseur :"
        '
        'liClient
        '
        Me.liClient.Location = New System.Drawing.Point(89, 64)
        Me.liClient.Name = "liClient"
        Me.liClient.Size = New System.Drawing.Size(457, 24)
        Me.liClient.TabIndex = 2
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(6, 64)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 16)
        Me.Label11.TabIndex = 90
        Me.Label11.Text = "Client :"
        '
        'liCommande
        '
        Me.liCommande.Location = New System.Drawing.Point(89, 16)
        Me.liCommande.Name = "liCommande"
        Me.liCommande.Size = New System.Drawing.Size(457, 24)
        Me.liCommande.TabIndex = 0
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(6, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 24)
        Me.Label10.TabIndex = 88
        Me.Label10.Text = "Commande :"
        '
        'dtFACTdate
        '
        Me.dtFACTdate.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcSousCommandes, "dateFactFournisseur", True))
        Me.dtFACTdate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFACTdate.Location = New System.Drawing.Point(149, 353)
        Me.dtFACTdate.Name = "dtFACTdate"
        Me.dtFACTdate.Size = New System.Drawing.Size(120, 20)
        Me.dtFACTdate.TabIndex = 9
        '
        'tbFACTtotalHT
        '
        Me.tbFACTtotalHT.AcceptsReturn = True
        Me.tbFACTtotalHT.BackColor = System.Drawing.SystemColors.Window
        Me.tbFACTtotalHT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbFACTtotalHT.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "totalHTFacture", True))
        Me.tbFACTtotalHT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbFACTtotalHT.Location = New System.Drawing.Point(149, 305)
        Me.tbFACTtotalHT.MaxLength = 0
        Me.tbFACTtotalHT.Name = "tbFACTtotalHT"
        Me.tbFACTtotalHT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbFACTtotalHT.Size = New System.Drawing.Size(121, 20)
        Me.tbFACTtotalHT.TabIndex = 7
        Me.tbFACTtotalHT.Text = "0"
        '
        'tbFACTreference
        '
        Me.tbFACTreference.AcceptsReturn = True
        Me.tbFACTreference.BackColor = System.Drawing.SystemColors.Window
        Me.tbFACTreference.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbFACTreference.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "refFactFournisseur", True))
        Me.tbFACTreference.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbFACTreference.Location = New System.Drawing.Point(149, 281)
        Me.tbFACTreference.MaxLength = 0
        Me.tbFACTreference.Name = "tbFACTreference"
        Me.tbFACTreference.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbFACTreference.Size = New System.Drawing.Size(121, 20)
        Me.tbFACTreference.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(5, 332)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(129, 17)
        Me.Label2.TabIndex = 84
        Me.Label2.Text = "Total TTC Facture :"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(5, 308)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(137, 17)
        Me.Label9.TabIndex = 82
        Me.Label9.Text = "Total H.T Facture :"
        '
        'tbFACTTotalTTC
        '
        Me.tbFACTTotalTTC.AcceptsReturn = True
        Me.tbFACTTotalTTC.BackColor = System.Drawing.SystemColors.Window
        Me.tbFACTTotalTTC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbFACTTotalTTC.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "totalTTCFacture", True))
        Me.tbFACTTotalTTC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbFACTTotalTTC.Location = New System.Drawing.Point(149, 329)
        Me.tbFACTTotalTTC.MaxLength = 0
        Me.tbFACTTotalTTC.Name = "tbFACTTotalTTC"
        Me.tbFACTTotalTTC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbFACTTotalTTC.Size = New System.Drawing.Size(121, 20)
        Me.tbFACTTotalTTC.TabIndex = 8
        Me.tbFACTTotalTTC.Text = "0"
        '
        'tbCMDtotalTTC
        '
        Me.tbCMDtotalTTC.AcceptsReturn = True
        Me.tbCMDtotalTTC.BackColor = System.Drawing.SystemColors.Window
        Me.tbCMDtotalTTC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCMDtotalTTC.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "totalTTC", True))
        Me.tbCMDtotalTTC.Enabled = False
        Me.tbCMDtotalTTC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCMDtotalTTC.Location = New System.Drawing.Point(150, 129)
        Me.tbCMDtotalTTC.MaxLength = 0
        Me.tbCMDtotalTTC.Name = "tbCMDtotalTTC"
        Me.tbCMDtotalTTC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCMDtotalTTC.Size = New System.Drawing.Size(121, 20)
        Me.tbCMDtotalTTC.TabIndex = 4
        Me.tbCMDtotalTTC.Text = "0"
        '
        'tbCMDTotalHT
        '
        Me.tbCMDTotalHT.AcceptsReturn = True
        Me.tbCMDTotalHT.BackColor = System.Drawing.SystemColors.Window
        Me.tbCMDTotalHT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCMDTotalHT.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommandes, "totalHT", True))
        Me.tbCMDTotalHT.Enabled = False
        Me.tbCMDTotalHT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCMDTotalHT.Location = New System.Drawing.Point(150, 105)
        Me.tbCMDTotalHT.MaxLength = 0
        Me.tbCMDTotalHT.Name = "tbCMDTotalHT"
        Me.tbCMDTotalHT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCMDTotalHT.Size = New System.Drawing.Size(121, 20)
        Me.tbCMDTotalHT.TabIndex = 3
        Me.tbCMDTotalHT.Text = "0"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(5, 284)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(137, 25)
        Me.Label7.TabIndex = 78
        Me.Label7.Text = "Ref Facture Fournisseur :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(6, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(137, 17)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Total HT commande :"
        '
        'cbAppliquer
        '
        Me.cbAppliquer.Location = New System.Drawing.Point(415, 520)
        Me.cbAppliquer.Name = "cbAppliquer"
        Me.cbAppliquer.Size = New System.Drawing.Size(112, 24)
        Me.cbAppliquer.TabIndex = 18
        Me.cbAppliquer.Text = "Appliquer"
        '
        'tbDateFactFourn
        '
        Me.tbDateFactFourn.BackColor = System.Drawing.SystemColors.Control
        Me.tbDateFactFourn.Cursor = System.Windows.Forms.Cursors.Default
        Me.tbDateFactFourn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.tbDateFactFourn.Location = New System.Drawing.Point(5, 357)
        Me.tbDateFactFourn.Name = "tbDateFactFourn"
        Me.tbDateFactFourn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbDateFactFourn.Size = New System.Drawing.Size(151, 16)
        Me.tbDateFactFourn.TabIndex = 80
        Me.tbDateFactFourn.Text = "Date Facture Fournisseur :"
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.BackColor = System.Drawing.SystemColors.Control
        Me.cbAfficher.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAfficher.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAfficher.Location = New System.Drawing.Point(8, 114)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAfficher.Size = New System.Drawing.Size(408, 23)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "A&fficher"
        Me.cbAfficher.UseVisualStyleBackColor = False
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(293, 60)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(90, 24)
        Me.cbRecherche.TabIndex = 3
        Me.cbRecherche.Text = "Recherche"
        '
        'm_dgvScmd
        '
        Me.m_dgvScmd.AllowUserToAddRows = False
        Me.m_dgvScmd.AllowUserToDeleteRows = False
        Me.m_dgvScmd.AllowUserToResizeRows = False
        Me.m_dgvScmd.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.m_dgvScmd.AutoGenerateColumns = False
        Me.m_dgvScmd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.m_dgvScmd.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.code, Me.TiersRS, Me.dateCommande, Me.totalHT, Me.totalTTC, Me.FournisseurRS, Me.montantCommission, Me.FournisseurCodeDataGridViewTextBoxColumn, Me.EtatLibelleDataGridViewTextBoxColumn, Me.TiersCodeDataGridViewTextBoxColumn})
        Me.m_dgvScmd.DataSource = Me.m_bsrcSousCommandes
        Me.m_dgvScmd.Location = New System.Drawing.Point(13, 143)
        Me.m_dgvScmd.Name = "m_dgvScmd"
        Me.m_dgvScmd.ReadOnly = True
        Me.m_dgvScmd.RowHeadersWidth = 10
        Me.m_dgvScmd.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.m_dgvScmd.Size = New System.Drawing.Size(403, 467)
        Me.m_dgvScmd.TabIndex = 102
        '
        'code
        '
        Me.code.DataPropertyName = "code"
        Me.code.HeaderText = "code"
        Me.code.Name = "code"
        Me.code.ReadOnly = True
        Me.code.Width = 56
        '
        'TiersRS
        '
        Me.TiersRS.DataPropertyName = "TiersRS"
        Me.TiersRS.HeaderText = "Client"
        Me.TiersRS.Name = "TiersRS"
        Me.TiersRS.ReadOnly = True
        Me.TiersRS.Width = 58
        '
        'dateCommande
        '
        Me.dateCommande.DataPropertyName = "dateCommande"
        Me.dateCommande.HeaderText = "date"
        Me.dateCommande.Name = "dateCommande"
        Me.dateCommande.ReadOnly = True
        Me.dateCommande.Width = 53
        '
        'totalHT
        '
        Me.totalHT.DataPropertyName = "totalHT"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle1.Format = "C2"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.totalHT.DefaultCellStyle = DataGridViewCellStyle1
        Me.totalHT.HeaderText = "H.T."
        Me.totalHT.Name = "totalHT"
        Me.totalHT.ReadOnly = True
        Me.totalHT.Width = 53
        '
        'totalTTC
        '
        Me.totalTTC.DataPropertyName = "totalTTC"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.Format = "C2"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.totalTTC.DefaultCellStyle = DataGridViewCellStyle2
        Me.totalTTC.HeaderText = "T.T.C."
        Me.totalTTC.Name = "totalTTC"
        Me.totalTTC.ReadOnly = True
        Me.totalTTC.Width = 62
        '
        'FournisseurRS
        '
        Me.FournisseurRS.DataPropertyName = "FournisseurRS"
        Me.FournisseurRS.HeaderText = "Producteur"
        Me.FournisseurRS.Name = "FournisseurRS"
        Me.FournisseurRS.ReadOnly = True
        Me.FournisseurRS.Width = 84
        '
        'montantCommission
        '
        Me.montantCommission.DataPropertyName = "montantCommission"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.Format = "C2"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.montantCommission.DefaultCellStyle = DataGridViewCellStyle3
        Me.montantCommission.HeaderText = "Mt Comm"
        Me.montantCommission.Name = "montantCommission"
        Me.montantCommission.ReadOnly = True
        Me.montantCommission.Width = 76
        '
        'FournisseurCodeDataGridViewTextBoxColumn
        '
        Me.FournisseurCodeDataGridViewTextBoxColumn.DataPropertyName = "FournisseurCode"
        Me.FournisseurCodeDataGridViewTextBoxColumn.HeaderText = "FournisseurCode"
        Me.FournisseurCodeDataGridViewTextBoxColumn.Name = "FournisseurCodeDataGridViewTextBoxColumn"
        Me.FournisseurCodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.FournisseurCodeDataGridViewTextBoxColumn.Width = 111
        '
        'EtatLibelleDataGridViewTextBoxColumn
        '
        Me.EtatLibelleDataGridViewTextBoxColumn.DataPropertyName = "EtatLibelle"
        Me.EtatLibelleDataGridViewTextBoxColumn.HeaderText = "EtatLibelle"
        Me.EtatLibelleDataGridViewTextBoxColumn.Name = "EtatLibelleDataGridViewTextBoxColumn"
        Me.EtatLibelleDataGridViewTextBoxColumn.ReadOnly = True
        Me.EtatLibelleDataGridViewTextBoxColumn.Width = 81
        '
        'TiersCodeDataGridViewTextBoxColumn
        '
        Me.TiersCodeDataGridViewTextBoxColumn.DataPropertyName = "TiersCode"
        Me.TiersCodeDataGridViewTextBoxColumn.HeaderText = "TiersCode"
        Me.TiersCodeDataGridViewTextBoxColumn.Name = "TiersCodeDataGridViewTextBoxColumn"
        Me.TiersCodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.TiersCodeDataGridViewTextBoxColumn.Width = 80
        '
        'grpDateScmd
        '
        Me.grpDateScmd.Controls.Add(Me.tbCodeFournisseur)
        Me.grpDateScmd.Controls.Add(Me.Label1)
        Me.grpDateScmd.Controls.Add(Me.cbRecherche)
        Me.grpDateScmd.Controls.Add(Me.dtdateFin)
        Me.grpDateScmd.Controls.Add(Me.lblDateFin)
        Me.grpDateScmd.Controls.Add(Me.dtDatedeb)
        Me.grpDateScmd.Controls.Add(Me.lblDateDebut)
        Me.grpDateScmd.Location = New System.Drawing.Point(27, 0)
        Me.grpDateScmd.Name = "grpDateScmd"
        Me.grpDateScmd.Size = New System.Drawing.Size(389, 86)
        Me.grpDateScmd.TabIndex = 103
        Me.grpDateScmd.TabStop = False
        '
        'tbCodeFournisseur
        '
        Me.tbCodeFournisseur.Location = New System.Drawing.Point(199, 60)
        Me.tbCodeFournisseur.Name = "tbCodeFournisseur"
        Me.tbCodeFournisseur.Size = New System.Drawing.Size(88, 20)
        Me.tbCodeFournisseur.TabIndex = 103
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(7, 65)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(186, 15)
        Me.Label1.TabIndex = 106
        Me.Label1.Text = "Fournisseur :"
        '
        'dtdateFin
        '
        Me.dtdateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdateFin.Location = New System.Drawing.Point(199, 36)
        Me.dtdateFin.Name = "dtdateFin"
        Me.dtdateFin.Size = New System.Drawing.Size(88, 20)
        Me.dtdateFin.TabIndex = 102
        '
        'lblDateFin
        '
        Me.lblDateFin.Location = New System.Drawing.Point(7, 41)
        Me.lblDateFin.Name = "lblDateFin"
        Me.lblDateFin.Size = New System.Drawing.Size(186, 15)
        Me.lblDateFin.TabIndex = 105
        Me.lblDateFin.Text = "Date de fin :"
        '
        'dtDatedeb
        '
        Me.dtDatedeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDatedeb.Location = New System.Drawing.Point(199, 12)
        Me.dtDatedeb.Name = "dtDatedeb"
        Me.dtDatedeb.Size = New System.Drawing.Size(88, 20)
        Me.dtDatedeb.TabIndex = 101
        '
        'lblDateDebut
        '
        Me.lblDateDebut.Location = New System.Drawing.Point(7, 17)
        Me.lblDateDebut.Name = "lblDateDebut"
        Me.lblDateDebut.Size = New System.Drawing.Size(186, 15)
        Me.lblDateDebut.TabIndex = 104
        Me.lblDateDebut.Text = "Date de début :"
        '
        'rbDateSCMD
        '
        Me.rbDateSCMD.AutoSize = True
        Me.rbDateSCMD.Checked = True
        Me.rbDateSCMD.Location = New System.Drawing.Point(8, 36)
        Me.rbDateSCMD.Name = "rbDateSCMD"
        Me.rbDateSCMD.Size = New System.Drawing.Size(14, 13)
        Me.rbDateSCMD.TabIndex = 104
        Me.rbDateSCMD.TabStop = True
        Me.rbDateSCMD.UseVisualStyleBackColor = True
        '
        'rbCodeScmd
        '
        Me.rbCodeScmd.AutoSize = True
        Me.rbCodeScmd.Location = New System.Drawing.Point(10, 91)
        Me.rbCodeScmd.Name = "rbCodeScmd"
        Me.rbCodeScmd.Size = New System.Drawing.Size(53, 17)
        Me.rbCodeScmd.TabIndex = 105
        Me.rbCodeScmd.Text = "Code "
        Me.rbCodeScmd.UseVisualStyleBackColor = True
        '
        'rbIDScmd
        '
        Me.rbIDScmd.AutoSize = True
        Me.rbIDScmd.Location = New System.Drawing.Point(198, 91)
        Me.rbIDScmd.Name = "rbIDScmd"
        Me.rbIDScmd.Size = New System.Drawing.Size(36, 17)
        Me.rbIDScmd.TabIndex = 106
        Me.rbIDScmd.Text = "ID"
        Me.rbIDScmd.UseVisualStyleBackColor = True
        '
        'tbCodeScmd
        '
        Me.tbCodeScmd.Location = New System.Drawing.Point(69, 90)
        Me.tbCodeScmd.Name = "tbCodeScmd"
        Me.tbCodeScmd.Size = New System.Drawing.Size(123, 20)
        Me.tbCodeScmd.TabIndex = 107
        '
        'tbIDScmd
        '
        Me.tbIDScmd.Location = New System.Drawing.Point(249, 90)
        Me.tbIDScmd.Name = "tbIDScmd"
        Me.tbIDScmd.Size = New System.Drawing.Size(167, 20)
        Me.tbIDScmd.TabIndex = 108
        '
        'frmGestionSCMD
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(997, 622)
        Me.Controls.Add(Me.tbIDScmd)
        Me.Controls.Add(Me.tbCodeScmd)
        Me.Controls.Add(Me.rbIDScmd)
        Me.Controls.Add(Me.rbCodeScmd)
        Me.Controls.Add(Me.rbDateSCMD)
        Me.Controls.Add(Me.grpDateScmd)
        Me.Controls.Add(Me.m_dgvScmd)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.grpDetailSousCommande)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmGestionSCMD"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Gestion des Sous-commandes"
        Me.grpDetailSousCommande.ResumeLayout(False)
        Me.grpDetailSousCommande.PerformLayout()
        CType(Me.m_bsrcSousCommandes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.m_dgvScmd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDateScmd.ResumeLayout(False)
        Me.grpDateScmd.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Méthodes Vinicom"
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarSaveEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False

    End Sub
    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenetre n'est pas libre")
        Dim bReturn As Boolean
        bReturn = setElementCourant2(New SousCommande(New CommandeClient(New Client("", "")), New Fournisseur))

        Return bReturn
    End Function
    Protected Shadows Function getElementCourant() As SousCommande
        Return CType(m_ElementCourant, SousCommande)
    End Function

    Protected Overrides Function frmSave() As Boolean
        '===============================================================================================
        ' Function : frmSave
        ' Description : Sauvegarde de la collection des souscommande
        '===============================================================================================

        Dim objSCMD As SousCommande

        For Each objSCMD In m_colSousCommandes
            Try
                objSCMD.Save()
            Catch ex As Exception
                DisplayError("frmGestSCMD.frmSave()", ex.ToString)
            End Try

        Next objSCMD
        setfrmNotUpdated()
    End Function

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Appliquer déclenche L'evenement Updated
        AddHandler cbAppliquer.Click, AddressOf ControlUpdated
    End Sub
    Protected Overrides Sub setfrmUpdated()
        setfrmNotUpdated()
    End Sub
    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        Dim objControl As Control
        MyBase.EnableControls(bEnabled)
        dtDatedeb.Enabled = True
        dtdateFin.Enabled = True
        tbCodeFournisseur.Enabled = True
        cbRecherche.Enabled = True
        cbAfficher.Enabled = True
        m_dgvScmd.Enabled = True
        For Each objControl In grpDetailSousCommande.Controls
            objControl.Enabled = True
        Next objControl
        For Each objControl In grpDateScmd.Controls
            objControl.Enabled = True
        Next objControl
        rbCodeScmd.Enabled = True
        rbDateSCMD.Enabled = True
        rbIDScmd.Enabled = True
        '        tbComfournisseur.Enabled = False
    End Sub



#End Region

#Region "Methodes privées"
    Private Sub initFenetre()
        dtDatedeb.Value = DateAdd(DateInterval.Month, -1, Now())
        dtdateFin.Value = Now()
        m_colSousCommandes = New Collection
        If setElementCourant2(Nothing) Then
            tbTxCommision.Text = Param.getConstante("CST_TX_COMMISSION")
            '        lbSousCommandes.HorizontalExtent = 1
            '        lbSousCommandes.HorizontalScrollbar = True
            If currentuser.role = vncEnums.userRole.ADMIN Or currentuser.role = vncEnums.userRole.COMPTABILITE Then
                tbComfournisseur.Visible = True
            Else
                tbComfournisseur.Visible = False
            End If
            Me.Text = getResume()
            setDateLabels()
            m_dgvScmd.Columns(0).Width = My.Settings.frmGestSCMD_COL1_WIDTH
            m_dgvScmd.Columns(1).Width = My.Settings.frmGestSCMD_COL2_WIDTH
            m_dgvScmd.Columns(2).Width = My.Settings.frmGestSCMD_COL3_WIDTH
            m_dgvScmd.Columns(3).Width = My.Settings.frmGestSCMD_COL4_WIDTH
            m_dgvScmd.Columns(4).Width = My.Settings.frmGestSCMD_COL5_WIDTH
            m_dgvScmd.Columns(5).Width = My.Settings.frmGestSCMD_COL6_WIDTH
            m_dgvScmd.Columns(6).Width = My.Settings.frmGestSCMD_COL7_WIDTH
        End If
        InitCriteres()
    End Sub
    Protected Sub InitCriteres()
        grpDateScmd.Enabled = rbDateSCMD.Checked
        tbCodeScmd.Enabled = rbCodeScmd.Checked
        tbIDScmd.Enabled = rbIDScmd.Checked
    End Sub
    Protected Overridable Sub setDateLabels()
        lblDateDebut.Text = "Début date Commande"
        lblDateFin.Text = "Fin date Commande"
    End Sub
    Protected Overridable Function appliqueModifications() As Boolean
        Dim bReturn As Boolean
        majElement()
        bReturn = getElementCourant.Save()
        Return bReturn
    End Function
    Protected Overridable Sub afficheListeSousCommande()

        debAffiche()
        Debug.Assert(Not m_colSousCommandes Is Nothing)
        Dim obj As SousCommande

        m_bsrcSousCommandes.RaiseListChangedEvents = False
        m_bsrcSousCommandes.Clear()
        For Each obj In m_colSousCommandes
            obj.load()
            m_bsrcSousCommandes.Add(obj)
        Next obj
        m_bsrcSousCommandes.RaiseListChangedEvents = True
        m_bsrcSousCommandes.ResetBindings(False)
        m_bsrcSousCommandes.Position = -1
        finAffiche()
        m_bsrcSousCommandes.Position = 0
    End Sub 'AfficheListeSousCommande
    Protected Overridable Function setListeSousCommandes() As Boolean
        Dim ddeb As Date
        Dim dfin As Date
        Dim codeFourn As String
        Dim col As New Collection
        Dim bReturn As Boolean
        Dim oSCMD As SousCommande
        Dim nId As Long
        Dim strCode As String
        Try
            If Not getElementCourant() Is Nothing Then
                If getElementCourant().bUpdated Then
                    If MsgBox("La sous commande courante n'a pas été sauvegardée, voulez-vous conservez vos modifications", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        appliqueModifications()
                    End If
                End If
            End If
            'Chargement de la liste de sous commandes en fonction des critères
            If rbDateSCMD.Checked Then
                ddeb = dtDatedeb.Value.ToShortDateString
                dfin = dtdateFin.Value.ToShortDateString
                codeFourn = tbCodeFournisseur.Text
                col = SousCommande.getListe(ddeb, dfin, codeFourn)
            End If
            If rbIDScmd.Checked Then
                nId = Integer.Parse(Me.tbIDScmd.Text)
                oSCMD = SousCommande.createandload(nId)
                If oSCMD.id = nId Then
                    col.Add(oSCMD)
                End If
            End If
            If rbCodeScmd.Checked Then
                strCode = Me.tbCodeScmd.Text
                oSCMD = SousCommande.createandload(strCode)
                If oSCMD.code.Equals(strCode) Then
                    col.Add(oSCMD)
                End If
            End If
            m_colSousCommandes = col
            bReturn = True
        Catch ex As Exception
            bReturn = False
            Debug.Assert(bReturn, ex.ToString)
        End Try

        Return bReturn
    End Function

    Protected Overridable Function selectionneSousCommande() As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Try
            If Not getElementCourant() Is Nothing Then
                If getElementCourant().bUpdated Then
                    If MsgBox("L'élement courant a été modifié, voulez-vous le sauvegarder ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        appliqueModifications()

                    End If
                End If
            End If
            If Not (m_bsrcSousCommandes.Current Is Nothing) Then
                bReturn = setElementCourant2(m_bsrcSousCommandes.Current)
                If bReturn Then
                    If getElementCourant.bResume Then
                        getElementCourant().load()
                    End If
                End If
            End If

        Catch ex As Exception
            Debug.Assert(False, ex.ToString)
            bReturn = False
        End Try

        Return bReturn

    End Function 'selectionneSousCommande

    Protected Function afficheSousCommande() As Boolean
        Dim objcommandeClient As CommandeClient
        Dim objFact As FactCom

        Debug.Assert(Not getElementCourant() Is Nothing, "La Sous Commande doit être renseignée")

        'Affichage de la commande Client
        debAffiche()
        objcommandeClient = New CommandeClient(New Client("", ""))
        objcommandeClient.load(getElementCourant().idCommandeClient)
        If objcommandeClient.id <> 0 Then
            liCommande.Tag = objcommandeClient.id
            liCommande.Text = objcommandeClient.shortResume
        End If

        'Affichage du Fournisseur
        liFournisseur.Tag = getElementCourant().oFournisseur.id
        liFournisseur.Text = getElementCourant().oFournisseur.shortResume

        'Affichage du client
        liClient.Tag = getElementCourant().oTiers.id
        liClient.Text = getElementCourant().oTiers.shortResume


        'Commission
        tbComfournisseur.Text = getElementCourant().oFournisseur.CommFacturation.comment

        'Facturee
        If getElementCourant().idFactCom <> 0 And getElementCourant().idFactCom <> -1 Then
            objFact = FactCom.createandload(getElementCourant().idFactCom)
            liFactCom.Tag = objFact.id
            liFactCom.Text = objFact.shortResume
        End If

        'AffichageEtat
        afficheEtatSCMD()
        finAffiche()

    End Function 'AfficheSousCommande
    Public Overrides Function majElement() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing, "Pas de sous commande courante")

        MAJEtatSCMD()
    End Function 'majElement
    Private Sub afficheEtatSCMD()
        debAffiche()
        liFactCom.Visible = False
        ckTransmiseFax.Checked = False
        ckExporteeInternet.Checked = False
        ckRapprochee.Checked = False
        ckImporteeInternet.Checked = False
        ckFacturee.Checked = False
        Select Case getElementCourant().etat.codeEtat
            Case vncEnums.vncEtatCommande.vncSCMDGeneree
            Case vncEnums.vncEtatCommande.vncSCMDtransmiseFax
                ckTransmiseFax.Checked = True
            Case vncEnums.vncEtatCommande.vncSCMDExporteeInt
                ckExporteeInternet.Checked = True
            Case vncEnums.vncEtatCommande.vncSCMDRapprochee
                ckTransmiseFax.Checked = Not getElementCourant().bExportInternet
                ckExporteeInternet.Checked = getElementCourant().bExportInternet
                ckRapprochee.Checked = True
            Case vncEnums.vncEtatCommande.vncSCMDRapprocheeInt
                ckTransmiseFax.Checked = Not getElementCourant().bExportInternet
                ckExporteeInternet.Checked = getElementCourant().bExportInternet
                ckImporteeInternet.Checked = True
            Case vncEnums.vncEtatCommande.vncSCMDFacturee
                ckTransmiseFax.Checked = Not getElementCourant().bExportInternet
                ckExporteeInternet.Checked = getElementCourant().bExportInternet
                ckRapprochee.Checked = Not getElementCourant().bExportInternet
                ckImporteeInternet.Checked = getElementCourant().bExportInternet
                ckFacturee.Checked = True
                liFactCom.Visible = True
            Case Else
                ckTransmiseFax.Checked = False
                ckExporteeInternet.Checked = False
                ckRapprochee.Checked = False
                ckImporteeInternet.Checked = False
                ckFacturee.Checked = False
        End Select
        finAffiche()
    End Sub ' afficheEtatSCMD
    Protected Sub MAJEtatSCMD()
        Debug.Assert(Not getElementCourant() Is Nothing, "Pas de souscommande Courante")
        ' Par defaut : Etat = Genérée
        getElementCourant().etat = New EtatSSCommandeGeneree(vncGenererSupprimer.vncRien, vncGenererSupprimer.vncRien)
        If ckTransmiseFax.Checked Then
            getElementCourant().etat = New EtatSSCommandeTransmise(vncGenererSupprimer.vncRien, vncGenererSupprimer.vncRien)
        End If
        If ckExporteeInternet.Checked Then
            getElementCourant().etat = New EtatSSCommandeExporteeInt(vncGenererSupprimer.vncRien, vncGenererSupprimer.vncRien)
        End If
        If ckRapprochee.Checked Then
            getElementCourant().etat = New EtatSSCommandeRapprochee(vncGenererSupprimer.vncRien, vncGenererSupprimer.vncRien)
        End If
        If ckImporteeInternet.Checked Then
            getElementCourant().etat = New EtatSSCommandeRapprocheeInt(vncGenererSupprimer.vncRien, vncGenererSupprimer.vncRien)
        End If
        If ckFacturee.Checked Then
            getElementCourant().etat = New EtatSSCommandeFacturee(vncGenererSupprimer.vncRien, vncGenererSupprimer.vncRien)
        End If

    End Sub ' afficheEtatSCMD
    Private Sub duppliqueinfosCMD()
        debAffiche()
        getElementCourant().DuppliqueInfosCmd()
        m_bsrcSousCommandes.ResetCurrentItem()
        finAffiche()
    End Sub 'duppliqueInfosCmd
    Private Function elementCourantSansModif() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = False
            If Not getElementCourant() Is Nothing Then
                bReturn = getElementCourant().bUpdated
            End If

        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Sub CalculMontantCom()

        getElementCourant().calcCommisionstandard(CalculCommScmd.CALCUL_COMMISSION_BASE_COMM)
        m_bsrcSousCommandes.ResetCurrentItem()
    End Sub
    Private Sub rechercheFournisseur()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.FOURNISSEUR, tbCodeFournisseur)

        If Not objTiers Is Nothing Then
            tbCodeFournisseur.Text = objTiers.code
        End If
    End Sub 'RechercheFournisseur

#End Region
#Region "Gestion des Evenements"
    Private Sub rbDateSCMD_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbDateSCMD.CheckedChanged
        InitCriteres()
    End Sub

    Private Sub rbCodeScmd_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbCodeScmd.CheckedChanged
        InitCriteres()
    End Sub

    Private Sub rbIDScmd_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbIDScmd.CheckedChanged
        InitCriteres()
    End Sub

    Private Sub frmGestionSCMD_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        My.Settings.frmGestSCMD_COL1_WIDTH = m_dgvScmd.Columns(0).Width
        My.Settings.frmGestSCMD_COL2_WIDTH = m_dgvScmd.Columns(1).Width
        My.Settings.frmGestSCMD_COL3_WIDTH = m_dgvScmd.Columns(2).Width
        My.Settings.frmGestSCMD_COL4_WIDTH = m_dgvScmd.Columns(3).Width
        My.Settings.frmGestSCMD_COL5_WIDTH = m_dgvScmd.Columns(4).Width
        My.Settings.frmGestSCMD_COL6_WIDTH = m_dgvScmd.Columns(5).Width
        My.Settings.frmGestSCMD_COL7_WIDTH = m_dgvScmd.Columns(6).Width
        My.Settings.Save()
    End Sub
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (Not DesignMode) Then
            initFenetre()
        End If
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        setcursorWait()
        If setListeSousCommandes() Then
            If setElementCourant2(Nothing) Then
                afficheListeSousCommande()
            End If
        End If
        grpDetailSousCommande.Enabled = False
        restoreCursor()
    End Sub

    Private Sub cbSelectionner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If selectionneSousCommande() Then
            afficheSousCommande()
            grpDetailSousCommande.Enabled = True
        End If
    End Sub

    Private Sub liFournisseur_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liFournisseur.LinkClicked
        afficheFenetreFournisseur(liFournisseur.Tag)
    End Sub

    Private Sub liClient_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liClient.LinkClicked
        afficheFenetreClient(liClient.Tag)
    End Sub

    Private Sub liCommande_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liCommande.LinkClicked
        afficheFenetreCommandeClient(liCommande.Tag)
    End Sub

    Private Sub cbAppliquer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAppliquer.Click
        appliqueModifications()
        grpDetailSousCommande.Enabled = False
    End Sub
#End Region



    Private Sub ckRapprochee_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckRapprochee.CheckedChanged
        If Not bAffichageEnCours() Then
            If ckRapprochee.Checked Then
                duppliqueinfosCMD()
                tbFACTreference.Focus()
            End If
        End If

    End Sub


    Private Sub cbCalc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCalc.Click
        CalculMontantCom()
    End Sub


    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        rechercheFournisseur()
    End Sub

    Private Sub flxSousCommandes_DblClick(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub tbTauxComm_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CalculMontantCom()
    End Sub
    Private Sub tbBaseComm_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CalculMontantCom()
    End Sub

    Private Sub cbTransmettreFax_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTransmettreFax.Click
        transmettreFax()
    End Sub

    Private Function transmettreFax() As Boolean
        Dim bReturn As Boolean = False
        Dim odlg As dlgVisuBonFournisseur
        Try
            sauvegardeElementCourant()
            odlg = New dlgVisuBonFournisseur
            odlg.setElement(getElementCourant())
            If odlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
                getElementCourant.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFaxer)
            End If
            afficheEtatSCMD()
            bReturn = True
        Catch ex As Exception
            Debug.Assert(False, ex.Message)
            bReturn = False

        End Try
        Return bReturn
    End Function
    Private Sub exporter()

        Dim objSCMD As SousCommande
        Dim bExportOK As Boolean = True
        Dim bReturn As Boolean
        Dim strFile As String
        Dim strSCMD_CSV As String
        Dim nFile As Integer
        Dim strPDFFileName As String
        Dim strFolder As String
        Dim oFTPvinicom As clsFTPVinicom
        Dim nSousCommandesPreparees As Integer
        Dim nSousCommandesExportees As Integer

        'Suppression - creation du répertoire temporaire
        Try
            nSousCommandesPreparees = 0
            setcursorWait()
            DisplayStatus("Création du répertoire Temporaire")
            strFolder = IMPORT_DIRECTORY & "/" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second
            If System.IO.Directory.Exists(strFolder) Then
                System.IO.Directory.Delete(strFolder, True)
            End If
            System.IO.Directory.CreateDirectory(strFolder)

            strFile = strFolder & "/" & EXPORTFTP_FILENAME

            'Génération des fichiers dans le répertoire temporaire
            nFile = FreeFile()
            FileOpen(nFile, strFile, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
            objSCMD = getElementCourant()
            DisplayStatus("Chargement de " & objSCMD.code)
            objSCMD.load()
            objSCMD.loadcolLignes()
            DisplayStatus("Chargement de " & objSCMD.code & " CSV")
            strSCMD_CSV = objSCMD.toCSV()
            Print(nFile, strSCMD_CSV)
            DisplayStatus("Chargement de " & objSCMD.code & " PDF")
            strPDFFileName = strFolder & "/" & objSCMD.code & ".PDF"
            objSCMD.genererPDF(PATHTOREPORTS, strPDFFileName)
            nSousCommandesPreparees = nSousCommandesPreparees + 1
            FileClose(nFile)

            bReturn = True
            DisplayStatus("Transferts des fichiers ")
            'Exporter les fichiers générés
            oFTPvinicom = New clsFTPVinicom 'Création avec les paramètres par defaut
            'oFTPvinicom.connect()
            nSousCommandesExportees = oFTPvinicom.uploadFromDir(strFolder)
            DisplayStatus("Fin de transfert des fichiers ")
            DisplayStatus("Nombre de commandes Préparées : " & nSousCommandesPreparees)
            DisplayStatus("Nombre de fichiers Exportés : " & nSousCommandesExportees)

            objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
            objSCMD.bExportInternet = True
            afficheEtatSCMD()

        Catch ex As Exception
            MsgBox("Erreur" + ex.Message)

        End Try
        Me.Cursor = Cursors.Default

    End Sub 'exporter

    Private Sub cbExportInternet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExportInternet.Click
        exporter()
    End Sub

    Private Sub m_dgvScmd_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub m_dgvScmd_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles m_dgvScmd.CellDoubleClick
        If selectionneSousCommande() Then
            afficheSousCommande()
            grpDetailSousCommande.Enabled = True
        End If
    End Sub

    Private Sub ckImporteeInternet_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckImporteeInternet.CheckedChanged

    End Sub

    Private Sub m_bsrcSousCommandes_CurrentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles m_bsrcSousCommandes.CurrentChanged
        If Not bAffichageEnCours() Then
            If selectionneSousCommande() Then
                afficheSousCommande()
                grpDetailSousCommande.Enabled = True
            End If
        End If
    End Sub

End Class