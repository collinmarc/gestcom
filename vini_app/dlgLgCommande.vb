Imports vini_DB
Public Class dlgLgCommande
    Inherits System.Windows.Forms.Form
    Private m_elementCourant As LgCommande
    Private m_TiersCourant As Tiers
    Private m_bModif As Boolean
    Private m_bAffichageEncours As Boolean
    Private m_typeProduit As vncTypeProduit
    'Modification ou Création de lignes

#Region "Code généré par le Concepteur Windows Form "
    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
        m_bModif = False
    End Sub
    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer
    Public WithEvents cbProduit As System.Windows.Forms.Button
    Public WithEvents tbCodeProduit As System.Windows.Forms.TextBox
    Public WithEvents tbNumLigne As textBoxNumeric
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents liProduit As System.Windows.Forms.LinkLabel
    Friend WithEvents grpQtePrix As System.Windows.Forms.GroupBox
    Public WithEvents tbQteFact As vini_app.textBoxNumeric
    Public WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents cbCalculTotal As System.Windows.Forms.Button
    Public WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents tbTotalTTC As textBoxCurrency
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Public WithEvents tbQteLiv As vini_app.textBoxNumeric
    Public WithEvents tbPrixUHT As textBoxCurrency
    Public WithEvents tbTotalHT As textBoxCurrency
    Public WithEvents ckGratuit As System.Windows.Forms.CheckBox
    Public WithEvents tbQteCom As vini_app.textBoxNumeric
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents cbAnnuler As System.Windows.Forms.Button
    Public WithEvents cbValider As System.Windows.Forms.Button
    Friend WithEvents laStockTheo As System.Windows.Forms.Label
    Friend WithEvents laStockCommande As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents laStockReel As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cbProduit = New System.Windows.Forms.Button
        Me.tbCodeProduit = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.liProduit = New System.Windows.Forms.LinkLabel
        Me.Label7 = New System.Windows.Forms.Label
        Me.grpQtePrix = New System.Windows.Forms.GroupBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.cbCalculTotal = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.ckGratuit = New System.Windows.Forms.CheckBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.cbAnnuler = New System.Windows.Forms.Button
        Me.cbValider = New System.Windows.Forms.Button
        Me.laStockTheo = New System.Windows.Forms.Label
        Me.laStockCommande = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.laStockReel = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.tbNumLigne = New vini_app.textBoxNumeric
        Me.tbQteFact = New vini_app.textBoxNumeric
        Me.tbTotalTTC = New vini_app.textBoxCurrency
        Me.tbQteLiv = New vini_app.textBoxNumeric
        Me.tbPrixUHT = New vini_app.textBoxCurrency
        Me.tbTotalHT = New vini_app.textBoxCurrency
        Me.tbQteCom = New vini_app.textBoxNumeric
        Me.grpQtePrix.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbProduit
        '
        Me.cbProduit.BackColor = System.Drawing.SystemColors.Control
        Me.cbProduit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbProduit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbProduit.Location = New System.Drawing.Point(232, 32)
        Me.cbProduit.Name = "cbProduit"
        Me.cbProduit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbProduit.Size = New System.Drawing.Size(25, 17)
        Me.cbProduit.TabIndex = 1
        Me.cbProduit.Text = "..."
        Me.cbProduit.UseVisualStyleBackColor = False
        '
        'tbCodeProduit
        '
        Me.tbCodeProduit.BackColor = System.Drawing.SystemColors.Window
        Me.tbCodeProduit.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCodeProduit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCodeProduit.Location = New System.Drawing.Point(96, 32)
        Me.tbCodeProduit.MaxLength = 0
        Me.tbCodeProduit.Name = "tbCodeProduit"
        Me.tbCodeProduit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCodeProduit.Size = New System.Drawing.Size(129, 19)
        Me.tbCodeProduit.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Code &produit"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "&Numéro Ligne"
        '
        'liProduit
        '
        Me.liProduit.Location = New System.Drawing.Point(264, 32)
        Me.liProduit.Name = "liProduit"
        Me.liProduit.Size = New System.Drawing.Size(496, 24)
        Me.liProduit.TabIndex = 28
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(368, 96)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 23)
        Me.Label7.TabIndex = 29
        Me.Label7.Text = "Qte en commande"
        '
        'grpQtePrix
        '
        Me.grpQtePrix.Controls.Add(Me.tbQteFact)
        Me.grpQtePrix.Controls.Add(Me.Label18)
        Me.grpQtePrix.Controls.Add(Me.cbCalculTotal)
        Me.grpQtePrix.Controls.Add(Me.Label17)
        Me.grpQtePrix.Controls.Add(Me.tbTotalTTC)
        Me.grpQtePrix.Controls.Add(Me.Label16)
        Me.grpQtePrix.Controls.Add(Me.tbQteLiv)
        Me.grpQtePrix.Controls.Add(Me.tbPrixUHT)
        Me.grpQtePrix.Controls.Add(Me.tbTotalHT)
        Me.grpQtePrix.Controls.Add(Me.ckGratuit)
        Me.grpQtePrix.Controls.Add(Me.tbQteCom)
        Me.grpQtePrix.Controls.Add(Me.Label12)
        Me.grpQtePrix.Controls.Add(Me.Label11)
        Me.grpQtePrix.Controls.Add(Me.Label10)
        Me.grpQtePrix.Controls.Add(Me.Label9)
        Me.grpQtePrix.Controls.Add(Me.Label6)
        Me.grpQtePrix.Controls.Add(Me.Label5)
        Me.grpQtePrix.Location = New System.Drawing.Point(8, 80)
        Me.grpQtePrix.Name = "grpQtePrix"
        Me.grpQtePrix.Size = New System.Drawing.Size(344, 192)
        Me.grpQtePrix.TabIndex = 2
        Me.grpQtePrix.TabStop = False
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(24, 64)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(81, 17)
        Me.Label18.TabIndex = 54
        Me.Label18.Text = "&Qte Facturée"
        '
        'cbCalculTotal
        '
        Me.cbCalculTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbCalculTotal.Location = New System.Drawing.Point(280, 112)
        Me.cbCalculTotal.Name = "cbCalculTotal"
        Me.cbCalculTotal.Size = New System.Drawing.Size(40, 24)
        Me.cbCalculTotal.TabIndex = 4
        Me.cbCalculTotal.Text = "&Calc"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(232, 168)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(41, 17)
        Me.Label17.TabIndex = 52
        Me.Label17.Text = "Euros"
        '
        'Label16
        '
        Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label16.Location = New System.Drawing.Point(24, 160)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(80, 23)
        Me.Label16.TabIndex = 51
        Me.Label16.Text = "Total &TTC"
        '
        'ckGratuit
        '
        Me.ckGratuit.BackColor = System.Drawing.SystemColors.Control
        Me.ckGratuit.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckGratuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.ckGratuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ckGratuit.Location = New System.Drawing.Point(24, 88)
        Me.ckGratuit.Name = "ckGratuit"
        Me.ckGratuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ckGratuit.Size = New System.Drawing.Size(104, 17)
        Me.ckGratuit.TabIndex = 40
        Me.ckGratuit.Text = "&Gratuit"
        Me.ckGratuit.UseVisualStyleBackColor = False
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(24, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(81, 17)
        Me.Label12.TabIndex = 50
        Me.Label12.Text = "&Qte Livrée"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(232, 120)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(41, 17)
        Me.Label11.TabIndex = 49
        Me.Label11.Text = "Euros"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(24, 112)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(81, 17)
        Me.Label10.TabIndex = 48
        Me.Label10.Text = "Prix&U HT"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(232, 144)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(41, 17)
        Me.Label9.TabIndex = 47
        Me.Label9.Text = "Euros"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(24, 136)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(73, 17)
        Me.Label6.TabIndex = 46
        Me.Label6.Text = "Total  &HT"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(24, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(88, 17)
        Me.Label5.TabIndex = 45
        Me.Label5.Text = "&Qte commandée"
        '
        'cbAnnuler
        '
        Me.cbAnnuler.BackColor = System.Drawing.SystemColors.Control
        Me.cbAnnuler.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAnnuler.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAnnuler.Location = New System.Drawing.Point(672, 232)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAnnuler.Size = New System.Drawing.Size(81, 25)
        Me.cbAnnuler.TabIndex = 4
        Me.cbAnnuler.Text = "&Annuler"
        Me.cbAnnuler.UseVisualStyleBackColor = False
        '
        'cbValider
        '
        Me.cbValider.BackColor = System.Drawing.SystemColors.Control
        Me.cbValider.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbValider.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbValider.Location = New System.Drawing.Point(576, 232)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbValider.Size = New System.Drawing.Size(81, 25)
        Me.cbValider.TabIndex = 3
        Me.cbValider.Text = "&Valider"
        Me.cbValider.UseVisualStyleBackColor = False
        '
        'laStockTheo
        '
        Me.laStockTheo.ForeColor = System.Drawing.Color.Red
        Me.laStockTheo.Location = New System.Drawing.Point(488, 112)
        Me.laStockTheo.Name = "laStockTheo"
        Me.laStockTheo.Size = New System.Drawing.Size(100, 16)
        Me.laStockTheo.TabIndex = 40
        Me.laStockTheo.Text = "Stock_Theorique"
        '
        'laStockCommande
        '
        Me.laStockCommande.Location = New System.Drawing.Point(488, 96)
        Me.laStockCommande.Name = "laStockCommande"
        Me.laStockCommande.Size = New System.Drawing.Size(168, 16)
        Me.laStockCommande.TabIndex = 39
        Me.laStockCommande.Text = "Qte_Commandee"
        '
        'Label13
        '
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(368, 112)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(100, 16)
        Me.Label13.TabIndex = 38
        Me.Label13.Text = "Stock Théorique"
        '
        'laStockReel
        '
        Me.laStockReel.ForeColor = System.Drawing.Color.Blue
        Me.laStockReel.Location = New System.Drawing.Point(488, 80)
        Me.laStockReel.Name = "laStockReel"
        Me.laStockReel.Size = New System.Drawing.Size(128, 16)
        Me.laStockReel.TabIndex = 37
        Me.laStockReel.Text = "Quantite_en_stock_reel"
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(368, 80)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 16)
        Me.Label8.TabIndex = 36
        Me.Label8.Text = "Stock Réel"
        '
        'tbNumLigne
        '
        Me.tbNumLigne.BackColor = System.Drawing.SystemColors.Window
        Me.tbNumLigne.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbNumLigne.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbNumLigne.Location = New System.Drawing.Point(96, 8)
        Me.tbNumLigne.MaxLength = 0
        Me.tbNumLigne.Name = "tbNumLigne"
        Me.tbNumLigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbNumLigne.Size = New System.Drawing.Size(49, 20)
        Me.tbNumLigne.TabIndex = 12
        Me.tbNumLigne.Text = "0"
        '
        'tbQteFact
        '
        Me.tbQteFact.AcceptsReturn = True
        Me.tbQteFact.BackColor = System.Drawing.SystemColors.Window
        Me.tbQteFact.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbQteFact.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbQteFact.Location = New System.Drawing.Point(112, 64)
        Me.tbQteFact.MaxLength = 0
        Me.tbQteFact.Name = "tbQteFact"
        Me.tbQteFact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbQteFact.Size = New System.Drawing.Size(112, 20)
        Me.tbQteFact.TabIndex = 2
        Me.tbQteFact.Text = "0"
        '
        'tbTotalTTC
        '
        Me.tbTotalTTC.AcceptsReturn = True
        Me.tbTotalTTC.Location = New System.Drawing.Point(112, 160)
        Me.tbTotalTTC.Name = "tbTotalTTC"
        Me.tbTotalTTC.Size = New System.Drawing.Size(112, 20)
        Me.tbTotalTTC.TabIndex = 6
        Me.tbTotalTTC.Text = "0"
        '
        'tbQteLiv
        '
        Me.tbQteLiv.BackColor = System.Drawing.SystemColors.Window
        Me.tbQteLiv.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbQteLiv.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbQteLiv.Location = New System.Drawing.Point(112, 40)
        Me.tbQteLiv.MaxLength = 0
        Me.tbQteLiv.Name = "tbQteLiv"
        Me.tbQteLiv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbQteLiv.Size = New System.Drawing.Size(112, 20)
        Me.tbQteLiv.TabIndex = 1
        Me.tbQteLiv.Text = "0"
        '
        'tbPrixUHT
        '
        Me.tbPrixUHT.BackColor = System.Drawing.SystemColors.Window
        Me.tbPrixUHT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbPrixUHT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbPrixUHT.Location = New System.Drawing.Point(112, 112)
        Me.tbPrixUHT.MaxLength = 0
        Me.tbPrixUHT.Name = "tbPrixUHT"
        Me.tbPrixUHT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbPrixUHT.Size = New System.Drawing.Size(113, 20)
        Me.tbPrixUHT.TabIndex = 3
        Me.tbPrixUHT.Text = "0"
        '
        'tbTotalHT
        '
        Me.tbTotalHT.BackColor = System.Drawing.SystemColors.Window
        Me.tbTotalHT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbTotalHT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbTotalHT.Location = New System.Drawing.Point(112, 136)
        Me.tbTotalHT.MaxLength = 0
        Me.tbTotalHT.Name = "tbTotalHT"
        Me.tbTotalHT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbTotalHT.Size = New System.Drawing.Size(113, 20)
        Me.tbTotalHT.TabIndex = 5
        Me.tbTotalHT.Text = "0"
        '
        'tbQteCom
        '
        Me.tbQteCom.BackColor = System.Drawing.SystemColors.Window
        Me.tbQteCom.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbQteCom.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbQteCom.Location = New System.Drawing.Point(112, 16)
        Me.tbQteCom.MaxLength = 0
        Me.tbQteCom.Name = "tbQteCom"
        Me.tbQteCom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbQteCom.Size = New System.Drawing.Size(112, 20)
        Me.tbQteCom.TabIndex = 0
        Me.tbQteCom.Text = "0"
        '
        'dlgLgCommande
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(762, 272)
        Me.Controls.Add(Me.laStockTheo)
        Me.Controls.Add(Me.laStockCommande)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.laStockReel)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.tbCodeProduit)
        Me.Controls.Add(Me.tbNumLigne)
        Me.Controls.Add(Me.grpQtePrix)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.liProduit)
        Me.Controls.Add(Me.cbProduit)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.cbValider)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.Green
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(184, 250)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgLgCommande"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Ligne de Commande [Num] [Nom Client]"
        Me.grpQtePrix.ResumeLayout(False)
        Me.grpQtePrix.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Accesseurs"
    Public Sub New()
        MyBase.New()
        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        m_elementCourant = New LgCommande(0)
        EnabledControlsQteetPrix(False)
        m_bModif = False
        m_typeProduit = vncEnums.vncTypeProduit.vncTous
    End Sub

    'Public Property bModif() As Boolean
    '    Get
    '        Return m_bModif
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        m_bModif = Value
    '    End Set
    'End Property

    Public Sub setElementCourant(ByRef p_objlg As LgCommande, ByVal p_objTiers As Tiers)
        Debug.Assert(Not p_objlg Is Nothing)
        Debug.Assert(Not p_objTiers Is Nothing)
        m_elementCourant = p_objlg
        m_TiersCourant = p_objTiers
        AfficheElement()
    End Sub
    Public Function getElementCourant() As LgCommande
        Return m_elementCourant
    End Function
    Public Sub setRechercheTypeProduit(ByVal pTypeProduit As vncTypeProduit)
        m_typeProduit = pTypeProduit
    End Sub
#End Region

#Region "Méthodes"
    'Affichagde de l'élement courant
    Private Sub AfficheElement()
        Debug.Assert(Not m_elementCourant Is Nothing)
        Debug.Assert(Not m_TiersCourant Is Nothing)
        m_bAffichageEncours = True
        tbNumLigne.Text = m_elementCourant.num.ToString()
        If Not m_elementCourant.oProduit Is Nothing Then
            AfficheProduit(m_elementCourant.oProduit)
        Else
            cbProduit.Visible = True
        End If
        tbQteCom.Text = m_elementCourant.qteCommande.ToString()
        tbQteLiv.Text = m_elementCourant.qteLiv.ToString()
        tbQteFact.Text = m_elementCourant.qteFact.ToString()
        tbPrixUHT.Text = m_elementCourant.prixU.ToString()
        tbTotalHT.Text = m_elementCourant.prixHT.ToString()
        tbTotalTTC.Text = m_elementCourant.prixTTC.ToString()


        ckGratuit.Checked = m_elementCourant.bGratuit

        If m_bModif Then
            tbNumLigne.Enabled = False
        Else
            tbNumLigne.Enabled = True
        End If
        m_bAffichageEncours = False
    End Sub
    'Affichage des caractéristiques du produit
    Private Sub AfficheProduit(ByVal p_oProduit As Produit)
        Debug.Assert(Not p_oProduit Is Nothing)
        Dim nStockReel As Decimal
        Dim nQteComm As Decimal
        Dim nStockTheo As Decimal
        Dim objProduit As Produit

        objProduit = Produit.createandload(p_oProduit.id)
        If Not objProduit Is Nothing Then

            tbCodeProduit.Text = objProduit.code
            liProduit.Text = objProduit.shortResume
            liProduit.Tag = objProduit.id

            nStockReel = objProduit.QteStock
            nQteComm = objProduit.qteCommande
            nStockTheo = nStockReel - nQteComm

            laStockReel.Text = CStr(nStockReel)
            laStockCommande.Text = CStr(nQteComm)
            laStockTheo.Text = nStockTheo.ToString()

            EnabledControlsQteetPrix(True)
        End If

    End Sub 'AfficheProduit
    'Récupération de la saisie dans l'objet courant
    Private Sub MAJElement(ByVal p_Element As LgCommande)
        Debug.Assert(Not p_Element Is Nothing)
        Dim bReturn As Boolean = True

        If m_bAffichageEncours Then
            Exit Sub
        End If
        p_Element.num = CInt(tbNumLigne.Text)
        p_Element.qteCommande = CDec(tbQteCom.Text)
        p_Element.qteLiv = CDec(tbQteLiv.Text)
        p_Element.qteFact = CDec(tbQteFact.Text)
        p_Element.bGratuit = ckGratuit.Checked
        p_Element.prixU = CDec(tbPrixUHT.Text)
        p_Element.prixHT = CDec(tbTotalHT.Text)
        p_Element.prixTTC = CDec(tbTotalTTC.Text)
        If ckGratuit.Checked Then
            p_Element.prixU = 0
            p_Element.prixHT = 0
            p_Element.prixTTC = 0
        End If

        If p_Element.oProduit Is Nothing Then
            p_Element.oProduit = Produit.createandload(CInt(liProduit.Tag))
        Else
            bReturn = p_Element.oProduit.load(CInt(liProduit.Tag))
        End If

        Debug.Assert(bReturn, "dlgLgcommande.MAJElement : " & LgCommande.getErreur)
    End Sub
    'Récupération des caractéristiques du produit
    'Le Paramètre est passé en byRef car il est mis à jour par cette procédure
    'Private Sub MAJProduit(ByRef p_oProduit As Produit)
    '    Debug.Assert(Not m_elementCourant Is Nothing)
    '    '        Debug.Assert(Not p_oProduit Is Nothing)
    '    Debug.Assert(IsNumeric(liProduit.Tag), "L'ID du produit n'est pas dans le tag")

    '    Dim bReturn As Boolean
    '    If p_oProduit Is Nothing Then
    '        p_oProduit = Produit.createandload(liProduit.Tag)
    '        bReturn = True
    '    Else
    '        bReturn = p_oProduit.load(liProduit.Tag)

    '    End If

    '    Debug.Assert(bReturn, "LoadProduit" & liProduit.Tag)

    'End Sub 'MAJ Produit

    '====================================================================================
    ' Fonction : rechercheProduit
    ' Description : REchrehce du produit à partir de son code
    '                si aucun produit n'est trouvé ou plus d'un une fenêtre est affichée 
    '                   avec la liste des produits correspondant
    '                   Les produits affichées sont ceux de la precommande du client
    '====================================================================================
    Private Sub rechercheProduit()
        Debug.Assert(Not m_elementCourant Is Nothing)
        Debug.Assert(Not m_TiersCourant Is Nothing)
        Dim colProduit As Collection
        Dim objProduit As Produit = Nothing
        Dim frm As frmRechercheDB
        Dim objClient As Client
        Dim objLgPrecom As lgPrecomm
        Dim qte As Decimal = 0
        Dim prixU As Decimal = 0

        If tbCodeProduit.Text <> "" Then
            colProduit = Produit.getListe(m_typeProduit, tbCodeProduit.Text)
        Else
            If m_TiersCourant.typeDonnee = vncTypeDonnee.FOURNISSEUR Then
                colProduit = Produit.getListe(m_typeProduit, , , , m_TiersCourant.id, )
            Else
                colProduit = Produit.getListe(m_typeProduit, , , , , m_TiersCourant.id)
            End If

        End If
        If colProduit.Count <> 1 Then
            'Création de la fenêtre de recherche
            frm = New frmRechercheDB
            frm.setTypeDonnees(vncEnums.vncTypeDonnee.PRODUIT)
            'Liste des produits de la PRECOMMANDE si Constantes positionnée
            If COMMANDECLIENT_LISTEPRODUIT_PRECOMMANDE Then
                If m_TiersCourant.typeDonnee = vncTypeDonnee.FOURNISSEUR Then
                    frm.setidFournisseur(m_TiersCourant.id)
                Else
                    frm.setidPrecommande(m_TiersCourant.id)
                End If
            End If
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
            If objProduit.bResume Then
                objProduit.load()
            End If
            m_elementCourant.oProduit = objProduit
            'Récupération des infos dans la précommande => Quantité Habituelle
            If m_TiersCourant.typeDonnee = vncEnums.vncTypeDonnee.CLIENT Then
                objClient = m_TiersCourant
                If objClient.LoadPreCommande() Then
                    If objClient.lgPrecomExists(objProduit.code) Then
                        objLgPrecom = objClient.getLgPrecom(objProduit.code)
                        qte = objLgPrecom.qteHab
                    End If
                End If
                ' Récupération du code tarif
                prixU = objProduit.Tarif(objClient.CodeTarif)
            End If
            m_elementCourant.qteCommande = qte
            m_elementCourant.prixU = prixU
        End If

        AfficheElement()
        tbQteCom.Focus()
    End Sub 'RechercheProduit

    Private Sub EnabledControlsQteetPrix(ByVal bEnabled As Boolean)
        grpQtePrix.Enabled = bEnabled
        cbValider.Enabled = bEnabled
    End Sub
    Private Sub setGratuit(ByVal bGratuit As Boolean)
        CalculPrix()
        tbPrixUHT.Enabled = Not bGratuit
        tbTotalHT.Enabled = Not bGratuit
        tbTotalTTC.Enabled = Not bGratuit
        cbCalculTotal.Enabled = Not bGratuit

    End Sub
    Private Sub CalculPrix()
        Debug.Assert(Not m_elementCourant Is Nothing)
        Dim oLgCom As LgCommande
        oLgCom = New LgCommande(0)

        MAJElement(oLgCom) 'Mise à jour de l'élement courant avant de calculer le prix
        oLgCom.calculPrixTotal()
        tbPrixUHT.Text = oLgCom.prixU
        tbTotalHT.Text = oLgCom.prixHT
        tbTotalTTC.Text = oLgCom.prixTTC
        cbValider.Focus()

    End Sub
    Private Sub AfficheFenetreProduit(ByVal tag As String)
        If Not IsNumeric(tag) Then
            Exit Sub
        End If
        Dim objfrmProduit As frmProduit
        Dim objProduit As Produit
        Dim nid As Long
        Dim bReturn As Boolean

        objfrmProduit = New frmProduit
        nid = tag
        objProduit = Produit.createandload(nid)
        bReturn = objfrmProduit.setElementCourant2(objProduit)
        If bReturn Then
            objfrmProduit.AfficheElementCourant()
        End If
        objfrmProduit.MdiParent = Me.MdiParent
        objfrmProduit.Show()


    End Sub

#End Region

#Region "Gestion des Evénements"
    Private Sub cbValider_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbValider.Click
        MAJElement(m_elementCourant)
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub cbAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnuler.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()

    End Sub

    Private Sub dlgLgCommande_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(27) Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If
    End Sub

    Private Sub tbCodeProduit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbCodeProduit.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            rechercheProduit()
        End If
    End Sub
    Private Sub cbProduit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbProduit.Click
        rechercheProduit()
    End Sub
    Private Sub cbCalculTotal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCalculTotal.Click
        CalculPrix()
    End Sub


    Private Sub ckGratuit_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckGratuit.CheckedChanged
        If Not m_bAffichageEncours Then
            setGratuit(ckGratuit.Checked)
        End If
    End Sub


    Private Sub tbPrixUHT_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbPrixUHT.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            CalculPrix()
        End If

    End Sub

    Private Sub dlgLgCommande_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.tbCodeProduit.Focus()
    End Sub
#End Region

    Private Sub tbQteCom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbQteCom.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            GetNextControl(tbQteCom, True).Focus()
            e.Handled = True
        End If '

    End Sub

    Private Sub tbQteCom_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbQteCom.Validated
        CalculPrix()
    End Sub

    Private Sub tbQteLiv_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbQteLiv.Validated
        CalculPrix()
    End Sub

    Private Sub tbQteFact_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbQteFact.Validated
        CalculPrix()
    End Sub
End Class
