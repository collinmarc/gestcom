Imports vini_DB
Friend Class dlgLgFactColisage
    Inherits System.Windows.Forms.Form
    Private m_elementCourant As LgFactColisage
    Private m_TiersCourant As Tiers
    Private m_bModif As Boolean
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtDateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbMontantHT As vini_app.textBoxCurrency
    Friend WithEvents tbStockFinal As vini_app.textBoxNumeric
    Friend WithEvents tbSorties As vini_app.textBoxNumeric
    Friend WithEvents tbEntrees As vini_app.textBoxNumeric
    Friend WithEvents tbStockInitial As vini_app.textBoxNumeric
    Friend WithEvents tbPrixUnitaire As vini_app.textBoxCurrency
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents m_bsrcLgFactColisage As System.Windows.Forms.BindingSource
    Private m_bAffichageEncours As Boolean
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
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    Public WithEvents cbAnnuler As System.Windows.Forms.Button
    Public WithEvents cbValider As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtDateDebut As System.Windows.Forms.DateTimePicker
    Friend WithEvents grpCMD As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.cbAnnuler = New System.Windows.Forms.Button
        Me.cbValider = New System.Windows.Forms.Button
        Me.grpCMD = New System.Windows.Forms.GroupBox
        Me.tbMontantHT = New vini_app.textBoxCurrency
        Me.m_bsrcLgFactColisage = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbStockFinal = New vini_app.textBoxNumeric
        Me.tbSorties = New vini_app.textBoxNumeric
        Me.tbEntrees = New vini_app.textBoxNumeric
        Me.tbStockInitial = New vini_app.textBoxNumeric
        Me.tbPrixUnitaire = New vini_app.textBoxCurrency
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtDateFin = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtDateDebut = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.grpCMD.SuspendLayout()
        CType(Me.m_bsrcLgFactColisage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbAnnuler
        '
        Me.cbAnnuler.BackColor = System.Drawing.SystemColors.Control
        Me.cbAnnuler.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAnnuler.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAnnuler.Location = New System.Drawing.Point(459, 270)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAnnuler.Size = New System.Drawing.Size(81, 25)
        Me.cbAnnuler.TabIndex = 1
        Me.cbAnnuler.Text = "&Annuler"
        Me.cbAnnuler.UseVisualStyleBackColor = False
        '
        'cbValider
        '
        Me.cbValider.BackColor = System.Drawing.SystemColors.Control
        Me.cbValider.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbValider.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbValider.Location = New System.Drawing.Point(357, 270)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbValider.Size = New System.Drawing.Size(81, 25)
        Me.cbValider.TabIndex = 0
        Me.cbValider.Text = "&Valider"
        Me.cbValider.UseVisualStyleBackColor = False
        '
        'grpCMD
        '
        Me.grpCMD.Controls.Add(Me.tbMontantHT)
        Me.grpCMD.Controls.Add(Me.tbStockFinal)
        Me.grpCMD.Controls.Add(Me.tbSorties)
        Me.grpCMD.Controls.Add(Me.tbEntrees)
        Me.grpCMD.Controls.Add(Me.tbStockInitial)
        Me.grpCMD.Controls.Add(Me.tbPrixUnitaire)
        Me.grpCMD.Controls.Add(Me.Label8)
        Me.grpCMD.Controls.Add(Me.Label7)
        Me.grpCMD.Controls.Add(Me.Label6)
        Me.grpCMD.Controls.Add(Me.Label5)
        Me.grpCMD.Controls.Add(Me.Label4)
        Me.grpCMD.Controls.Add(Me.Label3)
        Me.grpCMD.Controls.Add(Me.dtDateFin)
        Me.grpCMD.Controls.Add(Me.Label2)
        Me.grpCMD.Controls.Add(Me.dtDateDebut)
        Me.grpCMD.Controls.Add(Me.Label1)
        Me.grpCMD.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpCMD.Location = New System.Drawing.Point(8, 12)
        Me.grpCMD.Name = "grpCMD"
        Me.grpCMD.Size = New System.Drawing.Size(292, 296)
        Me.grpCMD.TabIndex = 3
        Me.grpCMD.TabStop = False
        '
        'tbMontantHT
        '
        Me.tbMontantHT.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "MontantHT", True))
        Me.tbMontantHT.Location = New System.Drawing.Point(104, 261)
        Me.tbMontantHT.Name = "tbMontantHT"
        Me.tbMontantHT.Size = New System.Drawing.Size(120, 20)
        Me.tbMontantHT.TabIndex = 7
        Me.tbMontantHT.Text = "0"
        '
        'm_bsrcLgFactColisage
        '
        Me.m_bsrcLgFactColisage.DataSource = GetType(vini_DB.LgFactColisage)
        '
        'tbStockFinal
        '
        Me.tbStockFinal.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "StockFinal", True))
        Me.tbStockFinal.Location = New System.Drawing.Point(104, 191)
        Me.tbStockFinal.Name = "tbStockFinal"
        Me.tbStockFinal.Size = New System.Drawing.Size(120, 20)
        Me.tbStockFinal.TabIndex = 5
        Me.tbStockFinal.Text = "0"
        '
        'tbSorties
        '
        Me.tbSorties.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "Sorties", True))
        Me.tbSorties.Location = New System.Drawing.Point(104, 156)
        Me.tbSorties.Name = "tbSorties"
        Me.tbSorties.Size = New System.Drawing.Size(120, 20)
        Me.tbSorties.TabIndex = 4
        Me.tbSorties.Text = "0"
        '
        'tbEntrees
        '
        Me.tbEntrees.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "Entrees", True))
        Me.tbEntrees.Location = New System.Drawing.Point(104, 121)
        Me.tbEntrees.Name = "tbEntrees"
        Me.tbEntrees.Size = New System.Drawing.Size(120, 20)
        Me.tbEntrees.TabIndex = 3
        Me.tbEntrees.Text = "0"
        '
        'tbStockInitial
        '
        Me.tbStockInitial.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "StockInitial", True))
        Me.tbStockInitial.Location = New System.Drawing.Point(104, 86)
        Me.tbStockInitial.Name = "tbStockInitial"
        Me.tbStockInitial.Size = New System.Drawing.Size(120, 20)
        Me.tbStockInitial.TabIndex = 2
        Me.tbStockInitial.Text = "0"
        '
        'tbPrixUnitaire
        '
        Me.tbPrixUnitaire.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "pu", True))
        Me.tbPrixUnitaire.Location = New System.Drawing.Point(104, 226)
        Me.tbPrixUnitaire.Name = "tbPrixUnitaire"
        Me.tbPrixUnitaire.Size = New System.Drawing.Size(120, 20)
        Me.tbPrixUnitaire.TabIndex = 6
        Me.tbPrixUnitaire.Text = "0"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 264)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(70, 13)
        Me.Label8.TabIndex = 30
        Me.Label8.Text = "Montant HT :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 229)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 13)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Prix unitaire :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 194)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(63, 13)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "Stock final :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 159)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "Sorties :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 124)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "Entrées :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 89)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Stock initial :"
        '
        'dtDateFin
        '
        Me.dtDateFin.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcLgFactColisage, "dFin", True))
        Me.dtDateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFin.Location = New System.Drawing.Point(104, 51)
        Me.dtDateFin.Name = "dtDateFin"
        Me.dtDateFin.Size = New System.Drawing.Size(120, 20)
        Me.dtDateFin.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Date fin :"
        '
        'dtDateDebut
        '
        Me.dtDateDebut.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcLgFactColisage, "dDeb", True))
        Me.dtDateDebut.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateDebut.Location = New System.Drawing.Point(104, 16)
        Me.dtDateDebut.Name = "dtDateDebut"
        Me.dtDateDebut.Size = New System.Drawing.Size(120, 20)
        Me.dtDateDebut.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Date début :"
        '
        'dlgLgFactColisage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(552, 319)
        Me.Controls.Add(Me.grpCMD)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.cbValider)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.Green
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(184, 250)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgLgFactColisage"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Ligne de Colisage"
        Me.grpCMD.ResumeLayout(False)
        Me.grpCMD.PerformLayout()
        CType(Me.m_bsrcLgFactColisage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Accesseurs"
    Public Sub New()
        MyBase.New()
        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        m_bModif = False
    End Sub

    'Public Property bModif() As Boolean
    '    Get
    '        Return m_bModif
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        m_bModif = Value
    '    End Set
    'End Property

    Public Sub setElementCourant(ByRef p_objLG As LgFactColisage)
        Debug.Assert(Not p_objLG Is Nothing)
        m_bsrcLgFactColisage.Add(p_objLG)
        m_elementCourant = p_objLG
    End Sub 'set element Courant
    Public Function getElementCourant() As LgFactColisage
        Return m_elementCourant
    End Function
    Public Sub setTiersCourant(ByVal poTiers As Fournisseur)
        Debug.Assert(Not poTiers Is Nothing, "Tiers Courant Non renseigné")
        m_TiersCourant = poTiers
    End Sub
#End Region

#Region "Méthodes"
    'Initialisation de l'élement Courant
    Private Sub MAJElement()
        Debug.Assert(Not m_elementCourant Is Nothing)
        Validate()
        m_elementCourant.qte = m_elementCourant.StockFinal
    End Sub 'MAJElement

    Private Sub CalcStockFinal()
        If (Not m_elementCourant Is Nothing) Then
            m_elementCourant.calculStockFinal()
            m_elementCourant.calculPrixTotal()
            m_bsrcLgFactColisage.ResetCurrentItem()
        End If
    End Sub
    Private Sub CalcMontantHT()
        If (Not m_elementCourant Is Nothing) Then
            m_elementCourant.calculPrixTotal()
            m_bsrcLgFactColisage.ResetCurrentItem()
        End If
    End Sub


#End Region

#Region "Gestion des Evénements"
    Private Sub cbValider_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbValider.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        MAJElement()
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

    Private Sub tbStockInitial_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbStockInitial.Validated
        CalcStockFinal()
    End Sub
#End Region

    Private Sub tbEntrees_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbEntrees.Validated
        CalcStockFinal()
    End Sub

    Private Sub tbSorties_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSorties.Validated
        CalcStockFinal()
    End Sub

    Private Sub tbStockFinal_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbStockFinal.Validated
        CalcMontantHT()
    End Sub

    Private Sub tbPrixUnitaire_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPrixUnitaire.Validated
        CalcMontantHT()
    End Sub


    Private Sub tbSorties_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles tbSorties.Validating
        Dim qte As Integer
        qte = tbSorties.Text
        If qte > 0 Then
            qte = qte * -1
            tbSorties.Text = qte
        End If
    End Sub

End Class
