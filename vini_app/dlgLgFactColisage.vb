Imports vini_DB
Friend Class dlgLgFactColisage
    Inherits System.Windows.Forms.Form


    Private m_elementCourant As LgFactColisage
    Private m_TiersCourant As Tiers
    Private m_bModif As Boolean
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbMontantHT As vini_app.textBoxCurrency
    Friend WithEvents tbQte As vini_app.textBoxNumeric
    Friend WithEvents tbPrixUnitaire As vini_app.textBoxCurrency
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents m_bsrcLgFactColisage As System.Windows.Forms.BindingSource
    Friend WithEvents btnRechercher As System.Windows.Forms.Button
    Friend WithEvents tbCodeProduit As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblProduit As System.Windows.Forms.Label
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
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
    Friend WithEvents grpCMD As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.cbAnnuler = New System.Windows.Forms.Button()
        Me.cbValider = New System.Windows.Forms.Button()
        Me.grpCMD = New System.Windows.Forms.GroupBox()
        Me.btnRechercher = New System.Windows.Forms.Button()
        Me.tbCodeProduit = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbMontantHT = New vini_app.textBoxCurrency()
        Me.m_bsrcLgFactColisage = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbQte = New vini_app.textBoxNumeric()
        Me.tbPrixUnitaire = New vini_app.textBoxCurrency()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.lblProduit = New System.Windows.Forms.Label()
        Me.grpCMD.SuspendLayout()
        CType(Me.m_bsrcLgFactColisage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbAnnuler
        '
        Me.cbAnnuler.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAnnuler.BackColor = System.Drawing.SystemColors.Control
        Me.cbAnnuler.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAnnuler.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAnnuler.Location = New System.Drawing.Point(459, 164)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAnnuler.Size = New System.Drawing.Size(81, 25)
        Me.cbAnnuler.TabIndex = 1
        Me.cbAnnuler.Text = "&Annuler"
        Me.cbAnnuler.UseVisualStyleBackColor = False
        '
        'cbValider
        '
        Me.cbValider.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbValider.BackColor = System.Drawing.SystemColors.Control
        Me.cbValider.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbValider.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbValider.Location = New System.Drawing.Point(357, 164)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbValider.Size = New System.Drawing.Size(81, 25)
        Me.cbValider.TabIndex = 0
        Me.cbValider.Text = "&Valider"
        Me.cbValider.UseVisualStyleBackColor = False
        '
        'grpCMD
        '
        Me.grpCMD.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpCMD.Controls.Add(Me.lblProduit)
        Me.grpCMD.Controls.Add(Me.btnRechercher)
        Me.grpCMD.Controls.Add(Me.tbCodeProduit)
        Me.grpCMD.Controls.Add(Me.Label1)
        Me.grpCMD.Controls.Add(Me.tbMontantHT)
        Me.grpCMD.Controls.Add(Me.tbQte)
        Me.grpCMD.Controls.Add(Me.tbPrixUnitaire)
        Me.grpCMD.Controls.Add(Me.Label8)
        Me.grpCMD.Controls.Add(Me.Label7)
        Me.grpCMD.Controls.Add(Me.Label6)
        Me.grpCMD.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpCMD.Location = New System.Drawing.Point(8, 12)
        Me.grpCMD.Name = "grpCMD"
        Me.grpCMD.Size = New System.Drawing.Size(343, 177)
        Me.grpCMD.TabIndex = 3
        Me.grpCMD.TabStop = False
        '
        'btnRechercher
        '
        Me.btnRechercher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRechercher.Location = New System.Drawing.Point(262, 11)
        Me.btnRechercher.Name = "btnRechercher"
        Me.btnRechercher.Size = New System.Drawing.Size(75, 23)
        Me.btnRechercher.TabIndex = 33
        Me.btnRechercher.Text = "Rechercher"
        Me.btnRechercher.UseVisualStyleBackColor = True
        '
        'tbCodeProduit
        '
        Me.tbCodeProduit.Location = New System.Drawing.Point(104, 11)
        Me.tbCodeProduit.Name = "tbCodeProduit"
        Me.tbCodeProduit.Size = New System.Drawing.Size(100, 20)
        Me.tbCodeProduit.TabIndex = 32
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Code Produit :"
        '
        'tbMontantHT
        '
        Me.tbMontantHT.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "prixHT", True))
        Me.tbMontantHT.Location = New System.Drawing.Point(104, 131)
        Me.tbMontantHT.Name = "tbMontantHT"
        Me.tbMontantHT.Size = New System.Drawing.Size(120, 20)
        Me.tbMontantHT.TabIndex = 7
        Me.tbMontantHT.Text = "0"
        '
        'm_bsrcLgFactColisage
        '
        Me.m_bsrcLgFactColisage.DataSource = GetType(vini_DB.LgFactColisage)
        '
        'tbQte
        '
        Me.tbQte.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "qteCommande", True))
        Me.tbQte.Location = New System.Drawing.Point(104, 61)
        Me.tbQte.Name = "tbQte"
        Me.tbQte.Size = New System.Drawing.Size(120, 20)
        Me.tbQte.TabIndex = 5
        Me.tbQte.Text = "0"
        '
        'tbPrixUnitaire
        '
        Me.tbPrixUnitaire.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcLgFactColisage, "prixU", True))
        Me.tbPrixUnitaire.Location = New System.Drawing.Point(104, 96)
        Me.tbPrixUnitaire.Name = "tbPrixUnitaire"
        Me.tbPrixUnitaire.Size = New System.Drawing.Size(120, 20)
        Me.tbPrixUnitaire.TabIndex = 6
        Me.tbPrixUnitaire.Text = "0"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 134)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(70, 13)
        Me.Label8.TabIndex = 30
        Me.Label8.Text = "Montant HT :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 99)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 13)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Prix unitaire :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(73, 13)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "Qte Colisage :"
        '
        'BindingSource1
        '
        Me.BindingSource1.DataSource = GetType(vini_DB.Produit)
        '
        'lblProduit
        '
        Me.lblProduit.AutoSize = True
        Me.lblProduit.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.BindingSource1, "nom", True))
        Me.lblProduit.Location = New System.Drawing.Point(104, 42)
        Me.lblProduit.Name = "lblProduit"
        Me.lblProduit.Size = New System.Drawing.Size(87, 13)
        Me.lblProduit.TabIndex = 34
        Me.lblProduit.Text = "Libelle du produit"
        '
        'dlgLgFactColisage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(552, 201)
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
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
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
        m_bsrcLgFactColisage.Clear()
        m_bsrcLgFactColisage.Add(p_objLG)
        BindingSource1.Clear()
        BindingSource1.Add(p_objLG.oProduit)
        m_elementCourant = p_objLG
    End Sub 'set element Courant
    Public Function getElementCourant() As LgFactColisage
        Return m_elementCourant
    End Function
#End Region

#Region "Méthodes"
    'Initialisation de l'élement Courant
    Private Sub MAJElement()
        Debug.Assert(Not m_elementCourant Is Nothing)
        Validate()
    End Sub 'MAJElement

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

#End Region



    Private Sub tbQte_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbQte.Validated
        CalcMontantHT()
    End Sub

    Private Sub tbPrixUnitaire_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPrixUnitaire.Validated
        CalcMontantHT()
    End Sub

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

        colProduit = Produit.getListe(vncTypeProduit.vncPlateforme, tbCodeProduit.Text)

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
            If objProduit.bResume Then
                objProduit.load()
            End If
            m_elementCourant.oProduit = objProduit
            BindingSource1.Clear()
            BindingSource1.Add(objProduit)
        End If

    End Sub 'RechercheProduit


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnRechercher.Click
        rechercheProduit()
    End Sub
End Class
