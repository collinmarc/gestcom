Option Strict Off
Imports vini_DB
Public Class frmGeneFactColisage
    Inherits frmDonBase
    Protected m_colMouvementsStock As Collection
    Private Const COL_STKDATE As Integer = 0
    Private Const COL_STKQTE As Integer = 1
    Private Const COL_STKLIBELLE As Integer = 2
    Private Const COL_STKINDEX As Integer = 3
    Friend WithEvents cbSave As System.Windows.Forms.Button
    Friend WithEvents dgvMvtStock As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcMvtStock As System.Windows.Forms.BindingSource
    Friend WithEvents cbxDossier As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DatemvtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CodeProduit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibelleDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NomProduit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents tbMontantHT As System.Windows.Forms.TextBox
    Friend WithEvents tbMontantTTC As System.Windows.Forms.TextBox
    Private Const COL_NBRECOL As Integer = 4

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
    Friend WithEvents lblFournisseur As System.Windows.Forms.Label
    Friend WithEvents dtDatedeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents cbGenerer As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtDateFacture As System.Windows.Forms.DateTimePicker
    Friend WithEvents grpFact As System.Windows.Forms.GroupBox
    Friend WithEvents dtDateFactCourante As System.Windows.Forms.DateTimePicker
    Friend WithEvents tbPeriodeFactCourante As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    'Friend WithEvents tbMontantTTCFactCourante As textBoxCurrency
    Friend WithEvents Label3 As System.Windows.Forms.Label
    'Friend WithEvents tbMontantHTFactCourante As textBoxCurrency
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents tbCodeFournisseur As System.Windows.Forms.TextBox
    Friend WithEvents liFacture As System.Windows.Forms.LinkLabel
    Friend WithEvents liTiers As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cbxDossier = New System.Windows.Forms.ComboBox()
        Me.dgvMvtStock = New System.Windows.Forms.DataGridView()
        Me.cbRecherche = New System.Windows.Forms.Button()
        Me.grpFact = New System.Windows.Forms.GroupBox()
        Me.cbSave = New System.Windows.Forms.Button()
        Me.dtDateFactCourante = New System.Windows.Forms.DateTimePicker()
        Me.tbPeriodeFactCourante = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.liTiers = New System.Windows.Forms.LinkLabel()
        Me.liFacture = New System.Windows.Forms.LinkLabel()
        Me.tbCodeFournisseur = New System.Windows.Forms.TextBox()
        Me.dtDateFacture = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbGenerer = New System.Windows.Forms.Button()
        Me.lblFournisseur = New System.Windows.Forms.Label()
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.m_bsrcMvtStock = New System.Windows.Forms.BindingSource(Me.components)
        Me.DatemvtDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CodeProduit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LibelleDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NomProduit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tbMontantTTC = New System.Windows.Forms.TextBox()
        Me.tbMontantHT = New System.Windows.Forms.TextBox()
        CType(Me.dgvMvtStock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpFact.SuspendLayout()
        CType(Me.m_bsrcMvtStock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 6)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 13)
        Me.Label5.TabIndex = 126
        Me.Label5.Text = "Dossier : "
        '
        'cbxDossier
        '
        Me.cbxDossier.FormattingEnabled = True
        Me.cbxDossier.Location = New System.Drawing.Point(232, 3)
        Me.cbxDossier.Name = "cbxDossier"
        Me.cbxDossier.Size = New System.Drawing.Size(103, 21)
        Me.cbxDossier.TabIndex = 125
        '
        'dgvMvtStock
        '
        Me.dgvMvtStock.AllowUserToAddRows = False
        Me.dgvMvtStock.AllowUserToDeleteRows = False
        Me.dgvMvtStock.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMvtStock.AutoGenerateColumns = False
        Me.dgvMvtStock.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvMvtStock.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMvtStock.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DatemvtDataGridViewTextBoxColumn, Me.CodeProduit, Me.LibelleDataGridViewTextBoxColumn, Me.QteDataGridViewTextBoxColumn, Me.NomProduit})
        Me.dgvMvtStock.DataSource = Me.m_bsrcMvtStock
        Me.dgvMvtStock.Location = New System.Drawing.Point(13, 111)
        Me.dgvMvtStock.Name = "dgvMvtStock"
        Me.dgvMvtStock.RowHeadersVisible = False
        Me.dgvMvtStock.RowHeadersWidth = 10
        Me.dgvMvtStock.Size = New System.Drawing.Size(347, 472)
        Me.dgvMvtStock.TabIndex = 124
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(150, 56)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(75, 23)
        Me.cbRecherche.TabIndex = 3
        Me.cbRecherche.Text = "Recherche"
        '
        'grpFact
        '
        Me.grpFact.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpFact.Controls.Add(Me.tbMontantHT)
        Me.grpFact.Controls.Add(Me.tbMontantTTC)
        Me.grpFact.Controls.Add(Me.cbSave)
        Me.grpFact.Controls.Add(Me.dtDateFactCourante)
        Me.grpFact.Controls.Add(Me.tbPeriodeFactCourante)
        Me.grpFact.Controls.Add(Me.Label10)
        Me.grpFact.Controls.Add(Me.Label7)
        Me.grpFact.Controls.Add(Me.Label3)
        Me.grpFact.Controls.Add(Me.Label2)
        Me.grpFact.Controls.Add(Me.liTiers)
        Me.grpFact.Controls.Add(Me.liFacture)
        Me.grpFact.Location = New System.Drawing.Point(426, 132)
        Me.grpFact.Name = "grpFact"
        Me.grpFact.Size = New System.Drawing.Size(536, 253)
        Me.grpFact.TabIndex = 15
        Me.grpFact.TabStop = False
        '
        'cbSave
        '
        Me.cbSave.Location = New System.Drawing.Point(416, 222)
        Me.cbSave.Name = "cbSave"
        Me.cbSave.Size = New System.Drawing.Size(97, 23)
        Me.cbSave.TabIndex = 141
        Me.cbSave.Text = "Sauvegarder"
        Me.cbSave.UseVisualStyleBackColor = True
        '
        'dtDateFactCourante
        '
        Me.dtDateFactCourante.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFactCourante.Location = New System.Drawing.Point(112, 72)
        Me.dtDateFactCourante.Name = "dtDateFactCourante"
        Me.dtDateFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.dtDateFactCourante.TabIndex = 2
        '
        'tbPeriodeFactCourante
        '
        Me.tbPeriodeFactCourante.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPeriodeFactCourante.Location = New System.Drawing.Point(112, 109)
        Me.tbPeriodeFactCourante.Name = "tbPeriodeFactCourante"
        Me.tbPeriodeFactCourante.Size = New System.Drawing.Size(200, 20)
        Me.tbPeriodeFactCourante.TabIndex = 4
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 112)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(88, 24)
        Me.Label10.TabIndex = 140
        Me.Label10.Text = "Periode"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 136
        Me.Label7.Text = "DateFacture"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 186)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 24)
        Me.Label3.TabIndex = 134
        Me.Label3.Text = "Montant TTC"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 149)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 132
        Me.Label2.Text = "Montant HT"
        '
        'liTiers
        '
        Me.liTiers.Location = New System.Drawing.Point(8, 48)
        Me.liTiers.Name = "liTiers"
        Me.liTiers.Size = New System.Drawing.Size(520, 16)
        Me.liTiers.TabIndex = 1
        Me.liTiers.TabStop = True
        Me.liTiers.Text = "RS Tiers"
        '
        'liFacture
        '
        Me.liFacture.Location = New System.Drawing.Point(8, 24)
        Me.liFacture.Name = "liFacture"
        Me.liFacture.Size = New System.Drawing.Size(520, 24)
        Me.liFacture.TabIndex = 0
        Me.liFacture.TabStop = True
        Me.liFacture.Text = "Reference Facture"
        '
        'tbCodeFournisseur
        '
        Me.tbCodeFournisseur.Location = New System.Drawing.Point(232, 56)
        Me.tbCodeFournisseur.Name = "tbCodeFournisseur"
        Me.tbCodeFournisseur.Size = New System.Drawing.Size(88, 20)
        Me.tbCodeFournisseur.TabIndex = 2
        '
        'dtDateFacture
        '
        Me.dtDateFacture.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtDateFacture.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFacture.Location = New System.Drawing.Point(525, 8)
        Me.dtDateFacture.Name = "dtDateFacture"
        Me.dtDateFacture.Size = New System.Drawing.Size(104, 20)
        Me.dtDateFacture.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(421, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 119
        Me.Label4.Text = "Date de Facture"
        '
        'cbGenerer
        '
        Me.cbGenerer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbGenerer.Location = New System.Drawing.Point(424, 80)
        Me.cbGenerer.Name = "cbGenerer"
        Me.cbGenerer.Size = New System.Drawing.Size(493, 24)
        Me.cbGenerer.TabIndex = 8
        Me.cbGenerer.Text = "Générer"
        '
        'lblFournisseur
        '
        Me.lblFournisseur.Location = New System.Drawing.Point(8, 56)
        Me.lblFournisseur.Name = "lblFournisseur"
        Me.lblFournisseur.Size = New System.Drawing.Size(136, 16)
        Me.lblFournisseur.TabIndex = 108
        Me.lblFournisseur.Text = "Code Fournisseur"
        '
        'dtDatedeb
        '
        Me.dtDatedeb.CustomFormat = "MMMM yyyy"
        Me.dtDatedeb.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtDatedeb.Location = New System.Drawing.Point(232, 30)
        Me.dtDatedeb.Name = "dtDatedeb"
        Me.dtDatedeb.Size = New System.Drawing.Size(103, 20)
        Me.dtDatedeb.TabIndex = 0
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(10, 30)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(176, 16)
        Me.Label14.TabIndex = 106
        Me.Label14.Text = "Mois de facturation"
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.BackColor = System.Drawing.SystemColors.Control
        Me.cbAfficher.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAfficher.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAfficher.Location = New System.Drawing.Point(13, 80)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAfficher.Size = New System.Drawing.Size(347, 24)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "A&fficher les Mouvements de Stocks"
        Me.cbAfficher.UseVisualStyleBackColor = False
        '
        'm_bsrcMvtStock
        '
        Me.m_bsrcMvtStock.DataSource = GetType(vini_DB.mvtStock)
        '
        'DatemvtDataGridViewTextBoxColumn
        '
        Me.DatemvtDataGridViewTextBoxColumn.DataPropertyName = "datemvt"
        Me.DatemvtDataGridViewTextBoxColumn.HeaderText = "Date"
        Me.DatemvtDataGridViewTextBoxColumn.Name = "DatemvtDataGridViewTextBoxColumn"
        Me.DatemvtDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CodeProduit
        '
        Me.CodeProduit.DataPropertyName = "CodeProduit"
        Me.CodeProduit.HeaderText = "CodeProduit"
        Me.CodeProduit.Name = "CodeProduit"
        Me.CodeProduit.ReadOnly = True
        '
        'LibelleDataGridViewTextBoxColumn
        '
        Me.LibelleDataGridViewTextBoxColumn.DataPropertyName = "libelle"
        Me.LibelleDataGridViewTextBoxColumn.HeaderText = "libelle"
        Me.LibelleDataGridViewTextBoxColumn.Name = "LibelleDataGridViewTextBoxColumn"
        Me.LibelleDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QteDataGridViewTextBoxColumn
        '
        Me.QteDataGridViewTextBoxColumn.DataPropertyName = "qte"
        Me.QteDataGridViewTextBoxColumn.HeaderText = "qte"
        Me.QteDataGridViewTextBoxColumn.Name = "QteDataGridViewTextBoxColumn"
        Me.QteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'NomProduit
        '
        Me.NomProduit.DataPropertyName = "NomProduit"
        Me.NomProduit.HeaderText = "NomProduit"
        Me.NomProduit.Name = "NomProduit"
        Me.NomProduit.ReadOnly = True
        '
        'TextBox1
        '
        Me.tbMontantTTC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMontantTTC.Location = New System.Drawing.Point(112, 144)
        Me.tbMontantTTC.Name = "TextBox1"
        Me.tbMontantTTC.Size = New System.Drawing.Size(200, 20)
        Me.tbMontantTTC.TabIndex = 142
        '
        'TextBox2
        '
        Me.tbMontantHT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMontantHT.Location = New System.Drawing.Point(112, 186)
        Me.tbMontantHT.Name = "TextBox2"
        Me.tbMontantHT.Size = New System.Drawing.Size(200, 20)
        Me.tbMontantHT.TabIndex = 143
        '
        'frmGeneFactColisage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(997, 630)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cbxDossier)
        Me.Controls.Add(Me.dgvMvtStock)
        Me.Controls.Add(Me.cbRecherche)
        Me.Controls.Add(Me.grpFact)
        Me.Controls.Add(Me.tbCodeFournisseur)
        Me.Controls.Add(Me.dtDateFacture)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbGenerer)
        Me.Controls.Add(Me.lblFournisseur)
        Me.Controls.Add(Me.dtDatedeb)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cbAfficher)
        Me.Name = "frmGeneFactColisage"
        Me.Text = "Génération de factures de colisage"
        CType(Me.dgvMvtStock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpFact.ResumeLayout(False)
        Me.grpFact.PerformLayout()
        CType(Me.m_bsrcMvtStock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Méthodes Vinicom"
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False
        m_ToolBarSaveEnabled = True
    End Sub

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        Dim objControl As Control
        MyBase.EnableControls(bEnabled)
        cbxDossier.Enabled = True
        dtDatedeb.Enabled = True
        tbCodeFournisseur.Enabled = True
        cbAfficher.Enabled = True
        dtDateFacture.Enabled = True
        cbRecherche.Enabled = True
        For Each objControl In grpFact.Controls
            objControl.Enabled = True
        Next objControl

    End Sub

    Protected Shadows Function getElementCourant() As FactColisageJ
        Return CType(m_ElementCourant, FactColisageJ)
    End Function

    Protected Overrides Function frmSave() As Boolean
        If MyBase.frmSave() Then
            afficheFactureCourante()
        End If
    End Function


#End Region

#Region "Methodes privées"
    Private Sub initFenetre()
        'Initialisation de la COMBO Dossier
        cbxDossier.Items.Clear()
        cbxDossier.Items.Add(Dossier.VINICOM)
        cbxDossier.Items.Add(Dossier.HOBIVIN)
        cbxDossier.Text = Dossier.VINICOM

        'Date de Début = 01 du mois Courant
        dtDatedeb.Value = "01/" & Now.Month() & "/" & Now.Year

        dtDateFacture.Value = Now()
        m_colMouvementsStock = New Collection
    End Sub
    Private Sub afficheListeMvtStock()

        setcursorWait()
        debAffiche()
        m_bsrcMvtStock.Clear()

        Debug.Assert(Not m_colMouvementsStock Is Nothing)
        Dim objMvt As mvtStock
        Dim nRow As Integer


        nRow = 0
        For Each objMvt In m_colMouvementsStock
            m_bsrcMvtStock.Add(objMvt)
        Next objMvt

        cbGenerer.Enabled = True

        finAffiche()
        restoreCursor()

    End Sub 'AfficheListeMvtStock

    ''' <summary>
    ''' Créé la liste des mouvements de stocks en fonction des paramètres
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    Private Function setListeMouvementsStock() As Boolean
        Dim ddeb As Date
        Dim dfin As Date
        Dim codeFournisseur As String
        Dim strDossier As String
        Dim col As Collection
        Dim bReturn As Boolean
        Dim oFRN As Fournisseur
        If tbCodeFournisseur.Text = "" And cbxDossier.Text = Dossier.VINICOM Then
            MsgBox("Saisie d'un code fournisseur Obligatoire pour le dossier VINICOM")
            Return False
        End If

        codeFournisseur = tbCodeFournisseur.Text
        strDossier = cbxDossier.Text

        debAffiche()
        setcursorWait()
        Try

            ddeb = dtDatedeb.Value
            dfin = ddeb.AddMonths(1).AddDays(-1)
            col = Nothing
            If strDossier = Dossier.VINICOM Then
                oFRN = Fournisseur.createandload(codeFournisseur)
                If Not oFRN Is Nothing Then
                    'Recupération de la liste des Mouvements de stocks
                    col = mvtStock.getListe2(ddeb, dfin, oFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
                End If
            End If
            If strDossier = Dossier.HOBIVIN Then
                'Recupération de la liste des Mouvements de stocks
                col = mvtStock.getListeDossierNonFacture(strDossier, ddeb, dfin)
            End If
            If col Is Nothing Then
                bReturn = False
            Else
                'On ne prend sur les Mouvements des produits en stock
                m_colMouvementsStock = col
                bReturn = True
            End If
        Catch ex As Exception
            bReturn = False
            Debug.Assert(bReturn, ex.ToString)
        End Try
        finAffiche()
        restoreCursor()
        Return bReturn
    End Function
    ''' <summary>
    ''' Generatin d'une facture de colisage avec les paramètres
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub genererFactures()
        Dim colFact As New Collection
        Dim oFRN As Fournisseur = Nothing
        Dim oFact As FactColisageJ
        If cbxDossier.Text = Dossier.VINICOM Then
            oFRN = Fournisseur.createandload(tbCodeFournisseur.Text)
        End If
        If cbxDossier.Text = Dossier.HOBIVIN Then
            oFRN = Fournisseur.getIntermediairePourUnDossier(Dossier.HOBIVIN)
        End If



        If (Not oFRN Is Nothing) Then
            setcursorWait()
            oFact = FactColisageJ.GenereFacture(dtDatedeb.Value, oFRN, cbxDossier.Text)
            If Not oFact Is Nothing Then
                oFact.dateFacture = dtDateFacture.Value.ToShortDateString()
                setElementCourant2(oFact)
            End If
            restoreCursor()
        End If

    End Sub 'genererFactures

    Private Function appliqueModifications() As Boolean
        Dim bReturn As Boolean
        majElement()
        bReturn = setElementCourant2(Nothing)
        Return bReturn
    End Function
    Public Overrides Function majElement() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing, "Pas de Facture courante")
        'Facture Fournisseur
        getElementCourant().dateCommande = dtDateFactCourante.Value
        getElementCourant().totalHT = tbMontantTTC.Text
        getElementCourant().totalTTC = tbMontantHT.Text
        getElementCourant().periode = tbPeriodeFactCourante.Text
        Return True
    End Function 'majElement


    Private Function afficheFactureCourante() As Boolean
        Debug.Assert(Not getElementCourant() Is Nothing, "La Facture courante doit être renseignée")

        'Affichage de la Facture
        debAffiche()
        liFacture.Text = getElementCourant().shortResume
        liFacture.Tag = getElementCourant().id
        If getElementCourant().id = 0 Then
            liFacture.Enabled = False
        Else
            liFacture.Enabled = True
        End If

        liTiers.Text = getElementCourant().oTiers.rs
        liTiers.Tag = getElementCourant().oTiers.id

        dtDateFactCourante.Value = getElementCourant().dateCommande
        tbPeriodeFactCourante.Text = getElementCourant().periode
        tbMontantTTC.Text = getElementCourant().totalHT
        tbMontantHT.Text = getElementCourant().totalTTC
        grpFact.Enabled = True

        finAffiche()
        Return True

    End Function 'AfficheFactureCourante

    Private Function elementCourantSansModif() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
            If Not getElementCourant() Is Nothing Then
                bReturn = bReturn And dtDateFactCourante.Value = getElementCourant().dateCommande
                bReturn = bReturn And tbPeriodeFactCourante.Text = getElementCourant().periode
                bReturn = bReturn And tbMontantTTC.Text = getElementCourant().totalHT
                bReturn = bReturn And tbMontantHT.Text = getElementCourant().totalTTC
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Sub rechercheClient()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.FOURNISSEUR, tbCodeFournisseur)

        If Not objTiers Is Nothing Then
            tbCodeFournisseur.Text = objTiers.code
        End If
    End Sub 'rechercheClient

#End Region
#Region "Gestion des Evenements"
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        If setListeMouvementsStock() Then
            'm_SCMDcourante = Nothing
            afficheListeMvtStock()
        End If
        dgvMvtStock.Enabled = True

    End Sub





#End Region

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Génerer déclenche L'evenement Updated
        'AddHandler cbAppliquer.Click, AddressOf ControlUpdated
        'AddHandler cbGenerer.Click, AddressOf ControlUpdated
    End Sub

    'Private Sub cbSelectionnerTout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSelectionnerTout.Click
    '    Call selectionnerTout()
    'End Sub

    Private Sub cbGenerer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbGenerer.Click
        Call sauvegardeElementCourant()
        Call genererFactures()
        Call afficheFactureCourante()
        If Not getElementCourant() Is Nothing Then
            cbGenerer.Enabled = False
            setfrmUpdated()
        End If
    End Sub


    Private Sub cbAppliquer_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If appliqueModifications() Then
            grpFact.Enabled = False
        End If
    End Sub

    Private Sub liTiers_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liTiers.LinkClicked
        afficheFenetreClient(liTiers.Tag)
    End Sub

    Private Sub liFactTRP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liFacture.LinkClicked
        afficheFenetreFactColisage(liFacture.Tag)
    End Sub

    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        rechercheClient()
    End Sub
    Private Sub cbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSave.Click
        frmSave()
    End Sub

    Private Sub cbxDossier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxDossier.SelectedIndexChanged
        If cbxDossier.Text = Dossier.VINICOM Then
            lblFournisseur.Visible = True
            tbCodeFournisseur.Visible = True
            cbRecherche.Visible = True

        End If
        If cbxDossier.Text = Dossier.HOBIVIN Then
            lblFournisseur.Visible = False
            tbCodeFournisseur.Visible = False
            cbRecherche.Visible = False

        End If
    End Sub
End Class
