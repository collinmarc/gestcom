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
    Friend WithEvents DatemvtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibelleDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtdateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtDatedeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents cbGenerer As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtDateFacture As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbPeriode As System.Windows.Forms.TextBox
    Friend WithEvents grpFact As System.Windows.Forms.GroupBox
    Friend WithEvents dtDateFactCourante As System.Windows.Forms.DateTimePicker
    Friend WithEvents tbPeriodeFactCourante As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents tbMontantTTCFactCourante As textBoxCurrency
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbMontantHTFactCourante As textBoxCurrency
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents tbCodeFournisseur As System.Windows.Forms.TextBox
    Friend WithEvents liFacture As System.Windows.Forms.LinkLabel
    Friend WithEvents liTiers As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.tbCodeFournisseur = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.cbGenerer = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtDateFacture = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.tbPeriode = New System.Windows.Forms.TextBox
        Me.grpFact = New System.Windows.Forms.GroupBox
        Me.cbSave = New System.Windows.Forms.Button
        Me.dtDateFactCourante = New System.Windows.Forms.DateTimePicker
        Me.tbPeriodeFactCourante = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.tbMontantTTCFactCourante = New vini_app.textBoxCurrency
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbMontantHTFactCourante = New vini_app.textBoxCurrency
        Me.Label2 = New System.Windows.Forms.Label
        Me.liTiers = New System.Windows.Forms.LinkLabel
        Me.liFacture = New System.Windows.Forms.LinkLabel
        Me.cbRecherche = New System.Windows.Forms.Button
        Me.dgvMvtStock = New System.Windows.Forms.DataGridView
        Me.DatemvtDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibelleDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcMvtStock = New System.Windows.Forms.BindingSource(Me.components)
        Me.grpFact.SuspendLayout()
        CType(Me.dgvMvtStock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcMvtStock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbCodeFournisseur
        '
        Me.tbCodeFournisseur.Location = New System.Drawing.Point(232, 56)
        Me.tbCodeFournisseur.Name = "tbCodeFournisseur"
        Me.tbCodeFournisseur.Size = New System.Drawing.Size(88, 20)
        Me.tbCodeFournisseur.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 108
        Me.Label1.Text = "Fournisseur"
        '
        'dtdateFin
        '
        Me.dtdateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdateFin.Location = New System.Drawing.Point(232, 32)
        Me.dtdateFin.Name = "dtdateFin"
        Me.dtdateFin.Size = New System.Drawing.Size(88, 20)
        Me.dtdateFin.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(176, 16)
        Me.Label8.TabIndex = 107
        Me.Label8.Text = "date de fin de Livraison"
        '
        'dtDatedeb
        '
        Me.dtDatedeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDatedeb.Location = New System.Drawing.Point(232, 8)
        Me.dtDatedeb.Name = "dtDatedeb"
        Me.dtDatedeb.Size = New System.Drawing.Size(88, 20)
        Me.dtDatedeb.TabIndex = 0
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 8)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(176, 16)
        Me.Label14.TabIndex = 106
        Me.Label14.Text = "date de début de Livraison"
        '
        'cbAfficher
        '
        Me.cbAfficher.BackColor = System.Drawing.SystemColors.Control
        Me.cbAfficher.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAfficher.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAfficher.Location = New System.Drawing.Point(8, 80)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAfficher.Size = New System.Drawing.Size(352, 24)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "A&fficher les Mouvements de Stocks"
        Me.cbAfficher.UseVisualStyleBackColor = False
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
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(421, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 119
        Me.Label4.Text = "Date de Facture"
        '
        'dtDateFacture
        '
        Me.dtDateFacture.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFacture.Location = New System.Drawing.Point(525, 8)
        Me.dtDateFacture.Name = "dtDateFacture"
        Me.dtDateFacture.Size = New System.Drawing.Size(104, 20)
        Me.dtDateFacture.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.Location = New System.Drawing.Point(423, 36)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 16)
        Me.Label6.TabIndex = 123
        Me.Label6.Text = "Période"
        '
        'tbPeriode
        '
        Me.tbPeriode.Location = New System.Drawing.Point(525, 34)
        Me.tbPeriode.Name = "tbPeriode"
        Me.tbPeriode.Size = New System.Drawing.Size(176, 20)
        Me.tbPeriode.TabIndex = 12
        '
        'grpFact
        '
        Me.grpFact.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpFact.Controls.Add(Me.cbSave)
        Me.grpFact.Controls.Add(Me.dtDateFactCourante)
        Me.grpFact.Controls.Add(Me.tbPeriodeFactCourante)
        Me.grpFact.Controls.Add(Me.Label10)
        Me.grpFact.Controls.Add(Me.Label7)
        Me.grpFact.Controls.Add(Me.tbMontantTTCFactCourante)
        Me.grpFact.Controls.Add(Me.Label3)
        Me.grpFact.Controls.Add(Me.tbMontantHTFactCourante)
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
        'tbMontantTTCFactCourante
        '
        Me.tbMontantTTCFactCourante.Location = New System.Drawing.Point(112, 183)
        Me.tbMontantTTCFactCourante.Name = "tbMontantTTCFactCourante"
        Me.tbMontantTTCFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantTTCFactCourante.TabIndex = 6
        Me.tbMontantTTCFactCourante.Text = "0"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 186)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 24)
        Me.Label3.TabIndex = 134
        Me.Label3.Text = "Montant TTC"
        '
        'tbMontantHTFactCourante
        '
        Me.tbMontantHTFactCourante.Location = New System.Drawing.Point(112, 146)
        Me.tbMontantHTFactCourante.Name = "tbMontantHTFactCourante"
        Me.tbMontantHTFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantHTFactCourante.TabIndex = 5
        Me.tbMontantHTFactCourante.Text = "0"
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
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(150, 56)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(75, 23)
        Me.cbRecherche.TabIndex = 3
        Me.cbRecherche.Text = "Recherche"
        '
        'dgvMvtStock
        '
        Me.dgvMvtStock.AllowUserToAddRows = False
        Me.dgvMvtStock.AllowUserToDeleteRows = False
        Me.dgvMvtStock.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMvtStock.AutoGenerateColumns = False
        Me.dgvMvtStock.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvMvtStock.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMvtStock.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DatemvtDataGridViewTextBoxColumn, Me.LibelleDataGridViewTextBoxColumn, Me.QteDataGridViewTextBoxColumn})
        Me.dgvMvtStock.DataSource = Me.m_bsrcMvtStock
        Me.dgvMvtStock.Location = New System.Drawing.Point(13, 111)
        Me.dgvMvtStock.Name = "dgvMvtStock"
        Me.dgvMvtStock.ReadOnly = True
        Me.dgvMvtStock.Size = New System.Drawing.Size(347, 472)
        Me.dgvMvtStock.TabIndex = 124
        '
        'DatemvtDataGridViewTextBoxColumn
        '
        Me.DatemvtDataGridViewTextBoxColumn.DataPropertyName = "datemvt"
        Me.DatemvtDataGridViewTextBoxColumn.HeaderText = "Date"
        Me.DatemvtDataGridViewTextBoxColumn.Name = "DatemvtDataGridViewTextBoxColumn"
        Me.DatemvtDataGridViewTextBoxColumn.ReadOnly = True
        Me.DatemvtDataGridViewTextBoxColumn.Width = 55
        '
        'LibelleDataGridViewTextBoxColumn
        '
        Me.LibelleDataGridViewTextBoxColumn.DataPropertyName = "libelle"
        Me.LibelleDataGridViewTextBoxColumn.HeaderText = "libelle"
        Me.LibelleDataGridViewTextBoxColumn.Name = "LibelleDataGridViewTextBoxColumn"
        Me.LibelleDataGridViewTextBoxColumn.ReadOnly = True
        Me.LibelleDataGridViewTextBoxColumn.Width = 58
        '
        'QteDataGridViewTextBoxColumn
        '
        Me.QteDataGridViewTextBoxColumn.DataPropertyName = "qte"
        Me.QteDataGridViewTextBoxColumn.HeaderText = "qte"
        Me.QteDataGridViewTextBoxColumn.Name = "QteDataGridViewTextBoxColumn"
        Me.QteDataGridViewTextBoxColumn.ReadOnly = True
        Me.QteDataGridViewTextBoxColumn.Width = 47
        '
        'm_bsrcMvtStock
        '
        Me.m_bsrcMvtStock.DataSource = GetType(vini_DB.mvtStock)
        '
        'frmGeneFactColisage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(997, 630)
        Me.Controls.Add(Me.dgvMvtStock)
        Me.Controls.Add(Me.cbRecherche)
        Me.Controls.Add(Me.grpFact)
        Me.Controls.Add(Me.tbPeriode)
        Me.Controls.Add(Me.tbCodeFournisseur)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dtDateFacture)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbGenerer)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtdateFin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtDatedeb)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cbAfficher)
        Me.Name = "frmGeneFactColisage"
        Me.Text = "Génération de factures de Colisage"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grpFact.ResumeLayout(False)
        Me.grpFact.PerformLayout()
        CType(Me.dgvMvtStock, System.ComponentModel.ISupportInitialize).EndInit()
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
        dtDatedeb.Enabled = True
        dtdateFin.Enabled = True
        tbCodeFournisseur.Enabled = True
        cbAfficher.Enabled = True
        dtDateFacture.Enabled = True
        tbPeriode.Enabled = True
        cbRecherche.Enabled = True
        For Each objControl In grpFact.Controls
            objControl.Enabled = True
        Next objControl

    End Sub

    Protected Shadows Function getElementCourant() As FactColisage
        Return CType(m_ElementCourant, FactColisage)
    End Function

    Protected Overrides Function frmSave() As Boolean
        If MyBase.frmSave() Then
            afficheFactureCourante()
        End If
    End Function


#End Region

#Region "Methodes privées"
    Private Sub initFenetre()
        Dim MoisSuivant As Date
        Dim premierMoisSuivant As Date
        Dim dernierMoisCourant As Date
        'Date de Début = 01 du mois Courant
        dtDatedeb.Value = "01/" & Now.Month() & "/" & Now.Year
        MoisSuivant = DateAdd(DateInterval.Month, +1, Now())
        premierMoisSuivant = "01/" & MoisSuivant.Month() & "/" & MoisSuivant.Year()
        dernierMoisCourant = DateAdd(DateInterval.Day, -1, premierMoisSuivant)

        dtdateFin.Value = dernierMoisCourant
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
        Dim col As Collection
        Dim bReturn As Boolean
        Dim oFRN As Fournisseur

        debAffiche()
        setcursorWait()
        Try

            ddeb = dtDatedeb.Value.ToShortDateString
            dfin = dtdateFin.Value.ToShortDateString
            codeFournisseur = tbCodeFournisseur.Text
            oFRN = Fournisseur.createandload(codeFournisseur)
            If Not oFRN Is Nothing Then
                'Recupération de la liste des Mouvements de stocks
                col = mvtStock.getListe2(ddeb, dfin, oFRN, vncEtatMVTSTK.vncMVTSTK_nFact)
                If col Is Nothing Then
                    bReturn = False
                Else
                    'On ne prend sur les Mouvements des produits en stock
                    m_colMouvementsStock = col
                    bReturn = True
                End If
            Else
                bReturn = False
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
        Dim oFRN As Fournisseur
        Dim oFact As FactColisage
        oFRN = Fournisseur.createandload(tbCodeFournisseur.Text)
        If (Not oFRN Is Nothing) Then
            setcursorWait()
            oFact = FactColisage.GenereFacture(dtDatedeb.Value.ToShortDateString(), dtdateFin.Value.ToShortDateString(), oFRN)
            If Not oFact Is Nothing Then
                oFact.periode = tbPeriode.Text
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
        getElementCourant().totalHT = tbMontantHTFactCourante.Text
        getElementCourant().totalTTC = tbMontantTTCFactCourante.Text
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
        tbMontantHTFactCourante.Text = getElementCourant().totalHT
        tbMontantTTCFactCourante.Text = getElementCourant().totalTTC
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
                bReturn = bReturn And tbMontantHTFactCourante.Text = getElementCourant().totalHT
                bReturn = bReturn And tbMontantTTCFactCourante.Text = getElementCourant().totalTTC
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
End Class
