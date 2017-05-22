Imports vini_DB
Public Class frmGeneFactCom
    Inherits frmDonBase
    Protected m_colSousCommandes As Collection
    Friend WithEvents m_bsrcFactCom As System.Windows.Forms.BindingSource
    Friend WithEvents dgvFactCom As System.Windows.Forms.DataGridView
    Friend WithEvents code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalHT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalTTC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cbAppliquer As System.Windows.Forms.Button
    Friend WithEvents m_bsrcSousCommande As System.Windows.Forms.BindingSource
    Friend WithEvents dgvSousCommandes As System.Windows.Forms.DataGridView
    Friend WithEvents DateReglementDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateStatistiqueDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PeriodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RefReglementDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MontantReglementDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShortResumeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BcolReglementLoadedDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents BExportInternetDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateCommandeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateFactureDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EtatDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EtatLibelleDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OTiersDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRSDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLignesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalHTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalTTCDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OTransporteurDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdParamTVADataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateValidationDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateLivraisonDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateEnlevementDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RefLivraisonDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaracteristiqueTiersDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CommCommandeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CommLivraisonDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CommFacturationDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CommLibreDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BcolLignesLoadedDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents BcolLignesUpatedDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents QteColisDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QtePalettesPrepareesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QtePalettesNonPrepareesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PoidsDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PuPalettesPrepareesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PuPalettesNonPrepareesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MontantTransportDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoutTransportDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LettreVoitureDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RefFactTRPDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MontantCommissionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BNewDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents BDeletedDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents BUpdatedDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents IdDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TypeDonneeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BResumeDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CodeDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FournisseurCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRSDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateCommandeDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalHTFacture As System.Windows.Forms.DataGridViewTextBoxColumn
    Protected m_colFact As ColEvent
    'Protected getElementCourant() As FactCom

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
    Friend WithEvents tbCodeFournisseur As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtdateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtDatedeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents cbGenerer As System.Windows.Forms.Button
    Friend WithEvents cbAnnGenerer As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtDateFacture As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtDateStatistique As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbPeriode As System.Windows.Forms.TextBox
    Friend WithEvents grpFact As System.Windows.Forms.GroupBox
    Friend WithEvents dtDateStatCourante As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtDateFactCourante As System.Windows.Forms.DateTimePicker
    Friend WithEvents tbPeriodeFactCourante As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents tbMontantTTCFactCourante As textBoxCurrency
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbMontantHTFactCourante As textBoxCurrency
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents liFournisseur As System.Windows.Forms.LinkLabel
    Friend WithEvents liFactCom As System.Windows.Forms.LinkLabel
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents cbSauvegarder As System.Windows.Forms.Button
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
        Me.cbAnnGenerer = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtDateFacture = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtDateStatistique = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.tbPeriode = New System.Windows.Forms.TextBox
        Me.grpFact = New System.Windows.Forms.GroupBox
        Me.dtDateStatCourante = New System.Windows.Forms.DateTimePicker
        Me.m_bsrcFactCom = New System.Windows.Forms.BindingSource(Me.components)
        Me.dtDateFactCourante = New System.Windows.Forms.DateTimePicker
        Me.cbAppliquer = New System.Windows.Forms.Button
        Me.tbPeriodeFactCourante = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.tbMontantTTCFactCourante = New vini_app.textBoxCurrency
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbMontantHTFactCourante = New vini_app.textBoxCurrency
        Me.Label2 = New System.Windows.Forms.Label
        Me.liFournisseur = New System.Windows.Forms.LinkLabel
        Me.liFactCom = New System.Windows.Forms.LinkLabel
        Me.cbRecherche = New System.Windows.Forms.Button
        Me.cbSauvegarder = New System.Windows.Forms.Button
        Me.dgvFactCom = New System.Windows.Forms.DataGridView
        Me.code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersRS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.totalHT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.totalTTC = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateReglementDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateStatistiqueDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PeriodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RefReglementDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MontantReglementDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ShortResumeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BcolReglementLoadedDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.BExportInternetDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateCommandeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateFactureDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EtatDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EtatLibelleDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.OTiersDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersRSDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLignesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TotalHTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TotalTTCDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.OTransporteurDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IdParamTVADataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateValidationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateLivraisonDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateEnlevementDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RefLivraisonDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CaracteristiqueTiersDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CommCommandeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CommLivraisonDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CommFacturationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CommLibreDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BcolLignesLoadedDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.BcolLignesUpatedDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.QteColisDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QtePalettesPrepareesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QtePalettesNonPrepareesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PoidsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PuPalettesPrepareesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PuPalettesNonPrepareesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MontantTransportDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoutTransportDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LettreVoitureDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RefFactTRPDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MontantCommissionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BNewDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.BDeletedDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.BUpdatedDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.IdDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TypeDonneeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BResumeDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.m_bsrcSousCommande = New System.Windows.Forms.BindingSource(Me.components)
        Me.dgvSousCommandes = New System.Windows.Forms.DataGridView
        Me.CodeDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FournisseurCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersRSDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateCommandeDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.totalHTFacture = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.grpFact.SuspendLayout()
        CType(Me.m_bsrcFactCom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvFactCom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcSousCommande, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSousCommandes, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.Label8.Text = "date de fin (Facture Fournisseur)"
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
        Me.Label14.Text = "date de début (Facture founisseur)"
        '
        'cbAfficher
        '
        Me.cbAfficher.BackColor = System.Drawing.SystemColors.Control
        Me.cbAfficher.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAfficher.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAfficher.Location = New System.Drawing.Point(8, 80)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAfficher.Size = New System.Drawing.Size(352, 25)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "A&fficher"
        Me.cbAfficher.UseVisualStyleBackColor = False
        '
        'cbGenerer
        '
        Me.cbGenerer.Location = New System.Drawing.Point(368, 160)
        Me.cbGenerer.Name = "cbGenerer"
        Me.cbGenerer.Size = New System.Drawing.Size(80, 24)
        Me.cbGenerer.TabIndex = 8
        Me.cbGenerer.Text = "- Générer ->"
        '
        'cbAnnGenerer
        '
        Me.cbAnnGenerer.Location = New System.Drawing.Point(368, 192)
        Me.cbAnnGenerer.Name = "cbAnnGenerer"
        Me.cbAnnGenerer.Size = New System.Drawing.Size(80, 24)
        Me.cbAnnGenerer.TabIndex = 16
        Me.cbAnnGenerer.Text = "<- Annule -<"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(480, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 119
        Me.Label4.Text = "Date de Facture"
        '
        'dtDateFacture
        '
        Me.dtDateFacture.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFacture.Location = New System.Drawing.Point(584, 8)
        Me.dtDateFacture.Name = "dtDateFacture"
        Me.dtDateFacture.Size = New System.Drawing.Size(104, 20)
        Me.dtDateFacture.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(480, 32)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 16)
        Me.Label5.TabIndex = 121
        Me.Label5.Text = "Date statistique"
        '
        'dtDateStatistique
        '
        Me.dtDateStatistique.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateStatistique.Location = New System.Drawing.Point(584, 32)
        Me.dtDateStatistique.Name = "dtDateStatistique"
        Me.dtDateStatistique.Size = New System.Drawing.Size(104, 20)
        Me.dtDateStatistique.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(480, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 16)
        Me.Label6.TabIndex = 123
        Me.Label6.Text = "Période"
        '
        'tbPeriode
        '
        Me.tbPeriode.Location = New System.Drawing.Point(584, 56)
        Me.tbPeriode.Name = "tbPeriode"
        Me.tbPeriode.Size = New System.Drawing.Size(176, 20)
        Me.tbPeriode.TabIndex = 12
        '
        'grpFact
        '
        Me.grpFact.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpFact.Controls.Add(Me.dtDateStatCourante)
        Me.grpFact.Controls.Add(Me.dtDateFactCourante)
        Me.grpFact.Controls.Add(Me.cbAppliquer)
        Me.grpFact.Controls.Add(Me.tbPeriodeFactCourante)
        Me.grpFact.Controls.Add(Me.Label10)
        Me.grpFact.Controls.Add(Me.Label9)
        Me.grpFact.Controls.Add(Me.Label7)
        Me.grpFact.Controls.Add(Me.tbMontantTTCFactCourante)
        Me.grpFact.Controls.Add(Me.Label3)
        Me.grpFact.Controls.Add(Me.tbMontantHTFactCourante)
        Me.grpFact.Controls.Add(Me.Label2)
        Me.grpFact.Controls.Add(Me.liFournisseur)
        Me.grpFact.Controls.Add(Me.liFactCom)
        Me.grpFact.Location = New System.Drawing.Point(456, 352)
        Me.grpFact.Name = "grpFact"
        Me.grpFact.Size = New System.Drawing.Size(536, 248)
        Me.grpFact.TabIndex = 15
        Me.grpFact.TabStop = False
        '
        'dtDateStatCourante
        '
        Me.dtDateStatCourante.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcFactCom, "dateStatistique", True))
        Me.dtDateStatCourante.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateStatCourante.Location = New System.Drawing.Point(112, 100)
        Me.dtDateStatCourante.Name = "dtDateStatCourante"
        Me.dtDateStatCourante.Size = New System.Drawing.Size(96, 20)
        Me.dtDateStatCourante.TabIndex = 3
        '
        'm_bsrcFactCom
        '
        Me.m_bsrcFactCom.DataSource = GetType(vini_DB.FactCom)
        '
        'dtDateFactCourante
        '
        Me.dtDateFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcFactCom, "dateFacture", True))
        Me.dtDateFactCourante.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFactCourante.Location = New System.Drawing.Point(112, 72)
        Me.dtDateFactCourante.Name = "dtDateFactCourante"
        Me.dtDateFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.dtDateFactCourante.TabIndex = 2
        '
        'cbAppliquer
        '
        Me.cbAppliquer.Location = New System.Drawing.Point(424, 216)
        Me.cbAppliquer.Name = "cbAppliquer"
        Me.cbAppliquer.Size = New System.Drawing.Size(104, 24)
        Me.cbAppliquer.TabIndex = 7
        Me.cbAppliquer.Text = "Appliquer"
        '
        'tbPeriodeFactCourante
        '
        Me.tbPeriodeFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactCom, "periode", True))
        Me.tbPeriodeFactCourante.Location = New System.Drawing.Point(112, 128)
        Me.tbPeriodeFactCourante.Name = "tbPeriodeFactCourante"
        Me.tbPeriodeFactCourante.Size = New System.Drawing.Size(200, 20)
        Me.tbPeriodeFactCourante.TabIndex = 4
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 131)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(88, 24)
        Me.Label10.TabIndex = 140
        Me.Label10.Text = "Periode"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(8, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(88, 23)
        Me.Label9.TabIndex = 137
        Me.Label9.Text = "Date Stat"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 76)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 136
        Me.Label7.Text = "DateFacture"
        '
        'tbMontantTTCFactCourante
        '
        Me.tbMontantTTCFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactCom, "totalTTC", True))
        Me.tbMontantTTCFactCourante.Location = New System.Drawing.Point(112, 184)
        Me.tbMontantTTCFactCourante.Name = "tbMontantTTCFactCourante"
        Me.tbMontantTTCFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantTTCFactCourante.TabIndex = 6
        Me.tbMontantTTCFactCourante.Text = "0"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(6, 187)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 15)
        Me.Label3.TabIndex = 134
        Me.Label3.Text = "Montant TTC"
        '
        'tbMontantHTFactCourante
        '
        Me.tbMontantHTFactCourante.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactCom, "totalHT", True))
        Me.tbMontantHTFactCourante.Location = New System.Drawing.Point(112, 156)
        Me.tbMontantHTFactCourante.Name = "tbMontantHTFactCourante"
        Me.tbMontantHTFactCourante.Size = New System.Drawing.Size(96, 20)
        Me.tbMontantHTFactCourante.TabIndex = 5
        Me.tbMontantHTFactCourante.Text = "0"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 159)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 132
        Me.Label2.Text = "Montant HT"
        '
        'liFournisseur
        '
        Me.liFournisseur.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactCom, "TiersRS", True))
        Me.liFournisseur.Location = New System.Drawing.Point(8, 48)
        Me.liFournisseur.Name = "liFournisseur"
        Me.liFournisseur.Size = New System.Drawing.Size(520, 16)
        Me.liFournisseur.TabIndex = 1
        Me.liFournisseur.TabStop = True
        Me.liFournisseur.Text = "RS Fournisseur"
        '
        'liFactCom
        '
        Me.liFactCom.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcFactCom, "shortResume", True))
        Me.liFactCom.Location = New System.Drawing.Point(8, 24)
        Me.liFactCom.Name = "liFactCom"
        Me.liFactCom.Size = New System.Drawing.Size(520, 24)
        Me.liFactCom.TabIndex = 0
        Me.liFactCom.TabStop = True
        Me.liFactCom.Text = "Reference Facture de Commission"
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(328, 56)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(75, 23)
        Me.cbRecherche.TabIndex = 3
        Me.cbRecherche.Text = "Recherche"
        '
        'cbSauvegarder
        '
        Me.cbSauvegarder.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSauvegarder.Location = New System.Drawing.Point(616, 320)
        Me.cbSauvegarder.Name = "cbSauvegarder"
        Me.cbSauvegarder.Size = New System.Drawing.Size(80, 23)
        Me.cbSauvegarder.TabIndex = 126
        Me.cbSauvegarder.Text = "Sauvegarder"
        '
        'dgvFactCom
        '
        Me.dgvFactCom.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvFactCom.AutoGenerateColumns = False
        Me.dgvFactCom.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvFactCom.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFactCom.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.code, Me.TiersRS, Me.totalHT, Me.totalTTC, Me.DateReglementDataGridViewTextBoxColumn, Me.DateStatistiqueDataGridViewTextBoxColumn, Me.PeriodeDataGridViewTextBoxColumn, Me.RefReglementDataGridViewTextBoxColumn, Me.MontantReglementDataGridViewTextBoxColumn, Me.ShortResumeDataGridViewTextBoxColumn, Me.BcolReglementLoadedDataGridViewCheckBoxColumn, Me.BExportInternetDataGridViewCheckBoxColumn, Me.CodeDataGridViewTextBoxColumn, Me.DateCommandeDataGridViewTextBoxColumn, Me.DateFactureDataGridViewTextBoxColumn, Me.EtatDataGridViewTextBoxColumn, Me.EtatLibelleDataGridViewTextBoxColumn, Me.OTiersDataGridViewTextBoxColumn, Me.TiersRSDataGridViewTextBoxColumn, Me.ColLignesDataGridViewTextBoxColumn, Me.TotalHTDataGridViewTextBoxColumn, Me.TotalTTCDataGridViewTextBoxColumn, Me.OTransporteurDataGridViewTextBoxColumn, Me.IdParamTVADataGridViewTextBoxColumn, Me.DateValidationDataGridViewTextBoxColumn, Me.DateLivraisonDataGridViewTextBoxColumn, Me.DateEnlevementDataGridViewTextBoxColumn, Me.RefLivraisonDataGridViewTextBoxColumn, Me.CaracteristiqueTiersDataGridViewTextBoxColumn, Me.CommCommandeDataGridViewTextBoxColumn, Me.CommLivraisonDataGridViewTextBoxColumn, Me.CommFacturationDataGridViewTextBoxColumn, Me.CommLibreDataGridViewTextBoxColumn, Me.BcolLignesLoadedDataGridViewCheckBoxColumn, Me.BcolLignesUpatedDataGridViewCheckBoxColumn, Me.QteColisDataGridViewTextBoxColumn, Me.QtePalettesPrepareesDataGridViewTextBoxColumn, Me.QtePalettesNonPrepareesDataGridViewTextBoxColumn, Me.PoidsDataGridViewTextBoxColumn, Me.PuPalettesPrepareesDataGridViewTextBoxColumn, Me.PuPalettesNonPrepareesDataGridViewTextBoxColumn, Me.MontantTransportDataGridViewTextBoxColumn, Me.CoutTransportDataGridViewTextBoxColumn, Me.LettreVoitureDataGridViewTextBoxColumn, Me.RefFactTRPDataGridViewTextBoxColumn, Me.MontantCommissionDataGridViewTextBoxColumn, Me.BNewDataGridViewCheckBoxColumn, Me.BDeletedDataGridViewCheckBoxColumn, Me.BUpdatedDataGridViewCheckBoxColumn, Me.IdDataGridViewTextBoxColumn, Me.TypeDonneeDataGridViewTextBoxColumn, Me.BResumeDataGridViewCheckBoxColumn})
        Me.dgvFactCom.DataSource = Me.m_bsrcFactCom
        Me.dgvFactCom.Location = New System.Drawing.Point(483, 104)
        Me.dgvFactCom.Name = "dgvFactCom"
        Me.dgvFactCom.RowHeadersVisible = False
        Me.dgvFactCom.Size = New System.Drawing.Size(509, 210)
        Me.dgvFactCom.TabIndex = 127
        '
        'code
        '
        Me.code.DataPropertyName = "code"
        Me.code.HeaderText = "code"
        Me.code.Name = "code"
        Me.code.Width = 56
        '
        'TiersRS
        '
        Me.TiersRS.DataPropertyName = "TiersRS"
        Me.TiersRS.HeaderText = "Raison sociale Fourn"
        Me.TiersRS.Name = "TiersRS"
        Me.TiersRS.ReadOnly = True
        Me.TiersRS.Width = 96
        '
        'totalHT
        '
        Me.totalHT.DataPropertyName = "totalHT"
        Me.totalHT.HeaderText = "H.T."
        Me.totalHT.Name = "totalHT"
        Me.totalHT.Width = 53
        '
        'totalTTC
        '
        Me.totalTTC.DataPropertyName = "totalTTC"
        Me.totalTTC.HeaderText = "T.T.C"
        Me.totalTTC.Name = "totalTTC"
        Me.totalTTC.Width = 59
        '
        'DateReglementDataGridViewTextBoxColumn
        '
        Me.DateReglementDataGridViewTextBoxColumn.DataPropertyName = "dateReglement"
        Me.DateReglementDataGridViewTextBoxColumn.HeaderText = "dateReglement"
        Me.DateReglementDataGridViewTextBoxColumn.Name = "DateReglementDataGridViewTextBoxColumn"
        Me.DateReglementDataGridViewTextBoxColumn.Width = 104
        '
        'DateStatistiqueDataGridViewTextBoxColumn
        '
        Me.DateStatistiqueDataGridViewTextBoxColumn.DataPropertyName = "dateStatistique"
        Me.DateStatistiqueDataGridViewTextBoxColumn.HeaderText = "dateStatistique"
        Me.DateStatistiqueDataGridViewTextBoxColumn.Name = "DateStatistiqueDataGridViewTextBoxColumn"
        Me.DateStatistiqueDataGridViewTextBoxColumn.Width = 102
        '
        'PeriodeDataGridViewTextBoxColumn
        '
        Me.PeriodeDataGridViewTextBoxColumn.DataPropertyName = "periode"
        Me.PeriodeDataGridViewTextBoxColumn.HeaderText = "periode"
        Me.PeriodeDataGridViewTextBoxColumn.Name = "PeriodeDataGridViewTextBoxColumn"
        Me.PeriodeDataGridViewTextBoxColumn.Width = 67
        '
        'RefReglementDataGridViewTextBoxColumn
        '
        Me.RefReglementDataGridViewTextBoxColumn.DataPropertyName = "refReglement"
        Me.RefReglementDataGridViewTextBoxColumn.HeaderText = "refReglement"
        Me.RefReglementDataGridViewTextBoxColumn.Name = "RefReglementDataGridViewTextBoxColumn"
        Me.RefReglementDataGridViewTextBoxColumn.Width = 95
        '
        'MontantReglementDataGridViewTextBoxColumn
        '
        Me.MontantReglementDataGridViewTextBoxColumn.DataPropertyName = "montantReglement"
        Me.MontantReglementDataGridViewTextBoxColumn.HeaderText = "montantReglement"
        Me.MontantReglementDataGridViewTextBoxColumn.Name = "MontantReglementDataGridViewTextBoxColumn"
        Me.MontantReglementDataGridViewTextBoxColumn.Width = 121
        '
        'ShortResumeDataGridViewTextBoxColumn
        '
        Me.ShortResumeDataGridViewTextBoxColumn.DataPropertyName = "shortResume"
        Me.ShortResumeDataGridViewTextBoxColumn.HeaderText = "shortResume"
        Me.ShortResumeDataGridViewTextBoxColumn.Name = "ShortResumeDataGridViewTextBoxColumn"
        Me.ShortResumeDataGridViewTextBoxColumn.ReadOnly = True
        Me.ShortResumeDataGridViewTextBoxColumn.Width = 94
        '
        'BcolReglementLoadedDataGridViewCheckBoxColumn
        '
        Me.BcolReglementLoadedDataGridViewCheckBoxColumn.DataPropertyName = "bcolReglementLoaded"
        Me.BcolReglementLoadedDataGridViewCheckBoxColumn.HeaderText = "bcolReglementLoaded"
        Me.BcolReglementLoadedDataGridViewCheckBoxColumn.Name = "BcolReglementLoadedDataGridViewCheckBoxColumn"
        Me.BcolReglementLoadedDataGridViewCheckBoxColumn.Width = 120
        '
        'BExportInternetDataGridViewCheckBoxColumn
        '
        Me.BExportInternetDataGridViewCheckBoxColumn.DataPropertyName = "bExportInternet"
        Me.BExportInternetDataGridViewCheckBoxColumn.HeaderText = "bExportInternet"
        Me.BExportInternetDataGridViewCheckBoxColumn.Name = "BExportInternetDataGridViewCheckBoxColumn"
        Me.BExportInternetDataGridViewCheckBoxColumn.Width = 85
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.Width = 56
        '
        'DateCommandeDataGridViewTextBoxColumn
        '
        Me.DateCommandeDataGridViewTextBoxColumn.DataPropertyName = "dateCommande"
        Me.DateCommandeDataGridViewTextBoxColumn.HeaderText = "dateCommande"
        Me.DateCommandeDataGridViewTextBoxColumn.Name = "DateCommandeDataGridViewTextBoxColumn"
        Me.DateCommandeDataGridViewTextBoxColumn.Width = 106
        '
        'DateFactureDataGridViewTextBoxColumn
        '
        Me.DateFactureDataGridViewTextBoxColumn.DataPropertyName = "dateFacture"
        Me.DateFactureDataGridViewTextBoxColumn.HeaderText = "dateFacture"
        Me.DateFactureDataGridViewTextBoxColumn.Name = "DateFactureDataGridViewTextBoxColumn"
        Me.DateFactureDataGridViewTextBoxColumn.Width = 89
        '
        'EtatDataGridViewTextBoxColumn
        '
        Me.EtatDataGridViewTextBoxColumn.DataPropertyName = "etat"
        Me.EtatDataGridViewTextBoxColumn.HeaderText = "etat"
        Me.EtatDataGridViewTextBoxColumn.Name = "EtatDataGridViewTextBoxColumn"
        Me.EtatDataGridViewTextBoxColumn.Width = 50
        '
        'EtatLibelleDataGridViewTextBoxColumn
        '
        Me.EtatLibelleDataGridViewTextBoxColumn.DataPropertyName = "EtatLibelle"
        Me.EtatLibelleDataGridViewTextBoxColumn.HeaderText = "EtatLibelle"
        Me.EtatLibelleDataGridViewTextBoxColumn.Name = "EtatLibelleDataGridViewTextBoxColumn"
        Me.EtatLibelleDataGridViewTextBoxColumn.ReadOnly = True
        Me.EtatLibelleDataGridViewTextBoxColumn.Width = 81
        '
        'OTiersDataGridViewTextBoxColumn
        '
        Me.OTiersDataGridViewTextBoxColumn.DataPropertyName = "oTiers"
        Me.OTiersDataGridViewTextBoxColumn.HeaderText = "oTiers"
        Me.OTiersDataGridViewTextBoxColumn.Name = "OTiersDataGridViewTextBoxColumn"
        Me.OTiersDataGridViewTextBoxColumn.Width = 61
        '
        'TiersRSDataGridViewTextBoxColumn
        '
        Me.TiersRSDataGridViewTextBoxColumn.DataPropertyName = "TiersRS"
        Me.TiersRSDataGridViewTextBoxColumn.HeaderText = "TiersRS"
        Me.TiersRSDataGridViewTextBoxColumn.Name = "TiersRSDataGridViewTextBoxColumn"
        Me.TiersRSDataGridViewTextBoxColumn.ReadOnly = True
        Me.TiersRSDataGridViewTextBoxColumn.Width = 70
        '
        'ColLignesDataGridViewTextBoxColumn
        '
        Me.ColLignesDataGridViewTextBoxColumn.DataPropertyName = "colLignes"
        Me.ColLignesDataGridViewTextBoxColumn.HeaderText = "colLignes"
        Me.ColLignesDataGridViewTextBoxColumn.Name = "ColLignesDataGridViewTextBoxColumn"
        Me.ColLignesDataGridViewTextBoxColumn.ReadOnly = True
        Me.ColLignesDataGridViewTextBoxColumn.Width = 77
        '
        'TotalHTDataGridViewTextBoxColumn
        '
        Me.TotalHTDataGridViewTextBoxColumn.DataPropertyName = "totalHT"
        Me.TotalHTDataGridViewTextBoxColumn.HeaderText = "totalHT"
        Me.TotalHTDataGridViewTextBoxColumn.Name = "TotalHTDataGridViewTextBoxColumn"
        Me.TotalHTDataGridViewTextBoxColumn.Width = 67
        '
        'TotalTTCDataGridViewTextBoxColumn
        '
        Me.TotalTTCDataGridViewTextBoxColumn.DataPropertyName = "totalTTC"
        Me.TotalTTCDataGridViewTextBoxColumn.HeaderText = "totalTTC"
        Me.TotalTTCDataGridViewTextBoxColumn.Name = "TotalTTCDataGridViewTextBoxColumn"
        Me.TotalTTCDataGridViewTextBoxColumn.Width = 73
        '
        'OTransporteurDataGridViewTextBoxColumn
        '
        Me.OTransporteurDataGridViewTextBoxColumn.DataPropertyName = "oTransporteur"
        Me.OTransporteurDataGridViewTextBoxColumn.HeaderText = "oTransporteur"
        Me.OTransporteurDataGridViewTextBoxColumn.Name = "OTransporteurDataGridViewTextBoxColumn"
        Me.OTransporteurDataGridViewTextBoxColumn.Width = 98
        '
        'IdParamTVADataGridViewTextBoxColumn
        '
        Me.IdParamTVADataGridViewTextBoxColumn.DataPropertyName = "idParamTVA"
        Me.IdParamTVADataGridViewTextBoxColumn.HeaderText = "idParamTVA"
        Me.IdParamTVADataGridViewTextBoxColumn.Name = "IdParamTVADataGridViewTextBoxColumn"
        Me.IdParamTVADataGridViewTextBoxColumn.Width = 91
        '
        'DateValidationDataGridViewTextBoxColumn
        '
        Me.DateValidationDataGridViewTextBoxColumn.DataPropertyName = "dateValidation"
        Me.DateValidationDataGridViewTextBoxColumn.HeaderText = "dateValidation"
        Me.DateValidationDataGridViewTextBoxColumn.Name = "DateValidationDataGridViewTextBoxColumn"
        Me.DateValidationDataGridViewTextBoxColumn.Width = 99
        '
        'DateLivraisonDataGridViewTextBoxColumn
        '
        Me.DateLivraisonDataGridViewTextBoxColumn.DataPropertyName = "dateLivraison"
        Me.DateLivraisonDataGridViewTextBoxColumn.HeaderText = "dateLivraison"
        Me.DateLivraisonDataGridViewTextBoxColumn.Name = "DateLivraisonDataGridViewTextBoxColumn"
        Me.DateLivraisonDataGridViewTextBoxColumn.Width = 95
        '
        'DateEnlevementDataGridViewTextBoxColumn
        '
        Me.DateEnlevementDataGridViewTextBoxColumn.DataPropertyName = "dateEnlevement"
        Me.DateEnlevementDataGridViewTextBoxColumn.HeaderText = "dateEnlevement"
        Me.DateEnlevementDataGridViewTextBoxColumn.Name = "DateEnlevementDataGridViewTextBoxColumn"
        Me.DateEnlevementDataGridViewTextBoxColumn.Width = 109
        '
        'RefLivraisonDataGridViewTextBoxColumn
        '
        Me.RefLivraisonDataGridViewTextBoxColumn.DataPropertyName = "refLivraison"
        Me.RefLivraisonDataGridViewTextBoxColumn.HeaderText = "refLivraison"
        Me.RefLivraisonDataGridViewTextBoxColumn.Name = "RefLivraisonDataGridViewTextBoxColumn"
        Me.RefLivraisonDataGridViewTextBoxColumn.Width = 86
        '
        'CaracteristiqueTiersDataGridViewTextBoxColumn
        '
        Me.CaracteristiqueTiersDataGridViewTextBoxColumn.DataPropertyName = "caracteristiqueTiers"
        Me.CaracteristiqueTiersDataGridViewTextBoxColumn.HeaderText = "caracteristiqueTiers"
        Me.CaracteristiqueTiersDataGridViewTextBoxColumn.Name = "CaracteristiqueTiersDataGridViewTextBoxColumn"
        Me.CaracteristiqueTiersDataGridViewTextBoxColumn.ReadOnly = True
        Me.CaracteristiqueTiersDataGridViewTextBoxColumn.Width = 124
        '
        'CommCommandeDataGridViewTextBoxColumn
        '
        Me.CommCommandeDataGridViewTextBoxColumn.DataPropertyName = "CommCommande"
        Me.CommCommandeDataGridViewTextBoxColumn.HeaderText = "CommCommande"
        Me.CommCommandeDataGridViewTextBoxColumn.Name = "CommCommandeDataGridViewTextBoxColumn"
        Me.CommCommandeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CommCommandeDataGridViewTextBoxColumn.Width = 114
        '
        'CommLivraisonDataGridViewTextBoxColumn
        '
        Me.CommLivraisonDataGridViewTextBoxColumn.DataPropertyName = "CommLivraison"
        Me.CommLivraisonDataGridViewTextBoxColumn.HeaderText = "CommLivraison"
        Me.CommLivraisonDataGridViewTextBoxColumn.Name = "CommLivraisonDataGridViewTextBoxColumn"
        Me.CommLivraisonDataGridViewTextBoxColumn.ReadOnly = True
        Me.CommLivraisonDataGridViewTextBoxColumn.Width = 103
        '
        'CommFacturationDataGridViewTextBoxColumn
        '
        Me.CommFacturationDataGridViewTextBoxColumn.DataPropertyName = "CommFacturation"
        Me.CommFacturationDataGridViewTextBoxColumn.HeaderText = "CommFacturation"
        Me.CommFacturationDataGridViewTextBoxColumn.Name = "CommFacturationDataGridViewTextBoxColumn"
        Me.CommFacturationDataGridViewTextBoxColumn.ReadOnly = True
        Me.CommFacturationDataGridViewTextBoxColumn.Width = 114
        '
        'CommLibreDataGridViewTextBoxColumn
        '
        Me.CommLibreDataGridViewTextBoxColumn.DataPropertyName = "CommLibre"
        Me.CommLibreDataGridViewTextBoxColumn.HeaderText = "CommLibre"
        Me.CommLibreDataGridViewTextBoxColumn.Name = "CommLibreDataGridViewTextBoxColumn"
        Me.CommLibreDataGridViewTextBoxColumn.ReadOnly = True
        Me.CommLibreDataGridViewTextBoxColumn.Width = 84
        '
        'BcolLignesLoadedDataGridViewCheckBoxColumn
        '
        Me.BcolLignesLoadedDataGridViewCheckBoxColumn.DataPropertyName = "bcolLignesLoaded"
        Me.BcolLignesLoadedDataGridViewCheckBoxColumn.HeaderText = "bcolLignesLoaded"
        Me.BcolLignesLoadedDataGridViewCheckBoxColumn.Name = "BcolLignesLoadedDataGridViewCheckBoxColumn"
        Me.BcolLignesLoadedDataGridViewCheckBoxColumn.ReadOnly = True
        '
        'BcolLignesUpatedDataGridViewCheckBoxColumn
        '
        Me.BcolLignesUpatedDataGridViewCheckBoxColumn.DataPropertyName = "bcolLignesUpated"
        Me.BcolLignesUpatedDataGridViewCheckBoxColumn.HeaderText = "bcolLignesUpated"
        Me.BcolLignesUpatedDataGridViewCheckBoxColumn.Name = "BcolLignesUpatedDataGridViewCheckBoxColumn"
        Me.BcolLignesUpatedDataGridViewCheckBoxColumn.ReadOnly = True
        Me.BcolLignesUpatedDataGridViewCheckBoxColumn.Width = 99
        '
        'QteColisDataGridViewTextBoxColumn
        '
        Me.QteColisDataGridViewTextBoxColumn.DataPropertyName = "qteColis"
        Me.QteColisDataGridViewTextBoxColumn.HeaderText = "qteColis"
        Me.QteColisDataGridViewTextBoxColumn.Name = "QteColisDataGridViewTextBoxColumn"
        Me.QteColisDataGridViewTextBoxColumn.Width = 69
        '
        'QtePalettesPrepareesDataGridViewTextBoxColumn
        '
        Me.QtePalettesPrepareesDataGridViewTextBoxColumn.DataPropertyName = "qtePalettesPreparees"
        Me.QtePalettesPrepareesDataGridViewTextBoxColumn.HeaderText = "qtePalettesPreparees"
        Me.QtePalettesPrepareesDataGridViewTextBoxColumn.Name = "QtePalettesPrepareesDataGridViewTextBoxColumn"
        Me.QtePalettesPrepareesDataGridViewTextBoxColumn.Width = 133
        '
        'QtePalettesNonPrepareesDataGridViewTextBoxColumn
        '
        Me.QtePalettesNonPrepareesDataGridViewTextBoxColumn.DataPropertyName = "qtePalettesNonPreparees"
        Me.QtePalettesNonPrepareesDataGridViewTextBoxColumn.HeaderText = "qtePalettesNonPreparees"
        Me.QtePalettesNonPrepareesDataGridViewTextBoxColumn.Name = "QtePalettesNonPrepareesDataGridViewTextBoxColumn"
        Me.QtePalettesNonPrepareesDataGridViewTextBoxColumn.Width = 153
        '
        'PoidsDataGridViewTextBoxColumn
        '
        Me.PoidsDataGridViewTextBoxColumn.DataPropertyName = "poids"
        Me.PoidsDataGridViewTextBoxColumn.HeaderText = "poids"
        Me.PoidsDataGridViewTextBoxColumn.Name = "PoidsDataGridViewTextBoxColumn"
        Me.PoidsDataGridViewTextBoxColumn.Width = 57
        '
        'PuPalettesPrepareesDataGridViewTextBoxColumn
        '
        Me.PuPalettesPrepareesDataGridViewTextBoxColumn.DataPropertyName = "puPalettesPreparees"
        Me.PuPalettesPrepareesDataGridViewTextBoxColumn.HeaderText = "puPalettesPreparees"
        Me.PuPalettesPrepareesDataGridViewTextBoxColumn.Name = "PuPalettesPrepareesDataGridViewTextBoxColumn"
        Me.PuPalettesPrepareesDataGridViewTextBoxColumn.Width = 130
        '
        'PuPalettesNonPrepareesDataGridViewTextBoxColumn
        '
        Me.PuPalettesNonPrepareesDataGridViewTextBoxColumn.DataPropertyName = "puPalettesNonPreparees"
        Me.PuPalettesNonPrepareesDataGridViewTextBoxColumn.HeaderText = "puPalettesNonPreparees"
        Me.PuPalettesNonPrepareesDataGridViewTextBoxColumn.Name = "PuPalettesNonPrepareesDataGridViewTextBoxColumn"
        Me.PuPalettesNonPrepareesDataGridViewTextBoxColumn.Width = 150
        '
        'MontantTransportDataGridViewTextBoxColumn
        '
        Me.MontantTransportDataGridViewTextBoxColumn.DataPropertyName = "montantTransport"
        Me.MontantTransportDataGridViewTextBoxColumn.HeaderText = "montantTransport"
        Me.MontantTransportDataGridViewTextBoxColumn.Name = "MontantTransportDataGridViewTextBoxColumn"
        Me.MontantTransportDataGridViewTextBoxColumn.Width = 115
        '
        'CoutTransportDataGridViewTextBoxColumn
        '
        Me.CoutTransportDataGridViewTextBoxColumn.DataPropertyName = "coutTransport"
        Me.CoutTransportDataGridViewTextBoxColumn.HeaderText = "coutTransport"
        Me.CoutTransportDataGridViewTextBoxColumn.Name = "CoutTransportDataGridViewTextBoxColumn"
        Me.CoutTransportDataGridViewTextBoxColumn.Width = 98
        '
        'LettreVoitureDataGridViewTextBoxColumn
        '
        Me.LettreVoitureDataGridViewTextBoxColumn.DataPropertyName = "lettreVoiture"
        Me.LettreVoitureDataGridViewTextBoxColumn.HeaderText = "lettreVoiture"
        Me.LettreVoitureDataGridViewTextBoxColumn.Name = "LettreVoitureDataGridViewTextBoxColumn"
        Me.LettreVoitureDataGridViewTextBoxColumn.Width = 88
        '
        'RefFactTRPDataGridViewTextBoxColumn
        '
        Me.RefFactTRPDataGridViewTextBoxColumn.DataPropertyName = "refFactTRP"
        Me.RefFactTRPDataGridViewTextBoxColumn.HeaderText = "refFactTRP"
        Me.RefFactTRPDataGridViewTextBoxColumn.Name = "RefFactTRPDataGridViewTextBoxColumn"
        Me.RefFactTRPDataGridViewTextBoxColumn.Width = 87
        '
        'MontantCommissionDataGridViewTextBoxColumn
        '
        Me.MontantCommissionDataGridViewTextBoxColumn.DataPropertyName = "montantCommission"
        Me.MontantCommissionDataGridViewTextBoxColumn.HeaderText = "montantCommission"
        Me.MontantCommissionDataGridViewTextBoxColumn.Name = "MontantCommissionDataGridViewTextBoxColumn"
        Me.MontantCommissionDataGridViewTextBoxColumn.ReadOnly = True
        Me.MontantCommissionDataGridViewTextBoxColumn.Width = 125
        '
        'BNewDataGridViewCheckBoxColumn
        '
        Me.BNewDataGridViewCheckBoxColumn.DataPropertyName = "bNew"
        Me.BNewDataGridViewCheckBoxColumn.HeaderText = "bNew"
        Me.BNewDataGridViewCheckBoxColumn.Name = "BNewDataGridViewCheckBoxColumn"
        Me.BNewDataGridViewCheckBoxColumn.Width = 41
        '
        'BDeletedDataGridViewCheckBoxColumn
        '
        Me.BDeletedDataGridViewCheckBoxColumn.DataPropertyName = "bDeleted"
        Me.BDeletedDataGridViewCheckBoxColumn.HeaderText = "bDeleted"
        Me.BDeletedDataGridViewCheckBoxColumn.Name = "BDeletedDataGridViewCheckBoxColumn"
        Me.BDeletedDataGridViewCheckBoxColumn.Width = 56
        '
        'BUpdatedDataGridViewCheckBoxColumn
        '
        Me.BUpdatedDataGridViewCheckBoxColumn.DataPropertyName = "bUpdated"
        Me.BUpdatedDataGridViewCheckBoxColumn.HeaderText = "bUpdated"
        Me.BUpdatedDataGridViewCheckBoxColumn.Name = "BUpdatedDataGridViewCheckBoxColumn"
        Me.BUpdatedDataGridViewCheckBoxColumn.Width = 60
        '
        'IdDataGridViewTextBoxColumn
        '
        Me.IdDataGridViewTextBoxColumn.DataPropertyName = "id"
        Me.IdDataGridViewTextBoxColumn.HeaderText = "id"
        Me.IdDataGridViewTextBoxColumn.Name = "IdDataGridViewTextBoxColumn"
        Me.IdDataGridViewTextBoxColumn.ReadOnly = True
        Me.IdDataGridViewTextBoxColumn.Width = 40
        '
        'TypeDonneeDataGridViewTextBoxColumn
        '
        Me.TypeDonneeDataGridViewTextBoxColumn.DataPropertyName = "typeDonnee"
        Me.TypeDonneeDataGridViewTextBoxColumn.HeaderText = "typeDonnee"
        Me.TypeDonneeDataGridViewTextBoxColumn.Name = "TypeDonneeDataGridViewTextBoxColumn"
        Me.TypeDonneeDataGridViewTextBoxColumn.ReadOnly = True
        Me.TypeDonneeDataGridViewTextBoxColumn.Width = 90
        '
        'BResumeDataGridViewCheckBoxColumn
        '
        Me.BResumeDataGridViewCheckBoxColumn.DataPropertyName = "bResume"
        Me.BResumeDataGridViewCheckBoxColumn.HeaderText = "bResume"
        Me.BResumeDataGridViewCheckBoxColumn.Name = "BResumeDataGridViewCheckBoxColumn"
        Me.BResumeDataGridViewCheckBoxColumn.ReadOnly = True
        Me.BResumeDataGridViewCheckBoxColumn.Width = 58
        '
        'm_bsrcSousCommande
        '
        Me.m_bsrcSousCommande.DataSource = GetType(vini_DB.SousCommande)
        '
        'dgvSousCommandes
        '
        Me.dgvSousCommandes.AllowUserToAddRows = False
        Me.dgvSousCommandes.AllowUserToDeleteRows = False
        Me.dgvSousCommandes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvSousCommandes.AutoGenerateColumns = False
        Me.dgvSousCommandes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSousCommandes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSousCommandes.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn1, Me.FournisseurCode, Me.TiersRSDataGridViewTextBoxColumn1, Me.DateCommandeDataGridViewTextBoxColumn1, Me.totalHTFacture})
        Me.dgvSousCommandes.DataSource = Me.m_bsrcSousCommande
        Me.dgvSousCommandes.Location = New System.Drawing.Point(10, 114)
        Me.dgvSousCommandes.Name = "dgvSousCommandes"
        Me.dgvSousCommandes.ReadOnly = True
        Me.dgvSousCommandes.RowHeadersVisible = False
        Me.dgvSousCommandes.Size = New System.Drawing.Size(352, 486)
        Me.dgvSousCommandes.TabIndex = 128
        '
        'CodeDataGridViewTextBoxColumn1
        '
        Me.CodeDataGridViewTextBoxColumn1.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn1.HeaderText = "code"
        Me.CodeDataGridViewTextBoxColumn1.Name = "CodeDataGridViewTextBoxColumn1"
        Me.CodeDataGridViewTextBoxColumn1.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn1.Width = 56
        '
        'FournisseurCode
        '
        Me.FournisseurCode.DataPropertyName = "FournisseurCode"
        Me.FournisseurCode.HeaderText = "FRN"
        Me.FournisseurCode.Name = "FournisseurCode"
        Me.FournisseurCode.ReadOnly = True
        Me.FournisseurCode.Width = 54
        '
        'TiersRSDataGridViewTextBoxColumn1
        '
        Me.TiersRSDataGridViewTextBoxColumn1.DataPropertyName = "TiersRS"
        Me.TiersRSDataGridViewTextBoxColumn1.HeaderText = "Client"
        Me.TiersRSDataGridViewTextBoxColumn1.Name = "TiersRSDataGridViewTextBoxColumn1"
        Me.TiersRSDataGridViewTextBoxColumn1.ReadOnly = True
        Me.TiersRSDataGridViewTextBoxColumn1.Width = 58
        '
        'DateCommandeDataGridViewTextBoxColumn1
        '
        Me.DateCommandeDataGridViewTextBoxColumn1.DataPropertyName = "dateCommande"
        Me.DateCommandeDataGridViewTextBoxColumn1.HeaderText = "date CMD"
        Me.DateCommandeDataGridViewTextBoxColumn1.Name = "DateCommandeDataGridViewTextBoxColumn1"
        Me.DateCommandeDataGridViewTextBoxColumn1.ReadOnly = True
        Me.DateCommandeDataGridViewTextBoxColumn1.Width = 80
        '
        'totalHTFacture
        '
        Me.totalHTFacture.DataPropertyName = "totalHTFacture"
        Me.totalHTFacture.HeaderText = "HT Fact"
        Me.totalHTFacture.Name = "totalHTFacture"
        Me.totalHTFacture.ReadOnly = True
        Me.totalHTFacture.Width = 71
        '
        'frmGeneFactCom
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(997, 630)
        Me.Controls.Add(Me.dgvSousCommandes)
        Me.Controls.Add(Me.dgvFactCom)
        Me.Controls.Add(Me.cbSauvegarder)
        Me.Controls.Add(Me.cbRecherche)
        Me.Controls.Add(Me.grpFact)
        Me.Controls.Add(Me.tbPeriode)
        Me.Controls.Add(Me.tbCodeFournisseur)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dtDateStatistique)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dtDateFacture)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbAnnGenerer)
        Me.Controls.Add(Me.cbGenerer)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtdateFin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtDatedeb)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cbAfficher)
        Me.Name = "frmGeneFactCom"
        Me.Text = "Génération de factures de commissions"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grpFact.ResumeLayout(False)
        Me.grpFact.PerformLayout()
        CType(Me.m_bsrcFactCom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvFactCom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcSousCommande, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSousCommandes, System.ComponentModel.ISupportInitialize).EndInit()
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
        m_ToolBarSaveEnabled = False
    End Sub

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        Dim objControl As Control
        MyBase.EnableControls(bEnabled)
        dtDatedeb.Enabled = True
        dtdateFin.Enabled = True
        tbCodeFournisseur.Enabled = True
        cbAfficher.Enabled = True
        dtDateFacture.Enabled = True
        dtDateStatistique.Enabled = True
        tbPeriode.Enabled = True
        cbRecherche.Enabled = True
        For Each objControl In grpFact.Controls
            objControl.Enabled = True
        Next objControl

    End Sub
    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenetre n'est pas libre")
        Dim bReturn As Boolean
        bReturn = setElementCourant2(New FactCom(New Client("", "")))

        Return bReturn
    End Function
    Protected Shadows Function getElementCourant() As FactCom
        Return CType(m_ElementCourant, FactCom)
    End Function


#End Region

#Region "Methodes privées"
    Private Sub initFenetre()
        dtDatedeb.Value = DateAdd(DateInterval.Year, -1, Now())
        dtdateFin.Value = Now()
        m_colSousCommandes = New Collection
        'm_SCMDcourante = Nothing
        'tbTxCommision.Text = Param.getConstante("CST_TX_COMMISSION")
    End Sub
    Private Sub afficheListeSousCommande()

        setcursorWait()
        debAffiche()
        m_bsrcSousCommande.Clear()

        Debug.Assert(Not m_colSousCommandes Is Nothing)
        Dim obj As SousCommande

        For Each obj In m_colSousCommandes
            m_bsrcSousCommande.Add(obj)
        Next obj

        cbGenerer.Enabled = True

        finAffiche()
        restoreCursor()

    End Sub 'AfficheListeSousCommande


    Private Function setListeSousCommandes() As Boolean
        Dim ddeb As Date
        Dim dfin As Date
        Dim codeFourn As String
        Dim col As Collection
        Dim bReturn As Boolean
        debAffiche()
        setcursorWait()
        Try

            ddeb = dtDatedeb.Value.ToShortDateString
            dfin = dtdateFin.Value.ToShortDateString
            codeFourn = tbCodeFournisseur.Text
            'Recupération de la liste des sous Commande A Facturer
            col = SousCommande.getListeAFacturer(ddeb, dfin, codeFourn)
            If col Is Nothing Then
                bReturn = False
            Else
                m_colSousCommandes = col
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
    'Private Function selectionnerTout()
    '    '===============================================================
    '    ' Function : selectionnerTout
    '    'Description : met à jour tous les items de la checkListBox à Checked
    '    '===============================================================
    '    Dim i As Integer


    '    For i = 0 To lbSousCommandes.Items.Count - 1
    '        lbSousCommandes.SetItemCheckState(i, CheckState.Checked)
    '    Next i

    '    If lbSousCommandes.CheckedItems.Count > 0 Then
    '        cbGenerer.Enabled = True
    '    Else
    '        cbGenerer.Enabled = False
    '    End If
    'End Function 'SelectionnerTout
    Private Sub genererFactures()
        '===============================================================
        ' Function : genererFactures
        'Description : Génere les facures de com à partir des sous commandes sélectionnées
        '               une fois les factures générées elles sont affichées
        '===============================================================
        '        Dim colSCMD As Collection
        Dim colFact As ColEvent

        Me.Cursor = Cursors.WaitCursor


        colFact = FactCom.createFactComs(m_colSousCommandes, dtDateFacture.Value, dtDateStatistique.Value, tbPeriode.Text)


        If Not colFact Is Nothing Then

            m_colFact = colFact
            afficheListeFactures()

            afficheListeSousCommande()
            setElementCourant2(Nothing)
        End If

        Me.Cursor = Cursors.Default

    End Sub 'genererFactures
    Private Function annuleGeneration() As Boolean
        '===============================================================
        ' Function : annuleGeneration
        'Description : Annule la génération de factures préalablement générées
        '               Si la Facture de com a été sauvegardé, elle est supprimée 
        '               Netoyage de la collection et de la liste Box
        '               Reaffichage de la liste des sous-commandes
        '===============================================================
        Dim objFactcom As FactCom = Nothing
        Dim bReturn As Boolean

        setcursorWait()
        bReturn = True
        For Each objFactcom In m_colFact
            If objFactcom.id <> 0 Then
                'La Facture a été Sauvegardée => Destruction
                objFactcom.bDeleted = True
                bReturn = bReturn And objFactcom.Save()
                Debug.Assert(bReturn, FactCom.getErreur())
            End If
        Next

        m_colFact = New ColEvent
        If setListeSousCommandes() Then
            afficheListeSousCommande()
        End If
        restoreCursor()

        Return bReturn
    End Function 'annuleGeneration

    Private Sub afficheListeFactures()
        Debug.Assert(Not m_colFact Is Nothing, "pColFact = Nothing")
        Dim objFactCom As FactCom


        'debAffiche()
        'lbFactCom.ClearSelected()
        'lbFactCom.Items.Clear()
        'lbFactCom.ValueMember = "Code"
        'lbFactCom.DisplayMember = "ShortResume"
        'For Each objFactCom In pcolFact
        '    lbFactCom.Items.Add(objFactCom)
        'Next objFactCom
        'lbFactCom.Enabled = True
        'finAffiche()

        debAffiche()
        setcursorWait()

        m_bsrcFactCom.Clear()
        For Each objFactCom In m_colFact
            m_bsrcFactCom.Add(objFactCom)
        Next objFactCom

        restoreCursor()
        finAffiche()

    End Sub
    Private Function appliqueModifications() As Boolean
        Dim objFactCom As FactCom
        Dim bReturn As Boolean
        Try
            objFactCom = Nothing
            If m_bsrcFactCom.Current IsNot Nothing Then
                objFactCom = CType(m_bsrcFactCom.Current, FactCom)
                objFactCom.Save()
            End If
            bReturn = True

        Catch ex As Exception
            Debug.Assert(False, ex.Message)
            bReturn = False
        End Try
        Return bReturn

    End Function 'appliqueModifications

    Private Function selectionneFactCom() As Boolean
        Dim bReturn As Boolean
        '===================================================================================================================================================
        ' Function : selectionnFactCom
        ' Description : Place la Facture selectionné en tant que facture courante
        '===================================================================================================================================================

        If Not elementCourantSansModif() Then
            If MsgBox("La Facture courante n'a pas été sauvegardée, voulez-vous conservez vos modifications", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                'Il faut mémorise l'index sélectionné car la maéthode appliquerModifications reaffiche la liste des souscommande et perd donc le selected index
                appliqueModifications()
                'Reselection de l'élement
            End If
        End If

        bReturn = setElementCourant2(m_bsrcFactCom.Current)
        Return bReturn
    End Function 'selectionneFactCom

    Private Function afficheFactureCourante() As Boolean

        Debug.Assert(Not getElementCourant() Is Nothing, "La Facture courante doit être renseignée")

        'Affichage de la Facture
        debAffiche()
        liFactCom.Text = getElementCourant().shortResume
        liFactCom.Tag = getElementCourant().id
        If getElementCourant().id = 0 Then
            liFactCom.Enabled = False
        Else
            liFactCom.Enabled = True
        End If

        liFournisseur.Text = getElementCourant().oTiers.rs
        liFournisseur.Tag = getElementCourant().oTiers.id

        dtDateFactCourante.Value = getElementCourant().dateCommande
        dtDateStatCourante.Value = getElementCourant().dateStatistique
        tbPeriodeFactCourante.Text = getElementCourant().periode
        tbMontantHTFactCourante.Text = getElementCourant().totalHT
        tbMontantTTCFactCourante.Text = getElementCourant().totalTTC

        finAffiche()

    End Function 'AfficheSousCommande

    Private Function elementCourantSansModif() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
            If Not getElementCourant() Is Nothing Then
                bReturn = bReturn And dtDateFactCourante.Value = getElementCourant().dateCommande
                bReturn = bReturn And dtDateStatCourante.Value = getElementCourant().dateStatistique
                bReturn = bReturn And tbPeriodeFactCourante.Text = getElementCourant().periode
                bReturn = bReturn And tbMontantHTFactCourante.Text = getElementCourant().totalHT
                bReturn = bReturn And tbMontantTTCFactCourante.Text = getElementCourant().totalTTC
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Sub rechercheFournisseur()
        Dim objTiers As Tiers

        objTiers = rechercheDonnee(vncEnums.vncTypeDonnee.FOURNISSEUR, tbCodeFournisseur)

        If Not objTiers Is Nothing Then
            tbCodeFournisseur.Text = objTiers.code
        End If
    End Sub 'RechercheFournisseur

    Private Function sauvegarderFactures() As Boolean
        Dim objFactCom As FactCom
        Dim bReturn As Boolean

        bReturn = False
        If Not m_colFact Is Nothing Then
            bReturn = True
            For Each objFactCom In m_colFact
                bReturn = bReturn And objFactCom.Save()
                Debug.Assert(bReturn, FactCom.getErreur())
            Next
        End If

        Return bReturn
    End Function 'sauvegarderFactures
#End Region
#Region "Gestion des Evenements"
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        sauvegardeElementCourant()
        If setListeSousCommandes() Then
            afficheListeSousCommande()
        End If
        dgvSousCommandes.Enabled = True

    End Sub

#End Region

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Génerer déclenche L'evenement Updated
        'AddHandler cbAppliquer.Click, AddressOf ControlUpdated
        'AddHandler cbGenerer.Click, AddressOf ControlUpdated
    End Sub


    Private Sub cbGenerer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbGenerer.Click
        Call genererFactures()
        If dgvFactCom.RowCount > 0 Then
            dgvFactCom.Enabled = True
            cbSauvegarder.Enabled = True
            cbAnnGenerer.Enabled = True
            cbGenerer.Enabled = False
        End If
    End Sub





    Private Sub cbAppliquer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAppliquer.Click
        appliqueModifications()
    End Sub

    Private Sub liFournisseur_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liFournisseur.LinkClicked
        afficheFenetreFournisseur(liFournisseur.Tag)
    End Sub

    Private Sub cbAnnGenerer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnGenerer.Click
        annuleGeneration()

        dgvFactCom.Enabled = False
        cbSauvegarder.Enabled = False
        cbAnnGenerer.Enabled = False
        cbGenerer.Enabled = True
    End Sub

    Private Sub liFactCom_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liFactCom.LinkClicked
        Dim objFactcom As FactCom
        objFactcom = Nothing
        If m_bsrcFactCom.Current IsNot Nothing Then
            objFactcom = CType(m_bsrcFactCom.Current, FactCom)
            If Not objFactcom.bNew Then
                afficheFenetreFactCom(objFactcom.id)
            End If
        End If
    End Sub

    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRecherche.Click
        rechercheFournisseur()
    End Sub

    Private Sub flxFactCom_DblClick(ByVal sender As Object, ByVal e As System.EventArgs)
        If selectionneFactCom() Then
            afficheFactureCourante()
        End If

    End Sub

    Private Sub cbSauvegarder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSauvegarder.Click
        If sauvegarderFactures() Then
            afficheListeFactures()
            grpFact.Enabled = True
        End If
    End Sub
End Class
