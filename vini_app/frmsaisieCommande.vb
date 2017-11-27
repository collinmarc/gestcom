Imports vini_DB
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
'Imports FAXCOMLib
Imports System.Windows.Forms.Cursors
Imports CrystalDecisions.Windows.Forms

Public Class frmSaisieCommande
    Inherits FrmDonBase
    Private Enum vncColLigneCommande
        COL_NUM = 0
        COL_CODEPRODUIT = 1
        COL_DESIGNATIONPRODUIT = 2
        COL_MILLESIME = 3
        COL_CONDITIONNEMENT = 4
        COL_COULEUR = 5
        COL_QTE_COM = 6
        COL_QTE_LIV = 7
        COL_QTE_FACT = 8
        COL_GRATUIT = 9
        COL_PRIX_U = 10
        COL_PRIX_HT = 11
        COL_PRIX_TTC = 12
        COL_IDPRODUIT = 15
        COL_NBRECOL = 16
    End Enum
    '    Private getCommandeCourante As Commande
    Private m_Tiers_Courant As Tiers

    Protected Function getCommandeCourante() As Commande
        Return CType(getElementCourant(), Commande)
    End Function

#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()
        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()


    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If

            DisposeCr(crwDetailCommandeClient)
            DisposeCr(crwBL)
            DisposeCr(crwFact)

        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents cbSupprimerLigne As System.Windows.Forms.Button
    Public WithEvents cbAjouterLigne As System.Windows.Forms.Button
    Public WithEvents cbModifierLigne As System.Windows.Forms.Button
    Public WithEvents tbTotalHT As textBoxCurrency
    Public WithEvents Label26 As System.Windows.Forms.Label
    Public WithEvents Label25 As System.Windows.Forms.Label
    Public WithEvents Label21 As System.Windows.Forms.Label
    Public WithEvents tpClient As System.Windows.Forms.TabPage
    Public WithEvents tpLignes As System.Windows.Forms.TabPage
    Public WithEvents tpCommentaires As System.Windows.Forms.TabPage
    Friend WithEvents tpValidCmd As System.Windows.Forms.TabPage
    Friend WithEvents tpRetourBL As System.Windows.Forms.TabPage
    Friend WithEvents laEtatCommande As System.Windows.Forms.Label
    Public WithEvents Label20 As System.Windows.Forms.Label
    Public WithEvents Label19 As System.Windows.Forms.Label
    Public WithEvents tbRib1 As textBoxStrNumeric
    Public WithEvents tbRib2 As textBoxStrNumeric
    Public WithEvents tbRib3 As textBoxStrNumeric
    Public WithEvents tbRib4 As textBoxStrNumeric
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents label As System.Windows.Forms.Label
    Friend WithEvents tbRaisonSociale As System.Windows.Forms.TextBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Labelbis As System.Windows.Forms.Label
    Public WithEvents Label18 As System.Windows.Forms.Label
    Public WithEvents Label17 As System.Windows.Forms.Label
    Public WithEvents Label16 As System.Windows.Forms.Label
    Public WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents cbFax As System.Windows.Forms.Button
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Public WithEvents tbCodeClient As System.Windows.Forms.TextBox
    Public WithEvents cbRechercheclient As System.Windows.Forms.Button
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents dtDateCommande As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents tbBanque As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents tpBL As System.Windows.Forms.TabPage
    Friend WithEvents tbNumFaxBC As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents tbRefBL As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents tbCode As System.Windows.Forms.TextBox
    Public WithEvents tbNomClient As System.Windows.Forms.TextBox
    Friend WithEvents tbAdrFact_Nom As System.Windows.Forms.TextBox
    Public WithEvents tbAdrFact_Rue1 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrFact_Rue2 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrFact_cp As textBoxStrNumeric
    Friend WithEvents tbAdrLiv_Nom As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLiv_Rue1 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLiv_Rue2 As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLiv_cp As textBoxStrNumeric
    Public WithEvents tbAdrLiv_Ville As System.Windows.Forms.TextBox
    Public WithEvents tbAdrLiv_Port As textBoxStrNumeric
    Public WithEvents tbAdrLiv_Tel As textBoxStrNumeric
    Public WithEvents tbAdrLiv_Fax As textBoxStrNumeric
    Public WithEvents tbAdrLiv_Email As System.Windows.Forms.TextBox
    Public WithEvents tbAdrFact_Ville As System.Windows.Forms.TextBox
    Public WithEvents tbAdrFact_Port As textBoxStrNumeric
    Public WithEvents tbAdrFact_Tel As textBoxStrNumeric
    Public WithEvents tbAdrFact_Fax As textBoxStrNumeric
    Public WithEvents tbAdrFact_Email As System.Windows.Forms.TextBox
    Friend WithEvents cboModeReglement As System.Windows.Forms.ComboBox
    Friend WithEvents laIdClient As System.Windows.Forms.Label
    Friend WithEvents liNomClient As System.Windows.Forms.LinkLabel
    Public WithEvents SSTabCommandeClient As System.Windows.Forms.TabControl
    Friend WithEvents rbTypeCmdPlateforme As System.Windows.Forms.RadioButton
    Friend WithEvents rb_TypeTRP_Depart As System.Windows.Forms.RadioButton
    Friend WithEvents rb_TypeTRP_Avance As System.Windows.Forms.RadioButton
    Friend WithEvents rb_TypeTrp_Franco As System.Windows.Forms.RadioButton
    Friend WithEvents tbCommentaireCommande As System.Windows.Forms.RichTextBox
    Friend WithEvents tbCommentaireLivraison As System.Windows.Forms.RichTextBox
    Friend WithEvents tbCommentaireFacturation As System.Windows.Forms.RichTextBox
    Friend WithEvents ckAdrIdentiques As System.Windows.Forms.CheckBox
    Friend WithEvents grpAdrLiv As System.Windows.Forms.GroupBox
    Friend WithEvents grpAdrFact As System.Windows.Forms.GroupBox
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Public WithEvents tbTotalTTC As textBoxCurrency
    Friend WithEvents tbNumFaxValidation As System.Windows.Forms.TextBox
    Friend WithEvents tbComValid As System.Windows.Forms.RichTextBox
    Friend WithEvents rbTypeCmdDirecte As System.Windows.Forms.RadioButton
    Friend WithEvents dtDateValidation As System.Windows.Forms.DateTimePicker
    Friend WithEvents ckCmdValide As System.Windows.Forms.CheckBox
    Friend WithEvents cbVisuDetailCommande As System.Windows.Forms.Button
    Friend WithEvents tbRecalcTotaux As System.Windows.Forms.Button
    Friend WithEvents grpTransporteur As System.Windows.Forms.GroupBox
    Friend WithEvents cbVisuBL As System.Windows.Forms.Button
    Friend WithEvents dtDateLivraison As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtDateEnlev As System.Windows.Forms.DateTimePicker
    Friend WithEvents tbTrpEmail As System.Windows.Forms.TextBox
    Friend WithEvents tbTrpFax As System.Windows.Forms.TextBox
    Friend WithEvents tbTrpPort As System.Windows.Forms.TextBox
    Friend WithEvents tbTrpTel As System.Windows.Forms.TextBox
    Friend WithEvents tbTrpVille As System.Windows.Forms.TextBox
    Friend WithEvents tbTrpCP As System.Windows.Forms.TextBox
    Friend WithEvents tbTrpRue2 As System.Windows.Forms.TextBox
    Friend WithEvents tbTrpRue1 As System.Windows.Forms.TextBox
    Friend WithEvents tbTrpNom As System.Windows.Forms.TextBox
    Friend WithEvents dtDateLivraisonReelle As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents dtDateEnlevementReelle As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbBLToutOK As System.Windows.Forms.Button
    Friend WithEvents cbAnnulerLivraison As System.Windows.Forms.Button
    Friend WithEvents cbEclatementCmde As System.Windows.Forms.Button
    Friend WithEvents tbSCMDTransporteurNom As System.Windows.Forms.TextBox
    Friend WithEvents tbSCMDCommentaire As System.Windows.Forms.RichTextBox
    Friend WithEvents liSCMDFournisseur As System.Windows.Forms.LinkLabel
    Friend WithEvents label50 As System.Windows.Forms.Label
    Friend WithEvents dtDateLivraisonPrevue As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtSCMDDateLiv As System.Windows.Forms.DateTimePicker
    Friend WithEvents tpEclatement As System.Windows.Forms.TabPage
    Friend WithEvents cbSCMDVoir As System.Windows.Forms.Button
    Friend WithEvents cbSCMDFaxerTout As System.Windows.Forms.Button
    Friend WithEvents cbAnnEclatement As System.Windows.Forms.Button
    Friend WithEvents grpTypeTransport As System.Windows.Forms.GroupBox
    Friend WithEvents grpTypeCommande As System.Windows.Forms.GroupBox
    Friend WithEvents grpEntete As System.Windows.Forms.GroupBox
    Public WithEvents LaCodeTiers As System.Windows.Forms.Label
    Friend WithEvents cbSCMDAppliquer As System.Windows.Forms.Button
    Friend WithEvents ckTransport As System.Windows.Forms.CheckBox
    Friend WithEvents liFactTRP As System.Windows.Forms.LinkLabel
    Friend WithEvents tbPiedPageBL As System.Windows.Forms.RichTextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents grpFactTRP As System.Windows.Forms.GroupBox
    Friend WithEvents grpInfFactureTRP As System.Windows.Forms.GroupBox
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents cbCalcMontantTransport As System.Windows.Forms.Button
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents tbMontantTransport As textBoxCurrency
    Friend WithEvents Label56 As System.Windows.Forms.Label
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents tbQtePallPrep As textBoxNumeric
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents tbQteColis As textBoxNumeric
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbPoids As textBoxNumeric
    Friend WithEvents tbQtePallNonPrep As textBoxNumeric
    Friend WithEvents tbPUPallNonPrep As textBoxCurrency
    Friend WithEvents tbPUPallPrep As textBoxCurrency
    Friend WithEvents cbFaxerBLTransporteur As System.Windows.Forms.Button
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents tbPoidsCmd As System.Windows.Forms.TextBox
    Friend WithEvents Label61 As System.Windows.Forms.Label
    Friend WithEvents tbQteColisCmd As System.Windows.Forms.TextBox
    Friend WithEvents cbCalcPoidsColis As System.Windows.Forms.Button
    Friend WithEvents tbFaxTRP As System.Windows.Forms.TextBox
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents cboCodeTRP As System.Windows.Forms.ComboBox
    Friend WithEvents tbMailPLTF As System.Windows.Forms.TextBox
    Friend WithEvents tbLettreVoiture As System.Windows.Forms.TextBox
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents Label63 As System.Windows.Forms.Label
    Friend WithEvents tbRefFactTRP As System.Windows.Forms.TextBox
    Friend WithEvents tbCoutTransport As vini_app.textBoxCurrency
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents tbMtComm As textBoxCurrency
    Friend WithEvents laMtComm As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcLignesCommande As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewCheckBoxColumn1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn15 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn16 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn17 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn18 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn19 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn20 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn21 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewCheckBoxColumn2 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn22 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn23 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn24 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents m_bsrcSousCommande As System.Windows.Forms.BindingSource
    Friend WithEvents m_dgvSCMD As System.Windows.Forms.DataGridView
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FournisseurRSDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalHTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateCommandeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EtatLibelleDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cbCalcPoidsColis_RL As System.Windows.Forms.Button
    Friend WithEvents Label66 As System.Windows.Forms.Label
    Friend WithEvents Label67 As System.Windows.Forms.Label
    Friend WithEvents cbCalcMontantTransport_RL As System.Windows.Forms.Button
    Friend WithEvents Label68 As System.Windows.Forms.Label
    Friend WithEvents Label69 As System.Windows.Forms.Label
    Friend WithEvents tbMontantTransport_RL As vini_app.textBoxCurrency
    Friend WithEvents tbPUPallNonPrep_RL As vini_app.textBoxCurrency
    Friend WithEvents tbPUPallPrep_RL As vini_app.textBoxCurrency
    Friend WithEvents Label70 As System.Windows.Forms.Label
    Friend WithEvents Label71 As System.Windows.Forms.Label
    Friend WithEvents tbQtePallNonPrep_RL As vini_app.textBoxNumeric
    Friend WithEvents tbQtePallPrep_RL As vini_app.textBoxNumeric
    Friend WithEvents Label72 As System.Windows.Forms.Label
    Friend WithEvents Label73 As System.Windows.Forms.Label
    Friend WithEvents tbPoids_RL As vini_app.textBoxNumeric
    Friend WithEvents Label74 As System.Windows.Forms.Label
    Friend WithEvents tbQteColis_RL As vini_app.textBoxNumeric
    Friend WithEvents Label75 As System.Windows.Forms.Label
    Friend WithEvents m_bsrcCommande As System.Windows.Forms.BindingSource
    Friend WithEvents crwDetailCommandeClient As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents crwBL As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents grpPrestashop As System.Windows.Forms.GroupBox
    Friend WithEvents NumDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProduitCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProduitNomDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProduitMilDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProduitConditionnementDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProduitCouleurDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteCommandeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteLivDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteFactDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BGratuitDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents PrixUDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PrixHTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PrixTTCDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cbxOrigine As System.Windows.Forms.ComboBox
    Friend WithEvents liPrestashop As System.Windows.Forms.LinkLabel
    Friend WithEvents laIntermediaires As System.Windows.Forms.Label
    Friend WithEvents cbxIntermédiaires As System.Windows.Forms.ComboBox
    Friend WithEvents m_bsrcIntermédiaires As System.Windows.Forms.BindingSource
    Friend WithEvents tpFactHbv As System.Windows.Forms.TabPage
    Friend WithEvents btnFactHBVAfficher As System.Windows.Forms.Button
    Friend WithEvents crwFact As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents DataGridView4 As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcLgFactHBV As System.Windows.Forms.BindingSource
    Friend WithEvents pnlFactHBV As System.Windows.Forms.Panel
    Friend WithEvents SplitFactHBV As System.Windows.Forms.SplitContainer
    Friend WithEvents ProduitCodeDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProduitNomDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QteFactDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BGratuitDataGridViewCheckBoxColumn1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents PrixUDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PrixHTDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PrixTTCDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cbMailBLPLTFRM As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle34 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle35 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle36 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle37 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle38 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle39 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle40 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle41 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle42 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle43 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle44 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle45 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle46 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle47 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle48 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle49 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle50 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle51 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle52 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle53 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle54 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle55 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle56 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle29 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle30 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle31 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle32 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle33 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.grpEntete = New System.Windows.Forms.GroupBox()
        Me.grpPrestashop = New System.Windows.Forms.GroupBox()
        Me.cbxOrigine = New System.Windows.Forms.ComboBox()
        Me.liPrestashop = New System.Windows.Forms.LinkLabel()
        Me.grpFactTRP = New System.Windows.Forms.GroupBox()
        Me.liFactTRP = New System.Windows.Forms.LinkLabel()
        Me.ckTransport = New System.Windows.Forms.CheckBox()
        Me.grpTypeTransport = New System.Windows.Forms.GroupBox()
        Me.rb_TypeTRP_Depart = New System.Windows.Forms.RadioButton()
        Me.rb_TypeTRP_Avance = New System.Windows.Forms.RadioButton()
        Me.rb_TypeTrp_Franco = New System.Windows.Forms.RadioButton()
        Me.grpTypeCommande = New System.Windows.Forms.GroupBox()
        Me.rbTypeCmdDirecte = New System.Windows.Forms.RadioButton()
        Me.rbTypeCmdPlateforme = New System.Windows.Forms.RadioButton()
        Me.dtDateCommande = New System.Windows.Forms.DateTimePicker()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.liNomClient = New System.Windows.Forms.LinkLabel()
        Me.laEtatCommande = New System.Windows.Forms.Label()
        Me.tbCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SSTabCommandeClient = New System.Windows.Forms.TabControl()
        Me.tpClient = New System.Windows.Forms.TabPage()
        Me.laIdClient = New System.Windows.Forms.Label()
        Me.tbBanque = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.cbRechercheclient = New System.Windows.Forms.Button()
        Me.tbCodeClient = New System.Windows.Forms.TextBox()
        Me.LaCodeTiers = New System.Windows.Forms.Label()
        Me.ckAdrIdentiques = New System.Windows.Forms.CheckBox()
        Me.grpAdrFact = New System.Windows.Forms.GroupBox()
        Me.tbAdrFact_Nom = New System.Windows.Forms.TextBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.tbAdrFact_Rue1 = New System.Windows.Forms.TextBox()
        Me.tbAdrFact_Rue2 = New System.Windows.Forms.TextBox()
        Me.tbAdrFact_cp = New vini_app.textBoxStrNumeric()
        Me.tbAdrFact_Ville = New System.Windows.Forms.TextBox()
        Me.tbAdrFact_Port = New vini_app.textBoxStrNumeric()
        Me.tbAdrFact_Tel = New vini_app.textBoxStrNumeric()
        Me.tbAdrFact_Fax = New vini_app.textBoxStrNumeric()
        Me.tbAdrFact_Email = New System.Windows.Forms.TextBox()
        Me.grpAdrLiv = New System.Windows.Forms.GroupBox()
        Me.tbAdrLiv_Nom = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Labelbis = New System.Windows.Forms.Label()
        Me.tbAdrLiv_Rue1 = New System.Windows.Forms.TextBox()
        Me.tbAdrLiv_Rue2 = New System.Windows.Forms.TextBox()
        Me.tbAdrLiv_cp = New vini_app.textBoxStrNumeric()
        Me.tbAdrLiv_Ville = New System.Windows.Forms.TextBox()
        Me.tbAdrLiv_Port = New vini_app.textBoxStrNumeric()
        Me.tbAdrLiv_Tel = New vini_app.textBoxStrNumeric()
        Me.tbAdrLiv_Fax = New vini_app.textBoxStrNumeric()
        Me.tbAdrLiv_Email = New System.Windows.Forms.TextBox()
        Me.tbRaisonSociale = New System.Windows.Forms.TextBox()
        Me.label = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboModeReglement = New System.Windows.Forms.ComboBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.tbRib1 = New vini_app.textBoxStrNumeric()
        Me.tbRib2 = New vini_app.textBoxStrNumeric()
        Me.tbRib3 = New vini_app.textBoxStrNumeric()
        Me.tbRib4 = New vini_app.textBoxStrNumeric()
        Me.tbNomClient = New System.Windows.Forms.TextBox()
        Me.tpLignes = New System.Windows.Forms.TabPage()
        Me.laMtComm = New System.Windows.Forms.Label()
        Me.tbMtComm = New vini_app.textBoxCurrency()
        Me.m_bsrcCommande = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbTotalTTC = New vini_app.textBoxCurrency()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.cbSupprimerLigne = New System.Windows.Forms.Button()
        Me.cbModifierLigne = New System.Windows.Forms.Button()
        Me.tbTotalHT = New vini_app.textBoxCurrency()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cbAjouterLigne = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.NumDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProduitCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProduitNomDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProduitMilDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProduitConditionnementDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProduitCouleurDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QteCommandeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QteLivDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QteFactDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BGratuitDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.PrixUDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrixHTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrixTTCDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_bsrcLignesCommande = New System.Windows.Forms.BindingSource(Me.components)
        Me.tbQteColisCmd = New System.Windows.Forms.TextBox()
        Me.Label61 = New System.Windows.Forms.Label()
        Me.tbPoidsCmd = New System.Windows.Forms.TextBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.tbRecalcTotaux = New System.Windows.Forms.Button()
        Me.tpCommentaires = New System.Windows.Forms.TabPage()
        Me.tbCommentaireFacturation = New System.Windows.Forms.RichTextBox()
        Me.tbCommentaireLivraison = New System.Windows.Forms.RichTextBox()
        Me.tbCommentaireCommande = New System.Windows.Forms.RichTextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.tpValidCmd = New System.Windows.Forms.TabPage()
        Me.crwDetailCommandeClient = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.dtDateLivraisonPrevue = New System.Windows.Forms.DateTimePicker()
        Me.label50 = New System.Windows.Forms.Label()
        Me.cbVisuDetailCommande = New System.Windows.Forms.Button()
        Me.tbComValid = New System.Windows.Forms.RichTextBox()
        Me.tbNumFaxBC = New System.Windows.Forms.Label()
        Me.ckCmdValide = New System.Windows.Forms.CheckBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.dtDateValidation = New System.Windows.Forms.DateTimePicker()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.tbNumFaxValidation = New System.Windows.Forms.TextBox()
        Me.cbFax = New System.Windows.Forms.Button()
        Me.tpBL = New System.Windows.Forms.TabPage()
        Me.crwBL = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.tbFaxTRP = New System.Windows.Forms.TextBox()
        Me.tbMailPLTF = New System.Windows.Forms.TextBox()
        Me.cbFaxerBLTransporteur = New System.Windows.Forms.Button()
        Me.grpInfFactureTRP = New System.Windows.Forms.GroupBox()
        Me.cbCalcPoidsColis = New System.Windows.Forms.Button()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.cbCalcMontantTransport = New System.Windows.Forms.Button()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.tbMontantTransport = New vini_app.textBoxCurrency()
        Me.tbPUPallNonPrep = New vini_app.textBoxCurrency()
        Me.tbPUPallPrep = New vini_app.textBoxCurrency()
        Me.Label56 = New System.Windows.Forms.Label()
        Me.Label55 = New System.Windows.Forms.Label()
        Me.tbQtePallNonPrep = New vini_app.textBoxNumeric()
        Me.tbQtePallPrep = New vini_app.textBoxNumeric()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.tbPoids = New vini_app.textBoxNumeric()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.tbQteColis = New vini_app.textBoxNumeric()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbPiedPageBL = New System.Windows.Forms.RichTextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.cbVisuBL = New System.Windows.Forms.Button()
        Me.grpTransporteur = New System.Windows.Forms.GroupBox()
        Me.cboCodeTRP = New System.Windows.Forms.ComboBox()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.tbTrpEmail = New System.Windows.Forms.TextBox()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.tbTrpFax = New System.Windows.Forms.TextBox()
        Me.tbTrpPort = New System.Windows.Forms.TextBox()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.tbTrpTel = New System.Windows.Forms.TextBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.tbTrpVille = New System.Windows.Forms.TextBox()
        Me.tbTrpCP = New System.Windows.Forms.TextBox()
        Me.tbTrpRue2 = New System.Windows.Forms.TextBox()
        Me.tbTrpRue1 = New System.Windows.Forms.TextBox()
        Me.tbTrpNom = New System.Windows.Forms.TextBox()
        Me.dtDateLivraison = New System.Windows.Forms.DateTimePicker()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.dtDateEnlev = New System.Windows.Forms.DateTimePicker()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.cbMailBLPLTFRM = New System.Windows.Forms.Button()
        Me.tpRetourBL = New System.Windows.Forms.TabPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cbCalcPoidsColis_RL = New System.Windows.Forms.Button()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.Label67 = New System.Windows.Forms.Label()
        Me.cbCalcMontantTransport_RL = New System.Windows.Forms.Button()
        Me.Label68 = New System.Windows.Forms.Label()
        Me.Label69 = New System.Windows.Forms.Label()
        Me.tbMontantTransport_RL = New vini_app.textBoxCurrency()
        Me.tbPUPallNonPrep_RL = New vini_app.textBoxCurrency()
        Me.tbPUPallPrep_RL = New vini_app.textBoxCurrency()
        Me.Label70 = New System.Windows.Forms.Label()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.tbQtePallNonPrep_RL = New vini_app.textBoxNumeric()
        Me.tbQtePallPrep_RL = New vini_app.textBoxNumeric()
        Me.Label72 = New System.Windows.Forms.Label()
        Me.Label73 = New System.Windows.Forms.Label()
        Me.tbPoids_RL = New vini_app.textBoxNumeric()
        Me.Label74 = New System.Windows.Forms.Label()
        Me.tbQteColis_RL = New vini_app.textBoxNumeric()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewCheckBoxColumn1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tbCoutTransport = New vini_app.textBoxCurrency()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.tbRefFactTRP = New System.Windows.Forms.TextBox()
        Me.tbLettreVoiture = New System.Windows.Forms.TextBox()
        Me.cbAnnulerLivraison = New System.Windows.Forms.Button()
        Me.dtDateEnlevementReelle = New System.Windows.Forms.DateTimePicker()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.tbRefBL = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.dtDateLivraisonReelle = New System.Windows.Forms.DateTimePicker()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.cbBLToutOK = New System.Windows.Forms.Button()
        Me.tpEclatement = New System.Windows.Forms.TabPage()
        Me.laIntermediaires = New System.Windows.Forms.Label()
        Me.cbxIntermédiaires = New System.Windows.Forms.ComboBox()
        Me.m_bsrcIntermédiaires = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_dgvSCMD = New System.Windows.Forms.DataGridView()
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FournisseurRSDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalHTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DateCommandeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EtatLibelleDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_bsrcSousCommande = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn17 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn18 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn19 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn20 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn21 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewCheckBoxColumn2 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataGridViewTextBoxColumn22 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn23 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn24 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cbSCMDAppliquer = New System.Windows.Forms.Button()
        Me.cbAnnEclatement = New System.Windows.Forms.Button()
        Me.tbSCMDCommentaire = New System.Windows.Forms.RichTextBox()
        Me.cbSCMDVoir = New System.Windows.Forms.Button()
        Me.cbSCMDFaxerTout = New System.Windows.Forms.Button()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.tbSCMDTransporteurNom = New System.Windows.Forms.TextBox()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.liSCMDFournisseur = New System.Windows.Forms.LinkLabel()
        Me.dtSCMDDateLiv = New System.Windows.Forms.DateTimePicker()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.cbEclatementCmde = New System.Windows.Forms.Button()
        Me.tpFactHbv = New System.Windows.Forms.TabPage()
        Me.pnlFactHBV = New System.Windows.Forms.Panel()
        Me.SplitFactHBV = New System.Windows.Forms.SplitContainer()
        Me.crwFact = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.btnFactHBVAfficher = New System.Windows.Forms.Button()
        Me.DataGridView4 = New System.Windows.Forms.DataGridView()
        Me.ProduitCodeDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProduitNomDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QteFactDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BGratuitDataGridViewCheckBoxColumn1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.PrixUDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrixHTDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrixTTCDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_bsrcLgFactHBV = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.grpEntete.SuspendLayout()
        Me.grpPrestashop.SuspendLayout()
        Me.grpFactTRP.SuspendLayout()
        Me.grpTypeTransport.SuspendLayout()
        Me.grpTypeCommande.SuspendLayout()
        Me.SSTabCommandeClient.SuspendLayout()
        Me.tpClient.SuspendLayout()
        Me.grpAdrFact.SuspendLayout()
        Me.grpAdrLiv.SuspendLayout()
        Me.tpLignes.SuspendLayout()
        CType(Me.m_bsrcCommande, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcLignesCommande, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpCommentaires.SuspendLayout()
        Me.tpValidCmd.SuspendLayout()
        Me.tpBL.SuspendLayout()
        Me.grpInfFactureTRP.SuspendLayout()
        Me.grpTransporteur.SuspendLayout()
        Me.tpRetourBL.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpEclatement.SuspendLayout()
        CType(Me.m_bsrcIntermédiaires, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_dgvSCMD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcSousCommande, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpFactHbv.SuspendLayout()
        Me.pnlFactHBV.SuspendLayout()
        CType(Me.SplitFactHBV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitFactHBV.Panel1.SuspendLayout()
        Me.SplitFactHBV.Panel2.SuspendLayout()
        Me.SplitFactHBV.SuspendLayout()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcLgFactHBV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpEntete
        '
        Me.grpEntete.Controls.Add(Me.grpPrestashop)
        Me.grpEntete.Controls.Add(Me.grpFactTRP)
        Me.grpEntete.Controls.Add(Me.grpTypeTransport)
        Me.grpEntete.Controls.Add(Me.grpTypeCommande)
        Me.grpEntete.Controls.Add(Me.dtDateCommande)
        Me.grpEntete.Controls.Add(Me.Label29)
        Me.grpEntete.Controls.Add(Me.liNomClient)
        Me.grpEntete.Controls.Add(Me.laEtatCommande)
        Me.grpEntete.Controls.Add(Me.tbCode)
        Me.grpEntete.Controls.Add(Me.Label1)
        Me.grpEntete.Location = New System.Drawing.Point(8, 0)
        Me.grpEntete.Name = "grpEntete"
        Me.grpEntete.Size = New System.Drawing.Size(992, 64)
        Me.grpEntete.TabIndex = 0
        Me.grpEntete.TabStop = False
        Me.grpEntete.Text = "Commande"
        '
        'grpPrestashop
        '
        Me.grpPrestashop.Controls.Add(Me.cbxOrigine)
        Me.grpPrestashop.Controls.Add(Me.liPrestashop)
        Me.grpPrestashop.Enabled = False
        Me.grpPrestashop.Location = New System.Drawing.Point(725, 0)
        Me.grpPrestashop.Name = "grpPrestashop"
        Me.grpPrestashop.Size = New System.Drawing.Size(136, 64)
        Me.grpPrestashop.TabIndex = 15
        Me.grpPrestashop.TabStop = False
        '
        'cbxOrigine
        '
        Me.cbxOrigine.FormattingEnabled = True
        Me.cbxOrigine.Items.AddRange(New Object() {"VINICOM", "HOBIVIN"})
        Me.cbxOrigine.Location = New System.Drawing.Point(6, 16)
        Me.cbxOrigine.Name = "cbxOrigine"
        Me.cbxOrigine.Size = New System.Drawing.Size(124, 21)
        Me.cbxOrigine.TabIndex = 1
        '
        'liPrestashop
        '
        Me.liPrestashop.AutoSize = True
        Me.liPrestashop.Location = New System.Drawing.Point(6, 43)
        Me.liPrestashop.Name = "liPrestashop"
        Me.liPrestashop.Size = New System.Drawing.Size(64, 13)
        Me.liPrestashop.TabIndex = 0
        Me.liPrestashop.TabStop = True
        Me.liPrestashop.Text = "liPrestashop"
        '
        'grpFactTRP
        '
        Me.grpFactTRP.Controls.Add(Me.liFactTRP)
        Me.grpFactTRP.Controls.Add(Me.ckTransport)
        Me.grpFactTRP.Enabled = False
        Me.grpFactTRP.Location = New System.Drawing.Point(581, 0)
        Me.grpFactTRP.Name = "grpFactTRP"
        Me.grpFactTRP.Size = New System.Drawing.Size(144, 64)
        Me.grpFactTRP.TabIndex = 14
        Me.grpFactTRP.TabStop = False
        '
        'liFactTRP
        '
        Me.liFactTRP.Location = New System.Drawing.Point(8, 40)
        Me.liFactTRP.Name = "liFactTRP"
        Me.liFactTRP.Size = New System.Drawing.Size(130, 16)
        Me.liFactTRP.TabIndex = 1
        Me.liFactTRP.TabStop = True
        Me.liFactTRP.Text = "Fact Transport"
        '
        'ckTransport
        '
        Me.ckTransport.Location = New System.Drawing.Point(8, 16)
        Me.ckTransport.Name = "ckTransport"
        Me.ckTransport.Size = New System.Drawing.Size(130, 24)
        Me.ckTransport.TabIndex = 0
        Me.ckTransport.Text = "Facture de transport"
        '
        'grpTypeTransport
        '
        Me.grpTypeTransport.Controls.Add(Me.rb_TypeTRP_Depart)
        Me.grpTypeTransport.Controls.Add(Me.rb_TypeTRP_Avance)
        Me.grpTypeTransport.Controls.Add(Me.rb_TypeTrp_Franco)
        Me.grpTypeTransport.Location = New System.Drawing.Point(495, 0)
        Me.grpTypeTransport.Name = "grpTypeTransport"
        Me.grpTypeTransport.Size = New System.Drawing.Size(94, 64)
        Me.grpTypeTransport.TabIndex = 2
        Me.grpTypeTransport.TabStop = False
        '
        'rb_TypeTRP_Depart
        '
        Me.rb_TypeTRP_Depart.Location = New System.Drawing.Point(6, 42)
        Me.rb_TypeTRP_Depart.Name = "rb_TypeTRP_Depart"
        Me.rb_TypeTRP_Depart.Size = New System.Drawing.Size(64, 16)
        Me.rb_TypeTRP_Depart.TabIndex = 2
        Me.rb_TypeTRP_Depart.Text = "Départ"
        '
        'rb_TypeTRP_Avance
        '
        Me.rb_TypeTRP_Avance.Location = New System.Drawing.Point(6, 26)
        Me.rb_TypeTRP_Avance.Name = "rb_TypeTRP_Avance"
        Me.rb_TypeTRP_Avance.Size = New System.Drawing.Size(64, 16)
        Me.rb_TypeTRP_Avance.TabIndex = 1
        Me.rb_TypeTRP_Avance.Text = "Avancé"
        '
        'rb_TypeTrp_Franco
        '
        Me.rb_TypeTrp_Franco.Location = New System.Drawing.Point(6, 10)
        Me.rb_TypeTrp_Franco.Name = "rb_TypeTrp_Franco"
        Me.rb_TypeTrp_Franco.Size = New System.Drawing.Size(64, 16)
        Me.rb_TypeTrp_Franco.TabIndex = 0
        Me.rb_TypeTrp_Franco.Text = "Franco"
        '
        'grpTypeCommande
        '
        Me.grpTypeCommande.Controls.Add(Me.rbTypeCmdDirecte)
        Me.grpTypeCommande.Controls.Add(Me.rbTypeCmdPlateforme)
        Me.grpTypeCommande.Location = New System.Drawing.Point(392, 0)
        Me.grpTypeCommande.Name = "grpTypeCommande"
        Me.grpTypeCommande.Size = New System.Drawing.Size(103, 64)
        Me.grpTypeCommande.TabIndex = 1
        Me.grpTypeCommande.TabStop = False
        '
        'rbTypeCmdDirecte
        '
        Me.rbTypeCmdDirecte.Location = New System.Drawing.Point(8, 40)
        Me.rbTypeCmdDirecte.Name = "rbTypeCmdDirecte"
        Me.rbTypeCmdDirecte.Size = New System.Drawing.Size(89, 16)
        Me.rbTypeCmdDirecte.TabIndex = 1
        Me.rbTypeCmdDirecte.Text = "Directe"
        '
        'rbTypeCmdPlateforme
        '
        Me.rbTypeCmdPlateforme.Location = New System.Drawing.Point(8, 8)
        Me.rbTypeCmdPlateforme.Name = "rbTypeCmdPlateforme"
        Me.rbTypeCmdPlateforme.Size = New System.Drawing.Size(89, 24)
        Me.rbTypeCmdPlateforme.TabIndex = 0
        Me.rbTypeCmdPlateforme.Text = "Plateforme"
        '
        'dtDateCommande
        '
        Me.dtDateCommande.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateCommande.Location = New System.Drawing.Point(280, 16)
        Me.dtDateCommande.Name = "dtDateCommande"
        Me.dtDateCommande.Size = New System.Drawing.Size(104, 20)
        Me.dtDateCommande.TabIndex = 13
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(168, 16)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(104, 16)
        Me.Label29.TabIndex = 12
        Me.Label29.Text = "Date de commande"
        '
        'liNomClient
        '
        Me.liNomClient.Location = New System.Drawing.Point(8, 40)
        Me.liNomClient.Name = "liNomClient"
        Me.liNomClient.Size = New System.Drawing.Size(376, 16)
        Me.liNomClient.TabIndex = 8
        '
        'laEtatCommande
        '
        Me.laEtatCommande.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.laEtatCommande.ForeColor = System.Drawing.Color.Red
        Me.laEtatCommande.Location = New System.Drawing.Point(867, 16)
        Me.laEtatCommande.Name = "laEtatCommande"
        Me.laEtatCommande.Size = New System.Drawing.Size(121, 40)
        Me.laEtatCommande.TabIndex = 7
        '
        'tbCode
        '
        Me.tbCode.Enabled = False
        Me.tbCode.Location = New System.Drawing.Point(48, 16)
        Me.tbCode.Name = "tbCode"
        Me.tbCode.Size = New System.Drawing.Size(112, 20)
        Me.tbCode.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "N°"
        '
        'SSTabCommandeClient
        '
        Me.SSTabCommandeClient.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SSTabCommandeClient.Controls.Add(Me.tpClient)
        Me.SSTabCommandeClient.Controls.Add(Me.tpLignes)
        Me.SSTabCommandeClient.Controls.Add(Me.tpCommentaires)
        Me.SSTabCommandeClient.Controls.Add(Me.tpValidCmd)
        Me.SSTabCommandeClient.Controls.Add(Me.tpBL)
        Me.SSTabCommandeClient.Controls.Add(Me.tpRetourBL)
        Me.SSTabCommandeClient.Controls.Add(Me.tpEclatement)
        Me.SSTabCommandeClient.Controls.Add(Me.tpFactHbv)
        Me.SSTabCommandeClient.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTabCommandeClient.Location = New System.Drawing.Point(0, 64)
        Me.SSTabCommandeClient.Name = "SSTabCommandeClient"
        Me.SSTabCommandeClient.SelectedIndex = 0
        Me.SSTabCommandeClient.Size = New System.Drawing.Size(1000, 583)
        Me.SSTabCommandeClient.TabIndex = 1
        '
        'tpClient
        '
        Me.tpClient.Controls.Add(Me.laIdClient)
        Me.tpClient.Controls.Add(Me.tbBanque)
        Me.tpClient.Controls.Add(Me.Label30)
        Me.tpClient.Controls.Add(Me.cbRechercheclient)
        Me.tpClient.Controls.Add(Me.tbCodeClient)
        Me.tpClient.Controls.Add(Me.LaCodeTiers)
        Me.tpClient.Controls.Add(Me.ckAdrIdentiques)
        Me.tpClient.Controls.Add(Me.grpAdrFact)
        Me.tpClient.Controls.Add(Me.grpAdrLiv)
        Me.tpClient.Controls.Add(Me.tbRaisonSociale)
        Me.tpClient.Controls.Add(Me.label)
        Me.tpClient.Controls.Add(Me.Label6)
        Me.tpClient.Controls.Add(Me.cboModeReglement)
        Me.tpClient.Controls.Add(Me.Label20)
        Me.tpClient.Controls.Add(Me.Label19)
        Me.tpClient.Controls.Add(Me.tbRib1)
        Me.tpClient.Controls.Add(Me.tbRib2)
        Me.tpClient.Controls.Add(Me.tbRib3)
        Me.tpClient.Controls.Add(Me.tbRib4)
        Me.tpClient.Controls.Add(Me.tbNomClient)
        Me.tpClient.Location = New System.Drawing.Point(4, 22)
        Me.tpClient.Name = "tpClient"
        Me.tpClient.Size = New System.Drawing.Size(992, 557)
        Me.tpClient.TabIndex = 0
        Me.tpClient.Text = "Client"
        '
        'laIdClient
        '
        Me.laIdClient.Location = New System.Drawing.Point(288, 8)
        Me.laIdClient.Name = "laIdClient"
        Me.laIdClient.Size = New System.Drawing.Size(40, 16)
        Me.laIdClient.TabIndex = 106
        Me.laIdClient.Visible = False
        '
        'tbBanque
        '
        Me.tbBanque.Location = New System.Drawing.Point(136, 424)
        Me.tbBanque.Name = "tbBanque"
        Me.tbBanque.Size = New System.Drawing.Size(376, 20)
        Me.tbBanque.TabIndex = 9
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(16, 424)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(112, 24)
        Me.Label30.TabIndex = 104
        Me.Label30.Text = "Banque"
        '
        'cbRechercheclient
        '
        Me.cbRechercheclient.BackColor = System.Drawing.SystemColors.Control
        Me.cbRechercheclient.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbRechercheclient.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbRechercheclient.Location = New System.Drawing.Point(192, 8)
        Me.cbRechercheclient.Name = "cbRechercheclient"
        Me.cbRechercheclient.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbRechercheclient.Size = New System.Drawing.Size(80, 24)
        Me.cbRechercheclient.TabIndex = 1
        Me.cbRechercheclient.Text = "Rechercher"
        Me.cbRechercheclient.UseVisualStyleBackColor = False
        '
        'tbCodeClient
        '
        Me.tbCodeClient.Location = New System.Drawing.Point(104, 8)
        Me.tbCodeClient.Name = "tbCodeClient"
        Me.tbCodeClient.Size = New System.Drawing.Size(72, 20)
        Me.tbCodeClient.TabIndex = 0
        '
        'LaCodeTiers
        '
        Me.LaCodeTiers.Location = New System.Drawing.Point(8, 8)
        Me.LaCodeTiers.Name = "LaCodeTiers"
        Me.LaCodeTiers.Size = New System.Drawing.Size(40, 16)
        Me.LaCodeTiers.TabIndex = 100
        Me.LaCodeTiers.Text = "Client"
        '
        'ckAdrIdentiques
        '
        Me.ckAdrIdentiques.Location = New System.Drawing.Point(8, 232)
        Me.ckAdrIdentiques.Name = "ckAdrIdentiques"
        Me.ckAdrIdentiques.Size = New System.Drawing.Size(144, 16)
        Me.ckAdrIdentiques.TabIndex = 5
        Me.ckAdrIdentiques.Text = "Adresses Identiques"
        '
        'grpAdrFact
        '
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_Nom)
        Me.grpAdrFact.Controls.Add(Me.Label38)
        Me.grpAdrFact.Controls.Add(Me.Label18)
        Me.grpAdrFact.Controls.Add(Me.Label17)
        Me.grpAdrFact.Controls.Add(Me.Label16)
        Me.grpAdrFact.Controls.Add(Me.Label14)
        Me.grpAdrFact.Controls.Add(Me.Label13)
        Me.grpAdrFact.Controls.Add(Me.Label12)
        Me.grpAdrFact.Controls.Add(Me.Label11)
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_Rue1)
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_Rue2)
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_cp)
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_Ville)
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_Port)
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_Tel)
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_Fax)
        Me.grpAdrFact.Controls.Add(Me.tbAdrFact_Email)
        Me.grpAdrFact.Location = New System.Drawing.Point(8, 256)
        Me.grpAdrFact.Name = "grpAdrFact"
        Me.grpAdrFact.Size = New System.Drawing.Size(944, 128)
        Me.grpAdrFact.TabIndex = 6
        Me.grpAdrFact.TabStop = False
        Me.grpAdrFact.Text = "Adresse Facturation"
        '
        'tbAdrFact_Nom
        '
        Me.tbAdrFact_Nom.Location = New System.Drawing.Point(96, 24)
        Me.tbAdrFact_Nom.Name = "tbAdrFact_Nom"
        Me.tbAdrFact_Nom.Size = New System.Drawing.Size(480, 20)
        Me.tbAdrFact_Nom.TabIndex = 0
        '
        'Label38
        '
        Me.Label38.Location = New System.Drawing.Point(16, 24)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(80, 16)
        Me.Label38.TabIndex = 102
        Me.Label38.Text = "Nom"
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(16, 96)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(73, 17)
        Me.Label18.TabIndex = 95
        Me.Label18.Text = "CP/Ville"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(16, 72)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(73, 17)
        Me.Label17.TabIndex = 96
        Me.Label17.Text = "Adresse2"
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label16.Location = New System.Drawing.Point(16, 48)
        Me.Label16.Name = "Label16"
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.Size = New System.Drawing.Size(73, 17)
        Me.Label16.TabIndex = 97
        Me.Label16.Text = "Adresse1"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(600, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(24, 17)
        Me.Label14.TabIndex = 98
        Me.Label14.Text = "Tél"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(600, 96)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(24, 17)
        Me.Label13.TabIndex = 99
        Me.Label13.Text = "Fax"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(760, 72)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(57, 17)
        Me.Label12.TabIndex = 100
        Me.Label12.Text = "Portable"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(760, 96)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(41, 17)
        Me.Label11.TabIndex = 101
        Me.Label11.Text = "Email"
        '
        'tbAdrFact_Rue1
        '
        Me.tbAdrFact_Rue1.AcceptsReturn = True
        Me.tbAdrFact_Rue1.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFact_Rue1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFact_Rue1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFact_Rue1.Location = New System.Drawing.Point(96, 48)
        Me.tbAdrFact_Rue1.MaxLength = 0
        Me.tbAdrFact_Rue1.Name = "tbAdrFact_Rue1"
        Me.tbAdrFact_Rue1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFact_Rue1.Size = New System.Drawing.Size(480, 20)
        Me.tbAdrFact_Rue1.TabIndex = 1
        '
        'tbAdrFact_Rue2
        '
        Me.tbAdrFact_Rue2.AcceptsReturn = True
        Me.tbAdrFact_Rue2.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFact_Rue2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFact_Rue2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFact_Rue2.Location = New System.Drawing.Point(96, 72)
        Me.tbAdrFact_Rue2.MaxLength = 0
        Me.tbAdrFact_Rue2.Name = "tbAdrFact_Rue2"
        Me.tbAdrFact_Rue2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFact_Rue2.Size = New System.Drawing.Size(480, 20)
        Me.tbAdrFact_Rue2.TabIndex = 2
        '
        'tbAdrFact_cp
        '
        Me.tbAdrFact_cp.AcceptsReturn = True
        Me.tbAdrFact_cp.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFact_cp.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFact_cp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFact_cp.Location = New System.Drawing.Point(96, 96)
        Me.tbAdrFact_cp.MaxLength = 0
        Me.tbAdrFact_cp.Name = "tbAdrFact_cp"
        Me.tbAdrFact_cp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFact_cp.Size = New System.Drawing.Size(73, 20)
        Me.tbAdrFact_cp.TabIndex = 3
        Me.tbAdrFact_cp.Text = "0"
        '
        'tbAdrFact_Ville
        '
        Me.tbAdrFact_Ville.AcceptsReturn = True
        Me.tbAdrFact_Ville.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFact_Ville.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFact_Ville.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFact_Ville.Location = New System.Drawing.Point(176, 96)
        Me.tbAdrFact_Ville.MaxLength = 0
        Me.tbAdrFact_Ville.Name = "tbAdrFact_Ville"
        Me.tbAdrFact_Ville.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFact_Ville.Size = New System.Drawing.Size(400, 20)
        Me.tbAdrFact_Ville.TabIndex = 4
        '
        'tbAdrFact_Port
        '
        Me.tbAdrFact_Port.AcceptsReturn = True
        Me.tbAdrFact_Port.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFact_Port.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFact_Port.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFact_Port.Location = New System.Drawing.Point(816, 72)
        Me.tbAdrFact_Port.MaxLength = 0
        Me.tbAdrFact_Port.Name = "tbAdrFact_Port"
        Me.tbAdrFact_Port.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFact_Port.Size = New System.Drawing.Size(112, 20)
        Me.tbAdrFact_Port.TabIndex = 7
        Me.tbAdrFact_Port.Text = "0"
        '
        'tbAdrFact_Tel
        '
        Me.tbAdrFact_Tel.AcceptsReturn = True
        Me.tbAdrFact_Tel.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFact_Tel.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFact_Tel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFact_Tel.Location = New System.Drawing.Point(624, 72)
        Me.tbAdrFact_Tel.MaxLength = 0
        Me.tbAdrFact_Tel.Name = "tbAdrFact_Tel"
        Me.tbAdrFact_Tel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFact_Tel.Size = New System.Drawing.Size(120, 20)
        Me.tbAdrFact_Tel.TabIndex = 5
        Me.tbAdrFact_Tel.Text = "0"
        '
        'tbAdrFact_Fax
        '
        Me.tbAdrFact_Fax.AcceptsReturn = True
        Me.tbAdrFact_Fax.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFact_Fax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFact_Fax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFact_Fax.Location = New System.Drawing.Point(624, 96)
        Me.tbAdrFact_Fax.MaxLength = 0
        Me.tbAdrFact_Fax.Name = "tbAdrFact_Fax"
        Me.tbAdrFact_Fax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFact_Fax.Size = New System.Drawing.Size(120, 20)
        Me.tbAdrFact_Fax.TabIndex = 6
        Me.tbAdrFact_Fax.Text = "0"
        '
        'tbAdrFact_Email
        '
        Me.tbAdrFact_Email.AcceptsReturn = True
        Me.tbAdrFact_Email.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrFact_Email.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrFact_Email.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrFact_Email.Location = New System.Drawing.Point(816, 96)
        Me.tbAdrFact_Email.MaxLength = 0
        Me.tbAdrFact_Email.Name = "tbAdrFact_Email"
        Me.tbAdrFact_Email.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrFact_Email.Size = New System.Drawing.Size(112, 20)
        Me.tbAdrFact_Email.TabIndex = 8
        '
        'grpAdrLiv
        '
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_Nom)
        Me.grpAdrLiv.Controls.Add(Me.Label15)
        Me.grpAdrLiv.Controls.Add(Me.Label5)
        Me.grpAdrLiv.Controls.Add(Me.Label3)
        Me.grpAdrLiv.Controls.Add(Me.Label4)
        Me.grpAdrLiv.Controls.Add(Me.Label7)
        Me.grpAdrLiv.Controls.Add(Me.Label8)
        Me.grpAdrLiv.Controls.Add(Me.Label9)
        Me.grpAdrLiv.Controls.Add(Me.Labelbis)
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_Rue1)
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_Rue2)
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_cp)
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_Ville)
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_Port)
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_Tel)
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_Fax)
        Me.grpAdrLiv.Controls.Add(Me.tbAdrLiv_Email)
        Me.grpAdrLiv.Location = New System.Drawing.Point(8, 96)
        Me.grpAdrLiv.Name = "grpAdrLiv"
        Me.grpAdrLiv.Size = New System.Drawing.Size(944, 128)
        Me.grpAdrLiv.TabIndex = 4
        Me.grpAdrLiv.TabStop = False
        Me.grpAdrLiv.Text = "Adresse Livraison"
        '
        'tbAdrLiv_Nom
        '
        Me.tbAdrLiv_Nom.Location = New System.Drawing.Point(96, 24)
        Me.tbAdrLiv_Nom.Name = "tbAdrLiv_Nom"
        Me.tbAdrLiv_Nom.Size = New System.Drawing.Size(480, 20)
        Me.tbAdrLiv_Nom.TabIndex = 0
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(16, 24)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(80, 16)
        Me.Label15.TabIndex = 85
        Me.Label15.Text = "Nom"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(16, 96)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(73, 17)
        Me.Label5.TabIndex = 78
        Me.Label5.Text = "CP/Ville"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 79
        Me.Label3.Text = "Adresse2"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(16, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(73, 17)
        Me.Label4.TabIndex = 80
        Me.Label4.Text = "Adresse1"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(592, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(25, 17)
        Me.Label7.TabIndex = 81
        Me.Label7.Text = "Tél"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(592, 96)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(24, 17)
        Me.Label8.TabIndex = 82
        Me.Label8.Text = "Fax"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(752, 72)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(48, 17)
        Me.Label9.TabIndex = 83
        Me.Label9.Text = "Portable"
        '
        'Labelbis
        '
        Me.Labelbis.BackColor = System.Drawing.SystemColors.Control
        Me.Labelbis.Cursor = System.Windows.Forms.Cursors.Default
        Me.Labelbis.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Labelbis.Location = New System.Drawing.Point(760, 96)
        Me.Labelbis.Name = "Labelbis"
        Me.Labelbis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Labelbis.Size = New System.Drawing.Size(41, 17)
        Me.Labelbis.TabIndex = 84
        Me.Labelbis.Text = "Email"
        '
        'tbAdrLiv_Rue1
        '
        Me.tbAdrLiv_Rue1.AcceptsReturn = True
        Me.tbAdrLiv_Rue1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbAdrLiv_Rue1.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLiv_Rue1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLiv_Rue1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLiv_Rue1.Location = New System.Drawing.Point(96, 48)
        Me.tbAdrLiv_Rue1.MaxLength = 0
        Me.tbAdrLiv_Rue1.Name = "tbAdrLiv_Rue1"
        Me.tbAdrLiv_Rue1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLiv_Rue1.Size = New System.Drawing.Size(480, 20)
        Me.tbAdrLiv_Rue1.TabIndex = 1
        '
        'tbAdrLiv_Rue2
        '
        Me.tbAdrLiv_Rue2.AcceptsReturn = True
        Me.tbAdrLiv_Rue2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbAdrLiv_Rue2.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLiv_Rue2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLiv_Rue2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLiv_Rue2.Location = New System.Drawing.Point(96, 72)
        Me.tbAdrLiv_Rue2.MaxLength = 0
        Me.tbAdrLiv_Rue2.Name = "tbAdrLiv_Rue2"
        Me.tbAdrLiv_Rue2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLiv_Rue2.Size = New System.Drawing.Size(480, 20)
        Me.tbAdrLiv_Rue2.TabIndex = 2
        '
        'tbAdrLiv_cp
        '
        Me.tbAdrLiv_cp.AcceptsReturn = True
        Me.tbAdrLiv_cp.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLiv_cp.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLiv_cp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLiv_cp.Location = New System.Drawing.Point(96, 96)
        Me.tbAdrLiv_cp.MaxLength = 0
        Me.tbAdrLiv_cp.Name = "tbAdrLiv_cp"
        Me.tbAdrLiv_cp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLiv_cp.Size = New System.Drawing.Size(73, 20)
        Me.tbAdrLiv_cp.TabIndex = 3
        Me.tbAdrLiv_cp.Text = "0"
        '
        'tbAdrLiv_Ville
        '
        Me.tbAdrLiv_Ville.AcceptsReturn = True
        Me.tbAdrLiv_Ville.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbAdrLiv_Ville.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLiv_Ville.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLiv_Ville.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLiv_Ville.Location = New System.Drawing.Point(176, 96)
        Me.tbAdrLiv_Ville.MaxLength = 0
        Me.tbAdrLiv_Ville.Name = "tbAdrLiv_Ville"
        Me.tbAdrLiv_Ville.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLiv_Ville.Size = New System.Drawing.Size(400, 20)
        Me.tbAdrLiv_Ville.TabIndex = 4
        '
        'tbAdrLiv_Port
        '
        Me.tbAdrLiv_Port.AcceptsReturn = True
        Me.tbAdrLiv_Port.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLiv_Port.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLiv_Port.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLiv_Port.Location = New System.Drawing.Point(808, 72)
        Me.tbAdrLiv_Port.MaxLength = 0
        Me.tbAdrLiv_Port.Name = "tbAdrLiv_Port"
        Me.tbAdrLiv_Port.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLiv_Port.Size = New System.Drawing.Size(120, 20)
        Me.tbAdrLiv_Port.TabIndex = 7
        Me.tbAdrLiv_Port.Text = "0"
        '
        'tbAdrLiv_Tel
        '
        Me.tbAdrLiv_Tel.AcceptsReturn = True
        Me.tbAdrLiv_Tel.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLiv_Tel.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLiv_Tel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLiv_Tel.Location = New System.Drawing.Point(624, 72)
        Me.tbAdrLiv_Tel.MaxLength = 0
        Me.tbAdrLiv_Tel.Name = "tbAdrLiv_Tel"
        Me.tbAdrLiv_Tel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLiv_Tel.Size = New System.Drawing.Size(120, 20)
        Me.tbAdrLiv_Tel.TabIndex = 5
        Me.tbAdrLiv_Tel.Text = "0"
        '
        'tbAdrLiv_Fax
        '
        Me.tbAdrLiv_Fax.AcceptsReturn = True
        Me.tbAdrLiv_Fax.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLiv_Fax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLiv_Fax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLiv_Fax.Location = New System.Drawing.Point(624, 96)
        Me.tbAdrLiv_Fax.MaxLength = 0
        Me.tbAdrLiv_Fax.Name = "tbAdrLiv_Fax"
        Me.tbAdrLiv_Fax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLiv_Fax.Size = New System.Drawing.Size(120, 20)
        Me.tbAdrLiv_Fax.TabIndex = 6
        Me.tbAdrLiv_Fax.Text = "0"
        '
        'tbAdrLiv_Email
        '
        Me.tbAdrLiv_Email.AcceptsReturn = True
        Me.tbAdrLiv_Email.BackColor = System.Drawing.SystemColors.Window
        Me.tbAdrLiv_Email.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbAdrLiv_Email.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbAdrLiv_Email.Location = New System.Drawing.Point(808, 96)
        Me.tbAdrLiv_Email.MaxLength = 0
        Me.tbAdrLiv_Email.Name = "tbAdrLiv_Email"
        Me.tbAdrLiv_Email.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbAdrLiv_Email.Size = New System.Drawing.Size(120, 20)
        Me.tbAdrLiv_Email.TabIndex = 8
        '
        'tbRaisonSociale
        '
        Me.tbRaisonSociale.Location = New System.Drawing.Point(104, 64)
        Me.tbRaisonSociale.Name = "tbRaisonSociale"
        Me.tbRaisonSociale.Size = New System.Drawing.Size(784, 20)
        Me.tbRaisonSociale.TabIndex = 3
        '
        'label
        '
        Me.label.Location = New System.Drawing.Point(8, 64)
        Me.label.Name = "label"
        Me.label.Size = New System.Drawing.Size(80, 16)
        Me.label.TabIndex = 95
        Me.label.Text = "Raison Sociale"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 24)
        Me.Label6.TabIndex = 94
        Me.Label6.Text = "Nom"
        '
        'cboModeReglement
        '
        Me.cboModeReglement.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboModeReglement.Location = New System.Drawing.Point(136, 400)
        Me.cboModeReglement.Name = "cboModeReglement"
        Me.cboModeReglement.Size = New System.Drawing.Size(288, 21)
        Me.cboModeReglement.TabIndex = 8
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Location = New System.Drawing.Point(16, 448)
        Me.Label20.Name = "Label20"
        Me.Label20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label20.Size = New System.Drawing.Size(97, 17)
        Me.Label20.TabIndex = 91
        Me.Label20.Text = "RIB"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.SystemColors.Control
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Location = New System.Drawing.Point(16, 400)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.Size = New System.Drawing.Size(105, 17)
        Me.Label19.TabIndex = 7
        Me.Label19.Text = "Mode de réglement"
        '
        'tbRib1
        '
        Me.tbRib1.AcceptsReturn = True
        Me.tbRib1.BackColor = System.Drawing.SystemColors.Window
        Me.tbRib1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRib1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRib1.Location = New System.Drawing.Point(136, 448)
        Me.tbRib1.MaxLength = 0
        Me.tbRib1.Name = "tbRib1"
        Me.tbRib1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRib1.Size = New System.Drawing.Size(49, 20)
        Me.tbRib1.TabIndex = 10
        Me.tbRib1.Text = "0"
        '
        'tbRib2
        '
        Me.tbRib2.AcceptsReturn = True
        Me.tbRib2.BackColor = System.Drawing.SystemColors.Window
        Me.tbRib2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRib2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRib2.Location = New System.Drawing.Point(192, 448)
        Me.tbRib2.MaxLength = 0
        Me.tbRib2.Name = "tbRib2"
        Me.tbRib2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRib2.Size = New System.Drawing.Size(57, 20)
        Me.tbRib2.TabIndex = 11
        Me.tbRib2.Text = "0"
        '
        'tbRib3
        '
        Me.tbRib3.AcceptsReturn = True
        Me.tbRib3.BackColor = System.Drawing.SystemColors.Window
        Me.tbRib3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRib3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRib3.Location = New System.Drawing.Point(256, 448)
        Me.tbRib3.MaxLength = 0
        Me.tbRib3.Name = "tbRib3"
        Me.tbRib3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRib3.Size = New System.Drawing.Size(193, 20)
        Me.tbRib3.TabIndex = 12
        Me.tbRib3.Text = "0"
        '
        'tbRib4
        '
        Me.tbRib4.AcceptsReturn = True
        Me.tbRib4.BackColor = System.Drawing.SystemColors.Window
        Me.tbRib4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbRib4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbRib4.Location = New System.Drawing.Point(456, 448)
        Me.tbRib4.MaxLength = 0
        Me.tbRib4.Name = "tbRib4"
        Me.tbRib4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbRib4.Size = New System.Drawing.Size(57, 20)
        Me.tbRib4.TabIndex = 13
        Me.tbRib4.Text = "0"
        '
        'tbNomClient
        '
        Me.tbNomClient.AcceptsReturn = True
        Me.tbNomClient.BackColor = System.Drawing.SystemColors.Window
        Me.tbNomClient.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbNomClient.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbNomClient.Location = New System.Drawing.Point(104, 40)
        Me.tbNomClient.MaxLength = 0
        Me.tbNomClient.Name = "tbNomClient"
        Me.tbNomClient.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbNomClient.Size = New System.Drawing.Size(784, 20)
        Me.tbNomClient.TabIndex = 2
        '
        'tpLignes
        '
        Me.tpLignes.Controls.Add(Me.laMtComm)
        Me.tpLignes.Controls.Add(Me.tbMtComm)
        Me.tpLignes.Controls.Add(Me.tbTotalTTC)
        Me.tpLignes.Controls.Add(Me.Label51)
        Me.tpLignes.Controls.Add(Me.cbSupprimerLigne)
        Me.tpLignes.Controls.Add(Me.cbModifierLigne)
        Me.tpLignes.Controls.Add(Me.tbTotalHT)
        Me.tpLignes.Controls.Add(Me.Label10)
        Me.tpLignes.Controls.Add(Me.cbAjouterLigne)
        Me.tpLignes.Controls.Add(Me.DataGridView1)
        Me.tpLignes.Controls.Add(Me.tbQteColisCmd)
        Me.tpLignes.Controls.Add(Me.Label61)
        Me.tpLignes.Controls.Add(Me.tbPoidsCmd)
        Me.tpLignes.Controls.Add(Me.Label40)
        Me.tpLignes.Controls.Add(Me.tbRecalcTotaux)
        Me.tpLignes.Location = New System.Drawing.Point(4, 22)
        Me.tpLignes.Name = "tpLignes"
        Me.tpLignes.Size = New System.Drawing.Size(992, 557)
        Me.tpLignes.TabIndex = 1
        Me.tpLignes.Text = "Lignes"
        Me.tpLignes.Visible = False
        '
        'laMtComm
        '
        Me.laMtComm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.laMtComm.AutoSize = True
        Me.laMtComm.DataBindings.Add(New System.Windows.Forms.Binding("Visible", Global.vini_app.My.MySettings.Default, "bCommVisible", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.laMtComm.Location = New System.Drawing.Point(789, 501)
        Me.laMtComm.Name = "laMtComm"
        Me.laMtComm.Size = New System.Drawing.Size(51, 13)
        Me.laMtComm.TabIndex = 68
        Me.laMtComm.Text = "Mt Comm"
        Me.laMtComm.Visible = Global.vini_app.My.MySettings.Default.bCommVisible
        '
        'tbMtComm
        '
        Me.tbMtComm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMtComm.DataBindings.Add(New System.Windows.Forms.Binding("Visible", Global.vini_app.My.MySettings.Default, "bCommVisible", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.tbMtComm.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "montantCommission", True))
        Me.tbMtComm.Enabled = False
        Me.tbMtComm.Location = New System.Drawing.Point(876, 498)
        Me.tbMtComm.Name = "tbMtComm"
        Me.tbMtComm.Size = New System.Drawing.Size(113, 20)
        Me.tbMtComm.TabIndex = 67
        Me.tbMtComm.Text = "0"
        Me.tbMtComm.Visible = Global.vini_app.My.MySettings.Default.bCommVisible
        '
        'm_bsrcCommande
        '
        Me.m_bsrcCommande.DataSource = GetType(vini_DB.Commande)
        '
        'tbTotalTTC
        '
        Me.tbTotalTTC.AcceptsReturn = True
        Me.tbTotalTTC.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTotalTTC.BackColor = System.Drawing.SystemColors.Window
        Me.tbTotalTTC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbTotalTTC.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "totalTTC", True))
        Me.tbTotalTTC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbTotalTTC.Location = New System.Drawing.Point(876, 474)
        Me.tbTotalTTC.MaxLength = 0
        Me.tbTotalTTC.Name = "tbTotalTTC"
        Me.tbTotalTTC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbTotalTTC.Size = New System.Drawing.Size(113, 20)
        Me.tbTotalTTC.TabIndex = 5
        Me.tbTotalTTC.Text = "0"
        '
        'Label51
        '
        Me.Label51.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label51.Location = New System.Drawing.Point(789, 477)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(64, 16)
        Me.Label51.TabIndex = 62
        Me.Label51.Text = "Total TTC"
        '
        'cbSupprimerLigne
        '
        Me.cbSupprimerLigne.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSupprimerLigne.BackColor = System.Drawing.SystemColors.Control
        Me.cbSupprimerLigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbSupprimerLigne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbSupprimerLigne.Location = New System.Drawing.Point(916, 524)
        Me.cbSupprimerLigne.Name = "cbSupprimerLigne"
        Me.cbSupprimerLigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbSupprimerLigne.Size = New System.Drawing.Size(73, 25)
        Me.cbSupprimerLigne.TabIndex = 3
        Me.cbSupprimerLigne.Text = "&Supprimer"
        Me.cbSupprimerLigne.UseVisualStyleBackColor = False
        '
        'cbModifierLigne
        '
        Me.cbModifierLigne.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbModifierLigne.BackColor = System.Drawing.SystemColors.Control
        Me.cbModifierLigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbModifierLigne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbModifierLigne.Location = New System.Drawing.Point(837, 524)
        Me.cbModifierLigne.Name = "cbModifierLigne"
        Me.cbModifierLigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbModifierLigne.Size = New System.Drawing.Size(73, 25)
        Me.cbModifierLigne.TabIndex = 2
        Me.cbModifierLigne.Text = "M&odifier "
        Me.cbModifierLigne.UseVisualStyleBackColor = False
        '
        'tbTotalHT
        '
        Me.tbTotalHT.AcceptsReturn = True
        Me.tbTotalHT.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTotalHT.BackColor = System.Drawing.SystemColors.Window
        Me.tbTotalHT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbTotalHT.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "totalHT", True))
        Me.tbTotalHT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbTotalHT.Location = New System.Drawing.Point(632, 474)
        Me.tbTotalHT.MaxLength = 0
        Me.tbTotalHT.Name = "tbTotalHT"
        Me.tbTotalHT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbTotalHT.Size = New System.Drawing.Size(113, 20)
        Me.tbTotalHT.TabIndex = 4
        Me.tbTotalHT.Text = "0"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(561, 476)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(65, 17)
        Me.Label10.TabIndex = 60
        Me.Label10.Text = "Total H.T."
        '
        'cbAjouterLigne
        '
        Me.cbAjouterLigne.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAjouterLigne.BackColor = System.Drawing.SystemColors.Control
        Me.cbAjouterLigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAjouterLigne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAjouterLigne.Location = New System.Drawing.Point(758, 524)
        Me.cbAjouterLigne.Name = "cbAjouterLigne"
        Me.cbAjouterLigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAjouterLigne.Size = New System.Drawing.Size(73, 25)
        Me.cbAjouterLigne.TabIndex = 1
        Me.cbAjouterLigne.Text = "A&jouter"
        Me.cbAjouterLigne.UseVisualStyleBackColor = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.NumDataGridViewTextBoxColumn, Me.ProduitCodeDataGridViewTextBoxColumn, Me.ProduitNomDataGridViewTextBoxColumn, Me.ProduitMilDataGridViewTextBoxColumn, Me.ProduitConditionnementDataGridViewTextBoxColumn, Me.ProduitCouleurDataGridViewTextBoxColumn, Me.QteCommandeDataGridViewTextBoxColumn, Me.QteLivDataGridViewTextBoxColumn, Me.QteFactDataGridViewTextBoxColumn, Me.BGratuitDataGridViewCheckBoxColumn, Me.PrixUDataGridViewTextBoxColumn, Me.PrixHTDataGridViewTextBoxColumn, Me.PrixTTCDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcLignesCommande
        Me.DataGridView1.Location = New System.Drawing.Point(8, 4)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidth = 10
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.Size = New System.Drawing.Size(981, 464)
        Me.DataGridView1.TabIndex = 69
        '
        'NumDataGridViewTextBoxColumn
        '
        Me.NumDataGridViewTextBoxColumn.DataPropertyName = "num"
        Me.NumDataGridViewTextBoxColumn.FillWeight = 38.57375!
        Me.NumDataGridViewTextBoxColumn.HeaderText = "Num"
        Me.NumDataGridViewTextBoxColumn.Name = "NumDataGridViewTextBoxColumn"
        Me.NumDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ProduitCodeDataGridViewTextBoxColumn
        '
        Me.ProduitCodeDataGridViewTextBoxColumn.DataPropertyName = "ProduitCode"
        DataGridViewCellStyle34.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ProduitCodeDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle34
        Me.ProduitCodeDataGridViewTextBoxColumn.FillWeight = 48.21718!
        Me.ProduitCodeDataGridViewTextBoxColumn.HeaderText = "Code Produit"
        Me.ProduitCodeDataGridViewTextBoxColumn.Name = "ProduitCodeDataGridViewTextBoxColumn"
        Me.ProduitCodeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ProduitNomDataGridViewTextBoxColumn
        '
        Me.ProduitNomDataGridViewTextBoxColumn.DataPropertyName = "ProduitNom"
        Me.ProduitNomDataGridViewTextBoxColumn.FillWeight = 385.7375!
        Me.ProduitNomDataGridViewTextBoxColumn.HeaderText = "Désignation Produit"
        Me.ProduitNomDataGridViewTextBoxColumn.Name = "ProduitNomDataGridViewTextBoxColumn"
        Me.ProduitNomDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ProduitMilDataGridViewTextBoxColumn
        '
        Me.ProduitMilDataGridViewTextBoxColumn.DataPropertyName = "ProduitMil"
        Me.ProduitMilDataGridViewTextBoxColumn.FillWeight = 48.21718!
        Me.ProduitMilDataGridViewTextBoxColumn.HeaderText = "Mill."
        Me.ProduitMilDataGridViewTextBoxColumn.Name = "ProduitMilDataGridViewTextBoxColumn"
        Me.ProduitMilDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ProduitConditionnementDataGridViewTextBoxColumn
        '
        Me.ProduitConditionnementDataGridViewTextBoxColumn.DataPropertyName = "ProduitConditionnement"
        Me.ProduitConditionnementDataGridViewTextBoxColumn.FillWeight = 38.57375!
        Me.ProduitConditionnementDataGridViewTextBoxColumn.HeaderText = "Cond."
        Me.ProduitConditionnementDataGridViewTextBoxColumn.Name = "ProduitConditionnementDataGridViewTextBoxColumn"
        Me.ProduitConditionnementDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ProduitCouleurDataGridViewTextBoxColumn
        '
        Me.ProduitCouleurDataGridViewTextBoxColumn.DataPropertyName = "ProduitCouleur"
        Me.ProduitCouleurDataGridViewTextBoxColumn.FillWeight = 46.0742!
        Me.ProduitCouleurDataGridViewTextBoxColumn.HeaderText = "Coul."
        Me.ProduitCouleurDataGridViewTextBoxColumn.Name = "ProduitCouleurDataGridViewTextBoxColumn"
        Me.ProduitCouleurDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QteCommandeDataGridViewTextBoxColumn
        '
        Me.QteCommandeDataGridViewTextBoxColumn.DataPropertyName = "qteCommande"
        DataGridViewCellStyle35.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.QteCommandeDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle35
        Me.QteCommandeDataGridViewTextBoxColumn.FillWeight = 53.57465!
        Me.QteCommandeDataGridViewTextBoxColumn.HeaderText = "Qte Comm"
        Me.QteCommandeDataGridViewTextBoxColumn.Name = "QteCommandeDataGridViewTextBoxColumn"
        Me.QteCommandeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QteLivDataGridViewTextBoxColumn
        '
        Me.QteLivDataGridViewTextBoxColumn.DataPropertyName = "qteLiv"
        DataGridViewCellStyle36.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.QteLivDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle36
        Me.QteLivDataGridViewTextBoxColumn.FillWeight = 53.57465!
        Me.QteLivDataGridViewTextBoxColumn.HeaderText = "Qte Liv"
        Me.QteLivDataGridViewTextBoxColumn.Name = "QteLivDataGridViewTextBoxColumn"
        Me.QteLivDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QteFactDataGridViewTextBoxColumn
        '
        Me.QteFactDataGridViewTextBoxColumn.DataPropertyName = "qteFact"
        DataGridViewCellStyle37.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.QteFactDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle37
        Me.QteFactDataGridViewTextBoxColumn.FillWeight = 53.57465!
        Me.QteFactDataGridViewTextBoxColumn.HeaderText = "Qte Fact"
        Me.QteFactDataGridViewTextBoxColumn.Name = "QteFactDataGridViewTextBoxColumn"
        Me.QteFactDataGridViewTextBoxColumn.ReadOnly = True
        '
        'BGratuitDataGridViewCheckBoxColumn
        '
        Me.BGratuitDataGridViewCheckBoxColumn.DataPropertyName = "bGratuit"
        Me.BGratuitDataGridViewCheckBoxColumn.FillWeight = 42.85972!
        Me.BGratuitDataGridViewCheckBoxColumn.HeaderText = "Gratuit"
        Me.BGratuitDataGridViewCheckBoxColumn.Name = "BGratuitDataGridViewCheckBoxColumn"
        Me.BGratuitDataGridViewCheckBoxColumn.ReadOnly = True
        '
        'PrixUDataGridViewTextBoxColumn
        '
        Me.PrixUDataGridViewTextBoxColumn.DataPropertyName = "prixU"
        DataGridViewCellStyle38.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle38.Format = "C2"
        Me.PrixUDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle38
        Me.PrixUDataGridViewTextBoxColumn.FillWeight = 60.0036!
        Me.PrixUDataGridViewTextBoxColumn.HeaderText = "Prix U"
        Me.PrixUDataGridViewTextBoxColumn.Name = "PrixUDataGridViewTextBoxColumn"
        Me.PrixUDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PrixHTDataGridViewTextBoxColumn
        '
        Me.PrixHTDataGridViewTextBoxColumn.DataPropertyName = "prixHT"
        DataGridViewCellStyle39.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle39.Format = "C2"
        Me.PrixHTDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle39
        Me.PrixHTDataGridViewTextBoxColumn.FillWeight = 75.00452!
        Me.PrixHTDataGridViewTextBoxColumn.HeaderText = "Montant HT"
        Me.PrixHTDataGridViewTextBoxColumn.Name = "PrixHTDataGridViewTextBoxColumn"
        Me.PrixHTDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PrixTTCDataGridViewTextBoxColumn
        '
        Me.PrixTTCDataGridViewTextBoxColumn.DataPropertyName = "prixTTC"
        DataGridViewCellStyle40.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle40.Format = "C2"
        Me.PrixTTCDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle40
        Me.PrixTTCDataGridViewTextBoxColumn.FillWeight = 75.00452!
        Me.PrixTTCDataGridViewTextBoxColumn.HeaderText = "Montant TTC"
        Me.PrixTTCDataGridViewTextBoxColumn.Name = "PrixTTCDataGridViewTextBoxColumn"
        Me.PrixTTCDataGridViewTextBoxColumn.ReadOnly = True
        '
        'm_bsrcLignesCommande
        '
        Me.m_bsrcLignesCommande.DataSource = GetType(vini_DB.LgCommande)
        '
        'tbQteColisCmd
        '
        Me.tbQteColisCmd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tbQteColisCmd.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "qteColis", True))
        Me.tbQteColisCmd.Enabled = False
        Me.tbQteColisCmd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbQteColisCmd.Location = New System.Drawing.Point(74, 477)
        Me.tbQteColisCmd.Name = "tbQteColisCmd"
        Me.tbQteColisCmd.Size = New System.Drawing.Size(100, 20)
        Me.tbQteColisCmd.TabIndex = 66
        '
        'Label61
        '
        Me.Label61.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label61.Location = New System.Drawing.Point(5, 478)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(64, 23)
        Me.Label61.TabIndex = 65
        Me.Label61.Text = "QteColis"
        '
        'tbPoidsCmd
        '
        Me.tbPoidsCmd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tbPoidsCmd.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "poids", True))
        Me.tbPoidsCmd.Enabled = False
        Me.tbPoidsCmd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPoidsCmd.Location = New System.Drawing.Point(74, 501)
        Me.tbPoidsCmd.Name = "tbPoidsCmd"
        Me.tbPoidsCmd.Size = New System.Drawing.Size(100, 20)
        Me.tbPoidsCmd.TabIndex = 64
        '
        'Label40
        '
        Me.Label40.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label40.Location = New System.Drawing.Point(3, 501)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(64, 24)
        Me.Label40.TabIndex = 63
        Me.Label40.Text = "Poids"
        '
        'tbRecalcTotaux
        '
        Me.tbRecalcTotaux.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbRecalcTotaux.Location = New System.Drawing.Point(465, 472)
        Me.tbRecalcTotaux.Name = "tbRecalcTotaux"
        Me.tbRecalcTotaux.Size = New System.Drawing.Size(75, 23)
        Me.tbRecalcTotaux.TabIndex = 6
        Me.tbRecalcTotaux.Text = "Re&Calcul"
        '
        'tpCommentaires
        '
        Me.tpCommentaires.Controls.Add(Me.tbCommentaireFacturation)
        Me.tpCommentaires.Controls.Add(Me.tbCommentaireLivraison)
        Me.tpCommentaires.Controls.Add(Me.tbCommentaireCommande)
        Me.tpCommentaires.Controls.Add(Me.Label26)
        Me.tpCommentaires.Controls.Add(Me.Label25)
        Me.tpCommentaires.Controls.Add(Me.Label21)
        Me.tpCommentaires.Location = New System.Drawing.Point(4, 22)
        Me.tpCommentaires.Name = "tpCommentaires"
        Me.tpCommentaires.Size = New System.Drawing.Size(992, 557)
        Me.tpCommentaires.TabIndex = 3
        Me.tpCommentaires.Text = "Commentaires"
        Me.tpCommentaires.Visible = False
        '
        'tbCommentaireFacturation
        '
        Me.tbCommentaireFacturation.Location = New System.Drawing.Point(96, 264)
        Me.tbCommentaireFacturation.Name = "tbCommentaireFacturation"
        Me.tbCommentaireFacturation.Size = New System.Drawing.Size(632, 120)
        Me.tbCommentaireFacturation.TabIndex = 2
        Me.tbCommentaireFacturation.Text = ""
        '
        'tbCommentaireLivraison
        '
        Me.tbCommentaireLivraison.Location = New System.Drawing.Point(96, 136)
        Me.tbCommentaireLivraison.Name = "tbCommentaireLivraison"
        Me.tbCommentaireLivraison.Size = New System.Drawing.Size(632, 120)
        Me.tbCommentaireLivraison.TabIndex = 1
        Me.tbCommentaireLivraison.Text = ""
        '
        'tbCommentaireCommande
        '
        Me.tbCommentaireCommande.Location = New System.Drawing.Point(96, 8)
        Me.tbCommentaireCommande.Name = "tbCommentaireCommande"
        Me.tbCommentaireCommande.Size = New System.Drawing.Size(632, 120)
        Me.tbCommentaireCommande.TabIndex = 0
        Me.tbCommentaireCommande.Text = ""
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.SystemColors.Control
        Me.Label26.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label26.Location = New System.Drawing.Point(8, 272)
        Me.Label26.Name = "Label26"
        Me.Label26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label26.Size = New System.Drawing.Size(64, 25)
        Me.Label26.TabIndex = 74
        Me.Label26.Text = "Facturation"
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.SystemColors.Control
        Me.Label25.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label25.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label25.Location = New System.Drawing.Point(8, 144)
        Me.Label25.Name = "Label25"
        Me.Label25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label25.Size = New System.Drawing.Size(72, 25)
        Me.Label25.TabIndex = 73
        Me.Label25.Text = "Livraison"
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.SystemColors.Control
        Me.Label21.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label21.Location = New System.Drawing.Point(8, 16)
        Me.Label21.Name = "Label21"
        Me.Label21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label21.Size = New System.Drawing.Size(80, 17)
        Me.Label21.TabIndex = 72
        Me.Label21.Text = "Commande"
        '
        'tpValidCmd
        '
        Me.tpValidCmd.Controls.Add(Me.crwDetailCommandeClient)
        Me.tpValidCmd.Controls.Add(Me.dtDateLivraisonPrevue)
        Me.tpValidCmd.Controls.Add(Me.label50)
        Me.tpValidCmd.Controls.Add(Me.cbVisuDetailCommande)
        Me.tpValidCmd.Controls.Add(Me.tbComValid)
        Me.tpValidCmd.Controls.Add(Me.tbNumFaxBC)
        Me.tpValidCmd.Controls.Add(Me.ckCmdValide)
        Me.tpValidCmd.Controls.Add(Me.Label28)
        Me.tpValidCmd.Controls.Add(Me.dtDateValidation)
        Me.tpValidCmd.Controls.Add(Me.Label27)
        Me.tpValidCmd.Controls.Add(Me.tbNumFaxValidation)
        Me.tpValidCmd.Controls.Add(Me.cbFax)
        Me.tpValidCmd.Location = New System.Drawing.Point(4, 22)
        Me.tpValidCmd.Name = "tpValidCmd"
        Me.tpValidCmd.Size = New System.Drawing.Size(992, 557)
        Me.tpValidCmd.TabIndex = 4
        Me.tpValidCmd.Text = "Validation de commande"
        '
        'crwDetailCommandeClient
        '
        Me.crwDetailCommandeClient.ActiveViewIndex = -1
        Me.crwDetailCommandeClient.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crwDetailCommandeClient.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crwDetailCommandeClient.Cursor = System.Windows.Forms.Cursors.Default
        Me.crwDetailCommandeClient.DisplayStatusBar = False
        Me.crwDetailCommandeClient.Location = New System.Drawing.Point(9, 8)
        Me.crwDetailCommandeClient.Name = "crwDetailCommandeClient"
        Me.crwDetailCommandeClient.Size = New System.Drawing.Size(603, 535)
        Me.crwDetailCommandeClient.TabIndex = 26
        Me.crwDetailCommandeClient.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'dtDateLivraisonPrevue
        '
        Me.dtDateLivraisonPrevue.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtDateLivraisonPrevue.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateLivraisonPrevue.Location = New System.Drawing.Point(744, 128)
        Me.dtDateLivraisonPrevue.Name = "dtDateLivraisonPrevue"
        Me.dtDateLivraisonPrevue.Size = New System.Drawing.Size(96, 20)
        Me.dtDateLivraisonPrevue.TabIndex = 4
        '
        'label50
        '
        Me.label50.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.label50.Location = New System.Drawing.Point(624, 128)
        Me.label50.Name = "label50"
        Me.label50.Size = New System.Drawing.Size(104, 32)
        Me.label50.TabIndex = 25
        Me.label50.Text = "Date de livraison prévue"
        '
        'cbVisuDetailCommande
        '
        Me.cbVisuDetailCommande.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbVisuDetailCommande.Location = New System.Drawing.Point(632, 8)
        Me.cbVisuDetailCommande.Name = "cbVisuDetailCommande"
        Me.cbVisuDetailCommande.Size = New System.Drawing.Size(96, 23)
        Me.cbVisuDetailCommande.TabIndex = 0
        Me.cbVisuDetailCommande.Text = "A&fficher"
        '
        'tbComValid
        '
        Me.tbComValid.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbComValid.Location = New System.Drawing.Point(624, 192)
        Me.tbComValid.Name = "tbComValid"
        Me.tbComValid.Size = New System.Drawing.Size(360, 104)
        Me.tbComValid.TabIndex = 5
        Me.tbComValid.Text = ""
        '
        'tbNumFaxBC
        '
        Me.tbNumFaxBC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbNumFaxBC.Location = New System.Drawing.Point(624, 40)
        Me.tbNumFaxBC.Name = "tbNumFaxBC"
        Me.tbNumFaxBC.Size = New System.Drawing.Size(112, 24)
        Me.tbNumFaxBC.TabIndex = 20
        Me.tbNumFaxBC.Text = "Numéro de Fax"
        '
        'ckCmdValide
        '
        Me.ckCmdValide.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ckCmdValide.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckCmdValide.Location = New System.Drawing.Point(624, 72)
        Me.ckCmdValide.Name = "ckCmdValide"
        Me.ckCmdValide.Size = New System.Drawing.Size(136, 16)
        Me.ckCmdValide.TabIndex = 2
        Me.ckCmdValide.Text = "Commande Validée"
        '
        'Label28
        '
        Me.Label28.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label28.Location = New System.Drawing.Point(624, 168)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(104, 24)
        Me.Label28.TabIndex = 15
        Me.Label28.Text = "Commentaires"
        '
        'dtDateValidation
        '
        Me.dtDateValidation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtDateValidation.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateValidation.Location = New System.Drawing.Point(744, 96)
        Me.dtDateValidation.Name = "dtDateValidation"
        Me.dtDateValidation.Size = New System.Drawing.Size(96, 20)
        Me.dtDateValidation.TabIndex = 3
        '
        'Label27
        '
        Me.Label27.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label27.Location = New System.Drawing.Point(624, 96)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(104, 16)
        Me.Label27.TabIndex = 13
        Me.Label27.Text = "Date de validation"
        '
        'tbNumFaxValidation
        '
        Me.tbNumFaxValidation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbNumFaxValidation.Location = New System.Drawing.Point(744, 40)
        Me.tbNumFaxValidation.Name = "tbNumFaxValidation"
        Me.tbNumFaxValidation.Size = New System.Drawing.Size(120, 20)
        Me.tbNumFaxValidation.TabIndex = 1
        '
        'cbFax
        '
        Me.cbFax.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbFax.Location = New System.Drawing.Point(880, 40)
        Me.cbFax.Name = "cbFax"
        Me.cbFax.Size = New System.Drawing.Size(64, 23)
        Me.cbFax.TabIndex = 7
        Me.cbFax.Text = "Fa&x"
        '
        'tpBL
        '
        Me.tpBL.Controls.Add(Me.crwBL)
        Me.tpBL.Controls.Add(Me.tbFaxTRP)
        Me.tpBL.Controls.Add(Me.tbMailPLTF)
        Me.tpBL.Controls.Add(Me.cbFaxerBLTransporteur)
        Me.tpBL.Controls.Add(Me.grpInfFactureTRP)
        Me.tpBL.Controls.Add(Me.tbPiedPageBL)
        Me.tpBL.Controls.Add(Me.Label39)
        Me.tpBL.Controls.Add(Me.cbVisuBL)
        Me.tpBL.Controls.Add(Me.grpTransporteur)
        Me.tpBL.Controls.Add(Me.dtDateLivraison)
        Me.tpBL.Controls.Add(Me.Label41)
        Me.tpBL.Controls.Add(Me.dtDateEnlev)
        Me.tpBL.Controls.Add(Me.Label42)
        Me.tpBL.Controls.Add(Me.cbMailBLPLTFRM)
        Me.tpBL.Location = New System.Drawing.Point(4, 22)
        Me.tpBL.Name = "tpBL"
        Me.tpBL.Size = New System.Drawing.Size(992, 557)
        Me.tpBL.TabIndex = 7
        Me.tpBL.Text = "Bon de Livraison"
        '
        'crwBL
        '
        Me.crwBL.ActiveViewIndex = -1
        Me.crwBL.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crwBL.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crwBL.Cursor = System.Windows.Forms.Cursors.Default
        Me.crwBL.DisplayStatusBar = False
        Me.crwBL.Location = New System.Drawing.Point(9, 4)
        Me.crwBL.Name = "crwBL"
        Me.crwBL.Size = New System.Drawing.Size(481, 550)
        Me.crwBL.TabIndex = 148
        Me.crwBL.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'tbFaxTRP
        '
        Me.tbFaxTRP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbFaxTRP.Enabled = False
        Me.tbFaxTRP.Location = New System.Drawing.Point(496, 528)
        Me.tbFaxTRP.Name = "tbFaxTRP"
        Me.tbFaxTRP.Size = New System.Drawing.Size(152, 20)
        Me.tbFaxTRP.TabIndex = 147
        '
        'tbMailPLTF
        '
        Me.tbMailPLTF.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMailPLTF.Location = New System.Drawing.Point(496, 496)
        Me.tbMailPLTF.Name = "tbMailPLTF"
        Me.tbMailPLTF.Size = New System.Drawing.Size(152, 20)
        Me.tbMailPLTF.TabIndex = 146
        '
        'cbFaxerBLTransporteur
        '
        Me.cbFaxerBLTransporteur.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbFaxerBLTransporteur.Location = New System.Drawing.Point(656, 528)
        Me.cbFaxerBLTransporteur.Name = "cbFaxerBLTransporteur"
        Me.cbFaxerBLTransporteur.Size = New System.Drawing.Size(128, 24)
        Me.cbFaxerBLTransporteur.TabIndex = 145
        Me.cbFaxerBLTransporteur.Text = "Faxer &Transporteur"
        '
        'grpInfFactureTRP
        '
        Me.grpInfFactureTRP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpInfFactureTRP.Controls.Add(Me.cbCalcPoidsColis)
        Me.grpInfFactureTRP.Controls.Add(Me.Label60)
        Me.grpInfFactureTRP.Controls.Add(Me.Label57)
        Me.grpInfFactureTRP.Controls.Add(Me.cbCalcMontantTransport)
        Me.grpInfFactureTRP.Controls.Add(Me.Label59)
        Me.grpInfFactureTRP.Controls.Add(Me.Label58)
        Me.grpInfFactureTRP.Controls.Add(Me.tbMontantTransport)
        Me.grpInfFactureTRP.Controls.Add(Me.tbPUPallNonPrep)
        Me.grpInfFactureTRP.Controls.Add(Me.tbPUPallPrep)
        Me.grpInfFactureTRP.Controls.Add(Me.Label56)
        Me.grpInfFactureTRP.Controls.Add(Me.Label55)
        Me.grpInfFactureTRP.Controls.Add(Me.tbQtePallNonPrep)
        Me.grpInfFactureTRP.Controls.Add(Me.tbQtePallPrep)
        Me.grpInfFactureTRP.Controls.Add(Me.Label54)
        Me.grpInfFactureTRP.Controls.Add(Me.Label53)
        Me.grpInfFactureTRP.Controls.Add(Me.tbPoids)
        Me.grpInfFactureTRP.Controls.Add(Me.Label52)
        Me.grpInfFactureTRP.Controls.Add(Me.tbQteColis)
        Me.grpInfFactureTRP.Controls.Add(Me.Label2)
        Me.grpInfFactureTRP.Location = New System.Drawing.Point(496, 224)
        Me.grpInfFactureTRP.Name = "grpInfFactureTRP"
        Me.grpInfFactureTRP.Size = New System.Drawing.Size(496, 120)
        Me.grpInfFactureTRP.TabIndex = 3
        Me.grpInfFactureTRP.TabStop = False
        Me.grpInfFactureTRP.Text = "Transport Facturé"
        '
        'cbCalcPoidsColis
        '
        Me.cbCalcPoidsColis.Location = New System.Drawing.Point(384, 16)
        Me.cbCalcPoidsColis.Name = "cbCalcPoidsColis"
        Me.cbCalcPoidsColis.Size = New System.Drawing.Size(104, 24)
        Me.cbCalcPoidsColis.TabIndex = 2
        Me.cbCalcPoidsColis.Text = "Calcul"
        '
        'Label60
        '
        Me.Label60.Location = New System.Drawing.Point(260, 42)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(80, 16)
        Me.Label60.TabIndex = 182
        Me.Label60.Text = "P.U."
        Me.Label60.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label57
        '
        Me.Label57.Location = New System.Drawing.Point(172, 42)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(56, 16)
        Me.Label57.TabIndex = 181
        Me.Label57.Text = "Qte"
        Me.Label57.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbCalcMontantTransport
        '
        Me.cbCalcMontantTransport.Location = New System.Drawing.Point(364, 66)
        Me.cbCalcMontantTransport.Name = "cbCalcMontantTransport"
        Me.cbCalcMontantTransport.Size = New System.Drawing.Size(16, 23)
        Me.cbCalcMontantTransport.TabIndex = 7
        Me.cbCalcMontantTransport.Text = "="
        '
        'Label59
        '
        Me.Label59.Location = New System.Drawing.Point(348, 82)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(8, 16)
        Me.Label59.TabIndex = 7
        Me.Label59.Text = "/"
        '
        'Label58
        '
        Me.Label58.Location = New System.Drawing.Point(348, 66)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(8, 23)
        Me.Label58.TabIndex = 7
        Me.Label58.Text = "\"
        '
        'tbMontantTransport
        '
        Me.tbMontantTransport.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "montantTransport", True))
        Me.tbMontantTransport.Location = New System.Drawing.Point(388, 66)
        Me.tbMontantTransport.Name = "tbMontantTransport"
        Me.tbMontantTransport.Size = New System.Drawing.Size(80, 20)
        Me.tbMontantTransport.TabIndex = 8
        Me.tbMontantTransport.Text = "0"
        '
        'tbPUPallNonPrep
        '
        Me.tbPUPallNonPrep.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "puPalettesNonPreparees", True))
        Me.tbPUPallNonPrep.Location = New System.Drawing.Point(260, 82)
        Me.tbPUPallNonPrep.Name = "tbPUPallNonPrep"
        Me.tbPUPallNonPrep.Size = New System.Drawing.Size(80, 20)
        Me.tbPUPallNonPrep.TabIndex = 6
        Me.tbPUPallNonPrep.Text = "0"
        '
        'tbPUPallPrep
        '
        Me.tbPUPallPrep.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "puPalettesPreparees", True))
        Me.tbPUPallPrep.Location = New System.Drawing.Point(260, 58)
        Me.tbPUPallPrep.Name = "tbPUPallPrep"
        Me.tbPUPallPrep.Size = New System.Drawing.Size(80, 20)
        Me.tbPUPallPrep.TabIndex = 4
        Me.tbPUPallPrep.Text = "0"
        '
        'Label56
        '
        Me.Label56.Location = New System.Drawing.Point(244, 82)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(16, 16)
        Me.Label56.TabIndex = 174
        Me.Label56.Text = "X"
        '
        'Label55
        '
        Me.Label55.Location = New System.Drawing.Point(244, 66)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(16, 16)
        Me.Label55.TabIndex = 173
        Me.Label55.Text = "X"
        '
        'tbQtePallNonPrep
        '
        Me.tbQtePallNonPrep.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "qtePalettesNonPreparees", True))
        Me.tbQtePallNonPrep.Location = New System.Drawing.Point(172, 82)
        Me.tbQtePallNonPrep.Name = "tbQtePallNonPrep"
        Me.tbQtePallNonPrep.Size = New System.Drawing.Size(64, 20)
        Me.tbQtePallNonPrep.TabIndex = 5
        Me.tbQtePallNonPrep.Text = "0"
        '
        'tbQtePallPrep
        '
        Me.tbQtePallPrep.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "qtePalettesPreparees", True))
        Me.tbQtePallPrep.Location = New System.Drawing.Point(172, 58)
        Me.tbQtePallPrep.Name = "tbQtePallPrep"
        Me.tbQtePallPrep.Size = New System.Drawing.Size(64, 20)
        Me.tbQtePallPrep.TabIndex = 3
        Me.tbQtePallPrep.Text = "0"
        '
        'Label54
        '
        Me.Label54.Location = New System.Drawing.Point(28, 82)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(128, 16)
        Me.Label54.TabIndex = 170
        Me.Label54.Text = "Pallettes non préparées"
        '
        'Label53
        '
        Me.Label53.Location = New System.Drawing.Point(28, 58)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(128, 16)
        Me.Label53.TabIndex = 169
        Me.Label53.Text = "Pallettes préparées"
        '
        'tbPoids
        '
        Me.tbPoids.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "poids", True))
        Me.tbPoids.Location = New System.Drawing.Point(260, 18)
        Me.tbPoids.Name = "tbPoids"
        Me.tbPoids.Size = New System.Drawing.Size(104, 20)
        Me.tbPoids.TabIndex = 1
        Me.tbPoids.Text = "0"
        '
        'Label52
        '
        Me.Label52.Location = New System.Drawing.Point(204, 18)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(40, 23)
        Me.Label52.TabIndex = 167
        Me.Label52.Text = "Poids"
        '
        'tbQteColis
        '
        Me.tbQteColis.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "qteColis", True))
        Me.tbQteColis.Location = New System.Drawing.Point(100, 18)
        Me.tbQteColis.Name = "tbQteColis"
        Me.tbQteColis.Size = New System.Drawing.Size(88, 20)
        Me.tbQteColis.TabIndex = 0
        Me.tbQteColis.Text = "0"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(28, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 165
        Me.Label2.Text = "Nbre Colis"
        '
        'tbPiedPageBL
        '
        Me.tbPiedPageBL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPiedPageBL.Location = New System.Drawing.Point(576, 352)
        Me.tbPiedPageBL.Name = "tbPiedPageBL"
        Me.tbPiedPageBL.Size = New System.Drawing.Size(408, 104)
        Me.tbPiedPageBL.TabIndex = 4
        Me.tbPiedPageBL.Text = ""
        '
        'Label39
        '
        Me.Label39.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label39.Location = New System.Drawing.Point(504, 352)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(56, 32)
        Me.Label39.TabIndex = 144
        Me.Label39.Text = "Pied de page"
        '
        'cbVisuBL
        '
        Me.cbVisuBL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbVisuBL.Location = New System.Drawing.Point(496, 8)
        Me.cbVisuBL.Name = "cbVisuBL"
        Me.cbVisuBL.Size = New System.Drawing.Size(104, 24)
        Me.cbVisuBL.TabIndex = 1
        Me.cbVisuBL.Text = "A&fficher"
        '
        'grpTransporteur
        '
        Me.grpTransporteur.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpTransporteur.Controls.Add(Me.cboCodeTRP)
        Me.grpTransporteur.Controls.Add(Me.Label62)
        Me.grpTransporteur.Controls.Add(Me.Label49)
        Me.grpTransporteur.Controls.Add(Me.tbTrpEmail)
        Me.grpTransporteur.Controls.Add(Me.Label48)
        Me.grpTransporteur.Controls.Add(Me.tbTrpFax)
        Me.grpTransporteur.Controls.Add(Me.tbTrpPort)
        Me.grpTransporteur.Controls.Add(Me.Label47)
        Me.grpTransporteur.Controls.Add(Me.tbTrpTel)
        Me.grpTransporteur.Controls.Add(Me.Label46)
        Me.grpTransporteur.Controls.Add(Me.Label45)
        Me.grpTransporteur.Controls.Add(Me.Label44)
        Me.grpTransporteur.Controls.Add(Me.Label43)
        Me.grpTransporteur.Controls.Add(Me.Label37)
        Me.grpTransporteur.Controls.Add(Me.tbTrpVille)
        Me.grpTransporteur.Controls.Add(Me.tbTrpCP)
        Me.grpTransporteur.Controls.Add(Me.tbTrpRue2)
        Me.grpTransporteur.Controls.Add(Me.tbTrpRue1)
        Me.grpTransporteur.Controls.Add(Me.tbTrpNom)
        Me.grpTransporteur.Location = New System.Drawing.Point(496, 32)
        Me.grpTransporteur.Name = "grpTransporteur"
        Me.grpTransporteur.Size = New System.Drawing.Size(496, 184)
        Me.grpTransporteur.TabIndex = 2
        Me.grpTransporteur.TabStop = False
        Me.grpTransporteur.Text = "Transporteur"
        '
        'cboCodeTRP
        '
        Me.cboCodeTRP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCodeTRP.Location = New System.Drawing.Point(72, 16)
        Me.cboCodeTRP.Name = "cboCodeTRP"
        Me.cboCodeTRP.Size = New System.Drawing.Size(240, 21)
        Me.cboCodeTRP.Sorted = True
        Me.cboCodeTRP.TabIndex = 143
        '
        'Label62
        '
        Me.Label62.Location = New System.Drawing.Point(8, 16)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(64, 23)
        Me.Label62.TabIndex = 142
        Me.Label62.Text = "Code TRP"
        '
        'Label49
        '
        Me.Label49.Location = New System.Drawing.Point(312, 160)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(64, 16)
        Me.Label49.TabIndex = 141
        Me.Label49.Text = "Email"
        '
        'tbTrpEmail
        '
        Me.tbTrpEmail.Location = New System.Drawing.Point(384, 160)
        Me.tbTrpEmail.Name = "tbTrpEmail"
        Me.tbTrpEmail.Size = New System.Drawing.Size(104, 20)
        Me.tbTrpEmail.TabIndex = 8
        '
        'Label48
        '
        Me.Label48.Location = New System.Drawing.Point(312, 136)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(56, 16)
        Me.Label48.TabIndex = 139
        Me.Label48.Text = "Fax"
        '
        'tbTrpFax
        '
        Me.tbTrpFax.Location = New System.Drawing.Point(384, 136)
        Me.tbTrpFax.Name = "tbTrpFax"
        Me.tbTrpFax.Size = New System.Drawing.Size(104, 20)
        Me.tbTrpFax.TabIndex = 7
        '
        'tbTrpPort
        '
        Me.tbTrpPort.Location = New System.Drawing.Point(160, 160)
        Me.tbTrpPort.Name = "tbTrpPort"
        Me.tbTrpPort.Size = New System.Drawing.Size(96, 20)
        Me.tbTrpPort.TabIndex = 6
        '
        'Label47
        '
        Me.Label47.Location = New System.Drawing.Point(72, 160)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(56, 16)
        Me.Label47.TabIndex = 7
        Me.Label47.Text = "Port"
        '
        'tbTrpTel
        '
        Me.tbTrpTel.Location = New System.Drawing.Point(160, 136)
        Me.tbTrpTel.Name = "tbTrpTel"
        Me.tbTrpTel.Size = New System.Drawing.Size(96, 20)
        Me.tbTrpTel.TabIndex = 5
        '
        'Label46
        '
        Me.Label46.Location = New System.Drawing.Point(72, 136)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(56, 16)
        Me.Label46.TabIndex = 5
        Me.Label46.Text = "Tel"
        '
        'Label45
        '
        Me.Label45.Location = New System.Drawing.Point(8, 112)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(56, 23)
        Me.Label45.TabIndex = 133
        Me.Label45.Text = "CP/Ville"
        '
        'Label44
        '
        Me.Label44.Location = New System.Drawing.Point(8, 88)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(56, 16)
        Me.Label44.TabIndex = 132
        Me.Label44.Text = "Rue2"
        '
        'Label43
        '
        Me.Label43.Location = New System.Drawing.Point(8, 64)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(56, 16)
        Me.Label43.TabIndex = 131
        Me.Label43.Text = "Rue1"
        '
        'Label37
        '
        Me.Label37.Location = New System.Drawing.Point(8, 40)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(56, 16)
        Me.Label37.TabIndex = 130
        Me.Label37.Text = "Nom"
        '
        'tbTrpVille
        '
        Me.tbTrpVille.Location = New System.Drawing.Point(152, 112)
        Me.tbTrpVille.Name = "tbTrpVille"
        Me.tbTrpVille.Size = New System.Drawing.Size(336, 20)
        Me.tbTrpVille.TabIndex = 4
        '
        'tbTrpCP
        '
        Me.tbTrpCP.Location = New System.Drawing.Point(72, 112)
        Me.tbTrpCP.Name = "tbTrpCP"
        Me.tbTrpCP.Size = New System.Drawing.Size(72, 20)
        Me.tbTrpCP.TabIndex = 3
        '
        'tbTrpRue2
        '
        Me.tbTrpRue2.Location = New System.Drawing.Point(72, 88)
        Me.tbTrpRue2.Name = "tbTrpRue2"
        Me.tbTrpRue2.Size = New System.Drawing.Size(416, 20)
        Me.tbTrpRue2.TabIndex = 2
        '
        'tbTrpRue1
        '
        Me.tbTrpRue1.Location = New System.Drawing.Point(72, 64)
        Me.tbTrpRue1.Name = "tbTrpRue1"
        Me.tbTrpRue1.Size = New System.Drawing.Size(416, 20)
        Me.tbTrpRue1.TabIndex = 1
        '
        'tbTrpNom
        '
        Me.tbTrpNom.Location = New System.Drawing.Point(72, 40)
        Me.tbTrpNom.Name = "tbTrpNom"
        Me.tbTrpNom.Size = New System.Drawing.Size(416, 20)
        Me.tbTrpNom.TabIndex = 0
        '
        'dtDateLivraison
        '
        Me.dtDateLivraison.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtDateLivraison.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateLivraison.Location = New System.Drawing.Point(872, 464)
        Me.dtDateLivraison.Name = "dtDateLivraison"
        Me.dtDateLivraison.Size = New System.Drawing.Size(96, 20)
        Me.dtDateLivraison.TabIndex = 6
        '
        'Label41
        '
        Me.Label41.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label41.Location = New System.Drawing.Point(680, 464)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(168, 24)
        Me.Label41.TabIndex = 118
        Me.Label41.Text = "Date de livraison souhaitée"
        '
        'dtDateEnlev
        '
        Me.dtDateEnlev.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtDateEnlev.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateEnlev.Location = New System.Drawing.Point(584, 464)
        Me.dtDateEnlev.Name = "dtDateEnlev"
        Me.dtDateEnlev.Size = New System.Drawing.Size(88, 20)
        Me.dtDateEnlev.TabIndex = 5
        '
        'Label42
        '
        Me.Label42.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label42.Location = New System.Drawing.Point(496, 464)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(88, 24)
        Me.Label42.TabIndex = 116
        Me.Label42.Text = "Date enlèvement"
        '
        'cbMailBLPLTFRM
        '
        Me.cbMailBLPLTFRM.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbMailBLPLTFRM.Location = New System.Drawing.Point(656, 496)
        Me.cbMailBLPLTFRM.Name = "cbMailBLPLTFRM"
        Me.cbMailBLPLTFRM.Size = New System.Drawing.Size(128, 24)
        Me.cbMailBLPLTFRM.TabIndex = 8
        Me.cbMailBLPLTFRM.Text = "Envo&yer plateforme"
        '
        'tpRetourBL
        '
        Me.tpRetourBL.Controls.Add(Me.GroupBox1)
        Me.tpRetourBL.Controls.Add(Me.DataGridView2)
        Me.tpRetourBL.Controls.Add(Me.tbCoutTransport)
        Me.tpRetourBL.Controls.Add(Me.Label65)
        Me.tpRetourBL.Controls.Add(Me.Label64)
        Me.tpRetourBL.Controls.Add(Me.Label63)
        Me.tpRetourBL.Controls.Add(Me.tbRefFactTRP)
        Me.tpRetourBL.Controls.Add(Me.tbLettreVoiture)
        Me.tpRetourBL.Controls.Add(Me.cbAnnulerLivraison)
        Me.tpRetourBL.Controls.Add(Me.dtDateEnlevementReelle)
        Me.tpRetourBL.Controls.Add(Me.Label31)
        Me.tpRetourBL.Controls.Add(Me.tbRefBL)
        Me.tpRetourBL.Controls.Add(Me.Label23)
        Me.tpRetourBL.Controls.Add(Me.dtDateLivraisonReelle)
        Me.tpRetourBL.Controls.Add(Me.Label22)
        Me.tpRetourBL.Controls.Add(Me.cbBLToutOK)
        Me.tpRetourBL.Location = New System.Drawing.Point(4, 22)
        Me.tpRetourBL.Name = "tpRetourBL"
        Me.tpRetourBL.Size = New System.Drawing.Size(992, 557)
        Me.tpRetourBL.TabIndex = 6
        Me.tpRetourBL.Text = "Retour de Livraison"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.cbCalcPoidsColis_RL)
        Me.GroupBox1.Controls.Add(Me.Label66)
        Me.GroupBox1.Controls.Add(Me.Label67)
        Me.GroupBox1.Controls.Add(Me.cbCalcMontantTransport_RL)
        Me.GroupBox1.Controls.Add(Me.Label68)
        Me.GroupBox1.Controls.Add(Me.Label69)
        Me.GroupBox1.Controls.Add(Me.tbMontantTransport_RL)
        Me.GroupBox1.Controls.Add(Me.tbPUPallNonPrep_RL)
        Me.GroupBox1.Controls.Add(Me.tbPUPallPrep_RL)
        Me.GroupBox1.Controls.Add(Me.Label70)
        Me.GroupBox1.Controls.Add(Me.Label71)
        Me.GroupBox1.Controls.Add(Me.tbQtePallNonPrep_RL)
        Me.GroupBox1.Controls.Add(Me.tbQtePallPrep_RL)
        Me.GroupBox1.Controls.Add(Me.Label72)
        Me.GroupBox1.Controls.Add(Me.Label73)
        Me.GroupBox1.Controls.Add(Me.tbPoids_RL)
        Me.GroupBox1.Controls.Add(Me.Label74)
        Me.GroupBox1.Controls.Add(Me.tbQteColis_RL)
        Me.GroupBox1.Controls.Add(Me.Label75)
        Me.GroupBox1.Location = New System.Drawing.Point(11, 82)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(976, 78)
        Me.GroupBox1.TabIndex = 86
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Transport Facturé"
        '
        'cbCalcPoidsColis_RL
        '
        Me.cbCalcPoidsColis_RL.Location = New System.Drawing.Point(384, 16)
        Me.cbCalcPoidsColis_RL.Name = "cbCalcPoidsColis_RL"
        Me.cbCalcPoidsColis_RL.Size = New System.Drawing.Size(104, 24)
        Me.cbCalcPoidsColis_RL.TabIndex = 2
        Me.cbCalcPoidsColis_RL.Text = "Calcul"
        '
        'Label66
        '
        Me.Label66.Location = New System.Drawing.Point(731, 6)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(80, 16)
        Me.Label66.TabIndex = 182
        Me.Label66.Text = "P.U."
        Me.Label66.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label67
        '
        Me.Label67.Location = New System.Drawing.Point(643, 6)
        Me.Label67.Name = "Label67"
        Me.Label67.Size = New System.Drawing.Size(56, 16)
        Me.Label67.TabIndex = 181
        Me.Label67.Text = "Qte"
        Me.Label67.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbCalcMontantTransport_RL
        '
        Me.cbCalcMontantTransport_RL.Location = New System.Drawing.Point(835, 30)
        Me.cbCalcMontantTransport_RL.Name = "cbCalcMontantTransport_RL"
        Me.cbCalcMontantTransport_RL.Size = New System.Drawing.Size(16, 23)
        Me.cbCalcMontantTransport_RL.TabIndex = 7
        Me.cbCalcMontantTransport_RL.Text = "="
        '
        'Label68
        '
        Me.Label68.Location = New System.Drawing.Point(821, 46)
        Me.Label68.Name = "Label68"
        Me.Label68.Size = New System.Drawing.Size(8, 16)
        Me.Label68.TabIndex = 7
        Me.Label68.Text = "/"
        '
        'Label69
        '
        Me.Label69.Location = New System.Drawing.Point(819, 30)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(8, 23)
        Me.Label69.TabIndex = 7
        Me.Label69.Text = "\"
        '
        'tbMontantTransport_RL
        '
        Me.tbMontantTransport_RL.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "montantTransport", True))
        Me.tbMontantTransport_RL.Location = New System.Drawing.Point(859, 30)
        Me.tbMontantTransport_RL.Name = "tbMontantTransport_RL"
        Me.tbMontantTransport_RL.Size = New System.Drawing.Size(80, 20)
        Me.tbMontantTransport_RL.TabIndex = 8
        Me.tbMontantTransport_RL.Text = "0"
        '
        'tbPUPallNonPrep_RL
        '
        Me.tbPUPallNonPrep_RL.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "puPalettesNonPreparees", True))
        Me.tbPUPallNonPrep_RL.Location = New System.Drawing.Point(731, 46)
        Me.tbPUPallNonPrep_RL.Name = "tbPUPallNonPrep_RL"
        Me.tbPUPallNonPrep_RL.Size = New System.Drawing.Size(80, 20)
        Me.tbPUPallNonPrep_RL.TabIndex = 6
        Me.tbPUPallNonPrep_RL.Text = "0"
        '
        'tbPUPallPrep_RL
        '
        Me.tbPUPallPrep_RL.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "puPalettesPreparees", True))
        Me.tbPUPallPrep_RL.Location = New System.Drawing.Point(731, 22)
        Me.tbPUPallPrep_RL.Name = "tbPUPallPrep_RL"
        Me.tbPUPallPrep_RL.Size = New System.Drawing.Size(80, 20)
        Me.tbPUPallPrep_RL.TabIndex = 4
        Me.tbPUPallPrep_RL.Text = "0"
        '
        'Label70
        '
        Me.Label70.Location = New System.Drawing.Point(715, 46)
        Me.Label70.Name = "Label70"
        Me.Label70.Size = New System.Drawing.Size(16, 16)
        Me.Label70.TabIndex = 174
        Me.Label70.Text = "X"
        '
        'Label71
        '
        Me.Label71.Location = New System.Drawing.Point(715, 30)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(16, 16)
        Me.Label71.TabIndex = 173
        Me.Label71.Text = "X"
        '
        'tbQtePallNonPrep_RL
        '
        Me.tbQtePallNonPrep_RL.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "qtePalettesNonPreparees", True))
        Me.tbQtePallNonPrep_RL.Location = New System.Drawing.Point(643, 46)
        Me.tbQtePallNonPrep_RL.Name = "tbQtePallNonPrep_RL"
        Me.tbQtePallNonPrep_RL.Size = New System.Drawing.Size(64, 20)
        Me.tbQtePallNonPrep_RL.TabIndex = 5
        Me.tbQtePallNonPrep_RL.Text = "0"
        '
        'tbQtePallPrep_RL
        '
        Me.tbQtePallPrep_RL.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "qtePalettesPreparees", True))
        Me.tbQtePallPrep_RL.Location = New System.Drawing.Point(643, 22)
        Me.tbQtePallPrep_RL.Name = "tbQtePallPrep_RL"
        Me.tbQtePallPrep_RL.Size = New System.Drawing.Size(64, 20)
        Me.tbQtePallPrep_RL.TabIndex = 3
        Me.tbQtePallPrep_RL.Text = "0"
        '
        'Label72
        '
        Me.Label72.Location = New System.Drawing.Point(499, 46)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(128, 16)
        Me.Label72.TabIndex = 170
        Me.Label72.Text = "Pallettes non préparées"
        '
        'Label73
        '
        Me.Label73.Location = New System.Drawing.Point(499, 22)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(128, 16)
        Me.Label73.TabIndex = 169
        Me.Label73.Text = "Pallettes préparées"
        '
        'tbPoids_RL
        '
        Me.tbPoids_RL.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "poids", True))
        Me.tbPoids_RL.Location = New System.Drawing.Point(260, 18)
        Me.tbPoids_RL.Name = "tbPoids_RL"
        Me.tbPoids_RL.Size = New System.Drawing.Size(104, 20)
        Me.tbPoids_RL.TabIndex = 1
        Me.tbPoids_RL.Text = "0"
        '
        'Label74
        '
        Me.Label74.Location = New System.Drawing.Point(204, 18)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(40, 23)
        Me.Label74.TabIndex = 167
        Me.Label74.Text = "Poids"
        '
        'tbQteColis_RL
        '
        Me.tbQteColis_RL.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcCommande, "qteColis", True))
        Me.tbQteColis_RL.Location = New System.Drawing.Point(100, 18)
        Me.tbQteColis_RL.Name = "tbQteColis_RL"
        Me.tbQteColis_RL.Size = New System.Drawing.Size(88, 20)
        Me.tbQteColis_RL.TabIndex = 0
        Me.tbQteColis_RL.Text = "0"
        '
        'Label75
        '
        Me.Label75.Location = New System.Drawing.Point(28, 18)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(64, 23)
        Me.Label75.TabIndex = 165
        Me.Label75.Text = "Nbre Colis"
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.AutoGenerateColumns = False
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8, Me.DataGridViewTextBoxColumn9, Me.DataGridViewCheckBoxColumn1, Me.DataGridViewTextBoxColumn10, Me.DataGridViewTextBoxColumn11, Me.DataGridViewTextBoxColumn12})
        Me.DataGridView2.DataSource = Me.m_bsrcLignesCommande
        Me.DataGridView2.Location = New System.Drawing.Point(6, 166)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.RowHeadersWidth = 10
        Me.DataGridView2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView2.Size = New System.Drawing.Size(981, 366)
        Me.DataGridView2.TabIndex = 85
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "num"
        Me.DataGridViewTextBoxColumn1.FillWeight = 36.0!
        Me.DataGridViewTextBoxColumn1.HeaderText = "Num"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "ProduitCode"
        DataGridViewCellStyle41.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn2.DefaultCellStyle = DataGridViewCellStyle41
        Me.DataGridViewTextBoxColumn2.FillWeight = 45.0!
        Me.DataGridViewTextBoxColumn2.HeaderText = "Code Produit"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "ProduitNom"
        Me.DataGridViewTextBoxColumn3.FillWeight = 360.0!
        Me.DataGridViewTextBoxColumn3.HeaderText = "Désignation Produit"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.ReadOnly = True
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "ProduitMil"
        Me.DataGridViewTextBoxColumn4.FillWeight = 45.0!
        Me.DataGridViewTextBoxColumn4.HeaderText = "Mill."
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.DataPropertyName = "ProduitConditionnement"
        Me.DataGridViewTextBoxColumn5.FillWeight = 36.0!
        Me.DataGridViewTextBoxColumn5.HeaderText = "Cond."
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.ReadOnly = True
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.DataPropertyName = "ProduitCouleur"
        Me.DataGridViewTextBoxColumn6.FillWeight = 43.0!
        Me.DataGridViewTextBoxColumn6.HeaderText = "Coul."
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.ReadOnly = True
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.DataPropertyName = "qteCommande"
        DataGridViewCellStyle42.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn7.DefaultCellStyle = DataGridViewCellStyle42
        Me.DataGridViewTextBoxColumn7.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn7.HeaderText = "Qte Comm"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.ReadOnly = True
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.DataPropertyName = "qteLiv"
        DataGridViewCellStyle43.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn8.DefaultCellStyle = DataGridViewCellStyle43
        Me.DataGridViewTextBoxColumn8.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn8.HeaderText = "Qte Liv"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.ReadOnly = True
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.DataPropertyName = "qteFact"
        DataGridViewCellStyle44.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn9.DefaultCellStyle = DataGridViewCellStyle44
        Me.DataGridViewTextBoxColumn9.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn9.HeaderText = "Qte Fact"
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.ReadOnly = True
        '
        'DataGridViewCheckBoxColumn1
        '
        Me.DataGridViewCheckBoxColumn1.DataPropertyName = "bGratuit"
        Me.DataGridViewCheckBoxColumn1.FillWeight = 40.0!
        Me.DataGridViewCheckBoxColumn1.HeaderText = "Gratuit"
        Me.DataGridViewCheckBoxColumn1.Name = "DataGridViewCheckBoxColumn1"
        Me.DataGridViewCheckBoxColumn1.ReadOnly = True
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.DataPropertyName = "prixU"
        DataGridViewCellStyle45.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle45.Format = "C2"
        Me.DataGridViewTextBoxColumn10.DefaultCellStyle = DataGridViewCellStyle45
        Me.DataGridViewTextBoxColumn10.FillWeight = 56.0!
        Me.DataGridViewTextBoxColumn10.HeaderText = "Prix U"
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        Me.DataGridViewTextBoxColumn10.ReadOnly = True
        '
        'DataGridViewTextBoxColumn11
        '
        Me.DataGridViewTextBoxColumn11.DataPropertyName = "prixHT"
        DataGridViewCellStyle46.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle46.Format = "C2"
        Me.DataGridViewTextBoxColumn11.DefaultCellStyle = DataGridViewCellStyle46
        Me.DataGridViewTextBoxColumn11.FillWeight = 70.0!
        Me.DataGridViewTextBoxColumn11.HeaderText = "Montant HT"
        Me.DataGridViewTextBoxColumn11.Name = "DataGridViewTextBoxColumn11"
        Me.DataGridViewTextBoxColumn11.ReadOnly = True
        '
        'DataGridViewTextBoxColumn12
        '
        Me.DataGridViewTextBoxColumn12.DataPropertyName = "prixTTC"
        DataGridViewCellStyle47.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle47.Format = "C2"
        Me.DataGridViewTextBoxColumn12.DefaultCellStyle = DataGridViewCellStyle47
        Me.DataGridViewTextBoxColumn12.FillWeight = 70.0!
        Me.DataGridViewTextBoxColumn12.HeaderText = "Montant TTC"
        Me.DataGridViewTextBoxColumn12.Name = "DataGridViewTextBoxColumn12"
        Me.DataGridViewTextBoxColumn12.ReadOnly = True
        '
        'tbCoutTransport
        '
        Me.tbCoutTransport.Location = New System.Drawing.Point(693, 32)
        Me.tbCoutTransport.Name = "tbCoutTransport"
        Me.tbCoutTransport.Size = New System.Drawing.Size(140, 20)
        Me.tbCoutTransport.TabIndex = 5
        Me.tbCoutTransport.Text = "0"
        '
        'Label65
        '
        Me.Label65.AutoSize = True
        Me.Label65.Location = New System.Drawing.Point(555, 59)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(132, 13)
        Me.Label65.TabIndex = 84
        Me.Label65.Text = "Ref Facture Transporteur :"
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Location = New System.Drawing.Point(555, 36)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(83, 13)
        Me.Label64.TabIndex = 83
        Me.Label64.Text = "Coût Transport :"
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Location = New System.Drawing.Point(555, 15)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(76, 13)
        Me.Label63.TabIndex = 82
        Me.Label63.Text = "Lettre Voiture :"
        '
        'tbRefFactTRP
        '
        Me.tbRefFactTRP.Location = New System.Drawing.Point(693, 56)
        Me.tbRefFactTRP.Name = "tbRefFactTRP"
        Me.tbRefFactTRP.Size = New System.Drawing.Size(140, 20)
        Me.tbRefFactTRP.TabIndex = 6
        '
        'tbLettreVoiture
        '
        Me.tbLettreVoiture.Location = New System.Drawing.Point(693, 9)
        Me.tbLettreVoiture.Name = "tbLettreVoiture"
        Me.tbLettreVoiture.Size = New System.Drawing.Size(140, 20)
        Me.tbLettreVoiture.TabIndex = 4
        '
        'cbAnnulerLivraison
        '
        Me.cbAnnulerLivraison.Location = New System.Drawing.Point(856, 16)
        Me.cbAnnulerLivraison.Name = "cbAnnulerLivraison"
        Me.cbAnnulerLivraison.Size = New System.Drawing.Size(120, 23)
        Me.cbAnnulerLivraison.TabIndex = 8
        Me.cbAnnulerLivraison.Text = "&Annuler Livraison"
        '
        'dtDateEnlevementReelle
        '
        Me.dtDateEnlevementReelle.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateEnlevementReelle.Location = New System.Drawing.Point(176, 8)
        Me.dtDateEnlevementReelle.Name = "dtDateEnlevementReelle"
        Me.dtDateEnlevementReelle.Size = New System.Drawing.Size(104, 20)
        Me.dtDateEnlevementReelle.TabIndex = 1
        '
        'Label31
        '
        Me.Label31.Location = New System.Drawing.Point(8, 8)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(144, 16)
        Me.Label31.TabIndex = 78
        Me.Label31.Text = "Date enlèvement réelle"
        '
        'tbRefBL
        '
        Me.tbRefBL.Location = New System.Drawing.Point(176, 56)
        Me.tbRefBL.Name = "tbRefBL"
        Me.tbRefBL.Size = New System.Drawing.Size(312, 20)
        Me.tbRefBL.TabIndex = 3
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(8, 56)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(136, 16)
        Me.Label23.TabIndex = 76
        Me.Label23.Text = "Reférence de Livraison"
        '
        'dtDateLivraisonReelle
        '
        Me.dtDateLivraisonReelle.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateLivraisonReelle.Location = New System.Drawing.Point(176, 32)
        Me.dtDateLivraisonReelle.Name = "dtDateLivraisonReelle"
        Me.dtDateLivraisonReelle.Size = New System.Drawing.Size(104, 20)
        Me.dtDateLivraisonReelle.TabIndex = 2
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(8, 32)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(152, 16)
        Me.Label22.TabIndex = 74
        Me.Label22.Text = "date de livraison réelle"
        '
        'cbBLToutOK
        '
        Me.cbBLToutOK.Location = New System.Drawing.Point(856, 48)
        Me.cbBLToutOK.Name = "cbBLToutOK"
        Me.cbBLToutOK.Size = New System.Drawing.Size(120, 24)
        Me.cbBLToutOK.TabIndex = 0
        Me.cbBLToutOK.Text = "Tout &OK"
        '
        'tpEclatement
        '
        Me.tpEclatement.Controls.Add(Me.laIntermediaires)
        Me.tpEclatement.Controls.Add(Me.cbxIntermédiaires)
        Me.tpEclatement.Controls.Add(Me.m_dgvSCMD)
        Me.tpEclatement.Controls.Add(Me.DataGridView3)
        Me.tpEclatement.Controls.Add(Me.cbSCMDAppliquer)
        Me.tpEclatement.Controls.Add(Me.cbAnnEclatement)
        Me.tpEclatement.Controls.Add(Me.tbSCMDCommentaire)
        Me.tpEclatement.Controls.Add(Me.cbSCMDVoir)
        Me.tpEclatement.Controls.Add(Me.cbSCMDFaxerTout)
        Me.tpEclatement.Controls.Add(Me.Label36)
        Me.tpEclatement.Controls.Add(Me.tbSCMDTransporteurNom)
        Me.tpEclatement.Controls.Add(Me.Label32)
        Me.tpEclatement.Controls.Add(Me.liSCMDFournisseur)
        Me.tpEclatement.Controls.Add(Me.dtSCMDDateLiv)
        Me.tpEclatement.Controls.Add(Me.Label24)
        Me.tpEclatement.Controls.Add(Me.cbEclatementCmde)
        Me.tpEclatement.Location = New System.Drawing.Point(4, 22)
        Me.tpEclatement.Name = "tpEclatement"
        Me.tpEclatement.Size = New System.Drawing.Size(992, 557)
        Me.tpEclatement.TabIndex = 5
        Me.tpEclatement.Text = "Eclatement Commande"
        '
        'laIntermediaires
        '
        Me.laIntermediaires.AutoSize = True
        Me.laIntermediaires.Location = New System.Drawing.Point(645, 211)
        Me.laIntermediaires.Name = "laIntermediaires"
        Me.laIntermediaires.Size = New System.Drawing.Size(78, 13)
        Me.laIntermediaires.TabIndex = 78
        Me.laIntermediaires.Text = "Intermédiaires :"
        '
        'cbxIntermédiaires
        '
        Me.cbxIntermédiaires.DataSource = Me.m_bsrcIntermédiaires
        Me.cbxIntermédiaires.DisplayMember = "shortResume"
        Me.cbxIntermédiaires.FormattingEnabled = True
        Me.cbxIntermédiaires.Location = New System.Drawing.Point(738, 208)
        Me.cbxIntermédiaires.Name = "cbxIntermédiaires"
        Me.cbxIntermédiaires.Size = New System.Drawing.Size(238, 21)
        Me.cbxIntermédiaires.TabIndex = 77
        '
        'm_bsrcIntermédiaires
        '
        Me.m_bsrcIntermédiaires.DataSource = GetType(vini_DB.Client)
        '
        'm_dgvSCMD
        '
        Me.m_dgvSCMD.AllowUserToAddRows = False
        Me.m_dgvSCMD.AllowUserToDeleteRows = False
        Me.m_dgvSCMD.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.m_dgvSCMD.AutoGenerateColumns = False
        Me.m_dgvSCMD.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.m_dgvSCMD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.m_dgvSCMD.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.FournisseurRSDataGridViewTextBoxColumn, Me.TotalHTDataGridViewTextBoxColumn, Me.DateCommandeDataGridViewTextBoxColumn, Me.EtatLibelleDataGridViewTextBoxColumn})
        Me.m_dgvSCMD.DataSource = Me.m_bsrcSousCommande
        Me.m_dgvSCMD.Location = New System.Drawing.Point(9, 239)
        Me.m_dgvSCMD.Name = "m_dgvSCMD"
        Me.m_dgvSCMD.ReadOnly = True
        Me.m_dgvSCMD.Size = New System.Drawing.Size(415, 273)
        Me.m_dgvSCMD.TabIndex = 75
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.FillWeight = 76.0!
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "Code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FournisseurRSDataGridViewTextBoxColumn
        '
        Me.FournisseurRSDataGridViewTextBoxColumn.DataPropertyName = "FournisseurRS"
        Me.FournisseurRSDataGridViewTextBoxColumn.FillWeight = 170.0!
        Me.FournisseurRSDataGridViewTextBoxColumn.HeaderText = "Producteur"
        Me.FournisseurRSDataGridViewTextBoxColumn.Name = "FournisseurRSDataGridViewTextBoxColumn"
        Me.FournisseurRSDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TotalHTDataGridViewTextBoxColumn
        '
        Me.TotalHTDataGridViewTextBoxColumn.DataPropertyName = "totalHT"
        DataGridViewCellStyle48.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle48.Format = "C2"
        Me.TotalHTDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle48
        Me.TotalHTDataGridViewTextBoxColumn.FillWeight = 70.0!
        Me.TotalHTDataGridViewTextBoxColumn.HeaderText = "Montant HT"
        Me.TotalHTDataGridViewTextBoxColumn.Name = "TotalHTDataGridViewTextBoxColumn"
        Me.TotalHTDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DateCommandeDataGridViewTextBoxColumn
        '
        Me.DateCommandeDataGridViewTextBoxColumn.DataPropertyName = "dateCommande"
        DataGridViewCellStyle49.Format = "d"
        DataGridViewCellStyle49.NullValue = Nothing
        Me.DateCommandeDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle49
        Me.DateCommandeDataGridViewTextBoxColumn.FillWeight = 60.0!
        Me.DateCommandeDataGridViewTextBoxColumn.HeaderText = "date Commande"
        Me.DateCommandeDataGridViewTextBoxColumn.Name = "DateCommandeDataGridViewTextBoxColumn"
        Me.DateCommandeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'EtatLibelleDataGridViewTextBoxColumn
        '
        Me.EtatLibelleDataGridViewTextBoxColumn.DataPropertyName = "EtatLibelle"
        Me.EtatLibelleDataGridViewTextBoxColumn.FillWeight = 12.0!
        Me.EtatLibelleDataGridViewTextBoxColumn.HeaderText = "Etat"
        Me.EtatLibelleDataGridViewTextBoxColumn.Name = "EtatLibelleDataGridViewTextBoxColumn"
        Me.EtatLibelleDataGridViewTextBoxColumn.ReadOnly = True
        '
        'm_bsrcSousCommande
        '
        Me.m_bsrcSousCommande.DataSource = GetType(vini_DB.SousCommande)
        '
        'DataGridView3
        '
        Me.DataGridView3.AllowUserToAddRows = False
        Me.DataGridView3.AllowUserToDeleteRows = False
        Me.DataGridView3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView3.AutoGenerateColumns = False
        Me.DataGridView3.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn13, Me.DataGridViewTextBoxColumn14, Me.DataGridViewTextBoxColumn15, Me.DataGridViewTextBoxColumn16, Me.DataGridViewTextBoxColumn17, Me.DataGridViewTextBoxColumn18, Me.DataGridViewTextBoxColumn19, Me.DataGridViewTextBoxColumn20, Me.DataGridViewTextBoxColumn21, Me.DataGridViewCheckBoxColumn2, Me.DataGridViewTextBoxColumn22, Me.DataGridViewTextBoxColumn23, Me.DataGridViewTextBoxColumn24})
        Me.DataGridView3.DataSource = Me.m_bsrcLignesCommande
        Me.DataGridView3.Location = New System.Drawing.Point(4, 3)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.ReadOnly = True
        Me.DataGridView3.RowHeadersWidth = 10
        Me.DataGridView3.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView3.Size = New System.Drawing.Size(981, 196)
        Me.DataGridView3.TabIndex = 74
        '
        'DataGridViewTextBoxColumn13
        '
        Me.DataGridViewTextBoxColumn13.DataPropertyName = "num"
        Me.DataGridViewTextBoxColumn13.FillWeight = 36.0!
        Me.DataGridViewTextBoxColumn13.HeaderText = "Num"
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        Me.DataGridViewTextBoxColumn13.ReadOnly = True
        '
        'DataGridViewTextBoxColumn14
        '
        Me.DataGridViewTextBoxColumn14.DataPropertyName = "ProduitCode"
        DataGridViewCellStyle50.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn14.DefaultCellStyle = DataGridViewCellStyle50
        Me.DataGridViewTextBoxColumn14.FillWeight = 45.0!
        Me.DataGridViewTextBoxColumn14.HeaderText = "Code Produit"
        Me.DataGridViewTextBoxColumn14.Name = "DataGridViewTextBoxColumn14"
        Me.DataGridViewTextBoxColumn14.ReadOnly = True
        '
        'DataGridViewTextBoxColumn15
        '
        Me.DataGridViewTextBoxColumn15.DataPropertyName = "ProduitNom"
        Me.DataGridViewTextBoxColumn15.FillWeight = 360.0!
        Me.DataGridViewTextBoxColumn15.HeaderText = "Désignation Produit"
        Me.DataGridViewTextBoxColumn15.Name = "DataGridViewTextBoxColumn15"
        Me.DataGridViewTextBoxColumn15.ReadOnly = True
        '
        'DataGridViewTextBoxColumn16
        '
        Me.DataGridViewTextBoxColumn16.DataPropertyName = "ProduitMil"
        Me.DataGridViewTextBoxColumn16.FillWeight = 45.0!
        Me.DataGridViewTextBoxColumn16.HeaderText = "Mill."
        Me.DataGridViewTextBoxColumn16.Name = "DataGridViewTextBoxColumn16"
        Me.DataGridViewTextBoxColumn16.ReadOnly = True
        '
        'DataGridViewTextBoxColumn17
        '
        Me.DataGridViewTextBoxColumn17.DataPropertyName = "ProduitConditionnement"
        Me.DataGridViewTextBoxColumn17.FillWeight = 36.0!
        Me.DataGridViewTextBoxColumn17.HeaderText = "Cond."
        Me.DataGridViewTextBoxColumn17.Name = "DataGridViewTextBoxColumn17"
        Me.DataGridViewTextBoxColumn17.ReadOnly = True
        '
        'DataGridViewTextBoxColumn18
        '
        Me.DataGridViewTextBoxColumn18.DataPropertyName = "ProduitCouleur"
        Me.DataGridViewTextBoxColumn18.FillWeight = 43.0!
        Me.DataGridViewTextBoxColumn18.HeaderText = "Coul."
        Me.DataGridViewTextBoxColumn18.Name = "DataGridViewTextBoxColumn18"
        Me.DataGridViewTextBoxColumn18.ReadOnly = True
        '
        'DataGridViewTextBoxColumn19
        '
        Me.DataGridViewTextBoxColumn19.DataPropertyName = "qteCommande"
        DataGridViewCellStyle51.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn19.DefaultCellStyle = DataGridViewCellStyle51
        Me.DataGridViewTextBoxColumn19.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn19.HeaderText = "Qte Comm"
        Me.DataGridViewTextBoxColumn19.Name = "DataGridViewTextBoxColumn19"
        Me.DataGridViewTextBoxColumn19.ReadOnly = True
        '
        'DataGridViewTextBoxColumn20
        '
        Me.DataGridViewTextBoxColumn20.DataPropertyName = "qteLiv"
        DataGridViewCellStyle52.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn20.DefaultCellStyle = DataGridViewCellStyle52
        Me.DataGridViewTextBoxColumn20.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn20.HeaderText = "Qte Liv"
        Me.DataGridViewTextBoxColumn20.Name = "DataGridViewTextBoxColumn20"
        Me.DataGridViewTextBoxColumn20.ReadOnly = True
        '
        'DataGridViewTextBoxColumn21
        '
        Me.DataGridViewTextBoxColumn21.DataPropertyName = "qteFact"
        DataGridViewCellStyle53.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn21.DefaultCellStyle = DataGridViewCellStyle53
        Me.DataGridViewTextBoxColumn21.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn21.HeaderText = "Qte Fact"
        Me.DataGridViewTextBoxColumn21.Name = "DataGridViewTextBoxColumn21"
        Me.DataGridViewTextBoxColumn21.ReadOnly = True
        '
        'DataGridViewCheckBoxColumn2
        '
        Me.DataGridViewCheckBoxColumn2.DataPropertyName = "bGratuit"
        Me.DataGridViewCheckBoxColumn2.FillWeight = 40.0!
        Me.DataGridViewCheckBoxColumn2.HeaderText = "Gratuit"
        Me.DataGridViewCheckBoxColumn2.Name = "DataGridViewCheckBoxColumn2"
        Me.DataGridViewCheckBoxColumn2.ReadOnly = True
        '
        'DataGridViewTextBoxColumn22
        '
        Me.DataGridViewTextBoxColumn22.DataPropertyName = "prixU"
        DataGridViewCellStyle54.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle54.Format = "C2"
        Me.DataGridViewTextBoxColumn22.DefaultCellStyle = DataGridViewCellStyle54
        Me.DataGridViewTextBoxColumn22.FillWeight = 56.0!
        Me.DataGridViewTextBoxColumn22.HeaderText = "Prix U"
        Me.DataGridViewTextBoxColumn22.Name = "DataGridViewTextBoxColumn22"
        Me.DataGridViewTextBoxColumn22.ReadOnly = True
        '
        'DataGridViewTextBoxColumn23
        '
        Me.DataGridViewTextBoxColumn23.DataPropertyName = "prixHT"
        DataGridViewCellStyle55.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle55.Format = "C2"
        Me.DataGridViewTextBoxColumn23.DefaultCellStyle = DataGridViewCellStyle55
        Me.DataGridViewTextBoxColumn23.FillWeight = 70.0!
        Me.DataGridViewTextBoxColumn23.HeaderText = "Montant HT"
        Me.DataGridViewTextBoxColumn23.Name = "DataGridViewTextBoxColumn23"
        Me.DataGridViewTextBoxColumn23.ReadOnly = True
        '
        'DataGridViewTextBoxColumn24
        '
        Me.DataGridViewTextBoxColumn24.DataPropertyName = "prixTTC"
        DataGridViewCellStyle56.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle56.Format = "C2"
        Me.DataGridViewTextBoxColumn24.DefaultCellStyle = DataGridViewCellStyle56
        Me.DataGridViewTextBoxColumn24.FillWeight = 70.0!
        Me.DataGridViewTextBoxColumn24.HeaderText = "Montant TTC"
        Me.DataGridViewTextBoxColumn24.Name = "DataGridViewTextBoxColumn24"
        Me.DataGridViewTextBoxColumn24.ReadOnly = True
        '
        'cbSCMDAppliquer
        '
        Me.cbSCMDAppliquer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSCMDAppliquer.Location = New System.Drawing.Point(880, 488)
        Me.cbSCMDAppliquer.Name = "cbSCMDAppliquer"
        Me.cbSCMDAppliquer.Size = New System.Drawing.Size(96, 24)
        Me.cbSCMDAppliquer.TabIndex = 10
        Me.cbSCMDAppliquer.Text = "Appliquer"
        '
        'cbAnnEclatement
        '
        Me.cbAnnEclatement.Location = New System.Drawing.Point(432, 208)
        Me.cbAnnEclatement.Name = "cbAnnEclatement"
        Me.cbAnnEclatement.Size = New System.Drawing.Size(144, 24)
        Me.cbAnnEclatement.TabIndex = 72
        Me.cbAnnEclatement.Text = "Annulation Eclatement"
        '
        'tbSCMDCommentaire
        '
        Me.tbSCMDCommentaire.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbSCMDCommentaire.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommande, "CommentaireFacturationText", True))
        Me.tbSCMDCommentaire.Location = New System.Drawing.Point(440, 336)
        Me.tbSCMDCommentaire.Name = "tbSCMDCommentaire"
        Me.tbSCMDCommentaire.Size = New System.Drawing.Size(536, 144)
        Me.tbSCMDCommentaire.TabIndex = 4
        Me.tbSCMDCommentaire.Text = ""
        '
        'cbSCMDVoir
        '
        Me.cbSCMDVoir.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSCMDVoir.Location = New System.Drawing.Point(440, 488)
        Me.cbSCMDVoir.Name = "cbSCMDVoir"
        Me.cbSCMDVoir.Size = New System.Drawing.Size(72, 24)
        Me.cbSCMDVoir.TabIndex = 7
        Me.cbSCMDVoir.Text = "Voir"
        '
        'cbSCMDFaxerTout
        '
        Me.cbSCMDFaxerTout.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSCMDFaxerTout.Location = New System.Drawing.Point(528, 488)
        Me.cbSCMDFaxerTout.Name = "cbSCMDFaxerTout"
        Me.cbSCMDFaxerTout.Size = New System.Drawing.Size(104, 24)
        Me.cbSCMDFaxerTout.TabIndex = 9
        Me.cbSCMDFaxerTout.Text = "Faxer Tout"
        '
        'Label36
        '
        Me.Label36.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label36.Location = New System.Drawing.Point(440, 320)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(80, 16)
        Me.Label36.TabIndex = 71
        Me.Label36.Text = "Commentaire"
        '
        'tbSCMDTransporteurNom
        '
        Me.tbSCMDTransporteurNom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbSCMDTransporteurNom.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommande, "TransporteurNom", True))
        Me.tbSCMDTransporteurNom.Location = New System.Drawing.Point(440, 296)
        Me.tbSCMDTransporteurNom.Name = "tbSCMDTransporteurNom"
        Me.tbSCMDTransporteurNom.Size = New System.Drawing.Size(536, 20)
        Me.tbSCMDTransporteurNom.TabIndex = 3
        '
        'Label32
        '
        Me.Label32.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label32.Location = New System.Drawing.Point(440, 280)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(96, 16)
        Me.Label32.TabIndex = 69
        Me.Label32.Text = "Transporteur"
        '
        'liSCMDFournisseur
        '
        Me.liSCMDFournisseur.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.liSCMDFournisseur.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommande, "FournisseurRS", True))
        Me.liSCMDFournisseur.DataBindings.Add(New System.Windows.Forms.Binding("Tag", Me.m_bsrcSousCommande, "Fournisseurid", True))
        Me.liSCMDFournisseur.Location = New System.Drawing.Point(440, 232)
        Me.liSCMDFournisseur.Name = "liSCMDFournisseur"
        Me.liSCMDFournisseur.Size = New System.Drawing.Size(408, 16)
        Me.liSCMDFournisseur.TabIndex = 68
        Me.liSCMDFournisseur.TabStop = True
        Me.liSCMDFournisseur.Text = "Désignation du Fournisseur"
        '
        'dtSCMDDateLiv
        '
        Me.dtSCMDDateLiv.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtSCMDDateLiv.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcSousCommande, "dateLivraison", True))
        Me.dtSCMDDateLiv.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtSCMDDateLiv.Location = New System.Drawing.Point(528, 256)
        Me.dtSCMDDateLiv.Name = "dtSCMDDateLiv"
        Me.dtSCMDDateLiv.Size = New System.Drawing.Size(96, 20)
        Me.dtSCMDDateLiv.TabIndex = 2
        '
        'Label24
        '
        Me.Label24.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label24.Location = New System.Drawing.Point(440, 256)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(72, 16)
        Me.Label24.TabIndex = 66
        Me.Label24.Text = "date livraison"
        '
        'cbEclatementCmde
        '
        Me.cbEclatementCmde.Location = New System.Drawing.Point(8, 208)
        Me.cbEclatementCmde.Name = "cbEclatementCmde"
        Me.cbEclatementCmde.Size = New System.Drawing.Size(416, 24)
        Me.cbEclatementCmde.TabIndex = 0
        Me.cbEclatementCmde.Text = "Eclatement de la commande client"
        '
        'tpFactHbv
        '
        Me.tpFactHbv.BackColor = System.Drawing.SystemColors.Control
        Me.tpFactHbv.Controls.Add(Me.pnlFactHBV)
        Me.tpFactHbv.Location = New System.Drawing.Point(4, 22)
        Me.tpFactHbv.Name = "tpFactHbv"
        Me.tpFactHbv.Padding = New System.Windows.Forms.Padding(3)
        Me.tpFactHbv.Size = New System.Drawing.Size(992, 557)
        Me.tpFactHbv.TabIndex = 8
        Me.tpFactHbv.Text = "FactureHobivin"
        '
        'pnlFactHBV
        '
        Me.pnlFactHBV.Controls.Add(Me.SplitFactHBV)
        Me.pnlFactHBV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFactHBV.Location = New System.Drawing.Point(3, 3)
        Me.pnlFactHBV.Name = "pnlFactHBV"
        Me.pnlFactHBV.Size = New System.Drawing.Size(986, 551)
        Me.pnlFactHBV.TabIndex = 152
        '
        'SplitFactHBV
        '
        Me.SplitFactHBV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitFactHBV.Location = New System.Drawing.Point(0, 0)
        Me.SplitFactHBV.Name = "SplitFactHBV"
        '
        'SplitFactHBV.Panel1
        '
        Me.SplitFactHBV.Panel1.Controls.Add(Me.crwFact)
        '
        'SplitFactHBV.Panel2
        '
        Me.SplitFactHBV.Panel2.Controls.Add(Me.btnFactHBVAfficher)
        Me.SplitFactHBV.Panel2.Controls.Add(Me.DataGridView4)
        Me.SplitFactHBV.Size = New System.Drawing.Size(986, 551)
        Me.SplitFactHBV.SplitterDistance = 328
        Me.SplitFactHBV.TabIndex = 0
        '
        'crwFact
        '
        Me.crwFact.ActiveViewIndex = -1
        Me.crwFact.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crwFact.Cursor = System.Windows.Forms.Cursors.Default
        Me.crwFact.DisplayStatusBar = False
        Me.crwFact.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crwFact.Location = New System.Drawing.Point(0, 0)
        Me.crwFact.Name = "crwFact"
        Me.crwFact.Size = New System.Drawing.Size(328, 551)
        Me.crwFact.TabIndex = 149
        Me.crwFact.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'btnFactHBVAfficher
        '
        Me.btnFactHBVAfficher.Location = New System.Drawing.Point(3, 3)
        Me.btnFactHBVAfficher.Name = "btnFactHBVAfficher"
        Me.btnFactHBVAfficher.Size = New System.Drawing.Size(104, 24)
        Me.btnFactHBVAfficher.TabIndex = 150
        Me.btnFactHBVAfficher.Text = "Afficher"
        '
        'DataGridView4
        '
        Me.DataGridView4.AllowUserToAddRows = False
        Me.DataGridView4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView4.AutoGenerateColumns = False
        Me.DataGridView4.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ProduitCodeDataGridViewTextBoxColumn1, Me.ProduitNomDataGridViewTextBoxColumn1, Me.QteFactDataGridViewTextBoxColumn1, Me.BGratuitDataGridViewCheckBoxColumn1, Me.PrixUDataGridViewTextBoxColumn1, Me.PrixHTDataGridViewTextBoxColumn1, Me.PrixTTCDataGridViewTextBoxColumn1})
        Me.DataGridView4.DataSource = Me.m_bsrcLgFactHBV
        Me.DataGridView4.Location = New System.Drawing.Point(3, 33)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.Size = New System.Drawing.Size(648, 515)
        Me.DataGridView4.TabIndex = 151
        '
        'ProduitCodeDataGridViewTextBoxColumn1
        '
        Me.ProduitCodeDataGridViewTextBoxColumn1.DataPropertyName = "ProduitCode"
        DataGridViewCellStyle29.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ProduitCodeDataGridViewTextBoxColumn1.DefaultCellStyle = DataGridViewCellStyle29
        Me.ProduitCodeDataGridViewTextBoxColumn1.HeaderText = "Code"
        Me.ProduitCodeDataGridViewTextBoxColumn1.Name = "ProduitCodeDataGridViewTextBoxColumn1"
        Me.ProduitCodeDataGridViewTextBoxColumn1.ReadOnly = True
        '
        'ProduitNomDataGridViewTextBoxColumn1
        '
        Me.ProduitNomDataGridViewTextBoxColumn1.DataPropertyName = "ProduitNom"
        DataGridViewCellStyle30.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ProduitNomDataGridViewTextBoxColumn1.DefaultCellStyle = DataGridViewCellStyle30
        Me.ProduitNomDataGridViewTextBoxColumn1.HeaderText = "Libellé"
        Me.ProduitNomDataGridViewTextBoxColumn1.Name = "ProduitNomDataGridViewTextBoxColumn1"
        Me.ProduitNomDataGridViewTextBoxColumn1.ReadOnly = True
        '
        'QteFactDataGridViewTextBoxColumn1
        '
        Me.QteFactDataGridViewTextBoxColumn1.DataPropertyName = "qteFact"
        Me.QteFactDataGridViewTextBoxColumn1.HeaderText = "Qte"
        Me.QteFactDataGridViewTextBoxColumn1.Name = "QteFactDataGridViewTextBoxColumn1"
        '
        'BGratuitDataGridViewCheckBoxColumn1
        '
        Me.BGratuitDataGridViewCheckBoxColumn1.DataPropertyName = "bGratuit"
        Me.BGratuitDataGridViewCheckBoxColumn1.HeaderText = "Gratuit ?"
        Me.BGratuitDataGridViewCheckBoxColumn1.Name = "BGratuitDataGridViewCheckBoxColumn1"
        '
        'PrixUDataGridViewTextBoxColumn1
        '
        Me.PrixUDataGridViewTextBoxColumn1.DataPropertyName = "prixU"
        DataGridViewCellStyle31.Format = "C2"
        DataGridViewCellStyle31.NullValue = Nothing
        Me.PrixUDataGridViewTextBoxColumn1.DefaultCellStyle = DataGridViewCellStyle31
        Me.PrixUDataGridViewTextBoxColumn1.HeaderText = "P.U."
        Me.PrixUDataGridViewTextBoxColumn1.Name = "PrixUDataGridViewTextBoxColumn1"
        '
        'PrixHTDataGridViewTextBoxColumn1
        '
        Me.PrixHTDataGridViewTextBoxColumn1.DataPropertyName = "prixHT"
        DataGridViewCellStyle32.Format = "C2"
        DataGridViewCellStyle32.NullValue = Nothing
        Me.PrixHTDataGridViewTextBoxColumn1.DefaultCellStyle = DataGridViewCellStyle32
        Me.PrixHTDataGridViewTextBoxColumn1.HeaderText = "Montant HT"
        Me.PrixHTDataGridViewTextBoxColumn1.Name = "PrixHTDataGridViewTextBoxColumn1"
        '
        'PrixTTCDataGridViewTextBoxColumn1
        '
        Me.PrixTTCDataGridViewTextBoxColumn1.DataPropertyName = "prixTTC"
        DataGridViewCellStyle33.Format = "C2"
        DataGridViewCellStyle33.NullValue = Nothing
        Me.PrixTTCDataGridViewTextBoxColumn1.DefaultCellStyle = DataGridViewCellStyle33
        Me.PrixTTCDataGridViewTextBoxColumn1.HeaderText = "Montant TTC"
        Me.PrixTTCDataGridViewTextBoxColumn1.Name = "PrixTTCDataGridViewTextBoxColumn1"
        '
        'm_bsrcLgFactHBV
        '
        Me.m_bsrcLgFactHBV.DataSource = GetType(vini_DB.LgFactHBV)
        '
        'Label34
        '
        Me.Label34.Location = New System.Drawing.Point(104, 24)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(128, 16)
        Me.Label34.TabIndex = 68
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(496, 232)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 24)
        Me.Button1.TabIndex = 67
        Me.Button1.Text = "Editer BaF"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(368, 232)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 24)
        Me.Button3.TabIndex = 10
        Me.Button3.Text = "Valider BaF"
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(8, 56)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(64, 16)
        Me.Label33.TabIndex = 9
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Location = New System.Drawing.Point(80, 48)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(488, 176)
        Me.RichTextBox2.TabIndex = 8
        Me.RichTextBox2.Text = ""
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(424, 136)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(88, 24)
        Me.Button4.TabIndex = 7
        Me.Button4.Text = "Appliquer"
        '
        'Label35
        '
        Me.Label35.Location = New System.Drawing.Point(8, 24)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(88, 16)
        Me.Label35.TabIndex = 0
        '
        'frmSaisieCommande
        '
        Me.ClientSize = New System.Drawing.Size(1016, 659)
        Me.Controls.Add(Me.SSTabCommandeClient)
        Me.Controls.Add(Me.grpEntete)
        Me.Name = "frmSaisieCommande"
        Me.Text = "Saisie Commande BA"
        Me.grpEntete.ResumeLayout(False)
        Me.grpEntete.PerformLayout()
        Me.grpPrestashop.ResumeLayout(False)
        Me.grpPrestashop.PerformLayout()
        Me.grpFactTRP.ResumeLayout(False)
        Me.grpTypeTransport.ResumeLayout(False)
        Me.grpTypeCommande.ResumeLayout(False)
        Me.SSTabCommandeClient.ResumeLayout(False)
        Me.tpClient.ResumeLayout(False)
        Me.tpClient.PerformLayout()
        Me.grpAdrFact.ResumeLayout(False)
        Me.grpAdrFact.PerformLayout()
        Me.grpAdrLiv.ResumeLayout(False)
        Me.grpAdrLiv.PerformLayout()
        Me.tpLignes.ResumeLayout(False)
        Me.tpLignes.PerformLayout()
        CType(Me.m_bsrcCommande, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcLignesCommande, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpCommentaires.ResumeLayout(False)
        Me.tpValidCmd.ResumeLayout(False)
        Me.tpValidCmd.PerformLayout()
        Me.tpBL.ResumeLayout(False)
        Me.tpBL.PerformLayout()
        Me.grpInfFactureTRP.ResumeLayout(False)
        Me.grpInfFactureTRP.PerformLayout()
        Me.grpTransporteur.ResumeLayout(False)
        Me.grpTransporteur.PerformLayout()
        Me.tpRetourBL.ResumeLayout(False)
        Me.tpRetourBL.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpEclatement.ResumeLayout(False)
        Me.tpEclatement.PerformLayout()
        CType(Me.m_bsrcIntermédiaires, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_dgvSCMD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcSousCommande, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpFactHbv.ResumeLayout(False)
        Me.pnlFactHBV.ResumeLayout(False)
        Me.SplitFactHBV.Panel1.ResumeLayout(False)
        Me.SplitFactHBV.Panel2.ResumeLayout(False)
        CType(Me.SplitFactHBV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitFactHBV.ResumeLayout(False)
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcLgFactHBV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Méthodes Vinicom"

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)
        If m_action = vncEnums.vncfrmAction.FRMLOAD Then
            Me.tbCode.Enabled = False
            Me.tbCodeClient.Enabled = False
            Me.cbRechercheclient.Enabled = False
            liPrestashop.Enabled = False
        End If
        If m_action = vncEnums.vncfrmAction.FRMNEW Then
            Me.tbCode.Enabled = False
            Me.tbCodeClient.Enabled = True
            Me.cbRechercheclient.Enabled = True
            liPrestashop.Enabled = True
        End If
    End Sub
    Public Overrides Function setElementCourant2(ByVal pElement As Persist) As Boolean
        debAffiche()
        Dim bReturn As Boolean
        m_Tiers_Courant = CType(pElement, Commande).oTiers
        bReturn = MyBase.setElementCourant2(pElement)
        finAffiche()
        Return bReturn
    End Function 'SetElement_Specifique
    'Interface ElementCourant -> Ecran
    Public Overrides Function AfficheElement() As Boolean

        Dim objParam As Param

        Debug.Assert(Not getCommandeCourante() Is Nothing)
        Debug.Assert(Not getCommandeCourante.oTiers Is Nothing)

        'Affichage des caractéristiques de la commande
        m_bsrcCommande.Clear()
        m_bsrcCommande.Add(getCommandeCourante())

        tbCode.Text = getCommandeCourante.code
        dtDateCommande.Text = getCommandeCourante.dateCommande


        'Etat de la commande
        AfficheEtat()
        dtDateValidation.Value = getCommandeCourante.dateValidation

        DisplayStatus("Affichage de l'onglet client")

        'Client
        tbCodeClient.Text = getCommandeCourante.oTiers.code
        tbNomClient.Text = getCommandeCourante.caracteristiqueTiers.nom
        tbRaisonSociale.Text = getCommandeCourante.caracteristiqueTiers.rs
        liNomClient.Text = getCommandeCourante.caracteristiqueTiers.nom
        liNomClient.Tag = getCommandeCourante.oTiers.id

        'Adresse Client
        setAdrFactEnabled(True)
        tbAdrLiv_Nom.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.nom
        tbAdrLiv_Rue1.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.rue1
        tbAdrLiv_Rue2.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.rue2
        tbAdrLiv_cp.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.cp
        tbAdrLiv_Ville.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.ville
        tbAdrLiv_Tel.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.tel
        tbAdrLiv_Fax.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.fax
        tbAdrLiv_Port.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.port
        tbAdrLiv_Email.Text = getCommandeCourante.caracteristiqueTiers.AdresseLivraison.Email
        tbAdrFact_Nom.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.nom
        tbAdrFact_Rue1.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.rue1
        tbAdrFact_Rue2.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.rue2
        tbAdrFact_cp.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.cp
        tbAdrFact_Ville.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.ville
        tbAdrFact_Tel.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.tel
        tbAdrFact_Fax.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.fax
        tbAdrFact_Port.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.port
        tbAdrFact_Email.Text = getCommandeCourante.caracteristiqueTiers.AdresseFacturation.Email
        ckAdrIdentiques.Checked = getCommandeCourante.caracteristiqueTiers.bAdressesIdentiques
        If getCommandeCourante.caracteristiqueTiers.bAdressesIdentiques Then
            setAdrFactEnabled(False)
        Else
            setAdrFactEnabled(True)
        End If

        'Mode de reglement
        For Each objParam In cboModeReglement.Items
            If objParam.id = getCommandeCourante.caracteristiqueTiers.idModeReglement Then
                cboModeReglement.SelectedItem = objParam
                Exit For
            End If
        Next
        'Coordonnées Bancaires
        tbBanque.Text = getCommandeCourante.caracteristiqueTiers.banque
        tbRib1.Text = getCommandeCourante.caracteristiqueTiers.rib1
        tbRib2.Text = getCommandeCourante.caracteristiqueTiers.rib2
        tbRib3.Text = getCommandeCourante.caracteristiqueTiers.rib3
        tbRib4.Text = getCommandeCourante.caracteristiqueTiers.rib4

        'Activation du bouton rechercher pour les commandes enCours
        If getCommandeCourante.EtatCode = vncEtatCommande.vncEnCoursSaisie Then
            cbRechercheclient.Enabled = True
        End If


        DisplayStatus("Affichage de l'onglet commentaire")
        'Commentaires
        tbCommentaireCommande.Text = getCommandeCourante.caracteristiqueTiers.CommCommande.comment
        tbCommentaireLivraison.Text = getCommandeCourante.caracteristiqueTiers.CommLivraison.comment
        tbCommentaireFacturation.Text = getCommandeCourante.caracteristiqueTiers.CommFacturation.comment
        tbComValid.Text = getCommandeCourante.caracteristiqueTiers.CommLibre.comment

        DisplayStatus("Affichage des lignes de commandes")
        'Lignes de Commandes
        affichecolLignes()
        afficheTotaux()

        'Onglat validation
        tbComValid.Text = getCommandeCourante.caracteristiqueTiers.CommCommande.comment

        'Onglet BL
        DisplayStatus("Affichage de l'onglet BL")

        'Code Transporteur
        Dim oTRP As Transporteur
        For Each oTRP In cboCodeTRP.Items
            If oTRP.code.Equals(getCommandeCourante.oTransporteur.code) Then
                cboCodeTRP.SelectedItem = oTRP
                Exit For
            End If
        Next
        tbTrpNom.Text = getCommandeCourante.oTransporteur.AdresseLivraison.nom
        tbTrpRue1.Text = getCommandeCourante.oTransporteur.AdresseLivraison.rue1
        tbTrpRue2.Text = getCommandeCourante.oTransporteur.AdresseLivraison.rue2
        tbTrpCP.Text = getCommandeCourante.oTransporteur.AdresseLivraison.cp
        tbTrpVille.Text = getCommandeCourante.oTransporteur.AdresseLivraison.ville
        tbTrpTel.Text = getCommandeCourante.oTransporteur.AdresseLivraison.tel
        tbTrpFax.Text = getCommandeCourante.oTransporteur.AdresseLivraison.fax
        tbTrpPort.Text = getCommandeCourante.oTransporteur.AdresseLivraison.port
        tbTrpEmail.Text = getCommandeCourante.oTransporteur.AdresseLivraison.Email

        dtDateLivraison.Value = getCommandeCourante.dateLivraison
        dtDateLivraisonPrevue.Value = getCommandeCourante.dateLivraison
        dtDateLivraisonReelle.Value = getCommandeCourante.dateLivraison
        dtDateEnlev.Value = getCommandeCourante.dateEnlevement
        dtDateEnlevementReelle.Value = getCommandeCourante.dateEnlevement
        'Logistique
        'Mise a jour par le bindingSource
        'tbQteColis.Text = getCommandeCourante.qteColis
        'tbPoids.Text = getCommandeCourante.poids
        'tbQtePallPrep.Text = getCommandeCourante.qtePalettesPreparees
        'tbQtePallNonPrep.Text = getCommandeCourante.qtePalettesNonPreparees
        'tbPUPallPrep.Text = getCommandeCourante.puPalettesPreparees
        'tbPUPallNonPrep.Text = getCommandeCourante.puPalettesNonPreparees
        'tbMontantTransport.Text = getCommandeCourante.montantTransport
        'tbMontantTransport.Text = getCommandeCourante.montantTransport
        tbPiedPageBL.Text = getCommandeCourante.caracteristiqueTiers.CommLivraison.comment

        tbRefBL.Text = getCommandeCourante.refLivraison

        tbLettreVoiture.Text = getCommandeCourante().lettreVoiture
        tbCoutTransport.Text = getCommandeCourante().coutTransport
        tbRefFactTRP.Text = getCommandeCourante().refFactTRP

        crwDetailCommandeClient.ReportSource = Nothing
        crwDetailCommandeClient.Refresh()
        crwBL.ReportSource = Nothing
        crwBL.Refresh()
        '        SSTabCommandeClient.SelectedTab = tpClient


        DisplayStatus("Fin Affichage Element courant")

        Return True
    End Function 'AfficheElement
    Protected Sub AfficheEtat()
        debAffiche()
        Me.laEtatCommande.Text = getCommandeCourante.etat.libelle
        laEtatCommande.Tag = getCommandeCourante.etat.codeEtat
        ckCmdValide.Enabled = True

        'Si la commande est Livrée on ne peu plus changer son type
        ' car cela pose des problèmes de stock

        Select Case getCommandeCourante.etat.codeEtat
            Case vncEnums.vncEtatCommande.vncEnCoursSaisie
                ckCmdValide.Checked = False
                dtDateValidation.Enabled = False
                ckCmdValide.Enabled = True
                'Type de Commande 
                rbTypeCmdDirecte.Enabled = True
                rbTypeCmdPlateforme.Enabled = True
            Case vncEnums.vncEtatCommande.vncValidee
                ckCmdValide.Checked = True
                dtDateValidation.Enabled = True
                ckCmdValide.Enabled = True
                'Type de Commande 
                rbTypeCmdDirecte.Enabled = True
                rbTypeCmdPlateforme.Enabled = True
            Case vncEnums.vncEtatCommande.vncLivree
                ckCmdValide.Checked = True
                dtDateValidation.Enabled = False
                ckCmdValide.Enabled = False
                'Type de Commande 
                rbTypeCmdDirecte.Enabled = False
                rbTypeCmdPlateforme.Enabled = False
            Case vncEnums.vncEtatCommande.vncEclatee
                ckCmdValide.Checked = True
                dtDateValidation.Enabled = False
                ckCmdValide.Enabled = False
                'Type de Commande 
                rbTypeCmdDirecte.Enabled = False
                rbTypeCmdPlateforme.Enabled = False
            Case vncEnums.vncEtatCommande.vncTransmiseQuadra
                ckCmdValide.Checked = True
                dtDateValidation.Enabled = False
                ckCmdValide.Enabled = False
                'Type de Commande 
                rbTypeCmdDirecte.Enabled = False
                rbTypeCmdPlateforme.Enabled = False
            Case vncEnums.vncEtatCommande.vncRapprochee
                ckCmdValide.Checked = True
                dtDateValidation.Enabled = False
                ckCmdValide.Enabled = False
                'Type de Commande 
                rbTypeCmdDirecte.Enabled = False
                rbTypeCmdPlateforme.Enabled = False
        End Select
        finAffiche()
    End Sub 'AfficheEtat
    Public Overrides Function MAJElement() As Boolean
        Debug.Assert(Not getCommandeCourante() Is Nothing, "Pas d'element courant")
        Dim bReturn As Boolean
        Dim nid As Integer

        bReturn = True
        If bReturn Then
            Try
                '        getCommandeCourante.code = tbCode.Text
                getCommandeCourante.dateCommande = dtDateCommande.Text

                'L'Etat de la commande est mise à jour en tant reel

                If IsDate(dtDateValidation.Value) Then
                    getCommandeCourante.dateValidation = dtDateValidation.Value
                End If

                'Rattachement du Tiers à la commande
                nid = liNomClient.Tag
                'getCommandeCourante.oTiers = m_Tiers_Courant

                getCommandeCourante.caracteristiqueTiers.nom = tbNomClient.Text
                getCommandeCourante.caracteristiqueTiers.rs = tbRaisonSociale.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.nom = tbAdrLiv_Nom.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.rue1 = tbAdrLiv_Rue1.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.rue2 = tbAdrLiv_Rue2.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.cp = tbAdrLiv_cp.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.ville = tbAdrLiv_Ville.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.tel = tbAdrLiv_Tel.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.fax = tbAdrLiv_Fax.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.port = tbAdrLiv_Port.Text
                getCommandeCourante.caracteristiqueTiers.AdresseLivraison.Email = tbAdrLiv_Email.Text

                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.nom = tbAdrFact_Nom.Text
                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.rue1 = tbAdrFact_Rue1.Text
                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.rue2 = tbAdrFact_Rue2.Text
                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.cp = tbAdrFact_cp.Text
                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.ville = tbAdrFact_Ville.Text
                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.tel = tbAdrFact_Tel.Text
                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.fax = tbAdrFact_Fax.Text
                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.port = tbAdrFact_Port.Text
                getCommandeCourante.caracteristiqueTiers.AdresseFacturation.Email = tbAdrFact_Email.Text

                getCommandeCourante.caracteristiqueTiers.bAdressesIdentiques = ckAdrIdentiques.Checked
                getCommandeCourante.caracteristiqueTiers.banque = tbBanque.Text
                getCommandeCourante.caracteristiqueTiers.rib1 = tbRib1.Text
                getCommandeCourante.caracteristiqueTiers.rib2 = tbRib2.Text
                getCommandeCourante.caracteristiqueTiers.rib3 = tbRib3.Text
                getCommandeCourante.caracteristiqueTiers.rib4 = tbRib4.Text

                If Not cboModeReglement.SelectedItem Is Nothing Then
                    getCommandeCourante.caracteristiqueTiers.idModeReglement = cboModeReglement.SelectedItem.id
                End If

                'Commentaires
                getCommandeCourante.caracteristiqueTiers.CommCommande.comment = tbCommentaireCommande.Text
                getCommandeCourante.caracteristiqueTiers.CommLivraison.comment = tbCommentaireLivraison.Text
                getCommandeCourante.caracteristiqueTiers.CommFacturation.comment = tbCommentaireFacturation.Text
                getCommandeCourante.caracteristiqueTiers.CommLibre.comment = tbComValid.Text

                'Lignes du tableau
                'la collection est mise à jour à chaque modification du tableau

                '        getCommandeCourante.totalHT = tbTotalHT.Text'BindingSource
                '       getCommandeCourante.totalTTC = tbTotalTTC.Text'BindingSource

                'Onglet BL
                If cboCodeTRP.Text <> "" Then
                    Dim oTRP As Transporteur
                    Dim nidTrp As Long

                    Try
                        oTRP = Transporteur.colTransporteur.Item(cboCodeTRP.Text)
                        nidTrp = oTRP.id
                        getCommandeCourante.oTransporteur = New Transporteur
                        If (nidTrp) <> 0 Then
                            getCommandeCourante.oTransporteur.load(nidTrp)
                        End If

                        getCommandeCourante.oTransporteur.nom = tbTrpNom.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.nom = tbTrpNom.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.rue1 = tbTrpRue1.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.rue2 = tbTrpRue2.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.cp = tbTrpCP.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.ville = tbTrpVille.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.tel = tbTrpTel.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.fax = tbTrpFax.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.port = tbTrpPort.Text
                        getCommandeCourante.oTransporteur.AdresseLivraison.Email = tbTrpEmail.Text
                    Catch ex As Exception
                        getCommandeCourante.oTransporteur = Transporteur.TransporteurDefault

                    End Try
                End If

                If IsDate(dtDateLivraisonReelle.Value) Then
                    getCommandeCourante.dateLivraison = dtDateLivraisonReelle.Value
                End If
                If IsDate(dtDateEnlevementReelle.Value) Then
                    getCommandeCourante.dateEnlevement = dtDateEnlevementReelle.Value
                End If

                '      getCommandeCourante.qteColis = tbQteColis.Text 'BindingSource
                '     getCommandeCourante.poids = tbPoids.Text 'BindingSource
                'getCommandeCourante.qtePalettesPreparees = tbQtePallPrep.Text
                'getCommandeCourante.qtePalettesNonPreparees = tbQtePallNonPrep.Text
                'getCommandeCourante.puPalettesPreparees = tbPUPallPrep.Text
                'getCommandeCourante.puPalettesNonPreparees = tbPUPallNonPrep.Text
                'getCommandeCourante.montantTransport = tbMontantTransport.Text

                getCommandeCourante.lettreVoiture = tbLettreVoiture.Text
                getCommandeCourante.coutTransport = tbCoutTransport.Text
                getCommandeCourante.refFactTRP = tbRefFactTRP.Text

                getCommandeCourante.refLivraison = tbRefBL.Text
            Catch ex As Exception
                bReturn = False
            End Try
        End If

        Return bReturn
    End Function 'MAJElement

    Private Sub selectTransporteur()
        Dim oTRP As Transporteur

        Try
            If Not cboCodeTRP.SelectedItem Is Nothing Then
                oTRP = cboCodeTRP.SelectedItem
                debAffiche()
                tbTrpNom.Text = oTRP.nom
                tbTrpRue1.Text = oTRP.AdresseLivraison.rue1
                tbTrpRue2.Text = oTRP.AdresseLivraison.rue2
                tbTrpCP.Text = oTRP.AdresseLivraison.cp
                tbTrpVille.Text = oTRP.AdresseLivraison.ville
                tbTrpTel.Text = oTRP.AdresseLivraison.tel
                tbTrpFax.Text = oTRP.AdresseLivraison.fax
                tbTrpPort.Text = oTRP.AdresseLivraison.port
                tbTrpEmail.Text = oTRP.AdresseLivraison.Email
                finAffiche()
            End If
        Catch ex As Exception

        End Try
    End Sub 'SelectTransporteur

    Public Overrides Function ControleAvantSauvegarde() As Boolean
        Dim bReturn As Boolean
        Debug.Assert(Not getCommandeCourante() Is Nothing, "Element Courant Requis")


        bReturn = isStringNum(tbAdrLiv_cp.Text)
        bReturn = bReturn And isStringNum(tbAdrLiv_Tel.Text)
        bReturn = bReturn And isStringNum(tbAdrLiv_Fax.Text)
        bReturn = bReturn And isStringNum(tbAdrLiv_Port.Text)
        bReturn = bReturn And isStringNum(tbAdrFact_cp.Text)
        bReturn = bReturn And isStringNum(tbAdrFact_Tel.Text)
        bReturn = bReturn And isStringNum(tbAdrFact_Fax.Text)
        bReturn = bReturn And isStringNum(tbAdrFact_Port.Text)

        If Not IsNumeric(liNomClient.Tag) Or liNomClient.Tag = 0 Then
            DisplayError("ControleAvantSauvegarde", "Le Client n'est pas renseigné")
            Return False
        End If
        Return bReturn
    End Function
    'Public Overrides Function getResume() As String 'Rend le caption de la fenêtre
    '    If (Not DesignMode) Then
    '        Return "CMD" & getCommandeCourante.shortResume
    '    Else
    '        Return String.Empty
    '    End If
    'End Function 'getResume
    'Descative les controles de la fenêtre 
    'Descative les controles de la fenêtre 
    Protected Overrides Function frmNew() As Boolean
        Dim breturn As Boolean
        breturn = MyBase.frmNew()
        'Le code est non saisissable en création 
        SSTabCommandeClient.SelectedTab = tpClient
        tbCode.Enabled = False
        tbCodeClient.Focus()
        Return breturn
    End Function
    Protected Overrides Function frmDel() As Boolean
        MyBase.frmDel()
    End Function ' frmDel

    Protected Overrides Sub setfrmUpdated()
        If Not getCommandeCourante() Is Nothing Then
            MAJElement()
            MyBase.setfrmUpdated()
        End If
    End Sub
    'Controle des Actions à Réaliser au moment de la sauvegarde
    Protected Overrides Function frmSave() As Boolean
        Dim bFenetreProduitOuverte As Boolean = False
        Dim bPremiereFenetreProduitOuverte As Boolean = False
        Dim bReturn As Boolean

        Trace.WriteLine(Now() & "Sauvegarde de la commande N°" & getCommandeCourante.code)
        Trace.Indent()

        If getbGestionMvtStock() Then
            If getCommandeCourante.getActionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer <> vncEnums.vncGenererSupprimer.vncRien Then
                'Controle s'il y a des fenêtre frmProduit ouvertes 
                Dim x As Integer
                Dim oFrmProduit As frmProduit
                For x = 0 To (Me.MdiParent.MdiChildren.Length) - 1
                    Try
                        oFrmProduit = CType(Me.MdiParent.MdiChildren(x), frmProduit)
                        If (Not bPremiereFenetreProduitOuverte) Then
                            ' Message d'alerte sur la première fenêtre
                            If MsgBox("Attention, vous avez des fenêtre produit ouvertes sur des produits, Veuillez les fermer avant de poursuivre ", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                                Exit Function
                            End If
                            bPremiereFenetreProduitOuverte = True
                        End If
                    Catch ex As Exception

                    End Try
                Next x
            End If
            'Mouvement de Stocks
            'If getCommandeCourante.getActionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer Then
            '    If MsgBox("Attention, la sauvegarde de l'élement courant entrainera la génération de mouvements de stocks, voulez-vous poursuivre", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            '        Exit Function
            '    End If
            'End If
            'If getCommandeCourante.getActionMvtStock = vncEnums.vncGenererSupprimer.vncSupprimer Then
            '    If MsgBox("Attention, la sauvegarde de l'élement courant entrainera la suppression des mouvements de stocks, voulez-vous poursuivre", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            '        Exit Function
            '    End If
            'End If
        End If

        If getbGestionSousCommande() Then
            'If getCommandeCourante.getActionSousCommande = vncEnums.vncGenererSupprimer.vncGenerer Then
            '    If MsgBox("Attention, la sauvegarde de l'élement courant entrainera la génération de sous-commandes, voulez-vous poursuivre", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            '        Exit Function
            '    End If
            'End If
            'If getCommandeCourante.getActionSousCommande = vncEnums.vncGenererSupprimer.vncSupprimer Then
            '    If MsgBox("Attention, la sauvegarde de l'élement courant entrainera la suppression des sous-commandes, voulez-vous poursuivre", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            '        Exit Function
            '    End If
            'End If
        End If

        bReturn = MyBase.frmSave()
        Trace.Unindent()
        Return bReturn
    End Function 'frmSave

    '==============================================================================================
    '==============================================================================================
    '==============================================================================================
#End Region
#Region "Methodes Privées"
    Protected Overridable Sub ActionLivrer()

    End Sub

    Protected Overridable Sub ActionAnnLivrer()

    End Sub
    Protected Sub initFenetre()
        debAffiche()
        crwDetailCommandeClient.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        crwDetailCommandeClient.DisplayToolbar = True
        crwBL.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        crwBL.DisplayToolbar = True
        InitComboModeReglement(cboModeReglement)
        initcboCodeTRP(Me.cboCodeTRP)
        finAffiche()
    End Sub 'initFenetre

    'Affiche la boite de dialogue d'ajout de ligne
    Protected Function ajouteruneLigne() As Boolean
        Debug.Assert(Not getCommandeCourante() Is Nothing)
        Debug.Assert(Not getCommandeCourante.oTiers Is Nothing)
        Dim odlg As dlgLgCommande
        Dim objLg As LgCommande
        Dim bReturn As Boolean


        objLg = New LgCommande(getCommandeCourante.id)
        objLg.num = getCommandeCourante.getNextNumLg()
        odlg = New dlgLgCommande
        odlg.setElementCourant(objLg, getCommandeCourante.oTiers)
        'Détermination du type de produit à Recchercher
        ' Commande Plateforme et BA = Produit Plateforme
        'Commande Directe => Tous
        odlg.setRechercheTypeProduit(vncEnums.vncTypeProduit.vncPlateforme)
        If m_TypeDonnees = vncEnums.vncTypeDonnee.COMMANDECLIENT Then
            If CType(getCommandeCourante(), CommandeClient).typeCommande = vncEnums.vncTypeCommande.vncCmdClientDirecte Then
                odlg.setRechercheTypeProduit(vncEnums.vncTypeProduit.vncTous)
            End If
        End If

        If odlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            objLg = getCommandeCourante.AjouteLigne(objLg, True)
            m_bsrcLignesCommande.Add(objLg)
            CalcPoidsColis()
            afficheTotaux()
        End If
        bReturn = True

        Return bReturn

    End Function 'Ajouteruneligne
    'Affiche la boite de dialogue de modification de ligne
    Protected Function modifieruneLigne() As Boolean
        Debug.Assert(Not getCommandeCourante() Is Nothing)
        Debug.Assert(Not getCommandeCourante.oTiers Is Nothing)
        Dim odlg As dlgLgCommande
        Dim objLg As LgCommande
        Dim bReturn As Boolean

        Try
            If Not m_bsrcLignesCommande.Current Is Nothing Then
                objLg = m_bsrcLignesCommande.Current
                odlg = New dlgLgCommande
                odlg.setElementCourant(objLg, getCommandeCourante.oTiers)
                If odlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    setcursorWait()
                    m_bsrcLignesCommande.ResetBindings(False)
                    If getCommandeCourante.estPartiellementLivree Then
                        ActionLivrer()
                    End If
                    getCommandeCourante.calculPrixTotal()
                    afficheTotaux()
                    AfficheEtat()
                    restoreCursor()
                End If
                setfrmUpdated()
            End If
            bReturn = True
        Catch ex As Exception
            DisplayError("frmSaisieCommande.modifierUneLigne", "Impossible de charger la ligne" & ex.ToString)
            bReturn = False
        End Try
        Return bReturn
    End Function 'Modifieruneligne
    'Affiche la boite de dialogue de modification de ligne
    Protected Function supprimeruneLigne() As Boolean
        Debug.Assert(Not getCommandeCourante() Is Nothing)
        Dim objLg As LgCommande
        Dim bReturn As Boolean

        Try
            If Not (m_bsrcLignesCommande.Current Is Nothing) Then
                objLg = m_bsrcLignesCommande.Current
                If MsgBox("Etes-vous sur de vouloir supprimer la ligne " & objLg.oProduit.code, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    getCommandeCourante.supprimeLigne(objLg.num)
                    m_bsrcLignesCommande.RemoveCurrent()
                End If
                setfrmUpdated()
            End If
            bReturn = True
        Catch ex As Exception
            bReturn = False
            DisplayError("frmSaisieCommande.SupprimerUneLigne", "Impossible de charger la ligne")
        End Try
        getCommandeCourante.calculPrixTotal()
        debAffiche()
        affichecolLignes()
        afficheTotaux()
        finAffiche()
        Return bReturn
    End Function 'Supprimeruneligne

    Protected Sub afficheTotaux()
        debAffiche()
        m_bsrcCommande.ResetCurrentItem()

        ' Ces Champs sont affiché grace au binding Source
        'tbTotalHT.Text = getCommandeCourante.totalHT
        'tbTotalTTC.Text = getCommandeCourante.totalTTC
        'tbQteColis.Text = getCommandeCourante.qteColis
        'tbPoids.Text = getCommandeCourante.poids
        'tbMtComm.Text = getCommandeCourante.montantCommission
        finAffiche()
    End Sub 'Affiche des totaux HT et TTC
    Protected Function affichecolLignes() As Boolean
        Debug.Assert(Not getCommandeCourante() Is Nothing)
        Debug.Assert(Not getCommandeCourante.colLignes Is Nothing)

        Dim objLg As LgCommande
        Dim nRow As Integer
        debAffiche()
        nRow = 0

        m_bsrcLignesCommande.Clear()
        For Each objLg In getCommandeCourante.colLignes
            nRow = nRow + 1
            DisplayStatus("Affichage ligne " & nRow)
            m_bsrcLignesCommande.Add(objLg)
        Next objLg
        afficheTotaux()
        m_bsrcLignesCommande.ResetBindings(False)
        finAffiche()
    End Function ' Affichage de des lignes de commandes
    ''Affiche l'entete du tableau des lignes de commande
    'Private Sub displayLstLignes(ByVal flx As AxMSFlexGridLib.AxMSFlexGrid)

    '    Dim nUnit As Integer
    '    flx.Cols = vncColLigneCommande.COL_NBRECOL
    '    nUnit = ((flx.Width - SCROLLBARWIDTH) / 32) * 16
    '    flx.Rows = 2
    '    flx.FixedRows = 1
    '    flx.Row = 0

    '    flx.Col = vncColLigneCommande.COL_NUM
    '    flx.Text = "Num."
    '    flx.set_ColWidth(vncColLigneCommande.COL_NUM, nUnit)

    '    flx.Col = vncColLigneCommande.COL_CODEPRODUIT
    '    flx.Text = "code Produit"
    '    flx.set_ColWidth(vncColLigneCommande.COL_CODEPRODUIT, nUnit * 2)

    '    flx.Col = vncColLigneCommande.COL_DESIGNATIONPRODUIT
    '    flx.Text = "Désignation Produit"
    '    flx.set_ColWidth(vncColLigneCommande.COL_DESIGNATIONPRODUIT, nUnit * 8)

    '    flx.Col = vncColLigneCommande.COL_MILLESIME
    '    flx.Text = "Mill."
    '    flx.set_ColWidth(vncColLigneCommande.COL_MILLESIME, nUnit)

    '    flx.Col = vncColLigneCommande.COL_CONDITIONNEMENT
    '    flx.Text = "Cond."
    '    flx.set_ColWidth(vncColLigneCommande.COL_CONDITIONNEMENT, nUnit * 2)

    '    flx.Col = vncColLigneCommande.COL_COULEUR
    '    flx.Text = "Coul."
    '    flx.set_ColWidth(vncColLigneCommande.COL_COULEUR, nUnit * 2)

    '    flx.Col = vncColLigneCommande.COL_QTE_COM
    '    flx.Text = "Qte Comm"
    '    flx.set_ColWidth(vncColLigneCommande.COL_QTE_COM, nUnit * 2)

    '    flx.Col = vncColLigneCommande.COL_QTE_LIV
    '    flx.Text = "Qte Liv"
    '    flx.set_ColWidth(vncColLigneCommande.COL_QTE_LIV, nUnit * 2)

    '    flx.Col = vncColLigneCommande.COL_QTE_FACT
    '    flx.Text = "Qte Fact"
    '    flx.set_ColWidth(vncColLigneCommande.COL_QTE_FACT, nUnit * 2)

    '    flx.Col = vncColLigneCommande.COL_GRATUIT
    '    flx.Text = "Gratuit"
    '    flx.set_ColWidth(vncColLigneCommande.COL_GRATUIT, nUnit / 2)

    '    flx.Col = vncColLigneCommande.COL_PRIX_U
    '    flx.Text = "Prix U"
    '    flx.set_ColWidth(vncColLigneCommande.COL_PRIX_U, nUnit * 2)

    '    flx.Col = vncColLigneCommande.COL_PRIX_HT
    '    flx.Text = "Montant HT"
    '    flx.set_ColWidth(vncColLigneCommande.COL_PRIX_HT, nUnit * 4)

    '    flx.Col = vncColLigneCommande.COL_PRIX_TTC
    '    flx.Text = "Montant TTC"
    '    flx.set_ColWidth(vncColLigneCommande.COL_PRIX_TTC, nUnit * 4)

    '    flx.Col = vncColLigneCommande.COL_IDPRODUIT
    '    flx.Text = "id"
    '    flx.set_ColWidth(vncColLigneCommande.COL_IDPRODUIT, 0)
    'End Sub

    'Rnd Vrai si un Tiers existe déjà pour cette commande
    Protected Function tiersdejaCharge() As Boolean
        Dim bReturn As Boolean
        If laIdClient.Text = "" Then
            bReturn = False
        Else
            bReturn = True
        End If
        Return bReturn
    End Function 'tiersdejacharge
    'Affiche les caractéristiques d'un client
    Protected Function afficheTiers(ByVal objClient As Tiers) As Boolean
        Dim bReturn As Boolean
        Dim objParam As Param
        debAffiche()
        liNomClient.Text = objClient.nom
        liNomClient.Tag = objClient.id

        tbCodeClient.Text = objClient.code
        tbNomClient.Text = objClient.nom
        tbRaisonSociale.Text = objClient.rs
        tbAdrLiv_Nom.Text = objClient.AdresseLivraison.nom
        tbAdrLiv_Rue1.Text = objClient.AdresseLivraison.rue1
        tbAdrLiv_Rue2.Text = objClient.AdresseLivraison.rue2
        tbAdrLiv_cp.Text = objClient.AdresseLivraison.cp
        tbAdrLiv_Ville.Text = objClient.AdresseLivraison.ville
        tbAdrLiv_Tel.Text = objClient.AdresseLivraison.tel
        tbAdrLiv_Fax.Text = objClient.AdresseLivraison.fax
        tbAdrLiv_Port.Text = objClient.AdresseLivraison.port
        tbAdrLiv_Email.Text = objClient.AdresseLivraison.Email


        tbAdrFact_Nom.Text = objClient.AdresseFacturation.nom
        tbAdrFact_Rue1.Text = objClient.AdresseFacturation.rue1
        tbAdrFact_Rue2.Text = objClient.AdresseFacturation.rue2
        tbAdrFact_cp.Text = objClient.AdresseFacturation.cp
        tbAdrFact_Ville.Text = objClient.AdresseFacturation.ville
        tbAdrFact_Tel.Text = objClient.AdresseFacturation.tel
        tbAdrFact_Fax.Text = objClient.AdresseFacturation.fax
        tbAdrFact_Port.Text = objClient.AdresseFacturation.port
        tbAdrFact_Email.Text = objClient.AdresseFacturation.Email

        ckAdrIdentiques.Checked = objClient.bAdressesIdentiques

        'Commentaires
        tbCommentaireCommande.Text = objClient.CommCommande.comment
        tbCommentaireLivraison.Text = objClient.CommLivraison.comment
        tbCommentaireFacturation.Text = objClient.CommFacturation.comment
        tbPiedPageBL.Text = objClient.CommLivraison.comment
        tbComValid.Text = objClient.CommCommande.comment

        For Each objParam In cboModeReglement.Items
            If objParam.id = objClient.idModeReglement Then
                cboModeReglement.SelectedItem = objParam
                Exit For
            End If
        Next

        tbBanque.Text = objClient.banque
        tbRib1.Text = objClient.rib1
        tbRib2.Text = objClient.rib2
        tbRib3.Text = objClient.rib3
        tbRib4.Text = objClient.rib4

        setfrmUpdated()

        finAffiche()
        Return bReturn
    End Function 'Affichetiers
    'Met à jour la Propriété Enabled des controles de la fentre
    'Private Sub setEnabled(ByVal pbEnabled As Boolean)

    '    tbCode.Enabled = pbEnabled
    '    dtDateCommande.Enabled = pbEnabled
    '    rb_TypeTRP_Avance.Enabled = pbEnabled
    '    rb_TypeTRP_Depart.Enabled = pbEnabled
    '    rb_TypeTrp_Franco.Enabled = pbEnabled
    '    rbTypeCmdDirecte.Enabled = pbEnabled
    '    rbTypeCmdPlateforme.Enabled = pbEnabled
    '    tbCodeClient.Enabled = pbEnabled
    '    cbRechercheclient.Enabled = pbEnabled
    '    tbNomClient.Enabled = pbEnabled
    '    tbRaisonSociale.Enabled = pbEnabled
    '    tbAdrLiv_Nom.Enabled = pbEnabled
    '    tbAdrLiv_Rue1.Enabled = pbEnabled
    '    tbAdrLiv_Rue2.Enabled = pbEnabled
    '    tbAdrLiv_cp.Enabled = pbEnabled
    '    tbAdrLiv_Ville.Enabled = pbEnabled
    '    tbAdrLiv_Tel.Enabled = pbEnabled
    '    tbAdrLiv_Fax.Enabled = pbEnabled
    '    tbAdrLiv_Port.Enabled = pbEnabled
    '    tbAdrLiv_Email.Enabled = pbEnabled
    '    tbAdrFact_Nom.Enabled = pbEnabled
    '    tbAdrFact_Rue1.Enabled = pbEnabled
    '    tbAdrFact_Rue2.Enabled = pbEnabled
    '    tbAdrFact_cp.Enabled = pbEnabled
    '    tbAdrFact_Ville.Enabled = pbEnabled
    '    tbAdrFact_Tel.Enabled = pbEnabled
    '    tbAdrFact_Fax.Enabled = pbEnabled
    '    tbAdrFact_Port.Enabled = pbEnabled
    '    tbAdrFact_Email.Enabled = pbEnabled
    '    cboModeReglement.Enabled = pbEnabled
    '    tbBanque.Enabled = pbEnabled
    '    tbRib1.Enabled = pbEnabled
    '    tbRib2.Enabled = pbEnabled
    '    tbRib3.Enabled = pbEnabled
    '    tbRib4.Enabled = pbEnabled

    '    SSTabCommandeClient.Enabled = pbEnabled

    'End Sub
    Private Sub rechercheTiers()
        Dim objTiers As Tiers
        Dim objCommande As Commande

        If tiersdejaCharge() Then
            If MsgBox("Cette commande a déja un tiers d'affecté. Souhaitez-vous le remplacer", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        objCommande = CType(getElementCourant(), Commande)
        objTiers = rechercheDonnee(objCommande.oTiers.typeDonnee, tbCodeClient)
        If Not objTiers Is Nothing Then
            If objTiers.bResume Then
                objTiers.load()
            End If

            If Not objTiers Is Nothing Then
                afficheTiers(objTiers)
                setTiers(objTiers)
            End If
        End If
    End Sub 'Rechercheclient
    Protected Overridable Sub setTiers(poTiers As Tiers)
        m_Tiers_Courant = poTiers
        getCommandeCourante.setTiers(poTiers)
    End Sub

    'Dupplique l'adresse de livraison sur l'adresse de facturation
    Private Sub duppliqueAdresses()
        If Not bAffichageEnCours() Then
            debAffiche()
            tbAdrFact_Nom.Text = tbAdrLiv_Nom.Text
            tbAdrFact_Rue1.Text = tbAdrLiv_Rue1.Text
            tbAdrFact_Rue2.Text = tbAdrLiv_Rue2.Text
            tbAdrFact_cp.Text = tbAdrLiv_cp.Text
            tbAdrFact_Ville.Text = tbAdrLiv_Ville.Text
            tbAdrFact_Tel.Text = tbAdrLiv_Tel.Text
            tbAdrFact_Fax.Text = tbAdrLiv_Fax.Text
            tbAdrFact_Port.Text = tbAdrLiv_Port.Text
            tbAdrFact_Email.Text = tbAdrLiv_Email.Text
            finAffiche()
        End If
    End Sub 'DuppliqueAdresses
    'Active ou désactive les zones d'adresses de facturation
    Private Sub setAdrFactEnabled(ByVal bEnabled As Boolean)
        grpAdrFact.Enabled = bEnabled
    End Sub

    Private Sub initTPValidCMD()
        tbNumFaxValidation.Text = tbAdrLiv_Fax.Text
    End Sub
    Public Overridable Sub initTPBL()
        debAffiche()
        crwBL.Enabled = False
        cbVisuBL.Enabled = False
        grpTransporteur.Enabled = False
        dtDateEnlev.Enabled = False
        dtDateLivraison.Enabled = False
        crwBL.Enabled = True
        cbVisuBL.Enabled = True
        grpTransporteur.Enabled = True
        dtDateEnlev.Enabled = True
        dtDateLivraison.Enabled = True
        cbMailBLPLTFRM.Enabled = True
        cbFaxerBLTransporteur.Enabled = True
        tbPiedPageBL.Text = tbCommentaireLivraison.Text
        finAffiche()
    End Sub

    Public Sub CalcPoidsColis()
        Debug.Assert(Not getCommandeCourante() Is Nothing)

        getCommandeCourante.CalcPoidsColis()

        afficheTotaux()

    End Sub

    '***********************************************************************************************************************************************************************
    Protected Sub calculMontantTransport()
        'Dim nQtePalPrep As Decimal
        'Dim nQtePalNonPrep As Decimal
        'Dim npuPalPrep As Decimal
        'Dim npuPalNonPrep As Decimal
        'Dim nMontant As Decimal

        Try
            Trace.WriteLine("calculMontantTransport")
            getCommandeCourante.CalcMontantTransport()
            m_bsrcCommande.ResetCurrentItem()
            'nQtePalPrep = tbQtePallPrep.Text
            'nQtePalNonPrep = tbQtePallNonPrep.Text
            'npuPalPrep = tbPUPallPrep.Text
            'npuPalNonPrep = tbPUPallNonPrep.Text

            'nMontant = (nQtePalPrep * npuPalPrep) + (nQtePalNonPrep * npuPalNonPrep)
            'tbMontantTransport.Text = nMontant
        Catch ex As Exception

        End Try

    End Sub
#End Region 'Méthodes Privées
#Region "Gestion Evenements"
    'Affichage de l'entete de la liste des lignes
    Private Sub cbAjouterLigne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAjouterLigne.Click
        If ajouteruneLigne() Then
            setfrmUpdated()
        End If
    End Sub
    Private Sub cbRechercheclient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRechercheclient.Click
        rechercheTiers()
    End Sub

    Private Sub liNomClient_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liNomClient.LinkClicked
        afficheFenetreClient(liNomClient.Tag)
    End Sub 'liNomClient_LinkClicked

    Private Sub tbCodeClient_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbCodeClient.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            rechercheTiers()
        End If

    End Sub

    Private Sub ckAdrIdentiques_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckAdrIdentiques.CheckedChanged
        If ckAdrIdentiques.Checked Then
            duppliqueAdresses()
            setAdrFactEnabled(False)
        Else
            setAdrFactEnabled(True)
        End If
    End Sub



    Private Sub cbModifierLigne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbModifierLigne.Click
        If Not m_bsrcLignesCommande.Current Is Nothing Then
            If modifieruneLigne() Then
                CalcPoidsColis()
                setfrmUpdated()
            End If
        End If
    End Sub

    Private Sub cbSupprimerLigne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSupprimerLigne.Click
        If Not m_bsrcLignesCommande.Current Is Nothing Then
            If supprimeruneLigne() Then
                setfrmUpdated()
            End If
        End If

    End Sub



    Protected Overridable Sub modifieruneLigneBL()
        modifieruneLigne()
    End Sub



    Private Sub SSTabCommandeClient_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SSTabCommandeClient.Click
        Try
            If Not DesignMode Then
                If getCommandeCourante.oTiers.id = 0 Then
                    MsgBox("Le Tiers (Client ou Forunisseur) doit être renseigné au préalable")
                    SSTabCommandeClient.SelectedTab = tpClient
                End If
                sauvegardeElementCourant()
                Select Case SSTabCommandeClient.SelectedTab.Name
                    Case tpValidCmd.Name
                        initTPValidCMD()
                    Case tpBL.Name
                        initTPBL()
                    Case tpEclatement.Name
                        initTPEclatement()
                End Select
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ckCmdValide_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckCmdValide.CheckedChanged
        If Not bAffichageEnCours() Then
            valideCommande()
        End If
    End Sub

    Private Sub cbVisuDetailCommande_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVisuDetailCommande.Click
        Dim oldCursor As Cursor
        sauvegardeElementCourant()
        oldCursor = Me.Cursor
        Me.Cursor = WaitCursor
        affichecrDetailCommande()
        Me.Cursor = oldCursor
    End Sub

    Private Sub cbFax_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFax.Click
        faxerValidationCommande()
    End Sub

    Private Sub tbRecalcTotaux_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRecalcTotaux.Click
        Debug.Assert(Not getCommandeCourante() Is Nothing, "Commande courante requiered")
        getCommandeCourante.calculPrixTotal()
        'Dans le cas d'une commande client on recalcule aussi le poids et la quantité de colis
        If m_TypeDonnees = vncEnums.vncTypeDonnee.COMMANDECLIENT Then
            CType(getCommandeCourante(), CommandeClient).CalcPoidsColis()
        End If
        afficheTotaux()
    End Sub

    Private Sub tbNumFaxValidation_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbNumFaxValidation.TextChanged
        If Trim(tbNumFaxValidation.Text) <> "" Then
            cbFax.Enabled = True
        Else
            cbFax.Enabled = False
        End If
    End Sub



    Private Sub cbVisuBL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVisuBL.Click
        Dim oldCursor As Cursor
        sauvegardeElementCourant()
        oldCursor = Me.Cursor
        Me.Cursor = WaitCursor
        affichecrBL()
        Me.Cursor = oldCursor
    End Sub

    Private Sub tbPiedPageBL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPiedPageBL.TextChanged
        If Not bAffichageEnCours() Then
            debAffiche()
            tbCommentaireLivraison.Text = tbPiedPageBL.Text
            finAffiche()
        End If
    End Sub

    Private Sub tbCommentaireLivraison_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCommentaireLivraison.TextChanged
        If Not bAffichageEnCours() Then
            debAffiche()
            tbPiedPageBL.Text = tbCommentaireLivraison.Text
            finAffiche()
        End If
    End Sub
    Private Sub cbFaxBLPLTFRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMailBLPLTFRM.Click
        exporterWEBEDI()
    End Sub

    Private Sub cbFaxerBLTransporteur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFaxerBLTransporteur.Click
        faxerBL(tbTrpFax.Text, getCommandeCourante.oTransporteur)
    End Sub
    Private Sub dtDateLivraison_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtDateLivraison.ValueChanged
        dtDateLivraisonReelle.Value = dtDateLivraison.Value
    End Sub


    Private Sub dtDateEnlev_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtDateEnlev.ValueChanged
        dtDateEnlevementReelle.Value = dtDateEnlev.Value
    End Sub

    Private Sub cbBLToutOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBLToutOK.Click
        Call blToutOK()
    End Sub

    Private Sub cbAnnulerLivraison_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnulerLivraison.Click
        annulerLivraison()
    End Sub

    Private Sub cbEclatementCmde_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbEclatementCmde.Click
        sauvegardeElementCourant()
        eclatementCommande()
        setfrmUpdated()
    End Sub
    Private Sub dtDateLivraisonPrevue_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtDateLivraisonPrevue.ValueChanged
        If Not bAffichageEnCours() Then
            dtDateLivraison.Value = dtDateLivraisonPrevue.Value
        End If
    End Sub

    Private Sub cbValider_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not bAffichageEnCours() Then
            If Not m_bsrcSousCommande.Current Is Nothing Then
                validerSousCommande()

            End If
        End If

    End Sub
    Private Sub cbSCMDFaxer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not m_bsrcSousCommande.Current Is Nothing Then
            faxerSousCommande()
        End If

    End Sub

    Private Sub cbSCMDValiderTout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        validerToutesLesSousCommandes()
    End Sub

    Private Sub cbSCMDFaxerTout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSCMDFaxerTout.Click
        faxerTouteslesSousCommandes()
    End Sub
    Private Sub cbSCMDVoir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSCMDVoir.Click
        If Not m_bsrcSousCommande.Current Is Nothing Then
            visualiserSousCommande()
        End If
    End Sub
    'Private Sub rbTypeCmdPlateforme_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbTypeCmdPlateforme.CheckedChanged
    'Modif du 22/11/06 : Dans tous les cas on peut modifier le transporteur
    '        If rbTypeCmdPlateforme.Checked Then
    '        grpTransporteur.Enabled = True
    '        Else
    '           grpTransporteur.Enabled = True
    'End If
    'End Sub

    Private Sub liSCMDFournisseur_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liSCMDFournisseur.LinkClicked
        afficheFenetreFournisseur(liSCMDFournisseur.Tag)
    End Sub
    Private Sub cbAnnEclatement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnEclatement.Click
        If (Not DesignMode) Then
            sauvegardeElementCourant()
            annulationEclatementCommande()
            setfrmUpdated()
        End If
    End Sub
    Private Sub frmSaisieCommande_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (Not DesignMode) Then
            initFenetre()
        End If
    End Sub

    Private Sub ckTransport_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckTransport.CheckedChanged
        If ckTransport.Checked Then
            '            grpInfFactureTRP.Enabled = True
            liFactTRP.Enabled = True
        Else
            '           grpInfFactureTRP.Enabled = False
            liFactTRP.Enabled = False
        End If
    End Sub




#End Region

#Region "Méthodes à redefinir"
    Protected Overridable Function exporterWEBEDI() As Boolean

    End Function
    Protected Overridable Function AppliquerModifScmd() As Boolean
    End Function
    Protected Overridable Function annulationEclatementCommande() As Boolean
    End Function
    Protected Overridable Function eclatementCommande() As Boolean
    End Function
    Protected Overridable Function faxerSousCommande() As Boolean
    End Function
    Protected Overridable Function faxerTouteslesSousCommandes() As Boolean
    End Function
    Protected Overridable Function initTPEclatement() As Boolean
    End Function
    Protected Overridable Function selectionneSousCommande() As Boolean
    End Function
    Protected Overridable Function setcrDetailCommandeClientParameters(ByVal objReport As ReportDocument) As Boolean
    End Function
    Protected Overridable Function validerSousCommande() As Boolean

    End Function
    Protected Overridable Function validerToutesLesSousCommandes() As Boolean

    End Function
    Protected Overridable Function visualiserSousCommande() As Boolean

    End Function
    Protected Overridable Function affichecrDetailCommande() As Boolean

    End Function
    Protected Overridable Function faxerValidationCommande() As Boolean

    End Function
    Protected Overridable Function affichecrBL() As Boolean

    End Function
    Protected Overridable Function faxerBL(ByVal pNumFax As String, Optional ByVal poTiers As Tiers = Nothing) As Boolean

    End Function
    Protected Overridable Sub blToutOK()
    End Sub 'BLToutOK
    Protected Overridable Function getbGestionMvtStock() As Boolean
        Debug.Assert(False, "Fonction Non implémentée")
    End Function
    Protected Overridable Function getbGestionSousCommande() As Boolean
        Debug.Assert(False, "Fonction Non implémentée")
    End Function

    'Fonction : AnnulerLivraison
    'Description : Annule la Livraison => changement d'état , suppression des mvt de stocks à la prochaine Sauvegarde
    Protected Overridable Sub annulerLivraison()
    End Sub

    Protected Overridable Sub valideCommande()
        Dim action As vncEnums.vncActionEtatCommande
        If ckCmdValide.Checked Then
            dtDateValidation.Value = Now()
            action = vncEnums.vncActionEtatCommande.vncActionValider
        Else
            action = vncEnums.vncActionEtatCommande.vncActionAnnValider
        End If
        getCommandeCourante.changeEtat(action)
        AfficheEtat()
    End Sub
    Protected Overridable Function affichecrFactHBV() As Boolean
        Return False
    End Function

#End Region


    Private Sub cbSCMDAppliquer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSCMDAppliquer.Click
        AppliquerModifScmd()
        setfrmUpdated()
    End Sub

    Private Sub cbCalcMontantTransport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCalcMontantTransport.Click
        calculMontantTransport()
    End Sub

    Private Sub liFactTRP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles liFactTRP.LinkClicked
        afficheFenetreFactTrp(liFactTRP.Tag)
    End Sub


    Private Sub cbCalcPoidsColis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCalcPoidsColis.Click
        CalcPoidsColis()
    End Sub

    Private Sub tbQteColis_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbQteColis.TextChanged
        tbQteColisCmd.Text = tbQteColis.Text
    End Sub

    Private Sub tbPoids_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbPoids.TextChanged
        tbPoidsCmd.Text = tbPoids.Text
    End Sub

    Private Sub tbTrpFax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbTrpFax.TextChanged
        tbFaxTRP.Text = tbTrpFax.Text

    End Sub

    Private Sub cboCodeTRP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCodeTRP.SelectedIndexChanged
        If Not bAffichageEnCours() Then
            selectTransporteur()
        End If
    End Sub


    Private Sub tbCommentaireCommande_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCommentaireCommande.TextChanged
        If Not bAffichageEnCours() Then
            debAffiche()
            tbComValid.Text = tbCommentaireCommande.Text
            finAffiche()
        End If
    End Sub

    Private Sub tbComValid_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbComValid.TextChanged
        If Not bAffichageEnCours() Then
            debAffiche()
            tbCommentaireCommande.Text = tbComValid.Text
            finAffiche()
        End If

    End Sub

    Private Sub Label64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label64.Click

    End Sub

    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        If Not m_bsrcLignesCommande.Current Is Nothing Then
            modifieruneLigne()
        End If
    End Sub
    Private Sub DataGridView1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        If Not m_bsrcLignesCommande.Current Is Nothing Then
            If e.KeyCode = Keys.Delete Then
                supprimeruneLigne()
            End If
        End If
    End Sub

    Private Sub DataGridView2_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseDoubleClick
        If Not m_bsrcLignesCommande.Current Is Nothing Then
            modifieruneLigne()
        End If
    End Sub

    Private Sub m_dgvSCMD_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles m_dgvSCMD.MouseDoubleClick
        If Not m_bsrcSousCommande.Current Is Nothing Then
            visualiserSousCommande()
        End If

    End Sub

    Private Sub cbCalcPoidsColis_RL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCalcPoidsColis_RL.Click
        CalcPoidsColis()
    End Sub

    Private Sub cbCalcMontantTransport_RL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCalcMontantTransport_RL.Click
        calculMontantTransport()
    End Sub


    Private Sub tbQtePallPrep_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbQtePallPrep.Validated
        calculMontantTransport()
    End Sub

    Private Sub tbQtePallNonPrep_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbQtePallNonPrep.Validated
        calculMontantTransport()
    End Sub

    Private Sub tbPUPallPrep_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPUPallPrep.Validated
        calculMontantTransport()
    End Sub

    Private Sub tbPUPallNonPrep_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPUPallNonPrep.Validated
        calculMontantTransport()
    End Sub

    Private Sub tbCodeClient_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCodeClient.TextChanged

    End Sub

    Private Sub tbQtePallPrep_RL_Validated(sender As System.Object, e As System.EventArgs) Handles tbQtePallPrep_RL.Validated
        calculMontantTransport()
    End Sub

    Private Sub tbQtePallNonPrep_RL_Validated(sender As System.Object, e As System.EventArgs) Handles tbQtePallNonPrep_RL.Validated
        calculMontantTransport()
    End Sub

    Private Sub tbPUPallPrep_RL_Validated(sender As System.Object, e As System.EventArgs) Handles tbPUPallPrep_RL.Validated
        calculMontantTransport()
    End Sub

    Private Sub tbPUPallNonPrep_RL_Validated(sender As System.Object, e As System.EventArgs) Handles tbPUPallNonPrep_RL.Validated
        calculMontantTransport()
    End Sub

    Private Sub DataGridView1_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DataGridView1.RowPrePaint
        Dim objlg As LgCommande
        Try

            objlg = CType(m_bsrcLignesCommande(e.RowIndex), LgCommande)
            If Not objlg.bStockDispo Then
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Red
            Else
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = SystemColors.Window
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnFactHBVAfficher_Click(sender As Object, e As EventArgs) Handles btnFactHBVAfficher.Click
        sauvegardeElementCourant()
        Me.Cursor = WaitCursor
        affichecrFactHBV()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cbxOrigine_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxOrigine.SelectedIndexChanged
        'If cbxOrigine.Text = "HOBIVIN" Then
        '    SSTabCommandeClient.TabPages.Add(tpFactHbv)
        'Else
        '    SSTabCommandeClient.TabPages.RemoveByKey(tpFactHbv.Name)

        'End If
    End Sub

End Class
