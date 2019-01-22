Option Strict Off
Option Explicit On
Imports vini_DB
Friend Class frmMain
    Inherits System.Windows.Forms.Form
    Private WithEvents m_frmActive As FrmVinicom
    Public m_title As String
    Friend WithEvents MnuCol_GenFactCol As System.Windows.Forms.MenuItem
    Friend WithEvents MnuCol_GestFactcolisage As System.Windows.Forms.MenuItem
    Friend WithEvents MnuCol_LstFactColisage As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_Tarif As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_EditTarif As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSettings As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTrp_RegistreLivraison As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTrp_RegistreAppro As System.Windows.Forms.MenuItem
    Private WithEvents dbConn As Persist = New Param
    Friend WithEvents mnuCompta As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCompta_ExportCompta As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil__Param As System.Windows.Forms.MenuItem
    Friend WithEvents mnuConstantes As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCompta_SaisieReglement As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCompta_ExportReglement As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCompta_ListeReglement As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCompta_TableaudeBordFactures As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_BackupRestore As System.Windows.Forms.MenuItem
    Friend WithEvents cbToolbarBackupRestore As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuStat_StatCAClientProducteur As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat_StatCAProducteurClient As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd_LstComAProvisionner As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd_LstcomLitige As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_Lstcommandes As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_GestTransporteur As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat_CAGroupesClients As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_FactComError As System.Windows.Forms.MenuItem
    Friend WithEvents cbTBCheck As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuCompta_EditionFactures As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_LstBonAppro As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCompta_ImportRglmt As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_Libere As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_GestModeReglmt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_GestTauxTVA As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_GestTypeClient As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_GestCouleurs As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_GestRegions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_GestConditionnement As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_GestContentants As System.Windows.Forms.MenuItem
    Friend WithEvents MnuUtil_checkDatabase As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_Droits As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_importCmd As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDB_Appelation As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEclatementCommandes As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportQuadra As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMAJLivraisonCommande As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_PurgePrecommande As System.Windows.Forms.MenuItem
    Friend WithEvents mnuImportTarif As MenuItem
    Friend WithEvents mnuVerifInfosLivraisons As System.Windows.Forms.MenuItem
    'Objet créé pour afficher l'evenement Connected/Disconnected
    Public m_currentuser As aut_user


#Region "Code généré par le Concepteur Windows Form "
    Public Sub New()
        MyBase.New()
        If m_vb6FormDefInstance Is Nothing Then
            m_vb6FormDefInstance = Me
        End If
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
    'Public WithEvents dlgCommonDialog As AxMSComDlg.AxCommonDialog
    'Public WithEvents sbStatusBar As AxMSComctlLib.AxStatusBar
    'Public WithEvents imlToolbarIcons As AxMSComctlLib.AxImageList
    'Public WithEvents Toolbar1 As AxMSComctlLib.AxToolbar
    'Public WithEvents sep1 As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
    Public WithEvents mnuFileOpen As System.Windows.Forms.MenuItem
    Public WithEvents mnuFileSave As System.Windows.Forms.MenuItem
    Public WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Public WithEvents mnuFile As System.Windows.Forms.MenuItem
    Public WithEvents mnuDB_Client As System.Windows.Forms.MenuItem
    Public WithEvents mnuDB_Fournisseur As System.Windows.Forms.MenuItem
    Public WithEvents mnuDB_Produit As System.Windows.Forms.MenuItem
    Public WithEvents mnuFenetre As System.Windows.Forms.MenuItem
    Public MainMenu1 As System.Windows.Forms.MainMenu
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    Friend WithEvents cbToolBarNew As System.Windows.Forms.ToolBarButton
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents cbToolBarOpen As System.Windows.Forms.ToolBarButton
    Friend WithEvents cbToolBarSave As System.Windows.Forms.ToolBarButton
    Friend WithEvents cbToolbarDelete As System.Windows.Forms.ToolBarButton
    Public WithEvents mnuFileNew As System.Windows.Forms.MenuItem
    Public WithEvents mnuGC_CmdCltSaisie As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_RecalculStock As System.Windows.Forms.MenuItem
    Friend WithEvents cbToolBarRefresh As System.Windows.Forms.ToolBarButton
    Public WithEvents mnuFileSuppr As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileRefresh As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_BonAppro As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_EtatStock As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_MouvementArticle As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat_CumulVentesArt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat_BilanComClient As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat_HistoriqueArticle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat_StatProduitFournisseur As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd_Gest As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd_RapproFactFourn As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd__Liste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd_GeneFactCom As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd_GestFactCom As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScmd_ListeFactCom As System.Windows.Forms.MenuItem
    Public WithEvents mnuDB As System.Windows.Forms.MenuItem
    Public WithEvents mnuGC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_Users As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_AutresEtats As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolBar As System.Windows.Forms.ToolBar
    Friend WithEvents mnuUtil_PurgeMvtStock As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_PurgeCommandesClients As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_PurgeBonAppro As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_PurgeFactCom As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUtil_ControleMvtStock As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_StatQteArticleMois As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTrp As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTrp_GeneFactTRP As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTrp_GestFactTransport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTrp_LstCmdTrp As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTrp_LstFactTrp As System.Windows.Forms.MenuItem
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarDB As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarError As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarEtat As System.Windows.Forms.StatusBarPanel
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_JournalAnnoncesLivraison As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_JournalAnnoncesAppro As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_RecapAnnonces As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat_CAClientMensuel As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStat_CA1Client As System.Windows.Forms.MenuItem
    Friend WithEvents MnuTrp_lstPalettes As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_LstEntreesPlateforme As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGC_LstSortiesPlateforme As System.Windows.Forms.MenuItem
    Friend WithEvents MnuCol As System.Windows.Forms.MenuItem
    Friend WithEvents MnuCol_GenMvtIvent As System.Windows.Forms.MenuItem
    Friend WithEvents MnuCol_Recapitulatifcolisage As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuFileNew = New System.Windows.Forms.MenuItem()
        Me.mnuFileOpen = New System.Windows.Forms.MenuItem()
        Me.mnuFileSave = New System.Windows.Forms.MenuItem()
        Me.mnuFileSuppr = New System.Windows.Forms.MenuItem()
        Me.mnuFileRefresh = New System.Windows.Forms.MenuItem()
        Me.mnuFileExit = New System.Windows.Forms.MenuItem()
        Me.mnuDB = New System.Windows.Forms.MenuItem()
        Me.mnuDB_Client = New System.Windows.Forms.MenuItem()
        Me.mnuDB_Fournisseur = New System.Windows.Forms.MenuItem()
        Me.mnuDB_Produit = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.mnuDB_Appelation = New System.Windows.Forms.MenuItem()
        Me.mnuDB_Tarif = New System.Windows.Forms.MenuItem()
        Me.mnuDB_EditTarif = New System.Windows.Forms.MenuItem()
        Me.mnuImportTarif = New System.Windows.Forms.MenuItem()
        Me.MenuItem20 = New System.Windows.Forms.MenuItem()
        Me.mnuDB_GestTransporteur = New System.Windows.Forms.MenuItem()
        Me.mnuDB_GestModeReglmt = New System.Windows.Forms.MenuItem()
        Me.mnuDB_GestTauxTVA = New System.Windows.Forms.MenuItem()
        Me.mnuDB_GestTypeClient = New System.Windows.Forms.MenuItem()
        Me.mnuDB_GestCouleurs = New System.Windows.Forms.MenuItem()
        Me.mnuDB_GestRegions = New System.Windows.Forms.MenuItem()
        Me.mnuDB_GestConditionnement = New System.Windows.Forms.MenuItem()
        Me.mnuDB_GestContentants = New System.Windows.Forms.MenuItem()
        Me.mnuGC = New System.Windows.Forms.MenuItem()
        Me.mnuGC_CmdCltSaisie = New System.Windows.Forms.MenuItem()
        Me.mnuGC_BonAppro = New System.Windows.Forms.MenuItem()
        Me.mnuGC_Lstcommandes = New System.Windows.Forms.MenuItem()
        Me.mnuGC_LstBonAppro = New System.Windows.Forms.MenuItem()
        Me.mnuGC_importCmd = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.mnuVerifInfosLivraisons = New System.Windows.Forms.MenuItem()
        Me.mnuEclatementCommandes = New System.Windows.Forms.MenuItem()
        Me.mnuExportQuadra = New System.Windows.Forms.MenuItem()
        Me.mnuMAJLivraisonCommande = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.mnuGC_EtatStock = New System.Windows.Forms.MenuItem()
        Me.mnuGC_MouvementArticle = New System.Windows.Forms.MenuItem()
        Me.mnuGC_StatQteArticleMois = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.mnuGC_JournalAnnoncesLivraison = New System.Windows.Forms.MenuItem()
        Me.mnuGC_JournalAnnoncesAppro = New System.Windows.Forms.MenuItem()
        Me.mnuGC_RecapAnnonces = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.mnuGC_LstEntreesPlateforme = New System.Windows.Forms.MenuItem()
        Me.mnuGC_LstSortiesPlateforme = New System.Windows.Forms.MenuItem()
        Me.mnuStat = New System.Windows.Forms.MenuItem()
        Me.mnuStat_CumulVentesArt = New System.Windows.Forms.MenuItem()
        Me.mnuStat_HistoriqueArticle = New System.Windows.Forms.MenuItem()
        Me.mnuStat_StatProduitFournisseur = New System.Windows.Forms.MenuItem()
        Me.MenuItem18 = New System.Windows.Forms.MenuItem()
        Me.mnuStat_BilanComClient = New System.Windows.Forms.MenuItem()
        Me.mnuStat_CAGroupesClients = New System.Windows.Forms.MenuItem()
        Me.mnuStat_CA1Client = New System.Windows.Forms.MenuItem()
        Me.mnuStat_CAClientMensuel = New System.Windows.Forms.MenuItem()
        Me.MenuItem19 = New System.Windows.Forms.MenuItem()
        Me.mnuStat_StatCAClientProducteur = New System.Windows.Forms.MenuItem()
        Me.mnuStat_StatCAProducteurClient = New System.Windows.Forms.MenuItem()
        Me.mnuScmd = New System.Windows.Forms.MenuItem()
        Me.mnuScmd_Gest = New System.Windows.Forms.MenuItem()
        Me.mnuScmd_RapproFactFourn = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.mnuScmd__Liste = New System.Windows.Forms.MenuItem()
        Me.mnuScmd_LstcomLitige = New System.Windows.Forms.MenuItem()
        Me.mnuScmd_LstComAProvisionner = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.mnuScmd_GeneFactCom = New System.Windows.Forms.MenuItem()
        Me.mnuScmd_GestFactCom = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.mnuScmd_ListeFactCom = New System.Windows.Forms.MenuItem()
        Me.MnuTrp = New System.Windows.Forms.MenuItem()
        Me.MnuTrp_GeneFactTRP = New System.Windows.Forms.MenuItem()
        Me.MnuTrp_GestFactTransport = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.MnuTrp_LstCmdTrp = New System.Windows.Forms.MenuItem()
        Me.MnuTrp_LstFactTrp = New System.Windows.Forms.MenuItem()
        Me.MnuTrp_lstPalettes = New System.Windows.Forms.MenuItem()
        Me.MenuItem13 = New System.Windows.Forms.MenuItem()
        Me.MnuTrp_RegistreLivraison = New System.Windows.Forms.MenuItem()
        Me.MnuTrp_RegistreAppro = New System.Windows.Forms.MenuItem()
        Me.MnuCol = New System.Windows.Forms.MenuItem()
        Me.MnuCol_GenMvtIvent = New System.Windows.Forms.MenuItem()
        Me.MnuCol_Recapitulatifcolisage = New System.Windows.Forms.MenuItem()
        Me.MnuCol_GenFactCol = New System.Windows.Forms.MenuItem()
        Me.MnuCol_GestFactcolisage = New System.Windows.Forms.MenuItem()
        Me.MnuCol_LstFactColisage = New System.Windows.Forms.MenuItem()
        Me.mnuCompta = New System.Windows.Forms.MenuItem()
        Me.mnuCompta_SaisieReglement = New System.Windows.Forms.MenuItem()
        Me.mnuCompta_ListeReglement = New System.Windows.Forms.MenuItem()
        Me.mnuCompta_TableaudeBordFactures = New System.Windows.Forms.MenuItem()
        Me.MenuItem15 = New System.Windows.Forms.MenuItem()
        Me.mnuCompta_ExportCompta = New System.Windows.Forms.MenuItem()
        Me.mnuCompta_ExportReglement = New System.Windows.Forms.MenuItem()
        Me.mnuCompta_ImportRglmt = New System.Windows.Forms.MenuItem()
        Me.mnuCompta_EditionFactures = New System.Windows.Forms.MenuItem()
        Me.mnuFenetre = New System.Windows.Forms.MenuItem()
        Me.mnuUtil = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_PurgeMvtStock = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_PurgeCommandesClients = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_PurgeFactCom = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_PurgeBonAppro = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_PurgePrecommande = New System.Windows.Forms.MenuItem()
        Me.MenuItem22 = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_Libere = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_Users = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_Droits = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_RecalculStock = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_AutresEtats = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_ControleMvtStock = New System.Windows.Forms.MenuItem()
        Me.MenuItem21 = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_FactComError = New System.Windows.Forms.MenuItem()
        Me.MenuItem12 = New System.Windows.Forms.MenuItem()
        Me.mnuUtil__Param = New System.Windows.Forms.MenuItem()
        Me.mnuSettings = New System.Windows.Forms.MenuItem()
        Me.mnuConstantes = New System.Windows.Forms.MenuItem()
        Me.MenuItem17 = New System.Windows.Forms.MenuItem()
        Me.MenuItem16 = New System.Windows.Forms.MenuItem()
        Me.mnuUtil_BackupRestore = New System.Windows.Forms.MenuItem()
        Me.MnuUtil_checkDatabase = New System.Windows.Forms.MenuItem()
        Me.mnuToolBar = New System.Windows.Forms.ToolBar()
        Me.cbToolBarNew = New System.Windows.Forms.ToolBarButton()
        Me.cbToolBarOpen = New System.Windows.Forms.ToolBarButton()
        Me.cbToolBarSave = New System.Windows.Forms.ToolBarButton()
        Me.cbToolbarDelete = New System.Windows.Forms.ToolBarButton()
        Me.cbToolBarRefresh = New System.Windows.Forms.ToolBarButton()
        Me.cbToolbarBackupRestore = New System.Windows.Forms.ToolBarButton()
        Me.cbTBCheck = New System.Windows.Forms.ToolBarButton()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.StatusBar1 = New System.Windows.Forms.StatusBar()
        Me.StatusBarDB = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarError = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarEtat = New System.Windows.Forms.StatusBarPanel()
        CType(Me.StatusBarDB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarError, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarEtat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuDB, Me.mnuGC, Me.mnuStat, Me.mnuScmd, Me.MnuTrp, Me.MnuCol, Me.mnuCompta, Me.mnuFenetre, Me.mnuUtil})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileNew, Me.mnuFileOpen, Me.mnuFileSave, Me.mnuFileSuppr, Me.mnuFileRefresh, Me.mnuFileExit})
        Me.mnuFile.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuFile.Tag = "mnuFile"
        Me.mnuFile.Text = "&Fichier"
        '
        'mnuFileNew
        '
        Me.mnuFileNew.Index = 0
        Me.mnuFileNew.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuFileNew.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuFileNew.Tag = "mnuFileNew"
        Me.mnuFileNew.Text = "&Nouveau"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Index = 1
        Me.mnuFileOpen.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuFileOpen.Shortcut = System.Windows.Forms.Shortcut.CtrlO
        Me.mnuFileOpen.Tag = "mnuFileOpen"
        Me.mnuFileOpen.Text = "&Ouvrir"
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Index = 2
        Me.mnuFileSave.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuFileSave.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuFileSave.Tag = "mnuFileSave"
        Me.mnuFileSave.Text = "&Sauvegarder"
        '
        'mnuFileSuppr
        '
        Me.mnuFileSuppr.Index = 3
        Me.mnuFileSuppr.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuFileSuppr.Shortcut = System.Windows.Forms.Shortcut.CtrlD
        Me.mnuFileSuppr.Tag = "mnuFileSuppr"
        Me.mnuFileSuppr.Text = "S&upprimer"
        '
        'mnuFileRefresh
        '
        Me.mnuFileRefresh.Index = 4
        Me.mnuFileRefresh.Shortcut = System.Windows.Forms.Shortcut.CtrlR
        Me.mnuFileRefresh.Tag = "mnuFileRefresh"
        Me.mnuFileRefresh.Text = "&Rafraichir"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Index = 5
        Me.mnuFileExit.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuFileExit.Tag = "mnuFileExit"
        Me.mnuFileExit.Text = "&Quitter"
        '
        'mnuDB
        '
        Me.mnuDB.Index = 1
        Me.mnuDB.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuDB_Client, Me.mnuDB_Fournisseur, Me.mnuDB_Produit, Me.MenuItem11, Me.mnuDB_Appelation, Me.mnuDB_Tarif, Me.mnuDB_EditTarif, Me.mnuImportTarif, Me.MenuItem20, Me.mnuDB_GestTransporteur, Me.mnuDB_GestModeReglmt, Me.mnuDB_GestTauxTVA, Me.mnuDB_GestTypeClient, Me.mnuDB_GestCouleurs, Me.mnuDB_GestRegions, Me.mnuDB_GestConditionnement, Me.mnuDB_GestContentants})
        Me.mnuDB.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuDB.Tag = "mnuDB"
        Me.mnuDB.Text = "&Données de base"
        '
        'mnuDB_Client
        '
        Me.mnuDB_Client.Index = 0
        Me.mnuDB_Client.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuDB_Client.Shortcut = System.Windows.Forms.Shortcut.F1
        Me.mnuDB_Client.Tag = "mnuDB_Client"
        Me.mnuDB_Client.Text = "&Client"
        '
        'mnuDB_Fournisseur
        '
        Me.mnuDB_Fournisseur.Index = 1
        Me.mnuDB_Fournisseur.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuDB_Fournisseur.Shortcut = System.Windows.Forms.Shortcut.F2
        Me.mnuDB_Fournisseur.Tag = "mnuDB_Fournisseur"
        Me.mnuDB_Fournisseur.Text = "&Fournisseur"
        '
        'mnuDB_Produit
        '
        Me.mnuDB_Produit.Index = 2
        Me.mnuDB_Produit.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuDB_Produit.Shortcut = System.Windows.Forms.Shortcut.F3
        Me.mnuDB_Produit.Tag = "mnuDB_Produit"
        Me.mnuDB_Produit.Text = "&Produit"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 3
        Me.MenuItem11.Text = "-"
        '
        'mnuDB_Appelation
        '
        Me.mnuDB_Appelation.Index = 4
        Me.mnuDB_Appelation.Tag = "mnuDB_Appelation"
        Me.mnuDB_Appelation.Text = "Gestion des appelations"
        Me.mnuDB_Appelation.Visible = False
        '
        'mnuDB_Tarif
        '
        Me.mnuDB_Tarif.Index = 5
        Me.mnuDB_Tarif.Tag = "mnuDB_Tarif"
        Me.mnuDB_Tarif.Text = "Gestion des Tarifs"
        '
        'mnuDB_EditTarif
        '
        Me.mnuDB_EditTarif.Index = 6
        Me.mnuDB_EditTarif.Tag = "mnuDB_EditTarif"
        Me.mnuDB_EditTarif.Text = "Edition du Tarif"
        '
        'mnuImportTarif
        '
        Me.mnuImportTarif.Index = 7
        Me.mnuImportTarif.Text = "Import des tarifs XLS"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 8
        Me.MenuItem20.Text = "-"
        '
        'mnuDB_GestTransporteur
        '
        Me.mnuDB_GestTransporteur.Index = 9
        Me.mnuDB_GestTransporteur.Tag = "mnuDB_GestTransporteur"
        Me.mnuDB_GestTransporteur.Text = "Gestion Transporteur"
        '
        'mnuDB_GestModeReglmt
        '
        Me.mnuDB_GestModeReglmt.Index = 10
        Me.mnuDB_GestModeReglmt.Tag = "mnuDB_GestModeReglmt"
        Me.mnuDB_GestModeReglmt.Text = "Gestion des modes de réglements"
        '
        'mnuDB_GestTauxTVA
        '
        Me.mnuDB_GestTauxTVA.Index = 11
        Me.mnuDB_GestTauxTVA.Tag = "mnuDB_GestTauxTVA"
        Me.mnuDB_GestTauxTVA.Text = "Gestion des taux de TVA"
        '
        'mnuDB_GestTypeClient
        '
        Me.mnuDB_GestTypeClient.Index = 12
        Me.mnuDB_GestTypeClient.Tag = "mnuDB_GestTypeClient"
        Me.mnuDB_GestTypeClient.Text = "Gestion des types de client"
        '
        'mnuDB_GestCouleurs
        '
        Me.mnuDB_GestCouleurs.Index = 13
        Me.mnuDB_GestCouleurs.Tag = "mnuDB_GestCouleurs"
        Me.mnuDB_GestCouleurs.Text = "Gestion des couleurs"
        '
        'mnuDB_GestRegions
        '
        Me.mnuDB_GestRegions.Index = 14
        Me.mnuDB_GestRegions.Tag = "mnuDB_GestRegions"
        Me.mnuDB_GestRegions.Text = "Gestion des régions"
        '
        'mnuDB_GestConditionnement
        '
        Me.mnuDB_GestConditionnement.Index = 15
        Me.mnuDB_GestConditionnement.Tag = "mnuDB_GestConditionnement"
        Me.mnuDB_GestConditionnement.Text = "Gestion des conditionnements"
        '
        'mnuDB_GestContentants
        '
        Me.mnuDB_GestContentants.Index = 16
        Me.mnuDB_GestContentants.Tag = "mnuDB_GestContentants"
        Me.mnuDB_GestContentants.Text = "Gestion des contenants"
        '
        'mnuGC
        '
        Me.mnuGC.Index = 2
        Me.mnuGC.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuGC_CmdCltSaisie, Me.mnuGC_BonAppro, Me.mnuGC_Lstcommandes, Me.mnuGC_LstBonAppro, Me.mnuGC_importCmd, Me.MenuItem7, Me.mnuVerifInfosLivraisons, Me.mnuEclatementCommandes, Me.mnuExportQuadra, Me.mnuMAJLivraisonCommande, Me.MenuItem10, Me.mnuGC_EtatStock, Me.mnuGC_MouvementArticle, Me.mnuGC_StatQteArticleMois, Me.MenuItem5, Me.mnuGC_JournalAnnoncesLivraison, Me.mnuGC_JournalAnnoncesAppro, Me.mnuGC_RecapAnnonces, Me.MenuItem9, Me.mnuGC_LstEntreesPlateforme, Me.mnuGC_LstSortiesPlateforme})
        Me.mnuGC.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuGC.Tag = "mnuGC"
        Me.mnuGC.Text = "Gestion &Commerciale"
        '
        'mnuGC_CmdCltSaisie
        '
        Me.mnuGC_CmdCltSaisie.Index = 0
        Me.mnuGC_CmdCltSaisie.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuGC_CmdCltSaisie.Shortcut = System.Windows.Forms.Shortcut.F5
        Me.mnuGC_CmdCltSaisie.Tag = "mnuGC_CmdCltSaisie"
        Me.mnuGC_CmdCltSaisie.Text = "&Commande Client"
        '
        'mnuGC_BonAppro
        '
        Me.mnuGC_BonAppro.Index = 1
        Me.mnuGC_BonAppro.Shortcut = System.Windows.Forms.Shortcut.F6
        Me.mnuGC_BonAppro.Tag = "mnuGC_BonAppro"
        Me.mnuGC_BonAppro.Text = "&Bon Approvisionnement"
        '
        'mnuGC_Lstcommandes
        '
        Me.mnuGC_Lstcommandes.Index = 2
        Me.mnuGC_Lstcommandes.Tag = "mnuGC_Lstcommandes"
        Me.mnuGC_Lstcommandes.Text = "Liste des Commandes"
        '
        'mnuGC_LstBonAppro
        '
        Me.mnuGC_LstBonAppro.Index = 3
        Me.mnuGC_LstBonAppro.Tag = "mnuGC_LstBonAppro"
        Me.mnuGC_LstBonAppro.Text = "Liste des Bons d'approvisionnement"
        '
        'mnuGC_importCmd
        '
        Me.mnuGC_importCmd.Index = 4
        Me.mnuGC_importCmd.Tag = "mnuGC_importCmd"
        Me.mnuGC_importCmd.Text = "Importation de commandes"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 5
        Me.MenuItem7.Text = "-"
        '
        'mnuVerifInfosLivraisons
        '
        Me.mnuVerifInfosLivraisons.Index = 6
        Me.mnuVerifInfosLivraisons.Text = "Vérification des informations de livraisons"
        '
        'mnuEclatementCommandes
        '
        Me.mnuEclatementCommandes.Index = 7
        Me.mnuEclatementCommandes.Text = "Eclatement des Commandes"
        '
        'mnuExportQuadra
        '
        Me.mnuExportQuadra.Index = 8
        Me.mnuExportQuadra.Text = "ExportQuadra"
        '
        'mnuMAJLivraisonCommande
        '
        Me.mnuMAJLivraisonCommande.Index = 9
        Me.mnuMAJLivraisonCommande.Text = "Mise à jour Infos livaison"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 10
        Me.MenuItem10.Text = "-"
        '
        'mnuGC_EtatStock
        '
        Me.mnuGC_EtatStock.Index = 11
        Me.mnuGC_EtatStock.Shortcut = System.Windows.Forms.Shortcut.F7
        Me.mnuGC_EtatStock.Tag = "mnuGC_EtatStock"
        Me.mnuGC_EtatStock.Text = "&Etat du Stock"
        '
        'mnuGC_MouvementArticle
        '
        Me.mnuGC_MouvementArticle.Index = 12
        Me.mnuGC_MouvementArticle.Shortcut = System.Windows.Forms.Shortcut.F8
        Me.mnuGC_MouvementArticle.Tag = "mnuGC_MouvementArticle"
        Me.mnuGC_MouvementArticle.Text = "&Mouvement Article"
        '
        'mnuGC_StatQteArticleMois
        '
        Me.mnuGC_StatQteArticleMois.Index = 13
        Me.mnuGC_StatQteArticleMois.Shortcut = System.Windows.Forms.Shortcut.F9
        Me.mnuGC_StatQteArticleMois.Tag = "mnuGC_StatQteArticleMois"
        Me.mnuGC_StatQteArticleMois.Text = "Qte / Article / mois"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 14
        Me.MenuItem5.Text = "-"
        '
        'mnuGC_JournalAnnoncesLivraison
        '
        Me.mnuGC_JournalAnnoncesLivraison.Index = 15
        Me.mnuGC_JournalAnnoncesLivraison.Tag = "mnuGC_JournalAnnoncesLivraison"
        Me.mnuGC_JournalAnnoncesLivraison.Text = "Journal d'annonces LIVRAISON"
        '
        'mnuGC_JournalAnnoncesAppro
        '
        Me.mnuGC_JournalAnnoncesAppro.Index = 16
        Me.mnuGC_JournalAnnoncesAppro.Tag = "mnuGC_JournalAnnoncesAppro"
        Me.mnuGC_JournalAnnoncesAppro.Text = "Journal d'annonces APPROVISIONNEMENT"
        '
        'mnuGC_RecapAnnonces
        '
        Me.mnuGC_RecapAnnonces.Index = 17
        Me.mnuGC_RecapAnnonces.Tag = "mnuGC_RecapAnnonces"
        Me.mnuGC_RecapAnnonces.Text = "Recapitulatif annonces Transporteur"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 18
        Me.MenuItem9.Text = "-"
        '
        'mnuGC_LstEntreesPlateforme
        '
        Me.mnuGC_LstEntreesPlateforme.Index = 19
        Me.mnuGC_LstEntreesPlateforme.Tag = "mnuGC_LstEntreesPlateforme"
        Me.mnuGC_LstEntreesPlateforme.Text = "Liste des entrées plateforme"
        '
        'mnuGC_LstSortiesPlateforme
        '
        Me.mnuGC_LstSortiesPlateforme.Index = 20
        Me.mnuGC_LstSortiesPlateforme.Tag = "mnuGC_LstSortiesPlateforme"
        Me.mnuGC_LstSortiesPlateforme.Text = "Liste des sorties plateforme"
        '
        'mnuStat
        '
        Me.mnuStat.Index = 3
        Me.mnuStat.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuStat_CumulVentesArt, Me.mnuStat_HistoriqueArticle, Me.mnuStat_StatProduitFournisseur, Me.MenuItem18, Me.mnuStat_BilanComClient, Me.mnuStat_CAGroupesClients, Me.mnuStat_CA1Client, Me.mnuStat_CAClientMensuel, Me.MenuItem19, Me.mnuStat_StatCAClientProducteur, Me.mnuStat_StatCAProducteurClient})
        Me.mnuStat.Tag = "mnuStat"
        Me.mnuStat.Text = "Statistiques Commerciales"
        '
        'mnuStat_CumulVentesArt
        '
        Me.mnuStat_CumulVentesArt.Index = 0
        Me.mnuStat_CumulVentesArt.Tag = "mnuStat_CumulVentesArt"
        Me.mnuStat_CumulVentesArt.Text = "Cumul des ventes par articles"
        '
        'mnuStat_HistoriqueArticle
        '
        Me.mnuStat_HistoriqueArticle.Index = 1
        Me.mnuStat_HistoriqueArticle.Tag = "mnuStat_HistoriqueArticle"
        Me.mnuStat_HistoriqueArticle.Text = "Historique Article"
        '
        'mnuStat_StatProduitFournisseur
        '
        Me.mnuStat_StatProduitFournisseur.Index = 2
        Me.mnuStat_StatProduitFournisseur.Tag = "mnuStat_StatProduitFournisseur"
        Me.mnuStat_StatProduitFournisseur.Text = "Statistiques Produit/Fournisseur"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 3
        Me.MenuItem18.Text = "-"
        '
        'mnuStat_BilanComClient
        '
        Me.mnuStat_BilanComClient.Index = 4
        Me.mnuStat_BilanComClient.Tag = "mnuStat_BilanComClient"
        Me.mnuStat_BilanComClient.Text = "Bilan Commercial Client"
        '
        'mnuStat_CAGroupesClients
        '
        Me.mnuStat_CAGroupesClients.Index = 5
        Me.mnuStat_CAGroupesClients.Tag = "mnuStat_CAGroupesClients"
        Me.mnuStat_CAGroupesClients.Text = "CA Groupes clients"
        '
        'mnuStat_CA1Client
        '
        Me.mnuStat_CA1Client.Index = 6
        Me.mnuStat_CA1Client.Tag = "mnuStat_CA1Client"
        Me.mnuStat_CA1Client.Text = "CA 1 Client"
        '
        'mnuStat_CAClientMensuel
        '
        Me.mnuStat_CAClientMensuel.Index = 7
        Me.mnuStat_CAClientMensuel.Tag = "mnuStat_CAClientMensuel"
        Me.mnuStat_CAClientMensuel.Text = "CA Mensuel Client "
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 8
        Me.MenuItem19.Text = "-"
        '
        'mnuStat_StatCAClientProducteur
        '
        Me.mnuStat_StatCAClientProducteur.Index = 9
        Me.mnuStat_StatCAClientProducteur.Tag = "mnuStat_StatCAClientProducteur"
        Me.mnuStat_StatCAClientProducteur.Text = "CA Client /Producteur"
        '
        'mnuStat_StatCAProducteurClient
        '
        Me.mnuStat_StatCAProducteurClient.Index = 10
        Me.mnuStat_StatCAProducteurClient.Tag = "mnuStat_StatCAProducteurClient"
        Me.mnuStat_StatCAProducteurClient.Text = "CA Producteur /Client"
        '
        'mnuScmd
        '
        Me.mnuScmd.Index = 4
        Me.mnuScmd.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuScmd_Gest, Me.mnuScmd_RapproFactFourn, Me.MenuItem6, Me.mnuScmd__Liste, Me.mnuScmd_LstcomLitige, Me.mnuScmd_LstComAProvisionner, Me.MenuItem8, Me.mnuScmd_GeneFactCom, Me.mnuScmd_GestFactCom, Me.MenuItem3, Me.mnuScmd_ListeFactCom})
        Me.mnuScmd.Tag = "mnuScmd"
        Me.mnuScmd.Text = "SousCommandes"
        '
        'mnuScmd_Gest
        '
        Me.mnuScmd_Gest.Index = 0
        Me.mnuScmd_Gest.Shortcut = System.Windows.Forms.Shortcut.ShiftF1
        Me.mnuScmd_Gest.Tag = "mnuScmd_Gest"
        Me.mnuScmd_Gest.Text = "Gestion des SousCommandes"
        '
        'mnuScmd_RapproFactFourn
        '
        Me.mnuScmd_RapproFactFourn.Index = 1
        Me.mnuScmd_RapproFactFourn.Shortcut = System.Windows.Forms.Shortcut.ShiftF2
        Me.mnuScmd_RapproFactFourn.Tag = "mnuScmd_RapproFactFourn"
        Me.mnuScmd_RapproFactFourn.Text = "Rapprochement Factures Fournisseurs"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 2
        Me.MenuItem6.Text = "-"
        '
        'mnuScmd__Liste
        '
        Me.mnuScmd__Liste.Index = 3
        Me.mnuScmd__Liste.Tag = "mnuScmd__Liste"
        Me.mnuScmd__Liste.Text = "Liste des Sous-Commandes"
        '
        'mnuScmd_LstcomLitige
        '
        Me.mnuScmd_LstcomLitige.Index = 4
        Me.mnuScmd_LstcomLitige.Tag = "mnuScmd_LstcomLitige"
        Me.mnuScmd_LstcomLitige.Text = "Liste des Sous-Commandes en litige"
        '
        'mnuScmd_LstComAProvisionner
        '
        Me.mnuScmd_LstComAProvisionner.Index = 5
        Me.mnuScmd_LstComAProvisionner.Tag = "mnuScmd_LstComAProvisionner"
        Me.mnuScmd_LstComAProvisionner.Text = "Liste des commissions à provisionner"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 6
        Me.MenuItem8.Text = "-"
        '
        'mnuScmd_GeneFactCom
        '
        Me.mnuScmd_GeneFactCom.Index = 7
        Me.mnuScmd_GeneFactCom.Shortcut = System.Windows.Forms.Shortcut.ShiftF4
        Me.mnuScmd_GeneFactCom.Tag = "mnuScmd_GeneFactCom"
        Me.mnuScmd_GeneFactCom.Text = "Génération des factures de commission"
        '
        'mnuScmd_GestFactCom
        '
        Me.mnuScmd_GestFactCom.Index = 8
        Me.mnuScmd_GestFactCom.Shortcut = System.Windows.Forms.Shortcut.ShiftF5
        Me.mnuScmd_GestFactCom.Tag = "mnuScmd_GestFactCom"
        Me.mnuScmd_GestFactCom.Text = "Gestion des factures de commission"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 9
        Me.MenuItem3.Text = "-"
        '
        'mnuScmd_ListeFactCom
        '
        Me.mnuScmd_ListeFactCom.Index = 10
        Me.mnuScmd_ListeFactCom.Tag = "mnuScmd_ListeFactCom"
        Me.mnuScmd_ListeFactCom.Text = "Liste des factures de commission"
        '
        'MnuTrp
        '
        Me.MnuTrp.Index = 5
        Me.MnuTrp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuTrp_GeneFactTRP, Me.MnuTrp_GestFactTransport, Me.MenuItem4, Me.MnuTrp_LstCmdTrp, Me.MnuTrp_LstFactTrp, Me.MnuTrp_lstPalettes, Me.MenuItem13, Me.MnuTrp_RegistreLivraison, Me.MnuTrp_RegistreAppro})
        Me.MnuTrp.Tag = "MnuTrp"
        Me.MnuTrp.Text = "Transport"
        '
        'MnuTrp_GeneFactTRP
        '
        Me.MnuTrp_GeneFactTRP.Index = 0
        Me.MnuTrp_GeneFactTRP.Shortcut = System.Windows.Forms.Shortcut.AltF1
        Me.MnuTrp_GeneFactTRP.Tag = "MnuTrp_GeneFactTRP"
        Me.MnuTrp_GeneFactTRP.Text = "Génération Factures"
        '
        'MnuTrp_GestFactTransport
        '
        Me.MnuTrp_GestFactTransport.Index = 1
        Me.MnuTrp_GestFactTransport.Shortcut = System.Windows.Forms.Shortcut.AltF2
        Me.MnuTrp_GestFactTransport.Tag = "MnuTrp_GestFactTransport"
        Me.MnuTrp_GestFactTransport.Text = "Gestion des factures de transport"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "-"
        '
        'MnuTrp_LstCmdTrp
        '
        Me.MnuTrp_LstCmdTrp.Index = 3
        Me.MnuTrp_LstCmdTrp.Shortcut = System.Windows.Forms.Shortcut.AltF4
        Me.MnuTrp_LstCmdTrp.Tag = "MnuTrp_LstCmdTrp"
        Me.MnuTrp_LstCmdTrp.Text = "Liste des Commandes en attente"
        '
        'MnuTrp_LstFactTrp
        '
        Me.MnuTrp_LstFactTrp.Index = 4
        Me.MnuTrp_LstFactTrp.Shortcut = System.Windows.Forms.Shortcut.AltF5
        Me.MnuTrp_LstFactTrp.Tag = "MnuTrp_LstFactTrp"
        Me.MnuTrp_LstFactTrp.Text = "Liste des factures de transport"
        '
        'MnuTrp_lstPalettes
        '
        Me.MnuTrp_lstPalettes.Index = 5
        Me.MnuTrp_lstPalettes.Tag = "MnuTrp_lstPalettes"
        Me.MnuTrp_lstPalettes.Text = "Liste des palettes transportées"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 6
        Me.MenuItem13.Text = "-"
        '
        'MnuTrp_RegistreLivraison
        '
        Me.MnuTrp_RegistreLivraison.Index = 7
        Me.MnuTrp_RegistreLivraison.Tag = "MnuTrp_RegistreLivraison"
        Me.MnuTrp_RegistreLivraison.Text = "Registre Livraison"
        '
        'MnuTrp_RegistreAppro
        '
        Me.MnuTrp_RegistreAppro.Index = 8
        Me.MnuTrp_RegistreAppro.Tag = "MnuTrp_RegistreAppro"
        Me.MnuTrp_RegistreAppro.Text = "Registre d'approvisionnement"
        '
        'MnuCol
        '
        Me.MnuCol.Index = 6
        Me.MnuCol.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuCol_GenMvtIvent, Me.MnuCol_Recapitulatifcolisage, Me.MnuCol_GenFactCol, Me.MnuCol_GestFactcolisage, Me.MnuCol_LstFactColisage})
        Me.MnuCol.Tag = "MnuCol"
        Me.MnuCol.Text = "Colisage"
        '
        'MnuCol_GenMvtIvent
        '
        Me.MnuCol_GenMvtIvent.Index = 0
        Me.MnuCol_GenMvtIvent.Tag = "MnuCol_GenMvtIvent"
        Me.MnuCol_GenMvtIvent.Text = "Génération des Mvts d'inventaire"
        '
        'MnuCol_Recapitulatifcolisage
        '
        Me.MnuCol_Recapitulatifcolisage.Index = 1
        Me.MnuCol_Recapitulatifcolisage.Tag = "MnuCol_Recapitulatifcolisage"
        Me.MnuCol_Recapitulatifcolisage.Text = "Récapitulatif Colisage"
        '
        'MnuCol_GenFactCol
        '
        Me.MnuCol_GenFactCol.Index = 2
        Me.MnuCol_GenFactCol.Tag = "MnuCol_GenFactCol"
        Me.MnuCol_GenFactCol.Text = "Génération des factures de colisage"
        '
        'MnuCol_GestFactcolisage
        '
        Me.MnuCol_GestFactcolisage.Index = 3
        Me.MnuCol_GestFactcolisage.Tag = "MnuCol_GestFactcolisage"
        Me.MnuCol_GestFactcolisage.Text = "Gestion des factures de colisage"
        '
        'MnuCol_LstFactColisage
        '
        Me.MnuCol_LstFactColisage.Index = 4
        Me.MnuCol_LstFactColisage.Tag = "MnuCol_LstFactColisage"
        Me.MnuCol_LstFactColisage.Text = "Liste des factures de colisage"
        '
        'mnuCompta
        '
        Me.mnuCompta.Index = 7
        Me.mnuCompta.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCompta_SaisieReglement, Me.mnuCompta_ListeReglement, Me.mnuCompta_TableaudeBordFactures, Me.MenuItem15, Me.mnuCompta_ExportCompta, Me.mnuCompta_ExportReglement, Me.mnuCompta_ImportRglmt, Me.mnuCompta_EditionFactures})
        Me.mnuCompta.Tag = "mnuCompta"
        Me.mnuCompta.Text = "Comptabilité"
        '
        'mnuCompta_SaisieReglement
        '
        Me.mnuCompta_SaisieReglement.Index = 0
        Me.mnuCompta_SaisieReglement.Tag = "mnuCompta_SaisieReglement"
        Me.mnuCompta_SaisieReglement.Text = "Saisie des réglements"
        '
        'mnuCompta_ListeReglement
        '
        Me.mnuCompta_ListeReglement.Index = 1
        Me.mnuCompta_ListeReglement.Tag = "mnuCompta_ListeReglement"
        Me.mnuCompta_ListeReglement.Text = "Liste des réglements"
        '
        'mnuCompta_TableaudeBordFactures
        '
        Me.mnuCompta_TableaudeBordFactures.Index = 2
        Me.mnuCompta_TableaudeBordFactures.Tag = "mnuCompta_TableaudeBordFactures"
        Me.mnuCompta_TableaudeBordFactures.Text = "Tableau de bord Facture"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 3
        Me.MenuItem15.Text = "-"
        '
        'mnuCompta_ExportCompta
        '
        Me.mnuCompta_ExportCompta.Index = 4
        Me.mnuCompta_ExportCompta.Tag = "mnuCompta_ExportCompta"
        Me.mnuCompta_ExportCompta.Text = "Export factures vers comptabilité"
        '
        'mnuCompta_ExportReglement
        '
        Me.mnuCompta_ExportReglement.Index = 5
        Me.mnuCompta_ExportReglement.Tag = "mnuCompta_ExportReglement"
        Me.mnuCompta_ExportReglement.Text = "Export réglements vers comptabilité"
        '
        'mnuCompta_ImportRglmt
        '
        Me.mnuCompta_ImportRglmt.Index = 6
        Me.mnuCompta_ImportRglmt.Tag = "mnuCompta_ImportRglmt"
        Me.mnuCompta_ImportRglmt.Text = "Import des réglements depuis la comptabilité"
        Me.mnuCompta_ImportRglmt.Visible = False
        '
        'mnuCompta_EditionFactures
        '
        Me.mnuCompta_EditionFactures.Index = 7
        Me.mnuCompta_EditionFactures.Tag = "mnuCompta_EditionFactures"
        Me.mnuCompta_EditionFactures.Text = "Edition des factures"
        '
        'mnuFenetre
        '
        Me.mnuFenetre.Index = 8
        Me.mnuFenetre.MdiList = True
        Me.mnuFenetre.MergeType = System.Windows.Forms.MenuMerge.Remove
        Me.mnuFenetre.Text = "Fe&nêtre"
        '
        'mnuUtil
        '
        Me.mnuUtil.Index = 9
        Me.mnuUtil.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuUtil_PurgeMvtStock, Me.mnuUtil_PurgeCommandesClients, Me.mnuUtil_PurgeFactCom, Me.mnuUtil_PurgeBonAppro, Me.mnuUtil_PurgePrecommande, Me.MenuItem22, Me.mnuUtil_Libere, Me.mnuUtil_Users, Me.mnuUtil_Droits, Me.MenuItem1, Me.mnuUtil_RecalculStock, Me.MenuItem2, Me.mnuUtil_AutresEtats, Me.mnuUtil_ControleMvtStock, Me.MenuItem21, Me.mnuUtil_FactComError, Me.MenuItem12, Me.mnuUtil__Param, Me.MenuItem16, Me.mnuUtil_BackupRestore, Me.MnuUtil_checkDatabase})
        Me.mnuUtil.Tag = "mnuUtil"
        Me.mnuUtil.Text = "Utilitaires"
        '
        'mnuUtil_PurgeMvtStock
        '
        Me.mnuUtil_PurgeMvtStock.Index = 0
        Me.mnuUtil_PurgeMvtStock.Tag = "mnuUtil_PurgeMvtStock"
        Me.mnuUtil_PurgeMvtStock.Text = "Purge des Mvts de Stocks"
        '
        'mnuUtil_PurgeCommandesClients
        '
        Me.mnuUtil_PurgeCommandesClients.Index = 1
        Me.mnuUtil_PurgeCommandesClients.Tag = "mnuUtil_PurgeCommandesClients"
        Me.mnuUtil_PurgeCommandesClients.Text = "Purge des commandes Clients"
        '
        'mnuUtil_PurgeFactCom
        '
        Me.mnuUtil_PurgeFactCom.Index = 2
        Me.mnuUtil_PurgeFactCom.Tag = "mnuUtil_PurgeFactCom"
        Me.mnuUtil_PurgeFactCom.Text = "Purge des Factures de commissions"
        '
        'mnuUtil_PurgeBonAppro
        '
        Me.mnuUtil_PurgeBonAppro.Index = 3
        Me.mnuUtil_PurgeBonAppro.Tag = "mnuUtil_PurgeBonAppro"
        Me.mnuUtil_PurgeBonAppro.Text = "Purge des Bons D'appro"
        '
        'mnuUtil_PurgePrecommande
        '
        Me.mnuUtil_PurgePrecommande.Index = 4
        Me.mnuUtil_PurgePrecommande.Tag = "mnuUtil_PurgePrecommande"
        Me.mnuUtil_PurgePrecommande.Text = "Purge des précommandes"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 5
        Me.MenuItem22.Text = "-"
        '
        'mnuUtil_Libere
        '
        Me.mnuUtil_Libere.Index = 6
        Me.mnuUtil_Libere.Tag = "mnuUtil_Libere"
        Me.mnuUtil_Libere.Text = "Libération des éléments bloqués"
        '
        'mnuUtil_Users
        '
        Me.mnuUtil_Users.Index = 7
        Me.mnuUtil_Users.Tag = "mnuUtil_Users"
        Me.mnuUtil_Users.Text = "Code utilisateur"
        '
        'mnuUtil_Droits
        '
        Me.mnuUtil_Droits.Index = 8
        Me.mnuUtil_Droits.Tag = "mnuUtil_Droits"
        Me.mnuUtil_Droits.Text = "Gestion des droits"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 9
        Me.MenuItem1.Text = "-"
        '
        'mnuUtil_RecalculStock
        '
        Me.mnuUtil_RecalculStock.Index = 10
        Me.mnuUtil_RecalculStock.Tag = "mnuUtil_RecalculStock"
        Me.mnuUtil_RecalculStock.Text = "RecalculStock"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 11
        Me.MenuItem2.Text = "-"
        '
        'mnuUtil_AutresEtats
        '
        Me.mnuUtil_AutresEtats.Index = 12
        Me.mnuUtil_AutresEtats.Shortcut = System.Windows.Forms.Shortcut.F12
        Me.mnuUtil_AutresEtats.Tag = "mnuUtil_AutresEtats"
        Me.mnuUtil_AutresEtats.Text = "Autres etats"
        '
        'mnuUtil_ControleMvtStock
        '
        Me.mnuUtil_ControleMvtStock.Index = 13
        Me.mnuUtil_ControleMvtStock.Tag = "mnuUtil_ControleMvtStock"
        Me.mnuUtil_ControleMvtStock.Text = "Controle des mouvements de stocks"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 14
        Me.MenuItem21.Text = "-"
        '
        'mnuUtil_FactComError
        '
        Me.mnuUtil_FactComError.Index = 15
        Me.mnuUtil_FactComError.Tag = "mnuUtil_FactComError"
        Me.mnuUtil_FactComError.Text = "Facture de commission en erreur"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 16
        Me.MenuItem12.Text = "-"
        '
        'mnuUtil__Param
        '
        Me.mnuUtil__Param.Index = 17
        Me.mnuUtil__Param.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSettings, Me.mnuConstantes, Me.MenuItem17})
        Me.mnuUtil__Param.Tag = "mnuUtil__Param"
        Me.mnuUtil__Param.Text = "Paramètres"
        '
        'mnuSettings
        '
        Me.mnuSettings.Index = 0
        Me.mnuSettings.Text = "Settings"
        '
        'mnuConstantes
        '
        Me.mnuConstantes.Index = 1
        Me.mnuConstantes.Text = "Constantes"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 2
        Me.MenuItem17.Text = "-"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 18
        Me.MenuItem16.Text = "-"
        '
        'mnuUtil_BackupRestore
        '
        Me.mnuUtil_BackupRestore.Index = 19
        Me.mnuUtil_BackupRestore.Tag = "mnuUtil_BackupRestore"
        Me.mnuUtil_BackupRestore.Text = "Sauvegarde/Restauration "
        '
        'MnuUtil_checkDatabase
        '
        Me.MnuUtil_checkDatabase.Index = 20
        Me.MnuUtil_checkDatabase.Tag = "MnuUtil_checkDatabase"
        Me.MnuUtil_checkDatabase.Text = "Check Database"
        '
        'mnuToolBar
        '
        Me.mnuToolBar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.cbToolBarNew, Me.cbToolBarOpen, Me.cbToolBarSave, Me.cbToolbarDelete, Me.cbToolBarRefresh, Me.cbToolbarBackupRestore, Me.cbTBCheck})
        Me.mnuToolBar.DropDownArrows = True
        Me.mnuToolBar.ImageList = Me.ImageList1
        Me.mnuToolBar.Location = New System.Drawing.Point(0, 0)
        Me.mnuToolBar.Name = "mnuToolBar"
        Me.mnuToolBar.ShowToolTips = True
        Me.mnuToolBar.Size = New System.Drawing.Size(890, 28)
        Me.mnuToolBar.TabIndex = 1
        Me.mnuToolBar.TextAlign = System.Windows.Forms.ToolBarTextAlign.Right
        '
        'cbToolBarNew
        '
        Me.cbToolBarNew.Enabled = False
        Me.cbToolBarNew.ImageIndex = 0
        Me.cbToolBarNew.Name = "cbToolBarNew"
        Me.cbToolBarNew.Tag = "NEW"
        Me.cbToolBarNew.ToolTipText = "Nouveau élément"
        '
        'cbToolBarOpen
        '
        Me.cbToolBarOpen.Enabled = False
        Me.cbToolBarOpen.ImageIndex = 1
        Me.cbToolBarOpen.Name = "cbToolBarOpen"
        Me.cbToolBarOpen.Tag = "OPEN"
        '
        'cbToolBarSave
        '
        Me.cbToolBarSave.Enabled = False
        Me.cbToolBarSave.ImageIndex = 2
        Me.cbToolBarSave.Name = "cbToolBarSave"
        Me.cbToolBarSave.Tag = "SAVE"
        '
        'cbToolbarDelete
        '
        Me.cbToolbarDelete.Enabled = False
        Me.cbToolbarDelete.ImageIndex = 3
        Me.cbToolbarDelete.Name = "cbToolbarDelete"
        Me.cbToolbarDelete.Tag = "DELETE"
        '
        'cbToolBarRefresh
        '
        Me.cbToolBarRefresh.Enabled = False
        Me.cbToolBarRefresh.ImageIndex = 4
        Me.cbToolBarRefresh.Name = "cbToolBarRefresh"
        Me.cbToolBarRefresh.Tag = "REFRESH"
        Me.cbToolBarRefresh.ToolTipText = "Refresh"
        '
        'cbToolbarBackupRestore
        '
        Me.cbToolbarBackupRestore.ImageIndex = 5
        Me.cbToolbarBackupRestore.Name = "cbToolbarBackupRestore"
        Me.cbToolbarBackupRestore.Tag = "BACKUP"
        '
        'cbTBCheck
        '
        Me.cbTBCheck.ImageKey = "criticalReport.ico"
        Me.cbTBCheck.Name = "cbTBCheck"
        Me.cbTBCheck.Tag = "CHECK"
        Me.cbTBCheck.ToolTipText = "Check"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "NewDocumentHS.png")
        Me.ImageList1.Images.SetKeyName(1, "openHS.png")
        Me.ImageList1.Images.SetKeyName(2, "saveHS.png")
        Me.ImageList1.Images.SetKeyName(3, "DeleteHS.png")
        Me.ImageList1.Images.SetKeyName(4, "RefreshDocViewHS.png")
        Me.ImageList1.Images.SetKeyName(5, "BackItUp_2.ico")
        Me.ImageList1.Images.SetKeyName(6, "criticalReport.ico")
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 508)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarDB, Me.StatusBarError, Me.StatusBarEtat})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(890, 22)
        Me.StatusBar1.TabIndex = 3
        '
        'StatusBarDB
        '
        Me.StatusBarDB.Name = "StatusBarDB"
        '
        'StatusBarError
        '
        Me.StatusBarError.Name = "StatusBarError"
        Me.StatusBarError.Width = 500
        '
        'StatusBarEtat
        '
        Me.StatusBarEtat.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.StatusBarEtat.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.StatusBarEtat.Name = "StatusBarEtat"
        Me.StatusBarEtat.Width = 273
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ClientSize = New System.Drawing.Size(890, 530)
        Me.Controls.Add(Me.StatusBar1)
        Me.Controls.Add(Me.mnuToolBar)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.IsMdiContainer = True
        Me.Location = New System.Drawing.Point(11, 57)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "VINICOM"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.StatusBarDB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarError, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarEtat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Prise en charge de la mise à niveau "
    Private Shared m_vb6FormDefInstance As frmMain
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmMain
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmMain
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal Value As frmMain)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    Public Property CurrentUser() As aut_user
        Get
            Return m_currentuser
        End Get
        Set(ByVal value As aut_user)
            m_currentuser = value
        End Set
    End Property

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim ota As dsVinicomTableAdapters.LOCKTableAdapter
        ota = New dsVinicomTableAdapters.LOCKTableAdapter()
        ota.Connection = Persist.oleDBConnection
        ota.DeletefromName(System.Environment.MachineName + CurrentUser.code)

    End Sub

    Private Sub frmMain_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave
    End Sub


    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Me.Text = Param.getConstante("CST_SOC_NOMSOC")
        'Mise à jour du menu en fonction du current user
        setMenuItems("", MainMenu1.MenuItems)
        SetToolBarEnabled()
        'Connection /Deconnection pour mettre à jour la barre d'état
        Persist.shared_connect()
        Persist.shared_disconnect()

    End Sub

    Private Sub setMenuItems(ByVal str As String, ByVal menuitems As Menu.MenuItemCollection)
        Try

            Debug.Write("setMenuItems: " & str)
            Dim oMnu As MenuItem
            For Each oMnu In menuitems
                If (Not oMnu.Tag Is Nothing) Then
                    oMnu.Enabled = m_currentuser.AccesAuthorise(oMnu.Tag)
                End If
                For Each oMnu1 As MenuItem In oMnu.MenuItems
                    If (Not oMnu1.Tag Is Nothing) Then
                        oMnu1.Enabled = m_currentuser.AccesAuthorise(oMnu1.Tag)
                    End If
                Next

            Next
        Catch ex As Exception
            Trace.WriteLine(ex.Message)
        End Try


    End Sub

    Private Sub frmMain_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
        If Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
            'SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, "Settings", "MainLeft", CStr(VB6.PixelsToTwipsX(Me.Left)))
            'SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, "Settings", "MainTop", CStr(VB6.PixelsToTwipsY(Me.Top)))
            'SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, "Settings", "MainWidth", CStr(VB6.PixelsToTwipsX(Me.Width)))
            'SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, "Settings", "MainHeight", CStr(VB6.PixelsToTwipsY(Me.Height)))
        End If
    End Sub


    Public Sub mnuDBClient_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDB_Client.Popup
        mnuDBClient_Click(eventSender, eventArgs)
    End Sub
    Public Sub mnuDBClient_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDB_Client.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmClient
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub
    Private Sub mnuWindowArrangeIcons_Click()
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.ArrangeIcons)
    End Sub

    Private Sub mnuWindowTileVertical_Click()
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileVertical)
    End Sub

    Private Sub mnuWindowTileHorizontal_Click()
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileHorizontal)
    End Sub

    Private Sub mnuWindowCascade_Click()
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.Cascade)
    End Sub
    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        Me.Close()
    End Sub

    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles mnuToolBar.ButtonClick
        If e.Button.Tag = "BACKUP" Then
            backupRestore()
        End If
        If e.Button.Tag = "CHECK" Then
            checkDataBase()
        End If

        If Me.MdiChildren.Length > 0 Then
            Select Case UCase(e.Button.Tag)
                Case "NEW"
                    m_frmActive.frmvncNew()
                Case "OPEN"
                    m_frmActive.frmvncLoad()
                Case "SAVE"
                    m_frmActive.frmvncSave()
                Case "DELETE"
                    m_frmActive.frmvncDel()
                Case "REFRESH"
                    m_frmActive.frmvncRefresh()
            End Select
            SetToolBarEnabled()

        End If
    End Sub
    Public Sub setFrmActive(ByRef ofrm As FrmVinicom)
        m_frmActive = ofrm
        SetToolBarEnabled()
    End Sub
    Private Sub SetToolBarEnabled()
        'Les boutons de la toolbar correspondent au 5 premier element du Premier Menu
        'He oui c'est comme chez McDonald !!!
        Me.mnuToolBar.Buttons(5).Enabled = m_currentuser.AccesAuthorise(mnuUtil_BackupRestore.Tag)
        Me.mnuToolBar.Buttons(6).Enabled = m_currentuser.AccesAuthorise(MnuUtil_checkDatabase.Tag)
        If mnuToolBar.Buttons.Count > 0 Then
            If Not m_frmActive Is Nothing Then
                Me.mnuToolBar.Buttons(0).Enabled = m_frmActive.m_ToolBarNewEnabled And m_currentuser.AccesAuthorise(mnuFileNew.Tag)
                Me.mnuToolBar.Buttons(1).Enabled = m_frmActive.m_ToolBarLoadEnabled And m_currentuser.AccesAuthorise(mnuFileOpen.Tag)
                Me.mnuToolBar.Buttons(2).Enabled = m_frmActive.m_ToolBarSaveEnabled And m_currentuser.AccesAuthorise(mnuFileSave.Tag)
                Me.mnuToolBar.Buttons(3).Enabled = m_frmActive.m_ToolBarDelEnabled And m_currentuser.AccesAuthorise(mnuFileSuppr.Tag)
                Me.mnuToolBar.Buttons(4).Enabled = m_frmActive.m_ToolBarRefreshEnabled And m_currentuser.AccesAuthorise(mnuFileRefresh.Tag)
            Else
                Me.mnuToolBar.Buttons(0).Enabled = False
                Me.mnuToolBar.Buttons(1).Enabled = False
                Me.mnuToolBar.Buttons(2).Enabled = False
                Me.mnuToolBar.Buttons(3).Enabled = False
                Me.mnuToolBar.Buttons(4).Enabled = False
            End If

            mnuFileNew.Enabled = Me.mnuToolBar.Buttons(0).Enabled
            mnuFileOpen.Enabled = Me.mnuToolBar.Buttons(1).Enabled
            mnuFileSave.Enabled = Me.mnuToolBar.Buttons(2).Enabled
            mnuFileSuppr.Enabled = Me.mnuToolBar.Buttons(3).Enabled
            mnuFileRefresh.Enabled = Me.mnuToolBar.Buttons(4).Enabled

        End If


    End Sub
    Private Sub mnuFileNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileNew.Click
        If Me.MdiChildren.Length > 0 Then
            m_frmActive.frmvncNew()
            SetToolBarEnabled()
        End If
    End Sub

    Private Sub mnuFileOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileOpen.Click
        If Me.MdiChildren.Length > 0 Then
            m_frmActive.frmvncLoad()
            SetToolBarEnabled()
        End If

    End Sub

    Private Sub mnuFileSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
        If Me.MdiChildren.Length > 0 Then
            m_frmActive.frmvncSave()
            SetToolBarEnabled()
        End If

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_AutresEtats.Click
        Dim ofrm As frmStat
        ofrm = New frmStat
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuCmdCltSaisie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_CmdCltSaisie.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmCommandeClient
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuDBFournisseur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_Fournisseur.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmFournisseurTab
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuDBProduit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_Produit.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmProduit
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuRecalculStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_RecalculStock.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmRecalculStock
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuFileSuppr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSuppr.Click
        If Me.MdiChildren.Length > 0 Then
            m_frmActive.frmvncDel()
            SetToolBarEnabled()
        End If

    End Sub

    Private Sub mnuFileRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileRefresh.Click
        If Me.MdiChildren.Length > 0 Then
            m_frmActive.frmvncRefresh()
            SetToolBarEnabled()
        End If

    End Sub

    Private Sub mnuBonAppro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_BonAppro.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmSaisieCommandeBonAppro
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuEtatStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_EtatStock.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmEtatStock
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuMouvementArticle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_MouvementArticle.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmMouvementArticle
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_CumulVentesArt.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmCumulVentesArticles
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuBilanComClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_BilanComClient.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmBilanClient
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuHistoriqueArticle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_HistoriqueArticle.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmHistoriqueArticles
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuCAGroupeClients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_CAClientMensuel.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmStatListeCAMensuelClient

        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuStatFournisseur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_StatProduitFournisseur.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmStatProduitFournisseur

        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub


    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_Users.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmuser
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuGestionScmd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScmd_Gest.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestionSCMD
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuRapproFactFourn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScmd_RapproFactFourn.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmRapproFactFourn
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub


    Private Sub mnucrLstSousCommandes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScmd__Liste.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeSousCommandes
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuGeneFactCom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScmd_GeneFactCom.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGeneFactCom
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub MnuGestFactCom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScmd_GestFactCom.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestFactCom
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub


    Private Sub mnuListeFactCom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScmd_ListeFactCom.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeFactCom
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuTablBordFourn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ofrm As FrmVinicom
        ofrm = New frmTableauDeBordFournisseur
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub



    Private Sub mnuPurgeMvtStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_PurgeMvtStock.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmPurgeMvtStock
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuPurgeCommandesClients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_PurgeCommandesClients.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmPurgeCommandeClient
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuPurgeFactCom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_PurgeFactCom.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmPurgeFactCom
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuPurgeBonAppro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_PurgeBonAppro.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmPurgeBonAppro
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub m_frmActive_ToolBarUpdated() Handles m_frmActive.evt_ToolBarUpdated
        '=====================================================================
        ' Evenement : toolbar updated 
        ' Description: la form courante à mis à jour sa toolbar
        '               Donc la fên^tre prinicpale retranscrire ces booleans en affichage
        '=====================================================================
        SetToolBarEnabled()
    End Sub 'm_frmActive_ToolBarUpdated()

    Private Sub mnuControleMvtStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_ControleMvtStock.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmControleStock
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuStatQteArticleMois_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_StatQteArticleMois.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmQteCmdArticleMois
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnu_GeneFactTRP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTrp_GeneFactTRP.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGeneFactTRP
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnu_GestFactTransport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTrp_GestFactTransport.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestFactTRP
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuLstCmdTrp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTrp_LstCmdTrp.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeCommandesTrp
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuLstFactTrp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTrp_LstFactTrp.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeFactTrp
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub


    Private Sub mnuJournalAnnoncesLivraison_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_JournalAnnoncesLivraison.Click
        Dim ofrm As frmJournalAnnonces
        ofrm = New frmJournalAnnonces
        ofrm.setbAnnoncesCommmandes(True) ' Journal D'annonces LIVRAISON
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub
    Private Sub mnuJournalAnnoncesAppro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_JournalAnnoncesAppro.Click
        Dim ofrm As frmJournalAnnonces
        ofrm = New frmJournalAnnonces
        ofrm.setbAnnoncesCommmandes(False) ' Journal D'annonces APPROVISIONNEMENT
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub
    Private Sub mnuRecapAnnonces_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_RecapAnnonces.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmRecapAnnonces
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub
    Private Sub mnuCA1Client_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_CA1Client.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmStatCAClient
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub
    Private Sub mnulstPalettes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTrp_lstPalettes.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListePalettesTRP
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub
    'Private Sub dbConn_Connected() Handles dbConn.Connected, dbConn.Disconnected
    '    Try

    '        If Persist.shared_isConnected() Then
    '            Me.StatusBarDB.Text = "Connected : " & Persist.nConnection
    '        Else
    '            StatusBarDB.Text = "Disconnected"
    '        End If
    '    Catch ex As Exception

    '    End Try

    'End Sub

    Private Sub m_frmActive_evt_DisplayError() Handles m_frmActive.evt_DisplayError
        Dim strMessage As String

        strMessage = m_frmActive.getErrrorMessage
        If strMessage = "" Then
            Me.StatusBarError.Text = ""
            Exit Sub
        End If
        Beep()
        Me.StatusBarError.Text = strMessage
        If ERR_DEBUG = 1 Then
            MsgBox(strMessage, MsgBoxStyle.Critical)
        End If
        Trace.WriteLine(strMessage)
    End Sub

    Private Sub m_frmActive_evt_DisplayEtat() Handles m_frmActive.evt_DisplayStatus
        Dim strMessage As String

        strMessage = m_frmActive.getStatusMessage
        If strMessage = "" Then
            Me.StatusBarEtat.Text = ""
            Exit Sub
        End If
        Me.StatusBarEtat.Text = strMessage
    End Sub

    Private Sub mnuLstEntreesPlateforme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_LstEntreesPlateforme.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmLstEntreesPlateforme
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuLstSortiesPlateforme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_LstSortiesPlateforme.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmLstSortiesPlateforme
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub


    Private Sub mnuGenMvtIvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuCol_GenMvtIvent.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGenerationMvtInventaire
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuRecapitulatifcolisage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuCol_Recapitulatifcolisage.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmcrRecapColisage
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGenFactCol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuCol_GenFactCol.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGeneFactColisage
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGestFactcolisage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuCol_GestFactcolisage.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestFactColisage
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuLstFactColisage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuCol_LstFactColisage.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeFactColisage()
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuTarif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_Tarif.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestTarif()
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuEditTarif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_EditTarif.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmcrListeTarif
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSettings.Click
        Dim ofrm As frmSettings
        ofrm = New frmSettings()
        ofrm.ShowDialog(Me)
    End Sub

    Private Sub mnuRegistreLivraison_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTrp_RegistreLivraison.Click
        frmRegistreLivraison()
    End Sub

    Private Sub frmRegistreLivraison()
        Dim ofrm As FrmVinicom
        ofrm = New frmcrRegistreLivraison
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuRegistreAppro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuTrp_RegistreAppro.Click
        frmRregistreAppro()

    End Sub
    Private Sub frmRregistreAppro()
        Dim ofrm As FrmVinicom
        ofrm = New frmcrRegistreAppro
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuExportCompta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompta_ExportCompta.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmExportFacture
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuConstantes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuConstantes.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmConstantes
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuSaisieReglement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompta_SaisieReglement.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmSaisieReglement
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuExportReglement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompta_ExportReglement.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmExportReglement
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuListeReglement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompta_ListeReglement.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeReglement
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuLstFactureSolde_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompta_TableaudeBordFactures.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmTableauBordFacture
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuBackupRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_BackupRestore.Click
        backupRestore()
    End Sub
    Private Sub backupRestore()
        Dim odlg As dlgBackupRestore
        odlg = New dlgBackupRestore

        odlg.ShowDialog(Me)


    End Sub

    Private Sub mnuStatCAClientProducteur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_StatCAClientProducteur.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmCAClientProducteur
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuStatCAProducteurClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_StatCAProducteurClient.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmCAProducteurClient
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuLstComAProvisionner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScmd_LstComAProvisionner.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeCommissionAProvisionner
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuLstcomLitige_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScmd_LstcomLitige.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmLstSCMDLitige
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuLstcommandes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_Lstcommandes.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeCommandes()
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGestTransporteur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_GestTransporteur.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmEditTransporteur()
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuCAGroupesClients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStat_CAGroupesClients.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmStatCAGroupeClients()
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuFactComError_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_FactComError.Click
        checkDataBase()

    End Sub

    Private Sub checkDataBase()
        Dim ofrm As frmCheckDatabase
        ofrm = New frmCheckDatabase()
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub


    Private Sub mnuItemEditionFactures_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompta_EditionFactures.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmEditfactures
        ofrm.MdiParent = Me
        ofrm.WindowState = FormWindowState.Maximized
        ofrm.Show()
    End Sub

    Private Sub mnuLstBonAppro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGC_LstBonAppro.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmListeBonAppro
        ofrm.MdiParent = Me
        ofrm.Show()
    End Sub

    Private Sub mnuImportRglmt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCompta_ImportRglmt.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmImportReglement
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuLibere_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUtil_Libere.Click
        Dim ofrm As frmLibere
        ofrm = New frmLibere
        ofrm.ShowDialog()
    End Sub

    Private Sub mnuGestModeReglmt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_GestModeReglmt.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestParamModeReglement
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGestTauxTVA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_GestTauxTVA.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestParamtauxTVA
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGestTypeClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_GestTypeClient.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestParamTypeClient
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGestCouleurs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_GestCouleurs.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestParamCouleur
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGestRegions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_GestRegions.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestParamRegion
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGestConditionnement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_GestConditionnement.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestParamConditionnement
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGestContentants_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDB_GestContentants.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmGestParamContenant
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub


    Private Sub MnuUtil_checkDatabase_Click(sender As System.Object, e As System.EventArgs) Handles MnuUtil_checkDatabase.Click
        checkDataBase()
    End Sub

    Private Sub MenuItem10_Click(sender As System.Object, e As System.EventArgs) Handles mnuUtil_Droits.Click
        Dim ofrm As FrmVinicom
        ofrm = New FrmUserRights
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuGC_importCmd_Click(sender As System.Object, e As System.EventArgs) Handles mnuGC_importCmd.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmImportcommandeClientPrestashop
        ofrm.MdiParent = Me
        'ofrm.WindowState = FormWindowState.Maximized
        ofrm.Show()

    End Sub


    Private Sub mnuEclatementCommandes_Click(sender As Object, e As EventArgs) Handles mnuEclatementCommandes.Click
        Dim ofrm As frmEclatementCommande
        ofrm = New frmEclatementCommande
        ofrm.MdiParent = Me
        'ofrm.WindowState = FormWindowState.Maximized
        ofrm.Show()

    End Sub

    Private Sub mnuExportQuadra_Click(sender As Object, e As EventArgs) Handles mnuExportQuadra.Click
        Dim ofrm As frmExportQuadra
        ofrm = New frmExportQuadra
        ofrm.MdiParent = Me
        'ofrm.WindowState = FormWindowState.Maximized
        ofrm.Show()

    End Sub

    Private Sub mnuMAJLivraisonCommande_Click(sender As Object, e As EventArgs) Handles mnuMAJLivraisonCommande.Click
        Dim ofrm As frmMajLivraisonCommande
        ofrm = New frmMajLivraisonCommande
        ofrm.MdiParent = Me
        'ofrm.WindowState = FormWindowState.Maximized
        ofrm.Show()

    End Sub

    Private Sub mnuUtil_PurgePrecommande_Click(sender As Object, e As EventArgs) Handles mnuUtil_PurgePrecommande.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmPurgePreCommande
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub cbToolBarSave_Disposed(sender As Object, e As EventArgs) Handles cbToolBarSave.Disposed

    End Sub

    Private Sub mnuImportTarif_Click(sender As Object, e As EventArgs) Handles mnuImportTarif.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmImportTarif
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub

    Private Sub mnuVerifInfosLivraisons_Click(sender As Object, e As EventArgs) Handles mnuVerifInfosLivraisons.Click
        Dim ofrm As FrmVinicom
        ofrm = New frmVerificationLivraison
        ofrm.MdiParent = Me
        ofrm.Show()

    End Sub
End Class