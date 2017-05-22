Imports vini_DB
Public Class frmMDI
    Inherits System.Windows.Forms.Form

    Private WithEvents m_frmActive As FrmVinicom
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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFichier As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportInternet As System.Windows.Forms.MenuItem
    Friend WithEvents mnuImportInternet As System.Windows.Forms.MenuItem
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarDB As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarError As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarEtat As System.Windows.Forms.StatusBarPanel
    Friend WithEvents tsToolBar As System.Windows.Forms.ToolBar
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents tsbExport As System.Windows.Forms.ToolBarButton
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSettings As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportDossier As System.Windows.Forms.MenuItem
    Friend WithEvents mnuImportDossier As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportPrecommande As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents tsbImport As System.Windows.Forms.ToolBarButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMDI))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFichier = New System.Windows.Forms.MenuItem()
        Me.mnuExportInternet = New System.Windows.Forms.MenuItem()
        Me.mnuImportInternet = New System.Windows.Forms.MenuItem()
        Me.mnuExportDossier = New System.Windows.Forms.MenuItem()
        Me.mnuImportDossier = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuExportPrecommande = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.mnuSettings = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.StatusBar1 = New System.Windows.Forms.StatusBar()
        Me.StatusBarDB = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarError = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarEtat = New System.Windows.Forms.StatusBarPanel()
        Me.tsToolBar = New System.Windows.Forms.ToolBar()
        Me.tsbExport = New System.Windows.Forms.ToolBarButton()
        Me.tsbImport = New System.Windows.Forms.ToolBarButton()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        CType(Me.StatusBarDB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarError, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarEtat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFichier})
        '
        'mnuFichier
        '
        Me.mnuFichier.Index = 0
        Me.mnuFichier.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExportInternet, Me.mnuImportInternet, Me.mnuExportDossier, Me.mnuImportDossier, Me.MenuItem2, Me.mnuExportPrecommande, Me.MenuItem1, Me.MenuItem3, Me.mnuSettings, Me.MenuItem4, Me.MenuItem5})
        Me.mnuFichier.Text = "&Fichier"
        '
        'mnuExportInternet
        '
        Me.mnuExportInternet.Index = 0
        Me.mnuExportInternet.Text = "&Export Internet"
        '
        'mnuImportInternet
        '
        Me.mnuImportInternet.Index = 1
        Me.mnuImportInternet.Text = "&Import Internet"
        '
        'mnuExportDossier
        '
        Me.mnuExportDossier.Index = 2
        Me.mnuExportDossier.Text = "Export d'un dossier"
        '
        'mnuImportDossier
        '
        Me.mnuImportDossier.Index = 3
        Me.mnuImportDossier.Text = "Import d'un dossier"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 4
        Me.MenuItem2.Text = "-"
        '
        'mnuExportPrecommande
        '
        Me.mnuExportPrecommande.Index = 5
        Me.mnuExportPrecommande.Text = "Export des PréCommandes"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 6
        Me.MenuItem1.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 7
        Me.MenuItem3.Text = "Etat du serveur"
        '
        'mnuSettings
        '
        Me.mnuSettings.Index = 8
        Me.mnuSettings.Text = "Settings"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 9
        Me.MenuItem4.Text = "-"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 10
        Me.MenuItem5.Text = "&Quitter"
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 603)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarDB, Me.StatusBarError, Me.StatusBarEtat})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(1016, 22)
        Me.StatusBar1.TabIndex = 1
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
        Me.StatusBarEtat.Name = "StatusBarEtat"
        Me.StatusBarEtat.Width = 400
        '
        'tsToolBar
        '
        Me.tsToolBar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.tsbExport, Me.tsbImport})
        Me.tsToolBar.DropDownArrows = True
        Me.tsToolBar.ImageList = Me.ImageList1
        Me.tsToolBar.Location = New System.Drawing.Point(0, 0)
        Me.tsToolBar.Name = "tsToolBar"
        Me.tsToolBar.ShowToolTips = True
        Me.tsToolBar.Size = New System.Drawing.Size(1016, 28)
        Me.tsToolBar.TabIndex = 3
        '
        'tsbExport
        '
        Me.tsbExport.ImageIndex = 0
        Me.tsbExport.Name = "tsbExport"
        Me.tsbExport.Tag = "EXPORT"
        '
        'tsbImport
        '
        Me.tsbImport.ImageIndex = 1
        Me.tsbImport.Name = "tsbImport"
        Me.tsbImport.Tag = "IMPORT"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Magenta
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        '
        'frmMDI
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 625)
        Me.Controls.Add(Me.tsToolBar)
        Me.Controls.Add(Me.StatusBar1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMenu1
        Me.Name = "frmMDI"
        Me.Text = "Vinicom : Internet"
        CType(Me.StatusBarDB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarError, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarEtat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub frmMDI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ofrmSettings As frmSettings

        initConstantes()
        If Not Persist.shared_connect() Then
            ofrmSettings = New frmSettings()
            ofrmSettings.ShowDialog()
            Exit Sub
        Else
            Persist.shared_disconnect()
        End If
        Param.LoadcolParams()
        Dim colTrp As Collection = Transporteur.colTransporteur
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Me.Close()
    End Sub

    Private Sub mnuExportInternet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportInternet.Click
        export()
    End Sub

    Private Sub mnuImportInternet_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuImportInternet.Click
        import()
    End Sub
    Private Sub import()
        m_frmActive = New frmImportInternet
        m_frmActive.MdiParent = Me
        m_frmActive.Show()

    End Sub
    Private Sub export()
        m_frmActive = New frmExportInternet
        m_frmActive.MdiParent = Me
        m_frmActive.WindowState = FormWindowState.Maximized
        m_frmActive.Show()
    End Sub


    Private Sub tsToolBar_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tsToolBar.ButtonClick
        Select Case UCase(e.Button.Tag.ToString())
            Case "EXPORT"
                export()
            Case "IMPORT"
                import()
        End Select

    End Sub

    Private Sub mnuSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSettings.Click
        Dim ofrm As frmSettings
        ofrm = New frmSettings
        ofrm.ShowDialog(Me)
    End Sub

    Private Sub mnuExportDossier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportDossier.Click
        Dim ofrm As frmExportDossier
        ofrm = New frmExportDossier()
        ofrm.MdiParent = Me
        ofrm.WindowState = FormWindowState.Maximized
        ofrm.Show()
    End Sub

    Private Sub mnuImportDossier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImportDossier.Click
        Dim ofrm As frmImportDossier
        ofrm = New frmImportDossier()
        ofrm.MdiParent = Me
        ofrm.WindowState = FormWindowState.Maximized
        ofrm.Show()

    End Sub

    Private Sub mnuExportPrecommande_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportPrecommande.Click
        Dim ofrm As frmExportPreCommande
        ofrm = New frmExportPreCommande
        ofrm.MdiParent = Me
        ofrm.WindowState = FormWindowState.Maximized
        ofrm.Show()

    End Sub

    Private Sub MenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles MenuItem3.Click
        Dim ofrm As frmEtatServeur
        ofrm = New frmEtatServeur
        ofrm.ShowDialog(Me)
    End Sub
End Class
