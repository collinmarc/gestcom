Public Class frm1
    Inherits System.Windows.Forms.Form

#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()
        initConstantes()

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
    Friend WithEvents mnuFichier As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportInternet As System.Windows.Forms.MenuItem
    Friend WithEvents mnuImportInternet As System.Windows.Forms.MenuItem
    Friend WithEvents mnuQuitter As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.mnuFichier = New System.Windows.Forms.MenuItem
        Me.mnuExportInternet = New System.Windows.Forms.MenuItem
        Me.mnuImportInternet = New System.Windows.Forms.MenuItem
        Me.mnuQuitter = New System.Windows.Forms.MenuItem
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFichier})
        '
        'mnuFichier
        '
        Me.mnuFichier.Index = 0
        Me.mnuFichier.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExportInternet, Me.mnuImportInternet, Me.mnuQuitter})
        Me.mnuFichier.Text = "Fichier"
        '
        'mnuExportInternet
        '
        Me.mnuExportInternet.Index = 0
        Me.mnuExportInternet.Text = "Export internet"
        '
        'mnuImportInternet
        '
        Me.mnuImportInternet.Index = 1
        Me.mnuImportInternet.Text = "Import Internet"
        '
        'mnuQuitter
        '
        Me.mnuQuitter.Index = 2
        Me.mnuQuitter.Shortcut = System.Windows.Forms.Shortcut.AltF4
        Me.mnuQuitter.Text = "Quitter"
        '
        'frm1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Menu = Me.MainMenu1
        Me.Name = "frm1"
        Me.Text = "Vinicom -Export/Import internet"

    End Sub

#End Region

    Private Sub mnuQuitter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuQuitter.Click
        Me.Close()
    End Sub

    Private Sub mnuExportInternet_Select(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExportInternet.Select
        Dim oFRm As Form
        oFRm = New frmExportInternet
    End Sub
End Class
