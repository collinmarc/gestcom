
Imports vini_DB
Public Class frmGestTarif
    Inherits frmDonBase
    Protected m_colProduit As Collection
    Friend WithEvents m_bsrcProduit As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcFournisseur As System.Windows.Forms.BindingSource
    Friend WithEvents lbFournisseur As System.Windows.Forms.ListBox
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NomDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MillesimeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibCouleurDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibConditionnementDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibContenantDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TarifADataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TarifBDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TarifCDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView

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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.m_bsrcProduit = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_bsrcFournisseur = New System.Windows.Forms.BindingSource(Me.components)
        Me.lbFournisseur = New System.Windows.Forms.ListBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NomDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MillesimeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibCouleurDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibConditionnementDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibContenantDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TarifADataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TarifBDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TarifCDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.m_bsrcProduit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcFournisseur, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'm_bsrcProduit
        '
        Me.m_bsrcProduit.DataSource = GetType(vini_DB.Produit)
        '
        'm_bsrcFournisseur
        '
        Me.m_bsrcFournisseur.DataSource = GetType(vini_DB.Fournisseur)
        '
        'lbFournisseur
        '
        Me.lbFournisseur.DataSource = Me.m_bsrcFournisseur
        Me.lbFournisseur.DisplayMember = "rs"
        Me.lbFournisseur.FormattingEnabled = True
        Me.lbFournisseur.Location = New System.Drawing.Point(12, 12)
        Me.lbFournisseur.Name = "lbFournisseur"
        Me.lbFournisseur.Size = New System.Drawing.Size(337, 602)
        Me.lbFournisseur.TabIndex = 0
        Me.lbFournisseur.ValueMember = "id"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.NomDataGridViewTextBoxColumn, Me.MillesimeDataGridViewTextBoxColumn, Me.LibCouleurDataGridViewTextBoxColumn, Me.LibConditionnementDataGridViewTextBoxColumn, Me.LibContenantDataGridViewTextBoxColumn, Me.TarifADataGridViewTextBoxColumn, Me.TarifBDataGridViewTextBoxColumn, Me.TarifCDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcProduit
        Me.DataGridView1.Location = New System.Drawing.Point(380, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(594, 588)
        Me.DataGridView1.TabIndex = 1
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "Code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn.Width = 50
        '
        'NomDataGridViewTextBoxColumn
        '
        Me.NomDataGridViewTextBoxColumn.DataPropertyName = "nom"
        Me.NomDataGridViewTextBoxColumn.HeaderText = "Libellé"
        Me.NomDataGridViewTextBoxColumn.Name = "NomDataGridViewTextBoxColumn"
        Me.NomDataGridViewTextBoxColumn.ReadOnly = True
        Me.NomDataGridViewTextBoxColumn.Width = 150
        '
        'MillesimeDataGridViewTextBoxColumn
        '
        Me.MillesimeDataGridViewTextBoxColumn.DataPropertyName = "millesime"
        Me.MillesimeDataGridViewTextBoxColumn.HeaderText = "millesime"
        Me.MillesimeDataGridViewTextBoxColumn.Name = "MillesimeDataGridViewTextBoxColumn"
        Me.MillesimeDataGridViewTextBoxColumn.ReadOnly = True
        Me.MillesimeDataGridViewTextBoxColumn.Width = 20
        '
        'LibCouleurDataGridViewTextBoxColumn
        '
        Me.LibCouleurDataGridViewTextBoxColumn.DataPropertyName = "libCouleur"
        Me.LibCouleurDataGridViewTextBoxColumn.HeaderText = "Couleur"
        Me.LibCouleurDataGridViewTextBoxColumn.Name = "LibCouleurDataGridViewTextBoxColumn"
        Me.LibCouleurDataGridViewTextBoxColumn.ReadOnly = True
        Me.LibCouleurDataGridViewTextBoxColumn.Width = 50
        '
        'LibConditionnementDataGridViewTextBoxColumn
        '
        Me.LibConditionnementDataGridViewTextBoxColumn.DataPropertyName = "libConditionnement"
        Me.LibConditionnementDataGridViewTextBoxColumn.HeaderText = "Conditionnement"
        Me.LibConditionnementDataGridViewTextBoxColumn.Name = "LibConditionnementDataGridViewTextBoxColumn"
        Me.LibConditionnementDataGridViewTextBoxColumn.ReadOnly = True
        Me.LibConditionnementDataGridViewTextBoxColumn.Width = 50
        '
        'LibContenantDataGridViewTextBoxColumn
        '
        Me.LibContenantDataGridViewTextBoxColumn.DataPropertyName = "libContenant"
        Me.LibContenantDataGridViewTextBoxColumn.HeaderText = "Contenant"
        Me.LibContenantDataGridViewTextBoxColumn.Name = "LibContenantDataGridViewTextBoxColumn"
        Me.LibContenantDataGridViewTextBoxColumn.ReadOnly = True
        Me.LibContenantDataGridViewTextBoxColumn.Width = 50
        '
        'TarifADataGridViewTextBoxColumn
        '
        Me.TarifADataGridViewTextBoxColumn.DataPropertyName = "TarifA"
        Me.TarifADataGridViewTextBoxColumn.HeaderText = "Tarif A"
        Me.TarifADataGridViewTextBoxColumn.Name = "TarifADataGridViewTextBoxColumn"
        Me.TarifADataGridViewTextBoxColumn.Width = 50
        '
        'TarifBDataGridViewTextBoxColumn
        '
        Me.TarifBDataGridViewTextBoxColumn.DataPropertyName = "TarifB"
        Me.TarifBDataGridViewTextBoxColumn.HeaderText = "Tarif B"
        Me.TarifBDataGridViewTextBoxColumn.Name = "TarifBDataGridViewTextBoxColumn"
        Me.TarifBDataGridViewTextBoxColumn.Width = 50
        '
        'TarifCDataGridViewTextBoxColumn
        '
        Me.TarifCDataGridViewTextBoxColumn.DataPropertyName = "TarifC"
        Me.TarifCDataGridViewTextBoxColumn.HeaderText = "Tarif C"
        Me.TarifCDataGridViewTextBoxColumn.Name = "TarifCDataGridViewTextBoxColumn"
        Me.TarifCDataGridViewTextBoxColumn.Width = 50
        '
        'frmGestTarif
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(997, 630)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.lbFournisseur)
        Me.Name = "frmGestTarif"
        Me.Text = "Gestion de tarif"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.m_bsrcProduit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcFournisseur, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

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
    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        MyBase.AddHandlerValidated(ocol)
        RemoveHandler lbFournisseur.SelectedValueChanged, AddressOf ControlUpdated
        RemoveHandler DataGridView1.Validated, AddressOf ControlUpdated
    End Sub



    Protected Overrides Function frmSave() As Boolean
        Dim objProduit As Produit

        For Each objProduit In m_colProduit
            objProduit.save()
        Next
        setfrmNotUpdated()
        Return True
    End Function

    Public Overrides Function getResume() As String 'Rend le caption de la fenêtre
        Return "Gestion des Tarifs"
    End Function 'getResume

#End Region

#Region "Methodes privées"
#End Region
#Region "Gestion des Evenements"


#End Region

    Private Sub lbFournisseur_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbFournisseur.SelectedIndexChanged
        Dim oPRD As Produit
        Dim oFRN As Fournisseur

        oFRN = lbFournisseur.SelectedItem

        sauvegardeElementCourant()

        If (Not oFRN Is Nothing) Then
            setcursorWait()
            m_colProduit = Produit.getListe(vncTypeProduit.vncTous, , , , oFRN.id)
            DataGridView1.Enabled = False
            m_bsrcProduit.Clear()
            For Each oPRD In m_colProduit
                oPRD.load()
                m_bsrcProduit.Add(oPRD)
            Next
            DataGridView1.Enabled = True
            restoreCursor()
        End If

        setElementCourant2(oFRN)

    End Sub

    Private Sub frmGestTarif_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim oColFRN As Collection
        Dim oFRN As Fournisseur

        oColFRN = Fournisseur.getListe()
        For Each oFRN In oColFRN
            m_bsrcFournisseur.Add(oFRN)
        Next

        lbFournisseur.Enabled = True
        setfrmNotUpdated()

    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If (Not getElementCourant() Is Nothing) Then
            setfrmUpdated()
        End If

    End Sub
End Class


