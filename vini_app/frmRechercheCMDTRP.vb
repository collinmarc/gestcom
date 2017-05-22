Imports vini_DB
Public Class frmRechercheCMDTRP
    Inherits System.Windows.Forms.Form
    Private m_TypeDonnees As vncTypeDonnee
    Private m_ElementSelectionne As Persist
    Friend WithEvents m_bsrcFacture As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRSDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateLivraisonDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MontantTransportDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Private m_ocol As Collection

#Region " Code généré par le Concepteur Windows Form "


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
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents tbSelectionner As System.Windows.Forms.Button
    Friend WithEvents cbAnnuler As System.Windows.Forms.Button
    Friend WithEvents laNom As System.Windows.Forms.Label
    Friend WithEvents laCode As System.Windows.Forms.Label
    Friend WithEvents tbNom As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtDateDeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtDateFin As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.tbSelectionner = New System.Windows.Forms.Button
        Me.cbAnnuler = New System.Windows.Forms.Button
        Me.laNom = New System.Windows.Forms.Label
        Me.laCode = New System.Windows.Forms.Label
        Me.tbNom = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtDateDeb = New System.Windows.Forms.DateTimePicker
        Me.dtDateFin = New System.Windows.Forms.DateTimePicker
        Me.m_bsrcFacture = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersRSDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateLivraisonDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MontantTransportDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.m_bsrcFacture, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(312, 80)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(136, 24)
        Me.cbAfficher.TabIndex = 4
        Me.cbAfficher.Text = "A&fficher"
        '
        'tbSelectionner
        '
        Me.tbSelectionner.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.tbSelectionner.Location = New System.Drawing.Point(232, 560)
        Me.tbSelectionner.Name = "tbSelectionner"
        Me.tbSelectionner.Size = New System.Drawing.Size(104, 23)
        Me.tbSelectionner.TabIndex = 5
        Me.tbSelectionner.Text = "&Sélectionner"
        '
        'cbAnnuler
        '
        Me.cbAnnuler.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cbAnnuler.Location = New System.Drawing.Point(344, 560)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.Size = New System.Drawing.Size(104, 24)
        Me.cbAnnuler.TabIndex = 6
        Me.cbAnnuler.Text = "&Annuler"
        '
        'laNom
        '
        Me.laNom.Location = New System.Drawing.Point(8, 56)
        Me.laNom.Name = "laNom"
        Me.laNom.Size = New System.Drawing.Size(80, 16)
        Me.laNom.TabIndex = 110
        Me.laNom.Text = "Nom"
        '
        'laCode
        '
        Me.laCode.Location = New System.Drawing.Point(8, 8)
        Me.laCode.Name = "laCode"
        Me.laCode.Size = New System.Drawing.Size(72, 24)
        Me.laCode.TabIndex = 109
        Me.laCode.Text = "Date Debut"
        '
        'tbNom
        '
        Me.tbNom.Location = New System.Drawing.Point(94, 56)
        Me.tbNom.Name = "tbNom"
        Me.tbNom.Size = New System.Drawing.Size(354, 20)
        Me.tbNom.TabIndex = 114
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 24)
        Me.Label1.TabIndex = 115
        Me.Label1.Text = "Date Fin"
        '
        'dtDateDeb
        '
        Me.dtDateDeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateDeb.Location = New System.Drawing.Point(96, 8)
        Me.dtDateDeb.Name = "dtDateDeb"
        Me.dtDateDeb.Size = New System.Drawing.Size(104, 20)
        Me.dtDateDeb.TabIndex = 116
        '
        'dtDateFin
        '
        Me.dtDateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFin.Location = New System.Drawing.Point(96, 32)
        Me.dtDateFin.Name = "dtDateFin"
        Me.dtDateFin.Size = New System.Drawing.Size(104, 20)
        Me.dtDateFin.TabIndex = 117
        '
        'm_bsrcFacture
        '
        Me.m_bsrcFacture.DataSource = GetType(vini_DB.CommandeClient)
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.TiersCodeDataGridViewTextBoxColumn, Me.TiersRSDataGridViewTextBoxColumn, Me.DateLivraisonDataGridViewTextBoxColumn, Me.MontantTransportDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcFacture
        Me.DataGridView1.Location = New System.Drawing.Point(13, 117)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(431, 437)
        Me.DataGridView1.TabIndex = 118
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "Code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TiersCodeDataGridViewTextBoxColumn
        '
        Me.TiersCodeDataGridViewTextBoxColumn.DataPropertyName = "TiersCode"
        Me.TiersCodeDataGridViewTextBoxColumn.FillWeight = 200.0!
        Me.TiersCodeDataGridViewTextBoxColumn.HeaderText = "Code CLT"
        Me.TiersCodeDataGridViewTextBoxColumn.Name = "TiersCodeDataGridViewTextBoxColumn"
        Me.TiersCodeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TiersRSDataGridViewTextBoxColumn
        '
        Me.TiersRSDataGridViewTextBoxColumn.DataPropertyName = "TiersRS"
        Me.TiersRSDataGridViewTextBoxColumn.HeaderText = "RS CLT"
        Me.TiersRSDataGridViewTextBoxColumn.Name = "TiersRSDataGridViewTextBoxColumn"
        Me.TiersRSDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DateLivraisonDataGridViewTextBoxColumn
        '
        Me.DateLivraisonDataGridViewTextBoxColumn.DataPropertyName = "dateLivraison"
        Me.DateLivraisonDataGridViewTextBoxColumn.HeaderText = "Date Liv."
        Me.DateLivraisonDataGridViewTextBoxColumn.Name = "DateLivraisonDataGridViewTextBoxColumn"
        Me.DateLivraisonDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MontantTransportDataGridViewTextBoxColumn
        '
        Me.MontantTransportDataGridViewTextBoxColumn.DataPropertyName = "montantTransport"
        DataGridViewCellStyle2.Format = "C2"
        Me.MontantTransportDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle2
        Me.MontantTransportDataGridViewTextBoxColumn.HeaderText = "Mt Transport"
        Me.MontantTransportDataGridViewTextBoxColumn.Name = "MontantTransportDataGridViewTextBoxColumn"
        Me.MontantTransportDataGridViewTextBoxColumn.ReadOnly = True
        '
        'frmRechercheCMDTRP
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(456, 590)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.dtDateFin)
        Me.Controls.Add(Me.dtDateDeb)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbNom)
        Me.Controls.Add(Me.laNom)
        Me.Controls.Add(Me.laCode)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.tbSelectionner)
        Me.Controls.Add(Me.cbAfficher)
        Me.KeyPreview = True
        Me.Name = "frmRechercheCMDTRP"
        Me.Text = "Recherche Commande avec transport"
        CType(Me.m_bsrcFacture, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
#Region "Accesseurs"
    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

    End Sub


    Public Function getElementSelectionne() As Persist
        Return m_ElementSelectionne
    End Function
    Public Sub setRSClient(ByVal pstrRSClient As String)
        tbNom.Text = pstrRSClient
    End Sub
#End Region
    Private Sub cbAfficher_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        Afficher()
    End Sub
    Private Sub Afficher()
        Dim col As Collection
        col = CommandeClient.getListe_TRP(dtDateDeb.Value.ToShortDateString, dtDateFin.Value.ToShortDateString, tbNom.Text)
        Debug.Assert(Not col Is Nothing, Persist.getErreur())
        If Not col Is Nothing Then
            m_ocol = col
            displayListe()
        End If
    End Sub

    Private Sub SelectListBoxItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSelectionner.Click
        selectionneEtSort()
    End Sub
    Private Sub selectionneEtSort()
        If Not m_bsrcFacture.Current Is Nothing Then
            Try
                m_ElementSelectionne = m_bsrcFacture.Current
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Catch ex As Exception
                Debug.Assert(False, ex.ToString)
            End Try
        End If

    End Sub 'selectionneEtSort
    Public Sub setListe(ByVal oCol As Collection)
        m_ocol = oCol
    End Sub

    Public Sub displayListe()
        Debug.Assert(Not m_ocol Is Nothing)
        Dim obj As Object


        m_bsrcFacture.Clear()
        For Each obj In m_ocol
            m_bsrcFacture.Add(obj)
        Next obj


        If m_ocol.Count > 0 Then
            m_bsrcFacture.Position = 1
        End If
        m_bsrcFacture.ResetBindings(False)

    End Sub

    Private Sub tbCode_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            Afficher()
        End If
    End Sub

    Private Sub lbTiers_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            selectionneEtSort()
            e.Handled = True
        End If

    End Sub


    Private Sub frmRechercheDB_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(27) Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If
    End Sub

    Private Sub frmRechercheDB_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
        'If lbTiers.Items.Count > 0 Then
        '    lbTiers.SelectedIndex = 0
        '    lbTiers.Focus()
        'Else
        '    lbTiers.SelectedIndex = -1
        '    tbCode.Focus()
        'End If
    End Sub


    Private Sub tbMotCle_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            Afficher()
        End If

    End Sub

    Private Sub tbNom_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            Afficher()
        End If

    End Sub


    Private Sub flxTiers_DblClick(ByVal sender As Object, ByVal e As System.EventArgs)
        selectionneEtSort()

    End Sub

    Private Sub laNom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles laNom.Click

    End Sub

    Private Sub DataGridView1_RowHeaderMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseDoubleClick
        selectionneEtSort()
    End Sub
End Class
