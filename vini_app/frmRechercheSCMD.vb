Imports vini_DB
Public Class frmRechercheSCMD
    Inherits System.Windows.Forms.Form
    Private m_TypeDonnees As vncTypeDonnee
    Private m_ElementSelectionne As Persist
    Friend WithEvents dgvLstSCMD As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcSourceCommande As System.Windows.Forms.BindingSource
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateCommandeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRSDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalHTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents montantCommission As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents tbCodeFournisseur As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtdateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtDatedeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.tbSelectionner = New System.Windows.Forms.Button
        Me.cbAnnuler = New System.Windows.Forms.Button
        Me.tbCodeFournisseur = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.dgvLstSCMD = New System.Windows.Forms.DataGridView
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DateCommandeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersRSDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TotalHTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.montantCommission = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcSourceCommande = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.dgvLstSCMD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcSourceCommande, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(308, 90)
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
        'tbCodeFournisseur
        '
        Me.tbCodeFournisseur.Enabled = False
        Me.tbCodeFournisseur.Location = New System.Drawing.Point(160, 48)
        Me.tbCodeFournisseur.Name = "tbCodeFournisseur"
        Me.tbCodeFournisseur.Size = New System.Drawing.Size(208, 20)
        Me.tbCodeFournisseur.TabIndex = 103
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 106
        Me.Label1.Text = "Code Fournisseur"
        '
        'dtdateFin
        '
        Me.dtdateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdateFin.Location = New System.Drawing.Point(160, 24)
        Me.dtdateFin.Name = "dtdateFin"
        Me.dtdateFin.Size = New System.Drawing.Size(88, 20)
        Me.dtdateFin.TabIndex = 102
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(0, 24)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(151, 16)
        Me.Label8.TabIndex = 105
        Me.Label8.Text = "date de fin Fact. Fourn."
        '
        'dtDatedeb
        '
        Me.dtDatedeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDatedeb.Location = New System.Drawing.Point(160, 0)
        Me.dtDatedeb.Name = "dtDatedeb"
        Me.dtDatedeb.Size = New System.Drawing.Size(88, 20)
        Me.dtDatedeb.TabIndex = 101
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(0, 4)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(154, 16)
        Me.Label14.TabIndex = 104
        Me.Label14.Text = "date de début Fact. Fourn."
        '
        'dgvLstSCMD
        '
        Me.dgvLstSCMD.AllowUserToAddRows = False
        Me.dgvLstSCMD.AllowUserToDeleteRows = False
        Me.dgvLstSCMD.AutoGenerateColumns = False
        Me.dgvLstSCMD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvLstSCMD.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.DateCommandeDataGridViewTextBoxColumn, Me.TiersRSDataGridViewTextBoxColumn, Me.TotalHTDataGridViewTextBoxColumn, Me.montantCommission})
        Me.dgvLstSCMD.DataSource = Me.m_bsrcSourceCommande
        Me.dgvLstSCMD.Location = New System.Drawing.Point(0, 122)
        Me.dgvLstSCMD.Name = "dgvLstSCMD"
        Me.dgvLstSCMD.ReadOnly = True
        Me.dgvLstSCMD.Size = New System.Drawing.Size(455, 432)
        Me.dgvLstSCMD.TabIndex = 108
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DateCommandeDataGridViewTextBoxColumn
        '
        Me.DateCommandeDataGridViewTextBoxColumn.DataPropertyName = "dateCommande"
        Me.DateCommandeDataGridViewTextBoxColumn.HeaderText = "Date Commande"
        Me.DateCommandeDataGridViewTextBoxColumn.Name = "DateCommandeDataGridViewTextBoxColumn"
        Me.DateCommandeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TiersRSDataGridViewTextBoxColumn
        '
        Me.TiersRSDataGridViewTextBoxColumn.DataPropertyName = "TiersRS"
        Me.TiersRSDataGridViewTextBoxColumn.HeaderText = "Client"
        Me.TiersRSDataGridViewTextBoxColumn.Name = "TiersRSDataGridViewTextBoxColumn"
        Me.TiersRSDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TotalHTDataGridViewTextBoxColumn
        '
        Me.TotalHTDataGridViewTextBoxColumn.DataPropertyName = "totalHT"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle1.Format = "C2"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.TotalHTDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle1
        Me.TotalHTDataGridViewTextBoxColumn.HeaderText = "H.T."
        Me.TotalHTDataGridViewTextBoxColumn.Name = "TotalHTDataGridViewTextBoxColumn"
        Me.TotalHTDataGridViewTextBoxColumn.ReadOnly = True
        '
        'montantCommission
        '
        Me.montantCommission.DataPropertyName = "montantCommission"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.Format = "C2"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.montantCommission.DefaultCellStyle = DataGridViewCellStyle2
        Me.montantCommission.HeaderText = "Commission"
        Me.montantCommission.Name = "montantCommission"
        Me.montantCommission.ReadOnly = True
        '
        'm_bsrcSourceCommande
        '
        Me.m_bsrcSourceCommande.DataSource = GetType(vini_DB.SousCommande)
        '
        'frmRechercheSCMD
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(456, 590)
        Me.Controls.Add(Me.dgvLstSCMD)
        Me.Controls.Add(Me.tbCodeFournisseur)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtdateFin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtDatedeb)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.tbSelectionner)
        Me.Controls.Add(Me.cbAfficher)
        Me.KeyPreview = True
        Me.Name = "frmRechercheSCMD"
        Me.Text = "Recherche"
        CType(Me.dgvLstSCMD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcSourceCommande, System.ComponentModel.ISupportInitialize).EndInit()
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
    Public Sub setCodeFourn(ByVal pstrCodeFourn As String)
        tbCodeFournisseur.Text = pstrCodeFourn
    End Sub
#End Region
    Private Sub cbAfficher_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        Afficher()
    End Sub
    Private Sub Afficher()
        Dim ddeb As Date
        Dim dfin As Date
        Dim codeFourn As String
        Dim col As Collection
        Dim oSCMD As SousCommande
        Try


            ddeb = dtDatedeb.Value.ToShortDateString
            dfin = dtdateFin.Value.ToShortDateString
            codeFourn = tbCodeFournisseur.Text
            col = SousCommande.getListeAFacturer(ddeb, dfin, codeFourn)
            Debug.Assert(Not col Is Nothing, Persist.getErreur())
            m_bsrcSourceCommande.Clear()
            If Not col Is Nothing Then
                For Each oSCMD In col
                    m_bsrcSourceCommande.Add(oSCMD)
                Next
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub SelectListBoxItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSelectionner.Click
        selectionneEtSort()
    End Sub
    Private Sub selectionneEtSort()
        If Not m_bsrcSourceCommande.Current Is Nothing Then
            Try
                m_ElementSelectionne = CType(m_bsrcSourceCommande.Current, SousCommande)
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Catch ex As Exception
                Debug.Assert(False, ex.ToString)
            End Try
        End If

    End Sub 'selectionneEtSort


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


    Private Sub dgvLstSCMD_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvLstSCMD.DoubleClick
        selectionneEtSort()
    End Sub
End Class
