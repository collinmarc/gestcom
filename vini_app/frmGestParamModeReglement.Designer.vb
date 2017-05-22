<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGestParamModeReglement
    Inherits vini_app.FrmVinicom

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.m_bsrcParamModeReglement = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ValeurDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Valeur2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DDebutEcheanceDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.DefautDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.cbOK = New System.Windows.Forms.Button
        Me.cbCancel = New System.Windows.Forms.Button
        CType(Me.m_bsrcParamModeReglement, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'm_bsrcParamModeReglement
        '
        Me.m_bsrcParamModeReglement.DataSource = GetType(vini_DB.ParamModeReglement)
        Me.m_bsrcParamModeReglement.Filter = "bDelete = false"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.ValeurDataGridViewTextBoxColumn, Me.Valeur2DataGridViewTextBoxColumn, Me.DDebutEcheanceDataGridViewTextBoxColumn, Me.DefautDataGridViewCheckBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcParamModeReglement
        Me.DataGridView1.Location = New System.Drawing.Point(3, 2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(697, 310)
        Me.DataGridView1.TabIndex = 0
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        '
        'ValeurDataGridViewTextBoxColumn
        '
        Me.ValeurDataGridViewTextBoxColumn.DataPropertyName = "valeur"
        Me.ValeurDataGridViewTextBoxColumn.HeaderText = "Libellé"
        Me.ValeurDataGridViewTextBoxColumn.Name = "ValeurDataGridViewTextBoxColumn"
        '
        'Valeur2DataGridViewTextBoxColumn
        '
        Me.Valeur2DataGridViewTextBoxColumn.DataPropertyName = "valeur2"
        Me.Valeur2DataGridViewTextBoxColumn.HeaderText = "Nbre jours"
        Me.Valeur2DataGridViewTextBoxColumn.Name = "Valeur2DataGridViewTextBoxColumn"
        '
        'DDebutEcheanceDataGridViewTextBoxColumn
        '
        Me.DDebutEcheanceDataGridViewTextBoxColumn.DataPropertyName = "dDebutEcheance"
        Me.DDebutEcheanceDataGridViewTextBoxColumn.HeaderText = "Début Calcul"
        Me.DDebutEcheanceDataGridViewTextBoxColumn.Items.AddRange(New Object() {"FDM", "FACT"})
        Me.DDebutEcheanceDataGridViewTextBoxColumn.Name = "DDebutEcheanceDataGridViewTextBoxColumn"
        Me.DDebutEcheanceDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DDebutEcheanceDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'DefautDataGridViewCheckBoxColumn
        '
        Me.DefautDataGridViewCheckBoxColumn.DataPropertyName = "defaut"
        Me.DefautDataGridViewCheckBoxColumn.HeaderText = "defaut"
        Me.DefautDataGridViewCheckBoxColumn.Name = "DefautDataGridViewCheckBoxColumn"
        '
        'cbOK
        '
        Me.cbOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbOK.Location = New System.Drawing.Point(529, 318)
        Me.cbOK.Name = "cbOK"
        Me.cbOK.Size = New System.Drawing.Size(75, 23)
        Me.cbOK.TabIndex = 1
        Me.cbOK.Text = "OK"
        Me.cbOK.UseVisualStyleBackColor = True
        '
        'cbCancel
        '
        Me.cbCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbCancel.Location = New System.Drawing.Point(621, 317)
        Me.cbCancel.Name = "cbCancel"
        Me.cbCancel.Size = New System.Drawing.Size(75, 23)
        Me.cbCancel.TabIndex = 2
        Me.cbCancel.Text = "Annuler"
        Me.cbCancel.UseVisualStyleBackColor = True
        '
        'frmGestParamModeReglement
        '
        Me.ClientSize = New System.Drawing.Size(712, 342)
        Me.Controls.Add(Me.cbCancel)
        Me.Controls.Add(Me.cbOK)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "frmGestParamModeReglement"
        Me.Text = "Gestion des modes de reglements"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.m_bsrcParamModeReglement, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents m_bsrcParamModeReglement As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents cbOK As System.Windows.Forms.Button
    Friend WithEvents cbCancel As System.Windows.Forms.Button
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ValeurDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Valeur2DataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DDebutEcheanceDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DefautDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn

End Class
