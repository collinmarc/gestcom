<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGestParamContenant
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
        Me.m_bsrcParam = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.cbOK = New System.Windows.Forms.Button
        Me.cbCancel = New System.Windows.Forms.Button
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LibelleDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CentDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BoutDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PoidsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DefautDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        CType(Me.m_bsrcParam, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'm_bsrcParam
        '
        Me.m_bsrcParam.DataSource = GetType(vini_DB.contenant)
        Me.m_bsrcParam.Filter = ""
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
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.LibelleDataGridViewTextBoxColumn, Me.CentDataGridViewTextBoxColumn, Me.BoutDataGridViewTextBoxColumn, Me.PoidsDataGridViewTextBoxColumn, Me.DefautDataGridViewCheckBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcParam
        Me.DataGridView1.Location = New System.Drawing.Point(3, 2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(697, 310)
        Me.DataGridView1.TabIndex = 0
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
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        '
        'LibelleDataGridViewTextBoxColumn
        '
        Me.LibelleDataGridViewTextBoxColumn.DataPropertyName = "libelle"
        Me.LibelleDataGridViewTextBoxColumn.HeaderText = "libelle"
        Me.LibelleDataGridViewTextBoxColumn.Name = "LibelleDataGridViewTextBoxColumn"
        '
        'CentDataGridViewTextBoxColumn
        '
        Me.CentDataGridViewTextBoxColumn.DataPropertyName = "cent"
        Me.CentDataGridViewTextBoxColumn.HeaderText = "Litres"
        Me.CentDataGridViewTextBoxColumn.Name = "CentDataGridViewTextBoxColumn"
        '
        'BoutDataGridViewTextBoxColumn
        '
        Me.BoutDataGridViewTextBoxColumn.DataPropertyName = "bout"
        Me.BoutDataGridViewTextBoxColumn.HeaderText = "equ bouteilles"
        Me.BoutDataGridViewTextBoxColumn.Name = "BoutDataGridViewTextBoxColumn"
        '
        'PoidsDataGridViewTextBoxColumn
        '
        Me.PoidsDataGridViewTextBoxColumn.DataPropertyName = "poids"
        Me.PoidsDataGridViewTextBoxColumn.HeaderText = "poids"
        Me.PoidsDataGridViewTextBoxColumn.Name = "PoidsDataGridViewTextBoxColumn"
        '
        'DefautDataGridViewCheckBoxColumn
        '
        Me.DefautDataGridViewCheckBoxColumn.DataPropertyName = "defaut"
        Me.DefautDataGridViewCheckBoxColumn.HeaderText = "defaut"
        Me.DefautDataGridViewCheckBoxColumn.Name = "DefautDataGridViewCheckBoxColumn"
        '
        'frmGestParamContenant
        '
        Me.ClientSize = New System.Drawing.Size(712, 342)
        Me.Controls.Add(Me.cbCancel)
        Me.Controls.Add(Me.cbOK)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "frmGestParamContenant"
        Me.Text = "Gestion des contenants"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.m_bsrcParam, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents m_bsrcParam As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents cbOK As System.Windows.Forms.Button
    Friend WithEvents cbCancel As System.Windows.Forms.Button
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LibelleDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CentDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BoutDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PoidsDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DefautDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn

End Class
