<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLibere
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbMachineName = New System.Windows.Forms.TextBox
        Me.cbAffiche = New System.Windows.Forms.Button
        Me.cbSave = New System.Windows.Forms.Button
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.DsVinicom1 = New vini_DB.dsVinicom
        Me.m_bsLOCK = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_taLOCK = New vini_DB.dsVinicomTableAdapters.LOCKTableAdapter
        Me.LCKTYPEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LCKPERSISTIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LCKNAMEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LCKDATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DsVinicom1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsLOCK, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Nom du poste de travail : "
        '
        'tbMachineName
        '
        Me.tbMachineName.Location = New System.Drawing.Point(147, 10)
        Me.tbMachineName.Name = "tbMachineName"
        Me.tbMachineName.Size = New System.Drawing.Size(100, 20)
        Me.tbMachineName.TabIndex = 1
        '
        'cbAffiche
        '
        Me.cbAffiche.Location = New System.Drawing.Point(499, 8)
        Me.cbAffiche.Name = "cbAffiche"
        Me.cbAffiche.Size = New System.Drawing.Size(75, 23)
        Me.cbAffiche.TabIndex = 2
        Me.cbAffiche.Text = "Affiche"
        Me.cbAffiche.UseVisualStyleBackColor = True
        '
        'cbSave
        '
        Me.cbSave.Location = New System.Drawing.Point(499, 342)
        Me.cbSave.Name = "cbSave"
        Me.cbSave.Size = New System.Drawing.Size(75, 23)
        Me.cbSave.TabIndex = 4
        Me.cbSave.Text = "Sauver"
        Me.cbSave.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.LCKTYPEDataGridViewTextBoxColumn, Me.LCKPERSISTIDDataGridViewTextBoxColumn, Me.LCKNAMEDataGridViewTextBoxColumn, Me.LCKDATEDataGridViewTextBoxColumn})
        Me.DataGridView1.DataMember = "LOCK"
        Me.DataGridView1.DataSource = Me.DsVinicom1
        Me.DataGridView1.Location = New System.Drawing.Point(12, 36)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(562, 301)
        Me.DataGridView1.TabIndex = 3
        '
        'DsVinicom1
        '
        Me.DsVinicom1.DataSetName = "dsVinicom"
        Me.DsVinicom1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'm_bsLOCK
        '
        Me.m_bsLOCK.DataMember = "LOCK"
        Me.m_bsLOCK.DataSource = Me.DsVinicom1
        '
        'm_taLOCK
        '
        Me.m_taLOCK.ClearBeforeFill = True
        '
        'LCKTYPEDataGridViewTextBoxColumn
        '
        Me.LCKTYPEDataGridViewTextBoxColumn.DataPropertyName = "LCK_TYPE"
        Me.LCKTYPEDataGridViewTextBoxColumn.HeaderText = "Type de l'élement"
        Me.LCKTYPEDataGridViewTextBoxColumn.Name = "LCKTYPEDataGridViewTextBoxColumn"
        Me.LCKTYPEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LCKPERSISTIDDataGridViewTextBoxColumn
        '
        Me.LCKPERSISTIDDataGridViewTextBoxColumn.DataPropertyName = "LCK_PERSISTID"
        Me.LCKPERSISTIDDataGridViewTextBoxColumn.HeaderText = "ID de l'élement"
        Me.LCKPERSISTIDDataGridViewTextBoxColumn.Name = "LCKPERSISTIDDataGridViewTextBoxColumn"
        Me.LCKPERSISTIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LCKNAMEDataGridViewTextBoxColumn
        '
        Me.LCKNAMEDataGridViewTextBoxColumn.DataPropertyName = "LCK_NAME"
        Me.LCKNAMEDataGridViewTextBoxColumn.HeaderText = "Poste "
        Me.LCKNAMEDataGridViewTextBoxColumn.Name = "LCKNAMEDataGridViewTextBoxColumn"
        Me.LCKNAMEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LCKDATEDataGridViewTextBoxColumn
        '
        Me.LCKDATEDataGridViewTextBoxColumn.DataPropertyName = "LCK_DATE"
        Me.LCKDATEDataGridViewTextBoxColumn.HeaderText = "Date du bloquage"
        Me.LCKDATEDataGridViewTextBoxColumn.Name = "LCKDATEDataGridViewTextBoxColumn"
        Me.LCKDATEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'frmLibere
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(586, 377)
        Me.Controls.Add(Me.cbSave)
        Me.Controls.Add(Me.cbAffiche)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.tbMachineName)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLibere"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Gestion des bloquages d'élements"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DsVinicom1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsLOCK, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbMachineName As System.Windows.Forms.TextBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DsVinicom1 As vini_DB.dsVinicom
    Friend WithEvents m_bsLOCK As System.Windows.Forms.BindingSource
    Friend WithEvents cbAffiche As System.Windows.Forms.Button
    Friend WithEvents m_taLOCK As vini_DB.dsVinicomTableAdapters.LOCKTableAdapter
    Friend WithEvents cbSave As System.Windows.Forms.Button
    Friend WithEvents LCKTYPEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LCKPERSISTIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LCKNAMEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LCKDATEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
