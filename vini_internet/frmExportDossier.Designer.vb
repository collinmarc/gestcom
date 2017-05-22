<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExportDossier
    Inherits System.Windows.Forms.Form

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
        Me.m_dlgDossier = New System.Windows.Forms.FolderBrowserDialog
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbDossier = New System.Windows.Forms.TextBox
        Me.tbBrowse = New System.Windows.Forms.Button
        Me.tbExporter = New System.Windows.Forms.Button
        Me.dgvStatus = New System.Windows.Forms.DataGridView
        Me.StatusDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.StatusMessageDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcStatus = New System.Windows.Forms.BindingSource(Me.components)
        Me.cbDeverrouillageFrom = New System.Windows.Forms.Button
        Me.cbDeverrouillageTo = New System.Windows.Forms.Button
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'm_dlgDossier
        '
        Me.m_dlgDossier.Description = "Choix du dossier à exporter"
        Me.m_dlgDossier.RootFolder = System.Environment.SpecialFolder.MyComputer
        Me.m_dlgDossier.ShowNewFolderButton = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Dossier : "
        '
        'tbDossier
        '
        Me.tbDossier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbDossier.Location = New System.Drawing.Point(71, 42)
        Me.tbDossier.Name = "tbDossier"
        Me.tbDossier.Size = New System.Drawing.Size(473, 20)
        Me.tbDossier.TabIndex = 1
        '
        'tbBrowse
        '
        Me.tbBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbBrowse.Location = New System.Drawing.Point(550, 40)
        Me.tbBrowse.Name = "tbBrowse"
        Me.tbBrowse.Size = New System.Drawing.Size(44, 23)
        Me.tbBrowse.TabIndex = 2
        Me.tbBrowse.Text = "..."
        Me.tbBrowse.UseVisualStyleBackColor = True
        '
        'tbExporter
        '
        Me.tbExporter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbExporter.Location = New System.Drawing.Point(18, 69)
        Me.tbExporter.Name = "tbExporter"
        Me.tbExporter.Size = New System.Drawing.Size(576, 23)
        Me.tbExporter.TabIndex = 3
        Me.tbExporter.Text = "Exporter"
        Me.tbExporter.UseVisualStyleBackColor = True
        '
        'dgvStatus
        '
        Me.dgvStatus.AllowUserToAddRows = False
        Me.dgvStatus.AllowUserToDeleteRows = False
        Me.dgvStatus.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvStatus.AutoGenerateColumns = False
        Me.dgvStatus.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStatus.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.StatusDateDataGridViewTextBoxColumn, Me.StatusMessageDataGridViewTextBoxColumn})
        Me.dgvStatus.DataSource = Me.m_bsrcStatus
        Me.dgvStatus.Location = New System.Drawing.Point(18, 105)
        Me.dgvStatus.Name = "dgvStatus"
        Me.dgvStatus.ReadOnly = True
        Me.dgvStatus.Size = New System.Drawing.Size(576, 219)
        Me.dgvStatus.TabIndex = 4
        '
        'StatusDateDataGridViewTextBoxColumn
        '
        Me.StatusDateDataGridViewTextBoxColumn.DataPropertyName = "statusDate"
        Me.StatusDateDataGridViewTextBoxColumn.HeaderText = "Date"
        Me.StatusDateDataGridViewTextBoxColumn.Name = "StatusDateDataGridViewTextBoxColumn"
        Me.StatusDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.StatusDateDataGridViewTextBoxColumn.Width = 55
        '
        'StatusMessageDataGridViewTextBoxColumn
        '
        Me.StatusMessageDataGridViewTextBoxColumn.DataPropertyName = "statusMessage"
        Me.StatusMessageDataGridViewTextBoxColumn.HeaderText = "Message"
        Me.StatusMessageDataGridViewTextBoxColumn.Name = "StatusMessageDataGridViewTextBoxColumn"
        Me.StatusMessageDataGridViewTextBoxColumn.ReadOnly = True
        Me.StatusMessageDataGridViewTextBoxColumn.Width = 75
        '
        'm_bsrcStatus
        '
        Me.m_bsrcStatus.DataSource = GetType(vini_internet.clsExportstatus)
        '
        'cbDeverrouillageFrom
        '
        Me.cbDeverrouillageFrom.Location = New System.Drawing.Point(18, 13)
        Me.cbDeverrouillageFrom.Name = "cbDeverrouillageFrom"
        Me.cbDeverrouillageFrom.Size = New System.Drawing.Size(126, 23)
        Me.cbDeverrouillageFrom.TabIndex = 5
        Me.cbDeverrouillageFrom.Text = "Déverrouillage From"
        Me.cbDeverrouillageFrom.UseVisualStyleBackColor = True
        '
        'cbDeverrouillageTo
        '
        Me.cbDeverrouillageTo.Location = New System.Drawing.Point(150, 13)
        Me.cbDeverrouillageTo.Name = "cbDeverrouillageTo"
        Me.cbDeverrouillageTo.Size = New System.Drawing.Size(126, 23)
        Me.cbDeverrouillageTo.TabIndex = 6
        Me.cbDeverrouillageTo.Text = "Déverrouillage To"
        Me.cbDeverrouillageTo.UseVisualStyleBackColor = True
        '
        'frmExportDossier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(606, 336)
        Me.Controls.Add(Me.cbDeverrouillageTo)
        Me.Controls.Add(Me.cbDeverrouillageFrom)
        Me.Controls.Add(Me.dgvStatus)
        Me.Controls.Add(Me.tbExporter)
        Me.Controls.Add(Me.tbBrowse)
        Me.Controls.Add(Me.tbDossier)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmExportDossier"
        Me.Text = "Export d'un dossier"
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcStatus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents m_dlgDossier As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbDossier As System.Windows.Forms.TextBox
    Friend WithEvents tbBrowse As System.Windows.Forms.Button
    Friend WithEvents tbExporter As System.Windows.Forms.Button
    Friend WithEvents m_bsrcStatus As System.Windows.Forms.BindingSource
    Friend WithEvents dgvStatus As System.Windows.Forms.DataGridView
    Friend WithEvents StatusDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusMessageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cbDeverrouillageFrom As System.Windows.Forms.Button
    Friend WithEvents cbDeverrouillageTo As System.Windows.Forms.Button
End Class
