Imports vini_app
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUserRights
    Inherits FrmVinicom

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
        Me.components = New System.ComponentModel.Container()
        Me.m_bsrcUserRights = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_dsVinicom = New vini_DB.dsVinicom()
        Me.RGHIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RGHTAGDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RGHROLEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RGHDROITDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.RGHTEXTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.RGHIDDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RGHTAGDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RGHROLEDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RGHDROITDataGridViewCheckBoxColumn1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.RGHTEXTDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cbValider = New System.Windows.Forms.Button()
        CType(Me.m_bsrcUserRights, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_dsVinicom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'm_bsrcUserRights
        '
        Me.m_bsrcUserRights.AllowNew = True
        Me.m_bsrcUserRights.DataMember = "USERSRIGHTS"
        Me.m_bsrcUserRights.DataSource = Me.m_dsVinicom
        '
        'm_dsVinicom
        '
        Me.m_dsVinicom.DataSetName = "dsVinicom"
        Me.m_dsVinicom.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'RGHIDDataGridViewTextBoxColumn
        '
        Me.RGHIDDataGridViewTextBoxColumn.Name = "RGHIDDataGridViewTextBoxColumn"
        '
        'RGHTAGDataGridViewTextBoxColumn
        '
        Me.RGHTAGDataGridViewTextBoxColumn.Name = "RGHTAGDataGridViewTextBoxColumn"
        '
        'RGHROLEDataGridViewTextBoxColumn
        '
        Me.RGHROLEDataGridViewTextBoxColumn.Name = "RGHROLEDataGridViewTextBoxColumn"
        '
        'RGHDROITDataGridViewCheckBoxColumn
        '
        Me.RGHDROITDataGridViewCheckBoxColumn.Name = "RGHDROITDataGridViewCheckBoxColumn"
        '
        'RGHTEXTDataGridViewTextBoxColumn
        '
        Me.RGHTEXTDataGridViewTextBoxColumn.Name = "RGHTEXTDataGridViewTextBoxColumn"
        '
        'DataGridView2
        '
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.AutoGenerateColumns = False
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.RGHIDDataGridViewTextBoxColumn1, Me.RGHTAGDataGridViewTextBoxColumn1, Me.RGHROLEDataGridViewTextBoxColumn1, Me.RGHDROITDataGridViewCheckBoxColumn1, Me.RGHTEXTDataGridViewTextBoxColumn1})
        Me.DataGridView2.DataSource = Me.m_bsrcUserRights
        Me.DataGridView2.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(647, 315)
        Me.DataGridView2.TabIndex = 0
        '
        'RGHIDDataGridViewTextBoxColumn1
        '
        Me.RGHIDDataGridViewTextBoxColumn1.DataPropertyName = "RGH_ID"
        Me.RGHIDDataGridViewTextBoxColumn1.HeaderText = "RGH_ID"
        Me.RGHIDDataGridViewTextBoxColumn1.Name = "RGHIDDataGridViewTextBoxColumn1"
        Me.RGHIDDataGridViewTextBoxColumn1.ReadOnly = True
        '
        'RGHTAGDataGridViewTextBoxColumn1
        '
        Me.RGHTAGDataGridViewTextBoxColumn1.DataPropertyName = "RGH_TAG"
        Me.RGHTAGDataGridViewTextBoxColumn1.HeaderText = "RGH_TAG"
        Me.RGHTAGDataGridViewTextBoxColumn1.Name = "RGHTAGDataGridViewTextBoxColumn1"
        '
        'RGHROLEDataGridViewTextBoxColumn1
        '
        Me.RGHROLEDataGridViewTextBoxColumn1.DataPropertyName = "RGH_ROLE"
        Me.RGHROLEDataGridViewTextBoxColumn1.HeaderText = "RGH_ROLE"
        Me.RGHROLEDataGridViewTextBoxColumn1.Name = "RGHROLEDataGridViewTextBoxColumn1"
        '
        'RGHDROITDataGridViewCheckBoxColumn1
        '
        Me.RGHDROITDataGridViewCheckBoxColumn1.DataPropertyName = "RGH_DROIT"
        Me.RGHDROITDataGridViewCheckBoxColumn1.HeaderText = "RGH_DROIT"
        Me.RGHDROITDataGridViewCheckBoxColumn1.Name = "RGHDROITDataGridViewCheckBoxColumn1"
        '
        'RGHTEXTDataGridViewTextBoxColumn1
        '
        Me.RGHTEXTDataGridViewTextBoxColumn1.DataPropertyName = "RGH_TEXT"
        Me.RGHTEXTDataGridViewTextBoxColumn1.HeaderText = "RGH_TEXT"
        Me.RGHTEXTDataGridViewTextBoxColumn1.Name = "RGHTEXTDataGridViewTextBoxColumn1"
        '
        'cbValider
        '
        Me.cbValider.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.cbValider.Location = New System.Drawing.Point(572, 329)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.Size = New System.Drawing.Size(75, 23)
        Me.cbValider.TabIndex = 1
        Me.cbValider.Text = "&Valider"
        Me.cbValider.UseVisualStyleBackColor = True
        '
        'FrmUserRights
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.ClientSize = New System.Drawing.Size(659, 364)
        Me.Controls.Add(Me.cbValider)
        Me.Controls.Add(Me.DataGridView2)
        Me.Name = "FrmUserRights"
        Me.Text = "Gestion des droits utilisateurs"
        CType(Me.m_bsrcUserRights, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_dsVinicom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents RGHIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RGHTAGDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RGHROLEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RGHDROITDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents RGHTEXTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents m_dsVinicom As vini_DB.dsVinicom
    Public WithEvents m_bsrcUserRights As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents RGHIDDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RGHTAGDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RGHROLEDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RGHDROITDataGridViewCheckBoxColumn1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents RGHTEXTDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cbValider As System.Windows.Forms.Button

End Class
