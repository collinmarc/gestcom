<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditTransporteur
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
        Me.components = New System.ComponentModel.Container
        Me.m_bsrcTransporteur = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_dsVinicom = New vini_DB.dsVinicom
        Me.m_datagrid = New System.Windows.Forms.DataGridView
        Me.TRPIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPCODEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPNOMDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLIVRUE1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLIVRUE2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLIVCPDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLIVVILLEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLIVTELDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLIVFAXDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLIVPORTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLIVEMAILDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPDEFAUTDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.TRPCOMMDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TRPLICENCEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.m_bsrcTransporteur, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_dsVinicom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_datagrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'm_bsrcTransporteur
        '
        Me.m_bsrcTransporteur.DataMember = "TRANSPORTEUR"
        Me.m_bsrcTransporteur.DataSource = Me.m_dsVinicom
        '
        'm_dsVinicom
        '
        Me.m_dsVinicom.DataSetName = "dsVinicom"
        Me.m_dsVinicom.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'm_datagrid
        '
        Me.m_datagrid.AutoGenerateColumns = False
        Me.m_datagrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.m_datagrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.m_datagrid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.TRPIDDataGridViewTextBoxColumn, Me.TRPCODEDataGridViewTextBoxColumn, Me.TRPNOMDataGridViewTextBoxColumn, Me.TRPLIVRUE1DataGridViewTextBoxColumn, Me.TRPLIVRUE2DataGridViewTextBoxColumn, Me.TRPLIVCPDataGridViewTextBoxColumn, Me.TRPLIVVILLEDataGridViewTextBoxColumn, Me.TRPLIVTELDataGridViewTextBoxColumn, Me.TRPLIVFAXDataGridViewTextBoxColumn, Me.TRPLIVPORTDataGridViewTextBoxColumn, Me.TRPLIVEMAILDataGridViewTextBoxColumn, Me.TRPDEFAUTDataGridViewCheckBoxColumn, Me.TRPCOMMDataGridViewTextBoxColumn, Me.TRPLICENCEDataGridViewTextBoxColumn})
        Me.m_datagrid.DataSource = Me.m_bsrcTransporteur
        Me.m_datagrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.m_datagrid.Location = New System.Drawing.Point(0, 0)
        Me.m_datagrid.Name = "m_datagrid"
        Me.m_datagrid.Size = New System.Drawing.Size(746, 380)
        Me.m_datagrid.TabIndex = 0
        '
        'TRPIDDataGridViewTextBoxColumn
        '
        Me.TRPIDDataGridViewTextBoxColumn.DataPropertyName = "TRP_ID"
        Me.TRPIDDataGridViewTextBoxColumn.HeaderText = "TRP_ID"
        Me.TRPIDDataGridViewTextBoxColumn.Name = "TRPIDDataGridViewTextBoxColumn"
        Me.TRPIDDataGridViewTextBoxColumn.ReadOnly = True
        Me.TRPIDDataGridViewTextBoxColumn.Width = 71
        '
        'TRPCODEDataGridViewTextBoxColumn
        '
        Me.TRPCODEDataGridViewTextBoxColumn.DataPropertyName = "TRP_CODE"
        Me.TRPCODEDataGridViewTextBoxColumn.HeaderText = "TRP_CODE"
        Me.TRPCODEDataGridViewTextBoxColumn.Name = "TRPCODEDataGridViewTextBoxColumn"
        Me.TRPCODEDataGridViewTextBoxColumn.Width = 90
        '
        'TRPNOMDataGridViewTextBoxColumn
        '
        Me.TRPNOMDataGridViewTextBoxColumn.DataPropertyName = "TRP_NOM"
        Me.TRPNOMDataGridViewTextBoxColumn.HeaderText = "TRP_NOM"
        Me.TRPNOMDataGridViewTextBoxColumn.Name = "TRPNOMDataGridViewTextBoxColumn"
        Me.TRPNOMDataGridViewTextBoxColumn.Width = 85
        '
        'TRPLIVRUE1DataGridViewTextBoxColumn
        '
        Me.TRPLIVRUE1DataGridViewTextBoxColumn.DataPropertyName = "TRP_LIV_RUE1"
        Me.TRPLIVRUE1DataGridViewTextBoxColumn.HeaderText = "TRP_LIV_RUE1"
        Me.TRPLIVRUE1DataGridViewTextBoxColumn.Name = "TRPLIVRUE1DataGridViewTextBoxColumn"
        Me.TRPLIVRUE1DataGridViewTextBoxColumn.Width = 111
        '
        'TRPLIVRUE2DataGridViewTextBoxColumn
        '
        Me.TRPLIVRUE2DataGridViewTextBoxColumn.DataPropertyName = "TRP_LIV_RUE2"
        Me.TRPLIVRUE2DataGridViewTextBoxColumn.HeaderText = "TRP_LIV_RUE2"
        Me.TRPLIVRUE2DataGridViewTextBoxColumn.Name = "TRPLIVRUE2DataGridViewTextBoxColumn"
        Me.TRPLIVRUE2DataGridViewTextBoxColumn.Width = 111
        '
        'TRPLIVCPDataGridViewTextBoxColumn
        '
        Me.TRPLIVCPDataGridViewTextBoxColumn.DataPropertyName = "TRP_LIV_CP"
        Me.TRPLIVCPDataGridViewTextBoxColumn.HeaderText = "TRP_LIV_CP"
        Me.TRPLIVCPDataGridViewTextBoxColumn.Name = "TRPLIVCPDataGridViewTextBoxColumn"
        Me.TRPLIVCPDataGridViewTextBoxColumn.Width = 96
        '
        'TRPLIVVILLEDataGridViewTextBoxColumn
        '
        Me.TRPLIVVILLEDataGridViewTextBoxColumn.DataPropertyName = "TRP_LIV_VILLE"
        Me.TRPLIVVILLEDataGridViewTextBoxColumn.HeaderText = "TRP_LIV_VILLE"
        Me.TRPLIVVILLEDataGridViewTextBoxColumn.Name = "TRPLIVVILLEDataGridViewTextBoxColumn"
        Me.TRPLIVVILLEDataGridViewTextBoxColumn.Width = 111
        '
        'TRPLIVTELDataGridViewTextBoxColumn
        '
        Me.TRPLIVTELDataGridViewTextBoxColumn.DataPropertyName = "TRP_LIV_TEL"
        Me.TRPLIVTELDataGridViewTextBoxColumn.HeaderText = "TRP_LIV_TEL"
        Me.TRPLIVTELDataGridViewTextBoxColumn.Name = "TRPLIVTELDataGridViewTextBoxColumn"
        Me.TRPLIVTELDataGridViewTextBoxColumn.Width = 102
        '
        'TRPLIVFAXDataGridViewTextBoxColumn
        '
        Me.TRPLIVFAXDataGridViewTextBoxColumn.DataPropertyName = "TRP_LIV_FAX"
        Me.TRPLIVFAXDataGridViewTextBoxColumn.HeaderText = "TRP_LIV_FAX"
        Me.TRPLIVFAXDataGridViewTextBoxColumn.Name = "TRPLIVFAXDataGridViewTextBoxColumn"
        Me.TRPLIVFAXDataGridViewTextBoxColumn.Width = 102
        '
        'TRPLIVPORTDataGridViewTextBoxColumn
        '
        Me.TRPLIVPORTDataGridViewTextBoxColumn.DataPropertyName = "TRP_LIV_PORT"
        Me.TRPLIVPORTDataGridViewTextBoxColumn.HeaderText = "TRP_LIV_PORT"
        Me.TRPLIVPORTDataGridViewTextBoxColumn.Name = "TRPLIVPORTDataGridViewTextBoxColumn"
        Me.TRPLIVPORTDataGridViewTextBoxColumn.Width = 112
        '
        'TRPLIVEMAILDataGridViewTextBoxColumn
        '
        Me.TRPLIVEMAILDataGridViewTextBoxColumn.DataPropertyName = "TRP_LIV_EMAIL"
        Me.TRPLIVEMAILDataGridViewTextBoxColumn.HeaderText = "TRP_LIV_EMAIL"
        Me.TRPLIVEMAILDataGridViewTextBoxColumn.Name = "TRPLIVEMAILDataGridViewTextBoxColumn"
        Me.TRPLIVEMAILDataGridViewTextBoxColumn.Width = 114
        '
        'TRPDEFAUTDataGridViewCheckBoxColumn
        '
        Me.TRPDEFAUTDataGridViewCheckBoxColumn.DataPropertyName = "TRP_DEFAUT"
        Me.TRPDEFAUTDataGridViewCheckBoxColumn.HeaderText = "TRP_DEFAUT"
        Me.TRPDEFAUTDataGridViewCheckBoxColumn.Name = "TRPDEFAUTDataGridViewCheckBoxColumn"
        Me.TRPDEFAUTDataGridViewCheckBoxColumn.Width = 84
        '
        'TRPCOMMDataGridViewTextBoxColumn
        '
        Me.TRPCOMMDataGridViewTextBoxColumn.DataPropertyName = "TRP_COMM"
        Me.TRPCOMMDataGridViewTextBoxColumn.HeaderText = "TRP_COMM"
        Me.TRPCOMMDataGridViewTextBoxColumn.Name = "TRPCOMMDataGridViewTextBoxColumn"
        Me.TRPCOMMDataGridViewTextBoxColumn.Width = 93
        '
        'TRPLICENCEDataGridViewTextBoxColumn
        '
        Me.TRPLICENCEDataGridViewTextBoxColumn.DataPropertyName = "TRP_LICENCE"
        Me.TRPLICENCEDataGridViewTextBoxColumn.HeaderText = "TRP_LICENCE"
        Me.TRPLICENCEDataGridViewTextBoxColumn.Name = "TRPLICENCEDataGridViewTextBoxColumn"
        Me.TRPLICENCEDataGridViewTextBoxColumn.Width = 105
        '
        'frmEditTransporteur
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(746, 380)
        Me.Controls.Add(Me.m_datagrid)
        Me.Name = "frmEditTransporteur"
        Me.Text = "frmEditTransporteur"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.m_bsrcTransporteur, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_dsVinicom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_datagrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents m_bsrcTransporteur As System.Windows.Forms.BindingSource
    Friend WithEvents m_dsVinicom As vini_DB.dsVinicom
    Friend WithEvents m_datagrid As System.Windows.Forms.DataGridView
    Friend WithEvents TRPIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPCODEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPNOMDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLIVRUE1DataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLIVRUE2DataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLIVCPDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLIVVILLEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLIVTELDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLIVFAXDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLIVPORTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLIVEMAILDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPDEFAUTDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents TRPCOMMDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRPLICENCEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
