<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCheckDatabase
    Inherits Form

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
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.DIFFERENCE = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbAnnee = New System.Windows.Forms.TextBox
        Me.cbExecuter = New System.Windows.Forms.Button
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.DataGridView3 = New System.Windows.Forms.DataGridView
        Me.Label5 = New System.Windows.Forms.Label
        Me.DataGridView4 = New System.Windows.Forms.DataGridView
        Me.CODEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CMDDATEDataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NUMDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QTEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcBACMDSansMVTSTK = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_dsError = New vini_DB.dsError
        Me.CMDCODEDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CMDDATEDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CLTNOMDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FRNCODEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FRNRSDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcError3 = New System.Windows.Forms.BindingSource(Me.components)
        Me.CMDCODEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CMDDATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NBREFRNDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NBRESCMDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcCMDERROR = New System.Windows.Forms.BindingSource(Me.components)
        Me.FCTCODEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FCTDATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FCTPERIODEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FRNNOMDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FCTTOTALHTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SOMMELIGNESDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcFactComError = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_oTAFactComError = New vini_DB.dsErrorTableAdapters.FACTCOMERRORTableAdapter
        Me.m_oTACmdErreur = New vini_DB.dsErrorTableAdapters.COMMANDE_ERRORTableAdapter
        Me.m_oTACMD_ERROR3 = New vini_DB.dsErrorTableAdapters.COMMANDE_SCMD_NULLTableAdapter
        Me.m_oTABASansMVTSTK = New vini_DB.dsErrorTableAdapters.BASansMVTSTKTableAdapter
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcBACMDSansMVTSTK, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_dsError, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcError3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcCMDERROR, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcFactComError, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FCTCODEDataGridViewTextBoxColumn, Me.FCTDATEDataGridViewTextBoxColumn, Me.FCTPERIODEDataGridViewTextBoxColumn, Me.FRNNOMDataGridViewTextBoxColumn, Me.FCTTOTALHTDataGridViewTextBoxColumn, Me.SOMMELIGNESDataGridViewTextBoxColumn, Me.DIFFERENCE})
        Me.DataGridView1.DataSource = Me.m_bsrcFactComError
        Me.DataGridView1.Location = New System.Drawing.Point(12, 51)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(717, 253)
        Me.DataGridView1.TabIndex = 0
        '
        'DIFFERENCE
        '
        Me.DIFFERENCE.DataPropertyName = "DIFFERENCE"
        DataGridViewCellStyle10.Format = "C2"
        DataGridViewCellStyle10.NullValue = Nothing
        Me.DIFFERENCE.DefaultCellStyle = DataGridViewCellStyle10
        Me.DIFFERENCE.HeaderText = "DIFFERENCE"
        Me.DIFFERENCE.Name = "DIFFERENCE"
        Me.DIFFERENCE.ReadOnly = True
        Me.DIFFERENCE.Width = 99
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Année : "
        '
        'tbAnnee
        '
        Me.tbAnnee.Location = New System.Drawing.Point(66, 5)
        Me.tbAnnee.Name = "tbAnnee"
        Me.tbAnnee.Size = New System.Drawing.Size(100, 20)
        Me.tbAnnee.TabIndex = 2
        '
        'cbExecuter
        '
        Me.cbExecuter.Location = New System.Drawing.Point(172, 3)
        Me.cbExecuter.Name = "cbExecuter"
        Me.cbExecuter.Size = New System.Drawing.Size(75, 23)
        Me.cbExecuter.TabIndex = 3
        Me.cbExecuter.Text = "Executer"
        Me.cbExecuter.UseVisualStyleBackColor = True
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.AutoGenerateColumns = False
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CMDCODEDataGridViewTextBoxColumn, Me.CMDDATEDataGridViewTextBoxColumn, Me.NBREFRNDataGridViewTextBoxColumn, Me.NBRESCMDDataGridViewTextBoxColumn})
        Me.DataGridView2.DataSource = Me.m_bsrcCMDERROR
        Me.DataGridView2.Location = New System.Drawing.Point(9, 340)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.Size = New System.Drawing.Size(424, 113)
        Me.DataGridView2.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(402, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "1- Liste des factures de commissions avec un total différents de la somme des lig" & _
            "nes"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(9, 307)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(367, 30)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "2- Liste des Commandes avec un nombre de sous-commandes <> au nombre de fournisse" & _
            "urs de la commande"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(448, 307)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(286, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "3- Liste des Commandes ayant perdu des sous-commandes"
        '
        'DataGridView3
        '
        Me.DataGridView3.AllowUserToAddRows = False
        Me.DataGridView3.AllowUserToDeleteRows = False
        Me.DataGridView3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView3.AutoGenerateColumns = False
        Me.DataGridView3.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CMDCODEDataGridViewTextBoxColumn1, Me.CMDDATEDataGridViewTextBoxColumn1, Me.CLTNOMDataGridViewTextBoxColumn, Me.FRNCODEDataGridViewTextBoxColumn, Me.FRNRSDataGridViewTextBoxColumn})
        Me.DataGridView3.DataSource = Me.m_bsrcError3
        Me.DataGridView3.Location = New System.Drawing.Point(439, 340)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.ReadOnly = True
        Me.DataGridView3.Size = New System.Drawing.Size(290, 113)
        Me.DataGridView3.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 456)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(354, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "4- Liste des bons d'appro  et commandes livrés sans mouvement de stock"
        '
        'DataGridView4
        '
        Me.DataGridView4.AllowUserToAddRows = False
        Me.DataGridView4.AllowUserToDeleteRows = False
        Me.DataGridView4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView4.AutoGenerateColumns = False
        Me.DataGridView4.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CODEDataGridViewTextBoxColumn, Me.CMDDATEDataGridViewTextBoxColumn2, Me.NUMDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn1, Me.QTEDataGridViewTextBoxColumn})
        Me.DataGridView4.DataSource = Me.m_bsrcBACMDSansMVTSTK
        Me.DataGridView4.Location = New System.Drawing.Point(9, 472)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.ReadOnly = True
        Me.DataGridView4.Size = New System.Drawing.Size(424, 127)
        Me.DataGridView4.TabIndex = 10
        '
        'CODEDataGridViewTextBoxColumn
        '
        Me.CODEDataGridViewTextBoxColumn.DataPropertyName = "CODE"
        Me.CODEDataGridViewTextBoxColumn.HeaderText = "Code"
        Me.CODEDataGridViewTextBoxColumn.Name = "CODEDataGridViewTextBoxColumn"
        Me.CODEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CMDDATEDataGridViewTextBoxColumn2
        '
        Me.CMDDATEDataGridViewTextBoxColumn2.DataPropertyName = "CMD_DATE"
        Me.CMDDATEDataGridViewTextBoxColumn2.HeaderText = "Date"
        Me.CMDDATEDataGridViewTextBoxColumn2.Name = "CMDDATEDataGridViewTextBoxColumn2"
        Me.CMDDATEDataGridViewTextBoxColumn2.ReadOnly = True
        '
        'NUMDataGridViewTextBoxColumn
        '
        Me.NUMDataGridViewTextBoxColumn.DataPropertyName = "NUM"
        Me.NUMDataGridViewTextBoxColumn.HeaderText = "NLg"
        Me.NUMDataGridViewTextBoxColumn.Name = "NUMDataGridViewTextBoxColumn"
        Me.NUMDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "PRD_CODE"
        Me.DataGridViewTextBoxColumn1.HeaderText = "Produit"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'QTEDataGridViewTextBoxColumn
        '
        Me.QTEDataGridViewTextBoxColumn.DataPropertyName = "QTE"
        Me.QTEDataGridViewTextBoxColumn.HeaderText = "QTE"
        Me.QTEDataGridViewTextBoxColumn.Name = "QTEDataGridViewTextBoxColumn"
        Me.QTEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'm_bsrcBACMDSansMVTSTK
        '
        Me.m_bsrcBACMDSansMVTSTK.DataMember = "BASansMVTSTK"
        Me.m_bsrcBACMDSansMVTSTK.DataSource = Me.m_dsError
        '
        'm_dsError
        '
        Me.m_dsError.DataSetName = "dsError"
        Me.m_dsError.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CMDCODEDataGridViewTextBoxColumn1
        '
        Me.CMDCODEDataGridViewTextBoxColumn1.DataPropertyName = "CMD_CODE"
        Me.CMDCODEDataGridViewTextBoxColumn1.HeaderText = "Commande"
        Me.CMDCODEDataGridViewTextBoxColumn1.Name = "CMDCODEDataGridViewTextBoxColumn1"
        Me.CMDCODEDataGridViewTextBoxColumn1.ReadOnly = True
        '
        'CMDDATEDataGridViewTextBoxColumn1
        '
        Me.CMDDATEDataGridViewTextBoxColumn1.DataPropertyName = "CMD_DATE"
        Me.CMDDATEDataGridViewTextBoxColumn1.HeaderText = "Date"
        Me.CMDDATEDataGridViewTextBoxColumn1.Name = "CMDDATEDataGridViewTextBoxColumn1"
        Me.CMDDATEDataGridViewTextBoxColumn1.ReadOnly = True
        '
        'CLTNOMDataGridViewTextBoxColumn
        '
        Me.CLTNOMDataGridViewTextBoxColumn.DataPropertyName = "CLT_NOM"
        Me.CLTNOMDataGridViewTextBoxColumn.HeaderText = "Client"
        Me.CLTNOMDataGridViewTextBoxColumn.Name = "CLTNOMDataGridViewTextBoxColumn"
        Me.CLTNOMDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FRNCODEDataGridViewTextBoxColumn
        '
        Me.FRNCODEDataGridViewTextBoxColumn.DataPropertyName = "FRN_CODE"
        Me.FRNCODEDataGridViewTextBoxColumn.HeaderText = "Producteur"
        Me.FRNCODEDataGridViewTextBoxColumn.Name = "FRNCODEDataGridViewTextBoxColumn"
        Me.FRNCODEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FRNRSDataGridViewTextBoxColumn
        '
        Me.FRNRSDataGridViewTextBoxColumn.DataPropertyName = "FRN_RS"
        Me.FRNRSDataGridViewTextBoxColumn.HeaderText = "RS Producteur"
        Me.FRNRSDataGridViewTextBoxColumn.Name = "FRNRSDataGridViewTextBoxColumn"
        Me.FRNRSDataGridViewTextBoxColumn.ReadOnly = True
        '
        'm_bsrcError3
        '
        Me.m_bsrcError3.DataMember = "COMMANDE_SCMD_NULL"
        Me.m_bsrcError3.DataSource = Me.m_dsError
        '
        'CMDCODEDataGridViewTextBoxColumn
        '
        Me.CMDCODEDataGridViewTextBoxColumn.DataPropertyName = "CMD_CODE"
        Me.CMDCODEDataGridViewTextBoxColumn.HeaderText = "Commande "
        Me.CMDCODEDataGridViewTextBoxColumn.Name = "CMDCODEDataGridViewTextBoxColumn"
        Me.CMDCODEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CMDDATEDataGridViewTextBoxColumn
        '
        Me.CMDDATEDataGridViewTextBoxColumn.DataPropertyName = "CMD_DATE"
        Me.CMDDATEDataGridViewTextBoxColumn.HeaderText = "Date de Commande"
        Me.CMDDATEDataGridViewTextBoxColumn.Name = "CMDDATEDataGridViewTextBoxColumn"
        Me.CMDDATEDataGridViewTextBoxColumn.ReadOnly = True
        '
        'NBREFRNDataGridViewTextBoxColumn
        '
        Me.NBREFRNDataGridViewTextBoxColumn.DataPropertyName = "NBRE_FRN"
        Me.NBREFRNDataGridViewTextBoxColumn.HeaderText = "Nbre Fourn"
        Me.NBREFRNDataGridViewTextBoxColumn.Name = "NBREFRNDataGridViewTextBoxColumn"
        Me.NBREFRNDataGridViewTextBoxColumn.ReadOnly = True
        '
        'NBRESCMDDataGridViewTextBoxColumn
        '
        Me.NBRESCMDDataGridViewTextBoxColumn.DataPropertyName = "NBRE_SCMD"
        Me.NBRESCMDDataGridViewTextBoxColumn.HeaderText = "NBbre SCMD"
        Me.NBRESCMDDataGridViewTextBoxColumn.Name = "NBRESCMDDataGridViewTextBoxColumn"
        Me.NBRESCMDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'm_bsrcCMDERROR
        '
        Me.m_bsrcCMDERROR.DataMember = "COMMANDE_ERROR"
        Me.m_bsrcCMDERROR.DataSource = Me.m_dsError
        '
        'FCTCODEDataGridViewTextBoxColumn
        '
        Me.FCTCODEDataGridViewTextBoxColumn.DataPropertyName = "FCT_CODE"
        Me.FCTCODEDataGridViewTextBoxColumn.HeaderText = "FCT_CODE"
        Me.FCTCODEDataGridViewTextBoxColumn.Name = "FCTCODEDataGridViewTextBoxColumn"
        Me.FCTCODEDataGridViewTextBoxColumn.ReadOnly = True
        Me.FCTCODEDataGridViewTextBoxColumn.Width = 88
        '
        'FCTDATEDataGridViewTextBoxColumn
        '
        Me.FCTDATEDataGridViewTextBoxColumn.DataPropertyName = "FCT_DATE"
        Me.FCTDATEDataGridViewTextBoxColumn.HeaderText = "FCT_DATE"
        Me.FCTDATEDataGridViewTextBoxColumn.Name = "FCTDATEDataGridViewTextBoxColumn"
        Me.FCTDATEDataGridViewTextBoxColumn.ReadOnly = True
        Me.FCTDATEDataGridViewTextBoxColumn.Width = 87
        '
        'FCTPERIODEDataGridViewTextBoxColumn
        '
        Me.FCTPERIODEDataGridViewTextBoxColumn.DataPropertyName = "FCT_PERIODE"
        Me.FCTPERIODEDataGridViewTextBoxColumn.HeaderText = "FCT_PERIODE"
        Me.FCTPERIODEDataGridViewTextBoxColumn.Name = "FCTPERIODEDataGridViewTextBoxColumn"
        Me.FCTPERIODEDataGridViewTextBoxColumn.ReadOnly = True
        Me.FCTPERIODEDataGridViewTextBoxColumn.Width = 106
        '
        'FRNNOMDataGridViewTextBoxColumn
        '
        Me.FRNNOMDataGridViewTextBoxColumn.DataPropertyName = "FRN_NOM"
        Me.FRNNOMDataGridViewTextBoxColumn.HeaderText = "FRN_NOM"
        Me.FRNNOMDataGridViewTextBoxColumn.Name = "FRNNOMDataGridViewTextBoxColumn"
        Me.FRNNOMDataGridViewTextBoxColumn.ReadOnly = True
        Me.FRNNOMDataGridViewTextBoxColumn.Width = 85
        '
        'FCTTOTALHTDataGridViewTextBoxColumn
        '
        Me.FCTTOTALHTDataGridViewTextBoxColumn.DataPropertyName = "FCT_TOTAL_HT"
        DataGridViewCellStyle11.Format = "C2"
        DataGridViewCellStyle11.NullValue = Nothing
        Me.FCTTOTALHTDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle11
        Me.FCTTOTALHTDataGridViewTextBoxColumn.HeaderText = "FCT_TOTAL_HT"
        Me.FCTTOTALHTDataGridViewTextBoxColumn.Name = "FCTTOTALHTDataGridViewTextBoxColumn"
        Me.FCTTOTALHTDataGridViewTextBoxColumn.ReadOnly = True
        Me.FCTTOTALHTDataGridViewTextBoxColumn.Width = 114
        '
        'SOMMELIGNESDataGridViewTextBoxColumn
        '
        Me.SOMMELIGNESDataGridViewTextBoxColumn.DataPropertyName = "SOMME_LIGNES"
        DataGridViewCellStyle12.Format = "C2"
        DataGridViewCellStyle12.NullValue = Nothing
        Me.SOMMELIGNESDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle12
        Me.SOMMELIGNESDataGridViewTextBoxColumn.HeaderText = "SOMME_LIGNES"
        Me.SOMMELIGNESDataGridViewTextBoxColumn.Name = "SOMMELIGNESDataGridViewTextBoxColumn"
        Me.SOMMELIGNESDataGridViewTextBoxColumn.ReadOnly = True
        Me.SOMMELIGNESDataGridViewTextBoxColumn.Width = 117
        '
        'm_bsrcFactComError
        '
        Me.m_bsrcFactComError.DataMember = "FACTCOMERROR"
        Me.m_bsrcFactComError.DataSource = Me.m_dsError
        '
        'm_oTAFactComError
        '
        Me.m_oTAFactComError.ClearBeforeFill = True
        '
        'm_oTACmdErreur
        '
        Me.m_oTACmdErreur.ClearBeforeFill = True
        '
        'm_oTACMD_ERROR3
        '
        Me.m_oTACMD_ERROR3.ClearBeforeFill = True
        '
        'm_oTABASansMVTSTK
        '
        Me.m_oTABASansMVTSTK.ClearBeforeFill = True
        '
        'frmCheckDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(741, 611)
        Me.Controls.Add(Me.DataGridView4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.DataGridView3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.cbExecuter)
        Me.Controls.Add(Me.tbAnnee)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "frmCheckDatabase"
        Me.Text = "frmFactComError"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcBACMDSansMVTSTK, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_dsError, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcError3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcCMDERROR, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcFactComError, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents m_oTAFactComError As vini_DB.dsErrorTableAdapters.FACTCOMERRORTableAdapter
    Friend WithEvents m_dsError As vini_DB.dsError
    Friend WithEvents m_bsrcFactComError As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbAnnee As System.Windows.Forms.TextBox
    Friend WithEvents cbExecuter As System.Windows.Forms.Button
    Friend WithEvents m_oTACmdErreur As vini_DB.dsErrorTableAdapters.COMMANDE_ERRORTableAdapter
    Friend WithEvents m_bsrcCMDERROR As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents CMDCODEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CMDDATEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NBREFRNDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NBRESCMDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcError3 As System.Windows.Forms.BindingSource
    Friend WithEvents m_oTACMD_ERROR3 As vini_DB.dsErrorTableAdapters.COMMANDE_SCMD_NULLTableAdapter
    Friend WithEvents PRDCODEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PRDLIBELLEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FCTCODEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FCTDATEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FCTPERIODEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FRNNOMDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FCTTOTALHTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SOMMELIGNESDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DIFFERENCE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CMDCODEDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CMDDATEDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CLTNOMDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FRNCODEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FRNRSDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents m_bsrcBACMDSansMVTSTK As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView4 As System.Windows.Forms.DataGridView
    Friend WithEvents m_oTABASansMVTSTK As vini_DB.dsErrorTableAdapters.BASansMVTSTKTableAdapter
    Friend WithEvents CODEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CMDDATEDataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NUMDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QTEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
