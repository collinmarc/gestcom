Imports vini_DB

Public Class frmuser
    Inherits FrmVinicom


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
    Friend WithEvents cbValider As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.cbValider = New System.Windows.Forms.Button()
        Me.m_bsrcUser = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_dsVinicom = New vini_DB.dsVinicom()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.USRIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.USRCODEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.USRPASSWORDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.USRROLEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.m_bsrcUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_dsVinicom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbValider
        '
        Me.cbValider.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbValider.Location = New System.Drawing.Point(503, 225)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.Size = New System.Drawing.Size(88, 24)
        Me.cbValider.TabIndex = 6
        Me.cbValider.Text = "&Valider"
        '
        'm_bsrcUser
        '
        Me.m_bsrcUser.DataMember = "USERS"
        Me.m_bsrcUser.DataSource = Me.m_dsVinicom
        '
        'm_dsVinicom
        '
        Me.m_dsVinicom.DataSetName = "dsVinicom"
        Me.m_dsVinicom.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.USRIDDataGridViewTextBoxColumn, Me.USRCODEDataGridViewTextBoxColumn, Me.USRPASSWORDDataGridViewTextBoxColumn, Me.USRROLEDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcUser
        Me.DataGridView1.Location = New System.Drawing.Point(13, 13)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(604, 196)
        Me.DataGridView1.TabIndex = 7
        '
        'USRIDDataGridViewTextBoxColumn
        '
        Me.USRIDDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.USRIDDataGridViewTextBoxColumn.DataPropertyName = "USR_ID"
        Me.USRIDDataGridViewTextBoxColumn.Frozen = True
        Me.USRIDDataGridViewTextBoxColumn.HeaderText = "USR_ID"
        Me.USRIDDataGridViewTextBoxColumn.Name = "USRIDDataGridViewTextBoxColumn"
        Me.USRIDDataGridViewTextBoxColumn.ReadOnly = True
        Me.USRIDDataGridViewTextBoxColumn.Width = 137
        '
        'USRCODEDataGridViewTextBoxColumn
        '
        Me.USRCODEDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.USRCODEDataGridViewTextBoxColumn.DataPropertyName = "USR_CODE"
        Me.USRCODEDataGridViewTextBoxColumn.Frozen = True
        Me.USRCODEDataGridViewTextBoxColumn.HeaderText = "USR_CODE"
        Me.USRCODEDataGridViewTextBoxColumn.Name = "USRCODEDataGridViewTextBoxColumn"
        Me.USRCODEDataGridViewTextBoxColumn.Width = 137
        '
        'USRPASSWORDDataGridViewTextBoxColumn
        '
        Me.USRPASSWORDDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.USRPASSWORDDataGridViewTextBoxColumn.DataPropertyName = "USR_PASSWORD"
        Me.USRPASSWORDDataGridViewTextBoxColumn.Frozen = True
        Me.USRPASSWORDDataGridViewTextBoxColumn.HeaderText = "USR_PASSWORD"
        Me.USRPASSWORDDataGridViewTextBoxColumn.Name = "USRPASSWORDDataGridViewTextBoxColumn"
        Me.USRPASSWORDDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.USRPASSWORDDataGridViewTextBoxColumn.Width = 137
        '
        'USRROLEDataGridViewTextBoxColumn
        '
        Me.USRROLEDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.USRROLEDataGridViewTextBoxColumn.DataPropertyName = "USR_ROLE"
        Me.USRROLEDataGridViewTextBoxColumn.HeaderText = "USR_ROLE"
        Me.USRROLEDataGridViewTextBoxColumn.Name = "USRROLEDataGridViewTextBoxColumn"
        Me.USRROLEDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.USRROLEDataGridViewTextBoxColumn.Width = 137
        '
        'frmuser
        '
        Me.ClientSize = New System.Drawing.Size(796, 526)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.cbValider)
        Me.Name = "frmuser"
        Me.Text = "Changement du code utilisateur"
        CType(Me.m_bsrcUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_dsVinicom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents m_bsrcUser As System.Windows.Forms.BindingSource
    Friend WithEvents m_dsVinicom As vini_DB.dsVinicom
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView

#End Region

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        DataGridView1.Enabled = True
    End Sub
    Private Sub frmuser_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim oTa As vini_DB.dsVinicomTableAdapters.USERSTableAdapter

        oTa = New vini_DB.dsVinicomTableAdapters.USERSTableAdapter()
        oTa.Connection = Persist.oleDBConnection
        oTa.Fill(m_dsVinicom.USERS)
    End Sub

    Private Sub cbValider_Click(sender As System.Object, e As System.EventArgs) Handles cbValider.Click
        Dim oTa As vini_DB.dsVinicomTableAdapters.USERSTableAdapter

        oTa = New vini_DB.dsVinicomTableAdapters.USERSTableAdapter()
        oTa.Connection = Persist.oleDBConnection
        oTa.Update(m_dsVinicom.USERS)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Friend WithEvents USRIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents USRCODEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents USRPASSWORDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents USRROLEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
