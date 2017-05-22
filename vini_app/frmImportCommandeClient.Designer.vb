<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportcommandeClient
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
        Me.components = New System.ComponentModel.Container()
        Me.cbImporter = New System.Windows.Forms.Button()
        Me.CONSTANTESTableAdapter = New vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.CONSTANTESBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DsVinicom = New vini_DB.dsVinicom()
        Me.lbMessages = New System.Windows.Forms.ListBox()
        Me.cbParcourir = New System.Windows.Forms.Button()
        Me.tbPath = New System.Windows.Forms.TextBox()
        Me.m_rbRepertoire = New System.Windows.Forms.RadioButton()
        Me.m_rbFichier = New System.Windows.Forms.RadioButton()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.CONSTANTESBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DsVinicom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbImporter
        '
        Me.cbImporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbImporter.Location = New System.Drawing.Point(608, 66)
        Me.cbImporter.Name = "cbImporter"
        Me.cbImporter.Size = New System.Drawing.Size(120, 23)
        Me.cbImporter.TabIndex = 18
        Me.cbImporter.Text = "Importer"
        '
        'CONSTANTESTableAdapter
        '
        Me.CONSTANTESTableAdapter.ClearBeforeFill = True
        '
        'CONSTANTESBindingSource
        '
        Me.CONSTANTESBindingSource.DataMember = "CONSTANTES"
        Me.CONSTANTESBindingSource.DataSource = Me.DsVinicom
        '
        'DsVinicom
        '
        Me.DsVinicom.DataSetName = "dsVinicom"
        Me.DsVinicom.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'lbMessages
        '
        Me.lbMessages.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbMessages.FormattingEnabled = True
        Me.lbMessages.Location = New System.Drawing.Point(12, 95)
        Me.lbMessages.Name = "lbMessages"
        Me.lbMessages.Size = New System.Drawing.Size(716, 121)
        Me.lbMessages.TabIndex = 22
        '
        'cbParcourir
        '
        Me.cbParcourir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbParcourir.Location = New System.Drawing.Point(607, 38)
        Me.cbParcourir.Name = "cbParcourir"
        Me.cbParcourir.Size = New System.Drawing.Size(121, 22)
        Me.cbParcourir.TabIndex = 25
        Me.cbParcourir.Text = "Parcourir"
        Me.cbParcourir.UseVisualStyleBackColor = True
        '
        'tbPath
        '
        Me.tbPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPath.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.CONSTANTESBindingSource, "CST_EXPORT_COMPTA_PATH", True))
        Me.tbPath.Location = New System.Drawing.Point(91, 38)
        Me.tbPath.Name = "tbPath"
        Me.tbPath.Size = New System.Drawing.Size(510, 20)
        Me.tbPath.TabIndex = 23
        '
        'm_rbRepertoire
        '
        Me.m_rbRepertoire.AutoSize = True
        Me.m_rbRepertoire.Checked = True
        Me.m_rbRepertoire.Location = New System.Drawing.Point(12, 12)
        Me.m_rbRepertoire.Name = "m_rbRepertoire"
        Me.m_rbRepertoire.Size = New System.Drawing.Size(74, 17)
        Me.m_rbRepertoire.TabIndex = 26
        Me.m_rbRepertoire.TabStop = True
        Me.m_rbRepertoire.Text = "Répertoire"
        Me.m_rbRepertoire.UseVisualStyleBackColor = True
        '
        'm_rbFichier
        '
        Me.m_rbFichier.AutoSize = True
        Me.m_rbFichier.Location = New System.Drawing.Point(12, 38)
        Me.m_rbFichier.Name = "m_rbFichier"
        Me.m_rbFichier.Size = New System.Drawing.Size(56, 17)
        Me.m_rbFichier.TabIndex = 27
        Me.m_rbFichier.Text = "Fichier"
        Me.m_rbFichier.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmImportcommandeClient
        '
        Me.ClientSize = New System.Drawing.Size(732, 252)
        Me.Controls.Add(Me.m_rbFichier)
        Me.Controls.Add(Me.m_rbRepertoire)
        Me.Controls.Add(Me.cbImporter)
        Me.Controls.Add(Me.lbMessages)
        Me.Controls.Add(Me.cbParcourir)
        Me.Controls.Add(Me.tbPath)
        Me.Name = "frmImportcommandeClient"
        Me.Text = "Import des commandes clients depuis les fichiers transmis par le site internet"
        CType(Me.CONSTANTESBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DsVinicom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbImporter As System.Windows.Forms.Button
    Friend WithEvents CONSTANTESTableAdapter As vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents CONSTANTESBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DsVinicom As vini_DB.dsVinicom
    Friend WithEvents lbMessages As System.Windows.Forms.ListBox
    Friend WithEvents cbParcourir As System.Windows.Forms.Button
    Friend WithEvents tbPath As System.Windows.Forms.TextBox
    Friend WithEvents m_rbRepertoire As System.Windows.Forms.RadioButton
    Friend WithEvents m_rbFichier As System.Windows.Forms.RadioButton
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog

End Class
