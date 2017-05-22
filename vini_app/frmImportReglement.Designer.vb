<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportReglement
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
        Me.cbImporter = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.CONSTANTESTableAdapter = New vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.CONSTANTESBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DsVinicom = New vini_DB.dsVinicom
        Me.lbMessages = New System.Windows.Forms.ListBox
        Me.cbParcourir = New System.Windows.Forms.Button
        Me.tbPath = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.rbFactCom = New System.Windows.Forms.RadioButton
        Me.rbFactTRP = New System.Windows.Forms.RadioButton
        Me.rbFactColisage = New System.Windows.Forms.RadioButton
        Me.cbVerifier = New System.Windows.Forms.Button
        CType(Me.CONSTANTESBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DsVinicom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbImporter
        '
        Me.cbImporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbImporter.Location = New System.Drawing.Point(867, 126)
        Me.cbImporter.Name = "cbImporter"
        Me.cbImporter.Size = New System.Drawing.Size(120, 23)
        Me.cbImporter.TabIndex = 18
        Me.cbImporter.Text = "Importer"
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label1.Location = New System.Drawing.Point(3, -102)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 20)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Date de début :"
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label2.Location = New System.Drawing.Point(3, -76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 20)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Date de fin :"
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
        Me.lbMessages.Location = New System.Drawing.Point(12, 155)
        Me.lbMessages.Name = "lbMessages"
        Me.lbMessages.Size = New System.Drawing.Size(975, 355)
        Me.lbMessages.TabIndex = 22
        '
        'cbParcourir
        '
        Me.cbParcourir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbParcourir.Location = New System.Drawing.Point(866, 69)
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
        Me.tbPath.Location = New System.Drawing.Point(62, 71)
        Me.tbPath.Name = "tbPath"
        Me.tbPath.Size = New System.Drawing.Size(768, 20)
        Me.tbPath.TabIndex = 23
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 74)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 13)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Fichier :"
        '
        'rbFactCom
        '
        Me.rbFactCom.AutoSize = True
        Me.rbFactCom.Checked = True
        Me.rbFactCom.Location = New System.Drawing.Point(12, 2)
        Me.rbFactCom.Name = "rbFactCom"
        Me.rbFactCom.Size = New System.Drawing.Size(209, 17)
        Me.rbFactCom.TabIndex = 28
        Me.rbFactCom.TabStop = True
        Me.rbFactCom.Text = "Reglement des factures de commission"
        Me.rbFactCom.UseVisualStyleBackColor = True
        '
        'rbFactTRP
        '
        Me.rbFactTRP.AutoSize = True
        Me.rbFactTRP.Location = New System.Drawing.Point(12, 25)
        Me.rbFactTRP.Name = "rbFactTRP"
        Me.rbFactTRP.Size = New System.Drawing.Size(196, 17)
        Me.rbFactTRP.TabIndex = 29
        Me.rbFactTRP.Text = "Reglement des factures de transport"
        Me.rbFactTRP.UseVisualStyleBackColor = True
        '
        'rbFactColisage
        '
        Me.rbFactColisage.AutoSize = True
        Me.rbFactColisage.Location = New System.Drawing.Point(12, 48)
        Me.rbFactColisage.Name = "rbFactColisage"
        Me.rbFactColisage.Size = New System.Drawing.Size(194, 17)
        Me.rbFactColisage.TabIndex = 30
        Me.rbFactColisage.Text = "Reglement des factures de colisage"
        Me.rbFactColisage.UseVisualStyleBackColor = True
        '
        'cbVerifier
        '
        Me.cbVerifier.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbVerifier.Location = New System.Drawing.Point(866, 98)
        Me.cbVerifier.Name = "cbVerifier"
        Me.cbVerifier.Size = New System.Drawing.Size(121, 23)
        Me.cbVerifier.TabIndex = 31
        Me.cbVerifier.Text = "Vérifier"
        Me.cbVerifier.UseVisualStyleBackColor = True
        '
        'frmImportReglement
        '
        Me.ClientSize = New System.Drawing.Size(995, 525)
        Me.Controls.Add(Me.cbVerifier)
        Me.Controls.Add(Me.rbFactColisage)
        Me.Controls.Add(Me.rbFactTRP)
        Me.Controls.Add(Me.rbFactCom)
        Me.Controls.Add(Me.cbImporter)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lbMessages)
        Me.Controls.Add(Me.cbParcourir)
        Me.Controls.Add(Me.tbPath)
        Me.Controls.Add(Me.Label3)
        Me.Name = "frmImportReglement"
        Me.Text = "Importation des réglements de facture depuis Quadra"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.CONSTANTESBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DsVinicom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbImporter As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CONSTANTESTableAdapter As vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents CONSTANTESBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DsVinicom As vini_DB.dsVinicom
    Friend WithEvents lbMessages As System.Windows.Forms.ListBox
    Friend WithEvents cbParcourir As System.Windows.Forms.Button
    Friend WithEvents tbPath As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents rbFactCom As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactTRP As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactColisage As System.Windows.Forms.RadioButton
    Friend WithEvents cbVerifier As System.Windows.Forms.Button

End Class
