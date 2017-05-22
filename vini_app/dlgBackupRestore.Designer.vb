<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgBackupRestore
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
        Me.cbBackup = New System.Windows.Forms.Button()
        Me.tbPath = New System.Windows.Forms.TextBox()
        Me.cbParcourirBackup = New System.Windows.Forms.Button()
        Me.cbRestore = New System.Windows.Forms.Button()
        Me.tbPathRestore = New System.Windows.Forms.TextBox()
        Me.cbParcourirRestore = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.m_dlgOpenFile = New System.Windows.Forms.OpenFileDialog()
        Me.m_dlgSaveFile = New System.Windows.Forms.SaveFileDialog()
        Me.SuspendLayout()
        '
        'cbBackup
        '
        Me.cbBackup.Location = New System.Drawing.Point(607, 71)
        Me.cbBackup.Name = "cbBackup"
        Me.cbBackup.Size = New System.Drawing.Size(89, 23)
        Me.cbBackup.TabIndex = 1
        Me.cbBackup.Text = "Sauvegarde"
        Me.cbBackup.UseVisualStyleBackColor = True
        '
        'tbPath
        '
        Me.tbPath.Location = New System.Drawing.Point(12, 73)
        Me.tbPath.Name = "tbPath"
        Me.tbPath.Size = New System.Drawing.Size(503, 20)
        Me.tbPath.TabIndex = 2
        '
        'cbParcourirBackup
        '
        Me.cbParcourirBackup.Location = New System.Drawing.Point(521, 73)
        Me.cbParcourirBackup.Name = "cbParcourirBackup"
        Me.cbParcourirBackup.Size = New System.Drawing.Size(75, 23)
        Me.cbParcourirBackup.TabIndex = 3
        Me.cbParcourirBackup.Text = "Parcourir"
        Me.cbParcourirBackup.UseVisualStyleBackColor = True
        '
        'cbRestore
        '
        Me.cbRestore.Enabled = False
        Me.cbRestore.Location = New System.Drawing.Point(607, 127)
        Me.cbRestore.Name = "cbRestore"
        Me.cbRestore.Size = New System.Drawing.Size(89, 23)
        Me.cbRestore.TabIndex = 4
        Me.cbRestore.Text = "Restauration"
        Me.cbRestore.UseVisualStyleBackColor = True
        '
        'tbPathRestore
        '
        Me.tbPathRestore.Location = New System.Drawing.Point(12, 127)
        Me.tbPathRestore.Name = "tbPathRestore"
        Me.tbPathRestore.Size = New System.Drawing.Size(503, 20)
        Me.tbPathRestore.TabIndex = 5
        '
        'cbParcourirRestore
        '
        Me.cbParcourirRestore.Location = New System.Drawing.Point(521, 125)
        Me.cbParcourirRestore.Name = "cbParcourirRestore"
        Me.cbParcourirRestore.Size = New System.Drawing.Size(75, 23)
        Me.cbParcourirRestore.TabIndex = 6
        Me.cbParcourirRestore.Text = "Parcourir"
        Me.cbParcourirRestore.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(108, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(420, 26)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Cette fenêtre vous permet de sauvegarder et de restaurer la base de données . " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "A" & _
    "ssurez-vous qu'aucun utilisateur n'est connecté au moment d'effectuer ces opérat" & _
    "ions."
        '
        'm_dlgOpenFile
        '
        Me.m_dlgOpenFile.DefaultExt = "*.bak"
        Me.m_dlgOpenFile.FileName = "vnc4.bak"
        Me.m_dlgOpenFile.Filter = " ""Fichiers Backup (*.bak)|*.bak|Tous les fichiers (*.*)|*.*"""
        Me.m_dlgOpenFile.RestoreDirectory = True
        Me.m_dlgOpenFile.Title = "Fichier de Sauvegarde de base de données"
        '
        'm_dlgSaveFile
        '
        Me.m_dlgSaveFile.DefaultExt = "*.bak"
        Me.m_dlgSaveFile.FileName = "vnc4.bak"
        Me.m_dlgSaveFile.Filter = " ""Fichiers Backup (*.bak)|*.bak|Tous les fichiers (*.*)|*.*"""
        Me.m_dlgSaveFile.Title = "Fichier de sauvegarde de base de données"
        '
        'dlgBackupRestore
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(702, 187)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbParcourirRestore)
        Me.Controls.Add(Me.tbPathRestore)
        Me.Controls.Add(Me.cbRestore)
        Me.Controls.Add(Me.cbParcourirBackup)
        Me.Controls.Add(Me.tbPath)
        Me.Controls.Add(Me.cbBackup)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgBackupRestore"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Sauvegarde / Restauration de la base de donnée"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbBackup As System.Windows.Forms.Button
    Friend WithEvents tbPath As System.Windows.Forms.TextBox
    Friend WithEvents cbParcourirBackup As System.Windows.Forms.Button
    Friend WithEvents cbRestore As System.Windows.Forms.Button
    Friend WithEvents tbPathRestore As System.Windows.Forms.TextBox
    Friend WithEvents cbParcourirRestore As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents m_dlgOpenFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents m_dlgSaveFile As System.Windows.Forms.SaveFileDialog

End Class
