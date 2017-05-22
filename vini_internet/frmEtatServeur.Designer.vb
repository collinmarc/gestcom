<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEtatServeur
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.laEtatServeur = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.laProtection = New System.Windows.Forms.Label()
        Me.cbUnlock = New System.Windows.Forms.Button()
        Me.cbLockFrom = New System.Windows.Forms.Button()
        Me.cbLockTo = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.laUrl = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.laUtilisateur = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.laPassword = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 100)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(26, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Etat"
        '
        'laEtatServeur
        '
        Me.laEtatServeur.AutoSize = True
        Me.laEtatServeur.ForeColor = System.Drawing.Color.Green
        Me.laEtatServeur.Location = New System.Drawing.Point(94, 100)
        Me.laEtatServeur.Name = "laEtatServeur"
        Me.laEtatServeur.Size = New System.Drawing.Size(63, 13)
        Me.laEtatServeur.TabIndex = 1
        Me.laEtatServeur.Text = "EtatServeur"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Protection"
        '
        'laProtection
        '
        Me.laProtection.AutoSize = True
        Me.laProtection.ForeColor = System.Drawing.Color.Green
        Me.laProtection.Location = New System.Drawing.Point(94, 128)
        Me.laProtection.Name = "laProtection"
        Me.laProtection.Size = New System.Drawing.Size(30, 13)
        Me.laProtection.TabIndex = 4
        Me.laProtection.Text = "Libre"
        '
        'cbUnlock
        '
        Me.cbUnlock.Location = New System.Drawing.Point(47, 195)
        Me.cbUnlock.Name = "cbUnlock"
        Me.cbUnlock.Size = New System.Drawing.Size(61, 27)
        Me.cbUnlock.TabIndex = 5
        Me.cbUnlock.Text = "Libérer"
        Me.cbUnlock.UseVisualStyleBackColor = True
        '
        'cbLockFrom
        '
        Me.cbLockFrom.Location = New System.Drawing.Point(125, 195)
        Me.cbLockFrom.Name = "cbLockFrom"
        Me.cbLockFrom.Size = New System.Drawing.Size(106, 27)
        Me.cbLockFrom.TabIndex = 6
        Me.cbLockFrom.Text = "Vérrouiller From"
        Me.cbLockFrom.UseVisualStyleBackColor = True
        '
        'cbLockTo
        '
        Me.cbLockTo.Location = New System.Drawing.Point(237, 198)
        Me.cbLockTo.Name = "cbLockTo"
        Me.cbLockTo.Size = New System.Drawing.Size(107, 24)
        Me.cbLockTo.TabIndex = 7
        Me.cbLockTo.Text = "Vérouiller To"
        Me.cbLockTo.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Url :"
        '
        'laUrl
        '
        Me.laUrl.AutoSize = True
        Me.laUrl.ForeColor = System.Drawing.Color.Green
        Me.laUrl.Location = New System.Drawing.Point(94, 9)
        Me.laUrl.Name = "laUrl"
        Me.laUrl.Size = New System.Drawing.Size(44, 13)
        Me.laUrl.TabIndex = 9
        Me.laUrl.Text = "Ftp://..."
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Utilisateur :"
        '
        'laUtilisateur
        '
        Me.laUtilisateur.AutoSize = True
        Me.laUtilisateur.ForeColor = System.Drawing.Color.Green
        Me.laUtilisateur.Location = New System.Drawing.Point(94, 34)
        Me.laUtilisateur.Name = "laUtilisateur"
        Me.laUtilisateur.Size = New System.Drawing.Size(61, 13)
        Me.laUtilisateur.TabIndex = 11
        Me.laUtilisateur.Text = "laUtilisateur"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 63)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(71, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Mot de passe"
        '
        'laPassword
        '
        Me.laPassword.AutoSize = True
        Me.laPassword.ForeColor = System.Drawing.Color.Green
        Me.laPassword.Location = New System.Drawing.Point(94, 64)
        Me.laPassword.Name = "laPassword"
        Me.laPassword.Size = New System.Drawing.Size(61, 13)
        Me.laPassword.TabIndex = 13
        Me.laPassword.Text = "laPassword"
        '
        'frmEtatServeur
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(377, 241)
        Me.Controls.Add(Me.laPassword)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.laUtilisateur)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.laUrl)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbLockTo)
        Me.Controls.Add(Me.cbLockFrom)
        Me.Controls.Add(Me.cbUnlock)
        Me.Controls.Add(Me.laProtection)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.laEtatServeur)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmEtatServeur"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Etat du serveur FTP"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents laEtatServeur As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents laProtection As System.Windows.Forms.Label
    Friend WithEvents cbUnlock As System.Windows.Forms.Button
    Friend WithEvents cbLockFrom As System.Windows.Forms.Button
    Friend WithEvents cbLockTo As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents laUrl As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents laUtilisateur As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents laPassword As System.Windows.Forms.Label
End Class
