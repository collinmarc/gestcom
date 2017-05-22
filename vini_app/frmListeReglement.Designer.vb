<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmListeReglement
    Inherits vini_app.frmStatistiques

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
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtDeb = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtFin = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbCodeTiers = New System.Windows.Forms.TextBox
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.cbRechercher = New System.Windows.Forms.Button
        Me.rbFactCommision = New System.Windows.Forms.RadioButton
        Me.rbFactTransport = New System.Windows.Forms.RadioButton
        Me.rbFactColisage = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Date de début :"
        '
        'dtDeb
        '
        Me.dtDeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDeb.Location = New System.Drawing.Point(96, 8)
        Me.dtDeb.Name = "dtDeb"
        Me.dtDeb.Size = New System.Drawing.Size(84, 20)
        Me.dtDeb.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(186, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Date de fin :"
        '
        'dtFin
        '
        Me.dtFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFin.Location = New System.Drawing.Point(257, 8)
        Me.dtFin.Name = "dtFin"
        Me.dtFin.Size = New System.Drawing.Size(82, 20)
        Me.dtFin.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(586, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Code Tiers  :"
        '
        'tbCodeTiers
        '
        Me.tbCodeTiers.Location = New System.Drawing.Point(659, 9)
        Me.tbCodeTiers.Name = "tbCodeTiers"
        Me.tbCodeTiers.Size = New System.Drawing.Size(100, 20)
        Me.tbCodeTiers.TabIndex = 4
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.Location = New System.Drawing.Point(873, 7)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(75, 23)
        Me.cbAfficher.TabIndex = 6
        Me.cbAfficher.Text = "Afficher"
        Me.cbAfficher.UseVisualStyleBackColor = True
        '
        'cbRechercher
        '
        Me.cbRechercher.Location = New System.Drawing.Point(777, 7)
        Me.cbRechercher.Name = "cbRechercher"
        Me.cbRechercher.Size = New System.Drawing.Size(75, 23)
        Me.cbRechercher.TabIndex = 5
        Me.cbRechercher.Text = "Rechercher"
        Me.cbRechercher.UseVisualStyleBackColor = True
        '
        'rbFactCommision
        '
        Me.rbFactCommision.AutoSize = True
        Me.rbFactCommision.Checked = True
        Me.rbFactCommision.Location = New System.Drawing.Point(345, 10)
        Me.rbFactCommision.Name = "rbFactCommision"
        Me.rbFactCommision.Size = New System.Drawing.Size(75, 17)
        Me.rbFactCommision.TabIndex = 7
        Me.rbFactCommision.TabStop = True
        Me.rbFactCommision.Text = "Commision"
        Me.rbFactCommision.UseVisualStyleBackColor = True
        '
        'rbFactTransport
        '
        Me.rbFactTransport.AutoSize = True
        Me.rbFactTransport.Location = New System.Drawing.Point(426, 10)
        Me.rbFactTransport.Name = "rbFactTransport"
        Me.rbFactTransport.Size = New System.Drawing.Size(70, 17)
        Me.rbFactTransport.TabIndex = 8
        Me.rbFactTransport.Text = "Transport"
        Me.rbFactTransport.UseVisualStyleBackColor = True
        '
        'rbFactColisage
        '
        Me.rbFactColisage.AutoSize = True
        Me.rbFactColisage.Location = New System.Drawing.Point(502, 10)
        Me.rbFactColisage.Name = "rbFactColisage"
        Me.rbFactColisage.Size = New System.Drawing.Size(65, 17)
        Me.rbFactColisage.TabIndex = 9
        Me.rbFactColisage.Text = "Colisage"
        Me.rbFactColisage.UseVisualStyleBackColor = True
        '
        'frmListeReglement
        '
        Me.ClientSize = New System.Drawing.Size(960, 342)
        Me.Controls.Add(Me.rbFactCommision)
        Me.Controls.Add(Me.rbFactColisage)
        Me.Controls.Add(Me.rbFactTransport)
        Me.Controls.Add(Me.cbRechercher)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.tbCodeTiers)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtFin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtDeb)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmListeReglement"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Liste des Règlements"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtDeb, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.dtFin, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.tbCodeTiers, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.cbRechercher, 0)
        Me.Controls.SetChildIndex(Me.rbFactTransport, 0)
        Me.Controls.SetChildIndex(Me.rbFactColisage, 0)
        Me.Controls.SetChildIndex(Me.rbFactCommision, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtDeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbCodeTiers As System.Windows.Forms.TextBox
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents cbRechercher As System.Windows.Forms.Button
    Friend WithEvents rbFactCommision As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactTransport As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactColisage As System.Windows.Forms.RadioButton

End Class
