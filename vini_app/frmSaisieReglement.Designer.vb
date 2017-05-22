<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSaisieReglement
    Inherits FrmVinicom

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
        Me.dtDateReglement = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbNumFactCom = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbReferenceReglement = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cbAppliquer = New System.Windows.Forms.Button
        Me.cbCancel = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbSolde = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.rbFactCommission = New System.Windows.Forms.RadioButton
        Me.rbFactColisage = New System.Windows.Forms.RadioButton
        Me.rbFactureTransport = New System.Windows.Forms.RadioButton
        Me.cbRechercheFacture = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbCommentaire = New System.Windows.Forms.TextBox
        Me.laTiers = New System.Windows.Forms.Label
        Me.laMontantFacture = New System.Windows.Forms.Label
        Me.tbMontant = New vini_app.textBoxCurrency
        Me.SuspendLayout()
        '
        'dtDateReglement
        '
        Me.dtDateReglement.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateReglement.Location = New System.Drawing.Point(175, 12)
        Me.dtDateReglement.Name = "dtDateReglement"
        Me.dtDateReglement.Size = New System.Drawing.Size(95, 20)
        Me.dtDateReglement.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Date du réglement : "
        '
        'tbNumFactCom
        '
        Me.tbNumFactCom.Location = New System.Drawing.Point(175, 103)
        Me.tbNumFactCom.Name = "tbNumFactCom"
        Me.tbNumFactCom.Size = New System.Drawing.Size(100, 20)
        Me.tbNumFactCom.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 182)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Montant Réglement"
        '
        'tbReferenceReglement
        '
        Me.tbReferenceReglement.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbReferenceReglement.Location = New System.Drawing.Point(175, 217)
        Me.tbReferenceReglement.Name = "tbReferenceReglement"
        Me.tbReferenceReglement.Size = New System.Drawing.Size(411, 20)
        Me.tbReferenceReglement.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 220)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(111, 13)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Reference Reglement"
        '
        'cbAppliquer
        '
        Me.cbAppliquer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAppliquer.Location = New System.Drawing.Point(371, 303)
        Me.cbAppliquer.Name = "cbAppliquer"
        Me.cbAppliquer.Size = New System.Drawing.Size(113, 23)
        Me.cbAppliquer.TabIndex = 9
        Me.cbAppliquer.Text = "Appliquer + nouveau"
        Me.cbAppliquer.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbAppliquer.UseVisualStyleBackColor = True
        '
        'cbCancel
        '
        Me.cbCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cbCancel.Location = New System.Drawing.Point(490, 303)
        Me.cbCancel.Name = "cbCancel"
        Me.cbCancel.Size = New System.Drawing.Size(96, 23)
        Me.cbCancel.TabIndex = 10
        Me.cbCancel.Text = "Quitter"
        Me.cbCancel.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(297, 182)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Solde : "
        '
        'tbSolde
        '
        Me.tbSolde.Enabled = False
        Me.tbSolde.ForeColor = System.Drawing.Color.Green
        Me.tbSolde.Location = New System.Drawing.Point(346, 180)
        Me.tbSolde.Name = "tbSolde"
        Me.tbSolde.Size = New System.Drawing.Size(100, 20)
        Me.tbSolde.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(172, 79)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(58, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "N° Facture"
        '
        'rbFactCommission
        '
        Me.rbFactCommission.AutoSize = True
        Me.rbFactCommission.Checked = True
        Me.rbFactCommission.Location = New System.Drawing.Point(12, 79)
        Me.rbFactCommission.Name = "rbFactCommission"
        Me.rbFactCommission.Size = New System.Drawing.Size(134, 17)
        Me.rbFactCommission.TabIndex = 1
        Me.rbFactCommission.TabStop = True
        Me.rbFactCommission.Text = "Facture de Commission"
        Me.rbFactCommission.UseVisualStyleBackColor = True
        '
        'rbFactColisage
        '
        Me.rbFactColisage.AutoSize = True
        Me.rbFactColisage.Location = New System.Drawing.Point(12, 125)
        Me.rbFactColisage.Name = "rbFactColisage"
        Me.rbFactColisage.Size = New System.Drawing.Size(119, 17)
        Me.rbFactColisage.TabIndex = 3
        Me.rbFactColisage.Text = "Facture de Colisage"
        Me.rbFactColisage.UseVisualStyleBackColor = True
        '
        'rbFactureTransport
        '
        Me.rbFactureTransport.AutoSize = True
        Me.rbFactureTransport.Location = New System.Drawing.Point(12, 102)
        Me.rbFactureTransport.Name = "rbFactureTransport"
        Me.rbFactureTransport.Size = New System.Drawing.Size(124, 17)
        Me.rbFactureTransport.TabIndex = 2
        Me.rbFactureTransport.Text = "Facture de Transport"
        Me.rbFactureTransport.UseVisualStyleBackColor = True
        '
        'cbRechercheFacture
        '
        Me.cbRechercheFacture.Location = New System.Drawing.Point(281, 100)
        Me.cbRechercheFacture.Name = "cbRechercheFacture"
        Me.cbRechercheFacture.Size = New System.Drawing.Size(75, 23)
        Me.cbRechercheFacture.TabIndex = 5
        Me.cbRechercheFacture.Text = "Rechercher"
        Me.cbRechercheFacture.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 249)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Commentaires"
        '
        'tbCommentaire
        '
        Me.tbCommentaire.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbCommentaire.Location = New System.Drawing.Point(175, 249)
        Me.tbCommentaire.Multiline = True
        Me.tbCommentaire.Name = "tbCommentaire"
        Me.tbCommentaire.Size = New System.Drawing.Size(411, 43)
        Me.tbCommentaire.TabIndex = 8
        '
        'laTiers
        '
        Me.laTiers.AutoSize = True
        Me.laTiers.ForeColor = System.Drawing.Color.Green
        Me.laTiers.Location = New System.Drawing.Point(175, 130)
        Me.laTiers.Name = "laTiers"
        Me.laTiers.Size = New System.Drawing.Size(66, 13)
        Me.laTiers.TabIndex = 22
        Me.laTiers.Text = "Tiers facturé"
        '
        'laMontantFacture
        '
        Me.laMontantFacture.AutoSize = True
        Me.laMontantFacture.ForeColor = System.Drawing.Color.Green
        Me.laMontantFacture.Location = New System.Drawing.Point(178, 147)
        Me.laMontantFacture.Name = "laMontantFacture"
        Me.laMontantFacture.Size = New System.Drawing.Size(82, 13)
        Me.laMontantFacture.TabIndex = 23
        Me.laMontantFacture.Text = "Montant facturé"
        '
        'tbMontant
        '
        Me.tbMontant.Location = New System.Drawing.Point(175, 179)
        Me.tbMontant.Name = "tbMontant"
        Me.tbMontant.Size = New System.Drawing.Size(100, 20)
        Me.tbMontant.TabIndex = 6
        Me.tbMontant.Text = "0"
        '
        'frmSaisieReglement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cbCancel
        Me.ClientSize = New System.Drawing.Size(593, 338)
        Me.Controls.Add(Me.tbMontant)
        Me.Controls.Add(Me.laMontantFacture)
        Me.Controls.Add(Me.laTiers)
        Me.Controls.Add(Me.tbCommentaire)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbRechercheFacture)
        Me.Controls.Add(Me.rbFactureTransport)
        Me.Controls.Add(Me.rbFactColisage)
        Me.Controls.Add(Me.rbFactCommission)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.tbSolde)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cbCancel)
        Me.Controls.Add(Me.cbAppliquer)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbReferenceReglement)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbNumFactCom)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtDateReglement)
        Me.Name = "frmSaisieReglement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Saisie des Règlements"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtDateReglement As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbNumFactCom As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbReferenceReglement As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbAppliquer As System.Windows.Forms.Button
    Friend WithEvents cbCancel As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbSolde As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents rbFactCommission As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactColisage As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactureTransport As System.Windows.Forms.RadioButton
    Friend WithEvents cbRechercheFacture As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbCommentaire As System.Windows.Forms.TextBox
    Friend WithEvents laTiers As System.Windows.Forms.Label
    Friend WithEvents laMontantFacture As System.Windows.Forms.Label
    Friend WithEvents tbMontant As vini_app.textBoxCurrency
End Class
