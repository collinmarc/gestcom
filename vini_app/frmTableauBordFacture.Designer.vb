<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTableauBordFacture
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
        Me.ckFactureSoldee = New System.Windows.Forms.CheckBox
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.rbFacturecommission = New System.Windows.Forms.RadioButton
        Me.Label3 = New System.Windows.Forms.Label
        Me.rbFactureTransport = New System.Windows.Forms.RadioButton
        Me.rbFactureColisage = New System.Windows.Forms.RadioButton
        Me.laCodetiers = New System.Windows.Forms.Label
        Me.tbCodeTiers = New System.Windows.Forms.TextBox
        Me.cbRecherche = New System.Windows.Forms.Button
        Me.ckDetail = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Date de début :"
        '
        'dtDeb
        '
        Me.dtDeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDeb.Location = New System.Drawing.Point(100, 8)
        Me.dtDeb.Name = "dtDeb"
        Me.dtDeb.Size = New System.Drawing.Size(88, 20)
        Me.dtDeb.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(194, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Date de fin : "
        '
        'dtFin
        '
        Me.dtFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFin.Location = New System.Drawing.Point(266, 9)
        Me.dtFin.Name = "dtFin"
        Me.dtFin.Size = New System.Drawing.Size(93, 20)
        Me.dtFin.TabIndex = 3
        '
        'ckFactureSoldee
        '
        Me.ckFactureSoldee.AutoSize = True
        Me.ckFactureSoldee.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckFactureSoldee.Location = New System.Drawing.Point(737, 11)
        Me.ckFactureSoldee.Name = "ckFactureSoldee"
        Me.ckFactureSoldee.Size = New System.Drawing.Size(122, 17)
        Me.ckFactureSoldee.TabIndex = 4
        Me.ckFactureSoldee.Text = "Facture non soldées"
        Me.ckFactureSoldee.UseVisualStyleBackColor = True
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.Location = New System.Drawing.Point(865, 8)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(75, 23)
        Me.cbAfficher.TabIndex = 5
        Me.cbAfficher.Text = "Afficher"
        Me.cbAfficher.UseVisualStyleBackColor = True
        '
        'rbFacturecommission
        '
        Me.rbFacturecommission.AutoSize = True
        Me.rbFacturecommission.Checked = True
        Me.rbFacturecommission.Location = New System.Drawing.Point(465, 10)
        Me.rbFacturecommission.Name = "rbFacturecommission"
        Me.rbFacturecommission.Size = New System.Drawing.Size(80, 17)
        Me.rbFacturecommission.TabIndex = 6
        Me.rbFacturecommission.TabStop = True
        Me.rbFacturecommission.Text = "Commission"
        Me.rbFacturecommission.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(365, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Type de Facture : "
        '
        'rbFactureTransport
        '
        Me.rbFactureTransport.AutoSize = True
        Me.rbFactureTransport.Location = New System.Drawing.Point(551, 10)
        Me.rbFactureTransport.Name = "rbFactureTransport"
        Me.rbFactureTransport.Size = New System.Drawing.Size(70, 17)
        Me.rbFactureTransport.TabIndex = 8
        Me.rbFactureTransport.Text = "Transport"
        Me.rbFactureTransport.UseVisualStyleBackColor = True
        '
        'rbFactureColisage
        '
        Me.rbFactureColisage.AutoSize = True
        Me.rbFactureColisage.Location = New System.Drawing.Point(627, 10)
        Me.rbFactureColisage.Name = "rbFactureColisage"
        Me.rbFactureColisage.Size = New System.Drawing.Size(65, 17)
        Me.rbFactureColisage.TabIndex = 9
        Me.rbFactureColisage.Text = "Colisage"
        Me.rbFactureColisage.UseVisualStyleBackColor = True
        '
        'laCodetiers
        '
        Me.laCodetiers.AutoSize = True
        Me.laCodetiers.Location = New System.Drawing.Point(365, 36)
        Me.laCodetiers.Name = "laCodetiers"
        Me.laCodetiers.Size = New System.Drawing.Size(64, 13)
        Me.laCodetiers.TabIndex = 10
        Me.laCodetiers.Text = "Code Tiers :"
        '
        'tbCodeTiers
        '
        Me.tbCodeTiers.Location = New System.Drawing.Point(465, 33)
        Me.tbCodeTiers.Name = "tbCodeTiers"
        Me.tbCodeTiers.Size = New System.Drawing.Size(100, 20)
        Me.tbCodeTiers.TabIndex = 12
        '
        'cbRecherche
        '
        Me.cbRecherche.Location = New System.Drawing.Point(583, 30)
        Me.cbRecherche.Name = "cbRecherche"
        Me.cbRecherche.Size = New System.Drawing.Size(75, 23)
        Me.cbRecherche.TabIndex = 13
        Me.cbRecherche.Text = "Recherche"
        Me.cbRecherche.UseVisualStyleBackColor = True
        '
        'ckDetail
        '
        Me.ckDetail.AutoSize = True
        Me.ckDetail.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckDetail.Location = New System.Drawing.Point(732, 35)
        Me.ckDetail.Name = "ckDetail"
        Me.ckDetail.Size = New System.Drawing.Size(127, 17)
        Me.ckDetail.TabIndex = 14
        Me.ckDetail.Text = "Détail des règlements"
        Me.ckDetail.UseVisualStyleBackColor = True
        '
        'frmTableauBordFacture
        '
        Me.ClientSize = New System.Drawing.Size(952, 646)
        Me.Controls.Add(Me.ckDetail)
        Me.Controls.Add(Me.cbRecherche)
        Me.Controls.Add(Me.tbCodeTiers)
        Me.Controls.Add(Me.laCodetiers)
        Me.Controls.Add(Me.rbFactureColisage)
        Me.Controls.Add(Me.rbFactureTransport)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.rbFacturecommission)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.ckFactureSoldee)
        Me.Controls.Add(Me.dtFin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtDeb)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmTableauBordFacture"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtDeb, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.dtFin, 0)
        Me.Controls.SetChildIndex(Me.ckFactureSoldee, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.rbFacturecommission, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.rbFactureTransport, 0)
        Me.Controls.SetChildIndex(Me.rbFactureColisage, 0)
        Me.Controls.SetChildIndex(Me.laCodetiers, 0)
        Me.Controls.SetChildIndex(Me.tbCodeTiers, 0)
        Me.Controls.SetChildIndex(Me.cbRecherche, 0)
        Me.Controls.SetChildIndex(Me.ckDetail, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtDeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents ckFactureSoldee As System.Windows.Forms.CheckBox
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents rbFacturecommission As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents rbFactureTransport As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactureColisage As System.Windows.Forms.RadioButton
    Friend WithEvents laCodetiers As System.Windows.Forms.Label
    Friend WithEvents tbCodeTiers As System.Windows.Forms.TextBox
    Friend WithEvents cbRecherche As System.Windows.Forms.Button
    Friend WithEvents ckDetail As System.Windows.Forms.CheckBox

End Class
