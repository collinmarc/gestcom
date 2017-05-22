<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditfactures
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
        Me.rbFactcom = New System.Windows.Forms.RadioButton()
        Me.rbFactTRP = New System.Windows.Forms.RadioButton()
        Me.rbFactCol = New System.Windows.Forms.RadioButton()
        Me.dtDateDeb = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtDateFin = New System.Windows.Forms.DateTimePicker()
        Me.rbCourier = New System.Windows.Forms.RadioButton()
        Me.rbFacture = New System.Windows.Forms.RadioButton()
        Me.rbReleve = New System.Windows.Forms.RadioButton()
        Me.ckEntete = New System.Windows.Forms.CheckBox()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.laTiers = New System.Windows.Forms.Label()
        Me.tbCodeTiers = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.grpTypeEditFactcom = New System.Windows.Forms.GroupBox()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.GroupBox1.SuspendLayout()
        Me.grpTypeEditFactcom.SuspendLayout()
        Me.SuspendLayout()
        '
        'rbFactcom
        '
        Me.rbFactcom.AutoSize = True
        Me.rbFactcom.Location = New System.Drawing.Point(6, 16)
        Me.rbFactcom.Name = "rbFactcom"
        Me.rbFactcom.Size = New System.Drawing.Size(133, 17)
        Me.rbFactcom.TabIndex = 0
        Me.rbFactcom.TabStop = True
        Me.rbFactcom.Text = "Factures de commision"
        Me.rbFactcom.UseVisualStyleBackColor = True
        '
        'rbFactTRP
        '
        Me.rbFactTRP.AutoSize = True
        Me.rbFactTRP.Location = New System.Drawing.Point(6, 39)
        Me.rbFactTRP.Name = "rbFactTRP"
        Me.rbFactTRP.Size = New System.Drawing.Size(125, 17)
        Me.rbFactTRP.TabIndex = 1
        Me.rbFactTRP.TabStop = True
        Me.rbFactTRP.Text = "Factures de transport"
        Me.rbFactTRP.UseVisualStyleBackColor = True
        '
        'rbFactCol
        '
        Me.rbFactCol.AutoSize = True
        Me.rbFactCol.Location = New System.Drawing.Point(6, 62)
        Me.rbFactCol.Name = "rbFactCol"
        Me.rbFactCol.Size = New System.Drawing.Size(123, 17)
        Me.rbFactCol.TabIndex = 2
        Me.rbFactCol.TabStop = True
        Me.rbFactCol.Text = "Factures de colisage"
        Me.rbFactCol.UseVisualStyleBackColor = True
        '
        'dtDateDeb
        '
        Me.dtDateDeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateDeb.Location = New System.Drawing.Point(297, 11)
        Me.dtDateDeb.Name = "dtDateDeb"
        Me.dtDateDeb.Size = New System.Drawing.Size(88, 20)
        Me.dtDateDeb.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(206, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Date de début :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(206, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Date de fin :"
        '
        'dtDateFin
        '
        Me.dtDateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFin.Location = New System.Drawing.Point(297, 37)
        Me.dtDateFin.Name = "dtDateFin"
        Me.dtDateFin.Size = New System.Drawing.Size(88, 20)
        Me.dtDateFin.TabIndex = 6
        '
        'rbCourier
        '
        Me.rbCourier.AutoSize = True
        Me.rbCourier.Location = New System.Drawing.Point(6, 16)
        Me.rbCourier.Name = "rbCourier"
        Me.rbCourier.Size = New System.Drawing.Size(61, 17)
        Me.rbCourier.TabIndex = 7
        Me.rbCourier.TabStop = True
        Me.rbCourier.Text = "Courrier"
        Me.rbCourier.UseVisualStyleBackColor = True
        '
        'rbFacture
        '
        Me.rbFacture.AutoSize = True
        Me.rbFacture.Location = New System.Drawing.Point(6, 39)
        Me.rbFacture.Name = "rbFacture"
        Me.rbFacture.Size = New System.Drawing.Size(61, 17)
        Me.rbFacture.TabIndex = 8
        Me.rbFacture.TabStop = True
        Me.rbFacture.Text = "Facture"
        Me.rbFacture.UseVisualStyleBackColor = True
        '
        'rbReleve
        '
        Me.rbReleve.AutoSize = True
        Me.rbReleve.Location = New System.Drawing.Point(6, 62)
        Me.rbReleve.Name = "rbReleve"
        Me.rbReleve.Size = New System.Drawing.Size(59, 17)
        Me.rbReleve.TabIndex = 9
        Me.rbReleve.TabStop = True
        Me.rbReleve.Text = "Relevé"
        Me.rbReleve.UseVisualStyleBackColor = True
        '
        'ckEntete
        '
        Me.ckEntete.AutoSize = True
        Me.ckEntete.Location = New System.Drawing.Point(584, 15)
        Me.ckEntete.Name = "ckEntete"
        Me.ckEntete.Size = New System.Drawing.Size(57, 17)
        Me.ckEntete.TabIndex = 10
        Me.ckEntete.Text = "Entete"
        Me.ckEntete.UseVisualStyleBackColor = True
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.Location = New System.Drawing.Point(913, 58)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(75, 23)
        Me.cbAfficher.TabIndex = 11
        Me.cbAfficher.Text = "Afficher"
        Me.cbAfficher.UseVisualStyleBackColor = True
        '
        'laTiers
        '
        Me.laTiers.AutoSize = True
        Me.laTiers.Location = New System.Drawing.Point(206, 63)
        Me.laTiers.Name = "laTiers"
        Me.laTiers.Size = New System.Drawing.Size(64, 13)
        Me.laTiers.TabIndex = 13
        Me.laTiers.Text = "Code Tiers :"
        '
        'tbCodeTiers
        '
        Me.tbCodeTiers.Location = New System.Drawing.Point(297, 60)
        Me.tbCodeTiers.Name = "tbCodeTiers"
        Me.tbCodeTiers.Size = New System.Drawing.Size(88, 20)
        Me.tbCodeTiers.TabIndex = 14
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbFactcom)
        Me.GroupBox1.Controls.Add(Me.rbFactTRP)
        Me.GroupBox1.Controls.Add(Me.rbFactCol)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(188, 87)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Type de factures"
        '
        'grpTypeEditFactcom
        '
        Me.grpTypeEditFactcom.Controls.Add(Me.rbCourier)
        Me.grpTypeEditFactcom.Controls.Add(Me.rbFacture)
        Me.grpTypeEditFactcom.Controls.Add(Me.rbReleve)
        Me.grpTypeEditFactcom.Location = New System.Drawing.Point(391, 0)
        Me.grpTypeEditFactcom.Name = "grpTypeEditFactcom"
        Me.grpTypeEditFactcom.Size = New System.Drawing.Size(107, 87)
        Me.grpTypeEditFactcom.TabIndex = 16
        Me.grpTypeEditFactcom.TabStop = False
        Me.grpTypeEditFactcom.Text = "Type édition"
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(13, 94)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(975, 572)
        Me.CrystalReportViewer1.TabIndex = 17
        '
        'frmEditfactures
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit
        Me.ClientSize = New System.Drawing.Size(1000, 678)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Controls.Add(Me.grpTypeEditFactcom)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.tbCodeTiers)
        Me.Controls.Add(Me.laTiers)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.ckEntete)
        Me.Controls.Add(Me.dtDateFin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtDateDeb)
        Me.Name = "frmEditfactures"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Edition des factures"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.grpTypeEditFactcom.ResumeLayout(False)
        Me.grpTypeEditFactcom.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rbFactcom As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactTRP As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactCol As System.Windows.Forms.RadioButton
    Friend WithEvents dtDateDeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtDateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents rbCourier As System.Windows.Forms.RadioButton
    Friend WithEvents rbFacture As System.Windows.Forms.RadioButton
    Friend WithEvents rbReleve As System.Windows.Forms.RadioButton
    Friend WithEvents ckEntete As System.Windows.Forms.CheckBox
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents laTiers As System.Windows.Forms.Label
    Friend WithEvents tbCodeTiers As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents grpTypeEditFactcom As System.Windows.Forms.GroupBox
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
