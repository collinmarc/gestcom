Imports CrystalDecisions.CrystalReports.Engine

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmLstSCMDLitige
    Inherits vini_app.frmStatistiques

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        If CrystalReportViewer1.ReportSource IsNot Nothing Then
            Dim oReport As ReportDocument
            oReport = CrystalReportViewer1.ReportSource
            oReport.Dispose()
            CrystalReportViewer1.ReportSource = Nothing

        End If
        CrystalReportViewer1.Dispose()

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
        Me.dtDateDeb = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtDateFin = New System.Windows.Forms.DateTimePicker
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbCodeFournisseur = New System.Windows.Forms.TextBox
        Me.cbRechercher = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "date de début :"
        '
        'dtDateDeb
        '
        Me.dtDateDeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateDeb.Location = New System.Drawing.Point(99, 5)
        Me.dtDateDeb.Name = "dtDateDeb"
        Me.dtDateDeb.Size = New System.Drawing.Size(93, 20)
        Me.dtDateDeb.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(224, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "date de fin :"
        '
        'dtDateFin
        '
        Me.dtDateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateFin.Location = New System.Drawing.Point(302, 5)
        Me.dtDateFin.Name = "dtDateFin"
        Me.dtDateFin.Size = New System.Drawing.Size(86, 20)
        Me.dtDateFin.TabIndex = 4
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(865, 4)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(75, 23)
        Me.cbAfficher.TabIndex = 5
        Me.cbAfficher.Text = "Afficher"
        Me.cbAfficher.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(409, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(98, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Code Fournisseur : "
        '
        'tbCodeFournisseur
        '
        Me.tbCodeFournisseur.Location = New System.Drawing.Point(513, 6)
        Me.tbCodeFournisseur.Name = "tbCodeFournisseur"
        Me.tbCodeFournisseur.Size = New System.Drawing.Size(100, 20)
        Me.tbCodeFournisseur.TabIndex = 7
        '
        'cbRechercher
        '
        Me.cbRechercher.Location = New System.Drawing.Point(637, 4)
        Me.cbRechercher.Name = "cbRechercher"
        Me.cbRechercher.Size = New System.Drawing.Size(75, 23)
        Me.cbRechercher.TabIndex = 8
        Me.cbRechercher.Text = "Rechercher"
        Me.cbRechercher.UseVisualStyleBackColor = True
        '
        'frmLstCommLitige
        '
        Me.ClientSize = New System.Drawing.Size(952, 646)
        Me.Controls.Add(Me.cbRechercher)
        Me.Controls.Add(Me.tbCodeFournisseur)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtDateFin)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.dtDateDeb)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmLstCommLitige"
        Me.Text = "Liste des commissions en litige"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtDateDeb, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.dtDateFin, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.tbCodeFournisseur, 0)
        Me.Controls.SetChildIndex(Me.cbRechercher, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtDateDeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtDateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbCodeFournisseur As System.Windows.Forms.TextBox
    Friend WithEvents cbRechercher As System.Windows.Forms.Button

End Class
