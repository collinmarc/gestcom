<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgInternet
    Inherits System.Windows.Forms.Form

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
        Me.WebBrowser4 = New System.Windows.Forms.WebBrowser
        Me.SuspendLayout()
        '
        'WebBrowser4
        '
        Me.WebBrowser4.DataBindings.Add(New System.Windows.Forms.Binding("Url", Global.vini_internet.My.MySettings.Default, "urlImportInternet", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.WebBrowser4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebBrowser4.Location = New System.Drawing.Point(0, 0)
        Me.WebBrowser4.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser4.Name = "WebBrowser4"
        Me.WebBrowser4.Size = New System.Drawing.Size(813, 463)
        Me.WebBrowser4.TabIndex = 3
        Me.WebBrowser4.Url = Global.vini_internet.My.MySettings.Default.urlImportInternet
        '
        'dlgInternet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(813, 463)
        Me.Controls.Add(Me.WebBrowser4)
        Me.Name = "dlgInternet"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents WebBrowser4 As System.Windows.Forms.WebBrowser
End Class
