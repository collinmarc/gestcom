<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportcommandeClientPrestashop
    Inherits vini_app.FrmVinicom

    'Form remplace la m�thode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE�: la proc�dure suivante est requise par le Concepteur Windows Form
    'Elle peut �tre modifi�e � l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas � l'aide de l'�diteur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.lbMessages = New System.Windows.Forms.ListBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'lbMessages
        '
        Me.lbMessages.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbMessages.FormattingEnabled = True
        Me.lbMessages.Location = New System.Drawing.Point(12, 17)
        Me.lbMessages.Name = "lbMessages"
        Me.lbMessages.Size = New System.Drawing.Size(726, 472)
        Me.lbMessages.TabIndex = 22
        '
        'Timer1
        '
        Me.Timer1.Interval = 300000
        '
        'frmImportcommandeClientPrestashop
        '
        Me.ClientSize = New System.Drawing.Size(742, 495)
        Me.Controls.Add(Me.lbMessages)
        Me.Name = "frmImportcommandeClientPrestashop"
        Me.Text = "Import des commandes clients cr�es le site internet"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lbMessages As System.Windows.Forms.ListBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
