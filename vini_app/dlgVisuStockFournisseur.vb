Option Strict Off
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
'Imports FAXCOMLib
Imports vini_DB

Public Class dlgVisuStockFournisseur
    Inherits System.Windows.Forms.Form
    Private m_objFournisseur As Fournisseur


#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    Friend WithEvents cbFaxer As System.Windows.Forms.Button
    Friend WithEvents tbNumFax As System.Windows.Forms.TextBox
    Friend WithEvents crwPrecommande As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cbFaxer = New System.Windows.Forms.Button()
        Me.tbNumFax = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.crwPrecommande = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'cbFaxer
        '
        Me.cbFaxer.Enabled = False
        Me.cbFaxer.Location = New System.Drawing.Point(584, 8)
        Me.cbFaxer.Name = "cbFaxer"
        Me.cbFaxer.Size = New System.Drawing.Size(96, 24)
        Me.cbFaxer.TabIndex = 85
        Me.cbFaxer.Text = "Fa&xer"
        '
        'tbNumFax
        '
        Me.tbNumFax.Location = New System.Drawing.Point(392, 8)
        Me.tbNumFax.Name = "tbNumFax"
        Me.tbNumFax.Size = New System.Drawing.Size(160, 20)
        Me.tbNumFax.TabIndex = 84
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(248, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 83
        Me.Label1.Text = "Numéro de Fax"
        '
        'crwPrecommande
        '
        Me.crwPrecommande.ActiveViewIndex = -1
        Me.crwPrecommande.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crwPrecommande.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crwPrecommande.Cursor = System.Windows.Forms.Cursors.Default
        Me.crwPrecommande.Location = New System.Drawing.Point(13, 36)
        Me.crwPrecommande.Name = "crwPrecommande"
        Me.crwPrecommande.Size = New System.Drawing.Size(663, 566)
        Me.crwPrecommande.TabIndex = 86
        '
        'dlgVisuStockFournisseur
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(688, 614)
        Me.Controls.Add(Me.crwPrecommande)
        Me.Controls.Add(Me.cbFaxer)
        Me.Controls.Add(Me.tbNumFax)
        Me.Controls.Add(Me.Label1)
        Me.Name = "dlgVisuStockFournisseur"
        Me.Text = "dlgVisuPrecommande"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Sub setFournisseur(ByVal objFournisseur As Fournisseur)
        Debug.Assert(Not objFournisseur Is Nothing)
        m_objFournisseur = objFournisseur
        tbNumFax.Text = m_objFournisseur.AdresseLivraison.fax
        Text = "[VSTKF]" & m_objFournisseur.shortResume
        Visu()
    End Sub

    Private Sub Visu()
        Dim objReport As ReportDocument

        objReport = New ReportDocument
        objReport.Load(CStr(PATHTOREPORTS) & "crStockFournisseur.rpt")
        objReport.SetParameterValue("IDFournisseur", m_objFournisseur.id)
        Persist.setReportConnection(objReport)
        crwPrecommande.ReportSource = objReport
    End Sub
    Private Sub tbNumFax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbNumFax.TextChanged
        If tbNumFax.Text <> "" Then
            cbFaxer.Enabled = True
        End If
    End Sub
    Private Sub cbFaxer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFaxer.Click
        'Dim diskOpts As New CrystalDecisions.Shared.DiskFileDestinationOptions
        'Dim objFax As clsFax
        'Dim objReport As ReportDocument


        'objReport = crwPrecommande.ReportSource
        'objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
        'objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
        'diskOpts.DiskFileName = FAX_STK_PATH & "PSTK_" & m_objFournisseur.code & ".doc"
        'objReport.ExportOptions.DestinationOptions = diskOpts
        'objReport.Export()

        'objFax = New clsFax

        'objFax.sendFax(FAX_NOM_INTERLOCUTEUR, FAX_TEL_INTERLOCUTEUR, FAX_STK_SUBJECT, FAX_STK_NOTES, FAX_STK_BSENDCOVERPAGE, diskOpts.DiskFileName, tbNumFax.Text, m_objFournisseur)
    End Sub


End Class
