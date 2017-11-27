Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.CrystalReports.Design
'Imports FAXCOMLib
Imports vini_DB

Public Class dlgVisuBonFournisseur
    Inherits System.Windows.Forms.Form
    Private m_objSCMD As SousCommande


#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        If crwPrecommande.ReportSource IsNot Nothing Then
            Dim oReport As ReportDocument
            oReport = crwPrecommande.ReportSource
            oReport.Dispose()
            crwPrecommande.ReportSource = Nothing

        End If
        crwPrecommande.Dispose()

        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    Friend WithEvents tbNumFax As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents crwPrecommande As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents cbFaxer As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbNumFax = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbFaxer = New System.Windows.Forms.Button()
        Me.crwPrecommande = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'tbNumFax
        '
        Me.tbNumFax.Location = New System.Drawing.Point(390, 4)
        Me.tbNumFax.Name = "tbNumFax"
        Me.tbNumFax.Size = New System.Drawing.Size(160, 20)
        Me.tbNumFax.TabIndex = 80
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(248, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 79
        Me.Label1.Text = "Numéro de Fax"
        '
        'cbFaxer
        '
        Me.cbFaxer.Location = New System.Drawing.Point(584, 2)
        Me.cbFaxer.Name = "cbFaxer"
        Me.cbFaxer.Size = New System.Drawing.Size(96, 24)
        Me.cbFaxer.TabIndex = 82
        Me.cbFaxer.Text = "Fa&xer"
        '
        'crwPrecommande
        '
        Me.crwPrecommande.ActiveViewIndex = -1
        Me.crwPrecommande.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crwPrecommande.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crwPrecommande.Cursor = System.Windows.Forms.Cursors.Default
        Me.crwPrecommande.DisplayStatusBar = False
        Me.crwPrecommande.Location = New System.Drawing.Point(13, 44)
        Me.crwPrecommande.Name = "crwPrecommande"
        Me.crwPrecommande.Size = New System.Drawing.Size(663, 490)
        Me.crwPrecommande.TabIndex = 83
        Me.crwPrecommande.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'dlgVisuBonFournisseur
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(688, 536)
        Me.Controls.Add(Me.crwPrecommande)
        Me.Controls.Add(Me.cbFaxer)
        Me.Controls.Add(Me.tbNumFax)
        Me.Controls.Add(Me.Label1)
        Me.Name = "dlgVisuBonFournisseur"
        Me.Text = "dlgVisuPrecommande"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Sub setElement(ByVal obj As SousCommande)
        Debug.Assert(Not obj Is Nothing)
        m_objSCMD = obj
        m_objSCMD.oFournisseur.load()
        tbNumFax.Text = m_objSCMD.oFournisseur.AdresseFacturation.fax
        Text = "[SCMD]" & m_objSCMD.shortResume
        Visu()
    End Sub

    Private Sub Visu()
        Dim objReport As ReportDocument

        objReport = m_objSCMD.genererReport(PATHTOREPORTS)
        If Not objReport Is Nothing Then
            crwPrecommande.ReportSource = objReport
        End If

    End Sub

    Private Sub tbNumFax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbNumFax.TextChanged
        If tbNumFax.Text <> "" Then
            cbFaxer.Enabled = True
        End If
    End Sub


    '    Private Sub crwPrecommande_ReportRefresh(ByVal source As Object, ByVal e As CrystalDecisions.Windows.Forms.ViewerEventArgs) Handles crwPrecommande.ReportRefresh
    '        Dim objReport As ReportDocument
    '        m_objSCMD.genererReport()
    '        objReport = m_objSCMD.report
    '        crwPrecommande.ReportSource = objReport
    '   End Sub

    Private Sub cbFaxer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFaxer.Click
        Dim strFileName As String

        strFileName = FAX_SCMD_PATH & m_objSCMD.code & ".doc"

        m_objSCMD.faxer(PATHTOREPORTS, FAX_NOM_INTERLOCUTEUR, FAX_TEL_INTERLOCUTEUR, FAX_SCMD_SUBJECT, FAX_SCMD_NOTES, FAX_SCMD_BSENDCOVERPAGE, strFileName, tbNumFax.Text)
        m_objSCMD.changeEtat(vncActionEtatCommande.vncActionSCMDFaxer)
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub dlgVisuBonFournisseur_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.crwPrecommande.ShowPageNavigateButtons = True
    End Sub
End Class
