Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
'Imports FAXCOMLib
Imports vini_DB

Public Class dlgVisuPrecommande
    Inherits System.Windows.Forms.Form
    Private m_objClient As Client


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
        Me.tbNumFax.Location = New System.Drawing.Point(392, 8)
        Me.tbNumFax.Name = "tbNumFax"
        Me.tbNumFax.Size = New System.Drawing.Size(160, 20)
        Me.tbNumFax.TabIndex = 80
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(248, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 79
        Me.Label1.Text = "Numéro de Fax"
        '
        'cbFaxer
        '
        Me.cbFaxer.Location = New System.Drawing.Point(584, 8)
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
        Me.crwPrecommande.Location = New System.Drawing.Point(13, 38)
        Me.crwPrecommande.Name = "crwPrecommande"
        Me.crwPrecommande.Size = New System.Drawing.Size(667, 490)
        Me.crwPrecommande.TabIndex = 83
        '
        'dlgVisuPrecommande
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(691, 540)
        Me.Controls.Add(Me.crwPrecommande)
        Me.Controls.Add(Me.cbFaxer)
        Me.Controls.Add(Me.tbNumFax)
        Me.Controls.Add(Me.Label1)
        Me.Name = "dlgVisuPrecommande"
        Me.Text = "dlgVisuPrecommande"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Sub setClient(ByVal objClient As Client)
        Debug.Assert(Not objClient Is Nothing)
        m_objClient = objClient
        tbNumFax.Text = m_objClient.AdresseLivraison.fax
        Text = "[VPCMD]" & m_objClient.shortResume
        Visu()
    End Sub

    Private Sub Visu()
        Dim objReport As ReportDocument

        objReport = New ReportDocument
        objReport.Load(PATHTOREPORTS & "crPrecommande.rpt")
        objReport.SetParameterValue("IDCLIENT", m_objClient.id)
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
        'Dim StrFileNAme As String
        'Dim StrFileNAme2 As String



        'objReport = crwPrecommande.ReportSource
        'objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
        'objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
        'StrFileNAme = FAX_PRECOMMANDE_PATH & "PCMD_" & m_objClient.code & ".doc"
        'StrFileNAme2 = FAX_PRECOMMANDE_PATH & "PCMDbis_" & m_objClient.code & ".doc"
        'diskOpts.DiskFileName = StrFileNAme
        'objReport.ExportOptions.DestinationOptions = diskOpts
        'objReport.Export()
        'objReport.ExportToDisk(ExportFormatType.WordForWindows, StrFileNAme2)
        'WaitnSeconds(5)

        'objFax = New clsFax

        'objFax.sendFax(FAX_NOM_INTERLOCUTEUR, FAX_TEL_INTERLOCUTEUR, FAX_PRECOMMANDE_SUBJECT, FAX_PRECOMMANDE_NOTES, FAX_PRECOMMANDE_BSENDCOVERPAGE, StrFileNAme, tbNumFax.Text, m_objClient)
        'objReport = Nothing
    End Sub
    Public Sub WaitnSeconds(ByVal n As Integer)
        If TimeOfDay >= #11:59:55 PM# Then
        End If
        Dim Start, Finish, TotalTime As Double
        Start = Microsoft.VisualBasic.DateAndTime.Timer
        Finish = Start + n   ' Set end time for 5-second duration.
        Do While Microsoft.VisualBasic.DateAndTime.Timer < Finish
            ' Do other processing while waiting for 5 seconds to elapse.
        Loop
        TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - Start
    End Sub
End Class
