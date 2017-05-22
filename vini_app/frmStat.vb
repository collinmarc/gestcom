Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmStat
    Inherits FrmVinicom

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
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents tbReport As System.Windows.Forms.TextBox
    Friend WithEvents cbRechercher As System.Windows.Forms.Button
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents cbVoir As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.tbReport = New System.Windows.Forms.TextBox()
        Me.cbRechercher = New System.Windows.Forms.Button()
        Me.cbVoir = New System.Windows.Forms.Button()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        '
        'tbReport
        '
        Me.tbReport.Location = New System.Drawing.Point(80, 8)
        Me.tbReport.Name = "tbReport"
        Me.tbReport.Size = New System.Drawing.Size(392, 20)
        Me.tbReport.TabIndex = 1
        '
        'cbRechercher
        '
        Me.cbRechercher.Location = New System.Drawing.Point(480, 8)
        Me.cbRechercher.Name = "cbRechercher"
        Me.cbRechercher.Size = New System.Drawing.Size(80, 24)
        Me.cbRechercher.TabIndex = 2
        Me.cbRechercher.Text = "Rechercher"
        '
        'cbVoir
        '
        Me.cbVoir.Location = New System.Drawing.Point(568, 8)
        Me.cbVoir.Name = "cbVoir"
        Me.cbVoir.Size = New System.Drawing.Size(88, 24)
        Me.cbVoir.TabIndex = 3
        Me.cbVoir.Text = "Afficher"
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.CrystalReportViewer1.DisplayStatusBar = False
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(13, 38)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(1003, 692)
        Me.CrystalReportViewer1.TabIndex = 4
        Me.CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'frmStat
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Controls.Add(Me.cbVoir)
        Me.Controls.Add(Me.cbRechercher)
        Me.Controls.Add(Me.tbReport)
        Me.Name = "frmStat"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Execution d'un état"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbRechercher.Click
        Dim strNomfichier As String = ""
        Dim dlgFichier As OpenFileDialog
        dlgFichier = New OpenFileDialog
        dlgFichier.InitialDirectory = PATHTOREPORTS
        dlgFichier.Filter = "Etat CrystalReport (*.rpt)|*.rpt"
        If (dlgFichier.ShowDialog()) Then
            strNomfichier = dlgFichier.FileName
        End If

        'strNomfichier = OpenFileDialog1.ShowDialog()
        Me.tbReport.Text = strNomfichier

    End Sub

    Private Sub cbVoir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVoir.Click
        Dim objReport As ReportDocument
        objReport = New ReportDocument()
        Try
            objReport.Load(tbReport.Text)
            Persist.setReportConnection(objReport)

            Me.CrystalReportViewer1.ReportSource = objReport

            Me.CrystalReportViewer1.RefreshReport()

        Catch

        End Try

    End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub frmStat_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tbReport.Text = PATHTOREPORTS
    End Sub
    Protected Overrides Sub EnableControls(bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub
End Class
