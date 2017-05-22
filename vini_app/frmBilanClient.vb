Option Explicit On 
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmBilanClient
    Inherits frmStatistiques

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents dtdeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbCodeClient As System.Windows.Forms.TextBox
    Friend WithEvents dtFin As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtdeb = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtFin = New System.Windows.Forms.DateTimePicker
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbCodeClient = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(9, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Date de début :"
        '
        'dtdeb
        '
        Me.dtdeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdeb.Location = New System.Drawing.Point(99, 8)
        Me.dtdeb.Name = "dtdeb"
        Me.dtdeb.Size = New System.Drawing.Size(136, 20)
        Me.dtdeb.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(241, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Date de fin :"
        '
        'dtFin
        '
        Me.dtFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFin.Location = New System.Drawing.Point(316, 8)
        Me.dtFin.Name = "dtFin"
        Me.dtFin.Size = New System.Drawing.Size(104, 20)
        Me.dtFin.TabIndex = 3
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.Location = New System.Drawing.Point(820, 5)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 23)
        Me.cbAfficher.TabIndex = 6
        Me.cbAfficher.Text = "Afficher"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(437, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(66, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Code client :"
        '
        'tbCodeClient
        '
        Me.tbCodeClient.Location = New System.Drawing.Point(510, 8)
        Me.tbCodeClient.Name = "tbCodeClient"
        Me.tbCodeClient.Size = New System.Drawing.Size(116, 20)
        Me.tbCodeClient.TabIndex = 5
        '
        'frmBilanClient
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(976, 678)
        Me.Controls.Add(Me.tbCodeClient)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.dtFin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtdeb)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmBilanClient"
        Me.Text = "Bilan commercial Client"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtdeb, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.dtFin, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.tbCodeClient, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        Dim objReport As ReportDocument
        Dim anneeN_1 As Integer
        Dim strCodeClient As String
        CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.GroupTree

        objReport = New ReportDocument
        objReport.Load(PATHTOREPORTS & "crStatBilanClient.rpt")
        objReport.SetParameterValue("ddeb", Me.dtdeb.Value.ToShortDateString())
        objReport.SetParameterValue("dfin", Me.dtFin.Value.ToShortDateString())
        anneeN_1 = Year(DateAdd(DateInterval.Year, -1, dtdeb.Value))
        objReport.SetParameterValue("N-1", anneeN_1)
        strCodeClient = tbCodeClient.Text
        strCodeClient = strCodeClient.Replace("%", "*")
        If String.IsNullOrEmpty(strCodeClient) Then
            strCodeClient = "*"
        End If
        objReport.SetParameterValue("codeClient", strCodeClient)

        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport


    End Sub

    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarSaveEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub frmBilanClient_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dtdeb.Value = CDate("01/01/" & Year(DateTime.Now))
    End Sub
End Class
