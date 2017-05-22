Option Explicit On 
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmRecapAnnonces
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
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtDebut As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboCodeTrp As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtFin As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtDebut = New System.Windows.Forms.DateTimePicker
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboCodeTrp = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtFin = New System.Windows.Forms.DateTimePicker
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(9, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Date de début"
        '
        'dtDebut
        '
        Me.dtDebut.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDebut.Location = New System.Drawing.Point(128, 8)
        Me.dtDebut.Name = "dtDebut"
        Me.dtDebut.Size = New System.Drawing.Size(136, 20)
        Me.dtDebut.TabIndex = 0
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(848, 8)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 23)
        Me.cbAfficher.TabIndex = 2
        Me.cbAfficher.Text = "&Afficher"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(478, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "code Transporteur"
        '
        'cboCodeTrp
        '
        Me.cboCodeTrp.ItemHeight = 13
        Me.cboCodeTrp.Location = New System.Drawing.Point(592, 8)
        Me.cboCodeTrp.Name = "cboCodeTrp"
        Me.cboCodeTrp.Size = New System.Drawing.Size(160, 21)
        Me.cboCodeTrp.Sorted = True
        Me.cboCodeTrp.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(280, 11)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "date de fin"
        '
        'dtFin
        '
        Me.dtFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFin.Location = New System.Drawing.Point(368, 8)
        Me.dtFin.Name = "dtFin"
        Me.dtFin.Size = New System.Drawing.Size(104, 20)
        Me.dtFin.TabIndex = 10
        '
        'frmRecapAnnonces
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(992, 694)
        Me.Controls.Add(Me.dtFin)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboCodeTrp)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.dtDebut)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmRecapAnnonces"
        Me.Text = "Recaptitulatif des journaux d'annonces"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtDebut, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.cboCodeTrp, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.dtFin, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public Overrides Function getResume() As String
        Return "Recapitulatif des journaux d'annonces"
    End Function
    Private Sub initFenetre()
        Dim strdate As String
        dtDebut.Value = "01/" & Now.Month & "/" & Now.Year
        strdate = "01/" & DateAdd(DateInterval.Month, 1, Now()).Month & "/" & DateAdd(DateInterval.Month, 1, Now()).Year
        dtFin.Value = strdate
        initcboCodeTRP(Me.cboCodeTrp)
    End Sub
    Private Sub afficheEtat()
        Dim objReport As ReportDocument


        objReport = New ReportDocument
        objReport.Load(PATHTOREPORTS & "crRecapAnnonces.rpt")


        objReport.SetParameterValue("dateDebut", Me.dtDebut.Value.ToShortDateString)

        objReport.SetParameterValue("dateFin", Me.dtFin.Value.ToShortDateString)

        objReport.SetParameterValue("codeTransporteur", Me.cboCodeTrp.Text)


        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport

    End Sub
    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        afficheEtat()
    End Sub


    Private Sub frmHistoriqueArticles_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub


End Class
