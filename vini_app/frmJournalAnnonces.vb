Option Explicit On 
Option Strict Off
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmJournalAnnonces
    Inherits frmStatistiques
    Private m_bAnnoncesCommandes As Boolean

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
    Friend WithEvents dtEnlev As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbFaxTRP As System.Windows.Forms.TextBox
    Friend WithEvents cbFaxer As System.Windows.Forms.Button
    Friend WithEvents cboCodeTrp As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtEnlev = New System.Windows.Forms.DateTimePicker
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboCodeTrp = New System.Windows.Forms.ComboBox
        Me.tbFaxTRP = New System.Windows.Forms.TextBox
        Me.cbFaxer = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(9, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Date d'enlèvement"
        '
        'dtEnlev
        '
        Me.dtEnlev.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtEnlev.Location = New System.Drawing.Point(128, 8)
        Me.dtEnlev.Name = "dtEnlev"
        Me.dtEnlev.Size = New System.Drawing.Size(136, 20)
        Me.dtEnlev.TabIndex = 0
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(568, 8)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 23)
        Me.cbAfficher.TabIndex = 2
        Me.cbAfficher.Text = "&Afficher"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(278, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "code Transporteur"
        '
        'cboCodeTrp
        '
        Me.cboCodeTrp.Location = New System.Drawing.Point(384, 8)
        Me.cboCodeTrp.Name = "cboCodeTrp"
        Me.cboCodeTrp.Size = New System.Drawing.Size(160, 21)
        Me.cboCodeTrp.Sorted = True
        Me.cboCodeTrp.TabIndex = 1
        '
        'tbFaxTRP
        '
        Me.tbFaxTRP.Location = New System.Drawing.Point(704, 8)
        Me.tbFaxTRP.Name = "tbFaxTRP"
        Me.tbFaxTRP.Size = New System.Drawing.Size(100, 20)
        Me.tbFaxTRP.TabIndex = 3
        '
        'cbFaxer
        '
        Me.cbFaxer.Location = New System.Drawing.Point(824, 8)
        Me.cbFaxer.Name = "cbFaxer"
        Me.cbFaxer.Size = New System.Drawing.Size(75, 23)
        Me.cbFaxer.TabIndex = 4
        Me.cbFaxer.Text = "Fa&xer"
        '
        'frmJournalAnnonces
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(992, 694)
        Me.Controls.Add(Me.cbFaxer)
        Me.Controls.Add(Me.tbFaxTRP)
        Me.Controls.Add(Me.cboCodeTrp)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.dtEnlev)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmJournalAnnonces"
        Me.Text = "Journal d'annonces"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtEnlev, 0)
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.cboCodeTrp, 0)
        Me.Controls.SetChildIndex(Me.tbFaxTRP, 0)
        Me.Controls.SetChildIndex(Me.cbFaxer, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Public Overrides Function getResume() As String
        If m_bAnnoncesCommandes Then
            Return "Journal d'annonces LIVRAISON"
        Else
            Return "Journal d'annonces APPROVISIONNEMENT"
        End If
    End Function
    ''' <summary>
    ''' Affecte le boolean decidant s'il s'agit du journal d'annonces de commandes ou d'Approvisionement
    ''' </summary>
    ''' <param name="bValue"></param>
    ''' <remarks></remarks>
    Public Sub setbAnnoncesCommmandes(ByVal bValue As Boolean)
        m_bAnnoncesCommandes = bValue
    End Sub
    Private Sub initFenetre()
        dtEnlev.Value = Now()
        initcboCodeTRP(Me.cboCodeTrp, True)
    End Sub
    Private Sub afficheEtat()
        Dim objReport As ReportDocument
        Dim objTRP As Transporteur
        Dim strCodeTRP As String
        setcursorWait()
        If Not cboCodeTrp.SelectedItem Is Nothing Then
            objTRP = cboCodeTrp.SelectedItem

            objReport = New ReportDocument
            If m_bAnnoncesCommandes Then
                objReport.Load(PATHTOREPORTS & "crJournalAnnoncesLivraison.rpt")
            Else
                objReport.Load(PATHTOREPORTS & "crJournalAnnoncesApprovisionnement.rpt")
            End If

            objReport.SetParameterValue("dateEnlevement", Me.dtEnlev.Value.ToShortDateString)
            If objTRP.code = "*/TOUS" Then
                strCodeTRP = "*"
            Else
                strCodeTRP = objTRP.code
            End If
            objReport.SetParameterValue("codetransporteur", strCodeTRP)

            Persist.setReportConnection(objReport)
            CrystalReportViewer1.ReportSource = objReport
        End If
        restoreCursor()

    End Sub
    Private Sub faxerEtat()
        'Dim diskOpts As New CrystalDecisions.Shared.DiskFileDestinationOptions
        'Dim objFax As clsFax
        'Dim objReport As ReportDocument
        'Dim objTRP As Transporteur
        'Dim strFileName As String

        'If tbFaxTRP.Text <> "" Then

        '    setcursorWait()
        '    If Not CrystalReportViewer1.ReportSource Is Nothing Then

        '        objTRP = cboCodeTrp.SelectedItem

        '        objReport = CrystalReportViewer1.ReportSource
        '        objReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
        '        objReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
        '        If m_bAnnoncesCommandes Then
        '            strFileName = FAX_JAL_PATH & "JAL" & Trim(Me.cboCodeTrp.Text) & "_" & Me.dtEnlev.Value.Year & Me.dtEnlev.Value.Month & Me.dtEnlev.Value.Day & ".doc"
        '        Else
        '            strFileName = FAX_JAL_PATH & "JAA" & Trim(Me.cboCodeTrp.Text) & "_" & Me.dtEnlev.Value.Year & Me.dtEnlev.Value.Month & Me.dtEnlev.Value.Day & ".doc"
        '        End If
        '        diskOpts.DiskFileName = strFileName
        '        objReport.ExportOptions.DestinationOptions = diskOpts
        '        Try
        '            objReport.Export()
        '            objFax = New clsFax
        '            If m_bAnnoncesCommandes Then
        '                objFax.sendFax(FAX_NOM_INTERLOCUTEUR, FAX_TEL_INTERLOCUTEUR, FAX_JAL_SUBJECT, FAX_JAL_NOTES, FAX_JAL_BSENDCOVERPAGE, diskOpts.DiskFileName, tbFaxTRP.Text, objTRP, True)
        '            Else
        '                objFax.sendFax(FAX_NOM_INTERLOCUTEUR, FAX_TEL_INTERLOCUTEUR, FAX_JAA_SUBJECT, FAX_JAA_NOTES, FAX_JAA_BSENDCOVERPAGE, diskOpts.DiskFileName, tbFaxTRP.Text, objTRP, True)
        '            End If

        '        Catch ex As Exception
        '            DisplayError("FaxerBL", "Erreur en Envoi de Fax|" & strFileName & "|" & ex.ToString)
        '        End Try
        '    Else
        '        MsgBox("Vous devez d'abord visualiser le Journal")
        '    End If
        '    restoreCursor()
        'End If

    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        afficheEtat()
    End Sub


    Private Sub frmHistoriqueArticles_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Private Sub cboCodeTrp_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeTrp.SelectedIndexChanged
        Dim oTRp As Transporteur
        If Not bAffichageEnCours() Then
            Try
                If Not cboCodeTrp.SelectedItem Is Nothing Then
                    oTRp = cboCodeTrp.SelectedItem
                    debAffiche()
                    tbFaxTRP.Text = oTRp.AdresseLivraison.fax
                    finAffiche()
                End If
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub cbFaxer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbFaxer.Click
        faxerEtat()
    End Sub
End Class
