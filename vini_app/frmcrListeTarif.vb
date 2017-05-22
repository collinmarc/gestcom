Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Public Class frmcrListeTarif
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
    Friend WithEvents m_bsrcFournisseur As System.Windows.Forms.BindingSource
    Friend WithEvents cbxFournisseur As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbxTarif As System.Windows.Forms.ComboBox
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.m_bsrcFournisseur = New System.Windows.Forms.BindingSource(Me.components)
        Me.cbxFournisseur = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cbxTarif = New System.Windows.Forms.ComboBox
        CType(Me.m_bsrcFournisseur, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbAfficher
        '
        Me.cbAfficher.Location = New System.Drawing.Point(868, 3)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(120, 23)
        Me.cbAfficher.TabIndex = 5
        Me.cbAfficher.Text = "Afficher"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Fournisseur :"
        '
        'm_bsrcFournisseur
        '
        Me.m_bsrcFournisseur.DataSource = GetType(vini_DB.Fournisseur)
        '
        'cbxFournisseur
        '
        Me.cbxFournisseur.DataSource = Me.m_bsrcFournisseur
        Me.cbxFournisseur.DisplayMember = "shortResume"
        Me.cbxFournisseur.FormattingEnabled = True
        Me.cbxFournisseur.Location = New System.Drawing.Point(109, 5)
        Me.cbxFournisseur.Name = "cbxFournisseur"
        Me.cbxFournisseur.Size = New System.Drawing.Size(495, 21)
        Me.cbxFournisseur.TabIndex = 7
        Me.cbxFournisseur.ValueMember = "code"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(610, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Code Tarif :"
        '
        'cbxTarif
        '
        Me.cbxTarif.FormattingEnabled = True
        Me.cbxTarif.Items.AddRange(New Object() {"A", "B", "C", "*"})
        Me.cbxTarif.Location = New System.Drawing.Point(679, 5)
        Me.cbxTarif.Name = "cbxTarif"
        Me.cbxTarif.Size = New System.Drawing.Size(51, 21)
        Me.cbxTarif.TabIndex = 9
        '
        'frmcrListeTarif
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1000, 678)
        Me.Controls.Add(Me.cbxTarif)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbxFournisseur)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbAfficher)
        Me.Name = "frmcrListeTarif"
        Me.Text = "Edition du Tarif"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.cbAfficher, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.cbxFournisseur, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.cbxTarif, 0)
        CType(Me.m_bsrcFournisseur, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Overrides Function getResume() As String
        Return "Edition du Tarif"
    End Function



    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click

        Dim objReport As ReportDocument
        Dim myConnectionInfo As ConnectionInfo = New ConnectionInfo()
        myConnectionInfo.ServerName = Persist.oleDBConnection.DataSource
        myConnectionInfo.DatabaseName = Persist.oleDBConnection.Database
        myConnectionInfo.IntegratedSecurity = True

        Dim r1 As New crListeTarif()
        Dim strReportName As String = r1.ResourceName

        objReport = New ReportDocument()
        objReport.Load(PATHTOREPORTS & strReportName)

        objReport.SetParameterValue("CODEFOURN", Me.cbxFournisseur.SelectedItem.Code)
        objReport.SetParameterValue("CODETARIF", Me.cbxTarif.Text)
        Persist.setReportConnection(objReport)
        CrystalReportViewer1.ReportSource = objReport
    End Sub

    Private Sub SetDBLogonForReport(ByVal myConnectionInfo As ConnectionInfo, ByVal myReportDocument As ReportDocument)


    End Sub


    Private Sub frmcrListeTarif_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim oCol As Collection
        Dim oFRN As Fournisseur

        oCol = Fournisseur.getListe()
        For Each oFRN In oCol
            m_bsrcFournisseur.Add(oFRN)
        Next
        oFRN = New Fournisseur("%", "Tous")
        m_bsrcFournisseur.Add(oFRN)

        cbxTarif.SelectedText = "*"

    End Sub
End Class
