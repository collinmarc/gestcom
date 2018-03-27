Imports vini_DB
Public Class frmImportTarif
    Inherits FrmVinicom
    Implements IObservateur

    Friend WithEvents Label1 As Label
    Friend WithEvents tbFilePath As TextBox
    Friend WithEvents btnParcourir As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents m_OpenFileDialog As OpenFileDialog
    Friend WithEvents nupNumColCode As NumericUpDown
    Friend WithEvents nupNumColTarif As NumericUpDown
    Friend WithEvents Label4 As Label
    Friend WithEvents tbNomFeuille As TextBox
    Private m_bLoad As Boolean = True
    Friend WithEvents lblMsg As Label
    Private m_oImportTarif As ImportTarifGESTCOM

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
    Friend WithEvents btnImporter As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents StatusDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusMessageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnImporter = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbFilePath = New System.Windows.Forms.TextBox()
        Me.btnParcourir = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.m_OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.nupNumColCode = New System.Windows.Forms.NumericUpDown()
        Me.nupNumColTarif = New System.Windows.Forms.NumericUpDown()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbNomFeuille = New System.Windows.Forms.TextBox()
        Me.lblMsg = New System.Windows.Forms.Label()
        CType(Me.nupNumColCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nupNumColTarif, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnImporter
        '
        Me.btnImporter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnImporter.Location = New System.Drawing.Point(14, 151)
        Me.btnImporter.Name = "btnImporter"
        Me.btnImporter.Size = New System.Drawing.Size(848, 24)
        Me.btnImporter.TabIndex = 4
        Me.btnImporter.Text = "Importer"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(14, 195)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(848, 23)
        Me.ProgressBar1.Step = 1
        Me.ProgressBar1.TabIndex = 132
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 13)
        Me.Label1.TabIndex = 133
        Me.Label1.Text = "Fichier de tarif (.xlsx) :"
        '
        'tbFilePath
        '
        Me.tbFilePath.Location = New System.Drawing.Point(126, 11)
        Me.tbFilePath.Name = "tbFilePath"
        Me.tbFilePath.Size = New System.Drawing.Size(626, 20)
        Me.tbFilePath.TabIndex = 134
        '
        'btnParcourir
        '
        Me.btnParcourir.Location = New System.Drawing.Point(781, 8)
        Me.btnParcourir.Name = "btnParcourir"
        Me.btnParcourir.Size = New System.Drawing.Size(75, 23)
        Me.btnParcourir.TabIndex = 135
        Me.btnParcourir.Text = "Parcourir"
        Me.btnParcourir.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(227, 13)
        Me.Label2.TabIndex = 136
        Me.Label2.Text = "Numéro de la colonne contenant le code (>=1)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(226, 13)
        Me.Label3.TabIndex = 138
        Me.Label3.Text = "Numéro de la colonne contenant le tarif (>=1) :"
        '
        'm_OpenFileDialog
        '
        Me.m_OpenFileDialog.DefaultExt = "*.*"
        Me.m_OpenFileDialog.FileName = "OpenFileDialog1"
        '
        'nupNumColCode
        '
        Me.nupNumColCode.Location = New System.Drawing.Point(257, 77)
        Me.nupNumColCode.Name = "nupNumColCode"
        Me.nupNumColCode.Size = New System.Drawing.Size(100, 20)
        Me.nupNumColCode.TabIndex = 140
        Me.nupNumColCode.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'nupNumColTarif
        '
        Me.nupNumColTarif.Location = New System.Drawing.Point(257, 112)
        Me.nupNumColTarif.Name = "nupNumColTarif"
        Me.nupNumColTarif.Size = New System.Drawing.Size(100, 20)
        Me.nupNumColTarif.TabIndex = 141
        Me.nupNumColTarif.Value = New Decimal(New Integer() {19, 0, 0, 0})
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 54)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 13)
        Me.Label4.TabIndex = 142
        Me.Label4.Text = "Nom de feuille :"
        '
        'tbNomFeuille
        '
        Me.tbNomFeuille.Location = New System.Drawing.Point(257, 51)
        Me.tbNomFeuille.Name = "tbNomFeuille"
        Me.tbNomFeuille.Size = New System.Drawing.Size(100, 20)
        Me.tbNomFeuille.TabIndex = 143
        Me.tbNomFeuille.Text = "Feuil1"
        '
        'lblMsg
        '
        Me.lblMsg.AutoSize = True
        Me.lblMsg.Location = New System.Drawing.Point(407, 221)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(39, 13)
        Me.lblMsg.TabIndex = 144
        Me.lblMsg.Text = "Label5"
        '
        'frmImportTarif
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(868, 470)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.tbNomFeuille)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.nupNumColTarif)
        Me.Controls.Add(Me.nupNumColCode)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnParcourir)
        Me.Controls.Add(Me.tbFilePath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnImporter)
        Me.Name = "frmImportTarif"
        Me.Text = "Import destarifs à partir du referentiel EXCEL"
        CType(Me.nupNumColCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nupNumColTarif, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Méthodes Vinicom"
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False
        m_ToolBarSaveEnabled = False
    End Sub

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)
        nupNumColCode.Enabled = True
        nupNumColTarif.Enabled = True
        btnParcourir.Enabled = True
        btnImporter.Enabled = True

    End Sub


#End Region

#Region "Methodes privées"
    Private Sub initFenetre()

        EnableControls(True)



    End Sub




#Region "Export"
#Region "Methodes backGoundWorker"
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        importer()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblMsg.Text = m_oImportTarif.message
        Catch ex As Exception

        End Try
    End Sub
#End Region


    ''' <summary>
    ''' Actualisation du travail
    ''' </summary>
    ''' <param name="pObj"></param>
    ''' <remarks></remarks>
    Public Sub Actualiser(pObj As Observable) Implements IObservateur.Actualiser
        BackgroundWorker1.ReportProgress(1)
    End Sub

#End Region


    Private Sub Log(ByVal strMessage As String)
        Trace.WriteLine(Now() + " " + strMessage)
    End Sub
#End Region
#Region "Gestion des Evenements"
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
        m_bLoad = False '' fin du load
    End Sub


#End Region

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Génerer déclenche L'evenement Updated
        'AddHandler cbAppliquer.Click, AddressOf ControlUpdated
        'AddHandler cbGenerer.Click, AddressOf ControlUpdated
    End Sub


    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        ProgressBar1.Value = ProgressBar1.Maximum
        MsgBox("Import Terminé")
        restoreCursor()
    End Sub

    Private Sub btnParcourir_Click(sender As Object, e As EventArgs) Handles btnParcourir.Click
        If m_OpenFileDialog.ShowDialog() Then
            tbFilePath.Text = m_OpenFileDialog.FileName
        End If
    End Sub

    Private Sub btnImporter_Click(sender As Object, e As EventArgs) Handles btnImporter.Click
        If MsgBox("Cet import va mettre à jour les TarifA ,TarifB et TarifC des produits, êtes-vous sûr ?", vbYesNo) = DialogResult.Yes Then
            m_oImportTarif = New ImportTarifGESTCOM(tbFilePath.Text, tbNomFeuille.Text, nupNumColCode.Value, nupNumColTarif.Value)
            m_oImportTarif.AjouteObservateur(Me)
            Me.ProgressBar1.Minimum = 0
            Me.ProgressBar1.Maximum = m_oImportTarif.getNbreLignes()
            Me.ProgressBar1.Value = Me.ProgressBar1.Minimum

            setcursorWait()
            BackgroundWorker1.RunWorkerAsync()
            restoreCursor()

        End If
    End Sub

    Private Sub importer()
        m_oImportTarif.ImportTarif()
    End Sub
End Class
