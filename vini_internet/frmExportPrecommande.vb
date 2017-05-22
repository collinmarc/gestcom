Imports vini_DB
Public Class frmExportPreCommande
    Inherits frmStatistiques
    Protected m_colCommandes As Collection

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
    Friend WithEvents cbExporter As System.Windows.Forms.Button
    Friend WithEvents ckFTP As System.Windows.Forms.CheckBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents tbCodeClient As System.Windows.Forms.TextBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents m_bsrcSCMD As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcStatus As System.Windows.Forms.BindingSource
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FournisseurCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusMessageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvStatus As System.Windows.Forms.DataGridView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.cbExporter = New System.Windows.Forms.Button
        Me.ckFTP = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbCodeClient = New System.Windows.Forms.TextBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.m_bsrcSCMD = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_bsrcStatus = New System.Windows.Forms.BindingSource(Me.components)
        Me.dgvStatus = New System.Windows.Forms.DataGridView
        Me.StatusDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.StatusMessageDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.m_bsrcSCMD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbExporter
        '
        Me.cbExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbExporter.Location = New System.Drawing.Point(217, 9)
        Me.cbExporter.Name = "cbExporter"
        Me.cbExporter.Size = New System.Drawing.Size(64, 24)
        Me.cbExporter.TabIndex = 4
        Me.cbExporter.Text = "Exporter"
        '
        'ckFTP
        '
        Me.ckFTP.Checked = CType(configurationAppSettings.GetValue("ckFTP.Checked", GetType(Boolean)), Boolean)
        Me.ckFTP.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckFTP.Location = New System.Drawing.Point(761, 8)
        Me.ckFTP.Name = "ckFTP"
        Me.ckFTP.Size = New System.Drawing.Size(104, 24)
        Me.ckFTP.TabIndex = 128
        Me.ckFTP.Text = "ckFTP"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(5, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 20)
        Me.Label1.TabIndex = 129
        Me.Label1.Text = "Code Client :"
        '
        'tbCodeClient
        '
        Me.tbCodeClient.Location = New System.Drawing.Point(79, 12)
        Me.tbCodeClient.Name = "tbCodeClient"
        Me.tbCodeClient.Size = New System.Drawing.Size(88, 20)
        Me.tbCodeClient.TabIndex = 2
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 39)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(839, 23)
        Me.ProgressBar1.TabIndex = 132
        '
        'm_bsrcStatus
        '
        Me.m_bsrcStatus.DataSource = GetType(vini_internet.clsExportstatus)
        '
        'dgvStatus
        '
        Me.dgvStatus.AllowUserToAddRows = False
        Me.dgvStatus.AllowUserToDeleteRows = False
        Me.dgvStatus.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvStatus.AutoGenerateColumns = False
        Me.dgvStatus.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStatus.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.StatusDateDataGridViewTextBoxColumn, Me.StatusMessageDataGridViewTextBoxColumn})
        Me.dgvStatus.DataSource = Me.m_bsrcStatus
        Me.dgvStatus.Location = New System.Drawing.Point(8, 69)
        Me.dgvStatus.Name = "dgvStatus"
        Me.dgvStatus.ReadOnly = True
        Me.dgvStatus.Size = New System.Drawing.Size(843, 315)
        Me.dgvStatus.TabIndex = 134
        '
        'StatusDateDataGridViewTextBoxColumn
        '
        Me.StatusDateDataGridViewTextBoxColumn.DataPropertyName = "statusDate"
        DataGridViewCellStyle1.Format = "G"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.StatusDateDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle1
        Me.StatusDateDataGridViewTextBoxColumn.HeaderText = "Date"
        Me.StatusDateDataGridViewTextBoxColumn.Name = "StatusDateDataGridViewTextBoxColumn"
        Me.StatusDateDataGridViewTextBoxColumn.ReadOnly = True
        '
        'StatusMessageDataGridViewTextBoxColumn
        '
        Me.StatusMessageDataGridViewTextBoxColumn.DataPropertyName = "statusMessage"
        Me.StatusMessageDataGridViewTextBoxColumn.HeaderText = "Message"
        Me.StatusMessageDataGridViewTextBoxColumn.Name = "StatusMessageDataGridViewTextBoxColumn"
        Me.StatusMessageDataGridViewTextBoxColumn.ReadOnly = True
        '
        'frmExportPreCommande
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(863, 396)
        Me.Controls.Add(Me.dgvStatus)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.tbCodeClient)
        Me.Controls.Add(Me.ckFTP)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbExporter)
        Me.Name = "frmExportPreCommande"
        Me.Text = "Exportation  des préCommandes  vers Internet"
        CType(Me.m_bsrcSCMD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcStatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).EndInit()
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


#End Region

#Region "Methodes privées"



    Private Function exporter() As Boolean

        Dim objClt As Client
        Dim bExportOK As Boolean = True
        Dim bReturn As Boolean
        Dim strFile As String
        Dim strSCMD_CSV As String
        Dim nFile As Integer
        Dim strFolder As String
        Dim oFTPvinicom As clsFTPVinicom
        Dim nficherExporte As Integer
        Dim strStatus As String
        Dim oColClient As Collection

        Try
            bReturn = False
            oColClient = Client.getListe(Me.tbCodeClient.Text)
            ProgressBar1.Maximum = oColClient.Count
            ProgressBar1.Value = 0
            setcursorWait()
            dgvStatus.Rows.Clear()
            strFolder = My.Settings.Default.PreCommandeFolder
            If My.Computer.FileSystem.DirectoryExists(strFolder) Then
                My.Computer.FileSystem.DeleteDirectory(strFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            My.Computer.FileSystem.CreateDirectory(strFolder)

            strFile = strFolder + "/PC" + Format(Now(), "yyyyMMdd") + ".csv"


            'Génération des fichiers dans le répertoire temporaire
            nFile = FreeFile()
            FileOpen(nFile, strFile, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
            For Each objClt In oColClient
                strStatus = ""
                DisplayStatus("Chargement de " & objClt.code)
                objClt.LoadPreCommande()
                strSCMD_CSV = objClt.oPrecommande.toCSV()
                Print(nFile, strSCMD_CSV)
                ProgressBar1.Increment(1)
            Next
            FileClose(nFile)

            If ckFTP.Checked Then
                DisplayStatus("Transferts des fichiers vers " + Param.getConstante("FTP_HOSTNAME"))
                'Exporter les fichiers générés
                oFTPvinicom = New clsFTPVinicom 'Création avec les paramètres par defaut
                'If oFTPvinicom.connect() Then
                If True Then
                    If (Not oFTPvinicom.IsLockFrom()) Then
                        nficherExporte = oFTPvinicom.uploadFromDir(strFolder)
                        If Not String.IsNullOrEmpty(oFTPvinicom.ErrorDescription) Then
                            DisplayStatus(oFTPvinicom.ErrorDescription())
                        Else
                            DisplayStatus("Fin de transfert des fichiers ")
                            DisplayStatus("Nombre de fichiers exportés : " & (nficherExporte))
                            bReturn = True
                        End If
                    Else
                        DisplayStatus("Serveur internet vérrouillé")
                    End If
                Else
                    DisplayStatus("Connexion impossible (" + Param.getConstante("FTP_USERNAME") + " /" + Param.getConstante("FTP_PASSWORD") + ")")
                End If

                WaitnSeconds(10)

            End If

        Catch ex As Exception
            MsgBox("Erreur" + ex.Message)

        End Try
        Me.Cursor = Cursors.Default

    End Function 'exporter

    Private Shadows Sub DisplayStatus(ByVal strMessage As String)

        Dim oStatus As clsExportstatus
        oStatus = New clsExportstatus()
        oStatus.statusDate = Now()
        oStatus.statusMessage = strMessage
        m_bsrcStatus.Add(oStatus)
        dgvStatus.Refresh()
        If dgvStatus.Rows.Count > dgvStatus.DisplayedRowCount(True) Then
            dgvStatus.FirstDisplayedScrollingRowIndex = dgvStatus.Rows.Count - dgvStatus.DisplayedRowCount(True) + 1
        End If
        Log(strMessage)
    End Sub
    Private Sub Log(ByVal strMessage As String)
        System.IO.File.AppendAllText("./vini_internet.trace", Now() + " " + strMessage + vbCrLf)
        Trace.WriteLine(Now() + " " + strMessage)
    End Sub
#End Region
#Region "Gestion des Evenements"
    Private Sub cbExporter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExporter.Click
        Call exporter()
    End Sub

#End Region

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Génerer déclenche L'evenement Updated
        'AddHandler cbAppliquer.Click, AddressOf ControlUpdated
        'AddHandler cbGenerer.Click, AddressOf ControlUpdated
    End Sub


    Private Sub tbCodeFournisseur_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCodeClient.TextChanged

    End Sub
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class
