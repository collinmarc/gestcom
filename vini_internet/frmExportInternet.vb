Imports vini_DB
Public Class frmExportInternet
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
    Friend WithEvents dtdateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtDatedeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents cbExporter As System.Windows.Forms.Button
    Friend WithEvents ckPDFs As System.Windows.Forms.CheckBox
    Friend WithEvents ckFTP As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbCodeFournisseur As System.Windows.Forms.TextBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents dgvSCmd As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcSCMD As System.Windows.Forms.BindingSource
    Friend WithEvents m_bsrcStatus As System.Windows.Forms.BindingSource
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbNbScmd As System.Windows.Forms.TextBox
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FournisseurCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalHT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dateCommande As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusMessageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents tbNbreTheorique As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dgvStatus As System.Windows.Forms.DataGridView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.cbExporter = New System.Windows.Forms.Button
        Me.ckPDFs = New System.Windows.Forms.CheckBox
        Me.ckFTP = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbCodeFournisseur = New System.Windows.Forms.TextBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.dgvSCmd = New System.Windows.Forms.DataGridView
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FournisseurCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TiersRS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.totalHT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dateCommande = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.m_bsrcSCMD = New System.Windows.Forms.BindingSource(Me.components)
        Me.m_bsrcStatus = New System.Windows.Forms.BindingSource(Me.components)
        Me.dgvStatus = New System.Windows.Forms.DataGridView
        Me.StatusDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.StatusMessageDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbNbScmd = New System.Windows.Forms.TextBox
        Me.tbNbreTheorique = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        CType(Me.dgvSCmd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcSCMD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtdateFin
        '
        Me.dtdateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdateFin.Location = New System.Drawing.Point(232, 32)
        Me.dtdateFin.Name = "dtdateFin"
        Me.dtdateFin.Size = New System.Drawing.Size(88, 20)
        Me.dtdateFin.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(176, 16)
        Me.Label8.TabIndex = 107
        Me.Label8.Text = "date de fin"
        '
        'dtDatedeb
        '
        Me.dtDatedeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDatedeb.Location = New System.Drawing.Point(232, 8)
        Me.dtDatedeb.Name = "dtDatedeb"
        Me.dtDatedeb.Size = New System.Drawing.Size(88, 20)
        Me.dtDatedeb.TabIndex = 0
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 8)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(176, 16)
        Me.Label14.TabIndex = 106
        Me.Label14.Text = "date de début "
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.BackColor = System.Drawing.SystemColors.Control
        Me.cbAfficher.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAfficher.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAfficher.Location = New System.Drawing.Point(8, 96)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAfficher.Size = New System.Drawing.Size(352, 23)
        Me.cbAfficher.TabIndex = 3
        Me.cbAfficher.Text = "A&fficher"
        Me.cbAfficher.UseVisualStyleBackColor = False
        '
        'cbExporter
        '
        Me.cbExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbExporter.Location = New System.Drawing.Point(366, 249)
        Me.cbExporter.Name = "cbExporter"
        Me.cbExporter.Size = New System.Drawing.Size(64, 24)
        Me.cbExporter.TabIndex = 4
        Me.cbExporter.Text = "Exporter"
        '
        'ckPDFs
        '
        Me.ckPDFs.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckPDFs.Checked = True
        Me.ckPDFs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckPDFs.Location = New System.Drawing.Point(352, 8)
        Me.ckPDFs.Name = "ckPDFs"
        Me.ckPDFs.Size = New System.Drawing.Size(144, 24)
        Me.ckPDFs.TabIndex = 127
        Me.ckPDFs.Text = "Exportation des PDFs"
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
        Me.Label1.Location = New System.Drawing.Point(8, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 129
        Me.Label1.Text = "Code Fournisseur"
        '
        'tbCodeFournisseur
        '
        Me.tbCodeFournisseur.Location = New System.Drawing.Point(232, 64)
        Me.tbCodeFournisseur.Name = "tbCodeFournisseur"
        Me.tbCodeFournisseur.Size = New System.Drawing.Size(88, 20)
        Me.tbCodeFournisseur.TabIndex = 2
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(438, 96)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(413, 23)
        Me.ProgressBar1.TabIndex = 132
        '
        'dgvSCmd
        '
        Me.dgvSCmd.AllowUserToAddRows = False
        Me.dgvSCmd.AllowUserToDeleteRows = False
        Me.dgvSCmd.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSCmd.AutoGenerateColumns = False
        Me.dgvSCmd.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSCmd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSCmd.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodeDataGridViewTextBoxColumn, Me.FournisseurCodeDataGridViewTextBoxColumn, Me.TiersRS, Me.totalHT, Me.dateCommande})
        Me.dgvSCmd.DataSource = Me.m_bsrcSCMD
        Me.dgvSCmd.Location = New System.Drawing.Point(8, 125)
        Me.dgvSCmd.Name = "dgvSCmd"
        Me.dgvSCmd.ReadOnly = True
        Me.dgvSCmd.Size = New System.Drawing.Size(352, 290)
        Me.dgvSCmd.TabIndex = 133
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        Me.CodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.CodeDataGridViewTextBoxColumn.Width = 56
        '
        'FournisseurCodeDataGridViewTextBoxColumn
        '
        Me.FournisseurCodeDataGridViewTextBoxColumn.DataPropertyName = "FournisseurCode"
        Me.FournisseurCodeDataGridViewTextBoxColumn.HeaderText = "Producteur"
        Me.FournisseurCodeDataGridViewTextBoxColumn.Name = "FournisseurCodeDataGridViewTextBoxColumn"
        Me.FournisseurCodeDataGridViewTextBoxColumn.ReadOnly = True
        Me.FournisseurCodeDataGridViewTextBoxColumn.Width = 84
        '
        'TiersRS
        '
        Me.TiersRS.DataPropertyName = "TiersRS"
        Me.TiersRS.HeaderText = "Client"
        Me.TiersRS.Name = "TiersRS"
        Me.TiersRS.ReadOnly = True
        Me.TiersRS.Width = 58
        '
        'totalHT
        '
        Me.totalHT.DataPropertyName = "totalHT"
        DataGridViewCellStyle1.Format = "C2"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.totalHT.DefaultCellStyle = DataGridViewCellStyle1
        Me.totalHT.HeaderText = "totalHT"
        Me.totalHT.Name = "totalHT"
        Me.totalHT.ReadOnly = True
        Me.totalHT.Width = 67
        '
        'dateCommande
        '
        Me.dateCommande.DataPropertyName = "dateCommande"
        DataGridViewCellStyle2.Format = "d"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.dateCommande.DefaultCellStyle = DataGridViewCellStyle2
        Me.dateCommande.HeaderText = "dateCommande"
        Me.dateCommande.Name = "dateCommande"
        Me.dateCommande.ReadOnly = True
        Me.dateCommande.Width = 106
        '
        'm_bsrcSCMD
        '
        Me.m_bsrcSCMD.DataSource = GetType(vini_DB.SousCommande)
        '
        'm_bsrcStatus
        '
        Me.m_bsrcStatus.DataSource = GetType(vini_internet.clsExportstatus)
        '
        'dgvStatus
        '
        Me.dgvStatus.AllowUserToAddRows = False
        Me.dgvStatus.AllowUserToDeleteRows = False
        Me.dgvStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvStatus.AutoGenerateColumns = False
        Me.dgvStatus.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStatus.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.StatusDateDataGridViewTextBoxColumn, Me.StatusMessageDataGridViewTextBoxColumn})
        Me.dgvStatus.DataSource = Me.m_bsrcStatus
        Me.dgvStatus.Location = New System.Drawing.Point(438, 126)
        Me.dgvStatus.Name = "dgvStatus"
        Me.dgvStatus.ReadOnly = True
        Me.dgvStatus.Size = New System.Drawing.Size(413, 290)
        Me.dgvStatus.TabIndex = 134
        '
        'StatusDateDataGridViewTextBoxColumn
        '
        Me.StatusDateDataGridViewTextBoxColumn.DataPropertyName = "statusDate"
        DataGridViewCellStyle3.Format = "G"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.StatusDateDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle3
        Me.StatusDateDataGridViewTextBoxColumn.HeaderText = "Date"
        Me.StatusDateDataGridViewTextBoxColumn.Name = "StatusDateDataGridViewTextBoxColumn"
        Me.StatusDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.StatusDateDataGridViewTextBoxColumn.Width = 55
        '
        'StatusMessageDataGridViewTextBoxColumn
        '
        Me.StatusMessageDataGridViewTextBoxColumn.DataPropertyName = "statusMessage"
        Me.StatusMessageDataGridViewTextBoxColumn.HeaderText = "Message"
        Me.StatusMessageDataGridViewTextBoxColumn.Name = "StatusMessageDataGridViewTextBoxColumn"
        Me.StatusMessageDataGridViewTextBoxColumn.ReadOnly = True
        Me.StatusMessageDataGridViewTextBoxColumn.Width = 75
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(367, 187)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 13)
        Me.Label2.TabIndex = 135
        Me.Label2.Text = "Nbre lu"
        '
        'tbNbScmd
        '
        Me.tbNbScmd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbNbScmd.Enabled = False
        Me.tbNbScmd.Location = New System.Drawing.Point(366, 203)
        Me.tbNbScmd.Name = "tbNbScmd"
        Me.tbNbScmd.Size = New System.Drawing.Size(64, 20)
        Me.tbNbScmd.TabIndex = 136
        '
        'tbNbreTheorique
        '
        Me.tbNbreTheorique.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbNbreTheorique.Enabled = False
        Me.tbNbreTheorique.Location = New System.Drawing.Point(370, 157)
        Me.tbNbreTheorique.Name = "tbNbreTheorique"
        Me.tbNbreTheorique.Size = New System.Drawing.Size(60, 20)
        Me.tbNbreTheorique.TabIndex = 137
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(370, 138)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 138
        Me.Label3.Text = "Théorique"
        '
        'frmExportInternet
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(863, 428)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbNbreTheorique)
        Me.Controls.Add(Me.tbNbScmd)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dgvStatus)
        Me.Controls.Add(Me.dgvSCmd)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.tbCodeFournisseur)
        Me.Controls.Add(Me.ckFTP)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ckPDFs)
        Me.Controls.Add(Me.cbExporter)
        Me.Controls.Add(Me.dtdateFin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtDatedeb)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cbAfficher)
        Me.Name = "frmExportInternet"
        Me.Text = "Exportation  des bons à facturer vers Internet"
        CType(Me.dgvSCmd, System.ComponentModel.ISupportInitialize).EndInit()
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

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)
        dtDatedeb.Enabled = True
        dtdateFin.Enabled = True
        cbAfficher.Enabled = True
        tbCodeFournisseur.Enabled = True
        dgvStatus.Enabled = True

    End Sub


#End Region

#Region "Methodes privées"
    Private Sub initFenetre()
        dtDatedeb.Value = Now()
        dtdateFin.Value = Now()
        m_colCommandes = New Collection
        ckPDFs.Enabled = True
        ckPDFs.Checked = True
        cbExporter.Enabled = False

    End Sub

    Private Sub afficheListeScmd()

        setcursorWait()
        debAffiche()

        m_bsrcSCMD.Clear()

        For Each objSCMD As SousCommande In m_colCommandes
            m_bsrcSCMD.Add(objSCMD)
        Next
        Dim objcommand As System.Data.OleDb.OleDbCommand
        objcommand = New OleDb.OleDbCommand
        objcommand.Connection = Persist.dbConnection.Connection
        objcommand.CommandText = "SELECT     COUNT(SOUSCOMMANDE.SCMD_ID) AS NBRE FROM SOUSCOMMANDE INNER JOIN FOURNISSEUR ON SOUSCOMMANDE.SCMD_FRN_ID = FOURNISSEUR.FRN_ID" & _
                                 " WHERE     SOUSCOMMANDE.SCMD_ETAT = 10 AND  FOURNISSEUR.FRN_BEXP_INTERNET = 1 " & _
                                 " AND SOUSCOMMANDE.SCMD_DATE >= '" & dtDatedeb.Value.ToShortDateString() & "' AND " & _
                                " SOUSCOMMANDE.SCMD_DATE <= '" & dtdateFin.Value.ToShortDateString() & "'"
        If Not String.IsNullOrEmpty(tbCodeFournisseur.Text) Then
            objcommand.CommandText = objcommand.CommandText & _
                                     " AND FOURNISSEUR.FRN_CODE LIKE '" & tbCodeFournisseur.Text & "'"
        End If
        objcommand.Connection.Open()
        tbNbreTheorique.Text = objcommand.ExecuteScalar().ToString()
        objcommand.Connection.Close()


        tbNbScmd.Text = m_bsrcSCMD.Count


        finAffiche()
        restoreCursor()

    End Sub 'AfficheListeFactTRP

    Private Function setListeScmd() As Boolean
        Dim ddeb As Date
        Dim dfin As Date
        Dim strCodeFourn As String
        Dim col As Collection
        Dim bReturn As Boolean
        debAffiche()
        setcursorWait()
        Try

            ddeb = dtDatedeb.Value.ToShortDateString
            dfin = dtdateFin.Value.ToShortDateString
            strCodeFourn = tbCodeFournisseur.Text
            col = SousCommande.getListeAExporter(ddeb, dfin, strCodeFourn)
            'Recupération de la liste des sous commande non Exportées
            If col Is Nothing Then
                bReturn = False
            Else
                m_colCommandes = col
                bReturn = True
            End If

        Catch ex As Exception
            bReturn = False
            Debug.Assert(bReturn, ex.ToString)
        End Try
        finAffiche()
        restoreCursor()
        Return bReturn
    End Function




    Private Function sauvegarder() As Boolean
        Dim objSCMD As SousCommande
        Dim bReturn As Boolean

        bReturn = False
        For Each objSCMD In m_colCommandes
            bReturn = bReturn And objSCMD.Save()
            Debug.Assert(bReturn, SousCommande.getErreur())
        Next

        Return bReturn
    End Function 'sauvegarderFactures

    Private Function exporter() As Boolean
        Debug.Assert(Not m_colCommandes Is Nothing)

        Dim objSCMD As SousCommande
        Dim bExportOK As Boolean = True
        Dim bReturn As Boolean
        Dim strFile As String
        Dim strSCMD_CSV As String
        Dim nFile As Integer
        Dim strPDFFileName As String
        Dim strFolder As String
        Dim oFTPvinicom As clsFTPVinicom
        Dim nSousCommandesPreparees As Integer
        Dim nSousCommandesExportees As Integer
        Dim strStatus As String

        '        Dim odlgInternet As dlgInternet
        'Suppression - creation du répertoire temporaire
        Try
            bReturn = False
            nSousCommandesPreparees = 0
            ProgressBar1.Maximum = m_colCommandes.Count
            ProgressBar1.Value = 0
            setcursorWait()
            dgvStatus.Rows.Clear()
            strFolder = IMPORT_DIRECTORY & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second
            If My.Computer.FileSystem.DirectoryExists(strFolder) Then
                My.Computer.FileSystem.DeleteDirectory(strFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            My.Computer.FileSystem.CreateDirectory(strFolder)

            strFile = strFolder & "/" & EXPORTFTP_FILENAME


            'Génération des fichiers dans le répertoire temporaire
            nFile = FreeFile()
            FileOpen(nFile, strFile, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
            For Each objSCMD In m_colCommandes
                strStatus = ""
                Log("Chargement de " & objSCMD.code)
                objSCMD.load()
                objSCMD.loadcolLignes()
                If objSCMD.colLignes.Count > 0 Then
                    strSCMD_CSV = objSCMD.toCSV()
                    strStatus = strStatus + " CSV OK"
                    Print(nFile, strSCMD_CSV)
                    If ckPDFs.Checked Then
                        strPDFFileName = strFolder & "/" & objSCMD.code & ".PDF"
                        If objSCMD.genererPDF(PATHTOREPORTS, strPDFFileName) Then
                            strStatus = strStatus + " PDF OK"
                        Else
                            DisplayStatus("Chargement de " & objSCMD.code & " PDF ERREUR" & objSCMD.getErreur())
                        End If

                    End If
                    nSousCommandesPreparees = nSousCommandesPreparees + 1
                    objSCMD.releasecolLignes()

                    'changement d'état de la sous Commandes
                    objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
                    objSCMD.bExportInternet = True
                    Log(strStatus)
                Else
                    DisplayStatus("Pas de ligne dans la sous-commande " + objSCMD.code)
                End If
                ProgressBar1.Increment(1)

            Next
            FileClose(nFile)

            DisplayStatus("Nombre de commandes préparées : " & nSousCommandesPreparees)
            If ckFTP.Checked Then
                DisplayStatus("Transferts des fichiers vers " + Param.getConstante("FTP_HOSTNAME"))
                'Exporter les fichiers générés
                oFTPvinicom = New clsFTPVinicom 'Création avec les paramètres par defaut
                'If oFTPvinicom.connect() Then
                If True Then
                    If (Not oFTPvinicom.IsLockFrom()) Then
                        nSousCommandesExportees = oFTPvinicom.uploadFromDir(strFolder)
                        If Not String.IsNullOrEmpty(oFTPvinicom.ErrorDescription) Then
                            DisplayStatus(oFTPvinicom.ErrorDescription())
                        Else
                            DisplayStatus("Fin de transfert des fichiers ")
                            DisplayStatus("Nombre de fichiers exportés : " & (nSousCommandesExportees - 1) & "+1")
                            bReturn = True
                        End If
                    Else
                        DisplayStatus("Serveur internet vérrouillé")
                    End If
                Else
                    DisplayStatus("Connexion impossible (" + Param.getConstante("FTP_USERNAME") + " /" + Param.getConstante("FTP_PASSWORD") + ")")
                End If

                WaitnSeconds(10)
                ActiverImportBAF()

                'oFTPvinicom.disconnect()
            End If
            If bReturn Then
                DisplayStatus("Validation des sous commandes (changement d'état)")
                For Each objSCMD In m_colCommandes
                    objSCMD.Save()
                Next
                DisplayStatus("Validation terminée")
            End If

            cbExporter.Enabled = False

        Catch ex As Exception
            MsgBox("Erreur" + ex.Message)

        End Try
        Me.Cursor = Cursors.Default

    End Function 'exporter

    Private Sub ActiverImportBAF()
        Dim odlg As dlgInternet = New dlgInternet()
        dlgInternet.ShowDialog()
    End Sub
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
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        If setListeScmd() Then
            afficheListeScmd()
        End If
        If (m_colCommandes.Count > 0) Then
            dgvSCmd.Enabled = True
            cbExporter.Enabled = True
        End If

    End Sub





#End Region

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Génerer déclenche L'evenement Updated
        'AddHandler cbAppliquer.Click, AddressOf ControlUpdated
        'AddHandler cbGenerer.Click, AddressOf ControlUpdated
    End Sub

    Private Sub cbExporter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExporter.Click
        Call exporter()
    End Sub

End Class
