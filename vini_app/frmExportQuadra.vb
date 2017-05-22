Imports vini_DB
Public Class frmExportQuadra
    Inherits FrmVinicom
    Protected m_colCommandes As List(Of SousCommande)

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbCodeFournisseur As System.Windows.Forms.TextBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents dgvSCmd As System.Windows.Forms.DataGridView
    Friend WithEvents m_bsrcSCMD As System.Windows.Forms.BindingSource
    Friend WithEvents CodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FournisseurCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TiersRS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents totalHT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dateCommande As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusMessageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ckSaveScmd As System.Windows.Forms.CheckBox
    Friend WithEvents tbExportQuadraFolder As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.cbExporter = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbCodeFournisseur = New System.Windows.Forms.TextBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.dgvSCmd = New System.Windows.Forms.DataGridView()
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FournisseurCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TiersRS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.totalHT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dateCommande = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_bsrcSCMD = New System.Windows.Forms.BindingSource(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.tbExportQuadraFolder = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ckSaveScmd = New System.Windows.Forms.CheckBox()
        CType(Me.dgvSCmd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcSCMD, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.cbAfficher.Size = New System.Drawing.Size(848, 23)
        Me.cbAfficher.TabIndex = 3
        Me.cbAfficher.Text = "A&fficher"
        Me.cbAfficher.UseVisualStyleBackColor = False
        '
        'cbExporter
        '
        Me.cbExporter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbExporter.Location = New System.Drawing.Point(8, 405)
        Me.cbExporter.Name = "cbExporter"
        Me.cbExporter.Size = New System.Drawing.Size(848, 24)
        Me.cbExporter.TabIndex = 4
        Me.cbExporter.Text = "Exporter"
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
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(8, 435)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(848, 23)
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
        Me.dgvSCmd.Size = New System.Drawing.Size(848, 274)
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
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'tbExportQuadraFolder
        '
        Me.tbExportQuadraFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbExportQuadraFolder.Location = New System.Drawing.Point(533, 31)
        Me.tbExportQuadraFolder.Name = "tbExportQuadraFolder"
        Me.tbExportQuadraFolder.Size = New System.Drawing.Size(323, 20)
        Me.tbExportQuadraFolder.TabIndex = 139
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(395, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 13)
        Me.Label2.TabIndex = 140
        Me.Label2.Text = "Répertoire export Quadra :"
        '
        'ckSaveScmd
        '
        Me.ckSaveScmd.AutoSize = True
        Me.ckSaveScmd.Location = New System.Drawing.Point(533, 6)
        Me.ckSaveScmd.Name = "ckSaveScmd"
        Me.ckSaveScmd.Size = New System.Drawing.Size(169, 17)
        Me.ckSaveScmd.TabIndex = 141
        Me.ckSaveScmd.Text = "SaveAutoDesSousCommande"
        Me.ckSaveScmd.UseVisualStyleBackColor = True
        '
        'frmExportQuadra
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(868, 470)
        Me.Controls.Add(Me.ckSaveScmd)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbExportQuadraFolder)
        Me.Controls.Add(Me.dgvSCmd)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.tbCodeFournisseur)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbExporter)
        Me.Controls.Add(Me.dtdateFin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtDatedeb)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cbAfficher)
        Me.Name = "frmExportQuadra"
        Me.Text = "Exportation  des bons à facturer vers Quadra"
        CType(Me.dgvSCmd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcSCMD, System.ComponentModel.ISupportInitialize).EndInit()
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
        tbExportQuadraFolder.Enabled = True

    End Sub


#End Region

#Region "Methodes privées"
    Private Sub initFenetre()
        Dim MoisSuivant As Date
        Dim premierMoisSuivant As Date
        Dim dernierMoisCourant As Date
        'Date de Début = 01 du mois Courant
        dtDatedeb.Value = "01/" & Now.Month() & "/" & Now.Year
        MoisSuivant = DateAdd(DateInterval.Month, +1, Now())
        premierMoisSuivant = "01/" & MoisSuivant.Month() & "/" & MoisSuivant.Year()
        dernierMoisCourant = DateAdd(DateInterval.Day, -1, premierMoisSuivant)

        dtdateFin.Value = dernierMoisCourant
        m_colCommandes = New List(Of SousCommande)
        cbExporter.Enabled = False
        tbExportQuadraFolder.Text = Param.getConstante("CST_EXPORT_COMPTA_PATH")

        ckSaveScmd.Checked = True
        ckSaveScmd.Enabled = True
    End Sub

    Private Sub afficheListeScmd()

        setcursorWait()
        debAffiche()

        m_bsrcSCMD.Clear()

        For Each objSCMD As SousCommande In m_colCommandes
            m_bsrcSCMD.Add(objSCMD)
        Next
        finAffiche()
        restoreCursor()

    End Sub 'AfficheListeFactTRP
    ''' <summary>
    ''' Récupération de la liste des Bons à facturer à exporter vers Quadra
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getListeScmd() As Boolean
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
            col = SousCommande.getListeAExporterQuadra(ddeb, dfin, strCodeFourn)
            'Recupération de la liste des sous commande 
            If col Is Nothing Then
                bReturn = False
            Else
                m_colCommandes.Clear()
                For Each oScmd As SousCommande In col
                    m_colCommandes.Add(oScmd)
                Next
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





    Private Sub ExportertBafVersQuadra()
        Me.ProgressBar1.Minimum = 0
        Me.ProgressBar1.Maximum = m_colCommandes.Count + 5
        setcursorWait()
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        exporter()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = ProgressBar1.Value + 1
    End Sub


    Private Function exporter() As Boolean
        Debug.Assert(Not m_colCommandes Is Nothing)

        Dim objSCMD As SousCommande
        Dim bExportOK As Boolean = True
        Dim bReturn As Boolean
        Dim strFile As String
        Dim strFolder As String
        Dim ofrn As Fournisseur

        Try
            bReturn = False
            strFolder = tbExportQuadraFolder.Text
            If Not My.Computer.FileSystem.DirectoryExists(strFolder) Then
                My.Computer.FileSystem.CreateDirectory(strFolder)
            End If
            strFile = strFolder & "/Export" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & ".csv"

            BackgroundWorker1.ReportProgress(1)

            'Parcours de la liste des souscommandes et génération du fichier CSV
            For Each objSCMD In m_bsrcSCMD
                objSCMD.load()
                'Normalement inutile car , seul les SousCommande exportables sont chargées
                ofrn = Fournisseur.createandload(objSCMD.Fournisseurid)
                If ofrn.bExportInternet = 2 Then
                    objSCMD.loadcolLignes()
                    If objSCMD.colLignes.Count > 0 Then
                        objSCMD.toCSVQuadra(strFile)
                        objSCMD.ValiderExportQuadra()
                        'Il faut sauvegarder les sous commandes car l'export a été réalisé
                        If ckSaveScmd.Checked Then
                            objSCMD.Save()
                        End If
                    End If
                    BackgroundWorker1.ReportProgress(1)
                End If

            Next

        Catch ex As Exception
            MsgBox("Erreur" + ex.Message)

        End Try
        Me.Cursor = Cursors.Default

    End Function 'exporter

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
        If getListeScmd() Then
            afficheListeScmd()
        End If
        If (m_colCommandes.Count > 0) Then
            cbExporter.Enabled = True
        End If

    End Sub

    Private Sub cbExporter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExporter.Click
        Call ExportertBafVersQuadra()
    End Sub




#End Region

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Génerer déclenche L'evenement Updated
        'AddHandler cbAppliquer.Click, AddressOf ControlUpdated
        'AddHandler cbGenerer.Click, AddressOf ControlUpdated
    End Sub

 
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        ProgressBar1.Value = ProgressBar1.Maximum
        cbExporter.Enabled = False
        restoreCursor()
    End Sub
End Class
