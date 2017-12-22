Imports vini_DB
Public Class frmExportQuadra
    Inherits FrmVinicom
    Implements IObservateur

    Protected m_oExportQuadra As ExportQuadra
    Private m_bLoad As Boolean = True



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
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents dgvSCmd As System.Windows.Forms.DataGridView
    Friend WithEvents StatusDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StatusMessageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ckSaveScmd As System.Windows.Forms.CheckBox
    Friend WithEvents rbBonAFactClient As System.Windows.Forms.RadioButton
    Friend WithEvents rbBonAchatFourn As System.Windows.Forms.RadioButton
    Friend WithEvents m_bsrcExportQuadra As BindingSource
    Friend WithEvents m_bsrcListCMD As BindingSource
    Friend WithEvents Selected As DataGridViewCheckBoxColumn
    Friend WithEvents CodeDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents TiersRSColumn As DataGridViewTextBoxColumn
    Friend WithEvents FournisseurRSColumn As DataGridViewTextBoxColumn
    Friend WithEvents totalHT As DataGridViewTextBoxColumn
    Friend WithEvents dateCommande As DataGridViewTextBoxColumn
    Friend WithEvents Origine As DataGridViewTextBoxColumn
    Friend WithEvents typeDonnee As DataGridViewTextBoxColumn
    Friend WithEvents tbExportQuadraFolder As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dtdateFin = New System.Windows.Forms.DateTimePicker()
        Me.m_bsrcExportQuadra = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label8 = New System.Windows.Forms.Label()
        Me.dtDatedeb = New System.Windows.Forms.DateTimePicker()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cbAfficher = New System.Windows.Forms.Button()
        Me.cbExporter = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.dgvSCmd = New System.Windows.Forms.DataGridView()
        Me.Selected = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TiersRSColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FournisseurRSColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.totalHT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dateCommande = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Origine = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.typeDonnee = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_bsrcListCMD = New System.Windows.Forms.BindingSource(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.tbExportQuadraFolder = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ckSaveScmd = New System.Windows.Forms.CheckBox()
        Me.rbBonAFactClient = New System.Windows.Forms.RadioButton()
        Me.rbBonAchatFourn = New System.Windows.Forms.RadioButton()
        CType(Me.m_bsrcExportQuadra, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSCmd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcListCMD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtdateFin
        '
        Me.dtdateFin.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcExportQuadra, "dateFin", True))
        Me.dtdateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdateFin.Location = New System.Drawing.Point(232, 32)
        Me.dtdateFin.Name = "dtdateFin"
        Me.dtdateFin.Size = New System.Drawing.Size(88, 20)
        Me.dtdateFin.TabIndex = 1
        '
        'm_bsrcExportQuadra
        '
        Me.m_bsrcExportQuadra.DataSource = GetType(vini_DB.ExportQuadra)
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
        Me.dtDatedeb.DataBindings.Add(New System.Windows.Forms.Binding("Value", Me.m_bsrcExportQuadra, "dateDeb", True))
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
        Me.dgvSCmd.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvSCmd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSCmd.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Selected, Me.CodeDataGridViewTextBoxColumn, Me.TiersRSColumn, Me.FournisseurRSColumn, Me.totalHT, Me.dateCommande, Me.Origine, Me.typeDonnee})
        Me.dgvSCmd.DataSource = Me.m_bsrcListCMD
        Me.dgvSCmd.Location = New System.Drawing.Point(8, 125)
        Me.dgvSCmd.Name = "dgvSCmd"
        Me.dgvSCmd.Size = New System.Drawing.Size(848, 274)
        Me.dgvSCmd.TabIndex = 133
        '
        'Selected
        '
        Me.Selected.DataPropertyName = "Selected"
        Me.Selected.FillWeight = 30.96447!
        Me.Selected.HeaderText = "Exp"
        Me.Selected.Name = "Selected"
        '
        'CodeDataGridViewTextBoxColumn
        '
        Me.CodeDataGridViewTextBoxColumn.DataPropertyName = "code"
        Me.CodeDataGridViewTextBoxColumn.FillWeight = 96.50592!
        Me.CodeDataGridViewTextBoxColumn.HeaderText = "code"
        Me.CodeDataGridViewTextBoxColumn.Name = "CodeDataGridViewTextBoxColumn"
        '
        'TiersRSColumn
        '
        Me.TiersRSColumn.DataPropertyName = "TiersRS"
        Me.TiersRSColumn.FillWeight = 96.50592!
        Me.TiersRSColumn.HeaderText = "Client"
        Me.TiersRSColumn.Name = "TiersRSColumn"
        Me.TiersRSColumn.ReadOnly = True
        '
        'FournisseurRSColumn
        '
        Me.FournisseurRSColumn.DataPropertyName = "FournisseurRS"
        Me.FournisseurRSColumn.HeaderText = "Producteur"
        Me.FournisseurRSColumn.Name = "FournisseurRSColumn"
        Me.FournisseurRSColumn.ReadOnly = True
        '
        'totalHT
        '
        Me.totalHT.DataPropertyName = "totalHT"
        DataGridViewCellStyle1.Format = "C2"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.totalHT.DefaultCellStyle = DataGridViewCellStyle1
        Me.totalHT.FillWeight = 96.50592!
        Me.totalHT.HeaderText = "totalHT"
        Me.totalHT.Name = "totalHT"
        '
        'dateCommande
        '
        Me.dateCommande.DataPropertyName = "dateCommande"
        DataGridViewCellStyle2.Format = "d"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.dateCommande.DefaultCellStyle = DataGridViewCellStyle2
        Me.dateCommande.FillWeight = 96.50592!
        Me.dateCommande.HeaderText = "dateCommande"
        Me.dateCommande.Name = "dateCommande"
        Me.dateCommande.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        '
        'Origine
        '
        Me.Origine.DataPropertyName = "Origine"
        Me.Origine.FillWeight = 96.50592!
        Me.Origine.HeaderText = "Origine"
        Me.Origine.Name = "Origine"
        Me.Origine.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        '
        'typeDonnee
        '
        Me.typeDonnee.DataPropertyName = "typeDonnee"
        Me.typeDonnee.FillWeight = 96.50592!
        Me.typeDonnee.HeaderText = "typeDonnee"
        Me.typeDonnee.Name = "typeDonnee"
        Me.typeDonnee.ReadOnly = True
        Me.typeDonnee.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        '
        'm_bsrcListCMD
        '
        Me.m_bsrcListCMD.DataMember = "ListCmd"
        Me.m_bsrcListCMD.DataSource = Me.m_bsrcExportQuadra
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'tbExportQuadraFolder
        '
        Me.tbExportQuadraFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbExportQuadraFolder.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.m_bsrcExportQuadra, "folder", True))
        Me.tbExportQuadraFolder.Location = New System.Drawing.Point(533, 61)
        Me.tbExportQuadraFolder.Name = "tbExportQuadraFolder"
        Me.tbExportQuadraFolder.Size = New System.Drawing.Size(323, 20)
        Me.tbExportQuadraFolder.TabIndex = 139
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(395, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 13)
        Me.Label2.TabIndex = 140
        Me.Label2.Text = "Répertoire export Quadra :"
        '
        'ckSaveScmd
        '
        Me.ckSaveScmd.AutoSize = True
        Me.ckSaveScmd.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Me.m_bsrcExportQuadra, "bSaveCmd", True))
        Me.ckSaveScmd.Location = New System.Drawing.Point(533, 6)
        Me.ckSaveScmd.Name = "ckSaveScmd"
        Me.ckSaveScmd.Size = New System.Drawing.Size(169, 17)
        Me.ckSaveScmd.TabIndex = 141
        Me.ckSaveScmd.Text = "SaveAutoDesSousCommande"
        Me.ckSaveScmd.UseVisualStyleBackColor = True
        '
        'rbBonAFactClient
        '
        Me.rbBonAFactClient.AutoSize = True
        Me.rbBonAFactClient.Checked = True
        Me.rbBonAFactClient.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Me.m_bsrcExportQuadra, "isExportBafClient", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.rbBonAFactClient.Location = New System.Drawing.Point(398, 30)
        Me.rbBonAFactClient.Name = "rbBonAFactClient"
        Me.rbBonAFactClient.Size = New System.Drawing.Size(124, 17)
        Me.rbBonAFactClient.TabIndex = 142
        Me.rbBonAFactClient.TabStop = True
        Me.rbBonAFactClient.Text = "Bon a Facturer Client"
        Me.rbBonAFactClient.UseVisualStyleBackColor = True
        '
        'rbBonAchatFourn
        '
        Me.rbBonAchatFourn.AutoSize = True
        Me.rbBonAchatFourn.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Me.m_bsrcExportQuadra, "isExportBaFournisseur", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.rbBonAchatFourn.Location = New System.Drawing.Point(533, 32)
        Me.rbBonAchatFourn.Name = "rbBonAchatFourn"
        Me.rbBonAchatFourn.Size = New System.Drawing.Size(132, 17)
        Me.rbBonAchatFourn.TabIndex = 143
        Me.rbBonAchatFourn.Text = "Bon Achat Fournisseur"
        Me.rbBonAchatFourn.UseVisualStyleBackColor = True
        '
        'frmExportQuadra
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(868, 470)
        Me.Controls.Add(Me.rbBonAchatFourn)
        Me.Controls.Add(Me.rbBonAFactClient)
        Me.Controls.Add(Me.ckSaveScmd)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbExportQuadraFolder)
        Me.Controls.Add(Me.dgvSCmd)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.cbExporter)
        Me.Controls.Add(Me.dtdateFin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dtDatedeb)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cbAfficher)
        Me.Name = "frmExportQuadra"
        Me.Text = "Exportation  des bons à facturer vers Quadra"
        CType(Me.m_bsrcExportQuadra, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSCmd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcListCMD, System.ComponentModel.ISupportInitialize).EndInit()
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
        ckSaveScmd.Enabled = True
        cbAfficher.Enabled = True
        tbExportQuadraFolder.Enabled = True
        rbBonAchatFourn.Enabled = True
        rbBonAFactClient.Enabled = True
        dgvSCmd.Enabled = True

    End Sub


#End Region

#Region "Methodes privées"
    Private Sub initFenetre()


        Dim MoisSuivant As Date
        Dim premierMoisSuivant As Date
        Dim dernierMoisCourant As Date
        Dim dDeb As Date

        'Date de Début = 01 du mois Courant
        dDeb = "01/" & Now.Month() & "/" & Now.Year
        MoisSuivant = DateAdd(DateInterval.Month, +1, dDeb)
        premierMoisSuivant = "01/" & MoisSuivant.Month() & "/" & MoisSuivant.Year()
        dernierMoisCourant = DateAdd(DateInterval.Day, -1, premierMoisSuivant)


        m_oExportQuadra = New ExportQuadra(dDeb, dernierMoisCourant, vncTypeExportQuadra.vncExportBafClient, Param.getConstante("CST_EXPORT_COMPTA_PATH"), True)
        m_bsrcExportQuadra.ResetBindings(False)
        FournisseurRSColumn.Visible = False
        cbExporter.Enabled = False

        m_bsrcExportQuadra.Clear()
        m_bsrcExportQuadra.Add(m_oExportQuadra)

        ckSaveScmd.Checked = True
#If DEBUG Then
        ckSaveScmd.Visible = True
#Else
        ckSaveScmd.visible = false
#End If


    End Sub

    ''' <summary>
    ''' Récupération de la liste des Bons à facturer à exporter vers Quadra
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getListeScmd() As Boolean
        Dim bReturn As Boolean
        debAffiche()
        setcursorWait()
        Try
            m_oExportQuadra.LoadListCmd()
            bReturn = True
        Catch ex As Exception
            bReturn = False
            Log("frmExportQuadra.getListScmd() ERR" + ex.Message)

        End Try
        finAffiche()
        restoreCursor()
        Return bReturn
    End Function



#Region "Export"
    ''' <summary>
    ''' Export vers quadra
    '''     Appel la Methode ExportBafQuadra au travers D'un backGroundWorker
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ExportertVersQuadra()
        Me.ProgressBar1.Minimum = 0
        Me.ProgressBar1.Maximum = (m_oExportQuadra.ListCmd.Count * 2) + 10
        Me.ProgressBar1.Value = Me.ProgressBar1.Minimum

        setcursorWait()
        BackgroundWorker1.RunWorkerAsync()
        restoreCursor()

    End Sub
#Region "Methodes backGoundWorker"
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        exporter()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            ProgressBar1.Value = ProgressBar1.Value + 1
        Catch ex As Exception

        End Try
    End Sub
#End Region

    Private Function exporter() As Boolean

        'je m'ajoute commme obervateur de l'évenement
        m_oExportQuadra.AjouteObservateur(Me)

        'J'execute le traitement
        Dim bReturn As Boolean = m_oExportQuadra.ExportBaf()
        If bReturn Then
            Dim oResult As DialogResult = vbYes
            oResult = MessageBox.Show("Export réalisé dans " & m_oExportQuadra.folder & vbNewLine & " Voulez-vous valider l'export ?", "Validation de l'export", MessageBoxButtons.YesNo)
            If oResult = DialogResult.Yes Then
                bReturn = m_oExportQuadra.ValiderExportBaf()
            End If
        End If
        Return bReturn
    End Function 'exporter

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

    Private Sub cbAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        If getListeScmd() Then
            m_bsrcExportQuadra.ResetBindings(False)
        End If
        If (m_oExportQuadra.ListCmd.Count > 0) Then
            cbExporter.Enabled = True
        End If

    End Sub

    Private Sub cbExporter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExporter.Click
        Call ExportertVersQuadra()
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


    Private Sub rbBonAFactClient_CheckedChanged(sender As Object, e As EventArgs) Handles rbBonAFactClient.CheckedChanged
        If Not m_bLoad Then
            If rbBonAFactClient.Checked Then
                Me.TiersRSColumn.Visible = True
                Me.FournisseurRSColumn.Visible = False
                m_bsrcListCMD.Clear()

            End If
        End If

    End Sub

    Private Sub rbBonAchatFourn_CheckedChanged(sender As Object, e As EventArgs) Handles rbBonAchatFourn.CheckedChanged
        If Not m_bLoad Then
            If rbBonAchatFourn.Checked Then
                Me.TiersRSColumn.Visible = False
                Me.FournisseurRSColumn.Visible = True
                m_bsrcListCMD.Clear()

            End If
        End If


    End Sub
End Class
