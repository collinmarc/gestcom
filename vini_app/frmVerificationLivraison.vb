Imports vini_DB
Imports System.ComponentModel

Public Class UserState
    Public nNum As Integer
    Public strMsg As String
End Class
Public Class frmVerificationLivraison
    Inherits FrmDonBase
    Protected m_lstCommandes As List(Of CommandeClient)
    Friend WithEvents m_bsrcMsgLivraison As System.Windows.Forms.BindingSource
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnTraiter As System.Windows.Forms.Button
    Friend WithEvents btnQuitter As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbFilePath As System.Windows.Forms.TextBox
    Friend WithEvents btnParcourir As System.Windows.Forms.Button
    Friend WithEvents NumCommandeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateMessageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NbreColisLivreDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MessageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ckAfficherCmdOK As System.Windows.Forms.CheckBox
    Friend WithEvents rbFTP As System.Windows.Forms.RadioButton
    Friend WithEvents rbManuel As System.Windows.Forms.RadioButton
    Friend WithEvents pnlFichierATraiter As System.Windows.Forms.Panel
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Protected m_colFact As ColEvent
    'Protected getElementCourant() As FactTRP

#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

        bTesting = True
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnlFichierATraiter = New System.Windows.Forms.Panel()
        Me.tbFilePath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnParcourir = New System.Windows.Forms.Button()
        Me.rbManuel = New System.Windows.Forms.RadioButton()
        Me.rbFTP = New System.Windows.Forms.RadioButton()
        Me.ckAfficherCmdOK = New System.Windows.Forms.CheckBox()
        Me.btnQuitter = New System.Windows.Forms.Button()
        Me.btnTraiter = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.NumCommandeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DateMessageDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NbreColisLivreDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MessageDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_bsrcMsgLivraison = New System.Windows.Forms.BindingSource(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.pnlFichierATraiter.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.m_bsrcMsgLivraison, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlFichierATraiter
        '
        Me.pnlFichierATraiter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlFichierATraiter.Controls.Add(Me.tbFilePath)
        Me.pnlFichierATraiter.Controls.Add(Me.Label1)
        Me.pnlFichierATraiter.Controls.Add(Me.btnParcourir)
        Me.pnlFichierATraiter.Location = New System.Drawing.Point(12, 58)
        Me.pnlFichierATraiter.Name = "pnlFichierATraiter"
        Me.pnlFichierATraiter.Size = New System.Drawing.Size(989, 38)
        Me.pnlFichierATraiter.TabIndex = 13
        '
        'tbFilePath
        '
        Me.tbFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbFilePath.Location = New System.Drawing.Point(122, 10)
        Me.tbFilePath.Name = "tbFilePath"
        Me.tbFilePath.Size = New System.Drawing.Size(758, 20)
        Me.tbFilePath.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Fichier à traiter : "
        '
        'btnParcourir
        '
        Me.btnParcourir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnParcourir.Location = New System.Drawing.Point(911, 8)
        Me.btnParcourir.Name = "btnParcourir"
        Me.btnParcourir.Size = New System.Drawing.Size(75, 23)
        Me.btnParcourir.TabIndex = 9
        Me.btnParcourir.Text = "Parcourir"
        Me.btnParcourir.UseVisualStyleBackColor = True
        '
        'rbManuel
        '
        Me.rbManuel.AutoSize = True
        Me.rbManuel.Location = New System.Drawing.Point(13, 35)
        Me.rbManuel.Name = "rbManuel"
        Me.rbManuel.Size = New System.Drawing.Size(113, 17)
        Me.rbManuel.TabIndex = 12
        Me.rbManuel.Text = "Traitement Manuel"
        Me.rbManuel.UseVisualStyleBackColor = True
        '
        'rbFTP
        '
        Me.rbFTP.AutoSize = True
        Me.rbFTP.Checked = True
        Me.rbFTP.Location = New System.Drawing.Point(13, 12)
        Me.rbFTP.Name = "rbFTP"
        Me.rbFTP.Size = New System.Drawing.Size(166, 17)
        Me.rbFTP.TabIndex = 11
        Me.rbFTP.TabStop = True
        Me.rbFTP.Text = "Traitement Automatique (FTP)"
        Me.rbFTP.UseVisualStyleBackColor = True
        '
        'ckAfficherCmdOK
        '
        Me.ckAfficherCmdOK.AutoSize = True
        Me.ckAfficherCmdOK.Location = New System.Drawing.Point(13, 102)
        Me.ckAfficherCmdOK.Name = "ckAfficherCmdOK"
        Me.ckAfficherCmdOK.Size = New System.Drawing.Size(226, 17)
        Me.ckAfficherCmdOK.TabIndex = 10
        Me.ckAfficherCmdOK.Text = "Affichage des commandes sans anomalies"
        Me.ckAfficherCmdOK.UseVisualStyleBackColor = True
        '
        'btnQuitter
        '
        Me.btnQuitter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnQuitter.Location = New System.Drawing.Point(930, 598)
        Me.btnQuitter.Name = "btnQuitter"
        Me.btnQuitter.Size = New System.Drawing.Size(75, 23)
        Me.btnQuitter.TabIndex = 6
        Me.btnQuitter.Text = "Quitter"
        Me.btnQuitter.UseVisualStyleBackColor = True
        '
        'btnTraiter
        '
        Me.btnTraiter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTraiter.Location = New System.Drawing.Point(13, 125)
        Me.btnTraiter.Name = "btnTraiter"
        Me.btnTraiter.Size = New System.Drawing.Size(991, 23)
        Me.btnTraiter.TabIndex = 4
        Me.btnTraiter.Text = "Traiter le fichier"
        Me.btnTraiter.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.NumCommandeDataGridViewTextBoxColumn, Me.DateMessageDataGridViewTextBoxColumn, Me.NbreColisLivreDataGridViewTextBoxColumn, Me.MessageDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.m_bsrcMsgLivraison
        Me.DataGridView1.Location = New System.Drawing.Point(13, 209)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(992, 374)
        Me.DataGridView1.TabIndex = 0
        '
        'NumCommandeDataGridViewTextBoxColumn
        '
        Me.NumCommandeDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.NumCommandeDataGridViewTextBoxColumn.DataPropertyName = "NumCommande"
        Me.NumCommandeDataGridViewTextBoxColumn.HeaderText = "Cmd"
        Me.NumCommandeDataGridViewTextBoxColumn.Name = "NumCommandeDataGridViewTextBoxColumn"
        Me.NumCommandeDataGridViewTextBoxColumn.ReadOnly = True
        Me.NumCommandeDataGridViewTextBoxColumn.Width = 53
        '
        'DateMessageDataGridViewTextBoxColumn
        '
        Me.DateMessageDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.DateMessageDataGridViewTextBoxColumn.DataPropertyName = "DateMessage"
        Me.DateMessageDataGridViewTextBoxColumn.HeaderText = "Date"
        Me.DateMessageDataGridViewTextBoxColumn.Name = "DateMessageDataGridViewTextBoxColumn"
        Me.DateMessageDataGridViewTextBoxColumn.ReadOnly = True
        Me.DateMessageDataGridViewTextBoxColumn.Width = 55
        '
        'NbreColisLivreDataGridViewTextBoxColumn
        '
        Me.NbreColisLivreDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.NbreColisLivreDataGridViewTextBoxColumn.DataPropertyName = "NbreColisLivre"
        Me.NbreColisLivreDataGridViewTextBoxColumn.HeaderText = "Colis Livrés"
        Me.NbreColisLivreDataGridViewTextBoxColumn.Name = "NbreColisLivreDataGridViewTextBoxColumn"
        Me.NbreColisLivreDataGridViewTextBoxColumn.ReadOnly = True
        Me.NbreColisLivreDataGridViewTextBoxColumn.Width = 85
        '
        'MessageDataGridViewTextBoxColumn
        '
        Me.MessageDataGridViewTextBoxColumn.DataPropertyName = "Message"
        Me.MessageDataGridViewTextBoxColumn.HeaderText = "Message"
        Me.MessageDataGridViewTextBoxColumn.Name = "MessageDataGridViewTextBoxColumn"
        Me.MessageDataGridViewTextBoxColumn.ReadOnly = True
        '
        'm_bsrcMsgLivraison
        '
        Me.m_bsrcMsgLivraison.DataSource = GetType(vini_DB.MsgLivraison)
        Me.m_bsrcMsgLivraison.Filter = "Resultat<>0"
        Me.m_bsrcMsgLivraison.Sort = ""
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "csv"
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = "Fichier CSV|*.csv|Tous les fichiers|*.*"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(13, 155)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(991, 23)
        Me.ProgressBar1.TabIndex = 15
        '
        'lblMessage
        '
        Me.lblMessage.AutoSize = True
        Me.lblMessage.Location = New System.Drawing.Point(13, 190)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(39, 13)
        Me.lblMessage.TabIndex = 16
        Me.lblMessage.Text = "Label2"
        '
        'frmVerificationLivraison
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1011, 633)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.pnlFichierATraiter)
        Me.Controls.Add(Me.rbManuel)
        Me.Controls.Add(Me.rbFTP)
        Me.Controls.Add(Me.ckAfficherCmdOK)
        Me.Controls.Add(Me.btnQuitter)
        Me.Controls.Add(Me.btnTraiter)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "frmVerificationLivraison"
        Me.Text = "Vérification des informations de livraison"
        Me.pnlFichierATraiter.ResumeLayout(False)
        Me.pnlFichierATraiter.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.m_bsrcMsgLivraison, System.ComponentModel.ISupportInitialize).EndInit()
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
        m_ToolBarSaveEnabled = isfrmUpdated()
    End Sub

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(bEnabled)

    End Sub

    Protected Overrides Function creerElement() As Boolean
        Return False
    End Function
    Protected Shadows Function getElementCourant() As FactTRP
        Return Nothing
    End Function
    ''' <summary>
    ''' Sauvegarde des commandes et des sous commandes modifiées
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function frmSave() As Boolean
        Debug.Assert(m_lstCommandes IsNot Nothing)
        Debug.Assert(m_lstCommandes.Count > 0)
        'Sauvegarde des commandes et des sousCommandes

        For Each oCmd As CommandeClient In m_lstCommandes
            If oCmd.bUpdated Then
                oCmd.bUpdatePrecommande = False 'Pas de Mise à jour des précommandes
                oCmd.save()
            End If
        Next
        setfrmNotUpdated()
        'm_bsrcScmd.Clear()

    End Function

#End Region

#Region "Methodes privées"
    Private Sub initFenetre()

    End Sub
    Private Sub afficheListeCommande()


    End Sub 'AfficheListeCommande

    ''' <summary>
    ''' Récupération de la liste des commandes
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getListeCommandes() As Boolean
    End Function
    Private Function sauver() As Boolean
        Dim bReturn As Boolean
        Try

            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Function suivant() As Boolean
        Dim bReturn As Boolean
        Try

            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    Private Function elementCourantSansModif() As Boolean
        Return True
    End Function
    Private Sub rechercheClient()
    End Sub 'rechercheClient

#End Region
#Region "Gestion des Evenements"
    Private Sub frmGestionSCMD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub






#End Region

    Protected Overrides Sub AddHandlerValidated(ByVal ocol As System.Windows.Forms.Control.ControlCollection)
        'Dans cette fenêtre seul le bouton Génerer déclenche L'evenement Updated
        'AddHandler cbAppliquer.Click, AddressOf ControlUpdated
        'AddHandler cbGenerer.Click, AddressOf ControlUpdated
    End Sub

    'Private Sub cbSelectionnerTout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSelectionnerTout.Click
    '    Call selectionnerTout()
    'End Sub


    Private Sub rbFTP_CheckedChanged(sender As Object, e As EventArgs) Handles rbFTP.CheckedChanged
        pnlFichierATraiter.Visible = rbManuel.Checked
    End Sub

    Private Sub rbManuel_CheckedChanged(sender As Object, e As EventArgs) Handles rbManuel.CheckedChanged
        pnlFichierATraiter.Visible = rbManuel.Checked
    End Sub

    Private Sub btnTraiter_Click(sender As Object, e As EventArgs) Handles btnTraiter.Click
        Me.Cursor = Cursors.WaitCursor
        btnTraiter.Enabled = False
        BackgroundWorker1.RunWorkerAsync(New UserState())
    End Sub

    Private Function VerificationInfosLivraison(bgw As BackgroundWorker, pFTP As Boolean, pAffichageOK As Boolean, Optional pFileName As String = "") As Boolean
        Dim olst As New List(Of MsgLivraison)
        '    Récupération du fichier à traiter
        If pFTP Then

            '        Récupération du fichier FTP
            Dim strSRV As String = Param.getConstante("CST_FTPEDI_SRV")
            Dim strUser As String = Param.getConstante("CST_FTPEDI_USER")
            Dim strpwd As String = Param.getConstante("CST_FTPEDI_PWD")
            Dim strPort As String = Param.getConstante("CST_FTPEDI_PORT")
            Dim strRepDistant As String = Param.getConstante("CST_FTPEDI_REP")
            Dim strRepLocal As String = Param.getConstante("CST_FTPEDI_REPLOCAL")
            Dim nFiles As Integer


            If Not System.IO.Directory.Exists(strRepLocal) Then
                System.IO.Directory.CreateDirectory(strRepLocal)
            End If
            bgw.ReportProgress(10, "Récupération des fichiers sur le serveur")
            '            DisplayStatus("Récupération des fichiers sur le serveur")
            nFiles = mvtEDI.getFilesCount(strSRV, strPort, strUser, strpwd, strRepDistant, strRepLocal)
            mvtEDI.getFilesFromFTP(strSRV, strPort, strUser, strpwd, strRepDistant, strRepLocal)


            Dim tabFiles As String() = System.IO.Directory.GetFiles(strRepLocal, "*.csv")

            ' Parcours des fichiers .CSV
            Dim nCount As Integer = 0
            For Each strFile As String In tabFiles
                nCount = nCount + 1
                bgw.ReportProgress((nCount / tabFiles.Length) * 90, "Traitement des fichiers")
                Dim olst1 As List(Of MsgLivraison)
                olst1 = mvtEDI.VerificationCommandes(strFile)
                olst.AddRange(olst1)
            Next
        Else
            Dim strFile As String
            strFile = pFileName
            If System.IO.File.Exists(strFile) Then
                bgw.ReportProgress(50, "Traitement du fichier")
                Dim nLines As Integer = 0
                nLines = nLines + System.IO.File.ReadAllLines(strFile).Length

                olst = mvtEDI.VerificationCommandes(strFile)
            End If
        End If

        'oLst la liste des messages issus de la vérification

        bgw.ReportProgress(100, olst)

    End Function


    Private Sub btnQuitter_Click(sender As Object, e As EventArgs) Handles btnQuitter.Click
        Me.Close()
    End Sub

    Private Sub ckAfficherCmdOK_CheckedChanged(sender As Object, e As EventArgs) Handles ckAfficherCmdOK.CheckedChanged
        If ckAfficherCmdOK.Checked Then
            m_bsrcMsgLivraison.RemoveFilter()
        Else
            m_bsrcMsgLivraison.Filter = "Resultat<>0"
        End If
    End Sub

    Private Sub btnParcourir_Click(sender As Object, e As EventArgs) Handles btnParcourir.Click
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            tbFilePath.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Dim strMsg As String

        If e.ProgressPercentage = 100 Then
            Dim olst As New List(Of MsgLivraison)
            olst = TryCast(e.UserState, List(Of MsgLivraison))
            If olst IsNot Nothing Then
                For Each oMsg As MsgLivraison In olst
                    m_bsrcMsgLivraison.Add(oMsg)
                Next

            End If
            lblMessage.Text = "Traietment terminé"
            ProgressBar1.Value = ProgressBar1.Maximum

        Else
            strMsg = e.UserState.ToString()
            lblMessage.Text = strMsg
            ProgressBar1.Value = e.ProgressPercentage

        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        btnTraiter.Enabled = True
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        btnTraiter.Enabled = False
        '        Me.Cursor = Cursors.WaitCursor
        VerificationInfosLivraison(Me.BackgroundWorker1, rbFTP.Checked, ckAfficherCmdOK.Checked, tbFilePath.Text)
    End Sub
End Class
