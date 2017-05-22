Imports vini_DB
Public Class frmImportInternet
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
    Friend WithEvents pbProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents laTraitement As System.Windows.Forms.Label
    Friend WithEvents cbImport As System.Windows.Forms.Button
    Friend WithEvents ckFTP As System.Windows.Forms.CheckBox
    Friend WithEvents tbNbreLignesTraitees As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbNbLignesATraiter As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lbErreurs As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbImport = New System.Windows.Forms.Button
        Me.pbProgressBar = New System.Windows.Forms.ProgressBar
        Me.laTraitement = New System.Windows.Forms.Label
        Me.ckFTP = New System.Windows.Forms.CheckBox
        Me.lbErreurs = New System.Windows.Forms.ListBox
        Me.tbNbreLignesTraitees = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbNbLignesATraiter = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(568, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Cette fenêtre permet de récupérer les rapprochements de factures des fournisseurs" & _
            " établis par le site internet"
        '
        'cbImport
        '
        Me.cbImport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbImport.Location = New System.Drawing.Point(8, 96)
        Me.cbImport.Name = "cbImport"
        Me.cbImport.Size = New System.Drawing.Size(576, 32)
        Me.cbImport.TabIndex = 2
        Me.cbImport.Text = "&Importer"
        '
        'pbProgressBar
        '
        Me.pbProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbProgressBar.Location = New System.Drawing.Point(8, 193)
        Me.pbProgressBar.Name = "pbProgressBar"
        Me.pbProgressBar.Size = New System.Drawing.Size(576, 24)
        Me.pbProgressBar.TabIndex = 2
        '
        'laTraitement
        '
        Me.laTraitement.Location = New System.Drawing.Point(8, 136)
        Me.laTraitement.Name = "laTraitement"
        Me.laTraitement.Size = New System.Drawing.Size(576, 16)
        Me.laTraitement.TabIndex = 3
        '
        'ckFTP
        '
        Me.ckFTP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ckFTP.Checked = CType(configurationAppSettings.GetValue("ckFTP.Checked", GetType(Boolean)), Boolean)
        Me.ckFTP.Location = New System.Drawing.Point(530, -3)
        Me.ckFTP.Name = "ckFTP"
        Me.ckFTP.Size = New System.Drawing.Size(54, 24)
        Me.ckFTP.TabIndex = 4
        Me.ckFTP.Text = "FTP"
        '
        'lbErreurs
        '
        Me.lbErreurs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbErreurs.Location = New System.Drawing.Point(8, 223)
        Me.lbErreurs.Name = "lbErreurs"
        Me.lbErreurs.Size = New System.Drawing.Size(576, 251)
        Me.lbErreurs.TabIndex = 8
        '
        'tbNbreLignesTraitees
        '
        Me.tbNbreLignesTraitees.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbNbreLignesTraitees.Enabled = False
        Me.tbNbreLignesTraitees.Location = New System.Drawing.Point(526, 167)
        Me.tbNbreLignesTraitees.Name = "tbNbreLignesTraitees"
        Me.tbNbreLignesTraitees.Size = New System.Drawing.Size(58, 20)
        Me.tbNbreLignesTraitees.TabIndex = 18
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(401, 170)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(118, 13)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Nbre de lignes traitées :"
        '
        'tbNbLignesATraiter
        '
        Me.tbNbLignesATraiter.Enabled = False
        Me.tbNbLignesATraiter.Location = New System.Drawing.Point(131, 167)
        Me.tbNbLignesATraiter.Name = "tbNbLignesATraiter"
        Me.tbNbLignesATraiter.Size = New System.Drawing.Size(65, 20)
        Me.tbNbLignesATraiter.TabIndex = 16
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 170)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(119, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Nbre de lignes à traiter :"
        '
        'frmImportInternet
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 486)
        Me.Controls.Add(Me.tbNbreLignesTraitees)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbNbLignesATraiter)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lbErreurs)
        Me.Controls.Add(Me.ckFTP)
        Me.Controls.Add(Me.laTraitement)
        Me.Controls.Add(Me.pbProgressBar)
        Me.Controls.Add(Me.cbImport)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmImportInternet"
        Me.Text = "Import Internet"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarSaveEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False

    End Sub


    Private Sub initFenetre()
        Dim dFindeMois As Date
        dFindeMois = Now().AddMonths(1)
        dFindeMois = dFindeMois.AddDays(-1 * (dFindeMois.Day - 1))
        dFindeMois = dFindeMois.AddDays(-1)

        ckFTP.Checked = True
        Me.Text = "Import des sous commandes depuis le site internet"
    End Sub
    Private Function import() As Boolean
        Dim oFTP As clsFTPVinicom = Nothing
        Dim bReturn As Boolean
        Dim nFile As Integer
        Dim strResult As String = ""
        Dim strImportFileName As String

        Dim strFolder As String
        setcursorWait()
        Try
            lbErreurs.Items.Clear()
            'Creation du répertoire temporaire
            strFolder = "ImportInternet" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second
            If My.Computer.FileSystem.DirectoryExists(strFolder) Then
                My.Computer.FileSystem.DeleteDirectory(strFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            My.Computer.FileSystem.CreateDirectory(strFolder)
            'Suppression du fichier toVinicom.csv
            If My.Computer.FileSystem.FileExists("./toVinicom.csv") Then
                My.Computer.FileSystem.DeleteFile("./tovinicom.csv")
            End If


            ' Ce champsest invisible et est lu dans le fichier appConfig
            If ckFTP.Checked Then
                oFTP = New clsFTPVinicom
                'If oFTP.connect() Then
                If True Then
                    If Not oFTP.IsLockTo() Then
                        oFTP.downloadToDir(".")
                    Else
                        lbErreurs.Items.Add("Serveur vérrouillé")
                    End If
                Else
                    lbErreurs.Items.Add("Connexion impossible (" + Param.getConstante("FTP_USERNAME") + " /" + Param.getConstante("FTP_PASSWORD") + ")")
                End If
            End If
            If My.Computer.FileSystem.FileExists("./toVinicom.csv") Then
                'Recopie du fichier d'import
                strImportFileName = strFolder & "/" & IMPORTFTP_FILENAME
                My.Computer.FileSystem.MoveFile("./toVinicom.csv", strImportFileName)
                'Traitement du fichier d'import
                TraitementImportfichier(strImportFileName)
                'Suppressoin du fichier d'import sur le seveur
                If ckFTP.Checked Then
                    If MsgBox("Import terminé, Suppression du fichier d'import", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        oFTP.deleteRemotefile(IMPORTFTP_FILENAME)
                    End If
                End If
            Else
                'pas de fichier à traité
                DisplayStatus("Pas de fichier d'import à traiter")
            End If

        Catch ex As Exception
            DisplayStatus("ERREUR :" & ex.ToString())
            bReturn = False
            FileClose(nFile)
        End Try
        restoreCursor()
        Return bReturn
    End Function
    Private Sub TraitementImportfichier(pstrImportfileName As String)

        Dim bReturn As Boolean
        Dim nFile As Integer
        Dim nLineNumber As Integer
        Dim strResult As String = ""
        Dim tabCSV As String()
        Dim nId As Integer
        Dim oSCMD As SousCommande
        Dim oColSMD As New Collection
        Dim nSousCommandeTraitees As Integer ' Nbre de sousCommandes traitées

        nFile = FreeFile()
        FileOpen(nFile, pstrImportfileName, OpenMode.Input, OpenAccess.Read)
        'Calcul du nombre de lignes à traiter
        nLineNumber = 0
        While Not EOF(nFile)
            nLineNumber = nLineNumber + 1
            LineInput(nFile)
        End While
        FileClose(nFile)
        pbProgressBar.Minimum = 0
        pbProgressBar.Maximum = nLineNumber
        tbNbLignesATraiter.Text = nLineNumber
        nSousCommandeTraitees = 0
        tbNbreLignesTraitees.Text = nSousCommandeTraitees
        nFile = FreeFile()
        FileOpen(nFile, pstrImportfileName, OpenMode.Input, OpenAccess.Read)
        nLineNumber = 0
        While Not EOF(nFile)
            Try
                bReturn = True
                pbProgressBar.Value = nLineNumber
                pbProgressBar.Refresh()
                strResult = LineInput(nFile)
                tabCSV = strResult.Split(";")
                nId = tabCSV(IMPORT_IDSCMD)
                oSCMD = SousCommande.createandload(nId)
                If (Not oSCMD.bNew) Then
                    oSCMD.refFactFournisseur = tabCSV(IMPORT_REFFACTFOURN)
                    oSCMD.dateFactFournisseur = Mid(tabCSV(IMPORT_DATEFACTFOURN), 1, 2) & "/" & Mid(tabCSV(IMPORT_DATEFACTFOURN), 3, 2) & "/" & Mid(tabCSV(IMPORT_DATEFACTFOURN), 5, 4)
                    oSCMD.totalHTFacture = tabCSV(IMPORT_TOTALHTFACTURE)
                    oSCMD.totalTTCFacture = tabCSV(IMPORT_TOTALTTCFACTURE)
                    'importation du taux de commission
                    'oSCMD.tauxCommission = tabCSV(IMPORT_TAUXCOMMISSION)
                    'Calcul du montant de la commission ave le montant facturé
                    oSCMD.calcCommisionstandard(CalculCommScmd.CALCUL_COMMISSION_HT_FACTURE)
                    oSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
                    'Controle de l'état
                    If (oSCMD.etat.codeEtat <> vncEnums.vncEtatCommande.vncSCMDRapprocheeInt) Then
                        DisplayStatus(oSCMD.code + "Erreur en changement d'état" + oSCMD.etat.codeEtat.ToString())
                        bReturn = False
                    End If
                    If bReturn Then
                        bReturn = oSCMD.Save()
                        If (bReturn) Then
                            Log(oSCMD.code + "Etat" + oSCMD.etat.codeEtat.ToString() + "(" + oSCMD.oFournisseur.rs + ") =" + oSCMD.refFactFournisseur + "," + oSCMD.dateFactFournisseur.ToString("d") + ":" + oSCMD.totalHT.ToString("c") + "->" + oSCMD.totalHTFacture.ToString("c"))
                            nSousCommandeTraitees = nSousCommandeTraitees + 1
                            'Ajout dans la collection pour traitement ultérieur
                            oColSMD.Add(oSCMD)
                            tbNbreLignesTraitees.Text = nSousCommandeTraitees
                        Else
                            DisplayStatus("Erreur en Sauvegarde de sous commande " + oSCMD.getErreur())
                            bReturn = False
                        End If
                    End If

                Else
                    DisplayStatus(strResult + "Sous commande inconnue")
                End If
            Catch Ex As Exception
                DisplayStatus(strResult + Ex.Message)
            End Try
            nLineNumber = nLineNumber + 1
            Me.Refresh()
        End While
        FileClose(nFile)
        DisplayStatus("Nbre d'éléments traités :" & nSousCommandeTraitees)
        bReturn = True

    End Sub
    Private Sub cdimport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbImport.Click
        import()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub frmImportInternet_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub
    Private Shadows Sub DisplayStatus(ByVal strMessage As String)

        lbErreurs.Items.Add(Now().ToShortTimeString() + " " + strMessage)
        Log(strMessage)
    End Sub
    Private Sub Log(ByVal strMessage As String)
        System.IO.File.AppendAllText("./vini_internet.trace", Now() + " " + strMessage + vbCrLf)
        Trace.WriteLine(Now() + " " + strMessage)
    End Sub

End Class
