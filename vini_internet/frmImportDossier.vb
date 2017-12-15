Imports vini_DB
Public Class frmImportDossier
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
    Friend WithEvents cbImport As System.Windows.Forms.Button
    Friend WithEvents tbPath As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dlg_openFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cbParcourir As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbNbLignesATraiter As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbNbreLignesTraitees As System.Windows.Forms.TextBox
    Friend WithEvents lbErreurs As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbImport = New System.Windows.Forms.Button
        Me.pbProgressBar = New System.Windows.Forms.ProgressBar
        Me.lbErreurs = New System.Windows.Forms.ListBox
        Me.tbPath = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.dlg_openFile = New System.Windows.Forms.OpenFileDialog
        Me.cbParcourir = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbNbLignesATraiter = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbNbreLignesTraitees = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(568, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Cette fenêtre permet de refaire le rapprochement après un echec lors de l'import " & _
            "internet"
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
        Me.pbProgressBar.Location = New System.Drawing.Point(8, 156)
        Me.pbProgressBar.Name = "pbProgressBar"
        Me.pbProgressBar.Size = New System.Drawing.Size(576, 24)
        Me.pbProgressBar.TabIndex = 3
        '
        'lbErreurs
        '
        Me.lbErreurs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbErreurs.Location = New System.Drawing.Point(8, 184)
        Me.lbErreurs.Name = "lbErreurs"
        Me.lbErreurs.Size = New System.Drawing.Size(576, 290)
        Me.lbErreurs.TabIndex = 4
        '
        'tbPath
        '
        Me.tbPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPath.Location = New System.Drawing.Point(12, 70)
        Me.tbPath.Name = "tbPath"
        Me.tbPath.Size = New System.Drawing.Size(482, 20)
        Me.tbPath.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Fichier à importer :"
        '
        'dlg_openFile
        '
        Me.dlg_openFile.AddExtension = False
        Me.dlg_openFile.DefaultExt = "csv"
        Me.dlg_openFile.FileName = "toVinicom.csv"
        Me.dlg_openFile.InitialDirectory = "./"
        Me.dlg_openFile.Title = "Fichier à importer"
        '
        'cbParcourir
        '
        Me.cbParcourir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbParcourir.Location = New System.Drawing.Point(501, 70)
        Me.cbParcourir.Name = "cbParcourir"
        Me.cbParcourir.Size = New System.Drawing.Size(75, 23)
        Me.cbParcourir.TabIndex = 1
        Me.cbParcourir.Text = "Parcourir"
        Me.cbParcourir.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 133)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(119, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Nbre de lignes à traiter :"
        '
        'tbNbLignesATraiter
        '
        Me.tbNbLignesATraiter.Enabled = False
        Me.tbNbLignesATraiter.Location = New System.Drawing.Point(129, 130)
        Me.tbNbLignesATraiter.Name = "tbNbLignesATraiter"
        Me.tbNbLignesATraiter.Size = New System.Drawing.Size(65, 20)
        Me.tbNbLignesATraiter.TabIndex = 12
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(399, 133)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(118, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Nbre de lignes traitées :"
        '
        'tbNbreLignesTraitees
        '
        Me.tbNbreLignesTraitees.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbNbreLignesTraitees.Enabled = False
        Me.tbNbreLignesTraitees.Location = New System.Drawing.Point(524, 130)
        Me.tbNbreLignesTraitees.Name = "tbNbreLignesTraitees"
        Me.tbNbreLignesTraitees.Size = New System.Drawing.Size(60, 20)
        Me.tbNbreLignesTraitees.TabIndex = 14
        '
        'frmImportDossier
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 486)
        Me.Controls.Add(Me.tbNbreLignesTraitees)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbNbLignesATraiter)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbParcourir)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbPath)
        Me.Controls.Add(Me.lbErreurs)
        Me.Controls.Add(Me.pbProgressBar)
        Me.Controls.Add(Me.cbImport)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmImportDossier"
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

    End Sub
    Private Function import() As Boolean
        Dim bReturn As Boolean
        Dim nFile As Integer
        Dim nLineNumber As Integer
        Dim strResult As String = ""
        Dim tabCSV As String()
        Dim nId As Integer
        Dim oSCMD As SousCommande
        Dim strImportFileName As String

        Dim nSousCommandes As Integer
        setcursorWait()
        Try
            nSousCommandes = 0
            lbErreurs.Items.Clear()
            strImportFileName = tbPath.Text
            If My.Computer.FileSystem.FileExists(strImportFileName) Then

                nFile = FreeFile()
                FileOpen(nFile, strImportFileName, OpenMode.Input, OpenAccess.Read)
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
                tbNbreLignesTraitees.Text = "0"

                nFile = FreeFile()
                FileOpen(nFile, strImportFileName, OpenMode.Input, OpenAccess.Read)
                nLineNumber = 0
                While Not EOF(nFile)
                    Try
                        bReturn = True
                        nLineNumber = nLineNumber + 1
                        pbProgressBar.Value = nLineNumber
                        pbProgressBar.Refresh()
                        strResult = LineInput(nFile)
                        tabCSV = strResult.Split(";")
                        nId = tabCSV(IMPORT_IDSCMD)
                        oSCMD = SousCommande.createandload(nId)
                        If (Not oSCMD.bNew) Then
                            Dim bSCMDATraiter As Boolean = True
                            'Controle des sousCommandes déjà facturées (import déjà effectué mais non validé)
                            If oSCMD.etat.codeEtat = vncEtatCommande.vncSCMDFacturee Or oSCMD.etat.codeEtat = vncEtatCommande.vncSCMDRapprocheeInt Then
                                If oSCMD.refFactFournisseur = tabCSV(IMPORT_REFFACTFOURN) Then
                                    bSCMDATraiter = False
                                End If
                            End If
                            If bSCMDATraiter Then
                                oSCMD.refFactFournisseur = tabCSV(IMPORT_REFFACTFOURN)
                                oSCMD.dateFactFournisseur = Mid(tabCSV(IMPORT_DATEFACTFOURN), 1, 2) & "/" & Mid(tabCSV(IMPORT_DATEFACTFOURN), 3, 2) & "/" & Mid(tabCSV(IMPORT_DATEFACTFOURN), 5, 4)
                                tabCSV(IMPORT_TOTALHTFACTURE) = tabCSV(IMPORT_TOTALHTFACTURE).Replace(".", System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator)
                                tabCSV(IMPORT_TOTALHTFACTURE) = tabCSV(IMPORT_TOTALHTFACTURE).Replace(",", System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator)
                                tabCSV(IMPORT_TOTALTTCFACTURE) = tabCSV(IMPORT_TOTALTTCFACTURE).Replace(".", System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator)
                                tabCSV(IMPORT_TOTALTTCFACTURE) = tabCSV(IMPORT_TOTALTTCFACTURE).Replace(",", System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator)
                                oSCMD.totalHTFacture = tabCSV(IMPORT_TOTALHTFACTURE)
                                oSCMD.totalTTCFacture = Convert.ToDecimal(tabCSV(IMPORT_TOTALTTCFACTURE))
                                'importation du taux de commission
                                'oSCMD.tauxCommission = tabCSV(IMPORT_TAUXCOMMISSION)
                                'Calcul du montant de la commission ave le montant facturé
                                oSCMD.calcCommisionstandard(CalculCommScmd.CALCUL_COMMISSION_HT_FACTURE)
                                oSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDImportInternet)
                                bReturn = oSCMD.Save()
                                If (bReturn) Then
                                    Log(oSCMD.code + "Etat" + oSCMD.etat.codeEtat.ToString() + "(" + oSCMD.oFournisseur.rs + ") =" + oSCMD.refFactFournisseur + "," + oSCMD.dateFactFournisseur.ToString("d") + ":" + oSCMD.totalHT.ToString("c") + "->" + oSCMD.totalHTFacture.ToString("c"))
                                    'Ajout dans la collection pour traitement ultérieur
                                    lbErreurs.Items.Add(oSCMD.code + "(" + oSCMD.oFournisseur.rs + ") =" + oSCMD.refFactFournisseur + "," + oSCMD.dateFactFournisseur.ToString("d") + ":" + oSCMD.totalHT.ToString("c") + "->" + oSCMD.totalHTFacture.ToString("c"))
                                    lbErreurs.Refresh()
                                    nSousCommandes = nSousCommandes + 1
                                    tbNbreLignesTraitees.Text = nLineNumber
                                End If
                            End If
                        Else
                            DisplayStatus(strResult + "Sous commande inconnue")
                        End If
                    Catch Ex As Exception
                        lbErreurs.Items.Add(strResult + Ex.Message)
                        lbErreurs.Refresh()
                    End Try

                End While
                FileClose(nFile)

                lbErreurs.Items.Add("Nbre d'éléments traités :" & nSousCommandes)
                lbErreurs.Refresh()
                bReturn = True
            End If

        Catch ex As Exception
            lbErreurs.Items.Add("ERREUR :" & ex.ToString())
            bReturn = False
            FileClose(nFile)
        End Try
        restoreCursor()
        Return bReturn
    End Function

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

    Private Sub cbParcourir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbParcourir.Click
        dlg_openFile.FileName = "ToVinicom.csv"
        If dlg_openFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            tbPath.Text = dlg_openFile.FileName
        End If
    End Sub
    Private Sub Log(ByVal strMessage As String)
        System.IO.File.AppendAllText("./vini_internet.trace", Now() + " " + strMessage + vbCrLf)
        Trace.WriteLine(Now() + " " + strMessage)
    End Sub
    Private Shadows Sub DisplayStatus(ByVal strMessage As String)

        lbErreurs.Items.Add(Now().ToShortTimeString() + " " + strMessage)
        lbErreurs.Refresh()
        Log(strMessage)
    End Sub
End Class
