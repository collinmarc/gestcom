Option Strict Off
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Imports System.IO
Public Class frmExportFacture
    Inherits FrmVinicom

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
    Friend WithEvents DsVinicom As vini_DB.dsVinicom
    Friend WithEvents CONSTANTESBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CONSTANTESTableAdapter As vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter
    Friend WithEvents cbParcourir As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbPath As System.Windows.Forms.TextBox
    Friend WithEvents dtFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbExporter As System.Windows.Forms.Button
    Friend WithEvents lbMessages As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtdeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents rbFactCommission As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactTransport As System.Windows.Forms.RadioButton
    Friend WithEvents rbFactColisage As System.Windows.Forms.RadioButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbCodeFacture As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents rbFactHobivin As System.Windows.Forms.RadioButton
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.CONSTANTESBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DsVinicom = New vini_DB.dsVinicom()
        Me.CONSTANTESTableAdapter = New vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.cbParcourir = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbPath = New System.Windows.Forms.TextBox()
        Me.lbMessages = New System.Windows.Forms.ListBox()
        Me.cbExporter = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtFin = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtdeb = New System.Windows.Forms.DateTimePicker()
        Me.rbFactCommission = New System.Windows.Forms.RadioButton()
        Me.rbFactTransport = New System.Windows.Forms.RadioButton()
        Me.rbFactColisage = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbCodeFacture = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.rbFactHobivin = New System.Windows.Forms.RadioButton()
        CType(Me.CONSTANTESBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DsVinicom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CONSTANTESBindingSource
        '
        Me.CONSTANTESBindingSource.DataMember = "CONSTANTES"
        Me.CONSTANTESBindingSource.DataSource = Me.DsVinicom
        '
        'DsVinicom
        '
        Me.DsVinicom.DataSetName = "dsVinicom"
        Me.DsVinicom.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CONSTANTESTableAdapter
        '
        Me.CONSTANTESTableAdapter.ClearBeforeFill = True
        '
        'cbParcourir
        '
        Me.cbParcourir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbParcourir.Location = New System.Drawing.Point(867, 127)
        Me.cbParcourir.Name = "cbParcourir"
        Me.cbParcourir.Size = New System.Drawing.Size(121, 22)
        Me.cbParcourir.TabIndex = 7
        Me.cbParcourir.Text = "Parcourir"
        Me.cbParcourir.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 130)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Répertoire :"
        '
        'tbPath
        '
        Me.tbPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPath.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.CONSTANTESBindingSource, "CST_EXPORT_COMPTA_PATH", True))
        Me.tbPath.Location = New System.Drawing.Point(93, 127)
        Me.tbPath.Name = "tbPath"
        Me.tbPath.Size = New System.Drawing.Size(768, 20)
        Me.tbPath.TabIndex = 6
        '
        'lbMessages
        '
        Me.lbMessages.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbMessages.FormattingEnabled = True
        Me.lbMessages.Location = New System.Drawing.Point(7, 184)
        Me.lbMessages.Name = "lbMessages"
        Me.lbMessages.Size = New System.Drawing.Size(981, 485)
        Me.lbMessages.TabIndex = 9
        '
        'cbExporter
        '
        Me.cbExporter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbExporter.Location = New System.Drawing.Point(868, 155)
        Me.cbExporter.Name = "cbExporter"
        Me.cbExporter.Size = New System.Drawing.Size(120, 23)
        Me.cbExporter.TabIndex = 8
        Me.cbExporter.Text = "Exporter"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(4, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 20)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Date de fin :"
        '
        'dtFin
        '
        Me.dtFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFin.Location = New System.Drawing.Point(93, 31)
        Me.dtFin.Name = "dtFin"
        Me.dtFin.Size = New System.Drawing.Size(122, 20)
        Me.dtFin.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(4, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Date de début :"
        '
        'dtdeb
        '
        Me.dtdeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdeb.Location = New System.Drawing.Point(93, 5)
        Me.dtdeb.Name = "dtdeb"
        Me.dtdeb.Size = New System.Drawing.Size(122, 20)
        Me.dtdeb.TabIndex = 0
        '
        'rbFactCommission
        '
        Me.rbFactCommission.AutoSize = True
        Me.rbFactCommission.Checked = True
        Me.rbFactCommission.Location = New System.Drawing.Point(263, 5)
        Me.rbFactCommission.Name = "rbFactCommission"
        Me.rbFactCommission.Size = New System.Drawing.Size(133, 17)
        Me.rbFactCommission.TabIndex = 3
        Me.rbFactCommission.TabStop = True
        Me.rbFactCommission.Text = "Facture de commission"
        Me.rbFactCommission.UseVisualStyleBackColor = True
        '
        'rbFactTransport
        '
        Me.rbFactTransport.AutoSize = True
        Me.rbFactTransport.Location = New System.Drawing.Point(263, 28)
        Me.rbFactTransport.Name = "rbFactTransport"
        Me.rbFactTransport.Size = New System.Drawing.Size(120, 17)
        Me.rbFactTransport.TabIndex = 4
        Me.rbFactTransport.Text = "Facture de transport"
        Me.rbFactTransport.UseVisualStyleBackColor = True
        '
        'rbFactColisage
        '
        Me.rbFactColisage.AutoSize = True
        Me.rbFactColisage.Location = New System.Drawing.Point(263, 51)
        Me.rbFactColisage.Name = "rbFactColisage"
        Me.rbFactColisage.Size = New System.Drawing.Size(118, 17)
        Me.rbFactColisage.TabIndex = 5
        Me.rbFactColisage.Text = "Facture de colisage"
        Me.rbFactColisage.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 76)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(71, 13)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Code Facture"
        '
        'tbCodeFacture
        '
        Me.tbCodeFacture.Location = New System.Drawing.Point(93, 73)
        Me.tbCodeFacture.Name = "tbCodeFacture"
        Me.tbCodeFacture.Size = New System.Drawing.Size(122, 20)
        Me.tbCodeFacture.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(68, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(19, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "ou"
        '
        'rbFactHobivin
        '
        Me.rbFactHobivin.AutoSize = True
        Me.rbFactHobivin.Location = New System.Drawing.Point(263, 71)
        Me.rbFactHobivin.Name = "rbFactHobivin"
        Me.rbFactHobivin.Size = New System.Drawing.Size(100, 17)
        Me.rbFactHobivin.TabIndex = 20
        Me.rbFactHobivin.TabStop = True
        Me.rbFactHobivin.Text = "Facture Hobivin"
        Me.rbFactHobivin.UseVisualStyleBackColor = True
        '
        'frmExportFacture
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1000, 678)
        Me.Controls.Add(Me.rbFactHobivin)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.tbCodeFacture)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.rbFactColisage)
        Me.Controls.Add(Me.rbFactTransport)
        Me.Controls.Add(Me.rbFactCommission)
        Me.Controls.Add(Me.lbMessages)
        Me.Controls.Add(Me.cbExporter)
        Me.Controls.Add(Me.cbParcourir)
        Me.Controls.Add(Me.tbPath)
        Me.Controls.Add(Me.dtFin)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtdeb)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Name = "frmExportFacture"
        Me.Text = "Exportation des factures vers Quadra"
        CType(Me.CONSTANTESBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DsVinicom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Overrides Function getResume() As String
        Return "Export Facture vers Comptabilité"
    End Function

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub

    Private Sub frmExportFacture_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO : cette ligne de code charge les données dans la table 'DsVinicom.CONSTANTES'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
        Me.CONSTANTESTableAdapter.Connection = Persist.oleDBConnection
        Me.CONSTANTESTableAdapter.Fill(Me.DsVinicom.CONSTANTES)
        rbFactCommission.Checked = True

    End Sub

    Private Sub cbParcourir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbParcourir.Click
        Me.FolderBrowserDialog1.Description = My.Resources.STR_EXPORTCOMPTA_PATH
        Me.FolderBrowserDialog1.SelectedPath = Me.tbPath.Text
        If (Me.FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Me.tbPath.Text = Me.FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub cbExporter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbExporter.Click
        exporter()
    End Sub

    ''' <summary>
    ''' Exporte les factures vers le répertoire indiqué
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub exporter()
        Dim strPath As String
        Dim strFile As String
        Dim dateDeb As Date
        Dim dateFin As Date
        Dim bExportFactureCommision As Boolean
        Dim bExportFactureColisage As Boolean
        Dim bExportFactureTransport As Boolean
        Dim bExportFactureHobivin As Boolean
        Dim bReturn As Boolean
        Dim colFact As Collection
        Dim bCodeFact As Boolean
        Dim strCode As String

        Try


            'Lecture des valeurs

            Me.setcursorWait()
            strPath = Trim(tbPath.Text)
            dateDeb = dtdeb.Value.ToShortDateString()
            dateFin = dtFin.Value.ToShortDateString()
            bExportFactureCommision = rbFactCommission.Checked
            bExportFactureColisage = rbFactColisage.Checked
            bExportFactureTransport = rbFactTransport.Checked
            bExportFactureHobivin = rbFactHobivin.Checked
            strCode = tbCodeFacture.Text
            bCodeFact = Not String.IsNullOrEmpty(strCode)

            lbMessages.Items.Clear()
            bReturn = True
            If bExportFactureCommision Then
                DisplayMessage("== EXPORTATION DES FACTURES DE COMMISSIONS ==")
                strFile = strPath + "/FACTCOM" + Format(Now(), "yyyyMMdd")
                If bCodeFact Then
                    colFact = FactCom.getListe(strCode)
                Else
                    colFact = FactCom.getListe(dateDeb, dateFin, , vncEtatCommande.vncFactComGeneree)
                End If
                bReturn = bReturn And exportFacture(strFile, colFact)
            End If
            If bExportFactureTransport Then
                DisplayMessage("== EXPORTATION DES FACTURES DE TRANSPORT ==")
                strFile = strPath + "/FACTTRP" + Format(Now(), "yyyyMMdd")
                If bCodeFact Then
                    colFact = FactTRP.getListe(strCode)
                Else
                    colFact = FactTRP.getListe(dateDeb, dateFin, , vncEtatCommande.vncFactTRPGeneree)
                End If
                bReturn = bReturn And exportFacture(strFile, colFact)
            End If
                If bExportFactureColisage Then
                    DisplayMessage("== EXPORTATION DES FACTURES DE COLISAGE ==")
                strFile = strPath + "/FACTCOL" + Format(Now(), "yyyyMMdd")
                If bCodeFact Then
                    colFact = FactColisageJ.getListe(strCode)
                Else
                    colFact = FactColisageJ.getListe(dateDeb, dateFin, , vncEtatCommande.vncFactCOLGeneree)
                End If
                bReturn = bReturn And exportFacture(strFile, colFact)
            End If
            If bExportFactureHobivin Then
                DisplayMessage("== EXPORTATION DES FACTURES HOBIVIN ==")
                strFile = strPath + "/FACTHBV" + Format(Now(), "yyyyMMdd")
                If bCodeFact Then
                    colFact = FactHBV.getListe(strCode)
                Else
                    colFact = FactHBV.getListe(dateDeb, dateFin, , vncEtatCommande.vncFactHBVGeneree)
                End If
                bReturn = bReturn And exportFacture(strFile, colFact)
            End If


            If bReturn Then
                MessageBox.Show(My.Resources.STR_EXPORTCOMPTA_OK)
                If (MessageBox.Show(My.Resources.STR_VALIDATION_EXPORT_FACTURE, "Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes) Then
                    ValiderExport()
                End If
            Else
                MessageBox.Show(My.Resources.STR_EXPORTCOMPTA_NOK, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            restoreCursor()
        Catch ex As Exception
            DisplayMessage(ex.Message)
        End Try


    End Sub

    ''' <summary>
    ''' Change l'état des factures à facturée
    ''' </summary>
    ''' <remarks></remarks>
    Private Function ValiderExport() As Boolean
        Dim strPath As String
        Dim dateDeb As Date
        Dim dateFin As Date
        Dim bExportFactureCommision As Boolean
        Dim bExportFactureColisage As Boolean
        Dim bExportFactureTransport As Boolean
        Dim bExportFactureHobivin As Boolean
        Dim bReturn As Boolean
        Dim colFact As Collection

        Try


            'Lecture des valeurs

            strPath = tbPath.Text
            dateDeb = dtdeb.Value.ToShortDateString()
            dateFin = dtFin.Value.ToShortDateString()
            bExportFactureCommision = rbFactCommission.Checked
            bExportFactureColisage = rbFactColisage.Checked
            bExportFactureTransport = rbFactTransport.Checked
            bExportFactureHobivin = rbFactHobivin.Checked

            bReturn = True
            If bExportFactureCommision Then
                DisplayMessage("== Validation Export Facture de commisions ==")
                colFact = FactCom.getListe(dateDeb, dateFin, , vncEtatCommande.vncFactComGeneree)
                For Each objFact As Facture In colFact
                    objFact.changeEtat(vncActionEtatCommande.vncActionFactComExporter)
                    objFact.Save()
                Next
            End If
            If bExportFactureColisage Then
                DisplayMessage("== Validation Export Facture de colisage ==")
                colFact = FactColisageJ.getListe(dateDeb, dateFin, , vncEtatCommande.vncFactCOLGeneree)
                For Each objFact As Facture In colFact
                    objFact.changeEtat(vncActionEtatCommande.vncActionFactCOLExporter)
                    objFact.Save()
                Next
            End If
            If bExportFactureTransport Then
                DisplayMessage("== Validation Export Facture de transport ==")
                colFact = FactTRP.getListe(dateDeb, dateFin, , vncEtatCommande.vncFactTRPGeneree)
                For Each objFact As Facture In colFact
                    objFact.changeEtat(vncActionEtatCommande.vncActionFactTRPExporter)
                    objFact.Save()
                Next
            End If
            If bExportFactureHobivin Then
                DisplayMessage("== Validation Export Facture de Hobivin ==")
                colFact = FactHBV.getListe(dateDeb, dateFin, , vncEtatCommande.vncFactHBVGeneree)
                For Each objFact As Facture In colFact
                    objFact.changeEtat(vncActionEtatCommande.vncActionFactHBVExporter)
                    objFact.Save()
                Next
            End If
            DisplayMessage("Validation terminée")
            bReturn = True

        Catch ex As Exception
            DisplayMessage(ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function
    ''' <summary>
    ''' Export les factures vers le répertoire indiqué
    ''' </summary>
    ''' <param name="pstrFile">répertoire d'exportation</param>
    ''' <param name="colFact">Collection de factures à exporter</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function exportFacture(ByVal pstrFile As String, ByVal colFact As Collection) As Boolean
        Dim bReturn As Boolean
        Dim bFileOk As Boolean
        Dim nTotal As Decimal
        Try
            If File.Exists(pstrFile) Then
                bFileOk = False
                If MsgBox("Le Fichier " & pstrFile & " existe voulez-vous l'écraser?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    File.Delete(pstrFile)
                    bFileOk = True
                End If
            Else
                bFileOk = True
            End If
            If bFileOk Then
                For Each objFact As Facture In colFact
                    DisplayMessage("Export Facture n°" + objFact.code + " " + objFact.oTiers.nom + " " + objFact.totalTTC.ToString("c"))
                    nTotal = nTotal + objFact.totalTTC
                    objFact.Exporter(pstrFile)
                Next
            End If
            DisplayMessage("==========================================")
            DisplayMessage("TOTAL exporté :" + nTotal.ToString("c"))
            bReturn = True
        Catch ex As Exception
            DisplayMessage(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'exportFacture
    Private Sub DisplayMessage(ByVal strMessage As String)
        lbMessages.Items.Add(DateAndTime.Now.ToLongTimeString() + ":" + strMessage)
        lbMessages.SetSelected(lbMessages.Items.Count - 1, True)
        lbMessages.Refresh()
    End Sub

End Class
