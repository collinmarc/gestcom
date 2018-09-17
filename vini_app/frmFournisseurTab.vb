Option Strict Off
Imports vini_DB
Public Class frmFournisseurTab
    Inherits frmTiers



#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        m_TypeDonnees = vncEnums.vncTypeDonnee.FOURNISSEUR
        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        initcboRegion(cboRegion)
        rtbCom3.BackColor = System.Drawing.Color.Yellow
        laModeReglmt.Text = "B.A."
        laModeRglmt1.Text = "Commision"
        LaModeReglmt2.Text = "Colisage"
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
    'Public WithEvents Label4 As System.Windows.Forms.Label
    '   Friend WithEvents listeClient1 As Vinicom.ListeClient
    'Friend WithEvents OdbcDataAdapter1 As System.Data.Odbc.OdbcDataAdapter
    'Friend WithEvents OdbcConnection1 As System.Data.Odbc.OdbcConnection
    'Friend WithEvents DsParamClient1 As Vinicom.dsParamClient
    'Friend WithEvents OdbcSelectCommand1 As System.Data.Odbc.OdbcCommand
    'Friend WithEvents OdbcInsertCommand1 As System.Data.Odbc.OdbcCommand
    'Friend WithEvents OdbcDataAdapter2 As System.Data.Odbc.OdbcDataAdapter
    'Friend WithEvents OdbcSelectCommand2 As System.Data.Odbc.OdbcCommand
    'Friend WithEvents OdbcInsertCommand2 As System.Data.Odbc.OdbcCommand
    'Friend WithEvents DsLstModeReglmt1 As Vinicom.dsLstModeReglmt
    'Friend WithEvents listeClient1 As Vinicom.ListeClient
    'Friend WithEvents flxlstProduits As AxMSFlexGridLib.AxMSFlexGrid
    '    Friend WithEvents Label33 As System.Windows.Forms.Label
    '    Friend WithEvents tbBanque As System.Windows.Forms.TextBox
    '    Public WithEvents tbRib4 As System.Windows.Forms.TextBox
    '    Public WithEvents tbRib3 As System.Windows.Forms.TextBox
    '    Public WithEvents tbRib2 As System.Windows.Forms.TextBox
    '    Public WithEvents tbRib1 As System.Windows.Forms.TextBox
    '    Friend WithEvents RIB As System.Windows.Forms.Label
    '    Friend WithEvents Label25 As System.Windows.Forms.Label
    '    Friend WithEvents tbTVAIntra As System.Windows.Forms.TextBox
    '    Friend WithEvents Label28 As System.Windows.Forms.Label
    '    Friend WithEvents tbSIRET As System.Windows.Forms.TextBox
    '    Friend WithEvents Label31 As System.Windows.Forms.Label
    '    '    Friend WithEvents ckAdrIdentiques As System.Windows.Forms.CheckBox
    '    Friend WithEvents Label3 As System.Windows.Forms.Label
    '    Friend WithEvents Label7 As System.Windows.Forms.Label
    '    Friend WithEvents Label8 As System.Windows.Forms.Label
    '    Friend WithEvents Label9 As System.Windows.Forms.Label
    '    Friend WithEvents Label15 As System.Windows.Forms.Label
    '    Friend WithEvents Label20 As System.Windows.Forms.Label
    '    Friend WithEvents tbAdrLivNom As System.Windows.Forms.TextBox
    '    Friend WithEvents Label21 As System.Windows.Forms.Label
    '    Friend WithEvents Label22 As System.Windows.Forms.Label
    '    Public WithEvents tbAdrLivPortable As System.Windows.Forms.TextBox
    '    Public WithEvents tbAdrLivRue1 As System.Windows.Forms.TextBox
    '    Public WithEvents tbAdrLivRue2 As System.Windows.Forms.TextBox
    '    Public WithEvents tbAdrLivCP As System.Windows.Forms.TextBox
    '    Public WithEvents tbAdrLivVille As System.Windows.Forms.TextBox
    '    Public WithEvents tbAdrLivTel As System.Windows.Forms.TextBox
    '    Public WithEvents tbAdrLivEmail As System.Windows.Forms.TextBox
    '    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    '    Friend WithEvents Label19 As System.Windows.Forms.Label
    '    Friend WithEvents Label18 As System.Windows.Forms.Label
    '    Public WithEvents Label14 As System.Windows.Forms.Label
    '    Public WithEvents Label13 As System.Windows.Forms.Label
    '    Friend WithEvents Label12 As System.Windows.Forms.Label
    '    Friend WithEvents Label11 As System.Windows.Forms.Label
    '    Friend WithEvents tbAdrFactNom As System.Windows.Forms.TextBox
    '    Friend WithEvents Label10 As System.Windows.Forms.Label
    'Friend WithEvents Label5 As System.Windows.Forms.Label
    'Public WithEvents tbAdrFactPortable As System.Windows.Forms.TextBox
    'Public WithEvents tbAdrFactRue1 As System.Windows.Forms.TextBox
    'Public WithEvents tbAdrFactRue2 As System.Windows.Forms.TextBox
    'Public WithEvents tbAdrFactCP As System.Windows.Forms.TextBox
    'Public WithEvents tbAdrFactVille As System.Windows.Forms.TextBox
    'Public WithEvents tbAdrFactTel As System.Windows.Forms.TextBox
    'Public WithEvents tbAdrFactFax As System.Windows.Forms.TextBox
    'Public WithEvents tbAdrFactEmail As System.Windows.Forms.TextBox
    'Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    'Public WithEvents Label17 As System.Windows.Forms.Label
    'Public WithEvents Label16 As System.Windows.Forms.Label
    'Friend WithEvents cboModeRglmt As System.Windows.Forms.ComboBox
    'Friend WithEvents Label23 As System.Windows.Forms.Label
    'Friend WithEvents Label26 As System.Windows.Forms.Label
    'Friend WithEvents Label24 As System.Windows.Forms.Label
    'Friend WithEvents Label32 As System.Windows.Forms.Label
    'Public WithEvents rtbComm1 As System.Windows.Forms.RichTextBox
    'Public WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    'Public WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
    'Public WithEvents RichTextBox3 As System.Windows.Forms.RichTextBox
    'Friend WithEvents flxLstProduits As AxMSFlexGridLib.AxMSFlexGrid
    'Friend WithEvents cbVisualiser As System.Windows.Forms.Button
    'Friend WithEvents cbPCAjouter As System.Windows.Forms.Button
    'Friend WithEvents cbPCModifier As System.Windows.Forms.Button
    'Friend WithEvents cbPCSuppr As System.Windows.Forms.Button
    'Friend WithEvents cbPCAppliquer As System.Windows.Forms.Button
    'Friend WithEvents tbPCQuantité As System.Windows.Forms.TextBox
    'Friend WithEvents Label30 As System.Windows.Forms.Label
    'Friend WithEvents tbPCPrix As System.Windows.Forms.TextBox
    'Friend WithEvents Label29 As System.Windows.Forms.Label
    'Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    'Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    'Friend WithEvents cbRechProduit As System.Windows.Forms.Button
    'Friend WithEvents tbPCCodeProduit As System.Windows.Forms.TextBox
    'Friend WithEvents Label27 As System.Windows.Forms.Label
    'Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents cboRegion As System.Windows.Forms.ComboBox
    Friend WithEvents cbStockFournisseur As System.Windows.Forms.Button
    Friend WithEvents cbCommissions As System.Windows.Forms.Button
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ckIntermediaire As System.Windows.Forms.CheckBox
    Friend WithEvents cbxDossier As System.Windows.Forms.ComboBox
    Friend WithEvents cbxExportBaf As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cboRegion = New System.Windows.Forms.ComboBox()
        Me.cbStockFournisseur = New System.Windows.Forms.Button()
        Me.cbCommissions = New System.Windows.Forms.Button()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.cbxExportBaf = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ckIntermediaire = New System.Windows.Forms.CheckBox()
        Me.cbxDossier = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'cboRegion
        '
        Me.cboRegion.BackColor = System.Drawing.SystemColors.Window
        Me.cboRegion.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboRegion.DisplayMember = "RQ_TypeClient.PAR_ID"
        Me.cboRegion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRegion.Enabled = False
        Me.cboRegion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboRegion.Location = New System.Drawing.Point(784, 7)
        Me.cboRegion.Name = "cboRegion"
        Me.cboRegion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboRegion.Size = New System.Drawing.Size(176, 21)
        Me.cboRegion.TabIndex = 2
        Me.cboRegion.ValueMember = "RQ_TypeClient.PAR_ID"
        '
        'cbStockFournisseur
        '
        Me.cbStockFournisseur.CausesValidation = False
        Me.cbStockFournisseur.Enabled = False
        Me.cbStockFournisseur.Location = New System.Drawing.Point(784, 55)
        Me.cbStockFournisseur.Name = "cbStockFournisseur"
        Me.cbStockFournisseur.Size = New System.Drawing.Size(176, 21)
        Me.cbStockFournisseur.TabIndex = 70
        Me.cbStockFournisseur.Text = "St&Ock"
        '
        'cbCommissions
        '
        Me.cbCommissions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbCommissions.Location = New System.Drawing.Point(784, 31)
        Me.cbCommissions.Name = "cbCommissions"
        Me.cbCommissions.Size = New System.Drawing.Size(176, 21)
        Me.cbCommissions.TabIndex = 73
        Me.cbCommissions.Text = "Com&missions"
        Me.cbCommissions.UseVisualStyleBackColor = True
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(731, 11)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(47, 13)
        Me.Label27.TabIndex = 74
        Me.Label27.Text = "Région :"
        '
        'cbxExportBaf
        '
        Me.cbxExportBaf.FormattingEnabled = True
        Me.cbxExportBaf.Items.AddRange(New Object() {"0-Pas d'export", "1-Export Internet", "2-Export Quadra"})
        Me.cbxExportBaf.Location = New System.Drawing.Point(604, 7)
        Me.cbxExportBaf.Name = "cbxExportBaf"
        Me.cbxExportBaf.Size = New System.Drawing.Size(121, 21)
        Me.cbxExportBaf.TabIndex = 75
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(510, 11)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 76
        Me.Label4.Text = "Export des Baf :"
        '
        'ckIntermediaire
        '
        Me.ckIntermediaire.AutoSize = True
        Me.ckIntermediaire.Checked = True
        Me.ckIntermediaire.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckIntermediaire.Location = New System.Drawing.Point(297, 82)
        Me.ckIntermediaire.Name = "ckIntermediaire"
        Me.ckIntermediaire.Size = New System.Drawing.Size(86, 17)
        Me.ckIntermediaire.TabIndex = 77
        Me.ckIntermediaire.Text = "Intermediaire"
        Me.ckIntermediaire.UseVisualStyleBackColor = True
        '
        'cbxDossier
        '
        Me.cbxDossier.FormattingEnabled = True
        Me.cbxDossier.Location = New System.Drawing.Point(400, 82)
        Me.cbxDossier.Name = "cbxDossier"
        Me.cbxDossier.Size = New System.Drawing.Size(121, 21)
        Me.cbxDossier.TabIndex = 78
        '
        'frmFournisseurTab
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(968, 734)
        Me.Controls.Add(Me.cbxDossier)
        Me.Controls.Add(Me.ckIntermediaire)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbxExportBaf)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.cbCommissions)
        Me.Controls.Add(Me.cbStockFournisseur)
        Me.Controls.Add(Me.cboRegion)
        Me.Name = "frmFournisseurTab"
        Me.Text = "Gestion des Fournisseurs"
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.tbCode, 0)
        Me.Controls.SetChildIndex(Me.tbRaisonSociale, 0)
        Me.Controls.SetChildIndex(Me.cboRegion, 0)
        Me.Controls.SetChildIndex(Me.cbStockFournisseur, 0)
        Me.Controls.SetChildIndex(Me.cbCommissions, 0)
        Me.Controls.SetChildIndex(Me.Label27, 0)
        Me.Controls.SetChildIndex(Me.cbxExportBaf, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.ckIntermediaire, 0)
        Me.Controls.SetChildIndex(Me.cbxDossier, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Overrides Function getResume() As String 'Rend le caption de la fenêtre
        If getElementCourant() Is Nothing Then
            Return "Gestion des Fournisseurs"
        Else
            Return "Fournisseur : " & getElementCourant.shortResume
        End If
    End Function 'getResume

    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated, "La fenetre n'est pas libre")

        setElementCourant2(New Fournisseur("", ""))

        Return True
    End Function
    Public Overrides Function AfficheElement() As Boolean
        Dim objFournisseur As Fournisseur
        Dim objParam As Param
        Dim bReturn As Boolean
        Debug.Assert(getElementCourant.GetType().Name.Equals("Fournisseur"), "Objet de type Fournisseur requis")


        bReturn = MyBase.AfficheElement()
        If bReturn Then
            objFournisseur = CType(getElementCourant(), Fournisseur)

            For Each objParam In cboRegion.Items
                If objParam.id = objFournisseur.idRegion Then
                    cboRegion.SelectedItem = objParam
                    Exit For
                End If
            Next

            cbxExportBaf.SelectedIndex = objFournisseur.bExportInternet
            ckIntermediaire.Checked = objFournisseur.bIntermdiaire
            cbxDossier.Text = objFournisseur.Dossier
        End If

        Return bReturn
    End Function 'AfficheElementCourant

    Public Overrides Function MAJElement() As Boolean
        Dim bReturn As Boolean
        Dim objFournisseur As Fournisseur
        bReturn = MyBase.MAJElement()
        If bReturn Then
            objFournisseur = CType(getElementCourant(), Fournisseur)
            Try
                objFournisseur.idRegion = cboRegion.SelectedItem.id
                objFournisseur.libregion = cboRegion.SelectedItem.Valeur
                objFournisseur.bExportInternet = cbxExportBaf.SelectedIndex
                objFournisseur.bIntermdiaire = ckIntermediaire.Checked
                objFournisseur.Dossier = cbxDossier.Text
                bReturn = True
            Catch ex As Exception
                DisplayError("frmFournisseur.MAJElement", ex.ToString())
                bReturn = False
            End Try

        End If
        Return bReturn
    End Function 'MAJElementCourant


    Private Sub cbStockFournisseur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStockFournisseur.Click
        Debug.Assert(Not getElementCourant() Is Nothing)
        Dim odlg As dlgVisuStockFournisseur
        Dim bOk As Boolean = True

        If isfrmUpdated() Then
            If MsgBox("L'élement courant a été modifié, souhaitez-vous l'enregister?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                If frmSave() Then
                    bOk = True
                Else
                    bOk = False
                End If
            Else
                bOk = True
            End If
        End If

        If bOk Then
            odlg = New dlgVisuStockFournisseur
            odlg.Show()
            odlg.setFournisseur(getElementCourant())
        End If

    End Sub

    Public Overrides Function ControleAvantSauvegarde() As Boolean
        Dim bReturn As Boolean
        bReturn = True
        If getElementCourant().bNew = True Then
            'controle de l'unicité du code
            If Fournisseur.getListe(tbCode.Text).Count > 0 Then
                DisplayError("Controleavantsauvegarde", "Le code Fournisseur doit être unique")
                bReturn = False
            End If
        End If
        bReturn = bReturn And MyBase.ControleAvantSauvegarde()
        Return bReturn

    End Function

    Private Sub frmFournisseurTab_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If currentuser.role = vncEnums.userRole.ADMIN Or currentuser.role = vncEnums.userRole.COMPTABILITE Then
            rtbCom3.Visible = True
            lartbCom3.Visible = True
            cbCommissions.Visible = True
        Else
            rtbCom3.Visible = False
            lartbCom3.Visible = False
            cbCommissions.Visible = False
        End If

        cbxDossier.Items.Clear()
        cbxDossier.Items.Add(Dossier.VINICOM)
        cbxDossier.Items.Add(Dossier.HOBIVIN)
        'Suppression du controle de l'evenement Validated
        '        RemoveHandler cbCommissions.Validated, AddressOf ControlUpdated
        'on peut aussi supprimer la génération de det evenement
        cbCommissions.CausesValidation = False

    End Sub

    Private Sub cbCommissions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCommissions.Click
        'Affiche les Taux de commisions du Fournisseur
        AfficheDlgcomm()
    End Sub

    ''' <summary>
    ''' Affiche la boite de disalogue de gestion des taux de commissions
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AfficheDlgcomm() As Boolean
        Dim bReturn As Boolean

        Try
            Dim odlg As dlgFrnComm
            odlg = New dlgFrnComm
            odlg.setFRNID(getElementCourant.id)
            odlg.ShowDialog(Me)
            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function

    Private Sub ckIntermediaire_CheckedChanged(sender As Object, e As EventArgs) Handles ckIntermediaire.CheckedChanged
        If ckIntermediaire.Checked = False Then
            cbxDossier.Visible = False
        Else
            cbxDossier.Visible = True
        End If
    End Sub
End Class
