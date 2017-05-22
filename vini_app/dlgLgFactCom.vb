Imports vini_DB
Friend Class dlgLgFactCom
    Inherits System.Windows.Forms.Form
    Private m_elementCourant As LgCommande
    Private m_SCMDCourante As SousCommande
    Private m_FournisseurCourant As Fournisseur
    Private m_bModif As Boolean
    Private m_bAffichageEncours As Boolean
    'Modification ou Création de lignes

#Region "Code généré par le Concepteur Windows Form "
    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
        m_bModif = False
    End Sub
    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer
    Public WithEvents Label2 As System.Windows.Forms.Label
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    Public WithEvents cbAnnuler As System.Windows.Forms.Button
    Public WithEvents cbValider As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbCodeCommande As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtDateCommande As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbCodeFactFourn As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtDateFactFourn As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents tbTotalHTFacture As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents tbBaseComm As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents tbMontantComm As System.Windows.Forms.TextBox
    Friend WithEvents liClient As System.Windows.Forms.LinkLabel
    Public WithEvents tbCode As System.Windows.Forms.TextBox
    Public WithEvents cbREcherche As System.Windows.Forms.Button
    Friend WithEvents grpSCMD As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cbREcherche = New System.Windows.Forms.Button
        Me.tbCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cbAnnuler = New System.Windows.Forms.Button
        Me.cbValider = New System.Windows.Forms.Button
        Me.grpSCMD = New System.Windows.Forms.GroupBox
        Me.liClient = New System.Windows.Forms.LinkLabel
        Me.tbMontantComm = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.tbBaseComm = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.tbTotalHTFacture = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.dtDateFactFourn = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.tbCodeFactFourn = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtDateCommande = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbCodeCommande = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.grpSCMD.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbREcherche
        '
        Me.cbREcherche.BackColor = System.Drawing.SystemColors.Control
        Me.cbREcherche.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbREcherche.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbREcherche.Location = New System.Drawing.Point(248, 32)
        Me.cbREcherche.Name = "cbREcherche"
        Me.cbREcherche.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbREcherche.Size = New System.Drawing.Size(25, 17)
        Me.cbREcherche.TabIndex = 1
        Me.cbREcherche.Text = "..."
        '
        'tbCode
        '
        Me.tbCode.AutoSize = False
        Me.tbCode.BackColor = System.Drawing.SystemColors.Window
        Me.tbCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCode.Location = New System.Drawing.Point(168, 32)
        Me.tbCode.MaxLength = 0
        Me.tbCode.Name = "tbCode"
        Me.tbCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCode.Size = New System.Drawing.Size(56, 19)
        Me.tbCode.TabIndex = 0
        Me.tbCode.Text = ""
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(144, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Code SousCommande"
        '
        'cbAnnuler
        '
        Me.cbAnnuler.BackColor = System.Drawing.SystemColors.Control
        Me.cbAnnuler.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAnnuler.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAnnuler.Location = New System.Drawing.Point(672, 304)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAnnuler.Size = New System.Drawing.Size(81, 25)
        Me.cbAnnuler.TabIndex = 4
        Me.cbAnnuler.Text = "&Annuler"
        '
        'cbValider
        '
        Me.cbValider.BackColor = System.Drawing.SystemColors.Control
        Me.cbValider.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbValider.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbValider.Location = New System.Drawing.Point(576, 304)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbValider.Size = New System.Drawing.Size(81, 25)
        Me.cbValider.TabIndex = 3
        Me.cbValider.Text = "&Valider"
        '
        'grpSCMD
        '
        Me.grpSCMD.Controls.Add(Me.liClient)
        Me.grpSCMD.Controls.Add(Me.tbMontantComm)
        Me.grpSCMD.Controls.Add(Me.Label9)
        Me.grpSCMD.Controls.Add(Me.tbBaseComm)
        Me.grpSCMD.Controls.Add(Me.Label8)
        Me.grpSCMD.Controls.Add(Me.tbTotalHTFacture)
        Me.grpSCMD.Controls.Add(Me.Label7)
        Me.grpSCMD.Controls.Add(Me.dtDateFactFourn)
        Me.grpSCMD.Controls.Add(Me.Label6)
        Me.grpSCMD.Controls.Add(Me.tbCodeFactFourn)
        Me.grpSCMD.Controls.Add(Me.Label5)
        Me.grpSCMD.Controls.Add(Me.Label4)
        Me.grpSCMD.Controls.Add(Me.dtDateCommande)
        Me.grpSCMD.Controls.Add(Me.Label3)
        Me.grpSCMD.Controls.Add(Me.tbCodeCommande)
        Me.grpSCMD.Controls.Add(Me.Label1)
        Me.grpSCMD.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpSCMD.Location = New System.Drawing.Point(8, 64)
        Me.grpSCMD.Name = "grpSCMD"
        Me.grpSCMD.Size = New System.Drawing.Size(552, 272)
        Me.grpSCMD.TabIndex = 5
        Me.grpSCMD.TabStop = False
        '
        'liClient
        '
        Me.liClient.Location = New System.Drawing.Point(120, 56)
        Me.liClient.Name = "liClient"
        Me.liClient.Size = New System.Drawing.Size(424, 24)
        Me.liClient.TabIndex = 16
        Me.liClient.TabStop = True
        Me.liClient.Text = "LinkLabel1"
        '
        'tbMontantComm
        '
        Me.tbMontantComm.Location = New System.Drawing.Point(120, 192)
        Me.tbMontantComm.Name = "tbMontantComm"
        Me.tbMontantComm.Size = New System.Drawing.Size(120, 20)
        Me.tbMontantComm.TabIndex = 15
        Me.tbMontantComm.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(8, 192)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(104, 24)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "Montant Comm"
        '
        'tbBaseComm
        '
        Me.tbBaseComm.Location = New System.Drawing.Point(120, 160)
        Me.tbBaseComm.Name = "tbBaseComm"
        Me.tbBaseComm.Size = New System.Drawing.Size(120, 20)
        Me.tbBaseComm.TabIndex = 13
        Me.tbBaseComm.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 160)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(104, 24)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "Base Comm"
        '
        'tbTotalHTFacture
        '
        Me.tbTotalHTFacture.Location = New System.Drawing.Point(120, 128)
        Me.tbTotalHTFacture.Name = "tbTotalHTFacture"
        Me.tbTotalHTFacture.Size = New System.Drawing.Size(120, 20)
        Me.tbTotalHTFacture.TabIndex = 11
        Me.tbTotalHTFacture.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 128)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 24)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "HT Facturé"
        '
        'dtDateFactFourn
        '
        Me.dtDateFactFourn.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtDateFactFourn.Location = New System.Drawing.Point(360, 88)
        Me.dtDateFactFourn.Name = "dtDateFactFourn"
        Me.dtDateFactFourn.Size = New System.Drawing.Size(120, 20)
        Me.dtDateFactFourn.TabIndex = 9
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(248, 88)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "date facture"
        '
        'tbCodeFactFourn
        '
        Me.tbCodeFactFourn.Location = New System.Drawing.Point(120, 88)
        Me.tbCodeFactFourn.Name = "tbCodeFactFourn"
        Me.tbCodeFactFourn.Size = New System.Drawing.Size(120, 20)
        Me.tbCodeFactFourn.TabIndex = 7
        Me.tbCodeFactFourn.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(104, 24)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Facture Fournisseur"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 16)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Client"
        '
        'dtDateCommande
        '
        Me.dtDateCommande.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtDateCommande.Location = New System.Drawing.Point(360, 24)
        Me.dtDateCommande.Name = "dtDateCommande"
        Me.dtDateCommande.Size = New System.Drawing.Size(120, 20)
        Me.dtDateCommande.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(248, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Date commande"
        '
        'tbCodeCommande
        '
        Me.tbCodeCommande.Location = New System.Drawing.Point(120, 24)
        Me.tbCodeCommande.Name = "tbCodeCommande"
        Me.tbCodeCommande.Size = New System.Drawing.Size(120, 20)
        Me.tbCodeCommande.TabIndex = 1
        Me.tbCodeCommande.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Code Commande"
        '
        'dlgLgFactCom
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(762, 352)
        Me.Controls.Add(Me.grpSCMD)
        Me.Controls.Add(Me.tbCode)
        Me.Controls.Add(Me.cbREcherche)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.cbValider)
        Me.Controls.Add(Me.Label2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.Green
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(184, 250)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgLgFactCom"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Ligne de commission"
        Me.grpSCMD.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Accesseurs"
    Public Sub New()
        MyBase.New()
        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        m_bModif = False
    End Sub

    'Public Property bModif() As Boolean
    '    Get
    '        Return m_bModif
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        m_bModif = Value
    '    End Set
    'End Property

    Public Sub setElementCourant(ByRef p_objLG As LgCommande)
        Dim objSCMD As SousCommande
        Debug.Assert(Not p_objLG Is Nothing)
        m_elementCourant = p_objLG
        Try
            If m_elementCourant.idSCmd <> 0 Then
                objSCMD = SousCommande.createandload(m_elementCourant.idSCmd)
                If objSCMD.id <> 0 Then
                    m_SCMDCourante = objSCMD
                End If
            End If

        Catch ex As Exception
            m_SCMDCourante = Nothing
        End Try
        AfficheElement()
    End Sub 'set element Courant
    Public Function getElementCourant() As LgCommande
        Return m_elementCourant
    End Function
    Public Sub setFournisseurCourant(ByVal poFourn As Fournisseur)
        Debug.Assert(Not poFourn Is Nothing, "Fournisseur Courant Non renseigné")
        m_FournisseurCourant = poFourn
    End Sub
#End Region

#Region "Méthodes"
    'Affichagde de l'élement courant
    Private Sub AfficheElement()
        Debug.Assert(Not m_elementCourant Is Nothing)
        Dim objCommande As CommandeClient
        m_bAffichageEncours = True
        Try
            If m_SCMDCourante Is Nothing Then
                grpSCMD.Enabled = False
            Else
                grpSCMD.Enabled = True
                tbCode.Text = m_SCMDCourante.code
                objCommande = CommandeClient.createandload(m_SCMDCourante.idCommandeClient)
                If objCommande.id <> 0 Then
                    tbCodeCommande.Text = objCommande.code
                    dtDateCommande.Value = objCommande.dateCommande
                    liClient.Text = objCommande.oTiers.shortResume
                    liClient.Tag = objCommande.oTiers.id
                    tbCodeFactFourn.Text = m_SCMDCourante.refFactFournisseur
                    dtDateFactFourn.Value = m_SCMDCourante.dateFactFournisseur
                    tbTotalHTFacture.Text = m_SCMDCourante.totalHTFacture
                    tbBaseComm.Text = m_SCMDCourante.baseCommission
                    tbMontantComm.Text = m_SCMDCourante.MontantCommission
                    cbValider.Enabled = True
                End If
            End If
        Catch ex As Exception

        End Try

        m_bAffichageEncours = False
    End Sub
    Private Sub rechercheSousCommande()
        Debug.Assert(Not m_elementCourant Is Nothing)
        Debug.Assert(Not m_FournisseurCourant Is Nothing, "Fournisseur Courant non renseigné")
        Dim objSCMD As SousCommande
        Dim frm As frmRechercheSCMD
        objSCMD = Nothing


        'Création de la fenêtre de recherche
        frm = New frmRechercheSCMD
        frm.setCodeFourn(m_FournisseurCourant.code)
        'Affichage de la fenêtre
        If frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Si on sort par OK
            objSCMD = frm.getElementSelectionne()
        End If
        If Not objSCMD Is Nothing Then
            m_elementCourant.idSCmd = objSCMD.id
            m_SCMDCourante = objSCMD
            AfficheElement()
        End If

    End Sub 'RechercheSousCommande

#End Region

#Region "Gestion des Evénements"
    Private Sub cbValider_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbValider.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub cbAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnuler.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgLgCommande_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(27) Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel

            Me.Close()
        End If
    End Sub

    Private Sub tbCode_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbCode.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            rechercheSousCommande()
        End If
    End Sub
    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbREcherche.Click
        rechercheSousCommande()
    End Sub
    Private Sub dlgLgCommande_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.tbCode.Focus()
        Me.cbValider.Enabled = False
    End Sub
#End Region

End Class
