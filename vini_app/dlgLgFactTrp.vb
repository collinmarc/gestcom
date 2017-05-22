Imports vini_DB
Friend Class dlgLgFactTRP
    Inherits System.Windows.Forms.Form
    Private m_elementCourant As LgFactTRP
    Private m_CMDCourante As CommandeClient
    Private m_TiersCourant As Client
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
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    Public WithEvents cbAnnuler As System.Windows.Forms.Button
    Public WithEvents cbValider As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtDateCommande As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents tbCode As System.Windows.Forms.TextBox
    Public WithEvents cbREcherche As System.Windows.Forms.Button
    Friend WithEvents tbREfLiv As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtDateLivraison As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbNomTransporteur As System.Windows.Forms.TextBox
    Friend WithEvents grpInfFactureTRP As System.Windows.Forms.GroupBox
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents cbCalcMontantTransport As System.Windows.Forms.Button
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents tbMontantTransport As vini_app.textBoxCurrency
    Friend WithEvents tbPUPallNonPrep As vini_app.textBoxCurrency
    Friend WithEvents tbPUPallPrep As vini_app.textBoxCurrency
    Friend WithEvents Label56 As System.Windows.Forms.Label
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents tbQtePallNonPrep As vini_app.textBoxNumeric
    Friend WithEvents tbQtePallPrep As vini_app.textBoxNumeric
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents tbPoids As vini_app.textBoxNumeric
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents tbQteColis As vini_app.textBoxNumeric
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents grpCMD As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cbREcherche = New System.Windows.Forms.Button()
        Me.tbCode = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbAnnuler = New System.Windows.Forms.Button()
        Me.cbValider = New System.Windows.Forms.Button()
        Me.grpCMD = New System.Windows.Forms.GroupBox()
        Me.grpInfFactureTRP = New System.Windows.Forms.GroupBox()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.cbCalcMontantTransport = New System.Windows.Forms.Button()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.tbMontantTransport = New vini_app.textBoxCurrency()
        Me.tbPUPallNonPrep = New vini_app.textBoxCurrency()
        Me.tbPUPallPrep = New vini_app.textBoxCurrency()
        Me.Label56 = New System.Windows.Forms.Label()
        Me.Label55 = New System.Windows.Forms.Label()
        Me.tbQtePallNonPrep = New vini_app.textBoxNumeric()
        Me.tbQtePallPrep = New vini_app.textBoxNumeric()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.tbPoids = New vini_app.textBoxNumeric()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.tbQteColis = New vini_app.textBoxNumeric()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbNomTransporteur = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtDateLivraison = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbREfLiv = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dtDateCommande = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.grpCMD.SuspendLayout()
        Me.grpInfFactureTRP.SuspendLayout()
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
        Me.cbREcherche.UseVisualStyleBackColor = False
        '
        'tbCode
        '
        Me.tbCode.BackColor = System.Drawing.SystemColors.Window
        Me.tbCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCode.Location = New System.Drawing.Point(168, 32)
        Me.tbCode.MaxLength = 0
        Me.tbCode.Name = "tbCode"
        Me.tbCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCode.Size = New System.Drawing.Size(56, 20)
        Me.tbCode.TabIndex = 0
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
        Me.Label2.Text = "Code Commande"
        '
        'cbAnnuler
        '
        Me.cbAnnuler.BackColor = System.Drawing.SystemColors.Control
        Me.cbAnnuler.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbAnnuler.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbAnnuler.Location = New System.Drawing.Point(664, 304)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbAnnuler.Size = New System.Drawing.Size(81, 25)
        Me.cbAnnuler.TabIndex = 5
        Me.cbAnnuler.Text = "&Annuler"
        Me.cbAnnuler.UseVisualStyleBackColor = False
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
        Me.cbValider.TabIndex = 4
        Me.cbValider.Text = "&Valider"
        Me.cbValider.UseVisualStyleBackColor = False
        '
        'grpCMD
        '
        Me.grpCMD.Controls.Add(Me.grpInfFactureTRP)
        Me.grpCMD.Controls.Add(Me.tbNomTransporteur)
        Me.grpCMD.Controls.Add(Me.Label4)
        Me.grpCMD.Controls.Add(Me.dtDateLivraison)
        Me.grpCMD.Controls.Add(Me.Label1)
        Me.grpCMD.Controls.Add(Me.tbREfLiv)
        Me.grpCMD.Controls.Add(Me.Label5)
        Me.grpCMD.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpCMD.Location = New System.Drawing.Point(8, 64)
        Me.grpCMD.Name = "grpCMD"
        Me.grpCMD.Size = New System.Drawing.Size(552, 272)
        Me.grpCMD.TabIndex = 3
        Me.grpCMD.TabStop = False
        '
        'grpInfFactureTRP
        '
        Me.grpInfFactureTRP.Controls.Add(Me.Label60)
        Me.grpInfFactureTRP.Controls.Add(Me.Label57)
        Me.grpInfFactureTRP.Controls.Add(Me.cbCalcMontantTransport)
        Me.grpInfFactureTRP.Controls.Add(Me.Label59)
        Me.grpInfFactureTRP.Controls.Add(Me.Label58)
        Me.grpInfFactureTRP.Controls.Add(Me.tbMontantTransport)
        Me.grpInfFactureTRP.Controls.Add(Me.tbPUPallNonPrep)
        Me.grpInfFactureTRP.Controls.Add(Me.tbPUPallPrep)
        Me.grpInfFactureTRP.Controls.Add(Me.Label56)
        Me.grpInfFactureTRP.Controls.Add(Me.Label55)
        Me.grpInfFactureTRP.Controls.Add(Me.tbQtePallNonPrep)
        Me.grpInfFactureTRP.Controls.Add(Me.tbQtePallPrep)
        Me.grpInfFactureTRP.Controls.Add(Me.Label54)
        Me.grpInfFactureTRP.Controls.Add(Me.Label53)
        Me.grpInfFactureTRP.Controls.Add(Me.tbPoids)
        Me.grpInfFactureTRP.Controls.Add(Me.Label52)
        Me.grpInfFactureTRP.Controls.Add(Me.tbQteColis)
        Me.grpInfFactureTRP.Controls.Add(Me.Label6)
        Me.grpInfFactureTRP.Location = New System.Drawing.Point(16, 120)
        Me.grpInfFactureTRP.Name = "grpInfFactureTRP"
        Me.grpInfFactureTRP.Size = New System.Drawing.Size(496, 120)
        Me.grpInfFactureTRP.TabIndex = 21
        Me.grpInfFactureTRP.TabStop = False
        Me.grpInfFactureTRP.Text = "Transport Facturé"
        '
        'Label60
        '
        Me.Label60.Location = New System.Drawing.Point(260, 42)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(80, 16)
        Me.Label60.TabIndex = 182
        Me.Label60.Text = "P.U."
        Me.Label60.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label57
        '
        Me.Label57.Location = New System.Drawing.Point(172, 42)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(56, 16)
        Me.Label57.TabIndex = 181
        Me.Label57.Text = "Qte"
        Me.Label57.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbCalcMontantTransport
        '
        Me.cbCalcMontantTransport.Location = New System.Drawing.Point(364, 66)
        Me.cbCalcMontantTransport.Name = "cbCalcMontantTransport"
        Me.cbCalcMontantTransport.Size = New System.Drawing.Size(16, 23)
        Me.cbCalcMontantTransport.TabIndex = 6
        Me.cbCalcMontantTransport.Text = "="
        '
        'Label59
        '
        Me.Label59.Location = New System.Drawing.Point(348, 82)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(8, 16)
        Me.Label59.TabIndex = 179
        Me.Label59.Text = "/"
        '
        'Label58
        '
        Me.Label58.Location = New System.Drawing.Point(348, 66)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(8, 23)
        Me.Label58.TabIndex = 6
        Me.Label58.Text = "\"
        '
        'tbMontantTransport
        '
        Me.tbMontantTransport.Location = New System.Drawing.Point(388, 66)
        Me.tbMontantTransport.Name = "tbMontantTransport"
        Me.tbMontantTransport.Size = New System.Drawing.Size(80, 20)
        Me.tbMontantTransport.TabIndex = 7
        Me.tbMontantTransport.Text = "0"
        '
        'tbPUPallNonPrep
        '
        Me.tbPUPallNonPrep.Location = New System.Drawing.Point(260, 82)
        Me.tbPUPallNonPrep.Name = "tbPUPallNonPrep"
        Me.tbPUPallNonPrep.Size = New System.Drawing.Size(80, 20)
        Me.tbPUPallNonPrep.TabIndex = 5
        Me.tbPUPallNonPrep.Text = "0"
        '
        'tbPUPallPrep
        '
        Me.tbPUPallPrep.Location = New System.Drawing.Point(260, 58)
        Me.tbPUPallPrep.Name = "tbPUPallPrep"
        Me.tbPUPallPrep.Size = New System.Drawing.Size(80, 20)
        Me.tbPUPallPrep.TabIndex = 3
        Me.tbPUPallPrep.Text = "0"
        '
        'Label56
        '
        Me.Label56.Location = New System.Drawing.Point(244, 82)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(16, 16)
        Me.Label56.TabIndex = 174
        Me.Label56.Text = "X"
        '
        'Label55
        '
        Me.Label55.Location = New System.Drawing.Point(244, 66)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(16, 16)
        Me.Label55.TabIndex = 173
        Me.Label55.Text = "X"
        '
        'tbQtePallNonPrep
        '
        Me.tbQtePallNonPrep.Location = New System.Drawing.Point(172, 82)
        Me.tbQtePallNonPrep.Name = "tbQtePallNonPrep"
        Me.tbQtePallNonPrep.Size = New System.Drawing.Size(64, 20)
        Me.tbQtePallNonPrep.TabIndex = 4
        Me.tbQtePallNonPrep.Text = "0"
        '
        'tbQtePallPrep
        '
        Me.tbQtePallPrep.Location = New System.Drawing.Point(172, 58)
        Me.tbQtePallPrep.Name = "tbQtePallPrep"
        Me.tbQtePallPrep.Size = New System.Drawing.Size(64, 20)
        Me.tbQtePallPrep.TabIndex = 2
        Me.tbQtePallPrep.Text = "0"
        '
        'Label54
        '
        Me.Label54.Location = New System.Drawing.Point(28, 82)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(128, 16)
        Me.Label54.TabIndex = 170
        Me.Label54.Text = "Pallettes non préparées"
        '
        'Label53
        '
        Me.Label53.Location = New System.Drawing.Point(28, 58)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(128, 16)
        Me.Label53.TabIndex = 169
        Me.Label53.Text = "Pallettes préparées"
        '
        'tbPoids
        '
        Me.tbPoids.Location = New System.Drawing.Point(260, 18)
        Me.tbPoids.Name = "tbPoids"
        Me.tbPoids.Size = New System.Drawing.Size(104, 20)
        Me.tbPoids.TabIndex = 1
        Me.tbPoids.Text = "0"
        '
        'Label52
        '
        Me.Label52.Location = New System.Drawing.Point(204, 18)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(40, 23)
        Me.Label52.TabIndex = 167
        Me.Label52.Text = "Poids"
        '
        'tbQteColis
        '
        Me.tbQteColis.Location = New System.Drawing.Point(100, 18)
        Me.tbQteColis.Name = "tbQteColis"
        Me.tbQteColis.Size = New System.Drawing.Size(88, 20)
        Me.tbQteColis.TabIndex = 0
        Me.tbQteColis.Text = "0"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(28, 18)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 23)
        Me.Label6.TabIndex = 165
        Me.Label6.Text = "Nbre Colis"
        '
        'tbNomTransporteur
        '
        Me.tbNomTransporteur.Location = New System.Drawing.Point(104, 80)
        Me.tbNomTransporteur.Name = "tbNomTransporteur"
        Me.tbNomTransporteur.Size = New System.Drawing.Size(440, 20)
        Me.tbNomTransporteur.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 23)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Transporteur"
        '
        'dtDateLivraison
        '
        Me.dtDateLivraison.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateLivraison.Location = New System.Drawing.Point(104, 16)
        Me.dtDateLivraison.Name = "dtDateLivraison"
        Me.dtDateLivraison.Size = New System.Drawing.Size(120, 20)
        Me.dtDateLivraison.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Date Livraison"
        '
        'tbREfLiv
        '
        Me.tbREfLiv.Location = New System.Drawing.Point(104, 48)
        Me.tbREfLiv.Name = "tbREfLiv"
        Me.tbREfLiv.Size = New System.Drawing.Size(440, 20)
        Me.tbREfLiv.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 24)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Recepissé Liv."
        '
        'dtDateCommande
        '
        Me.dtDateCommande.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateCommande.Location = New System.Drawing.Point(400, 32)
        Me.dtDateCommande.Name = "dtDateCommande"
        Me.dtDateCommande.Size = New System.Drawing.Size(120, 20)
        Me.dtDateCommande.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(288, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Date commande"
        '
        'dlgLgFactTRP
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(762, 352)
        Me.Controls.Add(Me.grpCMD)
        Me.Controls.Add(Me.tbCode)
        Me.Controls.Add(Me.cbREcherche)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.cbValider)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtDateCommande)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.Green
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(184, 250)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgLgFactTRP"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Ligne de Transport"
        Me.grpCMD.ResumeLayout(False)
        Me.grpCMD.PerformLayout()
        Me.grpInfFactureTRP.ResumeLayout(False)
        Me.grpInfFactureTRP.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

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

    Public Sub setElementCourant(ByRef p_objLG As LgFactTRP)
        Dim objCMD As Commande
        Debug.Assert(Not p_objLG Is Nothing)
        m_elementCourant = p_objLG
        Try
            If m_elementCourant.idCmdCLT <> 0 Then
                objCMD = CommandeClient.createandload(m_elementCourant.idCmdCLT)
                If objCMD.id <> 0 Then
                    m_CMDCourante = objCMD
                End If
            End If

        Catch ex As Exception
            m_CMDCourante = Nothing
        End Try
        AfficheElement()
    End Sub 'set element Courant
    Public Function getElementCourant() As LgFactTRP
        Return m_elementCourant
    End Function
    Public Sub setTiersCourant(ByVal poTiers As Client)
        Debug.Assert(Not poTiers Is Nothing, "Tiers Courant Non renseigné")
        m_TiersCourant = poTiers
    End Sub
#End Region

#Region "Méthodes"
    'Initialisation de l'élement Courant
    Private Sub MAJElement()
        Debug.Assert(Not m_elementCourant Is Nothing)

        If Not (m_CMDCourante Is Nothing) Then
            m_elementCourant.idCmdCLT = m_CMDCourante.id
        End If
        m_elementCourant.dateCommande = dtDateCommande.Text
        m_elementCourant.dateLivraison = dtDateLivraison.Text
        m_elementCourant.nomTransporteur = tbNomTransporteur.Text
        m_elementCourant.referenceLivraison = tbREfLiv.Text
        m_elementCourant.qteColis = tbQteColis.Text
        m_elementCourant.poids = tbPoids.Text
        m_elementCourant.qtePalettesNonPreparees = tbQtePallNonPrep.Text
        m_elementCourant.qtePalettesPreparees = tbQtePallPrep.Text
        m_elementCourant.puPalettesNonPreparees = tbPUPallNonPrep.Text
        m_elementCourant.puPalettesPreparees = tbPUPallPrep.Text
        m_elementCourant.prixHT = tbMontantTransport.Text
        m_elementCourant.calculPrixTotal()


    End Sub 'MAJElement
    'Affichagde de l'élement courant
    Private Sub AfficheElement()
        Debug.Assert(Not m_elementCourant Is Nothing)
        m_bAffichageEncours = True
        Try
            If m_CMDCourante Is Nothing Then
                tbCode.Text = ""
                dtDateCommande.Text = DATE_DEFAUT
                dtDateCommande.Enabled = False
            Else
                tbCode.Text = m_CMDCourante.code
                dtDateCommande.Text = m_elementCourant.dateCommande
                dtDateCommande.Enabled = True
            End If
            dtDateLivraison.Text = m_elementCourant.dateLivraison
            tbREfLiv.Text = m_elementCourant.referenceLivraison
            tbNomTransporteur.Text = m_elementCourant.nomTransporteur
            tbQteColis.Text = m_elementCourant.qteColis
            tbPoids.Text = m_elementCourant.poids
            tbQtePallNonPrep.Text = m_elementCourant.qtePalettesNonPreparees
            tbQtePallPrep.Text = m_elementCourant.qtePalettesPreparees
            tbPUPallNonPrep.Text = m_elementCourant.puPalettesNonPreparees
            tbPUPallPrep.Text = m_elementCourant.puPalettesPreparees
            tbMontantTransport.Text = m_elementCourant.prixHT

            dtDateCommande.Text = m_elementCourant.dateCommande
        Catch ex As Exception

        End Try

        m_bAffichageEncours = False
    End Sub
    Private Sub rechercheCommande()
        Debug.Assert(Not m_elementCourant Is Nothing)
        Debug.Assert(Not m_TiersCourant Is Nothing, "Fournisseur Courant non renseigné")
        Dim objCMD As CommandeClient
        Dim frm As frmRechercheCMDTRP
        objCMD = Nothing

        'Création de la fenêtre de recherche
        frm = New frmRechercheCMDTRP
        frm.setRSClient(m_TiersCourant.rs)
        'Affichage de la fenêtre
        If frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Si on sort par OK
            objCMD = frm.getElementSelectionne()
        End If
        If Not objCMD Is Nothing Then
            objCMD.load()
            m_elementCourant.dupliqueinfosCommande(objCMD)
            m_CMDCourante = objCMD
            AfficheElement()
        End If

    End Sub 'RechercheSousCommande
    Private Sub calculMontantTransport()

        Try
            m_elementCourant.calculPrixTotal()
        Catch ex As Exception

        End Try

    End Sub


#End Region

#Region "Gestion des Evénements"
    Private Sub cbValider_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbValider.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        MAJElement()
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
            rechercheCommande()
        End If
    End Sub
    Private Sub cbRecherche_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbREcherche.Click
        rechercheCommande()
    End Sub
    Private Sub dlgLgCommande_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.tbCode.Focus()
        Me.dtDateCommande.Enabled = False
    End Sub
#End Region

    Private Sub cbCalcMontantTransport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCalcMontantTransport.Click
        MAJElement()
        calculMontantTransport()
        tbMontantTransport.Text = m_elementCourant.prixHT
    End Sub

    Private Sub tbQtePallPrep_Validated(sender As System.Object, e As System.EventArgs) Handles tbQtePallPrep.Validated
        MAJElement()
        calculMontantTransport()
        tbMontantTransport.Text = m_elementCourant.prixHT
    End Sub

    Private Sub tbPUPallPrep_Validated(sender As System.Object, e As System.EventArgs) Handles tbPUPallPrep.Validated
        MAJElement()
        calculMontantTransport()
        tbMontantTransport.Text = m_elementCourant.prixHT

    End Sub



    Private Sub tbPoids_Validated(sender As System.Object, e As System.EventArgs) Handles tbPoids.Validated
        MAJElement()
        calculMontantTransport()
        tbMontantTransport.Text = m_elementCourant.prixHT

    End Sub

    Private Sub tbQteColis_Validated(sender As System.Object, e As System.EventArgs) Handles tbQteColis.Validated

        MAJElement()
        calculMontantTransport()
        tbMontantTransport.Text = m_elementCourant.prixHT
    End Sub

    Private Sub tbQtePallNonPrep_Validated(sender As System.Object, e As System.EventArgs) Handles tbQtePallNonPrep.Validated
        MAJElement()
        calculMontantTransport()
        tbMontantTransport.Text = m_elementCourant.prixHT

    End Sub


    Private Sub tbPUPallNonPrep_Validated(sender As System.Object, e As System.EventArgs) Handles tbPUPallNonPrep.Validated
        MAJElement()
        calculMontantTransport()
        tbMontantTransport.Text = m_elementCourant.prixHT

    End Sub
End Class
