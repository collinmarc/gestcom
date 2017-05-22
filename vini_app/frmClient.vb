Imports vini_DB
Public Class frmClient
    Inherits frmTiers



#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        m_TypeDonnees = vncEnums.vncTypeDonnee.CLIENT
        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

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
    Public WithEvents cboTypeClient As System.Windows.Forms.ComboBox
    'Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents cbxCodeTarif As System.Windows.Forms.ComboBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label231 As System.Windows.Forms.Label
    Friend WithEvents cbPrecommande As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cboTypeClient = New System.Windows.Forms.ComboBox
        Me.Label41 = New System.Windows.Forms.Label
        Me.cbPrecommande = New System.Windows.Forms.Button
        Me.Label231 = New System.Windows.Forms.Label
        Me.cbxCodeTarif = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'cboTypeClient
        '
        Me.cboTypeClient.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboTypeClient.BackColor = System.Drawing.SystemColors.Window
        Me.cboTypeClient.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboTypeClient.DisplayMember = "RQ_TypeClient.PAR_ID"
        Me.cboTypeClient.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTypeClient.Enabled = False
        Me.cboTypeClient.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTypeClient.Location = New System.Drawing.Point(784, 8)
        Me.cboTypeClient.Name = "cboTypeClient"
        Me.cboTypeClient.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboTypeClient.Size = New System.Drawing.Size(176, 21)
        Me.cboTypeClient.TabIndex = 2
        Me.cboTypeClient.ValueMember = "RQ_TypeClient.PAR_ID"
        '
        'Label41
        '
        Me.Label41.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label41.BackColor = System.Drawing.SystemColors.Control
        Me.Label41.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label41.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label41.Location = New System.Drawing.Point(672, 8)
        Me.Label41.Name = "Label41"
        Me.Label41.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label41.Size = New System.Drawing.Size(89, 17)
        Me.Label41.TabIndex = 16
        Me.Label41.Text = "Type de Client"
        '
        'cbPrecommande
        '
        Me.cbPrecommande.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbPrecommande.CausesValidation = False
        Me.cbPrecommande.Enabled = False
        Me.cbPrecommande.Location = New System.Drawing.Point(784, 80)
        Me.cbPrecommande.Name = "cbPrecommande"
        Me.cbPrecommande.Size = New System.Drawing.Size(176, 24)
        Me.cbPrecommande.TabIndex = 70
        Me.cbPrecommande.Text = "Pré-&Commande"
        '
        'Label231
        '
        Me.Label231.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label231.AutoSize = True
        Me.Label231.Location = New System.Drawing.Point(781, 56)
        Me.Label231.Name = "Label231"
        Me.Label231.Size = New System.Drawing.Size(56, 13)
        Me.Label231.TabIndex = 71
        Me.Label231.Text = "Code Tarif"
        '
        'cbxCodeTarif
        '
        Me.cbxCodeTarif.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxCodeTarif.FormattingEnabled = True
        Me.cbxCodeTarif.Items.AddRange(New Object() {"A", "B", "C"})
        Me.cbxCodeTarif.Location = New System.Drawing.Point(842, 53)
        Me.cbxCodeTarif.Name = "cbxCodeTarif"
        Me.cbxCodeTarif.Size = New System.Drawing.Size(117, 21)
        Me.cbxCodeTarif.TabIndex = 72
        '
        'frmClient
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(968, 734)
        Me.Controls.Add(Me.cbxCodeTarif)
        Me.Controls.Add(Me.Label231)
        Me.Controls.Add(Me.cbPrecommande)
        Me.Controls.Add(Me.Label41)
        Me.Controls.Add(Me.cboTypeClient)
        Me.Name = "frmClient"
        Me.Text = "Gestion des clients"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.tbCode, 0)
        Me.Controls.SetChildIndex(Me.tbRaisonSociale, 0)
        Me.Controls.SetChildIndex(Me.cboTypeClient, 0)
        Me.Controls.SetChildIndex(Me.Label41, 0)
        Me.Controls.SetChildIndex(Me.cbPrecommande, 0)
        Me.Controls.SetChildIndex(Me.Label231, 0)
        Me.Controls.SetChildIndex(Me.cbxCodeTarif, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Protected Overrides Function creerElement() As Boolean
        Debug.Assert(Not isfrmUpdated(), "La fenetre n'est pas libre")

        setElementCourant2(New Client("", ""))

        Return True
    End Function

    Public Overrides Function getResume() As String 'Rend le caption de la fenêtre
        If getElementCourant() Is Nothing Then
            Return "Gestion des Clients"
        Else
            Return "Client : " & getElementCourant().shortResume
        End If
    End Function 'getResume

    Public Overrides Function AfficheElement() As Boolean
        Dim objclient As Client
        Dim objParam As Param
        Dim bReturn As Boolean
        Dim str As String

        Debug.Assert(getElementCourant.GetType().Name.Equals("Client"), "Objet de type Client requis")


        bReturn = MyBase.AfficheElement()
        If bReturn Then
            objclient = CType(getElementCourant(), Client)

            For Each objParam In cboTypeClient.Items
                If objParam.id = objclient.idTypeClient Then
                    cboTypeClient.SelectedItem = objParam
                    Exit For
                End If
            Next
            For Each str In cbxCodeTarif.Items
                If str.Equals(objclient.CodeTarif) Then
                    cbxCodeTarif.SelectedItem = str
                    Exit For
                End If
            Next
        End If
        Return bReturn
    End Function 'AfficheElementCourant

    Public Overrides Function MAJElement() As Boolean
        Dim bReturn As Boolean
        Dim objClient As Client
        bReturn = MyBase.MAJElement()
        If bReturn Then

            objClient = getElementCourant()
            Try
                objClient.idTypeClient = cboTypeClient.SelectedItem.id
                objClient.CodeTarif = cbxCodeTarif.SelectedItem
                bReturn = True
            Catch ex As Exception
                DisplayError("frmClientTab.MAJElement", ex.ToString())
                bReturn = False
            End Try
        End If
        Return bReturn
    End Function 'MAJElementCourant

    Private Sub cbPrecommande_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPrecommande.Click
        Debug.Assert(Not getElementCourant() Is Nothing)
        Dim ofrmPreCommande As frmPrecommande
        Dim bOk As Boolean = True

        If isfrmUpdated() Then
            If MsgBox("L'élement courant a été modifié, souhaitez-vous l'enregister?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                If frmSave() Then
                    bOk = True
                Else
                    bOk = False
                End If
            Else
                bOk = False
            End If
        End If

        If bOk Then
            ofrmPreCommande = New frmPrecommande
            'Le parent MDI en le même que celui de la fenêtre courante
            ofrmPreCommande.MdiParent = MdiParent
            'Affiche la fenêtre 
            'il faut afficher la fentre avant l'init afin de déclencher l'évènement Load => initFenetre
            ofrmPreCommande.Show()
            'L'élément courant est le client courant
            If ofrmPreCommande.setElementCourant2(getElementCourant()) Then
                ofrmPreCommande.AfficheElementCourant()
            End If
        End If
    End Sub

    Private Sub initFenetre()
        Dim objTypeClient As Param
        debAffiche()
        laModeReglmt.Text = "Commande"
        laModeRglmt1.Text = "Transport"
        LaModeReglmt2.Visible = False
        cboModeReglement2.Visible = False

        cboTypeClient.DisplayMember = "Valeur"
        cboTypeClient.ValueMember = "id"
        For Each objTypeClient In Param.colTypeClient
            cboTypeClient.Items.Add(objTypeClient)
        Next
        finAffiche()
    End Sub

    Private Sub frmClientTab_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub

    Public Overrides Function ControleAvantSauvegarde() As Boolean
        Dim bReturn As Boolean
        bReturn = True
        If getElementCourant().bNew = True Then
            'controle de l'unicité du code
            If Client.getListe(tbCode.Text).Count > 0 Then
                DisplayError("Controleavantsauvegarde", "Le code Client doit être unique")
                bReturn = False
            End If
        End If
        bReturn = bReturn And MyBase.ControleAvantSauvegarde()
        Return bReturn

    End Function
End Class
