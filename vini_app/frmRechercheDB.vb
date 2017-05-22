Imports vini_DB
Imports System.Windows.Forms.Cursors

Public Class frmRechercheDB
    Inherits System.Windows.Forms.Form
    Private oldCursor As Cursor
    Private m_TypeDonnees As vncTypeDonnee
    Private m_ElementSelectionne As Persist
    Private m_ocol As Collection
    Private m_idPrecommande As Long
    Private m_idFournisseur As Long
    Friend WithEvents m_bsrc As System.Windows.Forms.BindingSource
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Private m_typeProduit As vncTypeProduit

#Region " Code généré par le Concepteur Windows Form "


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
    Friend WithEvents laCode As System.Windows.Forms.Label
    Friend WithEvents tbCode As System.Windows.Forms.TextBox
    Friend WithEvents tbMotCle As System.Windows.Forms.TextBox
    Friend WithEvents cbAfficher As System.Windows.Forms.Button
    Friend WithEvents tbNom As System.Windows.Forms.TextBox
    Friend WithEvents tbSelectionner As System.Windows.Forms.Button
    Friend WithEvents cbAnnuler As System.Windows.Forms.Button
    Friend WithEvents laNom As System.Windows.Forms.Label
    Friend WithEvents laMotCle As System.Windows.Forms.Label
    Friend WithEvents cboEtat As System.Windows.Forms.ComboBox
    Friend WithEvents laEtat As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.laCode = New System.Windows.Forms.Label
        Me.tbCode = New System.Windows.Forms.TextBox
        Me.laNom = New System.Windows.Forms.Label
        Me.tbNom = New System.Windows.Forms.TextBox
        Me.laMotCle = New System.Windows.Forms.Label
        Me.tbMotCle = New System.Windows.Forms.TextBox
        Me.cbAfficher = New System.Windows.Forms.Button
        Me.tbSelectionner = New System.Windows.Forms.Button
        Me.cbAnnuler = New System.Windows.Forms.Button
        Me.laEtat = New System.Windows.Forms.Label
        Me.cboEtat = New System.Windows.Forms.ComboBox
        Me.m_bsrc = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        CType(Me.m_bsrc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'laCode
        '
        Me.laCode.Location = New System.Drawing.Point(0, 8)
        Me.laCode.Name = "laCode"
        Me.laCode.Size = New System.Drawing.Size(72, 24)
        Me.laCode.TabIndex = 1
        Me.laCode.Text = "Code"
        '
        'tbCode
        '
        Me.tbCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbCode.Location = New System.Drawing.Point(80, 8)
        Me.tbCode.Name = "tbCode"
        Me.tbCode.Size = New System.Drawing.Size(368, 20)
        Me.tbCode.TabIndex = 1
        '
        'laNom
        '
        Me.laNom.Location = New System.Drawing.Point(0, 32)
        Me.laNom.Name = "laNom"
        Me.laNom.Size = New System.Drawing.Size(80, 16)
        Me.laNom.TabIndex = 3
        Me.laNom.Text = "Nom"
        '
        'tbNom
        '
        Me.tbNom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbNom.Location = New System.Drawing.Point(80, 32)
        Me.tbNom.Name = "tbNom"
        Me.tbNom.Size = New System.Drawing.Size(368, 20)
        Me.tbNom.TabIndex = 2
        '
        'laMotCle
        '
        Me.laMotCle.Location = New System.Drawing.Point(0, 56)
        Me.laMotCle.Name = "laMotCle"
        Me.laMotCle.Size = New System.Drawing.Size(80, 24)
        Me.laMotCle.TabIndex = 5
        Me.laMotCle.Text = "Mot Clé"
        '
        'tbMotCle
        '
        Me.tbMotCle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMotCle.Location = New System.Drawing.Point(80, 56)
        Me.tbMotCle.Name = "tbMotCle"
        Me.tbMotCle.Size = New System.Drawing.Size(368, 20)
        Me.tbMotCle.TabIndex = 3
        '
        'cbAfficher
        '
        Me.cbAfficher.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAfficher.Location = New System.Drawing.Point(312, 80)
        Me.cbAfficher.Name = "cbAfficher"
        Me.cbAfficher.Size = New System.Drawing.Size(136, 21)
        Me.cbAfficher.TabIndex = 5
        Me.cbAfficher.Text = "A&fficher"
        '
        'tbSelectionner
        '
        Me.tbSelectionner.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbSelectionner.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.tbSelectionner.Location = New System.Drawing.Point(232, 560)
        Me.tbSelectionner.Name = "tbSelectionner"
        Me.tbSelectionner.Size = New System.Drawing.Size(104, 23)
        Me.tbSelectionner.TabIndex = 7
        Me.tbSelectionner.Text = "&Sélectionner"
        '
        'cbAnnuler
        '
        Me.cbAnnuler.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAnnuler.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cbAnnuler.Location = New System.Drawing.Point(344, 560)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.Size = New System.Drawing.Size(104, 24)
        Me.cbAnnuler.TabIndex = 8
        Me.cbAnnuler.Text = "&Annuler"
        '
        'laEtat
        '
        Me.laEtat.Location = New System.Drawing.Point(0, 83)
        Me.laEtat.Name = "laEtat"
        Me.laEtat.Size = New System.Drawing.Size(64, 23)
        Me.laEtat.TabIndex = 8
        Me.laEtat.Text = "Etat"
        '
        'cboEtat
        '
        Me.cboEtat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboEtat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEtat.Location = New System.Drawing.Point(80, 80)
        Me.cboEtat.Name = "cboEtat"
        Me.cboEtat.Size = New System.Drawing.Size(121, 21)
        Me.cboEtat.TabIndex = 4
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(3, 107)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidth = 10
        Me.DataGridView1.Size = New System.Drawing.Size(441, 447)
        Me.DataGridView1.TabIndex = 9
        Me.DataGridView1.VirtualMode = True
        '
        'frmRechercheDB
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(456, 590)
        Me.ControlBox = False
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.cboEtat)
        Me.Controls.Add(Me.laEtat)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.tbSelectionner)
        Me.Controls.Add(Me.cbAfficher)
        Me.Controls.Add(Me.tbMotCle)
        Me.Controls.Add(Me.tbNom)
        Me.Controls.Add(Me.tbCode)
        Me.Controls.Add(Me.laMotCle)
        Me.Controls.Add(Me.laNom)
        Me.Controls.Add(Me.laCode)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRechercheDB"
        Me.ShowInTaskbar = False
        Me.Text = "Recherche"
        CType(Me.m_bsrc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
#Region "Accesseurs"
    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()
        m_idPrecommande = 0
        m_idFournisseur = 0
        m_typeProduit = False

    End Sub

    Public Sub setTypeDonnees(ByVal a As vncTypeDonnee)
        m_TypeDonnees = a
        Select Case m_TypeDonnees
            Case vncTypeDonnee.CLIENT
                Me.m_bsrc.DataSource = GetType(vini_DB.Client)
            Case vncTypeDonnee.FOURNISSEUR
                Me.m_bsrc.DataSource = GetType(vini_DB.Fournisseur)
            Case vncTypeDonnee.PRODUIT
                Me.m_bsrc.DataSource = GetType(vini_DB.Produit)
            Case vncTypeDonnee.COMMANDECLIENT
                Me.m_bsrc.DataSource = GetType(vini_DB.CommandeClient)
            Case vncTypeDonnee.BA
                Me.m_bsrc.DataSource = GetType(vini_DB.BonAppro)
            Case vncTypeDonnee.FACTCOMM
                Me.m_bsrc.DataSource = GetType(vini_DB.FactCom)
            Case vncTypeDonnee.FACTTRP
                Me.m_bsrc.DataSource = GetType(vini_DB.FactTRP)
            Case vncTypeDonnee.FACTCOL
                Me.m_bsrc.DataSource = GetType(vini_DB.FactColisage)
            Case vncTypeDonnee.FACTCOMM_NONREGLEE
                Me.m_bsrc.DataSource = GetType(vini_DB.FactCom)
            Case vncTypeDonnee.FACTTRP_NONREGLEE
                Me.m_bsrc.DataSource = GetType(vini_DB.FactTRP)
            Case vncTypeDonnee.FACTCOL_NONREGLEE
                Me.m_bsrc.DataSource = GetType(vini_DB.FactColisage)
        End Select
        Select Case m_TypeDonnees
            Case vncEnums.vncTypeDonnee.CLIENT
                laCode.Text = "Code"
                laNom.Text = "Nom"
                laMotCle.Text = "Raison Sociale"
                Me.Text = "Recherche de Client"
                laEtat.Enabled = False
                cboEtat.Enabled = False
            Case vncEnums.vncTypeDonnee.FOURNISSEUR
                laCode.Text = "Code"
                laNom.Text = "Nom"
                laMotCle.Text = "Raison Sociale"
                Me.Text = "Recherche de Fournisseur"
                laEtat.Enabled = False
                cboEtat.Enabled = False
            Case vncEnums.vncTypeDonnee.PRODUIT
                laCode.Text = "Code"
                laNom.Text = "Désignation"
                laMotCle.Text = "Mot Clé"
                Me.Text = "Recherche de Produit"
                laEtat.Enabled = False
                cboEtat.Enabled = False
            Case vncEnums.vncTypeDonnee.COMMANDECLIENT
                laCode.Text = "Code"
                laNom.Text = "Nom Client"
                laMotCle.Enabled = False
                tbMotCle.Enabled = False
                Me.Text = "Recherche de Commande"
                laEtat.Enabled = True
                cboEtat.Enabled = True
            Case vncEnums.vncTypeDonnee.BA
                laCode.Text = "Code"
                laNom.Text = "Nom Fourn"
                laMotCle.Enabled = False
                tbMotCle.Enabled = False
                Me.Text = "Recherche de Bon Appro"
                laEtat.Enabled = True
                cboEtat.Enabled = True
            Case vncEnums.vncTypeDonnee.FACTCOMM
                laCode.Text = "Code"
                laNom.Text = "Nom Fourn"
                laMotCle.Enabled = False
                tbMotCle.Enabled = False
                Me.Text = "Recherche de Facture de commission"
                laEtat.Enabled = True
                cboEtat.Enabled = True
            Case vncEnums.vncTypeDonnee.FACTCOL
                laCode.Text = "Code"
                laNom.Text = "Nom Fourn"
                laMotCle.Enabled = False
                tbMotCle.Enabled = False
                Me.Text = "Recherche de Facture de colisage"
                laEtat.Enabled = True
                cboEtat.Enabled = True
            Case vncEnums.vncTypeDonnee.FACTTRP
                laCode.Text = "Code"
                laNom.Text = "Nom Client"
                laMotCle.Enabled = False
                tbMotCle.Enabled = False
                Me.Text = "Recherche de Facture de Transport"
                laEtat.Enabled = True
                cboEtat.Enabled = True
            Case vncEnums.vncTypeDonnee.FACTCOMM_NONREGLEE
                laCode.Text = "Code"
                laNom.Text = "Nom Fourn"
                laMotCle.Enabled = False
                tbMotCle.Enabled = False
                Me.Text = "Recherche de Facture de commission non réglée"
                laEtat.Enabled = False
                cboEtat.Enabled = False
            Case vncEnums.vncTypeDonnee.FACTCOL_NONREGLEE
                laCode.Text = "Code"
                laNom.Text = "Nom Fourn"
                laMotCle.Enabled = False
                tbMotCle.Enabled = False
                Me.Text = "Recherche de Facture de colisage non réglée"
                laEtat.Enabled = False
                cboEtat.Enabled = False
            Case vncEnums.vncTypeDonnee.FACTTRP_NONREGLEE
                laCode.Text = "Code"
                laNom.Text = "Nom Client"
                laMotCle.Enabled = False
                tbMotCle.Enabled = False
                Me.Text = "Recherche de Facture de Transport non réglée"
                laEtat.Enabled = False
                cboEtat.Enabled = False
        End Select

        'Initialisation de la combo Etat en fonction du type de données
        cboEtat.Items.Clear()
        cboEtat.DisplayMember = "libelle"
        cboEtat.ValueMember = "codeEtat"
        Dim objEtat As EtatCommande

        Select Case m_TypeDonnees
            Case vncEnums.vncTypeDonnee.COMMANDECLIENT
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncRien)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncEnCoursSaisie)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncValidee)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncLivree)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncEclatee)
                cboEtat.Items.Add(objEtat)
                cboEtat.SelectedIndex = 0
            Case vncEnums.vncTypeDonnee.BA
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncRien)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncBAEnCours)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncBALivre)
                cboEtat.Items.Add(objEtat)
                cboEtat.SelectedIndex = 0
            Case vncEnums.vncTypeDonnee.FACTCOMM
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncRien)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactComGeneree)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactComExportee)
                cboEtat.Items.Add(objEtat)
                cboEtat.SelectedIndex = 0
            Case vncEnums.vncTypeDonnee.FACTTRP
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncRien)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactTRPGeneree)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactTRPExportee)
                cboEtat.Items.Add(objEtat)
                cboEtat.SelectedIndex = 0
            Case vncEnums.vncTypeDonnee.FACTCOL
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncRien)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactCOLGeneree)
                cboEtat.Items.Add(objEtat)
                objEtat = EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactCOLExportee)
                cboEtat.Items.Add(objEtat)
                cboEtat.SelectedIndex = 0
        End Select


    End Sub

    Public Sub setidPrecommande(ByVal p_id As Integer)
        Debug.Assert(m_TypeDonnees = vncEnums.vncTypeDonnee.PRODUIT)
        m_idPrecommande = p_id
    End Sub
    Public Sub setidFournisseur(ByVal p_id As Integer)
        Debug.Assert(m_TypeDonnees = vncEnums.vncTypeDonnee.PRODUIT)
        m_idFournisseur = p_id
    End Sub
    Public Sub setTypeProduit2(ByVal ptypeProduit As vncTypeProduit)
        Debug.Assert(m_TypeDonnees = vncEnums.vncTypeDonnee.PRODUIT)
        m_typeProduit = ptypeProduit
    End Sub

    Public Function getElementSelectionne() As Persist
        Return m_ElementSelectionne
    End Function
#End Region
    Private Sub setcursorWait()
        oldCursor = Me.Cursor
        Me.Cursor = WaitCursor
    End Sub
    Private Sub restoreCursor()
        Me.Cursor = oldCursor
    End Sub

    Private Sub cbAfficher_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAfficher.Click
        Afficher()
    End Sub
    Private Sub Afficher()
        setcursorWait()
        Select Case m_TypeDonnees
            Case vncTypeDonnee.CLIENT
                Me.m_bsrc.DataSource = GetType(vini_DB.Client)
                m_ocol = Client.getListe(tbCode.Text, tbNom.Text, tbMotCle.Text)
            Case vncTypeDonnee.FOURNISSEUR
                Me.m_bsrc.DataSource = GetType(vini_DB.Fournisseur)
                m_ocol = Fournisseur.getListe(tbCode.Text, tbNom.Text, tbMotCle.Text)
            Case vncTypeDonnee.PRODUIT
                Me.m_bsrc.DataSource = GetType(vini_DB.Produit)
                'Recherche des produits , les valeurs idFournisseur et idPrecommande sont initialisées à 0
                m_ocol = Produit.getListe(m_typeProduit, tbCode.Text, tbNom.Text, tbMotCle.Text, m_idFournisseur, m_idPrecommande)
            Case vncTypeDonnee.COMMANDECLIENT
                Me.m_bsrc.DataSource = GetType(vini_DB.CommandeClient)
                If (tbCode.Text = "" And tbNom.Text = "" And cboEtat.SelectedItem.codeEtat = vncEnums.vncEtatCommande.vncRien) Then
                    If MsgBox("Etes-vous sur ?", vbYesNo) = vbYes Then
                        m_ocol = CommandeClient.getListe(tbCode.Text, tbNom.Text, cboEtat.SelectedItem.codeEtat)
                    Else
                        Exit Sub
                    End If
                Else
                    m_ocol = CommandeClient.getListe(tbCode.Text, tbNom.Text, cboEtat.SelectedItem.codeEtat)
                End If
            Case vncTypeDonnee.BA
                Me.m_bsrc.DataSource = GetType(vini_DB.BonAppro)
                m_ocol = BonAppro.getListe(tbCode.Text, tbNom.Text, cboEtat.SelectedItem.codeEtat)
            Case vncTypeDonnee.FACTCOMM
                Me.m_bsrc.DataSource = GetType(vini_DB.FactCom)
                m_ocol = FactCom.getListe(tbCode.Text, tbNom.Text, cboEtat.SelectedItem.codeEtat)
            Case vncTypeDonnee.FACTTRP
                Me.m_bsrc.DataSource = GetType(vini_DB.FactTRP)
                m_ocol = FactTRP.getListe(tbCode.Text, tbNom.Text, cboEtat.SelectedItem.codeEtat)
            Case vncTypeDonnee.FACTCOL
                Me.m_bsrc.DataSource = GetType(vini_DB.FactColisage)
                m_ocol = FactColisage.getListe(tbCode.Text, tbNom.Text, cboEtat.SelectedItem.codeEtat)
            Case vncTypeDonnee.FACTCOMM_NONREGLEE
                Me.m_bsrc.DataSource = GetType(vini_DB.FactCom)
                m_ocol = FactCom.getListeNonReglee(tbCode.Text, tbNom.Text)
            Case vncTypeDonnee.FACTTRP_NONREGLEE
                Me.m_bsrc.DataSource = GetType(vini_DB.FactTRP)
                m_ocol = FactTRP.getListeNonReglee(tbCode.Text, tbNom.Text)
            Case vncTypeDonnee.FACTCOL_NONREGLEE
                Me.m_bsrc.DataSource = GetType(vini_DB.FactColisage)
                m_ocol = FactColisage.getListeNonReglee(tbCode.Text, tbNom.Text)
        End Select

        If Not m_ocol Is Nothing Then
            displayListe()
        Else
            MsgBox(CommandeClient.getErreur())
        End If
        restoreCursor()
    End Sub

    Private Sub SelectListBoxItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSelectionner.Click
        selectionneEtSort()
    End Sub
    Private Sub selectionneEtSort()

        If Not m_bsrc.Current Is Nothing Then
            Try
                m_ElementSelectionne = m_bsrc.Current
                m_ElementSelectionne.load()
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()

            Catch
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
            End Try
        End If

    End Sub 'selectionneEtSort
    Public Sub setListe(ByVal oCol As Collection)
        m_ocol = oCol
    End Sub
    Public Sub setCode(ByVal pcode As String)
        tbCode.Text = pcode
    End Sub
    Public Sub setNom(ByVal pnom As String)
        tbNom.Text = pnom
    End Sub
    Public Sub setMotCle(ByVal pmotCle As String)
        tbMotCle.Text = pmotCle
    End Sub

    Public Sub displayListe()
        Debug.Assert(Not m_ocol Is Nothing)
        Dim obj As Object
        Dim oCol As DataGridViewTextBoxColumn


        DataGridView1.DataSource = m_bsrc
        DataGridView1.Columns.Clear()
        Select Case m_TypeDonnees
            Case vncEnums.vncTypeDonnee.CLIENT, vncTypeDonnee.FOURNISSEUR

                oCol = New DataGridViewTextBoxColumn()
                oCol.DataPropertyName = "Code"
                oCol.HeaderText = "Code"
                oCol.Width = 100
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.DataPropertyName = "rs"
                oCol.HeaderText = "RS"
                oCol.Width = 250
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.Width = 100
                oCol.DataPropertyName = "nom"
                oCol.HeaderText = "Nom"
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.Width = 250
                oCol.DataPropertyName = "AdresseLivraisonVille"
                oCol.HeaderText = "Ville"
                DataGridView1.Columns.Add(oCol)

            Case vncEnums.vncTypeDonnee.PRODUIT
                oCol = New DataGridViewTextBoxColumn()
                oCol.DataPropertyName = "Code"
                oCol.HeaderText = "Code"
                oCol.FillWeight = 50
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.DataPropertyName = "nom"
                oCol.HeaderText = "Nom"
                oCol.FillWeight = 228
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 35
                oCol.DataPropertyName = "millesime"
                oCol.HeaderText = "Mil."
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 35
                oCol.DataPropertyName = "libcontenant"
                oCol.HeaderText = "cont."
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 35
                oCol.DataPropertyName = "libCouleur"
                oCol.HeaderText = "Coul."
                DataGridView1.Columns.Add(oCol)

            Case vncEnums.vncTypeDonnee.COMMANDECLIENT, vncEnums.vncTypeDonnee.BA
                oCol = New DataGridViewTextBoxColumn()
                oCol.DataPropertyName = "Code"
                oCol.HeaderText = "Code"
                oCol.FillWeight = 50
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.DataPropertyName = "TiersRS"
                oCol.HeaderText = "RS"
                oCol.FillWeight = 200
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 75
                oCol.DataPropertyName = "dateCommande"
                oCol.HeaderText = "Date"
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 47
                oCol.DataPropertyName = "TotalHT"
                oCol.HeaderText = "HT"
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 8
                oCol.DataPropertyName = "EtatLibelle"
                oCol.HeaderText = "Etat"
                DataGridView1.Columns.Add(oCol)

            Case vncEnums.vncTypeDonnee.FACTCOMM, vncTypeDonnee.FACTCOMM_NONREGLEE, vncTypeDonnee.FACTTRP, vncTypeDonnee.FACTTRP_NONREGLEE, vncTypeDonnee.FACTCOL, vncTypeDonnee.FACTCOL_NONREGLEE
                oCol = New DataGridViewTextBoxColumn()
                oCol.DataPropertyName = "Code"
                oCol.HeaderText = "Code"
                oCol.FillWeight = 50
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.DataPropertyName = "TiersRS"
                oCol.HeaderText = "RS"
                oCol.FillWeight = 200
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 75
                oCol.HeaderText = "Date"
                oCol.DataPropertyName = "dateCommande"
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 47
                oCol.HeaderText = "HT"
                oCol.DataPropertyName = "TotalHT"
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 47
                oCol.HeaderText = "TTC"
                oCol.DataPropertyName = "TotalTTC"
                DataGridView1.Columns.Add(oCol)

                oCol = New DataGridViewTextBoxColumn()
                oCol.FillWeight = 8
                oCol.HeaderText = "Etat"
                oCol.DataPropertyName = "EtatLibelle"
                DataGridView1.Columns.Add(oCol)
        End Select
        m_bsrc.Clear()
        For Each obj In m_ocol
            m_bsrc.Add(obj)
        Next obj


        If m_bsrc.Count > 0 Then
            m_bsrc.Position = 0
        End If
        m_bsrc.ResetBindings(False)
        DataGridView1.Focus()


    End Sub

    Private Sub tbCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbCode.KeyPress

        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            Afficher()
        End If
    End Sub



    Private Sub frmRechercheDB_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(27) Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If
    End Sub


    Private Sub tbMotCle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMotCle.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            Afficher()
        End If

    End Sub

    Private Sub tbNom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbNom.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            Afficher()
        End If

    End Sub

    Private Sub DataGridView1_RowHeaderMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseDoubleClick
        selectionneEtSort()
    End Sub

    Private Sub DataGridView1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridView1.KeyPress
    End Sub

    Private Sub DataGridView1_PreviewKeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles DataGridView1.PreviewKeyDown
        If e.KeyCode = Keys.Enter Then
            selectionneEtSort()
        End If
    End Sub

    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        selectionneEtSort()
    End Sub

    Private Sub frmRechercheDB_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
    End Sub
End Class
