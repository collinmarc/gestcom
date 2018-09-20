Option Strict Off
Imports vini_DB
Public Class frmControleStock
    Inherits frmStatistiques

#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()
        Me.CrystalReportViewer1.Visible = False

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
    Friend WithEvents pbProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents laTraitement As System.Windows.Forms.Label
    Friend WithEvents cbControle As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtdateDeb As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtdatefin As System.Windows.Forms.DateTimePicker
    Friend WithEvents lbErreurs As System.Windows.Forms.ListBox
    Friend WithEvents laProgress As System.Windows.Forms.Label
    Friend WithEvents ckRegenere As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbNumCmdClt As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbNumBA As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cbControle = New System.Windows.Forms.Button
        Me.pbProgressBar = New System.Windows.Forms.ProgressBar
        Me.laTraitement = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtdateDeb = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtdatefin = New System.Windows.Forms.DateTimePicker
        Me.lbErreurs = New System.Windows.Forms.ListBox
        Me.laProgress = New System.Windows.Forms.Label
        Me.ckRegenere = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbNumCmdClt = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbNumBA = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'cbControle
        '
        Me.cbControle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbControle.Location = New System.Drawing.Point(256, 48)
        Me.cbControle.Name = "cbControle"
        Me.cbControle.Size = New System.Drawing.Size(744, 32)
        Me.cbControle.TabIndex = 5
        Me.cbControle.Text = "&Contrôle des mouvements de stocks"
        '
        'pbProgressBar
        '
        Me.pbProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbProgressBar.Location = New System.Drawing.Point(8, 112)
        Me.pbProgressBar.Name = "pbProgressBar"
        Me.pbProgressBar.Size = New System.Drawing.Size(992, 24)
        Me.pbProgressBar.TabIndex = 2
        '
        'laTraitement
        '
        Me.laTraitement.Location = New System.Drawing.Point(256, 88)
        Me.laTraitement.Name = "laTraitement"
        Me.laTraitement.Size = New System.Drawing.Size(568, 16)
        Me.laTraitement.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Date &début"
        '
        'dtdateDeb
        '
        Me.dtdateDeb.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdateDeb.Location = New System.Drawing.Point(88, 0)
        Me.dtdateDeb.Name = "dtdateDeb"
        Me.dtdateDeb.Size = New System.Drawing.Size(96, 20)
        Me.dtdateDeb.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Date &fin"
        '
        'dtdatefin
        '
        Me.dtdatefin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtdatefin.Location = New System.Drawing.Point(88, 24)
        Me.dtdatefin.Name = "dtdatefin"
        Me.dtdatefin.Size = New System.Drawing.Size(96, 20)
        Me.dtdatefin.TabIndex = 1
        '
        'lbErreurs
        '
        Me.lbErreurs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbErreurs.HorizontalScrollbar = True
        Me.lbErreurs.Location = New System.Drawing.Point(8, 136)
        Me.lbErreurs.Name = "lbErreurs"
        Me.lbErreurs.Size = New System.Drawing.Size(984, 524)
        Me.lbErreurs.TabIndex = 6
        '
        'laProgress
        '
        Me.laProgress.Location = New System.Drawing.Point(56, 88)
        Me.laProgress.Name = "laProgress"
        Me.laProgress.Size = New System.Drawing.Size(184, 16)
        Me.laProgress.TabIndex = 9
        '
        'ckRegenere
        '
        Me.ckRegenere.ForeColor = System.Drawing.Color.Red
        Me.ckRegenere.Location = New System.Drawing.Point(8, 48)
        Me.ckRegenere.Name = "ckRegenere"
        Me.ckRegenere.Size = New System.Drawing.Size(240, 24)
        Me.ckRegenere.TabIndex = 4
        Me.ckRegenere.Text = "Regénération des Mouvements de stocks"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(208, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 16)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Numéro de commande"
        '
        'tbNumCmdClt
        '
        Me.tbNumCmdClt.Location = New System.Drawing.Point(328, 16)
        Me.tbNumCmdClt.Name = "tbNumCmdClt"
        Me.tbNumCmdClt.Size = New System.Drawing.Size(104, 20)
        Me.tbNumCmdClt.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(448, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 16)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Numéro de BA"
        '
        'tbNumBA
        '
        Me.tbNumBA.Location = New System.Drawing.Point(560, 16)
        Me.tbNumBA.Name = "tbNumBA"
        Me.tbNumBA.Size = New System.Drawing.Size(100, 20)
        Me.tbNumBA.TabIndex = 3
        '
        'frmControleStock
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1008, 670)
        Me.Controls.Add(Me.tbNumBA)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbNumCmdClt)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ckRegenere)
        Me.Controls.Add(Me.laProgress)
        Me.Controls.Add(Me.lbErreurs)
        Me.Controls.Add(Me.dtdatefin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtdateDeb)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.laTraitement)
        Me.Controls.Add(Me.pbProgressBar)
        Me.Controls.Add(Me.cbControle)
        Me.Name = "frmControleStock"
        Me.Text = "Controle des mouvements de stocks produits"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.cbControle, 0)
        Me.Controls.SetChildIndex(Me.pbProgressBar, 0)
        Me.Controls.SetChildIndex(Me.laTraitement, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.dtdateDeb, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.dtdatefin, 0)
        Me.Controls.SetChildIndex(Me.lbErreurs, 0)
        Me.Controls.SetChildIndex(Me.laProgress, 0)
        Me.Controls.SetChildIndex(Me.ckRegenere, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.tbNumCmdClt, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.tbNumBA, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private Enum vncCheckStock
        checkStockparDate = 1
        checkStockNumCMDCLT = 2
        checkStockNumBA = 3
    End Enum
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarSaveEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False

    End Sub

    Private Function controledesMvtsStocks(ByVal nTypeCheck As vncCheckStock, Optional ByVal pDateDeb As Date = DATE_DEFAUT, Optional ByVal pdateFin As Date = DATE_DEFAUT, Optional ByVal pnumCMDCLT As String = "", Optional ByVal pnumBA As String = "") As Boolean
        '===================================================================================
        ' Function : controledesMvtsStocks
        ' Description : Controles des Mvts de Stocks par rapport aux commandes
        '               Pour Chaque commande >= livrée
        '                   Pour Chaque ligne de commande
        '                       Vérification de l'existence d'un mvt de stocks
        '===================================================================================

        Dim colCommandes As Collection
        Dim colToutesCommandes As Collection
        Dim objCommande As Commande
        Dim colReturn As Collection
        Dim str As String
        Dim bReturn As Boolean

        colToutesCommandes = New Collection()
        setCursorWait()
        Select Case nTypeCheck

            Case vncCheckStock.checkStockparDate
                DisplayStatus("Lectures des Commandes Livrées")
                'Conacténation des commandes
                colToutesCommandes = CommandeClient.getListe(dtdateDeb.Value.ToShortDateString, dtdatefin.Value.ToShortDateString, vncEnums.vncEtatCommande.vncLivree, "")
                DisplayStatus("Lectures des Commandes Eclatées")
                colCommandes = CommandeClient.getListe(dtdateDeb.Value.ToShortDateString, dtdatefin.Value.ToShortDateString, vncEnums.vncEtatCommande.vncEclatee, "")
                'Conacténation des commandes
                DisplayStatus("Concaténation des commandes")
                For Each objCommande In colCommandes
                    colToutesCommandes.Add(objCommande)
                Next

                DisplayStatus("Lectures des BA Livrés")
                colCommandes = BonAppro.getListe(dtdateDeb.Value.ToShortDateString, dtdatefin.Value.ToShortDateString, , vncEnums.vncEtatCommande.vncBALivre)
                'Conacténation des BA
                DisplayStatus("Concaténation des bonAppros")
                For Each objCommande In colCommandes
                    colToutesCommandes.Add(objCommande)
                Next

            Case vncCheckStock.checkStockNumCMDCLT
                DisplayStatus("Lectures de la  commande")
                colToutesCommandes = CommandeClient.getListe(pnumCMDCLT, "", vncEtatCommande.vncRien, "")

            Case vncCheckStock.checkStockNumBA
                DisplayStatus("Lecture du Bon d'appro")
                colToutesCommandes = BonAppro.getListe(pnumBA)
        End Select


        lbErreurs.Items.Clear()
        pbProgressBar.Minimum = 0
        pbProgressBar.Maximum = colToutesCommandes.Count
        pbProgressBar.Value = 0

        For Each objCommande In colToutesCommandes
            bReturn = objCommande.load()
            Debug.Assert(bReturn, "Lecture de la commande en entier")
            laTraitement.Text = objCommande.toString()
            pbProgressBar.Increment(1)
            laProgress.Text = pbProgressBar.Value & "/" & colToutesCommandes.Count
            colReturn = objCommande.ControleMvtStock()
            DisplayStatus("Controle de l'élément " & objCommande.code)
            Debug.Assert(Not colReturn Is Nothing)
            For Each str In colReturn
                lbErreurs.Items.Add(str)
            Next
            If ckRegenere.Checked Then
                'Regenération des Mvts de stocks en cas d'erreur
                If colReturn.Count > 0 Then
                    'S'il y a une erreur
                    objCommande.regenereMvtStock()
                End If
            End If
            Me.Refresh()
        Next

        restoreCursor()
        DisplayStatus("")
        MsgBox("Terminé !!")
    End Function

    Private Sub cdcontrole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbControle.Click
        Dim typeCtrl As vncCheckStock

        If tbNumBA.Text <> "" Then
            dtdateDeb.Value = DATE_DEFAUT
            dtdatefin.Value = DATE_DEFAUT
            tbNumCmdClt.Text = ""
            typeCtrl = vncCheckStock.checkStockNumBA
        End If

        If tbNumCmdClt.Text <> "" Then
            dtdateDeb.Value = DATE_DEFAUT
            dtdatefin.Value = DATE_DEFAUT
            tbNumBA.Text = ""
            typeCtrl = vncCheckStock.checkStockNumCMDCLT
        End If

        If tbNumBA.Text = "" And tbNumCmdClt.Text = "" Then
            typeCtrl = vncCheckStock.checkStockparDate
        End If

        If MsgBox("Etes-vous sur de vouloir contrôler les mouvements de stocks ", MsgBoxStyle.YesNo, "Contrôle des mouvements de stocks") = MsgBoxResult.No Then
            Exit Sub
        End If
        If ckRegenere.Checked Then
            If MsgBox("Etes-vous sur de vouloir CORRIGER les mouvements de stocks ", MsgBoxStyle.YesNo, "Correction des mouvements de stocks") = MsgBoxResult.No Then
                Exit Sub
            End If
            If MsgBox("Avez-vous fait une sauvegarde ?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        controledesMvtsStocks(typeCtrl, dtdateDeb.Value.ToShortDateString(), dtdatefin.Value.ToShortDateString(), tbNumCmdClt.Text, tbNumBA.Text)
    End Sub

    Private Sub ckRegenere_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckRegenere.CheckedChanged
        If ckRegenere.Checked Then
            cbControle.Text = "Contrôle et Correction des annomalies de stocks"
        Else
            cbControle.Text = "Contrôle des mouvements stocks"
        End If
    End Sub

    Public Overrides Function getResume() As String
        Return "CTRL_STK: Contrôle des mouvements de stocks"
    End Function

End Class
