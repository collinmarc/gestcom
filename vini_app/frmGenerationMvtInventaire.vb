Imports vini_DB
Public Class frmGenerationMvtInventaire
    Inherits FrmVinicom
    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub


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
    Friend WithEvents cbGenMVI As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbCodeProduit As System.Windows.Forms.TextBox
    Friend WithEvents dtDateInventaire As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbGenMVI = New System.Windows.Forms.Button
        Me.pbProgressBar = New System.Windows.Forms.ProgressBar
        Me.laTraitement = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtDateInventaire = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbCodeProduit = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(216, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(568, 48)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ATTENTION Cette fenêtre permet de générer les mouvements de stocks d'inventaire. " & _
            "Assurrez-vous d'avoir conserver un état des stocks et d'avoir fait une sauvegard" & _
            "e."
        '
        'cbGenMVI
        '
        Me.cbGenMVI.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbGenMVI.Location = New System.Drawing.Point(8, 248)
        Me.cbGenMVI.Name = "cbGenMVI"
        Me.cbGenMVI.Size = New System.Drawing.Size(1000, 48)
        Me.cbGenMVI.TabIndex = 2
        Me.cbGenMVI.Text = "&Générer les Mouvements d'Inventaire"
        '
        'pbProgressBar
        '
        Me.pbProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbProgressBar.Location = New System.Drawing.Point(8, 368)
        Me.pbProgressBar.Name = "pbProgressBar"
        Me.pbProgressBar.Size = New System.Drawing.Size(992, 24)
        Me.pbProgressBar.TabIndex = 2
        '
        'laTraitement
        '
        Me.laTraitement.Location = New System.Drawing.Point(16, 320)
        Me.laTraitement.Name = "laTraitement"
        Me.laTraitement.Size = New System.Drawing.Size(984, 16)
        Me.laTraitement.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(495, 220)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "date d'inventaire :"
        '
        'dtDateInventaire
        '
        Me.dtDateInventaire.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtDateInventaire.Location = New System.Drawing.Point(597, 216)
        Me.dtDateInventaire.Name = "dtDateInventaire"
        Me.dtDateInventaire.Size = New System.Drawing.Size(96, 20)
        Me.dtDateInventaire.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(331, 190)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(341, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Normalement les mouvements d'inventaire sont générés au 1er du mois" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(324, 220)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Code Produit :"
        '
        'tbCodeProduit
        '
        Me.tbCodeProduit.Location = New System.Drawing.Point(405, 215)
        Me.tbCodeProduit.Name = "tbCodeProduit"
        Me.tbCodeProduit.Size = New System.Drawing.Size(84, 20)
        Me.tbCodeProduit.TabIndex = 0
        '
        'frmGenerationMvtInventaire
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 678)
        Me.Controls.Add(Me.tbCodeProduit)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtDateInventaire)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.laTraitement)
        Me.Controls.Add(Me.pbProgressBar)
        Me.Controls.Add(Me.cbGenMVI)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmGenerationMvtInventaire"
        Me.Text = "Génération des Mouvements d'Inventaire"
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

    Private Function generationMvtInventaire(ByVal pdtDateMvt As Date, ByVal pCodeProduit As String) As Boolean
        Dim objProduit As Produit
        Dim colProduit As Collection
        Dim bReturn As Boolean
        Dim nStock As Decimal

        setcursorWait()
        Try

            If String.IsNullOrEmpty(pCodeProduit) Then
                colProduit = Produit.getListe(vncTypeProduit.vncPlateforme)
            Else
                colProduit = Produit.getListe(vncTypeProduit.vncPlateforme, pCodeProduit)
            End If

            pbProgressBar.Minimum = 0
            pbProgressBar.Maximum = colProduit.Count
            pbProgressBar.Step = 1

            bReturn = True
            For Each objProduit In colProduit
                bReturn = objProduit.load()
                Debug.Assert(bReturn, "objProduit.load" & Produit.getErreur())
                bReturn = objProduit.loadcolmvtStock()
                Debug.Assert(bReturn, "objProduit.loadcolmvtStock" & Produit.getErreur())
                pbProgressBar.Increment(1)
                laTraitement.Text = "Traitement de " & objProduit.code
                Me.Refresh()
                nStock = objProduit.CalculeStockAu(pdtDateMvt)
                objProduit.genereMvtInventaire(pdtDateMvt, nStock)
                If bReturn Then
                    bReturn = objProduit.save()
                End If
                Debug.Assert(bReturn, "geneMvtInventaire" & Produit.getErreur())
            Next
        Catch ex As Exception
            bReturn = False
            DisplayError("Generation Mvt Inventaire", Produit.getErreur())
        End Try
        MsgBox("Génération des mouvements d'inventaire terminée")
        restoreCursor()
        Return bReturn
    End Function

    Private Sub cbGenMVI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbGenMVI.Click
        Dim dtDateMvt As Date
        dtDateMvt = CType(dtDateInventaire.Value.ToShortDateString(), Date)
        If MsgBox("Etes-vous sur de vouloir générer les Mouvements d'inventaire à la date du " + dtDateMvt.ToShortDateString(), MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            generationMvtInventaire(dtDateMvt, tbCodeProduit.Text)
        End If
    End Sub
End Class
