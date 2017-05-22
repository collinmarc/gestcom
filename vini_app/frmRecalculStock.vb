Imports vini_DB
Public Class frmRecalculStock
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cdRecalcul As System.Windows.Forms.Button
    Friend WithEvents pbProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents laTraitement As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRecalculStock))
        Me.Label1 = New System.Windows.Forms.Label
        Me.cdRecalcul = New System.Windows.Forms.Button
        Me.pbProgressBar = New System.Windows.Forms.ProgressBar
        Me.laTraitement = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(568, 48)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = resources.GetString("Label1.Text")
        '
        'cdRecalcul
        '
        Me.cdRecalcul.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cdRecalcul.Location = New System.Drawing.Point(8, 72)
        Me.cdRecalcul.Name = "cdRecalcul"
        Me.cdRecalcul.Size = New System.Drawing.Size(568, 32)
        Me.cdRecalcul.TabIndex = 1
        Me.cdRecalcul.Text = "&Recalculer tous les stocks"
        '
        'pbProgressBar
        '
        Me.pbProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbProgressBar.Location = New System.Drawing.Point(8, 136)
        Me.pbProgressBar.Name = "pbProgressBar"
        Me.pbProgressBar.Size = New System.Drawing.Size(576, 24)
        Me.pbProgressBar.TabIndex = 2
        '
        'laTraitement
        '
        Me.laTraitement.Location = New System.Drawing.Point(8, 120)
        Me.laTraitement.Name = "laTraitement"
        Me.laTraitement.Size = New System.Drawing.Size(576, 16)
        Me.laTraitement.TabIndex = 3
        '
        'frmRecalculStock
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(615, 286)
        Me.Controls.Add(Me.laTraitement)
        Me.Controls.Add(Me.pbProgressBar)
        Me.Controls.Add(Me.cdRecalcul)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmRecalculStock"
        Me.Text = "Recalcul des stocks produits"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.cdRecalcul, 0)
        Me.Controls.SetChildIndex(Me.pbProgressBar, 0)
        Me.Controls.SetChildIndex(Me.laTraitement, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Protected Overrides Sub setToolbarButtons()
        m_ToolBarNewEnabled = False
        m_ToolBarLoadEnabled = False
        m_ToolBarSaveEnabled = False
        m_ToolBarDelEnabled = False
        m_ToolBarRefreshEnabled = False

    End Sub

    Private Function recalculdesStocks() As Boolean
        Dim objProduit As Produit
        Dim colProduit As Collection
        Dim bReturn As Boolean

        setcursorWait()
        Try

            colProduit = Produit.getListe(vncTypeProduit.vncTous, )

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
                bReturn = objProduit.recalculStock()
                If bReturn Then
                    bReturn = objProduit.save()
                End If
                Debug.Assert(bReturn, "objProduit.reCalculStock" & Produit.getErreur())
            Next
        Catch ex As Exception
            bReturn = False
            DisplayError("RecaculdesStocks", Produit.getErreur())
        End Try
        MsgBox("Recalcul des stocks terminé")
        restoreCursor()
        Return bReturn
    End Function

    Private Sub cdRecalcul_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cdRecalcul.Click
        If MsgBox("Etes-vous sur de vouloir recalculer tous les stocks", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            recalculdesStocks()
        End If
    End Sub
End Class
