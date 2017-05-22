Imports vini_DB
Public Class frmPurgeBonAppro
    Inherits vini_app.frmStatistiques

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
    Friend WithEvents laTraitement As System.Windows.Forms.Label
    Friend WithEvents pbProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents cbAnnuler As System.Windows.Forms.Button
    Friend WithEvents cbValider As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtDatePurge As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.laTraitement = New System.Windows.Forms.Label
        Me.pbProgressBar = New System.Windows.Forms.ProgressBar
        Me.cbAnnuler = New System.Windows.Forms.Button
        Me.cbValider = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtDatePurge = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'laTraitement
        '
        Me.laTraitement.Location = New System.Drawing.Point(16, 160)
        Me.laTraitement.Name = "laTraitement"
        Me.laTraitement.Size = New System.Drawing.Size(688, 24)
        Me.laTraitement.TabIndex = 25
        '
        'pbProgressBar
        '
        Me.pbProgressBar.Location = New System.Drawing.Point(24, 192)
        Me.pbProgressBar.Name = "pbProgressBar"
        Me.pbProgressBar.Size = New System.Drawing.Size(680, 24)
        Me.pbProgressBar.TabIndex = 24
        '
        'cbAnnuler
        '
        Me.cbAnnuler.Location = New System.Drawing.Point(472, 224)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.Size = New System.Drawing.Size(104, 24)
        Me.cbAnnuler.TabIndex = 18
        Me.cbAnnuler.Text = "Annuler"
        '
        'cbValider
        '
        Me.cbValider.Location = New System.Drawing.Point(592, 224)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.Size = New System.Drawing.Size(96, 24)
        Me.cbValider.TabIndex = 20
        Me.cbValider.Text = "Purger"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(32, 136)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(272, 24)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Purger les bons d'approvissionnement antérieurs à "
        '
        'dtDatePurge
        '
        Me.dtDatePurge.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtDatePurge.Location = New System.Drawing.Point(312, 136)
        Me.dtDatePurge.Name = "dtDatePurge"
        Me.dtDatePurge.Size = New System.Drawing.Size(96, 20)
        Me.dtDatePurge.TabIndex = 17
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(608, 24)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Il n'y a pas de retour arrière possible, donc assurez vous d'avoir fait une sauve" & _
        "garde auparavant."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(608, 32)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Ce programme va supprimer les Bons d'Approvisionnement LIVRES antérieures à la da" & _
        "te que vous allez fournir. Ce programme ne supprime pas  les mouvements de stock" & _
        "s."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(0, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(688, 40)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Purge des Bons d'Approvissionnement"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmPurgeBonAppro
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(712, 342)
        Me.Controls.Add(Me.laTraitement)
        Me.Controls.Add(Me.pbProgressBar)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.cbValider)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dtDatePurge)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmPurgeBonAppro"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.dtDatePurge, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.cbValider, 0)
        Me.Controls.SetChildIndex(Me.cbAnnuler, 0)
        Me.Controls.SetChildIndex(Me.pbProgressBar, 0)
        Me.Controls.SetChildIndex(Me.laTraitement, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Overrides Function getResume() As String
        Return "Purge des Bons d'approvisionnement"
    End Function

    Private Sub purge(ByVal pDatePurge As Date)
        Dim objBA As BonAppro
        Dim colBA As Collection
        Dim bReturn As Boolean

        setcursorWait()
        Try
            'Chargement de tous les produits plateforme
            colBA = BonAppro.getListe(DATE_DEFAUT, DateAdd(DateInterval.Day, -1, pDatePurge), , vncEnums.vncEtatCommande.vncBALivre)
            pbProgressBar.Minimum = 0
            pbProgressBar.Maximum = colBA.Count
            pbProgressBar.Step = 1

            bReturn = True
            For Each objBA In colBA
                bReturn = objBA.load()
                Debug.Assert(bReturn, "objProduit.load" & BonAppro.getErreur())
                pbProgressBar.Increment(1)
                laTraitement.Text = "Traitement de " & objBA.shortResume
                Me.Refresh()
                bReturn = objBA.purge()
                Debug.Assert(bReturn, "purge" & BonAppro.getErreur())
            Next
        Catch ex As Exception
            bReturn = False
            DisplayError("frmPurgepurgeBonAppro", BonAppro.getErreur())
        End Try

        restoreCursor()
        MsgBox("Purge Terminée")


    End Sub 'purge
    Private Sub initFenetre()
        'Initialisation de la date à l'an dernier
        dtDatePurge.Value = DateAdd(DateInterval.Year, -1, Now())
    End Sub


    Private Sub cbAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnuler.Click
        Me.Hide()
    End Sub


    Private Sub cbValider_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbValider.Click
        If MsgBox("Etes-vous sur de vouloir purger les bons d'approvisionnement antérieurs au " & dtDatePurge.Value.ToShortDateString, MsgBoxStyle.OkCancel, "Purge des Mouvements de Stocks") = MsgBoxResult.Ok Then
            purge(dtDatePurge.Value.ToShortDateString)
            Me.Hide()
        End If


    End Sub

    Private Sub frmPurgeMvtStock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub
End Class

