Imports vini_DB
Public Class frmPurgeMvtStock
    Inherits vini_app.frmStatistiques

#Region " Code g�n�r� par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque apr�s l'appel InitializeComponent()

    End Sub

    'La m�thode substitu�e Dispose du formulaire pour nettoyer la liste des composants.
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

    'REMARQUE�: la proc�dure suivante est requise par le Concepteur Windows Form
    'Elle peut �tre modifi�e en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'�diteur de code.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtDatePurge As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbValider As System.Windows.Forms.Button
    Friend WithEvents cbAnnuler As System.Windows.Forms.Button
    Friend WithEvents pbProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents laTraitement As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtDatePurge = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.cbValider = New System.Windows.Forms.Button
        Me.cbAnnuler = New System.Windows.Forms.Button
        Me.pbProgressBar = New System.Windows.Forms.ProgressBar
        Me.laTraitement = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(688, 40)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Purge des Mouvements de Stocks"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(608, 32)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Ce programme va supprimer les mouvements de stocks ant�rieurs � la date que vous " & _
        "allez fournir.  Normalement ceci se fait apr�s une saisie d'inventaire compl�te." & _
        ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(608, 24)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Il n'y a pas de retour arri�re possible, donc assurez vous d'avoir fait une sauve" & _
        "garde auparavant."
        '
        'dtDatePurge
        '
        Me.dtDatePurge.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtDatePurge.Location = New System.Drawing.Point(288, 128)
        Me.dtDatePurge.Name = "dtDatePurge"
        Me.dtDatePurge.Size = New System.Drawing.Size(96, 20)
        Me.dtDatePurge.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(40, 128)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(240, 24)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Purger les mouvements de stocks ant�rieurs � "
        '
        'cbValider
        '
        Me.cbValider.Location = New System.Drawing.Point(600, 216)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.Size = New System.Drawing.Size(96, 24)
        Me.cbValider.TabIndex = 2
        Me.cbValider.Text = "Purger"
        '
        'cbAnnuler
        '
        Me.cbAnnuler.Location = New System.Drawing.Point(480, 216)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.Size = New System.Drawing.Size(104, 24)
        Me.cbAnnuler.TabIndex = 1
        Me.cbAnnuler.Text = "Annuler"
        '
        'pbProgressBar
        '
        Me.pbProgressBar.Location = New System.Drawing.Point(32, 184)
        Me.pbProgressBar.Name = "pbProgressBar"
        Me.pbProgressBar.Size = New System.Drawing.Size(680, 24)
        Me.pbProgressBar.TabIndex = 6
        '
        'laTraitement
        '
        Me.laTraitement.Location = New System.Drawing.Point(24, 152)
        Me.laTraitement.Name = "laTraitement"
        Me.laTraitement.Size = New System.Drawing.Size(688, 24)
        Me.laTraitement.TabIndex = 7
        '
        'frmPurgeMvtStock
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(720, 270)
        Me.Controls.Add(Me.laTraitement)
        Me.Controls.Add(Me.pbProgressBar)
        Me.Controls.Add(Me.cbAnnuler)
        Me.Controls.Add(Me.cbValider)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.dtDatePurge)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmPurgeMvtStock"
        Me.Text = "Purge des mouvements de stocks"
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
        Return "Purge des Mouvements de stocks"
    End Function

    Private Sub purge(ByVal pDatePurge As Date)
        Dim objProduit As Produit
        Dim colProduit As Collection
        Dim bReturn As Boolean

        setcursorWait()
        Try
            'Chargement de tous les produits plateforme
            colProduit = Produit.getListe(vncTypeProduit.vncPlateforme, )
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
                bReturn = objProduit.purgeMvtStock(pDatePurge)
                If bReturn Then
                    bReturn = objProduit.save()
                End If
                Debug.Assert(bReturn, "objProduit.reCalculStock" & Produit.getErreur())
            Next
        Catch ex As Exception
            bReturn = False
            DisplayError("RecacculdesStocks", Produit.getErreur())
        End Try

        restoreCursor()
        MsgBox("Purge Termin�e")


    End Sub 'purge
    Private Sub initFenetre()
        'Initialisation de la date � l'an dernier
        dtDatePurge.Value = DateAdd(DateInterval.Year, -1, Now())
    End Sub


    Private Sub cbAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnuler.Click
        Me.Hide()
    End Sub


    Private Sub cbValider_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbValider.Click
        If MsgBox("Etes-vous sur de vouloir purger les mouvements de stocks ant�rieurs au " & dtDatePurge.Value.ToShortDateString, MsgBoxStyle.OkCancel, "Purge des Mouvements de Stocks") = MsgBoxResult.Ok Then
            purge(dtDatePurge.Value.ToShortDateString)
            Me.Hide()
        End If


    End Sub

    Private Sub frmPurgeMvtStock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub
End Class
