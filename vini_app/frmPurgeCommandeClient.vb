Imports vini_DB
Public Class frmPurgeCommandeClient
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
        Me.laTraitement.Location = New System.Drawing.Point(24, 160)
        Me.laTraitement.Name = "laTraitement"
        Me.laTraitement.Size = New System.Drawing.Size(688, 24)
        Me.laTraitement.TabIndex = 16
        '
        'pbProgressBar
        '
        Me.pbProgressBar.Location = New System.Drawing.Point(32, 192)
        Me.pbProgressBar.Name = "pbProgressBar"
        Me.pbProgressBar.Size = New System.Drawing.Size(680, 24)
        Me.pbProgressBar.TabIndex = 15
        '
        'cbAnnuler
        '
        Me.cbAnnuler.Location = New System.Drawing.Point(480, 224)
        Me.cbAnnuler.Name = "cbAnnuler"
        Me.cbAnnuler.Size = New System.Drawing.Size(104, 24)
        Me.cbAnnuler.TabIndex = 9
        Me.cbAnnuler.Text = "Annuler"
        '
        'cbValider
        '
        Me.cbValider.Location = New System.Drawing.Point(600, 224)
        Me.cbValider.Name = "cbValider"
        Me.cbValider.Size = New System.Drawing.Size(96, 24)
        Me.cbValider.TabIndex = 11
        Me.cbValider.Text = "Purger"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(40, 136)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(240, 24)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Purger les commandes clients ant�rieures � "
        '
        'dtDatePurge
        '
        Me.dtDatePurge.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtDatePurge.Location = New System.Drawing.Point(288, 136)
        Me.dtDatePurge.Name = "dtDatePurge"
        Me.dtDatePurge.Size = New System.Drawing.Size(96, 20)
        Me.dtDatePurge.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(608, 24)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Il n'y a pas de retour arri�re possible, donc assurez vous d'avoir fait une sauve" & _
        "garde auparavant."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(608, 32)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Ce programme va supprimer les commandes clients ECLATEES et FACTUREES ant�rieures" & _
        "  � la date que vous allez fournir. Ce programme ne supprime ni les mouvements d" & _
        "e stocks ni les factures de commissions."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(688, 40)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Purge des Commandes Clients"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmPurgeCommandeClient
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
        Me.Name = "frmPurgeCommandeClient"
        Me.Text = "Purge des commandes clients"
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
        Return "Purge des Commandes Clients"
    End Function
    Private Sub purge(ByVal pDatePurge As Date)
        Dim objCommande As CommandeClient
        Dim colCommande As Collection
        Dim bReturn As Boolean

        setcursorWait()
        Try
            'Chargement de toutes les Commandes Clients ant�rieures � la date fix�e dans l'�tat Rapproch�e
            colCommande = CommandeClient.getListe(DATE_DEFAUT, DateAdd(DateInterval.Day, -1, pDatePurge), , vncEnums.vncEtatCommande.vncEclatee)
            pbProgressBar.Minimum = 0
            pbProgressBar.Maximum = colCommande.Count
            pbProgressBar.Step = 1

            bReturn = True
            For Each objCommande In colCommande
                bReturn = objCommande.load()
                Debug.Assert(bReturn, "objFact.load" & Commande.getErreur())
                pbProgressBar.Increment(1)
                laTraitement.Text = "Traitement de " & objCommande.shortResume
                Me.Refresh()
                bReturn = objCommande.purge()
            Next
        Catch ex As Exception
            bReturn = False
            DisplayError("purge", Commande.getErreur())
        End Try

        restoreCursor()
        MsgBox("Purge Termin�e")


    End Sub 'purge
    Private Sub initFenetre()
        'Initialisation de la date � l'an dernier
        dtDatePurge.Value = DateAdd(DateInterval.Year, -1, Now())
    End Sub

    Private Sub frmPurgeFactCom_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initFenetre()
    End Sub
    Private Sub cbAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnnuler.Click
        Me.Hide()
    End Sub
    Private Sub cbValider_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbValider.Click
        If MsgBox("Etes-vous sur de vouloir purger les commandes clients ant�rieures au " & dtDatePurge.Value.ToShortDateString, MsgBoxStyle.OkCancel, "Purge des commandes clients") = MsgBoxResult.Ok Then
            purge(dtDatePurge.Value.ToShortDateString)
            Me.Hide()
        End If
    End Sub
End Class
