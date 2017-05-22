Imports System.Diagnostics
Imports vini_DB
Imports System.Threading

Public Class vini_service

    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        Me.ServiceName = "vini_service"
        Me.CanShutdown = True
        Me.CanStop = True
    End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Ajoutez ici le code pour démarrer votre service. Cette méthode doit
        ' démarrer votre service.
        If Not EventLog.SourceExists(Me.ServiceName) Then
            EventLog.CreateEventSource(Me.ServiceName, Me.ServiceName)
        End If
        Trace.Listeners.Clear()
        Trace.Listeners.Add(New EventLogTraceListener(Me.ServiceName))
        Trace.WriteLine("Starting...")
        Trace.WriteLine("Imap = (" & My.Settings.ImapHost & ":" & My.Settings.ImapPort & "," & My.Settings.ImapUser & "," & My.Settings.ImapPassword & "," & My.Settings.ImapSSL & ")")
        Trace.WriteLine("dossier imap des messages traités = " & My.Settings.IMAPMSGfolder)
        Trace.WriteLine("Intervalle = " & My.Settings.ImapNSec & " Secondes")
        'Arement d'un timer qui démarera dans 1 secondes
        Dim tcb As TimerCallback = AddressOf ImporterCommandes
        Timer1 = New Timer(tcb, Nothing, 1, My.Settings.ImapNSec * 1000)
        Trace.WriteLine("Started...")
    End Sub

    Protected Overrides Sub OnStop()
        ' Ajoutez ici le code pour effectuer les destructions nécessaires à l'arrêt de votre service.
        Trace.WriteLine("Stopping...")
        '...
        Trace.WriteLine("Stopped...")
    End Sub

    Private Sub ImporterCommandes()
        Dim nSec As Integer
        Try
            Dim olst As List(Of CommandeClient)
            Persist.ConnectionString = My.Settings.MyCS
            If Persist.shared_connect() Then

                Dim oImport As New ImportPrestashop(My.Settings.ImapHost, My.Settings.ImapUser, My.Settings.ImapPassword, My.Settings.ImapPort, My.Settings.ImapSSL)
                oImport.MSGFolderName = My.Settings.IMAPMSGfolder
                nSec = My.Settings.ImapNSec
                olst = oImport.Import(True)
                If olst.Count > 0 Then
                    Trace.WriteLine("" & olst.Count.ToString() & " commandes importées")
                End If
            Else
                EventLog.WriteEntry(Me.ServiceName, "Erreur en Connection à la base de données" & Persist.ConnectionString, EventLogEntryType.Error)
            End If

        Catch ex As Exception
            EventLog.WriteEntry(Me.ServiceName, ex.Message, EventLogEntryType.Error)
        End Try

    End Sub

End Class
