Imports System.IO
Imports vini_DB
Imports System.Net.Mail

Public Class frmImportcommandeClientPrestashop

    Private m_oImportPrestashop As ImportPrestashop
    Public Overrides Function getResume() As String
        Return "Import des commandes clients crées sur le site internet"
    End Function

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub


    Private Sub cbImporter_Click(sender As System.Object, e As System.EventArgs)
        importCmd()
    End Sub

    Private Sub frmImportcommandeClientPrestashop_Load(sender As Object, e As EventArgs) Handles Me.Load
        formLoad()
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs)
        TestImap()
    End Sub
    Private Sub cbRecharger_Click(sender As Object, e As EventArgs)
        ChargerFichierLog()
    End Sub

    Private Sub importCmd()
        Me.Cursor = Cursors.WaitCursor
        Dim olst As List(Of CommandeClient)
        olst = m_oImportPrestashop.Import(True)
        Me.Cursor = Cursors.Default

        MessageBox.Show(olst.Count & " Commandes importées")
        ChargerFichierLog()

    End Sub
    Private Sub TestImap()
        If m_oImportPrestashop.TestLogin() Then
            Dim mail As MailMessage = New MailMessage(My.Settings.SMTPuser, m_oImportPrestashop.UserName)
            mail.Subject = "TEST"
            mail.IsBodyHtml = False
            '        mail.BodyEncoding = System.Text.ASCIIEncoding.UTF8
            Dim strBody As String = ""
            strBody = ""
            strBody = strBody & "[?xml version = ""1.0"" encoding=""utf-8"" standalone=""yes"" ?]" & vbCrLf
            strBody = strBody & "[cmdprestashop]" & vbCrLf
            strBody = strBody & "[id]36[/id]" & vbCrLf
            strBody = strBody & "[name]ESZARIWUG[/name]" & vbCrLf
            strBody = strBody & "[customer_id][/customer_id]" & vbCrLf
            strBody = strBody & "[livraison_company]MCII[/livraison_company]" & vbCrLf
            strBody = strBody & "[livraison_name]MCII[/livraison_name]" & vbCrLf
            strBody = strBody & "[livraison_firstname]MCII[/livraison_firstname]" & vbCrLf
            strBody = strBody & "[livraison_adress1]23, la mettrie[/livraison_adress1]" & vbCrLf
            strBody = strBody & "[livraison_adress2][/livraison_adress2]" & vbCrLf
            strBody = strBody & "[livraison_postalcode]35250[/livraison_postalcode]" & vbCrLf
            strBody = strBody & "[livraison_city]Chasné sur illet[/livraison_city]" & vbCrLf
            strBody = strBody & "[lignes]" & vbCrLf
            strBody = strBody & "[ligneprestashop]" & vbCrLf
            strBody = strBody & "[reference]demo_zzzz[/reference]" & vbCrLf
            strBody = strBody & "[quantite]1[/quantite]" & vbCrLf
            strBody = strBody & "[prixunitaire]5.5[/prixunitaire]" & vbCrLf
            strBody = strBody & "[/ligneprestashop]" & vbCrLf
            strBody = strBody & "[ligneprestashop]" & vbCrLf
            strBody = strBody & "[reference]demo_3[/reference]" & vbCrLf
            strBody = strBody & "[quantite]1[/quantite]" & vbCrLf
            strBody = strBody & "[prixunitaire]5.5[/prixunitaire]" & vbCrLf
            strBody = strBody & "[/ligneprestashop]" & vbCrLf
            strBody = strBody & "[/lignes]" & vbCrLf
            strBody = strBody & "[/cmdprestashop]" & vbCrLf
            strBody = strBody & "[/xml]" & vbCrLf

            Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(strBody, System.Text.Encoding.GetEncoding("UTF-8"), "text/plain")
            plainView.TransferEncoding = Net.Mime.TransferEncoding.SevenBit
            mail.AlternateViews.Add(plainView)
            mail.BodyEncoding = System.Text.Encoding.GetEncoding("UTF-8")

            Dim smtp As SmtpClient = New SmtpClient()
            smtp.Host = My.Settings.SMTPHost
            smtp.Port = My.Settings.SMTPPort
            smtp.Credentials = New System.Net.NetworkCredential(My.Settings.SMTPuser, My.Settings.SMTPpassword)
            smtp.EnableSsl = My.Settings.SMTPbSSL

            Try
                smtp.Send(mail)
            Catch ex As Exception
                Trace.WriteLine("frmImportCommandPrestashop.TestImap ERR", ex.Message)
            End Try


        Else
            MessageBox.Show("Connexion impossible", "Connexion IMAP", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub


    Private Sub ChargerFichierLog()
        Timer1.Enabled = False
        'Chargement du fichier de Log
        Me.Cursor = Cursors.WaitCursor
        lbMessages.Items.Clear()
        Dim oEntry As EventLogEntry
        Dim n As Integer
        Dim oEventLoag As EventLog = New EventLog("Application", System.Environment.MachineName)
        For n = (oEventLoag.Entries.Count - 1) To 0 Step -1
            oEntry = oEventLoag.Entries(n)
            If oEntry.Source = "vini_service" Then
                lbMessages.Items.Add(oEntry.TimeWritten & " " & oEntry.Message)
            End If
        Next
        Me.Cursor = Cursors.Default

        'If System.IO.File.Exists(ImportPrestashop.IMPORTLOGFILE) Then
        '    Dim st As StreamReader = New StreamReader(ImportPrestashop.IMPORTLOGFILE)
        '    Do While st.Peek() >= 0
        '        lbMessages.Items.Add(st.ReadLine())
        '    Loop
        '    st.Close()
        'End If
        Timer1.Enabled = True

    End Sub
    Private Sub formLoad()
        ChargerFichierLog()
        Timer1.Interval = 30000
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ChargerFichierLog()
    End Sub
End Class

