Imports vini_DB
Imports System.Net.Mail
Imports System.ServiceProcess

Partial Public Class frmConstantes
    Inherits vini_app.FrmVinicom

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub

    Protected Overrides Function frmSave() As Boolean
        Me.Validate()
        Me.CONSTANTESBindingSource.EndEdit()
        Me.CONSTANTESTableAdapter.Update(Me.DsVinicom.CONSTANTES)

    End Function

    Private Sub frmConstantes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO : cette ligne de code charge les données dans la table 'DsVinicom.CONSTANTES'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
        Me.CONSTANTESTableAdapter.Connection = Persist.oleDBConnection
        Me.CONSTANTESTableAdapter.Fill(Me.DsVinicom.CONSTANTES)

        Call init_vini_service()

    End Sub

    Private Sub cbTestWebEdi_Click(sender As System.Object, e As System.EventArgs) Handles cbTestWebEdi.Click

        Try
            Dim oMailClient As SmtpClient
            Dim oMailMessage As MailMessage
            Dim oAtt As Attachment
            Dim strTempDir As String

            Me.Cursor = Cursors.WaitCursor
            DisplayStatus("Connexion au serveur de messagerie : " & Me.tbWEBEDI_SMTPHOST.Text & ":" & Me.tbWEBEDI_SMTPPORT.Text)
            oMailClient = New SmtpClient(Me.tbWEBEDI_SMTPHOST.Text, Me.tbWEBEDI_SMTPPORT.Text)

            'Création du MailClient
            DisplayStatus("Creation du message de test depuis " & Me.tbWEBEDI_SMTPFROM.Text & " vers " & Me.tbWEBEDI_SMTPFROM.Text)
            oMailMessage = New MailMessage(Me.tbWEBEDI_SMTPFROM.Text, Me.tbWEBEDI_SMTPFROM.Text, "test", "Message de test")
            strTempDir = Me.tbWEBEDI_TEMP.Text + "/" + currentuser.code
            If Not My.Computer.FileSystem.DirectoryExists(strTempDir) Then
                My.Computer.FileSystem.CreateDirectory(strTempDir)
            End If

            DisplayStatus("Ajout d'une pièce attachée de test")
            My.Computer.FileSystem.WriteAllText(strTempDir + "/test.txt", "Ceci est un test", True)
            oAtt = New Attachment(strTempDir + "/test.txt")
            oMailMessage.Attachments.Add(oAtt)

            DisplayStatus("Envoi du message")
            oMailClient.Send(oMailMessage)
            oAtt.Dispose()

            My.Computer.FileSystem.DeleteFile(strTempDir + "/test.txt")
            MsgBox("Envoi d'un message de test terminé, vérifier la boite aux lettre de " & Me.tbWEBEDI_SMTPFROM.Text)
            DisplayStatus("")
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub CONSTANTESBindingSource_CurrentChanged(sender As System.Object, e As System.EventArgs) Handles CONSTANTESBindingSource.CurrentChanged

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim oftp As clsFTPVinicom
        DisplayStatus("")

        setcursorWait()
        oftp = New clsFTPVinicom(Me.FTP_HOSTNAMETextBox.Text, Me.FTP_USERNAMETextBox.Text, Me.FTP_PASSWORDTextBox.Text)

        If My.Computer.FileSystem.DirectoryExists("./TESTFTP") Then
            My.Computer.FileSystem.DeleteDirectory("./TESTFTP", FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
        My.Computer.FileSystem.CreateDirectory("./TESTFTP")

        Dim nFile As Integer
        Dim objSCMD As SousCommande
        Dim strSCMD_CSV As String
        Dim strPDFFileName As String
        nFile = FreeFile()
        FileOpen(nFile, "./TESTFTP/test.csv", OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
        objSCMD = SousCommande.createandload(Persist.GetSCMDMinID())
        DisplayStatus("Chargement de " & objSCMD.code)
        objSCMD.load()
        objSCMD.loadcolLignes()
        DisplayStatus("Chargement de " & objSCMD.code & " CSV")
        strSCMD_CSV = objSCMD.toCSV()
        Print(nFile, strSCMD_CSV)
        FileClose(nFile)
        DisplayStatus("Chargement de " & objSCMD.code & " PDF")
        strPDFFileName = "./TESTFTP/" & objSCMD.code & ".PDF"
        objSCMD.genererPDF(PATHTOREPORTS, strPDFFileName)

        'Envoi du fichier par FTP
        DisplayStatus("Envoi du fichier par FTP ")
        If oftp.uploadFromDir("./TESTFTP") Then
            DisplayStatus("Envoi du fichier OK")
            'Suppression et recréation du répertoire de test
            If My.Computer.FileSystem.DirectoryExists("./TESTFTP") Then
                My.Computer.FileSystem.DeleteDirectory("./TESTFTP", FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            My.Computer.FileSystem.CreateDirectory("./TESTFTP")
            strPDFFileName = objSCMD.code & ".PDF"

            DisplayStatus("Reception du fichier par FTP ")
            If (oftp.downloadToDir("./TESTFTP", strPDFFileName)) Then
                If My.Computer.FileSystem.FileExists("./TESTFTP/" & strPDFFileName) Then
                    DisplayStatus("Réception du fichier OK")
                Else
                    DisplayError("TESTFTP", "Le Fichier ./TESTFTP/" & strPDFFileName & " n'existe pas")
                End If
            Else
                DisplayError("TESTFTP", "Erreur en récupération de fichier")
            End If
        Else
            MsgBox("Erreur en Envoi de fichier")
        End If

        restoreCursor()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)
        Dim myService As New ServiceController("vini_service")
        Dim n As Integer

        Me.Cursor = Cursors.WaitCursor
        ' Si le service n'est pas démarré, alors on le démarre
        If Not myService.Status = ServiceControllerStatus.Running Then
            myService.Start()
        End If
        n = 0
        Do While myService.Status <> ServiceControllerStatus.Running And n < 100
            Threading.Thread.Sleep(1000) 'Attends 10 Secondes avant
            n = n + 1
            myService.Refresh()
            init_vini_service()
        Loop
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cbStopService_Click(sender As System.Object, e As System.EventArgs)
        Dim myService As New ServiceController("vini_service")

        Dim n As Integer
        ' Suspend l'exécution du service
        Try
            Me.Cursor = Cursors.WaitCursor
            ' Si le service peut être arrêté, alors on l'arrête
            If myService.CanStop Then
                myService.Stop()
                n = 0
                Do While myService.Status <> ServiceControllerStatus.Stopped And n < 100
                    Threading.Thread.Sleep(1000) 'Attends 10 Secondes avant
                    n = n + 1
                    myService.Refresh()
                    init_vini_service()
                Loop
            End If

        Catch ex As Exception
            MessageBox.Show("Impossible de suspendre le service")
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub init_vini_service()
        ' Instanciation d'un objet ServiceController
        'Dim myService As New ServiceController("vini_service")


        'cbStartService.Enabled = False
        'cbStopService.Enabled = False
        'tbCST_SRVCE_NBSEC.Enabled = False
        'tbCST_SRVCE_PATH.Enabled = False
        'tbCST_SRVCE_PATHERROR.Enabled = False

        '' Récupération de l'état actuel du service
        'If myService.Status = ServiceProcess.ServiceControllerStatus.Running Then
        '    cbStartService.Enabled = False
        '    cbStopService.Enabled = True
        'End If
        'If myService.Status = ServiceProcess.ServiceControllerStatus.Stopped Then
        '    cbStartService.Enabled = True
        '    cbStopService.Enabled = False
        '    tbCST_SRVCE_NBSEC.Enabled = True
        '    tbCST_SRVCE_PATH.Enabled = True
        '    tbCST_SRVCE_PATHERROR.Enabled = True
        'End If

    End Sub

    Private Sub tbImport_Click(sender As Object, e As EventArgs) Handles tbImport.Click
        Me.Cursor = Cursors.WaitCursor
        Dim olst As List(Of CommandeClient)
        Dim oImport As New ImportPrestashop(tbImapHost.Text, tbImapUser.Text, tbImapPwd.Text, tbImapPort.Text, ckImapSSL.Checked)
        oImport.MSGFolderName = tbImapFolder.Text
        olst = oImport.Import(True, ckCheck.Checked)
        Dim str As String = ""
        For Each oCmd As CommandeClient In olst
            If Not String.IsNullOrEmpty(str) Then
                str = str & ","
            End If
            str = str & oCmd.code & "(" & oCmd.NamePrestashop & ")"
        Next
        Me.Cursor = Cursors.Default
        MessageBox.Show(olst.Count & " Commandes importées : " & str)

    End Sub

    Private Sub tbSave_Click(sender As Object, e As EventArgs) Handles tbSave.Click
        My.Settings.Save()
        My.Settings.Reload()
    End Sub
End Class