Imports System.IO
Imports vini_DB

Public Class frmImportcommandeClient
    Public Overrides Function getResume() As String
        Return "Import des commandes clients depuis les fichier transmis par le site internet"
    End Function

    Protected Overrides Sub EnableControls(ByVal bEnabled As Boolean)
        MyBase.EnableControls(True)
    End Sub

    Private Sub cbParcourir_Click(sender As System.Object, e As System.EventArgs) Handles cbParcourir.Click
        If m_rbRepertoire.Checked Then
            FolderBrowserDialog1.ShowDialog()
            tbPath.Text = FolderBrowserDialog1.SelectedPath
        End If
        If m_rbFichier.Checked Then
            OpenFileDialog1.ShowDialog()
            tbPath.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub cbImporter_Click(sender As System.Object, e As System.EventArgs) Handles cbImporter.Click
        If m_rbFichier.Checked Then
            importerFichier(tbPath.Text)
        End If
        If m_rbRepertoire.Checked Then
            importerRepertoire(tbPath.Text)
        End If
    End Sub
    Private Sub importerFichier(pFileName As String)
        Dim oCmd As CommandeClient
        Dim bReturn As Boolean

        If My.Computer.FileSystem.FileExists(pFileName) Then
            oCmd = CommandeClient.importCommandeClient(pFileName)
            If Not oCmd Is Nothing Then
                bReturn = oCmd.save()
                If bReturn Then
                    displayMessage("Commande N°" + oCmd.code + " crée pour le client " + oCmd.oTiers.nom + "pour un montant de " + Format(oCmd.totalHT, "C") + "H.T.")
                Else
                    displayMessage("Erreur en sauvegarde de la commande")
                End If
            End If
        Else
            MsgBox("Fichier inexistant")
        End If
    End Sub
    Private Sub importerRepertoire(pFolderName As String)
        If My.Computer.FileSystem.DirectoryExists(pFolderName) Then
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(pFolderName, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*")
                importerFichier(foundFile)
            Next
        Else
            MsgBox("Répertoire inexistant")
        End If

    End Sub
    Private Sub displayMessage(ByVal strMessage As String)
        lbMessages.Items.Add(DateAndTime.Now.ToLongTimeString() + ":" + strMessage)
        lbMessages.SetSelected(lbMessages.Items.Count - 1, True)
        lbMessages.Refresh()
    End Sub

End Class

