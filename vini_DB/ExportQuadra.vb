Imports System.Collections.Generic

Public Class ExportQuadra
    Inherits Observable

    '            strFolder = tbExportQuadraFolder.Text
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return "ExpotQuadra"
        End Get
    End Property

    Public Sub ExportBaf(pColSCMD As List(Of SousCommande), pstrFolder As String, pTypeExport As vncTypeExportQuadra, Optional pbSaveScmd As Boolean = True)


        Debug.Assert(Not pColSCMD Is Nothing)

        Dim objSCMD As SousCommande
        Dim bExportOK As Boolean = True
        Dim bReturn As Boolean
        Dim strFile As String
        Dim ofrn As Fournisseur

        Try
            bReturn = False
            If Not My.Computer.FileSystem.DirectoryExists(pstrFolder) Then
                My.Computer.FileSystem.CreateDirectory(pstrFolder)
            End If
            strFile = pstrFolder & "/Export" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & ".csv"

            Call Notifier()

            'Parcours de la liste des souscommandes et génération du fichier CSV
            For Each objSCMD In pColSCMD
                objSCMD.load()
                ''Normalement inutile car , seul les SousCommande exportables sont chargées
                'ofrn = Fournisseur.createandload(objSCMD.Fournisseurid)
                'If ofrn.bExportInternet = 2 Then
                '    objSCMD.loadcolLignes()
                '    If objSCMD.colLignes.Count > 0 Then
                '        objSCMD.toCSVQuadra(strFile)
                '        objSCMD.ValiderExportQuadra()
                '        'Il faut sauvegarder les sous commandes car l'export a été réalisé
                '        If pbSaveScmd Then
                '            objSCMD.Save()
                '        End If
                '    End If
                'End If
                Notifier()

            Next

        Catch ex As Exception

        End Try


    End Sub



End Class
