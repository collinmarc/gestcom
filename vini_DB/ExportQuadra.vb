Imports System.Collections.Generic

Public Class ExportQuadra
    Inherits Observable
    Private _dateDeb As Date
    Public Property dateDeb() As Date
        Get
            Return _dateDeb
        End Get
        Set(ByVal value As Date)
            _dateDeb = value
        End Set
    End Property
    Private _datefin As Date
    Public Property dateFin() As Date
        Get
            Return _datefin
        End Get
        Set(ByVal value As Date)
            _datefin = value
        End Set
    End Property
    Private _TypeExport As vncTypeExportQuadra
    Public Property typeExport() As vncTypeExportQuadra
        Get
            Return _TypeExport
        End Get
        Set(ByVal value As vncTypeExportQuadra)
            _TypeExport = value
        End Set
    End Property

    Private _Folder As String
    Public Property folder() As String
        Get
            Return _Folder
        End Get
        Set(ByVal value As String)
            _Folder = value
        End Set
    End Property

    Private _bSaveCmd As Boolean
    Public Property bSaveCmd() As Boolean
        Get
            Return _bSaveCmd
        End Get
        Set(ByVal value As Boolean)
            _bSaveCmd = value
        End Set
    End Property

    Private _lst As List(Of Commande)
    Public Property ListCmd() As List(Of Commande)
        Get
            Return _lst
        End Get
        Private Set(ByVal value As List(Of Commande))
            _lst = value
        End Set
    End Property

    Public Sub New(pDateDeb As Date, pdateFin As Date, ptype As vncTypeExportQuadra, pFolder As String, Optional pbSaveCmd As Boolean = True)

        dateDeb = pDateDeb
        dateFin = pdateFin
        typeExport = ptype
        folder = pFolder
        bSaveCmd = pbSaveCmd

        ListCmd = New List(Of Commande)

    End Sub

    '            strFolder = tbExportQuadraFolder.Text
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return "ExportQuadra"
        End Get
    End Property
    ''' <summary>
    ''' Export des Commandes dans un fichier CSV pour Quadra
    ''' </summary>
    Public Function ExportBaf() As Boolean


        Debug.Assert(Not ListCmd Is Nothing)

        Dim objSCMD As Commande
        Dim bExportOK As Boolean = True
        Dim bReturn As Boolean
        Dim strFile As String


        Try
            Select Case typeExport
                Case vncTypeExportQuadra.vncExportBafClient
                    strFile = folder & "/ExportBAF" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & ".csv"
                Case vncTypeExportQuadra.vncExportBaFournisseur
                    strFile = folder & "/ExportAchat" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & ".csv"
                Case Else
                    strFile = folder & "/ExportBAF" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & ".csv"
            End Select

            bReturn = False
            If Not My.Computer.FileSystem.DirectoryExists(strFile) Then
                My.Computer.FileSystem.CreateDirectory(strFile)
            End If

            'Génération du nom de fichier en fonction du type d'export

            Call Notifier()

            'Parcours de la liste des souscommandes et génération du fichier CSV
            For Each objSCMD In ListCmd
                objSCMD.load()
                objSCMD.loadcolLignes()
                If objSCMD.colLignes.Count > 0 Then
                    objSCMD.toCSVQuadraFact(strFile, typeExport)
                    objSCMD.ValiderExportQuadra()
                    'Il faut sauvegarder les sous commandes car l'export a été réalisé
                    If Me.bSaveCmd Then
                        objSCMD.Save()
                    End If
                End If
                Notifier()
            Next
            bReturn = True
        Catch ex As Exception
            setError(Environment.StackTrace, ex.Message)
            bReturn = False
        End Try
        Return bReturn

    End Function

    ''' <summary>
    ''' Rend la liste des commandes à exporter
    ''' </summary>
    ''' <returns></returns>
    Public Function loadListCmd() As Boolean

        Dim bReturn As Boolean = False

        Try
            Me.ListCmd.Clear()

            Select Case typeExport
                Case vncTypeExportQuadra.vncExportBafClient
                    'Liste des SOUSCOMMANDES "VINICOM" avec des produits "HOBIVIN" (Fournisseur.Export = 2)
                    Me.ListCmd.AddRange(SousCommande.getListeAExporter(vncTypeExportScmd.vncExportQuadra, vncOrigineCmd.vncVinicom, dateDeb, dateFin))
                    'Liste des COMMMANDES "HOBIVIN" qqsoit le type D'export
                    Dim oCol As Collection
                    Dim strOrigine As String
                    strOrigine = "HOBIVIN"
                    oCol = CommandeClient.getListe(dateDeb, dateFin, "", vncEtatCommande.vncEclatee, strOrigine)
                    For Each oCmd As CommandeClient In oCol
                        Me.ListCmd.Add(oCmd)
                    Next
                Case vncTypeExportQuadra.vncExportBaFournisseur
                    'Liste des SousCommandes"HOBIVIN" avec des produits VINICOM (Fournisseur.Export <>2)
                    Me.ListCmd.AddRange(SousCommande.getListeAExporter(vncTypeExportScmd.vncExportInternet, vncOrigineCmd.vncHOBIVIN, dateDeb, dateFin))
                    Me.ListCmd.AddRange(SousCommande.getListeAExporter(vncTypeExportScmd.vncPasExport, vncOrigineCmd.vncHOBIVIN, dateDeb, dateFin))
            End Select
            bReturn = True
        Catch ex As Exception
            setError(Environment.StackTrace, ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function


End Class
