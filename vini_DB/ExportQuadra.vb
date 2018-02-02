Imports System.Collections.Generic
Imports vini_DB

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

    Public Property isExportBafClient As Boolean
        Get
            Return typeExport = vncTypeExportQuadra.vncExportBafClient
        End Get
        Set(ByVal value As Boolean)
            If value Then
                typeExport = vncTypeExportQuadra.vncExportBafClient

            Else
                typeExport = vncTypeExportQuadra.vncExportBaFournisseur

            End If
        End Set
    End Property
    Public Property isExportBaFournisseur As Boolean
        Get
            Return typeExport = vncTypeExportQuadra.vncExportBaFournisseur
        End Get
        Set(ByVal value As Boolean)
            If Not value Then
                typeExport = vncTypeExportQuadra.vncExportBafClient

            Else
                typeExport = vncTypeExportQuadra.vncExportBaFournisseur

            End If
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

        Dim objCMD As Commande
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
            If Not My.Computer.FileSystem.DirectoryExists(folder) Then
                My.Computer.FileSystem.CreateDirectory(folder)
            End If

            'Génération du nom de fichier en fonction du type d'export
            If System.IO.File.Exists(strFile) Then
                System.IO.File.Delete(strFile)
            End If

            Call Notifier()

            'Parcours de la liste des souscommandes et génération du fichier CSV
            For Each objCMD In ListCmd
                If objCMD.Selected Then
                    objCMD.load()
                    objCMD.loadcolLignes()
                    If objCMD.colLignes.Count > 0 Then
                        objCMD.toCSVQuadraFact(strFile, typeExport)
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
    ''' Valider l'export 
    ''' </summary>
    ''' <returns></returns>
    Public Function ValiderExportBaf() As Boolean
        Debug.Assert(Not ListCmd Is Nothing)

        Dim objCMD As Commande
        Dim bReturn As Boolean

        Try
            Call Notifier()

            'Parcours de la liste des souscommandes et génération du fichier CSV
            For Each objCMD In ListCmd
                If objCMD.Selected Then
                    objCMD.load()
                    objCMD.loadcolLignes()
                    If objCMD.colLignes.Count > 0 Then
                        objCMD.ValiderExportQuadra()
                        If Me.bSaveCmd Then
                            objCMD.Save()
                        End If
                    End If
                End If
                Call Notifier()
            Next
            bReturn = True
        Catch ex As Exception
            setError(Environment.StackTrace, ex.Message)
            bReturn = False
        End Try
        Return bReturn

    End Function

    ''' <summary>
    ''' Charge  la liste des commandes à exporter
    ''' </summary>
    ''' <returns></returns>
    Public Function LoadListCmd() As Boolean

        Dim bReturn As Boolean = False
        Dim oLst As New List(Of SousCommande)
        Try
            Me.ListCmd.Clear()

            Select Case typeExport
                Case vncTypeExportQuadra.vncExportBafClient
                    'Liste des SOUSCOMMANDES "VINICOM" avec des produits "HOBIVIN" (Fournisseur.Export = 2)
                    oLst = SousCommande.getListeAExporterQuadra(vncTypeExportScmd.vncExportQuadra, vncOrigineCmd.vncVinicom, dateDeb, dateFin)
                    For Each oScmd As SousCommande In oLst
                        Me.ListCmd.Add(oScmd)
                    Next
                    'Liste des COMMMANDES "HOBIVIN" qqsoit le type D'export
                    Dim oCol As Collection
                    Dim strOrigine As String
                    strOrigine = Dossier.HOBIVIN
                    oCol = CommandeClient.getListe(dateDeb, dateFin, "", vncEtatCommande.vncEclatee, strOrigine)
                    For Each oCmd As CommandeClient In oCol
                        Me.ListCmd.Add(oCmd)
                    Next
                Case vncTypeExportQuadra.vncExportBaFournisseur
                    'Liste des SousCommandes"HOBIVIN" avec des produits VINICOM (Fournisseur.Export <>2)

                    'Ces Sous Commandes peuvent avoir été exportées vers le producteur (avec le client intermédiaire)
                    oLst = SousCommande.getListeAExporterQuadra(vncTypeExportScmd.vncExportInternet, vncOrigineCmd.vncHOBIVIN, dateDeb, dateFin)
                    ListCmd.AddRange(oLst)
                    oLst = SousCommande.getListeAExporterQuadra(vncTypeExportScmd.vncPasExport, vncOrigineCmd.vncHOBIVIN, dateDeb, dateFin)
                    ListCmd.AddRange(oLst)

            End Select

            'For Each ocmd As Commande In ListCmd
            'If TypeOf (ocmd) Is SousCommande Then
            'Dim oScmd As SousCommande
            'oScmd = ocmd
            'oScmd.oFournisseur.load()
            'End If
            'Next

            'Tri de la Liste par Ordre de Numéro de Commande
            'Obligatoire pour Quadra
            ListCmd.Sort(New CodeCommandeSorter())
            bReturn = True
        Catch ex As Exception
            setError(Environment.StackTrace, ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function
    ''' <summary>
    ''' Comparateur de codeCommande
    ''' </summary>
    Public Class CodeCommandeSorter
        Implements IComparer(Of Commande)

        Private Function IComparer_Compare(x As Commande, y As Commande) As Integer Implements IComparer(Of Commande).Compare
            Return x.getCodeCommande().CompareTo(y.getCodeCommande())
        End Function
    End Class
End Class
