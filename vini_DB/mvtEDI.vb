Imports System.Collections.Generic
Public Class mvtEDI
    Inherits racine
    Implements IComparable(Of mvtEDI)
    Private _NumCMD As String
    Public Sub New()
        NumCMD = ""
        Entree = 0
        Sortie = 0
    End Sub
    Public Sub New(pNumCmd As String)
        NumCMD = pNumCmd
        Entree = 0
        Sortie = 0
    End Sub

    Public Sub New(pNumCmd As String, pEntree As String, pSortie As String)
        NumCMD = pNumCmd
        Entree = pEntree
        Sortie = pSortie
    End Sub


    Public Property NumCMD() As String
        Get
            Return _NumCMD
        End Get
        Set(ByVal value As String)
            If IsNumeric(value) Then
                Dim i As Integer
                i = CInt(value)
                _NumCMD = i.ToString()
            End If

        End Set
    End Property
    Private _Entree As String
    Public Property Entree() As String
        Get
            Return _Entree
        End Get
        Set(ByVal value As String)
            _Entree = value
        End Set
    End Property

    Private _sortie As String
    Public Property Sortie() As String
        Get
            Return _sortie
        End Get
        Set(ByVal value As String)
            _sortie = value
        End Set
    End Property

    Public Function CompareTo(other As mvtEDI) As Integer Implements IComparable(Of vini_DB.mvtEDI).CompareTo
        Return String.Compare(Me.NumCMD, other.NumCMD)
    End Function

    Friend Shared Function getListFromFile(pFilePath As String, Optional pbRenameFile As Boolean = True) As List(Of mvtEDI)
        Dim lstMvt As New List(Of mvtEDI)
        Dim omvtEDI As mvtEDI
        Dim bPremierLigne As Boolean = True
        Try
            If System.IO.File.Exists(pFilePath) Then
                Dim tabLines As String() = System.IO.File.ReadAllLines(pFilePath)
                'Traietement de chaque Ligne
                For Each oLine As String In tabLines
                    If bPremierLigne Then
                        bPremierLigne = False
                    Else
                        Try

                            Dim tabFields As String() = oLine.Split(";")
                            omvtEDI = New mvtEDI(tabFields(6), tabFields(4), tabFields(5))
                            'Vérification des entrées sorties
                            If omvtEDI.Sortie <> 0 Then
                                lstMvt.Add(omvtEDI)
                            End If

                        Catch ex As Exception
                            'Erreur dans le fichier d'origine
                            'on ne fait rien
                        End Try
                    End If
                Next
                If pbRenameFile Then
                    'Renommer le fichier (ajout extension .TRT)
                    System.IO.File.Copy(pFilePath, pFilePath & ".TRT", True)
                    System.IO.File.Delete(pFilePath)
                End If
            End If
        Catch ex As Exception
            lstMvt = New List(Of mvtEDI)()
        End Try
        Return lstMvt
    End Function

    Friend Shared Function getListCumuls(pListmvt As List(Of mvtEDI)) As List(Of mvtEDI)

        Dim bReturn As Boolean
        Dim lstcumuls As New List(Of mvtEDI)
        Dim omvtCumul As mvtEDI = Nothing
        Try
            'Tri de la liste par numero de commande
            pListmvt.Sort()
            Dim entreeMVT As Integer = 0
            Dim sortieMVT As Integer = 0
            For Each oMvt As mvtEDI In pListmvt
                Try
                    entreeMVT = CInt(oMvt.Entree)

                Catch ex As Exception
                    entreeMVT = 0
                End Try
                Try
                    sortieMVT = CInt(oMvt.Sortie)

                Catch ex As Exception
                    sortieMVT = 0
                End Try

                If omvtCumul Is Nothing Then
                    omvtCumul = New mvtEDI()
                    omvtCumul.NumCMD = oMvt.NumCMD
                    omvtCumul.Entree = 0
                    omvtCumul.Sortie = 0
                End If
                If oMvt.NumCMD = omvtCumul.NumCMD Then
                    'C'est la même commande , donc je cumule
                    omvtCumul.Entree = omvtCumul.Entree + entreeMVT
                    omvtCumul.Sortie = omvtCumul.Sortie + sortieMVT
                Else
                    'Ajout des cumuls à la liste
                    lstcumuls.Add(omvtCumul)
                    'Recréatino d'un nouvel mvtCumul
                    omvtCumul = New mvtEDI()
                    omvtCumul.NumCMD = oMvt.NumCMD
                    omvtCumul.Entree = entreeMVT
                    omvtCumul.Sortie = sortieMVT
                End If
            Next
            'Ajout des cumuls à la liste
            lstcumuls.Add(omvtCumul)



            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return lstcumuls
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <param name="pbRenameFile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getFilesFromFTP(pSRV As String, pPort As String, pUser As String, pPassword As String, pRepDistant As String, pRepLocal As String) As Boolean
        '        Récupération du fichier FTP
        Dim oftp As clsFTPVinicom
        oftp = New clsFTPVinicom(pSRV, pUser, pPassword, pRepDistant)


        If Not System.IO.Directory.Exists(pRepLocal) Then
            System.IO.Directory.CreateDirectory(pRepLocal)
        End If
        oftp.downloadDirToDir(pRepLocal)
        ' Suppression des fichier .CSV sur le répertoire Distant
        Dim tabFiles As String() = System.IO.Directory.GetFiles(pRepLocal, "*.csv")
        For Each strFile As String In tabFiles
            Dim oFileInfo As New System.IO.FileInfo(strFile)
            oftp.deleteRemotefile(pRepDistant & "/" & oFileInfo.Name)
        Next

    End Function
    Public Shared Function getFilesCount(pSRV As String, pPort As String, pUser As String, pPassword As String, pRepDistant As String, pRepLocal As String) As Integer
        Dim nReturn As Integer = 0
        Dim oftp As clsFTPVinicom
        oftp = New clsFTPVinicom(pSRV, pUser, pPassword, pRepDistant)
        Try
            nReturn = oftp.getRemoteFileCount()

        Catch ex As Exception
            nReturn = 0
            setError("MvtEDI.getFilesCount ERR " & ex.Message)
        End Try

        Return nReturn

    End Function

    Public Shared Function VerificationCommandes(pFileName As String) As List(Of MsgLivraison)
        Dim bReturn As Boolean
        Dim plstCumuls As List(Of mvtEDI)
        Dim plstMsg As New List(Of MsgLivraison)

        Try
            plstCumuls = getListFromFile(pFileName)
            plstCumuls = getListCumuls(plstCumuls)
            Dim ocmd As CommandeClient
            Dim oCol As Collection
            For Each omvt As mvtEDI In plstCumuls
                Dim omsg As New MsgLivraison()
                omsg.NumCommande = omvt.NumCMD
                omsg.NbreColisLivre = omvt.Sortie
                omsg.DateMessage = DateTime.Now
                'Chargement de la commande 'Tous dossiers Confondus'
                oCol = CommandeClient.getListe(omvt.NumCMD, pOrigine:="")
                If oCol.Count > 0 Then
                    ocmd = oCol(1)
                    ocmd.load()
                    If ocmd.etat.codeEtat = vncEtatCommande.vncValidee Then
                        If ocmd.qteColis = omvt.Sortie Then
                            ocmd.LivrerToutOK()
                            ocmd.save()
                            omsg.Message = "Commande correctement traitée"
                            omsg.Resultat = 0
                        Else
                            omsg.Message = "Commande non Livrée : " & ocmd.qteColis & " colis attendus, " & omvt.Sortie & "colis livrés"
                            omsg.Resultat = 1

                        End If
                    Else
                        omsg.Message = "Commande non Livrée : Commande non validée"
                        omsg.Resultat = 1
                    End If
                Else
                    omsg.Message = "Commande non traitée : Commande inconnue"
                    omsg.Resultat = 2
                End If
                plstMsg.Add(omsg)
            Next
        Catch ex As Exception
            Dim omsg As New MsgLivraison()
            omsg.Message = ex.Message
            omsg.Resultat = 2
            setError("mvtEDI.VerificationCommandes ERR" & ex.Message)
            bReturn = False
        End Try
        Return plstMsg
    End Function

    Public Overrides Function toString() As String
        Return NumCMD & " ;" & Entree & ";" & Sortie
    End Function
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return ""
        End Get
    End Property

End Class
