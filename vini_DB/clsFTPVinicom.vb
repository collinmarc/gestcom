Imports System.IO
Imports System.Collections.Generic
'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : <<NomClasse >>
' Description : <<Commentaire
'===================================================================================================================================
'Membres de Classes
'==================
'Public
'Protected
'Private
'Membres d'instances
'==================
'Public
'   toString()      : Rend 'objet sous forme de chaine
'   Equals()        : Test l'équivalence d'instance
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================
Public Class clsFTPVinicom
    Inherits racine
    Private m_FTP As NETFTPclient
    Private m_RemoteDir As String
    Private m_strLockFromFileName As String
    Private m_strLockToFileName As String
    Private m_strErrorDescription As String
    Private m_nWaitSeconds As Integer
    Private m_nWaitLoops As Integer
    Public Sub setWaitParam(nWaitSeconds As Integer, nWaitLoop As Integer)
        m_nWaitSeconds = nWaitSeconds
        m_nWaitLoops = nWaitLoop
    End Sub
    Public ReadOnly Property ftpConnection() As NETFTPclient
        Get
            Return m_FTP
        End Get
    End Property
    Public Property remoteDir() As String
        Get
            Return m_RemoteDir
        End Get
        Set(ByVal Value As String)
            m_RemoteDir = Value
        End Set
    End Property
    Public Property strLockFromFileName() As String
        Get
            Return m_strLockFromFileName
        End Get
        Set(ByVal Value As String)
            m_strLockFromFileName = Value
        End Set
    End Property
    Public Property strLockToFileName() As String
        Get
            Return m_strLockToFileName
        End Get
        Set(ByVal Value As String)
            m_strLockToFileName = Value
        End Set
    End Property
    Public ReadOnly Property ErrorDescription() As String
        Get
            Return m_strErrorDescription
        End Get
    End Property
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "clsFTPVinicom =(" & ")"
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Return True
    End Function

    Public Sub New(ByVal strHost As String, ByVal strUSer As String, ByVal strPassword As String, Optional ByVal strRemoteDir As String = "")
        m_FTP = New NETFTPclient(strHost, strUSer, strPassword)
        m_RemoteDir = strRemoteDir
        m_strLockFromFileName = Param.getConstante("FTP_LOCKFROMFILENAME")
        m_strLockToFileName = Param.getConstante("FTP_LOCKTOFILENAME")
        m_strErrorDescription = String.Empty
        m_nWaitSeconds = 2
        m_nWaitLoops = 100
    End Sub

    Public Sub New()
        m_FTP = New NETFTPclient(Param.getConstante("FTP_HOSTNAME"), Param.getConstante("FTP_USERNAME"), Param.getConstante("FTP_PASSWORD"))
        m_RemoteDir = Param.getConstante("FTP_REMOTEDIR")
        m_strLockFromFileName = Param.getConstante("FTP_LOCKFROMFILENAME")
        m_strLockToFileName = Param.getConstante("FTP_LOCKTOFILENAME")
        m_strErrorDescription = String.Empty
        m_nWaitSeconds = 2
        m_nWaitLoops = 100
    End Sub
    '=======================================================================
    'Fonction : connect()
    'Description : se connecte au site FTP avec changenement de repertoire s'il le faut
    'Détails    :  
    'Retour : Vari si la connection est OK
    '=======================================================================
    'Public Function connect() As Boolean
    '    Dim bReturn As Boolean
    '    bReturn = False
    '    m_strErrorDescription = String.Empty
    '    If m_FTP.Download() Then
    '        If Not m_RemoteDir.Equals("") Then
    '            m_FTP.bSetCurrentDir(m_RemoteDir)
    '        End If
    '        bReturn = True
    '    Else
    '        bReturn = False
    '    End If
    '    Return bReturn
    'End Function ' connect
    '=======================================================================
    'Fonction : lockfrom()
    'Description : Vérifie la présence du fichier de Lock, renomme le fichier "semaphore" en fichier lock
    'Détails    :  
    'Retour : Vari si l'opération s'est bien déroulé est OK
    '=======================================================================
    Public Function lockFrom() As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")
        Dim bReturn As Boolean
        m_strErrorDescription = String.Empty
        bReturn = lock(m_strLockFromFileName)
        Return bReturn
    End Function ' lockFrom
    '=======================================================================
    'Fonction : lockTo()
    'Description : Vérifie la présence du fichier de Lock, renomme le fichier "semaphore" en fichier lock
    'Détails    :  
    'Retour : Vari si l'opération s'est bien déroulé est OK
    '=======================================================================
    Public Function lockTo() As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")
        Dim bReturn As Boolean
        m_strErrorDescription = String.Empty
        bReturn = lock(m_strLockToFileName)
        Return bReturn
    End Function ' lockTo
    '=======================================================================
    'Fonction : unlockfrom()
    'Description : Vérifie la présence du fichier de Lock, renomme le fichier lock en "semaphore" 
    'Détails    :  
    'Retour : Vari si l'opération s'est bien déroulé est OK
    '=======================================================================
    Public Function unlockfrom() As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")

        Dim bReturn As Boolean

        m_strErrorDescription = String.Empty
        bReturn = unlock(m_strLockFromFileName)
        Return bReturn
    End Function ' unlockFrom
    '=======================================================================
    'Fonction : unlockTo()
    'Description : Vérifie la présence du fichier de Lock, renomme le fichier lock en "semaphore" 
    'Détails    :  
    'Retour : Vari si l'opération s'est bien déroulé est OK
    '=======================================================================
    Public Function unlockTo() As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")

        Dim bReturn As Boolean

        m_strErrorDescription = String.Empty
        bReturn = unlock(m_strLockToFileName)
        Return bReturn
    End Function ' unlockTo
    '=======================================================================
    'Fonction : lock()
    'Description : Vérifie la présence du fichier de Lock, renomme le fichier "semaphore" en fichier lock
    'Détails    :  
    'Retour : Vari si l'opération s'est bien déroulé est OK
    '=======================================================================
    Private Function lock(ByVal strFileName As String) As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")
        Dim bReturn As Boolean
        Dim i As Integer
        'Dim fso As New Scripting.FileSystemObject
        Dim nFile As Integer

        bReturn = False
        For i = 0 To m_nWaitLoops
            If Not m_FTP.FtpFileExists(strFileName) Then
                bReturn = True
                Exit For
            End If
            WaitnSeconds(m_nWaitSeconds)
        Next i

        If bReturn Then
            If m_FTP.FtpFileExists("semaphore") Then
                m_FTP.FtpRename("semaphore", strFileName)
            Else
                'Creation d'un fichier temporaire
                nFile = FreeFile()
                FileOpen(nFile, "./semaphore", OpenMode.Output, OpenAccess.Write)
                WriteLine(nFile, Now())
                FileClose(nFile)
                m_FTP.Upload("./semaphore", strFileName)
            End If
            bReturn = m_FTP.FtpFileExists(strFileName)
        End If

        Return bReturn
    End Function ' lock
    '=======================================================================
    'Fonction : lock()
    'Description : Vérifie la présence du fichier de Lock, renomme le fichier "semaphore" en fichier lock
    'Détails    :  
    'Retour : Vari si l'opération s'est bien déroulé est OK
    '=======================================================================
    Private Function unlock(ByVal strFileName As String) As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")
        Dim bReturn As Boolean
        If m_FTP.FtpFileExists(strFileName) Then
            m_FTP.FtpRename(strFileName, "semaphore")
        End If

        bReturn = m_FTP.FtpFileExists("semaphore")
        Return bReturn
    End Function ' unlock
    ''' <summary>
    ''' Rend vrai sur le dichier lockTo existe sur le connecteur FTP
    ''' </summary>
    ''' <returns>Vrai/Faux</returns>
    ''' <remarks>le Connecteur doit être connnecté</remarks>
    Public Function IsLockTo() As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")

        Dim bReturn As Boolean

        m_strErrorDescription = String.Empty
        bReturn = isLock(m_strLockToFileName)
        Return bReturn
    End Function 'IsLockTo
    ''' <summary>
    ''' Rend vrai sur le dichier lockFrom existe sur le connecteur FTP
    ''' </summary>
    ''' <returns>Vrai/Faux</returns>
    ''' <remarks>le Connecteur doit être connnecté</remarks>
    Public Function IsLockFrom() As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")
        Dim bReturn As Boolean
        m_strErrorDescription = String.Empty
        bReturn = isLock(m_strLockFromFileName)
        Return bReturn
    End Function 'IsLockFrom
    ''' <summary>
    ''' Rend varai si le fichier existe sur le répertoire distant
    ''' </summary>
    ''' <param name="strFileName">nom du fichier</param>
    ''' <returns>Vrai/Faux</returns>
    ''' <remarks>Le Connecteur doit être connecté</remarks>
    Private Function isLock(ByVal strFileName As String) As Boolean
        Debug.Assert(m_FTP.isConnected, "Le connecteur FTP n'est pas connecté")
        Dim bReturn As Boolean
        bReturn = m_FTP.FtpFileExists(strFileName)
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : uploadFromDir()
    'Description : Pose le verrou , Upload tous les fichiers du répertoire, libère le verrou
    'Détails    :  
    'Retour : Vari si l'opération s'est bien déroulé est OK
    '=======================================================================
    Public Function uploadFromDir(ByVal strLocalDirName As String) As Integer
        Dim nReturn As Integer
        Dim strFileName As String
        Dim oInf As FileInfo

        Try

            nReturn = 0
            m_strErrorDescription = String.Empty
            'Pose du verrou
            If lockFrom() Then
                'Upload des fichiers
                If Directory.Exists(strLocalDirName) Then

                    For Each strFileName In Directory.GetFiles(strLocalDirName)
                        oInf = New FileInfo(strFileName)
                        Dim targetFileName As String
                        If String.IsNullOrEmpty(remoteDir) Then
                            targetFileName = oInf.Name
                        Else
                            targetFileName = remoteDir & "/" & oInf.Name

                        End If

                        If m_FTP.Upload(strFileName, targetFileName) Then
                            nReturn = nReturn + 1
                        Else
                            If Not String.IsNullOrEmpty(m_strErrorDescription) Then
                                m_strErrorDescription = m_strErrorDescription + vbCrLf
                            End If
                            m_strErrorDescription = m_strErrorDescription + "Erreur FTP"
                        End If
                    Next
                End If
                'Libération du verrou
                WaitnSeconds(1)
                unlockfrom()
            End If
        Catch ex As Exception
            Debug.Assert(False, "clsFTPVinicom.uploadFromDir" & ex.Message)
            nReturn = -1
        End Try

        Return nReturn
    End Function ' uploadFromDir
    '=======================================================================
    'Fonction : downloadToDir()
    'Description : Pose le verrou , Download tous les fichiers du répertoire, libère le verrou
    'Détails    :  
    'Retour : Vari si l'opération s'est bien déroulé est OK
    '=======================================================================
    Public Function downloadToDir(ByVal strLocalDirName As String, Optional ByVal strFileName As String = "toVinicom.csv") As Boolean
        Dim bReturn As Boolean

        Try

            bReturn = False
            m_strErrorDescription = String.Empty
            'Pose du verrou
            If lockTo() Then
                'Download des fichiers
                If m_FTP.FtpFileExists(strFileName) Then
                    m_FTP.Download(strFileName, strLocalDirName + "/" + strFileName, True)
                End If
                'Libération du verrou
                bReturn = unlockTo()
            End If
        Catch ex As Exception
            Debug.Assert(False, "clsFTPVinicom.downloadToDir" & ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function ' downloadToDir
    ''' <summary>
    ''' Download All Files from Remote Directory
    ''' </summary>
    ''' <param name="strLocalDirName"></param>
    ''' <param name="strFileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function downloadDirToDir(ByVal strLocalDirName As String) As Boolean
        Dim bReturn As Boolean

        Try

            bReturn = False
            m_strErrorDescription = String.Empty
            'Pose du verrou
            If lockTo() Then
                Dim lstFile As List(Of String) = m_FTP.ListDirectory("/" & m_RemoteDir)
                For Each strFile As String In lstFile
                    If strFile <> "." And strFile <> ".." Then
                        If m_FTP.FtpFileExists("/" & m_RemoteDir & "/" & strFile) Then
                            m_FTP.Download("/" & m_RemoteDir & "/" & strFile, strLocalDirName & "/" & strFile, True)
                        End If
                    End If

                Next
                'Libération du verrou
                bReturn = unlockTo()
            End If
        Catch ex As Exception
            Debug.Assert(False, "clsFTPVinicom.downloadToDir" & ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function ' downloadToDir
    ''' <summary>
    ''' Rend le nombre de fichier dans le répertoire distant
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRemoteFileCount() As Integer
        Dim nReturn As Integer = 0

        Try

            nReturn = 0
            Dim lstFile As List(Of String) = m_FTP.ListDirectory("/" & m_RemoteDir)
            nReturn = lstFile.Count - 2
        Catch ex As Exception
            nReturn = 0
        End Try

        Return nReturn
    End Function ' getRemoteFileCount
    '=======================================================================
    'Fonction : shortResume()
    'Description : Rend un resumé de l'objet sous forme de chaine
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return ""
        End Get
    End Property

    Public Function deleteRemotefile(ByVal strFileName As String) As Boolean
        Dim bReturn As Boolean = False
        m_strErrorDescription = String.Empty
        Try
            If m_FTP.FtpFileExists(strFileName) Then
                bReturn = m_FTP.FtpDelete(strFileName)
            End If
        Catch ex As Exception
            Debug.Assert(False, ex.ToString())
            bReturn = False

        End Try
        Return bReturn
    End Function ' deleteRemoteFile
End Class
