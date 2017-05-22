'Option Strict Off
'Option Explicit On
'Public Class ftpwininet

'    'used for directory info
'    Private Const MAX_PATH As Short = 260
'    Private Const ERROR_NO_MORE_FILES As Short = 18

'    Private Structure FILETIME
'        Dim dwLowDateTime As Integer
'        Dim dwHighDateTime As Integer
'    End Structure

'    Private Structure WIN32_FIND_DATA
'        Dim dwFileAttributes As Integer
'        Dim ftCreationTime As FILETIME
'        Dim ftLastAccessTime As FILETIME
'        Dim ftLastWriteTime As FILETIME
'        Dim nFileSizeHigh As Integer
'        Dim nFileSizeLow As Integer
'        Dim dwReserved0 As Integer
'        Dim dwReserved1 As Integer
'        <VBFixedString(MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> Public cFileName As String
'        <VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=14)> Public cAlternate As String
'    End Structure

'    'collection to hold the directory listing
'    Private mcolItemList As ftpcolItem

'    'InternetOpen - lAccessType
'    Private Const INTERNET_OPEN_TYPE_PRECONFIG As Integer = 0
'    Private Const INTERNET_OPEN_TYPE_DIRECT As Integer = 1
'    Private Const INTERNET_OPEN_TYPE_PROXY As Integer = 3
'    'InternetOpen
'    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Integer, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Integer) As Integer

'    'InternetConnect - nServerPort
'    Private Const INTERNET_INVALID_PORT_NUMBER As Short = 0
'    'InternetConnect - lService
'    Private Const INTERNET_SERVICE_FTP As Integer = 1
'    'InternetConnect - lFlags
'    Private Const INTERNET_FLAG_PASSIVE As Integer = &H8000000
'    'InternetConnect
'    Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Integer, ByVal sServerName As String, ByVal nServerPort As Short, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Integer, ByVal lFlags As Integer, ByVal lContext As Integer) As Integer

'    'InternetCloseHandle
'    Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Integer) As Short

'    Private Const ERROR_INTERNET_EXTENDED_ERROR As Integer = 12003
'    'InternetGetLastResponseInfo
'    Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Integer, ByVal lpszBuffer As String, ByRef lpdwBufferLength As Integer) As Boolean

'    'InternetReadFile
'    Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Integer, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Integer, ByRef lpNumberOfBytesRead As Integer) As Boolean

'    'InternetWriteFile
'    Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Integer, ByVal lpBuffer As String, ByVal dwNumberOfBytesToWrite As Integer, ByRef lpdwNumberOfBytesWritten As Integer) As Boolean

'    'FtpPutFile/FtpGetFile - dwFlags
'    Private Const FTP_TRANSFER_TYPE_ASCII As Integer = &H1S
'    Private Const FTP_TRANSFER_TYPE_BINARY As Integer = &H2S

'    'FtpCreateDirectory
'    Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean

'    'FtpDeleteFile
'    Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Integer, ByVal lpszFileName As String) As Boolean

'    'FtpFindFirstFile
'    'UPGRADE_WARNING: La structure WIN32_FIND_DATA peut nécessiter que des attributs de marshaling soient passés en tant qu'argument dans cette instruction Declare. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
'    Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Integer, ByVal lpszSearchFile As String, ByRef lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Integer, ByVal dwContent As Integer) As Integer

'    'FtpGetCurrentDirectory
'    Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszCurrentDirectory As String, ByRef lpdwCurrentDirectory As Integer) As Boolean

'    'FtpGetFile
'    Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Integer, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Integer, ByVal dwFlags As Integer, ByVal dwContext As Integer) As Boolean

'    'FtpOpenFile - lAccess
'    Private Const GENERIC_READ As Integer = &H80000000
'    Private Const GENERIC_WRITE As Integer = &H40000000
'    'FtpOpenFile
'    Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Integer, ByVal sFileName As String, ByVal lAccess As Integer, ByVal lFlags As Integer, ByVal lContext As Integer) As Integer

'    'FtpPutFile
'    Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Integer, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Integer, ByVal dwContext As Integer) As Boolean

'    'FtpRemoveDirectory
'    Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean

'    'FtpRenameFile
'    Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Integer, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean

'    'FtpSetCurrentDirectory
'    Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean

'    'InternetFindNextFile
'    'UPGRADE_WARNING: La structure WIN32_FIND_DATA peut nécessiter que des attributs de marshaling soient passés en tant qu'argument dans cette instruction Declare. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
'    Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Integer, ByRef lpvFindData As WIN32_FIND_DATA) As Integer

'    'defaults - InternetOpen
'    Private Const DEFAULT_AGENT_NAME As String = "NIBLACK FTP DLL"
'    Private Const DEFAULT_ACCESS_TYPE As Integer = INTERNET_OPEN_TYPE_DIRECT

'    'defaults - InternetConnect
'    Private Const DEFAULT_USERID As String = "anonymous"
'    Private Const DEFAULT_PASSIVE_MODE As Boolean = True

'    'defaults - FtpGetFile/FtpPutFile
'    Private Const DEFAULT_TRANSFER_TYPE As Integer = FTP_TRANSFER_TYPE_BINARY
'    Private Const DEFAULT_OVER_WRITE As Short = 0

'    'defaults - FtpOpenFile
'    Private Const DEFAULT_FILE_ACCESS As Integer = GENERIC_WRITE

'    'class vars
'    Private mlngErrorNum As Integer 'error number
'    Private mstrErrorDesc As String 'error message

'    'class vars - InternetOpen
'    Private mstrAgentName As String 'agent name
'    Private mlngAccessType As Integer 'access type
'    Private mstrProxyName As String 'proxy name
'    Private mstrProxyBypass As String 'proxy bypass info
'    Private mlngINet As Integer 'handle from InternetOpen

'    'class vars - InternetConnect
'    Private mstrServerName As String 'ftp server name - dns/ip
'    Private mstrUserID As String 'user id
'    Private mstrPassword As String 'password
'    Private mblnPassiveMode As Boolean 'passive ftp syntax?
'    Private mlngINetConn As Integer 'handle from InternetConnect

'    'class vars - FtpOpenFile/InternetReadFile/InternetWriteFile
'    Private mlngINetConnFTP As Integer 'handle from FtpOpenFile
'    Private mlngFileAccess As Integer 'read or write flag for file

'    'class vars - FtpGetFile/FtpPutFile
'    Private mlngTransferType As Integer 'trasfer type - ascii/binary
'    Private mintOverWrite As Short 'overwrite existing file on get

'    Private Function bEnumDir(Optional ByVal strDir As String = "") As Boolean

'        Dim pData As WIN32_FIND_DATA = New WIN32_FIND_DATA()
'        Dim lngHINet As Integer
'        Dim intError As Short
'        Dim strTemp As String
'        Dim blnRC As Boolean
'        Dim blnDir As Boolean

'        'i'm optimistic
'        bEnumDir = True

'        'clear any existing items
'        mcolItemList.Clear()

'        'was a directory specfied?
'        If Len(strDir) > 0 Then

'            'move to the correct directory
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bSetCurrentDir(). Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            blnDir = bSetCurrentDir(strDir)

'        Else

'            'spoof rc
'            blnDir = True

'        End If

'        If blnDir Then

'            'init the filename buffer
'            pData.cFileName = New String(Chr(0), MAX_PATH)

'            'get the first file in the directory...
'            lngHINet = FtpFindFirstFile(mlngINetConn, "*.*", pData, 0, 0)

'            'how'd we do?
'            If lngHINet = 0 Then
'                'get the error from the findfirst call
'                intError = Err.LastDllError

'                'is the directory empty?
'                If intError <> ERROR_NO_MORE_FILES Then
'                    'whoa...a real error
'                    ErrorUpd(intError, Err.Description)

'                    'set bad return code
'                    bEnumDir = False
'                End If

'            Else

'                'we got some dir info...
'                'get the name
'                'strTemp = Left(pData.cFileName, InStr(1, pData.cFileName, New String(Chr(0), 1), CompareMethod.Binary) - 1)
'                strTemp = pData.cFileName
'                'store the item
'                mcolItemList.Add(strTemp, pData.dwFileAttributes)

'                'now loop through the rest of the files...
'                Do
'                    'init the filename buffer
'                    pData.cFileName = New String(Chr(0), MAX_PATH)

'                    'get the next item
'                    blnRC = InternetFindNextFile(lngHINet, pData)

'                    'how'd we do?
'                    If Not blnRC Then

'                        'get the error from the findnext call
'                        intError = Err.LastDllError

'                        'no more items
'                        If intError <> ERROR_NO_MORE_FILES Then
'                            'whoa...a real error
'                            ErrorUpd(intError, Err.Description)

'                            'set bad return code
'                            bEnumDir = False

'                            Exit Do

'                        Else

'                            'no more items...
'                            Exit Do

'                        End If

'                    Else

'                        'get the name
'                        '                        strTemp = Left(pData.cFileName, InStr(1, pData.cFileName, New String(Chr(0), 1), CompareMethod.Binary) - 1)
'                        strTemp = pData.cFileName
'                        'store the item
'                        mcolItemList.Add(strTemp, pData.dwFileAttributes)

'                    End If

'                Loop

'                'close the handle for the dir listing
'                InternetCloseHandle(lngHINet)

'            End If
'        End If
'    End Function


'    Public Function bQGetFile(ByVal strServerName As String, ByVal strUserID As String, ByVal strPassword As String, ByVal strSourceFile As String, ByVal strTargetFile As String, ByVal lngTransferType As Integer, ByVal blnOverWrite As Boolean) As Boolean

'        Dim blnRC As Boolean

'        'init the return code...i'm very optimistic
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQGetFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bQGetFile = True

'        'clean any old errors
'        ErrorClear()

'        'kill any exisiting connections
'        bDisconnect()

'        'set properties w/user's settings
'        sServerName = strServerName
'        sUserID = strUserID
'        sPassword = strPassword
'        lTransferType = lngTransferType
'        bOverWrite = blnOverWrite

'        'get a connection
'        If bConnect() Then

'            'put it!
'            blnRC = FtpGetFile(mlngINetConn, strSourceFile, strTargetFile, mintOverWrite, 0, mlngTransferType, 0)

'            'how'd we do?
'            If blnRC = False Then

'                'whoa...something happened
'                ErrorUpd(Err.Number, Err.Description)
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQGetFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQGetFile = False

'            Else

'                'we done good!
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQGetFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQGetFile = True

'            End If

'        Else

'            'setup error stuff
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."

'            'send bad return code
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQGetFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bQGetFile = False

'        End If

'        'clean up any handles
'        bDisconnect()

'    End Function

'    Public Function bQPutFile(ByVal strServerName As String, ByVal strUserID As String, ByVal strPassword As String, ByVal strSourceFile As String, ByVal strTargetFile As String, ByVal lngTransferType As Integer) As Boolean

'        Dim blnRC As Boolean

'        'init the return code...i'm very optimistic
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQPutFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bQPutFile = True

'        'clean any old errors
'        ErrorClear()

'        'kill any exisiting connections
'        bDisconnect()

'        'set properties w/user's settings
'        sServerName = strServerName
'        sUserID = strUserID
'        sPassword = strPassword
'        lTransferType = lngTransferType

'        'get a connection
'        If bConnect() Then

'            'put it!
'            blnRC = FtpPutFile(mlngINetConn, strSourceFile, strTargetFile, mlngTransferType, 0)

'            'how'd we do?
'            If blnRC = False Then

'                'whoa...something happened
'                ErrorUpd(Err.Number, Err.Description)
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQPutFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQPutFile = False

'            Else

'                'we done good!
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQPutFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQPutFile = True

'            End If

'        Else

'            'setup error stuff
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."

'            'send bad return code
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQPutFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bQPutFile = False

'        End If

'        'clean up any handles
'        bDisconnect()

'    End Function

'    Public Function sGetDirName(ByVal intItem As Short) As String

'        'clear any existing errors
'        ErrorClear()

'        'make sure the index is within our collection
'        If (intItem >= 1) And (intItem <= mcolItemList.Count) Then

'            'send back the name
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sGetDirName. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sGetDirName = mcolItemList.Item(intItem).Name

'        Else

'            'send back error
'            mlngErrorNum = -1
'            mstrErrorDesc = "Index out of range."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sGetDirName. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sGetDirName = ""

'        End If

'    End Function

'    Public Function lGetDirCount() As String

'        'clear any existing errors
'        ErrorClear()

'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lGetDirCount. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        lGetDirCount = mcolItemList.Count

'    End Function

'    Public Function bGetDir(Optional ByVal strDir As String = "", Optional ByVal intType As Short = 0) As Boolean

'        Dim lngEntries As Integer
'        Dim lngX As Integer
'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'I'm optimistic
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bGetDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bGetDir = True

'        'enumerate the directory listing
'        If Len(strDir) > 0 Then

'            'get specified directories contents
'            blnRC = bEnumDir(strDir)

'        Else

'            'get default directories contents
'            blnRC = bEnumDir()

'        End If

'        If blnRC Then

'            'get the number of entries in the dir listing
'            lngEntries = mcolItemList.Count

'            'do we have any?
'            If lngEntries > 0 Then

'                'yep...

'                'work through all the entries
'                For lngX = lngEntries To 1 Step -1

'                    'type flags?
'                    If intType > 0 Then

'                        'remove item if it doesn't match
'                        If Not (mcolItemList.Item(lngX).Attributes And intType) = intType Then
'                            mcolItemList.nRemove(lngX)
'                        End If

'                    End If

'                Next

'                'check for no entries matching
'                If mcolItemList.Count = 0 Then
'                    mlngErrorNum = 0
'                    mstrErrorDesc = "No Matching Entries"
'                    'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bGetDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                    bGetDir = True
'                End If

'            Else

'                'no entries...send back empty string
'                mlngErrorNum = 0
'                mstrErrorDesc = "Empty Directory"
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bGetDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bGetDir = True

'            End If

'        Else

'            'send back bad return code
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bGetDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bGetDir = False

'        End If

'    End Function

'    Public Function sListDir(Optional ByVal strDir As String = "", Optional ByVal intType As Short = 0) As String

'        Dim lngEntries As Integer
'        Dim lngX As Integer
'        Dim strTemp As String
'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'init return...send back empty string
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sListDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        sListDir = ""

'        'enumerate the directory listing
'        If Len(strDir) > 0 Then

'            'get specified directories contents
'            blnRC = bEnumDir(strDir)

'        Else

'            'get default directories contents
'            blnRC = bEnumDir()

'        End If

'        If blnRC Then

'            'get the number of entries in the dir listing
'            lngEntries = mcolItemList.Count

'            'do we have any?
'            If lngEntries > 0 Then

'                'yep...list 'em

'                'init return string
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sListDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                sListDir = ""

'                'work through all the entries
'                For lngX = 1 To lngEntries

'                    'init temp string
'                    strTemp = ""

'                    'get the name, if it matches
'                    If intType > 0 Then
'                        If mcolItemList.Item(lngX).Attributes And intType Then
'                            strTemp = mcolItemList.Item(lngX).Name
'                        End If
'                    Else
'                        strTemp = mcolItemList.Item(lngX).Name
'                    End If

'                    'did we find a match?
'                    If strTemp <> "" Then

'                        'add to the return string
'                        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sListDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                        If sListDir = "" Then
'                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sListDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                            sListDir = strTemp
'                        Else
'                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sListDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                            sListDir = sListDir & ";" & strTemp
'                        End If

'                    End If

'                Next

'                'check for no entries matching
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sListDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                If sListDir = "" Then
'                    mlngErrorNum = 0
'                    mstrErrorDesc = "No Matching Entries"
'                End If

'            Else

'                'no entries...send back empty string
'                mlngErrorNum = 0
'                mstrErrorDesc = "Empty Directory"
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sListDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                sListDir = ""

'            End If

'        End If

'    End Function

'    'UPGRADE_NOTE: Class_Initializea été mis à niveau vers Class_Initialize_Renamed. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
'    Private Sub Class_Initialize_Renamed()

'        'setup default class vars for this instance

'        'InternetOpen
'        mstrAgentName = DEFAULT_AGENT_NAME
'        mlngAccessType = DEFAULT_ACCESS_TYPE
'        mstrProxyName = ""
'        mstrProxyBypass = ""
'        mlngINet = 0

'        'InternetConnect
'        mstrServerName = ""
'        mstrUserID = DEFAULT_USERID
'        mstrPassword = ""
'        mblnPassiveMode = DEFAULT_PASSIVE_MODE
'        mlngINetConn = 0

'        'FtpGetFile/FtpPutFile
'        mlngTransferType = DEFAULT_TRANSFER_TYPE
'        mintOverWrite = DEFAULT_OVER_WRITE

'        'FtpOpenFile
'        mlngFileAccess = DEFAULT_FILE_ACCESS
'        mlngINetConnFTP = 0

'        'error stuff
'        mlngErrorNum = 0
'        mstrErrorDesc = ""

'        'listing stuff
'        mcolItemList = New ftpcolItem

'    End Sub
'    Public Sub New()
'        MyBase.New()
'        Class_Initialize_Renamed()
'    End Sub


'    'UPGRADE_NOTE: Class_Terminatea été mis à niveau vers Class_Terminate_Renamed. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
'    Private Sub Class_Terminate_Renamed()

'        'clean up after ourselves...incase the user didn't
'        bDisconnect()

'        'clear the item collection
'        'UPGRADE_NOTE: L'objet mcolItemList ne peut pas être détruit tant qu'il n'est pas récupéré par le garbage collector (ramasse-miettes). Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
'        mcolItemList = Nothing

'    End Sub
'    Protected Overrides Sub Finalize()
'        Class_Terminate_Renamed()
'        MyBase.Finalize()
'    End Sub
'    Public Function bDisconnect() As Boolean

'        'open file?
'        If mlngINetConnFTP <> 0 Then
'            InternetCloseHandle(mlngINetConnFTP)
'            mlngINetConnFTP = 0
'        End If

'        'active connection handle?
'        If mlngINetConn <> 0 Then
'            InternetCloseHandle(mlngINetConn)
'            mlngINetConn = 0
'        End If

'        'active open handle?
'        If mlngINet <> 0 Then
'            InternetCloseHandle(mlngINet)
'            mlngINet = 0
'        End If

'        'clear any item list
'        mcolItemList.Clear()

'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bDisconnect. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bDisconnect = True
'    End Function


'    Public Property lFileAccess() As Long
'        Get
'            If mlngFileAccess = GENERIC_WRITE Then
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lFileAccess. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                lFileAccess = 1
'            Else
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lFileAccess. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                lFileAccess = 2
'            End If
'        End Get
'        Set(ByVal Value As Long)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lngFileAccess. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            If Value = 1 Then
'                mlngFileAccess = GENERIC_WRITE
'            Else
'                mlngFileAccess = GENERIC_READ
'            End If
'        End Set
'    End Property


'    Public Property lAccessType() As Long
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lAccessType. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            lAccessType = mlngAccessType
'        End Get
'        Set(ByVal Value As Long)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lngAccessType. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            mlngAccessType = Value
'        End Set
'    End Property


'    Public Property sProxyName() As String
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sProxyName. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sProxyName = mstrProxyName
'        End Get
'        Set(ByVal Value As String)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet strProxyName. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            mstrProxyName = Trim(Value)
'        End Set
'    End Property


'    Public Property sProxyBypass() As String
'        Get
'            sProxyBypass = mstrProxyBypass
'        End Get
'        Set(ByVal Value As String)
'            mstrProxyBypass = Trim(Value)
'        End Set
'    End Property

'    Public ReadOnly Property sError() As String
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sError. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sError = mlngErrorNum & ": " & mstrErrorDesc
'        End Get
'    End Property

'    Public ReadOnly Property lErrorNum() As Long
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lErrorNum. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            lErrorNum = mlngErrorNum
'        End Get
'    End Property

'    Public ReadOnly Property sErrorDesc() As String
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sErrorDesc. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sErrorDesc = mstrErrorDesc
'        End Get
'    End Property


'    Public Property sServerName() As String
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sServerName. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sServerName = mstrServerName
'        End Get
'        Set(ByVal Value As String)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet strServerName. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            mstrServerName = Trim(Value)
'        End Set
'    End Property


'    Public Property sUserID() As String
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sUserID. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sUserID = mstrUserID
'        End Get
'        Set(ByVal Value As String)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet strUserID. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            mstrUserID = Trim(Value)
'        End Set
'    End Property


'    Public Property sPassword() As String
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sPassword. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sPassword = mstrPassword
'        End Get
'        Set(ByVal Value As String)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet strPassword. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            mstrPassword = Trim(Value)
'        End Set
'    End Property


'    Public Property bPassiveMode() As Boolean
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bPassiveMode. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bPassiveMode = mblnPassiveMode
'        End Get
'        Set(ByVal Value As Boolean)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet blnPassiveMode. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            mblnPassiveMode = Value
'        End Set
'    End Property


'    Public Property bOverWrite() As Boolean
'        Get
'            If mintOverWrite = 0 Then
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bOverWrite. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bOverWrite = True
'            Else
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bOverWrite. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bOverWrite = False
'            End If
'        End Get
'        Set(ByVal Value As Boolean)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet blnOverWrite. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            If Value Then
'                mintOverWrite = 0
'            Else
'                mintOverWrite = -1
'            End If
'        End Set
'    End Property


'    Public Property lTransferType() As Long
'        Get
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lTransferType. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            lTransferType = mlngTransferType
'        End Get
'        Set(ByVal Value As Long)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet lngTransferType. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            mlngTransferType = Value
'        End Set
'    End Property

'    Public Function bCloseFile() As Boolean

'        'clear any errors
'        ErrorClear()

'        'open file?
'        If mlngINetConnFTP <> 0 Then
'            InternetCloseHandle(mlngINetConnFTP)
'            mlngINetConnFTP = 0
'        End If

'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bCloseFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bCloseFile = True

'    End Function

'    Public Function sConvertCrLf(ByVal strData As String) As String

'        Dim lngLoc As Integer
'        Dim strTemp As String

'        'init the return string
'        strTemp = ""

'        Do
'            'get location of crlf
'            lngLoc = InStr(strData, vbCrLf)

'            'did we find it?
'            If lngLoc > 0 Then

'                'convert crlf to html br tag
'                strTemp = strTemp & Left(strData, lngLoc) & "<br>"

'                'reset string
'                strData = Mid(strData, lngLoc + 2)

'            Else

'                'wasn't found...just add rest of string
'                strTemp = strTemp & strData

'                'end looping
'                Exit Do

'            End If

'        Loop

'        'send back converted string
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sConvertCrLf. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        sConvertCrLf = strTemp

'    End Function

'    Public Function sReadFile() As String

'        Dim blnRC As Boolean
'        Dim lngBytes As Integer
'        Dim strTemp As String
'        Dim strBuffer As String

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have an open file
'        If mlngINetConnFTP <> 0 Then

'            'init temp string
'            strTemp = ""

'            'loop through the entire file
'            Do

'                'init the data buffer
'                strBuffer = New String(Chr(0), 1024)

'                'read upto 1024 bytes of data...
'                blnRC = InternetReadFile(mlngINetConnFTP, strBuffer, 1024, lngBytes)

'                'how'd we do?
'                If blnRC Then

'                    'done good...save the data
'                    strTemp = strTemp & Left(strBuffer, lngBytes)

'                    'check eof (lngBytes = 0)
'                    If lngBytes = 0 Then

'                        'file's complete...send back the data
'                        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sReadFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                        sReadFile = strTemp

'                        Exit Do

'                    End If

'                Else

'                    'send back error
'                    ErrorUpd(Err.LastDllError, Err.Description)
'                    'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sReadFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                    sReadFile = ""

'                    Exit Do

'                End If

'            Loop

'        Else

'            'send back error
'            mlngErrorNum = -1
'            mstrErrorDesc = "No File Opened."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sReadFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sReadFile = ""

'        End If

'    End Function

'    Public Function bWriteFile(ByVal strData As String) As Boolean

'        Dim blnRC As Boolean
'        Dim lngBytes As Integer

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have an open file
'        If mlngINetConnFTP <> 0 Then

'            'write the data to the file
'            blnRC = InternetWriteFile(mlngINetConnFTP, strData, Len(strData), lngBytes)

'            'how'd we do?
'            If blnRC Then

'                'check to make sure the number of bytes written match
'                If lngBytes = Len(strData) Then

'                    'done good!
'                    'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bWriteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                    bWriteFile = True

'                Else

'                    'send back error
'                    ErrorUpd(Err.LastDllError, Err.Description)
'                    'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bWriteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                    bWriteFile = False

'                End If

'            Else

'                'send back error
'                ErrorUpd(Err.LastDllError, Err.Description)
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bWriteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bWriteFile = False

'            End If

'        Else

'            'send back error
'            mlngErrorNum = -1
'            mstrErrorDesc = "No File Opened."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bWriteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bWriteFile = False

'        End If

'    End Function

'    Public Function bOpenFile(ByVal strFileName As String) As Boolean


'        'clean up any old errors
'        ErrorClear()

'        'I'm optimistic
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bOpenFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bOpenFile = True

'        'make sure we have a connection
'        If mlngINetConn = 0 Then

'            'whoa...we need a connection first
'            bConnect()

'        End If

'        'we better have a connection now
'        If mlngINetConn <> 0 Then

'            'open the file
'            mlngINetConnFTP = FtpOpenFile(mlngINetConn, strFileName, mlngFileAccess, mlngTransferType, 0)

'            'how'd we do?
'            If mlngINetConnFTP = 0 Then

'                'whoa...something happened
'                ErrorUpd(Err.LastDllError, Err.Description)
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bOpenFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bOpenFile = False

'            Else

'                'we done good!
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bOpenFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bOpenFile = True

'            End If
'        Else

'            'send back error
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bOpenFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bOpenFile = False

'        End If

'    End Function

'    Public Function bConnect() As Boolean

'        Dim lngFlags As Integer

'        'clean up any old errors
'        ErrorClear()

'        'do we have an open handle...just a safeguard for the users?
'        If mlngINet = 0 Then
'            'get new open handle based on proxy info
'            If mlngAccessType <> INTERNET_OPEN_TYPE_DIRECT Then
'                'use proxy info
'                mlngINet = InternetOpen(mstrAgentName, mlngAccessType, mstrProxyName, mstrProxyBypass, 0)
'            Else
'                'direct...
'                mlngINet = InternetOpen(mstrAgentName, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'            End If
'        End If

'        'check for an open handle again...
'        If mlngINet <> 0 Then
'            'we have a open handle...now get a connection handle

'            'do we already have a connection handle?
'            If mlngINetConn <> 0 Then

'                'connection exists...good return code
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bConnect. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bConnect = True

'            Else

'                If mstrServerName = "" Then
'                    'can't work without a server name...
'                    mstrErrorDesc = "Server name required."
'                    mlngErrorNum = -1

'                    'clean up handles
'                    bDisconnect()

'                    'set bad return code
'                    'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bConnect. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                    bConnect = False

'                Else

'                    'are we in passive mode?
'                    If mblnPassiveMode Then
'                        lngFlags = INTERNET_FLAG_PASSIVE
'                    Else
'                        lngFlags = 0
'                    End If

'                    'get connection
'                    mlngINetConn = InternetConnect(mlngINet, mstrServerName, INTERNET_INVALID_PORT_NUMBER, mstrUserID, mstrPassword, INTERNET_SERVICE_FTP, lngFlags, 0)

'                    'did we get a connection?
'                    If mlngINetConn <> 0 Then

'                        'pass back good return code
'                        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bConnect. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                        bConnect = True

'                    Else

'                        'whoa nellie...something bad happened
'                        ErrorUpd(Err.LastDllError, Err.Description)
'                        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bConnect. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                        bConnect = False

'                    End If
'                End If
'            End If
'        Else

'            'couldn't get a open handle
'            ErrorUpd(Err.LastDllError, Err.Description)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bConnect. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bConnect = False

'        End If

'    End Function

'    Public Function bRename(ByVal strCurrentName As String, ByVal strNewName As String) As Boolean

'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have a connection
'        If mlngINetConn = 0 Then

'            'whoa...we need a connection first
'            bConnect()

'        End If

'        'we better have a connection now
'        If mlngINetConn <> 0 Then
'            'put it!
'            blnRC = FtpRenameFile(mlngINetConn, strCurrentName, strNewName)

'            'how'd we do?
'            If blnRC = False Then

'                'whoa...something happened
'                ErrorUpd(Err.LastDllError, Err.Description)
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bRename. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bRename = False

'            Else

'                'we done good!
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bRename. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bRename = True

'            End If
'        Else

'            'send back error
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bRename. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bRename = False

'        End If

'    End Function

'    Public Function bGetFile(ByVal strSourceFile As String, ByVal strTargetFile As String) As Boolean

'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have a connection
'        If mlngINetConn = 0 Then

'            'whoa...we need a connection first
'            bConnect()

'        End If

'        'we better have a connection now
'        If mlngINetConn <> 0 Then
'            'put it!
'            blnRC = FtpGetFile(mlngINetConn, strSourceFile, strTargetFile, mintOverWrite, 0, mlngTransferType, 0)

'            'how'd we do?
'            If blnRC = False Then

'                'whoa...something happened
'                ErrorUpd(Err.LastDllError, Err.Description)
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bGetFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bGetFile = False

'            Else

'                'we done good!
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bGetFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bGetFile = True

'            End If
'        Else

'            'send back error
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bGetFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bGetFile = False

'        End If

'    End Function

'    Public Function bPutFile(ByVal strSourceFile As String, ByVal strTargetFile As String) As Boolean

'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have a connection
'        If mlngINetConn = 0 Then

'            'whoa...we need a connection first
'            bConnect()

'        End If

'        'we better have a connection now
'        If mlngINetConn <> 0 Then
'            'put it!
'            blnRC = FtpPutFile(mlngINetConn, strSourceFile, strTargetFile, mlngTransferType, 0)

'            'how'd we do?
'            If blnRC = False Then

'                'whoa...something happened
'                ErrorUpd(Err.LastDllError, Err.Description)
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bPutFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bPutFile = False

'            Else

'                'we done good!
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bPutFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bPutFile = True

'            End If
'        Else

'            'send back error
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bPutFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bPutFile = False

'        End If

'    End Function

'    Public Function bQRemoveDir(ByVal strServerName As String, ByVal strUserID As String, ByVal strPassword As String, ByVal strDir As String) As Boolean

'        Dim blnRC As Boolean

'        'init the return code...i'm very optimistic
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQRemoveDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bQRemoveDir = True

'        'clean up any old errors
'        ErrorClear()

'        'kill any exisiting connections
'        bDisconnect()

'        'set properties w/user's settings
'        sServerName = strServerName
'        sUserID = strUserID
'        sPassword = strPassword

'        'get a connection
'        If bConnect() Then

'            'delete the file
'            blnRC = FtpRemoveDirectory(mlngINetConn, strDir)

'            'how'd we do?
'            If blnRC Then

'                'send back good return code
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQRemoveDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQRemoveDir = True

'            Else

'                'send back the error
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQRemoveDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQRemoveDir = False
'                ErrorUpd(Err.LastDllError, Err.Description)

'            End If

'        Else

'            'whoa...we don't have a connection
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQRemoveDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bQRemoveDir = False

'        End If

'        'clean up any handles
'        bDisconnect()

'    End Function

'    Public Function bQMakeDir(ByVal strServerName As String, ByVal strUserID As String, ByVal strPassword As String, ByVal strDir As String) As Boolean

'        Dim blnRC As Boolean

'        'init the return code...i'm very optimistic
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQMakeDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bQMakeDir = True

'        'clean up any old errors
'        ErrorClear()

'        'kill any exisiting connections
'        bDisconnect()

'        'set properties w/user's settings
'        sServerName = strServerName
'        sUserID = strUserID
'        sPassword = strPassword

'        'get a connection
'        If bConnect() Then

'            'delete the file
'            blnRC = FtpCreateDirectory(mlngINetConn, strDir)

'            'how'd we do?
'            If blnRC Then

'                'send back good return code
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQMakeDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQMakeDir = True

'            Else

'                'send back the error
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQMakeDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQMakeDir = False
'                ErrorUpd(Err.LastDllError, Err.Description)

'            End If

'        Else

'            'whoa...we don't have a connection
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQMakeDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bQMakeDir = False

'        End If

'        'clean up any handles
'        bDisconnect()

'    End Function

'    Public Function bQRename(ByVal strServerName As String, ByVal strUserID As String, ByVal strPassword As String, ByVal strCurrentName As String, ByVal strNewName As String) As Boolean

'        Dim blnRC As Boolean

'        'init the return code...i'm very optimistic
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQRename. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bQRename = True

'        'clean up any old errors
'        ErrorClear()

'        'kill any exisiting connections
'        bDisconnect()

'        'set properties w/user's settings
'        sServerName = strServerName
'        sUserID = strUserID
'        sPassword = strPassword

'        'get a connection
'        If bConnect() Then

'            'delete the file
'            blnRC = FtpRenameFile(mlngINetConn, strCurrentName, strNewName)

'            'how'd we do?
'            If blnRC Then

'                'send back good return code
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQRename. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQRename = True

'            Else

'                'send back the error
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQRename. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQRename = False
'                ErrorUpd(Err.LastDllError, Err.Description)

'            End If

'        Else

'            'whoa...we don't have a connection
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQRename. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bQRename = False

'        End If

'        'clean up any handles
'        bDisconnect()

'    End Function

'    Public Function bQDeleteFile(ByVal strServerName As String, ByVal strUserID As String, ByVal strPassword As String, ByVal strFile As String) As Boolean

'        Dim blnRC As Boolean

'        'init the return code...i'm very optimistic
'        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQDeleteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'        bQDeleteFile = True

'        'clean up any old errors
'        ErrorClear()

'        'kill any exisiting connections
'        bDisconnect()

'        'set properties w/user's settings
'        sServerName = strServerName
'        sUserID = strUserID
'        sPassword = strPassword

'        'get a connection
'        If bConnect() Then

'            'delete the file
'            blnRC = FtpDeleteFile(mlngINetConn, strFile)

'            'how'd we do?
'            If blnRC Then

'                'send back good return code
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQDeleteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQDeleteFile = True

'            Else

'                'send back the error
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQDeleteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bQDeleteFile = False
'                ErrorUpd(Err.LastDllError, Err.Description)

'            End If

'        Else

'            'whoa...we don't have a connection
'            mlngErrorNum = -1
'            mstrErrorDesc = "Could not connect to server."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bQDeleteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bQDeleteFile = False

'        End If

'        'clean up any handles
'        bDisconnect()

'    End Function

'    Public Function bRemoveDir(ByVal strDir As String) As Boolean


'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have a connection
'        If mlngINetConn <> 0 Then

'            'delete the file
'            blnRC = FtpRemoveDirectory(mlngINetConn, strDir)

'            'how'd we do?
'            If blnRC Then

'                'send back good return code
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bRemoveDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bRemoveDir = True

'            Else

'                'send back the error
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bRemoveDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bRemoveDir = False
'                ErrorUpd(Err.LastDllError, Err.Description)

'            End If

'        Else

'            'whoa...we don't have a connection
'            mlngErrorNum = -1
'            mstrErrorDesc = "No available connection."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bRemoveDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bRemoveDir = False

'        End If

'    End Function

'    Public Function bMakeDir(ByVal strDir As String) As Boolean

'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have a connection
'        If mlngINetConn <> 0 Then

'            'delete the file
'            blnRC = FtpCreateDirectory(mlngINetConn, strDir)

'            'how'd we do?
'            If blnRC Then

'                'send back good return code
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bMakeDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bMakeDir = True

'            Else

'                'send back the error
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bMakeDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bMakeDir = False
'                ErrorUpd(Err.LastDllError, Err.Description)

'            End If

'        Else

'            'whoa...we don't have a connection
'            mlngErrorNum = -1
'            mstrErrorDesc = "No available connection."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bMakeDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bMakeDir = False

'        End If

'    End Function

'    Public Function bDeleteFile(ByVal strFile As String) As Boolean

'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have a connection
'        If mlngINetConn <> 0 Then

'            'delete the file
'            blnRC = FtpDeleteFile(mlngINetConn, strFile)

'            'how'd we do?
'            If blnRC Then

'                'send back good return code
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bDeleteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bDeleteFile = True

'            Else

'                'send back the error
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bDeleteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bDeleteFile = False
'                ErrorUpd(Err.LastDllError, Err.Description)

'            End If

'        Else

'            'whoa...we don't have a connection
'            mlngErrorNum = -1
'            mstrErrorDesc = "No available connection."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bDeleteFile. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bDeleteFile = False

'        End If

'    End Function

'    Public Function bSetCurrentDir(ByVal strDirName As String) As Boolean

'        Dim strTempDir As String
'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have a connection
'        If mlngINetConn <> 0 Then

'            'init the buffer for the returned directory info
'            strTempDir = Space(MAX_PATH)

'            'get the current directory name
'            blnRC = FtpSetCurrentDirectory(mlngINetConn, strDirName)

'            'how'd we do?
'            If blnRC Then

'                'send back the current directory name
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bSetCurrentDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bSetCurrentDir = True

'            Else

'                'send back the error
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bSetCurrentDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                bSetCurrentDir = False
'                ErrorUpd(Err.LastDllError, Err.Description)

'            End If

'        Else

'            'whoa...we don't have a connection
'            mlngErrorNum = -1
'            mstrErrorDesc = "No available connection."
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet bSetCurrentDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            bSetCurrentDir = False

'        End If

'    End Function

'    Private Sub ErrorClear()

'        'just clean up old error vars
'        mlngErrorNum = 0
'        mstrErrorDesc = ""

'    End Sub

'    Private Sub ErrorUpd(ByVal lngErr As Integer, ByVal strErrDesc As String)

'        Dim strError As String
'        Dim blnRC As Boolean
'        Dim lngErrX As Integer

'        'see if there's more info
'        If lngErr = ERROR_INTERNET_EXTENDED_ERROR Then

'            'get the additional info
'            strError = Space(255)
'            blnRC = InternetGetLastResponseInfo(lngErrX, strError, 255)

'            'did we get the additional info?
'            If blnRC Then

'                'we got something...need to clean it up
'                strError = Left(strError, InStr(1, strError, Chr(0), CompareMethod.Binary) - 1)

'                'update error stuff
'                mlngErrorNum = Val(strError)
'                mstrErrorDesc = Mid(strError, InStr(strError, CStr(mlngErrorNum)) + Len(CStr(mlngErrorNum)) + 1)

'            Else

'                'send back previous error
'                mlngErrorNum = lngErr
'                mstrErrorDesc = strErrDesc

'            End If

'        Else

'            'update error stuff
'            mlngErrorNum = lngErr
'            mstrErrorDesc = strErrDesc

'        End If

'    End Sub
'    Public Function sGetCurrentDir() As String

'        Dim strTempDir As String
'        Dim blnRC As Boolean

'        'clean up any old errors
'        ErrorClear()

'        'make sure we have a connection
'        If mlngINetConn <> 0 Then

'            'init the buffer for the returned directory info
'            strTempDir = Space(MAX_PATH)

'            'get the current directory name
'            blnRC = FtpGetCurrentDirectory(mlngINetConn, strTempDir, MAX_PATH)

'            'how'd we do?
'            If blnRC Then

'                'we got something...need to clean it up
'                strTempDir = Left(strTempDir, InStr(1, strTempDir, Chr(0), CompareMethod.Binary) - 1)

'                'send back the current directory name
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sGetCurrentDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                sGetCurrentDir = strTempDir

'            Else

'                'send back the error
'                ErrorUpd(Err.LastDllError, Err.Description)
'                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sGetCurrentDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'                sGetCurrentDir = mlngErrorNum & ": " & mstrErrorDesc

'            End If

'        Else

'            'whoa...we don't have a connection
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet sGetCurrentDir. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            sGetCurrentDir = "-1: No available connection."

'        End If


'    End Function
'    Public Function isConnected() As Boolean
'        Return (mlngINetConn > 0)
'    End Function
'    Public Function remoteFileExists(ByVal strFileName As String) As Boolean
'        Debug.Assert(isConnected, "Not is Connected")
'        Dim strDirContent As String

'        strDirContent = sListDir()
'        Return (InStr(strDirContent, strFileName) > 0)
'    End Function
'End Class