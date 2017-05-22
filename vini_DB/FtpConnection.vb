'Option Strict On
'Option Explicit On

'Imports System
'Imports System.Text
'Imports System.Net
'Imports System.Net.Sockets

''/// <summary>Represents a connection to an FTP server.</summary>
'Public Class FtpConnection
'    '/// <summary>Enumerates the different stream modes.</summary>
'    Public Enum StreamModes
'        ASCII = 65
'        Binary = 73
'    End Enum
'    '/// <summary>Raised when the object successfully connected to the remote host.</summary>
'    Event Connected()
'    '/// <summary>Raised when the connection to the remote host failed or when the connection to the host is lost.</summary>
'    Event ConnectionFailed()
'    '/// <summary>Raised when the connection to the remote host is closed.</summary>
'    Event Disconnected()
'    '/// <summary>Raised when an FTP command has been sent to the remote host.</summary>
'    Event CommandSent(ByVal Command As String)
'    '/// <summary>Raised when a command was successfully completed.</summary>
'    '/// <remarks>Do note that this does *not* mean the command has been successful.</remarks>
'    Event CommandCompleted()
'    '/// <summary>Raised when a the object received a reply from the remote host.</summary>
'    Event ReceivedReply(ByVal Reply As String, ByVal ReplyNumber As Short, ByVal ReplyType As Short)
'    '/// <summary>Raised when a new directory listing is available.</summary>
'    Event NewDirectoryListing()
'    '/// <summary>Specifies the remote port to use when the object connects to the remote host.</summary>
'    '/// <value>An Integer that specifies the remote port to use when the object connects to the remote host.</value>
'    Public Property Port() As Integer
'        Get
'            Return remotePort
'        End Get
'        Set(ByVal value As Integer)
'            remotePort = value
'        End Set
'    End Property
'    '/// <summary>Specifies the remote host to connect to.</summary>
'    '/// <value>A String that specifies the remote host to connect to.</value>
'    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null)<exceptions>
'    Public Property Hostname() As String
'        Get
'            Return remoteAddress
'        End Get
'        Set(ByVal value As String)
'            If value Is Nothing Then Throw New ArgumentNullException()
'            remoteAddress = value
'        End Set
'    End Property
'    '/// <summary>Specifies the Username to use when connecting to a remote host.</summary>
'    '/// <value>A String that specifies the Username to use when connecting to a remote host.</value>
'    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null)<exceptions>
'    Public Property Username() As String
'        Get
'            Return logonName
'        End Get
'        Set(ByVal value As String)
'            If value Is Nothing Then Throw New ArgumentNullException()
'            logonName = value
'        End Set
'    End Property
'    '/// <summary>Specifies the Password to use when connecting to a remote host.</summary>
'    '/// <value>A String that specifies the Password to use when connecting to a remote host.</value>
'    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null)<exceptions>
'    Public Property Password() As String
'        Get
'            Return logonPass
'        End Get
'        Set(ByVal value As String)
'            If value Is Nothing Then Throw New ArgumentNullException()
'            logonPass = value
'        End Set
'    End Property
'    '/// <summary>Specifies the WorkDir of the FTP connection.</summary>
'    '/// <value>A String that specifies the WorkDir of the FTP connection.</value>
'    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null)<exceptions>
'    Public Property WorkDir() As String
'        Get
'            Return remoteDirectory
'        End Get
'        Set(ByVal value As String)
'            If value Is Nothing Then Throw New ArgumentNullException()
'            If IsConnected Then
'                If Not remoteDirectory.Equals(value) Then ChangeWorkDir(value)
'            Else
'                remoteDirectory = value
'            End If
'        End Set
'    End Property
'    '/// <summary>Returns a Boolean that specifies whether the object is connected to a remote host or not.</summary>
'    '/// <value>A Boolean that specifies whether the object is connected to a remote host or not.</value>
'    Public ReadOnly Property IsConnected() As Boolean
'        Get
'            Return isNowConnected
'        End Get
'    End Property
'    '/// <summary>Specifies whether to use passive transfers or not.</summary>
'    '/// <value>A Boolean that specifies whether to use passive transfers or not.</value>
'    Public Property Passive() As Boolean
'        Get
'            Return passiveTransfers
'        End Get
'        Set(ByVal value As Boolean)
'            passiveTransfers = value
'        End Set
'    End Property
'    '/// <summary>Connects to the remote host and logs on.</summary>
'    Public Sub Connect()
'        Try
'            clientSocket = New Socket(AddressFamily.Unspecified, SocketType.Stream, ProtocolType.Tcp)
'            clientSocket.Connect(New IPEndPoint(Dns.GetHostEntry(remoteAddress).AddressList(0), remotePort))
'            isNowConnected = True
'            RaiseEvent Connected()
'        Catch e As SocketException
'            isNowConnected = False
'            RaiseEvent ConnectionFailed()
'            Exit Sub
'        End Try
'        WaitForResponse()
'        SendCommand("USER " + logonName)
'        If lastResponseType = 4 OrElse lastResponseType = 5 Then
'            isNowConnected = False
'            Exit Sub
'        End If
'        SendCommand("PASS " + logonPass)
'        If lastResponseType = 4 OrElse lastResponseType = 5 Then
'            isNowConnected = False
'            Exit Sub
'        End If
'        If remoteDirectory = "" Then
'            PrintWorkDir()
'        Else
'            ChangeWorkDir(remoteDirectory)
'        End If
'    End Sub
'    '/// <summary>Logs off and disconnects from the remote host.</summary>
'    Public Sub Disconnect()
'        If Not (isNowConnected) Then Exit Sub
'        SendCommand("QUIT")
'        Try
'            clientSocket.Shutdown(SocketShutdown.Both)
'        Finally
'            clientSocket.Close()
'            isNowConnected = False
'        End Try
'        RaiseEvent Disconnected()
'    End Sub
'    '/// <summary>Starts receiving data from the command socket, until it has received a complete reply.</summary>
'    Private Sub WaitForResponse()
'        If Not (isNowConnected) Then Exit Sub
'        Dim tempBuffer(1023) As Byte
'        Dim retBytes As Integer
'        Dim retString As String = ""
'        Do
'            Try
'                retBytes = clientSocket.Receive(tempBuffer)
'            Catch
'                Exit Do
'            End Try
'            retString = retString + ASCII.GetString(tempBuffer, 0, retBytes)
'        Loop Until IsValidResponse(retString)
'        lastResponse = retString
'        If retString.Length >= 3 Then
'            lastResponseNumber = Short.Parse(retString.Substring(0, 3))
'        Else
'            lastResponseNumber = 0
'        End If
'        lastResponseType = CType(Math.Floor(lastResponseNumber / 100), Short)
'        RaiseEvent ReceivedReply(lastResponse, lastResponseNumber, lastResponseType)
'    End Sub
'    '/// <summary>Checks whether a string is a valid reply or not.</summary>
'    '/// <param name="Input">The string that has to be checked for validity.</param>
'    Private Function IsValidResponse(ByVal Input As String) As Boolean
'        Dim Lines() As String = Input.Split(Convert.ToChar(10))
'        If Lines.Length > 1 Then
'            Try
'                If Lines(Lines.Length - 2).Replace(Convert.ToChar(13), "").Substring(3, 1).Equals(" ") Then Return True
'            Catch
'                Return False
'            End Try
'        End If
'        Return False
'    End Function
'    '/// <summary>Sends the specified command to the server.</summary>
'    '/// <param name="Command">The command to be sent to the server.</param>
'    Private Sub SendCommand(ByVal Command As String)
'        Try
'            clientSocket.Send(ASCII.GetBytes(Command + Convert.ToChar(13) + Convert.ToChar(10))) 'ControlChars.CrLf))
'            If Command.Length >= 4 AndAlso Command.Substring(0, 4).ToUpper.Equals("PASS") Then
'                RaiseEvent CommandSent("PASS ********")
'            Else
'                RaiseEvent CommandSent(Command)
'            End If
'            WaitForResponse()
'        Catch
'            RaiseEvent ConnectionFailed()
'            lastResponseNumber = 0
'            lastResponseType = 0
'            lastResponse = ""
'        End Try
'    End Sub
'    '/// <summary>Returns the last response number.</summary>
'    '/// <returns>A Short that represents the last response number.</returns>
'    Public Function GetLastResponseNumber() As Short
'        Return lastResponseNumber
'    End Function
'    '/// <summary>Returns the last response.</summary>
'    '/// <returns>A String that represents the last response.</returns>
'    Public Function GetLastResponse() As String
'        Return lastResponse
'    End Function
'    '/// <summary>Aborts an FTP transfer.</summary>
'    Public Sub Abort()
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("ABOR")
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Changes the current workdir to the specified directory.</summary>
'    '/// <param name="NewDirectory">The new directory.<param>
'    Public Sub ChangeWorkDir(ByVal NewDirectory As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("CWD " + NewDirectory)
'        PrintWorkDir()
'    End Sub
'    '/// <summary>Changes the current workdir to the parent directory.</summary>
'    Public Sub ChangeToUpDir()
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("CDUP")
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Changes the permissions of a file or directory.</summary>
'    '/// <param name="Filename">The filename(s) or directory name(s) to modify.<param>
'    '/// <param name="NewMode">The new permissions for the specified file(s)/dir(s).<param>
'    Public Sub ChMod(ByVal Filename As String, ByVal NewMode As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("SITE chmod " + NewMode + " " + Filename)
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Sends a custom command to the server.</summary>
'    '/// <param name="Command">The custom command to send.<param>
'    Public Sub CustomCommand(ByVal Command As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand(Command)
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Deletes a file from the server.</summary>
'    '/// <param name="Filename">The file to delete.<param>
'    Public Sub Delete(ByVal Filename As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("DELE " + Filename)
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Retrieves information about the FTP server.</summary>
'    Public Sub GetSystemInfo()
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("SYST")
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Retrieves help from the server.</summary>
'    Public Overloads Sub Help()
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("HELP")
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Retrieves help from the server about a specific command.</summary>
'    '/// <param name="CommandName">The command to get help for.<param>
'    Public Overloads Sub Help(ByVal CommandName As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("HELP " + CommandName)
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Creates a directory on the server.</summary>
'    '/// <param name="DirectoryName">The name of the new directory.<param>
'    Public Sub MakeDir(ByVal DirectoryName As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("MKD " + DirectoryName)
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Does nothing.</summary>
'    '/// <remarks>This command is used to keep connections from timing out.</remarks>
'    Public Sub NoOperation()
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("NOOP")
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Retrieves the current work dir.</summary>
'    Public Sub PrintWorkDir()
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("PWD")
'        Dim startPos As Integer = lastResponse.IndexOf("""")
'        Dim endPos As Integer
'        Try
'            If startPos >= 0 Then
'                endPos = lastResponse.IndexOf("""", startPos + 1)
'                If endPos >= 0 Then
'                    remoteDirectory = lastResponse.Substring(startPos + 1, endPos - startPos - 1)
'                Else
'                    remoteDirectory = lastResponse.Substring(4)
'                End If
'            Else
'                remoteDirectory = lastResponse.Substring(4)
'            End If
'        Catch
'        End Try
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Removes a directory from the server.</summary>
'    '/// <param name="DirectoryName">The directory to remove.<param>
'    Public Sub RemoveDir(ByVal DirectoryName As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("RMD " + DirectoryName)
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Renames a file on the server.</summary>
'    '/// <param name="OriginalName">The name of the original file.<param>
'    '/// <param name="NewName">The new name of the file.<param>
'    Public Sub Rename(ByVal OriginalName As String, ByVal NewName As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("RNFR " + OriginalName)
'        If lastResponseType = 3 Then SendCommand("RNTO " + NewName)
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Executes a site command on the server.</summary>
'    '/// <param name="Command">The site command to execute.<param>
'    Public Sub SiteCommand(ByVal Command As String)
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("SITE " + Command)
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Gets the status of the server.</summary>
'    Public Sub Status()
'        If Not (IsConnected) Then Exit Sub
'        SendCommand("STAT")
'        RaiseEvent CommandCompleted()
'    End Sub
'    '/// <summary>Retrieves the directory listing of the remote path.</summary>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Function GetList() As Boolean
'        If Not (IsConnected) Then Return False
'        SendCommand("TYPE A")
'        If Not CreateDataSocket(DataConnection.StreamDirections.Download) Then
'            RaiseEvent CommandCompleted()
'            Return False
'        End If
'        If Not passiveTransfers Then
'            Dim MyEndPoint As IPEndPoint = dataSocket.GetLocalEndPoint()
'            SendCommand("PORT " + MyEndPoint.Address.ToString.Replace(".", ",") + "," + CType(Math.Floor(MyEndPoint.Port / 256), Integer).ToString + "," + (MyEndPoint.Port Mod 256).ToString)
'        End If
'        SendCommand("LIST")
'        If lastResponseType <> 4 AndAlso lastResponseType <> 5 Then
'            Try
'                dataSocket.ReceiveFromSocket()
'            Catch
'                dataSocket.Close()
'                WaitForResponse()
'                RaiseEvent CommandCompleted()
'                Return False
'            End Try
'            WaitForResponse()
'        End If
'        ParseDirList(dataSocket.Data)
'        dataSocket.Close()
'        RaiseEvent NewDirectoryListing()
'        RaiseEvent CommandCompleted()
'        Return True
'    End Function
'    '/// <summary>Downloads a file from the remote server.</summary>
'    '/// <param name="RemoteFile">Specifies the remote file to download.</param>
'    '/// <param name="LocalFile">Specifies the local file to download to.</param>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Overloads Function DownloadFile(ByVal RemoteFile As String, ByVal LocalFile As String) As Boolean
'        Return DownloadFile(RemoteFile, LocalFile, StreamModes.Binary, DataConnection.FileModes.Overwrite, 0)
'    End Function
'    '/// <summary>Downloads a file from the remote server.</summary>
'    '/// <param name="RemoteFile">Specifies the remote file to download.</param>
'    '/// <param name="LocalFile">Specifies the local file to download to.</param>
'    '/// <param name="StreamMode">Specifies the transfer type.</param>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Overloads Function DownloadFile(ByVal RemoteFile As String, ByVal LocalFile As String, ByVal StreamMode As StreamModes) As Boolean
'        Return DownloadFile(RemoteFile, LocalFile, StreamMode, DataConnection.FileModes.Overwrite, 0)
'    End Function
'    '/// <summary>Downloads a file from the remote server.</summary>
'    '/// <param name="RemoteFile">Specifies the remote file to download.</param>
'    '/// <param name="LocalFile">Specifies the local file to download to.</param>
'    '/// <param name="FileMode">Specifies the file mode.</param>
'    '/// <param name="AppendFrom">Specifies where to append from.</param>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Overloads Function DownloadFile(ByVal RemoteFile As String, ByVal LocalFile As String, ByVal FileMode As DataConnection.FileModes, ByVal AppendFrom As Long) As Boolean
'        Return DownloadFile(RemoteFile, LocalFile, StreamModes.Binary, DataConnection.FileModes.Overwrite, AppendFrom)
'    End Function
'    '/// <summary>Downloads a file from the remote server.</summary>
'    '/// <param name="RemoteFile">Specifies the remote file to download.</param>
'    '/// <param name="LocalFile">Specifies the local file to download to.</param>
'    '/// <param name="StreamMode">Specifies the transfer type.</param>
'    '/// <param name="FileMode">Specifies the file mode.</param>
'    '/// <param name="AppendFrom">Specifies where to append from.</param>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Overloads Function DownloadFile(ByVal RemoteFile As String, ByVal LocalFile As String, ByVal StreamMode As StreamModes, ByVal FileMode As DataConnection.FileModes, ByVal AppendFrom As Long) As Boolean
'        SendCommand("TYPE " + Convert.ToChar(StreamMode))
'        If Not CreateDataSocket(LocalFile, FileMode, DataConnection.StreamDirections.Download, AppendFrom) Then
'            RaiseEvent CommandCompleted()
'            Return False
'        End If
'        If Not (passiveTransfers) Then
'            Dim MyEndPoint As IPEndPoint = dataSocket.GetLocalEndPoint()
'            SendCommand("PORT " + MyEndPoint.Address.ToString.Replace(".", ",") + "," + CType(Math.Floor(MyEndPoint.Port / 256), Integer).ToString + "," + (MyEndPoint.Port Mod 256).ToString)
'        End If
'        If FileMode = DataConnection.FileModes.Append AndAlso AppendFrom > 0 Then
'            SendCommand("REST " + AppendFrom.ToString)
'        End If
'        SendCommand("RETR " + RemoteFile)
'        If lastResponseType <> 4 AndAlso lastResponseType <> 5 Then
'            Try
'                dataSocket.ReceiveFromSocket()
'            Catch
'                dataSocket.Close()
'                WaitForResponse()
'                RaiseEvent CommandCompleted()
'                Return False
'            End Try
'            WaitForResponse()
'        End If
'        dataSocket.Close()
'        RaiseEvent CommandCompleted()
'        Return True
'    End Function
'    '/// <summary>Uploads a file to the remote server.</summary>
'    '/// <param name="LocalFile">Specifies the local file to upload.</param>
'    '/// <param name="RemoteFile">Specifies the remote file to upload to.</param>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Overloads Function UploadFile(ByVal LocalFile As String, ByVal RemoteFile As String) As Boolean
'        Return UploadFile(LocalFile, RemoteFile, StreamModes.Binary, DataConnection.FileModes.Overwrite, 0)
'    End Function
'    '/// <summary>Uploads a file to the remote server.</summary>
'    '/// <param name="LocalFile">Specifies the local file to upload.</param>
'    '/// <param name="RemoteFile">Specifies the remote file to upload to.</param>
'    '/// <param name="StreamMode">Specifies the transfer type.</param>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Overloads Function UploadFile(ByVal LocalFile As String, ByVal RemoteFile As String, ByVal StreamMode As StreamModes) As Boolean
'        Return UploadFile(LocalFile, RemoteFile, StreamMode, DataConnection.FileModes.Overwrite, 0)
'    End Function
'    '/// <summary>Uploads a file to the remote server.</summary>
'    '/// <param name="LocalFile">Specifies the local file to upload.</param>
'    '/// <param name="RemoteFile">Specifies the remote file to upload to.</param>
'    '/// <param name="FileMode">Specifies the file mode.</param>
'    '/// <param name="AppendFrom">Specifies where to resume from.</param>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Overloads Function UploadFile(ByVal LocalFile As String, ByVal RemoteFile As String, ByVal FileMode As DataConnection.FileModes, ByVal AppendFrom As Long) As Boolean
'        Return UploadFile(LocalFile, RemoteFile, StreamModes.Binary, FileMode, AppendFrom)
'    End Function
'    '/// <summary>Uploads a file to the remote server.</summary>
'    '/// <param name="LocalFile">Specifies the local file to upload.</param>
'    '/// <param name="RemoteFile">Specifies the remote file to upload to.</param>
'    '/// <param name="StreamMode">Specifies the transfer type.</param>
'    '/// <param name="FileMode">Specifies the file mode.</param>
'    '/// <param name="AppendFrom">Specifies where to resume from.</param>
'    '/// <returns>Returns True if the request was successful, False otherwise.</returns>
'    Public Overloads Function UploadFile(ByVal LocalFile As String, ByVal RemoteFile As String, ByVal StreamMode As StreamModes, ByVal FileMode As DataConnection.FileModes, ByVal AppendFrom As Long) As Boolean
'        SendCommand("TYPE " + Convert.ToChar(StreamMode))
'        If Not CreateDataSocket(LocalFile, FileMode, DataConnection.StreamDirections.Upload, AppendFrom) Then
'            RaiseEvent CommandCompleted()
'            Return False
'        End If
'        If Not (passiveTransfers) Then
'            Dim MyEndPoint As IPEndPoint = dataSocket.GetLocalEndPoint()
'            SendCommand("PORT " + MyEndPoint.Address.ToString.Replace(".", ",") + "," + CType(Math.Floor(MyEndPoint.Port / 256), Integer).ToString + "," + (MyEndPoint.Port Mod 256).ToString)
'        End If
'        If FileMode = DataConnection.FileModes.Append AndAlso AppendFrom > 0 Then
'            SendCommand("REST " + AppendFrom.ToString)
'            SendCommand("APPE " + RemoteFile)
'        Else
'            SendCommand("STOR " + RemoteFile)
'        End If
'        If lastResponseType <> 4 AndAlso lastResponseType <> 5 Then
'            Try
'                dataSocket.SendToSocket()
'            Catch
'                dataSocket.Close()
'                WaitForResponse()
'                RaiseEvent CommandCompleted()
'                Return False
'            End Try
'            WaitForResponse()
'        End If
'        dataSocket.Close()
'        RaiseEvent CommandCompleted()
'        Return True
'    End Function
'    Public Function remoteFileExists(ByVal strFileName As String) As Boolean
'        Dim bFound As Boolean
'        Dim oFile As FileItem
'        Dim oRep() As FileItem
'        Debug.Assert(IsConnected, "Not is Connected")
'        GetList()
'        oRep = GetDirectoryListing()
'        bFound = False
'        For Each oFile In oRep
'            Console.Out.WriteLine(oFile.FileTitle)
'            If oFile.FileTitle.Equals(strFileName) Then
'                bFound = True
'                Exit For
'            End If
'        Next
'        Return bFound
'    End Function
'    '/// <summary>Creates a data socket.</summary>
'    '/// <param name="StreamDirection">Specifies the direction of the stream.</param>
'    '/// <returns>Returns True if it was successfully created, False otherwise.</returns>
'    Private Overloads Function CreateDataSocket(ByVal StreamDirection As DataConnection.StreamDirections) As Boolean
'        Return CreateDataSocket("", DataConnection.FileModes.Overwrite, StreamDirection, 0)
'    End Function
'    '/// <summary>Creates a data socket.</summary>
'    '/// <param name="Filename">Specifies the local filename.</param>
'    '/// <param name="Filemode">Specifies the file mode.</param>
'    '/// <param name="StreamDirection">Specifies the direction of the stream.</param>
'    '/// <param name="AppendFrom">Specifies where to append from.</param>
'    '/// <returns>Returns True if it was successfully created, False otherwise.</returns>
'    Private Overloads Function CreateDataSocket(ByVal Filename As String, ByVal Filemode As DataConnection.FileModes, ByVal StreamDirection As DataConnection.StreamDirections, ByVal AppendFrom As Long) As Boolean
'        Try
'            dataSocket = New DataConnection
'            dataSocket.Filename = Filename
'            dataSocket.DownloadToFile = (Filename <> "")
'            dataSocket.AppendFrom = AppendFrom
'            dataSocket.FileMode = Filemode
'            dataSocket.StreamDirection = StreamDirection
'            If passiveTransfers Then
'                SendCommand("PASV")
'                If lastResponseType = 2 Then
'                    Dim BeginPos As Integer = lastResponse.IndexOf("(")
'                    Dim EndPos As Integer = lastResponse.IndexOf(")", BeginPos + 1)
'                    If BeginPos > 0 And EndPos > 0 Then
'                        Dim Output() As String = lastResponse.Substring(BeginPos + 1, EndPos - BeginPos - 1).Split(","c)
'                        If Output.Length() = 6 Then
'                            passiveEndPoint = New IPEndPoint(IPAddress.Parse(Output(0) + "." + Output(1) + "." + Output(2) + "." + Output(3)), Integer.Parse(Output(4)) * 256 + Integer.Parse(Output(5)))
'                            dataSocket.RemoteEndPoint = passiveEndPoint
'                            dataSocket.Connect()
'                        End If
'                    End If
'                End If
'            Else
'                dataSocket.Listen()
'            End If
'        Catch
'            Return False
'        End Try
'        Return True
'    End Function
'    '/// <summary>Parses a directory listing.</summary>
'    '/// <param name="DirList">The directory listing to parse.</param>
'    'permissions   number   owner   group   filesize   month   day   hour/year   name
'    Private Sub ParseDirList(ByVal DirList As String)
'        Dim Lines() As String = DirList.Split(Convert.ToChar(10))
'        Dim Items() As String
'        Dim Cnt As Integer
'        Dim tempDate As DateTime
'        Dim HourMin() As String
'        Dim Year As String
'        remoteFileCount = 0
'        If DirList Is Nothing Then Exit Sub
'        For Cnt = 0 To Lines.Length - 1
'            Items = RemoveRedundantItems(Lines(Cnt).Trim.Split(" "c))
'            If Items.Length > 9 Then
'                Items(8) = String.Join(" ", Items, 8, Items.Length - 8)
'                ReDim Preserve Items(8)
'            End If
'            If Items.Length = 9 Then
'                ReDim Preserve remoteFiles(remoteFileCount)
'                remoteFiles(remoteFileCount) = New FileItem
'                With remoteFiles(remoteFileCount)
'                    .FilePermissions = Items(0).Substring(1)
'                    .FileOwner = Items(2)
'                    .FileGroup = Items(3)
'                    .FileSize = Long.Parse(Items(4))
'                    If Items(7).IndexOf(":") >= 0 Then
'                        HourMin = Items(7).Split(":"c)
'                        Year = "2001"
'                    Else
'                        ReDim HourMin(1)
'                        HourMin(0) = "0"
'                        HourMin(1) = "0"
'                        Year = Items(7)
'                    End If
'                    tempDate = New DateTime(Integer.Parse(Year), MonthToNumber(Items(5)), Integer.Parse(Items(6)), Integer.Parse(HourMin(0)), Integer.Parse(HourMin(1)), 0)
'                    .FileDate = tempDate
'                    .FileTitle = Items(8)
'                    .FilePath = remoteDirectory
'                    .IsDirectory = (Items(0).ToLower.Chars(0) = "d"c)
'                End With
'                remoteFileCount = remoteFileCount + 1
'            End If
'        Next Cnt
'        Array.Sort(remoteFiles)
'    End Sub
'    '/// <summary>Returns the number of a month.</summary>
'    '/// <param name="Input">The month.</param>
'    '/// <returns>The number of the specified month.</returns>
'    Private Function MonthToNumber(ByVal Input As String) As Integer
'        Select Case Input.ToLower()
'            Case "jan" : Return 1
'            Case "feb" : Return 2
'            Case "mar" : Return 3
'            Case "apr" : Return 4
'            Case "may" : Return 5
'            Case "jun" : Return 6
'            Case "jul" : Return 7
'            Case "aug" : Return 8
'            Case "sep" : Return 9
'            Case "oct" : Return 10
'            Case "nov" : Return 11
'            Case "dec" : Return 12
'            Case Else : Return Integer.Parse(Input)
'        End Select
'    End Function
'    '/// <summary>Removes the redundant items from an array.</summary>
'    '/// <param name="Input">The array with the redundant items.</param>
'    '/// <returns>The cleaned up array.</returns>
'    Private Function RemoveRedundantItems(ByVal Input() As String) As String()
'        Dim Cnt As Integer, GoodItems As Integer
'        Dim ReturnArray() As String
'        ReDim ReturnArray(Input.Length - 1)
'        For Cnt = 0 To Input.Length - 1
'            If Input(Cnt) <> "" Then
'                ReturnArray(GoodItems) = Input(Cnt)
'                GoodItems = GoodItems + 1
'            End If
'        Next Cnt
'        ReDim Preserve ReturnArray(GoodItems - 1)
'        Return ReturnArray
'    End Function
'    '/// <summary>Retrieves the directory listing.</summary>
'    '/// <returns>The remote directory listing.</returns>
'    Public Function GetDirectoryListing() As FileItem()
'        Return remoteFiles
'    End Function
'    '/// <summary>Retrieves the number of remote files and directories.</summary>
'    '/// <returns>The number of remote files and directories.</returns>
'    Public Function GetDirectoryListCount() As Integer
'        Return remoteFileCount
'    End Function
'    '/// <summary>Retrieves the remote Dir.</summary>
'    '/// <returns>The remote directory.</returns>
'    Public Function GetRemoteDir() As String
'        PrintWorkDir()
'        Return remoteDirectory
'    End Function
'    ' Two methods that can be used by you
'    Private Sub DataReceived(ByVal BytesLength As Integer) Handles dataSocket.DataReceived
'        'Do Nothing
'        'If you want to be notified when the dataSocket receives some bytes,
'        'add your handler here
'    End Sub
'    Private Sub DataSent(ByVal BytesLength As Integer) Handles dataSocket.DataSent
'        'Do Nothing
'        'If you want to be notified when the dataSocket sends some bytes,
'        'add your handler here
'    End Sub
'    ' Private variables
'    Private clientSocket As Socket
'    Private remotePort As Integer = 21
'    Private remoteAddress As String = ""
'    Private logonName As String = "anonymous"
'    Private logonPass As String = "user@"
'    Private ASCII As New ASCIIEncoding
'    Private isNowConnected As Boolean = False
'    Private passiveTransfers As Boolean = False
'    Private lastResponseNumber As Short = 0
'    Private lastResponseType As Short = 0
'    Private lastResponse As String = ""
'    Private remoteDirectory As String = ""
'    Private WithEvents dataSocket As DataConnection
'    Private passiveEndPoint As IPEndPoint
'    Private remoteFiles() As FileItem
'    Private remoteFileCount As Integer = 0
'End Class
