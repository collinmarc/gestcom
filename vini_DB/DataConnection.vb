'Option Strict
'Option Explicit

'Imports System
'Imports System.IO
'Imports System.Text
'Imports System.Net
'Imports System.Net.Sockets

''/// <summary>Represents a data connection for an FTP session.</summary>
''/// <remarks>This class can be used to upload/download a file or to upload/download a string to a buffer.</remarks>
''/// <remarks>Although it is used in the FtpConnection class, it is not FTP dependant; hence, it can be used for other protocols as well.</remarks>
'Public Class DataConnection
'    '/// <summary>Enumerates all the possible values for the file mode.</summary>
'    Public Enum FileModes
'        Overwrite = 0
'        Append = 1
'    End Enum
'    '/// <summary>Enumerates all the possible values for the strean direction.</summary>
'    Public Enum StreamDirections
'        Download = 0
'        Upload = 1
'    End Enum
'    '/// <summary>Raised when the class receives new data.</summary>
'    Event DataReceived(ByVal ByteLen As Integer)
'    '/// <summary>Raised when the class sent new data.</summary>
'    Event DataSent(ByVal ByteLen As Integer)
'    '/// <summary>Specifies the byte where to append from.</summary>
'    '/// <remarks>Byte 0 is the first byte in the stream; if you want to download/upload the entire file, you should set this value to 0.</remarks>
'    '/// <remarks>If you specify a negative value for this property, it will be changed to 0.</remarks>
'    '/// <value>A Long that specifies the byte where to append from.</value>
'    Public Property AppendFrom() As Long
'        Get
'            Return m_AppendFrom
'        End Get
'        Set(ByVal value As Long)
'            If value < 0 Then
'                m_AppendFrom = 0
'            Else
'                m_AppendFrom = value
'            End If
'        End Set
'    End Property
'    '/// <summary>Specifies the filemode.</summary>
'    '/// <value>A member of the FileModes Enum that specifies the filemode.</value>
'    Public Property FileMode() As FileModes
'        Get
'            Return m_FileMode
'        End Get
'        Set(ByVal value As FileModes)
'            m_FileMode = value
'        End Set
'    End Property
'    '/// <summary>Specifies the filename.</summary>
'    '/// <value>A String that specifies the filename.</value>
'    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null)</exceptions>
'    Public Property Filename() As String
'        Get
'            Return m_Filename
'        End Get
'        Set(ByVal value As String)
'            If value Is Nothing Then Throw New ArgumentNullException()
'            m_Filename = value
'        End Set
'    End Property
'    '/// <summary>Returns the port the connection is using.</summary>
'    '/// <value>An Integer that specifies the port the connection is using.</value>
'    Public ReadOnly Property Port() As Integer
'        Get
'            Return m_Port
'        End Get
'    End Property
'    '/// <summary>Specifies the remote IPEndPoint.</summary>
'    '/// <value>An IPEndPoint that specifies the remote end point.</value>
'    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null)</exceptions>
'    Public Property RemoteEndPoint() As IPEndPoint
'        Get
'            Return m_RemoteEndPoint
'        End Get
'        Set(ByVal value As IPEndPoint)
'            If value Is Nothing Then Throw New ArgumentNullException()
'            m_RemoteEndPoint = value
'        End Set
'    End Property
'    '/// <summary>Specifies the direction of the stream.</summary>
'    '/// <value>A member of the StreamDirections Enum that specifies the direction of the stream.</value>
'    Public Property StreamDirection() As StreamDirections
'        Get
'            Return m_StreamDirection
'        End Get
'        Set(ByVal value As StreamDirections)
'            m_StreamDirection = value
'        End Set
'    End Property
'    '/// <summary>Specifies whether to download to a file or not.</summary>
'    '/// <value>A Boolean that specifies whether to download to a file or not.</value>
'    Public Property DownloadToFile() As Boolean
'        Get
'            Return m_DownloadToFile
'        End Get
'        Set(ByVal value As Boolean)
'            m_DownloadToFile = value
'        End Set
'    End Property
'    '/// <summary>Returns the number of transferred bytes.</summary>
'    '/// <value>A Long that specifies the number of transferred bytes.</value>
'    Public ReadOnly Property TransferredBytes() As Long
'        Get
'            Return m_TransferredBytes
'        End Get
'    End Property
'    '/// <summary>Specifies the data that has been downloaded or has to be uploaded.</summary>
'    '/// <value>A String that specifies the data that has been downloaded or has to be uploaded.</value>
'    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null)</exceptions>
'    Public Property Data() As String
'        Get
'            Return m_Data
'        End Get
'        Set(ByVal value As String)
'            If value Is Nothing Then Throw New ArgumentNullException()
'            m_Data = value
'        End Set
'    End Property
'    '/// <summary>Returns the local IPEndPoint.</summary>
'    '/// <value>An IPEndPoint that specifies the local end point.</value>
'    Public ReadOnly Property GetLocalEndPoint() As IPEndPoint
'        Get
'            Return CType(m_DataSocket.LocalEndPoint, IPEndPoint)
'        End Get
'    End Property
'    '/// <summary>Creates or opens the local data file.</summary>
'    '/// <exceptions cref="IOException">Thrown when there was an error while creating/opening the local data file.</exceptions>
'    Public Sub CreateDataFile()
'        If m_Filename = "" Then
'            m_DownloadToFile = False
'        ElseIf m_DownloadToFile Then
'            Try
'                If m_StreamDirection = StreamDirections.Download Then
'                    m_FileStream = New FileStream(m_Filename, System.IO.FileMode.OpenOrCreate)
'                Else
'                    m_FileStream = New FileStream(m_Filename, System.IO.FileMode.Open)
'                End If
'                If (m_FileMode = FileModes.Append) AndAlso (m_AppendFrom > 0) Then
'                    m_FileStream.Seek(m_AppendFrom, SeekOrigin.Begin)
'                    m_TransferredBytes = m_FileStream.Position()
'                End If
'            Catch
'                Throw New IOException()
'            End Try
'        End If
'    End Sub
'    '/// <summary>Connects to the remote host.</summary>
'    '/// <exceptions cref="IOException">Thrown when there was an error while creating/opening the local data file.</exceptions>
'    '/// <exceptions cref="SocketException">Thrown when there was an error while connecting to the remote host.</exceptions>
'    Public Sub Connect()
'        m_TransferredBytes = 0
'        m_Passive = True
'        CreateDataFile()
'        Try
'            m_DataSocket = New Socket(AddressFamily.Unspecified, SocketType.Stream, ProtocolType.Tcp)
'            m_DataSocket.Connect(m_RemoteEndPoint)
'        Catch
'            Throw New SocketException()
'        End Try
'    End Sub
'    '/// <summary>Opens a port and starts listening for connections.</summary>
'    '/// <exceptions cref="IOException">Thrown when there was an error while creating/opening the local data file.</exceptions>
'    '/// <exceptions cref="SocketException">Thrown when there was an error creating the socket.</exceptions>
'    Public Sub Listen()
'        m_TransferredBytes = 0
'        m_Passive = False
'        CreateDataFile()
'        Try
'            m_DataSocket = New Socket(AddressFamily.Unspecified, SocketType.Stream, ProtocolType.Tcp)
'            Dim localIP As IPAddress = Dns.GetHostEntry(Dns.GetHostName()).AddressList(0)
'            m_DataSocket.Bind(New IPEndPoint(localIP, 0))
'            m_Port = CType(m_DataSocket.LocalEndPoint, IPEndPoint).Port
'            m_DataSocket.Listen(0)
'        Catch
'            Throw New SocketException()
'        End Try
'    End Sub
'    '/// <summary>Starts receiving data from the socket.</summary>
'    '/// <exceptions cref="SocketException">Thrown when there was an error with the socket.</exceptions>
'    '/// <exceptions cref="IOException">Thrown when there was an error while creating/opening the local data file.</exceptions>
'    Public Sub ReceiveFromSocket()
'        Dim theSocket As Socket
'        Try
'            If m_Passive Then
'                theSocket = m_DataSocket
'            Else
'                theSocket = m_DataSocket.Accept()
'            End If
'        Catch
'            Throw New SocketException()
'        End Try
'        Try
'            Dim Ret As Integer, Buffer(1023) As Byte
'            Do
'                Ret = theSocket.Receive(Buffer)
'                If Ret > 0 Then
'                    If m_DownloadToFile Then
'                        m_FileStream.Write(Buffer, 0, Ret)
'                    Else
'                        m_Data = m_Data + ASCII.GetString(Buffer, 0, Ret)
'                    End If
'                    m_TransferredBytes = m_TransferredBytes + Ret
'                    RaiseEvent DataReceived(Ret)
'                End If
'            Loop Until Ret = 0
'        Catch ioe As IOException
'            Throw ioe
'        Catch
'            Throw New SocketException()
'        Finally
'            Try
'                theSocket.Shutdown(SocketShutdown.Both)
'            Catch
'            End Try
'            theSocket.Close()
'        End Try
'    End Sub
'    '/// <summary>Starts sending data to the socket.</summary>
'    '/// <exceptions cref="SocketException">Thrown when there was an error with the socket.</exceptions>
'    '/// <exceptions cref="IOException">Thrown when there was an error while creating/opening the local data file.</exceptions>
'    Public Sub SendToSocket()
'        Dim theSocket As Socket
'        Try
'            If m_Passive Then
'                theSocket = m_DataSocket
'            Else
'                theSocket = m_DataSocket.Accept()
'            End If
'        Catch
'            Throw New SocketException()
'        End Try
'        Try
'            Dim Ret As Integer, Buffer(1023) As Byte, StringBuf() As Byte
'            Do
'                If m_DownloadToFile Then
'                    m_FileStream.Seek(m_TransferredBytes, SeekOrigin.Begin)
'                    Ret = m_FileStream.Read(Buffer, 0, 1024)
'                Else
'                    If m_TransferredBytes + 1024 < m_Data.Length Then
'                        StringBuf = ASCII.GetBytes(m_Data.Substring(CType(m_TransferredBytes, Integer), 1024))
'                    Else
'                        StringBuf = ASCII.GetBytes(m_Data.Substring(CType(m_TransferredBytes, Integer)))
'                    End If
'                    Ret = StringBuf.Length
'                    Array.Copy(StringBuf, Buffer, Ret)
'                End If
'                Ret = theSocket.Send(Buffer, Ret, SocketFlags.None)
'                m_TransferredBytes = m_TransferredBytes + Ret
'                RaiseEvent DataSent(Ret)
'            Loop Until Ret = 0
'        Catch ioe As IOException
'            Throw ioe
'        Catch
'            Throw New SocketException()
'        Finally
'            Try
'                theSocket.Shutdown(SocketShutdown.Both)
'            Catch
'            End Try
'            theSocket.Close()
'        End Try
'    End Sub
'    '/// <summary>Closes the socket.</summary>
'    Public Sub Close()
'        Try
'            m_DataSocket.Shutdown(SocketShutdown.Both)
'        Catch
'        End Try
'        m_DataSocket.Close()
'        Try
'            If Not (m_FileStream Is Nothing) Then m_FileStream.Close()
'        Catch
'        End Try
'    End Sub
'    ' Private variables
'    Private m_AppendFrom As Long = 0
'    Private m_FileMode As FileModes
'    Private m_Filename As String = ""
'    Private m_Port As Integer = 0
'    Private m_StreamDirection As StreamDirections
'    Private m_DownloadToFile As Boolean
'    Private m_Hostname As String = ""
'    Private m_DataSocket As Socket
'    Private m_Data As String = ""
'    Private m_TransferredBytes As Long
'    Private ASCII As New ASCIIEncoding()
'    Private m_RemoteEndPoint As IPEndPoint
'    Private m_FileStream As FileStream
'    Private m_Passive As Boolean
'End Class
